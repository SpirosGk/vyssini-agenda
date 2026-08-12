// Βυσσινί Ατζέντα – Monday Production Pipeline
function mondayPipelineProduction() {

  const runStartedAt = Date.now();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    console.log('Υπάρχει ήδη ενεργή εκτέλεση. Θα ξαναδοκιμάσω αργότερα.');
    MPP_scheduleRetry_();
    return;
  }

  let props = null;

  try {

    props = PropertiesService.getScriptProperties();

    const cfg = MPP_config_();
    const week = MPP_getCurrentWeek_(cfg.timezone);

    // ==================================================
    // PRODUCTION SCHEDULE GUARD
    // ==================================================
    // Νέα κανονική εκτέλεση ξεκινά μόνο Δευτέρα (ώρα Ελλάδας).
    // Retry που ξεκίνησε Δευτέρα επιτρέπεται να συνεχιστεί αργότερα
    // όσο το MPP_PENDING_WEEK δείχνει την ίδια εβδομάδα.
    const isMonday = MPP_isMondayNow_(cfg.timezone);
    const pendingWeek = props.getProperty('MPP_PENDING_WEEK');

    if (!isMonday && pendingWeek !== week.key) {
      props.setProperty('MPP_LAST_STATUS', 'SKIPPED_NOT_MONDAY');
      props.setProperty('MPP_LAST_WEEK', week.key);
      console.log('Production guard: σήμερα δεν είναι Δευτέρα και δεν υπάρχει ενεργό retry. Καμία ενέργεια.');
      return;
    }

    if (isMonday && pendingWeek !== week.key) {
      props.setProperty('MPP_PENDING_WEEK', week.key);
    }

    const productionSS = SpreadsheetApp.openById(cfg.productionSpreadsheetId);
    const auditSS = SpreadsheetApp.openById(cfg.auditSpreadsheetId);

    // Production and audit workbooks are kept logically separate.

    // Production: διαβάζουμε SOURCES/HISTORY και γράφουμε ΜΟΝΟ POSTS.
    const sourcesSheet = productionSS.getSheetByName(cfg.sourcesSheet);
    const postsSheet = productionSS.getSheetByName(cfg.postsSheet);
    const historySheet = productionSS.getSheetByName(cfg.historySheet);

    // Audit/staging: συνεχίζουμε να κρατάμε τα TEST φύλλα για έλεγχο.
    const searchSheet = auditSS.getSheetByName(cfg.searchSheet);
    const finalSheet = auditSS.getSheetByName(cfg.finalSheet);
    const postSheet = auditSS.getSheetByName(cfg.postSheet);

    if (!sourcesSheet) throw new Error('Δεν βρέθηκε το production SOURCES.');
    if (!postsSheet) throw new Error('Δεν βρέθηκε το production POSTS.');
    if (!historySheet) throw new Error('Δεν βρέθηκε το production HISTORY.');
    if (!searchSheet) throw new Error('Δεν βρέθηκε το audit SEARCH_TEST_V3.');
    if (!finalSheet) throw new Error('Δεν βρέθηκε το audit FINAL_TEST.');
    if (!postSheet) throw new Error('Δεν βρέθηκε το audit POST_TEST.');

    props.setProperty('MPP_LAST_STATUS', 'RUNNING');
    props.setProperty('MPP_LAST_WEEK', week.key);
    props.setProperty('MPP_LAST_STARTED_AT', new Date().toISOString());

    // ==================================================
    // PENDING POST SHORTCUT
    // ==================================================
    // Αν προηγούμενο run ολοκλήρωσε όλη την έρευνα αλλά το POSTS ήταν
    // προσωρινά κατειλημμένο, ξαναδοκιμάζουμε μόνο το commit χωρίς Tavily/Gemini.
    const pendingPostWeek = props.getProperty('MPP_PENDING_POST_WEEK');
    const pendingPostText = props.getProperty('MPP_PENDING_POST_TEXT');
    const pendingPostDate = props.getProperty('MPP_PENDING_POST_DATE');

    if (pendingPostWeek === week.key && pendingPostText) {

      const pendingPostResult = {
        ready: true,
        status: 'PENDING_READY',
        date: pendingPostDate || week.startFull,
        post: pendingPostText
      };

      const pendingCommit = MPP_commitPostToProduction_(
        postsSheet,
        historySheet,
        pendingPostResult,
        props,
        week
      );

      if (pendingCommit.retryable) {
        props.setProperty('MPP_LAST_STATUS', pendingCommit.status);
        MPP_scheduleRetry_();
        return;
      }

      MPP_clearPendingPost_();
      props.deleteProperty('MPP_PENDING_WEEK');
      MPP_clearRetry_();

      props.setProperty('MPP_LAST_STATUS', 'SUCCESS_' + pendingCommit.status);
      props.setProperty('MPP_LAST_SUCCESS_AT', new Date().toISOString());
      props.deleteProperty('MPP_LAST_ERROR');

      console.log('Pending production commit ολοκληρώθηκε: ' + pendingCommit.status);
      return;
    }

    const tavilyKey = props.getProperty('TAVILY_API_KEY');
    const geminiKey = props.getProperty('GEMINI_API_KEY');
    const geminiModel = props.getProperty('GEMINI_MODEL') || 'gemini-3.6-flash';

    if (!tavilyKey) {
      throw new Error('Δεν βρέθηκε το TAVILY_API_KEY.');
    }

    if (!geminiKey) {
      throw new Error('Δεν βρέθηκε το GEMINI_API_KEY.');
    }

    // ==================================================
    // 1. TAVILY RESEARCH
    // ==================================================

    let candidates = [];

    const cacheWeek = props.getProperty('MPP_TAVILY_WEEK');
    const cacheAt = Number(props.getProperty('MPP_TAVILY_AT') || '0');

    const cacheIsFresh =
      cacheWeek === week.key &&
      cacheAt > 0 &&
      (Date.now() - cacheAt) < (30 * 60 * 1000) &&
      searchSheet.getLastRow() >= 2;

    if (cacheIsFresh) {

      console.log('Χρησιμοποιώ Tavily cache της ίδιας εβδομάδας.');
      candidates = MPP_readCandidatesFromSearch_(searchSheet, cfg);

    } else {

      console.log('Ξεκινά νέα Tavily έρευνα για ' + week.startFull + ' - ' + week.endFull);

      const research = MPP_runTavilyResearch_(
        sourcesSheet,
        tavilyKey,
        cfg,
        week,
        runStartedAt
      );

      if (!research.ok) {

        props.setProperty('MPP_LAST_STATUS', 'TAVILY_RETRY_SCHEDULED');
        props.setProperty('MPP_LAST_ERROR', String(research.status) + ' | ' + research.text);

        if (research.retryable) {
          MPP_scheduleRetry_();
          return;
        }

        throw new Error('Tavily API error ' + research.status + ': ' + research.text);
      }

      // Γράφουμε SEARCH_TEST_V3 μόνο αφού ολοκληρωθούν ΟΛΑ τα Tavily queries.
      MPP_writeSearch_(searchSheet, research.rows, week);

      candidates = research.candidates;

      props.setProperty('MPP_TAVILY_WEEK', week.key);
      props.setProperty('MPP_TAVILY_AT', String(Date.now()));

      console.log('Tavily ολοκληρώθηκε. Candidates: ' + candidates.length);
    }

    if (candidates.length === 0) {
      throw new Error('Δεν βρέθηκαν πραγματικά Tavily candidates με URL.');
    }

    // Προστασία από το όριο ~6 λεπτών του Apps Script.
    // Αν η αναζήτηση πήρε πολύ χρόνο, έχουμε ήδη αποθηκεύσει το Tavily cache,
    // οπότε το επόμενο retry θα ξεκινήσει από το Gemini.
    if ((Date.now() - runStartedAt) > (4 * 60 * 1000)) {
      props.setProperty('MPP_LAST_STATUS', 'TIME_BUDGET_RETRY_SCHEDULED');
      MPP_scheduleRetry_();
      return;
    }

    // ==================================================
    // 2. GEMINI VALIDATION
    // ==================================================

    const geminiCandidates = MPP_prepareGeminiCandidates_(candidates, cfg);

    console.log('Στέλνω στο Gemini candidates: ' + geminiCandidates.length);

    const geminiResult = MPP_runGeminiValidation_(
      geminiCandidates,
      geminiKey,
      geminiModel,
      cfg,
      week
    );

    if (!geminiResult.ok) {

      props.setProperty('MPP_LAST_STATUS', 'GEMINI_RETRY_SCHEDULED');
      props.setProperty('MPP_LAST_ERROR', String(geminiResult.status) + ' | ' + geminiResult.text);

      if (geminiResult.retryable) {
        MPP_scheduleRetry_();
        return;
      }

      throw new Error('Gemini API error ' + geminiResult.status + ': ' + geminiResult.text);
    }

    // ==================================================
    // 3. ΔΕΥΤΕΡΟ VALIDATION ΑΠΟ JAVASCRIPT
    // ==================================================

    const validatedMatches = MPP_validateMatches_(
      geminiResult.matches,
      geminiCandidates,
      cfg,
      week
    );

    // ==================================================
    // 4. URL CONTEXT SOURCE VERIFICATION
    // ==================================================

    let finalMatches = validatedMatches;

    if (validatedMatches.length > 0) {

      if ((Date.now() - runStartedAt) > (4.5 * 60 * 1000)) {
        props.setProperty('MPP_LAST_STATUS', 'URL_VERIFY_TIME_BUDGET_RETRY');
        MPP_scheduleRetry_();
        return;
      }

      const verifyPlan = MPP_buildUrlVerificationPlan_(
        validatedMatches,
        geminiCandidates,
        cfg
      );

      console.log(
        'URL verification: events=' +
        verifyPlan.events.length +
        ', urls=' +
        verifyPlan.urls.length
      );

      const verifyResult = MPP_runUrlVerification_(
        verifyPlan,
        geminiKey,
        geminiModel
      );

      if (!verifyResult.ok) {

        props.setProperty('MPP_LAST_STATUS', 'URL_VERIFY_RETRY_SCHEDULED');
        props.setProperty(
          'MPP_LAST_ERROR',
          String(verifyResult.status) + ' | ' + verifyResult.text
        );

        if (verifyResult.retryable) {
          MPP_scheduleRetry_();
          return;
        }

        throw new Error(
          'Gemini URL verification error ' +
          verifyResult.status +
          ': ' +
          verifyResult.text
        );
      }

      finalMatches = MPP_applyUrlVerification_(
        validatedMatches,
        verifyResult.verifications,
        verifyPlan,
        geminiCandidates,
        cfg
      );
    }

    // ==================================================
    // 5. ΓΡΑΦΗ FINAL_TEST
    // ==================================================

    MPP_writeFinal_(finalSheet, finalMatches);

    // ==================================================
    // 6. ΔΗΜΙΟΥΡΓΙΑ POST_TEST (AUDIT)
    // ==================================================

    const postResult = MPP_writePost_(postSheet, finalMatches, cfg, week);

    // ==================================================
    // 7. PRODUCTION COMMIT -> POSTS!A2:B2
    // ==================================================
    // Το post αποθηκεύεται πρώτα σε Script Properties ώστε, αν το POSTS
    // είναι προσωρινά κατειλημμένο, το retry να ΜΗΝ ξανατρέχει APIs.
    if (postResult.ready) {
      props.setProperty('MPP_PENDING_POST_WEEK', week.key);
      props.setProperty('MPP_PENDING_POST_DATE', postResult.date);
      props.setProperty('MPP_PENDING_POST_TEXT', postResult.post);
    }

    const commitResult = MPP_commitPostToProduction_(
      postsSheet,
      historySheet,
      postResult,
      props,
      week
    );

    if (commitResult.retryable) {
      props.setProperty('MPP_LAST_STATUS', commitResult.status);
      MPP_scheduleRetry_();
      return;
    }

    // ==================================================
    // 8. SUCCESS
    // ==================================================

    MPP_clearPendingPost_();
    props.deleteProperty('MPP_PENDING_WEEK');
    MPP_clearRetry_();

    props.setProperty('MPP_LAST_STATUS', 'SUCCESS_' + commitResult.status);
    props.setProperty('MPP_LAST_SUCCESS_AT', new Date().toISOString());
    props.deleteProperty('MPP_LAST_ERROR');

    console.log(
      'mondayPipelineProduction SUCCESS. Final verified events: ' +
      finalMatches.length +
      ' | commit=' +
      commitResult.status
    );

  } catch (error) {

    if (props) {
      props.setProperty('MPP_LAST_STATUS', 'FAILED_NON_RETRYABLE');
      props.setProperty('MPP_LAST_ERROR', String(error));
    }

    throw error;

  } finally {
    lock.releaseLock();
  }
}


// ======================================================
// RETRY HANDLER
// Ξεχωριστό handler ώστε να ΜΗΝ πειράζει μελλοντικό weekly trigger
// του mondayPipelineProduction().
// ======================================================

function mondayPipelineProductionRetry() {
  mondayPipelineProduction();
}


// ======================================================
// WEEKLY TRIGGER CREATOR
// Κρατά το ίδιο production handler, άρα ο υπάρχων trigger συνεχίζει να ισχύει.
// Τρέξ' το μόνο αν χρειαστεί ποτέ να ξαναδημιουργήσεις το Monday trigger.
// ======================================================

function createMondayProductionTrigger() {

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'mondayPipelineProduction') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('mondayPipelineProduction')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .nearMinute(0)
    .inTimezone('Europe/Athens')
    .create();

  console.log('Monday production trigger δημιουργήθηκε για ~08:00 Europe/Athens.');
}


// ======================================================
// ΠΡΟΑΙΡΕΤΙΚΟ MANUAL RESET CACHE
// Χρήσιμο μόνο αν θέλεις να αναγκάσεις νέα Tavily έρευνα
// μέσα στην ίδια εβδομάδα.
// ======================================================

function resetMondayPipelineProductionCache() {

  const props = PropertiesService.getScriptProperties();

  props.deleteProperty('MPP_TAVILY_WEEK');
  props.deleteProperty('MPP_TAVILY_AT');
  props.deleteProperty('MPP_LAST_ERROR');

  MPP_clearRetry_();

  console.log('Το Monday PRODUCTION cache καθαρίστηκε.');
}


// ======================================================
// CONFIG
// ======================================================

function MPP_config_() {

  const props = PropertiesService.getScriptProperties();

  const productionSpreadsheetId =
    props.getProperty('PRODUCTION_SPREADSHEET_ID');

  const auditSpreadsheetId =
    props.getProperty('AUDIT_SPREADSHEET_ID') ||
    productionSpreadsheetId;

  if (!productionSpreadsheetId) {
    throw new Error(
      'Δεν βρέθηκε το PRODUCTION_SPREADSHEET_ID στις Script Properties.'
    );
  }

  return {

    productionSpreadsheetId: productionSpreadsheetId,

    auditSpreadsheetId: auditSpreadsheetId,

    sourcesSheet: 'SOURCES',
    postsSheet: 'POSTS',
    historySheet: 'HISTORY',

    searchSheet: 'SEARCH_TEST_V3',
    finalSheet: 'FINAL_TEST',
    postSheet: 'POST_TEST',

    timezone: 'Europe/Athens',

    // PILOT FINAL: 13 ενεργά τμήματα.
    // BASKET_WOMEN και TRACK αφαιρέθηκαν από την εβδομαδιαία ατζέντα.
    allowedDepartments: [
      'FOOTBALL_MEN',
      'FOOTBALL_K19',
      'FOOTBALL_K17',
      'FOOTBALL_K15',
      'FOOTBALL_WOMEN',
      'FOOTBALL_WOMEN_K17',
      'FUTSAL',
      'BASKET_MEN',
      'BASKET_YOUTH',
      'VOLLEY_WOMEN',
      'VOLLEY_MEN',
      'VOLLEY_YOUTH',
      'BOXING'
    ],

    priorityIDs: [
      'FOOTBALL_MEN',
      'BASKET_MEN',
      'VOLLEY_WOMEN',
      'VOLLEY_MEN',
      'FUTSAL'
    ],

    searchTerms: {
      FOOTBALL_MEN: 'ΑΕΛ Λάρισα ποδόσφαιρο',
      FOOTBALL_K19: 'ΑΕΛ Λάρισας Κ19 ποδόσφαιρο',
      FOOTBALL_K17: 'ΑΕΛ Λάρισας Κ17 ποδόσφαιρο',
      FOOTBALL_K15: 'ΑΕΛ Λάρισας Κ15 ποδόσφαιρο',
      FOOTBALL_WOMEN: 'ΑΕΛ Λάρισας ποδόσφαιρο γυναικών',
      FOOTBALL_WOMEN_K17: 'ΑΕΛ Λάρισας Κ17 γυναικών ποδόσφαιρο',
      FUTSAL: 'ΑΕΛ Λάρισας futsal',
      BASKET_MEN: 'ΑΕΛ Λάρισας μπάσκετ ανδρών',
      BASKET_YOUTH: 'ΑΕΛ Λάρισας ακαδημίες μπάσκετ',
      VOLLEY_WOMEN: 'ΑΕΛ Λάρισας βόλεϊ γυναικών',
      VOLLEY_MEN: 'ΑΕΛ Λάρισας βόλεϊ ανδρών',
      VOLLEY_YOUTH: 'ΑΕΛ Λάρισας αναπτυξιακά βόλεϊ',
      BOXING: 'ΑΕΛ Larisa Boxing πυγμαχία'
    },

    excludedDomains: [
      'ael.com.cy',
      'ael-bc.com',
      'cbf.basketball',
      'volleyball.org.cy',
      'cfa.com.cy',
      'apoelfc.com.cy',
      'proinoslogosnews.gr'
    ],

    // ΚΥΜΑ 2: discovery aggregators. Δεν αρκούν μόνοι τους για CONFIRMED.
    aggregatorDomains: [
      'flashscore.com',
      'flashscore.gr',
      'flashscore.co.id',
      'sofascore.com'
    ],

    // ΚΥΜΑ 4: ισχυρές AEL-specific / τοπικές / περιφερειακές πηγές.
    localDomains: [
      'aelole.gr',
      'onlarissa.gr',
      'eleftheria.gr',
      'larissanet.gr',
      'athleticlarissa.gr',
      'aeliko-kafeneio.gr',
      'gegonota.news',
      'pressing.gr',
      'sportrikala.gr'
    ],

    // ΚΥΜΑ 6: μόνο βοηθητικά signals. Ποτέ μόνα τους ως τελική επιβεβαίωση.
    signalDomains: [
      'oddsportal.com',
      'betexplorer.com',
      'stoiximan.gr',
      'novibet.gr',
      'bet365.com'
    ]
  };
}

// ======================================================
// WEEK HELPERS
// ======================================================

function MPP_getCurrentWeek_(timezone) {

  const athensToday = Utilities.formatDate(
    new Date(),
    timezone,
    'yyyy-MM-dd'
  );

  const p = athensToday.split('-').map(Number);

  // UTC noon για να μην έχουμε μετατοπίσεις ημέρας από timezone/DST.
  const base = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12, 0, 0));

  const day = base.getUTCDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;

  const monday = new Date(base.getTime());
  monday.setUTCDate(base.getUTCDate() + diffToMonday);

  const sunday = new Date(monday.getTime());
  sunday.setUTCDate(monday.getUTCDate() + 6);

  const validDates = new Set();

  for (let i = 0; i < 7; i++) {
    const d = new Date(monday.getTime());
    d.setUTCDate(monday.getUTCDate() + i);
    validDates.add(MPP_formatUTCDate_(d, 'full'));
  }

  return {
    monday: monday,
    sunday: sunday,
    startFull: MPP_formatUTCDate_(monday, 'full'),
    endFull: MPP_formatUTCDate_(sunday, 'full'),
    startShort: MPP_formatUTCDate_(monday, 'short'),
    endShort: MPP_formatUTCDate_(sunday, 'short'),
    queryText:
      MPP_greekDateUTC_(monday) +
      ' έως ' +
      MPP_greekDateUTC_(sunday),
    key:
      MPP_formatUTCDate_(monday, 'full') +
      '_' +
      MPP_formatUTCDate_(sunday, 'full'),
    validDates: validDates
  };
}


function MPP_formatUTCDate_(date, mode) {

  const dd = String(date.getUTCDate()).padStart(2, '0');
  const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
  const yyyy = String(date.getUTCFullYear());

  if (mode === 'short') {
    return dd + '/' + mm;
  }

  return dd + '/' + mm + '/' + yyyy;
}


function MPP_greekDateUTC_(date) {

  const months = [
    'Ιανουαρίου',
    'Φεβρουαρίου',
    'Μαρτίου',
    'Απριλίου',
    'Μαΐου',
    'Ιουνίου',
    'Ιουλίου',
    'Αυγούστου',
    'Σεπτεμβρίου',
    'Οκτωβρίου',
    'Νοεμβρίου',
    'Δεκεμβρίου'
  ];

  return (
    date.getUTCDate() +
    ' ' +
    months[date.getUTCMonth()] +
    ' ' +
    date.getUTCFullYear()
  );
}


// ======================================================
// TAVILY RESEARCH
// ======================================================

function MPP_runTavilyResearch_(
  sourcesSheet,
  apiKey,
  cfg,
  week,
  runStartedAt
) {

  const lastRow = sourcesSheet.getLastRow();

  if (lastRow < 8) {
    return {
      ok: false,
      retryable: false,
      status: 'SOURCES_EMPTY',
      text: 'Το SOURCES δεν έχει τις αναμενόμενες γραμμές.'
    };
  }

  const sourceData = sourcesSheet
    .getRange(8, 1, lastRow - 7, 6)
    .getValues();

  const rows = [];
  const candidates = [];
  const seen = new Set();

  function runOne(id, department, sport, type, query, includeDomains) {

    // Αν πλησιάζουμε το Apps Script runtime limit, σταματάμε χωρίς partial write.
    if ((Date.now() - runStartedAt) > (4 * 60 * 1000)) {
      return {
        ok: false,
        retryable: true,
        status: 'TIME_BUDGET',
        text: 'Η Tavily έρευνα ξεπέρασε το ασφαλές time budget.'
      };
    }

    const payload = {
      query: query,
      topic: 'general',
      search_depth: 'basic',
      max_results: 7,
      include_answer: false,
      include_raw_content: false,
      exclude_domains: cfg.excludedDomains
    };

    if (includeDomains && includeDomains.length > 0) {
      payload.include_domains = includeDomains;
    } else {
      payload.country = 'greece';
    }

    const result = MPP_callTavilyWithRetry_(apiKey, payload);

    if (!result.ok) return result;

    const data = result.data;

    if (!data.results || data.results.length === 0) {

      rows.push([
        id,
        department,
        sport,
        type,
        query,
        'ΚΑΝΕΝΑ ΑΠΟΤΕΛΕΣΜΑ',
        '',
        '',
        '',
        '',
        ''
      ]);

      return { ok: true };
    }

    data.results.forEach(function(item) {

      const url = item.url || '';
      if (!url) return;

      const unique = id + '|' + url;

      // Το πρώτο κύμα που βρίσκει ένα URL κρατά το search_type του.
      if (seen.has(unique)) return;
      seen.add(unique);

      const candidate = {
        id: id,
        department: department,
        sport: sport,
        search_type: type,
        query: query,
        title: item.title || '',
        domain: MPP_getDomain_(url),
        url: url,
        score: item.score || '',
        official: type === 'W1_OFFICIAL' || type === 'W3_TEAM' ? 'ΝΑΙ' : 'ΟΧΙ',
        content: item.content || ''
      };

      candidates.push(candidate);

      rows.push([
        candidate.id,
        candidate.department,
        candidate.sport,
        candidate.search_type,
        candidate.query,
        candidate.title,
        candidate.domain,
        candidate.url,
        candidate.score,
        candidate.official,
        candidate.content
      ]);
    });

    return { ok: true };
  }

  function enoughEvidence(id) {
    return MPP_hasEnoughDiscoveryEvidence_(
      candidates.filter(function(c) { return c.id === id; }),
      week
    );
  }

  for (let i = 0; i < sourceData.length; i++) {

    const row = sourceData[i];

    const id = String(row[0] || '').trim();
    const department = String(row[1] || '').trim();
    const sport = String(row[2] || '').trim();
    const officialText = String(row[4] || '');
    const secondText = String(row[5] || '');

    if (!id || !department) continue;
    if (!cfg.allowedDepartments.includes(id)) continue;

    const base = cfg.searchTerms[id] || ('ΑΕΛ Λάρισας ' + department);

    // --------------------------------------------------
    // ΚΥΜΑ 1: επίσημη διοργανώτρια / ομοσπονδία
    // --------------------------------------------------
    const organizerDomains = MPP_extractDomains_(officialText);

    if (organizerDomains.length > 0) {
      const r1 = runOne(
        id,
        department,
        sport,
        'W1_OFFICIAL',
        base + ' αγώνας πρόγραμμα αγωνιστική ' + week.queryText,
        organizerDomains
      );
      if (!r1.ok) return r1;
    }

    // Πραγματικό adaptive escalation: αν η επίσημη διοργανώτρια
    // έδωσε ήδη σαφές event της εβδομάδας, σταματάμε εδώ.
    // Για συγκεντρωτικά youth IDs συνεχίζουμε ώστε να μη χαθούν
    // δεύτερες ηλικιακές κατηγορίες μέσα στην ίδια εβδομάδα.
    if (
      !['BASKET_YOUTH', 'VOLLEY_YOUTH'].includes(id) &&
      MPP_hasAuthoritativeWeekEvidence_(
        candidates.filter(function(c) { return c.id === id; }),
        week
      )
    ) {
      continue;
    }

    // --------------------------------------------------
    // ΚΥΜΑ 2A: Flashscore / Sofascore
    // --------------------------------------------------
    const r2a = runOne(
      id,
      department,
      sport,
      'W2_AGGREGATOR',
      base + ' αγώνας ' + week.queryText,
      cfg.aggregatorDomains
    );
    if (!r2a.ok) return r2a;

    // --------------------------------------------------
    // ΚΥΜΑ 2B: γενικό Tavily discovery
    // --------------------------------------------------
    const r2b = runOne(
      id,
      department,
      sport,
      'W2_BROAD',
      base + ' αγώνας πρόγραμμα ' + week.queryText + ' Ελλάδα',
      []
    );
    if (!r2b.ok) return r2b;

    // Αν έχουμε ήδη πολύ ισχυρό discovery evidence, δεν καίμε credits άσκοπα.
    if (enoughEvidence(id)) continue;

    // --------------------------------------------------
    // ΚΥΜΑ 3: επίσημη ΑΕΛ / επίσημος αντίπαλος / 2η επίσημη πηγή
    // Από τη στήλη "2η ΠΗΓΗ / ΑΕΛ" του SOURCES.
    // --------------------------------------------------
    const teamDomains = MPP_extractDomains_(secondText);

    if (teamDomains.length > 0) {
      const r3 = runOne(
        id,
        department,
        sport,
        'W3_TEAM',
        base + ' αγώνας ανακοίνωση πρόγραμμα ' + week.queryText,
        teamDomains
      );
      if (!r3.ok) return r3;
    }

    if (enoughEvidence(id)) continue;

    // --------------------------------------------------
    // ΚΥΜΑ 4: ισχυρές AEL-specific / τοπικές πηγές
    // --------------------------------------------------
    const r4 = runOne(
      id,
      department,
      sport,
      'W4_LOCAL',
      base + ' αγώνας πρόγραμμα ' + week.queryText,
      cfg.localDomains
    );
    if (!r4.ok) return r4;

    if (enoughEvidence(id)) continue;

    // --------------------------------------------------
    // ΚΥΜΑ 5: niche/deep search
    // Ειδικά για μικρές κατηγορίες, προγράμματα διαιτητών,
    // τοπικές ενώσεις και επίσημα sites αντιπάλων.
    // --------------------------------------------------
    const r5 = runOne(
      id,
      department,
      sport,
      'W5_NICHE',
      base +
        ' πρόγραμμα αγώνων αγωνιστική διαιτητές γήπεδο ώρα ' +
        week.queryText,
      []
    );
    if (!r5.ok) return r5;

    if (enoughEvidence(id)) continue;

    // --------------------------------------------------
    // ΚΥΜΑ 6: βοηθητικά signals / odds
    // ΔΕΝ επιτρέπεται να επιβεβαιώσουν μόνα τους έναν αγώνα.
    // --------------------------------------------------
    const r6 = runOne(
      id,
      department,
      sport,
      'W6_SIGNAL',
      base + ' αγώνας αποδόσεις odds ' + week.queryText,
      cfg.signalDomains
    );
    if (!r6.ok) return r6;
  }

  return {
    ok: true,
    retryable: false,
    status: 200,
    text: '',
    rows: rows,
    candidates: candidates
  };
}



function MPP_hasAuthoritativeWeekEvidence_(departmentCandidates, week) {

  const official = (departmentCandidates || []).filter(function(candidate) {
    return candidate.search_type === 'W1_OFFICIAL' && candidate.url;
  });

  return official.some(function(candidate) {

    const text = MPP_normText_(
      String(candidate.title || '') + ' ' +
      String(candidate.url || '') + ' ' +
      String(candidate.content || '').substring(0, 1000)
    );

    if (!MPP_mentionsAEL_(text)) return false;

    const hasEventLanguage = MPP_containsAny_(text, [
      'αγων', 'αναμετρ', 'προγραμμα', 'αγωνιστικ', 'κυπελλ', 'playoff'
    ]);

    if (!hasEventLanguage) return false;

    const dates = week && week.validDates
      ? Array.from(week.validDates)
      : [];

    return dates.some(function(dateText) {
      return MPP_textHasEventDate_(text, dateText);
    });
  });
}


function MPP_hasEnoughDiscoveryEvidence_(departmentCandidates, week) {

  const list = departmentCandidates || [];
  if (list.length === 0) return false;

  const eventLike = list.filter(function(candidate) {
    return MPP_candidateLooksEventLike_(candidate, week);
  });

  if (eventLike.length === 0) return false;

  // Ένα event-like αποτέλεσμα από επίσημη διοργανώτρια είναι αρκετό
  // για να σταματήσει το discovery escalation. Η τελική επιβεβαίωση
  // γίνεται αργότερα από Gemini + Hybrid URL verification.
  if (eventLike.some(function(c) {
    return c.search_type === 'W1_OFFICIAL';
  })) {
    return true;
  }

  const domains = new Set(eventLike.map(function(c) {
    return c.domain;
  }).filter(Boolean));

  const hasDirectNonAggregator = eventLike.some(function(c) {
    return !MPP_isAggregatorUrl_(c.url) && !MPP_isSignalUrl_(c.url);
  });

  const hasAggregator = eventLike.some(function(c) {
    return MPP_isAggregatorUrl_(c.url);
  });

  // Δύο ανεξάρτητα domains και τουλάχιστον ένα direct/non-aggregator
  // θεωρούνται αρκετό discovery evidence για να μη γίνουν άσκοπα βαθύτερα queries.
  if (domains.size >= 2 && hasDirectNonAggregator) return true;

  // Aggregator + direct broad source είναι επίσης επαρκές discovery signal.
  if (hasAggregator && hasDirectNonAggregator) return true;

  return false;
}


function MPP_candidateLooksEventLike_(candidate, week) {

  if (!candidate || !candidate.url) return false;

  const text = MPP_normText_(
    String(candidate.title || '') + ' ' +
    String(candidate.url || '') + ' ' +
    String(candidate.content || '').substring(0, 900)
  );

  if (!MPP_mentionsAEL_(text)) return false;

  const eventWords = [
    'αγων',
    'αναμετρ',
    'προγραμμα',
    'αγωνιστικ',
    'πρωταθλη',
    'κυπελλ',
    'playoff',
    'play out',
    'φιλικ',
    'διαιτητ',
    'γηπεδ',
    'ωρα'
  ];

  if (!MPP_containsAny_(text, eventWords)) return false;

  // Bonus ασφάλεια: αν υπάρχει ημερομηνία της εβδομάδας στο snippet/title,
  // θεωρείται σαφώς event-like. Αν όχι, αφήνουμε και direct πρόγραμμα/αγωνιστική
  // να περάσει ώστε να μη χάνουμε tables που δεν επαναλαμβάνουν πλήρη ημερομηνία.
  if (week && week.validDates) {
    const dates = Array.from(week.validDates);
    const hasWeekDate = dates.some(function(dateText) {
      return MPP_textHasEventDate_(text, dateText);
    });

    if (hasWeekDate) return true;
  }

  return (
    text.includes('προγραμμα') ||
    text.includes('αγωνιστικ') ||
    text.includes('διαιτητ') ||
    text.includes('playoff') ||
    text.includes('κυπελλ')
  );
}


function MPP_callTavilyWithRetry_(apiKey, payload) {

  const retryableCodes = new Set([429, 500, 502, 503, 504]);
  const waits = [0, 5000, 10000, 20000];

  let lastStatus = 0;
  let lastText = '';

  for (let attempt = 0; attempt < waits.length; attempt++) {

    if (waits[attempt] > 0) {
      console.log(
        'Tavily retry ' +
        (attempt + 1) +
        '/' +
        waits.length +
        ' σε ' +
        (waits[attempt] / 1000) +
        ' sec...'
      );

      Utilities.sleep(waits[attempt]);
    }

    try {

      const response = UrlFetchApp.fetch(
        'https://api.tavily.com/search',
        {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'Authorization': 'Bearer ' + apiKey
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        }
      );

      lastStatus = response.getResponseCode();
      lastText = response.getContentText();

      if (lastStatus === 200) {

        try {
          return {
            ok: true,
            retryable: false,
            status: 200,
            text: lastText,
            data: JSON.parse(lastText)
          };
        } catch (parseError) {
          lastStatus = 0;
          lastText = 'Tavily JSON parse error: ' + String(parseError);
          continue;
        }
      }

      if (!retryableCodes.has(lastStatus)) {
        return {
          ok: false,
          retryable: false,
          status: lastStatus,
          text: lastText
        };
      }

      console.log('Προσωρινό Tavily error: ' + lastStatus);

    } catch (error) {

      lastStatus = 0;
      lastText = String(error);
      console.log('Tavily UrlFetch exception: ' + lastText);
    }
  }

  return {
    ok: false,
    retryable: true,
    status: lastStatus,
    text: lastText
  };
}


function MPP_writeSearch_(sheet, rows, week) {

  sheet.clearContents();

  sheet.getRange('A1:K1').setValues([[
    'ID',
    'ΤΜΗΜΑ',
    'ΑΘΛΗΜΑ',
    'SEARCH TYPE',
    'QUERY',
    'ΤΙΤΛΟΣ',
    'DOMAIN',
    'URL',
    'SCORE',
    'OFFICIAL',
    'ΠΕΡΙΕΧΟΜΕΝΟ'
  ]]);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  sheet.setFrozenRows(1);

  sheet.getRange('M1').setValue('ΕΒΔΟΜΑΔΑ');
  sheet.getRange('N1').setValue(week.startFull + ' - ' + week.endFull);
  sheet.getRange('M2').setValue('STATUS');
  sheet.getRange('N2').setValue('TAVILY_OK');
  sheet.getRange('M3').setValue('UPDATED');
  sheet.getRange('N3').setValue(new Date());

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 110);
  sheet.setColumnWidth(5, 450);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 180);
  sheet.setColumnWidth(8, 480);
  sheet.setColumnWidth(9, 90);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 650);

  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 11).setWrap(true);
  }
}


function MPP_readCandidatesFromSearch_(sheet, cfg) {

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const values = sheet
    .getRange(2, 1, lastRow - 1, 11)
    .getDisplayValues();

  return values
    .map(function(r) {
      return {
        id: r[0],
        department: r[1],
        sport: r[2],
        search_type: r[3],
        query: r[4],
        title: r[5],
        domain: r[6],
        url: r[7],
        score: r[8],
        official: r[9],
        content: r[10]
      };
    })
    .filter(function(item) {
      return (
        cfg.allowedDepartments.includes(item.id) &&
        item.url &&
        !String(item.title).startsWith('ERROR') &&
        item.title !== 'ΚΑΝΕΝΑ ΑΠΟΤΕΛΕΣΜΑ'
      );
    });
}


// ======================================================
// GEMINI
// ======================================================

function MPP_prepareGeminiCandidates_(candidates, cfg) {

  const grouped = {};

  candidates.forEach(function(item) {

    if (!cfg.allowedDepartments.includes(item.id)) return;

    if (!grouped[item.id]) grouped[item.id] = [];

    grouped[item.id].push({
      id: item.id,
      department: item.department,
      sport: item.sport,
      search_type: item.search_type,
      query: item.query,
      title: item.title,
      domain: item.domain,
      url: item.url,
      score: item.score,
      official: item.official,
      content: String(item.content || '').substring(0, 950)
    });
  });

  const finalList = [];

  Object.keys(grouped).forEach(function(id) {

    grouped[id].sort(function(a, b) {
      return MPP_candidatePriorityScore_(b) - MPP_candidatePriorityScore_(a);
    });

    // Περιορίζουμε το token load χωρίς να χάνουμε τα βαθύτερα κύματα.
    grouped[id].slice(0, 18).forEach(function(item) {
      finalList.push(item);
    });
  });

  return finalList;
}


function MPP_candidatePriorityScore_(candidate) {

  const waveRank = {
    W1_OFFICIAL: 100,
    W3_TEAM: 88,
    W4_LOCAL: 68,
    W5_NICHE: 60,
    W2_BROAD: 56,
    W2_AGGREGATOR: 48,
    W6_SIGNAL: 10
  };

  let score = waveRank[candidate.search_type] || 30;

  const tavilyScore = Number(
    String(candidate.score || '0').replace(',', '.')
  ) || 0;

  score += Math.min(20, tavilyScore * 20);

  const quickText = MPP_normText_(
    String(candidate.title || '') + ' ' +
    String(candidate.content || '').substring(0, 500)
  );

  if (MPP_mentionsAEL_(quickText)) score += 8;

  if (MPP_containsAny_(quickText, [
    'αγων', 'αναμετρ', 'προγραμμα', 'αγωνιστικ', 'κυπελλ', 'playoff'
  ])) {
    score += 10;
  }

  if (MPP_isOfficialSportsUrl_(candidate.url)) score += 20;
  if (MPP_isTrustedSportsMediaUrl_(candidate.url)) score += 8;
  if (MPP_isSignalUrl_(candidate.url)) score -= 20;

  return score;
}

function MPP_runGeminiValidation_(
  candidates,
  apiKey,
  model,
  cfg,
  week
) {

  const prompt = `
Είσαι ο τελικός ελεγκτής της "Βυσσινί Ατζέντας".

Σου δίνονται ακατέργαστα αποτελέσματα web search από Tavily.
Πρέπει να εντοπίσεις ΟΛΑ και ΜΟΝΟ τα πραγματικά αθλητικά γεγονότα
της Αθλητικής Ένωσης Λάρισας (ΑΕΛ) που πραγματοποιούνται:

ΑΠΟ: ${week.startFull}
ΕΩΣ: ${week.endFull}

ΤΜΗΜΑΤΑ:
- FOOTBALL_MEN = ΠΑΕ ΑΕΛ Ανδρών
- FOOTBALL_K19 = ΑΕΛ Κ19
- FOOTBALL_K17 = ΑΕΛ Κ17
- FOOTBALL_K15 = ΑΕΛ Κ15
- FOOTBALL_WOMEN = Ποδόσφαιρο Γυναικών
- FOOTBALL_WOMEN_K17 = Ποδόσφαιρο Κ17 Γυναικών
- FUTSAL = ΑΕΛ Futsal
- BASKET_MEN = Μπάσκετ Ανδρών
- BASKET_YOUTH = Ακαδημίες / αναπτυξιακό μπάσκετ
- VOLLEY_WOMEN = Βόλεϊ Γυναικών
- VOLLEY_MEN = Βόλεϊ Ανδρών
- VOLLEY_YOUTH = Αναπτυξιακό βόλεϊ
- BOXING = Πυγμαχία / Larisa Boxing

ΑΥΣΤΗΡΟΙ ΚΑΝΟΝΕΣ:
1. Η ΑΕΛ είναι η Αθλητική Ένωση Λάρισας. Αγνόησε ΑΕΛ Λεμεσού και κάθε άλλη ΑΕΛ.
2. Μην μπερδεύεις τμήματα. Αν αποτέλεσμα εμφανίζεται σε query μπάσκετ αλλά αφορά ΠΑΕ, δεν είναι μπάσκετ.
3. Μην εμπιστεύεσαι μόνο QUERY, SCORE ή SEARCH TYPE. Έλεγξε title, content, URL, domain και department.
4. Επιτρέπονται μόνο γεγονότα από ${week.startFull} έως ${week.endFull}.
5. Αγνόησε παλιούς/μελλοντικούς αγώνες, μεταγραφές, προπονήσεις, παρουσιάσεις, ιστορικά άρθρα, συνεντεύξεις και γενικές ειδήσεις.
6. Για ποδόσφαιρο, μπάσκετ, βόλεϊ και futsal απαιτείται πραγματικός αγώνας.
7. Για πυγμαχία μπορεί να είναι επίσημη διοργάνωση όπου συμμετέχει το Larisa Boxing / τμήμα ΑΕΛ, μόνο με σαφή ταυτοποίηση συμμετοχής.
8. Μην επινοήσεις αντίπαλο, ημερομηνία, ώρα, έδρα ή διοργάνωση.
9. Αν ώρα/έδρα δεν επιβεβαιώνεται, επέστρεψε κενό string "".
10. Αν υπάρχουν διαφορετικές πληροφορίες, προτίμησε την πιο πρόσφατη επίσημη ενημέρωση.
11. Ιεράρχηση κυμάτων: W1_OFFICIAL (διοργανώτρια/ομοσπονδία) > W3_TEAM (επίσημη ΑΕΛ/ομάδα) > W4_LOCAL/W5_NICHE (άμεσες αξιόπιστες τοπικές/ειδικές πηγές) > W2_AGGREGATOR/W2_BROAD για discovery.
12. Flashscore/Sofascore βοηθούν στο discovery και στην επιβεβαίωση όταν συμφωνούν με direct πηγή, αλλά δεν αρκούν μόνα τους για CONFIRMED. W6_SIGNAL/στοιχηματικά signals ΠΟΤΕ δεν αρκούν μόνα τους για CONFIRMED.
13. Πολλά URLs μπορεί να αφορούν τον ίδιο αγώνα. Ενοποίησέ τα σε μία εγγραφή.
14. CONFIRMED = επαρκής επιβεβαίωση. UNCERTAIN = πιθανό αλλά όχι αρκετά επιβεβαιωμένο.
15. HIGH = ισχυρή επίσημη επιβεβαίωση. MEDIUM = αρκετά αξιόπιστο. LOW = αδύναμο.
16. Για BASKET_YOUTH απαιτείται σαφής ταυτοποίηση ότι η ομάδα/ακαδημία εκπροσωπεί την ΑΕΛ. Μην συμπεραίνεις σύνδεση μόνο από συνεργασία ή παρόμοιο όνομα.
17. primary_source και secondary_source πρέπει να είναι URLs που υπάρχουν ΑΚΡΙΒΩΣ στα δεδομένα που σου δίνω.
18. Αν υπάρχει μόνο μία αξιόπιστη πηγή, secondary_source = "".
19. home_team και away_team πρέπει να αποτυπώνουν τον τυπικό γηπεδούχο και φιλοξενούμενο. Η φυσική έδρα δεν αλλάζει απαραίτητα την τυπική γηπεδούχο ομάδα.
20. Για πυγμαχία χωρίς ζευγάρι ομάδων: home_team = "ΑΕΛ / Larisa Boxing" και away_team = διοργάνωση ή "-".
21. Αν η επίσημη πηγή δεν εμφανίζεται εύκολα αλλά δύο ανεξάρτητες direct πηγές συμφωνούν στα βασικά στοιχεία, μπορείς να προτείνεις CONFIRMED/MEDIUM· η JavaScript θα κάνει τελικό URL verification.

ΕΠΕΣΤΡΕΨΕ για κάθε πραγματικό γεγονός:
- department_id
- department
- sport
- home_team
- away_team
- date σε μορφή DD/MM/YYYY
- time σε μορφή HH:MM ή ""
- venue ή ""
- competition
- status
- confidence
- primary_source
- secondary_source
- reason

ΑΚΑΤΕΡΓΑΣΤΑ ΑΠΟΤΕΛΕΣΜΑΤΑ:
${JSON.stringify(candidates)}
`;

  const endpoint =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent';

  const payload = {
    contents: [
      {
        parts: [
          { text: prompt }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      responseJsonSchema: {
        type: 'object',
        properties: {
          matches: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                department_id: { type: 'string' },
                department: { type: 'string' },
                sport: { type: 'string' },
                home_team: { type: 'string' },
                away_team: { type: 'string' },
                date: { type: 'string' },
                time: { type: 'string' },
                venue: { type: 'string' },
                competition: { type: 'string' },
                status: {
                  type: 'string',
                  enum: ['CONFIRMED', 'UNCERTAIN']
                },
                confidence: {
                  type: 'string',
                  enum: ['HIGH', 'MEDIUM', 'LOW']
                },
                primary_source: { type: 'string' },
                secondary_source: { type: 'string' },
                reason: { type: 'string' }
              },
              required: [
                'department_id',
                'department',
                'sport',
                'home_team',
                'away_team',
                'date',
                'time',
                'venue',
                'competition',
                'status',
                'confidence',
                'primary_source',
                'secondary_source',
                'reason'
              ],
              additionalProperties: false
            }
          }
        },
        required: ['matches'],
        additionalProperties: false
      },
      temperature: 0,
      maxOutputTokens: 8192
    }
  };

  const apiResult = MPP_callGeminiWithRetry_(endpoint, apiKey, payload);

  if (!apiResult.ok) return apiResult;

  try {

    const data = JSON.parse(apiResult.text);

    if (
      !data.candidates ||
      !data.candidates.length ||
      !data.candidates[0].content ||
      !data.candidates[0].content.parts
    ) {
      return {
        ok: false,
        retryable: true,
        status: 200,
        text: 'Gemini 200 αλλά χωρίς κανονικό candidate/content.'
      };
    }

    const modelText = data.candidates[0].content.parts
      .filter(function(part) { return part.text; })
      .map(function(part) { return part.text; })
      .join('');

    const parsed = JSON.parse(modelText);

    return {
      ok: true,
      retryable: false,
      status: 200,
      text: apiResult.text,
      matches: parsed.matches || []
    };

  } catch (error) {

    return {
      ok: false,
      retryable: true,
      status: 200,
      text: 'Gemini structured-output parse error: ' + String(error)
    };
  }
}


function MPP_callGeminiWithRetry_(endpoint, apiKey, payload) {

  const retryableCodes = new Set([429, 500, 502, 503, 504]);
  const waits = [0, 5000, 10000, 20000, 40000];

  let lastStatus = 0;
  let lastText = '';

  for (let attempt = 0; attempt < waits.length; attempt++) {

    if (waits[attempt] > 0) {
      console.log(
        'Gemini retry ' +
        (attempt + 1) +
        '/' +
        waits.length +
        ' σε ' +
        (waits[attempt] / 1000) +
        ' sec...'
      );

      Utilities.sleep(waits[attempt]);
    }

    try {

      const response = UrlFetchApp.fetch(
        endpoint,
        {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'x-goog-api-key': apiKey
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        }
      );

      lastStatus = response.getResponseCode();
      lastText = response.getContentText();

      if (lastStatus === 200) {
        return {
          ok: true,
          retryable: false,
          status: 200,
          text: lastText
        };
      }

      if (!retryableCodes.has(lastStatus)) {
        return {
          ok: false,
          retryable: false,
          status: lastStatus,
          text: lastText
        };
      }

      console.log('Προσωρινό Gemini error: ' + lastStatus);

    } catch (error) {

      lastStatus = 0;
      lastText = String(error);
      console.log('Gemini UrlFetch exception: ' + lastText);
    }
  }

  return {
    ok: false,
    retryable: true,
    status: lastStatus,
    text: lastText
  };
}


// ======================================================
// VALIDATION
// ======================================================

function MPP_validateMatches_(matches, candidates, cfg, week) {

  const allowedIDs = new Set(cfg.allowedDepartments);
  const candidatesByDepartment = {};

  candidates.forEach(function(item) {
    if (!candidatesByDepartment[item.id]) {
      candidatesByDepartment[item.id] = [];
    }
    candidatesByDepartment[item.id].push(item);
  });

  const teamSports = new Set([
    'FOOTBALL_MEN',
    'FOOTBALL_K19',
    'FOOTBALL_K17',
    'FOOTBALL_K15',
    'FOOTBALL_WOMEN',
    'FOOTBALL_WOMEN_K17',
    'FUTSAL',
    'BASKET_MEN',
    'BASKET_YOUTH',
    'VOLLEY_WOMEN',
    'VOLLEY_MEN',
    'VOLLEY_YOUTH'
  ]);

  const cleaned = [];

  (matches || []).forEach(function(original) {

    const match = Object.assign({}, original);

    if (!allowedIDs.has(match.department_id)) return;
    if (!week.validDates.has(match.date)) return;

    const departmentCandidates =
      candidatesByDepartment[match.department_id] || [];

    // --------------------------------------------------
    // SOURCE SAFETY
    // --------------------------------------------------
    // Δεν εμπιστευόμαστε απλώς ένα URL επειδή είναι μέσα
    // στα Tavily αποτελέσματα. Πρέπει το ίδιο το αποτέλεσμα
    // να είναι άμεσα σχετικό με τον συγκεκριμένο αγώνα.
    // Αυτό προστατεύει από sitewide "Τελευταία Νέα" snippets.

    const rankedSources = departmentCandidates
      .map(function(candidate) {
        return {
          candidate: candidate,
          score: MPP_sourceRelevanceScore_(candidate, match)
        };
      })
      .filter(function(x) {
        return x.score >= 6;
      })
      .sort(function(a, b) {
        return b.score - a.score;
      });

    const selected = [];
    const usedDomains = new Set();

    rankedSources.forEach(function(x) {

      if (selected.length >= 2) return;

      const c = x.candidate;
      const domain = String(c.domain || '').toLowerCase();

      // Για δεύτερη πηγή προτιμάμε διαφορετικό domain.
      if (selected.length === 1 && usedDomains.has(domain)) {
        return;
      }

      selected.push(c.url);
      usedDomains.add(domain);
    });

    match.primary_source = selected[0] || '';
    match.secondary_source = selected[1] || '';

    // CONFIRMED χωρίς άμεσα σχετική πηγή δεν επιτρέπεται.
    if (
      match.status === 'CONFIRMED' &&
      !match.primary_source
    ) {
      match.status = 'UNCERTAIN';
      match.confidence = 'LOW';
    }

    if (
      match.status === 'CONFIRMED' &&
      match.confidence === 'LOW'
    ) {
      match.status = 'UNCERTAIN';
    }

    if (teamSports.has(match.department_id)) {

      const teamsText = (
        String(match.home_team || '') +
        ' ' +
        String(match.away_team || '')
      ).toUpperCase();

      const mentionsAEL =
        teamsText.includes('ΑΕΛ') ||
        teamsText.includes('AEL') ||
        teamsText.includes('ΛΑΡΙΣ');

      if (!mentionsAEL) return;

      if (!match.home_team || !match.away_team) {
        match.status = 'UNCERTAIN';
        match.confidence = 'LOW';
      }
    }

    cleaned.push(match);
  });

  // Deduplication
  const unique = [];
  const seen = new Set();

  cleaned.forEach(function(match) {

    const key = [
      match.department_id,
      match.home_team,
      match.away_team,
      match.date,
      match.time,
      match.competition
    ]
      .join('|')
      .toLowerCase();

    if (seen.has(key)) return;

    seen.add(key);
    unique.push(match);
  });

  unique.sort(function(a, b) {

    const aKey = MPP_dateSortKey_(a.date, a.time);
    const bKey = MPP_dateSortKey_(b.date, b.time);

    return aKey.localeCompare(bKey);
  });

  return unique;
}


function MPP_sourceRelevanceScore_(candidate, match) {

  const title = MPP_normText_(candidate.title || '');
  const url = MPP_normText_(candidate.url || '');
  const contentStart = MPP_normText_(
    String(candidate.content || '').substring(0, 500)
  );
  const fullContent = MPP_normText_(candidate.content || '');
  const domain = String(candidate.domain || '').toLowerCase();

  const home = MPP_normText_(match.home_team || '');
  const away = MPP_normText_(match.away_team || '');
  const competition = MPP_normText_(match.competition || '');
  const venue = MPP_normText_(match.venue || '');
  const time = MPP_normText_(match.time || '');

  const opponent = MPP_pickOpponent_(home, away);
  const opponentTokens = MPP_flexibleTokens_(MPP_keywords_(opponent));
  const competitionTokens = MPP_keywords_(competition);

  let score = 0;

  const titleUrl = title + ' ' + url;

  if (MPP_containsAny_(titleUrl, opponentTokens)) {
    score += 10;
  }

  if (MPP_containsAny_(contentStart, opponentTokens)) {
    score += 7;
  } else if (MPP_containsAny_(fullContent, opponentTokens)) {
    score += 2;
  }

  if (MPP_containsAny_(titleUrl, competitionTokens)) {
    score += 4;
  }

  if (MPP_containsAny_(contentStart, competitionTokens)) {
    score += 3;
  }

  if (
    time &&
    (titleUrl.includes(time) || contentStart.includes(time))
  ) {
    score += 2;
  }

  if (
    venue &&
    (titleUrl.includes(venue) || contentStart.includes(venue))
  ) {
    score += 2;
  }

  // Αναγνωρισμένοι επίσημοι αθλητικοί domains.
  if (
    domain === 'aelfc.gr' ||
    domain === 'epo.gr' ||
    domain === 'sl2.gr' ||
    domain === 'slgr.gr' ||
    domain === 'e-ael.gr' ||
    domain === 'aelbasket.gr' ||
    domain === 'eskath.gr' ||
    domain === 'volleyball.gr' ||
    domain === 'espekel.com' ||
    domain === 'epssalas.gr' ||
    domain === 'hellenicboxing.org.gr'
  ) {
    score += 3;
  }

  if (candidate.official === 'ΝΑΙ') {
    score += 2;
  }

  if (MPP_isAggregatorUrl_(candidate.url)) {
    score -= 1;
  }

  if (MPP_isSignalUrl_(candidate.url)) {
    score -= 8;
  }

  // Ισχυρή ποινή για άσχετο άρθρο μεταγραφής/διοικητικό
  // που απλώς περιέχει το event μέσα σε sidebar/latest-news.
  const noiseTitlePatterns = [
    'παικτης',
    'αποκτηση',
    'μεταγραφ',
    'συνεργασια',
    'μετοχικ',
    'διοικητικ',
    'νεο δ.σ',
    'ανακοινωνει την αποκτηση'
  ];

  if (
    MPP_containsAny_(title, noiseTitlePatterns) &&
    !MPP_containsAny_(titleUrl, opponentTokens)
  ) {
    score -= 12;
  }

  return score;
}


function MPP_pickOpponent_(home, away) {

  const isAEL = function(text) {
    return (
      text.includes('αελ') ||
      text.includes('ael') ||
      text.includes('λαρισ')
    );
  };

  if (isAEL(home) && !isAEL(away)) return away;
  if (isAEL(away) && !isAEL(home)) return home;

  return home + ' ' + away;
}


function MPP_keywords_(text) {

  const stop = new Set([
    'αο', 'fc', 'παε', 'αελ', 'ael',
    'novibet', 'ανδρων', 'γυναικων',
    'ελλαδας', 'superbet', 'league',
    'κυπελλο', 'πρωταθλημα', 'αγωνας'
  ]);

  return MPP_normText_(text)
    .split(/[^a-z0-9α-ω]+/)
    .filter(function(token) {
      return token.length >= 4 && !stop.has(token);
    });
}


function MPP_flexibleTokens_(tokens) {

  const out = [];

  (tokens || []).forEach(function(token) {

    const value = MPP_normText_(token || '');
    if (!value) return;

    if (!out.includes(value)) out.push(value);

    // Μικρό stem για ελληνικές κλίσεις ονομάτων ομάδων,
    // π.χ. Τρίκαλα -> Τρικάλων. Κρατάμε τουλάχιστον 6 γράμματα
    // ώστε να μην γίνει υπερβολικά γενικό το matching.
    if (value.length >= 7) {
      const stem = value.substring(0, value.length - 1);
      if (stem.length >= 6 && !out.includes(stem)) out.push(stem);
    }
  });

  return out;
}


function MPP_containsAny_(text, tokens) {

  if (!text || !tokens || tokens.length === 0) return false;

  return tokens.some(function(token) {
    return text.includes(token);
  });
}


function MPP_normText_(text) {

  return String(text || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/ς/g, 'σ');
}


function MPP_dateSortKey_(dateText, timeText) {

  const parts = String(dateText || '').split('/');

  if (parts.length !== 3) {
    return String(dateText || '') + ' ' + String(timeText || '');
  }

  return (
    parts[2] +
    '-' +
    parts[1] +
    '-' +
    parts[0] +
    ' ' +
    String(timeText || '')
  );
}



// ======================================================
// HYBRID SOURCE VERIFICATION (V4)
// ======================================================

function MPP_buildUrlVerificationPlan_(matches, candidates, cfg) {

  const byDepartment = {};

  candidates.forEach(function(item) {
    if (!byDepartment[item.id]) byDepartment[item.id] = [];
    byDepartment[item.id].push(item);
  });

  const eventPlans = [];

  (matches || []).forEach(function(match, index) {

    const eventKey = MPP_eventKey_(match, index);
    const pool = byDepartment[match.department_id] || [];

    const ranked = pool
      .map(function(candidate) {
        return {
          candidate: candidate,
          score: MPP_sourceRelevanceScore_(candidate, match) +
            MPP_directPageBonus_(candidate, match)
        };
      })
      .filter(function(x) {
        return x.candidate.url && x.score >= 3;
      })
      .sort(function(a, b) {
        return b.score - a.score;
      });

    const selected = [];
    const seen = new Set();

    [match.primary_source, match.secondary_source].forEach(function(url) {
      if (!url || seen.has(url)) return;
      const found = pool.find(function(c) { return c.url === url; });
      if (!found) return;

      // Δεν σπαταλάμε URL Context slot σε εμφανώς άσχετη παλιά επιλογή.
      if (MPP_directPageBonus_(found, match) < 0) return;

      selected.push(found);
      seen.add(url);
    });

    ranked.forEach(function(x) {
      if (selected.length >= 6) return;
      const c = x.candidate;
      if (seen.has(c.url)) return;
      selected.push(c);
      seen.add(c.url);
    });

    eventPlans.push({
      event_key: eventKey,
      department_id: match.department_id,
      home_team: match.home_team,
      away_team: match.away_team,
      date: match.date,
      time: match.time || '',
      venue: match.venue || '',
      competition: match.competition || '',
      candidate_urls: selected.map(function(c) { return c.url; })
    });
  });

  const urls = [];
  const seenUrls = new Set();

  for (let pass = 0; pass < 6 && urls.length < 20; pass++) {
    eventPlans.forEach(function(event) {
      if (urls.length >= 20) return;
      const url = event.candidate_urls[pass];
      if (!url || seenUrls.has(url)) return;
      urls.push(url);
      seenUrls.add(url);
    });
  }

  eventPlans.forEach(function(event) {
    event.candidate_urls = event.candidate_urls.filter(function(url) {
      return seenUrls.has(url);
    });
  });

  return {
    events: eventPlans,
    urls: urls
  };
}


function MPP_runUrlVerification_(plan, apiKey, model) {

  if (!plan || !plan.events || plan.events.length === 0) {
    return {
      ok: true,
      retryable: false,
      status: 200,
      text: '',
      verifications: []
    };
  }

  if (!plan.urls || plan.urls.length === 0) {
    return {
      ok: true,
      retryable: false,
      status: 200,
      text: '',
      verifications: plan.events.map(function(event) {
        return {
          event_key: event.event_key,
          source_assessments: [],
          reason: 'Δεν υπήρχε URL κατάλληλο για URL Context verification.'
        };
      })
    };
  }

  const prompt = `
Είσαι ο δεύτερος και αυστηρότερος ελεγκτής πηγών της "Βυσσινί Ατζέντας".

Χρησιμοποίησε το URL Context tool και προσπάθησε να ΑΝΟΙΞΕΙΣ τις πραγματικές URLs που σου δίνω.

ΣΤΟΧΟΣ:
Για ΚΑΘΕ event και για ΚΑΘΕ candidate URL του event, επέστρεψε αξιολόγηση της πραγματικής σελίδας.

ΠΟΛΥ ΣΗΜΑΝΤΙΚΟ:
Το ότι μια URL δεν μπορεί να ανοιχτεί από URL Context ΔΕΝ σημαίνει ότι είναι λανθασμένη πηγή.
Σε αυτή την περίπτωση πρέπει να επιστρέψεις access="UNAVAILABLE" και support="NONE".
Η JavaScript θα κάνει ξεχωριστό fallback έλεγχο με τα Tavily title/content δεδομένα.

ΚΑΝΟΝΕΣ:
1. access="OPENED" μόνο αν μπόρεσες πράγματι να διαβάσεις την πραγματική σελίδα με URL Context.
2. access="UNAVAILABLE" αν η σελίδα δεν μπόρεσε να φορτωθεί, μπλοκαρίστηκε, ήταν μη προσβάσιμη ή το tool δεν επέστρεψε το περιεχόμενό της.
3. DIRECT = ο ΤΙΤΛΟΣ, το ΚΥΡΙΟ ΑΡΘΡΟ/ΚΥΡΙΟ ΣΩΜΑ ή ένας επίσημος πίνακας/πρόγραμμα της ίδιας σελίδας αφορά άμεσα το συγκεκριμένο event.
4. Αναφορά μόνο σε sidebar, "Τελευταία Νέα", related articles, footer, navigation, ticker ή widget ΔΕΝ είναι DIRECT.
5. Αν η σελίδα είναι για μεταγραφή, σεμινάριο, διοικητικό θέμα ή άλλον αγώνα και απλώς εμφανίζει το event σε "Τελευταία Νέα", support="NONE".
6. PARTIAL = η κύρια σελίδα δίνει πραγματική αλλά ελλιπή/γενική πληροφορία για το event, χωρίς να είναι η άμεση σελίδα του αγώνα/προγράμματος.
7. Homepage ομοσπονδίας/ομάδας: DIRECT μόνο αν το event είναι εμφανώς κύριο τρέχον νέο ή επίσημο αποτέλεσμα/πρόγραμμα στην κύρια περιοχή. Διαφορετικά PARTIAL ή NONE.
8. Αν access="UNAVAILABLE", ΜΗΝ μαντέψεις το περιεχόμενο της URL και βάλε support="NONE".
9. Μην επινοήσεις URLs. Κάθε url πρέπει να είναι ακριβώς από candidate_urls του αντίστοιχου event.
10. ΑΕΛ = Αθλητική Ένωση Λάρισας, όχι ΑΕΛ Λεμεσού.
11. Επέστρεψε κατά προτίμηση assessment για ΚΑΘΕ candidate URL. Αν κάποια URL δεν μπόρεσες να εξετάσεις, βάλε UNAVAILABLE αντί να την παραλείψεις.

EVENTS ΚΑΙ ΥΠΟΨΗΦΙΕΣ ΠΗΓΕΣ:
${JSON.stringify(plan.events)}

URLS ΠΟΥ ΠΡΕΠΕΙ ΝΑ ΑΝΟΙΞΕΙΣ:
${plan.urls.join('\n')}
`;

  const endpoint =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    model +
    ':generateContent';

  const payload = {
    contents: [
      {
        parts: [
          { text: prompt }
        ]
      }
    ],
    tools: [
      {
        url_context: {}
      }
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      responseJsonSchema: {
        type: 'object',
        properties: {
          verifications: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                event_key: { type: 'string' },
                source_assessments: {
                  type: 'array',
                  items: {
                    type: 'object',
                    properties: {
                      url: { type: 'string' },
                      access: {
                        type: 'string',
                        enum: ['OPENED', 'UNAVAILABLE']
                      },
                      support: {
                        type: 'string',
                        enum: ['DIRECT', 'PARTIAL', 'NONE']
                      },
                      reason: { type: 'string' }
                    },
                    required: [
                      'url',
                      'access',
                      'support',
                      'reason'
                    ],
                    additionalProperties: false
                  }
                },
                reason: { type: 'string' }
              },
              required: [
                'event_key',
                'source_assessments',
                'reason'
              ],
              additionalProperties: false
            }
          }
        },
        required: ['verifications'],
        additionalProperties: false
      },
      temperature: 0,
      maxOutputTokens: 8192
    }
  };

  const apiResult = MPP_callGeminiWithRetry_(endpoint, apiKey, payload);

  if (!apiResult.ok) return apiResult;

  try {

    const data = JSON.parse(apiResult.text);

    if (
      !data.candidates ||
      !data.candidates.length ||
      !data.candidates[0].content ||
      !data.candidates[0].content.parts
    ) {
      return {
        ok: false,
        retryable: true,
        status: 200,
        text: 'URL Context Gemini 200 αλλά χωρίς candidate/content.'
      };
    }

    const modelText = data.candidates[0].content.parts
      .filter(function(part) { return part.text; })
      .map(function(part) { return part.text; })
      .join('');

    const parsed = JSON.parse(modelText);

    return {
      ok: true,
      retryable: false,
      status: 200,
      text: apiResult.text,
      verifications: parsed.verifications || []
    };

  } catch (error) {

    return {
      ok: false,
      retryable: true,
      status: 200,
      text: 'URL Context structured-output parse error: ' + String(error)
    };
  }
}


function MPP_applyUrlVerification_(matches, verifications, plan, candidates, cfg) {

  const verificationByKey = {};
  (verifications || []).forEach(function(v) {
    verificationByKey[v.event_key] = v;
  });

  const allowedByEvent = {};
  (plan.events || []).forEach(function(event) {
    allowedByEvent[event.event_key] = new Set(event.candidate_urls || []);
  });

  const candidateByUrl = {};
  (candidates || []).forEach(function(candidate) {
    if (candidate && candidate.url && !candidateByUrl[candidate.url]) {
      candidateByUrl[candidate.url] = candidate;
    }
  });

  return (matches || []).map(function(original, index) {

    const match = Object.assign({}, original);
    const eventKey = MPP_eventKey_(match, index);
    const v = verificationByKey[eventKey] || {
      event_key: eventKey,
      source_assessments: [],
      reason: 'Δεν επιστράφηκε URL verification record.'
    };

    const allowed = allowedByEvent[eventKey] || new Set();
    const assessments = {};

    (v.source_assessments || []).forEach(function(a) {
      if (!a || !a.url || !allowed.has(a.url)) return;
      assessments[a.url] = a;
    });

    const usableSources = [];
    const fallbackNotes = [];

    Array.from(allowed).forEach(function(url) {

      const candidate = candidateByUrl[url];
      if (!candidate) return;

      const assessment = assessments[url];

      // --------------------------------------------------
      // Α. Η πραγματική σελίδα άνοιξε και είναι DIRECT.
      // --------------------------------------------------
      if (
        assessment &&
        assessment.access === 'OPENED' &&
        assessment.support === 'DIRECT'
      ) {

        usableSources.push({
          url: url,
          candidate: candidate,
          evidence_type: 'URL_CONTEXT_DIRECT',
          evidence_strength: 'STRONG',
          evidence_score: 100,
          evidence_reason: assessment.reason || 'URL Context DIRECT.'
        });

        return;
      }

      // --------------------------------------------------
      // Β. Η πραγματική σελίδα άνοιξε αλλά ήταν PARTIAL/NONE.
      // Δεν επιτρέπουμε Tavily fallback, γιατί ξέρουμε ήδη
      // ότι το κύριο σώμα της σελίδας δεν είναι άμεση πηγή.
      // --------------------------------------------------
      if (
        assessment &&
        assessment.access === 'OPENED'
      ) {

        fallbackNotes.push(
          MPP_getDomain_(url) +
          ': URL Context=' +
          assessment.support +
          ', χωρίς Tavily fallback.'
        );

        return;
      }

      // --------------------------------------------------
      // Γ. URL Context UNAVAILABLE ή δεν επέστρεψε assessment.
      // Τότε και ΜΟΝΟ τότε εξετάζουμε title/url/Tavily content.
      // --------------------------------------------------
      const fallback = MPP_tavilyDirectEvidence_(candidate, match);

      if (fallback.direct) {

        usableSources.push({
          url: url,
          candidate: candidate,
          evidence_type: 'TAVILY_DIRECT_FALLBACK',
          evidence_strength: fallback.strength,
          evidence_score: fallback.score,
          evidence_reason: fallback.reason
        });

        fallbackNotes.push(
          MPP_getDomain_(url) +
          ': Tavily fallback ' +
          fallback.strength +
          ' (' +
          fallback.reason +
          ')'
        );

      } else {

        fallbackNotes.push(
          MPP_getDomain_(url) +
          ': fallback rejected (' +
          fallback.reason +
          ')'
        );
      }
    });

    // --------------------------------------------------
    // Τελική κατάταξη πηγών.
    // --------------------------------------------------
    usableSources.forEach(function(source) {
      source.rank = MPP_hybridSourceRank_(source, match);
    });

    usableSources.sort(function(a, b) {
      return b.rank - a.rank;
    });

    const primaryObj = usableSources.length ? usableSources[0] : null;
    let secondaryObj = null;

    if (primaryObj) {
      secondaryObj = usableSources.find(function(source) {
        return (
          source.url !== primaryObj.url &&
          MPP_getDomain_(source.url) !== MPP_getDomain_(primaryObj.url)
        );
      }) || null;
    }

    const primary = primaryObj ? primaryObj.url : '';
    const secondary = secondaryObj ? secondaryObj.url : '';

    match.primary_source = primary;
    match.secondary_source = secondary;

    const primaryOfficial = Boolean(
      primaryObj &&
      (
        MPP_isOfficialSportsUrl_(primary) ||
        primaryObj.candidate.official === 'ΝΑΙ'
      )
    );
    const secondaryOfficial = Boolean(
      secondaryObj &&
      (
        MPP_isOfficialSportsUrl_(secondary) ||
        secondaryObj.candidate.official === 'ΝΑΙ'
      )
    );

    const primaryStrong = primaryObj && primaryObj.evidence_strength === 'STRONG';
    const secondaryStrong = secondaryObj && secondaryObj.evidence_strength === 'STRONG';

    const twoIndependentStrong = Boolean(
      primaryObj &&
      secondaryObj &&
      primaryStrong &&
      secondaryStrong &&
      MPP_getDomain_(primary) !== MPP_getDomain_(secondary)
    );

    const oneOfficialStrong = Boolean(
      (primaryOfficial && primaryStrong) ||
      (secondaryOfficial && secondaryStrong)
    );

    const primaryAggregator = primary ? MPP_isAggregatorUrl_(primary) : false;
    const secondaryAggregator = secondary ? MPP_isAggregatorUrl_(secondary) : false;
    const primarySignal = primary ? MPP_isSignalUrl_(primary) : false;
    const secondarySignal = secondary ? MPP_isSignalUrl_(secondary) : false;

    const primaryDirectNonSignal = Boolean(
      primaryObj && !primaryAggregator && !primarySignal
    );
    const secondaryDirectNonSignal = Boolean(
      secondaryObj && !secondaryAggregator && !secondarySignal
    );

    const directPlusAggregator = Boolean(
      primaryObj &&
      secondaryObj &&
      primaryStrong &&
      secondaryStrong &&
      MPP_getDomain_(primary) !== MPP_getDomain_(secondary) &&
      (
        (primaryDirectNonSignal && secondaryAggregator) ||
        (secondaryDirectNonSignal && primaryAggregator)
      )
    );

    const twoDirectStrong = Boolean(
      twoIndependentStrong &&
      primaryDirectNonSignal &&
      secondaryDirectNonSignal
    );

    // --------------------------------------------------
    // CONFIRMED rules V5 / multi-wave pilot final
    // --------------------------------------------------
    if (oneOfficialStrong) {

      match.status = 'CONFIRMED';
      match.confidence = 'HIGH';

    } else if (twoDirectStrong || directPlusAggregator) {

      // Δύο ανεξάρτητα strong signals, με τουλάχιστον μία πραγματική
      // direct μη-aggregator πηγή, μπορούν να επιβεβαιώσουν MEDIUM.
      match.status = 'CONFIRMED';
      match.confidence = 'MEDIUM';

    } else if (primaryObj) {

      // Aggregator-only ή betting-signal-only ποτέ δεν γίνεται CONFIRMED.
      match.status = 'UNCERTAIN';
      match.confidence = 'MEDIUM';

    } else {

      match.status = 'UNCERTAIN';
      match.confidence = 'LOW';
    }

    const chosenText = [primaryObj, secondaryObj]
      .filter(Boolean)
      .map(function(source) {
        return (
          MPP_getDomain_(source.url) +
          '=' +
          source.evidence_type +
          '/' +
          source.evidence_strength
        );
      })
      .join(', ');

    match.reason =
      String(match.reason || '') +
      ' | Hybrid source verification: ' +
      String(v.reason || '') +
      (chosenText ? ' | Επιλεγμένες: ' + chosenText : '') +
      (fallbackNotes.length ? ' | ' + fallbackNotes.join(' ; ') : '');

    return match;
  });
}


function MPP_tavilyDirectEvidence_(candidate, match) {

  const title = MPP_normText_(candidate.title || '');
  const url = MPP_normText_(candidate.url || '');
  const content = MPP_normText_(candidate.content || '');
  const contentStart = content.substring(0, 700);
  const titleUrl = title + ' ' + url;

  const home = MPP_normText_(match.home_team || '');
  const away = MPP_normText_(match.away_team || '');
  const opponent = MPP_pickOpponent_(home, away);
  const opponentTokens = MPP_flexibleTokens_(MPP_keywords_(opponent));
  const competitionTokens = MPP_keywords_(match.competition || '');

  const titleOpponent = MPP_containsAny_(titleUrl, opponentTokens);
  const startOpponent = MPP_containsAny_(contentStart, opponentTokens);
  const fullOpponent = MPP_containsAny_(content, opponentTokens);

  const aelTitle = MPP_mentionsAEL_(titleUrl);
  const aelStart = MPP_mentionsAEL_(contentStart);
  const aelFull = MPP_mentionsAEL_(content);

  const competitionHit =
    MPP_containsAny_(titleUrl, competitionTokens) ||
    MPP_containsAny_(contentStart, competitionTokens);

  const dateHit = MPP_textHasEventDate_(titleUrl + ' ' + contentStart, match.date);
  const timeHit = match.time ? (titleUrl + ' ' + contentStart).includes(MPP_normText_(match.time)) : false;
  const venueHit = match.venue ? (titleUrl + ' ' + contentStart).includes(MPP_normText_(match.venue)) : false;

  const opponentIndex = MPP_firstTokenIndex_(content, opponentTokens);

  const obviousNoise = [
    'σεμιναριο',
    'sport radar',
    'παικτης',
    'αποκτηση',
    'μεταγραφ',
    'μετοχικ',
    'διοικητικ',
    'ημεριδα',
    'συνεργασια'
  ];

  if (
    MPP_containsAny_(title, obviousNoise) &&
    !titleOpponent
  ) {
    return {
      direct: false,
      strength: 'NONE',
      score: 0,
      reason: 'Ο τίτλος αφορά άλλο θέμα και ο αγώνας δεν είναι στον τίτλο/URL.'
    };
  }

  if (!fullOpponent) {
    return {
      direct: false,
      strength: 'NONE',
      score: 0,
      reason: 'Δεν εντοπίζεται ο αντίπαλος στο Tavily περιεχόμενο.'
    };
  }

  if (!aelFull && !aelTitle) {
    return {
      direct: false,
      strength: 'NONE',
      score: 0,
      reason: 'Δεν τεκμηριώνεται ότι το περιεχόμενο αφορά την ΑΕΛ Λάρισας.'
    };
  }

  // Αν ο αντίπαλος εμφανίζεται μόνο πολύ αργά στο body και όχι σε τίτλο/URL,
  // είναι τυπικό σημάδι sidebar / latest-news contamination.
  if (!titleOpponent && opponentIndex >= 700) {
    return {
      direct: false,
      strength: 'NONE',
      score: 0,
      reason: 'Η αναφορά του αντιπάλου εμφανίζεται πολύ αργά στο κείμενο (πιθανό sidebar/latest-news).'
    };
  }

  let score = 0;

  if (titleOpponent) score += 12;
  if (startOpponent) score += 8;
  else if (opponentIndex >= 0 && opponentIndex < 700) score += 4;

  if (aelTitle) score += 5;
  if (aelStart) score += 4;
  else if (aelFull) score += 1;

  if (competitionHit) score += 4;
  if (dateHit) score += 3;
  if (timeHit) score += 2;
  if (venueHit) score += 2;

  if (MPP_isOfficialSportsUrl_(candidate.url)) score += 4;
  if (MPP_isTrustedSportsMediaUrl_(candidate.url)) score += 2;

  // Event-specific article: αντίπαλος στον τίτλο/URL και ΑΕΛ στο title/body.
  const eventSpecificArticle =
    titleOpponent &&
    (aelTitle || aelStart) &&
    (startOpponent || competitionHit || dateHit || timeHit);

  // Επίσημη generic/homepage σελίδα επιτρέπεται μόνο όταν το event
  // εμφανίζεται πολύ νωρίς στο body και υπάρχουν επιπλέον match στοιχεία.
  const officialFrontPageEvidence =
    MPP_isOfficialSportsUrl_(candidate.url) &&
    opponentIndex >= 0 &&
    opponentIndex < 350 &&
    aelStart &&
    (competitionHit || dateHit || timeHit || venueHit);


  if (MPP_isSignalUrl_(candidate.url)) {
    return {
      direct: false,
      strength: 'NONE',
      score: score,
      reason: 'W6 betting/odds signal: χρησιμοποιείται μόνο για discovery, όχι ως direct τελική πηγή.'
    };
  }

  if (eventSpecificArticle && score >= 16) {
    return {
      direct: true,
      strength: 'STRONG',
      score: score,
      reason: 'Ο τίτλος/URL και η αρχή του Tavily content αφορούν άμεσα τον ίδιο αγώνα.'
    };
  }

  if (officialFrontPageEvidence && score >= 14) {
    return {
      direct: true,
      strength: 'STRONG',
      score: score,
      reason: 'Επίσημη σελίδα με το event στην αρχή του κύριου Tavily content και συγκεκριμένα στοιχεία αγώνα.'
    };
  }

  if (
    startOpponent &&
    aelStart &&
    score >= 12
  ) {
    return {
      direct: true,
      strength: 'MEDIUM',
      score: score,
      reason: 'Η αρχή του Tavily content αφορά τον αγώνα, αλλά η σελίδα δεν είναι αρκετά event-specific για STRONG.'
    };
  }

  return {
    direct: false,
    strength: 'NONE',
    score: score,
    reason: 'Το Tavily evidence δεν είναι αρκετά άμεσο για ασφαλές fallback.'
  };
}


function MPP_hybridSourceRank_(source, match) {

  let rank = 0;

  if (source.evidence_type === 'URL_CONTEXT_DIRECT') rank += 50;
  if (source.evidence_type === 'TAVILY_DIRECT_FALLBACK') rank += 30;

  if (source.evidence_strength === 'STRONG') rank += 25;
  if (source.evidence_strength === 'MEDIUM') rank += 10;

  if (MPP_isOfficialSportsUrl_(source.url)) rank += 60;
  else if (MPP_isTrustedSportsMediaUrl_(source.url)) rank += 25;
  else if (MPP_isAggregatorUrl_(source.url)) rank += 12;
  else if (MPP_isSignalUrl_(source.url)) rank -= 20;
  else rank += 10;

  rank += Math.min(
    20,
    Math.max(
      0,
      MPP_sourceRelevanceScore_(source.candidate, match)
    )
  );

  rank += Math.min(15, Number(source.evidence_score || 0) / 2);

  return rank;
}


function MPP_isTrustedSportsMediaUrl_(url) {

  const domain = MPP_getDomain_(url);

  const trusted = new Set([
    'sport24.gr',
    'sport-fm.gr',
    'gazzetta.gr',
    'eleftheria.gr',
    'onlarissa.gr',
    'larissanet.gr',
    'pressing.gr',
    'arenalarissa.gr',
    'athleticlarissa.gr',
    'aelole.gr',
    'aeliko-kafeneio.gr',
    'gegonota.news',
    'sportrikala.gr',
    'thessaliatv.gr'
  ]);

  return trusted.has(domain);
}


function MPP_mentionsAEL_(text) {

  const value = MPP_normText_(text || '');

  return (
    value.includes('αελ') ||
    value.includes('ael') ||
    value.includes('λαρισ')
  );
}


function MPP_firstTokenIndex_(text, tokens) {

  if (!text || !tokens || tokens.length === 0) return -1;

  let best = -1;

  tokens.forEach(function(token) {
    const idx = text.indexOf(token);
    if (idx < 0) return;
    if (best < 0 || idx < best) best = idx;
  });

  return best;
}


function MPP_textHasEventDate_(text, dateText) {

  const parts = String(dateText || '').split('/');
  if (parts.length !== 3) return false;

  const dd = parts[0];
  const mm = parts[1];
  const yyyy = parts[2];

  const d = String(Number(dd));
  const m = String(Number(mm));

  const variants = [
    dd + '/' + mm + '/' + yyyy,
    d + '/' + m + '/' + yyyy,
    dd + '-' + mm + '-' + yyyy,
    d + '-' + m + '-' + yyyy,
    dd + '/' + mm,
    d + '/' + m
  ];

  const normalized = MPP_normText_(text || '');

  return variants.some(function(v) {
    return normalized.includes(MPP_normText_(v));
  });
}


function MPP_eventKey_(match, index) {
  return [
    String(match.department_id || ''),
    String(match.home_team || ''),
    String(match.away_team || ''),
    String(match.date || ''),
    String(match.time || ''),
    String(index || 0)
  ].join('|');
}


function MPP_directPageBonus_(candidate, match) {

  const title = MPP_normText_(candidate.title || '');
  const url = MPP_normText_(candidate.url || '');
  const opponent = MPP_pickOpponent_(
    MPP_normText_(match.home_team || ''),
    MPP_normText_(match.away_team || '')
  );
  const opponentTokens = MPP_flexibleTokens_(MPP_keywords_(opponent));

  let bonus = 0;

  if (MPP_containsAny_(title + ' ' + url, opponentTokens)) {
    bonus += 12;
  }

  const obviousNoise = [
    'σεμιναριο',
    'sport radar',
    'παικτης',
    'αποκτηση',
    'μεταγραφ',
    'μετοχικ',
    'διοικητικ',
    'ημεριδα'
  ];

  if (
    MPP_containsAny_(title, obviousNoise) &&
    !MPP_containsAny_(title + ' ' + url, opponentTokens)
  ) {
    bonus -= 20;
  }

  return bonus;
}


function MPP_isOfficialSportsUrl_(url) {

  const domain = MPP_getDomain_(url);

  const officialDomains = new Set([
    'aelfc.gr',
    'epo.gr',
    'sl2.gr',
    'slgr.gr',
    'e-ael.gr',
    'aelbasket.gr',
    'eskath.gr',
    'volleyball.gr',
    'espekel.com',
    'epssalas.gr',
    'hellenicboxing.org.gr'
  ]);

  return officialDomains.has(domain);
}


function MPP_isAggregatorUrl_(url) {

  const domain = MPP_getDomain_(url);

  const aggregators = new Set([
    'flashscore.com',
    'flashscore.gr',
    'sofascore.com'
  ]);

  // Tavily/search engines μπορεί να επιστρέψουν locale subdomains.
  if (domain.indexOf('flashscore.') === 0) return true;

  return Array.from(aggregators).some(function(base) {
    return domain === base || domain.endsWith('.' + base);
  });
}


function MPP_isSignalUrl_(url) {

  const domain = MPP_getDomain_(url);

  const signals = new Set([
    'oddsportal.com',
    'betexplorer.com',
    'stoiximan.gr',
    'novibet.gr',
    'bet365.com'
  ]);

  return Array.from(signals).some(function(base) {
    return domain === base || domain.endsWith('.' + base);
  });
}

// ======================================================
// FINAL_TEST
// ======================================================

function MPP_writeFinal_(sheet, matches) {

  if (sheet.getMaxColumns() < 14) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      14 - sheet.getMaxColumns()
    );
  }

  // Γράφουμε μόνο μετά από επιτυχημένο Gemini + JS validation.
  sheet.clearContents();

  sheet.getRange(1, 1, 1, 14).setValues([[
    'ID',
    'ΤΜΗΜΑ',
    'ΑΘΛΗΜΑ',
    'ΓΗΠΕΔΟΥΧΟΣ',
    'ΦΙΛΟΞΕΝΟΥΜΕΝΟΣ',
    'ΗΜΕΡΟΜΗΝΙΑ',
    'ΩΡΑ',
    'ΕΔΡΑ',
    'ΔΙΟΡΓΑΝΩΣΗ',
    'STATUS',
    'CONFIDENCE',
    'ΚΥΡΙΑ ΠΗΓΗ',
    '2η ΠΗΓΗ',
    'ΑΙΤΙΟΛΟΓΗΣΗ'
  ]]);

  if (matches.length === 0) {

    sheet.getRange('A2').setValue('ΚΑΝΕΝΑ ΕΠΙΒΕΒΑΙΩΜΕΝΟ ΓΕΓΟΝΟΣ');

  } else {

    const rows = matches.map(function(match) {
      return [
        match.department_id,
        match.department,
        match.sport,
        match.home_team,
        match.away_team,
        match.date,
        match.time,
        match.venue,
        match.competition,
        match.status,
        match.confidence,
        match.primary_source,
        match.secondary_source,
        match.reason
      ];
    });

    sheet.getRange(2, 1, rows.length, 14).setValues(rows);
  }

  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 130);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 220);
  sheet.setColumnWidth(9, 220);
  sheet.setColumnWidth(10, 120);
  sheet.setColumnWidth(11, 120);
  sheet.setColumnWidth(12, 500);
  sheet.setColumnWidth(13, 500);
  sheet.setColumnWidth(14, 700);

  if (sheet.getLastRow() > 0) {
    sheet.getRange(1, 1, sheet.getLastRow(), 14).setWrap(true);
  }
}


// ======================================================
// POST_TEST
// ======================================================

function MPP_writePost_(sheet, matches, cfg, week) {

  const confirmed = matches.filter(function(match) {
    return (
      match.status === 'CONFIRMED' &&
      (match.confidence === 'HIGH' || match.confidence === 'MEDIUM')
    );
  });

  // Audit φύλλο μόνο. Το production POSTS δεν καθαρίζεται ΠΟΤΕ από εδώ.
  sheet.getRange('A2:D100').clearContent();

  sheet.getRange('A2').setValue(week.startFull);
  sheet.getRange('D2').setValue(new Date());

  if (confirmed.length === 0) {

    sheet.getRange('C2').setValue('NO_CONFIRMED_EVENTS');

    return {
      ready: false,
      status: 'NO_CONFIRMED_EVENTS',
      date: week.startFull,
      post: ''
    };
  }

  const blocks = confirmed.map(function(match) {

    const lines = [];

    lines.push(
      MPP_sportEmoji_(match.sport) +
      ' ' +
      MPP_shortDepartment_(match.department_id, match.department)
    );

    lines.push(
      match.home_team +
      ' – ' +
      match.away_team
    );

    lines.push(
      '🗓️ ' +
      MPP_greekDay_(match.date) +
      ' ' +
      MPP_shortDateText_(match.date)
    );

    if (match.time) {
      lines.push('🕒 ' + match.time);
    }

    if (match.venue) {
      lines.push('📍 ' + match.venue);
    }

    if (match.competition) {
      lines.push('🏆 ' + match.competition);
    }

    return lines.join('\n');
  });

  const post =
    '📅 Βυσσινί Ατζέντα | ' +
    week.startShort +
    ' – ' +
    week.endShort +
    '\n\n' +
    blocks.join('\n\n') +
    '\n\n' +
    '#ΑΕΛ #ΒυσσινίΑτζέντα #Λάρισα #AEL';

  sheet.getRange('B2').setValue(post);
  sheet.getRange('C2').setValue('TEST_READY');

  sheet.getRange('B2').setWrap(true);

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 650);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);

  return {
    ready: true,
    status: 'READY',
    date: week.startFull,
    post: post
  };
}


// ======================================================
// PRODUCTION COMMIT
// Γράφει ΜΟΝΟ POSTS!A2:B2.
// HISTORY: μόνο ανάγνωση για idempotency.
// ======================================================

function MPP_commitPostToProduction_(postsSheet, historySheet, postResult, props, week) {

  if (!postResult || !postResult.ready || !postResult.post) {
    return {
      status: 'NO_CONFIRMED_EVENTS',
      retryable: false
    };
  }

  // 1. Script-level idempotency.
  if (props.getProperty('MPP_COMMITTED_WEEK') === week.key) {
    return {
      status: 'ALREADY_COMMITTED',
      retryable: false
    };
  }

  // 2. Αν το Make έχει ήδη δημοσιεύσει αυτή την εβδομάδα,
  // δεν ξαναγεμίζουμε το POSTS ακόμη κι αν έχει ήδη αδειάσει.
  if (MPP_historyHasWeeklyPost_(historySheet, week.startFull)) {
    props.setProperty('MPP_COMMITTED_WEEK', week.key);
    return {
      status: 'ALREADY_PUBLISHED_HISTORY',
      retryable: false
    };
  }

  // 3. Δεν αντικαθιστούμε ποτέ υπάρχον queue item.
  const queue = postsSheet.getRange('A2:B2').getDisplayValues()[0];
  const existingDate = String(queue[0] || '').trim();
  const existingText = String(queue[1] || '').trim();

  if (existingDate || existingText) {

    // Αν είναι ακριβώς το ίδιο post, θεωρούμε ότι το commit έχει ήδη γίνει.
    if (
      existingText &&
      existingText === String(postResult.post).trim()
    ) {
      props.setProperty('MPP_COMMITTED_WEEK', week.key);
      return {
        status: 'ALREADY_QUEUED',
        retryable: false
      };
    }

    // Άλλο περιεχόμενο στο POSTS: περιμένουμε να το καταναλώσει το Make.
    return {
      status: 'POSTS_OCCUPIED_RETRY',
      retryable: true
    };
  }

  // 4. Τελικό atomic write: μόνο A2:B2.
  postsSheet
    .getRange('A2:B2')
    .setValues([[
      postResult.date,
      postResult.post
    ]]);

  SpreadsheetApp.flush();

  // 5. Read-after-write verification.
  const written = postsSheet.getRange('A2:B2').getDisplayValues()[0];

  if (
    String(written[0] || '').trim() !== String(postResult.date || '').trim() ||
    String(written[1] || '').trim() !== String(postResult.post || '').trim()
  ) {
    return {
      status: 'POST_WRITE_VERIFY_RETRY',
      retryable: true
    };
  }

  props.setProperty('MPP_COMMITTED_WEEK', week.key);
  props.setProperty('MPP_COMMITTED_AT', new Date().toISOString());

  return {
    status: 'WRITTEN_TO_POSTS',
    retryable: false
  };
}


function MPP_historyHasWeeklyPost_(historySheet, mondayDateText) {

  const lastRow = historySheet.getLastRow();

  if (lastRow < 2) return false;

  const rows = historySheet
    .getRange(2, 1, lastRow - 1, 3)
    .getDisplayValues();

  const targetDate = String(mondayDateText || '').trim();

  return rows.some(function(row) {

    const date = String(row[0] || '').trim();
    const type = MPP_normText_(row[1] || '');
    const text = String(row[2] || '').trim();

    const weeklyType =
      type.includes('εβδομαδι') ||
      type.includes('weekly');

    return (
      date === targetDate &&
      weeklyType &&
      text.length > 0
    );
  });
}


function MPP_clearPendingPost_() {

  const props = PropertiesService.getScriptProperties();

  props.deleteProperty('MPP_PENDING_POST_WEEK');
  props.deleteProperty('MPP_PENDING_POST_DATE');
  props.deleteProperty('MPP_PENDING_POST_TEXT');
}


function MPP_isMondayNow_(timezone) {

  const local = Utilities.formatDate(
    new Date(),
    timezone,
    'yyyy-MM-dd'
  );

  const p = local.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12, 0, 0));

  return d.getUTCDay() === 1;
}


function MPP_sportEmoji_(sport) {

  const value = String(sport || '').toLowerCase();

  if (value.includes('ποδόσφαιρο') || value.includes('futsal')) return '⚽';
  if (value.includes('μπάσκετ')) return '🏀';
  if (value.includes('βόλεϊ') || value.includes('βολλ')) return '🏐';
  if (value.includes('στίβ')) return '🏃';
  if (value.includes('πυγμαχ')) return '🥊';

  return '🔴';
}


function MPP_shortDepartment_(id, department) {

  const labels = {
    FOOTBALL_MEN: 'Ποδόσφαιρο | ΠΑΕ ΑΕΛ',
    FOOTBALL_K19: 'Ποδόσφαιρο | ΑΕΛ Κ19',
    FOOTBALL_K17: 'Ποδόσφαιρο | ΑΕΛ Κ17',
    FOOTBALL_K15: 'Ποδόσφαιρο | ΑΕΛ Κ15',
    FOOTBALL_WOMEN: 'Ποδόσφαιρο | Γυναικών',
    FOOTBALL_WOMEN_K17: 'Ποδόσφαιρο | Κ17 Γυναικών',
    FUTSAL: 'Futsal | ΑΕΛ',
    BASKET_MEN: 'Μπάσκετ | Ανδρών',
    BASKET_YOUTH: 'Μπάσκετ | Ακαδημίες',
    VOLLEY_WOMEN: 'Βόλεϊ | Γυναικών',
    VOLLEY_MEN: 'Βόλεϊ | Ανδρών',
    VOLLEY_YOUTH: 'Βόλεϊ | Αναπτυξιακό',
    BOXING: 'Πυγμαχία | ΑΕΛ'
  };

  return labels[id] || department;
}


function MPP_greekDay_(dateText) {

  const parts = String(dateText || '').split('/');

  if (parts.length !== 3) return '';

  const date = new Date(
    Date.UTC(
      Number(parts[2]),
      Number(parts[1]) - 1,
      Number(parts[0]),
      12,
      0,
      0
    )
  );

  const days = [
    'Κυριακή',
    'Δευτέρα',
    'Τρίτη',
    'Τετάρτη',
    'Πέμπτη',
    'Παρασκευή',
    'Σάββατο'
  ];

  return days[date.getUTCDay()];
}


function MPP_shortDateText_(dateText) {

  const parts = String(dateText || '').split('/');

  if (parts.length < 2) return dateText;

  return parts[0] + '/' + parts[1];
}


// ======================================================
// RETRY TRIGGER
// ======================================================

function MPP_scheduleRetry_() {

  MPP_clearRetry_();

  ScriptApp
    .newTrigger('mondayPipelineProductionRetry')
    .timeBased()
    .after(5 * 60 * 1000)
    .create();

  PropertiesService
    .getScriptProperties()
    .setProperty('MPP_RETRY_SCHEDULED_AT', new Date().toISOString());

  console.log('Προγραμματίστηκε νέο Monday PRODUCTION retry σε τουλάχιστον 5 λεπτά.');
}


function MPP_clearRetry_() {

  ScriptApp
    .getProjectTriggers()
    .forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'mondayPipelineProductionRetry') {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  PropertiesService
    .getScriptProperties()
    .deleteProperty('MPP_RETRY_SCHEDULED_AT');
}


// ======================================================
// SMALL HELPERS
// ======================================================

function MPP_extractDomains_(text) {

  if (!text) return [];

  const domains = [];

  const urlMatches = String(text).match(/https?:\/\/[^\s,;]+/g) || [];

  urlMatches.forEach(function(url) {

    const domain = MPP_getDomain_(url);

    if (domain && !domains.includes(domain)) {
      domains.push(domain);
    }
  });

  return domains;
}


function MPP_getDomain_(url) {

  try {
    return String(url)
      .replace(/^https?:\/\//i, '')
      .replace(/^www\./i, '')
      .split('/')[0]
      .split('?')[0]
      .split('#')[0]
      .toLowerCase();
  } catch (error) {
    return '';
  }
}
