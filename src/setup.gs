/**
 * Validates the non-secret configuration required by the project.
 *
 * Secrets are intentionally stored in Apps Script Script Properties
 * and are never committed to source control.
 */
function validateProjectConfiguration() {
  const props = PropertiesService.getScriptProperties();

  const required = [
    'PRODUCTION_SPREADSHEET_ID',
    'TAVILY_API_KEY',
    'GEMINI_API_KEY'
  ];

  const optional = [
    'AUDIT_SPREADSHEET_ID',
    'GEMINI_MODEL'
  ];

  const missing = required.filter(function (name) {
    return !props.getProperty(name);
  });

  console.log('Required configuration present: ' +
    required.filter(function (name) {
      return !!props.getProperty(name);
    }).join(', '));

  console.log('Optional configuration present: ' +
    optional.filter(function (name) {
      return !!props.getProperty(name);
    }).join(', '));

  if (missing.length > 0) {
    throw new Error(
      'Missing Script Properties: ' + missing.join(', ')
    );
  }

  console.log('Configuration check passed.');
}
