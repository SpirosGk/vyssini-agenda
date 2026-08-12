# Configuration

## Required Script Properties

Configure these in:

**Apps Script → Project Settings → Script Properties**

```text
PRODUCTION_SPREADSHEET_ID=<your production spreadsheet ID>
TAVILY_API_KEY=<your Tavily API key>
GEMINI_API_KEY=<your Gemini API key>
```

Optional:

```text
AUDIT_SPREADSHEET_ID=<your audit spreadsheet ID>
GEMINI_MODEL=gemini-3.6-flash
```

## Security rule

Never copy the real values into:

- `.gs` source files
- `README.md`
- screenshots
- GitHub Issues
- GitHub Discussions
- commit messages

If a secret is accidentally committed, revoke/rotate it immediately and remove it from the repository history.
