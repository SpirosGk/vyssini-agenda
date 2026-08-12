# Architecture

## High-level flow

The production pipeline is intentionally divided into discovery, validation, verification and commit stages.

```text
Monday trigger
     |
     v
Production guard
     |
     v
Tavily discovery
     |
     v
Candidate normalization
     |
     v
Gemini validation
     |
     v
Deterministic JavaScript validation
     |
     v
URL/context verification
     |
     v
Final verified matches
     |
     v
Post generation
     |
     v
Production POSTS
     |
     v
HISTORY / duplicate protection
     |
     v
Make publishing automation
```

## State

Google Apps Script `Script Properties` are used for runtime state such as:

- current pipeline status
- current reporting week
- Tavily cache timestamp
- pending retry
- pending post
- last error
- last successful execution

This allows a retry to continue from an appropriate stage instead of blindly restarting every external API operation.

## Failure handling

Retryable API responses and temporary sheet conflicts result in a delayed retry trigger.

The pipeline also uses:

- execution locks
- cache checks
- time-budget checks
- pending-post recovery
- duplicate-history checks

The goal is graceful degradation rather than repeated full executions.
