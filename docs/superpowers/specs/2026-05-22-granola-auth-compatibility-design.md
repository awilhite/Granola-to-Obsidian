# Granola Auth Compatibility Design

## Summary

Adapt the plugin to Granola's newer WorkOS-style auth/session model without assuming the plugin must fully own auth from day one.

The plugin should stop treating `stored-accounts.json` and `supabase.json` as static bearer-token files. Instead, it should introduce a Granola auth resolver that can:

1. Discover known Granola auth/session sources.
2. Normalize them into a common internal session shape.
3. Prefer the freshest usable session source.
4. Refresh access tokens directly when that path is verified.
5. Hand a valid access token to the sync/API layer.

This is a hybrid approach:
- compatible with Granola's current app-managed auth model
- capable of direct plugin-managed refresh where verified
- structured so hardening and clearer failure messaging can be added afterward

## Context

Recent sync failures showed that the plugin's current auth assumptions are no longer sufficient:

- Sync on May 21, 2026 failed before fetching documents.
- The plugin loaded an expired access token from `stored-accounts.json`.
- Re-running sync with Granola open did not rewrite either `stored-accounts.json` or `supabase.json`.
- Granola's current app bundle includes an Electron-managed session flow:
  - `get-session`
  - `session-changed`
  - `getRefreshedTokens()`
  - `refreshWorkOsAccessToken()`
- Granola's app now appears to manage WorkOS-style tokens in memory and refresh them proactively.

Current plugin behavior:
- `loadCredentials()` reads files directly and returns the first access token it finds.
- `fetchGranolaDocuments()` still calls `https://api.granola.ai/v2/get-documents`.
- The sync pipeline assumes that file-based token discovery is enough to reach Granola APIs.

The compatibility gap is not note rendering or duplicate detection. It is the plugin's outdated auth/session model.

## Goals

- Restore working sync against Granola's current auth model.
- Decouple sync logic from static file-based token assumptions.
- Support the current WorkOS-style token shape and current Granola auth lifecycle.
- Keep the solution compatible with the current plugin features:
  - template management
  - transcript fetch
  - authored notes
  - whole-vault duplicate detection

## Non-Goals

- Full auth hardening in this pass.
- User-facing troubleshooting polish in this pass.
- Re-architecting note rendering or sync behavior unrelated to auth compatibility.
- Making the plugin fully independent from the Granola desktop app in all scenarios.

## Architecture

### 1. GranolaAuthResolver

Introduce a dedicated auth/session resolver that owns Granola session discovery and token refresh.

Responsibilities:
- discover candidate auth/session sources
- parse newer WorkOS token payloads
- normalize session data into one internal shape
- return a valid access token to API callers
- perform direct refresh when the current refresh contract is verified

Normalized auth shape:

```js
{
  accessToken: string,
  refreshToken: string | null,
  sessionId: string | null,
  signInMethod: string | null,
  obtainedAt: number | null,
  expiresInSeconds: number | null,
  clientVersion: string,
  platform: string,
  osVersion: string,
  source: string,
  workspaceId: string | null,
  deviceId: string | null
}
```

Known source families:
- `stored-accounts.json`
- `supabase.json`
- any fresher persisted WorkOS/session representation found during implementation

Behavior:
- discover candidate sessions
- normalize all candidates
- choose the best candidate
- if the chosen candidate requires refresh and refresh is available, refresh it
- return the resolved auth context

### 2. GranolaApiClient

Move API interaction behind a single client that always receives its auth from `GranolaAuthResolver`.

Responsibilities:
- fetch documents
- fetch folders
- fetch transcripts
- perform template-management calls
- own endpoint selection and request headers

This removes direct token/file coupling from sync code.

### 3. Sync Orchestrator

Keep sync orchestration mostly unchanged:
- build diagnostics
- resolve auth once at sync start
- fetch documents through the API client
- continue with readiness checks, duplicate detection, rendering, and writes

The sync layer should not care whether the access token came from a file, a refreshed session, or a future session source.

## Compatibility Strategy

### Step A. Replace direct file-auth assumptions

Refactor `loadCredentials()` into an auth resolver entry point.

The current method:
- prefers `stored-accounts.json`
- falls back to `supabase.json`
- returns raw token data without lifecycle awareness

The new resolver should:
- parse all known candidates first
- support WorkOS token fields such as:
  - `access_token`
  - `refresh_token`
  - `expires_in`
  - `obtained_at`
  - `session_id`
  - `sign_in_method`
- preserve client/platform/device context for API headers

### Step B. Align with the newer Granola API model

Granola's own bundle now calls `get-documents-v2` through its API abstraction, while the plugin still calls `https://api.granola.ai/v2/get-documents`.

Implementation should:
- isolate document-fetch logic behind `GranolaApiClient`
- verify whether the plugin can continue using the current public endpoint with refreshed auth
- if needed, switch to the newer API path/shape used by Granola's app

This change should be limited to the fetch layer and not leak into note rendering logic.

### Step C. Add direct refresh support where verified

Once the refresh contract is confirmed, the resolver should refresh directly instead of waiting for Granola to rewrite files.

Expected behavior:
- when a resolved candidate includes a usable `refresh_token`, the resolver may refresh it
- the refreshed session should be normalized back into the same auth shape
- if Granola stores refreshed state back to disk in the app, that is a bonus, not a requirement

The plugin should not depend on Granola being open to rewrite on-disk tokens in order to sync successfully.

### Step D. Keep current features riding on the new auth layer

The following features should continue using the resolved auth context:
- document sync
- template management
- transcript fetch
- authored-notes hydration
- folder fetch

## Data Flow

1. User triggers manual or auto sync.
2. Sync orchestrator requests a valid session from `GranolaAuthResolver`.
3. Resolver:
   - discovers candidates
   - normalizes them
   - selects the best one
   - refreshes if needed and supported
4. Sync orchestrator passes resolved auth to `GranolaApiClient`.
5. API client fetches documents using the current compatible endpoint shape.
6. Sync continues normally for ready documents.

## Error Handling

This pass should keep behavior simple and explicit.

Failure categories:
- no recognizable Granola session source
- failed refresh attempt
- API rejection even after auth resolution
- unexpected response shape from newer document endpoints

For this pass:
- fail sync cleanly
- record the failure in diagnostics
- surface a meaningful internal error string

Detailed user-facing remediation and auth hardening can follow in a later pass.

## Testing Strategy

### Manual validation

1. Trigger sync with Granola open and a valid current session.
2. Confirm document fetch succeeds again.
3. Confirm at least one newly created Granola note syncs into Obsidian.
4. Confirm template management still applies the configured `Default` template when enabled.
5. Confirm transcript and authored-note features still work when re-enabled.

### Targeted validation

- Resolver can parse `stored-accounts.json` WorkOS token shape.
- Resolver can parse `supabase.json` WorkOS token shape.
- Resolver can normalize refresh-capable sessions.
- API client can fetch documents with the resolved auth context.
- Sync diagnostics reflect auth/fetch failures clearly.

## Risks

- Granola's refresh flow may depend on headers or state not yet fully mapped.
- The newer document endpoint may return a shape different from the plugin's current expectations.
- Template-management private endpoints may have additional auth assumptions compared to document fetch.
- The plugin may need to preserve device/workspace headers more faithfully than it currently does.

## Recommended Implementation Order

1. Add `GranolaAuthResolver` and normalized auth model.
2. Refactor sync to depend on the resolver instead of raw file reads.
3. Introduce `GranolaApiClient` for document fetch and migrate fetch logic into it.
4. Validate the compatible document-fetch path.
5. Add direct refresh support using the verified WorkOS refresh contract.
6. Validate current private feature flows on top of the new auth layer.

## Decision

Proceed with the hybrid compatibility path:
- adapt to Granola's current auth/session architecture now
- use direct refresh where verified
- delay hardening and polish until compatibility is restored
