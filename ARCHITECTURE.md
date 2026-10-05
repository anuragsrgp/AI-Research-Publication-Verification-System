# Architecture

## Data flow

1. `FORM_REGISTRY` stores registered Forms.
2. The system reads responses from each registered Form.
3. Form fields are normalized and mapped dynamically.
4. All submitted records are consolidated into `MASTER_DATA`.
5. Verification reads publication metadata from Master Data.
6. DOI is checked first when available.
7. Crossref title search is used when DOI is unavailable or cannot be resolved.
8. Verification results are written to `VERIFICATION`.
9. Groq AI is an optional supporting layer for ambiguous metadata comparisons.
10. Operational events are written to `SYSTEM_LOG`.

## Separation of concerns

- Form registration/synchronization: Form registry and sync functions
- Master data: submitted data only
- Verification: external verification output only
- AI: supporting analysis only
- Logging: diagnostics only

This separation prevents API metadata from overwriting user-submitted Master Data.
