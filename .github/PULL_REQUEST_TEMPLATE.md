<!--
Title format (checked by CI): {emoji} {Capitalised verb phrase}
  Good:  ✨ Add repeat-record dataframe support
  Bad:   :sparkles: Added repeat records.
Use the raw emoji in the title; use :gitmoji: codes in commit messages.

Add exactly ONE label: feature | bug | enhancement | documentation | internal
-->

## Summary

<!-- What changed, in a sentence or two. -->

## Motivation / context

<!-- Why this change? Link any related issue/discussion. -->

## Testing

<!-- How you verified it. Include commands. -->

```bash
PYTHONPATH=src python -m pytest -q
```

## Risks / migrations

<!-- Breaking changes, schema changes, or "none". -->

## Checklist

- [ ] One label applied (`feature` / `bug` / `enhancement` / `documentation` / `internal`)
- [ ] Tests pass locally (`PYTHONPATH=src python -m pytest -q`)
- [ ] Docs updated, or noted here why not
- [ ] `CLAUDE.md` updated if behavior/API/schema/placeholder changed
