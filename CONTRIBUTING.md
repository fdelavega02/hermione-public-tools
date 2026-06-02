# Contributing

This repo is a public-safe export of local assistant tooling. Before editing or pushing, check that the change is useful without exposing private context.

## Public-safety checklist

Do not commit:

- Credentials, tokens, cookies, session files, or browser state
- Mailbox contents, message bodies, inbox exports, screenshots, HTML dumps, or generated summaries
- Private local paths, account identifiers, channel IDs, or machine-specific secrets
- Named people or organizations from Francisco’s private context
- Concrete private alert filters, sender rules, terms, or matching examples
- Generated runtime state, logs, caches, debug outputs, or `node_modules/`

Prefer:

- `config.example.json` instead of real config
- Placeholder names and generic examples
- Local-only setup steps
- Explicit approval boundaries for any draft, send, or publish step
- Concise docs that explain the reusable pattern without private details

If a detail helps only because it reveals Francisco’s private context, leave it out.
