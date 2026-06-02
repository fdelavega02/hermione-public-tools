# Local-first Automation

Local-first automation keeps sensitive state on the user’s machine and exposes only the reusable pattern.

## Principles

- Authenticate manually in a local browser when possible.
- Save browser/session state only in ignored local folders.
- Keep config values local and provide public-safe examples.
- Write outputs to ignored folders such as `output/`, `state/`, or `logs/`.
- Prefer small scripts that can be inspected and run manually.
- Keep send, publish, or external actions behind explicit approval.

## Public examples

Public docs should explain how to adapt the tool without revealing private context. Use placeholders, generic filter names, and short notes about what each setting controls.

The public repo should show the shape of the automation, not Francisco’s private data.
