# Approval Boundaries

Automation in this repo should keep human approval clear.

## Default stance

- Draft locally first.
- Review before sending, publishing, or acting externally.
- Keep any send or publish action separate from detection, syncing, or drafting.
- Use explicit flags or separate commands for actions that leave the machine.

## Good patterns

- A sync command writes local output only.
- A draft command creates a review packet and sets `sendApproved: false`.
- A caller decides whether to notify, revise, or discard.
- External actions require a fresh, specific approval.

## Avoid

- Auto-sending mail from a draft generator
- Publishing generated text without review
- Hiding approval behind config defaults
- Mixing private runtime state into public examples
