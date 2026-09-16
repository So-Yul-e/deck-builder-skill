# Updates for clone-and-link installs

This package is a standalone skill folder, not a marketplace plugin. A skill's
newest instructions can be discovered only after its local files have changed;
starting a session does not fetch this GitHub repository by itself.

## Consent and trigger

At the beginning of a skill invocation run, before generation:

```bash
python3 <skill_dir>/scripts/update_skill.py --auto
```

Use an available Python 3 or the package's managed interpreter. On Windows use
`py -3` or `python`. If Python is missing, proceed with runtime preflight; do not
install it just for updating. Auto is off until the installer explicitly enables
it once. Enabling grants future fast-forward updates from the official origin/main
in the same clone; it does not grant resets, stashing, overwriting user edits,
installing new dependencies or global session hooks.

```bash
python3 <skill_dir>/scripts/update_skill.py --enable
python3 <skill_dir>/scripts/update_skill.py --check
python3 <skill_dir>/scripts/update_skill.py --disable
```

--check reports local consent/status without fetching. --auto updates only opted-in
clean main clones from the recorded official origin and throttles fetches to once
per hour. A changed result means reread SKILL.md and only the relevant references
once before continuing. Do not recursively call auto again in that invocation.
Normal runtime preflight still checks new dependency requirements without installing
anything. Warnings/skips leave the installed code usable; explain a meaningful
problem once, continue deck creation and label remote freshness as unverified.

## Supported install boundary

Works when the installed deck folder is linked to a clone's deck folder (or used
directly in that clone). The origin is this project's official GitHub URL, branch
is main, and the installer has enabled updating. The metadata lives in
`deck/.deck-update.json` and must remain ignored and local. Git and Python 3 are
required. Local edits, a different branch, changed origin, diverged history,
concurrent update or offline state stop updating. No local work is discarded.

Copied folders and marketplace cache copies lack the original checkout and cannot
self-update using this helper. Refresh the copy using the installer or migrate to
a clone-and-link installation. Never rewrite plugin caches through this helper.
Old installations without this script and entrypoint need one manual refresh before
they can opt in. Main commits must be published by the maintainer before clients
can receive them. A feature branch/local-only commit is not a released update.

## Host behavior and references

Codex detects local skill changes and supports linked folders; if changes do not
appear, restart Codex. This helper fetches repository changes during skill usage;
it is not an unconditional session-start hook. For workspace GitHub marketplace
imports, the host provides daily sync, which is a different distribution mechanism.
Do not claim daily marketplace sync for this standalone folder.

- https://learn.chatgpt.com/docs/build-skills
- https://learn.chatgpt.com/docs/enterprise/plugin-management
