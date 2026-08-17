# Project Burgenland: Live Power SPV Model

Interactive public-planning dashboard deployed through GitHub Pages.

## One editable source

The only supported local Git checkout is:

`C:\Users\ThomasKoenig\GitHub\burgenland-dc`

Edit, review, test, commit and push from that checkout only. The committed `main` branch is the
publication record and GitHub Pages deploys that record.

The similarly named OneDrive folder is deliberately **not** a Git checkout and must not be edited:

`C:\Users\ThomasKoenig\OneDrive - nexwell.com\Output\Austria DC\github-pages-deploy`

It is a generated, read-only working mirror for convenient access. After a release,
`scripts/sync_onedrive.ps1 -Apply` fetches `origin`, requires a clean canonical checkout with
`HEAD == origin/main`, runs the public-release checks, exports exactly the files declared in
`MIRROR_MANIFEST.json`, verifies their hashes, and replaces the mirror by same-volume renames. The
previous mirror remains beside it as a timestamped backup. The new mirror contains commit and
SHA-256 proof files but never a `.git` directory, undeclared extras, or local-only material.

See `DEPLOY_INSTRUCTIONS.md` for the operating procedure. Treat a file present in OneDrive but not
in the canonical Git commit as an output or backup, never as a second source of truth.

## Return convention

Wind and solar both use a 2029 first-generation year. Wind, solar and battery asset returns exclude
private-line capex because they are plant-only views. The consolidated Power SPV return includes all
assets assigned to that SPV, including the private line and interface scope. Its operating cash flow
uses the same 8,760-hour renewable, battery and residual-grid dispatch as the Supply and data-center
views; annual generation-minus-load netting is not used.

## Public release boundary

GitHub Pages is public and this release has no browser access-code gate. Any future client-side gate
would be presentation-only, not security. Run `scripts/check_release.ps1` before every publication.
The same check runs in GitHub Actions.
Confidential counterparty data, internal commercial terms and project-meter records must live in an
actually authenticated private environment, not in JavaScript delivered by this site or in Git
history.

The public production curves are deterministic planning profiles, not project meters and not a
bankable energy-yield study. Historical market-price records remain separate from those profiles.

The future-state masterplan is an illustrative visualization, not an approved site plan, survey,
engineering design or construction commitment.
