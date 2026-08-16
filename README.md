# Project Burgenland — Live Power SPV Model

Interactive dashboard deployed from this repository through GitHub Pages.

## Source of truth

The `main` branch of `tk-nexwell/burgenland-dc` is the definitive source. Never publish from a
OneDrive export and never copy an OneDrive edit back over GitHub. Make changes in a Git working
copy at `C:\Users\ThomasKoenig\GitHub\burgenland-dc`, commit them to `main`, wait for the Pages
build, and only then mirror the exact committed files to OneDrive. The sync script stages and hashes
the new mirror and archives the prior OneDrive folder intact before replacement.

The mirror manifest and detailed operating rules are in `MIRROR_MANIFEST.json` and
`DEPLOY_INSTRUCTIONS.md`.

## Return convention

Wind and solar both use a 2029 first-generation year. Wind, solar and battery asset returns exclude
private-line capex because they are plant-only views. The consolidated Power SPV return includes all
assets assigned to that SPV, including the private line and interface scope.

## Public release boundary

GitHub Pages is public and the in-browser access code is only a presentation gate. Run
`scripts/check_release.ps1` before every publication. Confidential counterparty data must live in an
actually authenticated private environment, not in JavaScript delivered by this site.

The public production curves are deterministic planning profiles, not project meters and not a
bankable energy-yield study. Historical market-price records remain separate from those profiles.
