# Project Burgenland — Live Power SPV Model

Interactive dashboard deployed from this repository through GitHub Pages.

## Source of truth

The `main` branch of `tk-nexwell/burgenland-dc` is the definitive source. The local working copy is
`Documents/GitHub/burgenland-dc`. A one-way mirror is maintained in OneDrive at
`Output/Austria DC/github-pages-deploy`. Never publish from the OneDrive copy and never copy an
OneDrive edit back over GitHub. Make changes in the GitHub working copy, commit them to `main`, wait
for the Pages build, and then mirror the exact committed files to OneDrive.

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
