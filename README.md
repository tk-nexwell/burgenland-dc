# Project Burgenland — Live Power SPV Model

Interactive dashboard deployed from this repository through GitHub Pages.

## Source of truth

The `main` branch of `tk-nexwell/burgenland-dc` is the definitive source. A one-way mirror is
maintained in OneDrive at `Output/Austria DC/github-pages-deploy`. Never publish from the OneDrive
copy and never copy an OneDrive edit back over GitHub. Make changes here, commit them to `main`,
wait for the Pages build, and then mirror the exact committed files to OneDrive.

The mirror manifest and detailed operating rules are in `MIRROR_MANIFEST.json` and
`DEPLOY_INSTRUCTIONS.md`.

## Return convention

Wind and solar both use a 2029 first-generation year. Every equity IRR displayed by the dashboard
excludes private direct-line capex. The line investment remains visible in total system capex and is
financed separately at Power SPV; direct-line losses remain in delivered generation.
