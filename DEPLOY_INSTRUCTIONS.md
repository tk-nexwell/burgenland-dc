# Deploying GDC Nickelsdorf

## One source, one direction

GitHub repository `tk-nexwell/burgenland-dc`, branch `main`, is the definitive source. The OneDrive
folder `Output/Austria DC/github-pages-deploy` is a one-way mirror for convenience and backup.

- Make every dashboard change in GitHub.
- Commit the complete change to `main`.
- GitHub Pages deploys automatically after the push.
- Verify the Pages build succeeded and the live site returns the new assets.
- Mirror the exact committed files listed in `MIRROR_MANIFEST.json` to OneDrive.
- Never push an OneDrive edit back to GitHub and never deploy from OneDrive.
- Do not modify or synchronize the legacy `.git` directory in OneDrive.

`deploy.bat` is retired and intentionally performs no deployment.

## Cache rule

When `gdc.css`, `gdc_data.js`, or `gdc_app.js` changes, also change the matching `?v=` asset token in
`index.html`. This prevents an investor's browser from retaining an older dashboard asset after a
successful deployment.

## Confidentiality

This repository is public. Raw 15-minute metering is published only as encrypted
`bedata_enc.bin`. Model assumptions and outputs present in JavaScript remain visible to anyone who
reads the public repository, regardless of the dashboard's client-side access gate.
