# Deploying GDC Nickelsdorf

## One source, one direction

GitHub repository `tk-nexwell/burgenland-dc`, branch `main`, is the definitive source. Work from
`Documents/GitHub/burgenland-dc`. The OneDrive folder `Output/Austria DC/github-pages-deploy` is a
one-way mirror for convenience and backup.

- Make every dashboard change in GitHub.
- Run `powershell -File scripts/check_release.ps1` and resolve every blocker.
- Commit the complete change to `main`.
- GitHub Pages deploys automatically after the push.
- Verify the Pages build succeeded and the live site returns the new assets.
- Run `powershell -File scripts/sync_onedrive.ps1` to mirror the verified clean commit to OneDrive.
- Never push an OneDrive edit back to GitHub and never deploy from OneDrive.
- Do not modify or synchronize the legacy `.git` directory in OneDrive.

`deploy.bat` is retired and intentionally performs no deployment.

## Cache rule

When `gdc.css`, `gdc_data.js`, or `gdc_app.js` changes, also change the matching `?v=` asset token in
`index.html`. This prevents an investor's browser from retaining an older dashboard asset after a
successful deployment.

## Confidentiality

This repository and every GitHub Pages asset are public. The browser access code does not provide
security because the JavaScript and data are downloaded before the gate is applied. `bedata_enc.bin`
protects one raw series, but clear-text aggregates or NDA-derived labels in `gdc_data.js` are also a
disclosure and are blocked by the release check. Publish only a sanitized public data module; host a
full counterparty model behind real server-side authentication.
