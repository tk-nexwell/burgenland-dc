# Deploying GDC Nickelsdorf

## One source, one direction

GitHub repository `tk-nexwell/burgenland-dc`, branch `main`, is the definitive source. Work from
`C:\Users\ThomasKoenig\GitHub\burgenland-dc`. The OneDrive folder `Output/Austria DC/github-pages-deploy` is a
one-way mirror for convenience and backup.

- Make every dashboard change in GitHub.
- Run `powershell -File scripts/check_release.ps1` and resolve every blocker.
- Commit the complete change to `main`.
- GitHub Pages deploys automatically after the push.
- Verify the Pages build succeeded and the live site returns the new assets.
- Run `powershell -File scripts/sync_onedrive.ps1` to create an exact, hash-verified OneDrive mirror.
- The previous OneDrive mirror is moved to a timestamped sibling archive before replacement, so stale
  files and its legacy `.git` cannot contaminate the new mirror and remain recoverable.
- Never push an OneDrive edit back to GitHub and never deploy from OneDrive.
- Do not work, commit or push from any archived OneDrive `.git` directory.

`deploy.bat` is retired and intentionally performs no deployment.

## Cache rule

When `gdc.css`, `gdc_data.js`, or `gdc_app.js` changes, also change the matching `?v=` asset token in
`index.html`. This prevents an investor's browser from retaining an older dashboard asset after a
successful deployment.

## Confidentiality

This repository and every GitHub Pages asset are public. The browser access code does not provide
security because the JavaScript and data are downloaded before the gate is applied. Encrypted
commercial blobs, meter data, clear-text aggregates and NDA-derived labels are all blocked by the
release check. Publish only the sanitized planning case; host a full counterparty model behind real
server-side authentication. The existing public Git history still requires a separate owner decision
because deleting a current file does not erase older commits.
