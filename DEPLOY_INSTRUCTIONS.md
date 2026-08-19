# Deploying GDC Nickelsdorf

## One source, one generated mirror

Use only this local Git checkout for authoring and publication:

`C:\Users\ThomasKoenig\GitHub\burgenland-dc`

It is the canonical checkout for `tk-nexwell/burgenland-dc`. GitHub `main` is the committed
publication record.

This OneDrive path is an output, not a second checkout:

`C:\Users\ThomasKoenig\OneDrive - nexwell.com\Output\Austria DC\github-pages-deploy`

Do not edit, commit from, or add a `.git` directory to the OneDrive mirror. Rebuild it from the
canonical checkout after publication.

## Publication is a push

GitHub Pages serves `main` at `gdc-nickelsdorf.com`. A push to `main` is the deployment; there is
no separate publish step and no batch file to run. `deploy.bat` has been deleted.

## Release sequence

1. Make changes only in `C:\Users\ThomasKoenig\GitHub\burgenland-dc`.
2. Review `git status` and `git diff`; stage only the intended public files and release controls.
   Windows editors write CRLF while this repository stores LF, so `git status` can mark every file
   as modified. `git diff --ignore-cr-at-eol` shows what actually changed.
3. Run `powershell -File scripts/check_release.ps1` and resolve every blocker. The tracked-file
   checks deliberately catch a new asset that was created but not staged.
4. Commit the complete change to `main` and push it to `origin`. That push publishes the site.
5. Wait for the `Public release check` workflow and the GitHub Pages deployment to succeed at the
   same commit SHA.
6. Bring the checkout back in step when you next need it: `git fetch origin` then
   `git reset --hard origin/main` from a clean tree.
7. To rebuild the mirror, preview with `powershell -File scripts/sync_onedrive.ps1`, then
   apply with `powershell -File scripts/sync_onedrive.ps1 -Apply`. The script exports committed Git
   bytes only, runs the release checks, verifies every declared file, moves the old mirror to a
   timestamped backup, installs the new mirror, and verifies it again.
8. Confirm `MIRROR_HEAD.txt` contains the published SHA and inspect `MIRROR_PROOF.json` for the
   per-file SHA-256 values. Then verify the live HTML, CSS, JavaScript, data, CSV and future-state
   image are served from that same SHA.

## Nothing runs on a schedule

There is no scheduled task, on this machine or any other. A push to `main` is the deployment, and
GitHub Pages needs nothing else. The OneDrive mirror is a convenience copy: rebuild it with
`scripts/sync_onedrive.ps1 -Apply` when somebody wants one, and leave it otherwise.

A previous `scripts/auto_sync.ps1` ran this every fifteen minutes through Task Scheduler. It has
been removed. If anything like it is ever reinstated, launch it through `conhost.exe --headless`
and hide every child process too: `-WindowStyle Hidden` on its own still allocates a console, which
paints a window on the desktop at every run.

Do not use `git add -A` without reviewing the result. Do not publish a transfer bundle, generated
proof file, encrypted meter archive, earlier visual candidate or other local-only artifact to Git.
The mirror script exports the exact tracked set declared by the manifest; an undeclared tracked
file blocks the release instead of silently appearing in OneDrive or GitHub Pages.

## Cache rule

When `gdc.css`, `gdc_data.js`, or `gdc_app.js` changes, also change the matching `?v=` asset token in
`index.html`. This prevents an investor's browser from retaining an older dashboard asset after a
successful deployment.

`model_export.js` is not loaded by `index.html`. It is fetched on the first click of the Excel
button, and `gdc_app.js` reads the release token off its own script tag and appends it to that
fetch. So bumping the token in `index.html` releases the workbook builder too, and nothing extra is
needed. Do not change that fetch back to a bare filename: GitHub Pages serves the file with a long
max-age, so a bare fetch leaves every returning reader building the workbook from whichever copy
their browser happened to cache first, while the dashboard around it updates normally. The release
check fails if the token is dropped.

## Confidentiality

This repository and every GitHub Pages asset are public. The current release has no access-code
gate; adding a browser-only gate would not provide security because JavaScript and data are
downloaded before such a presentation control is applied.
Encrypted project-meter archives, password verifiers, browser decryption code, internal commercial
terms, NDA-derived labels and unqualified counterparty claims are release blockers. Publish only a
sanitized public data module; host the full counterparty model behind real server-side
authentication.

## What the proof means

`MIRROR_PROOF.json` records the verified commit, remote URL and a SHA-256 plus Git blob ID for every
declared file in the installed mirror. `MIRROR_HEAD.txt` repeats the full commit SHA for quick
inspection. Together with a successful post-install verification, they prove that the local
OneDrive mirror was built from the fetched `origin/main` commit with no missing or extra files.

The proof does not establish that the OneDrive cloud client has finished uploading, that another
device has downloaded the mirror, or that GitHub Pages has deployed. Those remain separate checks.
Timestamped `github-pages-deploy.backup.*` folders are intentionally retained for recovery; review
and remove old backups manually only after the new mirror and cloud copy have been confirmed.
