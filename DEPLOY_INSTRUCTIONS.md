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
6. The checkout and the OneDrive mirror follow on their own if the scheduled task from
   `scripts/auto_sync.ps1 -Register` is installed. To move them immediately instead, run
   `powershell -File scripts/auto_sync.ps1`.
7. To rebuild the mirror by hand, preview with `powershell -File scripts/sync_onedrive.ps1`, then
   apply with `powershell -File scripts/sync_onedrive.ps1 -Apply`. The script exports committed Git
   bytes only, runs the release checks, verifies every declared file, moves the old mirror to a
   timestamped backup, installs the new mirror, and verifies it again.
8. Confirm `MIRROR_HEAD.txt` contains the published SHA and inspect `MIRROR_PROOF.json` for the
   per-file SHA-256 values. Then verify the live HTML, CSS, JavaScript, data, CSV and future-state
   image are served from that same SHA.

## Keeping the disk in step with GitHub

No scheduled task is installed at present. `scripts/auto_sync.ps1` is the unattended version of
steps 6 and 7, for whoever wants to install it.

    powershell -File scripts\auto_sync.ps1 -Register     install the scheduled task
    powershell -File scripts\auto_sync.ps1               run one synchronisation now
    powershell -File scripts\auto_sync.ps1 -Unregister   remove the scheduled task

It only ever moves work from GitHub to the disk. It fetches `origin/main`, resets the checkout to
the published commit when it is behind, clears line-ending-only noise, and rebuilds the mirror when
`MIRROR_HEAD.txt` is not already at that commit. If the checkout carries real local edits it stops
and names them instead of resetting, because those edits are the one thing a sync cannot recover.

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
