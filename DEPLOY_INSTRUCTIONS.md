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

## Release sequence

1. Make changes only in `C:\Users\ThomasKoenig\GitHub\burgenland-dc`.
2. Review `git status` and `git diff`; stage only the intended public files and release controls.
3. Run `powershell -File scripts/check_release.ps1` and resolve every blocker. The tracked-file
   checks deliberately catch a new asset that was created but not staged.
4. Commit the complete change to `main` and push it to `origin`.
5. Wait for the `Public release check` workflow and the GitHub Pages deployment to succeed at the
   same commit SHA.
6. With the canonical checkout clean, preview the mirror operation with
   `powershell -File scripts/sync_onedrive.ps1`. The preview still fetches and verifies the source,
   but does not copy, rename, or delete anything.
7. Build the mirror with `powershell -File scripts/sync_onedrive.ps1 -Apply`. The script exports
   committed Git bytes only, runs the release checks, verifies every declared file, moves the old
   mirror to a timestamped backup, installs the new mirror, and verifies it again.
8. Confirm `MIRROR_HEAD.txt` contains the published SHA and inspect `MIRROR_PROOF.json` for the
   per-file SHA-256 values. Then verify the live HTML, CSS, JavaScript, data, CSV and future-state
   image are served from that same SHA.

Do not use `git add -A` without reviewing the result. Do not publish a transfer bundle, generated
proof file, encrypted meter archive, earlier visual candidate or other local-only artifact to Git.
The mirror script exports the exact tracked set declared by the manifest; an undeclared tracked
file blocks the release instead of silently appearing in OneDrive or GitHub Pages.

`deploy.bat` is retired and intentionally performs no deployment or synchronization.

## Cache rule

When `gdc.css`, `gdc_data.js`, or `gdc_app.js` changes, also change the matching `?v=` asset token in
`index.html`. This prevents an investor's browser from retaining an older dashboard asset after a
successful deployment.

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
