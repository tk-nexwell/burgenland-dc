# GDC Nickelsdorf dashboard: independent review

Review date: 17 August 2026

## Release assessment

The reviewed dashboard is suitable as a public, illustrative planning interface. It is not an
authenticated data room, approved site plan, engineering design, executed commercial agreement,
bankable forecast or investment recommendation.

The release is intentionally explicit about what is proposed, modelled and still open. It does not
represent counterparty approval, an uptime/Tier outcome, a guaranteed tariff or supply price, or a
surveyed route.

## Material corrections

- Removed the browser access-code illusion and browser-delivered encrypted commercial terms.
- Replaced non-public project identities, precise coordinates, status dates and supplied commercial
  references with rounded, anonymised public planning assumptions.
- Added typed scenario validation, safe default merging, enum/range checks and import/share limits.
- Aligned the battery calculation and display around fractional tradeable duration, limited supported
  duration cases to 2/4/6/8 hours, capped ancillary capacity at 225 MW and removed free opening state
  of charge.
- Carried hourly residual-grid demand into both the SPV and data-center views instead of using annual
  generation-minus-load netting.
- Made IT load derive from facility load and PUE; fixed scenario-comparison units and the 357 MWp
  solar control.
- Added shareable page routes, browser history, focus/scroll handling, selected-state semantics,
  labelled controls and a compact mobile navigation row.
- Added an interactive future-state masterplan with six selectable assets and explicit Open actions.
- Withheld the Excel download. The workbook still lacks full parity for hourly dispatch, portfolio
  toggles, clipping, infrastructure scope and SPV reconciliation.

## Visual assessment

The future-state masterplan is now the clearest project visual: recognizable halls, battery blocks,
substation, solar fields, turbines, roads, farm context, shadows and a subordinate electrical
corridor make the proposal legible before the financial detail. The dark green, charcoal and gold
system is coherent and premium on desktop, and the phone layout has no page-level horizontal
overflow.

The Overview remains long and the tab row partly duplicates the Project Explorer. The next design
phase should emphasize progressive disclosure: a phase slider, map-layer toggles, compact evidence
and status cards, an investor/engineering view switch, and a mobile bottom sheet. Responsive
WebP/AVIF assets should replace the large PNG once the visual direction is approved.

## Residual risks

1. Earlier Git history still contains a removed meter-data archive and a cheap verifier for its old
   access code. Treat the code as exposed. Any history purge, repository replacement or visibility
   change requires the data owner's approval.
2. GitHub Pages is public static hosting. Confidential source material and negotiated terms require
   a separate server-authenticated application.
3. The battery can currently support both firming and merchant revenue in the model without a full
   co-optimization. Market eligibility, operational conflicts and tariff treatment remain open.
4. Generation, grid connection, route, permits, phasing, commercial structure and counterparties
   require dated evidence and qualified technical/legal validation.
5. A 500 MW / 4 GWh battery represents eight hours at full discharge; it is not proof of continuous
   campus uptime, Tier compliance or unlimited firm-power liability.

## Source-of-truth rule

`C:\Users\ThomasKoenig\GitHub\burgenland-dc` is the only editable checkout. The reviewed `main`
commit is the publication source. The OneDrive `github-pages-deploy` folder is a generated, non-Git
mirror of that exact commit, built only after release checks pass and verified by full commit ID,
file inventory and SHA-256 hashes. It must never be edited or used as a second Git worktree.
