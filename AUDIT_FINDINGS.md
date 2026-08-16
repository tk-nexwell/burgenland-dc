# GDC Nickelsdorf dashboard — independent review

Review date: 17 August 2026

## Release assessment

The revised current files are suitable as a public, illustrative planning dashboard, subject to the
two unresolved publication decisions below. They are not an authenticated data room, an agreed
commercial model, an engineering design, or investment advice.

Release checks and desktop/mobile browser walkthroughs pass. The live release should be mirrored
one way from the clean GitHub checkout; the OneDrive folder must not be used as a second Git worktree.

## Critical findings corrected

- **PPA tranche logic:** the interface described the first 200 MW at €100/MWh and the balance at
  €80/MWh, but the default wind and solar models still used €100/MWh. The calculated default blend
  is now applied consistently, including reset and scenario flows. On the reviewed case this changes
  wind equity IRR from roughly 18.5% to 12.8% and solar from 23.4% to 15.9%.
- **Battery integrity:** the hourly supply simulation previously began with a half-full battery,
  contributing about 1.68 GWh that was never charged. State of charge now starts at zero.
- **Workbook parity:** the workbook now uses the dashboard's day-by-day spread inputs, self-charge
  share, network-charge toggle and discounted LCOE method rather than a different average-day model.
- **Scenario safety:** imported and URL-carried scenarios now pass through a typed allowlist with
  size limits. Unknown keys, non-finite values and markup-bearing strings are rejected before render.
- **Browser confidentiality:** encrypted terms and in-browser decryption were removed. A browser
  access code is explicitly described as a presentation/navigation gate, not authentication.
- **Claims:** counterparty approval, fixed pricing, network-charge exemptions, route/topology,
  redundancy, delivery dates and contract status are no longer stated as established facts. The
  dashboard distinguishes published, counterparty-stated, modeled and illustrative material.
- **Navigation and accessibility:** every main view has a shareable hash route; tab changes open at
  the page heading; browser back/forward works; semantic buttons, focus states, active-state labels,
  compact mobile navigation and safer chart labels were added.
- **Solar control:** the range input now accepts 357 MWp exactly instead of snapping to the nearest
  five-megawatt step on first interaction.

## Model limitations that remain visible

- The consolidated Power SPV case still uses annual renewable/load netting. Hourly battery dispatch,
  losses, grid imports and delivery obligations are not yet integrated into the SPV cash flow.
- Battery arbitrage remains a historical perfect-foresight gross ceiling, not a forecast or a
  bankable revenue case. Revenue stacking and operational conflicts need a dispatch optimizer.
- A 500 MW / 4 GWh battery is an eight-hour reserve. It is not proof of campus uptime, Tier compliance
  or an independent electrical source.
- Connection design, route, permits, grid tariffs, tax, PPA terms, credit support and risk allocation
  require adviser and counterparty confirmation.

## Visual and product review

The dark-green Nexwell system, aerial masterplan and page-specific technical diagrams now form a
coherent visual language. The largest improvement is the future-state masterplan: recognizable halls,
solar fields, turbines, substation, storage and agricultural context make the project legible before
the user enters financial detail. The electrical route is deliberately subordinate rather than the
dominant visual feature.

Recommended next phase: evolve the masterplan into a lightweight digital twin.

1. Add a phase slider: existing land → first power → 200 MW → 500 MW build-out.
2. Add layer toggles for generation, grid, storage, campus, land/permits and water/fibre.
3. Let a selected building or asset open a small status card with capacity, phase, evidence status,
   owner and unresolved decision—not a paragraph of prose.
4. Animate energy flow only when the user requests it, with reduced-motion support.
5. Add two camera modes: investor overview and engineering schematic. Avoid pretending that an
   illustrative rendering is a surveyed or permitted site design.

## Publication decisions still required

1. **Repository history:** earlier public commits may retain material removed from current files.
   Restrict the repository or authorize a clean-history public replacement after a confidentiality
   review. Do not rewrite history without the owner's explicit approval.
2. **Real restricted access:** if detailed commercial or source material is to return, move it behind
   server-side identity and authorization. GitHub Pages and client-side JavaScript cannot protect it.

## Source-of-truth rule

- Canonical working source: clean GitHub checkout of `tk-nexwell/burgenland-dc`.
- Published source: reviewed commit on `main` after release checks pass.
- OneDrive: generated, hash-verified mirror of that exact commit, without `.git`, bundles or extras.
- Any intentional change starts in the canonical checkout, is reviewed and published, then mirrored.
