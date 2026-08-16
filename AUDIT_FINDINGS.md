# Independent Dashboard Audit

Review date: 16 August 2026

## Release decision

The current release files have been sanitized for a public narrative dashboard. Project-meter
aggregates, NDA source markers, and the encrypted meter archive have been removed; production
charts now use a deterministic planning profile labelled as illustrative. The browser access screen
remains a presentation control, not authentication.

The public Git history may still retain earlier versions. Current-file remediation therefore does
not replace a repository-history and confidentiality review. Choose one of these paths for full
remediation:

1. Publish a sanitized public narrative dashboard and move the detailed model to an authenticated
   private application. This is the recommended approach.
2. Move the complete dashboard to private hosting with server-side authentication and restrict the
   repository. GitHub Pages alone cannot provide that control.

Any history cleanup, repository visibility change, or data classification decision should be
approved by the relevant owner before it is performed.

## Commercial and model findings

- The previous consolidated Power SPV return omitted private-line and interface capex while the
  narrative assigned that infrastructure to the SPV. The consolidated return now includes it;
  wind, solar, and battery views remain plant-only returns.
- Several statements presented planning assumptions as agreed facts, including firm delivery,
  ownership, network-charge treatment, technical topology, and delivery dates. The dashboard now
  labels these as targets, candidate structures, or open decisions.
- The Power SPV proposition is strongest when its contracted scope follows assets and services it
  can finance, coordinate, operate, and measure. The dashboard now presents that proposed interface
  positively while preserving the network operator's and campus operator's distinct duties.
- The battery dispatch result is a perfect-foresight gross optimizer ceiling, not a forecast of
  achievable trading returns or proof of continuous firm power. It is now labelled accordingly.
- The proposed 500 MW / 4 GWh battery provides eight hours at full discharge. It cannot by itself
  support an 8,760-hour firm-power promise; grid import, generation diversity, curtailment rules,
  reserve margins, and contractual remedies remain open.

## Experience improvements

- Added an interactive isometric masterplan with selectable layers and decision-focused hotspots.
- Added a project explorer for direct navigation across the commercial and technical views.
- Added a party, payment, and risk matrix so the business proposition is visible alongside the
  engineering story.
- Reduced dense and repetitive copy, strengthened planning-status labels, and corrected mobile
  overflow, tooltip, segmented-control, and splash-screen behavior.
- Added release and one-way mirror checks so GitHub remains the source of truth and OneDrive becomes
  a verified snapshot rather than a competing codebase.

## Decisions still required

- Decide whether to make the repository private or perform an approved history rewrite for earlier
  sensitive versions.
- Confirm the proposed Power SPV ownership, payment flows, credit support, and risk allocation.
- Confirm the direct-line route, connection topology, redundancy standard, capex ownership, and
  network-charge treatment with qualified Austrian legal and grid advisers.
- Replace illustrative performance assumptions with sourced, dated inputs and show confidence or
  sensitivity ranges for every investment-critical output.
