# Independent Dashboard Audit

Review date: 16 August 2026

## Release decision

Do not publish the current repository as a public GitHub Pages site. `gdc_data.js` contains
clear-text Burgenland Energie production aggregates and an NDA-related source marker. The browser
access screen does not protect those files: every visitor can download the JavaScript and data
assets directly.

Removing the material from the next commit is not sufficient by itself because the public Git
history may retain earlier versions. Before the next public release, choose one of these paths:

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
- A third party has no durable role if it only introduces the utility and operator or resells the
  same energy. The added risk matrix makes the credible role test explicit: Nexwell must finance,
  guarantee, integrate, or operate a measurable service that neither bilateral party will provide.
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

## Decisions required before publication

- Confirm which material is suitable for a public site and which requires authenticated access.
- Confirm the proposed Power SPV ownership, payment flows, credit support, and risk allocation.
- Confirm the direct-line route, connection topology, redundancy standard, capex ownership, and
  network-charge treatment with qualified Austrian legal and grid advisers.
- Replace illustrative performance assumptions with sourced, dated inputs and show confidence or
  sensitivity ranges for every investment-critical output.
