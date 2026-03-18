# Minutes of Meeting

## Meeting Details

| Item | Details |
|------|---------|
| Meeting Title | Fintech Fuel Requirements - PGL Observations |
| Product | Fintech Fuel |
| Document Type | MoM |
| Date | Not specified in source notes |

## Participants

- Aamir Hafeez - Client, ODL
- Nida Wasif - Client, ODL
- Hassan Saeed Wattoo
- Imran Ashraf
- Hussain Ahmad
- Kamran Khalid

## Meeting Objective

Capture client observations and enhancement requirements for Fintech Fuel: dashboard visibility, reconciliation, DODO/COCO controls, alerts, reporting, sales targets, and head-office governance.

---

## Feature-Wise Discussion

### 1. Unified Platform and Multi-Site Visibility

- Single platform showing all plant sites; key data: storage tank volume and timestamped deliveries.
- Must scale to hundreds of sites (beyond a small pilot set). The client expects a performance upboost once pilot site testing is passed. Different roles (management, leadership, supervisors, operations) need different views; visibility and targets must follow user hierarchy.

### 2. Deliveries and Delivery Identification

- Distinguish **PGL** vs **non-PGL** deliveries (e.g. labels `PGL` / `Not PGL`). At COCO sites, current deliveries = PGL; at DODO sites, need identification of who decanted (multiple providers may share storage). Controls at DODO are weaker than COCO—client wants improvement.
- Use industry term **Delivery Note** (not Bill of Lading) on the UI; capture delivery/lading reference for imported products. Lading reference is reference-only for now; client wants it captured as KPI or future control even if initially wish-list.
- Highlight where actual decanted quantity exceeds expected. Manual delivery references alone are not sufficient for end-to-end control.

### 3. Dashboard Enhancements

- Flow: global picture → service station → product. Move from row/column history to visual, dashboard-driven presentation.
- Weekly fuel trend: support monthly and weekly scope; prefer sales-oriented visuals (e.g. bar charts). Tank widgets must clearly show whether levels are **real-time** or **filtered by date range**.
- Variance: support positive/negative highlighting (client asked for **green** positive, **orange** negative) and variance-based filtering.
- If dashboard works at hourly level only, remove minute-level filters from the dashboard.

### 4. Historical Comparison, Trend Analysis, and Forecasting

- Comparison and trend analysis for volumes and sales vs history (e.g. Gardenia site). Visibility of sales vs deliveries (demand vs supply) to support forecasting and to monitor/minimize external uplifting.

### 5. Reconciliation and Variance Analysis

- Reconciliation is critical (including monitoring external uplifting). Report should: show how often variation occurs; support **configurable thresholds** and **anomaly highlighting** above threshold; improve UI/UX (product-tile style—click product to see reconciliation; summary table on top with one row per product for the period).
- Clearly indicate when a site is not transmitting; show exclusion notes for missing data so variance is interpretable. Suggested rule: flag sites with no transmission for ≥60 minutes. Anomaly reasons should at least include data disconnectivity (internet or equipment). Reconciliation UI/UX quality is a possible road blocker.

### 6. Data Accuracy, Connectivity, and Manual Entry

- Two main failure cases: (1) internet disconnectivity—controller must buffer and transmit when back online; no data loss acceptable; (2) equipment/controller malfunction—system must alert and support **manual/offline entry** for lost data (admin-level tool; e.g. ATG down for two hours). Users must be able to tell whether reconciliation failure is due to connectivity vs equipment.

### 7. Alerts, Diagnostics, and Notifications

- Alerts must link to diagnostics and related reports (not isolated). Email notifications for alerts. Diagnostics/connectivity view for disconnected sites; configurable thresholds (e.g. 0.5% mentioned). If controller supports it, alert when device fails to ping within configured time.

### 8. Reporting, Terminology, and Units

- Use **Delivery Note** on UI (see §2). Report nomenclature (e.g. nozzle report) should be reviewed for business clarity. **Units of measure** (liters, currency) must be visible consistently across the application.

### 9. Nozzle/Dispenser and Meter Reading Visibility

- Improve nozzle-level sales visibility; management needs summaries, not only transaction detail. Single view for nozzle/pump-wise meter readings. Add **summarized report**: one entry per nozzle for selected date range, including meter readings (day-end is currently product-wise only).

### 10. Loss/Gain Analytics and Tank-Level Visibility

- Tank-wise visibility (in addition to product-wise) is **already covered by the Stock Levels report**; gap was client awareness and discoverability. Tank illustration is also required.

### 11. Sales Targets and Hierarchy

- Targets at yearly, quarterly, monthly, weekly (not daily). Weekly contribution to yearly targets matters. Client keeps targets in Excel at area/station level—wants **Excel upload** for bulk population (sample to be shared by client). Targets and sites must link to territory managers and org roles; current page does not fully reflect region/territory-manager hierarchy. Align with sales team on target model and expectations.

### 12. Date Range and Time Handling

- Date filter = full day(s): single day = 00:00–23:59; multi-day = start of first to end of last. No 8 AM–8 PM reconciliation window from UX perspective. Site-level time window possible with client approval. Transaction **end time** agreed for calculations; 5–10% tolerance discussed for midnight-crossing cases.

### 13. Price Monitoring and Price Change Control

- Head-office price to each dispenser (built for ALSONs; in progress for DOVER Fusion). If site sells above approved price, raise alert (exact match; no tolerance). Price comparison and history visible in app. Workaround needed when no transaction (PPU may not be received). Confirm whether FC controller communicates price changes directly.

### 14. Role-Based Access and Head-Office Control

- Ground staff: read-only. Head-office: edit and configuration. Manual entry and sensitive corrections restricted by role; client wants user roles and approval-through-hierarchy for manual data entry and control workflows.

### 15. Leak and Theft Alerts

- Client wants leak and theft alerting (e.g. tank dents, broken pipelines). Definitions and business rules need a separate session with Aamir.

---

## Decisions and Agreements

- Use **Delivery Note** on UI (not Bill of Lading).
- Remove minute-level filters from dashboard if only hourly aggregation is supported.
- Use **transaction end time** for date-based calculations.
- Add **summarized nozzle/meter reading report** for selected date ranges.
- Improve reconciliation UI/UX with anomaly visibility and missing-data indication.
- Alerts for connectivity and, where possible, equipment-down; configurable thresholds.
- Manual/offline data entry required when equipment prevents capture.

## Open Items

- Reliably identify PGL vs non-PGL deliveries (especially DODO).
- Whether controllers can provide equipment health / ping-based failure signals.
- Whether FC controller exposes price communication and price-change events.
- Configurable anomaly thresholds and application in reconciliation and dashboard.
- Align sales target model with sales team hierarchy and planning.
- Follow-up with Aamir on leak/theft alert requirements.
- Obtain sample sales target Excel from client.

## Action Items

| # | Action Item | Owner |
|---|-------------|--------|
| 1 | Dashboard: global→site→product navigation; clear real-time vs historical indicators | Octopus Digital |
| 2 | Deliveries: terminology to Delivery Note; capture lading/delivery reference | Octopus Digital |
| 3 | Propose control for PGL vs non-PGL (especially DODO) | Octopus Digital |
| 4 | Reconciliation: UI/UX, anomaly visibility, missing-data indication | Octopus Digital |
| 5 | Alerting for connectivity; investigate equipment-down via controller | Octopus Digital |
| 6 | Manual/offline data-entry workflow with role-based approval | Octopus Digital |
| 7 | Consistent unit-of-measure visibility across reports/screens | Octopus Digital |
| 8 | Add summarized nozzle/pump meter report for date range | Octopus Digital |
| 9 | Demonstrate Stock Levels report for tank-wise visibility to client | Octopus Digital |
| 10 | Share sample Excel for sales target upload | Client |
| 11 | Align target hierarchy with sales team | Client |
| 12 | Arrange discussion on leak/theft alerts | Client / Aamir / Octopus Digital |
