# Epics Planner — SEAL EPICS Guide

This guide explains how to use the **SEAL EPICS** feature on the Epics Planner (Epics Management) page. Sealing records an **approved date** for selected epics and locks their planned dates until you choose to revise them.

**Note:** Delivery Status, Remarks, and Actual Production Date are no longer edited on Epics Planner. They are managed per meeting in **IPP Meeting Planner** (Settings → IPP Meeting Planner) in the IPP Builder. The IPP Meeting Dashboard shows epics from the current Scheduled meeting and their per-meeting delivery status and remarks.

---

## Where to find it

- **Page:** Epics Planner (Epics Management)
- **URL:** `/settings/epics-management`
- **Navigation:** Settings / configuration → Epics Planner

---

## What sealing does

- **Sealing** saves a snapshot of the selected epics (metadata, phases, man-days, dates, status) and marks them as **sealed**.
- Sealed epics show a **lock icon** next to the epic name. Their planned dates are fixed until you **RE-BUDGET**.
- Each time you click **SEAL IT**, the system records a new **approved date**. You can review what was sealed on any past date.
- The list of **approved (sealed) dates** for each epic is visible in the **IPP Meeting Dashboard** → **RMI Details** drawer.

---

## Step-by-step: Sealing epics

### 1. Select epics

- In the Epics Planner table, use the **Select** column (first column).
- **Check the box** for each epic you want to seal. You can select multiple epics.
- Group rows (Project / Product Categorization / Component) and the draft row do not have checkboxes; only **epic rows** can be selected.

### 2. Open the seal modal

- When at least one epic is selected, the **SEAL EPICS** button (red) in the header bar becomes **enabled**.
- Click **SEAL EPICS**. A confirmation modal opens.

### 3. (Optional) Review a past sealed date

- The modal shows **Last sealed dates** — previous times someone clicked **SEAL IT**.
- Click any **date button** to load and review the snapshots from that seal:
  - Epic name, key, project, delivery status, plan status
  - Phase table: phase name, man-days, start date, end date
- Use **Close** in the review panel to hide it. This does not seal anything.

### 4. Confirm sealing

- Click **SEAL IT** in the modal footer.
- The selected epics are sealed: they get the lock icon, and a new **approved date** is stored.
- The modal closes, selection is cleared, and the table reloads so you see the lock icons.

---

## Revising a sealed epic (RE-BUDGET)

- For a **sealed** epic, the **Actions** column shows a **RE-BUDGET** button (in addition to Edit, Save, Sync Jira Epic, Delete).
- Click **RE-BUDGET** for that epic. The epic is **unsealed** (lock removed) so you can change dates and other plan data.
- After editing, you can **seal the epic again** by selecting it and using **SEAL EPICS** → **SEAL IT**. That creates a new approved date.

---

## Viewing sealed-date history (IPP Meeting Dashboard)

- Open **IPP Meeting Dashboard** (e.g. `/ipp_meeting_dashboard.html` or from Reports).
- Click an RMI (epic) on the roadmap to open the **RMI Details** drawer.
- In the drawer, find the section **Approved dates (sealed)**.
- It lists all **approved dates** when this epic was sealed (newest first). If it was never sealed, it shows: *No sealed dates recorded for this RMI.*

---

## Summary

| Action | Where | Result |
|--------|--------|--------|
| Select epics | Epics Planner — Select column | Enables SEAL EPICS button |
| Seal | SEAL EPICS → SEAL IT | Lock icon; new approved date; snapshot stored |
| Review past seal | SEAL EPICS modal → click a date | View snapshots (specs, phases, dates) for that date |
| Revise sealed epic | Actions → RE-BUDGET | Unseal; you can edit and seal again |
| See history | IPP Meeting Dashboard → RMI Details | Approved dates (sealed) list |

---

## Notes

- **Approved dates** are stored in the system and are **not** yet used by the Approved vs Planned Hours report; that integration is planned for later.
- Sealing does not change Jira; it only records the state in the Epics Planner database (`assignee_hours_capacity.db`).
