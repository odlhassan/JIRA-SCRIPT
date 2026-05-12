# EPR Tool Introduction

Report ID: `introduction`

## Purpose

The Introduction page establishes the EPR Tool positioning and clarifies the value chain from raw Jira data to leadership-ready reporting insights for Octopus Digital.

## Key Sections

| Section | What it explains | Typical contents |
| --- | --- | --- |
| Jira Input | What is available directly from Jira before local transformations. | Issue hierarchy, assignees, statuses, estimates, dates, worklogs, JQL-scoped project extraction. |
| EPR Enrichment | What EPR Tool orchestrates on top of Jira data. | Canonical normalization, planner/IPP merge, rollups, capacity/risk/penalty derivations, business rules. |
| EPR Output / Insights | What leaders and managers can act on. | Dashboards, schedule drift visibility, utilization/performance views, governance and recovery signals. |
| Detailed Software Architecture | How the platform is functionally layered internally. | Operational intake, standardization, business interpretation, canonical reporting memory, insight composition, experience and action modules. |
| Detailed Data-Flow Architecture | How information moves through the platform end to end. | Raw operational capture, trust shaping, business enrichment, persisted reporting truth, report composition, localhost delivery, user action. |

## Business Interpretation

- Jira is the operational source.
- EPR Tool is the transformation and orchestration layer.
- Leadership reports are the decision layer used for planning, execution, and governance actions.

## Architecture Notes

- The page now ends with two Mermaid-rendered diagrams so technical stakeholders can see both the system architecture and the staged data flow without leaving the frontend.
- The software architecture content is now written in functional-module language so readers understand what each architectural component does, why it exists, and how it supports adjacent modules.
- The system architecture diagram now has a dedicated full-width presentation area plus frontend zoom controls for easier inspection during leadership or PMO walkthroughs.
- The data-flow section explains the sequencing from raw Jira facts to canonical reporting state and finally to leadership-facing report consumption on localhost.
