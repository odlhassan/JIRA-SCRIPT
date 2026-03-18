-- Total work by projects (1. Total work by projects) – no bindings; edit dates below if needed.
-- Date range: 2026-02-01 to 2026-03-31 (change the 4 literals in the CTEs if needed).

WITH
run AS (
  SELECT last_success_run_id AS run_id
  FROM canonical_refresh_state
  WHERE id = 1
),

scoped_subtasks AS (
  SELECT
    i.issue_key,
    UPPER(i.project_key) AS project_key,
    UPPER(COALESCE(NULLIF(TRIM(i.epic_key), ''), 'NO_EPIC')) AS epic_key,
    i.assignee,
    i.issue_type,
    i.start_date,
    i.due_date,
    COALESCE(i.original_estimate_hours, 0) AS planned_hours
  FROM canonical_issues i
  JOIN run r ON r.run_id = i.run_id
  WHERE UPPER(i.project_key) <> 'RLT'
    AND (
      LOWER(i.issue_type) LIKE '%sub-task%'
      OR LOWER(i.issue_type) LIKE '%subtask%'
    )
    AND (
      (i.start_date <> '' AND i.start_date BETWEEN '2026-02-01' AND '2026-03-31')
      OR
      (i.due_date   <> '' AND i.due_date   BETWEEN '2026-02-01' AND '2026-03-31')
    )
),

worklogs AS (
  SELECT
    w.issue_key,
    SUM(COALESCE(w.hours_logged, 0)) AS total_hours_all,
    SUM(
      CASE
        WHEN w.started_date BETWEEN '2026-02-01' AND '2026-03-31' THEN COALESCE(w.hours_logged, 0)
        ELSE 0
      END
    ) AS total_hours_in_range
  FROM canonical_worklogs w
  JOIN run r ON r.run_id = w.run_id
  GROUP BY w.issue_key
),

epic_meta AS (
  SELECT
    UPPER(i.issue_key) AS epic_key,
    i.summary AS epic_summary,
    i.status  AS epic_status,
    i.start_date AS epic_start_date,
    i.due_date   AS epic_due_date
  FROM canonical_issues i
  JOIN run r ON r.run_id = i.run_id
  WHERE i.issue_key <> ''
)

SELECT
  s.project_key,
  s.epic_key,
  COALESCE(em.epic_summary, '') AS epic_summary,
  COALESCE(em.epic_status,  '') AS epic_status,
  COALESCE(em.epic_start_date, '') AS epic_start_date,
  COALESCE(em.epic_due_date,   '') AS epic_due_date,
  ROUND(SUM(s.planned_hours), 2) AS planned_hours,
  ROUND(SUM(COALESCE(w.total_hours_in_range, 0)), 2) AS actual_hours_log_date,
  ROUND(SUM(s.planned_hours) - SUM(COALESCE(w.total_hours_in_range, 0)), 2) AS plan_actual_difference_log_date,
  ROUND(SUM(COALESCE(w.total_hours_all, 0)), 2) AS actual_hours_planned_dates,
  ROUND(SUM(s.planned_hours) - SUM(COALESCE(w.total_hours_all, 0)), 2) AS plan_actual_difference_planned_dates,
  COUNT(*) AS subtask_count,
  COUNT(DISTINCT NULLIF(LOWER(TRIM(s.assignee)), '')) AS assignee_count
FROM scoped_subtasks s
LEFT JOIN worklogs w ON w.issue_key = s.issue_key
LEFT JOIN epic_meta em ON em.epic_key = s.epic_key
GROUP BY s.project_key, s.epic_key
ORDER BY planned_hours DESC, s.project_key, s.epic_key;
