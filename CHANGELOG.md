# Changelog

## v0.2.1 – 2026-11-13
- *Behind baseline* now counts documents whose baseline date has passed and that are not yet approved; documents approved ahead of their baseline no longer lower the figure
- *Archive Weekly Snapshot*: first run no longer leaves an empty row in *PM History*; re-running on the first day is safe; the summary reports the rows added this run
- README and rollout guide: which figures follow the status date and which show today's status
- Added `.gitignore` (data-policy safeguard) and MIT `LICENSE`

## v0.2 – 2026-11-10
- Work Package filter: *All*, a single WP, or *Custom selection* of several WPs via a side panel
- Document set filter (Key deliverables + any Yes/No set column), replacing the single-set filter
- New matrix: status per document set; WP matrix now follows the document-set filter
- Helper columns *Dash WP selected* and *Dash in set* in the table drive the filters
- Rollout guide (Markdown + PDF) and dashboard screenshots

## v0.1 – 2026-11-10
- First demo workbook with KPI tiles, S-curve (baseline vs. actual + forecast), document flow, status per WP, key-deliverables view, attention list
- Office Scripts: *Build PM Dashboard* and *Archive Weekly Snapshot*
- Copilot prompt
