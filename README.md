# Learning Progress Tracker v2.59.3 — TA Student Email Edit

## v2.59.3 — TA Student Email Edit

- TA can edit only a student email address from the Students page.
- Student ID, name, section/team-related profile fields, active status, roster import/add-student actions, access-link generation, and other roster management remain OWNER/INSTRUCTOR only.
- Server-side authorization enforces email-only updates for TA even if a crafted request includes other fields.
- TA email changes are recorded separately in the audit log as `UPDATE_STUDENT_EMAIL`.
- Schema remains 15; no `setupProject()` rerun is required.


This release makes the CLO Definitions page read-first and compact. Official CLO statements are always visible; editing opens only the selected CLO.

## Changes
- CLO Definitions now show a compact read-only list by default.
- Each row shows CLO code, official statement, and the optional student-friendly explanation when available.
- Edit opens only one CLO inline; other CLOs stay compact.
- Add CLO opens a single new-item editor instead of turning the whole page into a form.
- Official CLO statements are no longer hidden behind the student-friendly-description fields.
- Existing attainment and Self-Evaluation flows are unchanged.
- Schema remains 15; no `setupProject()` rerun is needed.


## Changes
- CLO definitions now include optional **Student-friendly description**.
- Generated Start / Midterm / Final Google Forms use five full 1–5 response descriptions.
- The official CLO statement remains the question title; the student-friendly explanation appears as question help text.
- Existing response import remains compatible because ratings are parsed from the leading 1–5 value.
- Existing courses migrate automatically from schema 14 to schema 15; do not run `setupProject()` again.

# v2.57.2 — TQF3 DOCX Import Hotfix

- Fixes valid `.docx` files failing at **ตรวจสอบข้อมูล** because Apps Script `Utilities.unzip()` was given a DOCX MIME blob instead of a ZIP blob.
- DOCX is now validated as an OOXML ZIP package (`PK` signature) and unzipped using `application/zip`.
- Improves the error message for files that are not valid OOXML DOCX packages.
- Keeps the v2.57.1 OBE export presets and selective TQF3 preview/apply behavior.
- Schema remains 14; do not rerun `setupProject()` on existing trackers.

- Adds real DOCX TQF3 import from the Course page.
- Preview before applying, grouped by Course Profile, Objectives, CLOs, Sections, Staff, Weekly Plan, Activities/Key Assessments, and detected CLO threshold.
- Item-level checkboxes; import only selected items.
- Existing data is preserved by default; optional explicit overwrite with confirmation.
- DOCX parsing is performed server-side using the OOXML document and tables; no external service is required.
- Schema remains 14. Existing courses do not run setupProject().

# v2.56.3 — Per-Section Actions & Editing

Improves Course > Section management so each Section is handled independently instead of opening every Section editor at once.

## New / Changed
- Each Section has its own vertical **⋮** action menu.
- **Edit Section** opens only that Section's details inline.
- **Delete Section** removes only the selected Section after confirmation.
- A Section that still has students assigned remains protected by the existing server-side guard; students must be moved/unassigned before the Section can be removed.
- **+ เพิ่ม Section** opens a single blank Section editor instead of switching the whole Section area into edit mode.
- Class schedule and exam rows can still be added/removed inside the selected Section editor.
- Section IDs are preserved when editing, including when the Section code itself changes.
- The Section summary is responsive and no longer relies on one large all-Section edit form.

## Compatibility
- Schema remains **14**.
- Existing trackers must **not** run `setupProject()` again.
- Google Form Self-Evaluation features from v2.56.0 and responsive Staff actions from v2.56.1 are retained.


## v2.56.3 — Course Setup Tools Placement
- Moves course-level data tools above the Getting Started checklist.
- Renames actions to concise Thai labels: ส่งออกรายวิชา, นำเข้าต้นแบบ, นำเข้า TQF3.
- Export and template import reuse the existing course data tools.
- TQF3 button is reserved for the upcoming reviewed preview/import flow.


## v2.57.2
- Renamed Full Course Archive to Full Course Backup.
- Added Learning Results / OBE Export preset and OBE-friendly summary sheets.
- Preserved Course Template and Custom export modes.

## v2.58.4

- Self-Evaluation imports that use an evaluation-period column can now be filtered to import only Start, Midterm, or Final/End-of-term rows.
- The preview reports how many rows from other periods were skipped before any data is written.
- Leaving the filter as “All periods in file” preserves the previous mixed-period import behavior.

## v2.58.2
Generated CLO Self-Evaluation Google Forms now collect verified Google email addresses, allow response editing, and limit each Google account to one response. Because FormApp does not expose the Google Forms 'Send responders a copy: Always' setting, the tracker installs a form-submit trigger that emails each respondent a copy of their submitted answers.


## v2.58.4 — Latest Self-Evaluation Radar Emphasis
- The latest available self-evaluation period (Final, otherwise Midterm, otherwise Start) is drawn as a solid, slightly thicker line on both class and student CLO radar charts.
- Earlier selected periods are dashed.
- Legends mark the latest period explicitly.



## v2.58.7 — Period-Specific Self-Evaluation Questions

Generated CLO Self-Evaluation Google Forms now tailor the post-CLO questions by evaluation period. Start and Midterm forms keep the optional Reflection item but show a period-specific prompt beneath it. Final/End-of-term forms replace Reflection with eight required 1–5 course-learning rating questions and four optional open-ended questions for course reflection and CQI. Existing forms are unchanged; regenerate/replace a form to use the new question set.

## v2.58.5 — Dashboard CLO Reference
- Adds hover descriptions to CLO codes in the Dashboard CLO summary.
- Adds a collapsible `View all CLO descriptions` reference below the summary table.
- Shows the official CLO statement and, when available, the student-friendly description.
- Schema remains 15.


## v2.58.8 — Compact Weekly Plan Summary
- Weekly-plan cards are summary-first: key topic items, CLO, short teaching/assessment summaries, and linked work only.
- Long content is intentionally summarized with + more counts instead of being visually clipped mid-sentence.
- A ดูรายละเอียด expander shows the full topic/description, teaching methods, assessment methods, linked activities, and any actual-teaching notes/adjustments.
- No schema change (schema 15).


## v2.59.0 — Compact Weekly Plan Cards

- Removed the expandable "ดูรายละเอียด" block from weekly-plan cards.
- Weekly cards now stay summary-only; full topic/CLO/method/activity details are available through the per-week edit action.
- Keeps the existing compact summary, weekly action menu, actual-teaching note/adjustment actions, and schema version 15.


## v2.59.0 — CLO Summary Bar Chart
- Adds Table / Bar chart toggle in Dashboard CLO Summary.
- Grouped bars per CLO: Evidence, Start, Midterm, Final on a 0–5 axis.
- Evidence uses the existing orange accent; Start/Midterm/Final use one blue family from light to dark.
- Missing periods are omitted rather than rendered as zero.
- Radar self-evaluation colors now use the same light-to-dark blue family for consistency.


## v2.59.1 — CLO Attainment Progress View
- Added a Cards / Progress bar view switch for dashboard CLO attainment.
- Progress view shows each CLO as a 0–100% bar with attained / evidence student counts.
- CLOs with no evidence remain N/A rather than being shown as 0%.
- Rows remain clickable to open CLO attainment details.


## v2.59.2 — CLO Attainment Progress Bars
- Replaces the old CLO attainment cards with the progress-bar view as the only dashboard presentation.
- Removes the Cards / Progress view switch.
- Keeps 0–100% attainment bars, attained / evidence student counts, N/A for CLOs without evidence, and click-through to CLO attainment details.
