# CLAUDE.md — Pacing Report Project Guide

> **Keep this file current.** At the end of any session involving significant structural changes,
> new workflows, or architectural decisions, update the relevant sections here before closing out.

---

## What This Project Is

A Google Apps Script project embedded in a Google Sheet for Oats Overnight's CX team.
It tracks daily agent pacing against ticket-handling goals, generates weekly KPI snapshots,
and produces a Team Dashboard for leadership. Data comes from the Gorgias helpdesk API,
the team's weekly schedule, and a separate Pulse Log spreadsheet tracking inbox health.

**Timezone:** America/Phoenix  
**Gorgias subdomain:** oatsovernight  
**Repo:** OTR15/cx-pacing-staffing (GitHub)

---

## File Map

| File | Owns |
|---|---|
| `00-config.js` | `CFG` object — all defaults, sheet names, checkpoints, baseline goals, staffing params, excluded agents, default roster |
| `01-utils.js` | Shared helpers: `normalizeName_()`, `round1_()`, `round2_()`, `num_()`, date helpers, name parsing |
| `02-setup.js` | Menu registration (`onOpen`), Team Guide builder, Case Use Summary builder, tab organization, mode switching (Internal/External), seed/setup functions |
| `03-schedule.js` | Schedule tab parsing, normalization into `Schedule_Normalized`, `getScheduleMapForDate_()`, `getScheduleForRep_()` |
| `04-roster.js` | `getDisplayRoster_()` — reads the Roster sheet; agent ID/name map |
| `05-goals.js` | `getGoalsMap_()`, `getEffectiveGoals_()`, `adjustGoalByShift_()` — per-agent goal overrides from the Goals tab |
| `06-gorgias.js` | All Gorgias API calls — closed tickets, replies, messages, CSAT; `getClosedPerAgentMap_()`, `getStatNumber_()`, `getCsatMetrics_()` |
| `07-daily.js` | Daily tab layout: `buildDailySheet_()`, `getLayout_()`, column validation, sort/filter by manager, `getDailySheetRowMap_()` |
| `08-publish.js` | Checkpoint publishing: `publishCheckpointForDate_()`, `applyGoalAdjustments()`, status block helpers, `applyPacingColor_()` |
| `09-automation.js` | Time-based trigger installation/removal, `onOpen` trigger wiring |
| `10-weekly.js` | Weekly tab builder: `buildWeeklyReportForWeek_()`, `writeWeeklyLeadershipSummary_()`, `getPreviousWeekSummary_()` |
| `11-debug.js` | Debug/test helpers — safe to ignore for most changes |
| `12-staffing.js` | Staffing model: `buildStaffingModelForDate_()`, coverage computation, send-home recommendations, observed data reads from Pulse Log |
| `13-kpi-adjustment.js` | Deprecated/removed — file is a one-line tombstone |
| `14-weekly-kpi.js` | Weekly KPI snapshot: `collectWeeklyKpiSnapshot_()`, `weeklyKpiCalcAgentScore_()`, status assignment, auto-fail logic |
| `15-team-dashboard.js` | Team Dashboard tab: `buildTeamDashboard_()`, all section writers, trend table, chart |

---

## Key Architecture

### CFG and Config Sheet
`CFG` in `00-config.js` holds all code defaults. At runtime, `getConfigValue_(key, fallback)`
reads the **Config sheet** (a hidden tab in the spreadsheet) for overrides. Most tunable
values — goals, thresholds, schedule tab names, external spreadsheet IDs — live there.
`setConfigValue_()` writes back to the Config sheet.

### Column Layout — `getLayout_()`
**This is the source of truth for all column positions in daily tabs.** Never hardcode
column numbers for daily tab work — always call `getLayout_()` and use its properties:

```
sections[]         — one per checkpoint, each with: startCol, closedCol, repliedCol, messagesCol, csatCol
progressStartCol   — first progress column (On Track)
notesCol           — progressStartCol + 4  (Column AE in a standard layout) ← reporting tools read this
reviewFlagCol      — notesCol + 1  (Goal Adjustment status)
reviewReasonCol    — notesCol + 2  (Reason dropdown)
reviewAdjustCol    — notesCol + 3  (Hours Removed)
lastCol            — reviewAdjustCol
```

`progressLabels` in CFG: `['On Track', 'On a Project', 'Actions Taken', 'EOD Goal Met', 'Notes']`  
The **"On a Project"** column (`progressStartCol + 1`) is **hidden** on all new daily tabs —
it is unused. Do not remove it; the column index must stay intact for downstream offsets.

### Checkpoints
Defined in `CFG.checkpoints`. Each has a `percent` — the fraction of daily goal expected
by that time. The `percent` drives `getCheckpointTarget_()` and colors.

| Key | Hour | % of Daily Goal |
|---|---|---|
| 7AM | 7 | 10% |
| 9AM | 9 | 25% |
| 11AM | 11 | 40% |
| 2PM | 14 | 60% |
| 6PM | 18 | 85% |
| EOD | 23 | 100% |

Scheduled work ends at 8PM. The 6PM checkpoint is 85% of goal; EOD (23:00) captures final
numbers. For reporting and leadership purposes, 8PM and EOD are treated as the same event.

### Name Matching
`normalizeName_()` (lowercase, trim, collapse spaces) is used for **all** agent name
comparisons across schedule, roster, Gorgias, and KPI data. Always use it when comparing
names — never raw string equality.

### Goal Adjustment Flow
1. Supervisor enters hours to remove + reason on the daily tab
2. Runs **Pacing Report → Apply Goal Adjustments** (`applyGoalAdjustments()` in `08-publish.js`)
3. Function reads effective hours from the **Notes column (notesCol)** via regex on `Effective: X`
4. Recomputes adjusted goals, repaints all checkpoint colors
5. Recalculates EOD Goal Met if EOD has been published
6. **Updates the Notes column (Column AE)** with new target values — reporting tools depend on this

### Auto-Fail (Weekly KPI)
Triggered in `weeklyKpiCalcAgentScore_()` (`14-weekly-kpi.js`) when:
- QA score ≤ 74%, OR
- Tickets Replied < ~40% of weekly goal (scales proportionally with shift hours)

**Auto-fail does NOT force overallPct to 0.** The agent's real calculated score is used.
Auto-fail agents are included in the team-wide Overall Avg in `computeKpiReportCardStats_()`.
They are surfaced separately as a count and name list on the Team Dashboard.
> This was a deliberate change (April 2026) — do not reintroduce the `overallPct = 0` override.

---

## Supervisor Workflow (accurate as of April 2026)

**What supervisors actually do:**
- Open the daily tab and read pacing colors throughout the day
- Enter hours removed + reason in Goal Adjustment columns when an agent's available hours change (CTO, VTO, project work, absence, performance)
- Run **Apply Goal Adjustments** after entering adjustments
- Use **Filter Active Tab to Manager** / **Sort by Manager** for per-manager views
- Review weekly KPI tabs and the Team Dashboard at end of week

**What supervisors do NOT do:**
- Fill in the Actions Taken column (auto-populated by publish for Working Lunch, CTO, VTO, Off)
- Fill in On Track or EOD Goal Met (auto-populated by publish)
- Use the On a Project column (hidden, unused)

---

## Weekly KPI Snapshot

The weekly KPI data lives in a **separate spreadsheet** (not this one).
Its ID is in `CFG.weekly.kpiSnapshot.spreadsheetId` and in the Config sheet.
`14-weekly-kpi.js` reads from that spreadsheet's `CONFIG` sheet for weights and thresholds,
then writes the snapshot table to `ADMIN_VIEW` starting at row 32.

KPI statuses: Exceeding (≥106%) · Meeting (≥100%) · Close (≥90%) · Not Meeting · Auto-Fail · Exempt

---

## Team Dashboard

Built by `buildTeamDashboard_()` in `15-team-dashboard.js`, called at the end of
`buildWeeklyReportForWeek_()`. Sections: Team Performance, KPI Report Card, Week-over-Week,
QA Lead, Inbox Health (Pulse Log), Trend chart + data table.

The trend table (`DASH_R_TREND_DATA` = row 48+) persists across rebuilds — only rows
above it are cleared. `TD_OVR_EXCL` (column 7) is kept for backward compatibility with
WoW comparisons but both overall columns now write the same `avg` value.

---

## Staffing Tool (`12-staffing.js`)

**Current state: legacy model active, shadow model scaffolded.**

The model computes per-checkpoint: projected work remaining vs projected capacity remaining,
producing a recommendation of SEND / HOLD / CAUTION / BLOCK with a send-home count.

Two data paths:
- **Legacy (active):** static inflow estimate (tickets/hour × hours remaining)
- **Shadow (collecting baseline):** reads from Pulse Log spreadsheet —
  `Workable Volume Log` (hourly workable open inbox) and `Overnight Inflow Log`
  (overnight inflow proxy = SOD workable open − EOD workable open)

Shadow model needs ≥10 days of observed data (`minimumObservedSampleDays`) before it
can be used. Currently `useObservedData: false`, `observedDataBlendWeight: 0`.
The shadow model surfaces data on the Staffing tab but does not influence recommendations.

---

## External Spreadsheet Connections

| Purpose | Location of ID | Sheets Used |
|---|---|---|
| Weekly KPI snapshot | `CFG.weekly.kpiSnapshot.spreadsheetId` | `CONFIG`, `ADMIN_VIEW` |
| Pulse Log (inbox health + staffing) | `CFG.staffing.pulseLogSpreadsheetId` | `WoW Summary`, `Workable Volume Log`, `Overnight Inflow Log` |
| QA Lead Report Card | Config sheet key `QA_LEAD_REPORT_CARD_ID` | `Weekly Review`, `Week Archive` |

### Gaps — fill in as known
- **Gorgias API auth:** method, token location, rate limit behavior — see `06-gorgias.js`
- **Pulse Log sheet structure:** column layout of `WoW Summary`, `Workable Volume Log`,
  `Overnight Inflow Log` — partially documented in `15-team-dashboard.js` constants but
  full schema not captured here
- **Weekly KPI spreadsheet:** full CONFIG key list, how weights are set, who manages it
- **QA Lead Report Card:** who maintains it, how the Weekly Review row is written
- **Trigger schedule:** exact times daily triggers are configured to fire

---

## Tab Visibility Modes

**Internal** (default for team): Team Guide, 7 daily tabs, 4 weekly tabs, Staffing  
**External** (for leadership sharing): Case Use Summary, Team Dashboard, 3 daily, 1 weekly  

Admin tabs always hidden from normal users: `Config`, `Roster`, `Goals`,
`Schedule_Normalized`, `Schedule`

---

## Important Conventions

- All column positions in daily tabs come from `getLayout_()` — never hardcode
- Agent name comparisons always use `normalizeName_()`
- `getConfigValue_(key, fallback)` for any runtime-tunable value
- `round1_()` / `round2_()` for display values; `num_()` for safe number coercion
- `applyPacingColor_(range, status)` for all pacing cell coloring — `status` is a string
  like `'green'`, `'yellow'`, `'red'`, `'cto'`, `'vto'`, `'exempt'`, `'unscheduled'`
- Daily tab row positions resolved via `getDailySheetRowMap_(sheet)` — never assume row
  number from roster index; supervisors can sort/filter the tab

---

## Recent Significant Changes (April 2026)

- **Auto-fail scoring:** removed `overallPct = 0` override; real scores now used in team avg
- **Goal adjustment → Column AE:** `applyGoalAdjustments()` now updates the Notes column
  with adjusted target values so reporting tools stay in sync
- **On a Project column:** hidden at build time (`progressStartCol + 1`); column index kept
- **Team Guide:** fully rewritten as a formatted supervisor FAQ/handoff guide
- **Case Use Summary:** em dashes removed; Auto-Fail metric definition added; Pulse Log
  context note added (how to read inbox data alongside KPIs)
- **"Overall Avg (excl. auto-fails)"** removed from Team Dashboard — replaced by single
  accurate Overall Avg; auto-fails surfaced as count + name list only
