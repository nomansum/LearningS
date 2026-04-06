# SNS's Study Plan — Final Implementation Plan v3

---

## 🗂️ Sheet Architecture

| # | Sheet Name | Purpose |
|---|---|---|
| 1 | ⚙️ Config | Hidden — dropdown source lists, named ranges |
| 2 | 🛤️ Learning Paths | Add paths directly, archived tasks live here |
| 3 | 📚 Study Log | Active tasks only — sorted by Path → Start Date, completed sink to bottom |
| 4 | 📋 Tasks by Category | Auto-grouped mirror of Study Log by path (Apps Script) |
| 5 | 📅 Task Scheduler | Today panel + This Week panel (non-completed tasks only) |
| 6 | 📊 Progress & Charts | Per-path doughnut charts + overall stats |
| 7 | 🔗 Resource Library | Resources linked to paths |
| 8 | ❓ How To Use | Guide |

---

## Sheet 1 — ⚙️ Config (hidden)

Named ranges powering every dropdown across all sheets:

| Named Range | Source | Used In |
|---|---|---|
| `LP_Names` | Learning Paths!B4:B1000 | Study Log col B, Resource Library col E, Scheduler col D |
| `Resource_Types` | Config!A2:A9 | Study Log col J, Resource Library col C |
| `Status_Options` | Config!B2:B6 | Study Log col H |
| `Done_Options` | Config!C2:C4 | Study Log col L, col N |
| `Session_Types` | Config!D2:D6 | Scheduler col E |
| `Platforms` | Config!E2:E9 | Learning Paths col D |
| `Path_Status` | Config!F2:F5 | Learning Paths col J |
| `Confidence` | Config!G2:G7 | Study Log col O |

---

## Sheet 2 — 🛤️ Learning Paths `[UPDATED]`

### Top section — Active Paths Table (rows 1–53)

Type directly into this table to add paths. No button needed.

```
A  #
B  Learning Path Name   ← LP_Names named range reads from here
C  Description
D  Platform / Source    (dropdown: Platforms)
E  Total Topics Planned (manual)
F  Topics Completed     =COUNTIF(Study Log col H filtered by path, "Completed")  [Apps Script fills]
G  % Complete           =IFERROR(F/E, 0)  → conditional format as color gradient bar
H  📍 Where I Left Off  (manual — the exact topic/video you stopped at)
I  Last Studied Date    =MAXIFS(Study Log Completed Date col, Path col, this path)  [Apps Script]
J  Status               (dropdown: Not Started / In Progress / On Hold / Completed)
K  Notes / Next Steps
```

### ⚡ Key behavior — Path Completion Archive:

When **Status in col J = "Completed"**:
- Apps Script `onEdit` fires
- All tasks in Study Log where Learning Path = this path name are **moved** (cut+paste) into an **archived sub-table** directly below this path's row in Learning Paths sheet
- The archived block has a **collapsed row group** (Google Sheets row grouping) so it stays tidy
- Study Log is then re-sorted

### Archived sub-table structure (per completed path, grouped/collapsible):
```
  ↳ [Archived Tasks — Pre-Security]          ← merged header, indented, grey
      Task | Tag | Start | Completed | Day-4 ✓ | Day-7 ✓ | Confidence | Resource
      ...rows of archived tasks...
```

### Sorting:
- Active paths: In Progress first → then Not Started → then On Hold → Completed last
- Completed paths with their archive blocks always at the bottom

---

## Sheet 3 — 📚 Study Log `[CORE]`

### Rules:
- **Only active (non-archived) tasks live here**
- Sorted automatically: by Learning Path (A→Z) → then by Start Date (oldest first)
- Completed tasks visually sink via conditional formatting + Apps Script re-sort on status change
- When a task's Status = Completed AND its Learning Path is also Completed → archive trigger fires

### Column Layout:
```
A  #                   (auto-number via Apps Script after sort)
B  Learning Path       (dropdown: LP_Names)
C  Topic / Task
D  Subject / Tag
E  Start Date          ← Apps Script date picker on click
F  Initiated Date      ← Apps Script date picker on click
G  Completed Date      ← Apps Script date picker on click
H  Status              (auto-set: Not Started→Initiated→In Progress→Completed)
                        also manually overrideable
I  Resource Link / Notes
J  Resource Type       (dropdown: Resource_Types)
K  Day-4 Date          =IF(G<>"", G+3, "")   [auto]
L  Day-4 Done?         (dropdown: Done_Options)
M  Day-7 Date          =IF(G<>"", G+6, "")   [auto]
N  Day-7 Done?         (dropdown: Done_Options)
O  Confidence (1–5)    (dropdown: Confidence)
P  Remarks
```

### Conditional Formatting (row-level):
| Condition | Color |
|---|---|
| Both Day-4 ✅ AND Day-7 ✅ | Pale green — Mastered |
| Day-4 due TODAY, not done | Light red |
| Day-7 due TODAY, not done | Orange |
| Day-4 OVERDUE, not done | Deeper red |
| Status = Initiated | Light blue |
| Status = In Progress | Yellow |
| Status = Completed (revisions pending) | Light green |
| Status = Completed + both revisions done | Bold pale green |

### Sorting logic (Apps Script, fires on status change or path change):
```
Primary sort:   Status != "Completed"  → top  |  Status = "Completed" → bottom
Secondary sort: Learning Path A → Z
Tertiary sort:  Start Date oldest → newest
```

---

## Sheet 4 — 📋 Tasks by Category `[NEW]`

### Purpose:
Read-only grouped view. Apps Script auto-generates one section per learning path.

### Per-path section structure:
```
┌─────────────────────────────────────────────────────────────┐
│ 🛤️ Pre-Security        3 / 8 completed   37%   In Progress  │  ← auto header
├────────────────┬──────┬────────┬────────┬──────┬────────────┤
│ Task           │ Tag  │ Start  │ Status │ D4 ✓ │ D7 ✓  │ ⭐ │
├────────────────┼──────┼────────┼────────┼──────┼────────────┤
│ ...tasks...    │      │        │        │      │       │    │
└────────────────┴──────┴────────┴────────┴──────┴────────────┘
  📎 Resources for this path:
     • [Resource Title] — [Type] — [URL]
```

### Refresh:
- **"🔄 Refresh View" button** at top → runs Apps Script `rebuildCategoryView()`
- Also auto-refreshes when Study Log is edited (debounced — max once per 30s)

### Completed paths:
- Their section shows archived tasks (pulled from Learning Paths archive block)
- Greyed out with a "✅ Path Completed" badge

---

## Sheet 5 — 📅 Task Scheduler `[NEW]`

### Panel A — 📌 TODAY (top, rows 1–18)

```
┌──────────────────────────────────────────────────┐
│ 📌 TODAY — Monday, 06 April 2026                 │
├──────────────────────────────────────────────────┤
│ ⚡ Auto-suggested (revision due today):           │
│   Task | Path | Type | Due Reason               │
│   ...FILTER(Study Log, Day-4 or Day-7 = today)  │
├──────────────────────────────────────────────────┤
│ ✏️ Manual slots (add extra tasks for today):      │
│   Task (dropdown) | Path (auto) | Session | Mins | Done? │
└──────────────────────────────────────────────────┘
```

### Panel B — 📆 THIS WEEK (rows 20+)

```
Week of: Mon DD MMM – Sun DD MMM     [← Prev Week]  [Next Week →]

Date | Day  | Task (dropdown)  | Learning Path | Session Type | Planned Mins | Actual Mins | Done? | Notes
Mon  |       |                  |               |              |              |             |       |
Tue  |       |                  |               |              |              |             |       |
...
```

### Task dropdown rules (enforced by Apps Script):
```javascript
// Only tasks where Status != "Completed" appear
=FILTER('Study Log'!C:C, 'Study Log'!H:H <> "Completed", 'Study Log'!C:C <> "")
```
- Selecting a task auto-fills Learning Path (col D) via INDEX/MATCH
- If a task gets completed in Study Log while it's in the scheduler → it greys out with strikethrough

### Color rules:
- Today's row → blue highlight
- Done = ✅ → grey + strikethrough
- Revision session types (Day-4/Day-7) → red/orange label badge

---

## Sheet 6 — 📊 Progress & Charts `[NEW]`

### Section A — Overall Stats strip (row 1–5)
```
Total Paths | Active Paths | Total Tasks | Completed Tasks | Overall % | Due Today | Overdue
```

### Section B — Per Learning Path Cards + Charts (dynamic, Apps Script generated)

For each learning path, a card is generated:

```
┌─────────────────────────────────────────────────┐
│ 🛤️ Pre-Security                    In Progress  │
│ Platform: TryHackMe                             │
│ Progress: ████████░░░░  3 / 8  (37%)           │
│ Last Studied: 04-Apr-2026                       │
│ 📍 Left off at: Introductory Networking         │
│                                                 │
│  [Doughnut Chart: Completed vs Remaining]       │
└─────────────────────────────────────────────────┘
```

### Charts:
- One **doughnut chart per path** auto-inserted by Apps Script
- Chart data range = dynamic named range per path (counts from Study Log)
- Chart title = path name
- Colors: Completed = green slice, Remaining = light grey slice
- Charts auto-update as Study Log changes (Google Sheets native chart behavior)

### Refresh:
- **"🔄 Rebuild Charts" button** → `rebuildProgressCharts()` — regenerates all cards + charts
- Run this after adding a new learning path

---

## Sheet 7 — 🔗 Resource Library `[UPDATED]`

```
A  #
B  Resource Title
C  Type               (dropdown: Resource_Types)
D  URL / Location
E  Learning Path      (dropdown: LP_Names)
F  Related Topic      (free text)
G  Quality ⭐          (dropdown: Confidence)
H  Notes
```

Resources surface in Tasks by Category under their linked path's section.

---

## 🔗 Full Data Linkage Map

```
⚙️ Config
  └─ Named Ranges ──────────────────────► all dropdowns everywhere

🛤️ Learning Paths (col B = LP_Names)
  ├─► 📚 Study Log          col B dropdown
  ├─► 🔗 Resource Library   col E dropdown
  ├─► 📅 Scheduler          col D auto-fill
  └─► 📊 Charts             path list for card generation
        │
        └── Completed Path trigger
              └─► Archives tasks FROM Study Log INTO Learning Paths sub-table

📚 Study Log (single source of truth for task data)
  ├─► 📋 Tasks by Category   FILTER/QUERY grouped by path
  ├─► 📅 Task Scheduler      FILTER non-completed tasks for dropdown
  ├─► 📊 Progress Charts     COUNTIFS per path for chart data
  └─► 🛤️ Learning Paths      MAXIFS for Last Studied Date, COUNTIF for completed count

🔗 Resource Library
  └─► 📋 Tasks by Category   shows resources per path block
```

---

## ⚙️ Apps Script Functions

| Function | Trigger | Does |
|---|---|---|
| `onEdit(e)` | Any edit | Router — calls relevant sub-functions |
| `showDatePicker(range)` | Click col E/F/G in Study Log | HTML dialog → writes date back |
| `autoSetStatus(row)` | Edit col F or G in Study Log | Sets H = status based on dates |
| `sortStudyLog()` | Status/Path change in Study Log | Sorts: active by Path→Date, completed to bottom |
| `archivePathTasks(pathName)` | Learning Path status → Completed | Moves tasks from Study Log to LP archive block, groups rows |
| `rebuildCategoryView()` | Button / Study Log edit | Rebuilds Tasks by Category sheet section by section |
| `rebuildProgressCharts()` | Button | Regenerates all path cards + doughnut charts |
| `updateSchedulerDropdown()` | Scheduler col C click | Filters Study Log for non-completed tasks |
| `autoFillPathInScheduler(row)` | Scheduler col C edit | INDEX/MATCH fills col D |
| `updateLPStats()` | Study Log edit | Pushes COUNTIF+MAXIFS results to Learning Paths cols F, I |
| `autoNumberRows(sheet)` | Any row add/sort | Renumbers col A sequentially |
| `setup()` | Run once manually | Creates named ranges, installs all triggers |

---

## 📦 Build Order

1. `setup()` Apps Script → named ranges + triggers
2. ⚙️ Config sheet
3. 🛤️ Learning Paths sheet + archive block structure
4. 📚 Study Log + sorting + conditional formatting + date picker
5. 🔗 Resource Library
6. 📋 Tasks by Category + `rebuildCategoryView()`
7. 📅 Task Scheduler (Today + Week panels)
8. 📊 Progress & Charts + `rebuildProgressCharts()`
9. ❓ How To Use