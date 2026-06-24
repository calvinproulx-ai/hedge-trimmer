# Project Hedge Trimmer — Deal Dashboard

Live deal dashboard for Phenix Salon Suites corporate location sales.

**Live URL:** `https://<your-github-username>.github.io/hedge-trimmer`

---

## How to update the dashboard

### Option 1 — Upload via GitHub website (no coding required)

1. Go to your repository on GitHub
2. Click into the `data/` folder
3. Click `CIM_Master_Data.xlsx`
4. Click the **pencil icon** (Edit) → then click **"Upload a file"** to replace it
5. Commit the change with a message like `"Update financials June 2026"`
6. GitHub automatically runs the workflow — dashboard updates in ~30 seconds

### Option 2 — From Claude (this tool)

Upload the updated xlsx here and say **"update the dashboard."**
Claude will regenerate `index.html` and you can download it.

---

## How to set up GitHub Pages (first time only)

1. Push this repo to GitHub: `https://github.com/new`
2. Go to **Settings → Pages**
3. Under **Source**, select **Deploy from a branch**
4. Choose `main` branch, `/ (root)` folder → click **Save**
5. Your dashboard will be live at `https://<username>.github.io/<repo-name>` within 1–2 minutes

---

## File structure

```
hedge-trimmer/
├── index.html                    ← Generated dashboard (served by GitHub Pages)
├── data/
│   └── CIM_Master_Data.xlsx      ← UPDATE THIS FILE to refresh the dashboard
├── scripts/
│   └── generate_dashboard.py     ← Reads xlsx, writes index.html
├── template/
│   └── dashboard.html            ← HTML/CSS/JS template (edit for design changes)
└── .github/
    └── workflows/
        └── generate.yml          ← Auto-runs generate_dashboard.py on xlsx changes
```

---

## What the dashboard shows

- All corporate locations grouped by market (Colorado, MN/FL, Midwest, WI/IL, Arizona, Texas, NC, California)
- Per-location: suites, occupancy, LTM revenue, LTM EBITDA, P12/mature EBITDA, ask price, AWR
- Click locations to build a custom package with live totals
- Export a summary table or generate a CIM prompt

## Data fields pulled from xlsx (Locations tab)

| Dashboard field | xlsx column |
|---|---|
| Name | Location Name |
| State | State |
| Region | Region / Package |
| Suites | Suites |
| Occupancy | Paid Occ% (5/5) |
| LTM Revenue | LTM Rev ($K) |
| LTM EBITDA | LTM EBITDA ($K) |
| P12/Mature EBITDA | Mature EBITDA ($K) |
| Ask Price | Ask Price ($K) — falls back to Val @ 5.5x ($K) |
| AWR | Wkly Rev/Suite |
| Notes (card badge) | Notes (first sentence) |

Package ask prices come from the **Packages tab** → `Total Ask Override ($K)`.
