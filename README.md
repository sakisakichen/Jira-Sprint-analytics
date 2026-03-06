# Jira Sprint Analytics Pipeline

An end-to-end automated data pipeline that extracts sprint data from Jira Cloud across **20+ Agile teams**, transforms it with Python, and delivers a formatted Excel dashboard — replacing hours of manual reporting with a single command.

🔗 **[Live Dashboard Preview](https://sakisakichen.github.io/Jira-Sprint-analytics)**

---

## The Problem

Each sprint cycle required manually pulling data from Jira across 20+ teams, copying it into Excel, and formatting reports by hand. This process was time-consuming, error-prone, and created reporting delays for stakeholders.

## The Solution

A two-script Python pipeline that:
1. **Extracts** all sprint issues via Jira REST API (with pagination to bypass the 1,000-record limit)
2. **Transforms** and validates the data — flagging missing assignees, missing hours, and incomplete story points
3. **Visualizes** results in a 4-sheet Excel dashboard with charts, KPI cards, and color-coded tables

---

## Impact

- ⏱ Reduced sprint reporting time from **manual hours → single command**
- 📊 Automated reporting across **20+ Agile teams** and **3,000+ Jira issues**
- 🚩 Surfaced data quality issues (missing assignees, missing hours) that were previously invisible
- 👥 Enabled cross-functional stakeholders to self-serve sprint performance insights

---

## Tech Stack

| Tool | Purpose |
|---|---|
| **Python 3.11** | Core scripting and orchestration |
| **Jira REST API v3** | Data extraction via `/rest/api/3/search/jql` |
| **pandas** | Data transformation, aggregation, quality flagging |
| **openpyxl** | Excel automation — charts, conditional formatting, multi-sheet workbooks |
| **HTTP Basic Auth** | Secure Jira Cloud authentication via API token |
| **Regex** | Sprint T-number extraction for correct numeric sorting |

---

## Pipeline Architecture

```
Jira Cloud API
      │
      ▼
jira_all_queries.py
  ├── 7 JQL queries (split by sprint range)
  ├── nextPageToken pagination (bypasses 1,000-record limit)
  ├── Deduplication across queries
  └── Output: jira_query_all.xlsx
      │
      ▼
visualization_excel_v5.py
  ├── Filter to included T-number teams
  ├── Compute sprint hours, story points, data quality flags
  └── Output: jira_dashboard.xlsx (4 sheets)
```

---

## Dashboard Sheets

| Sheet | Contents |
|---|---|
| **Summary** | KPI cards — total stories, hours logged, sprints below threshold, flags |
| **01 Hours per Sprint** | Bar + line combo chart, red highlight for sprints below 640hr standard |
| **02 Missing Story Points** | Pie chart + detail table of stories without estimates |
| **03 No Assignee or No Hours** | Color-coded flag breakdown by sprint and issue type |

---

## Key Technical Highlights

**Pagination beyond 1,000 records**
```python
# Uses nextPageToken to bypass Jira Cloud's hard limit
while True:
    params = {"jql": jql, "maxResults": 100, "fields": FIELDS}
    if next_page_token:
        params["nextPageToken"] = next_page_token
    data = requests.get(url, headers=headers, params=params, auth=auth).json()
    all_issues.extend(data.get("issues", []))
    next_page_token = data.get("nextPageToken")
    if not next_page_token:
        break
```

**Sprint T-number extraction for correct numeric sort**
```python
def extract_sprint_num(sprint_name):
    matches = re.findall(r'T(\d+)', str(sprint_name))
    return int(matches[-1]) if matches else 9999
```

**Combo chart: bar + threshold line overlay**
```python
chart = BarChart()
line  = LineChart()  # 640hr standard reference line
chart += line        # merged into single combo chart
```

---

## Setup

```bash
# Install dependencies
pip install pandas openpyxl requests urllib3

# Add your Jira credentials to jira_all_queries.py
EMAIL     = "your_email@company.com"
API_TOKEN = "your_jira_api_token"

# Run the pipeline
python jira_all_queries.py        # Step 1: Extract → jira_query_all.xlsx
python visualization_excel_v5.py  # Step 2: Visualize → jira_dashboard.xlsx
```

> **Note:** API credentials are not committed to this repo. Use environment variables or a `.env` file in production.

---

## Project Structure

```
├── jira_all_queries.py       # Step 1: Extract data from Jira API
├── visualization_excel_v5.py # Step 2: Transform + generate Excel dashboard
├── index.html                # Interactive portfolio page (live preview)
└── README.md
```

---

*Built by [Saki Chen](https://www.linkedin.com/in/sakichen/) — Data Analyst | Python · SQL · Jira API · Automation*
# Jira-Sprint-analytics
