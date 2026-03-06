import requests
from requests.auth import HTTPBasicAuth
import pandas as pd
import urllib3
import time

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ================= CONFIG =================
JIRA_URL   = "https://agilerocks.atlassian.net"
EMAIL      = "your_email_here"
API_TOKEN  = "your_api_token_here"
JIRA_LIMIT = 1000

auth = HTTPBasicAuth(EMAIL, API_TOKEN)
headers = {
    "Accept": "application/json"
}

# ================= FIELDS TO EXTRACT =================
FIELDS = [
    "summary", "status", "issuetype", "created",
    "assignee", "sprint", "story_points", "timespent",
    "customfield_10016",   # Story Points
    "customfield_10020",   # Sprint
    "customfield_10034",   # Sigma Time Spent (hrs) — adjust if needed
]

# ================= JQL QUERIES =================
# 7 queries split by sprint range to bypass Jira Cloud 1,000-record limit
JQL_QUERIES = {
    "Q1_T1_T10": """
        project = CS122
        AND Sprint in ("Sprint CS122 C2 T1","Sprint CS122 C2 T2","Sprint CS122 C2 T3",
                       "Sprint CS122 C2 T4","Sprint CS122 C2 T5","Sprint CS122 C2 T6",
                       "Sprint CS122 C2 T7","Sprint CS122 C2 T8","Sprint CS122 C2 T9",
                       "Sprint CS122 C2 T10")
        ORDER BY Sprint ASC
    """,
    "Q2_T11_T12": """
        project = CS122
        AND Sprint in ("Sprint CS122 C2 T11","Sprint CS122 C2 T12")
        ORDER BY Sprint ASC
    """,
    "Q3_T35_T37": """
        project = CS122
        AND Sprint in ("Sprint CS122 C2 T35","Sprint CS122 C2 T36","Sprint CS122 C2 T37")
        ORDER BY Sprint ASC
    """,
    "Q4_T38_T40": """
        project = CS122
        AND Sprint in ("Sprint CS122 C2 T38","Sprint CS122 C2 T39","Sprint CS122 C2 T40")
        ORDER BY Sprint ASC
    """,
    "Q5_T41_T43": """
        project = CS122
        AND Sprint in ("Sprint CS122 C2 T41","Sprint CS122 C2 T42","Sprint CS122 C2 T43")
        ORDER BY Sprint ASC
    """,
    "Q6_T44_T46": """
        project = CS122
        AND Sprint in ("Sprint CS122 C2 T44","Sprint CS122 C2 T45","Sprint CS122 C2 T46")
        ORDER BY Sprint ASC
    """,
    "Q7_ALL_TYPES": """
        project = CS122
        AND issuetype in (Story, Bug, Task, "Sprint Detail")
        AND Sprint in openSprints()
        ORDER BY Sprint ASC
    """,
}

# ================= FETCH FUNCTION =================
def fetch_issues(query_name, jql):
    """
    Fetches all issues for a given JQL using nextPageToken pagination
    to bypass Jira Cloud's 1,000-record limit.
    """
    url = f"{JIRA_URL}/rest/api/3/search/jql"
    all_issues      = []
    next_page_token = None

    print(f"\n[{query_name}] Starting fetch...")

    while True:
        params = {
            "jql":        jql.strip(),
            "maxResults": 100,
            "fields":     ",".join(FIELDS),
        }
        if next_page_token:
            params["nextPageToken"] = next_page_token

        response = requests.get(
            url,
            headers=headers,
            params=params,
            auth=auth,
            verify=False
        )

        if response.status_code != 200:
            print(f"  ✗ Error {response.status_code}: {response.text[:300]}")
            break

        data   = response.json()
        issues = data.get("issues", [])
        all_issues.extend(issues)
        print(f"  Fetched {len(issues)} records (Total so far: {len(all_issues)})")

        next_page_token = data.get("nextPageToken")
        if not next_page_token:
            break

        time.sleep(0.3)  # Rate limit buffer

    print(f"  ✓ [{query_name}] Complete — {len(all_issues)} total records")
    return all_issues

# ================= TRANSFORM FUNCTION =================
def transform_issues(issues):
    """
    Flattens raw Jira API response into a clean DataFrame.
    Extracts sprint name, story points, assignee, and time spent.
    """
    if not issues:
        return pd.DataFrame()

    rows = []
    for issue in issues:
        fields = issue.get("fields", {})

        # Assignee
        assignee = fields.get("assignee") or {}
        assignee_name = assignee.get("displayName")

        # Sprint (customfield_10020 — array, take last active)
        sprint_field = fields.get("customfield_10020") or []
        sprint_name  = None
        if isinstance(sprint_field, list) and sprint_field:
            sprint_name = sprint_field[-1].get("name")
        elif isinstance(sprint_field, dict):
            sprint_name = sprint_field.get("name")

        # Story Points (customfield_10016)
        story_points = fields.get("customfield_10016")

        # Sigma Time Spent (hrs) — customfield_10034 or timespent (seconds → hrs)
        sigma_hrs = fields.get("customfield_10034")
        if sigma_hrs is None:
            time_spent_sec = fields.get("timespent")
            sigma_hrs = round(time_spent_sec / 3600, 2) if time_spent_sec else None

        # SR Number (customfield — adjust key as needed)
        sr_number = fields.get("customfield_10100")

        rows.append({
            "Issue Key":              issue.get("key"),
            "Issue ID":               issue.get("id"),
            "Issue Type":             (fields.get("issuetype") or {}).get("name"),
            "SR Number":              sr_number,
            "Summary":                fields.get("summary"),
            "Status":                 (fields.get("status") or {}).get("name"),
            "Created":                fields.get("created"),
            "Assignee":               assignee_name,
            "Sprint":                 sprint_name,
            "Story Points":           story_points,
            "Sigma Time Spent (hrs)": sigma_hrs,
        })

    return pd.DataFrame(rows)

# ================= MAIN EXECUTION =================
OUTPUT_FILE = "jira_query_all.xlsx"

all_dfs = []

for query_name, jql in JQL_QUERIES.items():
    issues = fetch_issues(query_name, jql)
    df     = transform_issues(issues)
    if not df.empty:
        df["Source Query"] = query_name
        all_dfs.append(df)

# Combine all queries into one sheet
if all_dfs:
    df_all = pd.concat(all_dfs, ignore_index=True)
    df_all = df_all.drop_duplicates(subset=["Issue Key"])  # Deduplicate across queries

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_all.to_excel(writer, sheet_name="All_Queries", index=False)
        for query_name, df_q in zip(JQL_QUERIES.keys(), all_dfs):
            sheet_name = query_name[:31]  # Excel sheet name max 31 chars
            df_q.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n{'='*50}")
    print(f"  Export completed → {OUTPUT_FILE}")
    print(f"  Total unique issues: {len(df_all)}")
    print(f"  Queries run:         {len(all_dfs)}")
    print(f"{'='*50}")
else:
    print("\n  No data retrieved. Please check your JQL and credentials.")
