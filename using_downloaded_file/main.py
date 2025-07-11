import pandas as pd
import requests
import xml.etree.ElementTree as ET

input_file = r"windows_update_scraper\using_downloaded_file\Security Updates 2025-07-11-093335am.xlsx"
output_file = "filtered_updates.xlsx"
rss_url = "https://api.msrc.microsoft.com/update-guide/rss"

# Step 1: Parse RSS feed and build CVE → title dictionary
def get_cve_title_map(rss_url):
    resp = requests.get(rss_url, timeout=10)
    resp.raise_for_status()
    root = ET.fromstring(resp.content)
    cve_title = {}
    for item in root.findall(".//item"):
        guid = item.findtext("guid")
        title = item.findtext("title")
        if guid and title:
            # Remove CVE-xxxx-xxxx from the title to get only the description
            if title.startswith(guid):
                description = title[len(guid):].strip()
            else:
                description = title
            cve_title[guid] = description
    return cve_title

cve_title_map = get_cve_title_map(rss_url)

# Step 2: Read Excel and add Product column
df = pd.read_excel(input_file)
filtered_df = df[["Details", "Release date"]].copy()

# Step 3: Map Product column using the CVE number
filtered_df["Product"] = [
    cve_title_map.get(cve, "") if isinstance(cve, str) and cve.startswith("CVE-") else ""
    for cve in filtered_df["Details"]
]

# Step 4: Write to Excel with clickable CVE links
base_url = "https://msrc.microsoft.com/update-guide/vulnerability/"

with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    filtered_df.to_excel(writer, index=False, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']

    for row_num, value in enumerate(filtered_df["Details"], start=1):
        if pd.notna(value) and isinstance(value, str) and value.startswith("CVE-"):
            url = base_url + value
            worksheet.write_url(row_num, 0, url, string=value)
        else:
            worksheet.write(row_num, 0, value)

print(f"Saved filtered columns with clickable CVE links and Product titles to {output_file}")