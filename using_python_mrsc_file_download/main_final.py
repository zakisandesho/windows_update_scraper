import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import os

# ---- CONFIG ----
input_file = r"windows_update_scraper\using_python_mrsc_file_download\Security Updates 2025-09-12-014816pm.xlsx"
# Save output in the same folder as this script
output_file = os.path.join(os.path.dirname(__file__), "filtered_updates.xlsx")
base_url = "https://msrc.microsoft.com/update-guide/vulnerability/"

# ---- PHASE 1: Extract columns ----
def extract_columns(input_file):
    df = pd.read_excel(input_file)
    filtered_df = df.iloc[:, [8, 0]].copy()
    filtered_df.columns = ["Details", "Release date"]
    return filtered_df

# ---- PHASE 2: Fetch product titles ----
async def fetch_title(context, url):
    try:
        page = await context.new_page()
        await page.goto(url, timeout=30000)
        try:
            await page.wait_for_selector('.ms-Spinner', state='detached', timeout=10000)
        except:
            pass
        for _ in range(20):
            h1 = await page.query_selector("h1.ms-fontWeight-semibold")
            if h1:
                title = (await h1.inner_text()).strip()
                if title and "loading" not in title.lower():
                    title = title.split('\n')[0].split('<span')[0].strip()
                    await page.close()
                    return title
            await asyncio.sleep(0.5)
        await page.close()
        return "Unknown"
    except Exception as e:
        print(f"Error fetching title for {url}: {e}")
        return "Unknown"

async def add_product_titles(df):
    titles = []
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        context = await browser.new_context()
        for cve in df["Details"]:
            if isinstance(cve, str) and cve.startswith("CVE-"):
                url = f"{base_url}{cve}"
                print(f"Fetching: {url}")
                title = await fetch_title(context, url)
                titles.append(title)
            else:
                titles.append("Unknown")
        await browser.close()
    df["Product"] = titles
    return df

# ---- PHASE 3: Write Excel with clickable links ----
def write_with_links(df, output_file):
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        worksheet = writer.sheets['Sheet1']
        for row_num, value in enumerate(df["Details"], start=1):
            if isinstance(value, str) and value.startswith("CVE-"):
                url = base_url + value
                worksheet.write_url(row_num, 0, url, string=value)
            else:
                worksheet.write(row_num, 0, value)
    print(f"Saved clickable CVE links to {output_file}")

# ---- MAIN ----
def main():
    # Phase 1
    filtered_df = extract_columns(input_file)
    # Phase 2 (async)
    filtered_df = asyncio.run(add_product_titles(filtered_df))
    # Phase 3
    write_with_links(filtered_df, output_file)

if __name__ == "__main__":
    main()