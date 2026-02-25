import pandas as pd
import asyncio
from playwright.async_api import async_playwright
import os
from datetime import date

# ---- CONFIG ----
input_file = r"windows_update_scraper\using_python_mrsc_file_download\Security Updates 2026-02-11-111432am.xlsx"

# Save output in the same folder as this script
output_file = os.path.join(os.path.dirname(__file__), "filtered_updates.xlsx")
base_url = "https://msrc.microsoft.com/update-guide/vulnerability/"

# ---- PHASE 1: Extract columns ----
def extract_columns(input_file):
    df = pd.read_excel(input_file)
    filtered_df = df.iloc[:, [8, 0]].copy()
    filtered_df.columns = ["Details", "Release date"]
    
    # Convert Release date to ISO format
    filtered_df["Release date"] = pd.to_datetime(filtered_df["Release date"], errors='coerce').dt.strftime('%Y-%m-%d')
    
    return filtered_df

# ---- PHASE 2: Fetch product titles and exploitability ----
async def fetch_cve_data(context, url):
    """Fetch both the product title and exploitability assessment from a CVE page."""
    try:
        page = await context.new_page()
        await page.goto(url, timeout=30000)
        try:
            await page.wait_for_selector('.ms-Spinner', state='detached', timeout=10000)
        except:
            pass
        
        title = "Unknown"
        exploitability = "Unknown"
        
        # Fetch title
        for _ in range(20):
            h1 = await page.query_selector("h1.ms-fontWeight-semibold")
            if h1:
                title = (await h1.inner_text()).strip()
                if title and "loading" not in title.lower():
                    title = title.split('\n')[0].split('<span')[0].strip()
                    break
            await asyncio.sleep(0.5)
        
        # Fetch exploitability assessment
        try:
            # Wait a bit for dynamic content to load
            await asyncio.sleep(1)
            
            # Try multiple selector strategies
            # Strategy 1: Look for text containing "Exploitability assessment"
            elements = await page.query_selector_all("div, span, td, p")
            for element in elements:
                text = await element.inner_text()
                if "Exploitability assessment" in text or "Exploitability Assessment" in text:
                    # Try to find the value in the same element or nearby
                    parent = await element.evaluate_handle("el => el.parentElement")
                    parent_text = await parent.evaluate("el => el.innerText")
                    
                    # Parse out the assessment value
                    lines = parent_text.strip().split('\n')
                    for i, line in enumerate(lines):
                        if "Exploitability assessment" in line or "Exploitability Assessment" in line:
                            # Value might be on the same line or next line
                            if i + 1 < len(lines):
                                exploitability = lines[i + 1].strip()
                            else:
                                # Try to extract from same line
                                parts = line.split("Exploitability assessment")
                                if len(parts) > 1:
                                    exploitability = parts[1].strip().strip(':').strip()
                            break
                    if exploitability != "Unknown":
                        break
            
            # Strategy 2: If not found, try looking for common patterns
            if exploitability == "Unknown":
                page_content = await page.content()
                if "Exploitation More Likely" in page_content:
                    exploitability = "Exploitation More Likely"
                elif "Exploitation Less Likely" in page_content:
                    exploitability = "Exploitation Less Likely"
                elif "Exploitation Detected" in page_content:
                    exploitability = "Exploitation Detected"
                    
        except Exception as e:
            print(f"Could not fetch exploitability for {url}: {e}")
        
        await page.close()
        return title, exploitability
    except Exception as e:
        print(f"Error fetching data for {url}: {e}")
        return "Unknown", "Unknown"

async def add_product_titles(df):
    titles = []
    exploitabilities = []
    async with async_playwright() as playwright:
        browser = await playwright.chromium.launch(headless=True)
        context = await browser.new_context()
        for cve in df["Details"]:
            if isinstance(cve, str) and cve.startswith("CVE-"):
                url = f"{base_url}{cve}"
                print(f"Fetching: {url}")
                title, exploitability = await fetch_cve_data(context, url)
                titles.append(title)
                exploitabilities.append(exploitability)
            else:
                titles.append("Unknown")
                exploitabilities.append("Unknown")
        await browser.close()
    df["Product"] = titles
    df["Exploitability assessment"] = exploitabilities
    
    # Add today's date in ISO format
    df["Today's Date"] = date.today().isoformat()
    
    # Reorder columns: Details, Release date, Today's Date, Exploitability assessment, Product
    df = df[["Details", "Release date", "Today's Date", "Exploitability assessment", "Product"]]
    
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