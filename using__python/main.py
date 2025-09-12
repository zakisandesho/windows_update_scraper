import asyncio
from datetime import datetime

from playwright.async_api import async_playwright
import xlsxwriter


async def fetch_title(context, url):
    try:
        page = await context.new_page()
        await page.goto(url, timeout=30000)
        for _ in range(10):
            await page.wait_for_timeout(1000)
            try:
                title = await page.text_content("h1.ms-fontWeight-semibold")
                if title and title.strip() != "Loading...":
                    await page.close()
                    return title.strip()
            except:
                continue
        await page.close()
        return "Unknown"
    except Exception as e:
        print(f"Error fetching title for {url}: {e}")
        return "Unknown"


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False, slow_mo=100)
        context = await browser.new_context()
        page = await context.new_page()

        await page.goto("https://msrc.microsoft.com/update-guide", timeout=60000)

        print("Checking for cookie popup...")
        buttons = await page.query_selector_all("button")
        for btn in buttons:
            text = (await btn.inner_text()).strip()
            if text == "Accept":
                await btn.click()
                print("Cookie popup dismissed")
                break

        # Select Product Family: Windows
        print("Selecting Product Family: Windows")
        await page.click("text=Product Family")
        await page.wait_for_selector("text=Windows")
        await page.click("text=Windows")

        # Select Product: Windows Server 2016 and matching .NET Framework versions
        print("Selecting Product: Windows Server 2016 and matching .NET Framework versions")
        await page.click("text=Product")
        await page.wait_for_selector("text=Windows Server 2016")
        await page.click("text=Windows Server 2016")

        # Wait and select all .NET Framework 4.6* products
        items = await page.query_selector_all('span.ms-ContextualMenu-itemText')
        for item in items:
            text = (await item.inner_text()).strip()
            if text.startswith("Microsoft .NET Framework"):
                print(f"✔️ Clicking: {text}")
                await item.click()

        await page.mouse.click(100, 100)  # Dismiss dropdown
        await page.wait_for_timeout(3000)

        print("Waiting for result rows...")
        await page.wait_for_selector('div[role="rowgroup"] div[role="row"]', timeout=20000)

        print("Extracting all data by scrolling...")
        results_container = await page.query_selector('.ms-DetailsList-contentWrapper')
        data = []
        seen = set()
        last_seen_count = 0
        if results_container:
            await results_container.evaluate("(el) => el.scrollTo(0, 0)")
            for scroll_num in range(50):  # Adjust as needed for more rows
                rows = await page.query_selector_all('div[role="rowgroup"] div[role="row"]')
                print(f"Scroll {scroll_num}: Found {len(rows)} rows currently visible")

                # for row in rows:
                #     cells = await row.query_selector_all('div[role="gridcell"]')
                #     if len(cells) < 9:
                #         continue
                #     date = await cells[0].inner_text()
                #     details = await cells[8].inner_text()
                #     print(f"Row: date={date.strip()}, details={details.strip()}")
                #     key = f"{date.strip()}|{details.strip()}"
                #     if key not in seen:
                #         seen.add(key)
                #         data.append({"date": date.strip(), "details": details.strip()})


                for row in rows:
                    cells = await row.query_selector_all('div[role="gridcell"]')
                    if len(cells) < 9:
                        continue
                    date = await cells[0].inner_text()
                    details = await cells[8].inner_text()
                    key = f"{date.strip()}|{details.strip()}"
                    if key not in seen:
                        seen.add(key)
                        data.append({"date": date.strip(), "details": details.strip()})
                if len(seen) == last_seen_count:
                    print("No new rows found, stopping scroll.")
                    break
                last_seen_count = len(seen)
                await results_container.evaluate("(el) => el.scrollBy(0, 1000)")
                await page.wait_for_timeout(300)
        else:
            print("❌ Could not find scroll container for results")

        print(f"Extracted {len(data)} unique rows")

        print("Fetching CVE titles...")
        for row in data:
            cve = row["details"]
            if cve.startswith("CVE-"):
                url = f"https://msrc.microsoft.com/update-guide/vulnerability/{cve}"
                row["title"] = await fetch_title(context, url)
            else:
                row["title"] = "Unknown"

        # Sort by date (parsed to datetime)
        def parse_date(row):
            try:
                return datetime.strptime(row["date"], "%b %d, %Y")
            except ValueError:
                return datetime.min

        data.sort(key=parse_date, reverse=True)

        print("Writing to Excel...")
        workbook = xlsxwriter.Workbook("msrc_windows_server_2016.xlsx")
        worksheet = workbook.add_worksheet()

        headers = ["Article", "Date", "Title"]
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)

        hyperlink_format = workbook.add_format({'color': 'blue', 'underline': 1})

        for i, row in enumerate(data, start=1):
            if row["details"].startswith("CVE-"):
                url = f"https://msrc.microsoft.com/update-guide/vulnerability/{row['details']}"
                worksheet.write_url(i, 0, url, hyperlink_format, row["details"])
            else:
                worksheet.write(i, 0, row["details"])
            worksheet.write(i, 1, row["date"])
            worksheet.write(i, 2, row["title"])

        workbook.close()
        print("Excel saved as msrc_windows_server_2016.xlsx")
        await browser.close()


if __name__ == "__main__":
    asyncio.run(main())