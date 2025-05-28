
const puppeteer = require('puppeteer'); 
const fs = require('fs');
const Excel = require('exceljs');

async function fetchCveTitle(page, url) {
  try {
    await page.goto(url, { waitUntil: 'networkidle0', timeout: 20000 });
    await page.waitForSelector('h1.ms-fontWeight-semibold', { timeout: 10000 });
    const title = await page.$eval('h1.ms-fontWeight-semibold', el => el.innerText.trim());
    return title;
  } catch (error) {
    console.warn(`Could not fetch title for ${url}`);
    return 'Unknown';
  }
}


(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    slowMo: 100,
    defaultViewport: null,
  });

  const page = await browser.newPage();

  try {
    console.log('Navigating to MSRC...');
    await page.goto('https://msrc.microsoft.com/update-guide', {
      waitUntil: 'networkidle2',
      timeout: 60000,
    });

    // Accept cookie popup
    console.log('Checking for cookie consent...');
    const buttons = await page.$$('button');
    for (const btn of buttons) {
      const text = await page.evaluate(el => el.textContent.trim(), btn);
      if (text === 'Accept') {
        await btn.click();
        console.log('Cookie popup dismissed');
        await new Promise(res => setTimeout(res, 1000));
        break;
      }
    }

    // Product Family = Windows
    console.log('Opening "Product Family" filter...');
    await page.evaluate(() => {
      const spans = Array.from(document.querySelectorAll('span'));
      const target = spans.find(el => el.textContent.trim() === 'Product Family');
      if (target) target.click();
    });
    await page.waitForFunction(() => {
      return Array.from(document.querySelectorAll('span.ms-ContextualMenu-itemText'))
        .some(span => span.textContent.trim() === 'Windows');
    }, { timeout: 10000 });

    console.log('Selecting "Windows" under Product Family...');
    const selectedWindowsLabel = await page.evaluate(() => {
      const items = Array.from(document.querySelectorAll('span.ms-ContextualMenu-itemText'));
      const match = items.find(span => span.textContent.trim() === 'Windows');
      if (match) {
        match.click();
        return match.textContent.trim();
      }
      return null;
    });
    console.log('Found Product Family label:', selectedWindowsLabel);
    await new Promise(res => setTimeout(res, 1500));

    // Product = Windows Server 2016 
    console.log('Opening "Product" filter...');
    await page.evaluate(() => {
      const spans = Array.from(document.querySelectorAll('span'));
      const target = spans.find(el => el.textContent.trim() === 'Product');
      if (target) target.click();
    });
    await page.waitForFunction(() => {
      return Array.from(document.querySelectorAll('span.ms-ContextualMenu-itemText'))
        .some(span => span.textContent.trim() === 'Windows Server 2016');
    }, { timeout: 10000 });

    console.log('Selecting "Windows Server 2016"...');
    const clickedLabel = await page.evaluate(() => {
      const items = Array.from(document.querySelectorAll('span.ms-ContextualMenu-itemText'));
      const match = items.find(span => span.textContent.trim() === 'Windows Server 2016');
      if (match) {
        match.click();
        return match.textContent.trim();
      }
      return null;
    });
    console.log('Found Product label:', clickedLabel);

    // Force dropdown to close
    await page.mouse.click(100, 100);
    await new Promise(res => setTimeout(res, 3000));

    // Wait for Fluent UI data grid
    console.log('Waiting for result rows...');
    await page.waitForSelector('div[role="rowgroup"] div[role="row"]', { timeout: 20000 });

    // Extract data from Fluent UI rows
    const data = await page.evaluate(() => {
      const rows = Array.from(document.querySelectorAll('div[role="rowgroup"] div[role="row"]'));
      return rows.map(row => {
        const cells = row.querySelectorAll('div[role="gridcell"]');
        return {
          details: cells[8]?.innerText.trim() || '',
          date: cells[0]?.innerText.trim() || ''          
        };
      });
    });

    console.log(`Extracted ${data.length} rows`);


    const detailPage = await browser.newPage();

    for (let row of data) {
      const cve = row.details.trim();
      const isCVE = /^CVE-\d{4}-\d{4,}$/.test(cve);
      if (isCVE) {
        const url = `https://msrc.microsoft.com/update-guide/vulnerability/${cve}`;
        row.title = await fetchCveTitle(detailPage, url);
      } else {
        row.title = 'Unknown';
      }
    }

    await detailPage.close();


    // Save as Excel with hyperlink for CVEs
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('MSRC');
    sheet.columns = [
      { header: 'Details', key: 'details', width: 40 },
      { header: 'Date', key: 'date', width: 20 },
      { header: 'Title', key: 'title', width: 60 }
    ];

    data.forEach(row => {
      const isCVE = /^CVE-\d{4}-\d{4,}$/.test(row.details);
      sheet.addRow({
        date: row.date,
        details: isCVE
          ? { text: row.details, hyperlink: `https://msrc.microsoft.com/update-guide/vulnerability/${row.details}` }
          : row.details,
        title: row.title || ''
      });
    });

    await workbook.xlsx.writeFile('msrc_windows_server_2016.xlsx');
    console.log('Excel saved: msrc_windows_server_2016.xlsx');

  } catch (err) {
    console.error('Error occurred:', err);
    await page.screenshot({ path: 'error-state.png' });
    console.log('📸 Screenshot saved as error-state.png');
  } finally {
    await browser.close();
  }
})();
