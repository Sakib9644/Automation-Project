const puppeteer = require('puppeteer');
const ExcelJS = require('exceljs');

(async () => {
  const browser = await puppeteer.launch({
    headless: false,
    executablePath: "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
  });
  const page = await browser.newPage();

  const excelFile = new ExcelJS.Workbook();

  const keywords = ["Dhaka", "University", "Cricket", "Bombay", "Football", "Paper", "Knife"];
  const daysOfWeek = ["Saturday", "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];

  await page.goto('https://www.google.com', { waitUntil: 'domcontentloaded' });

  for (const day of daysOfWeek) {
    const worksheet = excelFile.addWorksheet(day);

    worksheet.columns = [
      { header: 'Day', key: 'day' },
      { header: 'Keyword', key: 'keyword' },
      { header: 'Longest Suggestion', key: 'longestSuggestion' },
      { header: 'Shortest Suggestion', key: 'shortestSuggestion' },
    ];

    for (const keyword of keywords) {
      await page.waitForSelector('input[name=q]', { timeout: 60000 });
      await page.type('input[name=q]', keyword);
      await page.waitForTimeout(9000);

      await page.waitForSelector('.sbct', { timeout: 10000 });

      const suggestionElements = await page.$$('.sbct');
      const suggestions = [];
      for (const suggestionElement of suggestionElements) {
        const suggestionText = await suggestionElement.evaluate(element => element.textContent);
        suggestions.push(suggestionText);
      }

      const longestSuggestion = suggestions.reduce((a, b) => (a.length > b.length ? a : b), '');
      const shortestSuggestion = suggestions.reduce((a, b) => (a.length < b.length ? a : b), suggestions[9]);

      worksheet.addRow({ day, keyword, longestSuggestion, shortestSuggestion });

      await page.$eval('input[name=q]', input => input.value = '');
    }

    const filePath = `.//${day}.xlsx`;
    await excelFile.xlsx.writeFile(filePath);
  }

  await browser.close();
})();
