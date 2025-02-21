function getPageSpeedData() {
  const sheetUrl =
    "https://docs.google.com/spreadsheets/d/1K0Uz-x6qnAS4HtvamizZ_tFQh2fpe7ohOfJh9WiPE3w/edit?gid=1697498024";
  const spreadsheet = SpreadsheetApp.openByUrl(sheetUrl);

  const urls = [
    { url: "https://boroux.com/", name: "Homepage" },
    // Add more URLs here as needed
  ];
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("PAGE_SPEED_API_KEY");

  const validStrategies = ["mobile", "desktop"];

  // Fetch data for each URL and strategy
  urls.forEach((site) => {
    validStrategies.forEach((strategy) => {
      const timestamp = new Date(); // Timestamp for each request
      const sheetName = `${site.name} ${strategy}`;
      let sheet = spreadsheet.getSheetByName(sheetName);

      // Create sheet if it doesn't exist
      if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
        const headers = [
          "Timestamp",
          "Performance",
          "Accessibility",
          "Best Practices",
          "SEO",
          "FCP (ms)",
          "LCP (ms)",
          "TBT (ms)",
          "CLS",
          "Fetch Time",
        ];
        sheet.appendRow(headers);
      }

      // Fetch PageSpeed data
      const apiUrl = `https://www.googleapis.com/pagespeedonline/v5/runPagespeed?url=${encodeURIComponent(
        site.url
      )}&key=${apiKey}&strategy=${encodeURIComponent(strategy)}`;
      const response = fetchWithRetry(apiUrl);

      if (!response) {
        console.error(`❌ API Request failed for ${site.url} (${strategy})`);
        sheet.appendRow([
          timestamp,
          "API Error",
          "API Error",
          "API Error",
          "API Error",
          "API Error",
          "API Error",
          "API Error",
          "API Error",
          "N/A",
        ]);
        return;
      }

      const data = JSON.parse(response.getContentText());

      // Check if Lighthouse data is available
      if (!data.lighthouseResult) {
        console.warn(
          `⚠️ API returned field data (CrUX) instead of Lighthouse data for ${site.url} (${strategy})`
        );
        sheet.appendRow([
          timestamp,
          "CrUX Data Only",
          "CrUX Data Only",
          "CrUX Data Only",
          "CrUX Data Only",
          "CrUX Data Only",
          "CrUX Data Only",
          "CrUX Data Only",
          "CrUX Data Only",
          "N/A",
        ]);
        return;
      }

      // Extract metrics
      const categories = data.lighthouseResult.categories || {};
      const audits = data.lighthouseResult.audits || {};

      const performance = categories.performance
        ? categories.performance.score * 100
        : "N/A";
      const accessibility = categories.accessibility
        ? categories.accessibility.score * 100
        : "N/A";
      const bestPractices = categories["best-practices"]
        ? categories["best-practices"].score * 100
        : "N/A";
      const seo = categories.seo ? categories.seo.score * 100 : "N/A";

      const fcp = audits["first-contentful-paint"]
        ? audits["first-contentful-paint"].numericValue
        : "N/A";
      const lcp = audits["largest-contentful-paint"]
        ? audits["largest-contentful-paint"].numericValue
        : "N/A";
      const tbt = audits["total-blocking-time"]
        ? audits["total-blocking-time"].numericValue
        : "N/A";
      const cls = audits["cumulative-layout-shift"]
        ? audits["cumulative-layout-shift"].numericValue
        : "N/A";

      const fetchTime = data.lighthouseResult.fetchTime || "N/A";

      // Prepare row data for the sheet
      const rowData = [
        timestamp,
        performance,
        accessibility,
        bestPractices,
        seo,
        fcp,
        lcp,
        tbt,
        cls,
        fetchTime,
      ];

      // Append data to the sheet
      sheet.appendRow(rowData);
    });
  });
}

// Helper function to fetch data with retries
function fetchWithRetry(url, retries = 3, delay = 1000) {
  for (let i = 0; i < retries; i++) {
    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (response.getResponseCode() === 200) {
        return response;
      }
    } catch (error) {
      console.warn(`Attempt ${i + 1} failed: ${error}`);
    }
    Utilities.sleep(delay);
  }
  return null;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 PageSpeed Tools")
    .addItem("Run PageSpeed Test", "getPageSpeedData")
    .addToUi();
}
