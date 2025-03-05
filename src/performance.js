function getPageSpeedData() {
  const spreadsheet = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1K0Uz-x6qnAS4HtvamizZ_tFQh2fpe7ohOfJh9WiPE3w/edit?gid=1697498024"
  );

  const urls = [
    { url: "https://boroux.com/", name: "Homepage" },
    {
      url: "https://boroux.com/collections/accessories-parts",
      name: "PLP",
    },
    // Add more URLs here as needed
  ];
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("PAGE_SPEED_API_KEY");

  const validStrategies = ["mobile", "desktop"];

  // Fetch data create sheet for each URL and strategy and append a row in each sheet populated with the data
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
      let psiData = fetchDataFromPSI(strategy, site.url);

      const rowData = [timestamp, ...psiData];
      console.log(rowData);

      // Append data to the sheet
      sheet.appendRow(rowData);
    });
  });
}

function pageSpeedApiEndpointUrl(strategy, url) {
  const apiBaseUrl =
    "https://www.googleapis.com/pagespeedonline/v5/runPagespeed";
  const apikey =
    PropertiesService.getScriptProperties().getProperty("PAGE_SPEED_API_KEY");
  const apiCategories = [
    "performance",
    "accessibility",
    "best-practices",
    "seo",
  ];
  const categoryParams = apiCategories
    .map((category) => `&category=${category}`)
    .join("");
  return `${apiBaseUrl}?url=${encodeURIComponent(
    url
  )}&key=${apikey}&strategy=${strategy}${categoryParams}`;
}

function fetchDataFromPSI(strategy, url) {
  const options = { muteHttpExceptions: true };
  const pageSpeedEndpointUrl = pageSpeedApiEndpointUrl(strategy, url);
  const response = UrlFetchApp.fetch(pageSpeedEndpointUrl, options);

  if (response.getResponseCode() !== 200) {
    throw new Error(`API request failed for ${url} (${strategy})`);
  }

  const json = response.getContentText();
  const parsedJson = JSON.parse(json);
  const lighthouse = parsedJson.lighthouseResult;

  if (!lighthouse) {
    throw new Error(`No Lighthouse data found for ${url} (${strategy})`);
  }

  return [
    (performanceScore = lighthouse.categories.performance.score * 100),
    (accessibilityScore = lighthouse.categories.accessibility.score * 100),
    (bestPracticesScore = lighthouse.categories["best-practices"].score * 100),
    (seoScore = lighthouse.categories.seo.score * 100),
    (fcp = lighthouse.audits["first-contentful-paint"].displayValue),
    (lcp = lighthouse.audits["largest-contentful-paint"].displayValue),
    (tti = lighthouse.audits.interactive.displayValue),
    (cls = lighthouse.audits["cumulative-layout-shift"].displayValue),
  ];
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 PageSpeed Tools")
    .addItem("Run PageSpeed Test", "getPageSpeedData")
    .addToUi();
}
