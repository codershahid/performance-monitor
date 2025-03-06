function getPageSpeedData() {
  const spreadsheet = SpreadsheetApp.openByUrl(
    "https://docs.google.com/spreadsheets/d/1K0Uz-x6qnAS4HtvamizZ_tFQh2fpe7ohOfJh9WiPE3w/edit?gid=1697498024"
  );

  const urls = [
    { url: "https://boroux.com/", name: "HP" },
    {
      url: "https://boroux.com/products/boroux-foundation-water-filter",
      name: "PDP",
    },
    { url: "https://boroux.com/collections/frontpage", name: "PLP" },
    // Add more URLs here as needed
  ];
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("PAGE_SPEED_API_KEY");

  const validStrategies = ["mobile", "desktop"];

  // Fetch data, create sheet for each URL and strategy, and append a row in each sheet populated with the data
  urls.forEach((site) => {
    validStrategies.forEach((strategy) => {
      const timestamp = formatTimestampUS(new Date()); // Format timestamp in US format
      const sheetName = `${site.name} ${strategy}`.substring(0, 31); // Ensure sheet name is within limit
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
          "TBT (s)",
          "CLS",
          "Fetch Time",
        ];
        sheet.appendRow(headers);
      }

      try {
        // Fetch PageSpeed data
        let psiData = fetchDataFromPSI(strategy, site.url);

        const rowData = [timestamp, ...psiData];
        console.log(rowData);

        // Append data to the sheet
        sheet.appendRow(rowData);
      } catch (error) {
        console.error(
          `Error fetching data for ${site.url} (${strategy}):`,
          error
        );
      }

      // Add a delay between requests to avoid exceeding execution time
      Utilities.sleep(5000); // 5-second delay
    });
  });
}

// Helper function to format timestamp in US format (MM/DD/YYYY HH:MM:SS)
function formatTimestampUS(date) {
  const month = String(date.getMonth() + 1).padStart(2, "0"); // Months are 0-indexed
  const day = String(date.getDate()).padStart(2, "0");
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  const seconds = String(date.getSeconds()).padStart(2, "0");

  return `${month}/${day}/${year} ${hours}:${minutes}:${seconds}`;
}

function fetchDataFromPSI(strategy, url) {
  const options = { muteHttpExceptions: true };
  const pageSpeedEndpointUrl = pageSpeedApiEndpointUrl(strategy, url);
  const response = UrlFetchApp.fetch(pageSpeedEndpointUrl, options);

  if (response.getResponseCode() !== 200) {
    const errorDetails = JSON.parse(response.getContentText());
    throw new Error(
      `API request failed for ${url} (${strategy}): ${
        errorDetails.error?.message || "Unknown error"
      }`
    );
  }

  const json = response.getContentText();
  const parsedJson = JSON.parse(json);
  const lighthouse = parsedJson.lighthouseResult;

  if (!lighthouse) {
    throw new Error(`No Lighthouse data found for ${url} (${strategy})`);
  }

  const performanceScore = lighthouse.categories.performance?.score * 100 || 0;
  const accessibilityScore =
    lighthouse.categories.accessibility?.score * 100 || 0;
  const bestPracticesScore =
    lighthouse.categories["best-practices"]?.score * 100 || 0;
  const seoScore = lighthouse.categories.seo?.score * 100 || 0;

  // Extract numeric values and convert to appropriate units
  const fcp =
    parseFloat(
      lighthouse.audits["first-contentful-paint"]?.displayValue?.replace(
        /s$/,
        ""
      )
    ) * 1000 || 0; // Convert to milliseconds
  const lcp =
    parseFloat(
      lighthouse.audits["largest-contentful-paint"]?.displayValue?.replace(
        /s$/,
        ""
      )
    ) * 1000 || 0; // Convert to milliseconds
  const tbt =
    parseFloat(
      lighthouse.audits["total-blocking-time"]?.displayValue?.replace(/s$/, "")
    ) || 0; // Already in milliseconds
  const cls = lighthouse.audits["cumulative-layout-shift"]?.numericValue || 0;

  return [
    performanceScore,
    accessibilityScore,
    bestPracticesScore,
    seoScore,
    fcp,
    lcp,
    tbt,
    cls,
  ];
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

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🚀 PageSpeed Tools")
    .addItem("Run PageSpeed Test", "getPageSpeedData")
    .addToUi();
}
