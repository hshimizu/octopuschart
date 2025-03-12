var pricingData = null;
var dateStrFromTwoDaysAgo = null;
var totalCost = null;

function formatDate(inputDate) {
  // Parse the date string into a Date object
  var dateObj = new Date(inputDate);
  
  // Get the day, month, and year
  var day = dateObj.getDate();
  var month = dateObj.toLocaleString('default', { month: 'long' }); // Get the full month name
  var year = dateObj.getFullYear();
  
  // Return the formatted date without ordinal suffix
  return day + ' ' + month + ' ' + year;
}

function fetchOctopusData() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Electricity Data");
  // If the sheet doesn't exist, create it
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Electricity Data");
  }


  // Fetch secrets
  var secretssheet = spreadsheet.getSheetByName("Secrets");
  if (!secretssheet) {
    sheet = spreadsheet.insertSheet("Secrets");
  }
  if (secretssheet.getRange("A1").getValue() == "" || secretssheet.getRange("A1").getValue() == null) {
    secretssheet.getRange("A1").setValue("Octopus API Key");
  }
  if (secretssheet.getRange("A2").getValue() == "" || secretssheet.getRange("A2").getValue() == null) {
    secretssheet.getRange("A2").setValue("MPAN");
  }
  if (secretssheet.getRange("A3").getValue() == "" || secretssheet.getRange("A3").getValue() == null) {
    secretssheet.getRange("A3").setValue("Meter Serial");
  }
  if (secretssheet.getRange("A4").getValue() == "" || secretssheet.getRange("A4").getValue() == null) {
    secretssheet.getRange("A4").setValue("Product");
  }
  if (secretssheet.getRange("A5").getValue() == "" || secretssheet.getRange("A5").getValue() == null) {
    secretssheet.getRange("A5").setValue("Tariff");
  }
  if (secretssheet.getRange("A6").getValue() == "" || secretssheet.getRange("A6").getValue() == null) {
    secretssheet.getRange("A6").setValue("Email");
  }
  var apiKey = secretssheet.getRange("B1").getValue();
  var mpan = secretssheet.getRange("B2").getValue();
  var serial = secretssheet.getRange("B3").getValue();
  var product = secretssheet.getRange("B4").getValue();
  var tariff = secretssheet.getRange("B5").getValue();

  var exit = false;
  if (apiKey == "" || apiKey == null) {
    Logger.log("API Key cell empty.");
    exit = true;
  }
  if (mpan == "" || mpan == null) {
    Logger.log("MPAN cell empty.");
    exit = true;
  }
  if (serial == "" || serial == null) {
    Logger.log("Meter Serial cell empty.");
    exit = true;
  }
  if (product == "" || product == null) {
    Logger.log("Product cell empty.");
    exit = true;
  }
  if (tariff == "" || tariff == null) {
    Logger.log("Tariff cell empty.");
    exit = true;
  }
  if (exit == true) {
    throw new Error("Set Octopus values in Secrets sheet.");
  }


  // Fetch values from Octopus
  var range = sheet.getRange("B:B");
  range.setNumberFormat("£#,##0.00");
  var now = new Date();
  now.setDate(now.getDate() - 2); // Get date 2 days ago
  var dateStr = now.toISOString().split("T")[0];
  dateStrFromTwoDaysAgo = dateStr;

  var start = dateStr + "T00:00:00Z";
  var end = dateStr + "T23:59:59Z";

  var url = `https://api.octopus.energy/v1/electricity-meter-points/${mpan}/meters/${serial}/consumption/?period_from=${start}&period_to=${end}&order_by=period`;

  var options = {
    "headers": { "Authorization": "Basic " + Utilities.base64Encode(apiKey + ":") }
  };

  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());

  // Fetch dynamic pricing for the same day
  var pricingUrl = `https://api.octopus.energy/v1/products/${product}/electricity-tariffs/${tariff}/standard-unit-rates/?period_from=${start}&period_to=${end}`;
  var pricingResponse = UrlFetchApp.fetch(pricingUrl, options);
  var pricingData = JSON.parse(pricingResponse.getContentText());

  // Clear old data
  sheet.clear();
  sheet.appendRow(["Timestamp", "Cost (£)", "Consumption (kWh)", "Price / kWh"]);

  // Create a map of time periods to prices
  var priceMap = {};
  pricingData.results.forEach(priceEntry => {
    priceMap[priceEntry.valid_from] = priceEntry.value_inc_vat; // Store price (including VAT)
  });

  data.results.forEach(entry => {
    var consumption = entry.consumption;
    var timestamp = entry.interval_start;

    var dateObj = new Date(timestamp);
    var formattedTime = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "HH:mm");
    
    // Find the closest price for this timestamp
    var closestPrice = getClosestPrice(timestamp, priceMap);
    var cost = consumption * closestPrice / 100;  // Calculate cost for the half-hour period
    totalCost = totalCost + cost;
    sheet.appendRow([formattedTime, cost, consumption, closestPrice/100]);
  });

  sheet.appendRow(["Total", "=sum(B2:B49)", "=sum(C2:C49)"])
  Logger.log("Data updated successfully");
}

// Helper function to get the closest price for a timestamp
function getClosestPrice(timestamp, priceMap) {
  var closestPrice = 0;
  var closestTimeDiff = Infinity;
  for (var time in priceMap) {
    var timeDiff = Math.abs(new Date(timestamp) - new Date(time));
    if (timeDiff < closestTimeDiff) {
      closestPrice = priceMap[time];
      closestTimeDiff = timeDiff;
    }
  }
  return closestPrice;
}

function createChart() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Electricity Data");
  var lastRow = 49; //sheet.getLastRow();

  // Clear previous charts
  var charts = sheet.getCharts();
  charts.forEach(chart => sheet.removeChart(chart));

  // Get the script's time zone
  var timeZone = Session.getScriptTimeZone();

  // Get the date for the chart title
  var dateValue = sheet.getRange("A2").getValue();
  var formattedDate = Utilities.formatDate(new Date(dateValue), timeZone, "yyyy-MM-dd");

  // **Overwrite column A with only HH:mm values**
  var timeRange = sheet.getRange("A2:A" + lastRow);
  var timeValues = timeRange.getValues();

  for (var i = 0; i < timeValues.length; i++) {
    if (timeValues[i][0] instanceof Date && !isNaN(timeValues[i][0])) {
      let dateObj = new Date(timeValues[i][0]);
      timeValues[i][0] = new Date(1970, 0, 1, dateObj.getHours(), dateObj.getMinutes(), 0); // Strip date, keep time
    }
  }
  timeRange.setValues(timeValues);
  timeRange.setNumberFormat("HH:mm"); // Force HH:mm format in cells

  // Define the data range
  var dataRange = sheet.getRange("A1:B" + lastRow);

  // Create the chart
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(dataRange)
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'Electricity Consumption and Cost for ' + formatDate(dateStrFromTwoDaysAgo))
    .setOption('hAxis', {
      title: 'Time',
      format: 'HH:mm',
      gridlines: { count: 12 }
    })
    .setOption('vAxis', { title: 'Value' })
    .setOption('series', {
      0: { targetAxisIndex: 0, color: 'green', labelInLegend: 'Cost' }
    })
    .setOption('vAxes', {
      0: { title: 'Cost (£)', format: '##0.00' }
    })
    .setOption('legend', { position: 'top' })
    .setOption('width', 1200)
    .setOption('height', 800)
    .build();

  // Insert the chart
  sheet.insertChart(chart);
}


function sendEmailWithChart() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Electricity Data");
  var chart = sheet.getCharts()[0]; // Get the first chart
  var secretssheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Secrets");

  var email = secretssheet.getRange("B6").getValue();;
  var subject = "Electricity Dynamic Cost Report (24h, Two Days Ago):" + " £" + totalCost;
  var body = "Attached is your electricity usage and dynamic cost report.";

  if (chart) {
    var blob = chart.getAs('image/png');
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      attachments: [blob]
    });
    Logger.log("Email sent with chart.");
  } else {
    Logger.log("No chart found.");
  }
}

function scheduleFetchAndEmail() {
  fetchOctopusData();
  createChart();
  sendEmailWithChart();
}
