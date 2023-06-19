function generateReport() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = spreadsheet.getSheetByName("Instantly Campaigns Data");
  var config = spreadsheet.getSheetByName("Config");   //Fetching the sheets from the current Spreadsheet

  var startDate = config.getRange("B1").getValue();   //Fetching start date from Config sheet
  var endDate = config.getRange("B2").getValue();   //Fetching end date from Config sheet
  var apiKey = config.getRange("B3").getValue();   //Fetching apiKey from Config sheet

  var options = {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric'
  };  //Setting the date format

  var formattedStartDate = startDate.toLocaleDateString('en-US', options);
  var formattedEndDate = endDate.toLocaleDateString('en-US', options);

  var apiUrl = 'https://api.instantly.ai/api/v1/analytics/campaign/count';
  var apiUrlSumm = 'https://api.instantly.ai/api/v1/analytics/campaign/summary';

  var requestUrl = `${apiUrl}?start_date=${formattedStartDate}&end_date=${formattedEndDate}&api_key=${apiKey}`;

  var response = UrlFetchApp.fetch(requestUrl);  //Fetching responses from Instantly via API

  var responseData = JSON.parse(response.getContentText());  //Parsing the JSON requests to 'data' variable

  // Assuming responseData and responseDataSumm are arrays of objects
  for (var i = 0; i < responseData.length; i++) {
    var campaignId = responseData[i].campaign_id;

    var requestUrlSumm = `${apiUrlSumm}?start_date=${formattedStartDate}&end_date=${formattedEndDate}&api_key=${apiKey}&campaign_id=${campaignId}`;

    var responseSumm = UrlFetchApp.fetch(requestUrlSumm); //Fetching response from Instantly for each campaign via API

    var responseDataSumm = JSON.parse(responseSumm.getContentText());

    responseData[i].bounced = responseDataSumm.bounced;
    responseData[i].unsubscribed = responseDataSumm.unsubscribed;
    responseData[i].completed = responseDataSumm.completed;
  }

  var data = reorderData(responseData);

  // Clear the existing data in the sheet
  dashboard.clear();

  // Insert column headers
  var headers = ["Sl. No", "Campaign ID", "Campaign Name", "Number of Contacts", "Emails Sent", "Emails Read", "Contacts Opened Email", "Contacts Replied", "Completed", "Bounced Leads", "Unsubscribed"];
  dashboard.appendRow(headers);

  // Insert data
  for (var i = 0; i < data.length; i++) {
    var row = Object.values(data[i]);
    row.unshift(i + 1);
    dashboard.appendRow(row);
  }
}

function reorderData(jsonData) {
  var reorderedData = jsonData.map(function (obj) {
    return {
      "campaign_id": obj.campaign_id,
      "campaign_name": obj.campaign_name,
      "new_leads_contacted": obj.new_leads_contacted,
      "emails_sent": obj.emails_sent,
      "emails_read": obj.emails_read,
      "leads_read": obj.leads_read,
      "leads_replied": obj.leads_replied,
      "completed": obj.completed,
      "bounced": obj.bounced,
      "unsubscribed": obj.unsubscribed
    };
  });

  return reorderedData;
}
