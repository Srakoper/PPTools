/*
 * @return: Date; date object for last day of previous month
 */
function getPrevMonth() {
  var previous = new Date();
  previous.setDate(1);
  previous.setHours(-1);
  return previous;
}

/**
 * Fetches OP from campaign name(s) by iterating over campaigns in campaign iterator.
 * OP format: OP + 7 digits (length: 9).
 * Requires OP to be present at the begining of campaign name(s).
 * If no OP is found, 'N/A' is returned.
 * @param campaignIterator: campaign iterator to check campaign names in
 * @return: str; OP or 'N/A' if no OP found
 */
function getOP(campaignIterator) {
  while (campaignIterator.hasNext()) {
    var OP = campaignIterator.next().getName().substring(0, 9);
    if (OP.substring(0,2) == "OP") return OP;
  }
  return "N/A";
}

/*
 * Gets number of impressions, clicks, CTR, total cost, and CPC for ALL cmpaigns in given period.
 * @param acc: account to fetch clicks and cost from
 * @param period: valid GAdW period for fetching stats (e.g 'LAST_MONTH', 'THIS_MONTH')
 * @return: array; number of impressions, clicks, CTR, total cost, and CPC for ALL cmpaigns in given period
 */
function getImpressionsClicksCTRCostCPC(acc, period) {
  MccApp.select(acc);
  var accImpressions = acc.getStatsFor(period).getImpressions();
  var accClicks = acc.getStatsFor(period).getClicks();
  var accCTR = acc.getStatsFor(period).getCtr();
  var accCost = acc.getStatsFor(period).getCost();
  var accCPC = Math.round(accCost / accClicks * 100) / 100;
  return [accImpressions, accClicks, accCTR, accCost, accCPC]
}

/*
 * Creates a report of GAdW stats for given period and retuns it as string, CSV string, and JSON string.
 * Stats: impressions, clicks, CTR, cost, CPC
 * @param period: str; valid GAdW period for fetching stats (e.g 'LAST_MONTH', 'THIS_MONTH')
 * @param label: str; valid label for selecting accounts containing it
 * @return: array; GAdW stats as string, CSV string, and JSON string
 */
function reportStats(period, label) {
  var today = new Date();
  if (today.getHours() > 14) today.setHours(24); // sets day +1 for timezone -9 h from Ljubljana
  if (period === "THIS_MONTH") { // variables for reporting current stats
    var year = today.getFullYear();
    var month = today.getMonth() + 1;
    var day = today.getDate();
    var string = "GAdW Statistics for current month, retrieved on: " + today.toUTCString() + "\n\n";
    var csv = "";
  } else if (period === "LAST_MONTH") { // variables for reporting last montht stats
  	var previous = getPrevMonth();
    var year = previous.getFullYear();
    var month = previous.getMonth() + 1;
    var day = previous.getDate();
    var string = "GAdW Statistics for " + year + "-" + ((month >= 10) ? month : "0" + month) + ", retrieved on: " + today.toUTCString() + "\n\n";
    var csv = "";
  }
  var json = {"date": [year, month, day], "stats": {}}
  var accountIterator = MccApp.accounts().withCondition('LabelNames CONTAINS "' + label + '"').get();
  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var accName = account.getName();
    MccApp.select(account);
    var OP = getOP(AdWordsApp.campaigns().get());
    var stats = getImpressionsClicksCTRCostCPC(account, period);
    var paused = true;
      var campaignIterator = AdWordsApp.campaigns().get();
      while (campaignIterator.hasNext()) {
        if (campaignIterator.next().isEnabled()) {
          paused = false;
          break;
        }
      }
    // building string
    string = string + 
      		 (paused ? accName + " [PAUSED]" : accName) + "\n" +
      		 "impressions: " + stats[0] + "\n" +
             "clicks: " + stats[1] + "\n" +
             "CTR: " + Math.round(stats[2] * 10000) / 100 + " %" + "\n" +
             "cost: " + stats[3] + " €" + "\n" +
             "CPC: " + stats[4] + " €" + "\n" +
             "\n";
	// building CSV
    csv = csv + OP + ";" + accName + ";" + stats[0] + ";" + stats[1] + "\n";
    // building JSON
    if (period === "THIS_MONTH") {
      if (paused) json["stats"][OP] = ["paused"];
      else json["stats"][OP] = [stats[1], stats[2], stats[3], stats[0], stats[4]];
    } else if (period === "LAST_MONTH") {
      json["stats"][OP] = [stats[1], stats[2], stats[3], stats[0], stats[4]];
    }
  }
  return [string, csv, JSON.stringify(json)];
}

/*
 * Sends email to a specified recipient, with specified subject and message.
 * @param recipient: recipient of email
 * @param subject: subject of email
 * @param body: body of email
 * @param message: message of email (can be included in body or as attachment)
 * @param attachment: content of attachment to be sent
 * @param mimetype: mimeType of attachment to be sent
 */
function sendEmail(recipient, subject, body, message, attachment, mimetype) {
  if (attachment) MailApp.sendEmail(recipient, subject, body, {attachments:[{fileName: attachment, mimeType: mimetype, content: message}]});
  else MailApp.sendEmail(recipient, subject, message);
}

function main() {
  var today = new Date();
  if (today.getHours() > 14) today.setHours(24); // sets day +1 for timezone -9 h from Ljubljana
  if (today.getDate() === 1) { // sends reports for previous month on 1st of current month
    var previous = getPrevMonth();
  	//var reports_DMI = reportStats("LAST_MONTH", "DMI");
    //var reports_MCE = reportStats("LAST_MONTH", "MČE");
    var reports_Aktivne = reportStats("LAST_MONTH", "Aktivne");
    // sendEmail("damjan.mihelic@tsmedia.si", "GAdW Stats Report for 'DMI', Previous Month ", reports_DMI[0], reports_DMI[2], "GAdW_JSON_prev_month.txt", 'text/plain');
    // sendEmail("damjan.mihelic@tsmedia.si", "GAdW Stats Report for 'MČE', Previous Month ", reports_MCE[0], reports_MCE[2], "GAdW_JSON_prev_month.txt", 'text/plain');
    // sendEmail("maja.cebulj@tsmedia.si", "GAdW Stats Report for 'MČE', Previous Month ", reports_MCE[0], reports_MCE[2], "GAdW_JSON_prev_month.txt", 'text/plain');
    sendEmail("damjan.mihelic@tsmedia.si", "GAdW Stats Report for 'Aktivne', Previous Month ", reports_Aktivne[0], reports_Aktivne[2], "GAdW_JSON_" + previous.getFullYear() + "-" + ((previous.getMonth() + 1 >= 10) ? previous.getMonth() + 1 : "0" + String(previous.getMonth() + 1)) + ".txt", 'text/plain');
    sendEmail("maja.cebulj@tsmedia.si", "GAdW Stats Report for 'Aktivne', Previous Month ", reports_Aktivne[0], reports_Aktivne[2], "GAdW_JSON_" + previous.getFullYear() + "-" + ((previous.getMonth() + 1 >= 10) ? previous.getMonth() + 1 : "0" + String(previous.getMonth() + 1)) + ".txt", 'text/plain');
    sendEmail("damjan.mihelic@tsmedia.si", "GAdW Stats Report for 'Aktivne', Previous Month, CSV ", reports_Aktivne[0], reports_Aktivne[1], "GAdW_CSV_" + previous.getFullYear() + "-" + ((previous.getMonth() + 1 >= 10) ? previous.getMonth() + 1 : "0" + String(previous.getMonth() + 1)) + ".csv", 'text/csv');
    //sendEmail("damjan.mihelic@tsmedia.si", "String test", reports[0]);
  }
}