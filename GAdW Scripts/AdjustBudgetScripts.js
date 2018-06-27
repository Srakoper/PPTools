/**
 * Array of accounts/companies to be ignored when processing accounts.
 * Modified manually.
 */
var ignore = ['Vsevid d.o.o.',
              'Hoteli Bernardin d.d.'];

/**
 * Object of monthly surpluses/deficits of clicks for GAdW accounts.
 * Deficits of clicks are represented by a negative value (e.g. -10).
 * Modified manually at the beginning of month.
 */
var surpluses =
{// surplus data for 2018-06
'OP0710307 - Impol servis, d.o.o.': 1,
'OP0716086 - Dvigala Bartol, d.o.o.': 4,
'OP0726257 - Agrokor, d.o.o.': 6,
'OP0729885 - Matjaž Zorman, s.p.': 9,
'OP0733080 - Novomat, d.o.o.': 3,
'OP0727117 - Rono senčila, d.o.o.': 9,
'OP0723074 - Granim, d.o.o.': 9,
'OP0733905 - ProFarm Košenina, d.o.o.': -12,
'OP0774728 - Interlogis, d.o.o.': 59,
'OP0737406 - Lokatrade Škofja Loka, d.o.o.': 19,
'OP0752118 - Avtosteklo, d.o.o.': 6,
'OP0752159 - Maksimum, d.o.o.': 6,
'OP0752753 - Dr. Konekt, d.o.o.': 4,
'OP0773577 - ZUKD Grad': -6,
'OP0760719 - Geodeti, d.o.o.': 25,
'OP0757821 - Orthops, d.o.o.': 2,
'OP0766681 - Slovensko združenje Taijiquan Hram CSN': 4,
'OP0761010 - HTE, d.o.o.': 3,
'OP0761613 - Pirh, d.o.o.': 10,
'OP0716593 - HM Storitve d.o.o.': 6,
'OP0716597 - KMC d.o.o.': 0,
'OP0730943 - Zip center d.o.o.': 19,
'OP0712780 - Odmera d.o.o.': 6,
'OP0712860 - Dental Design d.o.o.': 63,
'OP0748615 - Mlakar Jožko s.p.': 4,
'OP0715972 - Vsevid d.o.o.': -101,
'OP0716470 - Inventio d.o.o.': 3,
'OP0711400 - Saša d.o.o.': 34,
'OP0714615 - Converta d.o.o.': 8,
'OP0714511 - Dental Art d.o.o.': 0,
'OP0733404 - Bojan Burja s.p.': 4,
'OP0714167 - Topel dom, Zoran Šircelj s.p.': 8,
'OP0713997 - Birokrat IT d.o.o.': 6,
'OP0747107 - IZUM': 11,
'OP0728526 - Elektrostik, Bogdan Rupnik s.p.': 1,
'OP0747178 - Grajzar d.o.o.': 11,
'OP0714760 - Profit d.o.o., Ljubljana': 3,
'OP0710997 - Hernec d.o.o.': 3,
'OP0724668 - RTI d.o.o.': 18,
'OP0732703 - Tatalović d.o.o.': 18,
'OP0715538 - Očesni kirurški center dr. Pfeifer d.o.o.': 15,
'OP0710683 - Karnion d.o.o.': 10,
'OP0722693 - GM Skupina d.o.o.': 13,
'OP0749734 - Hoteli Bernardin d.d.': 14,
'OP0765539 - Termopol, d.o.o.': 5,
'OP0767222P - Prima Filtertehnika d.o.o.': 487,
'OP0768357P - Milan Matko s.p.': 446,
'OP0762351 - Superstar - Cvetana Rozman s.p.': 3,
'OP0768382P - Dinos d.d. - ZAČASNO USTAVLJENA': 500,
'OP0768553P - Srednja zdravstvena šola Celje - AKTIVACIJA ZOPET POLETI': 498,
'OP0768822 - Franc Kemperle': 5,
'OP0768826P - Stopar design, d.o.o.': 699,
'OP0769954 - Kast, d.o.o.': 2,
'OP0770034 - Primož Žunter, s.p.': 23,
'OP0770201P - Marnit d.o.o.': 274,
'OP0770606P - Avto Cerar, d.o.o.': 156,
'OP0767407P - TMD Invest d.o.o.': 467,
'OP0768999P - 7-Trade d.o.o.': 277,
'OP0770373P - Milena Žuran, s.p.': 582,
'OP0770041P - Istrabenz plini d.o.o.': 312,
'OP0770880P - Trakom, d.o.o.': 860,
'OP0772063 - Robert Repič, s.p.': 1,
'OP0771867P - Paloma PIS, d.o.o.': 87,
'OP0771875P - Baumuller Dravinja, d.o.o.': 448,
'OP0773171P - Ika, d.o.o.': 430,
'OP0774862P - Lichtenegger d.o.o.': 347,
'OP0774669P - Lassana d.d.': 488,
'OP0775384P - Ultramarin, d.o.o.': 473,
'OP0775289P - Trelleborg Slovenija, d.o.o.': 500,
'OP0775790P - Apolonija Strehar, s.p.': 366,
'OP0774854P - Vrtec Velenje': 273,
'OP0776274 - Peter Ocvirk s.p.': 0,
'OP0776132P - Tekoma Marguč, d.o.o.': 500,
'OP0775587P - Vaš dom d.o.o.': 150,
'OP0775997P - Liko Liboje d.d.': 500,
'OP0774661P - Kego d.o.o.': 448,
'OP0776282P - Marija Polak Rožič s.p.': 500,
'OP0777167P - Tomaž Luštrek, s.p.': 380,
'OP0777489P - SMM, d.o.o.': 800,
'OP0777761P - Gabrijela Ludvig-Zagožen': 122,
'OP0777500P - Alpeks, d.o.o.': 500
};

/**
 * Fetches OP from campaign name(s) by iterating over campaigns in campaign iterator.
 * OP format: OP + 7 digits (length: 9).
 * Requires OP to be present at the begining of campaign name(s).
 * If no OP is found, 'N/A' is returned.
 * @param campaignIterator: campaign iterator to check campaign names in
 * @return: str, OP or 'N/A' if no OP found
 */
function getOP(campaignIterator) {
  while (campaignIterator.hasNext()) {
      var OP = campaignIterator.next().getName().substring(0, 9);
      if (OP.substring(0,2) == "OP") return OP;
      else if (!(campaignIterator.hasNext())) return "N/A";
  }
}

/**
 * Company class for storing data of a company (account).
 * Stored data: company name/account name,
 *              OP label,
 *              package (49, 99, 199, 399),
 *              current GAdW clicks,
 *              current TSmedia clicks,
 *              monthly surplus/deficit of clicks,
 *              monthly GAdW clicks goal according to package + surplus/deficit,
 *              monthly TSmedia clicks goal according to package + surplus/deficit,
 *              monthly total clicks goal according to GAdW goal + TSmedia goal,
 *              current GAdW impressions,
 *              current GAdW cost,
 *              current GAdW CTR
 */
function Company(name, OP, package, clicksGAdW, clicksTSmedia, surplus, impressionsGAdW, cost, ctr) {
  this.packages = {49:[40,160],99:[80,320],199:[160,640],399:[320,1280]}
  this.name = name;
  this.OP = OP;
  this.package = package;
  this.clicksGAdW = clicksGAdW;
  this.clicksTSmedia = clicksTSmedia;
  this.surplus = surplus; // can be negative (deficit)
  this.goalGAdW = this.package in this.packages ? this.packages[this.package][0] : Math.round(this.package * 20) / 100; // checks if package not in packages
  this.goalTSmedia = this.package in this.packages ? this.packages[this.package][1] - this.surplus : Math.round(this.package * 80) / 100 - this.surplus; // checks if package not in packages
  this.goalTotal = this.goalGAdW + this.goalTSmedia;
  this.impressionsGAdW = impressionsGAdW;
  this.cost = cost;
  this.ctr = ctr;
  this.getName = function() {return this.name;};
  this.getOP = function() {return this.OP;};
  this.getPackage = function() {return this.package;};
  this.getClicksGAdW = function() {return this.clicksGAdW;};
  this.getClicksTSmedia = function() {return this.clicksTSmedia;};
  this.getSurplus = function() {return this.surplus;};
  this.getGoalGAdW = function() {return this.goalGAdW;};
  this.getGoalTSmedia = function() {return this.goalTSmedia;};
  this.getGoalTotal = function() {return this.goalTotal;};
  this.getImpressionsGAdW = function() {return this.impressionsGAdW;};
  this.getCost = function() {return this.cost;};
  this.getCtr = function() {return this.ctr;};
}

function getCPCFromLastDayPrevMonth(campaign) {
  return campaign.getStatsFor(getLastDayPrevMonth()).getAverageCpc();
}

/**
 * @return: date object for last day of previous month
 */
function getLastDayPrevMonth() {
  var previous = new Date();
  previous.setDate(1);
  previous.setHours(-1);
  return previous;
}

/**
 * Gets data from external source.
 * Intended to fetch TSmedia clicks data from Graphite AdServer via proxy url.
 * @param url: str; url to fetch data from.
 * @return: str; fetched data in string format
 */
function getExternalData(url) {
  var formData = {
   'email': 'damjan.mihelic@tsmedia.si',
   'password': 'L!4)ggdSkRH-N/Sqe9Qq-/379YbO4X'
  };
  var options = {
   'method': 'post',
   'payload': formData
  };
  return UrlFetchApp.fetch(url);
}

/**
 * Starts campaing with lowest CPC in previous month within account (starts SEARCH sampaign if CPCs tied).
 * @param acc: account to start campaigns within
 */
function startLowestCPCCampaigns(acc) {
  MccApp.select(acc);
  try {
    acc.removeLabel("StoppedByScript"); // removes any StoppedByScript label
  } catch (err) {}
  try {
    acc.removeLabel("PausedByScript"); // removes any PausedByScript label
  } catch (err) {}
  var campaignIterator = AdWordsApp.campaigns().get();
  var lowestCPC = Infinity;
  var lowestCmp;
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName().substring(0, 2).toLowerCase() === "op") { // ignores campaigns not starting with OP
      var lastCPC = campaign.getStatsFor("LAST_MONTH").getAverageCpc();
      if (lastCPC === 0) lastCPC = campaign.getStatsFor("ALL_TIME").getAverageCpc(); // gets all-time CPC if there is no CPC (CPC 0.0) from previous month
      if (lastCPC < lowestCPC) {
        lowestCPC = lastCPC;
        lowestCmp = campaign;
      } else if (lastCPC === lowestCPC) { // if lowest_CPC tied, only a SEARCH campaing can overwrite a previously set campaign
        if (campaign.getName().toLowerCase().search("search") > -1) {
          lowestCPC = campaign.getStatsFor("LAST_MONTH").getAverageCpc();
          lowestCmp = campaign;
        }
      }
      campaign.pause(); // pauses every campaing while iterating
    }
  }
  lowestCmp.enable(); // starts/restarts a campaign with lowest CPC
}

/*
 * Starts search campaign(s) if their CPC is lower than threshold and lower than display campaign(s) CPC, else starts display campaign(s).
 * @param acc: account to start campaigns within
 * @param threshold: float; threshold CPC value that search campaign(s) must not exceed in order to be started
 */
function startSearchIfLowCPCElseDisplay(acc, threshold) {
  MccApp.select(acc);
  removeLabels(acc, ["PausedByScript", "StoppedByScript", "GoalTotalEmailSent", "GoalTSmediaEmailSent"]);
  var lowestSearchCPC = [Infinity, ""];
  var lowestDisplayCPC = [Infinity, ""];
  var searchNotFound = true;
  var disabled = true;
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName().substring(0, 2).toLowerCase() === "op") { // ignores campaigns not starting with OP 
      if (campaign.getName().toLowerCase().search("display") > -1) { // processes display campaigns
        var lastDisplayCPC = campaign.getStatsFor("LAST_MONTH").getAverageCpc();
        if (lastDisplayCPC === 0) { // if display campaign CPC === 0, overall CPC value is fetched
          var lastDisplayCPC = campaign.getStatsFor("ALL_TIME").getAverageCpc();
        }
        if (lastDisplayCPC !== 0 && lastDisplayCPC < lowestDisplayCPC[0]) { // CPC must be non-zero
          lowestDisplayCPC = [lastDisplayCPC, campaign.getName()]
        }
      } else { // processes search campaigns; assumes search campaigns are ALL campains not containing "display"
        var lastSearchCPC = campaign.getStatsFor("LAST_MONTH").getAverageCpc();
        if (lastSearchCPC <= threshold) { // processes search campaigns with CPC <= threshold
          if (lastSearchCPC === 0 && searchNotFound) { // if search campaign CPC === 0 and another search campaign has not yet been enabled, overall CPC value is fetched
            var lastSearchCPC = campaign.getStatsFor("ALL_TIME").getAverageCpc();
          }
          if (lastSearchCPC !== 0) { // CPC must be non-zero
            campaign.enable(); // will enable ALL search campaigns with 0 < CPC <= threshold
            disabled = false;
            searchNotFound = false;
            if (lastSearchCPC < lowestSearchCPC[0]) {
              lowestSearchCPC = [lastSearchCPC, campaign.getName()];
            }
          }
        }
      }
    }
  }
  if (disabled) { // no search campaign was enabled
    if (lowestDisplayCPC[0] < lowestSearchCPC[0]) {
      var campaignIterator = AdWordsApp.campaigns().get();
      while (campaignIterator.hasNext()) {
        var campaign = campaignIterator.next();
        if (campaign.getName() === lowestDisplayCPC[1]) {
          campaign.enable(); // enables display campaign with lowest CPC (but > threshold) < lowerst search campaign CPC
          break;
        }
      }
    } else {
      var campaignIterator = AdWordsApp.campaigns().get();
      while (campaignIterator.hasNext()) {
        var campaign = campaignIterator.next();
        if (campaign.getName() === lowestSearchCPC[1]) {
          campaign.enable(); // enables search campaign with lowest CPC (but > threshold)
          break;
        }
      }
    }
  }
}

/**
 * Starts all campaings.
 * @param acc: account to start campaigns within
 */
function startAllCampaigns(acc) {
  MccApp.select(acc);
  try {
    acc.removeLabel("StoppedByScript"); // removes any StoppedByScript label
  } catch (err) {}
  try {
    acc.removeLabel("PausedByScript"); // removes any PausedByScript label
  } catch (err) {}
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName().substring(0, 2).toLowerCase() === "op") { // ignores campaigns not starting with OP
      campaign.enable();
    }
  }
}

/**
 * Starts all SEARCH campaings, pauses all other campaigns.
 * @param acc: account to start SEARCH campaigns/pause others within
 */
function startSearchCampaigns(acc) {
  MccApp.select(acc);
  try {
    acc.removeLabel("PausedByScript"); // removes any PausedByScript label
  } catch (err) {}
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName().substring(0, 2).toLowerCase() === "op") { // ignores campaigns not starting with OP
      if (campaign.getName().toLowerCase().search("search") > -1) campaign.enable();
      else campaign.pause();
    }
  }
}

/**
 * Starts all DISPLAY campaings, pauses all other campaigns.
 * @param acc: account to start DISPLAY campaigns/pause others within
 */
function startDisplayCampaigns(acc) {
  MccApp.select(acc);
  try {
    acc.removeLabel("StoppedByScript"); // removes any StoppedByScript label
  } catch (err) {}
  try {
    acc.removeLabel("PausedByScript"); // removes any PausedByScript label
  } catch (err) {}
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName().substring(0, 2).toLowerCase() === "op") { // ignores campaigns not starting with OP
      if (campaign.getName().toLowerCase().search("display") > -1) campaign.enable();
      else campaign.pause();
    }
  }
}

/**
 * Checks which campaigns in account are currently running: search, display, both or none.
 * @param acc: account to check running campaigns in
 * @return: str; "search" if search campaign(s) running, "display" if display campaign(s) running, "both" if both campaigns running, "none" if no campaign running
 */
function checkRunningCampaigns(acc) {
  MccApp.select(acc);
  var search = false;
  var display = false;
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName().substring(0, 2).toLowerCase() === "op") { // ignores campaigns not starting with OP
      if (campaign.isEnabled()) {
        if (campaign.getName().toLowerCase().search("display") > -1) {display = true;}
        else {search = true;}
      }
    }
  }
  if (search && display) {return "both";}
  else if (search) {return "search";}
  else if (display) {return "display";}
  else {return "none";}
}

/**
 * Sets budget of active campaings to a default value.
 * Values: {Poslovni49: 0.15, Poslovni99: 0.30, Poslovni199: 0.60, Poslovni399: 1.00}
 * @param acc: account to set default budget in
 * @param label: label of Poslovni paket (Poslovni49, Poslovni99, Poslovni199, Poslovni399)
 */
function setDefaultBudget(acc, label) {
  MccApp.select(acc);
  var activeCampaigns = [];
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.isEnabled()) {
      activeCampaigns.push(campaign.getName())
    }
  }
  var budgetIterator = AdWordsApp.budgets().get();
  while (budgetIterator.hasNext()) {
    var budget = budgetIterator.next();
    var budgetName = budget.getName();
    if (activeCampaigns.indexOf(budgetName) > -1) {
      Logger.log(budgetName);
      switch (label) {
        case "Poslovni 49":
          budget.setAmount(0.15);
          break;
        case "Poslovni 99":
          budget.setAmount(0.30);
          break;
        case "Poslovni 199":
          budget.setAmount(0.60);
          break;
        case "Poslovni 399":
          budget.setAmount(1.00);
          break;
      }
    }
  }
}

/**
 * Pauses all campaings.
 * @param acc: account to pause campaigns within
 */
function pauseAllCampaigns(acc) {
  MccApp.select(acc);
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    campaign = campaignIterator.next().pause();
  }
}

/**
 * Checks if at least one campaign in account is enabled.
 * @param acc: account to check campaigns within
 * @return: boolean; true if at least one campaingn in account is enabled, false otherwise
 */
function isEnabled(acc) {
  MccApp.select(acc);
  var enabled = false;
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next()
    if (campaign.isEnabled()) {
      enabled = true;
      break;
    }
  }
  return enabled;
}

/**
 * Pauses given campaign and enables alternative campaign(s).
 * If active search campaign -> switch to display campaign(s), if active display campaign -> switch to search campaign(s).
 * @param acc: account to switch campaigns within
 * @param OP: string, OP value of active campaign
 * @param activeCmpName: string, active campaign name
 * @param package: int, account package (49, 99, 199, 399)
 * @return: object; enabled and paused campaigns as key-value pairs: "enabled": [<campaign name>,...] or "paused": [<campaign name>,...]
 */
function switchCampaigns(acc, OP, activeCmpName, package) {
  MccApp.select(acc);
  var campaigns = {"enabled": [], "paused": []};
  var enabled = false;
  if (activeCmpName.toLowerCase().indexOf("display") !== -1) {
    var display = true;
  } else {
    var display = false;
  }
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.getName() !== activeCmpName) {
      if (display) { // enable all search campaigns containing OP
        if (!(campaign.getName().toLowerCase().indexOf("display") !== -1) && campaign.getName().substring(0, 9) === OP) {
          campaign.enable();
          enabled = true;
          campaigns["enabled"].push(campaign.getName());
          Logger.log("Campaign " + campaign.getName() + " enabled!");
        }
      } else { // enable all display campaigns containing OP
        if (campaign.getName().toLowerCase().indexOf("display") !== -1 && campaign.getName().substring(0, 9) === OP) {
          campaign.enable();
          enabled = true;
          campaigns["enabled"].push(campaign.getName());
          Logger.log("Campaign " + campaign.getName() + " enabled!");
        }
      }
    }
  }
  if (enabled) { // pauses active campaign iff one or more campaigns were enabled in previous step
    var campaignIterator = AdWordsApp.campaigns().get();
    while (campaignIterator.hasNext()) {
      if (campaign.getName() === activeCmpName) {
        campaign.pause(); // pause currently active campaign
        campaigns["paused"].push(campaign.getName());
        Logger.log("Campaign " + campaign.getName() + " paused!");
      }
    }
    setDefaultBudget(acc, "Poslovni " + package);
    return campaigns;
  }
  return false; // returns false if no campaigns were switched
}

/**
 * Adjusts budgets of GAdW campaigns to meet the monthly goal of clicks in GAdW.
 * @param acc: account to adjust budget in
 * @param company: company object belonging to account
 * @param owner: string; owner label
 * @param pausedByScript: string; label PausedByScript if account previously paused by adjustBudgetGAdW() or false if not
 * @param emailTotalSentByScript: string/boolean; GoalTotalEmailSent label if email alert previously sent, false otherwise
 * @param emailTSmediaSentByScript: string/boolean; GoalTSmediaEmailSent label if email alert previously sent, false otherwise
 * @param date: date object for current day
 * @return: array; returns array with account/company stats for paused or processed accounts/companies
 */
function adjustBudgetGAdW(acc, company, owner, pausedByScript, emailTotalSentByScript, emailTSmediaSentByScript, date) {
  MccApp.select(acc);
  var goalGAdW = company.getGoalGAdW();
  var goalTSmedia = company.getGoalTSmedia();
  var goalTotal = company.getGoalTotal();
  var clicksGAdW = company.getClicksGAdW();
  var clicksTSmedia = company.getClicksTSmedia();
  var clicksTotal = clicksGAdW + clicksTSmedia;
  var impressionsGAdW = company.getImpressionsGAdW();
  var cost = company.getCost();
  var ctr = company.getCtr();
  var clicksRemaining = goalTotal - clicksTotal;
  var daysRunning = new Date().getDate(); // assumes campaigns were started on 1st day of month
  var daysRemaining = getDaysRemaining();
  if (daysRemaining > 1) {daysRemaining -= 1;} // subtracts one day as if the month had one day less, unless only one day remaining
  var clicksPerDayRemaining = Math.ceil(clicksRemaining / daysRemaining);
  var clicksGAdWPerDay = clicksGAdW / daysRunning;
  var clicksGAdWProjected = Math.floor(clicksGAdWPerDay * daysRemaining);
  var clicksTSmediaPerDay = clicksTSmedia / daysRunning;
  var clicksTSmediaProjected = Math.floor(clicksTSmediaPerDay * daysRemaining);
  if (getDaysRemaining() <= 15 && (clicksTotal / goalTotal) / (daysRunning / (daysRunning + daysRemaining - 1)) < 0.75) { // send email alert if 15 or less days in month remaining and if campaign underperforming (ratio of generated total clicks / total goal to days running / total days is < 75%)
    var running = checkRunningCampaigns(acc);
    if (running === "search") {
      startDisplayCampaigns(acc);
      sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket campain underperforming", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + clicksTotal + "/" + goalTotal + "\n" + Math.round((clicksTotal / goalTotal) / (daysRunning / (daysRunning + daysRemaining - 1)) * 100) + "% performance\n\nDisplay campaign(s) activated!");
    } else if (running === "display") {
      startSearchCampaigns(acc);
      sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket campain underperforming", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + clicksTotal + "/" + goalTotal + "\n" + Math.round((clicksTotal / goalTotal) / (daysRunning / (daysRunning + daysRemaining - 1)) * 100) + "% performance\n\nSearch campaign(s) activated!");
    } else {
      sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket campain underperforming", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + clicksTotal + "/" + goalTotal + "\n" + Math.round((clicksTotal / goalTotal) / (daysRunning / (daysRunning + daysRemaining - 1)) * 100) + "% performance");
    }
    //sendEmail("maja.cebulj@tsmedia.si", "Poslovni paket campain underperforming", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + clicksTotal + "/" + goalTotal + "\n" + Math.round((clicksTotal / goalTotal) / (daysRunning / (daysRunning + daysRemaining - 1)) * 100) + "% performance");
  }
  if (company.getClicksTSmedia() >= company.getGoalTSmedia()) {
    if (!(emailTSmediaSentByScript)) { // send email alert if TSmedia clicks goal met and email not yet sent
      sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket TSmedia clicks goal reached", "", "Account: " + company.getName() + "\nTSmedia clicks/goal: " + company.getClicksTSmedia() + "/" + company.getGoalTSmedia());
      sendEmail("maja.cebulj@tsmedia.si", "Poslovni paket TSmedia clicks goal reached", "", "Account: " + company.getName() + "\nTSmedia clicks/goal: " + company.getClicksTSmedia() + "/" + company.getGoalTSmedia());
      acc.applyLabel("GoalTSmediaEmailSent");
      Logger.log("Label GoalTSmediaEmailSent applied to " + acc.getName());
    }
  }
  if (pausedByScript) { // account was previously paused by adjustBudgetGAdW()
    if (clicksGAdW >= goalGAdW && clicksTotal >= goalTotal) { // checks if campaign reached total goal while paused on GAdW
      acc.applyLabel("StoppedByScript"); // applies StoppedByScript label because total clicks goal has been met
      Logger.log("Label StoppedByScript applied to " + acc.getName());
      removeLabels(acc, ["PausedByScript"]); // removes PausedByScript label because it is replaced by StoppedByScript Label
      if (!(emailTotalSentByScript)) { // send email alert if total clicks goal reached and email not yet sent
        sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket total clicks goal reached", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + (company.getClicksGAdW() + company.getClicksTSmedia()) + "/" + company.getGoalTotal());
        sendEmail("maja.cebulj@tsmedia.si", "Poslovni paket total clicks goal reached", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + (company.getClicksGAdW() + company.getClicksTSmedia()) + "/" + company.getGoalTotal());
        acc.applyLabel("GoalTotalEmailSent");
        Logger.log("Label GoalTotalEmailSent applied to " + acc.getName());
      }
    }
    if (clicksTSmediaProjected < clicksRemaining) { // reactivate campaign with lowest CPC because projected TSmedia clicks are lower than remaining clicks and send email alert
      startLowestCPCCampaigns(acc);
      acc.removeLabel("PausedByScript"); // removes PausedByScript label because account reactivated
      sendEmail("damjan.mihelic@tsmedia.si", "GAdW Campaigns Reactivated", "", "Campaigns in account " + company.getName() + " reactivated on " + date.toDateString());
      sendEmail("maja.cebulj@tsmedia.si", "GAdW Campaigns Reactivated", "", "Campaigns in account " + company.getName() + " reactivated on " + date.toDateString());
    } else {return false;} // terminate if campaign not reactivated
  }
  if (clicksGAdW >= goalGAdW && clicksTotal >= goalTotal) { // pauses campaign that reached GAdW goal AND total goal and send email alert
    pauseAllCampaigns(acc);
    acc.applyLabel("StoppedByScript"); // applies StoppedByScript label because total clicks goal has been met
    Logger.log("Label StoppedByScript applied to " + acc.getName());
    removeLabels(acc, ["PausedByScript"]); // removes PausedByScript label because it is replaced by StoppedByScript Label
    if (!(emailTotalSentByScript)) { // send email alert if total clicks goal reached and email not yet sent
      sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket total clicks goal reached", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + (company.getClicksGAdW() + company.getClicksTSmedia()) + "/" + company.getGoalTotal());
      sendEmail("maja.cebulj@tsmedia.si", "Poslovni paket total clicks goal reached", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + (company.getClicksGAdW() + company.getClicksTSmedia()) + "/" + company.getGoalTotal());
      acc.applyLabel("GoalTotalEmailSent");
      Logger.log("Label GoalTotalEmailSent applied to " + acc.getName());
    }
    return [clicksGAdW, cost, impressionsGAdW, ctr];
  } else if (clicksGAdW >= goalGAdW && clicksTSmediaProjected >= clicksRemaining) { // pauses campaign that reached GAdW goal AND TSmedia projected clicks exceed total remaining clicks
    pauseAllCampaigns(acc);
    acc.applyLabel("PausedByScript"); // applies PausedByScript label that allows subsequent checking if clicks goal is still met
    Logger.log("Label PausedByScript applied to " + acc.getName());
    return [clicksGAdW, cost, impressionsGAdW, ctr];
  } else { // compute GAdW budget to bridge the gap between TSmedia projected clicks and total remaining clicks using GAdW clicks
    var clicksGap = clicksRemaining - clicksTSmediaProjected;
    var activeCmpsNamesCPCs = {}; // stores names and CPCs of campaigns currently active
    var activeCmpsNum = 0;
    var sumCPCs = 0; // stores sum of CPCs of campaigns currently active
    var campaignIterator = AdWordsApp.campaigns().get();
    while (campaignIterator.hasNext()) {
      var campaign = campaignIterator.next();
      if (campaign.isEnabled()) {
        var cmpCPC = Math.ceil(campaign.getStatsFor("THIS_MONTH").getAverageCpc() * 100) / 100; // CPC of campaign rounded UP to 2 decimals;
        if (cmpCPC === 0) {
          sendEmail("damjan.mihelic@tsmedia.si", "GAdW no CPC Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €0.00 -> NO CLICKS GENERATED!");
          sendEmail("maja.cebulj@tsmedia.si", "GAdW no CPC Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €0.00 -> NO CLICKS GENERATED!");
          break;
        }
        activeCmpsNamesCPCs[campaign.getName()] = cmpCPC;
        activeCmpsNum++;
        sumCPCs += cmpCPC;
        if (cmpCPC > 0.15) { // sends email alert if campaign CPC > €0.15
          var switched = switchCampaigns(acc, company.getOP(), campaign.getName(), company.getPackage());
          if (switched) {
            sendEmail("damjan.mihelic@tsmedia.si", "GAdW CPC Over €0.15 Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €" + cmpCPC + "\n\nCAMPAIGNS SWITCHED\nPaused campaign(s): " + switched["paused"] + "\nEnabled campaign(s): " + switched["enabled"]);
            sendEmail("maja.cebulj@tsmedia.si", "GAdW CPC Over €0.15 Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €" + cmpCPC + "\n\nCAMPAIGNS SWITCHED\nPaused campaign(s): " + switched["paused"] + "\nEnabled campaign(s): " + switched["enabled"]);
            //sendEmail("damjan.mihelic@tsmedia.si", "GAdW CPC > €0.15 Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €" + cmpCPC);
            //sendEmail("maja.cebulj@tsmedia.si", "GAdW CPC > €0.15 Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €" + cmpCPC);
          } else {
            sendEmail("damjan.mihelic@tsmedia.si", "GAdW CPC Over €0.15 Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €" + cmpCPC + "\n\nNO CAMPAIGNS TO SWITCH!");
            sendEmail("maja.cebulj@tsmedia.si", "GAdW CPC Over €0.15 Warning", "", "Account: " + company.getName() + "\nCampaign: " + campaign.getName() + "\nCPC: €" + cmpCPC + "\n\nNO CAMPAIGNS TO SWITCH!");
          }
        }
      }
    }
    if (switched) { // updates names and CPCs of campaigns currently active if campaigns have been switched
      activeCmpsNamesCPCs = {};
      activeCmpsNum = 0;
      sumCPCs = 0;
      var campaignIterator = AdWordsApp.campaigns().get();
      while (campaignIterator.hasNext()) {
        var campaign = campaignIterator.next();
        if (campaign.isEnabled()) {
          var cmpCPC = Math.ceil(campaign.getStatsFor("ALL_TIME").getAverageCpc() * 100) / 100; // CPC of switched campaign(s) rounded UP to 2 decimals;
          activeCmpsNamesCPCs[campaign.getName()] = cmpCPC;
          activeCmpsNum++;
          sumCPCs += cmpCPC;
        }
      }
    }
    if (activeCmpsNum === 1) { // only one active campaign
      var budgetIterator = AdWordsApp.budgets().forDateRange("YESTERDAY").get();
      while (budgetIterator.hasNext()) {
        var budget = budgetIterator.next();
        var budgetName = budget.getName();
        var budgetPrev = budget.getAmount();
        clicksGAdWProjected = Math.floor(budgetPrev / activeCmpsNamesCPCs[budgetName] * daysRemaining);
        if (clicksGAdWProjected === Infinity) {clicksGAdWProjected = 0;} // if current CPC === 0 because no clicks
        if (budgetName in activeCmpsNamesCPCs) {
          if (clicksGap <= 0 || clicksTSmedia === null || clicksTSmedia === undefined) { // there is no clicks gap or there is no TSmedia clicks data -> GAdW budget is set to meet GAdW package goal by end of month
            budget.setAmount(((goalGAdW - clicksGAdW) / daysRemaining) * activeCmpsNamesCPCs[budgetName] + 0.01);
            if (budget.getAmount() < cmpCPC) {budget.setAmount(cmpCPC + 0.01);} // sets budget equal to CPC + 0.01 if calculated budget is below CPC
            if (clicksTSmedia === undefined) {Logger.log(company.getName() + " - No TSmedia clicks data");}
          } else { // sets budget according to formula: (current campaign CPC * number of gap clicks / number of days remaining) + 1 * current campaign CPC as safety margin
            budget.setAmount((activeCmpsNamesCPCs[budgetName] * clicksGap / daysRemaining) + activeCmpsNamesCPCs[budgetName]);
            if (budget.getAmount() < cmpCPC) {budget.setAmount(cmpCPC + 0.01);} // sets budget equal to CPC + 0.01 if calculated budget is below CPC
          }
          return [budgetName,
                  clicksGAdW,
                  clicksTSmedia,
                  goalGAdW,
                  goalTSmedia,
                  goalTotal,
                  clicksGAdWProjected  + clicksGAdW,
                  clicksTSmediaProjected + clicksTSmedia,
                  cmpCPC,
                  budgetPrev,
                  budget.getAmount(),
                  (clicksGAdW + (Math.floor(budget.getAmount() / activeCmpsNamesCPCs[budgetName] * daysRemaining))),
                  (clicksGAdW + (Math.floor(budget.getAmount() / activeCmpsNamesCPCs[budgetName] * daysRemaining))) + clicksTSmediaProjected + clicksTSmedia,
                  cost,
                  Math.round(cost + budget.getAmount() * daysRemaining) * 100 / 100];
        }
      }
    } else { // more than one active campaign
      var reversedCmpsNamesCPCs = {}; // stores percentages of active campains contribution according to their CPCs (lower CPC -> larger percentage)
      for (var key in activeCmpsNamesCPCs) {
        reversedCmpsNamesCPCs[key] = (sumCPCs - activeCmpsNamesCPCs[key]) / (sumCPCs * (activeCmpsNum - 1));
      }
      var budgetsPrev = {};
      var budgetsNew = {};
      var budgetIterator = AdWordsApp.budgets().forDateRange("YESTERDAY").get();
      while (budgetIterator.hasNext()) {
        var budget = budgetIterator.next();
        var budgetName = budget.getName();
        var budgetPrev = budget.getAmount();
        budgetsPrev[budgetName] = budgetPrev;
        if (budgetName in activeCmpsNamesCPCs) {
          if (clicksGap <= 0) { // there is no clicks gap and GAdW budget is set to meet GAdW package goal by end of month
            budget.setAmount(((goalGAdW - clicksGAdW) / daysRemaining) * (sumCPCs / activeCmpsNum) * reversedCmpsNamesCPCs[budgetName] + 0.01);
            if (budget.getAmount() < cmpCPC) {budget.setAmount(cmpCPC + 0.01);} // sets budget equal to CPC + 0.01 if calculated budget is below CPC
            budgetsNew[budgetName] = budget.getAmount();
          } else { // sets budget according to formula: current campaign CPC * number of gap clicks in inverse proportion to current campaign CPC + 1 * current campaign CPC as safety margin
            budget.setAmount((clicksGap / daysRemaining) * (sumCPCs / activeCmpsNum) * reversedCmpsNamesCPCs[budgetName] + 0.01); /// CHANGED!!!
            if (budget.getAmount() < cmpCPC) {budget.setAmount(cmpCPC + 0.01);} // sets budget equal to CPC + 0.01 if calculated budget is below CPC
            budgetsNew[budgetName] = budget.getAmount();
          }
        }
      }
      var projectedGAdW = clicksGAdW;
      var projectedCost = cost;
      for (var budget in budgetsNew) {
        projectedGAdW += Math.round(budgetsNew[budget] / (sumCPCs / activeCmpsNum) * daysRemaining);
        projectedCost += budgetsNew[budget] / (sumCPCs / activeCmpsNum) * daysRemaining * activeCmpsNamesCPCs[budget];
      }
      return ["Multiple campaigns/budgets",
              clicksGAdW,
              clicksTSmedia,
              goalGAdW,
              goalTSmedia,
              goalTotal,
              clicksGAdWProjected  + clicksGAdW,
              clicksTSmediaProjected + clicksTSmedia,
              JSON.stringify(activeCmpsNamesCPCs),
              JSON.stringify(budgetsPrev),
              JSON.stringify(budgetsNew),
              projectedGAdW,
              projectedGAdW + clicksTSmediaProjected + clicksTSmedia,
              cost,
              Math.round(projectedCost * 100) / 100];
    }
  }
}

/**
 * Maximizes budgets in active campaigns of accounts by a given factor: budget = remaining total clicks * CPC * factor
 * To be used on the last day of month to increase budget if total goal not yet reached.
 * @param acc: account to adjust budget in
 * @company: company object belonging to account
 * @factor: factor to multiply budget by (default: 3)
 */
function maximizeBudget(acc, company, factor) {
  MccApp.select(acc);
  if (factor) {var f = factor;}
  else {var f = 3;}
  var goalTotal = company.getGoalTotal();
  var clicksGAdW = company.getClicksGAdW();
  var clicksTSmedia = company.getClicksTSmedia();
  var clicksTotal = clicksGAdW + clicksTSmedia;
  var clicksRemaining = goalTotal - clicksTotal;
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.isEnabled()) {
      var cmpName = campaign.getName();
      var cmpCPC = campaign.getStatsFor("LAST_7_DAYS").getAverageCpc();
      var budget = AdWordsApp.budgets().withCondition("BudgetName = " + cmpName).get();
      budget.next().setAmount(clicksRemaining * cmpCPC * f);
    }
  }
}

/**
 * Checks if total clicks exceed campaign goal (= goal met).
 * @param acc: account to check clicks against goal in
 * @param company: company object belonging to account
 * @param GAdWandTSmedia: boolean; true if applies to GAdW and TSmedia networks (total goal) or false if applies to TSmedia network only (TSmedia goal)
 * @return: boolean; true iff GAdW clicks >= GAdW goal and total clicks >= total goal, false otherwise
 */
function checkClicksAndGoal(acc, company, GAdWandTSmedia) {
  MccApp.select(acc);
  var goalGAdW = company.getGoalGAdW();
  var goalTSmedia = company.getGoalTSmedia();
  var goalTotal = company.getGoalTotal();
  var clicksGAdW = company.getClicksGAdW();
  var clicksTSmedia = company.getClicksTSmedia();
  var clicksTotal = clicksGAdW + clicksTSmedia;
  if (GAdWandTSmedia) {
    if (clicksGAdW >= goalGAdW && clicksTotal >= goalTotal) {return true;}
    else {return false;}
  } else {
    if (clicksTSmedia >= goalTSmedia) {return true;}
    else {return false;}
  }
}

/**
 * Returns the number of days in current month.
 * @return: int; number of days in current month
 */
function getDaysInMonth() {
  var today = new Date().getDate();
  var last = new Date();
  last.setMonth(last.getMonth() + 1);
  last.setDate(1);
  last.setHours(-1);
  last = last.getDate();
  return last;
}

/**
 * Returns number of remaining days in current month.
 * If remaining days is 0 (function executed on last day of month) 1 is returned instead.
 * @return: int; total number of days in month - current day in month + 1
 */
function getDaysRemaining() {
  var today = new Date().getDate();
  var last = new Date();
  last.setMonth(last.getMonth() + 1);
  last.setDate(1);
  last.setHours(-1);
  last = last.getDate();
  return last - today > 0 ? last - today : 1;
}

/**
 * Gets number of clicks for ALL campaigns in current month and total cost for ALL cmpaigns in current month.
 * @param acc: account to fetch clicks and cost from
 * @return: number of clicks for ALL campaigns in current month and total cost for ALL cmpaigns in current month
 */
function getClicksCost(acc) {
  MccApp.select(acc);
  var accClicks = acc.getStatsFor("THIS_MONTH").getClicks(); // gets clicks made in current month for ALL camapaigns
  var accCost = acc.getStatsFor("THIS_MONTH").getCost(); // gets cost accrued in current month for ALL campaigns
  return [accClicks, accCost]
}

function reportAsString() {
  var string = ""
  var accountIterator = MccApp.accounts().withCondition('LabelNames CONTAINS "DMI"').get();
  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var accName = account.getName();
    string = string + accName + "\n" + getClicksCost(account).toString();
  }
  Logger.log(string);
}

/**
 * Checks and sends email alert if click goals of custom Poslovni paketi (not in [49, 99, 199, 399]) have been met.
 * @param surplusesObject: str->int object; mapping of <OP - company name> strings to remaining clicks integers
 * @param companiesArray: array; companies objects to be compared against surplusesObject (OPs)
 * @param tsmediaDataObject: str->int object; mapping of <OP> strings to current clicks on TSmedia integers
 */
function checkCustomPoslovniPaket(surplusesObject, companiesArray, tsmediaDataObject) {
  for (var name in surplusesObject) { // entries with OP ending in "P" represent custom Poslovni paket
    if (name.substring(9, 10).toLowerCase() === "p") {
      for (var i = 0; i < tsmediaDataObject.length; i++) {
        if (name.substring(0,9) === tsmediaDataObject[i]["op"]) {
          if (tsmediaDataObject[i]["clicks"]["sum"] >= surplusesObject[name]) {
            sendEmail("damjan.mihelic@tsmedia.si", "Preusmeritve total clicks goal reached", "", "Account: " + name.substring(0, 9) + name.substring(10) + "\nTotal clicks/goal: " + tsmediaDataObject[i]["clicks"]["sum"] + "/" + surplusesObject[name]);
            sendEmail("maja.cebulj@tsmedia.si", "Preusmeritve total clicks goal reached", "", "Account: " + name.substring(0, 9) + name.substring(10) + "\nTotal clicks/goal: " + tsmediaDataObject[i]["clicks"]["sum"] + "/" + surplusesObject[name]);
          } else {
            Logger.log("Preusmeritve paket " + name.substring(0, 9) + name.substring(10) + " clicks goal not yet reached");
          }
        }
      }
    }
  }
}

/**
 * Removes specified labels from account if present.
 * @param account: account to remove labels from
 * @param labels: array of strings; labels to be removed from account
 * @return: true if at least one label removed, false otherwise
 */
function removeLabels(account, labels) {
  var removed = false;
  for (var i = 0; i < labels.length; i++) {
    try {
      account.removeLabel(labels[i]);
      removed = true;
    } catch (err) {}
  }
  return removed;
}

/**
 * Checks if at least one active campaign in account has end date defined.
 * Compares end day against given date if end day for active campaign found.
 * Returns true iff end date < given date (assumes whole account has ended), false otherwise.
 * @param account: account to check end date in
 * @param date: date object; date to compare end date against
 * @return account: boolean; true iff end date < date, false otherwise
 */
function checkEndDate(acc, date) {
  MccApp.select(acc);
  var endDate;
  var campaignIterator = AdWordsApp.campaigns().get();
  while (campaignIterator.hasNext()) {
    var campaign = campaignIterator.next();
    if (campaign.isEnabled()) {
      var end = campaign.getEndDate();
      if (end) {
        endDate = new Date(end["year"], end["month"] - 1, end["day"], 23, 59, 59);
        break;
      }
    }
  }
  if (endDate) {return endDate < date;}
  return false;
}

/**
 * Sends email to a specified recipient, with specified subject and message.
 * @param recipient: recipient of email
 * @param subject: subject of email
 * @param message: message of email
 */
function sendEmail(recipient, subject, body, message, attachment) {
  if (attachment) MailApp.sendEmail(recipient, subject, body, {attachments:[{fileName: attachment, mimeType: 'text/plain', content: message}]});
  else MailApp.sendEmail(recipient, subject, message);
}

function main() {
  var today = new Date();
  if (today.getHours() > 14) today.setHours(24); // sets day +1 for timezone -9 h from Ljubljana
  var year = today.getFullYear();
  var month = today.getMonth() + 1;
  var day = today.getDate();
  var reportPausedAsJSON = {"date": [year, month, day], "stats": {}}
  var paused = "";
  var processed = "";
  var tsmediaData = JSON.parse(getExternalData('https://atlas.siol.net/sandbox/graphite/'));
  var accountIterator = MccApp.accounts().withCondition('LabelNames CONTAINS "Aktivne"').get();
  var companies = [];
  while (accountIterator.hasNext()) {
    var account = accountIterator.next();
    var accName = account.getName();
    if (!(accName in ignore)) { // processes only accounts not in array ignore
      var accPoslovniLabel = account.labels().withCondition('Name CONTAINS "9"').get().next().getName();
      var accOwnerLabel = account.labels().withCondition('Name CONTAINS "M"').get().next().getName();
      var currentGAdWClicks = account.getStatsFor("THIS_MONTH").getClicks();
      try {
        var accPaused = account.labels().withCondition('Name = "PausedByScript"').get().next().getName();
      } catch (err) {
        var accPaused = false;
      }
      try {
        var emailTotalSent = account.labels().withCondition('Name = "GoalTotalEmailSent"').get().next().getName();
      } catch (err) {
        var emailTotalSent = false;
      }
      try {
        var emailTSmediaSent = account.labels().withCondition('Name = "GoalTSmediaEmailSent"').get().next().getName();
      } catch (err) {
        var emailTSmediaSent = false;
      }
      MccApp.select(account);
      if (today.getDate() === 1 && (!(checkEndDate(account, today)))) { // starts all paused accounts and sets default budgets on 1st of month
        startSearchIfLowCPCElseDisplay(account, 0.15);
        // startLowestCPCCampaigns(account);
        setDefaultBudget(account, accPoslovniLabel);
        removeLabels(account, ["PausedByScript", "StoppedByScript", "GoalTotalEmailSent", "GoalTSmediaEmailSent"]);
      }
      if (today.getDate() === 2) {
        if (checkEndDate(account, today)) {removeLabels(account, ["Aktivne"]);} // removes label Aktivne from account if active campaign(s) found to have end date < current date (needs to be removed on 2nd day of month otherwise monthly report from 1st day of month will not contain accounts with label "Aktivne" removed)
      }
      if (isEnabled(account) || accPaused) { // processes accounts with enabled campaign(s) or paused and having PausedByScript label
        if (isEnabled(account) && accPaused) {
          var accPaused = false;
          removeLabels(account, ["PausedByScript"]);
        }
        var campaignIterator = AdWordsApp.campaigns().get();
        var OP = getOP(campaignIterator);
        var currentTSmediaClicks;
        for (var i = 0; i < tsmediaData.length; i++) {
          var currentTSmediaClicks = 0;
          if (tsmediaData[i]["op"] === OP) {
            currentTSmediaClicks = tsmediaData[i]["clicks"]["sum"];
            break;
          }
        }
        var currentGAdWImpressions = account.getStatsFor("THIS_MONTH").getImpressions();
        var currentCost = account.getStatsFor("THIS_MONTH").getCost();
        var currentCtr = account.getStatsFor("THIS_MONTH").getCtr();
        var company = new Company(accName, OP, accPoslovniLabel.match(/\d+/)[0], currentGAdWClicks, currentTSmediaClicks, surpluses[OP + " - " + accName], currentGAdWImpressions, currentCost, currentCtr);
        companies.push(company);
        if (today.getDate() < 5) { // pauses all campaigns before 5th in month that met GAdW clicks goal AND total clicks goal, alerts if campaign paused, if TSmedia goal met
          if (checkClicksAndGoal(account, company, false)) {
            if (!(emailTSmediaSent)) { // sends email alert only if email previously not sent (label GoalTSmediaEmailSent not found)
              sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket TSmedia clicks goal reached", "", "Account: " + company.getName() + "\nTSmedia clicks/goal: " + company.getClicksTSmedia() + "/" + company.getGoalTSmedia());
              sendEmail("maja.cebulj@tsmedia.si", "Poslovni paket TSmedia clicks goal reached", "", "Account: " + company.getName() + "\nTSmedia clicks/goal: " + company.getClicksTSmedia() + "/" + company.getGoalTSmedia());
              acc.applyLabel("GoalTSmediaEmailSent");
            }
          }
          if (checkClicksAndGoal(account, company, true)) {
            pauseAllCampaigns(account);
            if (!(emailTotalSent)) { // sends email alert only if email previously not sent (label GoalTotalEmailSent not found)
              sendEmail("damjan.mihelic@tsmedia.si", "Poslovni paket total clicks goal reached", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + (company.getClicksGAdW() + company.getClicksTSmedia()) + "/" + company.getGoalTotal());
              sendEmail("maja.cebulj@tsmedia.si", "Poslovni paket total clicks goal reached", "", "Account: " + company.getName() + "\nTotal clicks/goal: " + (company.getClicksGAdW() + company.getClicksTSmedia()) + "/" + company.getGoalTotal());
              account.applyLabel("GoalTotalEmailSent");
            }
          }
        }
        if (today.getDate() >= 5) { // runs adjustBudgetGAdW() daily from 5th day in month
          var companyProcessed = adjustBudgetGAdW(account, company, accOwnerLabel, accPaused, emailTotalSent, emailTSmediaSent, today);
          if (companyProcessed) {
            if (companyProcessed.length <= 4) {
              reportPausedAsJSON["stats"][OP] = companyProcessed;
              paused += "Campaigns in account " + accName + " paused on " + today.toDateString() +
                        ".\nClicks GAdW: " + companyProcessed[0] +
                        "\nTotal cost: €" + companyProcessed[1] +
                        "\n\n";
            } else {
              processed += "Campaigns in account " + accName + " processed on " + today.toDateString() +
                           ".\nCampaign name: " + companyProcessed[0] +
                           "\nClicks GAdW: " + companyProcessed[1] +
                           "\nClicks TSmedia: " + companyProcessed[2] +
                           "\nClicks total: " + (companyProcessed[1] + companyProcessed[2]) +
                           "\nGoal GAdW: " + companyProcessed[3] +
                           "\nGoal TSmedia: " + companyProcessed[4] +
                           "\nGoal total: " + companyProcessed[5] +
                           "\nClicks GAdW projected old: " + companyProcessed[6] +
                           "\nClicks TSmedia projected old: " + companyProcessed[7] +
                           "\nClicks total projected old: " + (companyProcessed[6] + companyProcessed[7]) +
                           "\nCurrent CPC: €" + companyProcessed[8] +
                           "\nPrevious budget: €" + companyProcessed[9] + "/day" +
                           "\nNew budget: €" + companyProcessed[10] + "/day" +
                           "\nClicks GAdW projected new: " + companyProcessed[11] +
                           "\nClicks total projected new: " + companyProcessed[12] +
                           "\nTotal current cost: €" + companyProcessed[13] +
                           "\nTotal projected cost: €" + companyProcessed[14] +
                           "\n\n";
            }
          }
        }
      }
    }
  }
  if (paused || processed) {
    sendEmail("damjan.mihelic@tsmedia.si", "GAdW Budget Adjustment Report", "", paused + processed);
    sendEmail("maja.cebulj@tsmedia.si", "GAdW Budget Adjustment Report", "", paused + processed);
  }
  if (paused) {
    sendEmail("damjan.mihelic@tsmedia.si", "GAdW Campaigns Paused", paused, JSON.stringify(reportPausedAsJSON), "GAdW_JSON_paused_" + year + "-" + (month >= 10 ? month : "0" + month) + "-" + (day >= 10 ? day : "0" + day) + ".txt");
    sendEmail("maja.cebulj@tsmedia.si", "GAdW Campaigns Paused", paused, JSON.stringify(reportPausedAsJSON), "GAdW_JSON_paused_" + year + "-" + (month >= 10 ? month : "0" + month) + "-" + (day >= 10 ? day : "0" + day) + ".txt");
  }
  checkCustomPoslovniPaket(surpluses, companies, tsmediaData);
}