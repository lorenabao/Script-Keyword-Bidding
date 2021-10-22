//------------------------------------------------
// Report on Budget vs Spend (Hybrid)
// Created by: Remko van der Zwaag & PDDS
// remkovanderzwaag.nl & pdds.nl
// More info: http://goo.gl/d0KVTd
// CHANGELOG
// 12-06-2014: Combined two separate scripts into one Hybrid version (for MCC and accounts)
// 12-02-2015: Added column Spend yesterday and empty columns with days
// 18-06-2015: Added logic to only check campaigns with a specific label
// 03-02-2016: Will check for labelled adgroups in ignored campaigns
// 08-02-2016: Added label functionality to shopping campaigns
// 24-02-2016: Added extra fields and fixed up table layout
// 24-02-2016: Added option to switch label selection form OR to AND
// 04-04-2017: Changes converted clicks to conversions due to sunset
//------------------------------------------------

var spreadsheetId = '';// Google Spreadsheet with account info
var prefillSpreadsheet = false;                                    // When set to true, gets all accounts from the MCC account
                                                                   // and automagically adds their name and id to the spreadsheet
                                                                   // Use once, doesn't check for existing records
                                                                   // switch back to false after use
                                                                   // PREFERABLY RUN USING PREVIEW (true), CHANGE TO false AND SAVE

var emailAddr = "";                        // Where you want the report sent, only works when
var emailSubject = "INFORME DE GASTOS GOOGLE ADS LUNES";                         // What the email should be titled
var onlyReportProblems = false;                                     // Email will only contain accounts with errors
var ignoreNoBudgetCampaigns = true;                                // Never show campaigns with a 0 budget
var alwaysReport = true;                                          // Send an email even if no accounts need to be reported on
                                                                   // For use in combination with the previous options
var addWeekCols = false;                                            // Add extra columns for the days of the week
var features = {
  conversions: true,
  averageCpc: true,
  ctr: true,
  costPerConversion: true,
  ticketMedio: true,
  roas: true,
  yesterday: {
    cost: true,
    conversions: true,
    averageCpc: true,
    ctr: true,
    costPerConversion: true
  }
};
//I wanted to see if we can add Converted Clicks, average cost per click, ckick through rate, and cost per converted click into this?

// Report colors
var overspendColor = '#ffc7c1';
var underspendColor = '#fffec1';

var overCPAColor1 = '#EE7600';
var overCPAColor2 = 'red';
var underCPAColor1 = 'green';

var overSpendColor = 'red';
var rightSpendColor = 'green';
var underSpendColor = 'blue';

var overROASColor1 = 'green';
var underROASColor1 = 'red';

var iconCPA = '&#9679;';
var sizeIconCPA = '25px';

var iconOverspend = '&#9650;';
var iconRightSpend = '&#9648;';
var iconUnderspend = '&#9660;';
var sizeIconSpend = '20px';

var iconROAS = '&#9679;';
var sizeIconROAS = '25px';

// These are default values. Values found in the spreadsheet for the account will supercede these.
var defaultCombination = 'OR';
var defaultBudget = 0;    // Budget for the month in the AdWords account default currency
// var defaultOver = 0.1;    // The amount the current tally can go over the running target budget (incl. today) before a warning is sent
// var defaultUnder = 0.20;   // The amount the current tally can be under the running target budget. Set to 0 to ignore
                          // both are decimals representing percent: 0.1 = 10%, 1 = 100%, etc
var defaultOver1 = 10;
var defaultOver2 = 25;
var defaultOver3 = 35;
var defaultUnder1 = 20;
var defaultUnder2 = 30;
var defaultUnder3 = 60;

var defaultCPA = 0;
var defaultOverCPA1 = 15;
var defaultOverCPA2 = 30;
var defaultUnderCPA1 = -20;

var defaultTicketMedio = 1;
var defaultROASGoal = 3;
var defaultUnderROAS1 = 20;

// ADVANCED!
// --------------------------------------------------------
// Map spreadsheet info to fields we need
// Here you can ajust what column translates to a property
// 0 = A, 1 = B, etc.
// We only work with the first sheet,
// and only look at the first 26 columns (A-Z)
// ONLY USE IF YOU WANT TO USE A DIFFERENT FORMAT THAN THAT
// CREATED BY THE prefillSpreadsheet OPTION!
 function mapRowToInfo(row) {
   return {
      custId: row[1].trim(),
      cust: row[0],
      budget: row[2],
      labels: row[9].split(','),
      andOr: row[10],
      over1: row[3],
      over2: row[4],
      over3: row[5],
      under1: row[6],
      under2: row[7],
      under3: row[8],
      cpa: row[11],
      overCPA1: row[12],
      overCPA2: row[13],
      underCPA1: row[14],
      getTicketMedio: row[15],
      roasGoal: row[16],
      underROAS1: row[17]
   };
 }


// PLEASE DON'T TOUCH BELOW THIS LINE

var skipLabel = ''; // Fallback to retain backwards compatibility

function main() {
  Logger.log(getSpreadsheetIds());
  try {
    // Uses parallel execution. Is limited to 50 accounts by Google.
    if (prefillSpreadsheet) {
      MccApp.accounts()
        .withLimit(50)
        .executeInParallel("getSSAccountInfo","saveSSAccountInfo");
    } else {
      MccApp.accounts()
        .withIds(getSpreadsheetIds())
        .withLimit(50)
        .executeInParallel("processAccount", "processReports");
    }
  } catch (e) {
    processReport(processAccount());
  }
}

// Get account name and id
function getSSAccountInfo() {
  var result = {
    custId: AdWordsApp.currentAccount().getCustomerId(),
    cust: AdWordsApp.currentAccount().getName()
  };
  Logger.log(result);
  return JSON.stringify(result);
}

// Save account info to the spreadsheet
function saveSSAccountInfo(response) {
  var ss;
  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
  }
  ss = ss.getSheets()[0];
  ss.appendRow(["Account Name", "Account ID", "Budget", "Overspend Ratio", "Underspend Ratio", "Campaign labels", "Label AND/OR"]);
  for (var i in response) {
    if(!response[i].getReturnValue()) { continue; }
    var rep = JSON.parse(response[i].getReturnValue());
    Logger.log(rep);
    ss.appendRow([rep.cust, rep.custId]);
  }
}


//I wanted to see if we can add Converted Clicks, average cost per click, ckick through rate, and cost per converted click into this?
function costsForIterator(iterator, stats) {
  while (iterator.hasNext()) {
    var item = iterator.next();
    stats.skipList[item.getId()] = true;
    // We only look for costs made in the current month
    var iStats = item.getStatsFor("THIS_MONTH");
    stats.cost += iStats.getCost();
    stats.clicks += iStats.getClicks();
    stats.impressions += iStats.getImpressions();
    stats.conversions += iStats.getConversions();

    var iStatsYesterday = item.getStatsFor("YESTERDAY");
    stats.yesterday.cost += iStatsYesterday.getCost();
    stats.yesterday.clicks += iStatsYesterday.getClicks();
    stats.yesterday.impressions += iStatsYesterday.getImpressions();
    stats.yesterday.conversions += iStatsYesterday.getConversions();
  }
  return stats;
}

function costsForAdgroupsIterator(iterator, labels, operator, stats, skipList) {
  if (!operator) {
    operator = 'CONTAINS_ANY';
  }
  if (iterator) {
    while (iterator.hasNext()) {
      var campaign = iterator.next();
      if (skipList[campaign.getId()]) {
        continue;
      }
      var adGroups = campaign.adGroups();
      if (labels.length !== 0) {
        adGroups = adGroups.withCondition("LabelNames " + operator + " [" + labels.join(', ') + "]");
      }

      stats = costsForIterator(adGroups.get(), stats);
    }
  }
  return stats;
}

function processAccount() {
  var multiAccountInfo = getAccountInfo(),
      currency = AdWordsApp.currentAccount().getCurrencyCode(),
      results = [];

  for (var i in multiAccountInfo) {
    var accountInfo = multiAccountInfo[i],
        accountId = accountInfo.custId,
        budget = accountInfo.budget,
        over1 = accountInfo.over1,
        over2 = accountInfo.over2,
        over3 = accountInfo.over3,
        under1 = accountInfo.under1,
        under2 = accountInfo.under2,
        under3 = accountInfo.under3,
        account = accountInfo.cust,
        labels = accountInfo.labels,
        andOr = accountInfo.andOr,
        cpaLimit = accountInfo.cpa,
        overCPA1 = accountInfo.overCPA1,
        overCPA2 = accountInfo.overCPA2,
        underCPA1 = accountInfo.underCPA1,
        getTicketMedio = accountInfo.getTicketMedio,
        roasGoal = accountInfo.roasGoal,
        underROAS1 = accountInfo.underROAS1;

    var stats = {
          skipList: {},
          cost: 0,
          conversions: 0,
          clicks: 0,
          impressions: 0,
          yesterday: {
            cost: 0,
            conversions: 0,
            clicks: 0,
            impressions: 0
          }
        },
        date = new Date();

    // There is no good way to use the account's timezone
    // We use UTC, because it's close enough for use in Europe
    // The standard timezone is most likely PST

    // We adjust the current date so the amount of the day that
    // has past is taken into account when setting a target
    // Otherwise, running early in the day gives up to a day of
    // extra budget, even though there has been no opportunity to spend it.
    var today = date.getUTCDate() - (1 - date.getUTCHours()/23),
        days = 32 - new Date(date.getFullYear(), date.getMonth(), 32).getUTCDate();

    // This is a pretty naive way to plot this information
    // But unless your spending is significantly weighted within a month
    // it should be a decent predictor
    var partOfMonth = today/days,  // How far we are in the month
        maxInclTodayNoOver = partOfMonth * budget,      // The part of the budget allotted to the part of the month that has passed
        maxInclToday1 = maxInclTodayNoOver * (1 + (over1 / 100)), // The amount of money that has to be spent to warrant a warning mail
        maxInclToday2 = maxInclTodayNoOver * (1 + (over2 / 100)),
        maxInclToday3 = maxInclTodayNoOver * (1 + (over3 / 100)),
        minInclToday1 = maxInclTodayNoOver * (1 - (under1 / 100)),
        minInclToday2 = maxInclTodayNoOver * (1 - (under2 / 100)),
        minInclToday3 = maxInclTodayNoOver * (1 - (under3 / 100));

    // Get the campaigns for your account
    // You can add some conditions here, to limit the accounts counted etc.
    var campaignIterator = AdWordsApp.campaigns();
    var shoppingCampaignIterator = AdWordsApp.shoppingCampaigns();
    var operator = (andOr === 'AND') ? 'CONTAINS_ALL' : 'CONTAINS_ANY';
    var doAdgroups = false;
    if (labels.length === 1 && labels[0] == '') {
      labels = [];
    }
    if (labels.length !== 0) {
      for (var j in labels) {
        labels[j] = '\'' + labels[j].trim() + '\''
      }
      campaignIterator = campaignIterator.withCondition("LabelNames " + operator + " [" + labels.join(', ') + "]");
      shoppingCampaignIterator = shoppingCampaignIterator.withCondition("LabelNames " + operator + " [" + labels.join(', ') + "]");
      doAdgroups = true;
    } else if (hasLabel(skipLabel)) {
      campaignIterator = campaignIterator.withCondition("LabelNames CONTAINS_NONE ['" + skipLabel + "']");
      shoppingCampaignIterator = shoppingCampaignIterator.withCondition("LabelNames CONTAINS_NONE ['" + skipLabel + "']");
    }
    campaignIterator = campaignIterator.get();
    shoppingCampaignIterator = shoppingCampaignIterator.get();

    stats = costsForIterator(campaignIterator, stats);
    if (doAdgroups) {
      stats = costsForAdgroupsIterator(AdWordsApp.campaigns().get(), labels, operator, stats, stats.skipList);
      stats.skipList = {};
    }

    stats = costsForIterator(shoppingCampaignIterator, stats);
    if (doAdgroups) {
      stats = costsForAdgroupsIterator(AdWordsApp.shoppingCampaigns().get(), labels, operator, stats, stats.skipList);
      stats.skipList = {};
    }

    //Get the values that we can get by a report.
    /*var report = AdWordsApp.report("SELECT ConversionValue " +
       "FROM CAMPAIGN_PERFORMANCE_REPORT " +
       "DURING THIS_MONTH");*/
    var report = AdWordsApp.report("SELECT ConversionValue " +
       "FROM ACCOUNT_PERFORMANCE_REPORT " +
       "DURING THIS_MONTH");
    var rows = report.rows();

    var conversionValue = 0;
    while (rows.hasNext()) {
      var row = rows.next();
      var convValue = row['ConversionValue'];
      conversionValue += parseFloat(convValue);
    }

    Logger.log(accountId + ' - Val. Conv: ' + conversionValue + ' - Coste: ' + stats.cost);

    var cost = stats.cost;
    var diff = 0;
    if (cost > maxInclToday3) {
        diff = 3;
    } else if (cost > maxInclToday2) {
        diff = 2;
    } else if (cost > maxInclToday1) {
        diff = 1;
    } else if (cost < minInclToday3 && under3 !== 0) {
        diff = -3;
    } else if (cost < minInclToday2 && under2 !== 0) {
        diff = -2;
    } else if (cost < minInclToday1 && under1 !== 0) {
        diff = -1;
    }

    var diffLabel = {
        '-3': '<span style="font-size:' + sizeIconSpend + '; color:' + underSpendColor + ';">' + iconUnderspend + iconUnderspend + iconUnderspend + '</span>',
        '-2': '<span style="font-size:' + sizeIconSpend + '; color:' + underSpendColor + ';">' + iconUnderspend + iconUnderspend + '</span>',
        '-1': '<span style="font-size:' + sizeIconSpend + '; color:' + underSpendColor + ';">' + iconUnderspend + '</span>',
        '0': '<span style="font-size:' + sizeIconSpend + '; color:' + rightSpendColor + ';">' + iconRightSpend + '</span>',
        '1': '<span style="font-size:' + sizeIconSpend + '; color:' + overSpendColor + ';">' + iconOverspend + '</span>',
        '2': '<span style="font-size:' + sizeIconSpend + '; color:' + overSpendColor + ';">' + iconOverspend + iconOverspend + '</span>',
        '3': '<span style="font-size:' + sizeIconSpend + '; color:' + overSpendColor + ';">' + iconOverspend + iconOverspend + iconOverspend + '</span>'
    };

    // var diffLabel = {
    //     '-3': '<span style="font-size:' + sizeIconSpend + '; color:' + underSpendColor + ';">' + iconUnderspend + iconUnderspend + iconUnderspend + '</span> (' + cost + ') (' + minInclToday3 + ') (' + under3 + ')',
    //     '-2': '<span style="font-size:' + sizeIconSpend + '; color:' + underSpendColor + ';">' + iconUnderspend + iconUnderspend + '</span> (' + cost + ') (' + minInclToday2 + ') (' + under2 + ')',
    //     '-1': '<span style="font-size:' + sizeIconSpend + '; color:' + underSpendColor + ';">' + iconUnderspend + '</span> (' + cost + ') (' + minInclToday1 + ') (' + under1 + ')',
    //     '0': '<span style="font-size:' + sizeIconSpend + '; color:' + rightSpendColor + ';">' + iconRightSpend + '</span>',
    //     '1': '<span style="font-size:' + sizeIconSpend + '; color:' + overSpendColor + ';">' + iconOverspend + '</span> (' + cost + ') (' + maxInclToday1 + ') (' + over1 + ')',
    //     '2': '<span style="font-size:' + sizeIconSpend + '; color:' + overSpendColor + ';">' + iconOverspend + iconOverspend + '</span> (' + cost + ') (' + maxInclToday2 + ') (' + over2 + ')',
    //     '3': '<span style="font-size:' + sizeIconSpend + '; color:' + overSpendColor + ';">' + iconOverspend + iconOverspend + iconOverspend + '</span> (' + cost + ') (' + maxInclToday3 + ') (' + over3 + ')'
    // };

    var remainingBudget = budget - maxInclTodayNoOver,
        delta = cost - maxInclTodayNoOver,
        daysLeft = days - today;

    // Format results as an object. processReports decides what to do
    var result = {
      reportable: (diff !== 0),
      cust: account,
      custId: accountId,
      budget: fMoney(currency, budget),
      budgetInt: budget,
      diff: diff,
      status: diffLabel[diff],
      target: fMoney(currency, maxInclTodayNoOver),
      actual: fMoney(currency, cost),
      delta: fMoney(currency, delta) + ' <br>(' + twoDecPerc(delta/maxInclTodayNoOver) + ')',
      recommend: fMoney(currency, ((budget - cost) / daysLeft)),
    };
    var cpcc;
    if (features.conversions) {
      result.conversions = stats.conversions;
    }
    if (features.averageCpc) {
      result.averageCpc = fMoney(currency, stats.cost / stats.clicks);
    }
    if (features.ctr) {
      result.ctr = twoDecPerc(stats.clicks / stats.impressions);
    }
    if (features.costPerConversion) {
        cpcc = stats.conversions == 0 ? 'N/A'
            : twoDec(stats.cost / stats.conversions);
        if(cpcc != 0 && typeof cpaLimit !== 'undefined' && cpaLimit != '' && cpaLimit != 0) {
            var cpcc_percent_over = Math.ceil( ((cpcc - cpaLimit) / cpaLimit) *100 );
        } else {
            var cpcc_percent_over = 0;
        }
        
        if(cpcc_percent_over >= overCPA2) {
            result.costPerConversion = cpcc + '<span style="margin-left:5px; font-size:' + sizeIconCPA + '; color:' + overCPAColor2 + ';">' + iconCPA + '</span>';
        } else if(cpcc_percent_over >= overCPA1) {
            result.costPerConversion = cpcc + '<span style="margin-left:5px; font-size:' + sizeIconCPA + '; color:' + overCPAColor1 + ';">' + iconCPA + '</span>';
        } else if(cpcc_percent_over <= underCPA1) {
            result.costPerConversion = cpcc + '<span style="margin-left:5px; font-size:' + sizeIconCPA + '; color:' + underCPAColor1 + ';">' + iconCPA + '</span>';
        } else {
            result.costPerConversion = cpcc;
        }
    }
    
    if(features.ticketMedio) {
        if(getTicketMedio) {
            ticketMed = typeof conversionValue === 'undefined' || conversionValue == 0 || stats.conversions == 0 ?  fMoney(currency, '0.00') : fMoney(currency, twoDec(conversionValue / stats.conversions));
        } else {
            ticketMed = '-';
        }
        result.ticketMedio = ticketMed;
    }

    if(features.roas) {
        if(getTicketMedio) {
            roas = typeof conversionValue === 'undefined' || conversionValue == 0 || stats.cost == 0 ? 'N/A' : twoDec(conversionValue / stats.cost);

            if(roas != 0 && typeof roasGoal !== 'undefined' && roasGoal != '' && roasGoal != 0) {
                var roas_percent_over = Math.ceil( ((roas - roasGoal) / roasGoal) *100 );
            } else {
                var roas_percent_over = 0;
            }

            if(roas_percent_over > roasGoal) {
                result.roas = roas + '<span style="margin-left:5px; font-size:' + sizeIconROAS + '; color:' + overROASColor1 + ';">' + iconROAS + '</span>';
            } else if(roas_percent_over <= underROAS1) {
                result.roas = roas + '<span style="margin-left:5px; font-size:' + sizeIconROAS + '; color:' + underROASColor1 + ';">' + iconROAS + '</span>';
            } else {
                result.roas = roas;
            }
        } else {
            result.roas = '-';
        }
    }

    if (features.yesterday.cost) {
        if(stats.yesterday.cost == 0) {
            result.yesterday_cost = '<div style="width:100%; height:100%; background-color:#d92a2e; color:#fff; padding-left:2px;">' + fMoney(currency, stats.yesterday.cost) + '</div>';
        } else {
            result.yesterday_cost = fMoney(currency, stats.yesterday.cost);
        }
    }
    if (features.yesterday.conversions) {
      result.yesterday_conversions = stats.yesterday.conversions;
    }
    if (features.yesterday.averageCpc) {
      result.yesterday_averageCpc = fMoney(currency, stats.yesterday.cost / stats.yesterday.clicks);
    }
    if (features.yesterday.ctr) {
      result.yesterday_ctr = twoDecPerc(stats.yesterday.clicks / stats.yesterday.impressions);
    }
    if (features.yesterday.costPerConversion) {
      cpcc = stats.yesterday.conversions == 0 ? 'N/A'
        : twoDec(stats.yesterday.cost / stats.yesterday.conversions);
      result.yesterday_costPerConversion = cpcc;
    }

    if (addWeekCols) {
      result.wd_ma = ' ';
      result.wd_di = ' ';
      result.wd_wo = ' ';
      result.wd_do = ' ';
      result.wd_vr = ' ';
    }
    results.push(result);
  }

  return JSON.stringify(results);
}

function hasLabel(label) {
  return AdWordsApp.labels().withCondition("Name = '" + label + "'").
      get().hasNext();
}

// Process the results of a single
// Creates table, exports as html and sends to set emailaddress
function processReport(report) {
  // Define table(headers)
  var table = buildTable();
  rep = JSON.parse(report);
  for (var j in rep) {
    // Skip campaign if budget is 0 and ignoreNoBudgetCampaigns is on
    if (ignoreNoBudgetCampaigns && rep[j].budgetInt === 0) { continue; }
    // Only show records that have over/underspend if onlyReportProblems is true
    if (onlyReportProblems === false || rep[j].reportable === true) {
      var attrs = {};
      if (rep[j].diff <= -2) { attrs.style = 'background-color: ' + underspendColor; }
      else if (rep[j].diff >= 2) { attrs.style = 'background-color: ' + overspendColor; }
      add_row(table, rep[j], attrs);
    }
  }
  sendEmail(table);
}

// Process the results of all the accounts
// Creates table, exports as html and sends to set emailaddress
function processReports(reports) {
  // Define table(headers)
  var table = buildTable();
  for (var i in reports) {
    if(!reports[i].getReturnValue()) { continue; }
    var rep = JSON.parse(reports[i].getReturnValue());
    for (var j in rep) {
      // Skip campaign if budget is 0 and ignoreNoBudgetCampaigns is on
      if (ignoreNoBudgetCampaigns && rep[j].budgetInt === 0) { continue; }
      // Only show records that have over/underspend if onlyReportProblems is true
      if (onlyReportProblems === false || rep[j].reportable === true) {
        var attrs = {style: 'padding: 2px 4px;'};
        if (rep[j].diff <= -2) { attrs.style += 'background-color: ' + underspendColor; }
        else if (rep[j].diff >= 2) { attrs.style += 'background-color: ' + overspendColor; }
        add_row(table, rep[j], attrs);
      }
    }
  }
  sendEmail(table);
}

function sendEmail(table) {
   // Only send if there is something to report, or alwaysReport is set.
  if (alwaysReport || table.rows.length > 0) {
    var htmlBody = '<' + 'h1>' + emailSubject + '<' + '/h1>' + render_table(table, {border: 1, cellpadding: 0, cellspacing: 0, width: "95%", style: "border-collapse:collapse;"});
    MailApp.sendEmail(emailAddr, emailSubject, emailSubject, { htmlBody: htmlBody });
  }
}

function buildTable() {
  var tableCols = {
    cust: 'Customer',
    custId: 'Customer-ID',
    budget: 'Budget',
    status: 'Status',
    target: 'Target spend',
    actual: 'Actual spend',
    delta: 'Delta',
    recommend: 'Recommended daily spend',
  };

  if (features.conversions) {
    tableCols.conversions = 'Conversions';
  }
 
  if (features.costPerConversion) {
    tableCols.costPerConversion = 'Cost per conv';
  }

  if(features.ticketMedio) {
    tableCols.ticketMedio = 'Ticket Medio';
  }

  if(features.roas) {
    tableCols.roas = 'ROAS';
  }

  if (features.yesterday.cost) {
    tableCols.yesterday_cost = 'Spend (y)';
  }
 

  if (addWeekCols) {
    tableCols.wd_ma = 'Mo';
    tableCols.wd_di = 'Tu';
    tableCols.wd_wo = 'We';
    tableCols.wd_do = 'Th';
    tableCols.wd_vr = 'Fr';
  }
  return create_table(tableCols);
}

// Few formatting functions
function twoDec(i) {
  return parseFloat(i).toFixed(2);
}

function twoDecPerc(p) {
  return twoDec(p * 100) + '%';
}

function fMoney(currency, amount) {
  if (currency == 'EUR') {
    currency = '€';
  } else if (currency == 'USD') {
    currency = '$';
  }
  return currency + '&nbsp;' + twoDec(amount);
}

function getSpreadsheetIds() {
  var ids = [],
      ss,
      reAWId = /^([0-9]{3})-([0-9]{3})-([0-9]{4})$/;

  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return ids;
  }
  ss = ss.getSheets()[0];
  var rows = parseInt(ss.getLastRow());
  var range = ss.getRange("A1:Z" + rows).getValues();
  for (var i = 0; i < rows; i++) {
    var account = mapRowToInfo(range[i]);
    if (!reAWId.test(account.custId)) {
      continue;
    }
    ids.push(account.custId);
  }
  return ids;
}

// Fetch info for current account from the spreadsheet
// MCC scripts don't seem to support shared state between
// Parallel executions, so we need to do this fresh for every account

// Uses default info from 'defaults' set in script, and replaces with
// values from spreadsheet where possible
function getAccountInfo() {
  var ss;
  var reAWId = /^([0-9]{3})-([0-9]{3})-([0-9]{4})$/;
  var protoAccount = {
    custId: AdWordsApp.currentAccount().getCustomerId(),
    cust: AdWordsApp.currentAccount().getName(),
    budget: defaultBudget,
    labels: [],
    over1: defaultOver1,
    over2: defaultOver2,
    over3: defaultOver3,
    under1: defaultUnder1,
    under2: defaultUnder2,
    under3: defaultUnder3,
    andOr: defaultCombination,
    cpa: defaultCPA,
    overCPA1: defaultOverCPA1,
    overCPA2: defaultOverCPA2,
    underCPA1: defaultUnderCPA1,
    getTicketMedio: defaultTicketMedio,
    roasGoal: defaultROASGoal,
    underROAS1: defaultUnderROAS1
  };

  try {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    return [protoAccount];
  }
  ss = ss.getSheets()[0];
  var rows = parseInt(ss.getLastRow());
  var range = ss.getRange("A1:Z" + rows).getValues();
  var accounts = [];

  for (var i = 0; i < rows; i++) {
    var account = mapRowToInfo(range[i]);
    if (!reAWId.test(account.custId) || account.custId !== protoAccount.custId) {
      continue;
    }

    for(var key in account) {
      if (account[key] === '') {
        account[key] = protoAccount[key];
      }
    }
    accounts.push(account);
  }
  if (accounts.length === 0) {
    return [protoAccount];
  } else {
    return accounts;
  }
}

// Instantiate a table object with given column names
// Either as array or object/hash
function create_table(cols) {
  var table = { head: [], rows: [], row_attrs: [], row_names: undefined};
  if (cols instanceof Array) {
    table.head = cols;
  } else if (cols instanceof Object) {
    var i = 0;
    table.row_names = {};
    for (var key in cols) {
      table.head.push(cols[key]);
      table.row_names[key] = i;
      i++;
    }
  }
  return table;
}

// Add a row to the table object
// Either an clean array or an object
// with correct parameter names
function add_row(table, row, attrs) {
  if (row instanceof Array) {
    table.rows.push(row);
    return;
  }
  if (table.row_names === undefined) {
    return;
  }
  var new_row = [];
  for (var key in row) {
    if (table.row_names[key] === undefined) {
      continue;
    }
    new_row[table.row_names[key]] = row[key];
  }
  table.rows.push(new_row);
  table.row_attrs.push(attrs);
}

// Log the contents of the table object in a semi readable format
function log_table(table) {
  Logger.log('----------------------------------');
  Logger.log(table.head.join(' | '));
  Logger.log('----------------------------------');
  for (var i in table.rows) {
    Logger.log(table.rows[i].join(' | '));
  }
  Logger.log('----------------------------------');
}

// Turn the table object into an HTML table
// Add attributes to the table tag with the attrs param
// Takes an object/hash
function render_table(table, attrs) {
  function render_tag(content, tag_name, attrs) {
    var attrs_str = '';
    if (attrs instanceof Object) {
      for (var attr in attrs) {
        attrs_str += [' ',attr,'="', attrs[attr], '"'].join('');
      }
    }
    var tag = ['<' + tag_name + attrs_str + '>'];
    tag.push(content);
    tag.push('<!--' + tag_name + '-->');
    return tag.join('');
  }
  function render_row(row, field, row_attrs) {
    if (field === undefined) {
      field = 'td';
    }
    var row_ar = new Array(table.head.length);
    for (var col in row) {
      row_ar.push(render_tag(row[col], field, row_attrs));
    }
    return render_tag(row_ar.join(''), 'tr');
  }
  var table_ar = [];
  table_ar.push(render_row(table.head, 'th', {style:'background-color: #333333; color: #fefefe; padding: 2px 4px'}));
  for (var row in table.rows) {
    table_ar.push(render_row(table.rows[row], 'td', table.row_attrs[row]));
  }
  return render_tag(table_ar.join(''), 'table', attrs);
}