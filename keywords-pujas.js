// https://goo.gl/Au1RQF
var SPREADSHEET_URL = '';

function main() {
    var mccAccount = AdsApp.currentAccount();

    var spreadsheetAccess = new SpreadsheetAccess(SPREADSHEET_URL, 'Rules');

    spreadsheetAccess.spreadsheet.setSpreadsheetTimeZone(
        AdsApp.currentAccount().getTimeZone());
    prepareSheet(spreadsheetAccess);

    var row = spreadsheetAccess.nextRow();

    var keywordsEmail = [];

    while (row != null) {
        var argument;
        var stopLimit;
        try {
            argument = parseArgument(spreadsheetAccess, row);
            stopLimit = parseStopLimit(spreadsheetAccess, row);
        } catch (e) {
            logError(spreadsheetAccess, e);
            row = spreadsheetAccess.nextRow();
            continue;
        }
        var customerId = row[spreadsheetAccess.CUSTOMERID_INDEX];
        var account = null;
        try {
            var accountIterator = AdsManagerApp.accounts().withIds([customerId]).get();
            if (accountIterator.totalNumEntities() == 0) {
                throw ('Missing account: ' + customerId);
            } else {
                account = accountIterator.next();
            }
        } catch (e) {
            logError(spreadsheetAccess, e);
            row = spreadsheetAccess.nextRow();
            continue;
        }

        AdsManagerApp.select(account);

        var selector = buildSelector(spreadsheetAccess, row);
        var keywords = selector.get();

        try {
            keywords.hasNext();
        } catch (e) {
            logError(e);
            row = spreadsheetAccess.nextRow();
            continue;
        }

        var action = row[spreadsheetAccess.RULE_INDEX];
        var results = applyRules(keywords, action, argument, stopLimit, keywordsEmail);
        logResult(spreadsheetAccess, 'Fetched ' + results.fetched + '\nChanged ' +
            results.changed);
        row = spreadsheetAccess.nextRow();
    }

    sendSimpleTextEmail(keywordsEmail);

    spreadsheetAccess.spreadsheet.getRangeByName('last_execution')
        .setValue(new Date());
}

function sendSimpleTextEmail(keywordsEmail) {

    table = "<html><body><br><table border=0><tr><th>Palabras a revisar</th></tr>";

    for (var i = 0; i < keywordsEmail.length; i++) {
        table = table + "<tr><td>" + keywordsEmail[i] + "</td></tr>";

    }

    table = table + "</table></body></html>";


    MailApp.sendEmail('email@email.com',
        'Revisión manual pujas de keywords', '',
        { htmlBody: table});
}

/**
 * Prepares the spreadsheet for saving data.
 *
 * @param {Object} spreadsheetAccess the SpreadsheetAccess instance that
 *    handles the spreadsheet.
 */
function prepareSheet(spreadsheetAccess) {
    // Clear the results column.
    spreadsheetAccess.sheet.getRange(
        spreadsheetAccess.START_ROW,
        spreadsheetAccess.RESULTS_COLUMN_INDEX + spreadsheetAccess.START_COLUMN,
        spreadsheetAccess.MAX_COLUMNS, 1).clear();
}

/**
 * Builds a keyword selector based on the conditional column headers in the
 * spreadsheet.
 *
 * @param {Object} spreadsheetAccess the SpreadsheetAccess instance that
 *    handles the spreadsheet.
 * @param {Object} row the spreadsheet header row.
 * @return {Object} the keyword selector, based on spreadsheet header settings.
 */
function buildSelector(spreadsheetAccess, row) {
    var columns = spreadsheetAccess.getColumnHeaders();
    var selector = AdsApp.keywords();

    for (var i = spreadsheetAccess.FIRST_CONDITIONAL_COLUMN;
        i < spreadsheetAccess.RESULTS_COLUMN_INDEX; i++) {
        var header = columns[i];
        var value = row[i];
        if (!isNaN(parseFloat(value)) || value.length > 0) {
            if (header.indexOf("'") > 0) {
                value = value.replace(/\'/g, "\\'");
            } else if (header.indexOf('\"') > 0) {
                value = value.replace(/"/g, '\\\"');
            }
            var condition = header.replace('?', value);
            selector.withCondition(condition);
        }
    }
    selector.forDateRange(spreadsheetAccess.spreadsheet
        .getRangeByName('date_range').getValue());
    return selector;
}

/**
 * Applies the rules in the spreadsheet.
 *
 * @param {Object} keywords the keywords selector.
 * @param {String} action the action to be taken.
 * @param {String} argument the parameters for the operation specified by
 *   action.
 * @param {Number} stopLimit the upper limit to the bid value when applying
 *   rules.
 * @return {Object} the number of keywords that were fetched and modified.
 */
function applyRules(keywords, action, argument, stopLimit, keywordsEmail) {
    var fetched = 0;
    var changed = 0;

    while (keywords.hasNext()) {
        var keyword = keywords.next();
        var oldBid = keyword.bidding().getCpc();
        var newBid = 0;
        fetched++;
        if (action == 'Add') {
            if (stopLimit && oldBid > stopLimit) {
                newBid = oldBid;
                keywordsEmail.push(action + ' CPC Limite < Puja actual ' + keyword);
            } else {
                newBid = addToBid(oldBid, argument, stopLimit);
            }
        } else if (action == 'Multiply by') {
            if (stopLimit && oldBid > stopLimit) {
                newBid = oldBid;
                keywordsEmail.push(action + ' CPC Limite < Puja actual ' + keyword);
            } else {
                newBid = multiplyBid(oldBid, argument, stopLimit);
            }
        } else if (action == 'Lower Cpc') {
            newBid = lowerBid(oldBid, argument, stopLimit);
        } else if (action == 'Set to First Page Cpc') {
            newBid = keyword.getFirstPageCpc();
            if (stopLimit && (oldBid > newBid || oldBid > stopLimit)) {
                newBid = oldBid;
                keywordsEmail.push(action + ' CPC Limite < Puja actual ' + keyword);
            }
            if ((stopLimit && newBid > stopLimit) || (stopLimit && keyword.getFirstPageCpc() == 'undefined') || newBid == null) {
                newBid = stopLimit;
            }
        } else if (action == 'Set to Top of Page Cpc') {
            newBid = keyword.getTopOfPageCpc();
            if (stopLimit && (oldBid > newBid || oldBid > stopLimit)) {
                newBid = oldBid;
                keywordsEmail.push(action + ' CPC Limite < Puja actual ' + keyword);
            }
            if ((stopLimit && newBid > stopLimit) || (stopLimit && keyword.getTopOfPageCpc() == 'undefined') || newBid == null) {
                newBid = oldBid;
            }
        }

        if (newBid < 0) {
            newBid = 0.01;
        }
        newBid = newBid.toFixed(2);
        if (newBid != oldBid) {
            changed++;
        }
        keyword.bidding().setCpc(newBid);
        Logger.log(action + ': ' + keyword + ' nuestro CPC: ' + oldBid + ' Stop Limit: ' + stopLimit + ' CPC parte superior absoluta: ' + keyword.getTopOfPageCpc() + ' CPC parte superior: ' + keyword.getFirstPageCpc() + ' cambiado a ' + newBid)

    }
    return {
        'fetched': fetched,
        'changed': changed
    };
}

/**
 * Adds a value to an existing bid, while applying a stop limit.
 *
 * @param {Number} oldBid the existing bid.
 * @param {Number} argument the bid increment to apply.
 * @param {Number} stopLimit the cutoff limit for modified bid.
 * @return {Number} the modified bid.
 */
function addToBid(oldBid, argument, stopLimit) {
    return applyStopLimit(oldBid + argument, stopLimit, argument > 0);
}

/**
 * Multiplies an existing bid by a value, while applying a stop limit.
 *
 * @param {Number} oldBid the existing bid.
 * @param {Number} argument the bid multiplier.
 * @param {Number} stopLimit the cutoff limit for modified bid.
 * @return {Number} the modified bid.
 */
function multiplyBid(oldBid, argument, stopLimit) {
    return applyStopLimit(oldBid * argument, stopLimit, argument > 1);
}

function lowerBid(oldBid, argument, stopLimit) {
    var newBid = oldBid - oldBid * argument;
    if (stopLimit && newBid > stopLimit) {
        newBid = stopLimit;
    }
    return newBid;
}

/**
 * Applies a cutoff limit to a bid modification.
 *
 * @param {Number} newBid the modified bid.
 * @param {Number} stopLimit the bid cutoff limit.
 * @param {Boolean} isPositive true, if the stopLimit is an upper cutoff limit,
 *    false if it a lower cutoff limit.
 * @return {Number} the modified bid, after applying the stop limit.
 */
function applyStopLimit(newBid, stopLimit, isPositive) {
    if (stopLimit) {
        if (isPositive && newBid > stopLimit) {
            newBid = stopLimit;
        } else if (!isPositive && newBid < stopLimit) {
            newBid = stopLimit;
        }
    }
    return newBid;
}

/**
 * Parses the argument for an action on the spreadsheet.
 *
 * @param {Object} spreadsheetAccess the SpreadsheetAccess instance that
 *    handles the spreadsheet.
 * @param {Object} row the spreadsheet action row.
 * @return {Number} the parsed argument for the action.
 * @throws error if argument is missing, or is not a number.
 */
function parseArgument(spreadsheetAccess, row) {
    if (row[spreadsheetAccess.ARGUMENT_INDEX].length == 0 &&
        (row[spreadsheetAccess.RULE_INDEX] == 'Add' ||
            row[spreadsheetAccess.RULE_INDEX] == 'Multiply by')) {
        throw ('\"Argument\" must be specified.');
    }
    var argument = parseFloat(row[spreadsheetAccess.ARGUMENT_INDEX]);
    if (isNaN(argument)) {
        throw 'Bad Argument: must be a number.';
    }
    return argument;
}

/**
 * Parses the stop limit for an action on the spreadsheet.
 *
 * @param {Object} spreadsheetAccess the SpreadsheetAccess instance that
 *    handles the spreadsheet.
 * @param {Object} row the spreadsheet action row.
 * @return {Number} the parsed stop limit for the action.
 * @throws error if the stop limit is not a number.
 */
function parseStopLimit(spreadsheetAccess, row) {
    if (row[spreadsheetAccess.STOP_LIMIT_INDEX].length == 0) {
        return null;
    }
    var limit = parseFloat(row[spreadsheetAccess.STOP_LIMIT_INDEX]);
    if (isNaN(limit)) {
        throw 'Bad Argument: must be a number.';
    }
    return limit;
}

/**
 * Logs the error to the spreadsheet.
 *
 * @param {Object} spreadsheetAccess the SpreadsheetAccess instance that
 *    handles the spreadsheet.
 * @param {String} error the error message.
 */
function logError(spreadsheetAccess, error) {
    Logger.log(error);
    spreadsheetAccess.sheet.getRange(spreadsheetAccess.currentRow(),
        spreadsheetAccess.RESULTS_COLUMN_INDEX +
        spreadsheetAccess.START_COLUMN, 1, 1)
        .setValue(error)
        .setFontColor('#c00')
        .setFontSize(8)
        .setFontWeight('bold');
}

/**
 * Logs the results to the spreadsheet.
 *
 * @param {Object} spreadsheetAccess the SpreadsheetAccess instance that
 *    handles the spreadsheet.
 * @param {String} result the result message.
 */
function logResult(spreadsheetAccess, result) {
    spreadsheetAccess.sheet.getRange(spreadsheetAccess.currentRow(),
        spreadsheetAccess.RESULTS_COLUMN_INDEX +
        spreadsheetAccess.START_COLUMN, 1, 1)
        .setValue(result)
        .setFontColor('#444')
        .setFontSize(8)
        .setFontWeight('normal');
}

/**
 * Controls access to the data spreadsheet.
 *
 * @param {String} spreadsheetUrl the spreadsheet url.
 * @param {String} sheetName name of the spreadsheet that contains the bid
 *     rules.
 * @constructor
 */
function SpreadsheetAccess(spreadsheetUrl, sheetName) {

    /**
     * Gets the next row in sequence.
     *
     * @return {?Array.<Object> } the next row, or null if there are no more
     *     rows.
     * @this SpreadsheetAccess
     */
    this.nextRow = function () {
        for (; this.rowIndex < this.cells.length; this.rowIndex++) {
            if (this.cells[this.rowIndex][0]) {
                return this.cells[this.rowIndex++];
            }
        }
        return null;
    };

    /**
     * The current spreadsheet row.
     *
     * @return {Number} the current row.
     * @this SpreadsheetAccess
     */
    this.currentRow = function () {
        return this.rowIndex + this.START_ROW - 1;
    };

    /**
     * The total number of data columns for the spreadsheet.
     *
     * @return {Number} the total number of data columns.
     * @this SpreadsheetAccess
     */
    this.getTotalColumns = function () {
        var totalCols = 0;
        var columns = this.getColumnHeaders();
        for (var i = 0; i < columns.length; i++) {
            if (columns[i].length == 0 || columns[i] == this.RESULTS_COLUMN_HEADER) {
                totalCols = i;
                break;
            }
        }
        return totalCols;
    };

    /**
     * Gets the list of column beaders.
     *
     * @return {Array.<String>} the list of column headers.
     * @this SpreadsheetAccess
     */
    this.getColumnHeaders = function () {
        return this.sheet.getRange(
            this.HEADER_ROW,
            this.START_COLUMN,
            1,
            this.MAX_COLUMNS - this.START_COLUMN + 1).getValues()[0];
    };

    /**
     * Gets the results column index.
     *
     * @return {Number} the results column index.
     * @throws exception if results column is missing.
     * @this SpreadsheetAccess
     */
    this.getResultsColumn = function () {
        var columns = this.getColumnHeaders();
        var totalColumns = this.getTotalColumns();

        if (columns[totalColumns] != 'Results') {
            throw ('Results column is missing.');
        }
        return totalColumns;
    };

    /**
     * Initializes the class methods.
     *
     * @this SpreadsheetAccess
     */
    this.init = function () {

        this.HEADER_ROW = 5;

        this.FIRST_CONDITIONAL_COLUMN = 4;
        this.START_ROW = 6;
        this.START_COLUMN = 2;

        Logger.log('Using spreadsheet - %s.', spreadsheetUrl);
        this.spreadsheet = validateAndGetSpreadsheet(spreadsheetUrl);

        this.sheet = this.spreadsheet.getSheetByName(sheetName);
        this.RESULTS_COLUMN_HEADER = 'Results';

        this.MAX_ROWS = this.sheet.getMaxRows();
        this.MAX_COLUMNS = this.sheet.getMaxColumns();

        this.CUSTOMERID_INDEX = 0;
        this.RULE_INDEX = 1;
        this.ARGUMENT_INDEX = 2;
        this.STOP_LIMIT_INDEX = 3;
        this.RESULTS_COLUMN_INDEX = this.getResultsColumn();


        this.cells = this.sheet.getRange(this.START_ROW, this.START_COLUMN,
            this.MAX_ROWS, this.MAX_COLUMNS).getValues();
        this.rowIndex = 0;
    };

    this.init();
}

/**
 * Validates the provided spreadsheet URL
 * to make sure that it's set up properly. Throws a descriptive error message
 * if validation fails.
 *
 * @param {string} spreadsheeturl The URL of the spreadsheet to open.
 * @return {Spreadsheet} The spreadsheet object itself, fetched from the URL.
 * @throws {Error} If the spreadsheet URL hasn't been set
 */
function validateAndGetSpreadsheet(spreadsheeturl) {
    if (spreadsheeturl == 'YOUR_SPREADSHEET_URL') {
        throw new Error('Please specify a valid Spreadsheet URL. You can find' +
            ' a link to a template in the associated guide for this script.');
    }
    return SpreadsheetApp.openByUrl(spreadsheeturl);
}
