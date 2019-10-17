//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * Free Company portion added by Raven Ambree @ Excalibur.
// * Purpose:
// * Get FC Lodestone to load membership IDs + stats info
//////////////////////////////////////////////////////////////////////////////////////////////////////////

function runFCScript() {
    // Name of sheet to update
    _RosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FC Roster'),
        _AmountOfHeaders = 3,
        _RosterSheetFirstCharacterScannedRow = 4,
        _RosterSheetAmountOfRows = _RosterSheet.getLastRow() - _AmountOfHeaders,
        _FCID = 0,
        _FCRow = 2,
        _FCAPIInfo = "",
        _FCStoredIDs = [],
        _FCAvatarColumn = 1,
        _FCNameColumn = 2,
        _FCIDColumn = 3,
        _FCServerColumn = 4,
        _FCMemberCountColumn = 5,
        _FCServerRankColumn = 6;

    updateFreeCompany();
    updateCurrentRoster();
}

function updateFreeCompany() {
    if (isAPIKeyValid()) {
        _FCID = _RosterSheet.getRange(_FCRow, _CHIDColumn).getDisplayValue()

        // (Free Company Addon)
        // If an FC Lodestone ID is not present, error out.
        if (_FCID == "") {

            SpreadsheetApp.getUi().alert("There isn't an FC ID to parse. Please put one in!",
                SpreadsheetApp.getUi().ButtonSet.OK);

        } else {
            // Update the FC info, then update the roster list with the IDs found
            getFCInfo();
            getFCIDs();
        }
    }
}


// fetch FC info from xivapi, if it exists. If it does, do work with it. If not, it isn't found.
function getFCInfo() {

    _FCID = _RosterSheet.getRange(_FCRow, _CHIDColumn).getDisplayValue();
    _FCAPIInfo = JSON.parse(UrlFetchApp.fetch("https://xivapi.com/freecompany/" + _FCID + "?key=" + _APIKey + "&data=FCM").getContentText());

    if (_FCAPIInfo.hasOwnProperty('FreeCompany')) {
        updateFCHeader(_FCAPIInfo, _FCRow);
    } else {
        _RosterSheet.getRange(FCRow, _FCNameColumn).setValue("Not Found");
    }
}

function updateFCHeader(freecompany, row) {

    // add FC Logo
    _RosterSheet.getRange(row, _FCAvatarColumn).setValue("=IMAGE(\"https://xivapi.com/freecompany/" + freecompany.FreeCompany.ID + "/icon\",2)");

    // add FC Name
    _RosterSheet.getRange(row, _FCNameColumn).setValue("=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/freecompany/" + freecompany.FreeCompany.ID.replace('i', '') + "\", \"" + freecompany.FreeCompany.Name + " «" + freecompany.FreeCompany.Tag + "» " + "\")");

    // add server FC is from
    _RosterSheet.getRange(row, _FCServerColumn).setValue(freecompany.FreeCompany.Server + " (" + freecompany.FreeCompany.DC + ")");

    // add Membership Count
    _RosterSheet.getRange(row, _FCMemberCountColumn).setValue(freecompany.FreeCompanyMembers.length);

    // add Ranking (Weekly/Monthly)
    _RosterSheet.getRange(row, _FCServerRankColumn).setValue(freecompany.FreeCompany.Ranking.Weekly + " / " + freecompany.FreeCompany.Ranking.Monthly);
}



// Checks for existing IDs. If a Free Company Member has already been added to the sheet, don't include it in new array.
function getFCIDs() {

    var sheetIDs = [],
        parsedIDs = [],
        sortedIDs = [];

    // for each member, put the ID in parsed IDs, and collect the rest of the info to be pasted later.
    for (var i = 0; i < _FCAPIInfo.FreeCompanyMembers.length; i++) {

        // parsed this ID
        parsedIDs[i] = _FCAPIInfo.FreeCompanyMembers[i].ID
        sortedIDs[i] = _FCAPIInfo.FreeCompanyMembers[i].ID
    }
    // create a sorted list by arr ID
    sortedIDs.sort(sortBy(0));

    // pull all the data from the sheet at once. This puts it into a 2D array 6 across, and a number of rows down.
    if (_RosterSheetAmountOfRows > 0) {
        //Collect all character membership data we have.
        var sheet2DIDs = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHIDColumn, _RosterSheetAmountOfRows, 1).getDisplayValues()

        // check if there's an ID present in each row.
        for (i = 0; i < sheet2DIDs.length; i++) {
            var id = parseInt(sheet2DIDs[i][0]);
            sheetIDs[i] = id;
            if (isNaN(sheetIDs[i])) {
                sheetIDs[i] = ""
            }
        }
    }

    // Check for new IDs from the FC API parse.
    var newIDs = parsedIDs.filter(function (id) { return this.indexOf(id) < 0; }, sheetIDs);

    // smear newIDs into sheetIDs, filling in holes and adding to the end.
    sheetIDs = mergeArrays(newIDs, sheetIDs)

    // paste the IDs into the sheet
    var data = sheetIDs.map(function (elem) { return [elem]; });
    var leftRange = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHIDColumn, data.length, 1);
    leftRange.setValues(data);

    // update the roster with the new amount of rows.
    _RosterSheetAmountOfRows = _RosterSheet.getLastRow()

    // add times and equations for all empty time and equation cells.
    updateDates(_RosterSheetAmountOfRows - _AmountOfHeaders);
    updateTimeEquations(_RosterSheetAmountOfRows - _AmountOfHeaders);
}


// fills in blank rows with the current date
function updateDates(numRows) {


    if (numRows == 0) { return }
    // data placed. Update when we saw them
    var date2DArray = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHListDate, numRows).getDisplayValues();
    var dateArray = new Array();
    var date = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy")

    // convert dates to 1D array
    for (var i = 0; i < numRows; i++) {
        dateArray[i] = date2DArray[i][0];
        if (dateArray[i] == "") { dateArray[i] = date }
    }

    // prep for pasting onto the sheet
    var times = dateArray.map(function (elem) { return [elem]; });
    range = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHListDate, times.length);
    range.setValues(times);

}


// places formulas in as many rows as there are users, every update.
function updateTimeEquations(numRows) {
    if (numRows == 0) { return }
    // add time equations to all cells, in case they were missed previously.
    var offset = (_CHListDate - _CHTimeEquationColumn)
    var formulas = new Array()

    // generate formula column
    for (var i = 0; i < numRows; i++) {
        formulas[i] = "=IFERROR(IF(ISBLANK(R[0]C[" + offset + "]),,TODAY()-R[0]C[" + offset + "]))"
    }

    // paste it back into sheet
    var formulaMap = formulas.map(function (elem) { return [elem]; });
    var range = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHTimeEquationColumn, formulaMap.length, 1);
    range.setFormulasR1C1(formulaMap);
}



//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * End of Raven Ambree code segment added.
// * Thank you for the base to the script. I had lots of fun adding to this!
//////////////////////////////////////////////////////////////////////////////////////////////////////////