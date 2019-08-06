//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * FFXIV Automatic Spreadsheet
// * Created by Chakraa Arcana @ Leviathan.
// * Discord: Chakraa#1837
// * Created on: Jan 25th 2019
// * Purpose: Through the use of https://xivapi.com/, the goal is to display selected information of the
// * selected characters through their LodestoneID.
//////////////////////////////////////////////////////////////////////////////////////////////////////////


 /*
  * Declaration of a bunch of Global Variables used throughout the script.
  * The hard coded numbers represent the column/row number in which the information
  * need to be placed or pick off.
  * CH = Character
  * FC = Free Company
  */

var _RosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster'),
	_RosterSheetAmountOfRows = _RosterSheet.getLastRow(),
	_RosterSheetFirstCharacterScannedRow = 2,
	_InstructionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Instructions'),
    _InstructionsSheetAPIKeyRow = 4,
    _AmountOfHeaders = 1,
	_CHID = [], 
    _CHInfo = [], 
	_CHSheet = [],
    _CHSheet2 = [],
    _CHIDColumn = 3,
    _CHFCJoinDate = 39,	
	_CHLastUpdatedColumn = 36,
    _ClassOrder = [1,3,32,37,6,26,33,2,4,29,34,5,31,38,7,26,35,36,8,9,10,11,12,13,14,15,16,17,18], //The order is a bit weird, but the API is done like that. In the desired order (Tank, Heal, DPS, DoH, DoL). The IDs of the classes.
	_APIKey = "",
    _AlarmWrongilvl = false;

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * Added by Raven Ambree @ Excalibur.
// * Purpose: Get FC Lodestone to load membership data + other FC Info
//////////////////////////////////////////////////////////////////////////////////////////////////////////

	_FCRosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FC Roster'),
    _FCAmountOfHeaders = 3,
	_FCRosterSheetFirstCharacterScannedRow = 4,
	_FCRosterSheetAmountOfRows = _FCRosterSheet.getLastRow() - _FCAmountOfHeaders,
    _FCID = 0,
    _FCRow = 2,
    _FCInfo = "",
    _FCStoredIDs = [],
    _FCAvatarColumn = 1,
    _FCNameColumn = 2,
    _FCIDColumn = 3,
    _FCServerColumn = 4,
    _FCMemberCountColumn = 5,
    _FCServerRankColumn = 6,
    _FCTimeEquation = 38,
    _RunningFCScript = false;
      

function startFCRosterSheet() {
	if(isAPIKeyValid()){
      _RunningFCScript = true;
	  _FCID = _FCRosterSheet.getRange(_FCRow,_CHIDColumn).getDisplayValue()
      
	  // (Free Company Addon)
	  // If an FC Lodestone ID is present, get all members lodestone ids and list them in the sheet
	  if( _FCID !== "") {
		fetchFCInfo();
		addNewIDs();
	  }
	  
	  // If we are using the FCSheet, all the data of the RosterSheet becomes the one from FCRosterSheet
      _CHLastUpdatedColumn = _CHLastUpdatedColumn + 1;
	  _RosterSheet = _FCRosterSheet;
	  _RosterSheetAmountOfRows = _FCRosterSheetAmountOfRows;
	  _RosterSheetFirstCharacterScannedRow = _FCRosterSheetFirstCharacterScannedRow;
	  _AmountOfHeaders = _FCAmountOfHeaders;
	  fetchIDsFromSheet();
	}
}


// fetch FC info
function fetchFCInfo() {
  _FCID = _FCRosterSheet.getRange(_FCRow,_CHIDColumn).getDisplayValue();
  _FCInfo = JSON.parse(UrlFetchApp.fetch("https://xivapi.com/freecompany/" + _FCID + "?key=" + _APIKey + "&data=FCM").getContentText());

  if(_FCInfo.hasOwnProperty('FreeCompany')){
      updateFC(_FCInfo,_FCRow);
   }
    
   else{
      _FCRosterSheet.getRange(FCRow, _FCNameColumn).setValue("Not Found");
   }
}

function updateFC(freecompany, row) {
  
   // add FC Logo
   _FCRosterSheet.getRange(row, _FCAvatarColumn).setValue("=IMAGE(\"https://xivapi.com/freecompany/" + freecompany.FreeCompany.ID + "/icon\",2)");
  
  // add FC Name
  _FCRosterSheet.getRange(row, _FCNameColumn).setValue("=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/freecompany/" + freecompany.FreeCompany.ID.replace('i','') + "\", \"" + freecompany.FreeCompany.Name + " «" + freecompany.FreeCompany.Tag + "» " + "\")");
  
  // add server FC is from
  _FCRosterSheet.getRange(row, _FCServerColumn).setValue(freecompany.FreeCompany.Server + " (" + freecompany.FreeCompany.DC + ")");
  
  // add Membership Count
  _FCRosterSheet.getRange(row, _FCMemberCountColumn).setValue(freecompany.FreeCompanyMembers.length);
  
  // add Ranking (Weekly/Monthly)
  _FCRosterSheet.getRange(row, _FCServerRankColumn).setValue(freecompany.FreeCompany.Ranking.Weekly + " / " + freecompany.FreeCompany.Ranking.Monthly);

}



// Checks for existing IDs. If a Free Company Member has already been added to the sheet, don't include it in new array.
function addNewIDs() {
  var sheetIDs = [],
      parsedIDs = [],
      sorted2DIDs = [],
      data = [];
  
  // retrieve IDs from FC member list, put into parsedIDs
  for(var i = 0; i < _FCInfo.FreeCompanyMembers.length; i++) {
    parsedIDs[i] = _FCInfo.FreeCompanyMembers[i].ID
    
    var arr = [];
    
    arr[0] = _FCInfo.FreeCompanyMembers[i].ID
    arr[1] = _FCInfo.FreeCompanyMembers[i].Rank
    sorted2DIDs[i] = arr
  }
  // create a sorted list of parsed IDs
  sorted2DIDs.sort(sortNumber);
  
  
  
  // pull all the data from the sheet at once. This puts it into a 2D array 4 across, and a number of rows down. 
  if (_FCRosterSheetAmountOfRows > 0) {
    var sheet2DIDs = _FCRosterSheet.getRange(_FCRosterSheetFirstCharacterScannedRow, _CHIDColumn, _FCRosterSheetAmountOfRows,4).getDisplayValues()
    for (var i = 0; i < sheet2DIDs.length; i++) {
      var id = parseInt(sheet2DIDs[i][0])
      sheetIDs[i] = id
      if (isNaN(sheetIDs[i])) {
        sheetIDs[i] = ""
      }
    }
  }

    var newIDs = parsedIDs.filter( function(id) { return this.indexOf( id ) < 0; }, sheetIDs);
    data = mergeArrays(newIDs, sheetIDs)
    var timeRows = sheetIDs.length
    timeRows = data.length



    
    // Map the data into a 2D array for bulk placement (thank you, stackoverflow)
    // https://stackoverflow.com/questions/53090460/1d-array-to-fill-a-spreadsheet-column
    
    var newData = data.map(function (elem) { return [elem]; });
    if (newData.length > 0) {  
      var range = _FCRosterSheet.getRange(_FCRosterSheetFirstCharacterScannedRow ,_CHIDColumn, newData.length,1);
      range.setValues(newData);
    }
  
  // If FC member is already on sheet, get their parsed rank and update it.
  for (i = 0; i < data.length; i++){
    newData[i][0] = binarySearch(sorted2DIDs, data[i], 0, sorted2DIDs.length-1)
  }  
  var rankData = newData.map(function (elem) { return [elem]; });
  var range = _FCRosterSheet.getRange(_FCRosterSheetFirstCharacterScannedRow ,_FCServerRankColumn, rankData.length,1);
    range.setValues(rankData);
  
    // update the roster with the new amount of rows.
    _FCRosterSheetAmountOfRows = _FCRosterSheet.getLastRow()
  
    // add times and equations for all empty time and equation cells.
    updateDates(_FCRosterSheetAmountOfRows - _FCAmountOfHeaders);
    updateTimeEquations(_FCRosterSheetAmountOfRows - _FCAmountOfHeaders);
  }



// search through a sorted array for a key. Used recursively.
function binarySearch(arr, key, low, high) {
  var mid = parseInt((high + low) / 2);

  if (arr[mid][0] == key) {
    return arr[mid][1]
  } else if (low >= high) {
    return ""
  } else if (arr[mid][0] > key) {
    return binarySearch(arr, key, low, mid-1)
  } else {
    return binarySearch(arr, key, mid+1, high) 
  }
}

function sortNumber(a, b) {
  return a[0] - b[0];
}

// fills in blank rows with the current date
function updateDates(numRows) {
  
  
  if (numRows == 0) { return }
  // data placed. Update when we saw them
    var date2DArray = _FCRosterSheet.getRange(_FCRosterSheetFirstCharacterScannedRow, _CHFCJoinDate, numRows, 1).getDisplayValues();
    var dateArray = new Array();
    var date = Utilities.formatDate(new Date(), "GMT-8", "MM/dd/yyyy")
  
  // convert dates to 1D array
    for (var i = 0; i < numRows; i++) {
      dateArray[i] = date2DArray[i][0];
      if (dateArray[i] == "") { dateArray[i] = date }
    }
  
  // prep for pasting onto the sheet
    var times = dateArray.map(function (elem) { return [elem]; });
    range = _FCRosterSheet.getRange(_FCRosterSheetFirstCharacterScannedRow ,_CHFCJoinDate, times.length,1);
    range.setValues(times);
  
}


// places formulas in as many rows as there are users, every update.
function updateTimeEquations(numRows) {
  if (numRows == 0) { return }
  // add time equations to all cells, in case they were missed previously.
  var offset = (_CHFCJoinDate - _FCTimeEquation)
  var formulas = new Array()
  
  // generate formula column
  for (var i = 0; i < numRows; i++) {
    formulas[i] = "=IFERROR(IF(ISBLANK(R[0]C[" + offset + "]),,TODAY()-R[0]C[" + offset + "]))"
  }
  
  // paste it back into sheet
  var formulaMap = formulas.map(function (elem) { return [elem]; });
  var range = _FCRosterSheet.getRange(_FCRosterSheetFirstCharacterScannedRow, _FCTimeEquation, formulaMap.length, 1);
  range.setFormulasR1C1(formulaMap);
}

/*
 * Function
 * inputs: arrays a, b
 * returns: array c
 * 
 * example:
 * mergeArrays( [1,2,3,4],[a, ,b, ,c] )
 * gives the result
 * [a, 1, b, 2, c, 3, 4]
 * 
 */
function mergeArrays( aArr, bArr ) {
  
  var row = 0,
      bLen = bArr.length,
      aLen = aArr.length;
  var result = new Array();
  
  if (aLen == 0) {
    return bArr
  }
  
  for (var i = 0; i < (aLen); i++) {

    // find first vacant row to throw the ID in
    while (bArr[row] != "" && row < bArr.length) {
      result[row] = bArr[row]
      row++
    }
    
    // Vacant row found. put an ID from IDs in the empty space
    result[row] = aArr[i]
    row++
  }
  
  return result
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * End of Raven Ambree code segment added.
// * Thank you for the base to the script. I had lots of fun adding to this!
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * Beginning of the Roster Sheet Code by Chakraa Arcana
//////////////////////////////////////////////////////////////////////////////////////////////////////////

function fetchIDsFromSheet(){
	if(isAPIKeyValid()){
		_CHID = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHIDColumn, _RosterSheetAmountOfRows-_AmountOfHeaders, 1).getDisplayValues();
		fetchCharacterInfos();
	}
}

function isAPIKeyValid()
{
  if(_InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow,1).getDisplayValue() == ""){
    SpreadsheetApp.getUi().alert("Oh oh, something happened!", 
                                 "Make sure that you have your API Key in the Red Box in the Instructions Page.",
                                 SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
    
  }
  else if(/\s/g.test(_InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow,1).getDisplayValue())){
    SpreadsheetApp.getUi().alert("Oh oh, something happened!", 
                                 "Make sure that your API Key is correctly set in the Red Box of the Instructions Page and doesn't have any spaces, tabs or line returns.",
                                 SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
  else{
    _APIKey = _InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow,1).getDisplayValue();
    return true;
  }
}

function fetchCharacterInfos() {
  var pollingError = false;
  
  for(var i = 0; i < _CHID.length; ++i){
    if(_CHID[i] == ""){
		SpreadsheetApp.getUi().alert("Oh oh, something happened!", 
                                   "Make sure that you don't have data in any cells anywhere else than the rows you have Lodestone IDs in. If you aren't sure, clear everyone and rerun the script :)",
                                  SpreadsheetApp.getUi().ButtonSet.OK);
		pollingError = true;
		break;
    }
  }
  
  if(!pollingError){
	var apikey = _InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow,1).getDisplayValue();
	var temparray;
	var chunk = 2;
	for(var i = 0; i < _CHID.length; i += chunk){
        temparray = _CHID.slice(i,i+chunk);
		var options = {
            muteHttpExceptions : true,
			method : 'POST',
			payload : JSON.stringify({
			key : apikey,
			ids : temparray.toString(),
			data : 'FC',
			extended : 1
			})
		};
		_CHInfo = _CHInfo.concat(JSON.parse(UrlFetchApp.fetch("https://xivapi.com/characters", options).getContentText()));
	}
    if(_RunningFCScript)
      updateCharacterFCSheet();
    else
      updateCharacter();
  }
}

function updateCharacter(){
  for(var i = 0;i < _CHID.length;++i){
      _CHSheet[i] = [];
    if(_CHInfo[i].hasOwnProperty('Character')){
      _CHSheet[i][0] = "=HYPERLINK(\"" + _CHInfo[i].Character.Portrait + "\", IMAGE(\""+ _CHInfo[i].Character.Avatar + "\"))";
      _CHSheet[i][1] = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/character/" + _CHInfo[i].Character.ID + "\", \"" + _CHInfo[i].Character.Name + "\")";
      _CHSheet[i][2] = _CHInfo[i].Character.ID;
      _CHSheet[i][3] = _CHInfo[i].Character.Server + " (" + _CHInfo[i].Character.DC + ")";
      _CHSheet[i][4] = updateFreeCompany(_CHInfo[i]);
      _CHSheet[i][5] = updateCurrentClass(_CHInfo[i]);
      _CHSheet[i][6] = new Date(_CHInfo[i].Character.ParseDate * 1000);
      _CHSheet[i].splice.apply(_CHSheet[i],[6,0].concat(updateClassJobsAndTime(_CHInfo[i])));
    }
    else{
      for(var j = 0;j < _CHLastUpdatedColumn;++j){
        _CHSheet[i].push("");
      }
      _CHSheet[i][1] = "Not Found";
      _CHSheet[i][2] = _CHID[i];
    }
  }
  _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, 1, _RosterSheetAmountOfRows-_AmountOfHeaders, _CHLastUpdatedColumn).setValues(_CHSheet);   
  if(_AlarmWrongilvl)
    SpreadsheetApp.getUi().alert("Oh oh, something happened!", 
                                   "It looks like some of the equipment one of the characters is wearing hasn't been added to the database yet." + '\n' + "The shown equipped ilvl might be wrong, but the rest of the sheet is as normal." + '\n' +  "It should be fixed automatically shortly.",
                                  SpreadsheetApp.getUi().ButtonSet.OK);    
}

function updateCharacterFCSheet(){
  var ColumnNumberToSplit = 5;
  for(var i = 0;i < _CHID.length;++i){
      _CHSheet[i] = [],
      _CHSheet2[i] = [];
    if(_CHInfo[i].hasOwnProperty('Character')){
      _CHSheet[i][0] = "=HYPERLINK(\"" + _CHInfo[i].Character.Portrait + "\", IMAGE(\""+ _CHInfo[i].Character.Avatar + "\"))";
      _CHSheet[i][1] = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/character/" + _CHInfo[i].Character.ID + "\", \"" + _CHInfo[i].Character.Name + "\")";
      _CHSheet[i][2] = _CHInfo[i].Character.ID;
      _CHSheet[i][3] = _CHInfo[i].Character.Server + " (" + _CHInfo[i].Character.DC + ")";
      _CHSheet[i][4] = updateFreeCompany(_CHInfo[i]);
      _CHSheet2[i][0] = updateCurrentClass(_CHInfo[i]);
      _CHSheet2[i][1] = new Date(_CHInfo[i].Character.ParseDate * 1000);
      _CHSheet2[i].splice.apply(_CHSheet2[i],[1,0].concat(updateClassJobsAndTime(_CHInfo[i])));
    }
    else{
      for(var j = 0;j < ColumnNumberToSplit;++j){ //The 5 is really bad but I dont know how to do better
        _CHSheet[i].push("");
      }
      for(var j = ColumnNumberToSplit + 1;j < _CHLastUpdatedColumn;++j){
        _CHSheet2[i].push("");
      }
      _CHSheet[i][1] = "Not Found";
      _CHSheet[i][2] = _CHID[i];
    }
  }
  _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, 1, _RosterSheetAmountOfRows-_AmountOfHeaders, ColumnNumberToSplit).setValues(_CHSheet);
  _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, ColumnNumberToSplit + 2, _RosterSheetAmountOfRows-_AmountOfHeaders, _CHLastUpdatedColumn - (ColumnNumberToSplit + 1)).setValues(_CHSheet2);
  if(_AlarmWrongilvl)
    SpreadsheetApp.getUi().alert("Oh oh, something happened!", 
                                   "It looks like some of the equipment one of the characters is wearing hasn't been added to the database yet." + '\n' + "The shown equipped ilvl might be wrong, but the rest of the sheet is as normal." + '\n' +  "It should be fixed automatically shortly.",
                                  SpreadsheetApp.getUi().ButtonSet.OK);    
}

function updateFreeCompany(character){
	var FCName = "", FCRank = "";
  if(character.Character.FreeCompanyId !== null){
    if(character.FreeCompany !== null){
      FCName = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/freecompany/" + character.FreeCompany.ID.replace('i','') + "\", \"" + character.FreeCompany.Name + " «" + character.FreeCompany.Tag + "» " + "\")";
	}
	else
	{
      FCName = "Loading...";
	}
  }
  return FCName;
}

function updateCurrentClass(character){
  var ilvl = 0,
      gear = character.Character.GearSet.Gear,
      keys = Object.keys(gear),
      mainHandIndex = 0,
      offhand = false,
      soulcrystal = false;
  for(var amountofslots = 0; amountofslots < keys.length; ++amountofslots){
    if(keys[amountofslots].toString() !== "SoulCrystal" && keys[amountofslots].toString() !== "MainHand") //Ignore SoulCrystal dans total
      if(gear[keys[amountofslots]].Item !== null)
        ilvl += gear[keys[amountofslots]].Item.LevelItem;
      else
        _AlarmWrongilvl = true;
    if(keys[amountofslots].toString() == "MainHand")
      mainHandIndex = amountofslots;
    if(keys[amountofslots].toString() == "OffHand")
      offhand = true;
    if(keys[amountofslots].toString() == "SoulCrystal")
      soulcrystal = true;
  }
  
  //If you are a class wearing an offhand, you count the main hand once. If PLD or DoH only has its main hand equipped, will count twice.
  ilvl += (offhand ? gear[keys[mainHandIndex]].Item.LevelItem : gear[keys[mainHandIndex]].Item.LevelItem * 2);

  var ilvlstring = '\n' + Math.floor(ilvl/13) + " ilvl"; //Build the string to display with the average ilvl rounded to 2 decimal places. 13 is the number of gear slots.
  var abbreviation = (soulcrystal ? character.Character.GearSet.Job.Abbreviation : character.Character.GearSet.Class.Abbreviation);
  return abbreviation + ilvlstring;
}

function updateClassJobsAndTime(character){
  var classJobs = character.Character.ClassJobs;
  var keys  = Object.keys(classJobs);
  var AllClasses = [];
  for(var j = 0; j < _ClassOrder.length; ++j){
    //Add Class Levels
    AllClasses[j] = "-";
    for(var i = 0; i < _ClassOrder.length; ++i){
      if(classJobs[keys[i]] != null)
        if(classJobs[keys[i]].Class.ID === _ClassOrder[j]){
          if(classJobs[keys[i]].Level === 0){ //Put "-" instead of 0 when the person has not yet unlock his class.
            AllClasses[j] = "-";
          }
          else if(classJobs[keys[i]].ExpLevelTogo === 0){ //If level max, we just register it.
            AllClasses[j] = classJobs[keys[i]].Level;
          }
          else{ //We put the percentage of non-max levels, and we put it smaller than the written level.
            var level = classJobs[keys[i]].Level;
            var xpPercent = "\n(" + Math.round(classJobs[keys[i]].ExpLevel / classJobs[keys[i]].ExpLevelMax * 100) + "%)";
            AllClasses[j] = level + xpPercent;   
          }
          break;
        }
    }
  }
  return AllClasses;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// This Method is used to get the amount of milliseconds for debug purposes. 
// They will be displayed in Google Script's Logger.
// https://gist.github.com/dkodr/e1a131e09914c21042daf7a2ccdb48a4
//
// Usage:
// var runtimeCountStart = new Date();
// Logger.log("X Function Time in Milliseconds:");
// Logger.log(runtimeCountStop(runtimeCountStart));
//////////////////////////////////////////////////////////////////////////////////////////////////////////

function runtimeCountStop(start) {

  var props = PropertiesService.getScriptProperties();
  var currentRuntime = props.getProperty("runtimeCount");
  var stop = new Date();
  var newRuntime = Number(stop) - Number(start) + Number(currentRuntime);
  var setRuntime = {
    runtimeCount: newRuntime,
  }
  return setRuntime;
}