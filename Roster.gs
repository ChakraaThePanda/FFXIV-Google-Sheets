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
 * Prefix codes
 * _CH for Character
 * _FC for Free Company
 */

var _RosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster'),
  _RosterSheetAmountOfRows = _RosterSheet.getLastRow(),
  _RosterSheetFirstCharacterScannedRow = 2,
  _InstructionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Instructions'),
  _InstructionsSheetAPIKeyRow = 4,
  _AmountOfHeaders = 1,
  _CHID = [],
  _CHLastUpdated = [],
  _CHInfo = [],
  _CHSheet = [],
  _CHSheet2 = [],
  _CHIDColumn = 3,
  _CHFCJoinDate = 39,
  _CHLastUpdatedColumn = 36,
  _CHTimeEquation = 38,
  _ClassOrder = [19, 21, 32, 37, 24, 28, 33, 20, 22, 30, 34, 23, 31, 38, 25, 27, 35, 36, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18], //The order is a bit weird, but the API is done like that. In the desired order (Tank, Heal, DPS, DoH, DoL). The IDs of the classes.
  _APIKey = "",
  _AlarmWrongilvl = false;
  _RunningFCScript = false;
  _RunningLSScript = false;


/*
 * Collection of settings functions for the Roster, including name changes
 */

function runFCScript() {
  _RosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FC Roster'),
  _AmountOfHeaders = 3,
  _RosterSheetFirstCharacterScannedRow = 4,
  _FCRosterSheetAmountOfRows = _FCRosterSheet.getLastRow() - _FCAmountOfHeaders,
  _FCID = 0,
  _FCRow = 2,
  _FCAPIInfo = "",
  _FCStoredIDs = [],
  _FCAvatarColumn = 1,
  _FCNameColumn = 2,
  _FCIDColumn = 3,
  _FCServerColumn = 4,
  _FCMemberCountColumn = 5,
  _FCServerRankColumn = 6,
  _FCTimeEquation = 38;

  updateCurrentSheet();
}

function runLSScript() {


  updateCurrentSheet();
}




//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * Beginning of the Roster Sheet Code by Chakraa Arcana
//////////////////////////////////////////////////////////////////////////////////////////////////////////

function updateCurrentSheet() {
  if (isAPIKeyValid()) {
    _CHID = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHIDColumn, _RosterSheetAmountOfRows - _AmountOfHeaders, 1).getDisplayValues();
    fetchCharacterInfos();
  }
}

function isAPIKeyValid() {
  if (_InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue() == "") {
    SpreadsheetApp.getUi().alert("Oh oh, something happened!",
      "Make sure that you have your API Key in the Red Box in the Instructions Page.",
      SpreadsheetApp.getUi().ButtonSet.OK);
    return false;

  }
  else if (/\s/g.test(_InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue())) {
    SpreadsheetApp.getUi().alert("Oh oh, something happened!",
      "Make sure that your API Key is correctly set in the Red Box of the Instructions Page and doesn't have any spaces, tabs or line returns.",
      SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
  else {
    _APIKey = _InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue();
    return true;
  }
}

function fetchCharacterInfos() {
  var pollingError = false;

  for (var i = 0; i < _CHID.length; ++i) {
    if (_CHID[i] == "") {
      SpreadsheetApp.getUi().alert("Oh oh, something happened!",
        "Make sure that you don't have data in any cells anywhere else than the rows you have Lodestone IDs in. If you aren't sure, clear everyone and rerun the script :)",
        SpreadsheetApp.getUi().ButtonSet.OK);
      pollingError = true;
      break;
    }
  }

  if (!pollingError) {
    var apikey = _InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue();
    for (var i = 0; i < _CHID.length; i += chunk) {
      temparray = _CHID.slice(i, i + chunk);
      var options = {
        muteHttpExceptions: true,
        method: 'POST',
        payload: JSON.stringify({
          key: apikey,
          ids: temparray.toString(),
          data: 'FC',
          extended: 1
        })
      };
      _CHInfo = _CHInfo.concat(JSON.parse(UrlFetchApp.fetch("https://xivapi.com/characters", options).getContentText()));
    }
      updateCharacter();
  }
}

function updateCharacter() {
  for (var i = 0; i < _CHID.length; ++i) {
    _CHSheet = [];

    if (_CHInfo[i].hasOwnProperty('Character')) {

      // If the character ID is valid, and has data from it
      _CHSheet[0] = "=HYPERLINK(\"" + _CHInfo[i].Character.Portrait + "\", IMAGE(\"" + _CHInfo[i].Character.Avatar + "\"))";
      _CHSheet[1] = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/character/" + _CHInfo[i].Character.ID + "\", \"" + _CHInfo[i].Character.Name + "\")";
      _CHSheet[2] = _CHInfo[i].Character.ID;
      _CHSheet[3] = _CHInfo[i].Character.Server + " (" + _CHInfo[i].Character.DC + ")";
      _CHSheet[4] = updateFreeCompany(_CHInfo[i]);
      _CHSheet[5] = updateCurrentClass(_CHInfo[i]);
      _CHSheet[6] = new Date(_CHInfo[i].Character.ParseDate * 1000);
      _CHSheet.splice.apply(_CHSheet[i], [6, 0].concat(updateClassJobsAndTime(_CHInfo[i])));

    } else {
      // If the character ID is not valid, blank the line and state "Not Found"
      for (var j = 0; j < _CHLastUpdatedColumn; ++j) {
        _CHSheet[i].push("");
      }

      _CHSheet[1] = "Not Found";
      _CHSheet[2] = _CHID[i];
    }


    _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, 1, _RosterSheetAmountOfRows - _AmountOfHeaders, _CHLastUpdatedColumn).setValues(_CHSheet);

  }


  if (_AlarmWrongilvl)
    SpreadsheetApp.getUi().alert("Oh oh, something happened!",
      "It looks like some of the equipment one of the characters is wearing hasn't been added to the database yet." + '\n' + "The shown equipped ilvl might be wrong, but the rest of the sheet is as normal." + '\n' + "It should be fixed automatically shortly.",
      SpreadsheetApp.getUi().ButtonSet.OK);
}

function updateFreeCompany(character) {
  var FCName = "", FCRank = "";
  if (character.Character.FreeCompanyId !== null) {
    if (character.FreeCompany !== null) {
      FCName = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/freecompany/" + character.FreeCompany.ID.replace('i', '') + "\", \"" + character.FreeCompany.Name + " «" + character.FreeCompany.Tag + "» " + "\")";
    }
    else {
      FCName = "Loading...";
    }
  }
  return FCName;
}

function updateCurrentClass(character) {
  var ilvl = 0,
    gear = character.Character.GearSet.Gear,
    keys = Object.keys(gear),
    mainHandIndex = 0,
    offhand = false,
    soulcrystal = false;
  for (var amountofslots = 0; amountofslots < keys.length; ++amountofslots) {
    if (keys[amountofslots].toString() !== "SoulCrystal" && keys[amountofslots].toString() !== "MainHand") //Ignore SoulCrystal dans total
      if (gear[keys[amountofslots]].Item !== null)
        ilvl += gear[keys[amountofslots]].Item.LevelItem;
      else
        _AlarmWrongilvl = true;
    if (keys[amountofslots].toString() == "MainHand")
      mainHandIndex = amountofslots;
    if (keys[amountofslots].toString() == "OffHand")
      offhand = true;
    if (keys[amountofslots].toString() == "SoulCrystal")
      soulcrystal = true;
  }

  //If you are a class wearing an offhand, you count the main hand once. If PLD or DoH only has its main hand equipped, will count twice.
  ilvl += (offhand ? gear[keys[mainHandIndex]].Item.LevelItem : gear[keys[mainHandIndex]].Item.LevelItem * 2);

  var ilvlstring = '\n' + Math.floor(ilvl / 13) + " ilvl"; //Build the string to display with the average ilvl rounded to 2 decimal places. 13 is the number of gear slots.
  var abbreviation = (soulcrystal ? character.Character.GearSet.Job.Abbreviation : character.Character.GearSet.Class.Abbreviation);
  return abbreviation + ilvlstring;
}

function updateClassJobsAndTime(character) {
  var classJobs = character.Character.ClassJobs;
  var keys = Object.keys(classJobs);
  var AllClasses = [];
  for (var j = 0; j < _ClassOrder.length; ++j) {
    //Add Class Levels
    AllClasses[j] = "-";
    for (var i = 0; i < _ClassOrder.length; ++i) {
      if (classJobs[keys[i]] != null)
        if (classJobs[keys[i]].Job.ID === _ClassOrder[j]) {
          if (classJobs[keys[i]].Level === 0) { //Put "-" instead of 0 when the person has not yet unlock his class.
            AllClasses[j] = "-";
          }
          else if (classJobs[keys[i]].ExpLevelTogo === 0) { //If level max, we just register it.
            AllClasses[j] = classJobs[keys[i]].Level;
          }
          else { //We put the percentage of non-max levels, and we put it smaller than the written level.
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


/* After this line are helper methods used in both scripts. */