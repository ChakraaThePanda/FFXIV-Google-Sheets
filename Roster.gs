//////////////////////////////////////////////////////////////////////////////////////////////////////////
// * FFXIV Automatic Spreadsheet
// * Created by Chakraa Arcana @ Leviathan with contributions from Raven Ambree @ Excalibur
// * Discord: Chakraa#1837 or Onlyme#7038 (respectively)
// * Created on: Jan 25th 2019
// * Purpose: Through the use of https://xivapi.com/, display information of the
// * selected characters through their LodestoneID.
//////////////////////////////////////////////////////////////////////////////////////////////////////////


/**
 * Declaration of a bunch of global variables used throughout the script(s).
 * Hard coded integers represent column/row numbers.
 * Variable Prefix codes:
 * _CH for Character
 * _FC for Free Company
 * _LS for Linkshell
 */

/** Instruction sheet configuration data. Do not change */
var _InstructionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Instructions'),
  _InstructionsSheetAPIKeyRow = 4,

  /** What sheet are we using? */
  _RosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster'),
  _RosterSheetAmountOfRows = _RosterSheet.getLastRow(),
  _RosterSheetFirstCharacterScannedRow = 2,


  _AmountOfHeaders = 1,

  /** What kind of sheet is it? Just your friends, your FC, your Linkshell?
   * Options: "Character", "Free Company", "Linkshell"
   * Default _SheetType = "Character"
   */
  _SheetType = "Character",

  /** Character global values */
  _CHID = [],
  _CHLastUpdated = [],
  _CHIDColumn = 3,
  _CHLastUpdatedColumn = 37,
  _CHTimeEquationColumn = 38,
  _CHListDate = 39,

  /**
   * The order is a bit weird, but the API is done like that.
   * In the desired order (Tank, Heal, DPS, DoH, DoL). The IDs of the classes.
   */
  _ClassOrder = [19, 21, 32, 37, 24, 28, 33, 20, 22, 30, 34, 23, 31, 38, 25, 27, 35, 36, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18],
  _APIKey = "",
  _AlarmWrongilvl = false;


function updateCurrentRoster() {
  if (isAPIKeyValid()) {
    /** get every ID, when it was last checked, and which row it is from the sheet */
    _CHID = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHIDColumn, _RosterSheetAmountOfRows - _AmountOfHeaders, 1).getDisplayValues();
    _CHLastUpdated = _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, _CHLastUpdatedColumn, _RosterSheetAmountOfRows - _AmountOfHeaders, 1).getDisplayValues();
    /** update the characters found */
    updateCharacterInfos();
  }
}

function isAPIKeyValid() {
  /** Is the API key field in the instructions list blank? */
  if (_InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue() == "") {
    SpreadsheetApp.getUi().alert("Oh oh, something happened!",
      "Make sure that you have your API Key in the Red Box in the Instructions Page.",
      SpreadsheetApp.getUi().ButtonSet.OK);
    /**  If blank, put in a key by following the writing on the instructions page. */
    return false;

  }
  /** Is the key incorrectly set? */
  else if (/\s/g.test(_InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue())) {
    SpreadsheetApp.getUi().alert("Oh oh, something happened!",
      "Make sure that your API Key is correctly set in the Red Box of the Instructions Page and doesn't have any spaces, tabs or line returns.",
      SpreadsheetApp.getUi().ButtonSet.OK);
    // Gotta fix the formatting of that key
    return false;
  }
  else {
    /** API key is good! Use it */
    _APIKey = _InstructionsSheet.getRange(_InstructionsSheetAPIKeyRow, 1).getDisplayValue();
    return true;
  }
}



/** Update each row by pulling from the characters lodestone ID */
function updateCharacterInfos() {
  var pollingError = false;

  /** For each row in ID... */
  for (var i = 0; i < _CHID.length; ++i) {
    /** If the row is empty, skip it and pop an error window out. */
    if (_CHID[i] == "") {
      SpreadsheetApp.getUi().alert("Oh oh, something happened!",
        "Make sure that you don't have data in any cells anywhere else than the rows you have Lodestone IDs in. If you aren't sure, clear everyone and rerun the script :)",
        SpreadsheetApp.getUi().ButtonSet.OK);
      pollingError = true;
      break;
    }
  }


  /** If the ID is non-empty */
  if (!pollingError) {


    /** How many requests at a time */
    var chunk = 10,
      charParse = [],
      CHInfo = [];
    for (var i = 0; i < _CHID.length; i + chunk) {

      /** Format the requests to the api */
      for (var j = 0; i + j < _CHID.length; j++) {

        var charID = _CHID[i + j][0]
        
        charParse[j] = 'https://xivapi.com/character/' + charID + '?key=' + _APIKey + '&data=FC,FCM&extended=1'

        /**      Old way of doing things.
         * var options = {
                muteHttpExceptions: true,
                method: 'POST',
                payload: JSON.stringify({
                  id: charID,
                  private_key: _APIKey,
                  data: 'FCM',
                  extended: 1
                })
              };
         * var CHInfo = JSON.parse(UrlFetchApp.fetch("https://xivapi.com/character", options).getContentText());
         * var CHInfo = JSON.parse(UrlFetchApp.fetch("https://xivapi.com/character/" + charID + "?key=" + _APIKey + "&data=FCM&extended=1"))
         */
      }

      CHInfo = JSON.parse(UrlFetchApp.fetchAll(charParse).getContentText());
      updateCharacter(CHInfo);
    } /** end loop */
  }
}

/** update character info with raw parsed info provided. */
function updateCharacter(CHParse) {
  /** CHLine will be pasted back to the sheet */
  CHLine = [];

  if (CHParse.hasOwnProperty('Character')) {

    /** If the character ID is valid, get the character's data */
    CHLine[0] = "=HYPERLINK(\"" + CHParse.Character.Portrait + "\", IMAGE(\"" + CHParse.Character.Avatar + "\"))";
    CHLine[1] = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/character/" + CHParse.Character.ID + "\", \"" + CHParse.Character.Name + "\")";
    CHLine[2] = CHParse.Character.ID;
    CHLine[3] = CHParse.Character.Server + " (" + CHParse.Character.DC + ")";
    CHLine[4] = updateFreeCompany(CHParse);
    CHLine[5] = updateFreeCompanyRank(CHParse);
    CHLine[6] = updateCurrentClass(CHParse);
    CHLine[7] = new Date(CHParse.Character.ParseDate * 1000);
    CHLine.splice.apply(_CHLine, [6, 0].concat(updateClassJobsAndTime(CHParse)));

  } else {
    /** If the character ID is not valid, blank the line and state "Not Found" */
    for (var j = 0; j < _CHLastUpdatedColumn; ++j) {
      CHLine[i].push("");
    }

    CHLine[1] = "Not Found";
    CHLine[2] = _CHID[i];
  }


  _RosterSheet.getRange(_RosterSheetFirstCharacterScannedRow, 1, _RosterSheetAmountOfRows - _AmountOfHeaders, _CHLastUpdatedColumn).setValues(CHLine);


  if (_AlarmWrongilvl)
    SpreadsheetApp.getUi().alert("Oh oh, something happened!",
      "It looks like some of the equipment one of the characters is wearing hasn't been added to the database yet." + '\n' + "The shown equipped ilvl might be wrong, but the rest of the sheet is as normal." + '\n' + "It should be fixed automatically shortly.",
      SpreadsheetApp.getUi().ButtonSet.OK);
}

function updateFreeCompany(character) {
  /** Assume no FC */
  var FCName = "";

  /** If character is in a Free Company */
  if (character.Character.FreeCompanyId !== null) {

    /** Retrieve the name and tag */
    if (character.FreeCompany !== null) {
      FCName = "=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/freecompany/" + character.FreeCompany.ID.replace('i', '') + "\", \"" + character.FreeCompany.Name + " «" + character.FreeCompany.Tag + "» " + "\")";
    } else {
      FCName = "Loading...";
    }
  }
  return FCName;
}

function updateFreeCompanyRank(character) {
  /** Assume no rank */
  var FCRank = "";

  /** If character is in a free company, check the unsorted list for their spot, then get their rank. */
  for (var i = 0; i < character.FreeCompanyMembers.length; i++) {

    if (character.FreeCompanyMembers[i].ID == character.ID) {
      /** Found it. Update their rank */
      FCRank = character.FreeCompanyMembers[i].Rank
    }
  }

  return FCRank;
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