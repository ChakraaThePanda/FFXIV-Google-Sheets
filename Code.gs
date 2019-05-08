var _rosterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster'),
	_instructionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Instructions'),
    _instructionsAPIKeyRow = 4,
	_characterID = [], 
    _characterInfo = [], 
    _nombreDeRowsRegardes = _rosterSheet.getLastRow(),
    //Order un peu de la marde, l'API est fait comme ça. Dans l'ordre voulu (Tank, Heal, DPS, DoH, DoL). Les IDs des classes.
    _classOrder = [1,3,32,6,26,33,2,4,29,34,5,31,7,26,35,36,8,9,10,11,12,13,14,15,16,17,18], 
    _stateEnum = {"LOADING":1, "READY":2, "NOT_FOUND":3},
    _avatarColumn = 1,
    _nameColumn = 2,
    _lodestoneColumn = 3,
    _serverColumn = 4,
    _freeCompanyColumn = 5,
    _firstClassColumn = 8, 
    _firstCharacterScannedRow = 2
    _freeCompanyRankColumn = 6,
    _equippedilvlColumn = 7;

function fetchIDsFromSheet(){
  //_lodestoneColumn pour la column qui contient les Lodesone IDs
  for(var i = 0, j = _firstCharacterScannedRow; i < _nombreDeRowsRegardes; ++i, ++j){
    if(_rosterSheet.getRange(j,_lodestoneColumn).getDisplayValue() !== ""){
      _characterID[i] = _rosterSheet.getRange(j, _lodestoneColumn).getDisplayValue();
    } 
  }
  fetchCharacterInfos();
}

// http://xivapi.com/docs/Character/<lodestone_id> (States)
function fetchCharacterInfos() {
  var loading = false, error = false;
  for(var i = 0, row = _firstCharacterScannedRow; i < _characterID.length; ++i, ++row){
    _characterInfo[i] = JSON.parse(UrlFetchApp.fetch("https://xivapi.com/character/" + _characterID[i] + "?key=" + _instructionsSheet.getRange(_instructionsAPIKeyRow,1).getDisplayValue() + "&data=FC,FCM&extended=1").getContentText());
    
    if(_characterInfo[i].Info.Character.State === _stateEnum.LOADING){
      _rosterSheet.getRange(row, _nameColumn).setValue("Loading...");
    }
    
    if(_characterInfo[i].Info.Character.State === _stateEnum.NOT_FOUND){
      _rosterSheet.getRange(row, _nameColumn).setValue("Not Found");
    }
    
    if(_characterInfo[i].Info.Character.State === _stateEnum.READY){
      updateCharacter(_characterInfo[i], row);
    }
  }  
}

function updateCharacter(character, row){
  //Ajouter l'avatar du toon (_avatarColumn) + Un link vers full image
  _rosterSheet.getRange(row, _avatarColumn).setValue("=HYPERLINK(\"" + character.Character.Portrait + "\", IMAGE(\""+ character.Character.Avatar + "\"))");
  
  //Ajouter le nom du toon + Un link vers profile Lodestone
  _rosterSheet.getRange(row, _nameColumn).setValue("=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/character/" + character.Character.ID + "\", \"" + character.Character.Name + "\")");
  
  //Ajouter le serveur du toon
  _rosterSheet.getRange(row, _serverColumn).setValue(character.Character.Server);
  
  //Ajouter la FC du toon
  updateFreeCompany(character, row); 
  
  //Ajouter les classes du toon et le timestamp du last update
  updateClassJobsAndTime(character, row);
  
  //Ajouter Current Equipped class and ilvl
  updateCurrentClass(character, row);
  
}

function updateCurrentClass(character, row){
  var amountofslots = 0, //Nombre de slots d'equipped. S'ajuste au nombre que la personne wear.
      ilvl = 0,
      gear = character.Character.GearSet.Gear,
      keys = Object.keys(gear),
      mainHandIndex = 0,
      offhand = false; 
  for(var amountofslots; amountofslots < keys.length; ++amountofslots){
    if(keys[amountofslots].toString() !== "SoulCrystal" && keys[amountofslots].toString() !== "MainHand") //Ignore SoulCrystal dans total
       ilvl += gear[keys[amountofslots]].Item.LevelItem;
    if(keys[amountofslots].toString() == "MainHand")
      mainHandIndex = amountofslots;
    if(keys[amountofslots].toString() == "OffHand"){
      offhand = true;
    }
  }
  if(offhand) //Si on est une classe qui porte un offhand, on compte le main hand une seul fois. Peut RIP si exemple PLD avec pas d'OH d'equipped.
    ilvl += gear[keys[mainHandIndex]].Item.LevelItem;
  else
    ilvl += gear[keys[mainHandIndex]].Item.LevelItem * 2;
  
  var ilvlstring = '\n' + Math.floor(ilvl/13) + " ilvl"; //Build la string à afficher avec l'average ilvl arrondi à 2 decimales. 13 c'est le nombre de gear slots.
  
  var richText = SpreadsheetApp.newRichTextValue()
          .setText(character.Character.GearSet.Job.Abbreviation + ilvlstring)
          .setTextStyle(character.Character.GearSet.Job.Abbreviation.toString().length + 1, 
                        character.Character.GearSet.Job.Abbreviation.toString().length + ilvlstring.toString().length, 
                        SpreadsheetApp.newTextStyle()
                        .setFontSize(9)
                        .build()).build();;
          _rosterSheet.getRange(row, _equippedilvlColumn).setRichTextValue(richText);
}

function updateClassJobsAndTime(character, row){
  var classJobs = character.Character.ClassJobs;
  var keys  = Object.keys(classJobs);
  
  //Remplacer _firstClassColumn pour la column de départ de la première classe
  for(var j = 0, k = _firstClassColumn; j < keys.length; ++j, ++k){
    var found = false;
    
    //Ajouter les Class Levels
    for(var i = 0; i < keys.length && found === false; ++i){
      if(classJobs[keys[i]].Class.ID === _classOrder[j]){
        if(classJobs[keys[i]].Level === 0){ //Mettre des "-" au lieu de des 0 quand la personne n'as pas encore unlock sa classe.
          _rosterSheet.getRange(row, k).setValue("-");
        }
        else if(classJobs[keys[i]].ExpLevelTogo === 0){ //Si level max, on fait juste l'inscrire.
          _rosterSheet.getRange(row, k).setValue(classJobs[keys[i]].Level);
        }
        else{ //On met le pourcentage de levels non-max, et on le met plus petit que le level écrit.
          var level = classJobs[keys[i]].Level;
          var xpPercent = "\n(" + Math.round(classJobs[keys[i]].ExpLevel / classJobs[keys[i]].ExpLevelMax * 100) + "%)";
          var richText = SpreadsheetApp.newRichTextValue()
          .setText(classJobs[keys[i]].Level + xpPercent)
          .setTextStyle(level.toString().length + 1, level.toString().length + xpPercent.length, SpreadsheetApp.newTextStyle()
                        .setFontSize(7)
                        .build()).build();
          _rosterSheet.getRange(row, k).setRichTextValue(richText);
        }
        found = true;
      }
    }
  }
  //Ajouter l'Update Time
  var d = new Date(character.Character.ParseDate * 1000)
  _rosterSheet.getRange(row, k).setValue(d);
}

function updateFreeCompany(character, row){  
  if(character.Character.FreeCompanyId !== null){
    if(character.FreeCompany !== null){
      _rosterSheet.getRange(row, _freeCompanyColumn).setValue("=HYPERLINK(\"https://na.finalfantasyxiv.com/lodestone/freecompany/" + character.FreeCompany.ID.replace('i','') + "\", \"" + character.FreeCompany.Name + "\")");
	  for(var i = 0; i < character.FreeCompanyMembers.length; ++i)
        if(character.Character.Name == character.FreeCompanyMembers[i].Name)
          _rosterSheet.getRange(row, _freeCompanyRankColumn).setValue(character.FreeCompanyMembers[i].Rank);
	}
	else
      _rosterSheet.getRange(row, _freeCompanyColumn).setValue("Loading...");
  }
  else{
    _rosterSheet.getRange(row, _freeCompanyColumn).clearContent();
    _rosterSheet.getRange(row, _freeCompanyRankColumn).clearContent();
  }
}

// Fonction pour forcer les updates aux 6 heures. Appelée par TimeBased aux 6 heures.
function fetchUpdateIDsFromSheet(){
  //_lodestoneColumn pour la column qui contient les Lodesone IDs
  for(var i = 0, j = _firstCharacterScannedRow; i < _nombreDeRowsRegardes; ++i, ++j){
    if(_rosterSheet.getRange(j,_lodestoneColumn).getDisplayValue() !== ""){
      _characterID[i] = _rosterSheet.getRange(j, _lodestoneColumn).getDisplayValue();
    } 
  }
  for(var i = 0, row = _firstCharacterScannedRow; i < _characterID.length; ++i, ++row){
    UrlFetchApp.fetch("https://xivapi.com/character/" + _characterID[i] + "/update?key=" + _instructionsSheet.getRange(_instructionsAPIKeyRow,1).getDisplayValue()).getContentText();
  }  
}