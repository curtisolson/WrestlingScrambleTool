var rosterWrestlerID          = 0;
var rosterSchool              = 1;
var rosterFName               = 2;
var rosterLName               = 3;
var rosterGrade               = 4;
var rosterWeight              = 5;
var rosterExp                 = 6;
var rosterGender              = 7;
var rosterMatchCount          = 8;
var rosterBoutNumber          = 9;
var rosterMatchOpponentID     = 10;
var rosterMatchOpponentSchool = 11;
var rosterMatchOpponentShort  = 12;
var rosterExpPercent          = 13;
var rosterExpSortPercent      = 14;
var rosterRound               = 15;
var rosterMatchNumber         = 16;
var rosterLastPrintedBout     = 17;
var matchSet            = 0;
var matchNumber         = 1;
var matchSchool1        = 2;
var matchFName1         = 3;
var matchLName1         = 4;
var matchGrade1         = 5;
var matchWeight1        = 6;
var matchExp1           = 7;
var matchGender1        = 8;
var matchSchool2        = 9;
var matchFName2         = 10;
var matchLName2         = 11;
var matchGrade2         = 12;
var matchWeight2        = 13;
var matchExp2           = 14;
var matchGender2        = 15;
var matchWeightDiff     = 16;
var matchWeightPercent  = 17;
var matchExpPercent     = 18;
var matchExpSortPercent = 19;
var matchWrestlerID1    = 20;
var matchWrestlerID2    = 21;
var matchBoutNumber     = 22;
var teamSchool = 0;
var teamFName  = 1;
var teamLName  = 2;
var teamGrade  = 3;
var teamWeight = 4;
var teamExp    = 5;
var teamActive = 6;
var teamGender = 7;
var settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
var settingsData  = settings.getDataRange().getValues();
var gAgeFactor = getSetting("Age Factor", .05);
var gExperienceFactor = getSetting("Experience Factor", .05);
var gWeightLimit = getSetting("Allowed Weight Difference", .25);
var gWAELimit = getSetting("Allowed WAE Difference", .25);
var gPrintColor = getSetting("Print Color", "color");
var scrambleNotes = '';


function getSetting(settingId, defaultValue) {
  for (i=0; i<settingsData.length; i++) {
    if (settingsData[i][0] === settingId) {
      return settingsData[i][1];
    }
  }
  return defaultValue; 
}

function createMatchUps() {
//  Logger.clear();
  Logger.log("Start createMatchUps");
  //Get sheet references.
  var mainSheet      = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var schools        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Schools");
  var matches        = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CandidateMatches");
  var team1          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team1");
  var team2          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team2");
  var team3          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team3");
  var team4          = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team4");
  var boutSheets     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BoutSheets");
//  var boutSummary    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("BoutSummary");
  var combinedRoster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CombinedRoster");
  var schoolData = schools.getDataRange().getValues();
  var combinedRosterData = combinedRoster.getDataRange().getValues();
  var combinedRosterData = combinedRoster.getDataRange().getValues();
  updateRunStatus('Starting', mainSheet, init=true);
     
  qaPassed = validateData(mainSheet, team1, team2, team3, team4);
  
 
  combinedRosterData = combineTeamRosters(mainSheet, scrambleNotes);
//  combinedRosterData.sort(compareWeight);
  
  var counter = 0;
 
  var totalMatchCandidateData = createMatchCandidates(combinedRosterData, mainSheet);

  var matchData = [];
  var matchNumber = 0;

  for (var i = 0; i < totalMatchCandidateData.length; i++) {

    var wrestler1 = lookupWrestler(totalMatchCandidateData[i][0], combinedRosterData);      
    var wrestler2 = lookupWrestler(totalMatchCandidateData[i][1], combinedRosterData);
    if (wrestler1[rosterWeight] <= wrestler2[rosterWeight]) {
      var smallWrestler = wrestler1;        
      var bigWrestler = wrestler2;
    } else {
      var smallWrestler = wrestler2;        
      var bigWrestler = wrestler1;
    }
    var weightDiff = bigWrestler[rosterWeight]-smallWrestler[rosterWeight];
    var weightDiffPercent = weightDiff/smallWrestler[rosterWeight];
    var expDiffPercent = 1-(Math.abs((bigWrestler[rosterGrade]-smallWrestler[rosterGrade])*gAgeFactor+
                                       (bigWrestler[rosterExp]-smallWrestler[rosterExp])*gExperienceFactor+
                             weightDiffPercent));    
    
    if (smallWrestler[rosterSchool] === bigWrestler[rosterSchool]) {
      var expDiffSortPercent = '0-';
    } else {
      var expDiffSortPercent = '1-';
    }
    if (smallWrestler[rosterGender] !== bigWrestler[rosterGender]) {
      expDiffSortPercent += '0-';
    } else {
      expDiffSortPercent += '1-';
    }
    expDiffSortPercent += Math.floor(expDiffPercent*20).toString()+'-';
    expDiffSortPercent += 5-Math.abs(smallWrestler[rosterGrade] - bigWrestler[rosterGrade]).toString()+'-';
    expDiffSortPercent += expDiffPercent.toString();
        
    if (expDiffPercent >= 1-gWAELimit && weightDiffPercent <= gWeightLimit) {
      matchNumber++;
      matchData.push([' ', matchNumber, smallWrestler[rosterSchool], smallWrestler[rosterFName], smallWrestler[rosterLName], smallWrestler[rosterGrade], smallWrestler[rosterWeight], smallWrestler[rosterExp], smallWrestler[rosterGender], bigWrestler[rosterSchool], bigWrestler[rosterFName], bigWrestler[rosterLName], bigWrestler[rosterGrade], bigWrestler[rosterWeight], bigWrestler[rosterExp], bigWrestler[rosterGender], weightDiff, weightDiffPercent, expDiffPercent, expDiffSortPercent, smallWrestler[rosterWrestlerID], bigWrestler[rosterWrestlerID], '']);
      counter++;
    }
  }
  updateRunStatus('Found '+counter.toString()+' best matches using weight, age and experience.', mainSheet);
  
  // write candidate matches  
  writeMatchesSheet(matches, matchData);
  // Get match data after it has been sorted. The sort could be done in js a litte more efficiently.
  var matchData = matches.getDataRange().getValues();
   // Remove header row
  matchData.splice(0, 1);

  updateRunStatus('Choosing bouts from candidate matches.', mainSheet);
  var returnArray = pickMatchUps2(combinedRosterData, matchData, boutSheets, mainSheet);
  combinedRosterData = returnArray[0];
  matchData = returnArray[1];
  // write selected matches
  writeMatchesSheet(matches, matchData);

  //  create Bout Sheet;
//  combinedRosterData=createBoutSheets(combinedRosterData, matchData, boutSheets, mainSheet);

  // Update Roster with new columns
//  combinedRosterData.sort(compareSchoolGradeWeight);

  writeCombinedRosterSheet(combinedRoster, combinedRosterData, schoolData);


  getMatchStats(combinedRosterData, matchData, mainSheet);


  updateRunStatus('Run complete', mainSheet);
  Logger.log("End createMatchUps");
}

function validateData(mainSheet, team1, team2, team3, team4) {
  Logger.log("Start validateData");
  Logger.log("End validateData");

}

function getMatchStats(combinedRosterData, matchData, mainSheet) {
  Logger.log("Start getMatchStats");
  
/*  
var matchSet            = 0;
var matchNumber         = 1;
var matchSchool1        = 2;
var matchFName1         = 3;
var matchLName1         = 4;
var matchGrade1         = 5;
var matchWeight1        = 6;
var matchExp1           = 7;
var matchGender1        = 8;
var matchSchool2        = 9;
var matchFName2         = 10;
var matchLName2         = 11;
var matchGrade2         = 12;
var matchWeight2        = 13;
var matchExp2           = 14;
var matchGender2        = 15;
var matchWeightDiff     = 16;
var matchWeightPercent  = 17;
var matchExpPercent     = 18;
var matchExpSortPercent = 19;
var matchWrestlerID1    = 20;
var matchWrestlerID2    = 21;

*/
  
  var lastRound1Bout=0;
  var rosterBoutNumberSorted;
  var wrestlesTwice=0;
  var wrestlesTwiceArray=[];
  var wrestlesThrice=0;
  var wrestlesThriceArray=[];
  var round1Count=0;
  var weight1=0, weight2=0, weight3=0, weight4=0, weight5=0;
  var gradeSame=0, gradeDiff1=0, gradeDiff2=0;
  var genderSame=0, genderDiff=0;
  var schoolSame=0, schoolDiff=0;

  for (var i = 0; i < combinedRosterData.length; i++) {
    rosterBoutNumberSorted = combinedRosterData[i][rosterBoutNumber];
    rosterBoutNumberSorted.sort(compareNumbers);
    if (rosterBoutNumberSorted[0] > lastRound1Bout) {
      lastRound1Bout = rosterBoutNumberSorted[0];
    }
  }
  
  updateRunStatus('Everyone has had a match after bout ' + lastRound1Bout, mainSheet);

  for (var i = 0; i < combinedRosterData.length; i++) {
      round1Count = 0;    
    for (var ii = 0; ii < combinedRosterData[i][rosterBoutNumber].length; ii++) {
      if (combinedRosterData[i][rosterBoutNumber][ii] <= lastRound1Bout) { 
        round1Count++;
      } 
    }
    if (round1Count === 2) {
      wrestlesTwice++;
      wrestlesTwiceArray.push(combinedRosterData[i][rosterSchool] + " " + combinedRosterData[i][rosterFName] + " " + combinedRosterData[i][rosterLName]);
    } else if (round1Count === 3) {
      wrestlesThrice++;
      wrestlesThriceArray.push(combinedRosterData[i][rosterSchool] + " " + combinedRosterData[i][rosterFName] + " " + combinedRosterData[i][rosterLName]);
    }
  }
  updateRunStatus('Wrestlest twice: ' + wrestlesTwice, mainSheet);
  updateRunStatus('Wrestlest twice: ' + wrestlesTwiceArray, mainSheet);
  updateRunStatus('Wrestlest thrice: ' + wrestlesThrice, mainSheet);
  updateRunStatus('Wrestlest thrice: ' + wrestlesThriceArray, mainSheet);
  
  for (var i = 0; i < matchData.length; i++) {
    switch (Math.abs(matchData[i][matchGrade1] - matchData[i][matchGrade2])) {
      case 0:
        gradeSame++;
        break;
      case 1:
        gradeDiff1++;
        break;
      case 2:
        gradeDiff2++;
        break;
    }
    switch (matchData[i][matchGender1] === matchData[i][matchGender2]) {
      case true:
        genderSame++;
        break;
      case false:
        genderDiff++;
        break;
    }
    switch (matchData[i][matchSchool1] === matchData[i][matchSchool2]) {
      case true:
        schoolSame++;
        break;
      case false:
        schoolDiff++;
        break;
    }
    switch (Math.ceil(matchData[i][matchWeightPercent]*100/5)) {
      case 1:
        weight1++;
        break;
      case 2:
        weight2++;
        break;
      case 3:
        weight3++;
        break;
      case 4:
        weight4++;
        break;
      case 5:
        weight5++;
        break;
    }
  }

  updateRunStatus('WEIGHT', mainSheet);
  updateRunStatus('  0-5%: ' + weight1, mainSheet);
  updateRunStatus('  6-10%: ' + weight2, mainSheet);
  updateRunStatus('  11-25%: ' + weight3, mainSheet);
  updateRunStatus('  16-20%: ' + weight4, mainSheet);
  updateRunStatus('  21-25%: ' + weight5, mainSheet);
  updateRunStatus('GENDER Same: ' + genderSame + ' Diff: ' + genderDiff, mainSheet);
  updateRunStatus('GRADE Same: ' + gradeSame + ' Diff (1yr): ' + gradeDiff1 + ' Diff (2yr): ' + gradeDiff2, mainSheet);
  updateRunStatus('SCHOOL Diff: ' + schoolDiff + ' Same: ' + schoolSame, mainSheet);

  
  Logger.log("End getMatchStats");
}

function updateRunStatus(statusText, mainSheet, init) {
//  Logger.log("Start updateRunStatus");
  var statusRange = mainSheet.getRange(4, 5);
  if (init === true) { 
    statusRange.setValue('');
  }

  var statusLog = statusRange.getValues();
  var logDate = Date.now();
  
  statusLog += statusText+'\n';
  statusRange.setValue(statusLog);
//  Logger.log("End updateRunStatus");

}

function combineTeamRosters(mainSheet) {
  Logger.log("Start combineTeamRosters");

  var combinedRosterData = [];
  var teamName = mainSheet.getRange(4, 2).getValue();
  if (typeof teamName === "string" && teamName.length > 1) {
    Logger.log(teamName);
    var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team1");
    combinedRosterData = getTeamRoster(combinedRosterData, teamSheet);
  }
  var teamName = mainSheet.getRange(5, 2).getValue();
  if (typeof teamName === "string" && teamName.length > 1) {
    Logger.log(teamName);
    var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team2");
    combinedRosterData = getTeamRoster(combinedRosterData, teamSheet);
  }
  var teamName = mainSheet.getRange(6, 2).getValue();
  if (typeof teamName === "string" && teamName.length > 1) {
    Logger.log(teamName);
    var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team3");
    combinedRosterData = getTeamRoster(combinedRosterData, teamSheet);
  }
  var teamName = mainSheet.getRange(7, 2).getValue();
  if (typeof teamName === "string" && teamName.length > 1) {
    Logger.log(teamName);
    var teamSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Team4");
    combinedRosterData = getTeamRoster(combinedRosterData, teamSheet);
  }

  Logger.log("End combineTeamRosters");

  return combinedRosterData;
}  

function getTeamRoster(combinedRosterData, teamSheet) {
  var teamSheetData = teamSheet.getDataRange().getValues();
  var wrestlerID = combinedRosterData.length;

  for (var i = 1; i < teamSheetData.length; i++) {
    // i=1 will skip header rows
    
    
    if (teamSheetData[i][teamFName].length > 1 && teamSheetData[i][teamLName].length > 1) {

      if (teamSheetData[i][teamActive] !== 'Yes') {
        scrambleNotes += teamSheetData[i][teamSchool] + ': ';
        scrambleNotes += teamSheetData[i][teamFName] + ' ' + teamSheetData[i][teamLName] + ' skipped. Inactive.';
        scrambleNotes += '\n';    
      } else if (isNaN(teamSheetData[i][teamGrade]) || teamSheetData[i][teamGrade] < 5 || teamSheetData[i][teamGrade] > 8) {
        scrambleNotes += teamSheetData[i][teamSchool] + ': ';
        scrambleNotes += teamSheetData[i][teamFName] + ' ' + teamSheetData[i][teamLName] + ' skipped. No valid grade.';
        scrambleNotes += '\n';    
      } else if (isNaN(teamSheetData[i][teamWeight]) || teamSheetData[i][teamWeight] < 50 || teamSheetData[i][teamWeight] > 300) {
        scrambleNotes += teamSheetData[i][teamSchool] + ': ';
        scrambleNotes += teamSheetData[i][teamFName] + ' ' + teamSheetData[i][teamLName] + ' skipped. No valid weight.';
        scrambleNotes += '\n';    
      } else if (isNaN(teamSheetData[i][teamExp]) || teamSheetData[i][teamExp] < 1 || teamSheetData[i][teamExp] > 3) {
        scrambleNotes += teamSheetData[i][teamSchool] + ': ';
        scrambleNotes += teamSheetData[i][teamFName] + ' ' + teamSheetData[i][teamLName] + ' skipped. No valid experience.';
        scrambleNotes += '\n';    
/*
      } else if (!(teamSheetData[i][teamGender] === 'M' || teamSheetData[i][teamGender] === 'F')) {
        scrambleNotes += teamSheetData[i][teamSchool] + ': ';
        scrambleNotes += teamSheetData[i][teamFName] + ' ' + teamSheetData[i][teamLName] + ' skipped. Gender not M or F.';
        scrambleNotes += '\n';    
*/
      } else {
        wrestlerID++;
        combinedRosterData.push([
          wrestlerID,
          teamSheetData[i][teamSchool], 
          teamSheetData[i][teamFName], 
          teamSheetData[i][teamLName], 
          teamSheetData[i][teamGrade], 
          teamSheetData[i][teamWeight], 
          teamSheetData[i][teamExp],
          teamSheetData[i][teamGender],
          0, // match count
          [], // bout numbers
          [], // opponents IDs
          [], // opponents
          [], // opponents (short)
          [], // WAE Percent
          [], // WAE Percent for sorting
          [], // matched on round
          [], // match numbers
          -100 // last printed bout
          ])
      }
    }
  }
  return combinedRosterData;
}  

function createBoutSheets(combinedRosterData, matchData, boutSheets, mainSheet) {
  Logger.log("Start createBoutSheets");
//  var boutTracker = [];
  var matchTracker = [];
//  var recentWrestlerTracker = [];
  var deferredMatches = [];
  var queuedMatches = [];
  var wrestlerText1;
  var wrestlerText2;
  var schoolTextShort1;
  var schoolTextShort2;
  var wrestlerTextShort;
  var boutSheetData = [];
  var boutNumber = 0;
  var minRestMatches = Math.floor(Math.sqrt(combinedRosterData.length));
  if (minRestMatches > 10) {
    minRestMatches = 10;
  }
  var teams = mainSheet.getRange(4, 2, 4, 1).getValues();
  var matchByRound = [];
  var matchByRoundSorted = [];
  var matchList = [];
  var matchedWrestlerList = [];
  var skippedMatchList = [];
  var boutsPerPage = 10;
  var matchRow = [];
  var loopStart=0;
  var loopIncrement=1;
  
  for (var i = 0; i < 3; i++) {
    for (var ii = 0; ii < combinedRosterData.length; ii++) {
      // if match found and match hasn't already been added to bout list from opponent
      if (combinedRosterData[ii][rosterMatchCount] >= i+1 &&
          matchTracker.indexOf(combinedRosterData[ii][rosterMatchNumber][i]) === -1) {
        matchByRound.push([combinedRosterData[ii][rosterMatchNumber][i], combinedRosterData[ii][rosterExpSortPercent][i]]);
        matchTracker.push(combinedRosterData[ii][rosterMatchNumber][i]);
      }
    }
//    matchByRoundSorted = matchByRound.sort(compareMatchByRound);
    matchList=matchList.concat(matchByRound);
    matchByRound = [];
/*
    matchList.push(matchByRound);
    matchByRound = [];
*/
  }

  while (matchList.length > 0) {
    for (var ii = 0; ii < matchList.length; ii++) {
      for (var iii = 0; iii < matchData.length; iii++) {

        if (matchList[ii][0] === matchData[iii][matchNumber]) {
          matchRow = matchData[iii];
        }
      }
      if (matchedWrestlerList.indexOf(matchRow[matchWrestlerID1]) === -1 &&
          matchedWrestlerList.indexOf(matchRow[matchWrestlerID2]) === -1) {
        matchedWrestlerList.push(matchRow[matchWrestlerID1]);
        matchedWrestlerList.push(matchRow[matchWrestlerID2]);

        boutNumber++;
        
        [wrestlerText1, wrestlerTextShort2, schoolTextShort1] = formatWrestlerText(combinedRosterData[matchRow[matchWrestlerID1]-1]);
        [wrestlerText2, wrestlerTextShort2, schoolTextShort2] = formatWrestlerText(combinedRosterData[matchRow[matchWrestlerID2]-1]);
        boutSheetData.push([boutNumber, wrestlerText1]);
        boutSheetData.push(['', wrestlerText2]);
                                
        for (var iii = 0; iii < combinedRosterData[matchRow[matchWrestlerID1]-1][rosterMatchCount]; iii++) {
          if (combinedRosterData[matchRow[matchWrestlerID1]-1][rosterMatchNumber][iii] === matchRow[matchNumber]) {
            combinedRosterData[matchRow[matchWrestlerID1]-1][rosterBoutNumber][iii] = boutNumber;
          }
        }
        for (var iii = 0; iii < combinedRosterData[matchRow[matchWrestlerID2]-1][rosterMatchCount]; iii++) {
          if (combinedRosterData[matchRow[matchWrestlerID2]-1][rosterMatchNumber][iii] === matchRow[matchNumber]) {
            combinedRosterData[matchRow[matchWrestlerID2]-1][rosterBoutNumber][iii] = boutNumber;
          }
        }
      } else {
        skippedMatchList.push(matchList[ii]);
      }
    }
    // reverse to the to alternate top down and bottum up
    matchList = skippedMatchList.reverse();
    skippedMatchList = [];
    matchedWrestlerList = [];
  }
  
  
/*
  for (var i = 0; i < matchList.length; i++) {
    deferredMatches.push(matchList[i][0]);
    // try to set deferred matches
    queuedMatches = deferredMatches;
    deferredMatches = [];
    for (var ii = 0; ii < queuedMatches.length; ii++) {
      for (var iii = 0; iii < matchData.length; iii++) {
        if (matchData[iii][matchNumber] === queuedMatches[ii]) {
          // check for min rest or if at end of list, just set match
          if ((boutNumber - combinedRosterData[matchData[iii][matchWrestlerID1]-1][rosterLastPrintedBout] >= minRestMatches &&
              boutNumber - combinedRosterData[matchData[iii][matchWrestlerID2]-1][rosterLastPrintedBout] >= minRestMatches) ||
              i + minRestMatches >= matchList.length) {
            boutNumber++;
            
            [wrestlerText1, wrestlerTextShort2, schoolTextShort1] = formatWrestlerText(combinedRosterData[matchData[iii][matchWrestlerID1]-1]);
            [wrestlerText2, wrestlerTextShort2, schoolTextShort2] = formatWrestlerText(combinedRosterData[matchData[iii][matchWrestlerID2]-1]);
            boutSheetData.push([boutNumber, wrestlerText1]);
            boutSheetData.push(['', wrestlerText2]);
                                    
            combinedRosterData[matchData[iii][matchWrestlerID1]-1][rosterLastPrintedBout] = boutNumber;
            combinedRosterData[matchData[iii][matchWrestlerID2]-1][rosterLastPrintedBout] = boutNumber;
            for (var iiii = 0; iiii < combinedRosterData[matchData[iii][matchWrestlerID1]-1][rosterMatchCount]; iiii++) {
              if (combinedRosterData[matchData[iii][matchWrestlerID1]-1][rosterMatchNumber][iiii] === matchData[iii][matchNumber]) {
                combinedRosterData[matchData[iii][matchWrestlerID1]-1][rosterBoutNumber][iiii] = boutNumber;
              }
            }
            for (var iiii = 0; iiii < combinedRosterData[matchData[iii][matchWrestlerID2]-1][rosterMatchCount]; iiii++) {
              if (combinedRosterData[matchData[iii][matchWrestlerID2]-1][rosterMatchNumber][iiii] === matchData[iii][matchNumber]) {
                combinedRosterData[matchData[iii][matchWrestlerID2]-1][rosterBoutNumber][iiii] = boutNumber;
              }
            }

          } else {
            deferredMatches.push(queuedMatches[ii]);
          }  
        }
      }
    }
  }
*/

  // Write bout sheet data
  // Clear sheet but keep header
  var boutSheetRange = boutSheets.getRange(1, 1, 1, 30);
  var boutSheetHeader = boutSheetRange.getValues();
  boutSheets.clear({ formatOnly: false, contentsOnly: true });
  boutSheetRange.setValues(boutSheetHeader);
    
  var boutSheetRange = boutSheets.getRange(2, 1, boutSheetData.length, 2);
  boutSheetRange.setValues(boutSheetData);
  var boutSheetRange = boutSheets.getRange(2, 1, 400, 2);
  boutSheets.unhideRow(boutSheetRange);
  
  // visible area should be header, all bouts + 1 page of blank bouts
  var rowsToPrint = ((Math.ceil(boutNumber/boutsPerPage)+1)*2*boutsPerPage)+2;
  boutSheets.hideRows(rowsToPrint, 402-rowsToPrint);
  
  updateRunStatus('Created '+boutNumber.toString()+' bout sheets.', mainSheet);
  Logger.log("End createBoutSheets");
  return combinedRosterData;
}

function writeMatchesSheet(matches, matchData) {
  Logger.log("Start writeMatchesSheet");
  matches.clear({ formatOnly: false, contentsOnly: true });
  var matchRange = matches.getRange(1, 1, matchData.length, 23);
  matchRange.setValues(matchData);
  matches.sort(matchExpSortPercent+1, false);
  matches.insertRows(1);
  var header = matches.getRange("A1:W1");
  header.setValues([[' ', '#','School', 'First Name', 'Last Name', 'Grade', 'Weight', 'Experience', 'Gender', 'School', 'First Name', 'Last Name', 'Grade', 'Weight', 'Experience', 'Gender', 'Weight Diff', 'Weight Diff %', 'Exp Diff %', 'Exp Sort', 'Wrestler ID', 'Wrestler ID', 'Bout Number']]);
  Logger.log("End writeMatchesSheet");
}

function writeCombinedRosterSheet(combinedRoster, combinedRosterData, schoolData) {
  Logger.log("Start writeCombinedRosterSheet");

  var combinedRosterDataExploded = [];
  var combinedRosterRow = [];
  var combinedRosterSchool = '';
  for (var i = 0; i < combinedRosterData.length; i++) {
    // Push school as sub header
    if (combinedRosterSchool !== combinedRosterData[i][rosterSchool]) {
      combinedRosterDataExploded.push([combinedRosterData[i][rosterSchool], '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
      combinedRosterSchool = combinedRosterData[i][rosterSchool];
    }
    
    if (combinedRosterData[i][rosterMatchCount] === 0 ) {
      scrambleNotes += combinedRosterData[i][rosterSchool] + ': ';
      scrambleNotes += combinedRosterData[i][rosterFName] + ' ' + combinedRosterData[i][rosterLName] + ' has no matches.';
      scrambleNotes += '\n';    
    }
    
    combinedRosterRow = [
      combinedRosterData[i][rosterWrestlerID],
      combinedRosterData[i][rosterFName] + ' ' + combinedRosterData[i][rosterLName], 
      combinedRosterData[i][rosterGrade], 
      combinedRosterData[i][rosterWeight], 
      combinedRosterData[i][rosterExp],
      combinedRosterData[i][rosterGender]];
    
    for (var ii = 0; ii < 3; ii++) {
      if (combinedRosterData[i][rosterMatchCount] - 1 >= ii) {
        combinedRosterRow.push(combinedRosterData[i][rosterBoutNumber][ii]);
        combinedRosterRow.push(combinedRosterData[i][rosterMatchOpponentShort][ii]);
      } else {
        combinedRosterRow.push('');
        combinedRosterRow.push('');
      }
    }
    for (var ii = 0; ii < 3; ii++) {
      if (combinedRosterData[i][rosterMatchCount] - 1 >= ii) {
        combinedRosterRow.push(combinedRosterData[i][rosterExpPercent][ii]);
      } else {
        combinedRosterRow.push('');
      }
    }
    for (var ii = 0; ii < 3; ii++) {
      if (combinedRosterData[i][rosterMatchCount] - 1 >= ii) {
        combinedRosterRow.push(combinedRosterData[i][rosterRound][ii]);
      } else {
        combinedRosterRow.push('');
      }
    }
    for (var ii = 0; ii < 3; ii++) {
      if (combinedRosterData[i][rosterMatchCount] - 1 >= ii) {
        combinedRosterRow.push(combinedRosterData[i][rosterMatchNumber][ii]);
      } else {
        combinedRosterRow.push('');
      }
    }
    combinedRosterDataExploded.push(combinedRosterRow);
  }
  
  combinedRoster.clear();
 
  combinedRosterDataExploded.splice(0, 0, ['#', 'Name', 'Gr', 'Wt', 'Exp', 'Ge', 'B#', 'Opponent', 'B#', 'Opponent', 'B#', 'Opponent', 'WAE', 'WAE', 'WAE', 'Round1', 'Round2', 'Round3', 'Match1', 'Match2', 'Match3']);
  combinedRosterDataExploded.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  combinedRosterDataExploded.push(['SCRAMBLE NOTES', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  combinedRosterDataExploded.push(['--------------', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  combinedRosterDataExploded.push([scrambleNotes, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '']);
  var rosterRange = combinedRoster.getRange(1, 1, combinedRosterDataExploded.length, 21);
  rosterRange.setValues(combinedRosterDataExploded);
  //Format sheet
  combinedRoster.hideColumns(16, 6);
  var cells = combinedRoster.getRange("M:O");
  cells.setNumberFormat("0%");
  combinedRoster.setColumnWidth(1, 25);
  combinedRoster.setColumnWidth(2, 150);
  for (var i = 2; i < 21; i++) {
    combinedRoster.autoResizeColumn(i+1);
  }
  
  var colorsWAEGrid = [];
  var colorsWAERow = [];
  var colorsOpponentGrid = [];
  var colorsOpponentRow = [];
  var holdWAE;
  var holdOpponent;
  var currentSchool;
  var colorShade = (gPrintColor === 'color') ? '#DDDDDD' : '#DDDDDD';
  var colorScale1 = (gPrintColor === 'color') ? '#5BBA8B' : '#FFFFFF';
  var colorScale2 = (gPrintColor === 'color') ? '#ACC87C' : '#EFEFEF';
  var colorScale3 = (gPrintColor === 'color') ? '#FED56E' : '#CCCCCC';
  var colorScale4 = (gPrintColor === 'color') ? '#F2A972' : '#AEAEAE';
  var colorScale5 = (gPrintColor === 'color') ? '#E47C75' : '#8E8E8E';

  var horizontalAlignment = 'center';
  for (var i = 1; i < combinedRosterDataExploded.length; i++) {
    if (typeof combinedRosterDataExploded[i][0] === 'string') {
      if (combinedRosterDataExploded[i][0] === 'SCRAMBLE NOTES') {
        horizontalAlignment = 'left';
      }
      // School banner line
      currentSchool = combinedRosterDataExploded[i][0];
      cells = combinedRoster.getRange(i+1,1,1,21);
      cells.mergeAcross();
      cells.setHorizontalAlignment(horizontalAlignment);
      colorsOpponentGrid.push(['white', 'white', 'white', 'white', 'white', 'white']);
      colorsWAEGrid.push(['white', 'white', 'white']);
      for (var ii = 1; ii < schoolData.length; ii++) {
        if (combinedRosterDataExploded[i][0] === schoolData[ii][0]) {
          cells.setBackground(schoolData[ii][4]); 
          cells.setFontColor(schoolData[ii][3]);
        }
      }
    } else {
      for (var ii = 0; ii < 3; ii++) {
        // Color matches grid
        cells = combinedRoster.getRange(i+1,8+(2*ii),1,1);
        holdOpponent = cells.getValue();
        if (holdOpponent[0] === '[') {
          colorsOpponentRow.push(colorShade, colorShade);
        } else {
          colorsOpponentRow.push('white', 'white');
        }
        // Color WAE grid
        cells = combinedRoster.getRange(i+1,13+ii,1,1);
        holdWAE = cells.getValue();
        if (combinedRoster.getRange(i+1,6,1,1).getValue() === '') {
          // wrestler had no bouts
          colorsWAERow.push(colorScale5);
        } else if (typeof holdWAE !== 'number') {
          colorsWAERow.push('white');
        } else if (holdWAE >= .95) {
          colorsWAERow.push(colorScale1);
        } else if (holdWAE >= .90) {
          colorsWAERow.push(colorScale2);
        } else if (holdWAE >= .85) {
          colorsWAERow.push(colorScale3);
        } else if (holdWAE >= .80) {
          colorsWAERow.push(colorScale4);
        } else if (holdWAE >= .75) {
          colorsWAERow.push(colorScale5);
        }
      }
      colorsOpponentGrid.push(colorsOpponentRow);
      colorsOpponentRow=[];
      colorsWAEGrid.push(colorsWAERow);
      colorsWAERow=[];
    }
  }

  cells = combinedRoster.getRange(2,7,combinedRosterDataExploded.length-1,6);
  cells.setBackgrounds(colorsOpponentGrid);    
  cells = combinedRoster.getRange(2,13,combinedRosterDataExploded.length-1,3);
  cells.setBackgrounds(colorsWAEGrid);    

  Logger.log("End writeCombinedRosterSheet");
} 

function createMatchCandidates(combinedRosterData, mainSheet) {
  Logger.log("Start createMatchCandidates");
  updateRunStatus('Creating candidate matches for '+combinedRosterData.length.toString()+' wrestlers.', mainSheet);
  var totalMatchCandidates = [];
  var counter = 0;
  for (var i = 0; i < combinedRosterData.length; i++) {
    for (var ii = 0; ii < combinedRosterData.length; ii++) {
      if(combinedRosterData[i][rosterWrestlerID] < combinedRosterData[ii][rosterWrestlerID]){
        totalMatchCandidates.push([combinedRosterData[i][rosterWrestlerID], combinedRosterData[ii][rosterWrestlerID]]);
        counter++;
      }
    }
  }
  updateRunStatus(counter.toString()+' possible matches.', mainSheet);
  Logger.log("End createMatchCandidates");
  return totalMatchCandidates;
}

function formatWrestlerText(combinedRosterRow) {

  var wrestlerText = '';
  var wrestlerTextShort = '';
  var schoolTextShort = '';
  var upperLetter = /^[A-Z]+$/
  
  for (i=0; i<combinedRosterRow[rosterSchool].length; i++) {
    if (combinedRosterRow[rosterSchool][i].match(upperLetter)) {
      schoolTextShort+=combinedRosterRow[rosterSchool][i];
    }
  }
  wrestlerText += '('+combinedRosterRow[rosterSchool]+') ';
  wrestlerText += combinedRosterRow[rosterFName]+' '+combinedRosterRow[rosterLName]+'\n';
  wrestlerText += combinedRosterRow[rosterGrade].toString()+'/'+combinedRosterRow[rosterWeight].toString()+'/'+combinedRosterRow[rosterExp].toString()+'/'+combinedRosterRow[rosterGender];

  wrestlerTextShort += '('+schoolTextShort+') ';
  wrestlerTextShort += combinedRosterRow[rosterFName][0]+' '+combinedRosterRow[rosterLName]+' ';
  wrestlerTextShort += combinedRosterRow[rosterGrade].toString()+'/'+combinedRosterRow[rosterWeight].toString()+'/'+combinedRosterRow[rosterExp].toString()+'/'+combinedRosterRow[rosterGender];

  return [wrestlerText, wrestlerTextShort, schoolTextShort];
}

  
function pickMatchUps2(combinedRosterData, matchData, boutSheets, mainSheet) {
  Logger.log("Start pickMatchUps2");
  
//  var wrestlerIds=[];
//    wrestlerIds.push(combinedRosterData[i][rosterWrestlerID]);
  var matchCounts=[[0,0],[0,1],[0,2],[1,1],[1,2],[2,2]];
  var wrestlerText;
  var schoolTextShort;
  var wrestlerTextShort;
  var opponentText;
  var opponentSchoolTextShort;
  var opponentTextShort;
  var boutSheetData = [];
  var boutNumber = 0;
  var pickRound = 'not needed';
  var rowsPerBout = 6;
  var boutsPerPage = 4;
  var wrestlerMatched = false;
  
  for (var i = 0; i < matchCounts.length; i++) {
    //Find wrestler with low match count
    for (var ii = 0; ii < combinedRosterData.length; ii++) {
      wrestlerMatched = false;
      if (matchCounts[i][0] === combinedRosterData[ii][rosterMatchCount]) {
        Logger.log('Found wrestler ' + combinedRosterData[ii][rosterWrestlerID] + ' ' + matchCounts[i][0]);
        // Find match for that wrestler
        for (var iii = 0; iii < matchData.length; iii++) {

          if ((combinedRosterData[ii][rosterWrestlerID] === matchData[iii][matchWrestlerID1] || 
               combinedRosterData[ii][rosterWrestlerID] === matchData[iii][matchWrestlerID2]) &&
               matchData[iii][matchSet] !== 'X') {
            Logger.log('Found match ' + matchData[iii][matchNumber]);
            //Find opponent with correct match count
            for (var iiii = 0; iiii < combinedRosterData.length; iiii++) {
              if (((combinedRosterData[ii][rosterWrestlerID] === matchData[iii][matchWrestlerID1] &&
                    combinedRosterData[iiii][rosterWrestlerID] === matchData[iii][matchWrestlerID2]) || 
                   (combinedRosterData[iiii][rosterWrestlerID] === matchData[iii][matchWrestlerID1] &&
                    combinedRosterData[ii][rosterWrestlerID] === matchData[iii][matchWrestlerID2]))  &&
                   matchCounts[i][1] === combinedRosterData[iiii][rosterMatchCount]) {
            Logger.log('Found opponent ' + combinedRosterData[iiii][rosterWrestlerID] + ' ' + matchCounts[i][1]);
                
                wrestlerMatched = true;
                boutNumber++;
                // Update match
                matchData[iii][matchBoutNumber] = boutNumber;
                matchData[iii][matchSet] = 'X';
                
                // combinedRosterData[ii] = Wrestler
                // combinedRosterData[iiii] = Opponent
                
                [wrestlerText, wrestlerTextShort, schoolTextShort] = formatWrestlerText(combinedRosterData[ii]);
                [opponentText, opponentTextShort, opponentSchoolTextShort] = formatWrestlerText(combinedRosterData[iiii]);
                //Update Wrestler
                combinedRosterData[ii][rosterMatchCount]++;
                combinedRosterData[ii][rosterMatchOpponentID].push(combinedRosterData[iiii][rosterWrestlerID]);
                combinedRosterData[ii][rosterMatchOpponentSchool].push(combinedRosterData[iiii][rosterSchool]);
                combinedRosterData[ii][rosterMatchOpponentShort].push(opponentTextShort);                     
                combinedRosterData[ii][rosterBoutNumber].push(boutNumber);
                combinedRosterData[ii][rosterMatchNumber].push(matchData[iii][matchNumber]);
                combinedRosterData[ii][rosterExpPercent].push(matchData[iii][matchExpPercent]);
                combinedRosterData[ii][rosterExpSortPercent].push(matchData[iii][matchExpSortPercent]);
                combinedRosterData[ii][rosterRound].push(pickRound);

                //Update Opponent
                combinedRosterData[iiii][rosterMatchCount]++;
                combinedRosterData[iiii][rosterMatchOpponentID].push(combinedRosterData[ii][rosterWrestlerID]);
                combinedRosterData[iiii][rosterMatchOpponentSchool].push(combinedRosterData[ii][rosterSchool]);
                combinedRosterData[iiii][rosterMatchOpponentShort].push(wrestlerTextShort);
                combinedRosterData[iiii][rosterBoutNumber].push(boutNumber);
                combinedRosterData[iiii][rosterMatchNumber].push(matchData[iii][matchNumber]);
                combinedRosterData[iiii][rosterExpPercent].push(matchData[iii][matchExpPercent]);
                combinedRosterData[iiii][rosterExpSortPercent].push(matchData[iii][matchExpSortPercent]);
                combinedRosterData[iiii][rosterRound].push(pickRound);

                boutSheetData.push(['B#', 'Name']);
                boutSheetData.push([boutNumber, wrestlerText]);
                boutSheetData.push(['','']);
                boutSheetData.push(['','']);
                boutSheetData.push(['', opponentTextShort]);
                boutSheetData.push(['','']);
                
                break;
              }
            } 
          }
          if (wrestlerMatched) break;
        }
      }
    }
  }

Logger.log(combinedRosterData);
  
  // Write bout sheet data
Logger.log('====');
  var boutSheetRange = boutSheets.getRange(1, 1, 960, 2).clearContent();
Logger.log('====');
  boutSheets.unhideRow(boutSheetRange);
Logger.log(boutSheetData.length);
  var boutSheetRange = boutSheets.getRange(1, 1, boutSheetData.length, 2);
Logger.log('====');
  boutSheetRange.setValues(boutSheetData);
Logger.log('====');
  
  // visible area should be header, all bouts + 1 page of blank bouts
  
  var blankBoutsOnLastPage = boutsPerPage - (boutNumber % boutsPerPage);
  // print all abouts + remainder of page + 1 blank page.
  var rowsToPrint = (boutNumber+blankBoutsOnLastPage+boutsPerPage)*rowsPerBout;
  boutSheets.hideRows(rowsToPrint, 960-rowsToPrint);
Logger.log('====');
  
  updateRunStatus('Created '+boutNumber.toString()+' bout sheets.', mainSheet);
Logger.log('====');

  Logger.log("End pickMatchUps2");
  return [combinedRosterData, matchData];
}
    
function pickMatchUps(combinedRosterData, matchData) {
  Logger.log("Start pickMatchUps");
  var opponentID;
  var opponentsOpponentID;
  // 2 loops. on first, get everyone at least one good first match. on next, make best matches.
  for (var i = 0; i < matchData.length; i++) {
    if (!(combinedRosterData[matchData[i][matchWrestlerID1]-1][rosterMatchCount] === 3 || 
          combinedRosterData[matchData[i][matchWrestlerID2]-1][rosterMatchCount] === 3) &&
      (combinedRosterData[matchData[i][matchWrestlerID1]-1][rosterMatchCount] === 0 || 
      combinedRosterData[matchData[i][matchWrestlerID2]-1][rosterMatchCount] === 0)) {

      combinedRosterData = setMatch(combinedRosterData, matchData[i], 0);
      matchData[i][matchSet] = 'X';
    }  
  }
  for (var i = 0; i < matchData.length; i++) {
    if (matchData[i][matchSet] !== 'X' &&
       (!(combinedRosterData[matchData[i][matchWrestlerID1]-1][rosterMatchCount] === 3 || 
          combinedRosterData[matchData[i][matchWrestlerID2]-1][rosterMatchCount] === 3))) {
      combinedRosterData = setMatch(combinedRosterData, matchData[i], 0);
      matchData[i][matchSet] = 'X';
    }  
  }

  Logger.log("End pickMatchUps");
  return [combinedRosterData, matchData];}


function swapBouts(combinedRosterData, bout1, bout2) {
  for (var i = 0; i < combinedRosterData.length; i++) {
    for (var ii = 0; ii < combinedRosterData[i][rosterMatchCount]; ii++) {
      if (combinedRosterData[i][rosterBoutNumber][ii] === bout1) {
        combinedRosterData[i][rosterBoutNumber][ii] = bout2;
      } else if (combinedRosterData[i][rosterBoutNumber][ii] === bout2) {
        combinedRosterData[i][rosterBoutNumber][ii] = bout1;
      }
    }
  }
}

function setMatch(combinedRosterData, matchDataRow, pickRound) {
//  Logger.log("Start setMatch");
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterMatchCount]++;
  [opponentText, opponentTextShort] = formatWrestlerText(combinedRosterData[matchDataRow[matchWrestlerID2]-1]);
  // id same school matches so they can be formatted
  if (matchDataRow[matchSchool1] === matchDataRow[matchSchool2]) {
    opponentTextShort=opponentTextShort.replace("(","[");
    opponentTextShort=opponentTextShort.replace(")","]");
  }
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterMatchOpponentID].push(matchWrestlerID2);
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterMatchOpponentSchool].push(matchDataRow[matchSchool2]);
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterMatchOpponentShort].push(opponentTextShort);
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterMatchNumber].push(matchDataRow[matchNumber]);
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterExpPercent].push(matchDataRow[matchExpPercent]);
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterExpSortPercent].push(matchDataRow[matchExpSortPercent]);
  combinedRosterData[matchDataRow[matchWrestlerID1]-1][rosterRound].push(pickRound);

  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterMatchCount]++;
  [opponentText, opponentTextShort] = formatWrestlerText(combinedRosterData[matchDataRow[matchWrestlerID1]-1]);
  // id same school matches so they can be formatted
  if (matchDataRow[matchSchool1] === matchDataRow[matchSchool2]) {
    opponentTextShort=opponentTextShort.replace("(","[");
    opponentTextShort=opponentTextShort.replace(")","]");
  }
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterMatchOpponentID].push(matchWrestlerID1);
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterMatchOpponentSchool].push(matchDataRow[matchSchool1]);
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterMatchOpponentShort].push(opponentTextShort);
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterMatchNumber].push(matchDataRow[matchNumber]);
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterExpPercent].push(matchDataRow[matchExpPercent]);
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterExpSortPercent].push(matchDataRow[matchExpSortPercent]);
  combinedRosterData[matchDataRow[matchWrestlerID2]-1][rosterRound].push(pickRound);

//  Logger.log("End setMatch");
  return combinedRosterData;
}      
    
function lookupWrestler(wrestlerID, combinedRosterData) {
//  Logger.log("Start lookupWrestler");
  for (var i = 0; i < combinedRosterData.length; i++) {
    if(wrestlerID === combinedRosterData[i][rosterWrestlerID]){
      return combinedRosterData[i]
    }
  } 
//  Logger.log("End lookupWrestler");
} 


//
//   Sort Callback Functions
//   
function compareNumbers(a, b) {
  return a - b;
}

function compareSchoolGradeWeight(a, b) {
  if ((a[rosterSchool] <  b[rosterSchool]) ||
      (a[rosterSchool] ===  b[rosterSchool] &&
       a[rosterGrade] <  b[rosterGrade]) ||
      (a[rosterSchool] ===  b[rosterSchool] &&
       a[rosterGrade] ===  b[rosterGrade]&&
       a[rosterWeight] <  b[rosterWeight])) {
    return -1;
  } else if ((a[rosterSchool] >  b[rosterSchool]) ||
      (a[rosterSchool] ===  b[rosterSchool] &&
       a[rosterGrade] >  b[rosterGrade]) ||
      (a[rosterSchool] ===  b[rosterSchool] &&
       a[rosterGrade] ===  b[rosterGrade]&&
       a[rosterWeight] >  b[rosterWeight])) {
    return 1;
  } else {
    return 0;
  }
}

function compareMatchCountWeight(a, b) {
  if ((a[rosterMatchCount] <  b[rosterMatchCount]) ||
      (a[rosterMatchCount] ===  b[rosterMatchCount] &&
       a[rosterWeight] <  b[rosterWeight])) {
    return -1;
  } else if ((a[rosterMatchCount] >  b[rosterMatchCount]) ||
     (a[rosterMatchCount] ===  b[rosterMatchCount] &&
      a[rosterWeight] >  b[rosterWeight])) {
    return 1;
  } else {
    return 0;
  }
}

function compareWeight(a, b) {
  if (a[rosterWeight] <  b[rosterWeight]) {
    return -1;
  } else if (a[rosterWeight] >  b[rosterWeight]) {
    return 1;
  } else {
    return 0;
  }
}

function compareMatchByRound(a, b) {
  if (a[0] === b[0]) {
    return 0;
  }
  return (a[1] < b[1]) ? 1 : -1;
}

