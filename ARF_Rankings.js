var ssa = SpreadsheetApp.getActiveSpreadsheet();
var resultsSheet = ssa.getSheetByName('Results');
var playerSheet = ssa.getSheetByName('Players');
//resultsSheet.getRange('A2:Z').clear();
var countriesSheet =  ssa.getSheetByName('Countries');
var compRange = countriesSheet.getRange('A2:B').getValues();
var comp_mults = {};
for(var v = 0; v < compRange.length; v++){
 comp_mults[compRange[v][0]] = compRange[v][1]
}
var countries = Object.keys(comp_mults)
var tournamentsSheet = ssa.getSheetByName('Tournaments');
var tournamentsRange = tournamentsSheet.getRange('A2:B').getValues();

//Capitalizes first letter of string and lowercases every other letter
function PROPER_CASE(str) {
  if (typeof str == "string") {

  str = str.toLowerCase();

  var arr = str.split(' ');

  for(var l = 0; l < arr.length; l++){
    arr[l] = arr[l].charAt(0).toUpperCase() + arr[l].slice(1);
  }

  return arr.join(' ');
  }
}

//Pulls in results data from all country's spreadsheets and sheets
function getData() {

  let t = 0;

  for(z = 0; z < countries.length; z++)
  {
    try {
      var files = DriveApp.getFilesByName(countries[z]);
      var file = files.next();
    }
    catch(err) {
      continue
    }
    var ss = SpreadsheetApp.open(file);
    SpreadsheetApp.setActiveSpreadsheet(ss);
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    var country = ss.getName();

    for(var i = 2; i < sheets.length; i++)
    {
      var tournament = sheets[i].getName();
      var nameSheet = ss.getSheetByName(tournament);

      try {
        if(tournamentsRange[t][0] == tournament && tournamentsRange[t][1] == countries[z]){
          t++;
          continue};
        } catch(err) {}

      t++;
      var nameRange = nameSheet.getDataRange();
      var nameValues = nameRange.getValues();
      var date = nameValues[0][1];
      var division = ['Open', 'Women', 'Co-ed'];
      var open_results = [];
      var women_results = [];
      var coed_results = [];

      for(var y = 4; y < nameValues.length; y++)
      {
        var open = [];
        var women = [];
        var coed = [];

        if(typeof(nameValues[y][1]) == 'number'){
          var open_full_name = nameValues[y].slice(2,4).join(' ');
        } else {
          var open_full_name = null;
        }
        if(typeof(nameValues[y][11]) == 'number'){
          var women_full_name = nameValues[y].slice(12,14).join(' ');
        } else {
          var women_full_name = null;
        }
        if(typeof(nameValues[y][21]) == 'number'){
          var coed_full_name = nameValues[y].slice(22,24).join(' ');
        } else {
          var coed_full_name = null;
        }
        if(open_full_name != null){
        open.push(nameValues[y][0], date, country, tournament, division[0], nameValues[y][1], PROPER_CASE(open_full_name), PROPER_CASE(nameValues[y][4]), PROPER_CASE(nameValues[y][5]), PROPER_CASE(nameValues[y][6]), nameValues[y][7], nameValues[y][8], nameValues[y][9]);
        open_results.push(open);
        }
        if(women_full_name != null){
        women.push(nameValues[y][10], date, country, tournament, division[1], nameValues[y][11], PROPER_CASE(women_full_name), PROPER_CASE(nameValues[y][14]), PROPER_CASE(nameValues[y][15]), PROPER_CASE(nameValues[y][16]), nameValues[y][17], nameValues[y][18], nameValues[y][19])
        women_results.push(women);
        }
        if(coed_full_name != null){
        coed.push(nameValues[y][20], date, country, tournament, division[2], nameValues[y][21], PROPER_CASE(coed_full_name), PROPER_CASE(nameValues[y][24]), PROPER_CASE(nameValues[y][25]), PROPER_CASE(nameValues[y][26]), nameValues[y][27], nameValues[y][28], nameValues[y][29]);
        coed_results.push(coed);
        }
      }
      for(var m = 0; m < open_results.length; m++)
      {
        resultsSheet.appendRow(open_results[m]);
      }
      for(var w = 0; w < women_results.length; w++)
      {
        resultsSheet.appendRow(women_results[w]);
      }
      for(var c = 0; c < coed_results.length; c++)
      {
        resultsSheet.appendRow(coed_results[c]);
      }
      timestampRange = nameSheet.getRange(1,5);
      timestampRange.setValue("INGESTED")
      timestampRange.offset(0,1).setValue(new Date())
    }
  }
}

//Calculates and sorts ranking data in rankings sheets
function rankingSort() {

  var resultsRange = resultsSheet.getRange('A2:M');
  var resultsValues = resultsRange.getValues();
  var playerIdRange = playerSheet.getRange('A2:D');
  var playerIdRangeValues = playerIdRange.getValues();

  var maleRankSheet = ssa.getSheetByName('Male_Rankings');
  var femaleRankSheet = ssa.getSheetByName('Female_Rankings');
  try{
  maleRankSheet.deleteRows(2,maleRankSheet.getMaxRows()-1);  ;
  femaleRankSheet.deleteRows(2,femaleRankSheet.getMaxRows()-1);
  }
  catch(err){
    Logger.log("Rankings sheets have no rows to delete " + err.message);
  }
  var malePlayerRank = [];
  var femalePlayerRank = [];

  for(var i = 0; i < playerIdRangeValues.length; i++){
    var totalTourns = [];
    for (var j = 0; j < resultsValues.length; j++){
      if(playerIdRangeValues[i][0] == resultsValues[j][0]){
        totalTourns.push(resultsValues[j][12]);
      }
    }
    var points = totalTourns.sort(function(a, b){return b - a});
    if(points.length > 4){
      var x = parseInt(points.length*.6);
      var points = points.slice(0, x);
    }
    var points = points.reduce((partialSum, a) => partialSum + a, 0)
    if(playerIdRangeValues[i][2] == 'Male'){
      malePlayerRank.push('',playerIdRangeValues[i][1], playerIdRangeValues[i][2], playerIdRangeValues[i][3], points, playerIdRangeValues[i][0]);
      maleRankSheet.appendRow(malePlayerRank);
      malePlayerRank = [];
    }
    else if(playerIdRangeValues[i][2] == 'Female'){
      femalePlayerRank.push('',playerIdRangeValues[i][1], playerIdRangeValues[i][2], playerIdRangeValues[i][3], points, playerIdRangeValues[i][0]);
      femaleRankSheet.appendRow(femalePlayerRank);
      femalePlayerRank = [];
    }
    }
  var maleRankRange = maleRankSheet.getRange('A2:F');
  try{
    maleRankRange.sort(1);
    maleRankRange.setFontWeight("normal");
  }
  catch(err){
    Logger.log("No ranking data " + err.message);
  }
  var femaleRankRange = femaleRankSheet.getRange('A2:F');
  try{
    femaleRankRange.sort(1);
    femaleRankRange.setFontWeight("normal");
  }
  catch(err){
    Logger.log("No ranking data " + err.message);
  }
}

//Execute script
function MAIN(){

  getData();
  rankingSort();

}
