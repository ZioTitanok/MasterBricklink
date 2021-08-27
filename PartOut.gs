// Function: Download PartOut from Bricklink
function LoadPartOut(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var ConsumerKey = SheetSettings.getRange("B3").getValue();
  var ConsumerSecret = SheetSettings.getRange("B4").getValue();
  var TokenValue = SheetSettings.getRange("B5").getValue();
  var TokenSecret = SheetSettings.getRange("B6").getValue();
  
  var SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartOut');
  var PartOutMinRow = 4;
  var PartOutMaxRow = SheetPartOut.getMaxRows();  
  
  SheetPartOut.getRange(PartOutMinRow, 1, PartOutMaxRow, 8).clear({contentsOnly: true});

  // Data
  var SetNo = SheetPartOut.getRange("A2").getValue();
  var SetVar = SheetPartOut.getRange("B2").getValue()
  var BreakMinifigure = SheetPartOut.getRange("C2").getValue();
  if (BreakMinifigure == "YES"){
    BreakMinifigure = 'TRUE';
  } else if (BreakMinifigure == "NO"){
    BreakMinifigure = 'FALSE'
  }

  // API Request
  var Url = 'https://api.bricklink.com/api/store/v1' + '/items/set/' + SetNo + SetVar + '/subsets';
  var Options = {method: 'GET', contentType: 'application/json'};
  var Params = {
    break_minifigs: BreakMinifigure,
    break_subsets: 'TRUE'
  }; 
  
  urlFetch = OAuth1.withAccessToken(ConsumerKey, ConsumerSecret, TokenValue, TokenSecret);
  
  // Output 
  var OutputPartout = JSON.parse(urlFetch.fetch(Url, Params, Options));  
  var Partout = [];
  var i = 0;
  
  for (i in OutputPartout.data){
    Partout[i] = [i,
                  OutputPartout.data[i].entries[0].item.type,
                  OutputPartout.data[i].entries[0].item.no,
                  OutputPartout.data[i].entries[0].item.name,
                  OutputPartout.data[i].entries[0].item.category_id,
                  OutputPartout.data[i].entries[0].color_id,
                  OutputPartout.data[i].entries[0].quantity,
                  OutputPartout.data[i].match_no
                 ]    
  }

  SheetPartOut.getRange(PartOutMinRow, 1, Partout.length, 8).setValues(Partout);
  SheetPartOut.getRange(4, 1, Partout.length, 8).sort([2, 5, 3]);

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('PartOut', 'Download compled!', Ui.ButtonSet.OK);
}

// Function: Clear PartOut
function ClearPartOut(){
  var SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartOut');
  var PartOutMinRow = 4;
  var PartOutMaxRow = SheetPartOut.getMaxRows();  
  
  SheetPartOut.getRange(PartOutMinRow, 1, PartOutMaxRow, 8).clear({contentsOnly: true});

  // Ui
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('PartOut', 'PartOut is ready for new adventures!', Ui.ButtonSet.OK);
}