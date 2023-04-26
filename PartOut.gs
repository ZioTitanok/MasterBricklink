// Constants: PartOut
const PartOutRowMin = 4;

// Function: Download PartOut from Bricklink
function LoadPartOut(){
  const SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartOut');
  const PartOutRowMax = SheetPartOut.getMaxRows();  
  
  SheetPartOut.getRange(PartOutRowMin, 1, PartOutRowMax, 8).clearContent();

  // Data
  const SetNo = SheetPartOut.getRange("A2").getValue();
  const SetVar = SheetPartOut.getRange("B2").getValue()
  var BreakMinifigure = SheetPartOut.getRange("C2").getValue();
  if (BreakMinifigure == "YES"){
    BreakMinifigure = 'TRUE';
  } else if (BreakMinifigure == "NO"){
    BreakMinifigure = 'FALSE'
  }

  // API Request & Output
  const {BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret} = GetSettings();
  const Url = `${BrickLinkBaseUrl}/items/set/${SetNo}${SetVar}/subsets`;
  const Params = {break_minifigs: BreakMinifigure, break_subsets: 'TRUE'}; 
  urlFetch = OAuth1.withAccessToken(BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret);

  const OutputPartout = JSON.parse(urlFetch.fetch(Url, Params, BrickLinkOptions));  
  const Partout = OutputPartout.data.map((Item, Index) => {
    return [
      Index+1,
      Item.entries[0].item.type,
      Item.entries[0].item.no,
      Item.entries[0].item.name,
      Item.entries[0].item.category_id,
      Item.entries[0].color_id,
      Item.entries[0].quantity,
      Item.match_no
    ];
  });

  SheetPartOut.getRange(PartOutRowMin, 1, Partout.length, 8).setValues(Partout).setNumberFormat('@STRING@');
  SheetPartOut.getRange(4, 1, Partout.length, 9).sort([2, 9, 4]);

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('PartOut', 'Download compled!', Ui.ButtonSet.OK);
}

// Function: Clear PartOut
function ClearPartOut(){
  const SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PartOut');
  const PartOutRowMax = SheetPartOut.getMaxRows();  
  
  SheetPartOut.getRange(PartOutRowMin, 1, PartOutRowMax, 8).clearContent();

  // Ui
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('PartOut', 'PartOut is ready for new adventures!', Ui.ButtonSet.OK);
}