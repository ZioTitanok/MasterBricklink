// Variables from Lab
var LabMinRow = 4;
var LabColumnItemType = 1;
var LabColumnItemNo = 2;
var LabColumnQty = 4;
var LabColumnCondition = 5;
var LabColumnCompleteness = 6;
var LabColumnStock = 7;
var LabColumnColorID = 8;
var LabColumQtyInventory = 11;
var LabColumnPrice = 15;
var LabColumnPriceMin = 16;
var LabColumnPriceAvg = 17;
var LabColumnPriceAvgQty = 18;
var LabColumnPriceMax = 19;
var LabColumnPriceLot = 20;
var LabColumnPriceQty = 21;
var LabColumnLotID = 22;
var LabColumnDescription = 27;
var LabColumnRemarks = 28;

// Function: Download Prices (Batch)
function LoadPricesBatch(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var ConsumerKey = SheetSettings.getRange("B3").getValue();
  var ConsumerSecret = SheetSettings.getRange("B4").getValue();
  var TokenValue = SheetSettings.getRange("B5").getValue();
  var TokenSecret = SheetSettings.getRange("B6").getValue();
  var LabActive = SheetSettings.getRange("B8").getValue();
  var MaxRow = SheetSettings.getRange("B9").getValue();
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  // Data
  var PriceType = SheetLab.getRange("M2").getValue();
  var PriceRegion = SheetLab.getRange("K2").getValue();
  if (PriceRegion=="Worldwide") PriceRegion="";
  var Selection = SheetLab.getDataRange();
  var Cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var StartingRow = Cell.getRow();
 
  if (StartingRow < LabMinRow) return;
  
  // For Loop
  for (var Row = StartingRow; Row <= StartingRow+MaxRow; Row++){
    var CellCode = Selection.getCell(Row,LabColumnItemNo).getValue();
    if (CellCode == ""){
      break;
    } else {
      LoadPriceHistory(SheetLab, Row, PriceType, PriceRegion, ConsumerKey, ConsumerSecret, TokenValue, TokenSecret);
    }
    SheetLab.getRange("J2").setValue(Row);
  }
  
  // UI
  var WorkedRows = Row - StartingRow;  
  var Ui = SpreadsheetApp.getUi()
  Ui.alert('Lab', 'Updated ' + WorkedRows + ' rows!', Ui.ButtonSet.OK);
}

// Function: Download Prices (Single)
function LoadPricesSingle(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var ConsumerKey = SheetSettings.getRange("B3").getValue();
  var ConsumerSecret = SheetSettings.getRange("B4").getValue();
  var TokenValue = SheetSettings.getRange("B5").getValue();
  var TokenSecret = SheetSettings.getRange("B6").getValue();
  var LabActive = SheetSettings.getRange("B8").getValue();
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  // Data
  var PriceType = SheetLab.getRange("M2").getValue();
  var PriceRegion = SheetLab.getRange("K2").getValue();
  if (PriceRegion=="Worldwide") PriceRegion="";
  var Selection = SheetLab.getDataRange();
  var Cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var StartingRow = Cell.getRow();
  
  if (StartingRow < LabMinRow) return;

  var CellCode = Selection.getCell(StartingRow,LabColumnItemNo).getValue();
    if (CellCode==""){
       return;
    } else {
      Cell.setBackground("Red");
      LoadPriceHistory(SheetLab, StartingRow, PriceType, PriceRegion, ConsumerKey, ConsumerSecret, TokenValue, TokenSecret)
    }
   Cell.setBackground("White");

}

// Function: Download Price (Core)
function LoadPriceHistory(SheetLab, Row, PriceType, PriceRegion, ConsumerKey, ConsumerSecret, TokenValue, TokenSecret){
  var Selection = SheetLab.getDataRange();
  var CellItemType = Selection.getCell(Row,LabColumnItemType).getValue();
  var CellCode = Selection.getCell(Row,LabColumnItemNo).getValue();
  var CellColorID = Selection.getCell(Row,LabColumnColorID).getValue();
  var CellCondition = Selection.getCell(Row,LabColumnCondition).getValue();

  if (CellItemType == "PART"){
    var ItemType = 'PART'
    var Url = 'https://api.bricklink.com/api/store/v1' + '/items/part/' + CellCode + '/price';
  } else if (CellItemType == "MINIFIG"){
    var ItemType = 'MINIFIG'
    var Url = 'https://api.bricklink.com/api/store/v1' + '/items/minifig/' + CellCode + '/price';
  } else if (CellItemType == "SET"){
    var ItemType = 'SET'
    var Url = 'https://api.bricklink.com/api/store/v1' + '/items/set/' + CellCode + '/price';
    CellColorID = "";
  }

  // API Request
  var Options = {method: 'GET',contentType: 'application/json'};
  var Params = {
    no: CellCode,
    color_id: CellColorID,
    type: ItemType,
    new_or_used: CellCondition,
    guide_type: PriceType,
    region: PriceRegion,
    currency_code: 'EUR',
    vat: 'Y'
  }; 
          
  urlFetch = OAuth1.withAccessToken(ConsumerKey, ConsumerSecret, TokenValue, TokenSecret);

  // Output  
  var PriceGuide = JSON.parse(urlFetch.fetch(Url, Params, Options));
  SheetLab.getRange(Row, LabColumnPriceMin).setValue(PriceGuide.data.min_price);
  SheetLab.getRange(Row, LabColumnPriceAvg).setValue(PriceGuide.data.avg_price);
  SheetLab.getRange(Row, LabColumnPriceAvgQty).setValue(PriceGuide.data.qty_avg_price);
  SheetLab.getRange(Row, LabColumnPriceMax).setValue(PriceGuide.data.max_price);
  SheetLab.getRange(Row, LabColumnPriceLot).setValue(PriceGuide.data.unit_quantity);
  SheetLab.getRange(Row, LabColumnPriceQty).setValue(PriceGuide.data.total_quantity);
}

// Function: Import Inventory in Lab
function ImportInventory() {
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var LabActive = SheetSettings.getRange("B8").getValue()
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  var SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

  // Data
  var ItemType =  SheetLab.getRange("B2").getValue();
  var CategoryId = SheetLab.getRange("D2").getValue();
  var ColorId = SheetLab.getRange("H2").getValue();
  var Mode = SheetLab.getRange("A2").getValue();

  if (Mode == "ADD"){
    var data = SheetLab.getRange(LabMinRow, 2, SheetInventory.getLastRow(), 1).getValues().join('@').split('@');
    LabMinRow = LabMinRow + data.filter(Boolean).length;
    } else if (Mode = "CLEAR"){
    ClearLab()
  }

  // Output, For Loop
  var Data = [];
  Data = SheetInventory.getRange(4, 1, SheetInventory.getLastRow(), 18).getValues();
  var j = 0;

  for (var i=0; i < Data.length; i++){
    if (Data[i][1] == ""){
      break;
    } else {      
      if (Data[i][1] == ItemType || ItemType == ""){
        if (Data[i][3] == CategoryId || CategoryId == "-1"){
          if (Data[i][5] == ColorId || ColorId == ""){
            SheetLab.getRange(j+LabMinRow,1).setValue(Data[i][1]);
            SheetLab.getRange(j+LabMinRow,2).setValue(Data[i][2]);
            SheetLab.getRange(j+LabMinRow,3).setValue(Data[i][6]);
            SheetLab.getRange(j+LabMinRow,5).setValue(Data[i][12]);
            SheetLab.getRange(j+LabMinRow,6).setValue(Data[i][13]);
            SheetLab.getRange(j+LabMinRow,7).setValue(Data[i][15]);
            j++
          }
        }
      }
    }
  }

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Import', 'Import completed.', Ui.ButtonSet.OK);
}

// Function: Import PartOut in Lab
function ImportPartOut() {
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var LabActive = SheetSettings.getRange("B8").getValue()
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  
  var SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PartOut");

  // Data
  var ItemType =  SheetLab.getRange("B2").getValue();
  var CategoryId = SheetLab.getRange("D2").getValue();
  var ColorName = SheetLab.getRange("E2").getValue();
  var Mode = SheetLab.getRange("A2").getValue();
  var Conditions = SheetPartOut.getRange("I2").getValue();

  if (Mode == "ADD"){
    var data = SheetLab.getRange(LabMinRow, 2, SheetInventory.getLastRow(), 1).getValues().join('@').split('@');
    LabMinRow = LabMinRow + data.filter(Boolean).length;
    } else if (Mode = "CLEAR"){
    ClearLab()
  }

  // Output, For Loop
  var Data = [];
  Data = SheetPartOut.getRange(4, 1, SheetPartOut.getLastRow(), 9).getValues();
  var j = 0;

  for (var i=0; i < Data.length; i++){  
    if (Data[i][1] == ItemType || ItemType == ""){
      if (Data[i][4] == CategoryId || CategoryId == "-1"){
        if (Data[i][5] == ColorName || ColorName == ""){
          SheetLab.getRange(j+LabMinRow,1).setValue(Data[i][1]);
          SheetLab.getRange(j+LabMinRow,2).setValue(Data[i][2]);
          SheetLab.getRange(j+LabMinRow,3).setValue(Data[i][8]);
          SheetLab.getRange(j+LabMinRow,4).setValue(Data[i][6])
          SheetLab.getRange(j+LabMinRow,5).setValue(Conditions);
          SheetLab.getRange(j+LabMinRow,6).setValue("X");
          SheetLab.getRange(j+LabMinRow,7).setValue("NO");
          j++
        }
      }
    }
  }

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Import', 'Import completed!', Ui.ButtonSet.OK);
}

// Function: Clear Lab
function ClearLab(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var LabActive = SheetSettings.getRange("B8").getValue()
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  var LabMinRow = 4;
  var LabMaxRow = SheetLab.getMaxRows();

  ClearLabPrices()
  SheetLab.getRange("J2").clear({contentsOnly: true});
  SheetLab.getRange(LabMinRow, 1, LabMaxRow, 7).clear({contentsOnly: true});
}

// Function: Clear Lab (Prices)
function ClearLabPrices(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var LabActive = SheetSettings.getRange("B8").getValue()
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  var LabMinRow = 4;
  var LabMaxRow = SheetLab.getMaxRows();

  SheetLab.getRange(LabMinRow, 15, LabMaxRow, 7).clear({contentsOnly: true});
}
