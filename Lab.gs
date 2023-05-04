// Constants: Lab
const LabRowMin = 4;
const LabColumnItemType = 1;
const LabColumnItemNo = 2;
const LabColumnColorName = 3;
const LabColumnQty = 4;
const LabColumnCondition = 5;
const LabColumnCompleteness = 6;
const LabColumnStock = 7;
const LabColumQtyInventory = 10;
const LabColumnPrice = 15;
const LabColumnPriceMin = 16;
const LabColumnPriceAvg = 17;
const LabColumnPriceAvgQty = 18;
const LabColumnPriceMax = 19;
const LabColumnLotID = 22;
const LabColumnDescription = 27;
const LabColumnRemarks = 28;
const LabColumnColorID = 30;

// Function: Download Prices (Bulk)
function LoadPricesBulk(){
  const {BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret, LabActive, LabRowMaxPrice} = GetSettings();
  const {PriceType, PriceRegion} = GetPriceTypeAndRegion();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  // Data
  const Cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  const StartingRow = Cell.getRow();
  if (StartingRow < LabRowMin) return;
  
  var LabUsedRows = SheetLab.getRange(StartingRow, 2, SheetLab.getLastRow(), 1).getValues().join('@').split('@');
  if (LabUsedRows.filter(Boolean).length <= LabRowMaxPrice) {
    Input = SheetLab.getRange(StartingRow, 1, LabUsedRows.filter(Boolean).length, 30).getValues();
  } else {
    Input = SheetLab.getRange(StartingRow, 1, LabRowMaxPrice, 30).getValues();
  }
  
  // For Loop & API Request & Output
  var Output = [];
  for (var i in Input){
    var Url = `${BrickLinkBaseUrl}/items/${Input[i][LabColumnItemType-1]}/${Input[i][LabColumnItemNo-1]}/price`;
    console.log(Url)
    urlFetch = OAuth1.withAccessToken(BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret);
    var Params = {
      no: Input[i][LabColumnItemNo-1],
      color_id: Input[i][LabColumnColorID-1],
      type: Input[i][LabColumnItemType-1],
      new_or_used: Input[i][LabColumnCondition-1],
      guide_type: PriceType,
      region: PriceRegion,
      currency_code: 'EUR',
      vat: 'Y'
    }; 

    var PriceGuide = JSON.parse(urlFetch.fetch(Url, Params, BrickLinkOptions));
    Output[i] = [PriceGuide.data.min_price, 
                PriceGuide.data.avg_price, 
                PriceGuide.data.qty_avg_price, 
                PriceGuide.data.max_price,
                PriceGuide.data.unit_quantity, 
                PriceGuide.data.total_quantity];
  }

  SheetLab.getRange(StartingRow, LabColumnPriceMin, Output.length, 6).setValues(Output);  
  
  // UI
  const Ui = SpreadsheetApp.getUi();
  const WorkedRows = Output.length;
  SheetLab.getRange("N2").setValue(StartingRow+Output.length-1);
  Ui.alert('Lab', 'Updated ' + WorkedRows + ' rows!', Ui.ButtonSet.OK);
}

// Function: Download Prices (Rows)
function LoadPricesRows(){
  const {BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret, LabActive, LabRowMaxPrice} = GetSettings();
  const {PriceType, PriceRegion} = GetPriceTypeAndRegion();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  // Data
  const Selection = SheetLab.getDataRange();
  const Cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  const StartingRow = Cell.getRow();
  if (StartingRow < LabRowMin) return;
  
  // For Loop
  for (var Row = StartingRow; Row <= StartingRow+LabRowMaxPrice; Row++){
    var CellCode = Selection.getCell(Row,LabColumnItemNo).getValue();
    if (CellCode == ""){
      break;
    } else {
      LoadPriceHistory(SheetLab, Row, PriceType, PriceRegion, BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret);
    }
    SheetLab.getRange("N2").setValue(Row);
  }
  
  // UI
  const Ui = SpreadsheetApp.getUi();
  const WorkedRows = Row - StartingRow;  
  Ui.alert('Lab', 'Updated ' + WorkedRows + ' rows!', Ui.ButtonSet.OK);
}

// Function: Download Prices (Row)
function LoadPricesRow(){
  const {BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret, LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const {PriceType, PriceRegion} = GetPriceTypeAndRegion();

  // Data
  var Selection = SheetLab.getDataRange();
  var Cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  var StartingRow = Cell.getRow();
  if (StartingRow < LabRowMin) return;

  const CellCode = Selection.getCell(StartingRow,LabColumnItemNo).getValue();
    if (CellCode==""){
       return;
    } else {
      var Row = StartingRow;
      LoadPriceHistory(SheetLab, Row, PriceType, PriceRegion, BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret)
    }
}

// Function: Download Price (Core)
function LoadPriceHistory(SheetLab, Row, PriceType, PriceRegion, BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret){
  var Selection = SheetLab.getDataRange();
  var CellItemType = Selection.getCell(Row,LabColumnItemType).getValue();
  var CellCode = Selection.getCell(Row,LabColumnItemNo).getValue();
  var CellColorID = Selection.getCell(Row,LabColumnColorID).getValue();
  var CellCondition = Selection.getCell(Row,LabColumnCondition).getValue();

  // API Request & Output
  var Url = `${BrickLinkBaseUrl}/items/${CellItemType}/${CellCode}/price`;
  console.log(Url)
  urlFetch = OAuth1.withAccessToken(BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret);
  var Params = {
    no: CellCode,
    color_id: CellColorID,
    type: CellItemType,
    new_or_used: CellCondition,
    guide_type: PriceType,
    region: PriceRegion,
    currency_code: 'EUR',
    vat: 'Y'
  }; 

  var PriceGuide = JSON.parse(urlFetch.fetch(Url, Params, BrickLinkOptions));
  var Output = [PriceGuide.data.min_price, 
                PriceGuide.data.avg_price, 
                PriceGuide.data.qty_avg_price, 
                PriceGuide.data.max_price, 
                PriceGuide.data.unit_quantity, 
                PriceGuide.data.total_quantity];
  SheetLab.getRange(Row, LabColumnPriceMin, 1, 6).setValues([Output]);
}

// Function: Hint Prices
function HintPrices() {
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  SheetLab.getRange("O4").setValue("=IFERROR(IFS(((R4+Q4)/2)/K4>(1+$Q$2), ((R4+Q4)/2)*(1+$R$2), K4/((R4+Q4)/2)>(1+$Q$2), ((R4+Q4)/2)*(1+$R$2)), \"\")");;
  SheetLab.getRange("O4").copyTo(SheetLab.getRange("O5:O"));
  SpreadsheetApp.flush();
}

// Function: Import Inventory in Lab
function ImportInventory() {
  const SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);

  // Data
  const Mode = SheetLab.getRange("A2").getValue();
  const ItemType =  SheetLab.getRange("B2").getValue();
  const CategoryId = SheetLab.getRange("D1").getValue();
  const ColorId = SheetLab.getRange("D2").getValue();
  const Conditions = SheetLab.getRange("H2").getValue();
  const StockRoom = SheetLab.getRange("I2").getValue()

  if (Mode == "ADD"){
    const LabUsedRows = SheetLab.getRange(LabRowMin, 2, SheetLab.getLastRow(), 1).getValues().join('@').split('@');
    const RealLabRowMin = LabRowMin + LabUsedRows.filter(Boolean).length;
  } else if (Mode == "CLEAR"){ 
    ClearLab();
    var RealLabRowMin = LabRowMin
  }

  // Output, For Loop
  const InventoryUsedRows = SheetInventory.getRange(4, 1, SheetInventory.getLastRow(), 1).getValues().join('@').split('@');
  const Data = SheetInventory.getRange(4, 1, InventoryUsedRows.filter(Boolean).length, 18).getValues();

  var Output = [];
  var j = 0;
  for (var i in Data){
    if (Data[i][1] == ItemType || ItemType == ""){
      if (Data[i][3] == CategoryId || CategoryId == ""){
        if (Data[i][5] == ColorId || ColorId == ""){
          if (Data[i][12] == Conditions || Conditions == ""){
            if (Data[i][15] == StockRoom || StockRoom == ""){
              Output[j] = [Data[i][1], Data[i][2], Data[i][6], "", Data[i][12], Data[i][13], Data[i][15]];
              j++
            }
          }
        }
      }
    }
  }

  SheetLab.getRange(RealLabRowMin, 1, Output.length,7).setValues(Output);

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Import', 'Import completed.', Ui.ButtonSet.OK);
}

// Function: Import PartOut in Lab
function ImportPartOut() {
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PartOut");

  // Data
  const Mode = SheetLab.getRange("A2").getValue();
  const ItemType =  SheetLab.getRange("B2").getValue();
  const CategoryId = SheetLab.getRange("D1").getValue();
  const ColorId = SheetLab.getRange("D2").getValue();
  const Conditions = SheetLab.getRange("H2").getValue();
  const StockRoom = SheetLab.getRange("I2").getValue()

  if (Mode == "ADD"){
    const LabUsedRows = SheetLab.getRange(LabRowMin, 2, SheetLab.getLastRow(), 1).getValues().join('@').split('@');
    var RealLabRowMin = LabRowMin + LabUsedRows.filter(Boolean).length;
  } else if (Mode = "CLEAR"){ 
    ClearLab();
    var RealLabRowMin = LabRowMin
  }

  // Output
  const PartOutUsedRows = SheetPartOut.getRange(4, 1, SheetPartOut.getLastRow(), 1).getValues().join('@').split('@');
  const Data = SheetPartOut.getRange(4, 1, PartOutUsedRows.filter(Boolean).length, 9).getValues();

  var Output = [];
  var j = 0;
  for (var i in Data){
    if (Data[i][1] == ItemType || ItemType == ""){
      if (Data[i][4] == CategoryId || CategoryId == ""){
        if (Data[i][5] == ColorId || ColorId == ""){
          Output[j] = [Data[i][1], Data[i][2], Data[i][8], Data[i][6], Conditions, "", StockRoom];
          j++
        }
      }
    }
  }

  SheetLab.getRange(RealLabRowMin, 1, Output.length,7).setValues(Output);

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Import', 'Import completed!', Ui.ButtonSet.OK);
}

// Function: Clear Lab
function ClearLab(){
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const LabMaxRow = SheetLab.getMaxRows();

  ClearLabPrices()
  SheetLab.getRange(LabRowMin, 1, LabMaxRow, 7).clearContent();
  SheetLab.getRange(LabRowMin, 14, LabMaxRow, 1).clearContent().insertCheckboxes().setNumberFormat("General");
  SheetLab.getRange(LabRowMin, 27, LabMaxRow, 2).clearContent();
}

// Function: Clear Lab (Prices)
function ClearLabPrices(){
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const LabMaxRow = SheetLab.getMaxRows();

  SheetLab.getRange(LabRowMin, 15, LabMaxRow, 7).clearContent();
  HintPrices();
}

// Function: Get Price Type And Region
function GetPriceTypeAndRegion() {
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const PriceType = SheetLab.getRange("L2").getValue();
  const PriceRegion = SheetLab.getRange("J2").getValue();
  if (PriceRegion == "Worldwide") PriceRegion = "";
  
  return {PriceType, PriceRegion};
}