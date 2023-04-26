// Constants
const SheetColorSetting = "#000000";
const SheetColorDatabase = "#727272";
const SheetColorInventory = "#FFFF00";
const SheetColorPartOut = "#FF0000";
const SheetColorLab = "#0000FF";

const RowHeightMain = 75;
const RowHeightPlus = 45;
const RowHeightStandard = 21;

const ColoumColorPermanent = "#FFF2CC"
const ColoumColorInput = "#B6D7A8"
const CellColorPermanent = "#D9D9D9"
const CellColorInput = "#A4C2F4"

// Function: Regenerate Settings
function RegenerateSettings() {
  RegenerateSheet("Settings", SheetColorSetting, 14, 2)
  SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

  // Style
  SheetSettings.setColumnWidth(1, 200);
  SheetSettings.setColumnWidth(2, 350);
  SheetSettings.setRowHeights(1, SheetSettings.getMaxRows(), RowHeightStandard)
  SheetSettings.getRange(1,1, SheetSettings.getMaxRows(), SheetSettings.getMaxColumns()).setNumberFormat('@STRING@');
  
  SheetSettings.getRange("A1:B2").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetSettings.getRange("A1:B1").mergeAcross().setHorizontalAlignment("Center").setFontWeight("bold");
  SheetSettings.getRange("A2:B2").mergeAcross().setHorizontalAlignment("Center");

  SheetSettings.getRange("A7:B8").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetSettings.getRange("A7:B7").mergeAcross().setHorizontalAlignment("Center").setFontWeight("bold");
  SheetSettings.getRange("A8:B8").mergeAcross().setHorizontalAlignment("Center");
  
  SheetSettings.getRange("A11:B11").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetSettings.getRange("A11:B11").mergeAcross().setHorizontalAlignment("Center").setFontWeight("bold");

  SheetSettings.getRange("A1:B14").setBackground(CellColorPermanent);
  SheetSettings.getRange("B3:B6").setBackground(CellColorInput);
  SheetSettings.getRange("B9:B10").setBackground(CellColorInput);
  SheetSettings.getRange("B12:B14").setBackground(CellColorInput);
  SpreadsheetApp.flush();

  // Text  
  const ColumnA = [["Bricklink API Token"],["https://www.bricklink.com/v2/api/register_consumer.page"],["BL Consumer Key"], ["BL Consumer Secret"], ["BL Token Value"], ["BL Token Secret"], ["TurboBrickManager API Token"], ["https://ziotitanok.it/tbm"], ["TBM Token Value"], [""], ["Settings"], ["Database Auto Hide"], ["Lab Active"], ["Prices Row Max (Bulk/Batch)"]];
  SheetSettings.getRange("A1:A14").setValues(ColumnA);
  SheetSettings.getRange("B13").setValue("Lab");
  SheetSettings.getRange("B14").setValue("750");

  // Dropdowns
  const CheckBoxes = SpreadsheetApp.newDataValidation().requireValueInList(["YES", "NO"]).build();
  SheetSettings.getRange("B12").setDataValidation(CheckBoxes).setValue("YES");

  SpreadsheetApp.flush();
  
  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Settings', 'Settings is ready again!', Ui.ButtonSet.OK);
}

// Function: Regenerate DB-Colors
function RegenerateDBColors() {
  RegenerateSheet("DB-Colors", SheetColorDatabase, 250, 4)
  SheetDBColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors");
  const {DBAutoHide, TBMHeaders} = GetSettings();

  // Style
  SheetDBColors.setColumnWidth(1,175);
  SheetDBColors.setColumnWidths(2,3,75);
  SheetDBColors.setColumnWidth(4,100);
  SheetDBColors.setRowHeights(1, SheetDBColors.getMaxRows(), RowHeightStandard);
  SheetDBColors.setFrozenRows(1);
  SheetDBColors.getRange(1,1, SheetDBColors.getMaxRows(), SheetDBColors.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();

  // Text
  const TitlesA = ["Color Name", "Color ID", "RGB", "Type"];
  SheetDBColors.getRange("A1:D1").setBackground(CellColorPermanent).setFontWeight("bold").setValues([TitlesA]);

  // API Request & Output
  const Url = `${TBMBaseUrl}/BricklinkCatalogColors/`;
  const Response = UrlFetchApp.fetch(Url, {headers: TBMHeaders});
  const OutputColorGuide = JSON.parse(Response.getContentText());
  
  const ColorGuide = OutputColorGuide.map((Item) => {
    return [Item.colorname, Item.colorid, Item.rgb, Item.colortype];
  });

  SheetDBColors.getRange(3, 1, ColorGuide.length, 4).setValues(ColorGuide).sort([4,1]);
  SheetDBColors.deleteRows(4+ColorGuide.length, SheetDBColors.getMaxRows()-ColorGuide.length-4);
  SpreadsheetApp.flush();

  // AutoHide
  if (DBAutoHide == "YES") SheetDBColors.hideSheet();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Colors', 'DB-Colors is ready again!', Ui.ButtonSet.OK);
}

// Function: Regenerate DB-Category
function RegenerateDBCategories() {
  RegenerateSheet("DB-Categories", SheetColorDatabase, 1200, 2)
  SheetDBCategory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories");
  const {DBAutoHide, TBMHeaders} = GetSettings();

  // Style
  SheetDBCategory.setColumnWidth(1, 100);
  SheetDBCategory.setColumnWidth(2, 300);
  SheetDBCategory.setRowHeights(1, SheetDBCategory.getMaxRows(), RowHeightStandard);
  SheetDBCategory.setFrozenRows(1);
  SheetDBCategory.getRange(1,1, SheetDBCategory.getMaxRows(), SheetDBCategory.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  const TitlesA = ["Category ID", "Category Name"];
  SheetDBCategory.getRange("A1:B1").setBackground(CellColorPermanent).setFontWeight("bold").setValues([TitlesA]);

  // API Request & Output
  const Url = `${TBMBaseUrl}/BricklinkCatalogCategory/`;
  const Response = UrlFetchApp.fetch(Url, {headers: TBMHeaders});
  const OutputCategoryGuide = JSON.parse(Response.getContentText());
  
  const CategoryGuide = OutputCategoryGuide.map((Item) => {
    return [Item.categoryid, Item.categoryname];
  });

  SheetDBCategory.getRange(3, 1, CategoryGuide.length, 2).setValues(CategoryGuide).sort([2]);
  SheetDBCategory.deleteRows(3+CategoryGuide.length, SheetDBCategory.getMaxRows()-CategoryGuide.length-3);
  SpreadsheetApp.flush();

  // AutoHide
  if (DBAutoHide == "YES") SheetDBColors.hideSheet();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Category', 'DB-Category is ready again!', Ui.ButtonSet.OK);  
}

// Function: Regenerate DB-Part
function RegenerateDBPart() {
  RegenerateSheet("DB-Part", SheetColorDatabase, 100000, 4)
  SheetDBPart = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Part");
  const {DBAutoHide, TBMHeaders} = GetSettings();

  // Style
  SheetDBPart.setColumnWidth(1, 100);
  SheetDBPart.setColumnWidth(2, 200);
  SheetDBPart.setColumnWidth(3, 100);
  SheetDBPart.setColumnWidth(4, 1600);
  SheetDBPart.setRowHeights(1, SheetDBPart.getMaxRows(), RowHeightStandard);
  SheetDBPart.setFrozenRows(1);
  SheetDBPart.getRange(1,1, SheetDBPart.getMaxRows(), SheetDBPart.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  const TitlesA = ["Category ID", "Category Name", "Number", "Name"];
  SheetDBPart.getRange("A1:D1").setBackground(CellColorPermanent).setFontWeight("bold").setValues([TitlesA]);

  // API Request & Output
  const Url = `${TBMBaseUrl}/BricklinkCatalogPart/`;
  const Response = UrlFetchApp.fetch(Url, {headers: TBMHeaders});
  const OutputPartGuide = JSON.parse(Response.getContentText());

  const PartGuide = OutputPartGuide.map((Item) => {
    return [Item.categoryid, Item.categoryname, Item.partcode, Item.partname];
  });

  SheetDBPart.getRange(2, 1, PartGuide.length, 4).setValues(PartGuide);
  SheetDBPart.deleteRows(3+PartGuide.length, SheetDBPart.getMaxRows()-PartGuide.length-3);
  SpreadsheetApp.flush();

  // AutoHide
  if (DBAutoHide == "YES") SheetDBColors.hideSheet();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Part', 'DB-Part is ready again!', Ui.ButtonSet.OK); 
}

// Function: DB-Minifigure
function RegenerateDBMinifigure() {
  RegenerateSheet("DB-Minifigure", SheetColorDatabase, 20000, 4)
  SheetDBMinifigure = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Minifigure");
  const {DBAutoHide, TBMHeaders} = GetSettings();

  // Style
  SheetDBMinifigure.setColumnWidth(1, 100);
  SheetDBMinifigure.setColumnWidth(2, 200);
  SheetDBMinifigure.setColumnWidth(3, 100);
  SheetDBMinifigure.setColumnWidth(4, 1000);
  SheetDBMinifigure.setRowHeights(1, SheetDBMinifigure.getMaxRows(), RowHeightStandard);
  SheetDBMinifigure.setFrozenRows(1);
  SheetDBMinifigure.getRange(1,1, SheetDBMinifigure.getMaxRows(), SheetDBMinifigure.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  const TitlesA = ["Category ID", "Category Name", "Number", "Name"];
  SheetDBMinifigure.getRange("A1:D1").setBackground(CellColorPermanent).setFontWeight("bold").setValues([TitlesA]);

  // API Request
  const Url = `${TBMBaseUrl}/BricklinkCatalogMinifigure/`;
  const Response = UrlFetchApp.fetch(Url, {headers: TBMHeaders});
  const OutputMinifigureGuide = JSON.parse(Response.getContentText());
  
  const MinifigureGuide = OutputMinifigureGuide.map((Item) => {
    return [Item.categoryid, Item.categoryname, Item.minifigcode, Item.minifigname];
  });

  SheetDBMinifigure.getRange(2, 1, MinifigureGuide.length, 4).setValues(MinifigureGuide);
  SheetDBMinifigure.deleteRows(3+MinifigureGuide.length, SheetDBMinifigure.getMaxRows()-MinifigureGuide.length-3);
  SpreadsheetApp.flush();

  // AutoHide
  if (DBAutoHide == "YES") SheetDBColors.hideSheet();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Minifigure', 'DB-Minifigure is ready again!', Ui.ButtonSet.OK); 
}

// Function: DB-Set
function RegenerateDBSet() {
  RegenerateSheet("DB-Set", SheetColorDatabase, 20000, 4)
  SheetDBSet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Set");
  const {DBAutoHide, TBMHeaders} = GetSettings();
  
  // Style
  SheetDBSet.setColumnWidth(1, 100);
  SheetDBSet.setColumnWidth(2, 500);
  SheetDBSet.setColumnWidth(3, 100);
  SheetDBSet.setColumnWidth(4, 7500);
  SheetDBSet.setRowHeights(1, SheetDBSet.getMaxRows(), RowHeightStandard);
  SheetDBSet.setFrozenRows(1);
  SheetDBSet.getRange(1,1, SheetDBSet.getMaxRows(), SheetDBSet.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  const TitlesA = ["Category ID", "Category Name", "Number", "Name"];
  SheetDBSet.getRange("A1:D1").setBackground(CellColorPermanent).setFontWeight("bold").setValues([TitlesA]);

  // API Request
  const Url = `${TBMBaseUrl}/BricklinkCatalogSet/`;
  const Response = UrlFetchApp.fetch(Url, {headers: TBMHeaders});
  const OutputSetGuide = JSON.parse(Response.getContentText());
  
  const SetGuide = OutputSetGuide.map((Item) => {
    return [Item.categoryid, Item.categoryname, Item.setcode, Item.setname];
  });

  SheetDBSet.getRange(2, 1, SetGuide.length, 4).setValues(SetGuide);
  SheetDBSet.deleteRows(3+SetGuide.length, SheetDBSet.getMaxRows()-SetGuide.length-3);
  SpreadsheetApp.flush();

  // AutoHide
  if (DBAutoHide == "YES") SheetDBColors.hideSheet();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Set', 'DB-Set is ready again!', Ui.ButtonSet.OK); 
}

// Function: DB-Codes
function RegenerateDBCodes() {
  RegenerateSheet("DB-Codes", SheetColorDatabase, 100000, 3)
  SheetDBCodes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Codes");
  const {DBAutoHide, TBMHeaders} = GetSettings();

  // Style
  SheetDBCodes.setColumnWidth(1, 100);
  SheetDBCodes.setColumnWidth(2, 125);
  SheetDBCodes.setColumnWidth(3, 200);
  SheetDBCodes.setRowHeights(1, SheetDBCodes.getMaxRows(), RowHeightStandard);
  SheetDBCodes.setFrozenRows(1);
  SheetDBCodes.getRange(1,1, SheetDBCodes.getMaxRows(), SheetDBCodes.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  const TitlesA = ["Code", "Item No", "Color"];
  SheetDBCodes.getRange("A1:C1").setBackground(CellColorPermanent).setFontWeight("bold").setValues([TitlesA]);

  // API Request & Output
  const Url = `${TBMBaseUrl}/BricklinkCatalogCodes/`;
  const Response = UrlFetchApp.fetch(Url, {headers: TBMHeaders});
  const OutputCodesGuide = JSON.parse(Response.getContentText());
  
  const CodesGuide = OutputCodesGuide.map((Item) => {
    return [Item.legoid, Item.itemid, Item.colorname];
  });

  SheetDBCodes.getRange(2, 1, CodesGuide.length, 3).setValues(CodesGuide);
  SheetDBCodes.deleteRows(3+CodesGuide.length, SheetDBCodes.getMaxRows()-CodesGuide.length-3);
  SpreadsheetApp.flush();

  // AutoHide
  if (DBAutoHide == "YES") SheetDBColors.hideSheet();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Codes', 'DB-Codes is ready again!', Ui.ButtonSet.OK); 
}

// Function: Regenerate Inventory
function RegenerateInventory() {
  RegenerateSheet("Inventory", SheetColorInventory, 10000, 18)
  const SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');

  // Style
  SheetInventory.setColumnWidths(1, 4, 100);
  SheetInventory.setColumnWidth(5, 300);
  SheetInventory.setColumnWidth(6, 50);
  SheetInventory.setColumnWidths(7, 2, 250);
  SheetInventory.setColumnWidth(9, 50);
  SheetInventory.setColumnWidth(10, 100);
  SheetInventory.setColumnWidths(11, 2, 150);
  SheetInventory.setColumnWidths(13, 4, 50);
  SheetInventory.setColumnWidth(17, 100);
  SheetInventory.setColumnWidth(18, 200);
  SheetInventory.setRowHeights(1, SheetInventory.getMaxRows(), RowHeightStandard)
  SheetInventory.setRowHeight(3, RowHeightMain);
  SheetInventory.setFrozenRows(3);

  SheetInventory.getRange("A1:H2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetInventory.getRange("R1:R2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetInventory.getRange("A3:R3").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetInventory.getRange("A1:R3").setFontWeight("bold").setHorizontalAlignment("Center").setVerticalAlignment("Middle");

  SheetInventory.getRange(1,1, SheetInventory.getMaxRows(), SheetInventory.getMaxColumns()).setNumberFormat('@STRING@');
  SheetInventory.getRange(4, 10, SheetInventory.getMaxRows()-4, 1).setNumberFormat("##0.00[$€]");

  SheetInventory.getRange("A1:R2").setBackground(CellColorPermanent);
  SheetInventory.getRange("A2:E2").setBackground(CellColorInput);
  SheetInventory.getRange("G2:H2").setBackground(CellColorInput);
  SheetInventory.getRange("A3:R3").setBackground(ColoumColorPermanent).setTextRotation(45).setWrap(true);
  SpreadsheetApp.flush();

  // Text
  const TitlesA = ["Part", "Minifig", "Set", "Others", "Category Name","", "Color Name", "StockRoom", "", "", "", "", "", "", "", "", "","Last Download"];
  SheetInventory.getRange("A1:R1").setValues([TitlesA]);
  
  const TitlesC = ["i", "Item Type", "Item Code", "Category ID", "Item Name", "Color ID", "Color Name", "Index", "Qty", "Price", "Description", "Remarks",	"Condition", "Completeness", "Is Stock?", "Stock ID", "Inventory ID", "Date Created"];
  SheetInventory.getRange("A3:R3").setValues([TitlesC]);

  // Dropdowns
  SheetDBCategory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories");
  const CategoryRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBCategory.getRange("B2:B")).build();
  SheetInventory.getRange("E2").setDataValidation(CategoryRule);
  SheetInventory.getRange("F1").setValue("=XLOOKUP(E2,'DB-Categories'!B:B,'DB-Categories'!A:A,\"\")")

  SheetDBColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors");
  const ColorsRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBColors.getRange("A2:A")).build();
  SheetInventory.getRange("G2").setDataValidation(ColorsRule);
  SheetInventory.getRange("F2").setValue("=xLOOKUP(G2,'DB-Colors'!A:A,'DB-Colors'!B:B,\"\")")

  const StockRoomRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "A", "B", "C"]).build();
  SheetInventory.getRange("H2").setDataValidation(StockRoomRule);

  const CheckBoxes = SpreadsheetApp.newDataValidation().requireValueInList(["YES", "NO"]).build();
  SheetInventory.getRange("A2:D2").setDataValidation(CheckBoxes).setValue("YES");
  SpreadsheetApp.flush();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Inventory', 'Inventory is ready again!', Ui.ButtonSet.OK); 
}

// Function: Regenerate PartOut
function RegeneratePartOut() {
  RegenerateSheet("PartOut", SheetColorPartOut, 5000, 9)

  // Style
  SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PartOut");
  SheetPartOut.setColumnWidth(1, 75);
  SheetPartOut.setColumnWidths(2, 2, 100);
  SheetPartOut.setColumnWidth(4, 300);
  SheetPartOut.setColumnWidths(5, 4, 75);
  SheetPartOut.setColumnWidth(9, 250);
  SheetPartOut.setRowHeights(1, SheetPartOut.getMaxRows(), RowHeightStandard)
  SheetPartOut.setRowHeight(3, RowHeightMain);
  SheetPartOut.setFrozenRows(3);

  SheetPartOut.getRange("A1:D2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetPartOut.getRange("A3:I3").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetPartOut.getRange("A1:I3").setFontWeight("bold").setHorizontalAlignment("Center").setVerticalAlignment("Middle");

  SheetPartOut.getRange(1,1, SheetPartOut.getMaxRows(), SheetPartOut.getMaxColumns()).setNumberFormat('@STRING@');

  SheetPartOut.getRange("A1:I2").setBackground(CellColorPermanent);
  SheetPartOut.getRange("A2:D2").setBackground(CellColorInput);
  SheetPartOut.getRange("A3:I3").setBackground(ColoumColorPermanent).setTextRotation(45).setWrap(true);
  SpreadsheetApp.flush();

  // Text
  const TitlesA = ["Set", "Variant", "Break Minifig", "Set Name"];
  SheetPartOut.getRange("A1:D1").setValues([TitlesA]);

  const TitlesC = ["i", "Item Type", "Item Code", "Item Name", "Category ID", "Color ID", "Qty", "Match No", "Color Name"];
  SheetPartOut.getRange("A3:I3").setValues([TitlesC]);

  // Dropdowns
  const VariantRule = SpreadsheetApp.newDataValidation().requireValueInList(["-1", "-2", "-3", "-4", "-5", "-6", "-7", "-8", "-9", "-10", "-11", "-12", "-13", "-14", "-15", "-16"]).build();
  SheetPartOut.getRange("B2").setDataValidation(VariantRule);

  const BreakMinifigRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "YES"]).build();
  SheetPartOut.getRange("C2").setDataValidation(BreakMinifigRule);

  // Formula
  SheetPartOut.getRange("D2").setValue("=XLOOKUP(CONCATENATE(A2 & B2),'DB-Set'!C:C,'DB-Set'!D:D,\"\")");
  SheetPartOut.getRange("I4").setValue("=XLOOKUP(F4,'DB-Colors'!$B:B,'DB-Colors'!$A:A,\"\")");
  SheetPartOut.getRange("I4").copyTo(SheetPartOut.getRange("I5:I"));
  SpreadsheetApp.flush();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('PartOut', 'PartOut is ready again!', Ui.ButtonSet.OK); 
}

// Function: Regenerate Lab
function RegenerateLab() {
  RegenerateSheet("Lab", '#0000FF', 10000, 31)

  // Style
  SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lab");
  SheetLab.setColumnWidth(1, 85);
  SheetLab.setColumnWidths(2, 2, 130);
  SheetLab.setColumnWidths(4, 4, 50);
  SheetLab.setColumnWidth(8, 75);
  SheetLab.setColumnWidth(9, 150);
  SheetLab.setColumnWidths(10, 4, 70);
  SheetLab.setColumnWidth(14, 30);
  SheetLab.setColumnWidths(15, 7, 70);
  SheetLab.setColumnWidth(22, 100);
  SheetLab.setColumnWidth(23, 150);
  SheetLab.setColumnWidth(24, 100);
  SheetLab.setColumnWidths(25, 4, 150);
  SheetLab.setColumnWidths(29, 1, 200);
  SheetLab.setColumnWidth(30, 50);
  SheetLab.setColumnWidths(31, 1, 200);
  SheetLab.setRowHeights(1, 2, RowHeightStandard)
  SheetLab.setRowHeight(3, RowHeightMain);
  SheetLab.setRowHeights(4, SheetLab.getMaxRows()-3, RowHeightPlus);
  SheetLab.setFrozenRows(3);

  SheetLab.getRange("E1:G1").mergeAcross()
  SheetLab.getRange("E2:G2").mergeAcross()
  SheetLab.getRange("J1:K1").mergeAcross()
  SheetLab.getRange("J2:K2").mergeAcross()
  SheetLab.getRange("L1:M1").mergeAcross()
  SheetLab.getRange("L2:M2").mergeAcross()
  SheetLab.getRange("N1:P1").mergeAcross()
  SheetLab.getRange("N2:P2").mergeAcross()

  SheetLab.getRange("A1:S2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetLab.getRange("A3:AE3").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);

  SheetLab.getRange("A1:AE2").setBackground(CellColorPermanent);
  SheetLab.getRange("A2:C2").setBackground(CellColorInput);
  SheetLab.getRange("E2:M2").setBackground(CellColorInput);
  SheetLab.getRange("Q2:S2").setBackground(CellColorInput);
  SheetLab.getRange("A3:AE3").setBackground(ColoumColorPermanent).setTextRotation(45).setWrap(true);
  SheetLab.getRange("A3:G3").setBackground(ColoumColorInput);
  SheetLab.getRange("N3:O3").setBackground(ColoumColorInput);
  SheetLab.getRange("AA3:AB3").setBackground(ColoumColorInput);

  SheetLab.getRange(4, 1, SheetLab.getMaxRows()-3, 31).getBandings().forEach(function (banding) {banding.remove()});
  SheetLab.getRange(4, 1, SheetLab.getMaxRows()-3, 31).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  SheetLab.getRange("A1:AE3").setFontWeight("bold").setHorizontalAlignment("Center").setVerticalAlignment("Middle");
  SheetLab.getRange("A4:AE").setHorizontalAlignment("Center").setVerticalAlignment("Middle");
  SheetLab.getRange(1,1, SheetLab.getMaxRows(), SheetLab.getMaxColumns()).setNumberFormat('@STRING@');
  SheetLab.getRange(LabRowMin, 4, SheetLab.getMaxRows()-LabRowMin, 1).setNumberFormat("##0");
  SheetLab.getRange(LabRowMin, 9, SheetLab.getMaxRows()-LabRowMin, 1).setWrap(true);
  SheetLab.getRange(LabRowMin, 11, SheetLab.getMaxRows()-LabRowMin, 1).setNumberFormat("##0.00[$€]");
  SheetLab.getRange(LabRowMin, 12, SheetLab.getMaxRows()-LabRowMin, 2).setNumberFormat("##0.##%"); 
  SheetLab.getRange(LabRowMin, 14, SheetLab.getMaxRows()-LabRowMin, 1).insertCheckboxes().setNumberFormat("General");
  SheetLab.getRange(LabRowMin, 15, SheetLab.getMaxRows()-LabRowMin, 5).setNumberFormat("##0.00[$€]");  
  SpreadsheetApp.flush();

  const PercentageConditional = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue('#FFAFAF', SpreadsheetApp.InterpolationType.NUMBER, "100%")
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMinpointWithValue('#AFAFFF', SpreadsheetApp.InterpolationType.NUMBER, "-100%")
    .setRanges([SheetLab.getRange(LabRowMin, 12, SheetLab.getMaxRows(), 2)])
    .build();
  const PercentagesConditional = SheetLab.getConditionalFormatRules();
  PercentagesConditional.push(PercentageConditional);
  SheetLab.setConditionalFormatRules(PercentagesConditional);

  // Text
  const TitlesA = ["Mode", "Item Type", "Category", "", "Color", "", "", "N / U", "Stockroom", "Zone", "","Price Guide", "", "Last Worked Row", "", "","Tolerance", "VarPrice", "Angle"];
  SheetLab.getRange("A1:S1").setValues([TitlesA]);

  const TitlesC = ["Item Type", "Code", "Color Name", "Qty", "N / U", "Complete?", "Stock?", "Immagine", "Item Name", "On BL", "Price Inv.", "%Avg", "% Avg/Qty","Check", "Prezzo (O)", "Min", "Avg", "Avg/Qty", "Max", "Lotti", "Item Avaiable", "Link: LotID", "Link: Catalogo", "Link: Inventario", "Descrizione", "Remarks", "Descrizione (O)", "Remark (O)", "Date Created", "IDCol", "Index"];
  SheetLab.getRange("A3:AE3").setValues([TitlesC]);

  // Dropdowns
  const ModeRule = SpreadsheetApp.newDataValidation().requireValueInList(["ADD", "CLEAR"]).build();
  SheetLab.getRange("A2").setDataValidation(ModeRule).setValue("ADD");

  const ItemTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(["PART", "MINIFIG", "SET", "GEAR", "BOOK"]).build();
  SheetLab.getRange("B2").setDataValidation(ItemTypeRule);
  SheetLab.getRange("A4:A").setDataValidation(ItemTypeRule);

  SheetDBCategory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories");
  const CategoryRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBCategory.getRange("B2:B")).build();
  SheetLab.getRange("C2").setDataValidation(CategoryRule);
  SheetLab.getRange("D1").setValue("=XLOOKUP(C2,'DB-Categories'!B:B,'DB-Categories'!A:A,\"\")");

  SheetDBColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors");
  const ColorsRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBColors.getRange("A2:A")).build();
  SheetLab.getRange("E2").setDataValidation(ColorsRule);
  SheetLab.getRange("C4:C").setDataValidation(ColorsRule);
  SheetLab.getRange("D2").setValue("=XLOOKUP(E2,'DB-Colors'!A:A, 'DB-Colors'!B:B,\"\")");
  
  const StockRoomRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "A", "B", "C"]).build();
  SheetLab.getRange("I2").setDataValidation(StockRoomRule).setValue("NO");

  const ZoneRule = SpreadsheetApp.newDataValidation().requireValueInList(["Europe", "Worldwide"]).build();
  SheetLab.getRange("J2").setDataValidation(ZoneRule).setValue("Europe");

  const PriceRule = SpreadsheetApp.newDataValidation().requireValueInList(["Stock", "Sold"]).build();
  SheetLab.getRange("L2").setDataValidation(PriceRule).setValue("Stock");

  const ConditionRule = SpreadsheetApp.newDataValidation().requireValueInList(["N", "U"]).build();
  SheetLab.getRange("H2").setDataValidation(ConditionRule).setValue("N");
  SheetLab.getRange("E4:E").setDataValidation(ConditionRule);

  const CompletenessRule = SpreadsheetApp.newDataValidation().requireValueInList(["S", "C", "B", "X"]).build();
  SheetLab.getRange("F4:F").setDataValidation(CompletenessRule);
  
  const StockRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "A", "B", "C"]).build();
  SheetLab.getRange("G4:G").setDataValidation(StockRule);

  const ImageRule = SpreadsheetApp.newDataValidation().requireValueInList(["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17"]).build();
  SheetLab.getRange("S2").setDataValidation(ImageRule).setValue("01");

  // Parameters
  var ParametersB = ["0", "0"];
  SheetLab.getRange("Q2:R2").setValues([ParametersB]).setNumberFormat("##0.##%"); 

  SpreadsheetApp.flush()

  // Formulas
  SheetLab.getRange("H4").setValue("=(image(IFERROR(IFS(and(A4 = \"PART\", len(C4) > 0),CONCATENATE(\"http://img.bricklink.com/ItemImage/PN/\" & AD4 & \"/\" & B4 & \".png\"), and(A4 = \"PART\", len(C4) = 0), CONCATENATE(\"https://www.bricklink.com/3D/images/\" & B4 & \"/640/\" & $S$2 & \".png\"), A4 = \"MINIFIG\", CONCATENATE(\"http://img.bricklink.com/ItemImage/MN/0/\" & B4 & \".png\"), A4 = \"SET\", CONCATENATE(\"https://img.bricklink.com/ItemImage/SN/0/\" & B4 & \".png\"), len(A4) = 0,\"\"),\"\")))")
  SheetLab.getRange("H4").copyTo(SheetLab.getRange("H5:H"));
  SpreadsheetApp.flush();

  SheetLab.getRange("I4").setValue("=IFERROR(IFS(A4 = \"PART\",XLOOKUP(B4,'DB-Part'!$C:$C, 'DB-Part'!$D:$D,\"\",1), A4 = \"MINIFIG\", XLOOKUP(B4,'DB-Minifigure'!$C:$C, 'DB-Minifigure'!$D:$D, \"\", 1), A4 = \"SET\", XLOOKUP(B4,'DB-Set'!$C:$C,'DB-Set'!$D:$D, \"\",1)), \"\")");
  SheetLab.getRange("I4").copyTo(SheetLab.getRange("I5:I"));
  SpreadsheetApp.flush();

  SheetLab.getRange("J4").setValue("=XLOOKUP(AE4, Inventory!$H:$H, Inventory!$I:$I,\"\")");
  SheetLab.getRange("J4").copyTo(SheetLab.getRange("J5:J"));
  SpreadsheetApp.flush();

  SheetLab.getRange("K4").setValue("=XLOOKUP(AE4, Inventory!$H:$H,Inventory!$J:$J,\"\")");
  SheetLab.getRange("K4").copyTo(SheetLab.getRange("K5:K"));
  SpreadsheetApp.flush();

  SheetLab.getRange("L4").setValue("=IF(K4*Q4>0,(K4-Q4)/Q4,\"\")");
  SheetLab.getRange("L4").copyTo(SheetLab.getRange("L5:L"));
  SpreadsheetApp.flush();

  SheetLab.getRange("M4").setValue("=IF(K4*R4>0,(K4-R4)/R4,\"\")");
  SheetLab.getRange("M4").copyTo(SheetLab.getRange("M5:M"));
  SpreadsheetApp.flush();

  SheetLab.getRange("O4").setValue("=IFERROR(IFS(((R4+Q4)/2)/K4>(1+$Q$2), ((R4+Q4)/2)*(1+$R$2), K4/((R4+Q4)/2)>(1+$Q$2), ((R4+Q4)/2)*(1+$R$2)), \"\")");
  SheetLab.getRange("O4").copyTo(SheetLab.getRange("O5:O"));
  SpreadsheetApp.flush();

  SheetLab.getRange("V4").setValue("=IF(LEN(AE4)>0, HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?invID=\" & XLOOKUP(AE4, Inventory!$H:$H, Inventory!$Q:$Q,\"\")),(XLOOKUP(AE4,Inventory!$H:$H, Inventory!$Q:$Q,\"\"))), \"\")");
  SheetLab.getRange("V4").copyTo(SheetLab.getRange("V5:V"));
  SpreadsheetApp.flush();

  SheetLab.getRange("W4").setValue("=IFS(ISBLANK(B4),\"\", A4 = \"PART\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/v2/catalog/catalogitem.page?P=\" & B4 & \"#T=P&C=\" & AD4), CONCATENATE(B4 & \" - \" & C4)), A4 = \"MINIFIG\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/v2/catalog/catalogitem.page?M=\" & B4),B4), A4 = \"SET\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/v2/catalog/catalogitem.page?S=\" & B4),B4))");
  SheetLab.getRange("W4").copyTo(SheetLab.getRange("W5:W"));
  SpreadsheetApp.flush();

  SheetLab.getRange("X4").setValue("=IFS(ISBLANK(B4),\"\", A4=\"PART\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?q=\" & B4), B4), A4=\"MINIFIG\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?q=\" & B4), B4), A4=\"SET\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?q=\" & LEFT(B4, LEN(B4)-2)), CONCATENATE(B4)))");
  SheetLab.getRange("X4").copyTo(SheetLab.getRange("X5:X"));
  SpreadsheetApp.flush();

  SheetLab.getRange("Y4").setValue("=XLOOKUP(AE4, Inventory!$H:$H,Inventory!$K:$K,\"\")");
  SheetLab.getRange("Y4").copyTo(SheetLab.getRange("Y5:Y"));

  SheetLab.getRange("Z4").setValue("=XLOOKUP(AE4, Inventory!$H:$H,Inventory!$L:$L, \"\")");
  SheetLab.getRange("Z4").copyTo(SheetLab.getRange("Z5:Z"));

  SheetLab.getRange("AC4").setValue("=XLOOKUP(AE4, Inventory!$H:$H,Inventory!$R:$R,\"\")");
  SheetLab.getRange("AC4").copyTo(SheetLab.getRange("AC5:AC"));
  SpreadsheetApp.flush();

  SheetLab.getRange("AD4").setValue("=XLOOKUP(C4, 'DB-Colors'!A:A,'DB-Colors'!B:B ,\"\",1)");
  SheetLab.getRange("AD4").copyTo(SheetLab.getRange("AD5:AD"));

  SheetLab.getRange("AE4").setValue("=IFERROR(IFS(A4 = \"PART\", A4 & \"_\" & B4 & \"_\" & AD4 & \"_\" & E4 & \"_\" & G4, A4 = \"MINIFIG\", A4 & \"_\" & B4 & \"_\" & E4 & \"_\" & G4, A4 = \"SET\", A4 & \"_\" & B4 & \"_\" & E4 & \"_\" & F4 & \"_\" & G4),\"\")");
  SheetLab.getRange("AE4").copyTo(SheetLab.getRange("AE5:AE"));
  SpreadsheetApp.flush();

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Lab', 'Lab is ready again!', Ui.ButtonSet.OK); 
}

// Function: Regenerate XML
function RegenerateXML() {
  RegenerateSheet("XML", '#FF00FF', 5000, 2)

  // Style
  const SheetXML = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("XML")
  SheetXML.setColumnWidths(1, 2, 850);
  SheetXML.setRowHeights(1, SheetXML.getMaxRows(), RowHeightStandard)
  SheetXML.getRange("A1:B1").setBackground(CellColorPermanent).setFontWeight("bold");
  SheetXML.setFrozenRows(1);

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('XML', 'XML is ready again!', Ui.ButtonSet.OK); 
}

// Funzione: Generic Sheet
function RegenerateSheet(SheetName, SheetColor, Rows, Columns) {
/**
 * RegenerateSheet(SheetName, SheetIndex, Rows, Columns)
 * @param {String}  SheetName     Name of the new sheet
 * @param {String}  SheetColor    Color of the new sheet
 * @param {Number}  Rows          Vertical dimension of new sheet (default 0 means "system default", 1000)
 * @param {Number}  Columns       Horizontal dimension of new sheet (default 0 means "system default", 26)
 * @returns {Sheet}               Sheet object for chaining.
 */

  // Check
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName)){
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(SheetName)};
  const Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
  
  // Color
  Sheet.setTabColor(SheetColor);
  
  // Rows
  if (Rows !== 0) {
    const SheetRows = Sheet.getMaxRows();    
    if (Rows < SheetRows){
      Sheet.deleteRows(Rows+1, SheetRows-Rows);
    } else if (Rows > SheetRows) {
       Sheet.insertRowsAfter(SheetRows, Rows-SheetRows);
    }
  }
  
  // Columns
  if (Columns !== 0) {
    const SheetColumns = Sheet.getMaxColumns();
    if (Columns < SheetColumns) {
      Sheet.deleteColumns(Columns+1, SheetColumns-Columns);
    } else if (Columns > SheetColumns) {
      Sheet.insertColumnsAfter(SheetColumns,Columns-SheetColumns);
    }
  }

  // Return new Sheet object
  return Sheet;
}