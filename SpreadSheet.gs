// Function: Regenerate Settings
function RegenerateSettings() {
  RegenerateSheet("Settings", '#000000', 13, 2)
  SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");

  // Style
  SheetSettings.setColumnWidth(1, 200);
  SheetSettings.setColumnWidth(2, 350);
  SheetSettings.setRowHeights(1, SheetSettings.getMaxRows(), 21)
  SheetSettings.getRange(1,1, SheetSettings.getMaxRows(), SheetSettings.getMaxColumns()).setNumberFormat('@STRING@');
  
  SheetSettings.getRange("A1:B2").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetSettings.getRange("A1:B1").mergeAcross().setHorizontalAlignment("Center").setFontWeight("bold");
  SheetSettings.getRange("A2:B2").mergeAcross().setHorizontalAlignment("Center");

  SheetSettings.getRange("A7:B8").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetSettings.getRange("A7:B7").mergeAcross().setHorizontalAlignment("Center").setFontWeight("bold");
  SheetSettings.getRange("A8:B8").mergeAcross().setHorizontalAlignment("Center");
  
  SheetSettings.getRange("A11:B11").setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetSettings.getRange("A11:B11").mergeAcross().setHorizontalAlignment("Center").setFontWeight("bold");

  SheetSettings.getRange("A1:B13").setBackground('#D9D9D9');
  SheetSettings.getRange("B3:B6").setBackground('#A4C2F4');
  SheetSettings.getRange("B9:B10").setBackground('#A4C2F4');
  SheetSettings.getRange("B12:B13").setBackground('#A4C2F4');
  SpreadsheetApp.flush();

  // Text  
  var ColumnA = [["Bricklink API Token"],["https://www.bricklink.com/v2/api/register_consumer.page"],["Consumer Key"], ["Consumer Secret"], ["Token Value"], ["Token Secret"], ["TurboBrickManager API Token"], ["https://ziotitanok.it/tbm"], ["Token Value"], [""], ["Lab"], ["Lab Active"], ["Prices Row Max (Bulk/Batch)"]];
  SheetSettings.getRange("A1:A13").setValues(ColumnA);
  SheetSettings.getRange("B13").setValue("750");

  // Dropdowns
  var LabActiveRule = SpreadsheetApp.newDataValidation().requireValueInList(["Lab"]).build();
  SheetSettings.getRange("B12").setDataValidation(LabActiveRule).setValue("Lab");
  SpreadsheetApp.flush();
  
  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Settings', 'Settings is ready again!', Ui.ButtonSet.OK);
}

// Function: Regenerate DB-Colors
function RegenerateDBColors() {
  RegenerateSheet("DB-Colors", '#727272', 250, 4)
  SheetDBColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors");

  // Style
  SheetDBColors.setColumnWidth(1,175);
  SheetDBColors.setColumnWidths(2,3,75);
  SheetDBColors.setColumnWidth(4,100);
  SheetDBColors.setRowHeights(1, SheetDBColors.getMaxRows(), 21);
  SheetDBColors.setFrozenRows(1);
  SheetDBColors.getRange(1,1, SheetDBColors.getMaxRows(), SheetDBColors.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();

  // Text
  var TitlesA = ["Color Name", "Color ID", "RGB", "Type"];
  SheetDBColors.getRange("A1:D1").setBackground('#D9D9D9').setFontWeight("bold").setValues([TitlesA]);

  // Data
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var TBMToken = SheetSettings.getRange("B9").getValue();

  // API Request
  var url = 'https://django.ziotitanok.it/api/' + 'BricklinkCatalogColors/';
  var headers = {'Authorization': 'Token ' + TBMToken};
   
  var response = UrlFetchApp.fetch(url, {headers: headers});
  var OutputColorGuide = JSON.parse(response.getContentText());
  
  // Output, For Loop
  var i = 0;
  ColorGuide = [];

  for (i in OutputColorGuide){
    ColorGuide[i] = [OutputColorGuide[i].colorname,
                    OutputColorGuide[i].colorid,
                    OutputColorGuide[i].rgb,
                    OutputColorGuide[i].colortype
                    ]
  }

  SheetDBColors.getRange(3, 1, ColorGuide.length, 4).setValues(ColorGuide);
  SheetDBColors.deleteRows(4+ColorGuide.length, SheetDBColors.getMaxRows()-ColorGuide.length-4);
  SpreadsheetApp.flush();

  SheetDBColors.hideSheet();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Colors', 'DB-Colors is ready again!', Ui.ButtonSet.OK);
}

// Function: Regenerate DB-Category
function RegenerateDBCategories() {
  RegenerateSheet("DB-Categories", '#727272', 1200, 2)
  SheetDBCategory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories");

  // Style
  SheetDBCategory.setColumnWidth(1, 100);
  SheetDBCategory.setColumnWidth(2, 300);
  SheetDBCategory.setRowHeights(1, SheetDBCategory.getMaxRows(), 21);
  SheetDBCategory.setFrozenRows(1);
  SheetDBCategory.getRange(1,1, SheetDBCategory.getMaxRows(), SheetDBCategory.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  var TitlesA = ["Category ID", "Category Name"];
  SheetDBCategory.getRange("A1:B1").setBackground('#D9D9D9').setFontWeight("bold").setValues([TitlesA]);

  // Data
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var TBMToken = SheetSettings.getRange("B9").getValue();

  // API Request
  var url = 'https://django.ziotitanok.it/api/' + 'BricklinkCatalogCategory/';
  var headers = {'Authorization': 'Token ' + TBMToken};
   
  var response = UrlFetchApp.fetch(url, {headers: headers});
  var OutputCategoryGuide = JSON.parse(response.getContentText());
  
  // Output, For Loop
  var i = 0;
  CategoryGuide = [];

  for (i in OutputCategoryGuide){
    CategoryGuide[i] = [OutputCategoryGuide[i].categoryid,
                        OutputCategoryGuide[i].categoryname
                    ]
  }

  SheetDBCategory.getRange(3, 1, CategoryGuide.length, 2).setValues(CategoryGuide);
  SheetDBCategory.deleteRows(3+CategoryGuide.length, SheetDBCategory.getMaxRows()-CategoryGuide.length-3);
  SpreadsheetApp.flush();

  SheetDBCategory.hideSheet();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Category', 'DB-Category is ready again!', Ui.ButtonSet.OK);  
}

// Function: Regenerate DB-Part
function RegenerateDBPart() {
  RegenerateSheet("DB-Part", '#727272', 100000, 4)
  SheetDBPart = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Part");

  // Style
  SheetDBPart.setColumnWidth(1, 100);
  SheetDBPart.setColumnWidth(2, 200);
  SheetDBPart.setColumnWidth(3, 100);
  SheetDBPart.setColumnWidth(4, 1600);
  SheetDBPart.setRowHeights(1, SheetDBPart.getMaxRows(), 21);
  SheetDBPart.setFrozenRows(1);
  SheetDBPart.getRange(1,1, SheetDBPart.getMaxRows(), SheetDBPart.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  var TitlesA = ["Category ID", "Category Name", "Number", "Name"];
  SheetDBPart.getRange("A1:D1").setBackground('#D9D9D9').setFontWeight("bold").setValues([TitlesA]);

  // Data
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var TBMToken = SheetSettings.getRange("B9").getValue();

  // API Request
  var url = 'https://django.ziotitanok.it/api/' + 'BricklinkCatalogPart/';
  var headers = {'Authorization': 'Token ' + TBMToken};
  
  var response = UrlFetchApp.fetch(url, {headers: headers});
  var OutputPartGuide = JSON.parse(response.getContentText());
  
  // Output, For Loop
  var i = 0;
  PartGuide = [];

  for (i in OutputPartGuide){
    PartGuide[i] = [OutputPartGuide[i].categoryid,
                    OutputPartGuide[i].categoryname,
                    OutputPartGuide[i].partcode,
                    OutputPartGuide[i].partname
                    ]
  }

  SheetDBPart.getRange(2, 1, PartGuide.length, 4).setValues(PartGuide);
  SheetDBPart.deleteRows(3+PartGuide.length, SheetDBPart.getMaxRows()-PartGuide.length-3);
  SpreadsheetApp.flush();

  SheetDBPart.hideSheet();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Part', 'DB-Part is ready again!', Ui.ButtonSet.OK); 
}

// Function: DB-Minifigure
function RegenerateDBMinifigure() {
  RegenerateSheet("DB-Minifigure", '#727272', 20000, 4)
  SheetDBMinifigure = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Minifigure");

  // Style
  SheetDBMinifigure.setColumnWidth(1, 100);
  SheetDBMinifigure.setColumnWidth(2, 200);
  SheetDBMinifigure.setColumnWidth(3, 100);
  SheetDBMinifigure.setColumnWidth(4, 1000);
  SheetDBMinifigure.setRowHeights(1, SheetDBMinifigure.getMaxRows(), 21);
  SheetDBMinifigure.setFrozenRows(1);
  SheetDBMinifigure.getRange(1,1, SheetDBMinifigure.getMaxRows(), SheetDBMinifigure.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  var TitlesA = ["Category ID", "Category Name", "Number", "Name"];
  SheetDBMinifigure.getRange("A1:D1").setBackground('#D9D9D9').setFontWeight("bold").setValues([TitlesA]);

  // Data
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var TBMToken = SheetSettings.getRange("B9").getValue();

  // API Request
  var url = 'https://django.ziotitanok.it/api/' + 'BricklinkCatalogMinifigure/';
  var headers = {'Authorization': 'Token ' + TBMToken};
  
  var response = UrlFetchApp.fetch(url, {headers: headers});
  var OutputMinifigureGuide = JSON.parse(response.getContentText());
  
  // Output, For Loop
  var i = 0;
  MinifigureGuide = [];

  for (i in OutputMinifigureGuide){
    MinifigureGuide[i] = [OutputMinifigureGuide[i].categoryid,
                    OutputMinifigureGuide[i].categoryname,
                    OutputMinifigureGuide[i].minifigcode,
                    OutputMinifigureGuide[i].minifigname
                    ]
  }

  SheetDBMinifigure.getRange(2, 1, MinifigureGuide.length, 4).setValues(MinifigureGuide);
  SheetDBMinifigure.deleteRows(3+MinifigureGuide.length, SheetDBMinifigure.getMaxRows()-MinifigureGuide.length-3);
  SpreadsheetApp.flush();

  SheetDBMinifigure.hideSheet();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Minifigure', 'DB-Minifigure is ready again!', Ui.ButtonSet.OK); 
}

// Function: DB-Set
function RegenerateDBSet() {
  RegenerateSheet("DB-Set", '#727272', 20000, 4)
  SheetDBSet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Set");

  // Style
  SheetDBSet.setColumnWidth(1, 100);
  SheetDBSet.setColumnWidth(2, 500);
  SheetDBSet.setColumnWidth(3, 100);
  SheetDBSet.setColumnWidth(4, 7500);
  SheetDBSet.setRowHeights(1, SheetDBSet.getMaxRows(), 21);
  SheetDBSet.setFrozenRows(1);
  SheetDBSet.getRange(1,1, SheetDBSet.getMaxRows(), SheetDBSet.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  var TitlesA = ["Category ID", "Category Name", "Number", "Name"];
  SheetDBSet.getRange("A1:D1").setBackground('#D9D9D9').setFontWeight("bold").setValues([TitlesA]);

  // Data
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var TBMToken = SheetSettings.getRange("B9").getValue();

  // API Request
  var url = 'https://django.ziotitanok.it/api/' + 'BricklinkCatalogSet/';
  var headers = {'Authorization': 'Token ' + TBMToken};
  
  var response = UrlFetchApp.fetch(url, {headers: headers});
  var OutputSetGuide = JSON.parse(response.getContentText());
  
  // Output, For Loop
  var i = 0;
  SetGuide = [];

  for (i in OutputSetGuide){
    SetGuide[i] = [OutputSetGuide[i].categoryid,
                    OutputSetGuide[i].categoryname,
                    OutputSetGuide[i].setcode,
                    OutputSetGuide[i].setname
                    ]
  }

  SheetDBSet.getRange(2, 1, SetGuide.length, 4).setValues(SetGuide);
  SheetDBSet.deleteRows(3+SetGuide.length, SheetDBSet.getMaxRows()-SetGuide.length-3);
  SpreadsheetApp.flush();

  SheetDBSet.hideSheet();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Set', 'DB-Set is ready again!', Ui.ButtonSet.OK); 

  SheetDBSet.hideSheet();
}

// Function: DB-Codes
function RegenerateDBCodes() {
  RegenerateSheet("DB-Codes", '#727272', 100000, 3)
  SheetDBCodes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Codes");

  // Style
  SheetDBCodes.setColumnWidth(1, 100);
  SheetDBCodes.setColumnWidth(2, 125);
  SheetDBCodes.setColumnWidth(3, 200);
  SheetDBCodes.setRowHeights(1, SheetDBCodes.getMaxRows(), 21);
  SheetDBCodes.setFrozenRows(1);
  SheetDBCodes.getRange(1,1, SheetDBCodes.getMaxRows(), SheetDBCodes.getMaxColumns()).setNumberFormat('@STRING@');
  SpreadsheetApp.flush();
  
  // Text
  var TitlesA = ["Code", "Item No", "Color"];
  SheetDBCodes.getRange("A1:C1").setBackground('#D9D9D9').setFontWeight("bold").setValues([TitlesA]);

  // Data
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var TBMToken = SheetSettings.getRange("B9").getValue();

  // API Request
  var url = 'https://django.ziotitanok.it/api/' + 'BricklinkCatalogCodes/';
  var headers = {'Authorization': 'Token ' + TBMToken};
  
  var response = UrlFetchApp.fetch(url, {headers: headers});
  var OutputCodesGuide = JSON.parse(response.getContentText());
  
  // Output, For Loop
  var i = 0;
  CodesGuide = [];

  for (i in OutputCodesGuide){
    CodesGuide[i] = [OutputCodesGuide[i].legoid,
                    OutputCodesGuide[i].itemid,
                    OutputCodesGuide[i].colorname
                    ]
  }

  SheetDBCodes.getRange(2, 1, CodesGuide.length, 3).setValues(CodesGuide);
  SheetDBCodes.deleteRows(3+CodesGuide.length, SheetDBCodes.getMaxRows()-CodesGuide.length-3);
  SpreadsheetApp.flush();

  SheetDBCodes.hideSheet();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('DB-Codes', 'DB-Codes is ready again!', Ui.ButtonSet.OK); 

  SheetDBCodes.hideSheet();
}

// Function: Regenerate Inventory
function RegenerateInventory() {
  RegenerateSheet("Inventory", '#FFFF00', 10000, 18)

  // Style
  SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory");
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
  SheetInventory.setRowHeights(1, SheetInventory.getMaxRows(), 21)
  SheetInventory.setRowHeight(3, 75);
  SheetInventory.setFrozenRows(3);

  SheetInventory.getRange("A1:H2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetInventory.getRange("R1:R2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetInventory.getRange("A3:R3").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetInventory.getRange("A1:R3").setFontWeight("bold").setHorizontalAlignment("Center").setVerticalAlignment("Middle");

  SheetInventory.getRange(1,1, SheetInventory.getMaxRows(), SheetInventory.getMaxColumns()).setNumberFormat('@STRING@');
  SheetInventory.getRange(4, 10, SheetInventory.getMaxRows()-4, 1).setNumberFormat("##0.00[$€]");

  SheetInventory.getRange("A1:R2").setBackground('#D9D9D9');
  SheetInventory.getRange("A2:E2").setBackground('#A4C2F4');
  SheetInventory.getRange("G2:H2").setBackground('#A4C2F4');
  SheetInventory.getRange("A3:R3").setBackground('#FFF2CC').setTextRotation(45).setWrap(true);
  SpreadsheetApp.flush();

  // Text
  var TitlesA = ["Part", "Minifig", "Set", "All", "Category Name","", "Color Name", "StockRoom", "", "", "", "", "", "", "", "", "","Last Download"];
  SheetInventory.getRange("A1:R1").setValues([TitlesA]);
  
  var TitlesC = ["i", "Item Type", "Item Code", "Category ID", "Item Name", "Color ID", "Color Name", "Index", "Qty", "Price", "Description", "Remarks",	"Condition", "Completeness", "Is Stock?", "Stock ID", "Inventory ID", "Date Created"];
  SheetInventory.getRange("A3:R3").setValues([TitlesC]);

  // Dropdowns
  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories")) GenerateDBCategory();
  SheetDBCategory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories");
  var CategoryRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBCategory.getRange("B2:B")).build();
  SheetInventory.getRange("E2").setDataValidation(CategoryRule);
  SheetInventory.getRange("F1").setValue("=XLOOKUP(E2, 'DB-Categories'!B:B,'DB-Categories'!A:A,\"\",1)")

  if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors")) GenerateDBColors();
  SheetDBColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors");
  var ColorsRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBColors.getRange("A2:A")).build();
  SheetInventory.getRange("G2").setDataValidation(ColorsRule);
  SheetInventory.getRange("F2").setValue("=IFERROR(VLOOKUP(G2, 'DB-Colors'!A:B,2,FALSE),\"\")")

  var StockRoomRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "A", "B", "C"]).build();
  SheetInventory.getRange("H2").setDataValidation(StockRoomRule);

  var CheckBoxes = SpreadsheetApp.newDataValidation().requireValueInList(["YES", "NO"]).build();
  SheetInventory.getRange("A2:D2").setDataValidation(CheckBoxes);
  SpreadsheetApp.flush();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Inventory', 'Inventory is ready again!', Ui.ButtonSet.OK); 
}

function RegeneratePartOut() {
  RegenerateSheet("PartOut", '#FF0000', 5000, 9)

  // Style
  SheetPartOut = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PartOut");
  SheetPartOut.setColumnWidth(1, 75);
  SheetPartOut.setColumnWidths(2, 2, 100);
  SheetPartOut.setColumnWidth(4, 300);
  SheetPartOut.setColumnWidths(5, 4, 75);
  SheetPartOut.setColumnWidth(9, 250);
  SheetPartOut.setRowHeights(1, SheetPartOut.getMaxRows(), 21)
  SheetPartOut.setRowHeight(3, 75);
  SheetPartOut.setFrozenRows(3);

  SheetPartOut.getRange("A1:F2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetPartOut.getRange("A3:I3").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetPartOut.getRange("A1:I3").setFontWeight("bold").setHorizontalAlignment("Center").setVerticalAlignment("Middle");

   SheetPartOut.getRange(1,1, SheetPartOut.getMaxRows(), SheetPartOut.getMaxColumns()).setNumberFormat('@STRING@');

  SheetPartOut.getRange("A1:I2").setBackground('#D9D9D9');
  SheetPartOut.getRange("A2:F2").setBackground('#A4C2F4');
  SheetPartOut.getRange("A3:I3").setBackground('#FFF2CC').setTextRotation(45).setWrap(true);
  SpreadsheetApp.flush();

  // Text
  var TitlesA = ["Set", "Variant", "Break Minifig", "Set Name", "Condition", "StockRoom", "", "",""];
  SheetPartOut.getRange("A1:I1").setValues([TitlesA]);

  var TitlesC = ["i", "Item Type", "Item Code", "Item Name", "Category ID", "Color ID", "Qty", "Match No", "Color Name"];
  SheetPartOut.getRange("A3:I3").setValues([TitlesC]);

  // Dropdowns
  var VariantRule = SpreadsheetApp.newDataValidation().requireValueInList(["-1", "-2", "-3", "-4", "-5", "-6", "-7", "-8", "-9", "-10", "-11", "-12", "-13", "-14", "-15", "-16"]).build();
  SheetPartOut.getRange("B2").setDataValidation(VariantRule);

  var BreakMinifigRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "YES"]).build();
  SheetPartOut.getRange("C2").setDataValidation(BreakMinifigRule);

  var ConditionRule = SpreadsheetApp.newDataValidation().requireValueInList(["N", "U"]).build();
  SheetPartOut.getRange("E2").setDataValidation(ConditionRule);

  var StockRoomRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "A", "B", "C"]).build();
  SheetPartOut.getRange("F2").setDataValidation(StockRoomRule);
  SpreadsheetApp.flush();

  // Formula
  SheetPartOut.getRange("D2").setValue("=XLOOKUP(CONCATENATE(A2 & B2), 'DB-Set'!C:C, 'DB-Set'!D:D, \"\", 1)");

  SheetPartOut.getRange("I4").setValue("XLOOKUP(F4, 'DB-Colors'!$B:B, 'DB-Colors'!$A:A, \"\",1)");
  SheetPartOut.getRange("I4").copyTo(SheetPartOut.getRange("I5:I"));

  // UI
  var Ui = SpreadsheetApp.getUi();
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
  SheetLab.setRowHeights(1, 2, 21)
  SheetLab.setRowHeight(3, 75);
  SheetLab.setRowHeights(4, SheetLab.getMaxRows()-3, 45);
  SheetLab.setFrozenRows(3);

  SheetLab.getRange("E1:G1").mergeAcross()
  SheetLab.getRange("E2:G2").mergeAcross()
  SheetLab.getRange("J1:K1").mergeAcross()
  SheetLab.getRange("J2:K2").mergeAcross()
  SheetLab.getRange("L1:M1").mergeAcross()
  SheetLab.getRange("L2:M2").mergeAcross()

  SheetLab.getRange("A1:P2").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);
  SheetLab.getRange("A3:AE3").setBorder(true, true, true, true, true, null, "black", SpreadsheetApp.BorderStyle.SOLID);

  SheetLab.getRange("A1:AE2").setBackground('#D9D9D9');
  SheetLab.getRange("A2:C2").setBackground('#A4C2F4');
  SheetLab.getRange("E2").setBackground('#A4C2F4');
  SheetLab.getRange("H2").setBackground('#A4C2F4');
  SheetLab.getRange("J2").setBackground('#A4C2F4');
  SheetLab.getRange("L2").setBackground('#A4C2F4');
  SheetLab.getRange("O2").setBackground('#A4C2F4');
  SheetLab.getRange("P2").setBackground('#A4C2F4');
  SheetLab.getRange("A3:AE3").setBackground('#FFF2CC').setTextRotation(45).setWrap(true);
  SheetLab.getRange("A3:G3").setBackground('#B6D7A8');
  SheetLab.getRange("N3:O3").setBackground('#B6D7A8');
  SheetLab.getRange("AA3:AB3").setBackground('#B6D7A8');

  SheetLab.getRange(4, 1, SheetLab.getMaxRows()-3, 31).getBandings().forEach(function (banding) {banding.remove()});
  SheetLab.getRange(4, 1, SheetLab.getMaxRows()-3, 31).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  SheetLab.getRange("A1:AE3").setFontWeight("bold").setHorizontalAlignment("Center").setVerticalAlignment("Middle");
  SheetLab.getRange("A4:AE").setHorizontalAlignment("Center").setVerticalAlignment("Middle");
  SheetLab.getRange(1,1, SheetLab.getMaxRows(), SheetLab.getMaxColumns()).setNumberFormat('@STRING@');
  SheetLab.getRange(LabMinRow, 4, SheetLab.getMaxRows()-LabMinRow, 1).setNumberFormat("##0");
  SheetLab.getRange(LabMinRow, 9, SheetLab.getMaxRows()-LabMinRow, 1).setWrap(true);
  SheetLab.getRange(LabMinRow, 11, SheetLab.getMaxRows()-LabMinRow, 1).setNumberFormat("##0.00[$€]");
  SheetLab.getRange(LabMinRow, 12, SheetLab.getMaxRows()-LabMinRow, 2).setNumberFormat("##0.##%"); 
  SheetLab.getRange(LabMinRow, 14, SheetLab.getMaxRows()-LabMinRow, 1).insertCheckboxes().setNumberFormat("General");
  SheetLab.getRange(LabMinRow, 15, SheetLab.getMaxRows()-LabMinRow, 5).setNumberFormat("##0.00[$€]");  
  SpreadsheetApp.flush();

  var PercentageConditional = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMaxpointWithValue('#FFAFAF', SpreadsheetApp.InterpolationType.NUMBER, "100%")
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, "0")
    .setGradientMinpointWithValue('#AFAFFF', SpreadsheetApp.InterpolationType.NUMBER, "-100%")
    .setRanges([SheetLab.getRange(LabMinRow, 12, SheetLab.getMaxRows(), 2)])
    .build();
  var PercentagesConditional = SheetLab.getConditionalFormatRules();
  PercentagesConditional.push(PercentageConditional);
  SheetLab.setConditionalFormatRules(PercentagesConditional);

  // Text
  var TitlesA = ["Mode", "Item Type", "Category", "", "Color", "", "", "Angle", "Last Worked Row", "Zone", "","Price Guide", "", "", "Tolerance", "VarPrice"];
  SheetLab.getRange("A1:P1").setValues([TitlesA]);

  var TitlesC = ["Item Type", "Code", "Color Name", "Qty", "N / U", "Complete?", "Stock?", "Immagine", "Item Name", "On BL", "Price Inv.", "%Avg", "% Avg/Qty","Check", "Prezzo (O)", "Min", "Avg", "Avg/Qty", "Max", "Lotti", "Item Avaiable", "Link: LotID", "Link: Catalogo", "Link: Inventario", "Descrizione", "Remarks", "Descrizione (O)", "Remark (O)", "Date Created", "IDCol", "Index"];
  SheetLab.getRange("A3:AE3").setValues([TitlesC]);

  // Dropdowns
  var ModeRule = SpreadsheetApp.newDataValidation().requireValueInList(["ADD", "CLEAR"]).build();
  SheetLab.getRange("A2").setDataValidation(ModeRule).setValue("ADD");

  var ItemTypeRule = SpreadsheetApp.newDataValidation().requireValueInList(["PART", "MINIFIG", "SET", "GEAR", "BOOK"]).build();
  SheetLab.getRange("B2").setDataValidation(ItemTypeRule);
  SheetLab.getRange("A4:A").setDataValidation(ItemTypeRule);

  /* if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories")) GenerateDBCategory();
   * Should be done before to avoid time-out, assumed already ready
  */ 

  SheetDBCategory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Categories");
  var CategoryRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBCategory.getRange("B2:B")).build();
  SheetLab.getRange("C2").setDataValidation(CategoryRule);
  SheetLab.getRange("D1").setValue("=XLOOKUP(C2, 'DB-Categories'!B:B,'DB-Categories'!A:A,\"\",1)");

  /* if (!SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors")) GenerateDBColors();
  * Should be done before to avoid time-out, assumed already ready
  */ 

  SheetDBColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DB-Colors");
  var ColorsRule = SpreadsheetApp.newDataValidation().requireValueInRange(SheetDBColors.getRange("A2:A")).build();
  SheetLab.getRange("E2").setDataValidation(ColorsRule);
  SheetLab.getRange("C4:C").setDataValidation(ColorsRule);
  SheetLab.getRange("D2").setValue("=IFERROR(VLOOKUP(E2, 'DB-Colors'!A:B,2,FALSE),\"\")");

  SpreadsheetApp.flush();

  var ImageRule = SpreadsheetApp.newDataValidation().requireValueInList(["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17"]).build();
  SheetLab.getRange("H2").setDataValidation(ImageRule).setValue("01");

  var ZoneRule = SpreadsheetApp.newDataValidation().requireValueInList(["Europe", "Worldwide"]).build();
  SheetLab.getRange("J2").setDataValidation(ZoneRule).setValue("Europe");

  var PriceRule = SpreadsheetApp.newDataValidation().requireValueInList(["Stock", "Sold"]).build();
  SheetLab.getRange("L2").setDataValidation(PriceRule).setValue("Stock");

  var ConditionRule = SpreadsheetApp.newDataValidation().requireValueInList(["N", "U"]).build();
  SheetLab.getRange("E4:E").setDataValidation(ConditionRule);

  var CompletenessRule = SpreadsheetApp.newDataValidation().requireValueInList(["S", "C", "B", "X"]).build();
  SheetLab.getRange("F4:F").setDataValidation(CompletenessRule);
  
  var StockRule = SpreadsheetApp.newDataValidation().requireValueInList(["NO", "A", "B", "C"]).build();
  SheetLab.getRange("G4:G").setDataValidation(StockRule);

  // Parameters

  var ParametersB = ["0", "0"];
  SheetLab.getRange("O2:P2").setValues([ParametersB]).setNumberFormat("##0.##%"); 

  SpreadsheetApp.flush()

  // Formulas
  SheetLab.getRange("H4").setValue("=(image(IFERROR(IFS(and(A4 = \"PART\", len(C4) > 0),CONCATENATE(\"http://img.bricklink.com/ItemImage/PN/\" & AD4 & \"/\" & B4 & \".png\"), and(A4 = \"PART\", len(C4) = 0), CONCATENATE(\"https://www.bricklink.com/3D/images/\" & B4 & \"/640/\" & $H$2 & \".png\"), A4 = \"MINIFIG\", CONCATENATE(\"http://img.bricklink.com/ItemImage/MN/0/\" & B4 & \".png\"), A4 = \"SET\", CONCATENATE(\"https://img.bricklink.com/ItemImage/SN/0/\" & B4 & \".png\"), len(A4) = 0,\"\"),\"\")))")
  SheetLab.getRange("H4").copyTo(SheetLab.getRange("H5:H"));
  SpreadsheetApp.flush();

  SheetLab.getRange("I4").setValue("=IFERROR(IFS(A4 = \"PART\", VLOOKUP(B4,'DB-Part'!$C:$D,2,FALSE), A4 = \"MINIFIG\", VLOOKUP(B4,'DB-Minifigure'!$C:$D,2,FALSE), A4 = \"SET\", VLOOKUP(B4,'DB-Set'!$C:$D,2,FALSE)),\"\")");
  SheetLab.getRange("I4").copyTo(SheetLab.getRange("I5:I"));
  SpreadsheetApp.flush();

  SheetLab.getRange("J4").setValue("=IFERROR(VLOOKUP(AE4, Inventory!$H:$I,2,FALSE),\"\")");
  SheetLab.getRange("J4").copyTo(SheetLab.getRange("J5:J"));
  SpreadsheetApp.flush();

  SheetLab.getRange("K4").setValue("=IFERROR(VLOOKUP(AE4, Inventory!$H:$J,3,FALSE),\"\")");
  SheetLab.getRange("K4").copyTo(SheetLab.getRange("K5:K"));
  SpreadsheetApp.flush();

  SheetLab.getRange("L4").setValue("=IF(K4*Q4>0,(K4-Q4)/Q4,\"\")");
  SheetLab.getRange("L4").copyTo(SheetLab.getRange("L5:L"));
  SpreadsheetApp.flush();

  SheetLab.getRange("M4").setValue("=IF(K4*R4>0,(K4-R4)/R4,\"\")");
  SheetLab.getRange("M4").copyTo(SheetLab.getRange("M5:M"));
  SpreadsheetApp.flush();

  SheetLab.getRange("O4").setValue("=IFERROR(IFS(((R4+Q4)/2)/K4>(1+$O$2), ((R4+Q4)/2)*(1+$P$2), K4/((R4+Q4)/2)>(1+$O$2), ((R4+Q4)/2)*(1+$P$2)), \"\")");
  SheetLab.getRange("O4").copyTo(SheetLab.getRange("O5:O"));
  SpreadsheetApp.flush();

  SheetLab.getRange("V4").setValue("=IFERROR(HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?invID=\" & VLOOKUP(AE4, Inventory!$H:$Q,10,FALSE)),(VLOOKUP(AE4,Inventory!$H:$Q,10,FALSE))))");
  SheetLab.getRange("V4").copyTo(SheetLab.getRange("V5:V"));
  SpreadsheetApp.flush();

  SheetLab.getRange("W4").setValue("=IFS(ISBLANK(B4),\"\", A4 = \"PART\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/v2/catalog/catalogitem.page?P=\" & B4 & \"#T=P&C=\" & AD4), CONCATENATE(B4 & \" - \" & C4)), A4 = \"MINIFIG\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/v2/catalog/catalogitem.page?M=\" & B4),B4), A4 = \"SET\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/v2/catalog/catalogitem.page?S=\" & B4),B4))");
  SheetLab.getRange("W4").copyTo(SheetLab.getRange("W5:W"));
  SpreadsheetApp.flush();

  SheetLab.getRange("X4").setValue("=IFS(ISBLANK(B4),\"\", A4=\"PART\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?q=\" & B4), B4), A4=\"MINIFIG\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?q=\" & B4), B4), A4=\"SET\", HYPERLINK(CONCATENATE(\"https://www.bricklink.com/inventory_detail.asp?q=\" & LEFT(B4, LEN(B4)-2)), CONCATENATE(B4)))");
  SheetLab.getRange("X4").copyTo(SheetLab.getRange("X5:X"));
  SpreadsheetApp.flush();

  SheetLab.getRange("Y4").setValue("=IFERROR(VLOOKUP(AE4, Inventory!$H:$K,4,FALSE),\"\")");
  SheetLab.getRange("Y4").copyTo(SheetLab.getRange("Y5:Y"));

  SheetLab.getRange("Z4").setValue("=IFERROR(VLOOKUP(AE4, Inventory!$H:$M,5,FALSE),\"\")");
  SheetLab.getRange("Z4").copyTo(SheetLab.getRange("Z5:Z"));

  SheetLab.getRange("AC4").setValue("=IFERROR(VLOOKUP(AE4, Inventory!$H:$R,11,FALSE),\"\")");
  SheetLab.getRange("AC4").copyTo(SheetLab.getRange("AC5:AC"));
  SpreadsheetApp.flush();

  SheetLab.getRange("AD4").setValue("=XLOOKUP(C4, 'DB-Colors'!A:A,'DB-Colors'!B:B ,\"\",1)");
  SheetLab.getRange("AD4").copyTo(SheetLab.getRange("AD5:AD"));

  SheetLab.getRange("AE4").setValue("=IFERROR(IFS(A4 = \"PART\", A4 & \"_\" & B4 & \"_\" & AD4 & \"_\" & E4 & \"_\" & G4, A4 = \"MINIFIG\", A4 & \"_\" & B4 & \"_\" & E4 & \"_\" & G4, A4 = \"SET\", A4 & \"_\" & B4 & \"_\" & E4 & \"_\" & F4 & \"_\" & G4),\"\")");
  SheetLab.getRange("AE4").copyTo(SheetLab.getRange("AE5:AE"));
  SpreadsheetApp.flush();

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Lab', 'Lab is ready again!', Ui.ButtonSet.OK); 
}

// Function: Regenerate XML
function RegenerateXML() {
  RegenerateSheet("XML", '#FF00FF', 5000, 2)

  // Style
  var SheetXML = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("XML")
  SheetXML.setColumnWidths(1, 2, 850);
  SheetXML.setRowHeights(1, SheetXML.getMaxRows(), 21)
  SheetXML.getRange("A1:B1").setBackground('#D9D9D9').setFontWeight("bold");
  SheetXML.setFrozenRows(1);

  // UI
  var Ui = SpreadsheetApp.getUi();
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
  var Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
  
  // Color
  Sheet.setTabColor(SheetColor);
  
  // Rows
  if (Rows !== 0) {
    var SheetRows = Sheet.getMaxRows();    
    if (Rows < SheetRows){
      Sheet.deleteRows(Rows+1, SheetRows-Rows);
    } else if (Rows > SheetRows) {
       Sheet.insertRowsAfter(SheetRows, Rows-SheetRows);
    }
  }
  
  // Columns
  if (Columns !== 0) {
    var SheetColumns = Sheet.getMaxColumns();
    if (Columns < SheetColumns) {
      Sheet.deleteColumns(Columns+1, SheetColumns-Columns);
    } else if (Columns > SheetColumns) {
      Sheet.insertColumnsAfter(SheetColumns,Columns-SheetColumns);
    }
  }
  // Return new Sheet object
  return Sheet;
}