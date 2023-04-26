// OnOpen
function onOpen() {

    // Function: Menu
    SpreadsheetApp.getUi().createMenu('BrickLink Tool')

    .addSubMenu(SpreadsheetApp.getUi().createMenu('Inventory')
      .addItem('Download Inventory', 'LoadInventory')
      .addItem('Clear Inventory', 'ClearInventory'))

    .addSubMenu(SpreadsheetApp.getUi().createMenu('PartOut')
      .addItem('Download PartOut', 'LoadPartOut')
      .addItem('Clear PartOut', 'ClearPartOut'))
                
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Lab')
                .addItem('Update Prices (Bulk)', 'LoadPricesBulk')
                .addItem('Update Prices (Rows)', 'LoadPricesRows')
                .addItem('Update Prices (Row)', 'LoadPricesRow')
                .addSeparator()
                .addItem('Hint Prices','HintPrices')
                .addSeparator()
                .addItem('Import Inventory','ImportInventory')
                .addItem('Import PartOut', 'ImportPartOut')
                .addSeparator()
                .addItem('Clear Lab', 'ClearLab')
                .addItem('Clear Lab (Prices)', 'ClearLabPrices'))

    .addSubMenu(SpreadsheetApp.getUi().createMenu('XML')
                .addItem('Generate XML Upload/Update', 'XMLUploadUpdate')
                .addItem('Generate XML Wanted','XMLWanted')
                .addItem('Clear XML', 'ClearXML'))
    .addSeparator()

    .addSubMenu(SpreadsheetApp.getUi().createMenu('Regenerate')
                .addItem('Regenerate Settings', 'RegenerateSettings')
                .addSeparator()
                .addItem('Regenerate DB-Colors',"RegenerateDBColors")
                .addItem('Regenerate DB-Categories',"RegenerateDBCategories")
                .addItem('Regenerate DB-Part', 'RegenerateDBPart')
                .addItem('Regenerate DB-Minifigure', 'RegenerateDBMinifigure')
                .addItem('Regenerate DB-Set', 'RegenerateDBSet')
                .addItem('Regenerate DB-Codes', 'RegenerateDBCodes')
                .addSeparator()
                .addItem('Regenerate Inventory','RegenerateInventory')
                .addItem('Regenerate PartOut', 'RegeneratePartOut')
                .addItem('Regenerate XML', 'RegenerateXML')
                .addSeparator()
                .addItem('Regenerate Lab', 'RegenerateLab'))
    
    .addItem('Credits', 'Credits')
    .addToUi();  
  
}

// Function: Credits
function Credits(){
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Credits', 'MasterBricklink: Tools For AFOLs.\n\r\n\r\Developed by Nico Mascagni (ZioTitanok) and Gianluca Cannalire (GianCann).\n\r\Docs on gitub: https://github.com/ZioTitanok/MasterBricklink.', Ui.ButtonSet.OK);
}

// Constants
const BrickLinkBaseUrl = "https://api.bricklink.com/api/store/v1";
const BrickLinkOptions = {method: 'GET', contentType: 'application/json'};
const TBMBaseUrl = "https://django.ziotitanok.it/api";

// Function : Get Settings
function GetSettings() {
  const SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');

  const [BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret] = SheetSettings.getRange("B3:B6").getValues();
  
  const TBMTokenValue = SheetSettings.getRange("B9").getValue();
  const TBMHeaders = {'Authorization': 'Token ' + TBMTokenValue};
  
  const [DBAutoHide, LabActive, LabRowMaxPrice] = SheetSettings.getRange("B12:B14").getValues();

  return {BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret, TBMTokenValue, TBMHeaders, DBAutoHide, LabActive, LabRowMaxPrice};
}