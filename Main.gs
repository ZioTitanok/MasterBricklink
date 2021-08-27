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
                .addItem('Update Rows (Prices)', 'LoadPricesBatch')
                .addItem('Update Row (Prices)', 'LoadPricesSingle')
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

    .addSubMenu(SpreadsheetApp.getUi().createMenu('Rigenera')
                .addItem('Regenerate Settings', 'RegenerateSettings')
                .addSeparator()
                .addItem('Regenerate DB-Colori',"RegenerateDBColors")
                .addItem('Regenerate DB-Category',"RegenerateDBCategory")
                .addItem('Regenerate DB-Items (Manual Import)', 'RegenerateDBItems')
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
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Credits', 'MASTERBRICKLINK: TOOLS FOR AFOLs.\n\r\n\r\Developed by Nico Mascagni (ZioTitanok) and Gianluca Cannalire (GianCann).\n\r\Docs on gitub: https://github.com/ZioTitanok/MasterBricklink.', Ui.ButtonSet.OK);
}
