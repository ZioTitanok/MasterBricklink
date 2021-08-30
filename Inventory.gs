// Function: Download Inventory from Bricklink
function LoadInventory(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var ConsumerKey = SheetSettings.getRange("B3").getValue();
  var ConsumerSecret = SheetSettings.getRange("B4").getValue();
  var TokenValue = SheetSettings.getRange("B5").getValue();
  var TokenSecret = SheetSettings.getRange("B6").getValue();

  var SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var InventoryMinRow = 4;
  var InventoryMaxRow = SheetInventory.getMaxRows();

  SheetInventory.getRange(InventoryMinRow, 1, InventoryMaxRow, 18).clear({contentsOnly: true});

  // Data
  var ItemType = []
  if (SheetInventory.getRange("A2").getValue() == "YES") ItemType.push("PART");
  if (SheetInventory.getRange("B2").getValue() == "YES") ItemType.push("MINIFIG");
  if (SheetInventory.getRange("C2").getValue() == "YES") ItemType.push("SET");
  if (SheetInventory.getRange("D2").getValue() == "YES") ItemType = "";

  var CategoryId = SheetInventory.getRange("F1").getValue();
  var ColorId = SheetInventory.getRange("F2").getValue();
  var StockRoom = SheetInventory.getRange("H2").getValue();
  if (StockRoom == "NO") StockRoom = "Y";
  if (StockRoom == "A") StockRoom = "S";
  if (StockRoom == "B") StockRoom = "B";
  if (StockRoom == "C") StockRoom = "C";

  // API Request
  var Url = 'https://api.bricklink.com/api/store/v1' + '/inventories';
  var Options = {method: 'GET',contentType: 'application/json'};
  var Params = {
     item_type: ItemType,
     category_id: CategoryId,
     color_id: ColorId,
     status: StockRoom
  }; 
  urlFetch = OAuth1.withAccessToken(ConsumerKey, ConsumerSecret, TokenValue, TokenSecret);
  
  // Output 
  var OutputInventory = JSON.parse(urlFetch.fetch(Url, Params, Options));
  var Inventory = [];
  var i = 0;

  for (i in OutputInventory.data){
    var OutputStockRoom = OutputInventory.data[i].stock_room_id;
    if (OutputInventory.data[i].stock_room_id == undefined) OutputStockRoom = "NO";

    if (OutputInventory.data[i].item.type == "PART" || (OutputInventory.data[i].item.type != "MINIFIG" && OutputInventory.data[i].item.type != "SET")){
      var OutputIndex = OutputInventory.data[i].item.type + "_" + OutputInventory.data[i].item.no + "_" + OutputInventory.data[i].color_id + "_" + OutputInventory.data[i].new_or_used + "_" + OutputStockRoom
    } else if (OutputInventory.data[i].item.type == "MINIFIG"){
      var OutputIndex = OutputInventory.data[i].item.type + "_" + OutputInventory.data[i].item.no + "_" + OutputInventory.data[i].new_or_used + "_" + OutputStockRoom
    } else if (OutputInventory.data[i].item.type == "SET"){
      var OutputIndex = OutputInventory.data[i].item.type + "_" + OutputInventory.data[i].item.no + "_" + OutputInventory.data[i].new_or_used + "_" + OutputInventory.data[i].completeness + "_" + OutputStockRoom
    }
    
    Inventory[i] = [i,
                    OutputInventory.data[i].item.type,
                    OutputInventory.data[i].item.no,
                    OutputInventory.data[i].item.category_id,
                    OutputInventory.data[i].item.name,
                    OutputInventory.data[i].color_id,
                    OutputInventory.data[i].color_name,
                    OutputIndex,
                    OutputInventory.data[i].quantity,
                    OutputInventory.data[i].unit_price,
                    OutputInventory.data[i].description,
                    OutputInventory.data[i].remarks,
                    OutputInventory.data[i].new_or_used,
                    OutputInventory.data[i].completeness,
                    OutputInventory.data[i].is_stock_room,
                    OutputStockRoom,
                    OutputInventory.data[i].inventory_id,
                    OutputInventory.data[i].date_created
                   ]
  }
  i++;
  
  SheetInventory.getRange(InventoryMinRow, 1, Inventory.length, 18).setValues(Inventory);
  SheetInventory.getRange(InventoryMinRow, 1, Inventory.length, 18).sort([2, 7, 5]);
  SheetInventory.getRange(InventoryMinRow, 18).setValue(new Date());

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Inventory', 'Downloaded ' + i + ' items from your Bricklink inventory.', Ui.ButtonSet.OK);
}

// Function: Clear Inventory
function ClearInventory(){
  var SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  var InventoryMinRow = 4;
  var InventoryMaxRow = SheetInventory.getMaxRows();

  SheetInventory.getRange(InventoryMinRow, 1, InventoryMaxRow, 18).clear({contentsOnly: true});
  
  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('Inventory', 'Inventory is ready for new adventures!', Ui.ButtonSet.OK);
}
