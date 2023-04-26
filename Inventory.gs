// Constants: Inventory
const InventoryRowMin = 4;

// Function: Download Inventory from Bricklink
function LoadInventory(){
  const SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  const InventoryRowMax = SheetInventory.getMaxRows();
  SheetInventory.getRange(InventoryRowMin, 1, InventoryRowMax, 18).clearContent();

  // Data
  const ItemTypeValues = SheetInventory.getRange("A2:D2").getValues()[0];
  const ItemTypeLables = ["PART", "MINIFIG", "SET", "BOOK, GEAR, CATALOG, INSTRUCTION, UNSORTED_LOT, ORIGINAL_BOX"];

  var ItemType = [];
  for (var i = 0; i < ItemTypeLables.length; i++) {
    if (ItemTypeValues[i] == "YES") {
      ItemType.push(ItemTypeLables[i]);
    } else {
      if (i == ItemTypeLables.length - 1) {
        var Words = ItemTypeLables[i].split(", ");
        for (var j = 0; j < Words.length; j++) {
          ItemType.push("-" + Words[j]);
        }
      } else {
        ItemType.push("-" + ItemTypeLables[i]);
      }
    }
  }

  const CategoryId = SheetInventory.getRange("F1").getValue();
  const ColorId = SheetInventory.getRange("F2").getValue();
  var StockRoom = SheetInventory.getRange("H2").getValue();
  if (StockRoom == "NO") StockRoom = "Y";
  if (StockRoom == "A") StockRoom = "S";
  if (StockRoom == "B") StockRoom = "B";
  if (StockRoom == "C") StockRoom = "C";

  // API Request
  const {BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret} = GetSettings();
  const Url = `${BrickLinkBaseUrl}/inventories`;
  urlFetch = OAuth1.withAccessToken(BLConsumerKey, BLConsumerSecret, BLTokenValue, BLTokenSecret);
  var Params = {
     item_type: ItemType,
     category_id: CategoryId,
     color_id: ColorId,
     status: StockRoom
  }; 

  // Output 
  const OutputInventory = JSON.parse(urlFetch.fetch(Url, Params, BrickLinkOptions));
  var Inventory = OutputInventory.data.map((Item, Index) => {
    var OutputStockRoom = Item.stock_room_id || "NO";
    var OutputIndex = "";
    
    if (Item.item.type === "PART" || (Item.item.type !== "MINIFIG" && Item.item.type !== "SET")) {
      OutputIndex = `${Item.item.type}_${Item.item.no}_${Item.color_id}_${Item.new_or_used}_${OutputStockRoom}`;
    } else if (Item.item.type === "MINIFIG") {
      OutputIndex = `${Item.item.type}_${Item.item.no}_${Item.new_or_used}_${OutputStockRoom}`;
    } else if (Item.item.type === "SET") {
      OutputIndex = `${Item.item.type}_${Item.item.no}_${Item.new_or_used}_${Item.completeness}_${OutputStockRoom}`;
    }
    
    return [
      Index+1,
      Item.item.type,
      Item.item.no,
      Item.item.category_id,
      Item.item.name,
      Item.color_id,
      Item.color_name,
      OutputIndex,
      Item.quantity,
      Item.unit_price,
      Item.description,
      Item.remarks,
      Item.new_or_used,
      Item.completeness,
      Item.is_stock_room,
      OutputStockRoom,
      Item.inventory_id,
      Item.date_created
    ];
  });
  
  SheetInventory.getRange(InventoryRowMin, 1, Inventory.length, 18).setValues(Inventory).sort([2, 7, 5]);
  SheetInventory.getRange("R2").setValue(new Date());

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Inventory', 'Downloaded ' + Inventory.length + ' items from your Bricklink inventory.', Ui.ButtonSet.OK);
}
  
// Function: Clear Inventory
function ClearInventory(){
  const SheetInventory = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventory');
  const InventoryRowMax = SheetInventory.getMaxRows();
  SheetInventory.getRange(InventoryRowMin, 1, InventoryRowMax, 18).clearContent();
  SheetInventory.getRange("R2").clearContent();
  
  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('Inventory', 'Inventory is ready for new adventures!', Ui.ButtonSet.OK);
}