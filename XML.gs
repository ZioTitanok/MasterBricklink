// Function: XML Wanted
function XMLWanted(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var LabActive = SheetSettings.getRange("B8").getValue()
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  
  var SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');
  var Selection = SheetLab.getDataRange();
  var StartingRow = 4;
  var EndingRow = SheetLab.getMaxRows();
 
  ClearXML()
  
  // Output, For Loop
  SheetXml.getRange(1,1).setValue('XML Wanted');
  SheetXml.getRange(2,1).setValue('<INVENTORY>');  
  var i = 1;
  var OutputWanted = [];

  for (var CurrentRow = StartingRow; CurrentRow <= EndingRow; CurrentRow++){
    var CellCodice = Selection.getCell(CurrentRow,LabColumnItemNo).getValue();
    
    if (CellCodice == ""){
      break;
    } else {
      
      var StringWanted = "<ITEM>" + "<ITEMTYPE>"
      if (Selection.getCell(CurrentRow,LabColumnItemType).getValue() == "PART"){
        StringWanted = StringWanted + "P</ITEMTYPE>"
      } else if (Selection.getCell(CurrentRow,LabColumnItemType).getValue() == "MINIFIG"){
        StringWanted = StringWanted + "M</ITEMTYPE>"
      } else if (Selection.getCell(CurrentRow,LabColumnItemType).getValue() == "SET"){
        StringWanted = StringWanted + "S</ITEMTYPE>"
        StringWanted = StringWanted + "<SUBCONDITION>" + Selection.getCell(CurrentRow,LabColumnCompleteness).getValue() + "</SUBCONDITION>"
      }

      StringWanted = StringWanted + "<ITEMID>" + Selection.getCell(CurrentRow,LabColumnItemNo).getValue() + "</ITEMID>";
      StringWanted = StringWanted + "<COLOR>" + Selection.getCell(CurrentRow,LabColumnColorID).getValue() + "</COLOR>";
      StringWanted = StringWanted + "<CONDITION>" + Selection.getCell(CurrentRow,LabColumnCondition).getValue() + "</CONDITION>";
      if (Selection.getCell(CurrentRow,LabColumnQty).getValue() != ""){
        StringWanted = StringWanted + "<MINQTY>" + Selection.getCell(CurrentRow,LabColumnQty).getValue() + "</MINQTY>";
      }
      if (Selection.getCell(CurrentRow, LabColumnCondition).getValue() != ""){
        StringWanted = StringWanted + "<CONDITION>" + Selection.getCell(CurrentRow, LabColumnCondition).getValue() + "</CONDITION>"
      }
      if (Selection.getCell(CurrentRow,LabColumnRemarks).getValue() != ""){
         StringWanted = StringWanted + "<REMARKS>" + Selection.getCell(CurrentRow,LabColumnRemarks).getValue() + "</REMARKS>";
      }
      StringWanted = StringWanted + "</ITEM>";
      
      OutputWanted[i] = StringWanted;
      i++;
      }
  }

  for (var k=1; k <= OutputWanted.length; k++) SheetXml.getRange(k+2,1).setValue(OutputWanted[k]);
  SheetXml.getRange(k+1,1).setValue('</INVENTORY>');

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('XML', 'XML Wanted created.' , Ui.ButtonSet.OK);
  
}

// Function: XML Upload and Update
function XMLUploadUpdate(){
  var SheetSettings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  var LabActive = SheetSettings.getRange("B8").getValue()
  var SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  
  var SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');
  var Selection = SheetLab.getDataRange();
  var StartingRow = 4;
  var EndingRow = SheetLab.getMaxRows();

  ClearXML()
  
  // Output, For Loop
  SheetXml.getRange(1,1).setValue('XML Update');
  SheetXml.getRange(1,2).setValue('XML Upload');
  SheetXml.getRange(2,1).setValue('<INVENTORY>');
  SheetXml.getRange(2,2).setValue('<INVENTORY>');
  var i = 1;
  var j = 1;
  var OutputUpdate = [];
  var OutputUpload = [];

  for (var CurrentRow = StartingRow; CurrentRow < EndingRow; CurrentRow++) {
    
    var CellCode = Selection.getCell(CurrentRow,LabColumnItemNo).getValue();
    var LotID = Selection.getCell(CurrentRow,LabColumnLotID).getValue();
    
    if (CellCode == ""){
      break;
    } else {
                 
      if (LotID != ""){
        // YES LotID: Update  
        if (Selection.getCell(CurrentRow,LabColumnQty).getValue() != "" || Selection.getCell(CurrentRow,LabColumnPrice).getValue() != "" || Selection.getCell(CurrentRow,LabColumnDescription).getValue() != "" || Selection.getCell(CurrentRow,LabColumnRemarks).getValue() != ""){
        // LotID presente (se almeno uno dei paramentri è cambiato): Update
          var StringUpdate = "<ITEM>";
          StringUpdate = StringUpdate + "<LOTID>" + Selection.getCell(CurrentRow,LabColumnLotID).getValue() + "</LOTID>";
          
          if (Selection.getCell(CurrentRow,LabColumnQty).getValue() != ""){
            if (Selection.getCell(CurrentRow, LabColumQtyInventory).getValue() + Selection.getCell(CurrentRow, LabColumnQty).getValue() == 0){
              StringUpdate = StringUpdate + "<DELETE>Y</DELETE>"
            } else {
              if (Selection.getCell(CurrentRow,LabColumnQty).getValue() > 0){
                var Sign = "+";
              } else {
                var Sign = "";
              }
              StringUpdate = StringUpdate + "<QTY>" + Sign + Selection.getCell(CurrentRow,LabColumnQty).getValue() + "</QTY>";
            }
          }

          if (Selection.getCell(CurrentRow,LabColumnPrice).getValue() != ""){
            StringUpdate = StringUpdate + "<PRICE>" + Selection.getCell(CurrentRow,LabColumnPrice).getValue() + "</PRICE>";
          }
          if (Selection.getCell(CurrentRow,LabColumnDescription).getValue() != ""){
            StringUpdate = StringUpdate + "<DESCRIPTION>" + Selection.getCell(CurrentRow,LabColumnDescription).getValue() + "</DESCRIPTION>";
          }
          if (Selection.getCell(CurrentRow,LabColumnRemarks).getValue() != ""){
            StringUpdate = StringUpdate + "<REMARKS>" + Selection.getCell(CurrentRow,LabColumnRemarks).getValue() + "</REMARKS>";
          }
          StringUpdate = StringUpdate + "</ITEM>";
          
          OutputUpdate[i] = StringUpdate;
          i++;
        }
        
      } else {
        // NO LotID: Upload        
        var StringUpload = "<ITEM>" + "<CATEGORY></CATEGORY>" + "<ITEMTYPE>"
        if (Selection.getCell(CurrentRow,LabColumnItemType).getValue() == "PART"){
          StringUpload = StringUpload + "P</ITEMTYPE>"
        } else if (Selection.getCell(CurrentRow,LabColumnItemType).getValue() == "MINIFIG"){
            StringUpload = StringUpload + "M</ITEMTYPE>"
        } else if (Selection.getCell(CurrentRow,LabColumnItemType).getValue() == "SET"){
            StringUpload = StringUpload + "S</ITEMTYPE>"
            StringUpload = StringUpload + "<SUBCONDITION>" + Selection.getCell(CurrentRow,LabColumnCompleteness).getValue() + "</SUBCONDITION>"
        }
        StringUpload = StringUpload + "<ITEMID>" + Selection.getCell(CurrentRow,LabColumnItemNo).getValue() + "</ITEMID>";
        StringUpload = StringUpload + "<COLOR>" + Selection.getCell(CurrentRow,LabColumnColorID).getValue() + "</COLOR>";
        StringUpload = StringUpload + "<CONDITION>" + Selection.getCell(CurrentRow,LabColumnCondition).getValue() + "</CONDITION>";
        StringUpload = StringUpload + "<QTY>" + Selection.getCell(CurrentRow,LabColumnQty).getValue() + "</QTY>";
        
        if (Selection.getCell(CurrentRow,LabColumnPrice).getValue() == ""){
          StringUpload = StringUpload + "<PRICE>" + Selection.getCell(CurrentRow,LabColumnPriceAvg).getValue() + "</PRICE>";
        } else {
          StringUpload = StringUpload + "<PRICE>" + Selection.getCell(CurrentRow,LabColumnPrice).getValue() + "</PRICE>";
        }
        
        if (Selection.getCell(CurrentRow,LabColumnDescription).getValue() != ""){
          StringUpload = StringUpload + "<DESCRIPTION>" + Selection.getCell(CurrentRow,LabColumnDescription).getValue() + "</DESCRIPTION>";
        }
        
        if (Selection.getCell(CurrentRow,LabColumnRemarks).getValue() != ""){
          StringUpload = StringUpload + "<REMARKS>" + Selection.getCell(CurrentRow,LabColumnRemarks).getValue() + "</REMARKS>";
        }
        
        if (Selection.getCell(CurrentRow,LabColumnStock).getValue() != "NO"){
          StringUpload = StringUpload + "<STOCKROOM>" + "Y" + "</STOCKROOM>";
          StringUpload = StringUpload + "<STOCKROOMID>" + Selection.getCell(CurrentRow,LabColumnStock).getValue() + "</STOCKROOMID>";
        }        
        StringUpload = StringUpload + "</ITEM>";

        OutputUpload[j] = StringUpload;
        j++;
      }
    }
  }

  for (var k=1; k <= OutputUpdate.length; k++) SheetXml.getRange(k+2,1).setValue(OutputUpdate[k]);
  SheetXml.getRange(k+1,1).setValue('</INVENTORY>');

  for (var l=1; l <= OutputUpload.length; l++) SheetXml.getRange(l+2,2).setValue(OutputUpload[l]);
  SheetXml.getRange(l+1,2).setValue('</INVENTORY>');

  // UI
  var Ui = SpreadsheetApp.getUi();
  Ui.alert('XML', 'XML Upload and Update created.', Ui.ButtonSet.OK);
}

// Function: Clear XML
function ClearXML(){
  var SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');
  var SheetXmlMinRow = 1;
  var SheetXmlMaxRow = SheetXml.getMaxRows();

  SheetXml.getRange(SheetXmlMinRow, 1, SheetXmlMaxRow, 2).clear({contentsOnly: true});
}