////////////////////////////////////////
/////            XLM.gs            /////
////////////////////////////////////////

// Constants: XML
const SheetXmlRowMin = 2;

// Function: XML Add/Update/Want
function XMLGenerate(){
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');

  ClearXML()

  // Output, For Loop
  SheetXml.getRange(2,1).setValue('<INVENTORY>'); 
  SheetXml.getRange(2,2).setValue('<INVENTORY>'); 
  SheetXml.getRange(2,3).setValue('<INVENTORY>'); 

  const LabRowUsed = SheetLab.getRange(LabRowMin, 2, SheetLab.getLastRow(), 1).getValues().join('@').split('@');
  const Input = SheetLab.getRange(LabRowMin, 1, LabRowUsed.filter(Boolean).length, 30).getValues();

  var OutputAdd = [];
  var OutputUpdate = [];
  var OutputWanted = [];
  
  var a = 0;
  var u = 0;
  var w = 0;

  for (var i in Input){
    // LotID: YES, Update or Delete+Wanted
    if (Input[i][LabColumnLotID-1] != ""){
      if (Input[i][LabColumnQty-1] != "" || Input[i][LabColumnPrice-1] != "" || Input[i][LabColumnDescription-1] != "" || Input[i][LabColumnRemarks-1] != ""){

        // Qty: YES
        if (Input[i][LabColumnQty-1] != ""){
          var QtyDelta = parseInt(Input[i][LabColumQtyInventory-1]) + parseInt(Input[i][LabColumnQty-1]);

          if (QtyDelta > 0){
            var StringUpdate = "<ITEM>";
            StringUpdate += "<LOTID>" + Input[i][LabColumnLotID-1] + "</LOTID>";
            StringUpdate += "<QTY>" + Input[i][LabColumnQty-1] + "</QTY>";
            
            if (Input[i][LabColumnPrice-1] !=""){
              StringUpdate += "<PRICE>" + +Input[i][LabColumnPrice-1].toFixed(2) + "</PRICE>";
            }
            if (Input[i][LabColumnDescription-1] != ""){
              StringUpdate += "<DESCRIPTION>" + Input[i][LabColumnDescription-1] + "</DESCRIPTION>";
            }
            if (Input[i][LabColumnRemarks-1] != ""){
              StringUpdate += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
            }
            StringUpdate += "</ITEM>";
            
            OutputUpdate[u] = [StringUpdate];
            u++;
          }
          
          if (QtyDelta == 0){
            var StringUpdate = "<ITEM>";
            StringUpdate += "<LOTID>" + Input[i][LabColumnLotID-1] + "</LOTID>";

            StringUpdate += "<DELETE>Y</DELETE>"
            StringUpdate += "</ITEM>";
            
            OutputUpdate[u] = [StringUpdate];
            u++;
          }

          if (QtyDelta < 0){
            var StringUpdate = "<ITEM>";
            StringUpdate += "<LOTID>" + Input[i][LabColumnLotID-1] + "</LOTID>";

            StringUpdate += "<DELETE>Y</DELETE>"
            StringUpdate += "</ITEM>";
            
            OutputUpdate[u] = [StringUpdate];
            u++;

            var StringWanted = "<ITEM>" + "<ITEMTYPE>";
        
            if (Input[i][LabColumnItemType-1] == "PART"){
              StringWanted += "P</ITEMTYPE>"
            } else if (Input[i][LabColumnItemType-1] == "MINIFIG"){
              StringWanted += "M</ITEMTYPE>"
            } else if (Input[i][LabColumnItemType-1] == "SET"){
              StringWanted += "S</ITEMTYPE>";
              StringWanted += "<SUBCONDITION>" + Input[i][LabColumnCompleteness-1] + "</SUBCONDITION>"
            }

            StringWanted += "<ITEMID>" + Input[i][LabColumnItemNo-1] + "</ITEMID>";
            StringWanted += "<COLOR>" + Input[i][LabColumnColorID-1] + "</COLOR>";
            StringWanted += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>";
            
            StringWanted += "<MINQTY>" + Math.abs(QtyDelta) + "</MINQTY>";

            if (Input[i][LabColumnCondition-1] != ""){
              StringWanted += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>"
            }
            if (Input[i][LabColumnRemarks-1] != ""){
              StringWanted += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
            }
            StringWanted += "</ITEM>";
            
            OutputWanted[w] = [StringWanted];
            w++

          }

        } else {
        // Qty: NO
          var StringUpdate = "<ITEM>";
          StringUpdate += "<LOTID>" + Input[i][LabColumnLotID-1] + "</LOTID>";

          if (Input[i][LabColumnPrice-1] !=""){
            StringUpdate += "<PRICE>" + +Input[i][LabColumnPrice-1].toFixed(2) + "</PRICE>";
          }
          if (Input[i][LabColumnDescription-1] != ""){
            StringUpdate += "<DESCRIPTION>" + Input[i][LabColumnDescription-1] + "</DESCRIPTION>";
          }
          if (Input[i][LabColumnRemarks-1] != ""){
            StringUpdate += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
          }
          StringUpdate += "</ITEM>";
          
          OutputUpdate[u] = [StringUpdate];
          u++;
        }
      }
    } else {
    // LotID: NO, Add or Wanted
      if (Input[i][LabColumnQty-1] > 0){         
        var StringAdd = "<ITEM>" + "<CATEGORY></CATEGORY>" + "<ITEMTYPE>"

        if (Input[i][LabColumnItemType-1] == "PART"){
          StringAdd += "P</ITEMTYPE>"
        } else if (Input[i][LabColumnItemType-1] == "MINIFIG"){
            StringAdd+= "M</ITEMTYPE>"
        } else if (Input[i][LabColumnItemType-1] == "SET"){
            StringAdd += "S</ITEMTYPE>"
            StringAdd += "<SUBCONDITION>" + Input[i][LabColumnCompleteness-1] + "</SUBCONDITION>"
        }

        StringAdd += "<ITEMID>" + Input[i][LabColumnItemNo-1] + "</ITEMID>";
        StringAdd += "<COLOR>" + Input[i][LabColumnColorID-1] + "</COLOR>";
        StringAdd += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>";
        StringAdd += "<QTY>" + Input[i][LabColumnQty-1] + "</QTY>";
        
        if (Input[i][LabColumnPrice-1] == ""){
          StringAdd += "<PRICE>" + +Input[i][LabColumnPriceAvg-1].toFixed(2) + "</PRICE>";
        } else {
          StringAdd += "<PRICE>" + +Input[i][LabColumnPrice-1].toFixed(2) + "</PRICE>";
        }
            
        if (Input[i][LabColumnDescription-1] != ""){
          StringAdd += "<DESCRIPTION>" + Input[i][LabColumnDescription-1] + "</DESCRIPTION>";
        }
            
        if (Input[i][LabColumnRemarks-1] != ""){
          StringAdd += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
        }
        
        if (Input[i][LabColumnStock-1] != "NO"){
          StringAdd += "<STOCKROOM>" + "Y" + "</STOCKROOM>";
          StringAdd += "<STOCKROOMID>" + Input[i][LabColumnStock-1] + "</STOCKROOMID>";
        }        
        StringAdd += "</ITEM>";

        OutputAdd[a] = [StringAdd];
        a++;

      } else {
        var StringWanted = "<ITEM>" + "<ITEMTYPE>";
    
        if (Input[i][LabColumnItemType-1] == "PART"){
          StringWanted += "P</ITEMTYPE>"
        } else if (Input[i][LabColumnItemType-1] == "MINIFIG"){
          StringWanted += "M</ITEMTYPE>"
        } else if (Input[i][LabColumnItemType-1] == "SET"){
          StringWanted += "S</ITEMTYPE>";
          StringWanted += "<SUBCONDITION>" + Input[i][LabColumnCompleteness-1] + "</SUBCONDITION>"
        }

        StringWanted += "<ITEMID>" + Input[i][LabColumnItemNo-1] + "</ITEMID>";
        StringWanted += "<COLOR>" + Input[i][LabColumnColorID-1] + "</COLOR>";
        StringWanted += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>";
        if (Input[i][LabColumnQty-1] != ""){
          StringWanted += "<MINQTY>" + Math.abs(Input[i][LabColumnQty-1]) + "</MINQTY>";
        }
        if (Input[i][LabColumnCondition-1] != ""){
          StringWanted += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>"
        }
        if (Input[i][LabColumnRemarks-1] != ""){
          StringWanted += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
        }
        StringWanted += "</ITEM>";
        
        OutputWanted[w] = [StringWanted];
        w++
      }
    }
  }

  if (OutputAdd.length > 0){
    SheetXml.getRange(3, 1, OutputAdd.length, 1).setValues(OutputAdd);
    SheetXml.getRange(OutputAdd.length+3,1).setValue('</INVENTORY>');
  } else { SheetXml.getRange(3,1).setValue('</INVENTORY>') };

  if (OutputUpdate.length > 0){
    SheetXml.getRange(3, 2, OutputUpdate.length, 1).setValues(OutputUpdate);
    SheetXml.getRange(OutputUpdate.length+3,2).setValue('</INVENTORY>');
  } else { SheetXml.getRange(3,2).setValue('</INVENTORY>') };

  if (OutputWanted.length > 0){
    SheetXml.getRange(3, 3, OutputWanted.length, 1).setValues(OutputWanted);
    SheetXml.getRange(OutputWanted.length+3,3).setValue('</INVENTORY>');
  } else { SheetXml.getRange(3,3).setValue('</INVENTORY>') };

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('XML', 'XML generated.', Ui.ButtonSet.OK);
}

// Function: Clear XML
function ClearXML(){
  const SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');
  const SheetXmlRowMax = SheetXml.getMaxRows();

  SheetXml.getRange(SheetXmlRowMin, 1, SheetXmlRowMax, 3).clearContent();
}