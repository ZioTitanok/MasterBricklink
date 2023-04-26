// Constants: XML
const SheetXmlRowMin = 1;

// Function: XML Wanted
function XMLWanted(){
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);
  const SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');
 
  ClearXML()

  // Output, For Loop
  SheetXml.getRange(1,1).setValue('XML Wanted');
  SheetXml.getRange(2,1).setValue('<INVENTORY>');  

  const LabRowUsed = SheetLab.getRange(LabRowMin, 2, SheetLab.getLastRow(), 1).getValues().join('@').split('@');
  const Input = SheetLab.getRange(LabRowMin, 1, LabRowUsed.filter(Boolean).length, 30).getValues();

  var OutputWanted = [];
  for (var i in Input){
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
      StringWanted += "<MINQTY>" + Input[i][LabColumnQty-1] + "</MINQTY>";
    }
    if (Input[i][LabColumnCondition-1] != ""){
      StringWanted += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>"
    }
    if (Input[i][LabColumnRemarks-1] != ""){
       StringWanted += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
    }
    StringWanted += "</ITEM>";
    
    OutputWanted[i] = [StringWanted];
  }
  
  SheetXml.getRange(3, 1, OutputWanted.length, 1).setValues(OutputWanted);
  SheetXml.getRange(OutputWanted.length+3,1).setValue('</INVENTORY>');

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('XML', 'XML Wanted created.' , Ui.ButtonSet.OK);
  
}

// Function: XML Upload and Update
function XMLUploadUpdate(){
  const {LabActive} = GetSettings();
  const SheetLab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LabActive);  
  const SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');

  ClearXML()
  
  // Output, For Loop
  SheetXml.getRange(1,1).setValue('XML Update');
  SheetXml.getRange(1,2).setValue('XML Upload');
  SheetXml.getRange(2,1).setValue('<INVENTORY>');
  SheetXml.getRange(2,2).setValue('<INVENTORY>');

  const LabRowUsed = SheetLab.getRange(LabRowMin, 2, SheetLab.getLastRow(), 1).getValues().join('@').split('@');
  const Input = SheetLab.getRange(LabRowMin, 1, LabRowUsed.filter(Boolean).length, 30).getValues();

  var OutputUpdate = [];
  var OutputUpload = [];
  var j = 0;
  var k = 0;

  for (var i in Input){
    // YES LotID: Update
    if (Input[i][LabColumnLotID-1] != ""){
      if (Input[i][LabColumnQty-1] != "" || Input[i][LabColumnPrice-1] != "" || Input[i][LabColumnDescription-1] != "" || Input[i][LabColumnRemarks-1] != ""){

        var StringUpdate = "<ITEM>";
        StringUpdate += "<LOTID>" + Input[i][LabColumnLotID-1] + "</LOTID>";

        if (Input[i][LabColumnQty-1] != ""){
          if (Input[i][LabColumQtyInventory-1] + Input[i][LabColumnQty-1] == 0){
            StringUpdate += "<DELETE>Y</DELETE>"
          } else {
            if (Input[i][LabColumnQty-1] > 0){
              var Sign = "+";
            } else {
              var Sign = "";
            }
            StringUpdate += "<QTY>" + Sign + Input[i][LabColumnQty-1] + "</QTY>";
          }
        }
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
        
        OutputUpdate[j] = [StringUpdate];
        j++;
      }

    } else {
      // NO LotID: Upload        
      var StringUpload = "<ITEM>" + "<CATEGORY></CATEGORY>" + "<ITEMTYPE>"
      if (Input[i][LabColumnItemType-1] == "PART"){
        StringUpload += "P</ITEMTYPE>"
      } else if (Input[i][LabColumnItemType-1] == "MINIFIG"){
          StringUpload+= "M</ITEMTYPE>"
      } else if (Input[i][LabColumnItemType-1] == "SET"){
          StringUpload += "S</ITEMTYPE>"
          StringUpload += "<SUBCONDITION>" + Input[i][LabColumnCompleteness-1] + "</SUBCONDITION>"
      }
      StringUpload += "<ITEMID>" + Input[i][LabColumnItemNo-1] + "</ITEMID>";
      StringUpload += "<COLOR>" + Input[i][LabColumnColorID-1] + "</COLOR>";
      StringUpload += "<CONDITION>" + Input[i][LabColumnCondition-1] + "</CONDITION>";
      StringUpload += "<QTY>" + Input[i][LabColumnQty-1] + "</QTY>";
      
      if (Input[i][LabColumnPrice-1] == ""){
        StringUpload += "<PRICE>" + +Input[i][LabColumnPriceAvg-1].toFixed(2) + "</PRICE>";
      } else {
        StringUpload += "<PRICE>" + +Input[i][LabColumnPrice-1].toFixed(2) + "</PRICE>";
        console.log(Input[i][LabColumnPrice-1].toFixed(2))
      }
          
      if (Input[i][LabColumnDescription-1] != ""){
        StringUpload += "<DESCRIPTION>" + Input[i][LabColumnDescription-1] + "</DESCRIPTION>";
      }
          
      if (Input[i][LabColumnRemarks-1] != ""){
        StringUpload += "<REMARKS>" + Input[i][LabColumnRemarks-1] + "</REMARKS>";
      }
      
      if (Input[i][LabColumnStock-1] != "NO"){
        StringUpload += "<STOCKROOM>" + "Y" + "</STOCKROOM>";
        StringUpload += "<STOCKROOMID>" + Input[i][LabColumnStock-1] + "</STOCKROOMID>";
      }        
      StringUpload += "</ITEM>";

      OutputUpload[k] = [StringUpload];
      k++;
    }
  }

  if (OutputUpdate.length > 0){
    SheetXml.getRange(3, 1, OutputUpdate.length, 1).setValues(OutputUpdate);
    SheetXml.getRange(OutputUpdate.length+3,1).setValue('</INVENTORY>');
  } else { SheetXml.getRange(3,1).setValue('</INVENTORY>') };

  if (OutputUpload.length > 0){
    SheetXml.getRange(3, 2, OutputUpload.length, 1).setValues(OutputUpload);
    SheetXml.getRange(OutputUpload.length+3,2).setValue('</INVENTORY>');
  } else { SheetXml.getRange(3,2).setValue('</INVENTORY>') };

  // UI
  const Ui = SpreadsheetApp.getUi();
  Ui.alert('XML', 'XML Upload and Update created.', Ui.ButtonSet.OK);
}

// Function: Clear XML
function ClearXML(){
  const SheetXml = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('XML');
  const SheetXmlRowMax = SheetXml.getMaxRows();

  SheetXml.getRange(SheetXmlRowMin, 1, SheetXmlRowMax, 2).clearContent();
}