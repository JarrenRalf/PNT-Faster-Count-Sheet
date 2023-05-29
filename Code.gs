/**
 * This function runs the installed on edit trigger, which detects scanned barcodes, item searches, and various other commands 
 * that allow us to update the upc database.
 * 
 * @author Jarren Ralf
 */
function installedOnEdit(e)
{
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  var sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === "Manual Scan") // Check if a barcode has been scanned
      manualScan(e, spreadsheet, sheet)
    else if (sheetName === "Item Search") 
      search(e, spreadsheet, sheet)
    else if (sheetName === 'Scan Log')
      updateCounts(e, sheet)
  } 
  catch (error) 
  {
    Logger.log(error)
    Browser.msgBox(error)
  }
}

/**
 * This function allows the user to select items on the Scan Log page and move them to the UPC Database and Manually Added UPCs pages.
 * In turn, this will now allow the items to be searchable via a Manual Scan.
 * 
 * @author Jarren Ralf
 */
function addItemsToUpcData()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");

  if (sheet.getSheetName() === 'Manual Scan')
  { 
    const ui = SpreadsheetApp.getUi();
    const barcodeInputRange = sheet.getRange(1, 1);
    const values = barcodeInputRange.getValue().split('\n');

    const response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + values[0] +'.', ui.ButtonSet.OK_CANCEL)
    {
      if (response.getSelectedButton() == ui.Button.OK)
      {
        const item = values[0].split(' - ');
        const upc = response.getResponseText();
        const numUpcs = upcDatabaseSheet.getLastRow();
        var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);

        var l = 0; // Lower-bound
        var u = numUpcs - 1; // Upper-bound
        var m = Math.ceil((u + l)/2) // Midpoint

        while (l < m && u > m)
        {
          if (upc < upcDatabase[m][0])
            u = m;   
          else if (upc > upcDatabase[m][0])
            l = m;

          m = Math.ceil((u + l)/2) // Midpoint
        }

        if (upc < upcDatabase[0][0])
          upcDatabase.splice(0, 0, [upc, item[0]])
        else if (upc < upcDatabase[m][0])
          upcDatabase.splice(m, 0, [upc, item[0]])
        else
          upcDatabase.splice(m + 1, 0, [upc, item[0]])
          
        upcDatabaseSheet.getRange(1, 1, numUpcs + 1, 2).setValues(upcDatabase);
        manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([[item[0], upc, values[0]]]);
        barcodeInputRange.activate();
      }
    }
  }
  else if (sheet.getSheetName() === 'Scan Log' || sheet.getSheetName() === 'Item Search')
  {
    const ui = SpreadsheetApp.getUi();
    const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
    var activeRanges = sheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
    var firstRows = [], lastRows = [], numRows = [], itemValues = [[[]]], response, item, upc, l, u, m;
    
    // Find the first row and last row in the the set of all active ranges
    for (var r = 0; r < activeRanges.length; r++)
    {
      firstRows[r] = activeRanges[r].getRow();
      lastRows[r] = activeRanges[r].getLastRow()
    }
    
    for (var r = 0; r < activeRanges.length; r++)
    {
        numRows[r] = lastRows[r] - firstRows[r] + 1;
      itemValues[r] = sheet.getSheetValues(firstRows[r], 1, numRows[r], 1);
    }
    
    var itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
    var numItems = itemVals.length;
    var numUpcs = upcDatabaseSheet.getLastRow();
    var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);

    var upcTemporaryValues = itemVals.map(v => {
      response = ui.prompt('Manually Add UPCs', 'Please scan the barcode for:\n\n' + v[0] +'.', ui.ButtonSet.OK_CANCEL)

      if (ui.Button.OK === response.getSelectedButton())
      {
        item = v[0].split(' - ');
        upc = response.getResponseText();
        
        l = 0; // Lower-bound
        u = numUpcs - 1; // Upper-bound
        m = Math.ceil((u + l)/2) // Midpoint

        while (l < m && u > m)
        {
          if (upc < upcDatabase[m][0])
            u = m;   
          else if (upc > upcDatabase[m][0])
            l = m;

          m = Math.ceil((u + l)/2) // Midpoint
        }

        if (upc < upcDatabase[0][0])
          upcDatabase.splice(0, 0, [upc, v[0]])
        else if (upc < upcDatabase[m][0])
          upcDatabase.splice(m, 0, [upc, v[0]])
        else
          upcDatabase.splice(m + 1, 0, [upc, v[0]])

        numUpcs = upcDatabase.length

        return [item[0], upc, v[0]]
      }
      else
        return [null, null, null]
    });

    upcDatabaseSheet.getRange(1, 1, upcDatabase.length, 2).setValues(upcDatabase);
    manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, numItems, 3).setNumberFormat('@').setValues(upcTemporaryValues);
    spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
  }
}

/**
 * This function allows the user to select items on the Scan Log page and move them to the UPC Database and Manually Added UPCs pages.
 * In turn, this will now allow the items to be searchable via a Manual Scan. In this case, the item/s in question appear not to be found in the Adagio database.
 * 
 * @author Jarren Ralf
 */
function addItemsToUpcData_ItemsNotFound()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  const inventorySheet = spreadsheet.getSheetByName("Inventory");
  const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
  const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
  var activeRanges = sheet.getActiveRangeList().getRanges(); // The selected ranges on the item search sheet
  var firstRows = [], lastRows = [], numRows = [], itemValues = [[[]]], response, response2, item, itemJoined, upc, itemTemporaryValues = [], l, u, m;
  
  // Find the first row and last row in the the set of all active ranges
  for (var r = 0; r < activeRanges.length; r++)
  {
    firstRows[r] = activeRanges[r].getRow();
    lastRows[r] = activeRanges[r].getLastRow()
  }
  
  for (var r = 0; r < activeRanges.length; r++)
  {
       numRows[r] = lastRows[r] - firstRows[r] + 1;
    itemValues[r] = sheet.getSheetValues(firstRows[r], 1, numRows[r], 1);
  }
  
  var itemVals = [].concat.apply([], itemValues); // Concatenate all of the item values as a 2-D array
  var numItems = itemVals.length;
  var numUpcs = upcDatabaseSheet.getLastRow();
  var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);

  var upcTemporaryValues = itemVals.map(() => {
    response = ui.prompt('Item Not Found', 'Please enter a new description:', ui.ButtonSet.OK_CANCEL)

    if (ui.Button.OK === response.getSelectedButton())
    {
      item = response.getResponseText().split(' - ');
      item[0] = 'MAKE_NEW_SKU';
      itemJoined = item.join(' - ')
      response2 = ui.prompt('Item Not Found', 'Please scan the barcode for:\n\n' + itemJoined +'.', ui.ButtonSet.OK_CANCEL)

      if (ui.Button.OK === response2.getSelectedButton())
      {
        upc = response2.getResponseText();

        l = 0; // Lower-bound
        u = numUpcs - 1; // Upper-bound
        m = Math.ceil((u + l)/2) // Midpoint

        while (l < m && u > m)
        {
          if (upc < upcDatabase[m][0])
            u = m;   
          else if (upc > upcDatabase[m][0])
            l = m;

          m = Math.ceil((u + l)/2) // Midpoint
        }

        if (upc < upcDatabase[0][0])
          upcDatabase.splice(0, 0, [upc, itemJoined])
        else if (upc < upcDatabase[m][0])
          upcDatabase.splice(m, 0, [upc, itemJoined])
        else
          upcDatabase.splice(m + 1, 0, [upc, itemJoined])

        itemTemporaryValues.push([itemJoined])
        numUpcs = upcDatabase.length

        return ['MAKE_NEW_SKU', upc, itemJoined]
      }
      else
      {
        itemTemporaryValues.push([null])
        return [null, null, null]
      }
    }
    else
      return [null, null, null]
  });

  upcDatabaseSheet.getRange(1, 1, upcDatabase.length, 2).setValues(upcDatabase);
  inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1, numItems).setValues(itemTemporaryValues)
  manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, numItems, 3).setNumberFormat('@').setValues(upcTemporaryValues);
  spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
}

/**
 * This function checks if a given value is precisely a non-blank string.
 * 
 * @param  {String}  value : A given string.
 * @return {Boolean} Returns a boolean based on whether an inputted string is not-blank or not.
 * @author Jarren Ralf
 */
function isNotBlank(value)
{
  return value !== '';
}

/**
* This function checks if the given input is a number or not.
*
* @param {Object} num The inputted argument, assumed to be a number.
* @return {Boolean} Returns a boolean reporting whether the input paramater is a number or not
* @author Jarren Ralf
*/
function isNumber(num)
{
  return !(isNaN(Number(num)));
}

/**
 * This function watches two cells and if the left one is edited then it searches the UPC Database for the upc value (the barcode that was scanned).
 * It then checks if the item is on the scan log and stores the relevant data in the left cell. If the right cell is edited, then the function
 * uses the data in the left cell and moves the item over to the scan log with the updated quantity.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function manualScan(e, spreadsheet, sheet)
{
  const barcodeInputRange = e.range;

  if (barcodeInputRange.columnEnd === 1) // Barcode is scanned
  {
    const upcCode = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
      .setFontFamily("Arial").setFontColor("black").setFontSize(25)                     // Set the font parameters
      .setVerticalAlignment("middle").setHorizontalAlignment("center")                  // Set the alignment parameters
      .getValue();

    if (isNotBlank(upcCode)) // The user may have hit the delete key
    {
      const scanLogPage = spreadsheet.getSheetByName("Scan Log");
      const lastRow = scanLogPage.getLastRow();
      const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
      const upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, upcDatabaseSheet.getLastRow(), 1);

      var l = 0; // Lower-bound
      var u = upcDatabase.length - 1; // Upper-bound
      var m = Math.ceil((u + l)/2) // Midpoint

      if (lastRow == 0) // There are no items on the scan log
      {
        while (l < m && u > m)
        {
          if (upcCode < upcDatabase[m][0])
            u = m;   
          else if (upcCode > upcDatabase[m][0])
            l = m;
          else
          {
            barcodeInputRange.setValue(upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0] + '\nwill be added to the Scan Log at line :\n' + 1);
            break; // Item was found, therefore stop searching
          }

          m = Math.ceil((u + l)/2) // Midpoint
        }
      }
      else // There are existing items on the scan log
      {
        const row = lastRow + 1;
        const scanLog = scanLogPage.getSheetValues(1, 1, row - 1, 3);

        while (l < m && u > m)
        {
          if (upcCode < upcDatabase[m][0])
            u = m;   
          else if (upcCode > upcDatabase[m][0])
            l = m;
          else
          {
            const itemDescription = upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0];

            for (var j = 0; j < scanLog.length; j++) // Loop through the scan log
            {
              if (scanLog[j][0] === itemDescription) // The description matches
              {
                barcodeInputRange.setValue(itemDescription  + '\nwas found on the Scan Log at line :\n' + (j + 1) 
                                                                                + '\nCurrent Manual Count :\n' + scanLog[j][1] 
                                                                                + '\nCurrent Running Sum :\n'  + scanLog[j][2]);
                break; // Item was found on the scan log, therefore stop searching
              }
            }

            if (j === scanLog.length) // Item was not found on the scan log
              barcodeInputRange.setValue(upcDatabaseSheet.getSheetValues(m + 1, 2, 1, 1)[0][0] + '\nwill be added to the Scan Log at line :\n' + row);

            break; // Item was found, therefore stop searching
          }

          m = Math.ceil((u + l)/2) // Midpoint
        }
      }

      if (l >= m || u <= m)
      {
        if (upcCode.toString().length > 25)
          barcodeInputRange.offset(0, 0, 1, 2).setValues([['Barcode is Not Found.', '']]);
        else
          barcodeInputRange.offset(0, 0, 1, 2).setValues([['Barcode:\n\n' + upcCode + '\n\n is NOT FOUND.', '']]);

        barcodeInputRange.activate()
      }
      else
        barcodeInputRange.offset(0, 1, 1, 1).setValue('').activate();
    }
  }
  else if (barcodeInputRange.columnStart !== 1) // Quantity is entered
  {
    const quantity = barcodeInputRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP) // Wrap strategy for the cell
      .setFontFamily("Arial").setFontColor("black").setFontSize(25)                      // Set the font parameters
      .setVerticalAlignment("middle").setHorizontalAlignment("center")                   // Set the alignment parameters
      .getValue();

    if (isNotBlank(quantity)) // The user may have hit the delete key
    {
      const scanLogPage = spreadsheet.getSheetByName("Scan Log");
      const item = sheet.getSheetValues(1, 1, 1, 1)[0][0].split('\n'); // The information from the left cell that is used to move the item to the scan log

      if (quantity <= 100000) // If false, Someone probably scanned a barcode in the quantity cell (not likely to have counted an inventory amount of 100 000)
      {
        if (item.length !== 1) // The cell to the left contains valid item information
        {
          if (item[1].split(' ')[0] === 'was') // The item was already on the scan log
          {
            const range = scanLogPage.getRange(item[2], 2, 1, 2);
            const itemValues = range.getValues()
            const updatedCount = Number(itemValues[0][0]) + quantity;
            const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(quantity) === 1 || Math.sign(quantity) === 0)  ? 
                                                                  String(itemValues[0][1]) + ' \+ ' + String(   quantity)  : 
                                                                  String(itemValues[0][1]) + ' \- ' + String(-1*quantity)) :
                                                                    ((isNotBlank(itemValues[0][0])) ? 
                                                                      String(itemValues[0][0]) + ' \+ ' + String(quantity) : 
                                                                      String(quantity));
            range.setNumberFormats([['#.#', '@']]).setValues([[updatedCount, runningSum]])
            sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Scan Log at line :\n' + item[2] 
                                                            + '\nCurrent Manual Count :\n' + updatedCount 
                                                            + '\nCurrent Running Sum :\n' + runningSum,
                                                            '']]);
          }
          else
          {
            scanLogPage.getRange(scanLogPage.getLastRow() + 1, 1, 1, 3).setNumberFormats([['@', '#.#', '@']]).setValues([[item[0], quantity, '\'' + String(quantity)]])
            sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas added to the Scan Log at line :\n' + item[2] 
                                                            + '\nCurrent Manual Count :\n' + quantity,
                                                            '']]);
          }
        }
        else // The cell to the left does not contain the necessary item information to be able to move it to the manual counts page
          barcodeInputRange.setValue('Please scan your barcode in the left cell again.')

        sheet.getRange(1, 1).activate();
      }
      else if (quantity.toString().toLowerCase() === 'clear')
      {
        scanLogPage.getRange(item[2], 2, 1, 2).setValues([['', '']])
        sheet.getRange(1, 1, 1, 2).setValues([[item[0]  + '\nwas found on the Scan Log at line :\n' + item[2] 
                                                        + '\nCurrent Manual Count :\n\nCurrent Running Sum :\n',
                                                        '']]);
      }
      else if (quantity.toString().split(' ', 1)[0] === 'uuu') // Unmarry upc
      {
        const upc = quantity.split(' ')[1];

        if (upc > 100000)
        {
          const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
          unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[upc, item[0]]]);
          barcodeInputRange.setValue('UPC Code has been added to the unmarry list.')
          sheet.getRange(1, 1).activate();
        }
        else
          barcodeInputRange.setValue('Please enter a valid UPC Code to unmarry.')
      }
      else if (quantity.toString().split(' ', 1)[0] === 'mmm') // Marry upc
      {
        const upc = quantity.split(' ')[1];

        const sku = item[0].split(' - ', 1)[0];
        const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
        const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
        const numUpcs = upcDatabaseSheet.getLastRow();
        var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);

        var l = 0; // Lower-bound
        var u = numUpcs - 1; // Upper-bound
        var m = Math.ceil((u + l)/2) // Midpoint

        while (l < m && u > m)
        {
          if (upc < upcDatabase[m][0])
            u = m;   
          else if (upc > upcDatabase[m][0])
            l = m;

          m = Math.ceil((u + l)/2) // Midpoint
        }

        if (upc < upcDatabase[0][0])
          upcDatabase.splice(0, 0, [upc, item[0]])
        else if (upc < upcDatabase[m][0])
          upcDatabase.splice(m, 0, [upc, item[0]])
        else
          upcDatabase.splice(m + 1, 0, [upc, item[0]])
          
        upcDatabaseSheet.getRange(1, 1, numUpcs + 1, 2).setValues(upcDatabase);
        manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([[sku, upc, item[0]]]);
        barcodeInputRange.setValue('UPC Code has been added to the database temporarily.')
        sheet.getRange(1, 1).activate();
      }
      else 
        barcodeInputRange.setValue('Please enter a valid quantity.')
    }
  }
}

/**
 * This function moves the user to the search box on the Item Search page
 * 
 * @author Jarren Ralf
 */
function moveToItemSearch()
{
  SpreadsheetApp.getActive().getSheetByName('Item Search').getRange(1, 1).activate();
}

/**
 * This function moves the user to the barcode input cell (left) on the Manual Scan page
 * 
 * @author Jarren Ralf
 */
function moveToManualScan()
{
  SpreadsheetApp.getActive().getSheetByName('Manual Scan').getRange(1, 1).activate();
}

/**
 * This function moves the user to the Scan Log page.
 * 
 * @author Jarren Ralf
 */
function moveToScanLog()
{
  SpreadsheetApp.getActive().getSheetByName('Scan Log').activate();
}

/**
 * This function moves the user to the UPC Database page.
 * 
 * @author Jarren Ralf
 */
function moveToUpcDatabse()
{
  SpreadsheetApp.getActive().getSheetByName('UPC Database').activate();
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the Inventory page for the items in question.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function search(e, spreadsheet, sheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const colEnd = range.columnEnd;

  if (row == range.rowEnd && (colEnd == null || col == colEnd)) // Check and make sure only a single cell is being edited
  {
    if (row === 1 && col === 1) // Check if the search box is edited
    {
      const startTime = new Date()
      const MAX_NUM_ITEMS = 1500;
      const searchWords = sheet.getRange(1, 1).clearFormat()                                            // Clear the formatting of the range of the search box
        .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
        .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
        .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
        .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
        .getValue().toString().toLowerCase().split(/\s+/);                                              // Split the search string at whitespacecharacters into an array of search words

      const itemSearchFullRange = sheet.getRange(2, 1, MAX_NUM_ITEMS + 1, 2); // The entire range of the Item Search page

      if (isNotBlank(searchWords[0])) // If the value in the search box is NOT blank, then compute the search
      {
        const inventorySheet = spreadsheet.getSheetByName('INVENTORY');
        const data = inventorySheet.getSheetValues(1, 1, inventorySheet.getLastRow(), 1);
        const numSearchWords = searchWords.length - 1; // The number of search words - 1
        const output = [];

        for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
        {
          for (var j = 0; j <= numSearchWords; j++) // Loop through each word in the User's query
          {
            if (data[i][0].toString().toLowerCase().includes(searchWords[j])) // Does the i-th item description contain the j-th search word
            {
              if (j === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                output.push([data[i][0], '']);
            }
            else
              break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
          }
        }

        const numItems = output.length;

        if (numItems === 0) // No items were found
        {
          sheet.getRange('A1').activate(); // Move the user back to the seachbox
          itemSearchFullRange.clearContent(); // Clear content
          sheet.getRange('A2').setValue('No results found.\t\tPlease try again.')
        }
        else
        {
          if (numItems > MAX_NUM_ITEMS) // Over MAX_NUM_ITEMS items were found
          {
            sheet.getRange('B3').activate(); // Move the user to the top of the search items
            output.splice(MAX_NUM_ITEMS); // Slice off all the entries after MAX_NUM_ITEMS
            output.unshift([numItems + " results found, only " + MAX_NUM_ITEMS + " displayed.", "Count"])
            itemSearchFullRange.setValues(output);
          }
          else // Less than MAX_NUM_ITEMS items were found
          {
            sheet.getRange('B3').activate(); // Move the user to the top of the search items
            itemSearchFullRange.clearContent(); // Clear content and reset the text format
            output.unshift([numItems + " results found.", "Count"])
            sheet.getRange(2, 1, numItems + 1, 2).setValues(output);
          }
        }
      }
      else
        itemSearchFullRange.clearContent(); // Clear content 

      sheet.getRange('B1').setValue((new Date().getTime() - startTime)/1000 + " s");
    }
    else if (col === 2) // Check for the user trying to marry / unmarry upcs, add a new item, or update the count in the Scan Log
    {
      if (userHasNotPressedDelete(e.value))
      {
        if (isNumber(e.value))
        {
          if (quantity <= 100000) // If false, Someone probably scanned a barcode in the quantity cell (not likely to have counted an inventory amount of 100 000)
          {
            const scanLogPage = spreadsheet.getSheetByName("Scan Log");
            const lastRow = scanLogPage.getLastRow();

            if (lastRow == 0) // There are no items on the scan log
              scanLogPage.appendRow([range.setValue(e.value + ' added to Scan Log').offset(0, -1).getValue(), e.value, '\'' + String(e.value)])
            else // There are existing items on the scan log
            {
              const scanLog = scanLogPage.getSheetValues(1, 1, lastRow, 1);
              const itemDescription = range.offset(0, -1).getValue()

              for (var j = lastRow - 1; j >= 0; j--) // Loop through the scan log
              {
                if (scanLog[j][0] === itemDescription) // The description matches
                {
                  const scanLogRange = scanLogPage.getRange(j + 1, 2, 1, 2);
                  const itemValues = scanLogRange.getValues()
                  const updatedCount = Number(itemValues[0][0]) + e.value;
                  const runningSum = (isNotBlank(itemValues[0][1])) ? ((Math.sign(e.value) === 1 || Math.sign(e.value) === 0)  ? 
                                                                        String(itemValues[0][1]) + ' \+ ' + String(   e.value)  : 
                                                                        String(itemValues[0][1]) + ' \- ' + String(-1*e.value)) :
                                                                          ((isNotBlank(itemValues[0][0])) ? 
                                                                            String(itemValues[0][0]) + ' \+ ' + String(e.value) : 
                                                                            String(e.value));
                  scanLogRange.setNumberFormats([['#.#', '@']]).setValues([[updatedCount, runningSum]])
                  range.setValue(e.value + ' added to Scan Log')
                  break; // Item was found on the scan log, therefore stop searching
                }
              }

              if (j === scanLog.length) // Item was not found on the scan log
                scanLogPage.appendRow([range.setValue(e.value + ' added to Scan Log').offset(0, -1).getValue(), e.value, '\'' + String(e.value)])      
            }
          }
          else 
            Browser.msgBox('Invalid Entry', 'Please enter a valid quantity.', Browser.Buttons.OK)
        }
        else
        {
          const value = e.value.split(' ', 2);
          range.setValue(e.oldValue);

          if (isNumber(value[1])) // Assume the entry after the space is a valid upc code
          {
            if (value[0].toLowerCase() === 'mmm')
            {
              const description = sheet.getSheetValues(row, 1, 1, 4)[0][0]
              const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
              const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
              const numUpcs = upcDatabaseSheet.getLastRow();
              var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);

              var l = 0; // Lower-bound
              var u = numUpcs - 1; // Upper-bound
              var m = Math.ceil((u + l)/2) // Midpoint

              while (l < m && u > m)
              {
                if (value[1] < upcDatabase[m][0])
                  u = m;   
                else if (value[1] > upcDatabase[m][0])
                  l = m;

                m = Math.ceil((u + l)/2) // Midpoint
              }

              if (value[1] < upcDatabase[0][0])
                upcDatabase.splice(0, 0, [value[1], description])
              else if (value[1] < upcDatabase[m][0])
                upcDatabase.splice(m, 0, [value[1], description])
              else
                upcDatabase.splice(m + 1, 0, [value[1], description])
                
              upcDatabaseSheet.getRange(1, 1, numUpcs + 1, 2).setValues(upcDatabase);
              manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([[description.split(' - ', 1)[0], value[1], description]]);
              spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
            }
            else if (value[0].toLowerCase() === 'uuu')
            {
              const item = sheet.getSheetValues(row, 1, 1, 1)[0][0];
              const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
              unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[value[1], item]]);
              spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
            }
            else if (value[0].toLowerCase() === 'aaa')
            {
              const description = sheet.getSheetValues(row, 1, 1, 1)[0][0]
              const newItem = description.split(' - ')
              newItem[0] = 'MAKE_NEW_SKU'
              const itemJoined = newItem.join(' - ')
              const inventorySheet = spreadsheet.getSheetByName('Inventory');
              const upcDatabaseSheet = spreadsheet.getSheetByName("UPC Database");
              const manAddedUPCsSheet = spreadsheet.getSheetByName("Manually Added UPCs");
              const numUpcs = upcDatabaseSheet.getLastRow()
              var upcDatabase = upcDatabaseSheet.getSheetValues(1, 1, numUpcs, 2);

              var l = 0; // Lower-bound
              var u = numUpcs - 1; // Upper-bound
              var m = Math.ceil((u + l)/2) // Midpoint

              while (l < m && u > m)
              {
                if (value[1] < upcDatabase[m][0])
                  u = m;   
                else if (value[1] > upcDatabase[m][0])
                  l = m;

                m = Math.ceil((u + l)/2) // Midpoint
              }

              if (value[1] < upcDatabase[0][0])
                upcDatabase.splice(0, 0, [value[1], itemJoined])
              else if (value[1] < upcDatabase[m][0])
                upcDatabase.splice(m, 0, [value[1], itemJoined])
              else
                upcDatabase.splice(m + 1, 0, [value[1], itemJoined])

              upcDatabaseSheet.getRange(1, 1, numUpcs + 1, 2).setValues(upcDatabase);
              manAddedUPCsSheet.getRange(manAddedUPCsSheet.getLastRow() + 1, 1, 1, 3).setNumberFormat('@').setValues([['MAKE_NEW_SKU', value[1], itemJoined]]);
              inventorySheet.getRange(inventorySheet.getLastRow() + 1, 1).setValue(itemJoined); // Add the 'MAKE_NEW_SKU' item to the inventory sheet
              spreadsheet.getSheetByName("Manual Scan").getRange(1, 1).activate();
            }
            else
              Browser.msgBox('Invalid Entry', 'Please begin the command with either mmm , uuu, or aaa, orthwise input a valid number to update the Scan Log.', Browser.Buttons.OK)
          }
          else
            Browser.msgBox('Invalid UPC Code', 'Please type either mmm, uuu, or aaa, followed by SPACE and the UPC Code.', Browser.Buttons.OK)
        }
      }
      else // User has hit the delete key on one of the counts
      {
        const scanLog = spreadsheet.getSheetByName('Scan Log');
        const lastRow = scanLog.getLastRow();
        const scannedValues = scanLog.getSheetValues(1, 1, lastRow, 1)
        const description = range.offset(0, -1).getValue()

        for (var i = scannedValues.length - 1; i >= 0; i--)
        {
          if (scannedValues[i][0] === description)
          {
            scanLog.deleteRow(i + 1)
            break;
          }
        }

        range.setValue('Item removed from Scan Log')
      }
    }
  }
}

/**
 * This function takes the item that was just scanned on the manual scan page and copies it to the list of UPCs to unmarry from the countersales data.
 * A user interface is launched that accepts the UPC value to unmarry
 * 
 * @author Jarren Ralf
 */
function unmarryUPC()
{
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActive()
  const item = spreadsheet.getActiveSheet().getSheetValues(1, 1, 1, 1)[0][0].toString().split('\n')
  const response = ui.prompt('Unmarry UPCs', 'Please scan the barcode for:\n\n' + item[0] +'.', ui.ButtonSet.OK_CANCEL)

  if (ui.Button.OK === response.getSelectedButton())
  {
    const unmarryUpcSheet = spreadsheet.getSheetByName("UPCs to Unmarry");
    unmarryUpcSheet.getRange(unmarryUpcSheet.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[response.getResponseText(), item[0]]]);
  }
}

/**
 * This function watches two cells and if the left one is edited then it searches the UPC Database for the upc value (the barcode that was scanned).
 * It then checks if the item is on the scan log and stores the relevant data in the left cell. If the rightr cell is edited, then the function
 * uses the data in the left cell and moves the item over to the scan log with the updated quantity.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf
 */
function updateCounts(e, sheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;

  if (col == 2)
  {
    if (e.oldValue !== undefined) // Old value is NOT blank
    {
      if (userHasNotPressedDelete(e.value)) // New value is NOT blank
      {
        if (isNumber(e.value))
        {
          if (isNumber(e.oldValue))
          {
            const difference  = e.value - e.oldValue;
            const runningSumRange = sheet.getRange(row, 3);
            var runningSumValue = runningSumRange.getValue().toString();

            if (runningSumValue === '')
              runningSumValue = Math.round(e.oldValue).toString();

            (difference > 0) ? 
              runningSumRange.setValue(runningSumValue.toString() + ' + ' + difference.toString()) : 
              runningSumRange.setValue(runningSumValue.toString() + ' - ' + (-1*difference).toString());
          }
          else // Old value is not a number
          {
            const runningSumRange = sheet.getRange(row, 3);
            var runningSumValue = runningSumRange.getValue().toString();

            if (isNotBlank(runningSumValue))
              runningSumRange.setValue(runningSumValue + ' + ' + Math.round(e.value).toString());
            else
              runningSumRange.setValue(Math.round(e.value).toString());
          }
        }
        else if (e.value.toString().split(' ', 1)[0] === 'a') // The number is preceded by the letter 'a' and a space, in order to trigger an "add" operation
        {
          const newCountAndSumRange = sheet.getRange(row, 2, 1, 2);
          var newCountAndSumValue = newCountAndSumRange.getValues()
          newCountAndSumValue[0][0] = e.value.toString().split(' ')[1]

          if (isNumber(newCountAndSumValue[0][0])) // New Count is a number
          {
            if (isNumber(e.oldValue))
            {
              if (isNotBlank(newCountAndSumValue[0][1]))
                newCountAndSumRange.setNumberFormat('@').setValues([[parseInt(e.oldValue) + parseInt(newCountAndSumValue[0][0]), 
                  newCountAndSumValue[0][1].toString() + ' + ' + newCountAndSumValue[0][0].toString()]])

              else
                newCountAndSumRange.setNumberFormat('@').setValues([[parseInt(e.oldValue) + parseInt(newCountAndSumValue[0][0]), 
                  parseInt(e.oldValue).toString() + ' + ' + newCountAndSumValue[0][0].toString()]])
            }
            else
            {
              if (isNotBlank(newCountAndSumValue[0][1]))
                newCountAndSumRange.setNumberFormat('@').setValues([[newCountAndSumValue[0][0], 
                  newCountAndSumValue[0][1].toString() + ' + ' + NaN.toString() + ' + ' + newCountAndSumValue[0][0].toString()]])
                
              else
                newCountAndSumRange.setNumberFormat('@').setValues([[newCountAndSumValue[0][0], newCountAndSumValue[0][0].toString()]])
            }
          }
          else
          {
            if (isNumber(e.oldValue))
            {
              if (isNotBlank(newCountAndSumValue[0][1])) // Running Sum is not blank
                newCountAndSumRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), newCountAndSumValue[0][1].toString() + ' + ' + NaN.toString()]])
              else
                newCountAndSumRange.setNumberFormat('@').setValues([[Math.round(e.oldValue).toString(), Math.round(e.oldValue).toString() + ' + ' + NaN.toString()]])
            }

            SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
          }
        }
        else // New value is not a number
        {
          const runningSumRange = sheet.getRange(row, 3);
          const runningSumValue = runningSumRange.getValue().toString();

          if (isNumber(e.oldValue))
          {
            if (isNotBlank(runningSumValue))
              runningSumRange.setNumberFormat('@').setValue(runningSumValue + ' + ' + NaN.toString())
            else
              runningSumRange.setNumberFormat('@').setValue(Math.round(e.oldValue).toString())
          }

          SpreadsheetApp.getUi().alert("The quantity you entered is not a number.")
        }
      }
      else // New value IS blank
        sheet.getRange(row, 3).setValue(''); // Clear the running sum
    }
    else
      if (!isNumber(e.value)) SpreadsheetApp.getUi().alert("The quantity you entered is not a number.");
      //(isNumber(e.value)) ? sheet.getRange(row, 2).setValue(e.value) : SpreadsheetApp.getUi().alert("The quantity you entered is not a number.");
  }
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning false if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is not-undefined or not.
* @author Jarren Ralf
*/
function userHasNotPressedDelete(value)
{
  return value !== undefined;
}