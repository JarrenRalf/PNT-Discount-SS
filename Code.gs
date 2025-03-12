/**
 * This function runs the item search sheet.
 * 
 * @param {Event Object} : The event object
 * @author Jarren Ralf
 */
function installedOnEdit(e)
{
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  var sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === "Item Search") // Check if the user is searching for an item or trying to marry, unmarry or add a new item to the upc database
      searchV2(e, spreadsheet, sheet);
    else if (sheetName === "Discount Percentages") // Check if the user is changing the discount values
      changeDiscountStructure(e, spreadsheet, sheet);
  } 
  catch (err) 
  {
    var error = err['stack'];
    Logger.log(error)
    Browser.msgBox(error)
    throw new Error(error);
  }
}

/**
 * This function loads a Menu where the user can click a button to update the Adagio price. 
 * It also puts a filter on the Discount Percentages page as to only display the items that do not have discounts assigned yet.
 * 
 * @param {Event Object} : The event object
 * @author Jarren Ralf
 */
function onOpen(e)
{
  const spreadsheet = e.source;

  SpreadsheetApp.getUi().createMenu('PNT Controls')
    .addItem('Update Price', 'updateAdagioBasePrice')
    .addItem('Update Items (from inventory.csv)', 'updateAdagioDatabase')
    .addToUi();

  spreadsheet.getSheetByName('Item Search').getRange(1, 3).uncheck();
  spreadsheet.getSheetByName('Shopify Update').hideSheet();

  const filter = spreadsheet.getSheetByName('Discount Percentages').getFilter();
  filter.setColumnFilterCriteria( 5, SpreadsheetApp.newFilterCriteria().whenCellEmpty().build())
        .setColumnFilterCriteria(15, SpreadsheetApp.newFilterCriteria().whenTextEqualTo("0").build());
}

/**
 * Display the items that have a discount structure such that Lodge > Wholesale or Guide > Lodge or Guide > Wholesale.
 * 
 * @author Jarren Ralf
 */
function itemsWithNonLinearDiscounts()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const discountPercentagesSheet = spreadsheet.getSheetByName('Discount Percentages');
  const discountPercentages = discountPercentagesSheet.getSheetValues(2, 11, discountPercentagesSheet.getLastRow() - 1, 5);
  const lodge_GreaterThan_Wholesale = [], guide_GreaterThan_Lodge = [], guide_GreaterThan_Wholesale = [];

  discountPercentages.map(item => {
    if (item[3] > item[4])
      lodge_GreaterThan_Wholesale.push(item)
    else if (item[2] > item[3])
      guide_GreaterThan_Lodge.push(item)
    else if (item[2] > item[4])
      guide_GreaterThan_Wholesale.push(item)
  })

  spreadsheet.getSheetByName('Lodge > Wholesale').clearContents()
    .getRange(1, 1, lodge_GreaterThan_Wholesale.unshift(['Google Description',	'Base Price',	'Guide', 'Lodge',	'Wholesale']), 5).setValues(lodge_GreaterThan_Wholesale)
  spreadsheet.getSheetByName('Guide > Lodge')    .clearContents()
    .getRange(1, 1, guide_GreaterThan_Lodge.unshift(    ['Google Description',	'Base Price',	'Guide', 'Lodge',	'Wholesale']), 5).setValues(guide_GreaterThan_Lodge)
  spreadsheet.getSheetByName('Guide > Wholesale').clearContents()
    .getRange(1, 1, guide_GreaterThan_Wholesale.unshift(['Google Description',	'Base Price',	'Guide', 'Lodge',	'Wholesale']), 5).setValues(guide_GreaterThan_Wholesale)
}

/**
 * This function takes the values that the user has just changed on the Discount Percentages page, specifically changes to the 3 discount structures that we use,
 * and it logs those changes on the Shopify Update sheet.
 * 
 * @param {Range}               range : The active range that was just editted by the user.
 * @param {Number}                row : The first row that was editted by the user.
 * @param {Number}                col : The first column that was editted by the user.
 * @param {Number}            numRows : The number of rows that were editted by the user.
 * @param {Number}            numCols : The number of columns that were editted by the user.
 * @param {Object[][]}         values : The new values that the user has just changed on the sheet.
 * @param {Spreadsheet}   spreadsheet : The active spreadsheet
 * @param {Sheet}               sheet : The sheet that is being edited
 * @param {Boolean}      isSingleCell : Whether a single cell was editted or not.
 * @param {Boolean} isItemSearchSheet : Whether the active sheet is the item search page or not.
 * @author Jarren Ralf
 */
function addItemToShopifyUpdatePage(range, row, col, numRows, numCols, values, spreadsheet, sheet, isSingleCell, isItemSearchSheet)
{
  if (isSingleCell)
    spreadsheet.toast('Adding your change to the shopify update page...', 'Shopify Updating', -1)
  else
    spreadsheet.toast('This may take up to 30 seconds. Adding your change to the shopify update page...', 'Shopify Updating', -1)

  const shopifyUpdateSheet = spreadsheet.getSheetByName('Shopify Update');
  const lastRow = shopifyUpdateSheet.getLastRow();

  if (isItemSearchSheet)
  {
    const removePriceColumns = u => [u[0].split(' - ').pop().toString().toUpperCase(), '%', (Number(u[2]) < 1) ? u[2]*100 : u[2], (Number(u[4]) < 1) ? u[4]*100 : u[4], (Number(u[6]) < 1) ? u[6]*100 : u[6]];

    if (lastRow > 1)
    {
      const recentlyUpdatedItems = shopifyUpdateSheet.getSheetValues(2, 1, lastRow - 1, 5)
      var idx;

      if (isSingleCell)
      {
        const newItem = range.offset(0, 2 - col, numRows, 7).getValues()[0]
        idx = recentlyUpdatedItems.findIndex(item => item[0] == newItem[0])

        if (idx !== -1)
          recentlyUpdatedItems[idx][Math.floor(col/2)] = values; // This was the single change that was made by the user
        else
        {
          newItem[(col % 2 == 0) ? col - 2 : col - 3] = values;
          recentlyUpdatedItems.push(...[newItem].map(removePriceColumns));
        }
      }
      else
      {
        const colIdx_ShopifyUpdate = Math.floor(col/2);
        const colIdx_ItemSearch = (col % 2 == 0) ? col - 2 : col - 3;

        range.offset(0, 2 - col, numRows, 8).getValues().map((newItem, r) => {

          idx = recentlyUpdatedItems.findIndex(item => item[0] == newItem[0])

          if (idx !== -1)
            recentlyUpdatedItems[idx][colIdx_ShopifyUpdate] = (newItem[colIdx_ItemSearch] < 1) ? newItem[colIdx_ItemSearch]*100 : newItem[colIdx_ItemSearch];
          else
            recentlyUpdatedItems.push(...[newItem].map(removePriceColumns))
        })
      }

      shopifyUpdateSheet.getRange(2, 1, recentlyUpdatedItems.length, 5).setValues(recentlyUpdatedItems);
    }
    else // There are no other items currently on the list therefore add the recent change straight to the list
      shopifyUpdateSheet.getRange(2, 1, numRows, 5).setValues(range.offset(0, 2 - col, numRows, 7).getValues().map(removePriceColumns))
  }
  else // Discount Percentages sheet
  {
    const reformat = u => [u[0].split(' - ').pop().toString().toUpperCase(), '%', u[2], u[3], u[4]];

    if (lastRow > 1)
    {
      const recentlyUpdatedItems = shopifyUpdateSheet.getSheetValues(2, 1, lastRow - 1, 5)
      var idx;

      if (isSingleCell)
      {
        const newItem = range.offset(0, 11 - col, numRows, 5).getValues()[0]
        idx = recentlyUpdatedItems.findIndex(item => item[0] == newItem[0])

        if (idx !== -1)
          recentlyUpdatedItems[idx][col - 11] = values[0][0]; // This was the single change that was made by the user
        else
          recentlyUpdatedItems.push(newItem.map(reformat))
      }
      else if (numCols == 3 && col == 13)
      {
        range.offset(0, 11 - col, numRows, 5).getValues().map((newItem, r) => {

          if (!sheet.isRowHiddenByFilter(row + r))
          {
            idx = recentlyUpdatedItems.findIndex(item => item[0] == newItem[0])

            if (idx !== -1)
            {
              recentlyUpdatedItems[idx][2] = newItem[2];
              recentlyUpdatedItems[idx][3] = newItem[3];
              recentlyUpdatedItems[idx][4] = newItem[4];
            }
            else
              recentlyUpdatedItems.push(newItem.map(reformat))
          }
        })
      }
      else if (numCols == 2)
      {
        const colIdx_1 = col - 11;
        const colIdx_2 = col - 10;

        range.offset(0, 11 - col, numRows, 5).getValues().map((newItem, r) => {

          if (!sheet.isRowHiddenByFilter(row + r))
          {
            idx = recentlyUpdatedItems.findIndex(item => item[0] == newItem[0])

            if (idx !== -1)
            {
              recentlyUpdatedItems[idx][colIdx_1] = newItem[colIdx_1];
              recentlyUpdatedItems[idx][colIdx_2] = newItem[colIdx_2];
            }
            else
              recentlyUpdatedItems.push(newItem.map(reformat))
          }
        })
      }
      else if (numCols == 1)
      {
        const colIdx = col - 11;

        range.offset(0, 11 - col, numRows, 5).getValues().map((newItem, r) => {

          if (!sheet.isRowHiddenByFilter(row + r))
          {
            idx = recentlyUpdatedItems.findIndex(item => item[0] == newItem[0])

            if (idx !== -1)
              recentlyUpdatedItems[idx][colIdx] = newItem[colIdx];
            else
              recentlyUpdatedItems.push(newItem.map(reformat))
          }
        })
      }

      shopifyUpdateSheet.getRange(2, 1, recentlyUpdatedItems.length, 5).setValues(recentlyUpdatedItems);
    }
    else // There are no other items currently on the list therefore add the recent change straight to the list
      shopifyUpdateSheet.getRange(2, 1, numRows, 5).setValues(range.offset(0, 11 - col, numRows, 5).getValues().map(reformat))
  }

  spreadsheet.toast('Update COMPLETED', 'Shopify Updated')
}

/**
 * This function checks if the user has made changes to any of the 3 relevant discscount markup columns.
 * 
 * @param {Event Object}     e      : The event object generated by an on edit event.
 * @param {Spreadsheet} spreadsheet : The active spreadsheet
 * @param    {Sheet}        sheet   : The sheet that is being edited
 * @author Jarren Ralf
 */
function changeDiscountStructure(e, spreadsheet, sheet)
{
  const range = e.range;
  const col = range.columnStart;

  if (range.rowStart > 1 && col > 12 && col < 16) // Discount Markup 1, 2, or 3 are being edited
  {
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();
    const values = range.getValues()

    if (numRows > 1) // Multiple rows were changed
    {
      if (col + numCols > 16) // Too many columns were changed
        if (isEveryValueBlank(values)) // Every value is blank, therefore this is the user clicking delete
          spreadsheet.toast('Press Ctrl + Z to undo your changes', 'Too many columns changed', -1);
        else // Assumed to be an undo 
          spreadsheet.toast('Undo: Successful');
      else if (isEveryValueBlank(values)) // Every value is blank, therefore this is the user clicking delete
      {
        range.offset(0, 13 - col, numRows, 3).setValue('').offset(0, -8, numRows, 1).setValue(new Date().toDateString());
        addItemToShopifyUpdatePage(range, range.rowStart, col, numRows, numCols, values, spreadsheet, sheet, false)
      }
      else if (values.some(vals => vals.some(num => isNaN(Number(num))))) // Atleast one of the entries contains a letter
      {
        range.offset(0, 13 - col, numRows, 3).setValue(0).offset(0, -8, numRows, 1).setValue('');
        spreadsheet.toast('Numerals only')
      }
      else // Assumed to be a user make an edit to a multiple rows and 1 or more columns
      {
        range.offset(0, 5 - col, numRows, 1).setValue(new Date().toDateString());
        addItemToShopifyUpdatePage(range, range.rowStart, col, numRows, numCols, values, spreadsheet, sheet, false)
      }
    }
    else if (numCols > 1) // Multiple columns were changed in 1 row
    {
      if (col + numCols > 16) // Too many columns were changed
        if (isEveryValueBlank(values)) // Every value is blank, therefore this is the user clicking delete
          spreadsheet.toast('Press Ctrl + Z to undo your changes', 'Too many columns changed', -1);
        else // Assumed to be an undo 
          spreadsheet.toast('Undo: Successful');
      else if (isEveryValueBlank(values)) // Every value is blank, therefore this is the user clicking delete
      {
        range.offset(0, 13 - col, 1, 3).setValues([['', '', '']]).offset(0, -8, 1, 1).setValue(new Date().toDateString());
        SpreadsheetApp.flush();
        addItemToShopifyUpdatePage(range, range.rowStart, col, numRows, numCols, values, spreadsheet, sheet, false)
      }
        
      else if (values[0].some(num => isNaN(Number(num)))) // Atleast one of the entries contains a letter
      {
        range.offset(0, 13 - col, 1, 3).setValues([[0, 0, 0]]).offset(0, -8, 1, 1).setValue('');
        spreadsheet.toast('Numerals only')
      }
      else // Assumed to be a user make an edit to a single row but multiple columns
      {
        range.offset(0, 5 - col, 1, 1).setValue(new Date().toDateString());
        addItemToShopifyUpdatePage(range, range.rowStart, col, numRows, numCols, values, spreadsheet, sheet, false)
      }
    }
    else // One cell is being changed given only 1 range in the range list
    {
      const oldValue = e.oldValue;

      if (isNotBlank(values[0][0])) // The user is assumed to be typing a number
      {
        if (isNaN(Number(values[0][0]))) // User inputted a non-number
        {
          range.setValue(oldValue).offset(0, 13 - col, 1, 3).setValues([[0, 0, 0]]).offset(0, -8, 1, 1).setValue('');
          spreadsheet.toast('Numerals only')
        }
        else
        {
          range.offset(0, 5 - col, 1, 1).setValue(new Date().toDateString());
          addItemToShopifyUpdatePage(range, range.rowStart, col, numRows, numCols, values, spreadsheet, sheet, true)
        }
      }
      else if (oldValue == 0) // It appears that the user has pressed delete over a value of zero, meaning that 0 is the chosen discount
      {
        range.offset(0, 13 - col, 1, 3).setValues([['', '', '']]).offset(0, -8, 1, 1).setValue(new Date().toDateString());
        spreadsheet.toast('0% Discount: Confirmed');
        SpreadsheetApp.flush()
        addItemToShopifyUpdatePage(range, range.rowStart, col, numRows, numCols, values, spreadsheet, sheet, true)
      }
    }
  }
}

/**
 * This function checks if every value in the import multi-array is blank, which means that the user has
 * highlighted and deleted all of the data.
 * 
 * @param {Object[][]} values : The import data
 * @return {Boolean} Whether the import data is deleted or not
 * @author Jarren Ralf
 */
function isEveryValueBlank(values)
{
  return values.every(arr => arr.every(val => val === '') === true);
}

/**
 * This function checks if the given string is not blank.
 * 
 * @param {String} str : The given string.
 * @return {Boolean} Returns true if the given string is blank, false otherwise.
 */
function isNotBlank(str)
{
  return str !== ''
}

/**
 * This function removes the dashes from the SKUs that come from Adagio.
 * 
 * @author Jarren Ralf
 */
function removeDashes()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName('Discount Percentages');
  const numRows = sheet.getLastRow() - 1;
  const range = sheet.getRange(2, 1, numRows, 1)
  const values = range.getValues().map(v => [v[0].toString().substring(0, 4) + v[0].toString().substring(5, 9) + v[0].toString().substring(10)])
  range.setValues(values)
}

/**
 * This function first applies the standard formatting to the search box, then it seaches the SearchData page for the items in question.
 * It also highlights the items that are already on the shipped page and already on the order page.
 * 
 * @param {Event Object}      e      : An instance of an event object that occurs when the spreadsheet is editted
 * @param {Spreadsheet}  spreadsheet : The spreadsheet that is being edited
 * @param    {Sheet}        sheet    : The sheet that is being edited
 * @author Jarren Ralf 
 */
function searchV2(e, spreadsheet, sheet)
{
  const range = e.range;
  const row = range.rowStart;
  const col = range.columnStart;
  const rowEnd = range.rowEnd;
  const colEnd = range.columnEnd;

  if (col == colEnd)
  {
    if (col === 2) // Column two is being edited
    {
      const startTime = new Date().getTime();
      const searchResultsDisplayRange = sheet.getRange(1, 1); // The range that will display the number of items found by the search
      const functionRunTimeRange = sheet.getRange(2, 1);      // The range that will display the runtimes for the search and formatting
      const itemSearchFullRange = sheet.getRange(4, 1, sheet.getMaxRows() - 3, 9); // The entire range of the Item Search page

      if (row == rowEnd && row === 1) // The search box is being edited
      {
        const output = [];
        const searchesOrNot = sheet.getRange(1, 2).clearFormat()                                          // Clear the formatting of the range of the search box
          .setBorder(true, true, true, true, null, null, 'black', SpreadsheetApp.BorderStyle.SOLID_THICK) // Set the border
          .setFontFamily("Arial").setFontColor("black").setFontWeight("bold").setFontSize(14)             // Set the various font parameters
          .setHorizontalAlignment("center").setVerticalAlignment("middle")                                // Set the alignment
          .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)                                              // Set the wrap strategy
          .merge().trimWhitespace()                                                                       // Merge and trim the whitespaces at the end of the string
          .getValue().toString().toLowerCase().split(' not ')                                             // Split the search string at the word 'not'

        const searches = searchesOrNot[0].split(' or ').map(words => words.split(/\s+/)) // Split the search values up by the word 'or' and split the results of that split by whitespace

        if (isNotBlank(searches[0][0])) // If the value in the search box is NOT blank, then compute the search
        {
          spreadsheet.toast('Searching...')
          const inventorySheet = spreadsheet.getSheetByName('Discount Percentages');
          const data = inventorySheet.getSheetValues(2, 10, inventorySheet.getLastRow() - 1, 6)
            .map(d => [d[0], d[1], d[2], Number(d[3])/100, (d[2]*(100 - d[3])/100).toFixed(2), Number(d[4])/100, (d[2]*(100 - d[4])/100).toFixed(2), Number(d[5])/100, (d[2]*(100 - d[5])/100).toFixed(2)]);
          const numSearches = searches.length; // The number searches
          var numSearchWords;

          if (searchesOrNot.length === 1) // The word 'not' WASN'T found in the string
          {
            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      output.push(data[i]);
                      break loop;
                    }
                  }
                  else
                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item
                }
              }
            }
          }
          else // The word 'not' was found in the search string
          {
            var dontIncludeTheseWords = searchesOrNot[1].split(/\s+/);

            for (var i = 0; i < data.length; i++) // Loop through all of the descriptions from the search data
            {
              loop: for (var j = 0; j < numSearches; j++) // Loop through the number of searches
              {
                numSearchWords = searches[j].length - 1;

                for (var k = 0; k <= numSearchWords; k++) // Loop through each word in each set of searches
                {
                  if (data[i][1].toString().toLowerCase().includes(searches[j][k])) // Does the i-th item description contain the k-th search word in the j-th search
                  {
                    if (k === numSearchWords) // The last search word was succesfully found in the ith item, and thus, this item is returned in the search
                    {
                      for (var l = 0; l < dontIncludeTheseWords.length; l++)
                      {
                        if (!data[i][1].toString().toLowerCase().includes(dontIncludeTheseWords[l]))
                        {
                          if (l === dontIncludeTheseWords.length - 1)
                          {
                            output.push(data[i]);
                            break loop;
                          }
                        }
                        else
                          break;
                      }
                    }
                  }
                  else
                    break; // One of the words in the User's query was NOT contained in the ith item description, therefore move on to the next item 
                }
              }
            }
          }

          const numItems = output.length;

          if (numItems === 0) // No items were found
          {
            sheet.getRange('B1').activate(); // Move the user back to the seachbox
            itemSearchFullRange.clearContent(); // Clear content
            const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
            const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
            searchResultsDisplayRange.setRichTextValue(message);
          }
          else
          {
            const numberFormats_ItemSearch = [...Array(numItems)].map(e => ['@', '@', '$0.00', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00']); // Currency format
            const fontWeights = [...Array(numItems)].map(e => ['bold', 'bold', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold']); // Currency format
            sheet.getRange('B4').activate(); // Move the user to the top of the search items
            itemSearchFullRange.clearContent().offset(0, 0, numItems, 9).setNumberFormats(numberFormats_ItemSearch).setFontWeights(fontWeights).setValues(output);
            (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
          }

          spreadsheet.toast('Searching Complete.')
        }
        else if (isNotBlank(e.oldValue) && userHasPressedDelete(e.value)) // If the user deletes the data in the search box, then ...
        {
          itemSearchFullRange.setBackground('white').setValue('');
          searchResultsDisplayRange.setValue('');
        }
        else
        {
          itemSearchFullRange.clearContent(); // Clear content 
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
          const message = SpreadsheetApp.newRichTextValue().setText("Invalid search.\nPlease try again.").setTextStyle(0, 14, textStyle).build();
          searchResultsDisplayRange.setRichTextValue(message);
        }

        functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
      }
      else if (row != rowEnd && row > 3) // Multiple lines are being pasted
      {
        const values = range.getValues().filter(blank => isNotBlank(blank[0]))

        if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
        {
          const inventorySheet = spreadsheet.getSheetByName('Discount Percentages');
          const data = inventorySheet.getSheetValues(2, 10, inventorySheet.getLastRow() - 1, 6)
            .map(d => [d[0], d[1], d[2], d[3]/100, (d[2]*(100 - d[3])/100).toFixed(2), d[4]/100, (d[2]*(100 - d[4])/100).toFixed(2), d[5]/100, (d[2]*(100 - d[5])/100).toFixed(2)]);
          var someSKUsNotFound = false, skus;

          if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
          {
            skus = values.map(item => {
            
              for (var i = 0; i < data.length; i++)
              {
                if (data[i][1].toString().split(" - ").pop().toUpperCase() == item[0].toString().split(" - ").pop().toUpperCase())
                  return data[i];
              }

              someSKUsNotFound = true;

              return ['SKU Not Found:', item[0].toString().split(" - ").pop().toUpperCase(), '', '', '', '', '']
            });
          }
          else if (values[0][0].toString().includes('-'))
          {
            skus = values.map(sku => (sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).trim()).map(item => {
            
              for (var i = 0; i < data.length; i++)
              {
                if (data[i][1].toString().split(" - ").pop().toUpperCase() == item.toString().toUpperCase())
                  return data[i];
              }

              someSKUsNotFound = true;

              return ['SKU Not Found:', item, '', '', '', '', '', '', '']
            });
          }
          else
          {
            skus = values.map(item => {
            
              for (var i = 0; i < data.length; i++)
              {
                if (data[i][1].toString().split(" - ").pop().toUpperCase() == item[0].toString().toUpperCase())
                  return data[i];
              }

              someSKUsNotFound = true;

              return ['SKU Not Found:', item[0], '', '', '', '', '', '', '']
            });
          }
      
          if (someSKUsNotFound)
          {
            const skusNotFound = [];
            var isSkuFound;

            const skusFound = skus.filter(item => {
              isSkuFound = item[0] !== 'SKU Not Found:'

              if (!isSkuFound)
                skusNotFound.push(item)

              return isSkuFound;
            })

            const numSkusFound = skusFound.length;
            const numSkusNotFound = skusNotFound.length;
            const items = [].concat.apply([], [skusNotFound, skusFound]); // Concatenate all of the item values as a 2-D array
            var numItems = items.length
            const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'right', 'right', 'right', 'right', 'right', 'right', 'right'])
            const numberFormats_ItemSearch = [...Array(numItems)].map(e => ['@', '@', '$0.00', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00']); // Currency format
            const fontWeights = [...Array(numItems)].map(e => ['bold', 'bold', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold']); // Currency format
            const WHITE = new Array(9).fill('white')
            const YELLOW = new Array(9).fill('#ffe599')
            const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

            itemSearchFullRange.clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
              .offset(0, 0, numItems, 9)
                .setFontFamily('Arial').setFontWeight('bold').setFontSize(12).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours)
                .setBorder(false, null, false, null, false, false).setNumberFormats(numberFormats_ItemSearch).setFontWeights(fontWeights).setValues(items).activate();

            if (numSkusFound > 0)
              itemSearchFullRange.offset(numSkusNotFound, 0, numSkusFound, 9).activate()
          }
          else // All SKUs were succefully found
          {
            var numItems = skus.length
            const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'right', 'right', 'right', 'right', 'right', 'right', 'right'])
            const numberFormats_ItemSearch = [...Array(numItems)].map(e => ['@', '@', '$0.00', '#%', '$0.00', '#%', '$0.00', '#%', '$0.00']); // Currency format
            const fontWeights = [...Array(numItems)].map(e => ['bold', 'bold', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold']); // Currency format

            itemSearchFullRange.clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
              .offset(0, 0, numItems, 9)
                .setFontFamily('Arial').setFontWeight('bold').setFontSize(12).setHorizontalAlignments(horizontalAlignments)
                .setBorder(false, null, false, null, false, false).setNumberFormats(numberFormats_ItemSearch).setFontWeights(fontWeights).setValues(skus).activate()
          }

          (numItems !== 1) ? searchResultsDisplayRange.setValue(numItems + " results found.") : searchResultsDisplayRange.setValue(numItems + " result found.");
          spreadsheet.toast('Searching Complete.')
          functionRunTimeRange.setValue((new Date().getTime() - startTime)/1000 + " seconds");
        }
        else
        {
          sheet.getRange('B1').activate(); // Move the user back to the seachbox
          itemSearchFullRange.clearContent(); // Clear content
          const textStyle = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('yellow').build();
          const message = SpreadsheetApp.newRichTextValue().setText("No results found.\nPlease try again.").setTextStyle(0, 16, textStyle).build();
          searchResultsDisplayRange.setRichTextValue(message);
        }
      }
    }
    else if (col > 3 && col <= 9 && row > 3) // The discount or price is being changed
    {
      if (range.offset(1 - row, 3 - col, 1, 1).isChecked()) // The top right hand cell of this spreadsheet has a checkbox that needs to be click before prices can by changed
      {
        spreadsheet.toast('', 'Updating discount structure...', -1)

        if (rowEnd > row) // This is a drag action with more than 1 item changing
        {
          const discountDataSheet = spreadsheet.getSheetByName('Discount Percentages');
          const discountData_SKUs = discountDataSheet.getSheetValues(2, 1, discountDataSheet.getLastRow() - 1, 1);
          
          const numRows = rowEnd - row + 1;
          var idx, itemRange, itemValues;

          switch (col)
          {
            case 4: // Percentages Changed
            case 6:
            case 8:

              const updatedPrices = [];

              range.offset(0, 2 - col, numRows, col - 1).getValues().map(itemDiscountValues => {
                idx = discountData_SKUs.findIndex(sku => sku[0] == itemDiscountValues[0].split(' - ').pop()) + 2;

                if (idx !== 1)
                { 
                  itemRange = discountDataSheet.getRange(idx, 5, 1, 11);
                  itemValues = itemRange.getValues()[0];
                  itemValues[0] = new Date().toDateString(); // Date (Category)
                  itemValues[col/2 + 6] = itemDiscountValues[col - 2]*100; // Change the discount column
                  itemRange.setNumberFormat('@').setValues([itemValues])
                  updatedPrices.push([Number(itemDiscountValues[1])*(100 - Number(itemValues[col/2 + 6]))/100]); // Price
                }
              })

              range.offset(0, 1, numRows).setNumberFormat('$0.00').setValues(updatedPrices)
              break;
            case 5: // Price Changed
            case 7:
            case 9:

              const updatedPercentagesAndPrices = [];

              range.offset(0, 2 - col, numRows, col - 1).getValues().map(itemDiscountValues => {
                idx = discountData_SKUs.findIndex(sku => sku[0] == itemDiscountValues[0].split(' - ').pop()) + 2;

                if (idx !== 1)
                { 
                  itemRange = discountDataSheet.getRange(idx, 5, 1, 11);
                  itemValues = itemRange.getValues()[0];
                  itemValues[0] = new Date().toDateString(); // Date (Category)
                  itemValues[col/2 + 5.5] = (1 - Number(itemDiscountValues[col - 2])/Number(itemDiscountValues[1]))*100; // Change the appropriate discount column
                  itemRange.setNumberFormat('@').setValues([itemValues])
                  updatedPercentagesAndPrices.push([Number(itemValues[col/2 + 5.5])/100, itemDiscountValues[col - 2]]); // Percentage and price
                }
              })

              range.offset(0, -1, numRows, 2).setNumberFormats(new Array(numRows).fill(['#%', '$0.00'])).setValues(updatedPercentagesAndPrices)
              break;
          }

          SpreadsheetApp.flush()
          addItemToShopifyUpdatePage(range, row, col, numRows, colEnd - col + 1, '', spreadsheet, sheet, false, true)
          spreadsheet.toast('', 'Discounts Updated.')
        }
        else if (!e.oldValue) // If e.oldValue is undefined then the user used the drag function to change 1 item
        {
          const sku = range.offset(0, 2 - col).getValue().split(' - ').pop().toUpperCase();
          const discountDataSheet = spreadsheet.getSheetByName('Discount Percentages');
          const itemIndex = discountDataSheet.getSheetValues(2, 11, discountDataSheet.getLastRow() - 1, 1).findIndex(description => description[0].split(' - ').pop().toUpperCase() === sku);

          if (itemIndex !== -1)
          {
            switch (col)
            {
              case 4: // Percentages Changed
              case 6:
              case 8:
                var newPercentage = range.getValue();
                discountDataSheet.getRange(itemIndex + 2, col/2 + 11).setNumberFormat('@').setValue(newPercentage*100);
                addItemToShopifyUpdatePage(range, row, col, 1, colEnd - col + 1, Number(newPercentage)*100, spreadsheet, sheet, true, true)
                range.offset(0, 0, 1, 2).setNumberFormats([['#%', '$0.00']]).setValues([[newPercentage, Number(range.offset(0, 3 - col).getValue())*(1 - newPercentage)]])
                break;
              case 5: // Price Changed
              case 7:
              case 9:
                const newPrice = range.getValue();
                var newPercentage = 1 - Number(newPrice)/Number(range.offset(0, 3 - col).getValue())
                discountDataSheet.getRange(itemIndex + 2, col/2 + 10.5).setNumberFormat('@').setValue(newPercentage*100);
                addItemToShopifyUpdatePage(range, row, col, 1, colEnd - col + 1, Number(newPercentage)*100, spreadsheet, sheet, true, true)
                range.offset(0, -1, 1, 2).setNumberFormats([['#%', '$0.00']]).setValues([[newPercentage, newPrice]])
                break;
            }

            spreadsheet.toast('', 'Discount Updated.')
          }
          else
          {
            range.setValue(e.oldValue)
            SpreadsheetApp.flush()
            Browser.msgBox('Item not found on the Discount Percentages sheet.')
          }
        }
        else // Single row
        {
          const sku = range.offset(0, 2 - col).getValue().split(' - ').pop().toUpperCase();
          const discountDataSheet = spreadsheet.getSheetByName('Discount Percentages');
          const itemIndex = discountDataSheet.getSheetValues(2, 11, discountDataSheet.getLastRow() - 1, 1).findIndex(description => description[0].split(' - ').pop().toUpperCase() === sku);

          if (itemIndex !== -1)
          {
            switch (col)
            {
              case 4: // Percentages Changed
              case 6:
              case 8:
                var newPercentage = range.getValue();
                discountDataSheet.getRange(itemIndex + 2, col/2 + 11).setNumberFormat('@').setValue(newPercentage*100);
                addItemToShopifyUpdatePage(range, row, col, 1, colEnd - col + 1, Number(newPercentage)*100, spreadsheet, sheet, true, true)
                range.offset(0, 0, 1, 2).setNumberFormats([['#%', '$0.00']]).setValues([[newPercentage, Number(range.offset(0, 3 - col).getValue())*(1 - newPercentage)]])
                break;
              case 5: // Price Changed
              case 7:
              case 9:
                const newPrice = range.getValue();
                var newPercentage = 1 - Number(newPrice)/Number(range.offset(0, 3 - col).getValue())
                discountDataSheet.getRange(itemIndex + 2, col/2 + 10.5).setNumberFormat('@').setValue(newPercentage*100);
                addItemToShopifyUpdatePage(range, row, col, 1, colEnd - col + 1, Number(newPercentage)*100, spreadsheet, sheet, true, true)
                range.offset(0, -1, 1, 2).setNumberFormats([['#%', '$0.00']]).setValues([[newPercentage, newPrice]])
                break;
            }
            
            spreadsheet.toast('', 'Discount Updated.')
          }
          else
          {
            range.setValue(e.oldValue)
            SpreadsheetApp.flush()
            Browser.msgBox('Item not found on the Discount Percentages sheet.')
          }
        }
      }
      else if (col == 4 || col == 6 || col == 8) // Multiple rows of percentages were changed without authorization
      {
        (row === rowEnd) ? 
          (e.oldValue != undefined) ? 
            range.setValue(e.oldValue).setNumberFormat('#%') : 
          range.setValue(0).setNumberFormat('#%') : 
        range.setNumberFormat('#%').setValues(new Array(rowEnd - row + 1).fill([0]));

        SpreadsheetApp.flush();
        Browser.msgBox('You are NOT authorized to change these discounts. **These percentages have not officially changed, compute your search again to re-display the accurate numbers.')
      }
      else // Multiple rows of prices were changed without authorization
      {
        (row === rowEnd) ? 
          (e.oldValue != undefined) ? 
            range.setValue(e.oldValue).setNumberFormat('$0.00') : 
          range.setValue(0).setNumberFormat('$0.00') : 
        range.setNumberFormat('$0.00').setValues(new Array(rowEnd - row + 1).fill([0]));

        SpreadsheetApp.flush();
        Browser.msgBox('You are NOT authorized to change these discounts. **These prices have not officially changed, compute your search again to re-display the accurate numbers.')
      }
    }
  }
}

/**
 * This function creates the google description from the relevant information.
 * 
 * @author Jarren Ralf
 */
function setGoogleDescription()
{
  const sheet = SpreadsheetApp.getActive().getSheetByName('Discount Percentages');
  const numRows = sheet.getLastRow() - 1;
  const values = sheet.getSheetValues(2, 1, numRows, 9)
    .map(v => [v[0].toString().substring(0, 4) + v[0].toString().substring(5, 9) + v[0].toString().substring(10) + ' - ' + v[2] + ' - ' + v[1] + ' - ' + v[8] + ' - ' + v[3]])
  sheet.getRange(2, 16, numRows).setValues(values)
}

/**
 * This function sorts the data by the vendor name.
 * 
 * @author Jarren Ralf
 */
function sortDataByVendorName(a, b)
{
  a[1] = a[1].toString().toUpperCase();
  b[1] = b[1].toString().toUpperCase();

  return (a[1] === b[1]) ? 0 : (a[1] === '') ? -1 : (b[1] === '') ? -1 : (a[1] < b[1]) ? -1 : 1;
}

/**
 * This function sorts the discount percentages sheet.
 * 
 * @author Jarren Ralf
 */
function sortDiscountPercentagesSheet()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const discountDataSheet = spreadsheet.getSheetByName('Discount Percentages');
  const discountDataRange = discountDataSheet.getRange(2, 1, discountDataSheet.getLastRow() - 1, 12)
  const discountData = discountDataRange.getValues().sort(sortDataByVendorName).sort(sortDataByVendorName)
  discountDataRange.setValues(discountData)
}

/**
 * This function creates all of the triggers that make the spreadsheet function properly.
 * 
 * @author Jarren Ralf
 */
function trigger_CreateAll()
{
  ScriptApp.newTrigger('installedOnEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create()
  ScriptApp.newTrigger('updateAdagioDatabase').timeBased().everyDays(1).atHour(5).create()
  ScriptApp.newTrigger('updateAdagioBasePrice').timeBased().everyDays(1).atHour(6).create()
  ScriptApp.newTrigger('updateDiscountsSheetOnPntShopifyUpdater').timeBased().everyDays(1).atHour(5).create()
}

/**
 * This function deletes all of the triggers.
 * 
 * @author Jarren Ralf
 */
function trigger_DeleteAll()
{
  ScriptApp.getProjectTriggers().map(trigger => ScriptApp.deleteTrigger(trigger));
}

/**
* This function updates the base price from the shopify item comparison.
*
* @author Jarren Ralf
*/
function updateAdagioBasePrice()
{
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.toast('Updating base price...')
  const discountDataSheet = spreadsheet.getSheetByName('Discount Percentages');
  const numRows = discountDataSheet.getLastRow() - 1;
  const discountDataRange = discountDataSheet.getRange(2, 1, numRows, 15)
  const discountData = discountDataRange.getValues()
  const ss = SpreadsheetApp.openById('1sLhSt5xXPP5y9-9-K8kq4kMfmTuf6a9_l9Ohy0r82gI');
  const adagioDataSheet = ss.getSheetByName('FromAdagio');
  const lastUpdated = ss.getSheetByName('Dashboard').getSheetValues(24, 11, 1, 1)[0][0]
  const adagioData = adagioDataSheet.getSheetValues(2, 17, adagioDataSheet.getLastRow() - 1, 2);

  for (var j = 0; j < discountData.length; j++)
  {
    // If any of the disocunt values are blank, then replace them with zeros
    if (discountData[j][12] === '')
      discountData[j][12] = 0;

    if (discountData[j][13] === '')
      discountData[j][13] = 0;

    if (discountData[j][14] === '')
      discountData[j][14] = 0;

    for (var i = 0; i < adagioData.length; i++)
    {
      if (adagioData[i][0].toString().toUpperCase().trim() == discountData[j][0].toString().toUpperCase().trim())
      {
        discountData[j][ 5] = adagioData[i][1];
        discountData[j][11] = adagioData[i][1];
      }
    }
  }

  const numberFormats = new Array(numRows).fill(['@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@', '0', '0', '0', '0'])
  discountDataRange.setNumberFormats(numberFormats).setValues(discountData)
  const text = 'The prices in this spreadsheet were last updated at ' + Utilities.formatDate(lastUpdated, spreadsheet.getSpreadsheetTimeZone(), 'h:mm a   dd MMM yyyy');
  const richTextValue = SpreadsheetApp.newRichTextValue().setText(text)
    .setTextStyle(0, 52, SpreadsheetApp.newTextStyle().setFontSize(18).setBold(false).build())
    .setTextStyle(52, text.length, SpreadsheetApp.newTextStyle().setFontSize(18).setBold(true).build()).build()
  spreadsheet.getSheetByName('Item Search').getRange(2, 2).setRichTextValue(richTextValue).activate()
  spreadsheet.toast('Price update complete.')
}

/**
* This function updates the base price from the shopify item comparison.
*
* @author Jarren Ralf
*/
function updateAdagioDatabase()
{
  const spreadsheet = SpreadsheetApp.getActive();
  const discountDataSheet = spreadsheet.getSheetByName('Discount Percentages');
  const numItems_Initial = discountDataSheet.getLastRow() - 1;
  const numCols = 17;
  const discountData = discountDataSheet.getSheetValues(2, 1, numItems_Initial, numCols)
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  const header = csvData.shift(); // Remove the header
  const sku = header.indexOf('Item #')
  const uom = header.indexOf('Price Unit')
  const description = header.indexOf('Item List')

  const numItems_CSV = csvData.length;
  var googleDescription, vendor, category, fullDescription;

  for (var i = 0; i < numItems_CSV; i++)
  {
    for (var j = 0; j < numItems_Initial; j++)
    {
      if (discountData[j][0].toString().toUpperCase().trim() == csvData[i][sku].toString().toUpperCase().trim()) // SKU
      {
        googleDescription = csvData[i][description].split(' - ');
        googleDescription.pop() // Remove the sku
        googleDescription.pop() // Remove the unit of measure
        category = googleDescription.pop();
        vendor = googleDescription.pop();
        fullDescription = googleDescription.join(' - ');

        discountData[j][ 1] = vendor;
        discountData[j][ 2] = fullDescription;
        discountData[j][ 3] = csvData[i][uom].toString().toUpperCase();
        discountData[j][ 8] = category;
        discountData[j][ 9] = csvData[i][uom].toString().toUpperCase();
        discountData[j][10] = csvData[i][description];
        break;
      }
    }  

    if (j === numItems_Initial) // Item was not found in the newest Adagio data, therefore add it to the discount spreadsheet
    {
      googleDescription = csvData[i][1].split(' - ');
      googleDescription.pop() // Remove the sku
      googleDescription.pop() // Remove the unit of measure
      category = googleDescription.pop();
      vendor = googleDescription.pop();
      fullDescription = googleDescription.join(' - ');

      discountData.push([
        csvData[i][sku], // SKU
        vendor, // Vendor 1 Name
        fullDescription, // Description
        csvData[i][uom], // Unit
        '', // Category Code
        0, // Base Price
        '', // Comments 1
        '', // Comments 2
        category, // Comments 3
        csvData[i][uom], // Unit
        csvData[i][description], // Google Description
        0, // Base Price
        0, // Discount 1 (Guide)
        0, // Discount 2 (Lodge)
        0, // Discount 3 (Wholesale)
        0, // Discount 4 
        0  // Discount 5
      ])
    }
  }

  const numItems_Final = discountData.length;
  const numberFormats = [...Array(numItems_Final)].map(e => ['@', '@', '@', '@', '@', '0', '@', '@', '@', '@', '@', '0', '0', '0', '0', '0', '0']); // Currency format

  Logger.log('numItems_Final: ' + numItems_Final)
  Logger.log('numItems_InitialL: ' + numItems_Initial)

  if (numItems_Final > numItems_Initial)
  {
    Logger.log('Add Items')
    discountData.sort(sortDataByVendorName).sort(sortDataByVendorName) // For some reason I found that double sorting produces the desired result
    discountDataSheet.getRange(2, 1, discountData.length, numCols).setNumberFormats(numberFormats).setValues(discountData)
  }
  else if (numItems_Final === numItems_Initial)
  {
    Logger.log('No Change for items:')
    discountDataSheet.getRange(2, 1, numItems_Initial, numCols).setNumberFormats(numberFormats).setValues(discountData)
  }
}

/**
 * This function updates the Discounts sheet on the PNT Shopify Updater so that AJ is informed as to which items need to be updated on the website.
 * 
 * @author Jarren Ralf
 */
function updateDiscountsSheetOnPntShopifyUpdater()
{
  // Remove the items that are not on the website.
  // Don't include items on sale
  const shopifyUpdateSheet = SpreadsheetApp.getActive().getSheetByName('Shopify Update');
  const lastRow = shopifyUpdateSheet.getLastRow();

  if (lastRow > 1)
  {
    const pntShopifyUpdaterSS = SpreadsheetApp.openById('1sLhSt5xXPP5y9-9-K8kq4kMfmTuf6a9_l9Ohy0r82gI');
    const fromWebsiteDiscountSheet = pntShopifyUpdaterSS.getSheetByName('FromWebsiteDiscount')
    const fromShopifySheet = pntShopifyUpdaterSS.getSheetByName('FromShopify')
    const numRows_FromShopify = fromShopifySheet.getLastRow() - 1;
    const variantSkus = fromShopifySheet.getSheetValues(2, fromShopifySheet.getSheetValues(1, 1, 1, fromShopifySheet.getLastColumn())[0].indexOf('Variant SKU') + 1, numRows_FromShopify, 1); 
    const fromWebsiteDiscounts = fromWebsiteDiscountSheet.getSheetValues(2, 1, fromWebsiteDiscountSheet.getLastRow() - 1, 6)
    const range = shopifyUpdateSheet.getRange(2, 1, lastRow - 1, 5);
    var idx_fromWebsiteDiscounts = -1, idx_fromShopify = -1;

    const itemValues = range.getValues().filter(discountedItem => {

      idx_fromWebsiteDiscounts = fromWebsiteDiscounts.findIndex(shopifyItem => shopifyItem[1] == discountedItem[0])
      
      if (idx_fromWebsiteDiscounts !== -1) // Found the item in the from Website Discount list
        return (discountedItem[2] != fromWebsiteDiscounts[idx_fromWebsiteDiscounts][3] || // Guide
                discountedItem[3] != fromWebsiteDiscounts[idx_fromWebsiteDiscounts][4] || // Lodge
                discountedItem[4] != fromWebsiteDiscounts[idx_fromWebsiteDiscounts][5]) ? // Wholesale
                true : false; 
      else
      {
        idx_fromShopify = variantSkus.findIndex(shopifyItem => shopifyItem[0] == discountedItem[0])

        return (idx_fromShopify !== -1) ? true : false // // Found the item in the from Shopify list
      }
    });

    const numNewDiscounts = itemValues.length;
    
    if (numNewDiscounts !== 0)
    {
      const masterSkus = fromShopifySheet.getSheetValues(2, 1, numRows_FromShopify, 1); 
      const newDiscounts = itemValues.map(item => {item.unshift(masterSkus[variantSkus.findIndex(sku => sku[0] == item[0])][0]); return item}); // Add the master sku to the front
      pntShopifyUpdaterSS.getSheetByName('Discounts').clearContents().getRange(1, 1, newDiscounts.unshift(['Handle', 'SKU', 'Price Type', 'Guide', 'Lodge', 'Wholesale']), 6).setNumberFormat('@').setValues(newDiscounts);
      Logger.log('New discount changes have been written to the Discounts sheet on the PNT Shopify Updater');
    }

    range.clearContent();
  }
  else
    Logger.log('No new discount changes.')
}

/**
* This function checks if the user has pressed delete on a certain cell or not, returning true if they have.
*
* @param {String or Undefined} value : An inputed string or undefined
* @return {Boolean} Returns a boolean reporting whether the event object new value is undefined or not.
* @author Jarren Ralf
*/
function userHasPressedDelete(value)
{
  return value === undefined;
}