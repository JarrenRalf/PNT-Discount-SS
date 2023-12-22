/**
 * This function runs the item search sheet.
 * 
 * @author Jarren Ralf
 */
function onEdit(e)
{
  var spreadsheet = e.source;
  var sheet = spreadsheet.getActiveSheet(); // The active sheet that the onEdit event is occuring on
  var sheetName = sheet.getSheetName();

  try
  {
    if (sheetName === "Item Search") // Check if the user is searching for an item or trying to marry, unmarry or add a new item to the upc database
      searchV2(e, spreadsheet, sheet);
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
 * 
 * @author Jarren Ralf
 */
function onOpen()
{
  SpreadsheetApp.getUi().createMenu('PNT Controls')
    .addItem('Update Price', 'updateAdagioBasePrice')
    //.addItem('Update Items (from inventory.csv)', 'updateAdagioDatabase')
    .addToUi();
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

  if (col == colEnd && col === 2) // Column two is being edited
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
          .map(d => [d[0], d[1], d[2], d[3] + '%', (d[2]*(100 - d[3])/100).toFixed(2), d[4] + '%', (d[2]*(100 - d[4])/100).toFixed(2), d[5] + '%', (d[2]*(100 - d[5])/100).toFixed(2)]);
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
          const numberFormats = [...Array(numItems)].map(e => ['@', '@', '$0.00', '@', '$0.00', '@', '$0.00', '@', '$0.00']); // Currency format
          const fontWeights = [...Array(numItems)].map(e => ['bold', 'bold', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold']); // Currency format
          sheet.getRange('B4').activate(); // Move the user to the top of the search items
          itemSearchFullRange.clearContent().offset(0, 0, numItems, 9).setNumberFormats(numberFormats).setFontWeights(fontWeights).setValues(output);
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
    else if (row != rowEnd && row > 3)
    {
      const values = range.getValues().filter(blank => isNotBlank(blank[0]))

      if (values.length !== 0) // Don't run function if every value is blank, probably means the user pressed the delete key on a large selection
      {
        const inventorySheet = spreadsheet.getSheetByName('Discount Percentages');
        const data = inventorySheet.getSheetValues(2, 10, inventorySheet.getLastRow() - 1, 6)
          .map(d => [d[0], d[1], d[2], d[3] + '%', (d[2]*(100 - d[3])/100).toFixed(2), d[4] + '%', (d[2]*(100 - d[4])/100).toFixed(2), d[5] + '%', (d[2]*(100 - d[5])/100).toFixed(2)]);
        var someSKUsNotFound = false, skus;

        if (values[0][0].toString().includes(' - ')) // Strip the sku from the first part of the google description
        {
          skus = values.map(item => {
          
            for (var i = 0; i < data.length; i++)
            {
              if (data[i][1].toString().split(" - ", 1)[0].toUpperCase() == item[0].toString().split(" - ", 1)[0].toUpperCase())
                return data[i]
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0].toString().split(" - ", 1)[0].toUpperCase(), '', '', '', '', '']
          });
        }
        else if (values[0][0].toString().includes('-'))
        {
          skus = values.map(sku => sku[0].substring(0,4) + sku[0].substring(5,9) + sku[0].substring(10)).map(item => {
          
            for (var i = 0; i < data.length; i++)
            {
              if (data[i][1].toString().split(" - ", 1)[0].toUpperCase() == item.toString().toUpperCase())
                return data[i]
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item, '', '', '', '', '']
          });
        }
        else
        {
          skus = values.map(item => {
          
            for (var i = 0; i < data.length; i++)
            {
              if (data[i][1].toString().split(" - ", 1)[0].toUpperCase() == item[0].toString().toUpperCase())
                return data[i]
            }

            someSKUsNotFound = true;

            return ['SKU Not Found:', item[0], '', '', '', '', '']
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
          const numberFormats = [...Array(numItems)].map(e => ['@', '@', '$0.00', '@', '$0.00', '@', '$0.00', '@', '$0.00']); // Currency format
          const fontWeights = [...Array(numItems)].map(e => ['bold', 'bold', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold']); // Currency format
          const WHITE = new Array(9).fill('white')
          const YELLOW = new Array(9).fill('#ffe599')
          const colours = [].concat.apply([], [new Array(numSkusNotFound).fill(YELLOW), new Array(numSkusFound).fill(WHITE)]); // Concatenate all of the item values as a 2-D array

          itemSearchFullRange.clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 9)
              .setFontFamily('Arial').setFontWeight('bold').setFontSize(12).setHorizontalAlignments(horizontalAlignments).setBackgrounds(colours)
              .setBorder(false, null, false, null, false, false).setNumberFormats(numberFormats).setFontWeights(fontWeights).setValues(items)
            .offset(numSkusNotFound, 0, numSkusFound, 9).activate()
        }
        else // All SKUs were succefully found
        {
          var numItems = skus.length
          const horizontalAlignments = new Array(numItems).fill(['center', 'left', 'right', 'right', 'right', 'right', 'right', 'right', 'right'])
          const numberFormats = [...Array(numItems)].map(e => ['@', '@', '$0.00', '@', '$0.00', '@', '$0.00', '@', '$0.00']); // Currency format
          const fontWeights = [...Array(numItems)].map(e => ['bold', 'bold', 'bold', 'normal', 'bold', 'normal', 'bold', 'normal', 'bold']); // Currency format

          itemSearchFullRange.clearContent().setBackground('white').setFontColor('black').setBorder(true, true, true, true, false, false)
            .offset(0, 0, numItems, 9)
              .setFontFamily('Arial').setFontWeight('bold').setFontSize(12).setHorizontalAlignments(horizontalAlignments)
              .setBorder(false, null, false, null, false, false).setNumberFormats(numberFormats).setFontWeights(fontWeights).setValues(skus).activate()
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
  const discountDataSheet = spreadsheet.getSheetByName('Copy of Discount Percentages');
  const discountDataRange = discountDataSheet.getRange(2, 1, discountDataSheet.getLastRow() - 1, 12)
  const discountData = discountDataRange.getValues().sort(sortDataByVendorName).sort(sortDataByVendorName)
  discountDataRange.setValues(discountData)
}

/**
 * This function creates a trigger that will update the pricing daily.
 * 
 * @author Jarren Ralf
 */
function triggerPriceUpdate()
{
  ScriptApp.newTrigger('updateAdagioBasePrice').timeBased().everyDays(1).atHour(6).create()
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
  const discountDataRange = discountDataSheet.getRange(2, 1, discountDataSheet.getLastRow() - 1, 12)
  const discountData = discountDataRange.getValues()
  const ss = SpreadsheetApp.openById('1sLhSt5xXPP5y9-9-K8kq4kMfmTuf6a9_l9Ohy0r82gI');
  const adagioDataSheet = ss.getSheetByName('FromAdagio');
  const lastUpdated = ss.getSheetByName('Dashboard').getSheetValues(24, 11, 1, 1)[0][0]
  const adagioData = adagioDataSheet.getSheetValues(2, 17, adagioDataSheet.getLastRow() - 1, 2);

  for (var i = 0; i < adagioData.length; i++)
  {
    for (var j = 0; j < discountData.length; j++)
    {
      if (adagioData[i][0] === discountData[j][0].toString().trim())
      {
        discountData[j][ 5] = adagioData[i][1];
        discountData[j][11] = adagioData[i][1];
        break;
      }
    }
  }

  discountDataRange.setNumberFormat('@').setValues(discountData)
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
  const discountDataSheet = spreadsheet.getSheetByName('Copy of Discount Percentages');
  const numItems_Initial = discountDataSheet.getLastRow() - 1;
  const numCols = 17;
  const discountData = discountDataSheet.getSheetValues(2, 1, numItems_Initial, numCols)
  const csvData = Utilities.parseCsv(DriveApp.getFilesByName("inventory.csv").next().getBlob().getDataAsString());
  csvData.shift(); // Remove the header
  const numItems_CSV = csvData.length;
  var googleDescription;

  for (var i = 0; i < numItems_CSV; i++)
  {
    for (var j = 0; j < numItems_Initial; j++)
      if (discountData[j][0] == csvData[i][6].toString().toUpperCase().trim()) // SKU
        break;

    if (j === numItems_Initial) // Item was not found in the newest Adagio data, therefore add it to the discount spreadsheet
    {
      googleDescription = csvData[i][1].split(' - ');

      discountData.push([
        csvData[i][6], // SKU
        googleDescription[2], // Vendor 1 Name
        googleDescription[1], // Description
        csvData[i][0], // Unit
        '', // Category Code
        0, // Base Price
        '', // Comments 1
        '', // Comments 2
        googleDescription[3], // Comments 3
        csvData[i][0], // Unit
        csvData[i][1], // Google Description
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

  Logger.log(numItems_Final)
  Logger.log(numItems_Initial)

  if (numItems_Final > numItems_Initial)
  {
    Logger.log('Add Items')
    //discountData.sort(sortDataByVendorName).sort(sortDataByVendorName) // For some reason I found that double sorting produces the desired result
    discountDataSheet.getRange(2, 1, discountData.length, numCols).setNumberFormats(numberFormats).setValues(discountData)
  }
  else if (numItems_Final === numItems_Initial)
  {
    Logger.log('No Change for items:')
    discountDataSheet.getRange(2, 1, numItems_Initial, numCols).setNumberFormats(numberFormats).setValues(discountData)
  }
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