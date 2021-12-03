/*====================================================================================================================================*
  getAllCrypto by Melviin Khoo
  ====================================================================================================================================
  Version:      1.0.0
  Project Page: https://github.com/melviinkhoo/sheets-crypto
  Help Guide:   https://github.com/melviinkhoo/sheets-crypto/blob/main/README.md
  Copyright:    (c) 20021-2022 by Melviin Khoo
  License:      MIT License
                https://github.com/melviinkhoo/sheets-crypto/blob/main/LICENSE
                
  For future enhancements see https://github.com/melviinkhoo/sheets-crypto/issues?q=is%3Aopen+label%3Aenhancement
  For bug reports see https://github.com/melviinkhoo/sheets-crypto/issues
  ------------------------------------------------------------------------------------------------------------------------------------
  ImportJSON library is required for importing JSON feeds into Google spreadsheets.
     ImportJSON            For use by end users to import a JSON feed from a URL (by Brad Jasper and Trevor Lohrbeer)
  For source code see https://github.com/bradjasper/ImportJSON/blob/master/ImportJSON.gs
  
  ------------------------------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.0  Initial release
 *====================================================================================================================================*/

/**
 * Coingecko API only allows maximum of 1 page, estimated 250 coins data to be returned.  
 * This getAllCrypto() function uses ImportJSON and a for loop to get multiple page data and concatenates the data together before outputing the data.
 * Note that the function will only works when ImportJSON being added to the App Script. See Help Guide for more info.
 * Coingecko API Reference: https://www.coingecko.com/en/api/documentation
 *
 * Four parameter url, query, parseOptions,pageNum is required to be inputted to the function.
 *    @param {url}            The Coingecko API URL with param, 
*                             such as "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=1000&page=1&sparkline=false"
 *    @param {query}          Query data to be returned such as "/name,/symbol,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h"
 *    @param {parseOptions}   Refer to below on changing behavior such as "noTruncate,noHeaders"
 *    @param {loopNum}        Number of pages to be loop, such as "8"
 *
 * Imports a JSON feed and returns the results to be inserted into a Google Spreadsheet. The JSON feed is flattened to create 
 * a two-dimensional array. The first row contains the headers, with each column header indicating the path to that data in 
 * the JSON feed. The remaining rows contain the data. 
 *
 * By default, data gets transformed so it looks more like a normal data import. Specifically:
 *   - Data from parent JSON elements gets inherited to their child elements, so rows representing child elements contain the values 
 *      of the rows representing their parent elements.
 *   - Values longer than 256 characters get truncated.
 *   - Headers have slashes converted to spaces, common prefixes removed and the resulting text converted to title case. 
 *
 * To change this behavior, pass in one of these values in the options parameter:
 *
 *    noInherit:     Don't inherit values from parent elements
 *    noTruncate:    Don't truncate values
 *    rawHeaders:    Don't prettify headers
 *    noHeaders:     Don't include headers, only the data
 *    allHeaders:    Include all headers from the query parameter in the order they are listed
 *    debugLocation: Prepend each value with the row & column it belongs in
 *
 * Full example of working formula:
 *
 *   =getAllCrypto("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=1000&page=1&sparkline=false",
 *                 "/name,/symbol,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h","noTruncate,noHeaders",
 *                 "8",
 *                  doNotDelete!$A$1)
 *
 * 
 **/
function getAllCrypto(url, query, parseOptions,pageNum){
  var fResult={};
  if(url.indexOf("&page=")<0){ //if no &page param found
    fResult = ImportCrypto(url+"&page=1", query, parseOptions)
    
    for(var i=2;i<=parseInt(pageNum);i++){
      var editedURL = url+"&page="+i;
      var result = ImportCrypto(editedURL, query, parseOptions);
      fResult=fResult.concat(result);
    }
  }
else{  //if &page param found
  var searchIndex1 = url.indexOf("&page=");
  var searchIndex2 = url.substring(searchIndex1+6,url.length).indexOf("&");
  fResult = ImportCrypto(url, query, parseOptions)

  for(var i=2;i<=parseInt(pageNum);i++){
    var editedURL = url.substring(0,searchIndex1+6)+i+url.substring(searchIndex1+6+searchIndex2,url.length);
    var result = ImportCrypto(editedURL, query, parseOptions);
    fResult=fResult.concat(result);
    }
  }
return fResult;
}

/**
 * The following function is obtained from the Import CoinGecko Cryptocurrency Data into Google Sheets [Updated 2021]
 *  Souce URL: https://blog.coingecko.com/import-coingecko-cryptocurrency-data-into-google-sheets/
 * 
 *  The function helps to refresh the formula.
 *
 *  A new sheet with the sheet name doNotDelete is required. 
 *  Add the function triggerAutoRefresh() to trigger is requied. 
 *  For the steps to implement, see Help Guide
 *
 **/

function triggerAutoRefresh() {  
    SpreadsheetApp.getActive().getSheetByName('doNotDelete').getRange(1, 1).setValue(getRandomInt(1, 200));
}
// Basic Math.random() function
function getRandomInt(min, max) {
    min = Math.ceil(min);
    max = Math.floor(max);
    return Math.floor(Math.random() * (max - min + 1)) + min;
}
