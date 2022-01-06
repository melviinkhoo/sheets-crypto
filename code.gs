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
 * This getAllCrypto() function uses ImportJSON and a for loop to get multiple page data and concatenates the data together before outputing the data to sheet.
 * Note that the function will only works when ImportJSON being added to the App Script. See Help Guide for more info.
 * Coingecko API Reference: https://www.coingecko.com/en/api/documentation
 *
 * Four parameter url, query, parseOptions,pageNum is required to be inputted to the function.
 *    @param {url}            The Coingecko API URL with param, 
 *                            such as "https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=1000&page=1&sparkline=false"
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
 *    noHeaders:     Don't include headers, only the data (recommended)
 *    allHeaders:    Include all headers from the query parameter in the order they are listed
 *    debugLocation: Prepend each value with the row & column it belongs in
 *
 * Full example of working formula:
 *
 *   =getAllCrypto("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=1000&page=1&sparkline=false",
 *                 "/name,/symbol,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h","noTruncate,noHeaders",
 *                 "30",
 *                  doNotDelete!$A$1)
 *
 *  The doNotDelete!$A$1 is used to auto-refresh the formula when triggerAutoRefresh() function is triggered.
 *  The sheet with the name doNotDelete is required for auto-refresh to work.
 *  The loopNum must be lesser than 34, Google Sheets limit custom formula of 30 seconds run time.
 *  Please use minimal loop number, else error below will be obtained:
 *  Exception: Request failed for https://api.coingecko.com returned code 429. Truncated server response: error code: 1015
 *    Due to when Apps Script runs a script, the script is assigned to one of the Google Cloud nodes. 
 *    This notes makes an outbound IP connection to fetch the data from Coinmarketcap. 
 *    When one node (ip address) is generating to much traffic on Coinmarketcap it may get banned for a period of time.
 *    See the source of explanation: https://www.reddit.com/r/Cointrexer/comments/8lqtfo/coinmarketcap_error_429/
 *
 **/
function getAllCrypto(url, query, parseOptions,loopNum){
  //Logger.log("#URL: "+url+" #query: "+query+" #parseOptions: "+parseOptions+" #loopNum: "+loopNum) //troubleshooting use only
  var fResult={};
  try{
    if(url.indexOf("&page=")<0){ //if no &page param found
      fResult = ImportJSON(url+"&page=1", query, parseOptions)
      for(var i=2;i<parseInt(loopNum);i++){
        Logger.log("loop "+i); //To record in log. (Optional)
        var editedURL = url+"&page="+i;
        var result = ImportJSON(editedURL, query, parseOptions);
        fResult=fResult.concat(result);
        Utilities.sleep(100); //to reduce number of api call, set lower number if maximum time execution error occur.
      }
    }else{  //if &page param found
      var searchIndex1 = url.indexOf("&page=");
      var searchIndex2 = url.substring(searchIndex1+6,url.length).indexOf("&");
      var pageNum = url.substring(searchIndex1+6,searchIndex1+6+searchIndex2)
      fResult = ImportJSON(url, query, parseOptions)
      for(var i=parseInt(pageNum)+1;i<parseInt(loopNum);i++){
        Logger.log("loop "+i); //To record in log. (Optional)
        var editedURL = url.substring(0,searchIndex1+6)+i+url.substring(searchIndex1+6+searchIndex2,url.length);
        var result = ImportJSON(editedURL, query, parseOptions);
        fResult=fResult.concat(result);
        Utilities.sleep(100);//to reduce number of api call, set lower number if maximum time execution error occur.
      }
    }
    
    return fResult;
  }
  catch(e){
    Logger.log("Error: "+e);
    if(fResult[0]==undefined){
      return "Error loading due to too much request made to Coingekco, try again later."
    }else{
      return fResult
    }
  }
}

/**
 * 
 *  The function uses getAllCrypto function and write the crypto data directly to the sheet name getAllCoins
 *  
 *  It is an alternative to getAllCrypto function as getAllCrypto formula sometimes will get the error below, and returned no data:
 *  Exception: Request failed for https://api.coingecko.com returned code 429. Truncated server response: error code: 1015
 *    Due to when Apps Script runs a script, the script is assigned to one of the Google Cloud nodes. 
 *    This notes makes an outbound IP connection to fetch the data from Coinmarketcap. 
 *    When one node (ip address) is generating to much traffic on Coinmarketcap it may get banned for a period of time.
 *    See the source of explanation: https://www.reddit.com/r/Cointrexer/comments/8lqtfo/coinmarketcap_error_429/
 *
 * It needs to be triggered with a Google Apps Scripts trigger at https://script.google.com/home/:
 *   - Select project and add trigger
 *   - Choose which function to run: triggerAutoRefresh
 *   - Select event source: Time-driven
 *   - Select type of time based trigger: Minutes timer
 *   - Select minute interval: 5 minutes (to avoid too many requests)
 * 
 *  Sheet with name getAllCoins in required with the following parameter settings:
 *  Param need to be at specific cell 
 *    @param {url}            Cell B1
 *    @param {query}          Cell D1
 *    @param {parseOptions}   Cell F1
 *    @param {loopNum}        Cell H1
 *  Please refer to getAllCrypto function to udnerstand more about the parameter. 
 *  The loopNum can go up to 200 (tested) 
 *
 *  Example parameter input in cell:
 *    @param {B1}           https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=300&page=1&sparkline=false
 *    @param {D1}           "/symbol,/symbol,/name,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h"
 *    @param {F1}           noTruncate,noHeaders
 *    @param {H1}           200
 *
 *  Please refer to Help Guide for more info
 **/

function writeDataToSheet(){
  var ss = SpreadsheetApp.getActive().getSheetByName('getAllCoins');
  var params = ss.getRange(1,1,1,8).getDisplayValues();
  var getData,dateRange;
  var dateNow = new Date();
  
  try{
    getData = getAllCrypto(params[0][1].toString(),params[0][3].toString(),params[0][5].toString(),params[0][7].toString());
    ss.getRange(2, 1,getData.length,getData[0].length).setValues(getData);
    dateRange=ss.getRange(2, getData[0].length+1,getData.length,1);
    dateRange.setValue(dateNow); 
    dateRange.setNumberFormat("dd-mmm-yy h:mm:ss");
    Logger.log("passed data: "+getData.length); //To record in log. (Optional)
  }
  catch(e){
    if(getData.length==undefined){
      Logger.log("Error in Write: "+e); //To record in log. (Optional)
    }else{
      Logger.log("passed data but error: "+getData.length); //To record in log. (Optional)
      
      if(getData[0].length!=1){
        ss.getRange(2, 1,getData.length,getData[0].length).setValues(getData);
        dateRange=ss.getRange(2, getData[0].length+1,getData.length,1);
        dateRange.setValue(dateNow); 
        dateRange.setNumberFormat("dd-mmm-yy h:mm:ss");
      }
    }
  }
}


/**
 * The following function is obtained from the Import CoinGecko Cryptocurrency Data into Google Sheets [Updated 2021]
 *  Souce URL: https://blog.coingecko.com/import-coingecko-cryptocurrency-data-into-google-sheets/
 * 
 *  The function helps to refresh the formula.
 *
 * This function by Vadorequest generates a random number in the "randomNumber" sheet.
 *
 * It needs to be triggered with a Google Apps Scripts trigger at https://script.google.com/home/:
 *   - Select project and add trigger
 *   - Choose which function to run: triggerAutoRefresh
 *   - Select event source: Time-driven
 *   - Select type of time based trigger: Minutes timer
 *   - Select minute interval: 10 minutes (to avoid too many requests)
 * Please refer to Help Guide for more info
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
