sheets-crypto
# Google Sheets Get All Crypto Coin App Script with Coingecko API

**Google Sheets Get All Crypto Coin App Script** allows users to get the Crypto Coin data in full without max limit of 250 coins per page.

Coingecko API only allows maximum of 1 page, estimated 250 coins data to be returned in Google App Script. 

This repo is build to get multiple page data and concatenates the data together before outputing the data to sheet.

### There is two functions that is developed to get the mentioned function:
1. getAllCrypto() function allows user to use formula to obtain the results
2. writeDataToSheet() function allows data to be written directly to the sheet name getAllCoins

**Advantages & Disadvantages**
|   Function Name  |   Advantage                                             |  Disadvantage                                 |   Type   |
| ---------------- | ------------------------------------------------------- | --------------------------------------------- | -------- |
|   getAllCrypto   | Easy to use, use it in any cell of the sheets           | Low loop count, api error will clear all data | formula  |
| writeDataToSheet | High loop count, data remain when api error, time stamp | Must have specific sheet                      | function |


## Requirement to set up
1. Get ready your Google Sheets to use.
2. Select which function (1) or (2) to be use, both have different requirement and it is not advisible to run both.
3. Go to your Google Sheets, click on "Extentions > App Script".
4. Click on the "+" icon besides the file, then select "Script".
5. Name it as ImportJSON or anything you like.
6. Go to [here](https://github.com/bradjasper/ImportJSON/blob/master/ImportJSON.gs) and copy raw content (the whole code) and paste it to the script file you just created.
7. Click on "Code.gs" in your App Script.
8. Go to [this project's code.gs](https://github.com/melviinkhoo/sheets-crypto/blob/main/code.gs) and copy raw content (the whole code) and paste it to the selected script file.

  ### If you plan to use (1)getAllCrypto custom formula
  9. Click on "Trigger" on the left hand panel in the Script Editor.
  10. Select "Add Trigger" then follow the settings below and press "Save":
  ![Adding trigger for getAllCrypto](/assets/images/trigger_getAllCrypto.PNG)
  
  <br />(You will get error if the next steps is not done within 5 minutes. The error can be ignored.)
  <br />(You can also delete the trigger and add it again after finishing the last step, or change it to notify me weekly if you find the error annoying.)
  
  12. Back to your Google Sheets, create a new sheet tab with the name "doNotDelete".
  13. Input the following formula to any of the cell in any of the sheet tab:
  
  ```
  =getAllCrypto(url,
  query,parseOptions,
  loopNum,
  doNotDelete!$A$1)
  ```
  
  The following param needs to but inputted into the formula:
  <br />**url**:            The Coingecko API URL with param, such as **"https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=id_asc&per_page=1000&page=1&sparkline=false"**
  <br />                (refer to [Coingecko API documentation](https://www.coingecko.com/en/api/documentation))
  <br />
  <br />**query**:          Query data to be returned such as **"/name,/symbol,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h"**
  <br />                (refer to [Coingecko API documentation](https://www.coingecko.com/en/api/documentation))
  <br />
  <br />**parseOptions**:   Changing behavior such as **"noTruncate,noHeaders"**
  <br />                (refer to [ImportJSON](https://github.com/bradjasper/ImportJSON/blob/master/ImportJSON.gs))
  <br />
  <br />**loopNum**:        Number of pages to be loop, such as **"8"**
  <br />                !The loopNum must be lesser than 34, Google Sheets limit custom formula of 30 seconds run time.
  
  Example of fully working formula:
   ```
   =getAllCrypto("https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=market_cap_desc&per_page=1000&page=1&sparkline=false",
   "/name,/symbol,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h","noTruncate,noHeaders",
   "30",
   doNotDelete!$A$1)
   ```
   The formula will start to run and it will automatically refresh every 5 mins (or when triggerAutoRefresh() function is run).
   Please refer to code.gs on the explanation of each function.
   
   Note that you will get error such as "Error loading due to too much request made to Coingekco, try again later". This is due to API limitation/error.
   If you found that the coin data obtained is lesser than the loop, this is also due to API limitation/error, the formula will return all data obtained prior to the error.
   Coingecko API error/limitation explaination:
    Please use minimal loop number, this is due to when Apps Script runs a script, the script is assigned to one of the Google Cloud nodes. 
    This notes makes an outbound IP connection to fetch the data from Coinmarketcap. 
    When one node (ip address) is generating to much traffic on Coinmarketcap it may get banned for a period of time.
      See the source of explanation: https://www.reddit.com/r/Cointrexer/comments/8lqtfo/coinmarketcap_error_429/

  
  ### If you plan to use (2)writeDataToSheet function
  9. Click on "Trigger" on the left hand panel in the Script Editor.
  10. Select "Add Trigger" then follow the settings below and press "Save":
  ![Adding trigger for writeDataToSheet](/assets/images/trigger_writeDataToSheet.PNG)
  (You will get error if the next steps is not done within 5 minutes. The error can be ignored.) 
  (You can also delete the trigger and add it again after finishing the last step, or change it to notify me weekly if you find the error annoying.)
  
  12. Back to your Google Sheets, create a new sheet tab with the name "getAllCoins".
  13. Input the following parameters at the following cell:
        
        Cell A1: **URL:**
        Cell B1: url
        Cell C1: **Query:**
        Cell D1: query
        Cell E1: **parseOptions:**
        Cell F1: parseOptions
        Cell G1: **loopNum:**
        Cell H1: loopNum
        
  
  The following param needs to but inputted into the formula:
  <br />**url**:            The Coingecko API URL with param, such as **https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=id_asc&per_page=1000&page=1&sparkline=false**
  <br />                (refer to [Coingecko API documentation](https://www.coingecko.com/en/api/documentation))
  <br />
  <br />**query**:          Query data to be returned such as **"/name,/symbol,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h"**
  <br />                (refer to [Coingecko API documentation](https://www.coingecko.com/en/api/documentation))
  <br />
  <br />**parseOptions**:   Changing behavior such as **noTruncate,noHeaders**
  <br />                (refer to [ImportJSON](https://github.com/bradjasper/ImportJSON/blob/master/ImportJSON.gs))
  <br />
  <br />**loopNum**:        Number of pages to be loop, such as **150**

  
  Example of fully working cell items:
  
    Cell B1: https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&order=id_asc&per_page=300&page=1&sparkline=false
    Cell D1: "/symbol,/symbol,/name,/current_price,/market_cap,/price_change,/total_volume,/high_24h,/low_24h"
    Cell F1: noTruncate,noHeaders
    Cell H1: 150
  
   The function  will start to run every 5 mins and it will automatically refresh every 5 mins (or when writeDataToSheet() function is run).
   Please refer to code.gs on the explanation of each function.
   
   The time stamp shows the time of the data is refreshed and you may possibly see coin data with different time stamp. 
   This is due to API limitation/error, the formula will return all data obtained prior to the error, hence differnt time stamp.
   Coingecko API error/limitation explaination:
    This is due to when Apps Script runs a script, the script is assigned to one of the Google Cloud nodes. 
    This notes makes an outbound IP connection to fetch the data from Coinmarketcap. 
    When one node (ip address) is generating to much traffic on Coinmarketcap it may get banned for a period of time.
      See the source of explanation: https://www.reddit.com/r/Cointrexer/comments/8lqtfo/coinmarketcap_error_429/
     <br /><br />The suggestion is to change the api to sort the coin (such as order by market cap) so that all your preferred coin is at the top for better refresh time. 


## You can then, use Vlookup formula to select information to be pull to your specific cell.
For example: 

```
=VLOOKUP(lower(A1), getAllCoins!$A$2:$K,3, false)
```

Input the symbol of the coin (eg: ETH) to cell A1 and the formula on cell A2. The formula will then return the price for ETH.<br />
<a href="https://docs.google.com/spreadsheets/d/1Nqnynzybk-VThI06GcfTq9o4-wjxZ9BWZNItK_9vOQ4/edit?usp=sharing" target="_blank">View Only Google Sheets sample</a> for this project (please do not request access).<br />

END
------------------------------------------------------------------------------------------------------------------------------------
