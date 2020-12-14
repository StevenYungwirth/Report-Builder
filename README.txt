The report builder takes a .csv file exported from Total Rebalance Expert (TRX), creates a trade recommendations report, and prepares the recommendations letter to be sent to the client.

Instructions:
1. Open Report Builder.xlsm and TradeRecommendationsExport.csv.
2. With TradeRecommendationsExport.csv open, click on "Build Report from TradeRecommendationsExport.csv".

The buttons for building the report will do the following:
1. The report will be generated as a new tab in the exported .csv
1a. The trades will be separated by account, and the accounts' name, type, and custodian will be shown.
1b. The trades within each account will be split between buys and sells and sorted alphabetically by fund name.
1c. The client's equity target will be placed in the top-right corner, and the client's name and the current date will be put into the header.
2. The macro will open a save dialog, first attempting to save the report in the client's folder on the network drive (Z drive). If the location is unavailable, it will default to the user's default save location.
3. The print preview will be shown so the report may be printed immediately.

The buttons for building the letter will do the following:
1. The letter will be opened in Microsoft Word, but it will be hidden until the letter is processed.
2. The macro will iterate through each word, looking for key words to replace with necessary information to complete the letter.
2a. "DATE" will be replaced with the current date.
2b. "Dear:" will be replaced with "Dear: " and the client's name.
2c. "BUYSELLTARGET" will be replaced with the amount of equities either bought or sold, and the client's equity target.
2d. If one of the Insert Options is selected on the report builder, "INSERT" will be replaced with a paragraph regarding withdrawals or tax loss harvesting.
2e. When "Regards" is found, a subsearch will be run looking for the advisors' names. When they are found, they will be replaced with the advisors checked off on the report builder.
3. The macro will open a save dialog, first attempting to save the report in the client's folder on the network drive (Z drive). If the location is unavailable, it will default to the user's default save location.
4. The print preview will be shown so the report may be printed immediately.

Note: For both the report and letter, if the macros are run after 3PM, the dates will be one day ahead (Or three days, if run after 3PM on a Friday). This is simply the cutoff at FPIS so there is enough time to send the letters out for the day.