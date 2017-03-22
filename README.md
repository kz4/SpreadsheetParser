# How to use the SpreadSheetParser
1. Browse a spreadsheet that you want to parSelect
2. Select the Column, StartRow and EndRow you want to choose from
3. Select how many words you want on a single line
4. Click the Parse the button to parse it

![alt tag](https://raw.githubusercontent.com/kz4/SpreadsheetParser/master/SpreadsheetParser/Images/SpreadSheet.PNG)

![alt tag](https://raw.githubusercontent.com/kz4/SpreadsheetParser/master/SpreadsheetParser/Images/Result.PNG)

# How to use the ConnectWise API Updater
1. Click the CW Helper on the main page

2. Enter the 
```
PublicKey = "";
PrivateKey = "";
```
in ConnectWiseService.cs

3. File Name is basically the ConnectWise URL you need to update: for example, in a spreadsheet you have a column of IDs 123, 234, 345. In this case, we will get:
```
https://connectwiselab.yourcompany.com/v4_6_release/apis/3.0/service/tickets/123
https://connectwiselab.yourcompany.com/v4_6_release/apis/3.0/service/tickets/234
https://connectwiselab.yourcompany.com/v4_6_release/apis/3.0/service/tickets/345 
```

4. Change the companyname to your company name, and the rest of parameters per your needs

![alt tag](https://raw.githubusercontent.com/kz4/SpreadsheetParser/master/SpreadsheetParser/Images/ConnectWiseUpdateScreen.PNG)

![alt tag](https://raw.githubusercontent.com/kz4/SpreadsheetParser/master/SpreadsheetParser/Images/ConnectWiseApi1.PNG)