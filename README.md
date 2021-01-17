# fotoleu Toolbox
A toolbox to create documents used at fotoleu based on Excel and Word. It creates based on word templates and data retrieved from Excel final word documents to be send to customer, like bill. It includes a Swiss QR code creator for bills used within Switzerland.

Up to V2.x the values for the QR code have been read from an excel sheet with name "SwissQRCode" from fixed cells. This was not very convienend.

Starting from V3.x the values for the QR code are read from a table named "TabQRCode". This is more flexible and allows better integration into other excel workbooks.

## Overview
The main logic is in fotoleuToolbox.cs, which contains three main functions:
1. generateAuftragsblatt: Generates a document, loads a word template and replaces placeholders with real values from excel table named "TabABBookmarks"
2. generateQRCodeV2: Generates QR bitmap code, replaces them in the template and stores the newly generated file under the name passed by strFilePath.
3. generateRechnung: Main function, combines the two methods generateAuftragsblatt and generateQRCodeV2 and creates combined final document.

## How to setup in your environment
You can easily setup this toolbox in your environment, be following the steps below.
1. Install the published toolbox; this install the ribbon button "fotoleu Toolbox" in your excel
2. Copy the two tables "TabABBookmarks" and "TabQRCode" form example excel workbook: fotoleuToolbox.xlsx
3. Link/adjust the values in second column of the tables above, with your values (e.g. link them with values out of your excel workbook)
4. In case you like to trace debug messages within excel, copy also the sheet "SwissQRCode-Debug" to you excel; this allows you to enable/disable debug messages
5. Done :-)
 

# SwissQRCodeExcel
This project was built on existing project SwissQRCodeExcel, see https://github.com/barnstee/SwissQRCodeExcel

QR code generator for Microsoft Excel to be used in the Swiss banking sector. The requirements for the new Swiss QR bill can be read [here](https://www.moneytoday.ch/lexikon/qr-rechnung/).

The QR code generator is an Excel ribbon button that, when clicked, will read the data from the supplied SwissQRBill.xlsx Excel template, generate the QR code and then place it in the correct position in the SwissQRBill.xlsx.
