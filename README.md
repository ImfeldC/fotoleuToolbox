# fotoleu Toolbox
A toolbox to create documents used at fotoleu based on Excel and Word. It creates based on word templates and date retrieved from Excel final word documents to be send to customer, like bill. It includes a Swiss QR code creator for bills used within Switzerland.

# SwissQRCodeExcel
This project was built on existing project SwissQRCodeExcel, see https://github.com/barnstee/SwissQRCodeExcel

QR code generator for Microsoft Excel to be used in the Swiss banking sector. The requirements for the new Swiss QR bill can be read [here](https://www.moneytoday.ch/lexikon/qr-rechnung/).

The QR code generator is an Excel ribbon button that, when clicked, will read the data from the supplied SwissQRBill.xlsx Excel template, generate the QR code and then place it in the correct position in the SwissQRBill.xlsx.
