"VIMPACT"
# Visma Integration & Maconomy - Payroll Accounting Conversion Tool
Date: 30th of October 2024

This tool facilitates the conversion of payroll accounting data file from Visma Payroll to Delktek Maconomy ERP.
Please configure Visma Payroll to output H&L fixed width file format. Currenty, its the only format that supports VAT codes.

## Features

- Seamless data conversion
- Customization
- High accuracy and reliability

## Usage

1) Export the H & L accounting file from Visma Payroll. You don't have to move it away from the Dowloads folder.
2) Export the "Transaksjoner, detaljert" Excel file from Visma Payroll.

Specify the following columns:
Lønnsperiode, ansattnummer, lønnsart, beløp, tekst and reiseregningID.
Please filter on Lønnsartgrupper = Expense. Remember to adjust Fra/til lønnskjøring to filter out the transaction target.

3) Modify the file mapping.xlsx and enter the relationship between account/activity and task number. 
You can also edit the project listing for special handeling of projects with VAT.  

4) Copy the content of out.xlsx to a text file (copy - paste). Do not try Save As text file from Excel. Excel add quotation marks to text strings that contains special characters. 

5) Import file in Maconomy - Import General Journal. Rembember to check "internal popup names"

## Contributing

We welcome contributions! 

## License

This project is licensed under the Apache License 2.0. 


## Contact

For any questions or feedback, please contact us at post@cicero.oslo.no.
