"VIMPACT"
# Visma Integration & Maconomy - Payroll Accounting Conversion Tool
Date: 30th of October 2024

This tool facilitates the conversion of payroll accounting data file from Visma Payroll to Delktek Maconomy ERP.
It solves the following challenges: VAT code, textual description, employee dimension and company specific customizations.

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
You can also edit the project listing for special handeling of projects with VAT.  ALternativly use API to fetch data from Maconomy

4) Format the Debit and Credit cells: general number with two digits.

5) Copy the content of out.xlsx to a text file (copy - paste). Do not try Save As text file from Excel. Excel add quotation marks to text strings that contains special characters. 

6) Import file in Maconomy - Import General Journal. Rembember to check "internal popup names"

## Contributing

We welcome contributions! 

## License

This project is licensed under the Apache License 2.0. 


## Contact

For any questions or feedback, please contact us at post@cicero.oslo.no.
