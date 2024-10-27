"VIMPACT"
# Visma Integration & Maconomy Payroll Accounting Conversion Tool
Date: 27th of October 2024

This tool facilitates the conversion of payroll accounting data file between Visma Payroll and Delktek Maconomy ERP systems.
Please configure Visma Payroll to output H&L fixed width file format. Currenty, its the only format that supports VAT codes.

## Features

- Seamless data conversion
- Customization
- High accuracy and reliability

## Usage

1) Export the H & L accounting file from Visma Payroll. You can leave it in the Dowloads folder.
2) Export the "Transaksjoner, detaljert" Excel file from Visma Payroll.

Specify the following columns:
Lønnsperiode, ansattnummer, lønnsart, beløp, tekst, reiseregningID
Please filter on Lønnsartgrupper = Expense. Remember to adjust Fra/til lønnskjøring to filter out the transaction target.
3) Modify the file mapping.xlsx and enter the relationship between account/activity and task number. 
You can also edit the project listing for special handeling of projects with VAT.  

## Contributing

We welcome contributions! 

## License

This project is licensed under the Apache License 2.0. See the [LICENSE](LICENSE) file for details.


## Contact

For any questions or feedback, please contact us at post@cicero.oslo.no.
