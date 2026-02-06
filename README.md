# sap-finance-automation

This repository serves as a portfolio showcase demonstrating how VBA and SAP GUI Scripting can be used to eliminate repetitive manual tasks in Enterprise Resource Planning (ERP) environments.

```

The Excel sheet for the F28 Input Automation script should look like this:

		 A  |  B  |  C  |             D              |             E              |       F       | ...
------+-----+-----+-----+----------------------------+----------------------------+---------------+----
  1   |     |     |     |                            |                            |               |
------+-----+-----+-----+----------------------------+----------------------------+---------------+----
  2   |     |     |     |               F-28 Input Automation (MAX. 990)          |               |  [ BUTTON TO CLEAR CELLS ]
------+-----+-----+-----+----------------------------+----------------------------+---------------+----
  3   |     |     |     |      Invoice number		 |   Partial payment amount   |               |
      |     |     |     |                  		     |                            |               |
------+-----+-----+-----+----------------------------+----------------------------+---------------+----
  4   |     |     |     | [ START LISTING DATA HERE ]| [ START LISTING DATA HERE ]|               |   
      |     |     |     |         -> Cell D4         |         -> Cell E4         |               |
------+-----+-----+-----+----------------------------+----------------------------+---------------+----
  5   |     |     |     | ...                        | ...                        |               |
  6   |     |     |     | ...                        | ...                        |               |
```
## F28 Input Automation Script
This VBA script automates the input process for the **SAP F-28 (Incoming Payments)** transaction. It is designed to handle large volumes of partial payment allocations where manual data entry is slow and error-prone. 

The script connects Excel directly to the SAP GUI Scripting API, reads invoice and payment data from the active worksheet, and inputs it into the SAP interface.

## Core Functionality
The primary purpose of this tool is to overcome the UI limitations of the SAP Table Control:
1.  **Selection Screen Pagination:** It handles lists larger than the default window by automatically filling out the form as many times as needed.
2.  **Open Item Table Scrolling:** In the payment window, it automatically fills the rows & scrolls down automatically in order to continue filling, as long as there is data.

## The algorithm
1.  **Connection:** Attaches to the running SAP GUI session.
2.  **Data Parsing:** Reads the Document Number (Col D) and Payment Amount (Col E) into arrays. (The same array is reused for memory efficiency)
3.  **Invoice Selection:** Iterates through the array and populates the input fields, sending a 'Refresh' command (Enter) or adjusting the slider when fields are full.

## Requirements
* Microsoft Excel (VBA enabled)
* SAP Logon (GUI Scripting must be enabled on client and server side)
* Access to transaction F-28 (or compatible Z-transaction)

## Usage
1.  Open the Excel file containing the payment data.
2.  Ensure Document Numbers are in **Column D** and Amounts in **Column E** (starting row 4).
3.  Log in to SAP.
4.  Enter F28 and get to the part where you need to input multiple invoice numbers.
5.  Run the script.
6.  Use the 'ClearInputCells' subroutine to easily clear the data from the Excel file.
