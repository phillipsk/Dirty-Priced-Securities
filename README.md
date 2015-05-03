# Dirty-Priced-Securities
Reconcile Dirty-priced securities from a client list against an internal list. Adding any dirty-priced securities not found on the internal EXCEL spreadsheet but on the external client excel spreadsheet, to the internal spreadsheet via the Python programming language and the Python-Excel reader library openpyxl. No Microsoft Excel Macro Visual Basic code needed.

## Installation

After cloning the repository first time, create a virtual environment with the virtualenv command:
``virtualenv env``

Then install the needed packages:
``pip install -r requirements.txt``

This latter step needs to be done each time you update the repository (just in case new requirements have been added
to the project).

## Run

To run the application type
``python merge.py name_of_reconciled_xlsx name_of_daily_data_xlsx``

That's it.
