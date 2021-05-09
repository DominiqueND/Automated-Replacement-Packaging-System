# Automated-Replacement-Packaging-System
The project this code was designed for was a device replacement and packaging system that would take information from a list of devices that need to be replaced, and create a new list of information about the replacement devices.This Python code will read the files from an Excel file (Agency Test Names), anaylze it, scan in new data, and create a new Excel file (New Barcode Scans 1) with the new information.
Python will search the Excel file for:
  Agency Name: The city the "agency" is in.
  Batch Number: There are 6 devices in a batch, and each batch goes to only one "agency". This is the number an old device's serial number is in a batch, starting over at 6 and                 at each new "agency".
  Agency Number: The number an old device's serial number is for an "agency".
  Old Serial Number: The old serial number accosiated with that agency.
Python will ask the user for input of:
  New Serial Number: The replacement device's serial number that is scanned during the packaging process.
Hand Scan is a refrence column if you wish to check the results by hand.
All of this will be collected and writen into a new Excel that can be named at the bottom of the code (New Barcode Scans 1). It will be saved into the same location as the old Excel (Agency Test Names).
