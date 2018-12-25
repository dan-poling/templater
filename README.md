# templater

template_filler:
This uses a excel sheet to fill in a word doc template to make a new document (template + sheet = result).

The docx template needs unique placeholders that match up with the 
first column in the xlsx sheet. The second column of the xlsx sheet
should contain the values that will flow into the final form.

This program uses the following libraries: re to find/replace, python-docx to read & save .docx, and pandas to read in .xlsx

TODO: make this work with google docs & sheets.
