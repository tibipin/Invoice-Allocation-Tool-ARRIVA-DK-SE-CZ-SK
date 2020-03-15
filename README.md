# Invoice-Allocation-Tool-ARRIVA-DK-SE-CZ-SK

This invoice allocation tool was created for DB Schenker GBS Bucharest's AP Department.
The tool was created for four different countries within the same entity. 
The tools works as follows:
  - It imports WINI_OOP.py which is a partial xpath mapping of the WINI Website - the Website on which the AP department processes the invoices from the suppliers. 
  The tool accesses the WINI website, with the credentials input by the user, and automatically downloads the reports based on the country the user is working for.
  - The tool downloads these reports in a temporary folder, merges them and performs calculations of the daily invoice backlog, inflow and processing performance. The merge and calculations are performed using the pandas module.
  - The tool also has a basic GUI created in tkinter so it can be easily be used by users which do not use scripting.
