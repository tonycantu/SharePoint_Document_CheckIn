# SharePoint Document Check-In
Iterates through a document library using PowerShell and REST calls to report and check in files that are currently checked out.

This PowerShell script iterates through a SharePoint site looking for document libraries that have the "Check In" feature eneabled. If it finds one, it goes into that document library and startes looping through all the documents in the library looking for documents that are checked out. if it finds one, it attempts to check the document back in.

The script also creates two files: a log file and an Excel file which is used for reporting. 
