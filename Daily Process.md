<h2>Update Database</h2>

- Click on <b>DBLastGaspUpdate</b> (Smiley Face)
- <b>Yesterday's Date</b> will be prefilled (Last Gasp is always 1 day late)
- Click <b>Done</b>
- It will either update the database or tell you it has already been updated.

<h2>Process SilverSprings Data</h2>

- Download Silver Springs data into Excel
- Click on <b>SSNSplitFile</b> (This will save the files by time segment.)
- Repeat until all .csv have been processed.  
  At this point you may have several files SSN-YYYY-MM-DD-HHMMSS-HHMMSS.xlsx  
  YYYY-MM-DD is the date of the SSN data  
  the first HHMMSS is the beginning time of the SSN data  
  the second HHMMSS is the ending time of the SSN data
- The next step is to <b>merge</b> the time segment files


<h2>Merge SSN Files</h2>

- Click on <b>SSNMerge</b> (this will merge the time segment files.)  
  At this point you should have a single file SSN-YYYY-MM-DD.xlsx  
  The time segment files have been moved to subdirectory \Processed SSN Downloads

<h2>Download Last Gasp Report</h2>

- Click on <b>Query</b>
- Select <b>Last Gasp Daily</b>
- Click on <b>Submit</b>
- Popup will tell you how many records were found (Defaults to all records.)
- Click on <b>Submit</b>
- Records will be loaded into Excel (progress should show in lower lefthand corner)
- "Teradata Download Finished" will apppear when completed.
