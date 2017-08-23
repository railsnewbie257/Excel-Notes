<h2>Update Database</h2>

- Click on <b>DBLastGaspUpdate</b> (Smiley Face)
- <b>Yesterday's Date</b> will be prefilled (Last Gasp is always 1 day late)
- Click <b>Done</b>
- It will either update the database or tell you it has already been updated.

<h2>Process SilverSprings Data</h2>

- Download Silver Springs data into Excel
- Click on <b>SSNSplitFile</b> (This will save the files by time segment.)
- Repeat until all .csv have been processed.  
  <em>At this point you may have several files SSN-YYYY-MM-DD-HHMMSS-HHMMSS.xlsx  
  YYYY-MM-DD is the date of the SSN data  
  the first HHMMSS is the beginning time of the SSN data  
  the second HHMMSS is the ending time of the SSN data</em>
- The next step is to <b>merge</b> the time segment files


<h2>Merge SSN Files</h2>

- Click on <b>SSNMerge</b> (this will merge the time segment files.)  
  <em>At this point you should have a single file SSN-YYYY-MM-DD.xlsx  
  The time segment files have been moved to subdirectory \Processed SSN Downloads</em>

<h2>Download Last Gasp Report</h2>

- Click on <b>Query</b>
- Select <b>Last Gasp Daily</b>
- Click on <b>Submit</b>
- Popup will tell you how many records were found (Defaults to all records.)
- Click on <b>Submit</b>
- Records will be loaded into Excel (progress should show in lower lefthand corner)
- "<b>Teradata Download Finished</b>" will apppear when completed.

<h2>Update Meter Status from SSN</h2>

- Click <b>SSNMeterStatus</b>  
  <em>The <b>RunDate</b> (leftmost colum) is used to determine the corresponding SSN file</em>
- "<b>SSN Meter Status Finished.</b>" will appear when done.  
  The column header "src_admin_state" will be highlighted in blue.
  
<h2>Collect Disconnected Meters</h2>

- Click on <b>Disconnected</b>
- The number of Disconnected meters will be shown.
- Click <b>OK</b>

  <em>A new <b>Disconnected</b> tab will appear</em>

<h2>Create the Proximity Tab</h2>

- Click on <b>PromximityColumns</b>  
  A new <b>Proximity</b> tabe will appear.
  
<h2>Process Meters</h2>

- Click on "<b>LG_NEXT</b>" and "<b>LG_BACK" to step through the meters
- Click on "<b>LG_KEEP</b>" to save a meter to the <b>Keep</b> and <b>Ticket</b> tabs

  
  
<h2>
