/**

* Recap Sheet Synchronization
*
* This script synchronizes dynamically fetched datasets
* into structured recap or reporting sheets in Google Sheets.
*
* The automation handles datasets whose row counts change
* depending on query results by inserting rows and copying
* values while preserving formulas in recap sheets.
*
* Typical workflow:
*
* BI Platform / API
* → Google Sheets (raw dataset)
* → Apps Script synchronization
* → Recap / reporting sheet
*
* This helps automate reporting pipelines and removes the
* need for manual copy-paste operations.
  */
