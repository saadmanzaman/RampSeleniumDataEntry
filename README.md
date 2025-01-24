# RampSeleniumDataEntry
 
Since Ramp doesn't have an out of the box API to update the Ramp fields, I created this to automate the task of updating hundreds of transactions.

This file is a snippet of a larger Excel addin file that is designed to update Ramp expense transactions data fields using selenium and chrome.
It cannot be run as is, and needs to be tweaked to fit the needs of the user.

Notes:
Ramp updates their UI often, breaking the xpath location.
This is only useful until Ramp develops a method to edit their transactions through their API.
