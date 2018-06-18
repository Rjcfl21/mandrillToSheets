# mandrillToSheets

easy to configure python3 script to grab email stats from Mandrill and push to Google Sheets.

## dependencies
* `config`
* `gspread`
* `mandrill`
* `google_api_python_client`
* `pandas`

## usage
First, the `config.json` and `config.py` files need to be edited.
The JSON file contains all of the configuration for Mandrill, like an APIKEY.
For my particular use-case, I had to split different groups of tags into different sheets - so that functionality is there under campaigns.
Fill in the tags you need reporting on and it will dump the stats to that particular sheet.
The only stats the script reports on are 'Sends', 'Opens', and 'Clicks'.


In `config.py`, you will need to enter all of the service account information, mainly the public key.
I use a dictionary to store this as I kept on forgetting to copy the key across..
If you get stuck with the service account set-up, [gspread](http://gspread.readthedocs.io/en/latest/oauth2.html) has some good documentation on it.

Once all the set-up is complete, it can be run as such:
`python3 mandrillToSheets.py --campaign campaign1`
And your sheet should populate within a minute or so with the stats.

the two command line options are `--campaign` which is required, and `--today` which is optional. `--today` will grab only Today's email stats.