This script reads the rows of a Google sheet and generates HTML pages showing thumbnails of objects with pop-ups describing the object.

This script is hardcoded on sheet, column names and picture IDs; it is not a general purpose script.  It is used by a museum to generate HTML pages for objects on a per-room basis.  The Google Sheet is used as a collections object database, where each row of the sheet has fields about the object.

There are 3 parameters, with the first 2 are required.  The default logging is WARNING; with the default logging on, there a Google client library API warnings - not sure why.

`--file_id <your-sheet-id> --team_drive_id <your-team-drive-id> --loglevel="DEBUG"`

Refer to [this](https://stackoverflow.com/questions/54300175/how-to-get-teamdriveid-when-use-google-drive-api) on how to obtain the team-drive-id

Libraries that need to be installed to run build_html.py:
* The [Google python client library](https://developers.google.com/sheets/api/quickstart/python).  Another reference is the [Google drive quickstart](https://developers.google.com/drive/api/v3/quickstart/python).
* [Dominate](https://github.com/Knio/dominate): creates HTML docs using a DOM API

You will need to create a credentials.json file, as per the instructions in the python library quickstart.

The first time build_html is run, a browser window will open, you'll go through the google sign-in process, and you will be asked to give permissions to drive-access-1 to access your sheets/drive; the code will create a token.pickle file. Subsequent runs will use token.pickle file.
