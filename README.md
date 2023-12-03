# Custom Slack emojis to Excel
A simple script written to help going through lots of custom emojis in Slack workspaces.

## Description
My friends and I share a Slack workspace, that is about 7 years old. Before that, for about 15 years, we have been migrating through different messaging platforms and collecting custom emojis as we go.
Old emojis have been transferred from old platforms to new ones and that resulted us having about 500 custom emojis, of which some haven't been used in years, some have doubles, some could use better names..
I created this simple script to automatically create and Excel file of all the custom emojis so we could go through them efficiently.
Hope this is of use to you!

I am by no means a programmer or a software developer, nor do I aim to become one. 
I created this script for fun, so it has zero to none error handling. Hope it is still of use to you!

## How-to

### Dependencies
* Python 3.7 or higher
* Slack app token, both User and Bot tokens seem to work.
* Probably need Excel installed, not sure if it works with alternative solutions like OpenOffice.

### Installing
* You need to edit emojis.py and change TOKENHERE toyour Slack app token.
	* You could also modify the script to use a presaved emoji.json, which are accessible to non-admins as well (according to Google.)
* Make sure you have downloaded the following modules: requests, xlsxwriter, Pillow, DateTime

### Executing program
* Save emojis.py to a location of your choice.
	* For me, I created a separate directory on my desktop for this, just for clarity!
* From where you saved emojis.py, run CMD or navigate CMD to the save location
* Start the program with command: python emojis.py
* Choose whether you want to
	* 1) Save all custom emojis as images, with emoji names as filenames.
	* 2) Save all emoji names and images to excel.
	* 3) Exit :)

## Help
It is likely you get an error, as there is no error handling. Sorry about that!

* Make sure your Slack app token has emoji.read permissions. See Slack API site for more info!
* If you get an error while saving emojis to Excel, it is likely a temporary directory was created but not deleted. Temporary directories are always named tempemojis and created to same directory where emojis.py is, so make sure to delete it manually before trying again.

## Authors
Me! Please don't contact me though.

## Version history
* 0.1
	* Initial release.
	
## License
This project is licensed under the terms of the MIT license - see the LICENSE.md file for details.
