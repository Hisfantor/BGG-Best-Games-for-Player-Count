# BGG-Best-Games-for-Player-Count
a script for google sheets to import your ownd games form [Board Game Geek API](https://boardgamegeek.com/xmlapi2) and sort them by how good they play for a given Player Count

## Options:
- **Username**, add your [bgg](https://boardgamegeek.com/) username here
- **include Expansions** can be set to yes or no
- **Player Count** allows any number of players
- **Sort by** lets you sort the resulting games by the table headers(name, published, rating, playtime, weight)
- **Sort all regardles of playability** can be set to yes or no, standard is no which shows you three seperate areas of playability as seen here:

![all settings shown in a screenshot, Sort all regardles of playability set to no](/Sort_all_regardles_of_playability-no.png "Sort all regardles of playability: no")

if you set "**Sort all regardles of playability: yes**" it combines all the games for this player count in one list and sorts them like this:

![all settings shown in a screenshot, Sort all regardles of playability set to yes](/Sort_all_regardles_of_playability-yes.png "Sort all regardles of playability: yes")

"**Load Player Data**" also sorts the games by the currently assigned settings

"**Sort**" resorts the last loaded games (also works after reloading the page)

**weight** is color coded similar to BGG 1-2 green, 2-3 orange, 3-5 red
## Installation:
- open the [google sheets link](https://docs.google.com/spreadsheets/d/1Yz4JlLDtu8P97KRSHnSmc9yVc6uMFrvkScRSY4cZuLc/edit?usp=sharing) and make a copy of the file, this will also copy the Apps Script

or

- download the "BGG Best Games for Player Count.xlsx" file
- open a new spreadsheet in [google sheets](https://docs.google.com/spreadsheets) and import the downloaded file
- go to Extensions -> Apps Script, copy the code from "Code.gs" and save
- close the editor
- assign the scripts to the buttons by right clicking them, select the dots -> Assign a Script

	- Load Player Data = getUserGames
	- Sort = sortgames


when first using the scripts you'll be asked to give them Access to the sheet, you have to allow this for the script to work
