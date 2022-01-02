# FFXIV Sheet Gear Tracker
A script for Google Sheets designed streamline the raid gear distribution process.

## How to Set up
1) Create a new Google Sheet and name the tab "BiS".
2) Go to Extensions > Apps Script.
3) Click the plus by `Files` and create a script called `Gear`. Fill it in with the contents of the `Gear.gs` file in this repository under `scripts/`.
4) Click the plus icon again, choosing `HTML` this time, and create two files: `NewPlayer` and `RemovePlayer` filling both in with the contents of the identically named files in this repository under `html`.
5) Reload your spreadsheet a new option at the top should appear called "Manage Players". Fill out the form and all the players from your static. The character is verified against xivapi.com, so do be patient after hitting "Submit".

## How to Use
Each player has an Etro link they can optionally place into the `Etro Link` section. If they choose to do so, all the gear specified in the Etro gear set will be downloaded and parsed into whether the character needs a Coffer, Twine, Coating, or crafted item for that gear slot. Removing the Etro link will not remove the items from the `Needs` column.

Once a player has looted something, they can check the box relating to the piece of gear they got and which slot they're using it on.

At the bottom of the sheet is the "Drops Still Needed" table which will automatically tally the number of twines and coatings a player needs to reach BiS. Once a Twine or Coating has been checked off the tallies will automatically update with how many they still need.