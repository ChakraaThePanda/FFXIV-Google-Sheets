# FFXIV Automatic Spreadsheet

### ⚠️ This project is no longer maintained ⚠️  
I (Chakraa) am taking a most likely definite break from FFXIV and am moving on to different games.  
You can fork this project if you want to work on it yourself, or contact Raven (Discord: RavenAmbree#7038) for support.  
Good luck to all my fellow Warriors of Light!  

---

<ins>**Created by**</ins>:  
Chakraa Arcana @ Famfrit.  
Discord: Chakraa#0001

<ins>**Maintainer(s)**</ins>:  
Raven Ambree @ Excalibur.  
Discord: RavenAmbree#7038

<ins>**Purpose**</ins>:  
Through the use of [XIVAPI](https://xivapi.com/), the goal is to display selected information of the
selected characters through their LodestoneID in a Google Sheets Project.  
The Sheets Project will be made available to copy by anyone and be used to manage their own team, FC, Linkshell, etc.

<ins>**Limitations**</ins>:  
Due to how the API works and Google's limitation on Spreadsheet scripts, the FC portion of the sheet will not work for FCs with a high member count, past i'd say roughly 280 members.
You can still track people individually in the regular Roster sheet, perfect for your favorite FC members!

<ins>**Where to Start/Update your Current Copy**</ins>:  
You can copy the blank template located below, and add your informations after following the little tutorial.
https://docs.google.com/spreadsheets/d/1c9wFqLBBDqp7QmtggCNuAqxAMyNzWQSCzVTHmz8lwj0/
Unfortunately, there isn't a good and practical way to update your current copy of the Spreadsheet with the up-to-date informations.  
If you want to have the latest version of the Spreadsheet, you will have to make a copy of the blank version listed just above and transfer your data over.

<ins>**Latest Github Release**</ins>:  
https://github.com/ChakraaThePanda/FFXIV-Google-Sheets/releases  
***Unless specified otherwise, even if the latest release doesn't reflect the current FFXIV Patch, it should still be working. It just means nothing breaking was added to the game since the last sheet update.***

## The Sheets

### Instructions
![Instructions](https://imgur.com/JskAW0g.png)
On this page you have the basic breakdown on how to setup the project properly for your personal use.  
The API key you will have received on [XIVAPI](https://xivapi.com/) will have to be entered in the **red box**.

### Roster
![Roster](https://imgur.com/55oIpUO.png)
On this page, you will have to enter the LodestoneID of every character that you want to track.  
If you are unsure of what the LodestodeID of the character you want to add is, [search for the character you want to add](https://na.finalfantasyxiv.com/lodestone/character/) and grab all the numbers in the address bar.  
Example: [https://na.finalfantasyxiv.com/lodestone/character/**20723319**/](https://na.finalfantasyxiv.com/lodestone/character/20723319/)  
Pressing the Black button in the top-left of the page will start the updating process.  
If you know what you are doing, you can add a Time-Trigger in Google Script to automate the whole process on a time interval.

### FC Roster
![FC Roster](https://imgur.com/lVsGg5s.png)
On this page, you will need your Free Company's LodestoneID.
You can search for your Free Company [here.](https://na.finalfantasyxiv.com/lodestone/freecompany/)  
Updating this page using the Black button in the top-left will add every character in your FC automatically and get their informations.

## Quick FAQ
* If I have a question, where should I ask it?
  * If you have any questions, I would suggest contacting Raven on Discord: RavenAmbree#7038
* The FC Sheet only pulls the IDs but doesnt show the information on my members, why?
  * Most likely your FC is too big for us to handle. The sheet works for FCs that have an amount of members ranging from 1 to more or less 280.
* I added my character in the list, but it seems to be outdated. How do I update it?
  * The spreadsheet will not update itself to have your character's info real time. The data we use on the sheet is taken directly on the Lodestone page of your character. If the Lodestone page is outdated, we also will be. Lodestone can take a couple hours to have all the correct informations
* I entered my character's LodestoneID and it cannot find my character.
  * If your character is brand new, it may take a little while for it to appear in the Lodestone database. Retry a bit later :smile:
* I have found an issue with the page/something doesn't seem to work.
  * Add Raven on Discord: RavenAmbree#7038. He will be happy to assist you :wink:
* Do you have a donation link at all?
  * Nah, no worries, its about the only little side project I have that im keeping up to date, so it doesnt take much of my time. I do it just for fun :smile:
