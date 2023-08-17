# source-doc-generator

**NOTE: DO NOT SHARE THIS PROJECT WITH ANYONE OUTSIDE THE TEAM. IF, IN ANY CASE YOU WANT TO SHARE IT, PLEASE MAKE SURE TO DELETE THE `credentials.json` FILE IN THE FOLDER BEFORE SHARING.**

This bot will help the writers create a source file for any script written in Sheets and place the newly created sources doc in the Sources folder in the shared Google Drive (as long as the conditions are met).

___

## How to use
* You will require 2 things:
  * **Sheet name:** Name of the sheet version. (Usually, it is `Script vX`). Make sure to give the final version of the Sheet.
  * **Sheet ID:** You can copy this from the URL. For eg., the Bold part is the Sheet ID (Do not include `/`).
    * docs.google.com/spreadsheets/d/ _**1zwPujGeAQALu20sfTxQP5oRCb8OS4sCgQEuR1rtHEbU**_ /edit#gid=94820894
  
* After you have the 2 things, run the `source doc generator.exe` file.
  * Ignore all the system warning about how unsafe the app is.
  * (Your data is safe with me 😇. JK. I don't have access to any of your data 😝).
* It will open a command prompt and also a browser window prompting you to sign in using your Google account.
  * ![image](https://github.com/ameymane09/source-doc-generator/assets/71630371/f5acb772-3331-4b61-a6f0-5cc948d85f98)
* Click continue.
  * ![image](https://github.com/ameymane09/source-doc-generator/assets/71630371/198bb716-d33c-41be-b79d-abdbb1070d4c)
* Give all the permissions.
  * ![image](https://github.com/ameymane09/source-doc-generator/assets/71630371/1af99f9b-1376-4022-90a2-02ce0f54c7db)

* After the authentication is complete, you should be greeted with the following message.
  * ```The authentication flow has completed. You may close this window.```
* The terminal will ask you for the sheet name and sheet id. Enter the values and after a few seconds, the doc should appear in the Sources Folder. (If not, try refreshing the Sources Folder).

___

## Things to keep in mind while writing the script
I've designed the bot to accomodate your current writing style. You do not need to make any changes to it, but please make sure you stick to the below points (which you already follow anyway).
* Use the same format across all writers. (Duh! ChatGPT jaisa bot nahi hai mera 😔)
* Cell A1 Should be the final title of the video in Bold.
* The section headings should always be in Bold.
* Don't leave blank cells in between the script. Use <> symbol to represent blank line.
* All the sources should be in the first column only. 2nd column and onwards should be used for animation cues, comments, etc.
* The lines ommitted/not recorded/deleted from the final version of the script must be in strikethrough.
* A link must not be randomly pasted in the cell. It should always be a hyperlink (i.e. text attached to that link.)
* The heading cell should not contain a hyperlink.
* Nothing other than the headings should be in Bold.

___

## Known Bugs (in v1)
* **Missing Sections**:
  * Sometimes the sources doc will be missing sections.
  * This happens due to the heading cell not being cleared of previous link formatting.
  * Basically, the hyperlinks are not cleared from the cells (even if they are not present) and the bot thinks its a link instead of a section heading.
  * Quick Fix (takes less than 1 min to apply):
    * Select all headings.
    * Press `Ctrl`+`\`. This will clear all the formatting from the cell. This does not clear the text.
    * Press `Ctrl`+`b` to format all the headings in Bold.
    * Press `Alt`+`o`+`a`+`c` to center align the text again.
  * Run the bot again after applying the quick fix, should work :)
 
* **Incomplete sources (not a bug)**:
   * In my testing, I found that this happens only when there is a blank row in between the lines.
   * Delete the row and run the bot again, the fix should work.
 
___

Enjoy! Use your saved time to treat me :)
