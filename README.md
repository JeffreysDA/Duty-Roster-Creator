# Duty Roster Creator
Fairly distributes Duty position responsibilities. Originally created for Marine Corps’ 8th Comm Bn Camp on Lejeune, but updated to hopefully work for most any military branch and personnel size.



### Summary
The Duty Roster Creator is comprised of several excel files that work together and are dependent upon each other. It was started as an attempt to minimize human error, and ensure fair distribution of duties for not only each company, but also each section within those companies.

The main document is the Able Body Report Roster, which imports and stores information from three separate personnel rosters. The information from these combined rosters identifies who can stand duty, and how many of each duty position to give each company and each of their sections while ensuring that no section or company is unnecessarily over or under tasked. This information is then exported to a new Duty Roster for the selected month and year where holidays and weekends are automatically highlighted. Said Duty Roster is then duplicated and renamed according to each company, then merged at a later date to create a complete battalion Duty Roster.



### Compatibility:
Microsoft Office 2013 and newer. *(Untested, but might work with older versions)*



### Required Files
The Primary file is the **Able Body Report Roster.xlsm** (ABRR). The ABRR is the powerhouse behind the other files and is the bulk of this tool.
It requires the **2_TEMPLATES** folder and all of its contents to properly work. The files can work from almost any location, but may have issues 
if used from a network drive (Share drive, OneDrive, etc.). Aside from ensuring that you have the ABRR, and the **2_TEMPLATES** folder (with all of its contents)
the only other requirement is that the file structure remain the same. This just means that both the ABRR and the **2_TEMPLATES** folder need to be located in the same spot.
This also means that the files cannot be renamed.

*If changes need to be made to the Duty Roster Creator. Use the **Developer Mode** feature to make your changes.*



### Not Required, but Useful
The ABRR imports documents to better determine how to do its calculations and help ensure it is as accurate as possible. Examples of these files can be found in the
**1_EXAMPLES** folder. The **1_EXAMPLES** folder is not required in order for the Duty Roster Creator to work, but rather to give the user an idea of what the imported documents
should look like, how they should be named, and how they should be structured internally. Likewise, the instruction manuals for the ABRR and the Duty Roster are not
required, but will help give the users a better understanding of each function and feature, and how to use it to its full capabilities.



### Developer Mode
To improve versatility, I have implemented a **Developer Mode** into this Macro Enabled Excel document.

Developer Mode allows a user to safely make certain edits to the document such as changing the options in a dropdown list, the picture associated with a worksheet, the password used to lock and protect the document, and more. To enter Developer Mode, you must start by opening the document, enabling Macros, and clicking on any unlocked/editable cell. Once you have done that simply press *Ctrl*+*Shift*+*Alt*+*D*.

If done correctly, a window will appear informing you that you are about to enter Developer Mode. While in this mode, you will be able to safely make changed to certain aspects of the document. If further or more in-depth changes need to be made, please contact me. See my bio for the appropriate contact information.



### Making edits beyond what can be made in Developer Mode

Edits beyond what is possible while in Developer Mode can be made by unlocking the document entirely with the password given in the files Developer Mode sheet, or by entering Developer Mode, and while on the **DevMode** sheet pressing *Ctrl*+*Shift*+*~*.
**I STRONGLY URGE YOU TO USE CAUTION IF YOU MAKE THESE KINDS OF CHANGES. ONLY DO SO IF YOU KNOW EXACTLY WHAT YOU ARE DOING.**

It is important to note that changes beyond what can be made while in Developer Mode are very likely to render the file broken as there are hard coded references to cells, sheets, and other aspects of the document that can be broken if changes are made and the VBA code is not updated to reflect those changes. That being said, if you understand how to read and write VBA macros, and you are proficient in doing so, then you are more than welcome to attempt to do this on your own, but at your own risk of breaking the file (and possibly other associated files) by doing so.
