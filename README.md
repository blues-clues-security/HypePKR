# HypePKR
"Hype Packer" AKA Hyper-Kinetic Position Reverser is a vBA script to reformat and coalesce Nessus output into a more readable and succinct format  
### Intro  
Currently, the way this is used is by copy and pasting the code into a macro enabled workbook and turning it into a .xlsm .  
You'll notice in the "Mod1_6" text there is a note for '''Sheet 1''' and '''Start Workbook'''. Paste the code accordingly, and run each function in the first sheet. This will generate the buttons as well as the extra sheet required to generate hostname information.  
### How-To
To run HypePKR after the buttons have been generated, press start on the first sheet of the workbook. Afterwards, you'll be asked whether you filled out the "IP List" worksheet. This is just to remind you to paste your hostname information in the correct column before moving forward. After pressing "Yes" next you'll need to insert the file path to where your Nessus CSV files are located.  
  
*** _NOTE_ ***   
Any .csv files within the same directory will attempt to be ingested so keep that in mind!  
  
  
After entering the file path, Excel will open each .csv and paste the information into the first worksheet in order to manipulate the data. If there are any IP's that do not have an associated hostname, the script will put "[IP]???" into the hostname column.  
  
### Closing Remarks
If you have any suggestions please let me know! My next plan with this is to make the script accessible via a .vbs file to eliminate having to copy and paste code for first time use. 
