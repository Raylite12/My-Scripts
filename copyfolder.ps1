Start-Transcript -Path "\\corp.nsa.gov\NCMD\Home\vol32031\ldraye\Private\Larry\Script\notepad\Test.txt" -NoClobber

#Copy folder
#Copy-Item -Path "c: the folder you want copy" -Destination "c: destination" -Recurse

Start-Process -FilePath "notepad++" New-Item 

#Create a directory/Folder
New-Item -Path "\\corp.nsa.gov\NCMD\Home\vol32031\ldraye\Private\Larry\Script\notepad" -ItemType Directory
Stop-Transcript