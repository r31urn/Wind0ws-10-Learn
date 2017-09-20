'This script may or may not work'
'No assurance is provided'
'Use it at your own risk'
'Copy and ditribute the code, link the credits to my repo'


Option Explicit 

Dim strComputer, objWMIService, objItem, Caption, colItems
'Create wscript.shell object 
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)
For Each objItem in colItems
    Caption = objItem.Caption  
Next

If InStr(Caption,"Microsoft Windows 7") > 0 Or InStr(Caption,"Microsoft Windows 8") > 0 Or InStr(Caption,"Microsoft Windows 10")Then 
	Dim objshell,path,DigitalID,Path2,DigitalIDIE, Result 
	Set objshell = CreateObject("WScript.Shell")
	'Set registry key path
	Path = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\"
	'Registry key value
	DigitalID = objshell.RegRead(Path & "DigitalProductId")
	Dim ProductName,ProductID,ProductKey,ProductData,Build,Release,ProductKeyIE
	'Get ProductName, ProductID, ProductKey
	ProductName = "Product Name: " & objshell.RegRead(Path & "ProductName")
	ProductID = "Product ID: " & objshell.RegRead(Path & "ProductID")
	ProductKey = "Installed Key: " & ConvertToKey(DigitalID)
	Build = "Build: " & objshell.RegRead(Path & "BuildBranch") 
	Release ="Release:" & objshell.RegRead(Path & "BuildLabEx")
	Path2 ="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Registration\"
	DigitalIDIE = objshell.RegRead(Path2 & "DigitalProductId")
	ProductKeyIE = "IE KEY: " & ConvertToKey(DigitalIDIE)
	ProductData = ProductName  & vbNewLine & ProductID  & vbNewLine & Build  & vbNewLine & Release  & vbNewLine & ProductKey & vbNewLine & ProductKeyIE
	'Show messbox if save to a file 	
	If vbYes = MsgBox(ProductData  & vblf & vblf & "Save to a file?", vbYesNo + vbQuestion, "BackUp Windows Key Information") then
	   Save ProductData 
	End If
	
Else
	MsgBox "Please run this script in Windows 8.x or Windows 10"	
End If

'Convert binary to chars
Function ConvertToKey(Key)
    Const KeyOffset = 52
    Dim isWin8, Maps, i, j, Current, KeyOutput, Last, keypart1, insert
    'Check if OS is Windows 8
    isWin8 = (Key(66) \ 6) And 1
    Key(66) = (Key(66) And &HF7) Or ((isWin8 And 2) * 4)
    i = 24
    Maps = "BCDFGHJKMPQRTVWXY2346789"
    Do
       	Current= 0
        j = 14
        Do
           Current = Current* 256
           Current = Key(j + KeyOffset) + Current
           Key(j + KeyOffset) = (Current \ 24)
           Current=Current Mod 24
            j = j -1
        Loop While j >= 0
        i = i -1
        KeyOutput = Mid(Maps,Current+ 1, 1) & KeyOutput
        Last = Current
    Loop While i >= 0 
    keypart1 = Mid(KeyOutput, 2, Last)
    insert = "N"
    KeyOutput = Replace(KeyOutput, keypart1, keypart1 & insert, 2, 1, 0)
    If Last = 0 Then KeyOutput = insert & KeyOutput
    ConvertToKey = Mid(KeyOutput, 1, 5) & "-" & Mid(KeyOutput, 6, 5) & "-" & Mid(KeyOutput, 11, 5) & "-" & Mid(KeyOutput, 16, 5) & "-" & Mid(KeyOutput, 21, 5)
   
    
End Function
'Save data to a file
Function Save(Data)
    Dim fso, fName, txt,objshell,UserName
    Set objshell = CreateObject("wscript.shell")
    'Get current user name 
    UserName = objshell.ExpandEnvironmentStrings("%UserName%") 
    'Create a text file on desktop 
    fName = "C:\Users\" & UserName & "\Desktop\WindowsKeyInfo.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set txt = fso.CreateTextFile(fName)
    txt.Writeline Data
    txt.Close
End Function
