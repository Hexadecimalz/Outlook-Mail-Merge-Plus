' Outlook Mail Merge Attachment & Change Sender 
' 
' This script adds an attachment or changes the sender to all emails
' in the Microsoft Office Outlook outbox. The script is tested with 
' Microsoft Outlook 2003, 2007, 2010, 2013, and 2016.  
'
' The emails are modified by passing keystrokes. Please do not touch the keyboard or mouse while in process.
'
' This script has been modified to include a portion to change the sender to a delegate or generic mailbox
' The publication of a modification of the code has been published with permission from the original author. 
'
' Modification by Hexadecimalz: https://github.com/Hexadecimalz 
'
' For more information, visit http://omma.sourceforge.net or contact
' westerveld@users.sourceforge.net.
'
' Version 1.2.0 Beta, 19 October 2015 
'
' Copyright (C) 2006-2015 Wouter Westerveld
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
' 
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
' 
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'

SubOutlookMailMergeAttachment

Sub SubOutlookMailMergeAttachment		
	' Script version
	strProgamName = "Outlook Mail Merge Attachment (v1.1.9 Beta)"
	strProgamVersion = "Outlook Mail Merge Attachment (v1.1.9 Beta)"	
	
	' Set manual line-breaks in message box texts for Windows versions < 6.
	strBoxCr = vbCrLf
	On Error Resume Next
	Set SystemSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem") 	
	For each System in SystemSet 				
		If System.Version >= 6 Then		
			strBoxCr = ""
		End If
		sWindowsVersion = System.Caption
	Next 
	On Error Goto 0
		
	' Welcome dialog
	strDialog = "This script adds an attachment and/or changes the sender of all emails that are currently in the Microsoft Office Outlook outbox. " & strBoxCr & _
	    "The script is tested with Microsoft Outlook 2003, 2007, 2010, 2013, and 2016." & vbCrLf & _
		"" & vbCrLf & _
		"Usage:" & vbCrLf & _
		"1.  Put Outlook in OFFLINE Mode (We can start this in the next step), then send your mail merge." & vbCrLf & _
		"2.  Execute (Double-Click) 'Outlook Mail Merge Attachment.vbs',"  & vbCrLf & _
		"3.  Run through the script where you will be prompted to change the sender or add attachments." & vbCrLf & _
		"4.  After your input the script will change the sender and/or add attachments on your mail merge (all items in the Outbox)" & vbCrLf & _
		"5.  You may preview the changes in your outbox. If it worked, then put Outlook into online mode. Leave the keyboard and mouse alone." & vbCrLf & _
		"" & vbCrLf & _			
		"The emails are sent by passing keystrokes. Please do not touch the keyboard or mouse while in " & strBoxCr & _ 
		"process." & vbCrLf & _		
		"" & vbCrLf & _
		"Do you want to continue?" & vbCrLf & _		
		"" & vbCrLf & _		
    	"http://omma.sourceforge.net" & vbCrLf & _	
    	"westerveld@users.sourceforge.net (C) 2006-2015" & vbCrLf & _			
		"" & vbCrLf & _			
		"Modification by Hexadecimalz: https://github.com/Hexadecimalz " 
	'''''''''''''''''''''''''''''''''''''''''''''''
	' Initialize, load objects, check
	'''''''''''''''''''''''''''''''''''''''''''''''
	
    If MsgBox(strDialog, vbOKCancel + vbInformation, strProgamName) = vbCancel Then
        ' fout  
	    Exit Sub                  
    End If     
        
    ' Outlook and Word Constants
    intFolderOutbox = 4
    msoFileDialogOpen = 1
    
    ' Load requied objects
    Set WshShell = WScript.CreateObject("WScript.Shell")	' Windows Shell
    Set ObjWord = CreateObject("Word.Application")      ' File Open dialog    
    Set ObjOlApp = CreateObject("Outlook.Application")      ' Outlook
    Set ns = ObjOlApp.GetNamespace("MAPI")                  ' Outlook
    Set box = ns.GetDefaultFolder(intFolderOutbox)          ' Outlook   

    ' Check if we can detect problems in the outlook configuration
    sProblems = ""    
    sBuild = Left(ObjOlApp.Version, InStr(1, ObjOlApp.Version, ".") + 1)
    
    ' check spelling check just before sending
    On Error Resume Next
    r = WshShell.RegRead("HKCU\Software\Microsoft\Office\" & sBuild & "\Outlook\Options\Spelling\Check")    
    If Not(Err) And (r = 1) Then
    	sProblems = sProblems & _    	
    	"Your Outlook spell check is configured such that it gives a pop-up box when sending emails. Please disable " & strBoxCr & _
    	"the 'Always check spelling before sending' option in your Outlook. (ErrorCode = 101)" & vbCrLf &vbCrLf
    End If
    On Error Goto 0
    
	' For outlook 2000, 2002, 2003
	If sBuild = "9.0" Or sBuild = "10.0" Or sBuild = "11.0" Then
	
	    ' Check for word as email editor.
	    On Error Resume Next
		intEditorPrefs = WshShell.RegRead("HKCU\Software\Microsoft\Office\" & sBuild & "\Outlook\Options\Mail\EditorPreference")		
		If Not(Err) Then
			If intEditorPrefs = 131073 Or intEditorPrefs = 196609 Or intEditorPrefs = 65537 Then
				' HTML = 131072, HTML & Word To Edit = 131073, Rich Text = 196610, Rich Text & Word To Edit = 196609, Plain Text = 65536, Plain Text & Word To Edit = 65537			
				sProblems = sProblems & _			
				"Your Outlook is configured to use Word as email editor. Please change this to the internal outlook editor in " & strBoxCr & _
				"your outlook settings. (ErrorCode = 102)" & vbCrLf &vbCrLf				
			End If
		End If		
		On Error Goto 0
	End If

	If sProblems <> "" Then				    
		sProblems = "The OMMA script detected settings in your Outlook settings that need to be changed for the software to work." & vbCrLf & vbCrLf & sProblems
		MsgBox 	sProblems, vbExclamation, strProgamName			
		'fout
		Exit Sub
	End If
  
    ' Check if there are messages
    If box.Items.Count = 0 Then
        MsgBox "There are no messages in the Outbox.", vbExclamation, strProgamName           
       	' fout
       Exit Sub
    End If
    
    ' Give a warning if there already is an attachment
    If box.Items(1).Attachments.Count > 0 Then
        If MsgBox("The first email in your outbox has already " & box.Items(1).Attachments.Count & " attachment(s). Do you want to continue?", vbOKCancel + vbQuestion, strProgamName) = vbCancel Then
            ' fout  
		    Exit Sub            
        End If
    End If
	
	' Do you want to change the sender (i.e. for delegates or Generic Mailbox)    
    intChangeSender = _
    Msgbox("Would you like to change the sender of the e-mail?", _
        vbYesNo)
		
	'If the User said yes, then set the var for changing, and if not proceed.	
    If intChangeSender = vbYes Then
		strUserEmail = InputBox( "Enter the e-mail you want to send with: " )
    Else
		Msgbox "Okay, the message will be sent with the default e-mail."
    End If
	
	' Add CC field 
	intCCfield = _
    Msgbox("Would you like to CC a contact?", _
        vbYesNo)
		
	'If the User said yes, then set the var for changing, and if not proceed.	
    If intCCfield = vbYes Then
		strUserCC = InputBox( "Enter the e-mail you want to CC: " )
    Else
		Msgbox "Okay, no CC is added."
    End If
	
	' Do you REALLY want to add an attachment? If not, that is okay. 
	
    intAttachment = _
    Msgbox("Would you like to add an attachment to this mail merge?", _
        vbYesNo)
	
    '''''''''''''''''''''''''''''''''''''''''''''''
    ' Ask user for Filenames, add attachment, and 
    ' Add attachment and save email
    '''''''''''''''''''''''''''''''''''''''''''''''         
    ' Ask user to open a file
    ' Select the attachment filename 
	
	if intAttachment = vbYes Then 
 
	ObjWord.ChangeFileOpenDirectory(CreateObject("Wscript.Shell").SpecialFolders("Desktop"))	
	ObjWord.FileDialog(msoFileDialogOpen).Title = "Attach file(s)..."
	ObjWord.FileDialog(msoFileDialogOpen).AllowMultiSelect = True
	
	okEscape = False	
	If ObjWord.FileDialog(1).Show = -1 Then
		If ObjWord.FileDialog(1).SelectedItems.Count > 0 Then		
			okEscape = True
		End If 
	End If 
	
	If Not okEscape Then
		ObjWord.Quit
		MsgBox "Cancel was pressed, no attachments where added.", vbExclamation, strProgamName
		Exit Sub   	
	End If 
	
	End If 
	
    WScript.Sleep(800)               
        
    ' Add the attachment to each email
    For Each Item In box.Items        
    	For Each objFile in ObjWord.FileDialog(1).SelectedItems
			If intAttachment = vbYes Then
			Item.Attachments.Add(objFile)
			End If 
			If intChangeSender = vbYes then 
            Item.SentOnBehalfOfName = strUserEmail
			End If
			If intCCfield = vbYes then
			Item.CC = strUserCC
			End If 
        Next             
        Item.Save
    Next 

	ObjWord.Quit
 	
 	'''''''''''''''''''''''''''''''''''''''''''''''
 	' Send the emails using keystrokes
 	'''''''''''''''''''''''''''''''''''''''''''''''
 	
    For i = 1 to box.Items.Count
        
        ' Wait 5 extra seconds after 50 emails
        If (i Mod 50) = 0 Then
    		WScript.Sleep(5000)    	
        End If
        
        ' Open email
        Set objItem = box.Items(i)
		Set objInspector = objItem.GetInspector
		objInspector.Activate		
		WshShell.AppActivate(objInspector.Caption)		
		objInspector.Activate
	
		' wait upto 10 seconds until the window has focus		
		okEscape = False
		For j = 1 To 100
			WScript.Sleep(100)
			If (objInspector Is ObjOlApp.ActiveWindow) Then
				okEscape = True
				Exit For
			End	If
		Next
		If Not(okEscape) Then			        		
	        MsgBox "Internal error while opening email in outbox. Please read the how-to and the troubleshooting sections in the " & strBoxCr & "documentation. (ErrorCode = 103)", vbError, strProgamName
	       ' fout
	       Exit Sub			
		End If
		
		' send te email by typing ALT+S
		WshShell.SendKeys("%S")
						
		' wait upto 10 seconds for the sending to complete
		okEscape = False
		For j = 1 To 100
			WScript.Sleep(100)
			boolSent = False
			On Error Resume Next
			boolSent = objItem.Sent
			If Err Then
				boolSent = True
			End	If
			On Error Goto 0
			If boolSent Then
				okEscape = True
				Exit For
			End	If
		Next						
		If Not(okEscape) Then					
			' Error			       
	        MsgBox "Internal error while sending email. Perhaps the email window was not activated. Please read the how-to and " & strBoxCr & "the troubleshooting sections in the documentation. (ErrorCode = 104)", vbExclamation, strProgamName
	       ' fout
	       Exit Sub						
		End If
    Next 
 
    ' Finished    
    strDialog = "Successfully added the attachment and/or changed sender to " & box.Items.Count & " emails." & vbCrLf & vbCrLf & _    	
    	"OMMA is free software, please let the author know whether OMMA worked properly. " &strBoxCr & _
    	"Did you already fill the feedback form?" & vbCrLf & vbCrLf & _ 
    	"Answer 'No' will open the feedback form in your browser."  & vbCrLf & _  
    	"Answer 'Yes' just exit the script." 
    	
    If MsgBox(strDialog, vbYesNo + vbInformation, strProgamName) = vbNo Then
		WshShell.Run "http://omma.sourceforge.net/feedback.php?worksok=yes&verOmma=" & escape(strProgamVersion) & "&verWindows=" & escape(sWindowsVersion) & "&verOutlook=" & escape(sBuild)
    End If         
    
End Sub