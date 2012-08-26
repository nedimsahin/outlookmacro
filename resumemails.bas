Attribute VB_Name = "Module1"
' Outlook Macro for moving specific mails and reading their content
' Version: 1.0 (VBA version: 7.0)
' Date: 8/26/2012
' Author: Nedim Sahin
' Web site: www.birvesifir.com
' Contact: nedim@nedimsahin.net
'
' Summary: What does this macro do? (Step by step)
'  - Find the mails which have the specific subject in Inbox folder
'    (In this case, subject is "New resume has been received!")
'  - Creating a folder named like "Resumes 8/26/2012 1:29:50 PM" under Inbox.
'  - Moving those mails to "Inbox > Resumes 8/26/2012 1:29:50 PM"
'  - Creating a folder named like "Resumes 8/26/2012 1:29:50 PM" under C: drive
'  - Saving those mails' attachments under "C:\Resumes 8/26/2012 1:29:50 PM"
'  - Reading content of those mails and creating Excel file
'
' Notes:
'  - A reference named "Microsoft Excel 14.0 Object Library" must be added
'  - It searchs mails under only Inbox folder
'
' Special thanks:
'  - http://www.techonthenet.com/excel


Sub resumemails() ' Main subroutine

    ' Declare all variables.
    Dim objOutlook As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objSourceFolder As Outlook.MAPIFolder
    Dim objDestFolder As Outlook.MAPIFolder
    Dim objVariant As Variant
    Dim lngMovedMailItems As Long
    Dim intCount As Integer
    Dim title As String
    Dim strDestFolder As String
    Dim folderName As String
    Dim folderNameCdrive As String
    Dim fs As Object
    Dim Item As Object
    Dim Atmt As Attachment
    Dim FileName As String
    Dim temp As String
    Dim xlApp As Excel.Application
    Dim xlWB As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim olItem As Outlook.MailItem
    Dim vText As Variant
    Dim sText As String
    Dim vItem As Variant
    Dim i As Long
    Dim rCount As Long
    
    ' Setting first values
    Set objOutlook = Application
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objSourceFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    
    ' Creating and setting folders
    folderName = "Resumes " & DateTime.Now()
    Set objDestFolder = objSourceFolder.Folders.Add(folderName)
    Set fs = CreateObject("Scripting.FileSystemObject")
    folderNameCdrive = "C:\" & "Resumes " & Replace(DateTime.Now(), ":", "-")
    folderNameCdrive = Replace(folderNameCdrive, "/", "-")
    fs.CreateFolder folderNameCdrive
                
    ' Move resume mails to "Inbox\Resumes DateTime"
    For intCount = objSourceFolder.Items.Count To 1 Step -1
        Set objVariant = objSourceFolder.Items.Item(intCount)
        DoEvents
        If objVariant.Class = olMail Then
            Debug.Print objVariant.SentOn
            title = objVariant.Subject
            If title = "New resume has been received!" Then
                objVariant.Move objDestFolder
                lngMovedMailItems = lngMovedMailItems + 1
            End If
        End If
    Next
    
    ' Save attached resumes to "C:\Resume DateTime"
    For Each Item In objDestFolder.Items
        For Each Atmt In Item.Attachments
            If LCase(Right(Atmt.FileName, Len(ExtString))) = LCase(ExtString) Then
                FileName = folderNameCdrive & "\" & Atmt.FileName
                Atmt.SaveAsFile FileName
                FileName = Mid(FileName, 1, InStrRev(FileName, ".") - 1)
            End If
        Next Atmt
    Next Item
    
    ' Setting Excel values
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Add
    Set xlSheet = xlWB.Worksheets(1)
    rCount = rCount + 1
    
    ' Process each resume mail
    For Each olItem In objDestFolder.Items
        sText = olItem.Body
        vText = Split(sText, Chr(13))
        
        rCount = rCount + 1
    
        ' Columns names
        xlSheet.Range("A1") = "Full name"
        xlSheet.Range("B1") = "Email"
        xlSheet.Range("C1") = "Phone"
        xlSheet.Range("D1") = "Preferred time to contact"
        xlSheet.Range("E1") = "Visa status"
        xlSheet.Range("F1") = "Position"
        xlSheet.Range("G1") = "CV File Name"
    
        'Check each line of text in the message body
        For i = UBound(vText) To 0 Step -1
            If InStr(1, vText(i), "Full name:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("A" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
    
            If InStr(1, vText(i), "Email address:") > 0 Then
                temp2 = InStrRev(vText(i), ":")
                temp3 = InStr(temp2 + 1, vText(i), Chr(34))
                If temp3 < temp2 Then ' If there is a link, this "If" statement handles it
                    vItem = Split(vText(i), Chr(58))
                Else
                    vItem(1) = Mid(vText(i), temp2 + 1, temp3 - temp2 - 1)
                End If
                xlSheet.Range("B" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
    
            If InStr(1, vText(i), "Phone number:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("C" & rCount).NumberFormat = "@"
                xlSheet.Range("C" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
            
            If InStr(1, vText(i), "Preferred time to contact you:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("D" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
            
            If InStr(1, vText(i), "Visa status:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("E" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
            
            If InStr(1, vText(i), "Position you are applying for:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("F" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
            
            If InStr(1, vText(i), "Upload your CV:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("G" & rCount) = LTrim(cleanIt(CStr(vItem(1))))
            End If
        Next i
    Next olItem
    
    ' Beautification
    xlSheet.Columns("A:G").AutoFit
    xlSheet.Range("A1:G1").Interior.Color = RGB(30, 144, 255)
    xlSheet.Range("A1:G1").Font.Color = RGB(248, 248, 255)
    
    ' Relasing the object
    Set objDestFolder = Nothing
    
End Sub

Function cleanIt(text As String) As String ' Clean spaces and unprintable characters

    Dim temp1 As String
    Dim temp2 As Integer
    
    temp1 = Trim(text)
    
    If temp1 <> "" Then
        temp2 = Asc(Left(temp1, 1))
        If temp2 < 33 Or temp2 = 160 Then
            temp1 = Right(temp1, Len(temp1) - 1)
        End If
    End If
    
    cleanIt = temp1
    
End Function

