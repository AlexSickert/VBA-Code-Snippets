
Dim excApp As Object, _
    excWkb As Object, _
    excWks As Object, _
    intVersion As Integer, _
    intMessages As Integer, _
    lngRow As Long
    
Dim dict
Dim dictSender
Dim dictRecipient
Dim currentPath

' Program start function
Sub ProgStart()

    ' loop through the folders
    
    ' for each folder process the entries for dictionary
    ' for each mail save mail in a folder based on time
    ' makeFoldes(basePath, email, sendDate)
    
    Dim strFilename As String, olkSto As Outlook.Store
    
    Set dict = New Scripting.Dictionary
    Set dictSender = New Scripting.Dictionary
    Set dictRecipient = New Scripting.Dictionary
 
    Debug.Print "start"

        intMessages = 0
        intVersion = GetOutlookVersion()
       
        For Each olkSto In Session.Stores
            lngRow = 2
            'ProcessFolderForCsv olkSto.GetRootFolder(), "root"
            ProcessFolderForDictionaryAndFile olkSto.GetRootFolder(), "root"
           
        Next
        
    Debug.Print "finished"
    
    'export the dictionary to a csv file
    Call exportDictionaries
   

End Sub

Sub ProcessFolderForDictionaryAndFile(olkFld As Outlook.MAPIFolder, ByVal parentFolderName As String)

    Dim tmpCounter
    tmpCounter = 0
    Dim tmpEmail
    Dim tmpStr
    Dim tmpLen
    Dim res
    Dim line
    Dim tmpFolderPath
    
    Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    
    
    'If Not pos > 1 Then
    If 1 = 1 Then
    
        Debug.Print "processing folder: " & olkFld.name
        
        tmpFolderPath = parentFolderName & "/" & olkFld.name
    
        'MsgBox ("found")
        
        If InStr(1, tmpFolderPath, "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx") > 1 Then
        
            
    
                On Error Resume Next
                    
                    For Each olkMsg In olkFld.Items
                        
                        'Only export messages, not receipts or appointment requests, etc.
                        'On Error Resume Next
                        If olkMsg.Class = olMail Then
                            'Add a row for each field in the message you want to export
                            'excWks.Cells(lngRow, 1) = olkFld.Name
                            'excWks.Cells(lngRow, 2) = GetSMTPAddress(olkMsg, intVersion)
                            'excWks.Cells(lngRow, 3) = olkMsg.SenderName
                            'excWks.Cells(lngRow, 4) = olkMsg.ReceivedTime
                            'excWks.Cells(lngRow, 5) = olkMsg.ReceivedByName
                            'excWks.Cells(lngRow, 6) = olkMsg.Subject
                            'excWks.Cells(lngRow, 7) = Left(olkMsg.Body, 500)
                            
                            tmpEmail = GetSMTPAddress(olkMsg, intVersion)
                            
                            'line = """" & makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & """, """ & makeCleanString(getUrl(tmpEmail)) & """, " & """" & makeCleanString(tmpFolderPath) & """"
                            
                            
                             'line = makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & ", " & makeCleanString(getUrl(tmpEmail)) & ", " & makeCleanString(tmpFolderPath)
                            
                            
                            line = """" & makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & """, """ & makeCleanString(getUrl(tmpEmail)) & """"
                            
                            ' Debug.Print line
                            
                            lineToDict tmpEmail, line, "sender"
                            
                            'process recipients
                            For Each Recipient In olkMsg.Recipients
                                Debug.Print "Recipient.Address = " & Recipient.Address
                                Debug.Print "Recipient.Name = " & Recipient.name
                                
                                line = """" & makeCleanString(Recipient.Address) & """, """ & makeCleanString(Recipient.name) & """, """ & makeCleanString(getUrl(Recipient.Address)) & """"
                                                            
                                lineToDict Recipient.Address, line, "recipient"
                                
                            Next Recipient
                            
                            ' save the mail in folder
                            saveMailInFolder (olkMsg)
                            
                            tmpCounter = tmpCounter + 1
                           
                            lngRow = lngRow + 1
                            intMessages = intMessages + 1
                            
                            If intMessages = 100 Then
                                exportDictcomplex (intMessages)
                                End
                            End If
                            
                        End If
                        'On Error GoTo 0
                        
                    Next
                
                On Error GoTo 0
        
        End If
        
        
    End If
    
    'Debug.Print "moving to next folder"
    
    Set olkMsg = Nothing
    
    For Each olkSub In olkFld.Folders
        If InStr(1, tmpFolderPath, "andreas@ciklum") > 1 Then
            ProcessFolderForDictionaryAndFile olkSub, tmpFolderPath
        End If
    Next
    
    Set olkSub = Nothing
    
End Sub


Sub saveMailInFolder(olkMsg)




End Sub


Sub makeFoldes(basePath, email, sendDate)

    If Len(Dir(basePath, vbDirectory)) = 0 Then
       MkDir basePath
    End If
    
    Dim tmp
    tmp = basePath & email
    
    If Len(Dir(tmp, vbDirectory)) = 0 Then
       MkDir tmp
    End If
    
    tmp = basePath & email & sendDate
    
    If Len(Dir(tmp, vbDirectory)) = 0 Then
       MkDir tmp
    End If
    
    currentPath = tmp


End Sub


Sub exportDictionaries()


     For Each x In dict.Keys()
        Debug.Print dict(x)
        Debug.Print dictSender(x)
        Debug.Print dictRecipient(x)
    Next
  


  

    Open counter & "-" & "dictionary.txt" For Append As #4
    
    Dim lineComplete
        
    
    For Each ele In dictComplex
        lineComplete = ele("email") & """, """ & ele("sender") & """, """ & ele("recipient") & """, """ & ele("line")
        Debug.Print "exporting line: " & lineComplete
        Print #4, lineComplete
    Next
    
    Close #4

    Debug.Print "exort finished"






End Sub

Sub testDict()

    Set dict = New Scripting.Dictionary
    Set dictSender = New Scripting.Dictionary
    Set dictRecipient = New Scripting.Dictionary
    
    dict("a") = "sdfgsdfg"
    dictSender("a") = "11111111111"
    dictRecipient("a") = "3333333333333"
    
    dict("b") = "4444"
    dictSender("b") = "5555"
    dictRecipient("b") = "6666"
    


    For Each x In dict.Keys()
        Debug.Print dict(x)
        Debug.Print dictSender(x)
        Debug.Print dictRecipient(x)
    Next
  


End Sub


Sub ExportMessagesToCsvSimple()

    Dim strFilename As String, olkSto As Outlook.Store
    
    Set dict = New Scripting.Dictionary
    Set dictSender = New Scripting.Dictionary
    Set dictRecipient = New Scripting.Dictionary
 
    
    Debug.Print "start"

        intMessages = 0
        intVersion = GetOutlookVersion()
       
        For Each olkSto In Session.Stores
           
            lngRow = 2
            
            'ProcessFolderForCsv olkSto.GetRootFolder(), "root"
            ProcessFolderForDictionary olkSto.GetRootFolder(), "root"
           
           
        Next
        
    Debug.Print "finished"
       
End Sub



Sub ProcessFolderForDictionary(olkFld As Outlook.MAPIFolder, ByVal parentFolderName As String)

    Dim tmpCounter
    tmpCounter = 0
    Dim tmpEmail
    Dim tmpStr
    Dim tmpLen
    Dim res
    Dim line
    Dim tmpFolderPath
    
    Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    'Write messages to spreadsheet
    
    'Debug.Print "processing folder: " & olkFld.name
    
    Open "folders.txt" For Append As #1
    Write #1, olkFld.name
    Close #1
    
    Dim tst As String
    tst = " xxx" & olkFld.name & "xxx"
    Dim pos
    pos = InStr(1, tst, "All Mail")
    
    'If Not pos > 1 Then
    If 1 = 1 Then
    
        Debug.Print "processing folder: " & olkFld.name
        
        tmpFolderPath = parentFolderName & "/" & olkFld.name
    
        'MsgBox ("found")
        
        If InStr(1, tmpFolderPath, "andreas@ciklum") > 1 Then
        
            
    
                On Error Resume Next
                    
                    For Each olkMsg In olkFld.Items
                        
                        'Only export messages, not receipts or appointment requests, etc.
                        'On Error Resume Next
                        If olkMsg.Class = olMail Then
                            'Add a row for each field in the message you want to export
                            'excWks.Cells(lngRow, 1) = olkFld.Name
                            'excWks.Cells(lngRow, 2) = GetSMTPAddress(olkMsg, intVersion)
                            'excWks.Cells(lngRow, 3) = olkMsg.SenderName
                            'excWks.Cells(lngRow, 4) = olkMsg.ReceivedTime
                            'excWks.Cells(lngRow, 5) = olkMsg.ReceivedByName
                            'excWks.Cells(lngRow, 6) = olkMsg.Subject
                            'excWks.Cells(lngRow, 7) = Left(olkMsg.Body, 500)
                            
                            tmpEmail = GetSMTPAddress(olkMsg, intVersion)
                            
                            'line = """" & makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & """, """ & makeCleanString(getUrl(tmpEmail)) & """, " & """" & makeCleanString(tmpFolderPath) & """"
                            
                            
                             'line = makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & ", " & makeCleanString(getUrl(tmpEmail)) & ", " & makeCleanString(tmpFolderPath)
                            
                            
                            line = """" & makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & """, """ & makeCleanString(getUrl(tmpEmail)) & """"
                            
                            ' Debug.Print line
                            
                            lineToDict tmpEmail, line, "sender"
                            
                            'process recipients
                            For Each Recipient In olkMsg.Recipients
                                Debug.Print "Recipient.Address = " & Recipient.Address
                                Debug.Print "Recipient.Name = " & Recipient.name
                                
                                line = """" & makeCleanString(Recipient.Address) & """, """ & makeCleanString(Recipient.name) & """, """ & makeCleanString(getUrl(Recipient.Address)) & """"
                                                            
                                lineToDict Recipient.Address, line, "recipient"
                                
                            Next Recipient
                            
                      
                            
                            tmpCounter = tmpCounter + 1
                           
                            
                            
                            lngRow = lngRow + 1
                            intMessages = intMessages + 1
                            
                            If intMessages = 100 Then
                                exportDictcomplex (intMessages)
                                End
                            End If
                            
                            
                        End If
                        'On Error GoTo 0
                    Next
                
                On Error GoTo 0
        
            
        End If
        
        
    End If
    
    'Debug.Print "moving to next folder"
    
    Set olkMsg = Nothing
    
    For Each olkSub In olkFld.Folders
        If InStr(1, tmpFolderPath, "andreas@ciklum") > 1 Then
            ProcessFolderForDictionary olkSub, tmpFolderPath
        End If
    Next
    
    Set olkSub = Nothing
    
End Sub

Sub lineToDict(email, line, senderOrRecipient)

    On Error GoTo err
    'dictComplex
    Dim lineDict


    If dict.Exists(email) Then
        Debug.Print "email exists: " & email
    Else
        Debug.Print "new email email : " & email
        dict(email) = line
        dictRecipient(email) = 0
        dictSender(email) = 0
                
        'Open "dictionary.txt" For Append As #3
        'Print #3, line
        'Close #3
        
        Debug.Print "size of dictionary now : " & dict.Count
    End If
    
    If senderOrRecipient = "sender" Then
        dictSender(email) = dictSender(email) + 1
    End If
    
    If senderOrRecipient = "recipient" Then
        dictRecipient(email) = dictRecipient(email) + 1
    End If
    
    Debug.Print "email: " & email
    Debug.Print "dictSender(email): " & dictSender(email)
    Debug.Print "dictRecipient(email): " & dictRecipient(email)
        
    Debug.Print "____________________________________________"

    'add parameters to the line
    
    
err:
    
    Debug.Print "error: " & err.Description
    Debug.Print "error: " & err.Source
    On Error Resume Next

End Sub


Sub exportDictcomplex(counter)

    

    Open counter & "-" & "dictionary.txt" For Append As #4
    
    Dim lineComplete
        
    
    For Each ele In dictComplex
        lineComplete = ele("email") & """, """ & ele("sender") & """, """ & ele("recipient") & """, """ & ele("line")
        Debug.Print "exporting line: " & lineComplete
        Print #4, lineComplete
    Next
    
    Close #4

    Debug.Print "exort finished"
    
End Sub

Sub ProcessFolderForCsv(olkFld As Outlook.MAPIFolder, ByVal parentFolderName As String)

    Dim tmpCounter
    tmpCounter = 0
    Dim tmpEmail
    Dim tmpStr
    Dim tmpLen
    Dim res
    Dim line
    Dim tmpFolderPath
    
    Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    'Write messages to spreadsheet
    
    'Debug.Print "processing folder: " & olkFld.name
    
    Open "folders.txt" For Append As #1
    Write #1, olkFld.name
    Close #1
    
    Dim tst As String
    tst = " xxx" & olkFld.name & "xxx"
    Dim pos
    pos = InStr(1, tst, "All Mail")
    
    'If Not pos > 1 Then
    If 1 = 1 Then
    
        Debug.Print "processing folder: " & olkFld.name
        
        tmpFolderPath = parentFolderName & "/" & olkFld.name
    
        'MsgBox ("found")
        
        If InStr(1, tmpFolderPath, "andreas@ciklum") > 1 Then
        
            Open "contacts.txt" For Append As #2
    
                On Error Resume Next
                    
                    For Each olkMsg In olkFld.Items
                        
                        'Only export messages, not receipts or appointment requests, etc.
                        'On Error Resume Next
                        If olkMsg.Class = olMail Then
                            'Add a row for each field in the message you want to export
                            'excWks.Cells(lngRow, 1) = olkFld.Name
                            'excWks.Cells(lngRow, 2) = GetSMTPAddress(olkMsg, intVersion)
                            'excWks.Cells(lngRow, 3) = olkMsg.SenderName
                            'excWks.Cells(lngRow, 4) = olkMsg.ReceivedTime
                            'excWks.Cells(lngRow, 5) = olkMsg.ReceivedByName
                            'excWks.Cells(lngRow, 6) = olkMsg.Subject
                            'excWks.Cells(lngRow, 7) = Left(olkMsg.Body, 500)
                            
                            tmpEmail = GetSMTPAddress(olkMsg, intVersion)
                            
                            line = """" & makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & """, """ & makeCleanString(getUrl(tmpEmail)) & """, " & """" & makeCleanString(tmpFolderPath) & """"
                            
                            
                             'line = makeCleanString(tmpEmail) & """, """ & makeCleanString(olkMsg.SenderName) & ", " & makeCleanString(getUrl(tmpEmail)) & ", " & makeCleanString(tmpFolderPath)
                            
                            
                            Debug.Print line
                            
                            ' print stuff to the file
                            Print #2, line
                            
                            
                            tmpCounter = tmpCounter + 1
                            
                            If tmpCounter = 10 Then
                                'Debug.Print "sleeping..."
                                'Application.Wait (Now + TimeValue("00:00:01"))
                                'tmpCounter = 1
                            End If
                            
                            
                            lngRow = lngRow + 1
                            intMessages = intMessages + 1
                            
                            If intMessages = 100 Then
                                'End
                            End If
                            
                            
                        End If
                        'On Error GoTo 0
                    Next
                
                On Error GoTo 0
        
            Close #2
        End If
        
        
    End If
    
    'Debug.Print "moving to next folder"
    
    Set olkMsg = Nothing
    
    For Each olkSub In olkFld.Folders
        ProcessFolderForCsv olkSub, tmpFolderPath
    Next
    
    Set olkSub = Nothing
    
End Sub

Sub test()

    Dim f
    f = makeCleanString("asdfafds""asdfasdf""asdffasdf")

    Debug.Print (f)

End Sub

Function makeCleanString(s)

    s = Replace(s, """", " ")
    s = Trim(s)

    makeCleanString = s

End Function



Sub ExportMessagesToExcelAsArray()
    Dim strFilename As String, olkSto As Outlook.Store
    
    Set dic = New Scripting.Dictionary
    Set dic2 = New Scripting.Dictionary
    
    Debug.Print 1

        intMessages = 0
        intVersion = GetOutlookVersion()
        'Set excApp = CreateObject("Excel.Application")
        Debug.Print 2
        'Set excWkb = excApp.Workbooks.Add
        For Each olkSto In Session.Stores
            Debug.Print 3
            'Set excWks = excWkb.Worksheets.Add()
            'excWks.name = "Output1"
            'Write Excel Column Headers
            Debug.Print 4
           ' With excWks
            '    .Cells(1, 1) = "EmailAddress"
             '   .Cells(1, 2) = "Name"
              '  .Cells(1, 3) = "URL"
           ' End With
            
            'excWkb.Save
             
             
            Debug.Print 5
            lngRow = 2
            Debug.Print 5
            ProcessFolderForArray olkSto.GetRootFolder()
            'ProcessFolderNoPRINT olkSto.GetRootFolder()
            Debug.Print 6
        Next
        'excWkb.Save
 
    'Set excWks = Nothing
    'Set excWkb = Nothing
    'excApp.Quit
    'Set excApp = Nothing
    'MsgBox "Process complete.  A total of " & intMessages & " messages were exported.", vbInformation + vbOKOnly, "Export messages to Excel"
End Sub



Sub ProcessFolderForArray(olkFld As Outlook.MAPIFolder)

    Dim tmpCounter
    tmpCounter = 0
    Dim tmpEmail
    Dim tmpStr
    Dim tmpLen
    Dim res
    
    Debug.Print 11
    Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    'Write messages to spreadsheet
    For Each olkMsg In olkFld.Items
        
        'Only export messages, not receipts or appointment requests, etc.
        On Error Resume Next
        If olkMsg.Class = olMail Then
            'Add a row for each field in the message you want to export
            'excWks.Cells(lngRow, 1) = olkFld.Name
            'excWks.Cells(lngRow, 2) = GetSMTPAddress(olkMsg, intVersion)
            'excWks.Cells(lngRow, 3) = olkMsg.SenderName
            'excWks.Cells(lngRow, 4) = olkMsg.ReceivedTime
            'excWks.Cells(lngRow, 5) = olkMsg.ReceivedByName
            'excWks.Cells(lngRow, 6) = olkMsg.Subject
            'excWks.Cells(lngRow, 7) = Left(olkMsg.Body, 500)
            
            tmpEmail = GetSMTPAddress(olkMsg, intVersion)
            tmpLen = getLengthFromDic(tmpEmail)
            If tmpLen > 1 Then
                Debug.Print "Address exists already:  " & olkMsg.SenderName
                'check if new addresss is better
                tmpStr = Len(olkMsg.SenderName & getUrl(tmpEmail))
                
                If tmpLen >= tmpStr Then
                    ' do nothing
                Else
                   res = setToDic(tmpEmail, olkMsg.SenderName, getUrl(tmpEmail))
                End If
                
            Else
                Debug.Print "adding Address:  " & tmpEmail
                res = setToDic(tmpEmail, olkMsg.SenderName, getUrl(tmpEmail))
                
            End If
            
            Debug.Print "Dictionary size:  " & dic.Count
            
            
            
            'Debug.Print olkFld.name & " message number " & intMessages
            'Debug.Print " message number " & olkMsg.SenderName
            'Debug.Print GetSMTPAddress(olkMsg, intVersion)
            'Debug.Print olkMsg.ReceivedTime
            'Debug.Print olkMsg.ReceivedByName
            'Debug.Print olkMsg.Subject
            'Debug.Print Left(olkMsg.Body, 500)
            
            
            
            
         '   If tmpCounter = 100 Then
         '       tmpCounter = 0
          '      Debug.Print olkFld.name & " message number " & intMessages & " now saving file "
           '     excWkb.Save
           ' End If
            
            tmpCounter = tmpCounter + 1
            
            lngRow = lngRow + 1
            intMessages = intMessages + 1
        End If
        On Error GoTo 0
    Next
    
    'excWkb.Save
    
    
    Debug.Print 13
    Set olkMsg = Nothing
    For Each olkSub In olkFld.Folders
        ProcessFolderForArray olkSub
    Next
    Set olkSub = Nothing
End Sub

Function getUrl(email)
    Dim pos
    Dim ret
    pos = InStr(1, email, "@")
    
    If pos > 1 Then
        ret = "http://www." & Mid(email, pos + 1)
    Else
        ret = "-"
    End If

    getUrl = ret
End Function
Function getLengthFromDic(s)

    Dim ret As Integer
    Dim tmpStr As String
    Dim row
    
    ret = 0

    If dic.Exists(s) Then
        row = dic(s)
        tmpStr = row("name") & row("url")
        Debug.Print "in getLengthFromDic: " & tmpStr
        ret = Len(tmpStr)
    Else
        Debug.Print "in getLengthFromDic - address does not exist: " & s
    End If
    
    getLengthFromDic = ret
            
End Function

Function setToDic(email, name, url)
    Dim tmpDic
    Set dic(email) = New Scripting.Dictionary
    
    tmpDic = dic(email)
    tmpDic("name") = name
    tmpDic("url") = url
    Debug.Print " setToDic done for  " & email
    
   

End Function

Sub ExportMessagesToExcelSimple()
    Dim strFilename As String, olkSto As Outlook.Store
    strFilename = "C:\Users\Alex Sickert\Desktop\test.xlsx"
    Debug.Print 1
    If strFilename <> "" Then
        intMessages = 0
        intVersion = GetOutlookVersion()
        Set excApp = CreateObject("Excel.Application")
        Debug.Print 2
        Set excWkb = excApp.Workbooks.Add
        For Each olkSto In Session.Stores
            Debug.Print 3
            Set excWks = excWkb.Worksheets.Add()
            excWks.name = "Output1"
            'Write Excel Column Headers
            Debug.Print 4
            With excWks
                .Cells(1, 1) = "Folder"
                .Cells(1, 2) = "Sender"
                .Cells(1, 3) = "Sender Name"
                .Cells(1, 4) = "Received"
                .Cells(1, 5) = "Sent To"
                .Cells(1, 6) = "Subject"
                .Cells(1, 7) = "Content"
            End With
            
            excWkb.Save
             
             
            Debug.Print 5
            lngRow = 2
            Debug.Print 5
            ProcessFolder olkSto.GetRootFolder()
            'ProcessFolderNoPRINT olkSto.GetRootFolder()
            Debug.Print 6
        Next
        excWkb.SaveAs strFilename
    End If
    Set excWks = Nothing
    Set excWkb = Nothing
    excApp.Quit
    Set excApp = Nothing
    MsgBox "Process complete.  A total of " & intMessages & " messages were exported.", vbInformation + vbOKOnly, "Export messages to Excel"
End Sub




Sub ProcessFolder(olkFld As Outlook.MAPIFolder)

    Dim tmpCounter
    tmpCounter = 0
    
    Debug.Print 11
    Dim olkMsg As Object, olkSub As Outlook.MAPIFolder
    'Write messages to spreadsheet
    For Each olkMsg In olkFld.Items
        
        'Only export messages, not receipts or appointment requests, etc.
        On Error Resume Next
        If olkMsg.Class = olMail Then
            'Add a row for each field in the message you want to export
            excWks.Cells(lngRow, 1) = olkFld.name
            excWks.Cells(lngRow, 2) = GetSMTPAddress(olkMsg, intVersion)
            excWks.Cells(lngRow, 3) = olkMsg.SenderName
            excWks.Cells(lngRow, 4) = olkMsg.ReceivedTime
            excWks.Cells(lngRow, 5) = olkMsg.ReceivedByName
            excWks.Cells(lngRow, 6) = olkMsg.Subject
            excWks.Cells(lngRow, 7) = Left(olkMsg.Body, 500)
            
            Debug.Print olkFld.name & " message number " & intMessages
            Debug.Print " message number " & olkMsg.SenderName
            'Debug.Print GetSMTPAddress(olkMsg, intVersion)
            'Debug.Print olkMsg.ReceivedTime
            'Debug.Print olkMsg.ReceivedByName
            'Debug.Print olkMsg.Subject
            'Debug.Print Left(olkMsg.Body, 500)
            
            
            
            
            If tmpCounter = 100 Then
                tmpCounter = 0
                Debug.Print olkFld.name & " message number " & intMessages & " now saving file "
                excWkb.Save
            End If
            
            tmpCounter = tmpCounter + 1
            
            lngRow = lngRow + 1
            intMessages = intMessages + 1
        End If
        On Error GoTo 0
    Next
    
    excWkb.Save
    
    
    Debug.Print 13
    Set olkMsg = Nothing
    For Each olkSub In olkFld.Folders
        ProcessFolder olkSub
    Next
    Set olkSub = Nothing
End Sub

Private Function GetSMTPAddress(Item As Outlook.MailItem, intOutlookVersion As Integer) As String
    Dim olkSnd As Outlook.AddressEntry, olkEnt As Object
    On Error Resume Next
    Select Case intOutlookVersion
        Case Is < 14
            If Item.SenderEmailType = "EX" Then
                GetSMTPAddress = SMTP2007(Item)
            Else
                GetSMTPAddress = Item.SenderEmailAddress
            End If
        Case Else
            Set olkSnd = Item.Sender
            If olkSnd.AddressEntryUserType = olExchangeUserAddressEntry Then
                Set olkEnt = olkSnd.GetExchangeUser
                GetSMTPAddress = olkEnt.PrimarySmtpAddress
            Else
                GetSMTPAddress = Item.SenderEmailAddress
            End If
    End Select
    On Error GoTo 0
    Set olkPrp = Nothing
    Set olkSnd = Nothing
    Set olkEnt = Nothing
End Function

Function GetOutlookVersion() As Integer
    Dim arrVer As Variant
    arrVer = Split(Outlook.Version, ".")
    GetOutlookVersion = arrVer(0)
End Function

Function SMTP2007(olkMsg As Outlook.MailItem) As String
    Dim olkPA As Outlook.PropertyAccessor
    On Error Resume Next
    Set olkPA = olkMsg.PropertyAccessor
    SMTP2007 = olkPA.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x5D01001E")
    On Error GoTo 0
    Set olkPA = Nothing
End Function

