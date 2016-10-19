Attribute VB_Name = "Module1"
Sub BeyondCompare()
 ' Set reference to VB Script library
 ' Microsoft VBScript Regular Expressions 5.5
 
    Dim olMail As Outlook.MailItem
    Dim RegVersion As RegExp
    Dim RegFile As RegExp
    Dim M1 As MatchCollection
    Dim M As Match
    Dim Version As String
    Version = ""
    
    Dim LeftCompare As String
    Dim RightCompare As String
    LeftCompare = "Baseline"
    RightCompare = "Modified"
        
    Set olMail = Application.ActiveExplorer().Selection(1)
    
    Set RegVersion = New RegExp
    Set RegFile = New RegExp
    
    With RegFile
        .Pattern = "Baseline(.*) folders."
        .Global = True
    End With
    If RegFile.Test(olMail.Body) Then
    
        ' determine version of comparison
        Set M1 = RegFile.Execute(olMail.Body)
        For Each M In M1
            
            Dim v As String
            
            v = M.Value
            v = Replace(v, " folders.", "")
            v = Replace(v, LeftCompare, "")
            v = Replace(v, "(", "")
            v = Replace(v, ")", "")
            v = Replace(v, "v", "")
            v = Replace(v, " ", "")

            If StrComp(v, "") = 1 Then
                Version = " (v" + v + ")"
            End If
            
            Exit For
          
        Next
    End If
    
    With RegFile
        .Pattern = "<file:///(.*)>"
        .Global = True
    End With
    If RegFile.Test(olMail.Body) Then
    
        Set M1 = RegFile.Execute(olMail.Body)
        For Each M In M1
            ' look for file link
            Dim path As String
            path = Replace(M.Value, "<file:///", "")
            path = Replace(path, ">", "")
            
            Dim exeStr As String
            
            ' create BComp.exe execution string
            exeString = "BComp "
            exeString = exeString + """" + path + LeftCompare
            exeString = exeString + Version + """ "
            exeString = exeString + """" + path + RightCompare
            exeString = exeString + Version + """"
            
            exeString = Replace(exeString, "%20", " ") ' html space

            ' execute in shell
            Shell exeString
       
        Next
    End If
    
End Sub

