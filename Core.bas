Sub ProcessMultipleChina()

Dim userID, TemplateFolder, sourceFolder, destinationFolder As String
Dim file, filsystem, folder As Object
Dim files As Collection
Dim filewithE, pairedFile, fileNumber As String
Dim doc, templatedoc, doc1, doc2, copiedTemplateDoc As Document
Dim matched, isTemplateOpen As Boolean
Dim newTemplateFileName As String
    
Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    FileDialog.Title = "Select the Source Folder Containing the agent copies"
        If FileDialog.Show = -1 Then
            sourceFolder = FileDialog.SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Process cancelled."
            Exit Sub
        End If
        
        Debug.Print "sourceFolder: " & sourceFolder
        
Set FileDialog = Application.FileDialog(msoFileDialogFolderPicker)
    FileDialog.Title = "Select the Destination Folder to save the completed reports"
        If FileDialog.Show = -1 Then
            destinationFolder = FileDialog.SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Process cancelled."
            Exit Sub
        End If
        
        Debug.Print "destination folder: " & destinationFolder
       
Set filsystem = CreateObject("Scripting.FileSystemObject")
Set folder = filsystem.GetFolder(sourceFolder)
Set files = New Collection
    
For Each file In folder.files
    filewithE = file.name
    Debug.Print "Found Files: " & filewithE
    If InStr(filewithE, "(E).doc") >= 8 Then
        Set doc1 = Documents.Open(file.path)
        fileNumber = Left(filewithE, 8)
            
        Debug.Print "fileNumber: " & fileNumber
        Debug.Print "First file (E) opened: " & filewithE
        
        matched = False
        For Each pairedFile In folder.files
            If StrComp(pairedFile.name, filewithE, vbTextCompare) <> 0 And Left(pairedFile.name, Len(fileNumber)) = fileNumber And InStr(pairedFile.name, "(E)") = 0 Then
                matched = True
                Exit For
            End If
        Next pairedFile
        
        Debug.Print "pairedFile: " & pairedFile
        
        If matched Then
            Set doc2 = Documents.Open(pairedFile.path)
            Debug.Print "Paired file opened: " & pairedFile.name
            
            Dim TemplateFolderAndFile As String
            
            TemplateFolderAndFile = "C:\CHNROC\TemplateFile.docm"

            Debug.Print "Template Folder path: " & TemplateFolderAndFile
            Set templatedoc = Documents.Open(TemplateFolderAndFile)
            
            Call SlotChRC(templatedoc, doc1, doc2, destinationFolder)
            
            If Not doc1 Is Nothing Then
                doc1.Close savechanges:=False
            End If
            
            If Not doc2 Is Nothing Then
                doc2.Close savechanges:=False
            End If
            
         End If
       End If
Next file

End Sub


Sub SlotChRC(ByVal templatedoc As Document, ByVal doc1 As Document, ByVal doc2 As Document, ByVal destinationFolder As String)

Dim newFileName, savePath As String

If Len(doc1) > 0 And Len(doc2) > 0 And Len(templatedoc) > 0 Then

Application.ScreenUpdating = False

Documents(doc1).Activate

With Documents(doc1)

    For Each mytable In ActiveDocument.Tables
        mytable.Range.Editors.Add wdEditorEveryone
        mytable.Rows.Alignment = wdAlignRowLeft
        mytable.PreferredWidth = InchesToPoints(6.25)
    Next
    
    ActiveDocument.SelectAllEditableRanges (wdEditorEveryone)
    ActiveDocument.DeleteAllEditableRanges (wdEditorEveryone)
    
    Selection.ParagraphFormat.SpaceBefore = 0
    Selection.ParagraphFormat.SpaceAfter = 0
    Selection.ParagraphFormat.LeftIndent = InchesToPoints(0)
    Selection.Shading.BackgroundPatternColor = wdColorAutomatic
    Selection.Font.name = "Arial"
    Selection.Font.Size = 9
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    Selection.ParagraphFormat.LineSpacing = LinesToPoints(1)
    Selection.WholeStory
    DoEvents
    
    Selection.Collapse direction:=wdCollapseEnd
    
    Selection.GoTo wdGoToPage, wdGoToAbsolute, 1
    
    Dim EnglishTitle, companyNo, companyName, Address, ZipCode, telephone, fax, website, AICNo, SocialCreditCode, registry, RegZipCode, RegisterAdds, EstablishmentDate As String
    Dim DurationOfOps, legalstatus, LegalRep, RegisterCapital, BusinessScope As String
    Dim openBracket, closeBracket As Long
    Dim RegisterAmount, dollarsign As String
    Dim telResult, layerSentence1, layerSentence2 As String
    Dim layerImage As InlineShape
    Dim layerRange As Range
    Dim IvgTable As table
    Dim cell As cell
    Dim rng As Range
    Dim sentencesBelow As Range
    Dim layerTable As table
    Dim Ivgresult As String
    
    SearchText "Business Registration Report", False, False
        If Selection.Find.Found Then
            Selection.MoveDown unit:=wdLine, Count:=1
            Selection.Expand unit:=wdParagraph
            EnglishTitle = Selection.Text
            EnglishTitle = Replace(EnglishTitle, Chr(13), "")
            EnglishTitle = Trim(EnglishTitle)
            Selection.Text = EnglishTitle
            Debug.Print EnglishTitle
        End If
        
    SearchText "Your Reference:", False, False
        If Selection.Find.Found Then
        companyNo = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
             
    SearchText "INVESTIGATION NOTES", True, True
        If Selection.Find.Found Then
            If Selection.Information(wdWithInTable) Then
                    Set IvgTable = Selection.Tables(1)
                        If IvgTable.Range.Text <> Chr(13) & Chr(7) Then
                            For Each cell In IvgTable.Range.Cells
                            Set cellRange = cell.Range
                            cellRange.ParagraphFormat.Alignment = wdAlignParagraphJustify
                            Next cell
                            IvgTable.Range.Copy
                        End If
            End If
            Else
            Ivgresult = "No investigation notes."
        End If
          
    SearchText ("PROFILE")
    If Selection.Find.Found Then
    
        SearchText "Subject Name:", False, False
        If Selection.Find.Found Then
        companyName = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText "Telephone:", False, False
        If Selection.Find.Found Then
        telephone = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
            Else
            SearchText "Hotline:", False, False
                If Selection.Find.Found Then
                    telephone = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
                End If
        End If
        
        If telephone = "" Then
            telephone = "NA"
        End If
                
        If telephone = "NA" Then
            Address = "NA"
        Else
            SearchText "N.O.C.:", False, False
            If Selection.Find.Found Then
            Selection.MoveDown unit:=wdLine, Count:=1
            Selection.MoveLeft unit:=wdWord, Count:=1, Extend:=wdMove
                SearchText "Address:", False, False
                If Selection.Find.Found Then
                Address = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
                    If InStr(Address, "Town") > 0 Or InStr(Address, "District") > 0 Or InStr(Address, "City") > 0 Then
                        Address = Replace(Address, "Town", "Town, ")
                        Address = Replace(Address, "District", "District, ")
                        Address = Replace(Address, "City", "City, ")
                        Address = Replace(Address, "Zone", "Zone, ")
                        Address = Replace(Address, "Park", "Park, ")
                        Address = Replace(Address, "Road", "Road, ")
                    End If
                End If
            End If
        End If
                    
        If Address = "NA" Then
            ZipCode = "NA"
        Else
            SearchText "Zip Code:", False, False
                If Selection.Find.Found Then
                ZipCode = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
                End If
        End If
        
        SearchText "Facsimile:", False, False
        If Selection.Find.Found Then
        fax = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        Else
        fax = "NA"
        End If
        
        SearchText "Website:", False, False
        If Selection.Find.Found Then
            website = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        Else
            website = "NA"
        End If
    End If
    
    SearchText ("REGISTRATION")
    If Selection.Find.Found Then
    
        SearchText "Establishment Date:", False, False
        If Selection.Find.Found Then
            EstablishmentDate = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText "Registered Address:", False, False
        If Selection.Find.Found Then
            RegisterAdds = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        Selection.MoveDown unit:=wdLine, Count:=1
        Selection.HomeKey unit:=wdLine
               
        SearchText "Zip Code:", False, False
        If Selection.Find.Found Then
            RegZipCode = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText "Registry:", False, False
        If Selection.Find.Found Then
            registry = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText "Legal Rep.:", False, False
        If Selection.Find.Found Then
            LegalRep = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText "AIC No.:", False, False
        If Selection.Find.Found Then
            AICNo = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText "Unified Social Credit Code:", False, False
        If Selection.Find.Found Then
            SocialCreditCode = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
             
        SearchText "Legal Status:", False, False
        If Selection.Find.Found Then
            legalstatus = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
             
        SearchText "Duration of Operation:", False, False
        If Selection.Find.Found Then
            DurationOfOps = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
             
       SearchText "Registered Capital:", False, False
        If Selection.Find.Found Then
            RegisterCapital = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        openBracket = (InStrRev(RegisterCapital, "("))
        If openBracket > 0 Then
            closeBracket = (InStr(openBracket, RegisterCapital, ")"))
                If closeBracket > 0 Then
                    dollarsign = Mid(RegisterCapital, openBracket + 1, closeBracket - openBracket - 1)
                Else
                    dollarsign = ""
                End If
        End If
        
    RegisterAmount = Left(RegisterCapital, Len(RegisterCapital) - 6)
        
        SearchText "Business Scope:", False, False
        If Selection.Find.Found Then
            BusinessScope = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
    End If
    
    SearchText ("Changes in Registration")
    If Selection.Find.Found Then
        Selection.GoToNext wdGoToTable
        Set ChangeInRegister = Selection.Tables(1)
    End If
    
    SearchText ("SHAREHOLDERS AND SHARES")
    If Selection.Find.Found Then
        SearchText ("Shareholders")
        Set ShareHolders1 = Selection.Tables(1)
        
        SearchText ("Shareholders")
        Set ShareHolders2 = Selection.Tables(1)
    End If
      
    SearchText ("LAYERS OF SHAREHOLDING")
    If Selection.Find.Found Then
        Selection.MoveDown unit:=wdLine, Count:=1
            If Selection.InlineShapes.Count > 0 Then
                Set layerImage = Selection.InlineShapes(1)
                layerImage.Range.Copy
            End If
    End If
    
        Selection.GoToNext wdGoToTable
        Set layerTable = Selection.Tables(1)
        
    If Not layerTable Is Nothing Then
        layerTable.Range.Select
        Selection.MoveDown unit:=wdLine, Count:=2
        Selection.Expand unit:=wdParagraph
            If Selection.Find.Found Then
                layerSentence1 = RemoveExtSpacing(Selection.Text)
            End If
            
        Selection.MoveDown unit:=wdLine, Count:=2
        Selection.Expand unit:=wdParagraph
            If Selection.Find.Found Then
                layerSentence2 = RemoveExtSpacing(Selection.Text)
            End If
    End If
    
End With

Documents(doc2).Activate

With Documents(doc2)
    
    Selection.WholeStory
    Selection.Font.name = "SimSun"
    Selection.Font.Size = 9
    Selection.Font.Color = wdColorAutomatic
    With Selection.ParagraphFormat
        .FirstLineIndent = InchesToPoints(0)
        .LeftIndent = InchesToPoints(0)
        .SpaceBefore = 0
        .SpaceBeforeAuto = False
        .SpaceAfter = 0
        .SpaceAfterAuto = False
        .LineSpacingRule = wdLineSpaceSingle
        .LineUnitBefore = 0
        .LineUnitAfter = 0
    End With
    
    Dim CNCompanyName, CNBusinessAdds, CNLegalRep, CNRegisterAds As String
    Dim ChineseNameShareHolders1 As Collection, ChineseNameShareHolders2 As Collection
      
    Set ChineseNameShareHolders1 = New Collection
    Set ChineseNameShareHolders2 = New Collection
    
    SearchText ChrW(20844) & ChrW(21496) & ChrW(21517) & ChrW(31216), False, False
    If Selection.Find.Found Then
     CNCompanyName = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
    End If
    
    SearchText ChrW(30005) & ChrW(-29731) & ChrW(-230), False, False
        If Selection.Find.Execute Then
            Selection.MoveUp unit:=wdLine, Count:=2
                SearchText ChrW(22320) & ChrW(22336) & ":", False, False
                If Selection.Find.Found Then
                    CNBusinessAdds = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
                End If
        Else
                    CNBusinessAdds = ""
        End If
       
    SearchText ChrW(27880) & ChrW(20876) & ChrW(36164) & ChrW(26009)
    If Selection.Find.Found Then
        
        SearchText ChrW(27880) & ChrW(20876) & ChrW(22320) & ChrW(22336) & ":", False, False
        If Selection.Find.Found Then
            CNRegisterAds = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
        
        SearchText ChrW(27861) & ChrW(23450) & ChrW(20195) & ChrW(34920) & ChrW(20154) & ":", False, False
        If Selection.Find.Found Then
            CNLegalRep = RemoveExtSpacing(Selection.Cells(1).Next.Range.Text)
        End If
    End If
    
    SearchText ChrW(32929) & ChrW(19996) & ChrW(21450) & ChrW(32929) & ChrW(20221)
    If Selection.Find.Found Then
        
        Selection.GoToNext wdGoToTable
        With Selection.Tables(1)
            For r = 2 To .Rows.Count - 2
                ChineseNameShareHolders1.Add Replace(RemoveSubString(.cell(r, 1).Range.Text, ""), Chr(13), vbNullString)
            Next r
        End With
        
        Selection.GoToNext wdGoToTable
        With Selection.Tables(1)
            For r = 2 To .Rows.Count
                ChineseNameShareHolders2.Add Replace(RemoveSubString(.cell(r, 1).Range.Text, ""), Chr(13), vbNullString)
            Next r
        End With
        
    End If

End With

Documents(templatedoc).Activate

With Documents(templatedoc)

    .Bookmarks("ENCompanyName").Range.Text = EnglishTitle & vbCr & CNCompanyName
    .Bookmarks("Month").Range.Text = MonthName(Month(Date))
    .Bookmarks("companyNo").Range.Text = companyNo
    .Bookmarks("ENCompanyName1").Range.Text = EnglishTitle & vbNewLine & CNCompanyName
        If Address <> "NA" Then
    .Bookmarks("ENBusinessAdds").Range.Text = Address & ", China" & vbCr & ChrW(20013) & ChrW(22269) & CNBusinessAdds
        Else
    .Bookmarks("ENBusinessAdds").Range.Text = Address
        End If
    .Bookmarks("ZipCode").Range.Text = ZipCode
    .Bookmarks("Telephone").Range.Text = telephone
    .Bookmarks("Fax").Range.Text = fax
    .Bookmarks("Website").Range.Text = website
    .Bookmarks("AICNo").Range.Text = AICNo
    .Bookmarks("SocialCreditCode").Range.Text = SocialCreditCode
    .Bookmarks("RegistrationAgent").Range.Text = registry
    .Bookmarks("ENRegistrationAdds").Range.Text = RegisterAdds & "," & ChrW(32) & RegZipCode & "," & ChrW(32) & "China"
    .Bookmarks("CNRegistrationAdds").Range.Text = ChrW(20013) & ChrW(22269) & CNRegisterAds & "," & ChrW(32) & RegZipCode
    .Bookmarks("DateOfEstablishment").Range.Text = EstablishmentDate
    .Bookmarks("DurationOfOps").Range.Text = DurationOfOps
    .Bookmarks("LegalForm").Range.Text = legalstatus
    .Bookmarks("dollarsign").Range.Text = Trim(dollarsign)
    If RegisterAmount <> "" then
        .Bookmarks("RegisterAmount").Range.Text = RegisterAmount
    End if
    .Bookmarks("LegalRep").Range.Text = LegalRep & " " & CNLegalRep
    .Bookmarks("BussScope").Range.Text = BusinessScope
    
    If IsEmpty(ChangeInRegister) = True Then
       .Bookmarks("ChangeInRegister").Range.Delete
       .Bookmarks("ChangeInRegisterDELALL").Range.Delete
    Else
       .Bookmarks("ChangeInRegister").Range.FormattedText = ChangeInRegister.Range.FormattedText
       .Bookmarks("ChangeInRegisterDEL").Range.Delete
    End If
    
    .Bookmarks("ShareHolders1").Range.FormattedText = ShareHolders1.Range.FormattedText
    .Bookmarks("ShareHolders2").Range.FormattedText = ShareHolders2.Range.FormattedText
       
    Set layerRange = .Bookmarks("layerImage").Range
    layerRange.Paste
    
    If IsEmpty(layerTable) = False Then
        .Bookmarks("layerTable").Range.FormattedText = layerTable.Range.FormattedText
    End If
    
    .Bookmarks("sentencesBelow").Range.Text = layerSentence1 & vbLf & vbLf & layerSentence2
    If Ivgresult <> "" Then
        .Bookmarks("IvgTable").Range.Text = Ivgresult
        Else
        .Bookmarks("IvgTable").Range.FormattedText = IvgTable.Range.FormattedText
    End If
    
    Selection.GoTo wdGoToPage, wdGoToAbsolute, 1
    
    SearchText "Business Address", True, False
        If Selection.Find.Found Then
            Selection.Cells(1).Range.Select
            Selection.Expand unit:=wdParagraph
            Selection.Font.name = "Arial"
        End If

    SearchText "Telephone", True, False
    If Selection.Find.Found Then
        telResult = Selection.Cells(1).Next.Range.Text
            If InStr(1, telResult, "NA", vbcompare) > 0 Then
                .Bookmarks("NoContact").Range.Text = "No contact information of Subject was found during the course of investigation. The information in the report is based on the registration information of Subject with the local authority."
            End If
    End If
        
        SearchText "Registered Address", True, False
        If Selection.Find.Found Then
            Selection.Cells(1).Next.Range.Select
            Selection.Expand unit:=wdParagraph
            Selection.Font.name = "Arial"
        End If
        
        SearchText "INVESTIGATION NOTES", True, True
        If Selection.Find.Found Then
            Selection.Rows(1).Delete
        End If
    
        SearchText ("Shareholders")
        If Selection.Find.Found Then
            Selection.GoToNext wdGoToTable
        With Selection.Tables(1)
            CountCNShare = 1
            For r = 2 To .Rows.Count - 2
                .cell(r, 1).Range.Text = .cell(r, 1).Range.Text & ChineseNameShareHolders1(CountCNShare)
                CountCNShare = CountCNShare + 1
            Next r
        End With
        
        Selection.GoToNext wdGoToTable
        With Selection.Tables(1)
            CountCNShare2 = 1
            For r = 2 To .Rows.Count
                .cell(r, 1).Range.Text = .cell(r, 1).Range.Text & ChineseNameShareHolders2(CountCNShare2)
                CountCNShare2 = CountCNShare2 + 1
            Next r
        End With
        
        SearchText "LAYERS OF SHAREHOLDING", True, True
        If Selection.Find.Found Then
            Selection.Rows(1).Delete
        End If
                
        If InStr(1, companyName, "(Given by Official Sources)", vbTextCompare) > 0 Then
            .Bookmarks("SourceEnglishName").Range.Text = "Subject's official English name is as reflected in the report."
        ElseIf InStr(1, companyName, "(Literal Translation)", vbTextCompare) > 0 Then
            .Bookmarks("SourceEnglishName").Range.Text = "No official English name was found during the course of investigation. The literally translated English name is provided in the report."
        ElseIf InStr(1, companyName, "(Given by Subject's Homepage)", vbTextCompare) > 0 Then
            .Bookmarks("SourceEnglishName").Range.Text = "No official English name was found during the course of investigation. The English name obtained from Subject's homepage is adopted in the report."
        ElseIf InStr(1, companyName, "(Given by Subject)", vbTextCompare) > 0 Then
            .Bookmarks("SourceEnglishName").Range.Text = "No official English name was found during the course of investigation. The English name provided by Subject is adopted in the report."
        ElseIf InStr(1, companyName, "(Given by the Client)", vbTextCompare) > 0 Then
            .Bookmarks("SourceEnglishName").Range.Text = "No official English name was found during the course of investigation. The supplied English name is adopted in the report."
        End If
    End If
    
      For Each bkm In templatedoc.Bookmarks
        bkm.Delete
      Next
     
      newFileName = destinationFolder & companyNo & "_CHN_" & EnglishTitle & ".docx"
      Debug.Print "newFileName: " & newFileName
      templatedoc.SaveAs2 FileName:=newFileName, Fileformat:=wdFormatXMLDocument
      Debug.Print "destinationFolder: " & destinationFolder
    
    templatedoc.Close savechanges:=False
    
End With

End If

End Sub

Function RemoveExtSpacing(ByVal Target As String) As String
    RemoveExtSpacing = Replace(RemoveSubString(Target, ""), Chr(13), vbNullString)
End Function
Function RemoveSubString(ByVal Target As String, pattern As String) As String
    RemoveSubString = Replace(Target, pattern, vbNullString)
End Function
Function SearchText(ByVal ExtText As String, Optional ByVal Bolded As Boolean = True, Optional ByVal UpperCase As Boolean = True)
    With Selection.Find
        .ClearFormatting
        .Font.Bold = Bolded
        .MatchCase = UpperCase
        .Text = ExtText
        .Execute Wrap:=wdFindContinue
    End With
End Function
Function FindFileEXT(Optional Lang As String = vbNullString) As String

Dim openFDialog As FileDialog
Set openFDialog = Application.FileDialog(msoFileDialogOpen)

If Lang <> vbNullString Then
    Des2 = Lang & " Agent Copy."
Else
    Des2 = "Template File."
End If

With openFDialog
    .Title = "Please select file for " & Des2
    .ButtonName = "I Select " & Des2
    .AllowMultiSelect = False
    .Filters.Clear
    .Filters.Add "doc files", "*.doc"
    If .Show = 0 Then
        Exit Function
    Else
        .Execute
        FindFileEXT = .SelectedItems(1)
    End If
End With

End Function
