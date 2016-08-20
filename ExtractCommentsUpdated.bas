Public Sub ExtractCommentsToNewDoc() 
 
 'Original source Macro created 2007 by Lene Fredborg, DocTools
 'amended by Phil Thomas May 2012 [URL="http://www.grosvenorsc.com"]www.grosvenorsc.com[/URL]
 'with help from Frosty at VBAExpress.com
     
 
 
 'The macro creates a new document
 'and extracts all comments from the active document
 'incl. metadata
 
 'Minor adjustments are made to the styles used
 'You may need to change the style settings and table layout to fit your needs
 '=========================
 'Setup all the variables
    Dim oDoc As Document 
    Dim oNewDoc As Document 
    Dim oTable As Table 
    Dim nCount As Long 
    Dim n As Long 
    Dim Title As String 
 
 'Setup intital values
    Title = "Extract All Comments to New Document" 
    Set oDoc = ActiveDocument 
    nCount = ActiveDocument.Comments.Count 
 
 'Check if document has any comments in it and if it does, then check this is what the user wants to do
    If nCount = 0 Then 
        MsgBox "The active document contains no comments.", vbOKOnly, Title 
        GoTo ExitHere 
    Else 
 'Stop if user does not click Yes
        If MsgBox("Do  you want to extract all comments to a new document?", _ 
                vbYesNo + vbQuestion, Title) <> vbYes Then 
            GoTo ExitHere 
        End If 
    End If 
 
 'Turned on as recommendation from MSDN Technet article
    Application.ScreenUpdating = True 
 'Create a new document for the comments, based on Normal.dot
    Set oNewDoc = Documents.Add 
 'Set to landscape
    oNewDoc.PageSetup.Orientation = wdOrientLandscape 
 'Insert a 10-column table for the comments
    With oNewDoc 
        .Content = "" 
        Set oTable = .Tables.Add _ 
            (Range:=Selection.Range, _ 
            NumRows:=nCount + 1, _ 
            NumColumns:=9) 
    End With 
 
 'Insert info in header - change date format as you wish
    oNewDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _ 
        "Document Review Record - " & "Comments extracted from: " & oDoc.Name & vbCr & _ 
        "Created by: " & Application.UserName & _ 
        " Creation date: " & Format(Date, "MMMM d, yyyy") & _ 
        "  - All page and line numbers are with Final: Show Markup turned on" 
 'insert page number into footer
    oNewDoc.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight 
 
 
 'Adjust the Normal style and Header style
    With oNewDoc.Styles(wdStyleNormal) 
        .Font.Name = "Arial" 
        .Font.Size = 8 
        .ParagraphFormat.LeftIndent = 0 
        .ParagraphFormat.SpaceAfter = 6 
    End With 
 
    With oNewDoc.Styles(wdStyleHeader) 
        .Font.Size = 8 
        .ParagraphFormat.SpaceAfter = 0 
    End With 
 
 'Format the table appropriately
    With oTable 
        .AllowAutoFit = False 
        .Style = "Table Grid" 
        .PreferredWidthType = wdPreferredWidthPercent 
        .PreferredWidth = 100 
        .Columns(1).PreferredWidth = 5 
        .Columns(2).PreferredWidth = 20 
        .Columns(3).PreferredWidth = 5 
        .Columns(4).PreferredWidth = 5 
        .Columns(5).PreferredWidth = 20 
        .Columns(6).PreferredWidth = 20 
        .Columns(7).PreferredWidth = 10 
        .Columns(8).PreferredWidth = 15 
        .Columns(8).Shading.BackgroundPatternColor = -570359809 
        .Columns(9).PreferredWidth = 20 
        .Columns(9).Shading.BackgroundPatternColor = -570359809 
        .Rows(1).HeadingFormat = True 
    End With 
 
 'Insert table headings
    With oTable.Rows(1) 
        .Range.Font.Bold = True 
        .Shading.BackgroundPatternColor = 5296274 
        .Cells(1).Range.Text = "Comment" 
        .Cells(2).Range.Text = "Section Heading" 
        .Cells(3).Range.Text = "Page" 
        .Cells(4).Range.Text = "Line on Page" 
        .Cells(5).Range.Text = "Comment scope" 
        .Cells(6).Range.Text = "Comment text" 
        .Cells(7).Range.Text = "Author" 
        .Cells(8).Range.Text = "Response Summary (Accept/ Reject/ Defer)" 
        .Cells(9).Range.Text = "Response to comment" 
    End With 
 
 'Repaginate - Start MSDN bug fix on Knowledgebase article
    ActiveDocument.Repaginate 
 
 'Toggle nonprinting characters twice
    ActiveWindow.ActivePane.View.ShowAll = Not _ 
      ActiveWindow.ActivePane.View.ShowAll 
 
    ActiveWindow.ActivePane.View.ShowAll = Not _ 
      ActiveWindow.ActivePane.View.ShowAll 
 'End MSDN KB code
 
 
 'Get info from each comment from oDoc and insert in table, no way to currently insert Criticality of comment.
 'Suggest either done afterwards or simply include C,S,M in the start of each comment and remove Col 8
    For n = 1 To nCount 
        With oTable.Rows(n + 1) 
            .Cells(1).Range.Text = n 
 'call function to get section heading
            .Cells(2).Range.Text = fGetNearestParaTextStyledIn(oDoc.Comments(n).Scope) 
 'Page number
            .Cells(3).Range.Text = oDoc.Comments(n).Scope.Information(wdActiveEndPageNumber) 
 ' The line number
            .Cells(4).Range.Text = oDoc.Comments(n).Scope.Information(wdFirstCharacterLineNumber) 
 'The text marked by the comment
            .Cells(5).Range.Text = oDoc.Comments(n).Scope 
 'The comment itself
            .Cells(6).Range.Text = oDoc.Comments(n).Range.Text 
 'The comment author
            .Cells(7).Range.Text = oDoc.Comments(n).Author 
        End With 
    Next n 
 
    Application.ScreenUpdating = True 
    Application.ScreenRefresh 
 
 'Tell them its finished
    oNewDoc.Activate 
    MsgBox nCount & " comments found. Finished creating comments document.", vbOKOnly, Title 
 
ExitHere: 
    Set oDoc = Nothing 
    Set oNewDoc = Nothing 
    Set oTable = Nothing 
 
End Sub 
 
 
 '----------------------------------------------------------------------------------------------------------
 ' COMPLEX SEARCH METHOD:
 ' Uses the Find object (which is always faster) to search an array of style names
 ' and return the text of the paragraph nearest to the original range
 '----------------------------------------------------------------------------------------------------------
Public Function fGetNearestParaTextStyledIn(Optional rngOriginal As Range, _ 
    Optional sStyleNames As String = "Heading 1|Heading 2|Heading 3", _ 
    Optional bLookDown As Boolean = False, _ 
    Optional bIncludeParagraphMark As Boolean = False) As String 
 
    Dim oDoc As Document 
    Dim aryStyleNames() As String 
    Dim colFoundRanges As Collection 
    Dim rngReturn As Range 
    Dim i As Integer 
    Dim sReturnText As String 
    Dim s1ReturnText As String 
    Dim s2ReturnText As String 
    Dim lDistance As Long 
 
    On Error GoTo l_err 
 'set a default if we didn't pass it
    If rngOriginal Is Nothing Then 
        Set rngOriginal = Selection.Range.Duplicate 
    End If 
 
 'create a new instance of a collection
    Set colFoundRanges = New Collection 
 
 'get our array of style names to look for
    aryStyleNames = Split(sStyleNames, "|") 
 
 'loop through the array
    For i = 0 To UBound(aryStyleNames) 
 'if you wanted to add additional styles, you could change the optional parameter, or
 'pass in different values
        Set rngReturn = fGetNearestParaRange(rngOriginal.Duplicate, aryStyleNames(i), bLookDown) 
 'if we found it in the search direction
        If Not rngReturn Is Nothing Then 
 'then add it to the collection
            colFoundRanges.Add rngReturn 
        End If 
    Next 
 
 'if we found anything in our collection, then we can go through it,
 'and see which range is closest to our original range, depending on our search direction
    If colFoundRanges.Count > 0 Then 
 'start with an initial return
        Set rngReturn = colFoundRanges(1) 
 'and an initial distance value as an absolute number
        lDistance = Abs(rngOriginal.Start - rngReturn.Start) 
 'then go through the rest of them, and return the one with the lowest distance between
        For i = 2 To colFoundRanges.Count 
            If lDistance > Abs(rngOriginal.Start - colFoundRanges(i).Start) Then 
 'set a new range
                Set rngReturn = colFoundRanges(i) 
 'and a new distance test
                lDistance = Abs(rngOriginal.Start - rngReturn.Start) 
            End If 
        Next 
 
 'now get the text we're going to return
        s1ReturnText = rngReturn.ListFormat.ListString 
        s2ReturnText = rngReturn.Text 
        sReturnText = s1ReturnText & " - " & s2ReturnText 
 'and whether to include the paragraph mark
        If bIncludeParagraphMark = False Then 
            sReturnText = Replace(sReturnText, vbCr, "") 
        End If 
    End If 
 
 
l_exit: 
    fGetNearestParaTextStyledIn = sReturnText 
    Exit Function 
l_err: 
 'black box, so that any errors return an empty string
    sReturnText = "" 
    Resume l_exit 
End Function 
 '----------------------------------------------------------------------------------------------------------
 'return the nearest paragraph range styled
 'defaults to Heading 1
 'NOTE: if searching forward, starts searching from the *beginning* of the passed range
 '      if searching backward, starts searching from the *end* of the passed range
 '----------------------------------------------------------------------------------------------------------
Public Function fGetNearestParaRange(rngWhere As Range, _ 
    Optional sStyleName As String = "Heading 1", _ 
    Optional bSearchForward As Boolean = False) As Range 
    Dim rngSearch As Range 
 
    On Error GoTo l_err 
    Set rngSearch = rngWhere.Duplicate 
 'if searching down, then start at the beginning of our search range
    If bSearchForward Then 
        rngSearch.Collapse wdCollapseStart 
 'otherwise, search from the end
    Else 
        rngSearch.Collapse wdCollapseEnd 
    End If 
 
 'find the range
    With rngSearch.Find 
        .Wrap = wdFindStop 
        .Forward = bSearchForward 
        .Style = sStyleName 
 'if we found it, return it
        If .Execute Then 
            Set fGetNearestParaRange = rngSearch 
        Else 
            Set fGetNearestParaRange = Nothing 
        End If 
    End With 
 
l_exit: 
    Exit Function 
l_err: 
 'black box- any errors, return nothing
    Set rngSearch = Nothing 
    Resume l_exit 
End Function 