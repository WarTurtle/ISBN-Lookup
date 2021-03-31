Sub createTable()
'
' Macro2 Macro
'
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Value"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "ISBN"
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "2"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "9781101984598"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "9781101984598"
    Range("B1").Select
    Application.CutCopyMode = False
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$B$3"), , xlYes).Name = _
        "qTable"
End Sub


Sub books()
'
' books Macro
' Run createTable to set up the query table
'    createTable
'
' You will need to change the privacy seetings of your workbook.
' Your full table will be in the final merge query
' Data>Queries>Merge#


    Dim nextQuery As String
    Dim rowsInQuery As Integer
    rowsInQuery = ActiveSheet.ListObjects("qTable").DataBodyRange.Rows.Count()
    Dim formula1 As String
    Dim formula2 As String
    Dim formula3 As String
    Dim formula4 As String
    Dim formula5 As String
    Dim formula6 As String
    Dim formula7 As String
    Dim formulaQTable As String
    
    Dim pq As Object
    For Each pq In ThisWorkbook.Queries
       pq.Delete
    Next
    ' Name Your sheet Books if you want.
    Sheets("Books").Select
    
    ' Deletes the ISBN table when you want to recreate it
    ' ActiveSheet.ListObjects(CellInTable(ActiveSheet.Range("C1"))).Delete
    

    
    ActiveWorkbook.Queries.Add Name:="Query1", Formula:= _
        "let Parameter=(TableName,ParameterLabel) =>" & Chr(10) & " " & Chr(10) & "let" & Chr(10) & "   Source = Excel.CurrentWorkbook(){[Name=TableName]}[Content]," & Chr(10) & " value = Source{[Value=ParameterLabel]}[ISBN]" & Chr(10) & "in" & Chr(10) & "    value" & Chr(10) & "" & Chr(10) & "in Parameter"

    For I = 1 To rowsInQuery
        formula1 = "let" & Chr(10)
        formula2 = "    ISBNraw =Query1(""qTable""," & I & ")," & Chr(10)
        formula3 = "    ISBN = if ISBNraw=null then ""null"" else if Value.Is(ISBNraw,Int64.Type) then ""isbn+"" & Number.ToText(ISBNraw) else try ""isbn+"" & Number.ToText(Text.ToNumber(ISBNraw)) otherwise ISBNraw, " & Chr(10)
        formula4 = "    Source = Json.Document(Web.Contents(""https://www.googleapis.com/books/v1/volumes?q="" & ""isbn+"" &  ISBN))," & Chr(10)
        formula5 = "    #""Converted to Table"" = Record.ToTable(Source)," & Chr(10)
        formula6 = "    Value = #""Converted to Table""{2}[Value]," & Chr(10)
        formula7 = "    Value1 = Value{0}," & Chr(10)
        formula8 = "    volumeInfo = Value1[volumeInfo]," & Chr(10)
        formula9 = "    #""Converted to Table1"" = Record.ToTable(volumeInfo)," & Chr(10) & ""
        formula10 = "     #""Check for Empty"" = if ISBN=""null"" then  Table.FromRecords(" & Chr(10)
        formula11 = "{" & Chr(10)
        formula12 = "[Name = ""title"", Value = ""-----------"" ]," & Chr(10)
      formula13 = "[Name = ""subtitle"", Value = ""----------"" ]," & Chr(10)
      formula14 = "[Name = ""authors"", Value = {[ Author = ""------""]} ]," & Chr(10)
      formula15 = "[Name = ""publisher"", Value = ""-------"" ]," & Chr(10)
      formula16 = "[Name = ""publishedDate"", Value = ""0000"" ]," & Chr(10)
      formula17 = "[Name = ""description"", Value = ""---------"" ]," & Chr(10)
      formula18 = "[Name = ""industryIdentifiers"", Value = {[ type = ""ISBN_13"", ISBN13=""0000000000000""], [type = ""ISBN_10"", ISBN10 = ""0000000""]} ]," & Chr(10)
      formula19 = "[Name = ""readingModes"", Value = [somethin = ""------""] ]," & Chr(10)
      formula20 = "[Name = ""pageCount"", Value = 0 ]," & Chr(10)
      formula21 = "[Name = ""printType"", Value = ""------"" ]," & Chr(10)
      formula22 = "[Name = ""categories"", Value = [ somethin = ""-------""] ]," & Chr(10)
      formula23 = "[Name = ""averageRating"", Value = 0 ]," & Chr(10)
      formula24 = "[Name = ""ratingsCount"", Value = 0 ]," & Chr(10)
      formula25 = "[Name = ""maturityRating"", Value = ""-----------"" ]," & Chr(10)
      formula26 = "[Name = ""allowAnonLogging"", Value = ""False"" ]," & Chr(10)
      formula27 = "[Name = ""contentVersion"", Value = ""------------"" ]," & Chr(10)
      formula28 = "[Name = ""panelizationSummary"", Value = [something = ""-------""]]," & Chr(10)
      formula29 = "[Name = ""imageLinks"", Value = [ somethin = ""-------"" ]]," & Chr(10)
      formula30 = "[Name = ""language"", Value = ""--"" ]," & Chr(10)
      formula31 = "[Name = ""previewLink"", Value = ""----------"" ]," & Chr(10)
      formula32 = "[Name = ""infoLink"", Value = ""----------"" ]" & Chr(10)
    formula33 = "}" & Chr(10)
  formula34 = ")" & Chr(10)
  formula35 = "else #""Converted to Table1"""
        formula36 = "in" & Chr(10)
        formula37 = "    #""Check for Empty"""
        formulaQueries = formula1 & formula2 & formula3 & formula4 & formula5 & formula6 & formula7 & formula8 & formula9 & formula10 & formula11 & formula12 & formula13 & formula14 & formula15 & formula16 & formula17 & formula18 & formula19 & formula20 & formula21 & formula22 & formula23 & formula24 & formula25 & formula26 & formula27 & formula28 & formula29 & formula30 & formula31 & formula32 & formula33 & formula34 & formula35 & formula36 & formula37

    
    ActiveWorkbook.Queries.Add Name:="APIQuery" & I, Formula:= _
        formulaQueries
        
    Next
    
    For I = 1 To rowsInQuery
        If I = 1 Then
            ActiveWorkbook.Queries.Add Name:="Merge1", Formula:= _
                "let" & Chr(13) & "" & Chr(10) & "    Source = Table.NestedJoin(APIQuery1,{""Name""},APIQuery2,{""Name""},""APIQuery2"",JoinKind.FullOuter)," & Chr(13) & "" & Chr(10) & "    #""Expanded APIQuery2"" = Table.ExpandTableColumn(Source, ""APIQuery2"", {""Value""}, {""APIQuery2.Value""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Expanded APIQuery2"""
        ElseIf I = rowsInQuery Then
        ElseIf I = rowsInQuery - 1 Then
            Dim finalMerge As String
            finalMerge = "Merge" & I
            nextQuery = "APIQuery" & I + 1
            
                formula1 = "let" & Chr(13) & "" & Chr(10) & "    Source = Table.NestedJoin(" & "Merge" & I - 1 & ",{""Name""}," & nextQuery & ",{""Name""},""" & nextQuery & """,JoinKind.FullOuter)," & Chr(13) & "" & Chr(10)
                formula2 = "    #""Expanded " & nextQuery & """ = Table.ExpandTableColumn(Source, """ & nextQuery & """, {""Value""}, {""" & nextQuery & ".Value""})," & Chr(13) & "" & Chr(10)
                formula3 = "    #""Filtered Rows"" = Table.SelectRows(#""Expanded " & nextQuery & """, each ([Name] = ""authors"" or [Name] = ""categories"" or [Name] = ""description"" or [Name] = ""industryIdentifiers"" or [Name] = ""pageCount"" or [Name] = ""publishedDate"" or [Name] = ""title""))," & Chr(13) & "" & Chr(10)
                formula4 = "    #""Transposed Table"" = Table.Transpose(#""Filtered Rows"")," & Chr(13) & "" & Chr(10)
                formula5 = "    #""Promoted Headers"" = Table.PromoteHeaders(#""Transposed Table"", [PromoteAllScalars=true])," & Chr(13) & "" & Chr(10)
                formula6 = "    #""Extracted First Characters"" = Table.TransformColumns(#""Promoted Headers"", {{""publishedDate"", each Text.Start(_, 4), type text}})," & Chr(10)
                formula7 = "    #""Changed Type"" = Table.TransformColumnTypes(#""Extracted First Characters"",{{""title"", type text}, {""authors"", type any}, {""publishedDate"", Int64.Type}, {""description"", type text}, {""industryIdentifiers"", type any}, {""pageCount"", Int64.Type}, {""categories"", type any}})," & Chr(13) & "" & Chr(10)
                formula8 = "    #""Extracted Values"" = Table.TransformColumns(#""Changed Type"", {""authors"", each Text.Combine(List.Transform(_, Text.From), "", ""), type text})," & Chr(13) & "" & Chr(10)
                formula9 = "    #""Expanded industryIdentifiers"" = Table.ExpandListColumn(#""Extracted Values"", ""industryIdentifiers"")," & Chr(13) & "" & Chr(10)
                formula10 = "   #""Expanded industryIdentifiers1"" = Table.ExpandRecordColumn(#""Expanded industryIdentifiers"", ""industryIdentifiers"", {""type"", ""identifier""}, {""industryIdentifiers.type"", ""industryIdentifiers.identifier""})," & Chr(13) & "" & Chr(10)
                formula11 = "   #""Filtered Rows1"" = Table.SelectRows(#""Expanded industryIdentifiers1"", each ([industryIdentifiers.type] <> ""ISBN_10""))," & Chr(13) & "" & Chr(10)
                formula12 = "   #""Removed Columns"" = Table.RemoveColumns(#""Filtered Rows1"",{""industryIdentifiers.type""})," & Chr(13) & "" & Chr(10)
                formula13 = "   #""Extracted Values1"" = Table.TransformColumns(#""Removed Columns"", {""categories"", each Text.Combine(List.Transform(_, Text.From), "", ""), type text})," & Chr(13) & "" & Chr(10)
                formula14 = "   #""Reordered Columns"" = Table.ReorderColumns(#""Extracted Values1"",{""industryIdentifiers.identifier"", ""title"", ""authors"", ""publishedDate"", ""pageCount"", ""categories"", ""description""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Reordered Columns"""
    
        formulaQTable = formula1 & formula2 & formula3 & formula4 & formula5 & formula6 & formula7 & formula8 & formula9 & formula10 & formula11 & formula12 & formula13 & formula14
            
            ActiveWorkbook.Queries.Add Name:=finalMerge, Formula:= _
                formulaQTable
                
            With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
                "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & finalMerge & ";Extended Properties=""""" _
                , Destination:=Range("$C$1")).QueryTable
                .CommandType = xlCmdSql
                .CommandText = Array("SELECT * FROM [" & finalMerge & "]")
                .RowNumbers = False
                .FillAdjacentFormulas = False
                .PreserveFormatting = True
                .RefreshOnFileOpen = False
                .BackgroundQuery = True
                .RefreshStyle = xlInsertDeleteCells
                .SavePassword = False
                .SaveData = True
                .AdjustColumnWidth = True
                .RefreshPeriod = 0
                .PreserveColumnInfo = True
                .ListObject.DisplayName = finalMerge

    End With
        Else
            nextQuery = "APIQuery" & I + 1
            ActiveWorkbook.Queries.Add Name:="Merge" & I, Formula:= _
                "let" & Chr(13) & "" & Chr(10) & "    Source = Table.NestedJoin(Merge" & I - 1 & ",{""Name""}," & nextQuery & ",{""Name""},""" & nextQuery & """,JoinKind.FullOuter)," & Chr(13) & "" & Chr(10) & "    #""Expanded " & nextQuery & """ = Table.ExpandTableColumn(Source, """ & nextQuery & """, {""Value""}, {""" & nextQuery & ".Value""})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Expanded " & nextQuery & """"
        End If
    Next
            ActiveWorkbook.RefreshAll
End Sub

Function CellInTable(thisCell As Range) As String
    Dim tableName As String
    tableName = ""
    On Error Resume Next
    tableName = thisCell.ListObject.Name
    CellInTable = tableName
End Function
