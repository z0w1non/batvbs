word = Replace(WScript.Arguments(0), """", "")
target = Replace(WScript.Arguments(1), """", "")

Set isMatch = CreateObject("VBScript.RegExp")
With isMatch
    .Pattern = word
    .IgnoreCase = True
End With

Set isExcelBookPath = CreateObject("VBScript.RegExp")
With isExcelBookPath
    .Pattern = "\.xls[xm]*$"
    .IgnoreCase = True
End With

Set excel = CreateObject("Excel.Application")

Set fso = CreateObject("Scripting.FileSystemObject")

Search target

Sub Search(path)
    If fso.FileExists(path) Then
        SearchFile path
    ElseIf fso.FolderExists(path) Then
        RecursiveSearchFolder path
    End If
End Sub

Sub SearchFile(path)
    If isExcelBookPath.Test(path) Then
        SearchExcelBook path
    Else
        SearchTextFile path
    End If
End Sub

Sub RecursiveSearchFolder(path)
    With fso.GetFolder(path)
        For Each file In .Files
            SearchFile file
        Next
        For Each subfolder In .SubFolders
            RecursiveSearchFolder subfolder
        Next
    End With
End Sub

Sub SearchExcelBook(path)
    With excel.Workbooks.Open(path)
        For Each worksheet In .Worksheets
            For Each cell In worksheet.UsedRange.Cells
                If isMatch.Test(cell.Value) Then
                    WScript.StdOut.WriteLine path & "(" & worksheet.Name & ":" & cell.Address & "):" & cell.Value
                End If
            Next
        Next
        .Close
    End With
End Sub

Sub SearchTextFile(path)
    With fso.OpenTextFile(path)
        row = 1
        Do Until .AtEndOfStream
            line = .ReadLine()
            If isMatch.Test(line) Then
                WScript.StdOut.WriteLine path & "(" & row & "):" & line
            End If
            row = row + 1
        Loop
        .Close
    End With
End Sub
