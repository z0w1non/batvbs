'Parse the command line arguments.
bookpath = Replace(WScript.Arguments(0), """", "")
target_color = Eval(Replace(WScript.Arguments(1), """", ""))
replacement_color = Eval(Replace(WScript.Arguments(2), """", ""))

Set excel = CreateObject("Excel.Application")

'Replace the target color to the replacement color each all cell of all sheet of the workbook.
With excel.Workbooks.Open(bookpath)
    For Each worksheet In .Worksheets
        For Each cell In worksheet.UsedRange.Cells
            If cell.Font.Color = target_color Then
                cell.Font.Color = replacement_color
            End If
            If cell.Interior.Color = target_color Then
                cell.Interior.Color = replacement_color
            End If
            For Each border In cell.Borders
                If border.Color = target_color Then
                    border.Color = replacement_color
                End If
            Next
        Next
    Next
    .Save
    .Close True
End With
