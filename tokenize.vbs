Set word = CreateObject("VBScript.RegExp")
With word
    .Pattern = "\b\w+\b"
    .Global = True
End With

With WScript.StdIn
    Do Until .AtEndOfStream
        For Each match In word.Execute(.ReadLine())
            WScript.StdOut.WriteLine match.Value
        Next
    Loop
End With
