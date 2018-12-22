Const ForWriting = 2
Const ForAppending = 8

output = WScript.Arguments(0)
iomode = ForWriting
If WScript.Arguments.Count >= 2 Then
    If InStr(WScript.Arguments(1), "/a") > 0 Then
        iomode = ForAppending
    End If
End If

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(output) Then
    fso.CreateTextFile(output).Close
End If
Set output_stream = fso.OpenTextFile(output, iomode)
With WScript.StdIn
    Do Until .AtEndOfStream
        line = .ReadLine()
        output_stream.WriteLine line
        WScript.StdOut.WriteLine line
    Loop
End With
output_stream.Close
