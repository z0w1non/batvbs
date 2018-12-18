'Get the path of a item dropped.
path = WScript.Arguments(0)
WScript.StdOut.WriteLine(path)

'Show contents of the text file while prepending line number.
Set fso = CreateObject("Scripting.FileSystemObject")
With fso.OpenTextFile(path)
    row = 1
    Do Until .AtEndOfStream
        line = .ReadLine()
        WScript.StdOut.WriteLine Replace(Space(4 - Len(row)), Space(1), "0") & row & ": " & line
        row = row + 1
    Loop
End With
