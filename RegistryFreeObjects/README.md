# RegFree

## Metadata

Source: https://www.vbforums.com/showthread.php?910948-Formal-Request-to-Olaf-to-migrate-RichClient-5-6-to-64bit-operation&p=5681753&viewfull=1#post5681753
Author: VanGoghGaming

## Post

> It should be pretty straightforward to make it work in 64-bit (just replace some Longs with LongPtr for pointer types) 
> although I haven't tested it in twinBASIC to confirm this.

Confirmed, this reg-free technique also works in 64-bit:

```vb
Dim objRegExp As Object, objMatch As Object, objDictionary As Object, objFSO As Object

Private Sub Main()
    #If Win64 Then
      const sys = "C:\Windows\SysWow64"
    #Else
      const sys = "C:\Windows\System32"
    #End If
    
    With New cRegFree
      Set objRegExp = .CreateObj("RegExp", sys & "\vbscript.dll", 3)
    End With

    With New cRegFree
      Set objDictionary = .CreateObj("Dictionary", sys & "\scrrun.dll")
      Set objFSO = .CreateObj("FileSystemObject")
    End With

    With objRegExp
        .Global = True
        .IgnoreCase = True
        .Pattern = "the"
        For Each objMatch In .Execute("The quick brown fox jumps over the lazy dog")
            objDictionary.Add objMatch.FirstIndex, objMatch
        Next objMatch
    End With

    Dim vKey As Variant
    With objDictionary
        For Each vKey In .Keys
            Debug.Print vKey, .Item(vKey)
        Next vKey
    End With

    Debug.Print """Windows"" folder exists:", objFSO.FolderExists("C:\Windows")
End Sub
```