<div align="center">

## Fax


</div>

### Description

I was looking for code to fax on PSC. found some fax stuff but people had some problems. Hopefully this solves it. It worked for me. Put this code in a Class Module.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bazzapr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bazzapr.md)
**Level**          |Advanced
**User Rating**    |4.5 (18 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bazzapr-fax__1-40131/archive/master.zip)





### Source Code

```
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Function Fax(ByVal FileName As String, ByVal FaxNumber As String)
Dim FaxServer As Object
Dim FaxDoc As Object
Dim ComputerName As String
  ComputerName = String(50, Chr(0))
  Call GetComputerName(ComputerName, 50)
  ComputerName = Trim(Replace(ComputerName, Chr(0), ""))
  Set FaxServer = CreateObject("FaxServer.FaxServer")
  FaxServer.Connect ("\\" & ComputerName)
  Set FaxDoc = FaxServer.CreateDocument(FileName)
  With FaxDoc
    .FaxNumber = FaxNumber
    .Send
  End With
  Set FaxDoc = Nothing
  Set FaxServer = Nothing
End Function
```

