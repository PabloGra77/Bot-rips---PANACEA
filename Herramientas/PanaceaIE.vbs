' Simple lanzador usando Internet Explorer (si está disponible en el sistema)
On Error Resume Next
Dim ie
Set ie = CreateObject("InternetExplorer.Application")
If Err.Number <> 0 Or IsEmpty(ie) Then
  WScript.Echo "No se pudo iniciar Internet Explorer en este equipo."
  WScript.Quit 1
End If
On Error GoTo 0

ie.Visible = True
ie.MenuBar = False
ie.ToolBar = False
ie.AddressBar = False
ie.Navigate "http://181.51.196.194/Panacea"

Do While ie.Busy
  WScript.Sleep 100
Loop
