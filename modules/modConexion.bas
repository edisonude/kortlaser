Attribute VB_Name = "modConexion"
Option Explicit

'Variables de configuracion
Public SQL As String
Public pathBD As String
Public keyBD As String

'Almacena la ruta del archivo de configuraciones
Dim fileConfigPath As String

'Conexion ADOB
'Public mysqlCon As New ADODB.Connection
'Public properties As New CProperty

Sub Main()

On Local Error GoTo Control
frmMenu.Show
Exit Sub
Control:
MsgBox "error inesperado"
End Sub

Public Function getNewConection() As ADODB.Connection
'Se lee la configuracion de conexion a la base de datos
fileConfigPath = left(App.Path, InStrRev(App.Path, "\")) & "config.ini"

keyBD = modFiles.readPropertyFile(fileConfigPath, "pass", "")
pathBD = modFiles.readPropertyFile(fileConfigPath, "pathBD", "")

Dim myCon As New ADODB.Connection
With myCon
    .CursorLocation = adUseClient
    .Open "Provider=Microsoft.Jet.OLEDB.4.0; " & _
            "Data Source=" & pathBD & ";" & _
            "Jet OLEDB:Database Password=" & keyBD & ""
End With

Set getNewConection = myCon
End Function




