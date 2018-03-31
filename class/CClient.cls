VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Clase CClient, mapea directamente con la tabla client para la gestión de clientes

'Attributes
Public id As Integer
Public document As String
Public name As String
Public phone As String

Public Sub loadClient(id As Integer, document As String, name As String, phone As String)
    Me.id = id
    Me.document = document
    Me.name = name
    Me.phone = phone
End Sub

'Encuentra un cliente por su documento
Public Function findByDocument(document As String) As cClient
Dim conBd As ADODB.Connection
Dim rec As New ADODB.Recordset

'Se solicita una conexion a la bd
Set conBd = modConexion.getNewConection
rec.CursorLocation = adUseClient

rec.Open "select * from client where document='" & document & "'", conBd, adOpenStatic, adLockOptimistic
'rec.Open "select * from client", conBd, adOpenStatic, adLockOptimistic

Dim client As cClient
If (rec.RecordCount >= 1) Then
    Set client = New cClient
    Call client.loadClient(rec("id"), rec("document"), rec("name"), rec("phone"))
End If

Set findByDocument = client
rec.Close
End Function

'Encuentra un cliente por su documento
Public Function findById(id As Integer) As cClient
Dim conBd As ADODB.Connection
Dim rec As New ADODB.Recordset

'Se solicita una conexion a la bd
Set conBd = modConexion.getNewConection
rec.CursorLocation = adUseClient

rec.Open "select * from client where id=" & id & "", conBd, adOpenStatic, adLockOptimistic

Dim client As cClient
If (rec.RecordCount >= 1) Then
    Set client = New cClient
    Call client.loadClient(rec("id"), rec("document"), rec("name"), rec("phone"))
End If

Set findById = client
rec.Close
End Function
