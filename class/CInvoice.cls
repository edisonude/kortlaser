VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Clase cInvoice, mapea directamente con la tabla invoice para la gestión de facturas

'Attributes
Public id As Integer
Public id_client As Integer
Public date_invoice As Date
Public total_value As Double
Public residue_value As Double

Public Sub loadInvoice(id As Integer, id_client As Integer, date_invoice As Date, total_value As Double, residue_value As Double)
    Me.id = id
    Me.id_client = id_client
    Me.date_invoice = date_invoice
    Me.total_value = total_value
    Me.residue_value = residue_value
End Sub

'Encuentra la ultima factura guardad
Public Function findLastInvoice() As cInvoice
Dim conBd As ADODB.Connection
Dim rec As New ADODB.Recordset

'Se solicita una conexion a la bd
Set conBd = modConexion.getNewConection
rec.CursorLocation = adUseClient

rec.Open "Select Top 1 invoice.* from invoice order by 1 desc", conBd, adOpenStatic, adLockOptimistic

Dim invoice As cInvoice
If (rec.RecordCount >= 1) Then
    Set invoice = New cInvoice
    Call invoice.loadInvoice(rec("id"), rec("id_client"), rec("date_invoice"), rec("total_value"), rec("residue_value"))
End If

Set findLastInvoice = invoice
rec.Close
End Function

Public Sub addDetail(detail As cInvoiceDetail)
Dim conBd As ADODB.Connection
'Se solicita una conexion a la bd
Set conBd = modConexion.getNewConection

SQL = "INSERT INTO invoice_detail " & _
    "(id_invoice,description, detail_value) VALUES " & _
    "(" & detail.id_invoice & ",'" & detail.description & "'," & detail.detail_value & ")"
conBd.Execute (SQL)
End Sub

Public Sub addPayment(payment As cInvoicePayment)
Dim conBd As ADODB.Connection
'Se solicita una conexion a la bd
Set conBd = modConexion.getNewConection

SQL = "INSERT INTO invoice_payment " & _
    "(id_invoice,date_payment,payment_value,payment_change,residue_value) VALUES " & _
    "(" & payment.id_invoice & ",#" & modFormater.convertDateToAccesDate(payment.date_payment) & "#," & payment.payment_value & "," & payment.payment_change & "," & payment.residue_value & ")"
conBd.Execute (SQL)
End Sub

