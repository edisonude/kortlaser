VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvoiceDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'Clase cInvoiceDetail, mapea directamente con la tabla invoice_detail para la gestión de los detalles de las facturas

'Attributes
Public id As Integer
Public id_invoice As Integer
Public description As String
Public detail_value As Double

Public Sub load(id As Integer, id_invoice As Integer, description As String, detail_value As Double)
    Me.id = id
    Me.id_invoice = id_invoice
    Me.description = description
    Me.detail_value = detail_value
End Sub


