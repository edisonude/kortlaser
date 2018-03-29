VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cInvoicePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Clase cInvoicePayment, mapea directamente con la tabla invoice_payment para la gestión de los pagos de una factura

'Attributes
Public id As Integer
Public id_invoice As Integer
Public date_payment As Date
Public payment_value As Double
Public payment_change As Double
Public residue_value As Double


Public Sub load(id As Integer, id_invoice As Integer, date_payment As Date, payment_value As Double, payment_change As Double, residue_value As Double)
    Me.id = id
    Me.id_invoice = id_invoice
    Me.date_payment = date_payment
    Me.payment_value = payment_value
    Me.payment_change = payment_change
    Me.residue_value = residue_value
End Sub



