Attribute VB_Name = "Ap"
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'Instancia de cliente utilitaria
Public cClient As New cClient

'Instancia de factura utilitaria
Public cInvoice As New cInvoice

'Instancia del men�
Public frmMenu As frmMenu

'Acci�n de busqueda
Public Const ACT_SEARCH = 1
