VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInvoice 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12210
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tPhone 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   2
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox tResidueValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10185
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "$0"
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox tChangeValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "$0"
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox tPaymentValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2400
      TabIndex        =   17
      Text            =   "$0"
      Top             =   6840
      Width           =   1815
   End
   Begin VB.TextBox tDescription 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   7320
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
   End
   Begin VB.TextBox tValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7320
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.TextBox tDocument 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox tName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
   Begin MSComctlLib.ListView listData 
      Height          =   2415
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descripcion"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label tResidueValueTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00373436&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   10200
      TabIndex        =   25
      Top             =   6150
      Width           =   1815
   End
   Begin VB.Label label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo restante"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   8
      Left            =   8040
      TabIndex        =   24
      Top             =   6150
      Width           =   2070
   End
   Begin VB.Label tIdClient 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   23
      Top             =   840
      Width           =   255
   End
   Begin VB.Label label 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   12
      Left            =   240
      TabIndex        =   22
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Image Image6 
      Height          =   390
      Left            =   3720
      Picture         =   "frmInvoice.frx":0000
      Top             =   1200
      Width           =   405
   End
   Begin VB.Image Image5 
      Height          =   405
      Left            =   7800
      Picture         =   "frmInvoice.frx":08CA
      Top             =   7560
      Width           =   2025
   End
   Begin VB.Image cmdFinish 
      Height          =   405
      Left            =   9960
      Picture         =   "frmInvoice.frx":3414
      Top             =   7560
      Width           =   2025
   End
   Begin VB.Label label 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo restante"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   11
      Left            =   8160
      TabIndex        =   20
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label label 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor cambio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   9
      Left            =   4320
      TabIndex        =   18
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Line Line6 
      BorderColor     =   &H001A1A1A&
      X1              =   12000
      X2              =   240
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label label 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor pagado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   10
      Left            =   480
      TabIndex        =   16
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label label 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   7
      Left            =   8295
      TabIndex        =   15
      Top             =   5700
      Width           =   1815
   End
   Begin VB.Label tTotalValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00373436&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   10215
      TabIndex        =   14
      Top             =   5700
      Width           =   1815
   End
   Begin VB.Line Line5 
      BorderColor     =   &H003933ED&
      X1              =   8640
      X2              =   12000
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Image cmdAdd 
      Height          =   405
      Left            =   9960
      Picture         =   "frmInvoice.frx":5F5E
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Line Line4 
      BorderColor     =   &H001A1A1A&
      BorderWidth     =   3
      X1              =   5640
      X2              =   5640
      Y1              =   720
      Y2              =   2760
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   5760
      Picture         =   "frmInvoice.frx":8AA8
      Top             =   720
      Width           =   405
   End
   Begin VB.Label label 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label label 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   5
      Left            =   5760
      TabIndex        =   12
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label label 
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   4
      Left            =   6240
      TabIndex        =   11
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label label 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label label 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label label 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label label 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalles"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H003933ED&
      X1              =   240
      X2              =   1320
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "frmInvoice.frx":9372
      Top             =   720
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H001A1A1A&
      BorderWidth     =   5
      X1              =   0
      X2              =   15
      Y1              =   0
      Y2              =   7575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Facturas"
      BeginProperty Font 
         Name            =   "Adobe Gothic Std B"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H003933ED&
      BorderWidth     =   2
      X1              =   240
      X2              =   1680
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Conexión activa de BD  para asignar un servicio
Dim conBd As ADODB.Connection
Dim rec As New ADODB.Recordset

Dim paymentValue As Double
Dim netValue As Double
Dim changeValue As Double
Dim residueValue As Double

Private Sub cmdAdd_Click()
If Me.tDescription = "" Then
    MsgBox "Debe ingresar la descripción del trabajo realizado", vbCritical
    Me.tDescription.SetFocus
    Exit Sub
End If

If Me.tValue = "" Or modFormater.convertCurrencyToValue(Me.tValue) = 0 Then
    MsgBox "Debe ingresar el valor del trabajo realizado", vbCritical
    Me.tValue.SetFocus
    Exit Sub
End If

Set li = Me.listData.ListItems.Add(, , Me.tDescription)
    li.SubItems(1) = modFormater.convertValueToCurrency(Me.tValue, 0)

Call calculateTotal
Me.tDescription = ""
Me.tValue = ""
End Sub

Private Sub cmdFinish_Click()
If Not validateOrRegisterClient Then Exit Sub

If Me.listData.ListItems.Count < 1 Then
    MsgBox "No existen detalles de los trabajos para generar la factura", vbCritical
    Me.tDescription.SetFocus
    Exit Sub
End If

'Guarda la factura
Dim invoice As cInvoice
Set invoice = saveInvoice

If invoice Is Nothing Then
    MsgBox "No se pudo crear la factura", vbCritical, "Error administrador"
    Exit Sub
End If

'Agrega detalles a la factura
Dim item As Integer
For item = 1 To Me.listData.ListItems.Count
    Dim invoiceDetail As New cInvoiceDetail
    Call invoiceDetail.load(0, invoice.id, Me.listData.ListItems(item), Me.listData.ListItems(item).SubItems(1))
    invoice.addDetail invoiceDetail
Next

If Val(Me.tPaymentValue) > 0 Then
    Dim invoicePayment As New cInvoicePayment
    Call invoicePayment.load(0, invoice.id, Now(), paymentValue, changeValue, residueValue)
    invoice.addPayment invoicePayment
End If

Call imprimirFactura(invoice.id)

MsgBox "fin"

'SQL = "INSERT INTO job " & _
'    "(date_invoice, total_value, residue_value) VALUES " & _
'    "(#" & modFormater.convertDateToAccesDate(Now) & "#," & modFormater.convertCurrencyToValue(Me.tTotalValue) & "," & modFormater.convertCurrencyToValue(Me.tResidueValue) & ")"
'
'conBd.Execute (SQL)
End Sub

Private Function saveInvoice() As cInvoice
SQL = "INSERT INTO invoice " & _
    "(id_client,date_invoice, total_value, residue_value) VALUES " & _
    "('" & Me.tIdClient & "',#" & modFormater.convertDateToAccesDate(Now) & "#," & modFormater.convertCurrencyToValue(Me.tTotalValue) & "," & modFormater.convertCurrencyToValue(Me.tResidueValue) & ")"
conBd.Execute (SQL)
Sleep 800
Set saveInvoice = Ap.cInvoice.findLastInvoice
End Function

Private Function validateOrRegisterClient()
Dim client As cClient
Set client = Ap.cClient.findByDocument(Me.tDocument)
validateOrRegisterClient = False

If client Is Nothing Then
    'Si el cliente no existe valida que se pueda registrar
    
    If (Me.tDocument = "") Then
        MsgBox "Debe ingresar el documento del cliente", vbCritical
        Me.tDocument.SetFocus
        Exit Function
    End If
    
    If (Me.tName = "") Then
        MsgBox "Debe ingresar el nombre del cliente", vbCritical
        Me.tName.SetFocus
        Exit Function
    End If
    
    If (Me.tPhone = "") Then
        MsgBox "Debe ingresar el teléfono del cliente", vbCritical
        Me.tPhone.SetFocus
        Exit Function
    End If
    
    'Si se cumplen las validaciones se procede a crear el cliente
    SQL = "INSERT INTO client " & _
    "(document, name, phone) VALUES " & _
    "('" & Me.tDocument & "','" & Me.tName & "','" & Me.tPhone & "')"
    conBd.Execute (SQL)
    Sleep 800
    
    Set client = Ap.cClient.findByDocument(Me.tDocument)
End If

If client Is Nothing Then
    MsgBox "Error inesperado. No se pudo crear o cargar el cliente", vbCritical, "Error Administrador"
    Exit Function
Else
    Me.tIdClient = client.id
    validateOrRegisterClient = True
End If
    
End Function

Private Sub Form_Load()
Call createConexion

Dim ancho As Double
ancho = Me.listData.Width
Me.listData.ColumnHeaders(1).Width = ancho * 0.8
Me.listData.ColumnHeaders(2).Width = ancho * 0.19
End Sub

'Se solicita una conexion a la bd
Private Function createConexion()
Set conBd = modConexion.getNewConection
rec.CursorLocation = adUseClient
End Function

Private Sub listData_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 46 And Not Me.listData.SelectedItem Is Nothing) Then
    Me.listData.ListItems.Remove (Me.listData.SelectedItem.Index)
    Call calculateTotal
End If
End Sub

Private Sub tDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdAdd_Click
End If
End Sub

Private Sub tDocument_LostFocus()
Dim client As cClient
Set client = Ap.cClient.findByDocument(Me.tDocument)

If client Is Nothing Then
    Me.tIdClient = ""
    Me.tName = ""
    Me.tPhone = ""
Else
    Me.tIdClient = client.id
    Me.tDocument = client.document
    Me.tName = client.name
    Me.tPhone = client.phone
End If
End Sub

Private Sub Text5_Change()

End Sub

Private Sub tPaymentValue_Change()
Call calculateTotal
End Sub

Private Sub tPaymentValue_GotFocus()
tPaymentValue = modFormater.convertCurrencyToValue(Me.tPaymentValue)
If Me.tPaymentValue = "0" Then
    Me.tPaymentValue = ""
End If
tPaymentValue.SelStart = Len(tPaymentValue)
End Sub

Private Sub tPaymentValue_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
End Sub

Private Sub tPaymentValue_LostFocus()
tPaymentValue = modFormater.convertValueToCurrency(Me.tPaymentValue, 0)
End Sub

Private Sub tValue_GotFocus()
tValue = modFormater.convertCurrencyToValue(Me.tValue)
If Me.tValue = "0" Then
    Me.tValue = ""
End If
tValue.SelStart = Len(tValue)
End Sub

Private Sub tValue_KeyPress(KeyAscii As Integer)
KeyAscii = SoloNumeros(KeyAscii)
If KeyAscii = 13 Then
    Call cmdAdd_Click
End If
End Sub

Private Sub tValue_LostFocus()
tValue = modFormater.convertValueToCurrency(Me.tValue, 0)
End Sub

Private Function calculateTotal()
Dim item As Integer

netValue = 0
For item = 1 To Me.listData.ListItems.Count
    netValue = netValue + modFormater.convertCurrencyToValue(listData.ListItems(item).SubItems(1))
Next
Me.tTotalValue = modFormater.convertValueToCurrency(netValue, 0)

paymentValue = modFormater.convertCurrencyToValue(Me.tPaymentValue)
changeValue = paymentValue - netValue
changeValue = IIf(changeValue < 0, 0, changeValue)
residueValue = netValue - paymentValue
residueValue = IIf(residueValue < 0, 0, residueValue)

Me.tChangeValue = modFormater.convertValueToCurrency(changeValue, 0)
Me.tResidueValue = modFormater.convertValueToCurrency(residueValue, 0)

End Function

Private Sub imprimirFactura(id As Integer)
Dim oAcces As Access.APPLICATION
Set oAcces = New Access.APPLICATION

oAcces.OpenCurrentDatabase pathBD, False, keyBD
oAcces.Visible = False
oAcces.DoCmd.OpenReport "invoice", acViewPreview, , "id_invoice=" & id

oAcces.DoCmd.PrintOut acPrintAll
oAcces.CloseCurrentDatabase
oAcces.Quit
Set oAcces = Nothing
End Sub
