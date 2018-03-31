VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientDetail 
   BackColor       =   &H00373436&
   BorderStyle     =   0  'None
   ClientHeight    =   8295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   3720
      TabIndex        =   25
      Top             =   2595
      Width           =   495
   End
   Begin VB.CheckBox optOnlyPendinfInvoices 
      BackColor       =   &H00373436&
      Caption         =   "Ver sólo facturas pendientes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0004D3FF&
      Height          =   285
      Left            =   7320
      TabIndex        =   24
      Top             =   2160
      Width           =   3135
   End
   Begin VB.CommandButton cmdAddFocus 
      Height          =   195
      Left            =   7920
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
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
      Left            =   6360
      TabIndex        =   7
      Top             =   1080
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
      Left            =   360
      TabIndex        =   6
      Top             =   1080
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
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox tFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   2595
      Width           =   495
   End
   Begin VB.TextBox tFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   900
      TabIndex        =   3
      Top             =   2595
      Width           =   495
   End
   Begin VB.TextBox tFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1380
      TabIndex        =   2
      Top             =   2595
      Width           =   495
   End
   Begin VB.TextBox tFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   1920
      TabIndex        =   1
      Top             =   2595
      Width           =   495
   End
   Begin VB.TextBox tFiltro 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   3120
      TabIndex        =   0
      Top             =   2595
      Width           =   495
   End
   Begin MSComctlLib.ListView listData 
      Height          =   4215
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7435
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "id"
         Object.Width           =   5644
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saldo restante"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "id_client"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label tSumInvoiceResidue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00373436&
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CD7C10&
      Height          =   375
      Left            =   7200
      TabIndex        =   23
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label tSumInvoiceValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00373436&
      Caption         =   "$0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00CD7C10&
      Height          =   375
      Left            =   4080
      TabIndex        =   22
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label tTotalInvoice 
      BackColor       =   &H00373436&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0004D3FF&
      Height          =   375
      Left            =   7560
      TabIndex        =   21
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label label 
      BackColor       =   &H00373436&
      Caption         =   "Total de facturas"
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
      Left            =   5040
      TabIndex        =   20
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label tTotalInvoicePending 
      BackColor       =   &H00373436&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0004D3FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   19
      Top             =   1620
      Width           =   615
   End
   Begin VB.Label label 
      BackColor       =   &H00373436&
      Caption         =   "Facturas con saldo pendiente"
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
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   1680
      Width           =   3375
   End
   Begin VB.Image cmdExit 
      Height          =   405
      Left            =   10080
      Picture         =   "frmClientDetail.frx":0000
      Top             =   7680
      Width           =   2025
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
      Left            =   6360
      TabIndex        =   17
      Top             =   720
      Width           =   1455
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
      Left            =   2400
      TabIndex        =   16
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label label 
      BackColor       =   &H00373436&
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
      Left            =   360
      TabIndex        =   15
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label label 
      BackColor       =   &H00373436&
      Caption         =   "Lista de facturas"
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
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H003933ED&
      X1              =   240
      X2              =   2205
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "frmClientDetail.frx":2B4A
      Top             =   120
      Width           =   405
   End
   Begin VB.Line Line2 
      BorderColor     =   &H001A1A1A&
      BorderWidth     =   5
      X1              =   0
      X2              =   0
      Y1              =   120
      Y2              =   8040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   840
      TabIndex        =   13
      Top             =   120
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H003933ED&
      BorderWidth     =   2
      X1              =   240
      X2              =   2400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image cmdAdd 
      Height          =   405
      Left            =   8520
      Picture         =   "frmClientDetail.frx":33C0
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Label tNoInvoices 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "X clientes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label cmdCleanFilters 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00373436&
      Caption         =   "Limpiar filtros"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0004D3FF&
      Height          =   285
      Index           =   1
      Left            =   10200
      TabIndex        =   11
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label tIdClient 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmClientDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Conexión activa de BD  para asignar un servicio
Dim conBd As ADODB.Connection
Dim rec As New ADODB.Recordset

'SQL actual para los reportes
Dim baseSQL As String
Dim filtersApplied As Integer

Public action As Integer
Public parent As Form

Private Sub cmdAdd_Click()
'Valida la información del cliente
If (Me.tDocument = "") Then
    MsgBox "Debe ingresar el documento del cliente", vbCritical
    Me.tDocument.SetFocus
    Exit Sub
End If

If (Me.tName = "") Then
    MsgBox "Debe ingresar el nombre del cliente", vbCritical
    Me.tName.SetFocus
    Exit Sub
End If

If (Me.tPhone = "") Then
    MsgBox "Debe ingresar el teléfono del cliente", vbCritical
    Me.tPhone.SetFocus
    Exit Sub
End If

'Si existe se actualizan los datos
SQL = "UPDATE client SET document='" & Me.tDocument & "'," & _
"name='" & Me.tName & "',phone='" & Me.tPhone & "' WHERE id=" & Me.tIdClient & ""
conBd.Execute (SQL)
Sleep 800

MsgBox "El cliente fue actualizado con éxito", vbInformation
End Sub

Private Sub cmdAddFocus_Click()
Call cmdAdd_Click
End Sub

Private Sub cmdCleanFilters_Click(Index As Integer)
modComponents.cleanFilters tFiltro, -1
filtersApplied = 0
Me.listData.Sorted = False
Me.loadInvoices
loadTotalsInvoice
loadSumInvoice
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call createConexion

'width for the columns
Dim widthTotal As Double
Dim widthCols(5) As Double

widthTotal = Me.listData.Width
widthCols(1) = widthTotal * 0 'id
widthCols(2) = widthTotal * 0.33 'fecha
widthCols(3) = widthTotal * 0.33 'valor total
widthCols(4) = widthTotal * 0.34 'saldo restante
widthCols(5) = widthTotal * 0 'cliente

modComponents.setWidthForColumnsAndFilters tFiltro, listData, widthCols

Me.tFiltro(0).Visible = False
Me.tFiltro(1).Visible = False
Me.tFiltro(2).Visible = False

Me.tSumInvoiceValue.Width = Me.tFiltro(3).Width
Me.tSumInvoiceValue.left = Me.tFiltro(3).left - 100

Me.tSumInvoiceResidue.Width = Me.tFiltro(4).Width
Me.tSumInvoiceResidue.left = Me.tFiltro(4).left - 100

filtersApplied = 0

Me.Top = frmMenu.source.Top
Me.left = frmMenu.source.left
End Sub

'Se solicita una conexion a la bd
Private Function createConexion()
Set conBd = modConexion.getNewConection
rec.CursorLocation = adUseClient
End Function

Private Sub Form_Unload(Cancel As Integer)
Me.parent.refreshData
End Sub

Private Sub listData_DblClick()
If Me.listData.SelectedItem.Index > 0 Then
    frmInvoice.Show , Ap.frmMenu
    frmInvoice.loadInvoice getInvoiceSelected()
End If
End Sub

Private Function getInvoiceSelected() As cInvoice
Dim invoice As New cInvoice
invoice.loadInvoice Me.listData.SelectedItem, Me.listData.SelectedItem.SubItems(4), Me.listData.SelectedItem.SubItems(1), Me.listData.SelectedItem.SubItems(2), Me.listData.SelectedItem.SubItems(3)
Set getInvoiceSelected = invoice
End Function

Private Sub optOnlyPendinfInvoices_Click()
Call queryWithParameters
End Sub

Private Sub tDocument_KeyPress(KeyAscii As Integer)
Call executeRegister(KeyAscii)
End Sub

Private Sub executeRegister(KeyAscii)
If (13 = KeyAscii) Then
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

Private Sub loadList(SQL As String)

SQL = SQL & " order by date_invoice DESC"
rec.Open SQL, conBd, adOpenStatic, adLockOptimistic
Me.listData.ListItems.Clear
Do Until rec.EOF
    Set li = Me.listData.ListItems.Add(, , rec("id"))
        li.SubItems(1) = rec("date_invoice")
        li.SubItems(2) = modFormater.convertValueToCurrency(rec("total_value"), 0)
        li.SubItems(3) = modFormater.convertValueToCurrency(rec("residue_value"), 0)
        li.SubItems(4) = rec("id_client")
    rec.MoveNext
Loop
Me.tNoInvoices = rec.RecordCount & " facturas registradas."
rec.Close
End Sub

Public Sub loadInvoices()
Call loadList("Select * from invoice where id_client=" & Me.tIdClient & "")
End Sub

Private Sub tFiltro_Change(Index As Integer)
If modComponents.cleaningFilters Then Exit Sub
Call queryWithParameters
End Sub

'Agrega los parametros al SQL para su cosulta según los criterios de filtro
Private Function queryWithParameters()

SQL = "Select * from invoice where id_client=" & Me.tIdClient & ""
filtersApplied = 1

'Verifica y agrega los criterios de los filtros
On Error GoTo control
Dim countFilters As Integer
For countFilters = 2 To Me.tFiltro.count - 1
    If Me.tFiltro(countFilters).Text <> "" Then
        Select Case countFilters
            Case 3
                addParameter "total_value like '%" & tFiltro(countFilters) & "%'"
            Case 4
                addParameter "residue_value like '%" & tFiltro(countFilters) & "%'"
        End Select
    End If
Next

If (Me.optOnlyPendinfInvoices.value = 1) Then
    addParameter "residue_value >0"
End If

Call loadList(SQL)
loadSumInvoice
filtersApplied = 0
Exit Function
control:
If Err.Number = 503 Then
    Resume Next
End If
End Function

Private Function addParameter(parameter As String)
If filtersApplied = 0 Then
    SQL = SQL & " WHERE "
Else
    SQL = SQL & " AND "
End If
SQL = SQL & parameter
filtersApplied = filtersApplied + 1
End Function

Private Sub tName_KeyPress(KeyAscii As Integer)
Call executeRegister(KeyAscii)
End Sub

Private Sub tPhone_KeyPress(KeyAscii As Integer)
Call executeRegister(KeyAscii)
End Sub

Public Sub loadClient(client As cClient)
Me.tIdClient = client.id
Me.tName = client.name
Me.tDocument = client.document
Me.tPhone = client.phone
Me.tDocument.SetFocus
Call loadInvoices
loadTotalsInvoice
loadSumInvoice
End Sub

Public Sub loadTotalsInvoice()
Dim item As Integer
Dim invoicesPending As Integer
invoicesPending = 0
For item = 1 To Me.listData.ListItems.count
    If modFormater.convertCurrencyToValue(listData.ListItems(item).SubItems(3)) > 0 Then
        invoicesPending = invoicesPending + 1
    End If
Next
Me.tTotalInvoice = Me.listData.ListItems.count
Me.tTotalInvoicePending = invoicesPending
End Sub

Public Sub loadSumInvoice()
Dim item As Integer
Dim sumInvoiceTotal As Double
Dim sumInvoiceResidue As Double
sumInvoiceTotal = 0
sumInvoiceResidue = 0
For item = 1 To Me.listData.ListItems.count
    sumInvoiceTotal = sumInvoiceTotal + modFormater.convertCurrencyToValue(listData.ListItems(item).SubItems(2))
    sumInvoiceResidue = sumInvoiceResidue + modFormater.convertCurrencyToValue(listData.ListItems(item).SubItems(3))
Next
Me.tSumInvoiceResidue = modFormater.convertValueToCurrency(sumInvoiceResidue, 0)
Me.tSumInvoiceValue = modFormater.convertValueToCurrency(sumInvoiceTotal, 0)
End Sub

