VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmInvTransacQry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Transacciones de inventarios"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lsw 
      Height          =   4695
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#Boleta"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Estado"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "#Documento"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Detalle"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Usr.Solicita"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Fec.Solicita"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Usr.Autoriza"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fec.Autoriza"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Usr.Procesa"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Fec.Procesa"
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.TextBox txtLineas 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5280
      TabIndex        =   12
      Text            =   "500"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      TabIndex        =   10
      Top             =   840
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   38201
   End
   Begin VB.ComboBox cboUsrBase 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.ComboBox cboFecBase 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox cboEstado 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   5280
      TabIndex        =   9
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   111345667
      CurrentDate     =   38201
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   612
      Left            =   6840
      TabIndex        =   14
      Top             =   240
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   2
      Picture         =   "frmInvTransacQry.frx":0000
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Líneas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   7
      Left            =   4560
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   8400
      X2              =   0
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   6
      Left            =   4560
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   2640
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   4
      Left            =   480
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   3
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmInvTransacQry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cboFecBase_Click()
If Mid(cboFecBase.Text, 1, 1) = "T" Then
 dtpInicio.Enabled = False
Else
 dtpInicio.Enabled = True
End If
dtpCorte.Enabled = dtpInicio.Enabled
End Sub

Private Sub cboUsrBase_Click()

If Mid(cboUsrBase.Text, 1, 1) = "T" Then
 txtUser.Enabled = False
Else
 txtUser.Enabled = True
End If

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select TOP " & txtLineas & " * from pv_invTransac where Tipo = '" & gInvTran.Tipo & "'"

If Mid(cboEstado.Text, 1, 1) <> "T" Then
  strSQL = strSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
End If

Select Case Mid(cboFecBase.Text, 1, 1)
  Case "S"
    strSQL = strSQL & " and genera_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "A"
    strSQL = strSQL & " and autoriza_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "P"
    strSQL = strSQL & " and procesa_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "I"
    strSQL = strSQL & " and fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select

Select Case Mid(cboUsrBase.Text, 1, 1)
  Case "S"
    strSQL = strSQL & " and genera_user like '" & txtUser & "%'"
  Case "A"
    strSQL = strSQL & " and autoriza_user like '" & txtUser & "%'"
  Case "P"
    strSQL = strSQL & " and procesa_user like '" & txtUser & "%'"
End Select

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Boleta)
     itmX.SubItems(1) = Format(rs!fecha, "dd/mm/yyyy")
     Select Case rs!Estado
       Case "S"
         itmX.SubItems(2) = "Solicitada"
       Case "A"
         itmX.SubItems(2) = "Autorizada"
       Case "P"
         itmX.SubItems(2) = "Procesada"
       Case "R"
         itmX.SubItems(2) = "Rechazada"
     End Select
     itmX.SubItems(3) = rs!Documento & ""
     itmX.SubItems(4) = rs!notas & ""
     itmX.SubItems(5) = rs!genera_user & ""
     itmX.SubItems(6) = Format(rs!genera_fecha, "dd/mm/yyyy")
     itmX.SubItems(7) = rs!Autoriza_user & ""
     itmX.SubItems(8) = Format(rs!Autoriza_Fecha, "dd/mm/yyyy")
     itmX.SubItems(9) = rs!Procesa_user & ""
     itmX.SubItems(10) = Format(rs!Procesa_Fecha, "dd/mm/yyyy")

 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()

Select Case gInvTran.Tipo
  Case "E"
    Me.Caption = Me.Caption & " / ENTRADAS"
  Case "S"
    Me.Caption = Me.Caption & " / SALIDAS"
  Case "T"
    Me.Caption = Me.Caption & " / TRASLADOS ENTRE BODEGAS"
  Case "R"
    Me.Caption = Me.Caption & " / REQUISICIONES"
End Select

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

With cboEstado
 .Clear
 .AddItem "Solicitada"
 .AddItem "Autorizada"
 .AddItem "Procesada"
 .AddItem "Rechazada"
 .AddItem "TODOS"
 .Text = "TODOS"
End With

With cboFecBase
 .Clear
 .AddItem "Inventario"
 .AddItem "Solicita"
 .AddItem "Autoriza"
 .AddItem "Procesa"
 .AddItem "TODOS"
 .Text = "TODOS"
 cboFecBase_Click
End With

With cboUsrBase
 .Clear
 .AddItem "Solicita"
 .AddItem "Autoriza"
 .AddItem "Procesa"
 .AddItem "TODOS"
 .Text = "TODOS"
 cboUsrBase_Click
End With



End Sub


Private Sub lsw_Click()

If lsw.ListItems.Count > 0 Then
 gInvTran.Boleta = lsw.SelectedItem
 Unload Me
End If

End Sub
