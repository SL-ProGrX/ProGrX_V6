VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCR_RetencionBitacora 
   Caption         =   "Bitácora Especial de Retenciones y Creditos"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   Icon            =   "frmCR_RetencionBitacora.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5220
   ScaleMode       =   0  'User
   ScaleWidth      =   13695
   WindowState     =   2  'Maximized
   Begin VB.Frame fraRevision 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   10080
      TabIndex        =   18
      Top             =   0
      Width           =   6375
      Begin VB.ComboBox cboRevision 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   20
         Top             =   120
         Width           =   1935
      End
      Begin VB.CheckBox chkRevision 
         Appearance      =   0  'Flat
         Caption         =   "Buscar Usuario/Fecha Revisión"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3240
         TabIndex        =   19
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox chkMovimientos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2640
      TabIndex        =   6
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox chkFechas 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2760
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chkUsuarios 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2760
      TabIndex        =   4
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.TextBox txtCedula 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox cboTipo 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4335
      Left            =   3960
      TabIndex        =   1
      Top             =   795
      Width           =   13455
      _Version        =   524288
      _ExtentX        =   23733
      _ExtentY        =   7646
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   496
      SpreadDesigner  =   "frmCR_RetencionBitacora.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionBitacora.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionBitacora.frx":6FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RetencionBitacora.frx":D842
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   360
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   635
      ButtonWidth     =   1852
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            Key             =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reporte"
            Key             =   "Reporte"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exportar"
            Key             =   "Exportar"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Excel"
                  Text            =   "Microsoft Excel"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "HTML"
                  Text            =   "HTML"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lswMovimientos 
      Height          =   4590
      Left            =   120
      TabIndex        =   8
      Top             =   2550
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   8096
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5539
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   157220865
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   1320
      TabIndex        =   10
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   157220865
      CurrentDate     =   37678
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   7800
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   10
      Left            =   480
      TabIndex        =   16
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   9
      Left            =   480
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   480
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas y Usuario del Movimiento ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   480
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
End
Attribute VB_Name = "frmCR_RetencionBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim vPaso As Boolean



Private Sub cboRevision_Click()
If cboRevision.ListCount = 0 Then Exit Sub
Call sbBuscar
End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If

End Sub

Private Sub chkMovimientos_Click()
Dim i As Integer

For i = 1 To lswMovimientos.ListItems.Count
  lswMovimientos.ListItems.Item(i).Checked = chkMovimientos.Value
Next i

End Sub


Private Sub chkRevision_Click()
If chkRevision.Value = vbChecked Then
   txtUsuario.BackColor = cboRevision.BackColor
Else
   txtUsuario.BackColor = vbWhite
End If
End Sub

Private Sub chkUsuarios_Click()
 If chkUsuarios.Value = vbChecked Then
   txtUsuario.Enabled = False
 Else
   txtUsuario.Enabled = True
   txtUsuario = "(Presione F4)"
 End If
  
End Sub

Private Sub sbBuscar()
Dim rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass
   
vGrid.MaxRows = 0
vGrid.MaxCols = 12

vPaso = True

rs.Open fxSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!id_Credito_SuBit
  
  vGrid.Col = 2
  vGrid.Text = rs!Usuario
  vGrid.Col = 3
  vGrid.Text = rs!Fecha 'Format(rs!Fecha, "dd/mm/yyyy")
  vGrid.Col = 4
  vGrid.Text = rs!MovimientoDesc
  vGrid.Col = 5
  vGrid.Text = CStr(rs!ID_SOLICITUD)
  vGrid.Col = 6
  vGrid.Text = rs!Codigo
  vGrid.Col = 7
  vGrid.Text = rs!Detalle
  vGrid.Col = 8
  vGrid.Text = Trim(rs!Cedula)
  vGrid.Col = 9
  vGrid.Text = rs!Nombre
  
  vGrid.Col = 10
  vGrid.Text = rs!notas & ""
  
  vGrid.Col = 11
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.Col = 12
  vGrid.Text = rs!Revisado_Fecha & ""
  
  
  vGrid.RowHeight(vGrid.Row) = vGrid.MaxTextRowHeight(vGrid.Row)
  rs.MoveNext
Loop
rs.Close

vPaso = False
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle

cboRevision.AddItem "TODOS"
cboRevision.AddItem "Pendientes"
cboRevision.AddItem "Revisados"
cboRevision.Text = "TODOS"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

lswMovimientos.ListItems.Clear
strSQL = "select MOVIMIENTO,DESCRIPCION from US_MOVIMIENTOS_BE WHERE MODULO = " & vModulo & " ORDER BY MOVIMIENTO"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswMovimientos.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!Movimiento
     itmX.Checked = chkMovimientos.Value
 rs.MoveNext
Loop
rs.Close


cboTipo.Clear
cboTipo.AddItem "Créditos"
cboTipo.AddItem "Retenciones"
cboTipo.AddItem "TODOS"
cboTipo.Text = "Créditos"


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - (520 + vGrid.Left)
vGrid.Height = Me.Height - (vGrid.Top + 1700)

lswMovimientos.Height = Me.Height - (lswMovimientos.Top + 1550)

End Sub


Private Function fxSQL() As String
Dim strSQL As String
Dim vCadena As String, i As Integer


strSQL = "select C.*,S.cedula,S.nombre,M.Descripcion as MovimientoDesc,case when C.revisado_fecha is null then 0 else 1 end as 'Revisado'" _
       & " from credito_subit C inner join reg_Creditos R on C.id_solicitud = R.id_solicitud" _
       & " inner join Socios S on S.cedula = R.cedula" _
       & " inner join US_MOVIMIENTOS_BE M on C.Movimiento = M.Movimiento" _
       & " Where M.Modulo = " & vModulo

If Len(Trim(txtCedula.Text)) > 0 Then
  strSQL = strSQL & " and S.cedula like '%" & txtCedula.Text & "%'"
End If
       
If chkFechas.Value = vbUnchecked Then
   If chkRevision.Value = vbChecked Then
        strSQL = strSQL & " and C.Revisado_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
   Else
        strSQL = strSQL & " and C.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
   End If
End If

'Lista de Tipos de Movimientos
vCadena = " and C.movimiento in('"
For i = 1 To lswMovimientos.ListItems.Count
  If lswMovimientos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswMovimientos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"

If chkUsuarios.Value = vbUnchecked Then
  If txtUsuario <> "" And txtUsuario <> "(Presione F4)" Then
     If chkRevision.Value = vbChecked Then
             strSQL = strSQL & " and C.Revisado_Usuario = '" & txtUsuario & "'"
     Else
             strSQL = strSQL & " and C.Usuario = '" & txtUsuario & "'"
     End If
  End If
End If

Select Case Mid(cboTipo.Text, 1, 1)
   Case "C" 'Creditos
        strSQL = strSQL & " and C.Tipo = 'C'"
   Case "R" 'Retenciones
        strSQL = strSQL & " and C.Tipo = 'R'"
   Case "T" 'Todos
End Select

Select Case Mid(cboRevision.Text, 1, 1)
   Case "P" 'Pendientes
        strSQL = strSQL & " and C.Revisado_Fecha is null"
   Case "R" 'Revisados
        strSQL = strSQL & " and C.Revisado_Fecha is not null"
   Case "T" 'Todos
End Select

If chkRevision.Value = vbChecked Then
    strSQL = strSQL & " order by C.Revisado_fecha"
Else
    strSQL = strSQL & " order by C.fecha"
End If

fxSQL = strSQL

 
End Function



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  
  Case "Reporte"
        vGrid.PrintHeader = "Créditos: Bitácora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Select
End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "Revisado?"
    vHeaders.Headers(2) = "Usuario"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Operación"
    vHeaders.Headers(6) = "Línea"
    vHeaders.Headers(7) = "Detalle"
    vHeaders.Headers(8) = "Cédula"
    vHeaders.Headers(9) = "Nombre"
    vHeaders.Headers(10) = "Notas"
    vHeaders.Headers(11) = "Revisado por"
    vHeaders.Headers(12) = "Revisado Fecha"
    
Select Case ButtonMenu.Key
  Case "Excel"
      Call sbSIFGridExportar(vGrid, vHeaders, "Creditos_BitacoraEspecial")
  Case "HTML"
      Call sbSIFGridExportar(vGrid, vHeaders, "Creditos_BitacoraEspecial", "HTML")
End Select
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select Cedula,Nombre From Socios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtCedula.Text = Trim(gBusquedas.Resultado)
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Or Col > 1 Or Not fraRevision.Enabled Then Exit Sub
 
vGrid.Row = Row
vGrid.Col = 1
If vGrid.Value = vbChecked Then
   strSQL = "update CREDITO_SUBIT set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = getdate()" _
          & " where id_Credito_SuBit = " & vGrid.CellTag
   glogon.Conection.Execute strSQL

   vGrid.Col = 11
   vGrid.Text = glogon.Usuario
   vGrid.Col = 12
   vGrid.Text = Date
   
End If
End Sub
