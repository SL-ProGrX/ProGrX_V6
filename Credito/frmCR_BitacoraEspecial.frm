VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_BitacoraEspecial 
   Caption         =   "Bit�cora Especial de Cr�ditos y Retenciones"
   ClientHeight    =   7296
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   16944
   Icon            =   "frmCR_BitacoraEspecial.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6602.961
   ScaleMode       =   0  'User
   ScaleWidth      =   16947.41
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   228
      Left            =   3000
      TabIndex        =   18
      Top             =   480
      Width           =   216
      _Version        =   1245185
      _ExtentX        =   370
      _ExtentY        =   409
      _StockProps     =   79
      Caption         =   "Todas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ListView lswMovimientos 
      Height          =   4692
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   3732
      _Version        =   1245185
      _ExtentX        =   6583
      _ExtentY        =   8276
      _StockProps     =   77
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
   End
   Begin VB.Frame fraRevision 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   10080
      TabIndex        =   9
      Top             =   0
      Width           =   6375
      Begin XtremeSuiteControls.ComboBox cboRevision 
         Height          =   312
         Left            =   960
         TabIndex        =   17
         Top             =   120
         Width           =   2172
         _Version        =   1245185
         _ExtentX        =   3831
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkRevision 
         Height          =   252
         Left            =   3240
         TabIndex        =   21
         Top             =   120
         Width           =   2892
         _Version        =   1245185
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Buscar Usuario/Fecha Revisi�n"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Revisi�n ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
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
         TabIndex        =   10
         Top             =   120
         Width           =   5535
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4332
      Left            =   4080
      TabIndex        =   0
      Top             =   792
      Width           =   13452
      _Version        =   524288
      _ExtentX        =   23728
      _ExtentY        =   7641
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   496
      SpreadDesigner  =   "frmCR_BitacoraEspecial.frx":000C
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_BitacoraEspecial.frx":07DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_BitacoraEspecial.frx":7041
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_BitacoraEspecial.frx":D8A3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   312
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   3612
      _ExtentX        =   6371
      _ExtentY        =   550
      ButtonWidth     =   1757
      ButtonHeight    =   550
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
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   348
      Left            =   1440
      TabIndex        =   11
      Top             =   480
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   614
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   348
      Left            =   1440
      TabIndex        =   12
      Top             =   840
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   614
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1440
      TabIndex        =   13
      Top             =   1320
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1440
      TabIndex        =   14
      Top             =   1680
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   8280
      TabIndex        =   16
      Top             =   120
      Width           =   1692
      _Version        =   1245185
      _ExtentX        =   2985
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   228
      Left            =   3000
      TabIndex        =   19
      Top             =   1320
      Width           =   216
      _Version        =   1245185
      _ExtentX        =   370
      _ExtentY        =   409
      _StockProps     =   79
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkMovimientos 
      Height          =   228
      Left            =   3600
      TabIndex        =   20
      Top             =   2160
      Width           =   216
      _Version        =   1245185
      _ExtentX        =   370
      _ExtentY        =   409
      _StockProps     =   79
      Caption         =   "Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
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
      TabIndex        =   8
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   10
      Left            =   360
      TabIndex        =   7
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   9
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   8
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas y Usuario del Movimiento ...:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos ...:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificaci�n"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmCR_BitacoraEspecial.frx":14105
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4008
   End
End
Attribute VB_Name = "frmCR_BitacoraEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
  
  vGrid.col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!id_credito_suBit
  
  vGrid.col = 2
  vGrid.Text = rs!Usuario
  vGrid.col = 3
  vGrid.Text = rs!fecha 'Format(rs!Fecha, "dd/mm/yyyy")
  vGrid.col = 4
  vGrid.Text = rs!MovimientoDesc
  vGrid.col = 5
  vGrid.Text = CStr(rs!Id_Solicitud)
  vGrid.col = 6
  vGrid.Text = rs!Codigo
  vGrid.col = 7
  vGrid.Text = rs!Detalle
  vGrid.col = 8
  vGrid.Text = Trim(rs!Cedula)
  vGrid.col = 9
  vGrid.Text = rs!Nombre
  
  vGrid.col = 10
  vGrid.Text = rs!notas & ""
  
  vGrid.col = 11
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.col = 12
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
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle

With lswMovimientos.ColumnHeaders
    .Clear
    .Add , , "", 3150
End With


cboRevision.AddItem "TODOS"
cboRevision.AddItem "Pendientes"
cboRevision.AddItem "Revisados"
cboRevision.Text = "TODOS"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

lswMovimientos.ListItems.Clear
strSQL = "select MOVIMIENTO,DESCRIPCION from US_MOVIMIENTOS_BE WHERE MODULO = " & vModulo & " ORDER BY MOVIMIENTO"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswMovimientos.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!Movimiento
     itmX.Checked = chkMovimientos.Value
 rs.MoveNext
Loop
rs.Close


cboTipo.Clear
cboTipo.AddItem "Cr�ditos"
cboTipo.AddItem "Retenciones"
cboTipo.AddItem "TODOS"
cboTipo.Text = "Cr�ditos"


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - (520 + vGrid.Left)
vGrid.Height = Me.Height - (vGrid.top + 1700)

lswMovimientos.Height = Me.Height - (lswMovimientos.top + 1700)

imgBanner.Height = Me.Height

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
        vGrid.PrintHeader = "Cr�ditos: Bit�cora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
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
    vHeaders.Headers(5) = "Operaci�n"
    vHeaders.Headers(6) = "L�nea"
    vHeaders.Headers(7) = "Detalle"
    vHeaders.Headers(8) = "C�dula"
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

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Or col > 1 Or Not fraRevision.Enabled Then Exit Sub
 
vGrid.Row = Row
vGrid.col = 1
If vGrid.Value = vbChecked Then
   strSQL = "update CREDITO_SUBIT set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = dbo.MyGetdate()" _
          & " where id_Credito_SuBit = " & vGrid.CellTag
   Call ConectionExecute(strSQL)

   vGrid.col = 11
   vGrid.Text = glogon.Usuario
   vGrid.col = 12
   vGrid.Text = Date
   
End If
End Sub
