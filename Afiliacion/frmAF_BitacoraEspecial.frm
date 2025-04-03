VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_BitacoraEspecial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácora Especial de Afiliación"
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15165
   Icon            =   "frmAF_BitacoraEspecial.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   15165
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   204
      Left            =   3000
      TabIndex        =   21
      Top             =   600
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.GroupBox fraRevision 
      Height          =   492
      Left            =   6000
      TabIndex        =   12
      Top             =   470
      Width           =   5652
      _Version        =   1441793
      _ExtentX        =   9970
      _ExtentY        =   868
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkRevision 
         Height          =   252
         Left            =   3120
         TabIndex        =   15
         Top             =   120
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Buscar Usuario/Fecha Revisión"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
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
      Begin XtremeSuiteControls.ComboBox cboRevision 
         Height          =   312
         Left            =   1440
         TabIndex        =   14
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Revisión ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   120
         Width           =   7692
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5895
      Left            =   4080
      TabIndex        =   0
      Top             =   960
      Width           =   11055
      _Version        =   524288
      _ExtentX        =   19500
      _ExtentY        =   10398
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
      MaxCols         =   9
      SpreadDesigner  =   "frmAF_BitacoraEspecial.frx":6852
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   556
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
      Height          =   312
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   556
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
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      Top             =   1320
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   1320
      TabIndex        =   10
      Top             =   1680
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswMovimientos 
      Height          =   4692
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
      _ExtentY        =   8276
      _StockProps     =   77
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
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkRevTodos 
      Height          =   252
      Left            =   4920
      TabIndex        =   16
      Top             =   550
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Revisados?"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   4920
      TabIndex        =   17
      Top             =   24
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_BitacoraEspecial.frx":6F97
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   372
      Left            =   6120
      TabIndex        =   18
      Top             =   24
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_BitacoraEspecial.frx":7697
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   372
      Left            =   7680
      TabIndex        =   19
      Top             =   24
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_BitacoraEspecial.frx":7D9E
   End
   Begin XtremeSuiteControls.PushButton btnRevisar 
      Height          =   372
      Left            =   9360
      TabIndex        =   20
      Top             =   24
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Revisar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_BitacoraEspecial.frx":7F08
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   204
      Left            =   3000
      TabIndex        =   22
      Top             =   1320
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkMovimientos 
      Height          =   204
      Left            =   3600
      TabIndex        =   23
      Top             =   2160
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1092
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
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
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   8
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1092
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
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   9
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1092
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
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   10
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmAF_BitacoraEspecial.frx":862F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3972
   End
End
Attribute VB_Name = "frmAF_BitacoraEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 9
    vHeaders.Headers(1) = "Revisado?"
    vHeaders.Headers(2) = "Usuario"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Cédula"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Detalle"
    vHeaders.Headers(8) = "Revisado por"
    vHeaders.Headers(9) = "Revisado Fecha"
      
    Call sbSIFGridExportar(vGrid, vHeaders, "Afiliación_BitacoraEspecial")

End Sub

Private Sub btnInforme_Click()
        vGrid.PrintHeader = "Clientes: Bitácora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Sub

Private Sub btnRevisar_Click()
 Call sbRevisar
End Sub

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

Private Sub chkRevTodos_Click()
Dim i As Long

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  vGrid.Value = chkRevTodos.Value
Next i

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
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select C.*,S.cedula,S.nombre,M.Descripcion as MovimientoDesc, case when C.revisado_fecha is null then 0 else 1 end as 'Revisado'" _
       & " from Afi_Bitacora_especial C inner join  Socios S on S.cedula = C.cedula" _
       & " inner join US_MOVIMIENTOS_BE M on C.Movimiento = M.Movimiento" _
       & " Where M.Modulo = " & vModulo
       
 
 
If Len(Trim(txtCedula.Text)) > 0 Then
  strSQL = strSQL & " and C.cedula like '%" & txtCedula.Text & "%'"
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

'
If chkUsuarios.Value = vbUnchecked Then
  If txtUsuario <> "" And txtUsuario <> "(Presione F4)" Then
     If chkRevision.Value = vbChecked Then
             strSQL = strSQL & " and C.Revisado_Usuario = '" & txtUsuario & "'"
     Else
             strSQL = strSQL & " and C.Usuario = '" & txtUsuario & "'"
     End If
  End If
End If

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

vPaso = True
vGrid.MaxRows = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!id_Bitacora
  
  vGrid.col = 2
  vGrid.Text = rs!Usuario
  vGrid.col = 3
  vGrid.Text = rs!fecha
  vGrid.col = 4
  vGrid.Text = rs!MovimientoDesc
  vGrid.col = 5
  vGrid.Text = rs!Cedula
  vGrid.col = 6
  vGrid.Text = rs!Nombre
  vGrid.col = 7
  vGrid.Text = rs!Detalle & ""
  vGrid.col = 8
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.col = 9
  vGrid.Text = rs!Revisado_Fecha & ""
  
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



Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

lswMovimientos.ColumnHeaders.Add , , "", 3150

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

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - (500 + vGrid.Left)
vGrid.Height = Me.Height - (vGrid.top + 750)

lswMovimientos.Height = Me.Height - (lswMovimientos.top + 750)

imgBanner.Height = Me.Height

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



Private Sub sbRevisar()
Dim strSQL As String, i As Long, IdBitacora As Long

If vPaso Or Not fraRevision.Enabled Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


With vGrid

  For i = 1 To .MaxRows
     .Row = i
     .col = 1
     If .Value = vbChecked Then
        IdBitacora = .CellTag
        .col = 8
        If Trim(.Text) = "" Then
            strSQL = "update AFI_BITACORA_ESPECIAL set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = dbo.MyGetdate()" _
                   & " where id_Bitacora = " & IdBitacora
            Call ConectionExecute(strSQL)
    
            vGrid.col = 8
            vGrid.Text = glogon.Usuario
            vGrid.col = 9
            vGrid.Text = Date
        End If
      End If
   
   Next i
End With
 
Me.MousePointer = vbDefault
MsgBox "Revisión aplicada satisfactoriamente!", vbInformation
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
 
End Sub
