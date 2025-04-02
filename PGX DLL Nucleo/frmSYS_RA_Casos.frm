VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmSYS_RA_Casos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Expedientes RA: Control de Casos"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   18075
   LinkTopic       =   "Form1"
   ScaleHeight     =   10470
   ScaleWidth      =   18075
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lswAut 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   6840
      Width           =   9135
      _Version        =   1310723
      _ExtentX        =   16113
      _ExtentY        =   7223
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswCon 
      Height          =   4095
      Left            =   9240
      TabIndex        =   2
      Top             =   6840
      Width           =   8655
      _Version        =   1310723
      _ExtentX        =   15266
      _ExtentY        =   7223
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.CheckBox chkVencimiento 
      Height          =   375
      Left            =   12600
      TabIndex        =   19
      Top             =   120
      Width           =   2655
      _Version        =   1310723
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Vencimiento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   2
      Appearance      =   17
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtPersonaId 
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   17655
      _Version        =   524288
      _ExtentX        =   31141
      _ExtentY        =   9128
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
      MaxCols         =   11
      ScrollBars      =   2
      SpreadDesigner  =   "frmSYS_RA_Casos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   4695
      _Version        =   1310723
      _ExtentX        =   8281
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   495
      Left            =   16560
      TabIndex        =   11
      Top             =   360
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Casos.frx":1266
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   15480
      TabIndex        =   12
      Top             =   360
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_RA_Casos.frx":1B37
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   8760
      TabIndex        =   14
      Top             =   480
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   10680
      TabIndex        =   16
      Top             =   480
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   330
      Index           =   0
      Left            =   12600
      TabIndex        =   17
      Top             =   480
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   330
      Index           =   1
      Left            =   13920
      TabIndex        =   18
      Top             =   480
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   135
      Index           =   4
      Left            =   10800
      TabIndex        =   15
      Top             =   240
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Tipo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   135
      Index           =   3
      Left            =   8880
      TabIndex        =   13
      Top             =   240
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   135
      Index           =   2
      Left            =   4080
      TabIndex        =   10
      Top             =   240
      Width           =   4695
      _Version        =   1310723
      _ExtentX        =   8281
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Nombre"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   135
      Index           =   1
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Identificación"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   135
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   238
      _StockProps     =   79
      Caption         =   "Persona Id"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption scConsultas 
      Height          =   375
      Left            =   9240
      TabIndex        =   5
      Top             =   6480
      Width           =   8655
      _Version        =   1310723
      _ExtentX        =   15266
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Consulta de Expedientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scAutorizacion 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6480
      Width           =   9135
      _Version        =   1310723
      _ExtentX        =   16113
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Autorizaciones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Alignment       =   1
   End
End
Attribute VB_Name = "frmSYS_RA_Casos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vPaso As Boolean

Private Sub btnBuscar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtPersonaId.Text = fxSysCleanTxtInject(txtPersonaId.Text)
txtIdentificacion.Text = fxSysCleanTxtInject(txtIdentificacion.Text)
txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)

strSQL = "select 0, 0, 0, Persona_Id, Cedula, Nombre, EstadoDesc, TipoDesc, Fecha_Vence, Registro_Fecha, Registro_Usuario " _
       & " from vSYS_RA_Casos" _
       & " Where Cedula like '%" & txtIdentificacion.Text & "%'"

If Len(txtNombre.Text) > 0 Then
    strSQL = strSQL & " And Nombre like '%" & txtNombre.Text & "%'"
End If

If Len(txtPersonaId.Text) > 0 Then
    strSQL = strSQL & " And Persona_Id = " & txtPersonaId.Text
End If

If cboEstado.Text <> "TODOS" Then
    strSQL = strSQL & " And EstadoDesc = '" & cboEstado.Text & "'"
End If

If cboTipo.Text <> "TODOS" Then
    strSQL = strSQL & " And Tipo_Id = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
End If

If chkVencimiento.Value = xtpChecked Then
    strSQL = strSQL & " And Fecha_Vence between '" & Format(dtpVence(0).Value, "yyyy-mm-dd") _
            & " 00:00:00' and '" & Format(dtpVence(1).Value, "yyyy-mm-dd") & " 23:59:59'"
End If


strSQL = strSQL & " order by Persona_Id"
vPaso = True
    Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL, True)
vPaso = False


vGrid.MaxRows = vGrid.MaxRows - 1

scAutorizacion.Caption = "Autorizaciones"
scAutorizacion.Tag = "0"

scConsultas.Caption = "Seleccione una Autorización"
scConsultas.Tag = "0"

lswAut.ListItems.Clear
lswCon.ListItems.Clear



Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 11
    vHeaders.Headers(1) = ""
    vHeaders.Headers(2) = ""
    vHeaders.Headers(3) = ""
    vHeaders.Headers(4) = "Persona Id"
    vHeaders.Headers(5) = "Identificación"
    vHeaders.Headers(6) = "Nombre"
    vHeaders.Headers(7) = "Estado"
    vHeaders.Headers(8) = "Tipo"
    vHeaders.Headers(9) = "Vencimiento"
    vHeaders.Headers(10) = "Rec.Fecha"
    vHeaders.Headers(11) = "Rec. Usuario"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_RA_Consulta")

End Sub

Private Sub chkVencimiento_Click()

If chkVencimiento.Value = xtpChecked Then
    dtpVence(0).Enabled = True
Else
    dtpVence(0).Enabled = False
End If

dtpVence(1).Enabled = dtpVence(0).Enabled

End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 10


cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.AddItem "Inactiva"
cboEstado.AddItem "Vencida"
cboEstado.AddItem "TODOS"
cboEstado.Text = "TODOS"


strSQL = "SELECT TIPO_ID AS 'IDX', DESCRIPCION AS 'ItmX' FROM SYS_EXP_TIPOS WHERE ACTIVO = 1"
Call sbCbo_Llena_New(cboTipo, strSQL, True, True)

vGrid.MaxRows = 0

With lswAut.ColumnHeaders
    .Clear
    .Add , , "Id Autorización", 1800
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Horas", 1200, vbCenter
    .Add , , "Vence", 1800, vbCenter
    .Add , , "Autorizador", 3800, vbCenter
    .Add , , "Autorizado", 3800, vbCenter
End With


With lswCon.ColumnHeaders
    .Clear
    .Add , , "Id Autorización", 1800
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Usuario", 3800, vbCenter
    .Add , , "Autorizador", 3800, vbCenter
End With

chkVencimiento.Value = xtpUnchecked

dtpVence(0).Value = fxFechaServidor
dtpVence(1).Value = dtpVence(0).Value

Call chkVencimiento_Click

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Resize()
On Error Resume Next

scAutorizacion.Top = Me.Height - (lswAut.Height + scAutorizacion.Height + 600)
lswAut.Top = scAutorizacion.Top + scAutorizacion.Height + 100

scConsultas.Top = scAutorizacion.Top
lswCon.Top = lswAut.Top

scAutorizacion.Width = (Me.Width - 300) / 2
lswAut.Width = scAutorizacion.Width

scConsultas.Width = scAutorizacion.Width
lswCon.Width = lswAut.Width

scConsultas.Left = scAutorizacion.Width + scAutorizacion.Left + 50
lswCon.Left = scConsultas.Left

vGrid.Height = scAutorizacion.Top - (vGrid.Top + 150)
vGrid.Width = Me.Width - 250


End Sub


Private Sub sbAutorizaciones_Load(pPersonaId As Long)

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lswAut.ListItems.Clear

scConsultas.Tag = 0
scConsultas.Caption = "Selecciones una Autorización"

lswCon.ListItems.Clear

vPaso = True

strSQL = "select * from vSYS_RA_Autorizaciones where Persona_Id = " & pPersonaId _
       & " order by Autorizacion_Id desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAut.ListItems.Add(, , rs!Autorizacion_Id)
      itmX.SubItems(1) = rs!registro_Fecha
      itmX.SubItems(2) = rs!Horas
      itmX.SubItems(3) = rs!Fecha_Vence
      itmX.SubItems(4) = rs!Usuario_Autorizador
      itmX.SubItems(5) = rs!Usuario_Autorizado
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



Private Sub sbAccesos_Load(pAutorizacionId As Long)

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


scConsultas.Tag = pAutorizacionId
scConsultas.Caption = "Autorización Id: " & pAutorizacionId

lswCon.ListItems.Clear

vPaso = True

strSQL = "select * from vSYS_RA_Accesos where Autorizacion_Id = " & pAutorizacionId _
       & " order by registro_fecha desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCon.ListItems.Add(, , rs!Autorizacion_Id)
      itmX.SubItems(1) = rs!registro_Fecha
      itmX.SubItems(2) = rs!Usuario_Autorizado
      itmX.SubItems(3) = rs!Usuario_Autorizador
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


Private Sub lswAut_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Call sbAccesos_Load(Item.Text)

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

Dim pPersonaId As Long
Dim frm As Form


vGrid.Row = Row

'Expediente
If Col = 1 Then
  vGrid.Col = 4
  pPersonaId = vGrid.Text
   

 Call sbFormsCall("frmSYS_RA_Personas")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmSYS_RA_Personas") Then
     Call frm.sbConsulta_Externa(pPersonaId)
     Exit For
   End If
 Next frm
  
   
End If


'Registro Autorizacion
If Col = 2 Then
  vGrid.Col = 4
  pPersonaId = vGrid.Text
   
 Call sbFormsCall("frmSYS_RA_Autorizaciones")
 For Each frm In Forms
   If UCase(frm.Name) = UCase("frmSYS_RA_Autorizaciones") Then
     Call frm.sbConsulta_Externa(pPersonaId)
     Exit For
   End If
 Next frm
  
   
End If


'Consulta
If Col = 3 Then
   vGrid.Col = 4
   pPersonaId = vGrid.Text
    
   scAutorizacion.Tag = vGrid.Text
   
   vGrid.Col = 6
   scAutorizacion.Caption = "Autorizaciones para: " & vGrid.Text
    
   Call sbAutorizaciones_Load(pPersonaId)

End If


End Sub
