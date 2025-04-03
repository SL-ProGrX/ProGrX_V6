VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSys_Gestiones_Bitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitácora de Gestiones de Clientes"
   ClientHeight    =   10590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16425
   LinkTopic       =   "Form1"
   ScaleHeight     =   10590
   ScaleWidth      =   16425
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   13920
      Top             =   120
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Todas las Fecha"
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   8520
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Appearance      =   17
      Picture         =   "frmSys_Gestiones_Bitacora.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7335
      Left            =   0
      TabIndex        =   2
      Top             =   1800
      Width           =   16815
      _Version        =   524288
      _ExtentX        =   29660
      _ExtentY        =   12938
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
      MaxCols         =   6
      ScrollBars      =   2
      SpreadDesigner  =   "frmSys_Gestiones_Bitacora.frx":0700
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   120
      Width           =   1935
      _Version        =   1441793
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
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1680
      TabIndex        =   5
      Top             =   840
      Width           =   1935
      _Version        =   1441793
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
   Begin XtremeSuiteControls.ComboBox cboGestion 
      Height          =   330
      Left            =   1680
      TabIndex        =   6
      Top             =   480
      Width           =   6255
      _Version        =   1441793
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   5280
      TabIndex        =   7
      Top             =   840
      Width           =   1335
      _Version        =   1441793
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   6600
      TabIndex        =   8
      Top             =   840
      Width           =   1335
      _Version        =   1441793
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   495
      Left            =   9840
      TabIndex        =   9
      Top             =   960
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Appearance      =   17
      Picture         =   "frmSys_Gestiones_Bitacora.frx":0D0E
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Gestión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   10
      Top             =   840
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Fecha Registro"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSys_Gestiones_Bitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean


Private Sub sbInicializa()

On Error GoTo vError

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -3, dtpCorte.Value)


vPaso = True
    
strSQL = "select cod_gestion as 'IdX', rtrim(descripcion) as 'ItmX' " _
       & " from SYS_GESTIONES_TIPOS Where Activa = 1" _
       & " order by descripcion"

Call sbCbo_Llena_New(cboGestion, strSQL, True, True)

vPaso = False

Exit Sub

vError:



End Sub

Private Sub btnBuscar_Click()
Dim pWhere As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)
txtCedula.Text = fxSysCleanTxtInject(txtCedula.Text)
txtNombre.Text = fxSysCleanTxtInject(txtNombre.Text)


pWhere = False

strSQL = "select CEDULA, NOMBRE, REGISTRO_FECHA, REGISTRO_USUARIO, DESCRIPCION, NOTAS" _
       & " from  vSys_Bitacora_Operaciones" _


If chkFechas.Value = xtpUnchecked Then
  If Not pWhere Then
        pWhere = True
        strSQL = strSQL & " WHERE "
  Else
        strSQL = strSQL & " AND "
  End If
  
    strSQL = strSQL & "REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
    
End If


If Len(txtUsuario.Text) > 0 Then
  If Not pWhere Then
        pWhere = True
        strSQL = strSQL & " WHERE "
  Else
        strSQL = strSQL & " AND "
  End If
    strSQL = strSQL & "REGISTRO_USUARIO like '%" & txtUsuario.Text & "%'"
End If


If Len(txtCedula.Text) > 0 Then
  If Not pWhere Then
        pWhere = True
        strSQL = strSQL & " WHERE "
  Else
        strSQL = strSQL & " AND "
  End If
    strSQL = strSQL & "Cedula like '%" & txtCedula.Text & "%'"
End If

If Len(txtNombre.Text) > 0 Then
  If Not pWhere Then
        pWhere = True
        strSQL = strSQL & " WHERE "
  Else
        strSQL = strSQL & " AND "
  End If
    strSQL = strSQL & "Nombre like '%" & txtNombre.Text & "%'"
End If


If cboGestion.Text <> "TODOS" Then
  If Not pWhere Then
        pWhere = True
        strSQL = strSQL & " WHERE "
  Else
        strSQL = strSQL & " AND "
  End If
    strSQL = strSQL & "COD_GESTION= '" & cboGestion.ItemData(cboGestion.ListIndex) & "'"
End If

Call sbCargaGrid(vGrid, 6, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders

    vHeaders.Columnas = 6
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Usuario"
    vHeaders.Headers(5) = "Gestión"
    vHeaders.Headers(6) = "Notas"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Gestiones_Clientes_Log")

End Sub



Private Sub chkFechas_Click()
If chkFechas.Value = xtpChecked Then
    dtpInicio.Enabled = False
    dtpCorte.Enabled = False
Else
    dtpInicio.Enabled = True
    dtpCorte.Enabled = True
End If

End Sub

Private Sub Form_Load()
vModulo = 10


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 230
vGrid.Height = Me.Height - (vGrid.Top + 450)

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado <> "" Then
        txtCedula.Text = Trim(gBusquedas.Resultado)
        txtNombre.Text = gBusquedas.Resultado2
    End If
End If

End Sub


