VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCxC_ContratosPagadores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pagadores"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   615
      Left            =   7680
      TabIndex        =   7
      Top             =   720
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4471
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Mostratr solo los asociados a este contrato ?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6012
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   10932
      _Version        =   524288
      _ExtentX        =   19283
      _ExtentY        =   10605
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCxC_ContratosPagadores.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   6615
      _Version        =   1310723
      _ExtentX        =   11668
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   315
      Left            =   9600
      TabIndex        =   6
      Top             =   240
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   556
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
      Text            =   "1000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   556
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
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   960
      Width           =   5295
      _Version        =   1310723
      _ExtentX        =   9340
      _ExtentY        =   556
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCxC_ContratosPagadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub chkTodos_Click()
Call sbConsulta
End Sub

Private Sub Form_Activate()
vModulo = 31
End Sub

Private Sub Form_Load()
vModulo = 31

vGrid.MaxCols = 5
vGrid.MaxRows = 0

txtCodigo.Text = GLOBALES.gTag
txtDescripcion.Text = GLOBALES.gTag2

Call sbConsulta


Call Formularios(Me)
Call RefrescaTags(Me)
End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim bWhere As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

bWhere = True

strSQL = "select Per.CEDULA,Per.NOMBRE,ISNULL(Cnt.registro_usuario,'') as 'Usuario', ISNULL(Cnt.registro_fecha,'') as 'Fecha'" _
       & ", case when ISNULL(cod_contrato, 'NoExiste') = 'NoExiste' then 0 else 1 end as 'Activo'"

If chkTodos.Value = vbUnchecked Then
   strSQL = strSQL & " from CXC_PERSONAS Per left join CXC_CONTRATOS_PAGADORES Cnt" _
       & " on Per.CEDULA = Cnt.Cedula and Cnt.cod_contrato = '" & txtCodigo.Text & "'"
Else
   strSQL = strSQL & " from CXC_PERSONAS Per inner join CXC_CONTRATOS_PAGADORES Cnt" _
       & " on Per.CEDULA = Cnt.Cedula and Cnt.cod_contrato = '" & txtCodigo.Text & "'"
End If

If Len(Trim(txtCedula.Text)) > 0 Then
  If bWhere Then
     strSQL = strSQL & " Where Per.cedula like '%" & txtCedula.Text & "%'"
     bWhere = False
  Else
     strSQL = strSQL & " And Per.cedula like '%" & txtCedula.Text & "%'"
  End If
End If


If Len(Trim(txtNombre.Text)) > 0 Then
  If bWhere Then
     strSQL = strSQL & " Where Per.Nombre like '%" & txtNombre.Text & "%'"
     bWhere = False
  Else
     strSQL = strSQL & " And Per.Nombre like '%" & txtNombre.Text & "%'"
  End If
End If

If bWhere Then
    strSQL = strSQL & " Where Per.Rol_Pagador = 1" _
           & " order by Per.nombre,Cnt.cod_contrato desc"
Else
    strSQL = strSQL & " and Per.Rol_Pagador = 1" _
           & " order by Per.nombre,Cnt.cod_contrato desc"
End If
Call sbCargaGrid(vGrid, 5, strSQL, True)


vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsulta
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.col = col

If vGrid.Value = vbChecked Then
    vGrid.col = 1
    strSQL = "insert CxC_Contratos_Pagadores(cod_contrato,cedula,registro_fecha,registro_usuario) values('" _
           & txtCodigo.Text & "','" & vGrid.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
    Call Bitacora("Registra", "Pagador Id.:" & vGrid.Text & " de Contrato No.:" & txtCodigo.Text)
Else
    vGrid.col = 1
    strSQL = "Delete CxC_Contratos_Pagadores where cod_contrato = '" & txtCodigo.Text & "' and cedula = '" _
           & vGrid.Text & "'"
    Call Bitacora("Borra", "Pagador Id.:" & vGrid.Text & " de Contrato No.:" & txtCodigo.Text)
End If
Call ConectionExecute(strSQL)

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub
