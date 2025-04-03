VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_Ajustes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ajustes "
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10545
   HelpContextID   =   1012
   Icon            =   "frmAF_Ajustes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbAplicar 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
      _ExtentY        =   1720
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   495
         Left            =   8760
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_Ajustes.frx":000C
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.RadioButton optAjustes 
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Tipo de Identificación"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin VB.Frame fraTrabajo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1692
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   8295
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   330
         Left            =   1680
         TabIndex        =   7
         Top             =   120
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11245
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
      End
      Begin XtremeSuiteControls.FlatEdit txtCTDesc 
         Height          =   330
         Left            =   2520
         TabIndex        =   12
         Top             =   1320
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
         Height          =   330
         Left            =   1680
         TabIndex        =   13
         Top             =   600
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
         Height          =   330
         Left            =   2520
         TabIndex        =   14
         Top             =   600
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
         Height          =   330
         Left            =   1680
         TabIndex        =   15
         Top             =   960
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSecDesc 
         Height          =   330
         Left            =   2520
         TabIndex        =   16
         Top             =   960
         Width           =   5535
         _Version        =   1441793
         _ExtentX        =   9763
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCTCodigo 
         Height          =   330
         Left            =   1680
         TabIndex        =   17
         Top             =   1320
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblDepartamento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblSeccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblCentroTrabajo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro Trabajo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Institución"
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
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8493
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   312
      Left            =   3840
      TabIndex        =   6
      Top             =   1440
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8493
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   End
   Begin XtremeSuiteControls.RadioButton optAjustes 
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado de la Persona"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.RadioButton optAjustes 
      Height          =   852
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Empresa/Institución"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Persona"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Id."
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ajustes "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1884
      TabIndex        =   0
      Top             =   360
      Width           =   3612
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmAF_Ajustes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strInstAnterior As String
Dim strIdentAnterior As String
Dim strEstadoAnterior As String
Dim vDetalle As String, vTipoCambio As String
Dim vInstitucion As String
Dim bCarga As Boolean


Private Sub btnAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vTipoCambio = ""
vDetalle = ""

  
    Select Case True
      Case optAjustes(0).Value 'tipo identificacion
           If cboTipoId.Tag <> cboTipoId.ItemData(cboTipoId.ListIndex) Then
              Call sbCambioIdentificacion
           End If
           
           
           
      Case optAjustes(1).Value ' Estado
           If cboEstado.Tag <> cboEstado.ItemData(cboEstado.ListIndex) Then
              Call sbCambioEstado
           End If
      
      Case optAjustes(2).Value ' institucion
          If GLOBALES.SysASEVersion Then
             Call sbCambioInstitucionAseccss
          Else
             Call sbCambioinstitucion
          End If
           
    End Select
    
        
If vDetalle <> "" Then
    Call Bitacora("Modifica", vDetalle)

    If vParametros.BitacoraEspecial Then
      If vTipoCambio <> "" Then
        Call sbgAFIBitacora(vTipoCambio, vDetalle, Trim(GLOBALES.gCedulaActual))
      End If
    End If
End If
MsgBox "Información Actualizada satisfactoriamente...", vbInformation
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()

If bCarga Then
    
    Call sbCargaDatos
    
    txtDeptCodigo.Tag = Trim(txtDeptCodigo.Text)
    txtSecCodigo.Tag = Trim(txtSecCodigo.Text)
    txtCTCodigo.Tag = Trim(txtCTCodigo.Text)
     
    cboTipoId.Tag = cboTipoId.ItemData(cboTipoId.ListIndex)
    cboInstitucion.Tag = cboInstitucion.ItemData(cboInstitucion.ListIndex)
    
    strInstAnterior = Trim(cboInstitucion.Text)
    strIdentAnterior = Trim(cboTipoId.Text)
    strEstadoAnterior = Trim(cboEstado.Text)
    cboEstado.Tag = GLOBALES.gTag
    
End If

bCarga = False

End Sub

Private Sub Form_Load()
Dim strSQL As String, i As Integer

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Me.Caption = "Ajustes  [Cédula : " & GLOBALES.gCedulaActual & " " & fxNombre(GLOBALES.gCedulaActual) & "]"
bCarga = True

If GLOBALES.SysASEVersion Then
   lblDepartamento.Caption = "U. Programatica"
   lblSeccion.Caption = "U. Trabajo"
   
Else
   lblDepartamento.Caption = "Departamento"
   lblSeccion.Caption = "Sección"
'   lblCentroTrabajo.Visible = False
   txtCTCodigo.Visible = False
   txtCTDesc.Visible = False
End If

strSQL = "select COD_INSTITUCION as Idx, rtrim(Descripcion) as ItmX from INSTITUCIONES ORDER BY COD_INSTITUCION"
Call sbCbo_Llena_New(cboInstitucion, strSQL, False, True)

strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS"
Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
   
'Carga Estados de la Persona
strSQL = "select RTRIM(E.COD_ESTADO) as 'IdX', rtrim(E.DESCRIPCION) AS itmX" _
       & " from AFI_ESTADOS_PERSONA E" _
       & " where E.ACTIVO  =  1"
Call sbCbo_Llena_New(cboEstado, strSQL, False, True)


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub optAjustes_Click(Index As Integer)

If optAjustes(2).Value Then
    cboTipoId.Enabled = False
    fraTrabajo.Enabled = True
Else
    cboTipoId.Enabled = True
    fraTrabajo.Enabled = False
End If

End Sub

Private Sub txtCTCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then txtCTDesc.Text = fxgAFIDepartamento(cboInstitucion.ItemData(cboInstitucion.ListIndex), txtCTCodigo.Text)

If KeyCode = vbKeyF4 Then
 If GLOBALES.SysASEVersion Then
    gBusquedas.Columna = "codigo"
    gBusquedas.Orden = "codigo"
    gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
    gBusquedas.Filtro = ""
 Else
    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
 End If
  
  frmBusquedas.Show vbModal
  txtCTCodigo.Text = gBusquedas.Resultado
  txtCTDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCTCodigo_LostFocus()
 txtCTDesc.Text = fxgAFIDepartamento(cboInstitucion.ItemData(cboInstitucion.ListIndex), txtCTCodigo.Text)
End Sub

Private Sub txtCTDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
 If GLOBALES.SysASEVersion Then
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
    gBusquedas.Filtro = ""
 Else
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
 End If
  
  frmBusquedas.Show vbModal
  txtCTCodigo.Text = gBusquedas.Resultado
  txtCTDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then txtDeptDesc.Text = fxgAFIDepartamento(cboInstitucion.ItemData(cboInstitucion.ListIndex), txtDeptCodigo.Text)

If KeyCode = vbKeyF4 Then
    If GLOBALES.SysASEVersion Then
       gBusquedas.Columna = "codigo"
       gBusquedas.Orden = "codigo"
       gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
       gBusquedas.Filtro = ""
    Else
       gBusquedas.Columna = "cod_departamento"
       gBusquedas.Orden = "cod_departamento"
       gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
       gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    End If
     
     frmBusquedas.Show vbModal
     txtDeptCodigo.Text = gBusquedas.Resultado
     txtDeptDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDeptCodigo_LostFocus()
 txtDeptDesc.Text = fxgAFIDepartamento(cboInstitucion.ItemData(cboInstitucion.ListIndex), txtDeptCodigo.Text)
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 If GLOBALES.SysASEVersion Then
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
    gBusquedas.Filtro = ""
 Else
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from AFDepartamentos"
    gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
 End If
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If


End Sub

Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then txtSecDesc.Text = fxgAFISeccion(cboInstitucion.ItemData(cboInstitucion.ListIndex), txtDeptCodigo.Text, txtSecCodigo.Text)
If KeyCode = vbKeyF4 Then
  If GLOBALES.SysASEVersion Then
        gBusquedas.Columna = "ut_codigo"
        gBusquedas.Orden = "ut_codigo"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from utrabajo"
        gBusquedas.Filtro = ""
  Else
        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from AFSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                  & " and cod_departamento = '" & txtDeptCodigo & "'"
  End If
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtSecCodigo_LostFocus()
 txtSecDesc.Text = fxgAFISeccion(cboInstitucion.ItemData(cboInstitucion.ListIndex), txtDeptCodigo.Text, txtSecCodigo.Text)
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  
 If GLOBALES.SysASEVersion Then
        gBusquedas.Columna = "ut_descripcion"
        gBusquedas.Orden = "ut_descripcion"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from utrabajo"
        gBusquedas.Filtro = ""
  Else
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from AFSecciones"
        gBusquedas.Filtro = " and cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
                  & " and cod_departamento = '" & txtDeptCodigo & "'"
  End If
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
  
End If

End Sub

Private Sub sbCambioIdentificacion()
Dim strSQL As String

vDetalle = "Identificación de " & strIdentAnterior & " a " & cboTipoId & " " & GLOBALES.gCedulaActual
vTipoCambio = "26"
strSQL = "Update socios set tipo_id = " & cboTipoId.ItemData(cboTipoId.ListIndex) & " where cedula = '" & GLOBALES.gCedulaActual & "' "
Call ConectionExecute(strSQL)
        
End Sub

Private Sub sbCambioEstado()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""

'Verificar el Estado de la Persona si está autorizado en la institución
strSQL = "select count(*) as Existe from AFI_ESTADOS_INSTITUCIONES" _
       & " where cod_estado = '" & cboEstado.ItemData(cboEstado.ListIndex) _
       & "' and cod_institucion in(select cod_institucion from socios where cedula = '" & GLOBALES.gCedulaActual & "')"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = vMensaje & " - El ESTADO de la Persona a modificar o incluir no está autorizado en esta institución: " & cboInstitucion.Text & vbCrLf
End If
rs.Close


strSQL = "Select  isnull(aporte,0) as aporte from Ahorro_consolidado where  cedula = '" & GLOBALES.gCedulaActual & "' and Aporte > 0"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Or Not rs.BOF Then
    If rs!Aporte > 0 Then
        vMensaje = vMensaje & " - No procede el cambio de estado  por cuanto esta persona ya tiene Aporte registrado, verifique..." & vbCrLf
    End If
End If
rs.Close

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If

vDetalle = "Estado de  de " & strEstadoAnterior & " a " & cboEstado & " " & GLOBALES.gCedulaActual
vTipoCambio = "27"

strSQL = "Update socios set estadoActual = '" & cboEstado.ItemData(cboEstado.ListIndex) & "' where cedula = '" & GLOBALES.gCedulaActual & "' "
Call ConectionExecute(strSQL)

End Sub


Private Sub sbCambioInstitucionAseccss()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strUp As String, strUt As String, strCt As String
Dim vMensaje As String

vMensaje = ""

'Verificar el Estado de la Persona si está autorizado en la institución
strSQL = "select count(*) as Existe from AFI_ESTADOS_INSTITUCIONES" _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and cod_estado in(select estadoActual from socios where cedula = '" & GLOBALES.gCedulaActual & "')"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = vMensaje & " - El ESTADO de la Persona a modificar o incluir no está autorizado en esta institución: " & cboInstitucion.Text & vbCrLf
End If
rs.Close

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If


vDetalle = "Ajustes en valores tabajo " & GLOBALES.gCedulaActual

If cboInstitucion.Tag <> cboInstitucion.ItemData(cboInstitucion.ListIndex) Then
    vDetalle = "Institución de " & strInstAnterior & " a " & cboInstitucion & " " & GLOBALES.gCedulaActual
    vTipoCambio = "01"
    vInstitucion = "cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
Else
    vInstitucion = ""
    vTipoCambio = ""
End If

If txtDeptCodigo.Tag <> Trim(txtDeptCodigo.Text) Then
    strUp = "up = '" & txtDeptCodigo.Text & "'"
Else
    strUp = ""
End If

If txtSecCodigo.Tag <> Trim(txtSecCodigo.Text) Then
    strUt = "Ut = '" & txtSecCodigo.Text & "'"
Else
    strUt = ""
End If

If txtCTCodigo.Tag <> Trim(txtCTCodigo.Text) Then
    strCt = "ct = '" & txtCTCodigo.Text & "'"
Else
    strCt = ""
End If

strSQL = "update socios set "
strSQL = strSQL & IIf(vInstitucion = "", "", vInstitucion)

If vInstitucion = "" Then
    strSQL = strSQL & strUp
Else
    strSQL = strSQL & "," & strUp
End If

If strUp = "" Then
    strSQL = strSQL & strUt
Else
    strSQL = strSQL & "," & strUt
End If

If strUt = "" Then
    strSQL = strSQL & strCt
Else
    strSQL = strSQL & "," & strCt
End If

strSQL = IIf(Mid(strSQL, Len(strSQL), 1) = ",", Mid(strSQL, 1, Len(strSQL) - 1) & " ", strSQL)

If Trim(strSQL) <> "update socios set " Then
    strSQL = strSQL & " where cedula = '" & GLOBALES.gCedulaActual & "'"
    Call ConectionExecute(strSQL)
End If

End Sub

Private Sub sbCambioinstitucion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, vCambia As Boolean
Dim vDetalle As String

vMensaje = ""
vCambia = False

'Verificar el Estado de la Persona si está autorizado en la institución
strSQL = "select count(*) as Existe from AFI_ESTADOS_INSTITUCIONES" _
       & " where cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " and cod_estado in(select estadoActual from socios where cedula = '" & GLOBALES.gCedulaActual & "')"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = vMensaje & " - El ESTADO de la Persona a modificar o incluir no está autorizado en esta institución: " & cboInstitucion.Text & vbCrLf
End If
rs.Close

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If



vDetalle = ""
If cboInstitucion.Tag <> cboInstitucion.ItemData(cboInstitucion.ListIndex) Then
    vDetalle = "Institución: " & cboInstitucion.Tag & " -> " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
    vCambia = True
End If
vInstitucion = "cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""

If txtDeptCodigo.Tag <> Trim(txtDeptCodigo.Text) Then
    vCambia = True
    vDetalle = vDetalle & "¦ Dept: " & txtDeptCodigo.Tag & " -> " & Trim(txtDeptCodigo.Text)
End If

If txtSecCodigo.Tag <> Trim(txtSecCodigo.Text) Then
    vCambia = True
    vDetalle = vDetalle & "¦ Sección: " & txtSecCodigo.Tag & " -> " & Trim(txtSecCodigo.Text)
End If

'Registra Cambio
strSQL = "update socios set cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) _
       & " ,cod_departamento = '" & txtDeptCodigo.Text _
       & "',cod_seccion = '" & txtSecCodigo.Text _
       & "' where cedula = '" & GLOBALES.gCedulaActual & "'"
Call ConectionExecute(strSQL)

'Registra Bitácora
If vCambia Then
    Call sbgAFIBitacora("01", vDetalle, Trim(GLOBALES.gCedulaActual))
End If


'Tipo_Cambio = '01' > Bitacora

End Sub


Private Sub sbCargaDatos()

Dim strSQL As String, rs As New ADODB.Recordset

If Not GLOBALES.SysASEVersion Then
    strSQL = "Select S.*,Est.Descripcion as 'EstadoPersonaDesc',Est.Cod_Estado + ' - ' + Est.Descripcion as 'EstadoPersona'" _
           & ",I.descripcion as DescInst,D.descripcion as DescDept,X.descripcion as DescSec" _
           & ",Tid.Descripcion as TipoIdDesc" _
           & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
           & " left join AFDepartamentos D on S.cod_institucion = D.cod_institucion and S.cod_departamento = D.cod_departamento" _
           & " left join AFSecciones X on S.cod_institucion = X.cod_institucion" _
           & "  and S.cod_departamento = X.cod_departamento and S.cod_seccion = X.cod_seccion" _
           & " inner join AFI_ESTADOS_PERSONA Est on S.EstadoActual = Est.Cod_Estado" _
           & " left join AFI_TIPOS_IDS Tid on S.tipo_id = Tid.tipo_id" _
           & " where cedula='" & Trim(GLOBALES.gCedulaActual) & "'"
Else
    strSQL = "Select S.*,UT as 'Cod_Seccion',UP as 'Cod_Departamento',C.descripcion as 'CentroDesc'" _
           & ",Est.Descripcion as 'EstadoPersonaDesc',Est.Cod_Estado + ' - ' + Est.Descripcion as 'EstadoPersona'" _
           & ",I.descripcion as DescInst,D.descripcion as DescDept,X.ut_descripcion as DescSec" _
           & ",Tid.Descripcion as TipoIdDesc" _
           & " From socios S inner join Instituciones I on S.cod_institucion = I.cod_institucion" _
           & " left join uprogramatica D on S.UP = D.codigo" _
           & " left join utrabajo X on S.ut = X.ut_codigo" _
           & " left join uprogramatica C on S.CT = C.codigo" _
           & " inner join AFI_ESTADOS_PERSONA Est on S.EstadoActual = Est.Cod_Estado" _
           & " left join AFI_TIPOS_IDS Tid on S.tipo_id = Tid.tipo_id" _
           & " where cedula='" & Trim(GLOBALES.gCedulaActual) & "'"
End If
Call OpenRecordSet(rs, strSQL)

txtDeptCodigo = rs!cod_departamento & ""
txtDeptDesc = Trim(rs!descDept & "")
txtSecCodigo = rs!cod_seccion & ""
txtSecDesc = Trim(rs!DescSec & "")

If Not IsNull(rs!TipoIdDesc) Then
  Call sbCboAsignaDato(cboTipoId, Trim(rs!TipoIdDesc), True, rs!Tipo_Id)
End If

Call sbCboAsignaDato(cboInstitucion, Trim(rs!DescInst), True, rs!cod_institucion)

Call sbCboAsignaDato(cboEstado, rs!EstadoPersonaDesc, True, rs!EstadoActual)

If GLOBALES.SysASEVersion Then
    lblCentroTrabajo.Visible = True
    txtCTCodigo.Visible = True
    txtCTDesc.Visible = True
    
    txtCTCodigo.Text = rs!CT & ""
    txtCTDesc.Text = rs!CentroDesc & ""
End If

rs.Close
End Sub

