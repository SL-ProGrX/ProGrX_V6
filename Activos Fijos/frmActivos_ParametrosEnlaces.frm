VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmActivos_ParametrosEnlaces 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parámetros de Enlaces"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   5
      Top             =   3276
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox chkProveedores 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actualiza Lista de Proveedores vrs ProGrX_ Comercial"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   1800
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox chkReponsables 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actualiza listado de Responsables vrs Planillas"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox chkSecciones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actualiza listado de Secciones vrs Planillas"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox chkPeriodos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actualiza Periodos para Cierres vrs Contabilidad"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1440
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin VB.CheckBox chkDepartamentos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Actualiza listado de Departamentos vrs Planillas"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Value           =   1  'Checked
      Width           =   4815
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      ToolTipText     =   "Importa Catálogo de Cuentas por Pagar"
      Top             =   2520
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Guardar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmActivos_ParametrosEnlaces.frx":0000
   End
   Begin VB.Label lblX 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   0
      TabIndex        =   6
      Top             =   2520
      Width           =   4332
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6720
      X2              =   0
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmActivos_ParametrosEnlaces"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbDepartamentos()

End Sub


Private Sub sbSecciones()

End Sub


Private Sub sbResponsables()

End Sub


Private Sub sbProveedores()
Dim strSQL As String, rs As New ADODB.Recordset

lblX.Caption = "Registrando Nuevos Proveedores..."
lblX.Refresh

'Registra Nuevos Proveedores
strSQL = "insert into Activos_proveedores(cod_proveedor,descripcion) (select cod_proveedor,descripcion" _
       & " From cxp_proveedores where cod_proveedor not in(select cod_proveedor from Activos_proveedores))"
Call ConectionExecute(strSQL)

End Sub


Private Sub cmdAplicar_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass
prgBar.Visible = True

If chkDepartamentos.Value = vbChecked Then Call sbDepartamentos
If chkSecciones.Value = vbChecked Then Call sbSecciones
If chkReponsables.Value = vbChecked Then Call sbResponsables
If chkProveedores.Value = vbChecked Then Call sbProveedores


lblX.Caption = ""
prgBar.Visible = False
Me.MousePointer = vbDefault

MsgBox "Actualizacion Finalizada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Sub Form_Load()
vModulo = 36
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub
