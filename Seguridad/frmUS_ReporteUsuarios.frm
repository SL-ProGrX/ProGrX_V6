VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmUS_ReporteUsuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reporte Generales de Seguridad"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   255
      Left            =   5040
      TabIndex        =   9
      Top             =   3360
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   7815
      _Version        =   1441793
      _ExtentX        =   13785
      _ExtentY        =   2566
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   615
         Left            =   5880
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Informe"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmUS_ReporteUsuarios.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
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
         Height          =   330
         Left            =   2160
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas"
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
   Begin XtremeSuiteControls.ComboBox cboReporte 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   1920
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   3360
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.ComboBox cboVinculacion 
      Height          =   330
      Left            =   2280
      TabIndex        =   11
      Top             =   2880
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.ComboBox cboRol 
      Height          =   315
      Left            =   2280
      TabIndex        =   17
      Top             =   3840
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Roles"
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Vinculación"
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
   Begin XtremeShortcutBar.ShortcutCaption lblCliente 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1320
      Width           =   8055
      _Version        =   1441793
      _ExtentX        =   14208
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   7
      Top             =   3360
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
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
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Reporte"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   8535
      _Version        =   1441793
      _ExtentX        =   15055
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Informes de Seguridad de Cuentas de Usuarios"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1230
      Left            =   0
      Top             =   0
      Width           =   13560
   End
End
Attribute VB_Name = "frmUS_ReporteUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
 dtpInicio.Enabled = False
 dtpCorte.Enabled = False
Else
 dtpInicio.Enabled = True
 dtpCorte.Enabled = True
End If
End Sub

Private Sub chkUsuarios_Click()
If chkUsuarios.Value = vbChecked Then
   txtUsuario.BackColor = vbWhite
Else
   txtUsuario.BackColor = vbGrayText
End If
End Sub

Private Sub btnReporte_Click()


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "SystemSecurity"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxFecha='Emitido el:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxUsuario='Usuario:" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxEmpresa='" & lblCliente.Caption & "'"
 
 .Formulas(3) = "fxSubTitulo='Estado: " & cboEstado.Text & " ¦ Vinculación: " & cboVinculacion.Text & "'"

If txtUsuario.Text = "" Then
    txtUsuario.Tag = "-x-"
Else
    txtUsuario.Tag = txtUsuario.Text
End If

 Select Case Mid(cboReporte.Text, 1, 2)
    
    Case "01" 'Listado de Usuarios
      'StoreProc:  spSEG_Informe_Usuarios_Lista(@EmpresaId int, @Usuario varchar(30) = Null, @Estado char(1) = Null, @Vinculado varchar(20) = Null, @Contabiliza smallint = 1)
      
      .ReportFileName = SIFGlobal.fxPathReportes("SysSec_Listado_Usuarios.rpt")
      .StoredProcParam(0) = lblCliente.Tag
      .StoredProcParam(1) = txtUsuario.Tag
      .StoredProcParam(2) = Mid(cboEstado.Text, 1, 1)
      .StoredProcParam(3) = cboVinculacion.Text
      .StoredProcParam(4) = 1
      
      
    Case "02"  ' Listado de Roles por Usuarios
      'StoreProc: spSEG_Informe_Usuarios_Roles(@EmpresaId int, @Usuario varchar(30) = Null, @Contabiliza smallint = 1)
      
      .ReportFileName = SIFGlobal.fxPathReportes("SysSec_Listado_Usuarios_Roles.rpt")
      
      .StoredProcParam(0) = lblCliente.Tag
      .StoredProcParam(1) = txtUsuario.Tag
      .StoredProcParam(2) = 1
      
    Case "03" 'Permisos x Usuarios
      
      'StoreProc: spSEG_Informe_Usuarios_Permisos(@EmpresaId int, @Usuario varchar(30) = Null, @Contabiliza smallint = 1)
      
      .ReportFileName = SIFGlobal.fxPathReportes("SysSec_Listado_Permisos_Usuarios.rpt")
      
      .StoredProcParam(0) = lblCliente.Tag
      .StoredProcParam(1) = txtUsuario.Tag
      .StoredProcParam(2) = 1
      
      
    Case "04" 'Permisos x Roles
      'StoreProc: spSEG_Informe_Roles_Permisos(@RolId varchar(10) = Null)
      
      .ReportFileName = SIFGlobal.fxPathReportes("SysSec_Listado_Permisos_Roles.rpt")
      
      If cboRol.Text = "TODOS" Then
          .StoredProcParam(0) = "-x-"
      Else
          .StoredProcParam(0) = cboRol.ItemData(cboRol.ListIndex)
      End If
 
 End Select
 
 .PrintReport
 
End With

End Sub


Private Sub Form_Load()
vModulo = 13


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

dtpInicio.Value = Format(fxFechaServidor, "dd/mm/yyyy")
dtpCorte.Value = dtpInicio.Value

lblCliente.Tag = gPortal.Empresa_Id
lblCliente.Caption = gPortal.Empresa_Name


cboEstado.AddItem "Activos"
cboEstado.AddItem "Inactivos"
cboEstado.AddItem "Todos"
cboEstado.Text = "Activos"

cboVinculacion.AddItem "Vigente"
cboVinculacion.AddItem "Eliminada"
cboVinculacion.AddItem "Todos"
cboVinculacion.Text = "Vigente"


cboReporte.AddItem "01 - Listado de Usuarios"
cboReporte.AddItem "02 - Listado de Usuarios y Roles"
cboReporte.AddItem "03 - Permisos por Usuario"
cboReporte.AddItem "04 - Permisos por Roles"
cboReporte.Text = "01 - Listado de Usuarios"

strSQL = "select COD_ROL as 'IdX', DESCRIPCION as 'ItmX'" _
       & " From US_ROLES Where ACTIVO = 1 order by DESCRIPCION"
Call sbCbo_Llena_New(cboRol, strSQL, True, True)

Call chkFechas_Click
Call chkUsuarios_Click

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Usuario, Nombre, Estado_Desc, Vinculacion From vPGX_Usuarios_Empresa_Historica"
    gBusquedas.Filtro = " and cod_Empresa = " & lblCliente.Tag
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    
    frmBusquedas.Show vbModal
    txtUsuario.Transparent = Trim(gBusquedas.Resultado)
End If

vError:

End Sub

