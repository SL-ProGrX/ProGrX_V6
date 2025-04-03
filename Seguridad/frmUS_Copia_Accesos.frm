VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmUS_Copia_Accesos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Copia de Accesos entre Usuarios"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2655
      Left            =   240
      TabIndex        =   20
      Top             =   5040
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8493
      _ExtentY        =   4683
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
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbAplica 
      Height          =   855
      Left            =   120
      TabIndex        =   18
      Top             =   7920
      Width           =   11175
      _Version        =   1441792
      _ExtentX        =   19711
      _ExtentY        =   1508
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   615
         Left            =   9480
         TabIndex        =   19
         Top             =   240
         Width           =   1695
         _Version        =   1441792
         _ExtentX        =   2990
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Copiar"
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
         Picture         =   "frmUS_Copia_Accesos.frx":0000
      End
   End
   Begin XtremeSuiteControls.CheckBox chkRS_Inicializa 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario destino conserva sus roles actuales?"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCliente 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   11415
      _Version        =   1441792
      _ExtentX        =   20135
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsBase 
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtUsDestino 
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtUsBaseDesc 
      Height          =   330
      Left            =   4320
      TabIndex        =   6
      Top             =   1560
      Width           =   5895
      _Version        =   1441792
      _ExtentX        =   10398
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
   Begin XtremeSuiteControls.FlatEdit txtUsDestinoDesc 
      Height          =   330
      Left            =   4320
      TabIndex        =   7
      Top             =   2040
      Width           =   5895
      _Version        =   1441792
      _ExtentX        =   10398
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
   Begin XtremeSuiteControls.CheckBox chkRO_Inicializa 
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   3240
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario destino conserva sus roles actuales?"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Creditos 
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   4080
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CREDITOS: Copia Equipos de Trabajo y Niveles de Resolución?"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Cobros 
      Height          =   375
      Left            =   5640
      TabIndex        =   13
      Top             =   4560
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "COBROS: Copia Ejecutivo de Cobros"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Bancos 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   5040
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "BANCOS: Copia Perfil de Acceso a Cuentas Bancarias"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Contabilidad 
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   6000
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CONTABILIDAD: Copia Acceso a Contabilidades"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Inventarios 
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   6960
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "INVENTARIOS: Perfiles de Acceso y Autorización"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Compras 
      Height          =   375
      Left            =   5640
      TabIndex        =   17
      Top             =   7440
      Width           =   5775
      _Version        =   1441792
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "COMPRAS: Perfiles de Autorizaciones"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRS_Roles 
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   3720
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Copia Roles de Seguridad?"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRS_Horarios 
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   4080
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Copia Horarios de Trabajo?"
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
   End
   Begin XtremeSuiteControls.CheckBox chkRS_Estaciones 
      Height          =   255
      Left            =   600
      TabIndex        =   23
      Top             =   4440
      Width           =   4575
      _Version        =   1441792
      _ExtentX        =   8070
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Copia Estaciones de acceso?"
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
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Cajas 
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   5520
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "CAJAS: Copia Cajas Default"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Presupuesto 
      Height          =   375
      Left            =   5640
      TabIndex        =   25
      Top             =   6480
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "PRESUPUESTO: Copia Perfiles de Acceso "
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkRO_Deducciones 
      Height          =   375
      Left            =   5640
      TabIndex        =   26
      Top             =   3600
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "DEDUCCIONES: Copia Acceso a Deductoras "
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
      Value           =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   2640
      Width           =   6495
      _Version        =   1441792
      _ExtentX        =   11456
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Roles de Operativos"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   4935
      _Version        =   1441792
      _ExtentX        =   8705
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Roles de Seguridad"
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11535
      _Version        =   1441792
      _ExtentX        =   20346
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Copia de Accesos entre Usuarios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario Destino:"
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
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario Base:"
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
Attribute VB_Name = "frmUS_Copia_Accesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbRoles_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass



Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCopiar_Click()

On Error GoTo vError

If txtUsBase.Text = "" Then
  MsgBox "No Indicó al usuario BASE para referencia de Copia!", vbExclamation
  Exit Sub
End If

If txtUsDestino.Text = "" Then
  MsgBox "No Indicó al usuario DESTINO para la Copia!", vbExclamation
  Exit Sub
End If


If txtUsDestino.Text = txtUsBase.Text Then
  MsgBox "No Indicó al usuario ORIGEN Y DESTINO no puede ser el mismo!", vbExclamation
  Exit Sub
End If

Me.MousePointer = vbHourglass

'Copia Roles de Seguridad
strSQL = "exec spSEG_Copia_Permisos " & txtCliente.Tag & ",'" & txtUsBase.Text & "','" & txtUsDestino.Text & "','" & glogon.Usuario _
        & "', " & chkRS_Roles.Value & ", " & chkRS_Estaciones.Value & ", " & chkRS_Horarios.Value & ", " & chkRS_Inicializa.Value
Call ConectionExecute(strSQL)

'Copia Roles Operativos
Call sbCore_Copy_Permisos(txtCliente.Tag, txtUsBase.Text, txtUsDestino.Text, 1, chkRO_Deducciones.Value, chkRO_Contabilidad.Value _
            , chkRO_Creditos.Value, chkRO_Creditos.Value, chkRO_Cobros.Value, chkRO_Cajas.Value, chkRO_Bancos.Value _
            , chkRO_Presupuesto.Value, chkRO_Inventarios.Value, chkRO_Compras.Value, chkRO_Inicializa.Value)

Call Bitacora("Aplica", "Copia de Permisos Empresa [" & txtCliente.Tag & "] del Usuario: " & txtUsBase.Text & " --> " & txtUsDestino.Text)

Me.MousePointer = vbDefault

MsgBox "Permisos Copiados del Usuario: " & txtUsBase.Text & " --> " & txtUsDestino.Text, vbInformation

txtUsBase.Text = ""
txtUsBaseDesc.Text = ""

txtUsDestino.Text = ""
txtUsDestinoDesc.Text = ""

lsw.ListItems.Clear

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
vModulo = 13

txtCliente.Tag = gPortal.Empresa_Id
txtCliente.Text = gPortal.Empresa_Name

With lsw.ColumnHeaders
    .Clear
    .Add , , "", lsw.Width - 100
End With

End Sub


Private Sub txtUsBase_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select Usuario,Nombre from vPGX_Usuarios_Empresa"
    gBusquedas.Filtro = " and cod_Empresa = " & gPortal.Empresa_Id
    frmBusquedas.Show vbModal
    txtUsBase.Text = gBusquedas.Resultado
    txtUsBaseDesc.Text = gBusquedas.Resultado2
End If
End Sub


Private Sub txtUsDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select Usuario,Nombre from vPGX_Usuarios_Empresa"
    gBusquedas.Filtro = " and cod_Empresa = " & gPortal.Empresa_Id
    frmBusquedas.Show vbModal
    txtUsDestino.Text = gBusquedas.Resultado
    txtUsDestinoDesc.Text = gBusquedas.Resultado2
End If
End Sub
