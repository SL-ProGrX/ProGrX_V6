VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmSYS_CORE_Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información de Usuarios: Operativos"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   5520
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   10455
      _Version        =   1441793
      _ExtentX        =   18441
      _ExtentY        =   9128
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Usuario"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "txtNombre"
      Item(0).Control(1)=   "txtTelCelular"
      Item(0).Control(2)=   "txtEmail"
      Item(0).Control(3)=   "txtNotas"
      Item(0).Control(4)=   "Label2"
      Item(0).Control(5)=   "Label1(3)"
      Item(0).Control(6)=   "Label1(7)"
      Item(0).Control(7)=   "Label1(6)"
      Item(0).Control(8)=   "Label3"
      Item(0).Control(9)=   "txtUsuarioSistema"
      Item(0).Control(10)=   "gbImportar"
      Item(0).Control(11)=   "chkActivo"
      Item(1).Caption =   "Miembro de"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "tcAux"
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1441793
         _ExtentX        =   18441
         _ExtentY        =   10398
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
         Appearance      =   4
         Color           =   32
         ItemCount       =   2
         Item(0).Caption =   "UENS"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lswMiembros"
         Item(1).Caption =   "Roles"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "vgRoles"
         Begin XtremeSuiteControls.ListView lswMiembros 
            Height          =   4335
            Left            =   0
            TabIndex        =   25
            Top             =   360
            Width           =   10455
            _Version        =   1441793
            _ExtentX        =   18441
            _ExtentY        =   7646
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
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin FPSpreadADO.fpSpread vgRoles 
            Height          =   4335
            Left            =   -70000
            TabIndex        =   26
            Top             =   360
            Visible         =   0   'False
            Width           =   10935
            _Version        =   524288
            _ExtentX        =   19288
            _ExtentY        =   7646
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
            SpreadDesigner  =   "frmSYS_CORE_Usuarios.frx":0000
            VScrollSpecialType=   2
            Appearance      =   1
            AppearanceStyle =   1
         End
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   255
         Left            =   7800
         TabIndex        =   23
         Top             =   600
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activo?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuarioSistema 
         Height          =   315
         Left            =   2280
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
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
         Left            =   2280
         TabIndex        =   2
         Top             =   960
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11239
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtTelCelular 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   315
         Left            =   2280
         TabIndex        =   4
         Top             =   1680
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11239
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1155
         Left            =   2280
         TabIndex        =   5
         Top             =   2040
         Width           =   6375
         _Version        =   1441793
         _ExtentX        =   11245
         _ExtentY        =   2037
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbImportar 
         Height          =   1335
         Left            =   360
         TabIndex        =   21
         Top             =   3720
         Width           =   9255
         _Version        =   1441793
         _ExtentX        =   16325
         _ExtentY        =   2355
         _StockProps     =   79
         Caption         =   "Importar de Usuarios del Sistema"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnImportar 
            Height          =   615
            Left            =   3480
            TabIndex        =   22
            Top             =   480
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Importar Usuarios del Sistema Activos"
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
            Picture         =   "frmSYS_CORE_Usuarios.frx":335A
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario Sistema"
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
         Left            =   720
         TabIndex        =   10
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Tel. Celular"
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
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Email"
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
         Index           =   7
         Left            =   720
         TabIndex        =   8
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Index           =   6
         Left            =   720
         TabIndex        =   7
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
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
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
   End
   Begin XtremeSuiteControls.PushButton btnExiste 
      Height          =   315
      Left            =   4440
      TabIndex        =   11
      Top             =   600
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Existe?"
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
      Appearance      =   17
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtUserName 
      Height          =   435
      Left            =   1080
      TabIndex        =   13
      Top             =   600
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_CORE_Usuarios.frx":3A62
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2160
      TabIndex        =   15
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_CORE_Usuarios.frx":4094
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   2520
      TabIndex        =   16
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_CORE_Usuarios.frx":468F
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   17
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_CORE_Usuarios.frx":4C33
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_CORE_Usuarios.frx":5364
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   3960
      TabIndex        =   19
      ToolTipText     =   "Reporte"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_CORE_Usuarios.frx":5A64
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.Label Label7 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Usuario:"
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSYS_CORE_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vBusca As Integer, vScroll As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String


Select Case Index
    Case 0 'NUEVO
      vEdita = False
      Call sbLimpiaPantalla
      txtUserName.SetFocus
      Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
        vEdita = True
        txtNombre.SetFocus
        Call sbBarra_Accion("Editar")
     
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
     
      Call sbBarra_Accion("Editar")
      If txtUserName.Text = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      Else
        Call sbConsulta(txtUserName.Text)
      End If
    
    Case 5 'REPORTES

   
End Select

End Sub


Private Sub btnExiste_Click()
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select count(*) as 'Existe' from CORE_USUARIOS" _
       & " where CORE_USUARIO = '" & txtUserName.Text & "'"
       
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = "USUARIO: Libre!"
Else
    vMensaje = "USUARIO: Ocupado!"
End If
rs.Close

Me.MousePointer = vbDefault

MsgBox vMensaje, vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnImportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCORE_Usuarios_Importar"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Usuarios del Sistema Sincronizados/Importados Satisfactoriamente!", vbInformation
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 CORE_USUARIO from CORE_USUARIOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where CORE_USUARIO > '" & txtUserName & "'"
    Else
       strSQL = strSQL & " where CORE_USUARIO < '" & txtUserName & "'"
    End If
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " order by CORE_USUARIO asc"
    Else
       strSQL = strSQL & " order by CORE_USUARIO desc"
    End If
    
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtUserName = rs!CORE_USUARIO
      Call sbConsulta(txtUserName)
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()

vModulo = 10

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True


With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "", lswMiembros.Width - 100
End With

 vEdita = True
 
 Call sbBarra_Accion("Activo")
 
 Call sbLimpiaPantalla
 

 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
End Sub


Private Sub sbLimpiaPantalla()

vBusca = 1

tcMain.Item(0).Selected = True

txtUserName.Text = ""
txtNombre.Text = ""
txtNotas.Text = ""


txtUsuarioSistema.Text = ""
txtEmail.Text = ""
txtTelCelular.Text = ""

chkActivo.Value = xtpChecked

tcMain.Item(0).Selected = True
tcMain.Item(1).Enabled = False

End Sub



Private Sub lswMiembros_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswMiembros.SortKey = ColumnHeader.Index - 1
  If lswMiembros.SortOrder = 0 Then lswMiembros.SortOrder = 1 Else lswMiembros.SortOrder = 0
  lswMiembros.Sorted = True
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spSys_UENS_Miembros_Registro '" & Item.Tag & "', '" & txtUserName.Text _
       & "', '" & glogon.Usuario & "', '" & IIf((Item.Checked), "A", "E") & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Pass = 1 Then
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbMiembros_List()

If vPaso Then Exit Sub
If txtUserName.Text = "" Then Exit Sub

Me.MousePointer = vbHourglass

vPaso = True


With lswMiembros.ListItems
 .Clear
  
  
    strSQL = "exec spSys_CORE_Users_UENs_Miembros_Consultas '" & txtUserName.Text & "',''"
    Call OpenRecordSet(rs, strSQL)
    
     .Clear
     Do While Not rs.EOF
      Set itmX = .Add(, , rs!DESCRIPCION)
          itmX.Tag = rs!COD_UNIDAD
          
          If rs!ASIGNADO = 1 Then
             itmX.ForeColor = vbBlue
             itmX.Checked = True
          End If
          
      rs.MoveNext
     Loop
    rs.Close
    
  
End With


vPaso = False


Me.MousePointer = vbDefault

End Sub


Private Sub sbRoles_List()

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "exec spSys_CORE_Users_UENs_Roles_Consultas '" & txtUserName.Text & "', ''"
Call OpenRecordSet(rs, strSQL)

With vgRoles
    .MaxRows = 0
    
    Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     
     .Col = 1
     .Text = rs!COD_UNIDAD
     .Col = 2
     .Text = rs!DESCRIPCION
     .Col = 3
     .Value = rs!ROL_SOLICITA
     .Col = 4
     .Value = rs!ROL_CONSULTA
     .Col = 5
     .Value = rs!ROL_AUTORIZA
     .Col = 6
     .Value = rs!ROL_ENCARGADO
     
     rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault

vPaso = False


Exit Sub

vError:
    Me.MousePointer = vbDefault
    vPaso = False
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0
        Call sbMiembros_List
    Case 1
        Call sbRoles_List
End Select


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vError

Select Case Item.Index
  Case 0 'Usuarios
  Case 1 'Miembro de...
    tcAux(0).Selected = True
    Call sbMiembros_List
End Select



Exit Sub


vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Public Sub sbConsulta(pUsuario As String)

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "Select *" _
       & " from CORE_USUARIOS" _
       & " where CORE_USUARIO = '" & pUsuario & "'"
       
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    Call sbBarra_Accion("activo")

    vEdita = True
    
    tcMain.Item(0).Selected = True
    tcMain.Item(1).Enabled = True
    
    
    txtUserName = Trim(rs!CORE_USUARIO)
    txtUsuarioSistema.Text = rs!Usuario_Ref & ""
    
    txtNombre = IIf(IsNull(rs!Nombre), "", rs!Nombre)
    txtNotas = IIf(IsNull(rs!NOTAS), "", rs!NOTAS)
    
    txtEmail.Text = rs!EMAIL & ""
    txtTelCelular.Text = rs!Tel_Movil & ""
        
    chkActivo.Value = rs!Activo
Else
   Call sbLimpiaPantalla
End If

rs.Close
Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""

fxValida = True


'Valida que el nombre de CORE_USUARIO esté desocupado
If Not vEdita Then
    strSQL = "select count(*) as 'Existe' from CORE_USUARIOS" _
           & " where CORE_USUARIO = '" & txtUserName.Text & "'"
           
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
        vMensaje = vMensaje & vbCrLf & " - Este USUARIO se encuentra -Ocupado- debe cambiar el nombre de USUARIO por otro que se encuentre -Libre-"
    End If
End If


vMensaje = ""

If Trim(txtUserName.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - No a indicado el Nombre de CORE_USUARIO!"
End If

If Trim(txtNombre.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - No a indicado el Nombre de la Persona!"
End If

If Trim(txtEmail.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - Indicar un Email válido"
End If


If Trim(txtTelCelular.Text) = "" Then
   vMensaje = vMensaje & vbCrLf & " - No a indicado un número de teléfono Móvil!"
End If
  
If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbGuardar()

On Error GoTo vError
 
If Not vEdita Then
  
  strSQL = "Insert CORE_USUARIOS(CORE_USUARIO, Usuario_Ref ,Nombre, Registro_Fecha, Registro_Usuario" _
         & ", Activo, Notas, Email, Tel_Movil )" _
         & " values('" & Trim(txtUserName) & "','" & Trim(txtUsuarioSistema.Text) & "','" & Trim(txtNombre) _
         & "', getdate(),'" & glogon.Usuario & "', " & chkActivo.Value _
         & ", '" & Trim(txtNotas) & "', '" & Trim(txtEmail) & "', '" & txtTelCelular.Text & "')"
        
Else
  strSQL = "Update CORE_USUARIOS Set Nombre = '" & Trim(txtNombre) & "', Activo = " & chkActivo.Value _
         & " , Usuario_Ref = '" & txtUsuarioSistema.Text _
         & "', Notas = '" & Trim(txtNotas.Text) _
         & "', Email = '" & Trim(txtEmail) & "',Tel_Movil = '" & Trim(txtTelCelular.Text) _
         & "', Modifica_Fecha = dbo.MyGetDate(), Modifica_Usuario = '" & glogon.Usuario _
         & "' Where CORE_USUARIO = '" & txtUserName.Text & "'"
End If
Call ConectionExecute(strSQL)


Call sbBarra_Accion("activo")

Call sbConsulta(txtUserName)
        
vEdita = True
        
MsgBox "Información guardada satisfactoriamente...", vbInformation

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then

'  strSQL = "delete Miembros where Nombre = '" & txtUserName & "'"
'  Call ConectionExecute(strSQL)
'
'  strSQL = "delete permisos where Tipo = 'U' and Nombre = '" & txtUserID & "'"
'  Call ConectionExecute(strSQL)
'
'  strSQL = "delete CORE_USUARIOs where UserID = " & txtUserID
'  Call ConectionExecute(strSQL)
'
'  Call Bitacora("Elimina", "CORE_USUARIO: " & txtUserName)

  Call sbLimpiaPantalla
  Call sbBarra_Accion("Nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtUsuarioSistema_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelCelular.SetFocus
End Sub


Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUsuarioSistema.SetFocus
End Sub

Private Sub txtTelCelular_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub


Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then

 If vEdita Then Call sbConsulta(txtUserName)
 txtUsuarioSistema.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    gBusquedas.Consulta = "select CORE_USUARIO,Nombre from CORE_USUARIOS"
    
    gBusquedas.Filtro = ""
    
    frmBusquedas.Show vbModal
    
    txtUserName = gBusquedas.Resultado
    txtNombre = gBusquedas.Resultado2
    
    Call sbConsulta(txtUserName)
    txtUserName.SetFocus
End If

End Sub

Private Sub vgRoles_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

On Error GoTo vError

Dim pUEN As String
Dim pRSolicita As Integer, pRConsulta As Integer, pRAutoriza As Integer, pREncargado As Integer

With vgRoles
    .Row = Row
    .Col = 1
    pUEN = .Text
    .Col = 3
    pRSolicita = .Value
    .Col = 4
    pRConsulta = .Value
    .Col = 5
    pRAutoriza = .Value
    .Col = 6
    pREncargado = .Value
    
    strSQL = "exec spSys_UENS_Roles_Registro '" & pUEN & "', '" & txtUserName.Text & "', " & pRSolicita _
           & ", " & pRConsulta & ", " & pRAutoriza & ", " & pREncargado & ", '" & glogon.Usuario & "'"
    
    Call ConectionExecute(strSQL)
    
End With

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

