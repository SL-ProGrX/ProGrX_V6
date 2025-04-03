VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmUS_Roles_Membrecias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Roles: Membrecías"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswUsuarios 
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   9763
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
   Begin XtremeSuiteControls.ListView lswClienteUsuarios 
      Height          =   5535
      Left            =   4440
      TabIndex        =   10
      Top             =   2040
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   9763
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
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   240
      Top             =   120
   End
   Begin XtremeSuiteControls.PushButton btnOtros 
      Height          =   375
      Index           =   0
      Left            =   9360
      TabIndex        =   12
      Top             =   7680
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estaciones"
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
      Picture         =   "frmUS_Roles_Membrecias.frx":0000
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5535
      Left            =   8760
      TabIndex        =   11
      Top             =   2040
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
      _ExtentY        =   9763
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
      ItemCount       =   3
      Item(0).Caption =   "Roles"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lswRoles"
      Item(1).Caption =   "Estaciones"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "chkAcceso_Equipo"
      Item(1).Control(1)=   "lswEstaciones"
      Item(2).Caption =   "Horarios"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "chkAcceso_Horario"
      Item(2).Control(1)=   "lswHorarios"
      Begin XtremeSuiteControls.ListView lswRoles 
         Height          =   5175
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   9128
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
      Begin XtremeSuiteControls.ListView lswEstaciones 
         Height          =   4575
         Left            =   -70000
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   8070
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
      Begin XtremeSuiteControls.ListView lswHorarios 
         Height          =   4575
         Left            =   -70000
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   8070
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
      Begin XtremeSuiteControls.CheckBox chkAcceso_Equipo 
         Height          =   375
         Left            =   -69880
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Este usuario tiene acceso restringido por equipo?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkAcceso_Horario 
         Height          =   375
         Left            =   -69880
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Este usuario tiene acceso restringido por equipo?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
   End
   Begin XtremeSuiteControls.CheckBox chkContabiliza 
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   7680
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Usuarios Contabilizados"
      ForeColor       =   16711680
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
      Value           =   2
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCliente 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   13575
      _Version        =   1441793
      _ExtentX        =   23945
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "..."
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltroUsuarioAsg 
      Height          =   330
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.PushButton btnOtros 
      Height          =   375
      Index           =   1
      Left            =   10680
      TabIndex        =   13
      Top             =   7680
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Horarios"
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
      Picture         =   "frmUS_Roles_Membrecias.frx":0720
   End
   Begin XtremeSuiteControls.PushButton btnOtros 
      Height          =   375
      Index           =   2
      Left            =   12000
      TabIndex        =   21
      Top             =   7680
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Copia"
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
      Picture         =   "frmUS_Roles_Membrecias.frx":0FCC
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltroAccess 
      Height          =   330
      Left            =   8760
      TabIndex        =   22
      Top             =   1680
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
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
   Begin XtremeSuiteControls.Label lblCount 
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   20
      Top             =   7680
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Items:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label lblCount 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   7680
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Items:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Direcctorio de Usuarios de la Plataforma"
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   330
      Index           =   1
      Left            =   4440
      TabIndex        =   8
      Top             =   1320
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Usuarios vinculados al Cliente"
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
   Begin XtremeShortcutBar.ShortcutCaption lblUsuario 
      Height          =   330
      Left            =   8760
      TabIndex        =   6
      Top             =   1320
      Width           =   4935
      _Version        =   1441793
      _ExtentX        =   8705
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Usuario!"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuarios Vinculados al Cliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Usuarios, Accesos y Roles de Seguridad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmUS_Roles_Membrecias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnOtros_Click(Index As Integer)

Select Case Index
    Case 0 'Equipos
        Call sbFormsCall("frmUS_Access_Estaciones", vbModal, , , False, Me)
        
      
    Case 1 'Horarios
        Call sbFormsCall("frmUS_Access_Horarios", vbModal, , , False, Me)

    Case 2 'Copia
        Call sbFormsCall("frmUS_Copia_Accesos", vbModal, , , False, Me)
End Select

End Sub

Private Sub chkAcceso_Equipo_Click()

If vPaso Then Exit Sub

On Error GoTo vError
'spPGX_Usuario_Estacion_Limita (@Cliente int, @UsuarioLimita varchar(30), @Usuario varchar(30), @Limita smallint = 0)

If chkAcceso_Equipo.Value = vbChecked Then
    strSQL = "exec spPGX_Usuario_Estacion_Limita " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & glogon.Usuario & "',1"
Else
    strSQL = "exec spPGX_Usuario_Estacion_Limita " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & glogon.Usuario & "',0"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkAcceso_Horario_Click()

If vPaso Then Exit Sub

On Error GoTo vError
'spPGX_Usuario_Estacion_Limita (@Cliente int, @UsuarioLimita varchar(30), @Usuario varchar(30), @Limita smallint = 0)

If chkAcceso_Horario.Value = vbChecked Then
    strSQL = "exec spPGX_Usuario_Horario_Limita " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & glogon.Usuario & "',1"
Else
    strSQL = "exec spPGX_Usuario_Horario_Limita " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & glogon.Usuario & "',0"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkContabiliza_Click()
Call sbConsultaUsuariosVinculados
End Sub

Private Sub Form_activate()
vModulo = 13
End Sub

Private Sub Form_Load()
vModulo = 13

txtCliente.Tag = gPortal.Empresa_Id
txtCliente.Text = gPortal.Empresa_Name

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lswClienteUsuarios.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1800
    .Add , , "Nombre", lswClienteUsuarios.Width - 2000
End With

With lswUsuarios.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1800
    .Add , , "Nombre", lswUsuarios.Width - 2000
End With

With lswRoles.ColumnHeaders
    .Clear
    .Add , , "Rol Id", 1600
    .Add , , "Descripción", lswRoles.Width - 1800
End With

With lswEstaciones.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Descripción", lswEstaciones.Width - 2000
End With

With lswHorarios.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Descripción", lswHorarios.Width - 2000
End With


Call sbLimpia
Call sbConsultaUsuariosVinculados

Call Formularios(Me)
Call RefrescaTags(Me)



End Sub

Private Sub sbLimpia()

vPaso = True

txtUsuario.Text = ""
txtFiltroUsuarioAsg.Text = ""

lswUsuarios.ListItems.Clear

lswRoles.ListItems.Clear

vPaso = False

If gPortal.Empresa_Id <= 0 Then
   lswUsuarios.Enabled = False
   lswClienteUsuarios.Enabled = False
   lswRoles.Enabled = False
   
   MsgBox "Debe Consultar un Cliente para poder vincular usuarios y roles!", vbExclamation
   
Else
   lswUsuarios.Enabled = True
   lswClienteUsuarios.Enabled = True
   lswRoles.Enabled = True

End If

End Sub

Private Sub sbConsultaUsuario()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)


strSQL = "select Usuario,Nombre,UserID" _
       & " from US_Usuarios" _
       & " where Estado = 'A' and (Usuario like '%" & Trim(txtUsuario.Text) & "%' or Nombre like '%" & Trim(txtUsuario.Text) & "%')"


If Not gAdminAccess.Rol_AdminView Then
    strSQL = strSQL & " AND isnull(key_admin,0) = 0"
End If
    
If Not gAdminAccess.Rol_DirGlobal Then
    'Solo Usuarios que han formado parte de este cliente anteriormente, si por error fue desvinculado
    strSQL = strSQL & " AND usuario in(select usuario from PGX_CLIENTES_USERS_H" _
        & " Where cod_Empresa = " & gPortal.Empresa_Id & ")"
End If


vPaso = True

With lswUsuarios.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Usuario)
          itmX.SubItems(1) = rs!Nombre
          itmX.Tag = rs!UserID
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

lblCount(0).Caption = "Items: " & Format(lswUsuarios.ListItems.Count, "###,##0")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbConsultaUsuariosVinculados()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltroUsuarioAsg.Text = fxSysCleanTxtInject(txtFiltroUsuarioAsg.Text)

strSQL = "select U.Usuario,U.Nombre,U.UserID,A.registro_Fecha,A.Registro_Usuario" _
       & " from US_Usuarios U inner join PGX_Clientes_USERS A on U.Usuario = A.usuario and A.cod_Empresa = " & txtCliente.Tag _
       & " where (U.Usuario like '%" & Trim(txtFiltroUsuarioAsg.Text) & "%' or U.Nombre like '%" & Trim(txtFiltroUsuarioAsg.Text) & "%')"
       
Select Case chkContabiliza.Value
  Case xtpChecked
    strSQL = strSQL & " and U.Contabiliza = 1"
  Case xtpUnchecked
    strSQL = strSQL & " and U.Contabiliza = 0"
  Case xtpGrayed
End Select


If Not gAdminAccess.Rol_AdminView Then
    strSQL = strSQL & " AND isnull(U.key_admin,0) = 0"
End If
    
strSQL = strSQL & " order by U.Nombre"

vPaso = True

With lswClienteUsuarios.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Usuario)
          itmX.SubItems(1) = rs!Nombre
          itmX.Tag = rs!UserID
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

lblCount(1).Caption = "Items: " & Format(lswClienteUsuarios.ListItems.Count, "###,##0")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbConsultaRoles()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltroAccess.Text = fxSysCleanTxtInject(txtFiltroAccess.Text)

strSQL = "select R.COD_ROL,R.DESCRIPCION,  case when ISNULL( m.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end as 'Asignado'" _
       & "       ,M.REGISTRO_FECHA, M.REGISTRO_USUARIO" _
       & "  from US_ROLES R" _
       & "         left join US_ROL_MIEMBROS M on R.COD_ROL = M.COD_ROL and M.COD_EMPRESA = " & txtCliente.Tag _
       & "         and M.USUARIO = '" & lblUsuario.Tag & "'" _
       & "  where R.ACTIVO = 1 and isnull(R.COD_EMPRESA," & txtCliente.Tag & ") = " & txtCliente.Tag _
       & "    and R.descripcion like '%" & txtFiltroAccess.Text & "%'" _
       & " order by   case when ISNULL( m.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end desc  , R.Descripcion asc"

vPaso = True

With lswRoles.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!cod_rol)
          itmX.SubItems(1) = rs!Descripcion
          itmX.Checked = rs!Asignado
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbConsultaEstaciones()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltroAccess.Text = fxSysCleanTxtInject(txtFiltroAccess.Text)


strSQL = "select E.ESTACION,E.DESCRIPCION,  case when ISNULL( A.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end as 'Asignado'" _
       & "       ,A.REGISTRO_FECHA, A.REGISTRO_USUARIO" _
       & "  from PGX_CLIENTES_ESTACIONES E" _
       & "         left join PGX_CLIENTES_ESTACIONES_USERS A on E.ESTACION = A.ESTACION and E.COD_EMPRESA = " & txtCliente.Tag _
       & "         and A.USUARIO = '" & lblUsuario.Tag & "'" _
       & "  where E.ACTIVA = 1 and isnull(E.COD_EMPRESA," & txtCliente.Tag & ") = " & txtCliente.Tag _
       & "    and E.descripcion like '%" & txtFiltroAccess.Text & "%'" _
       & " order by   case when ISNULL( A.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end desc  , E.Descripcion asc"

vPaso = True

With lswEstaciones.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!ESTACION)
          itmX.SubItems(1) = rs!Descripcion
          itmX.Checked = rs!Asignado
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbConsultaHorarios()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltroAccess.Text = fxSysCleanTxtInject(txtFiltroAccess.Text)

strSQL = "select E.COD_HORARIO,E.DESCRIPCION,  case when ISNULL( A.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end as 'Asignado'" _
       & "       ,A.REGISTRO_FECHA, A.REGISTRO_USUARIO" _
       & "  from PGX_CLIENTES_HORARIOS E" _
       & "         left join PGX_CLIENTES_HORARIOS_USERS A on E.COD_HORARIO = A.COD_HORARIO and E.COD_EMPRESA = " & txtCliente.Tag _
       & "         and A.USUARIO = '" & lblUsuario.Tag & "'" _
       & "  where E.ACTIVO = 1 and isnull(E.COD_EMPRESA," & txtCliente.Tag & ") = " & txtCliente.Tag _
       & "    and E.descripcion like '%" & txtFiltroAccess.Text & "%'" _
       & " order by   case when ISNULL( A.REGISTRO_USUARIO  ,'') = '' then 0 else 1 end desc  , E.Descripcion asc"

vPaso = True

With lswHorarios.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!cod_Horario)
          itmX.SubItems(1) = rs!Descripcion
          itmX.Checked = rs!Asignado
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub lswClienteUsuarios_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswClienteUsuarios.SortKey = ColumnHeader.Index - 1
  If lswClienteUsuarios.SortOrder = 0 Then lswClienteUsuarios.SortOrder = 1 Else lswClienteUsuarios.SortOrder = 0
  lswClienteUsuarios.Sorted = True
End Sub

Private Sub lswClienteUsuarios_DblClick()
Dim i As Integer

If vPaso Or lswClienteUsuarios.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError
 i = MsgBox("Esta seguro que desea desvincular a este USUARIO con el CLIENTE?", vbYesNo)
 If i = vbNo Then Exit Sub

'Asigna Usuario a un Cliente
strSQL = "exec spPGX_Usuario_Cliente_Asigna " & txtCliente.Tag & ",'" & lswClienteUsuarios.SelectedItem & "','" & glogon.Usuario & "','E',''"
Call ConectionExecute(strSQL)

'Sincroniza Core
Call spCore_Usuario_Sincroniza(txtCliente.Tag, lswClienteUsuarios.SelectedItem, lswClienteUsuarios.SelectedItem.SubItems(1), "I")


If lblUsuario.Tag = lswClienteUsuarios.SelectedItem Then
    lblUsuario.Tag = ""
    lblUsuario.Caption = "Usuario.:"
    lswRoles.ListItems.Clear
End If

Call sbConsultaUsuariosVinculados


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswClienteUsuarios_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Or lswClienteUsuarios.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

lblUsuario.Caption = "Usuario ..: " & Item.SubItems(1)
lblUsuario.Tag = Item.Text


strSQL = "select isnull(Limita_Acceso_Estacion,0) as 'Estacion', isnull(Limita_Acceso_Horario,0) as 'Horario'" _
       & " from PGX_Clientes_Users where cod_empresa = " & txtCliente.Tag _
       & " and usuario = '" & lblUsuario.Tag & "'"
Call OpenRecordSet(rs, strSQL, 1)

vPaso = True

If rs.EOF And rs.BOF Then
   chkAcceso_Equipo.Value = vbUnchecked
   chkAcceso_Horario.Value = vbUnchecked
Else
   chkAcceso_Equipo.Value = rs!ESTACION
   chkAcceso_Horario.Value = rs!Horario
End If
rs.Close

vPaso = False

tcMain(0).Selected = True
Call tcMain_SelectedChanged(tcMain.Item(0))
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswEstaciones_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswEstaciones.SortKey = ColumnHeader.Index - 1
  If lswEstaciones.SortOrder = 0 Then lswEstaciones.SortOrder = 1 Else lswEstaciones.SortOrder = 0
  lswEstaciones.Sorted = True
End Sub

Private Sub lswEstaciones_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Or lswEstaciones.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spPGX_Usuario_Estacion_Asigna " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & Item.Text & "','" & glogon.Usuario & "',1"
    Call ConectionExecute(strSQL)
Else
    strSQL = "exec spPGX_Usuario_Estacion_Asigna " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & Item.Text & "','" & glogon.Usuario & "',0"
    Call ConectionExecute(strSQL)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswHorarios_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswHorarios.SortKey = ColumnHeader.Index - 1
  If lswHorarios.SortOrder = 0 Then lswHorarios.SortOrder = 1 Else lswHorarios.SortOrder = 0
  lswHorarios.Sorted = True
End Sub

Private Sub lswHorarios_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Or lswHorarios.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spPGX_Usuario_Horario_Asigna " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & Item.Text & "','" & glogon.Usuario & "',1"
    Call ConectionExecute(strSQL)
Else
    strSQL = "exec spPGX_Usuario_Horario_Asigna " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & Item.Text & "','" & glogon.Usuario & "',0"
    Call ConectionExecute(strSQL)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswRoles_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRoles.SortKey = ColumnHeader.Index - 1
  If lswRoles.SortOrder = 0 Then lswRoles.SortOrder = 1 Else lswRoles.SortOrder = 0
  lswRoles.Sorted = True
End Sub


Private Sub lswRoles_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Or lswRoles.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spPGX_Usuario_Rol_Asigna " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & Item.Text & "','" & glogon.Usuario & "','I'"
Else
    strSQL = "exec spPGX_Usuario_Rol_Asigna " & txtCliente.Tag & ",'" & lblUsuario.Tag & "','" & Item.Text & "','" & glogon.Usuario & "','E'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswUsuarios_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswUsuarios.SortKey = ColumnHeader.Index - 1
  If lswUsuarios.SortOrder = 0 Then lswUsuarios.SortOrder = 1 Else lswUsuarios.SortOrder = 0
  lswUsuarios.Sorted = True
End Sub




Private Sub TimerX_Timer()

TimerX.Interval = 0

If Not gAdminAccess.Rol_Permisos Then
    MsgBox "No cuenta con el Rol para conceder permisos con este cliente", vbExclamation
    Unload Me
End If

End Sub

Private Sub txtFiltroAccess_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Select Case tcMain.SelectedItem
        Case 0 'Roles
            Call sbConsultaRoles
        Case 1 'Estaciones
            Call sbConsultaEstaciones
        Case 2 'Horarios
            Call sbConsultaHorarios
    End Select
End If
End Sub

Private Sub txtFiltroUsuarioAsg_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbConsultaUsuariosVinculados
End If
End Sub


Private Sub txtUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbConsultaUsuario
End If
End Sub


Private Sub lswUsuarios_DblClick()


If vPaso Or lswUsuarios.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError
'Asigna Usuario a un Cliente
strSQL = "exec spPGX_Usuario_Cliente_Asigna " & txtCliente.Tag & ",'" & lswUsuarios.SelectedItem & "','" & glogon.Usuario & "','I',''"
Call ConectionExecute(strSQL)

'Sincroniza Core
Call spCore_Usuario_Sincroniza(txtCliente.Tag, lswUsuarios.SelectedItem, lswUsuarios.SelectedItem.SubItems(1), "A")

Call sbConsultaUsuariosVinculados

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
    Case 0 'Roles
        Call sbConsultaRoles
    Case 1 'Estaciones
        Call sbConsultaEstaciones
    Case 2 'Horarios
        Call sbConsultaHorarios
End Select
  
End Sub
