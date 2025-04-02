VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmUS_Admin_Rol 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Usuarios Administradores de la Seguridad"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   15900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lswClientes 
      Height          =   4575
      Left            =   8640
      TabIndex        =   6
      Top             =   1680
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
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
   Begin XtremeSuiteControls.ListView lswAdmin 
      Height          =   4575
      Left            =   4320
      TabIndex        =   3
      Top             =   1680
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswUsuarios 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   13573
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
   Begin XtremeSuiteControls.PushButton btnAdminApl 
      Height          =   375
      Index           =   0
      Left            =   5760
      TabIndex        =   18
      Top             =   8880
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplicar al Usuario"
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
      Picture         =   "frmUS_Admin_Rol.frx":0000
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Global_Search 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   11
      Top             =   6720
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Consulta Directorio Global"
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
   End
   Begin XtremeSuiteControls.CheckBox chkClientes 
      Height          =   210
      Left            =   8880
      TabIndex        =   9
      Top             =   1035
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   1320
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
   Begin XtremeSuiteControls.FlatEdit txtAdministrador 
      Height          =   330
      Left            =   4320
      TabIndex        =   4
      Top             =   1320
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
   Begin XtremeSuiteControls.FlatEdit txtCliente 
      Height          =   330
      Left            =   8640
      TabIndex        =   7
      Top             =   1320
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
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
   Begin XtremeSuiteControls.CheckBox chk_R_Local_User 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   12
      Top             =   7080
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Administra Usuarios Locales"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Local_Reset 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   13
      Top             =   7800
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Reset Claves U. Locales"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Admin_Review 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   14
      Top             =   8160
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Puede ver otros Administradores"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Propagar 
      Height          =   255
      Left            =   5160
      TabIndex        =   15
      Top             =   8520
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Propagar Roles a todos los clientes"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Global_Search 
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   19
      Top             =   6720
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Consulta Directorio Global"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Local_User 
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   20
      Top             =   7080
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Administra Usuarios Locales"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Local_Reset 
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   21
      Top             =   7800
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Reset Claves U. Locales"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Admin_Review 
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   22
      Top             =   8160
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Puede ver otros Administradores"
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
   End
   Begin XtremeSuiteControls.PushButton btnAdminApl 
      Height          =   375
      Index           =   1
      Left            =   10920
      TabIndex        =   23
      Top             =   8880
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplicar en el Cliente"
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
      Picture         =   "frmUS_Admin_Rol.frx":0727
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Local_Grants 
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   24
      Top             =   7440
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Asigna Permisos por Roles"
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
   End
   Begin XtremeSuiteControls.CheckBox chk_R_Local_Grants 
      Height          =   255
      Index           =   1
      Left            =   9600
      TabIndex        =   25
      Top             =   7440
      Width           =   2655
      _Version        =   1441793
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Asigna Permisos por Roles"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scCliente 
      Height          =   330
      Left            =   8640
      TabIndex        =   17
      Top             =   6240
      Width           =   7215
      _Version        =   1441793
      _ExtentX        =   12726
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scAdmin 
      Height          =   330
      Left            =   4320
      TabIndex        =   16
      Top             =   6240
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   240
      Width           =   6375
      _Version        =   1441793
      _ExtentX        =   11245
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Administradores de la Seguridad"
      ForeColor       =   16777215
      BackColor       =   -2147483633
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
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Directorio de Usuarios de la Plataforma"
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
      Left            =   4320
      TabIndex        =   5
      Top             =   960
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Usuarios Administradores"
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
      Index           =   2
      Left            =   8640
      TabIndex        =   8
      Top             =   960
      Width           =   7335
      _Version        =   1441793
      _ExtentX        =   12938
      _ExtentY        =   582
      _StockProps     =   14
      Caption         =   "Clientes Asignados"
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
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15975
   End
End
Attribute VB_Name = "frmUS_Admin_Rol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vPaso As Boolean


Private Sub btnAdminApl_Click(Index As Integer)

On Error GoTo vError

Select Case Index
    Case 0 'Roles General del Admin
        If scAdmin.Tag = "" Then Exit Sub
        
        strSQL = "exec spSEG_Admin_Local_Add '" & scAdmin.Tag & "', 'A', '" & glogon.Usuario _
                & "', " & chk_R_Local_Grants(0).Value & ", " & chk_R_Local_User(0).Value _
                & ", " & chk_R_Local_Reset(0).Value & ", " & chk_R_Global_Search(0).Value _
                & ", " & chk_R_Admin_Review(0).Value & ", " & chk_R_Propagar.Value
        Call ConectionExecute(strSQL)
    
    
    Case 1 'Roles General del Admin por Clientes
        If scCliente.Tag = "" Or scCliente.Tag = "0" Then Exit Sub
        
            strSQL = "exec spSEG_Admin_Clients_Roles_Add '" & scAdmin.Tag & "', " & scCliente.Tag & " , 'A', '" & glogon.Usuario _
                    & "', " & chk_R_Local_Grants(1).Value & ", " & chk_R_Local_User(1).Value _
                    & ", " & chk_R_Local_Reset(1).Value & ", " & chk_R_Global_Search(1).Value _
                    & ", " & chk_R_Admin_Review(1).Value
            
            Call ConectionExecute(strSQL)

End Select

MsgBox "Roles aplicados satisfactoriamente!", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_activate()
vModulo = 13
End Sub

Private Sub Form_Load()
vModulo = 13

txtCliente.Tag = gPortal.Empresa_Id
txtCliente.Text = gPortal.Empresa_Name

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lswAdmin.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1800
    .Add , , "Nombre", lswAdmin.Width - 2000
End With

With lswUsuarios.ColumnHeaders
    .Clear
    .Add , , "Usuario", 1800
    .Add , , "Nombre", lswUsuarios.Width - 2000
End With

With lswClientes.ColumnHeaders
    .Clear
    .Add , , "Corto", 2100
    .Add , , "Nombre", lswClientes.Width - 2100
End With

lswClientes.Checkboxes = True


Call sbLimpia
Call sbAdministradores_Vinculados

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()

vPaso = True

txtUsuario.Text = ""
txtAdministrador.Text = ""
txtCliente.Text = ""

lswUsuarios.ListItems.Clear
lswClientes.ListItems.Clear

vPaso = False

End Sub

'Private Sub sbConsultaUsuario()
'
'On Error GoTo vError
'
'Me.MousePointer = vbHourglass
'
'txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)
'
'
'strSQL = "select Usuario,Nombre,UserID" _
'       & " from US_Usuarios" _
'       & " where Estado = 'A' and (Usuario like '%" & Trim(txtUsuario.Text) & "%' or Nombre like '%" & Trim(txtUsuario.Text) & "%')"
'
'vPaso = True
'
'With lswUsuarios.ListItems
'    .Clear
'    Call OpenRecordSet(rs, strSQL)
'    Do While Not rs.EOF
'      Set itmX = .Add(, , rs!Usuario)
'          itmX.SubItems(1) = rs!Nombre
'          itmX.Tag = rs!UserID
'      rs.MoveNext
'    Loop
'    rs.Close
'End With
'
'vPaso = False
'
''lblCount(0).Caption = "Items: " & Format(lswUsuarios.ListItems.Count, "###,##0")
'
'Me.MousePointer = vbDefault
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'
'End Sub


Private Sub sbAdministradores_Vinculados()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtAdministrador.Text = fxSysCleanTxtInject(txtAdministrador.Text)


strSQL = "exec spSEG_Admin_Local_List '" & txtAdministrador.Text & "'"

vPaso = True

With lswAdmin.ListItems
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

'lblCount(1).Caption = "Items: " & Format(lswAdmin.ListItems.Count, "###,##0")

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub sbUsuario_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtUsuario.Text = fxSysCleanTxtInject(txtUsuario.Text)

strSQL = "exec spSEG_Usuarios_List '" & txtUsuario.Text & "'"

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


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub sbClientes_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtCliente.Text = fxSysCleanTxtInject(txtCliente.Text)

strSQL = "exec spSEG_Admin_Clients_Load '" & scAdmin.Tag & "', '" & txtCliente.Text & "'"

vPaso = True

With lswClientes.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Nombre_Corto)
          itmX.SubItems(1) = rs!Nombre_Largo
          itmX.Tag = rs!cod_Empresa
          
          If rs!Asignado = 1 Then
            itmX.Checked = True
          End If
          
      rs.MoveNext
    Loop
    rs.Close
End With

chk_R_Global_Search(1).Value = xtpUnchecked
chk_R_Local_User(1).Value = xtpUnchecked
chk_R_Local_Grants(1).Value = xtpUnchecked
chk_R_Local_Reset(1).Value = xtpUnchecked
chk_R_Admin_Review(1).Value = xtpUnchecked
chk_R_Local_Grants(1).Value = xtpUnchecked


vPaso = False


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub sbClientes_Roles_Load(pCliente As Long)
On Error GoTo vError

'Consulta Permisos del Admin: General
strSQL = "exec  spSEG_Admin_Clients_Roles_Load '" & scAdmin.Tag & "', " & pCliente
Call OpenRecordSet(rs, strSQL)

    
    scCliente.Tag = pCliente
    scCliente.Caption = "-Asignando-"

chk_R_Global_Search(1).Value = xtpUnchecked
chk_R_Local_User(1).Value = xtpUnchecked
chk_R_Local_Grants(1).Value = xtpUnchecked
chk_R_Local_Reset(1).Value = xtpUnchecked
chk_R_Admin_Review(1).Value = xtpUnchecked


If Not rs.EOF And Not rs.BOF Then
    
    scCliente.Tag = pCliente
    scCliente.Caption = rs!Nombre
    
    chk_R_Global_Search(1).Value = rs!R_Global_Dir_Search
    chk_R_Local_User(1).Value = rs!R_Local_Users
    chk_R_Local_Grants(1).Value = rs!R_Local_Grants
    chk_R_Local_Reset(1).Value = rs!R_Local_Key_Reset
    chk_R_Admin_Review(1).Value = rs!R_Admin_Review
End If



Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAdmin_Load(pAdmin As String)
On Error GoTo vError

'Consulta Permisos del Admin: General
strSQL = "exec  spSEG_Admin_Local_Load '" & pAdmin & "'"
Call OpenRecordSet(rs, strSQL)


scAdmin.Tag = pAdmin
scAdmin.Caption = rs!Nombre

chk_R_Global_Search(0).Value = xtpUnchecked
chk_R_Local_User(0).Value = xtpUnchecked
chk_R_Local_Grants(0).Value = xtpUnchecked
chk_R_Local_Reset(0).Value = xtpUnchecked
chk_R_Admin_Review(0).Value = xtpUnchecked

chk_R_Propagar.Value = xtpUnchecked

If Not rs.EOF And Not rs.BOF Then
    chk_R_Global_Search(0).Value = rs!R_Global_Dir_Search
    chk_R_Local_User(0).Value = rs!R_Local_Users
    chk_R_Local_Grants(0).Value = rs!R_Local_Grants
    chk_R_Local_Reset(0).Value = rs!R_Local_Key_Reset
    chk_R_Admin_Review(0).Value = rs!R_Admin_Review
End If


'Consulta a Clientes vinculados con el Administrador
Call sbClientes_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswAdmin_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

Call sbAdmin_Load(Item.Text)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswClientes_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Or lswClientes.ListItems.Count = 0 Then Exit Sub

If scCliente.Tag = "" Or scCliente.Tag = "0" Then Exit Sub

If Item.Checked Then
    strSQL = "exec spSEG_Admin_Clients_Roles_Add '" & scAdmin.Tag & "', " & Item.Tag & " , 'A', '" & glogon.Usuario _
            & "', 0, 0, 0, 0, 0"
Else
    strSQL = "exec spSEG_Admin_Clients_Roles_Add '" & scAdmin.Tag & "', " & Item.Tag & " , 'E', '" & glogon.Usuario _
            & "', 0, 0, 0, 0, 0"
End If

Call ConectionExecute(strSQL)


Call sbClientes_Roles_Load(Item.Tag)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswClientes_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

If vPaso Or lswClientes.ListItems.Count = 0 Then Exit Sub

Call sbClientes_Roles_Load(Item.Tag)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswUsuarios_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswUsuarios.SortKey = ColumnHeader.Index - 1
  If lswUsuarios.SortOrder = 0 Then lswUsuarios.SortOrder = 1 Else lswUsuarios.SortOrder = 0
  lswUsuarios.Sorted = True
End Sub


Private Sub lswUsuarios_DblClick()

If vPaso Or lswUsuarios.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

'Asigna Usuario a un Cliente
strSQL = "exec  spSEG_Admin_Local_Add '" & lswUsuarios.SelectedItem & "', 'A', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbAdministradores_Vinculados

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtAdministrador_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbAdministradores_Vinculados
End If
End Sub


Private Sub txtCliente_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbClientes_Load
    
End If

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbUsuario_Consulta
End If
End Sub
