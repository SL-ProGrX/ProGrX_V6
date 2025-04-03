VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Begin VB.Form frmPosCajaClave 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Clave de Caja"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdCambiar 
      Height          =   492
      Left            =   3360
      TabIndex        =   5
      Top             =   3840
      Width           =   1932
      _Version        =   1310720
      _ExtentX        =   3408
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cambiar Clave"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmPosCajaClave.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   2280
      TabIndex        =   8
      Top             =   1080
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5313
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2280
      TabIndex        =   9
      Top             =   1440
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5318
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
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   2280
      TabIndex        =   10
      Top             =   2520
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5313
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConfirma 
      Height          =   312
      Left            =   2280
      TabIndex        =   11
      Top             =   3240
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5313
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNueva 
      Height          =   312
      Left            =   2280
      TabIndex        =   7
      Top             =   2880
      Width           =   3012
      _Version        =   1310720
      _ExtentX        =   5313
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   5892
      _Version        =   1310720
      _ExtentX        =   10393
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Cambio de Clave"
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
      Alignment       =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Realice la Renovación de su contraseña de Cajero de POS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   5652
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   5400
      X2              =   360
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Asignada"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave del Sistema"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   480
      TabIndex        =   0
      Top             =   2520
      Width           =   1812
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPosCajaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbo_Click()
txtClave = ""
txtNueva = ""
txtConfirma = ""
End Sub

Private Sub cmdCambiar_Click()
Dim strSQL As String

On Error GoTo vError
      
If txtClave.Text <> glogon.Clave Then
   MsgBox "La clave de Inicio de Sesión al Sistema es incorrecta!", vbExclamation
   Exit Sub
End If

If txtNueva.Text <> txtConfirma.Text Then
  MsgBox "Clave Nueva y su Confirmación no concuerdan...", vbExclamation
  Exit Sub
End If
  
strSQL = "update pv_cajas set clave = '" & fxPosEncrypta(txtNueva.Text) & "'" _
       & " where usuario = '" & txtUsuario & "' and cod_caja = '" _
       & cbo.ItemData(cbo.ListIndex) & "'"
Call ConectionExecute(strSQL)
    
Call Bitacora("Aplica", "Cambio de Clave Caja:" & cbo.ItemData(cbo.ListIndex) & ".US." & txtUsuario)

MsgBox "Cambio de Clave realizado...", vbInformation
Unload Me

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 33

txtUsuario = glogon.Usuario

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtUsuario = glogon.Usuario
txtClave.Text = ""
txtNueva.Text = ""
txtConfirma.Text = ""

strSQL = "select rtrim(cod_caja) as 'IdX' , rtrim(nombre) as 'ItmX'" _
       & " from pv_cajas where estado = 'A' and usuario = '" _
       & glogon.Usuario & "' order by cod_caja"

Call sbCbo_Llena_New(cbo, strSQL, False, True)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNueva.SetFocus
End Sub

Private Sub txtConfirma_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdCambiar.SetFocus
End Sub

Private Sub txtNueva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConfirma.SetFocus
End Sub
