VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPosCajasDesBloqueo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Desbloqueo de Cajas"
   ClientHeight    =   3336
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   6084
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3336
   ScaleWidth      =   6084
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdDesBloquear 
      Height          =   492
      Left            =   3360
      TabIndex        =   0
      Top             =   2520
      Width           =   1932
      _Version        =   1245187
      _ExtentX        =   3408
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "&Des Bloquear Caja"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmPosCajasDesBloqueo.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   2280
      TabIndex        =   2
      Top             =   1800
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboUser 
      Height          =   312
      Left            =   2280
      TabIndex        =   4
      Top             =   1440
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Caja Asignada"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Desbloquear Cajas con Bloqueo Activado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   5772
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPosCajasDesBloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cmdDesBloquear_Click()
Dim strSQL As String

On Error GoTo vError


strSQL = "update pv_cajas set bloqueo = 0" _
      & " where usuario = '" & cboUser.ItemData(cboUser.ListIndex) _
      & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Desbloque de Caja: " & cbo.ItemData(cbo.ListIndex) & ".US." & cboUser.ItemData(cboUser.ListIndex))

MsgBox "Caja desBloqueada Satisfactoriamente...", vbInformation
Unload Me

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLlenaCboCaja(vUsuario As String)
Dim strSQL As String

strSQL = "select rtrim(cod_caja) as 'IdX',  rtrim(nombre) as 'ItmX'" _
       & " from pv_cajas where estado = 'A' and usuario = '" _
       & vUsuario & "' and bloqueo = 1 order by cod_caja"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

End Sub


Private Sub cboUser_Click()
If vPaso Then Exit Sub

Call sbLlenaCboCaja(cboUser.ItemData(cboUser.ListIndex))

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 33

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True

strSQL = "select rtrim(Usuario) as 'IdX',  rtrim(Usuario) + ': ' + nombre as 'Itmx'" _
       & " from pv_cajas where estado = 'A' and bloqueo = 1 order by cod_caja"
Call sbCbo_Llena_New(cboUser, strSQL, False, True)

vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



