VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmTES_AutorizaChKey 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Contraseña de Autorizadores"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   528
      Left            =   4440
      TabIndex        =   4
      Top             =   2280
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   931
      _StockProps     =   79
      Caption         =   "Cambiar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      Appearance      =   16
      Picture         =   "frmTES_AutorizaChKey.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1680
      TabIndex        =   5
      Top             =   240
      Width           =   4212
      _Version        =   1245185
      _ExtentX        =   7429
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
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtActual 
      Height          =   312
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   4212
      _Version        =   1245185
      _ExtentX        =   7429
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
      PasswordChar    =   "*"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNueva 
      Height          =   312
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   4212
      _Version        =   1245185
      _ExtentX        =   7429
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
      PasswordChar    =   "*"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtConfirma 
      Height          =   312
      Left            =   1680
      TabIndex        =   8
      Top             =   1680
      Width           =   4212
      _Version        =   1245185
      _ExtentX        =   7429
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
      PasswordChar    =   "*"
      Appearance      =   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirma Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Clave"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave Actual"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00404040&
      Height          =   312
      Index           =   4
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1692
   End
End
Attribute VB_Name = "frmTES_AutorizaChKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset


On Error GoTo vError

'Verifica que Exista la clave actual

strSQL = "select isnull(count(*),0) as Existe from tes_autorizaciones" _
       & " where nombre = '" & txtUsuario & "' and clave = '" & fxTESCifrado(txtActual) & "'"

Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  rs.Close
  MsgBox "La Clave Actual no es válida o no fue localizada...", vbExclamation
  Exit Sub
End If
rs.Close

'Verifica que la Clave nueva y la confirmacion sean iguales
If txtNueva.Text <> txtConfirma.Text Then
  MsgBox "La Clave nueva no corresponde a la confirmada...", vbExclamation
  Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "update tes_autorizaciones set clave = '" & fxTESCifrado(txtNueva) _
       & "' where nombre = '" & txtUsuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Contraseña establecida...", vbInformation

Unload Me
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
vModulo = 9


Call Formularios(Me)
Call RefrescaTags(Me)

txtUsuario.Text = glogon.Usuario

End Sub

Private Sub txtActual_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtNueva.SetFocus
End Sub

Private Sub txtConfirma_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then cmdAplicar.SetFocus
vError:
End Sub

Private Sub txtNueva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtConfirma.SetFocus
End Sub
