VERSION 5.00
Begin VB.Form frmPosCambioClaveCaja 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Clave de Caja"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "&Cambiar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3960
      Picture         =   "frmPosCambioClaveCaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtConfirma 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtNueva 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtClave 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   5160
      X2              =   120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Confirmación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Nueva"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   5040
      X2              =   240
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Caja Asignada"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Clave Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "frmPosCambioClaveCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer


Private Sub cbo_Click()
txtClave = ""
txtNueva = ""
txtConfirma = ""
End Sub

Private Sub cmdCambiar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from pv_cajas where usuario = '" _
       & txtUsuario & "' and cod_caja = '" & fxCodigoCbo(cbo) & "' and clave = '" _
       & fxPosEncrypta(txtClave) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  MsgBox "No se encontró caja: verifique Clave y Cajas para su usuario...", vbExclamation
Else
  If txtNueva <> txtConfirma Then
    MsgBox "Clave Nueva y su Confirmación no concuerdan...", vbExclamation
  Else
    strSQL = "update pv_cajas set clave = '" & fxPosEncrypta(txtNueva) & "'" _
           & " where usuario = '" & txtUsuario & "' and cod_caja = '" & fxCodigoCbo(cbo) _
           & "' and clave = '" & fxPosEncrypta(txtClave) & "'"
    Call ConectionExecute(strSQL)
        
    Call Bitacora("Aplica", "Cambio de Clave Caja:" & fxCodigoCbo(cbo) & ".US." & txtUsuario)
    
    MsgBox "Cambio de Clave realizado...", vbInformation
    Unload Me
  End If
End If
rs.Close
End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 33

i = 0

txtUsuario = glogon.Usuario

strSQL = "select (rtrim(cod_caja) + ' - ' + rtrim(nombre)) as Caja" _
       & " from pv_cajas where estado = 'A' and usuario = '" _
       & glogon.Usuario & "' order by cod_caja"
Call OpenRecordSet(rs, strSQL)

cbo.Clear
Do While Not rs.EOF
 cbo.AddItem rs!Caja
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
  rs.MoveFirst
  cbo.Text = rs!Caja
End If
rs.Close

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
