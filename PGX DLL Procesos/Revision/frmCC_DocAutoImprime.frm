VERSION 5.00
Begin VB.Form frmCC_DocAutoImprime 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorización de ReImpresion de Documentos"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtClave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox txtUsuario 
      Height          =   315
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   $"frmCC_DocAutoImprime.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Contraseña"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "frmCC_DocAutoImprime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   'Verifica Usuario
  strSQL = "exec spSEGLogon '" & glogon.Usuario & "','" & fxgSegCifrado(glogon.Clave) & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs!existe = 0 Then
        MsgBox "No se puedo realizar el Login, verifique el Usuario y Contraseña...", vbExclamation
        Exit Sub
    End If
  rs.Close
 
    strSQL = "select coalesce(count(*),0) as Existe from ase_usr_Autoriza where usuario = '" _
           & txtUsuario & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs!existe = 0 Then
      MsgBox "Este Usuario no está autorizado a reImprimir documentos...", vbExclamation
    Else
       Call sbImprimeRecibo(frmCC_Documentos.txtDocumento, fxTipoASEDoc(frmCC_Documentos.cboTipo.Text), True)
    End If
    rs.Close
    
    Unload Me
End If

Exit Sub

vError:
  MsgBox "No se puedo realizar el Login, verifique el Usuario y Contraseña...", vbCritical

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtClave.SetFocus
End Sub
