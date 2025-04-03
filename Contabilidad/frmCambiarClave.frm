VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCambiaClave 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar Contraseña"
   ClientHeight    =   4044
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   10632
   HelpContextID   =   9014
   Icon            =   "frmCambiarClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCambiarClave.frx":030A
   ScaleHeight     =   4044
   ScaleWidth      =   10632
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnCambiar 
      Height          =   492
      Left            =   8040
      TabIndex        =   7
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   3120
      Width           =   2292
      _Version        =   1245185
      _ExtentX        =   4043
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cambiar Contraseña"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmCambiarClave.frx":4B55
      ImageAlignment  =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambiarClave.frx":54C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtViejo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7980
      PasswordChar    =   "*"
      TabIndex        =   0
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   1500
      Width           =   2415
   End
   Begin VB.TextBox txtNuevo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7980
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   1920
      Width           =   2415
   End
   Begin VB.TextBox txtConfirm 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   7980
      PasswordChar    =   "*"
      TabIndex        =   2
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   2340
      Width           =   2415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      Tag             =   "SIFGlobal.fxStringCifrado"
      X1              =   5400
      X2              =   10680
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirmar nueva contraseña"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   5850
      TabIndex        =   5
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   2340
      Width           =   2850
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nueva contraseña"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   5850
      TabIndex        =   4
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   1920
      Width           =   2145
   End
   Begin VB.Label lblPassword 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Contraseña anterior"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      Tag             =   "SIFGlobal.fxStringCifrado"
      Top             =   1500
      Width           =   2250
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cambio de clave de acceso "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   10785
   End
End
Attribute VB_Name = "frmCambiaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function fxVerifica() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, vPaso As Boolean

On Error GoTo vError

vMensaje = ""

strSQL = "select * from us_parametros"
Call OpenRecordSet(rs, strSQL, 1)

If txtViejo.Text <> glogon.Clave Then
 vMensaje = vMensaje & " - La contraseña anterior digitada no corresponde a la contraseña actual" & vbCrLf
End If

If txtNuevo.Text <> txtConfirm.Text Then
    vMensaje = vMensaje & " - La confirmación de la contraseña no corresponde con la nueva contraseña." & vbCrLf
End If

'Largo Minimo de la Contraseña
If Len(txtNuevo.Text) < rs!KEY_LENMIN Then
    vMensaje = vMensaje & " - El largo de la contraseña no es válido, el mínimo es de " & rs!KEY_LENMIN _
             & " caracteres" & vbCrLf
End If

'Largo Maximo de la Contraseña
If Len(txtNuevo.Text) > rs!KEY_LENMAX Then
    vMensaje = vMensaje & " - El largo de la contraseña no es válido, el máximo es de " & rs!KEY_LENMAX _
             & " caracteres" & vbCrLf
End If

'Valida que tenga los numeros
If fxClaveContarNumeros(txtNuevo.Text) < rs!key_NumChar Then
    vMensaje = vMensaje & " - La clave no contiene la cantidad requerida de números, cantidad requerida: " & rs!key_NumChar _
             & " números" & vbCrLf
End If

'Valida que tenga caracteres especiales
If fxClaveCaracteresEspeciales(txtNuevo.Text) < rs!key_SimChar Then
    vMensaje = vMensaje & " - La clave no contiene la cantidad requerida de caracteres especiales $#@^&*(), cantidad requerida: " & rs!key_SimChar _
             & " caracteres" & vbCrLf
End If

'Valida que tenga caracteres especiales
If fxClaveCaracteresMayusculas(txtNuevo.Text) < rs!key_CapChar Then
    vMensaje = vMensaje & " - La clave no contiene la cantidad requerida de mayúsculas, cantidad requerida: " & rs!key_CapChar _
             & " mayúsculas" & vbCrLf
End If

'Busca la contraseña, en el historial para que no duplique
strSQL = "select top " & rs!KEY_HISTORY & " * from us_keyHistory" _
       & " where idkeysec in(select userID from US_usuarios where Usuario = '" _
       & glogon.Usuario & "') order by id desc"
rs.Close
Call OpenRecordSet(rs, strSQL, 1)
vPaso = True
Do While Not rs.EOF
 If Trim(SIFGlobal.fxStringCifrado(txtNuevo)) = Trim(rs!KEYSEC) Then
   vPaso = False
   Exit Do
 End If
 rs.MoveNext
Loop
rs.Close

If Not vPaso Then
    vMensaje = vMensaje & " - La contraseña nueva ya ha sido utilizada con anterioridad, por favor ingrese una nueva" & vbCrLf
End If

If Len(vMensaje) = 0 Then
   fxVerifica = True
Else
   fxVerifica = False
   MsgBox vMensaje, vbExclamation
End If

Exit Function

vError:
 fxVerifica = False
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Function

Private Sub sbCambiar()
Dim strSQL As String

On Error GoTo error
   
If fxVerifica Then
  glogon.Clave = txtNuevo.Text
  
  strSQL = "exec spSEG_Password " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "','" & SIFGlobal.fxStringCifrado(txtViejo.Text) _
         & "','" & SIFGlobal.fxStringCifrado(txtNuevo.Text) & "',0"
  Call ConectionExecute(strSQL, 1)

  Call sbSEGCuentaLog("12")
  
  MsgBox "La clave de acceso ha sido cambiada.", vbExclamation
  Unload Me
End If

Exit Sub

Salir:
    Exit Sub
error:
End Sub



Private Sub btnCambiar_Click()
Call sbCambiar
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then Call sbCambiar
End Sub

Private Sub txtNuevo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtConfirm.SetFocus
End Sub


Private Sub txtViejo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtNuevo.SetFocus
End Sub
