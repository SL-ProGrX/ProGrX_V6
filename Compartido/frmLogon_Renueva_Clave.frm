VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmLogon_Renueva_Clave 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Renueva Contraseña:"
   ClientHeight    =   3285
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnCambiar 
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   2760
      Width           =   1575
      _Version        =   1245185
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cambiar"
      ForeColor       =   16711680
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
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmLogon_Renueva_Clave.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.FlatEdit txtNuevo 
      Height          =   315
      Left            =   5640
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
      _Version        =   1245185
      _ExtentX        =   4683
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConfirm 
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
      _Version        =   1245185
      _ExtentX        =   4683
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirma Contraseña"
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
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   4
      Top             =   1635
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Contraseña"
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
      Height          =   255
      Index           =   0
      Left            =   5640
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.Image imgX 
      Height          =   3300
      Left            =   0
      Picture         =   "frmLogon_Renueva_Clave.frx":096B
      Stretch         =   -1  'True
      Tag             =   "1200"
      Top             =   0
      Width           =   9555
   End
End
Attribute VB_Name = "frmLogon_Renueva_Clave"
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
  
  strSQL = "exec spSEG_Password " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "','" & SIFGlobal.fxStringCifrado(txtConfirm.Text) _
         & "','" & SIFGlobal.fxStringCifrado(txtNuevo.Text) & "',0"
  Call ConectionExecute(strSQL, 1)

  Call sbSEGCuentaLog("12")
  
  MsgBox "La contraseña de acceso ha sido renovada, puede iniciar sesión con su nueva contraseña.", vbExclamation
  End 'Sale de la aplicación
  
End If

Exit Sub

Salir:
    Exit Sub
error:
End Sub


Private Sub btnCambiar_Click()
Call sbCambiar
End Sub


Private Sub Form_Load()

Me.BackColor = RGB(70, 111, 178)
Me.ScaleMode = 3

End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then Call sbCambiar
End Sub

Private Sub txtNuevo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtConfirm.SetFocus
End Sub



