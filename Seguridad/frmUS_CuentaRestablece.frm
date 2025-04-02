VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmUS_CuentaRestablece 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Restablecer Contraseña"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Aceptar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmUS_CuentaRestablece.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkCambiaSesion 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   2280
      Width           =   3255
      _Version        =   1310723
      _ExtentX        =   5741
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cambiar al Iniciar Sesión"
      ForeColor       =   8388608
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
   Begin XtremeSuiteControls.FlatEdit txtNuevo 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1200
      Width           =   3255
      _Version        =   1310723
      _ExtentX        =   5741
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConfirm 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
      _Version        =   1310723
      _ExtentX        =   5741
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
      PasswordChar    =   "*"
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   6735
      _Version        =   1310723
      _ExtentX        =   11880
      _ExtentY        =   2355
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   4800
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cancelar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmUS_CuentaRestablece.frx":0727
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   6735
      _Version        =   1310723
      _ExtentX        =   11880
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2655
      _Version        =   1310723
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Notas:"
      ForeColor       =   8388608
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   2655
      _Version        =   1310723
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Confirmar"
      ForeColor       =   8388608
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2655
      _Version        =   1310723
      _ExtentX        =   4683
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Nueva Contraseña"
      ForeColor       =   8388608
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7215
      _Version        =   1310723
      _ExtentX        =   12726
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Restablecer la Contraseña de Acceso"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   510
      X2              =   7080
      Y1              =   4650
      Y2              =   4650
   End
End
Attribute VB_Name = "frmUS_CuentaRestablece"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAccion_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo error
   
   
Select Case Index
    Case 0 'Cambiar

        If fxVerifica Then
        
        'Actualiza Usuario y el Historial de Claves
        strSQL = "update us_usuarios set KEY_RENEW_SESION = " & chkCambiaSesion.Value _
               & " where UserID = " & gEntidad.UserID
        Call ConectionExecute(strSQL, 1)
         
        
          strSQL = "select isnull(max(id),0) + 1 as Ultimo from us_keyHistory" _
                 & " where idkeysec = " & gEntidad.UserID
          Call OpenRecordSet(rs, strSQL, 1)
          
            strSQL = "insert us_keyHistory(ID,IDKEYSEC,KEYING,KEYSEC) values(" & rs!Ultimo _
                   & "," & gEntidad.UserID & ",Getdate(),'" & SIFGlobal.fxStringCifrado(txtNuevo) & "')"
            Call ConectionExecute(strSQL, 1)
          rs.Close
          
          Call sbSEGCuentaLog("04", txtNotas.Text, glogon.Usuario, gEntidad.Usuario)
          
          If chkCambiaSesion.Value = vbChecked Then
             Call sbSEGCuentaLog("18", IIf((chkCambiaSesion.Value = vbChecked), "Activa", "Inactiva") & vbCrLf & vbCrLf & txtNotas.Text, glogon.Usuario, gEntidad.Usuario)
          End If
          
          
          MsgBox "La clave de acceso ha sido restablecida.", vbExclamation
          Unload Me
        End If

    Case 1 'Cancelar
        Unload Me
End Select


Exit Sub

Salir:
    Exit Sub
error:

End Sub



Private Function fxVerifica() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, vPaso As Boolean

On Error GoTo vError

vMensaje = ""

strSQL = "select * from us_parametros"
Call OpenRecordSet(rs, strSQL, 1)


If Len(txtNotas.Text) <= 10 Then
    vMensaje = vMensaje & " - La anotación no es válida... verifique.!!!" & vbCrLf
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

If chkCambiaSesion.Value = vbUnchecked Then

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
    
End If

'Busca la contraseña, en el historial para que no duplique
strSQL = "select top " & rs!KEY_HISTORY & " * from us_keyHistory" _
       & " where idkeysec = " & gEntidad.UserID & "  order by id desc"
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


Private Sub cbo_Click()

If cbo.ListCount = 0 Then Exit Sub

  txtNotas.Text = cbo.Text
End Sub

Private Sub Form_Load()


cbo.Clear
cbo.AddItem "Clave Inicial del Usuario"
cbo.AddItem "Usuario olvida contraseña"
cbo.AddItem "Usuario No logra Restablecer la Clave"
cbo.AddItem "Cambio por Solicitud del Cliente"
cbo.AddItem "Revisión de Soporte Técnico"


End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then btnAccion(0).SetFocus
End Sub


Private Sub txtNuevo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtConfirm.SetFocus
End Sub

