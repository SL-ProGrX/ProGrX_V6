VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmUS_Cuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de Cuenta"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   240
      Top             =   0
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   1
      Top             =   5520
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
      Picture         =   "frmUS_Cuentas.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   3960
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
      Left            =   5280
      TabIndex        =   3
      Top             =   5520
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
      Picture         =   "frmUS_Cuentas.frx":0727
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkBloqueo 
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   5055
      _Version        =   1310723
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta se encuentra bloqueada temporalmente ?"
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
   Begin XtremeSuiteControls.CheckBox chkBloqueoI 
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   5055
      _Version        =   1310723
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta con Bloqueo Indefinido"
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
   Begin XtremeSuiteControls.CheckBox chkCuentaCaduca 
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2520
      Width           =   5055
      _Version        =   1310723
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta Nunca Caduca ?"
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
   Begin XtremeSuiteControls.CheckBox chkAdmin 
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2880
      Width           =   5055
      _Version        =   1310723
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cuenta de Administración"
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
   Begin XtremeSuiteControls.CheckBox chkCambioContrasena 
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   5055
      _Version        =   1310723
      _ExtentX        =   8916
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cambia Contraseña en el próximo Logon"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   3720
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
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6975
      _Version        =   1310723
      _ExtentX        =   12303
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Revision de Cuenta"
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
End
Attribute VB_Name = "frmUS_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mBloqueo As Integer, mBloqueoT As Integer, mCtaAdmin As Integer, mCaduca As Integer, mRenueva As Integer



Private Sub cmdCancelar_Click()
Unload Me
End Sub


Private Sub btnAccion_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError



Select Case Index

    Case 0 'Aplicar

        If Len(txtNotas.Text) < 10 Then
           MsgBox "Especifique una Anotación válida para registrar el movimiento...!", vbExclamation
           Exit Sub
        End If
        
        strSQL = "update US_usuarios set key_bloqueo = " & IIf((chkBloqueo.Value = vbChecked), "Getdate()", "Null") _
               & ",key_bloqueoI = " & chkBloqueoI.Value & ",key_admin = " & chkAdmin.Value _
               & ",key_caducidad = " & chkCuentaCaduca.Value & ",KEY_RENEW_SESION = " & chkCambioContrasena.Value _
               & " where Usuario = '" & scTitulo.Caption & "'"
        Call ConectionExecute(strSQL)
              
        txtNotas.Text = " ...: " & txtNotas.Text
              
        If mCaduca <> chkCuentaCaduca.Value Then
         Call sbSEGCuentaLog("17", IIf((chkCuentaCaduca.Value = vbChecked), "Activa", "Inactiva") & vbCrLf & vbCrLf & txtNotas.Text, glogon.Usuario, scTitulo.Caption)
        End If
        
        If mCtaAdmin <> chkAdmin.Value Then
         Call sbSEGCuentaLog("05", IIf((chkAdmin.Value = vbChecked), "Activa", "Inactiva") & vbCrLf & vbCrLf & txtNotas.Text, glogon.Usuario, scTitulo.Caption)
        End If
        
        If mRenueva <> chkCambioContrasena.Value Then
         Call sbSEGCuentaLog("18", IIf((chkCambioContrasena.Value = vbChecked), "Activa", "Inactiva") & vbCrLf & vbCrLf & txtNotas.Text, glogon.Usuario, scTitulo.Caption)
        End If
        
        If mBloqueo <> chkBloqueoI.Value Then
         Call sbSEGCuentaLog("06", IIf((chkBloqueoI.Value = vbChecked), "Activa", "Inactiva") & vbCrLf & vbCrLf & txtNotas.Text, glogon.Usuario, scTitulo.Caption)
        End If
        
        If mBloqueoT <> chkBloqueo.Value Then
          txtNotas.Text = IIf((chkBloqueo.Value = vbChecked), "Activa", "Inactiva")
         Call sbSEGCuentaLog("07", txtNotas.Text, glogon.Usuario, scTitulo.Caption)
        End If
      
    Case 1 'Cancelar
       Unload Me

End Select


      
MsgBox "Cambios Realizados Satisfactoriamente...", vbInformation
      
Unload Me

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub Form_Load()
    scTitulo.Caption = gEntidad.Usuario
    scTitulo.Tag = gEntidad.UserID
End Sub


Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

TimerX.Interval = 0

strSQL = "exec spSEG_Bloqueo '" & scTitulo.Caption & "'"
Call OpenRecordSet(rs, strSQL)
    If rs!bloqueo = 1 Then chkBloqueo.Value = vbChecked
    If rs!BloqueoI = 1 Then chkBloqueoI.Value = vbChecked
    If IsNull(rs!bloqueoT) Then chkBloqueo.Value = vbUnchecked
rs.Close

strSQL = "select key_caducidad,key_admin,KEY_RENEW_SESION" _
       & " from US_usuarios where Usuario = '" & scTitulo.Caption & "'"
Call OpenRecordSet(rs, strSQL)
    chkCambioContrasena.Value = IIf(IsNull(rs!KEY_RENEW_SESION), 0, rs!KEY_RENEW_SESION)
    chkCuentaCaduca.Value = IIf(IsNull(rs!key_caducidad), 0, rs!key_caducidad)
    chkAdmin.Value = IIf(IsNull(rs!key_admin), 0, rs!key_admin)
rs.Close


mCaduca = chkCuentaCaduca.Value
mCtaAdmin = chkAdmin.Value
mRenueva = chkCambioContrasena.Value
mBloqueo = chkBloqueoI.Value
mBloqueoT = chkBloqueo.Value


'Desactiva todas las Opciones, Estas se activan hasta que se especifique una anotación válida
Call txtNotas_KeyUp(0, 0)

End Sub


Private Sub txtNotas_KeyUp(KeyCode As Integer, Shift As Integer)

If Len(txtNotas.Text) <= 10 Then
    chkCuentaCaduca.Enabled = False
    chkAdmin.Enabled = False
    chkCambioContrasena.Enabled = False
    chkBloqueoI.Enabled = False
    chkBloqueo.Enabled = False
Else
    chkCuentaCaduca.Enabled = True
    chkAdmin.Enabled = True
    chkCambioContrasena.Enabled = True
    chkBloqueoI.Enabled = True
    chkBloqueo.Enabled = True
End If

End Sub
