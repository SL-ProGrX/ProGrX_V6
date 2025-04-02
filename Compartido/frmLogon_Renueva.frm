VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmLogon_Renueva 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acceso: Solicita Renovación de Cuenta"
   ClientHeight    =   3435
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   10800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton RbtToken 
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   210
      _Version        =   1310720
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   4
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton RbtToken 
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   210
      _Version        =   1310720
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   4
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit_Token_Usuario 
      Height          =   315
      Left            =   7200
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
      _Version        =   1310720
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit_Token 
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
      _Version        =   1310720
      _ExtentX        =   4043
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnToken_Activa 
      Height          =   495
      Left            =   7200
      TabIndex        =   5
      Top             =   2400
      Width           =   2295
      _Version        =   1310720
      _ExtentX        =   4043
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Renovar Contraseña"
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
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmLogon_Renueva.frx":0000
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit_Usuario 
      Height          =   315
      Left            =   5040
      TabIndex        =   11
      Top             =   1560
      Width           =   4455
      _Version        =   1310720
      _ExtentX        =   7853
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit FlatEdit_Email 
      Height          =   315
      Left            =   5040
      TabIndex        =   12
      Top             =   1920
      Width           =   4455
      _Version        =   1310720
      _ExtentX        =   7853
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnToken_Envia 
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
      _Version        =   1310720
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Enviar TOKEN"
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
      TextAlignment   =   1
      Appearance      =   6
      Picture         =   "frmLogon_Renueva.frx":096B
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.Label lblOpcionTitle 
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
      _Version        =   1310720
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Ya Cuento con el TOKEN y quiero renovar mi contraseña"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblOpcionTitle 
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   16
      Top             =   1200
      Width           =   1575
      _Version        =   1310720
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Solicitar el TOKEN para Renovar mi contraseña"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   15
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   14
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label lblToken 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Token:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   1920
      Width           =   2055
   End
   Begin XtremeSuiteControls.Label lblOpcion 
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   1200
      Width           =   6975
      _Version        =   1310720
      _ExtentX        =   12303
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Opcion?"
      ForeColor       =   16777215
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblIngreso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Token:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6000
      TabIndex        =   7
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblIngreso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   6000
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TOKEN digital para Renovar su Contraseña:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   8292
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmLogon_Renueva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String

Private Sub btnToken_Activa_Click()
Call sbValida_Token

If Len(vMensaje) = 0 Then
    glogon.Usuario = FlatEdit_Token_Usuario.Text
    frmLogon_Renueva_Clave.Show vbModal
    Unload Me
Else
    MsgBox vMensaje, vbExclamation
End If
End Sub

Private Sub btnToken_Envia_Click()

Call sbValida_Datos

If Len(vMensaje) = 0 Then
    Call sbToken_Envia
Else
    MsgBox vMensaje, vbExclamation
End If
    
End Sub

Private Sub Form_Load()
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 lblOpcion.BackColor = RGB(70, 111, 178)
 
 Call RbtToken_Click(0)
 Call PGX_Portal_Access
 
 
End Sub

Private Sub RbtToken_Click(Index As Integer)

lblToken(0).Visible = False
lblToken(1).Visible = False

FlatEdit_Usuario.Visible = False
FlatEdit_Email.Visible = False

btnToken_Envia.Visible = False

lblToken(2).Visible = False
lblToken(3).Visible = False

FlatEdit_Token_Usuario.Visible = False
FlatEdit_Token.Visible = False

btnToken_Activa.Visible = False

 If Index = 0 Then
    lblOpcion.Caption = "Indique un código de Usuario y Correo Electrónico registrado en el sistema"
    lblToken(0).Visible = True
    lblToken(1).Visible = True
    
    FlatEdit_Usuario.Visible = True
    FlatEdit_Email.Visible = True
    
    btnToken_Envia.Visible = True
    
 Else
    lblOpcion.Caption = "Ingrese su Usuario y Token recibido en su correo:"
    lblToken(2).Visible = True
    lblToken(3).Visible = True
 
    FlatEdit_Token_Usuario.Visible = True
    FlatEdit_Token.Visible = True
    
    btnToken_Activa.Visible = True
 
 End If
 
End Sub

Private Function Aleatorio(Minimo As Long, Maximo As Long) As Long
    Randomize ' inicializar la semilla
    Aleatorio = CLng((Minimo - Maximo) * Rnd + Maximo)
End Function


Private Function fxToken_Genera() As String
Dim i As Integer, Largo As Integer, Cadena As String, x As Integer

 Cadena = ""
 Largo = Aleatorio(8, 16)
 For x = 1 To Largo
        i = Aleatorio(40, 90)
        Cadena = Cadena & Chr(i)
 Next x
 
 fxToken_Genera = Cadena
End Function


Private Sub sbValida_Datos()
Dim strSQL As String, rs As New ADODB.Recordset

vMensaje = ""

strSQL = "select dbo.fxSEG_Token_Valida_Usuario('" & FlatEdit_Usuario.Text & "','" & FlatEdit_Email.Text & "','') as 'Resultado'"
Call OpenRecordSet(rs, strSQL, 1)
If rs!Resultado = 0 Then
    vMensaje = "Datos del Usuario y Correo no Localizados!"
End If

End Sub

Private Sub sbValida_Token()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTokenKey As String

vMensaje = ""
vTokenKey = SIFGlobal.fxStringCifrado(FlatEdit_Token.Text)

strSQL = "select dbo.fxSEG_Token_Valida('" & FlatEdit_Token_Usuario.Text & "','" & vTokenKey & "') as 'Resultado'"
Call OpenRecordSet(rs, strSQL, 1)
Select Case rs!Resultado
    Case 1 'Ok
    Case 0 'No Existe
        vMensaje = "Token no es válido!"
    Case -1 'Está vencido
        vMensaje = "El Token suministrado se encuentra vencido!"
End Select

rs.Close

End Sub


Private Sub sbToken_Envia()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vToken As String, vTokenKey As String

On Error GoTo vError

vToken = fxToken_Genera
vTokenKey = SIFGlobal.fxStringCifrado(vToken)

strSQL = "exec spSEG_Token_Registra '" & FlatEdit_Usuario.Text & "','" & vToken & "','" & vTokenKey & "'"
Call ConectionExecute(strSQL, 1)

FlatEdit_Usuario.Text = ""
FlatEdit_Email.Text = ""

MsgBox "Se ha enviado un TOKEN para renovación de su contraseña a su correo electrónico. Favor revisar su correo e ingresa a la opción de Ingreso del Token", vbInformation

Call RbtToken_Click(1)

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

