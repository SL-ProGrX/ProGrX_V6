VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCajas_Acceso 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acceso a Cajas"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmCajas_Acceso.frx":0000
   ScaleHeight     =   3720
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnConectar 
      Height          =   612
      Left            =   6960
      TabIndex        =   6
      Top             =   2640
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Conectar"
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
      Appearance      =   21
      Picture         =   "frmCajas_Acceso.frx":6061
   End
   Begin VB.TextBox txtClave 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   5640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   5640
      TabIndex        =   7
      Top             =   1320
      Width           =   3012
      _Version        =   1572864
      _ExtentX        =   5318
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16579836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16579836
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Login de Cajas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   495
      Index           =   1
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmCajas_Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConectar_Click()
 Call sbAbreCaja
End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 5

txtUsuario = glogon.Usuario
txtClave = ""

ModuloCajas.mApertura = 0
ModuloCajas.mCaja = ""

'Abre Cajas Disponibles y Abiertas para el Usuario
If ModuloCajas.mCierreActiva Then
    strSQL = "exec spCajas_CierreCajasDisponibles '" & txtUsuario.Text & "'"
Else
    strSQL = "exec spCajas_AperturaCajasDisponibles '" & txtUsuario.Text & "'"
End If
Call sbCbo_Llena_New(cbo, strSQL, False, True)

End Sub


Private Sub sbAbreCaja()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select count(*) as 'Aceptado'" _
       & " from cajas_usuarios where usuario= '" & txtUsuario.Text _
       & "' and contrasena = '" & SIFGlobal.fxStringCifrado(txtClave) _
       & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "'"

Call OpenRecordSet(rs, strSQL)

If rs!aceptado > 0 Then
     rs.Close
     strSQL = "exec spCajas_AbreCaja '" & cbo.ItemData(cbo.ListIndex) & "','" & txtUsuario.Text & "', 'ProGrX_" & glogon.AppVersion & "'"
     Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF And Not rs.BOF Then
       ModuloCajas.mApertura = rs!Cod_Apertura
       ModuloCajas.mCaja = rs!Cod_Caja
       ModuloCajas.mUsuario = txtUsuario.Text
       ModuloCajas.mDescripcion = rs!CajaDesc
       ModuloCajas.mTipoCierre = rs!Cierre_Tipo
       
       ModuloCajas.mOficina = rs!COD_OFICINA
       ModuloCajas.mOficinaDesc = rs!OficinaDesc
       ModuloCajas.mUnidad = rs!Cod_Unidad
       ModuloCajas.mCentroCosto = rs!Cod_Centro_Costo
       
       Unload Me
       Exit Sub
     Else
       MsgBox "No existe Apertura Disponible para esta caja o se encuentra en uso por otro usuario!", vbExclamation
     End If
     
Else
   MsgBox "No se encuentra autorizado para utilizar esta caja...", vbCritical

End If
rs.Close
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
ModuloCajas.mCierreActiva = False
End Sub


Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
 Call sbAbreCaja
End If

End Sub


 


