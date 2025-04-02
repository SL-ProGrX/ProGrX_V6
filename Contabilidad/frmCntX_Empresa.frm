VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCntX_Empresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro Empresa"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   HelpContextID   =   14
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRegistro 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   3360
      Width           =   5535
   End
   Begin VB.TextBox txtContacto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   3000
      Width           =   5535
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   15
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox txtDireccion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1368
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   1200
      Width           =   5772
   End
   Begin VB.TextBox txtFax 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   13
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtAptoPostal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      TabIndex        =   12
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtTelefono 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   10
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   9
      Top             =   120
      Width           =   5535
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   612
      Left            =   5040
      TabIndex        =   18
      Top             =   3840
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Guardar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCntX_Empresa.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Contacto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Apto.Postal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   3600
      TabIndex        =   3
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Ced.Jur"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   732
   End
End
Attribute VB_Name = "frmCntX_Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbLeeDatos()
Dim rs As New ADODB.Recordset, strSQL As String
strSQL = "select * from CntX_Empresa_Registro"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtNombre = IIf(IsNull(rs!Nombre), "", rs!Nombre)
  txtCedula = IIf(IsNull(rs!cedula_juridica), "", rs!cedula_juridica)
  txtAptoPostal = IIf(IsNull(rs!apto_postal), "", rs!apto_postal)
  txtContacto = IIf(IsNull(rs!contacto), "", rs!contacto)
  txtDireccion = IIf(IsNull(rs!Direccion), "", rs!Direccion)
  txtEmail = IIf(IsNull(rs!Email), "", rs!Email)
  txtFax = IIf(IsNull(rs!fax), "", rs!fax)
  txtRegistro = ""
  txtRegistro.Enabled = False
  If Len(txtRegistro) > 0 Then
   txtRegistro.PasswordChar = "*"
   txtRegistro.Enabled = False
  End If
  txtTelefono = IIf(IsNull(rs!telefono), "", rs!telefono)
End If
rs.Close
End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String

On Error GoTo vError

  strSQL = "delete CntX_Empresa_Registro"
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "insert into CntX_Empresa_Registro(nombre,cedula_juridica,direccion,apto_postal" _
         & ",email,telefono,fax,contacto) values('" & UCase(txtNombre) _
         & "','" & txtCedula & "','" & UCase(txtDireccion) & "','" & txtAptoPostal _
         & "','" & txtEmail & "','" & txtTelefono & "','" & txtFax _
         & "','" & UCase(txtContacto) & "')"
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Registra", "Empresa : " & UCase(txtNombre))
  
  Call sbLeeDatos
  MsgBox "Empresa Registrada Satisfactoriamente...", vbInformation
  Unload Me

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
vModulo = 20
Set Me.Icon = frmContenedor.Icon

Call sbLeeDatos

End Sub

Private Sub txtAptoPostal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTelefono.SetFocus
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAptoPostal.SetFocus
End Sub

Private Sub txtContacto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRegistro.SetFocus
End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail.SetFocus
End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContacto.SetFocus
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub

Private Sub txtRegistro_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cmdGuardar.SetFocus
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFax.SetFocus
End Sub
