VERSION 5.00
Begin VB.Form frmPosCajaClave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Confirmación de Caja"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.TextBox txtClave 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtUsuario 
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
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   120
      Picture         =   "frmPosClaveCaja.frx":0000
      Top             =   240
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5160
      X2              =   0
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccione su Caja y digite su contraseña"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmPosCajaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

i = 0

gCajas.Apertura = 0

txtUsuario = glogon.Usuario
txtClave = ""
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

End Sub

Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

'1. Verificar que el Estado este Activa / en el Load esta validado
'2. Que no se encuentre Bloqueada
'3. Verificar si la caja esta abierta (Apertura) y Sacar el Consecutivo
'   de la apertura.

If i > 3 And KeyCode = vbKeyReturn Then
  MsgBox "No se permiten más intentos...", vbExclamation
  Unload Me
End If

If KeyCode = vbKeyReturn Then
 i = i + 1
 strSQL = "select bloqueo from pv_cajas where usuario = '" _
        & txtUsuario & "' and cod_caja = '" & fxCodigoCbo(cbo) & "' and clave = '" _
        & fxPosEncrypta(txtClave) & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs.EOF And rs.BOF Then
   MsgBox "Caja: verifique su Usuario y Clave para Esta Caja ...", vbExclamation
 Else
  If rs!bloqueo = 0 Then
     gCajas.Caja = fxCodigoCbo(cbo)
     gCajas.Usuario = txtUsuario
     
     'Consulta la caja para verificar que tenga una apertura existente
     rs.Close
     strSQL = "select cod_ac from pv_cajas_ac where cod_caja = '" & gCajas.Caja _
            & "' and usuario = '" & gCajas.Usuario & "' and estado = 'A'"
     Call OpenRecordSet(rs, strSQL)
     If Not rs.EOF And Not rs.BOF Then
        gCajas.Apertura = rs!cod_ac
        Unload Me
     Else
        MsgBox "Esta caja no tiene apertura existente, debe abrirla primero...", vbExclamation
     End If
     
    
  Else
    MsgBox "La Caja se encuentra Bloqueada...", vbExclamation
  End If 'Bloqueo
 End If 'Select cajas
 rs.Close

End If

End Sub
