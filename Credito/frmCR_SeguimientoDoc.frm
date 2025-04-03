VERSION 5.00
Begin VB.Form frmCR_SeguimientoDoc 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Digite el Número de CK"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Verificar"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtDoc2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtDoc1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   0
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   0
      X2              =   6240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label2 
      Caption         =   "# Verificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "# Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Left            =   2880
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Digite el Número de Cheque, con el cual se va a desembolsar este préstamo"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
End
Attribute VB_Name = "frmCR_SeguimientoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

If Len(Trim(txtDoc1)) = 0 Or Len(Trim(txtDoc2)) = 0 Then
  MsgBox "No se ha especificado el número del documento...", vbExclamation
  Exit Sub
End If

If Trim(txtDoc1) <> Trim(txtDoc2) Then
  MsgBox "El número del documento no concuerda con su verificación...", vbExclamation
  Exit Sub
End If

strSQL = "select isnull(count(*),0) as Existe from Tes_Transacciones where ndocumento = '" _
       & Trim(txtDoc1) & "' and Tipo = 'CK' and id_banco in(select cod_banco from" _
       & " reg_creditos where id_solicitud = " & Operacion.Operacion & ")"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  MsgBox "El documento especificado ya existe registrado en Tesorería...", vbExclamation
  rs.Close
  Exit Sub
End If
rs.Close

i = MsgBox("Esta seguro que este " & txtDoc1 & " de documento es correcto", vbYesNo)
If i = vbYes Then
  Operacion.Valida = True
  Operacion.Documento = txtDoc1
  Unload Me
Else
  Operacion.Valida = False
  Operacion.Documento = ""
End If


End Sub

Private Sub Form_Load()
Operacion.Valida = False
Operacion.Documento = ""
End Sub

Private Sub txtDoc1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDoc2.SetFocus
End Sub

Private Sub txtDoc2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
End Sub
