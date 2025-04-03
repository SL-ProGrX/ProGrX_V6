VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCajas_Clave 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cajas: Cambio de Clave"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswCajas 
      Height          =   2052
      Left            =   2280
      TabIndex        =   6
      Top             =   2280
      Width           =   5052
      _Version        =   1310723
      _ExtentX        =   8911
      _ExtentY        =   3619
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnCambiar 
      Height          =   615
      Left            =   5640
      TabIndex        =   7
      Top             =   4560
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2778
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "Cambiar"
      BackColor       =   -2147483633
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCajas_Clave.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtClaveSistema 
      Height          =   312
      Left            =   2280
      TabIndex        =   8
      Top             =   1320
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   2280
      TabIndex        =   9
      Top             =   1800
      Width           =   3012
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   315
      Left            =   2280
      TabIndex        =   10
      Top             =   960
      Width           =   3015
      _Version        =   1310723
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "(...Para Validación!)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   1320
      Width           =   1860
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave (Sistema)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label lblCajas 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambiar clave para las cajas siguientes ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1980
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Clave Nueva"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   2
      Left            =   720
      TabIndex        =   0
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio de Clave de Usuario de Cajas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   6615
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmCajas_Clave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnCambiar_Click()
 Call sbCambioClave
End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 5


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtUsuario = glogon.Usuario
txtClave = ""

With lswCajas.ColumnHeaders
    .Clear
    .Add , , "Código", 1440
    .Add , , "Descripción", 3440
End With

vPaso = True
    strSQL = "select C.cod_caja,C.Descripcion,PERIOCIDAD_CONTRASENA" _
            & " from cajas_definicion C inner join cajas_usuarios U on C.cod_caja = U.cod_caja and U.usuario = '" & glogon.Usuario & "'" _
            & " where C.Activa = 1 order by C.cod_caja"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswCajas.ListItems.Add(, , rs!cod_caja)
         itmX.SubItems(1) = rs!Descripcion
         itmX.Tag = rs!Periocidad_Contrasena
         itmX.Checked = True
     rs.MoveNext
    Loop
    rs.Close
vPaso = False


End Sub


Private Sub sbCambioClave()
Dim strSQL As String, vClave As String
Dim i As Integer

On Error GoTo vError

If Len(Trim(txtClave.Text)) = 0 Then
   MsgBox "No se ha especificado ninguna clave!", vbExclamation
   Exit Sub
End If

If txtClaveSistema.Text <> glogon.Clave Then
 MsgBox " - La Clave del Sistema digitada no corresponde a su usuario de Sistema!", vbExclamation
 Exit Sub
End If

vClave = SIFGlobal.fxStringCifrado(Trim(txtClave.Text))


With lswCajas.ListItems
 For i = 1 To .Count
   If .Item(i).Checked Then
        strSQL = "update cajas_usuarios set contrasena = '" & vClave _
               & "', Contrasena_Renovacion = dateadd(d," & .Item(i).Tag & ",dbo.MyGetdate())" _
               & " where cod_Caja = '" & .Item(i).Text & "' and usuario = '" & txtUsuario.Text & "'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Cambio de Clave de Caja..:" & .Item(i).Text & " ..Usuario.:" & txtUsuario.Text)
        
   End If
 Next i
End With

MsgBox "Cambio de Clave de Usuario de Cajas Existoso!", vbInformation
Unload Me

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

