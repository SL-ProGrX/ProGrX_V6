VERSION 5.00
Begin VB.Form frmSIF_EstadoCuentaTextoCtaCor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Cuenta : Encabezado y Pie de Páagina"
   ClientHeight    =   4980
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7752
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7752
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      Picture         =   "frmSIF_EstadoCuentaTexto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox txtEC_Encabezado 
      Height          =   1395
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   5295
   End
   Begin VB.TextBox txtEC_PiePagina 
      Height          =   1395
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   5295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   7560
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Image img1 
      Height          =   576
      Index           =   2
      Left            =   120
      Picture         =   "frmSIF_EstadoCuentaTexto.frx":013A
      Top             =   0
      Width           =   576
   End
   Begin VB.Label Label3 
      Caption         =   "Notas para Estados de Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   6135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   7920
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Caption         =   "Encabezado"
      Height          =   255
      Index           =   11
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Pie de Página"
      Height          =   255
      Index           =   12
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frmSIF_EstadoCuentaTextoCtaCor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGuardar_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update sif_empresa set ec_Nota01 = '" & txtEC_Encabezado.Text _
          & "',ec_Nota02 = '" & txtEC_PiePagina.Text & "'"

Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Cambio de Encabezado y Pie de Pag. E.C.")

MsgBox "Infomación Guardada Satisfactoriamente...", vbInformation

UnLoad Me
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 10

Me.Icon = Me.Picture

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select * from sif_empresa"
Call OpenRecordSet(rs, strSQL)
  txtEC_Encabezado = rs!EC_Nota01 & ""
  txtEC_PiePagina = rs!EC_Nota02 & ""
rs.Close


End Sub
