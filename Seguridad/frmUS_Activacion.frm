VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmUS_Activacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   3840
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Activar"
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
      Picture         =   "frmUS_Activacion.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2280
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
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Inactivar"
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
      Picture         =   "frmUS_Activacion.frx":0727
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.FlatEdit txtUserName 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   1320
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4471
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Usuario:"
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
      TabIndex        =   3
      Top             =   0
      Width           =   6975
      _Version        =   1310723
      _ExtentX        =   12303
      _ExtentY        =   1508
      _StockProps     =   14
      Caption         =   "Activación de la Cuenta"
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
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   2040
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
End
Attribute VB_Name = "frmUS_Activacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset



Private Sub btnAccion_Click(Index As Integer)

On Error GoTo vError

strSQL = "select Estado from US_usuarios where UserId = " & gEntidad.UserID
Call OpenRecordSet(rs, strSQL)
 If rs!ESTADO = "A" And Index = 0 Then
    Unload Me
 End If
 If rs!ESTADO = "I" And Index = 1 Then
    Unload Me
 End If
rs.Close

Select Case Index
  Case 0 'Activar
     strSQL = "Update US_Usuarios set Estado = 'A' where UserID = " & gEntidad.UserID
     Call ConectionExecute(strSQL)
     
     Call sbSEGCuentaLog("02", txtNotas.Text, glogon.Usuario, gEntidad.Usuario)
  Case 1 'Inactivar
     strSQL = "Update US_Usuarios set Estado = 'I' where UserID = " & gEntidad.UserID
     Call ConectionExecute(strSQL)
     
     Call sbSEGCuentaLog("03", txtNotas.Text, glogon.Usuario, gEntidad.Usuario)

End Select

Unload Me

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub Form_Load()

'vModulo = 13

txtUserName.Text = gEntidad.Usuario

strSQL = "select Estado from US_usuarios where UserId = " & gEntidad.UserID
Call OpenRecordSet(rs, strSQL)
 If rs!ESTADO = "A" Then
    btnAccion.Item(0).Enabled = False
 Else
    btnAccion.Item(1).Enabled = False
 End If
rs.Close


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

