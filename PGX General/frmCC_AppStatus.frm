VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCC_AppStatus 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6396
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10296
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6396
   ScaleWidth      =   10296
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   372
      Left            =   9120
      TabIndex        =   9
      Top             =   5880
      Width           =   972
      _Version        =   1245187
      _ExtentX        =   1714
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cerrar"
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
      Appearance      =   16
   End
   Begin VB.TextBox txtLink 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1728
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmCC_AppStatus.frx":0000
      Top             =   4680
      Width           =   3012
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Text            =   "..."
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox txtVersion 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Text            =   "..."
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   2640
      Top             =   1920
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Link: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Versión: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos de la última versión:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCC_AppStatus.frx":0004
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Versión disponible!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image Image2 
      Height          =   885
      Left            =   0
      Picture         =   "frmCC_AppStatus.frx":0093
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3360
   End
   Begin VB.Image imgInfo 
      Height          =   6372
      Left            =   3360
      Picture         =   "frmCC_AppStatus.frx":43FF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7620
   End
End
Attribute VB_Name = "frmCC_AppStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()

imgInfo.top = 0
imgInfo.Height = Me.Height
imgInfo.Width = Me.Width - imgInfo.Left

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

With glogon
  .strSQL = "exec spSEG_App_Update_Notes '" & .AppName & "'"
  Call OpenRecordSet(.Recordset, .strSQL, 1)
  
  txtVersion.Text = .Recordset!Version
  txtFecha.Text = Format(.Recordset!fecha, "dd/mm/yyyy")
  txtLink.Text = .Recordset!notas

End With

End Sub
