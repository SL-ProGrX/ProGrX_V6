VERSION 5.00
Begin VB.Form frmTestKey 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TestKey"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtTO 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Text            =   "15"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtDb 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Time Out"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Data Base"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Server"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   6120
      X2              =   0
      Y1              =   1800
      Y2              =   1800
   End
End
Attribute VB_Name = "frmTestKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdTest_Click()
Dim xSQL As String, cn As New ADODB.Connection
Dim RootKey As String, RootName As String
Dim RootServer As String, RootDb As String

On Error GoTo vError

RootKey = "HitStoreSSMain"
RootName = "AseccssIn"
RootServer = txtServer
RootDb = txtDb


xSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & RootServer & ";UID=" _
     & RootName & ";PWD=" & RootKey & ";Database=" & RootDb & ";APP=SecuritySystem;"

cn.CommandTimeout = txtTO
cn.Open xSQL

cn.Close

MsgBox "Llave Accesada Correctamente", vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical


End Sub
