VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmTES_Endosos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emisión de Endosos"
   ClientHeight    =   3348
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9516
   Icon            =   "frmTES_Endosos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3348
   ScaleWidth      =   9516
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   552
      Left            =   7680
      TabIndex        =   1
      Top             =   2400
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   974
      _StockProps     =   79
      Caption         =   "&Aplicar"
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
      Appearance      =   16
      Picture         =   "frmTES_Endosos.frx":6852
   End
   Begin XtremeSuiteControls.ComboBox cboEndoso 
      Height          =   348
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2892
      _Version        =   1245187
      _ExtentX        =   5101
      _ExtentY        =   614
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtEndoso 
      Height          =   348
      Left            =   3000
      TabIndex        =   3
      Top             =   1440
      Width           =   6252
      _Version        =   1245187
      _ExtentX        =   11028
      _ExtentY        =   614
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   8640
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Endosos"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_Endosos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()

If Trim(txtEndoso.Text) <> "" Then
    With frmContenedor.Crt
         .Reset
         .ReportFileName = SIFGlobal.fxPathReportes("Banking_Endoso.rpt")
         .Formulas(0) = "Endoso='" & cboEndoso.Text & Trim(UCase(txtEndoso.Text)) & "'"
         .Destination = crptToPrinter
         .PrintReport
    End With
End If

End Sub

Private Sub Form_Load()
vModulo = 9
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboEndoso.Clear
cboEndoso.AddItem "Léase Correctamente : "
cboEndoso.AddItem "Endoso este Cheque a : "
cboEndoso.AddItem "Cancela Operación No. "
cboEndoso.Text = "Endoso este Cheque a : "
End Sub

