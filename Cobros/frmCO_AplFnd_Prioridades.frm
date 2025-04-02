VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCO_AplFnd_Prioridades 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aplicación de Fondos a Mora: Prioridades de Garantías"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Fondos a Créditos con Morosidad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Checked         =   -1  'True
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   10455
      _Version        =   524288
      _ExtentX        =   18441
      _ExtentY        =   9128
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   491
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_AplFnd_Prioridades.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   1320
      Width           =   4215
      _Version        =   1441793
      _ExtentX        =   7435
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aporte Obrero a Créditos con Morosidad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prioridades para Aplicación de Fondos a Mora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   7815
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   11535
   End
End
Attribute VB_Name = "frmCO_AplFnd_Prioridades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String

Private Sub btnOpcion_Click(Index As Integer)

btnOpcion(Index).Checked = True
If Index = 0 Then
    btnOpcion(1).Checked = False
Else
    btnOpcion(0).Checked = False
End If
End Sub

Private Sub Form_Load()

vModulo = 4
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

