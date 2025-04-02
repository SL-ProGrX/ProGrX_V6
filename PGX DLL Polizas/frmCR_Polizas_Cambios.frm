VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Polizas_Cambios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizaciones Pendientes de Pólizas"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6735
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   13815
      _Version        =   1441793
      _ExtentX        =   24368
      _ExtentY        =   11880
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
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   0
      Left            =   11280
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Actualizar"
      BackColor       =   16777215
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
      Appearance      =   17
      Picture         =   "frmCR_Polizas_Cambios.frx":0000
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   1275
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   1
      Left            =   12480
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplicar"
      BackColor       =   16777215
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
      Appearance      =   17
      Picture         =   "frmCR_Polizas_Cambios.frx":0700
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   13815
      _Version        =   1441793
      _ExtentX        =   24368
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Lista de Cambios Pendientes de Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualización de Pólizas"
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
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Polizas de Vivienda y Prendario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmCR_Polizas_Cambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Operacion", 2100
    
    .Add , , "Poliza", 1100, vbCenter
    .Add , , "Cuota Actual", 2500, vbRightJustify
    .Add , , "Cuota Diferencia", 2500, vbRightJustify
    .Add , , "Cuota Nueva", 2500, vbRightJustify
    .Add , , "F. Registro", 2100, vbCenter
End With

End Sub
