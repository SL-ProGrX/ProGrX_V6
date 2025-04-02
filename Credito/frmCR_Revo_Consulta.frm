VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_Revo_Consulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Contratos Revolutivos"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   14880
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      _Version        =   1310723
      _ExtentX        =   25003
      _ExtentY        =   1931
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit feContrato 
         Height          =   312
         Left            =   1200
         TabIndex        =   2
         Top             =   600
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   495
         Left            =   9960
         TabIndex        =   1
         Top             =   480
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Buscar"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCR_Revo_Consulta.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit feOperacion 
         Height          =   312
         Left            =   2520
         TabIndex        =   4
         Top             =   600
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feCedula 
         Height          =   312
         Left            =   3840
         TabIndex        =   6
         Top             =   600
         Width           =   1572
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feNombre 
         Height          =   315
         Left            =   5400
         TabIndex        =   8
         Top             =   600
         Width           =   4335
         _Version        =   1310723
         _ExtentX        =   7646
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit feLinea 
         Height          =   312
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1092
         _Version        =   1310723
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   495
         Left            =   11160
         TabIndex        =   13
         Top             =   480
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Excel"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCR_Revo_Consulta.frx":0700
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnNuevo 
         Height          =   495
         Left            =   12600
         TabIndex        =   14
         Top             =   480
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmCR_Revo_Consulta.frx":0FD1
         ImageAlignment  =   4
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Línea"
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
         Height          =   252
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
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
         Height          =   252
         Index           =   3
         Left            =   5400
         TabIndex        =   9
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cédula"
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
         Height          =   252
         Index           =   2
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Operación"
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
         Height          =   252
         Index           =   1
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Contrato"
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
         Height          =   252
         Index           =   0
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1572
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6975
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   13815
      _Version        =   524288
      _ExtentX        =   24368
      _ExtentY        =   12303
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
      MaxCols         =   18
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Revo_Consulta.frx":1603
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmCR_Revo_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next

GroupBox1.Width = Me.Width - 300

vGrid.Width = GroupBox1.Width
vGrid.Height = Me.Height - (vGrid.top + 580)

End Sub
