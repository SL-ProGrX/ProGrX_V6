VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCO_AplFnd_Acuerdos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acuerdos de Aplicación de Fondos a Mora"
   ClientHeight    =   10860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10860
   ScaleWidth      =   12840
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   12840
      _Version        =   1572864
      _ExtentX        =   22648
      _ExtentY        =   11668
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.GroupBox gbAcuerdos 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12615
      _Version        =   1572864
      _ExtentX        =   22251
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Acuerdo:"
      ForeColor       =   4210752
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkApl_Creditos 
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Aplica Créditos"
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
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   0
         Left            =   9840
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmCO_AplFnd_Acuerdos.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   1
         Left            =   11160
         TabIndex        =   4
         Top             =   480
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "frmCO_AplFnd_Acuerdos.frx":0632
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   2
         Left            =   11640
         TabIndex        =   5
         Top             =   480
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         Picture         =   "frmCO_AplFnd_Acuerdos.frx":0D63
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   3360
         TabIndex        =   7
         Top             =   480
         Width           =   6015
         _Version        =   1572864
         _ExtentX        =   10604
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   1560
         TabIndex        =   8
         Top             =   840
         Width           =   7815
         _Version        =   1572864
         _ExtentX        =   13785
         _ExtentY        =   1402
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
      Begin XtremeSuiteControls.DateTimePicker dtpFirma 
         Height          =   330
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   255
         Left            =   8040
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Activo"
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
      End
      Begin XtremeSuiteControls.CheckBox chkApl_Obrero 
         Height          =   255
         Left            =   5160
         TabIndex        =   19
         Top             =   1800
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Aplica Aporte Obrero"
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
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Observaciones"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha Firma"
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   375
      Index           =   3
      Left            =   12360
      TabIndex        =   15
      ToolTipText     =   "Exportar Lista"
      Top             =   3840
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmCO_AplFnd_Acuerdos.frx":0DEC
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   3840
      Width           =   10095
      _Version        =   1572864
      _ExtentX        =   17806
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   18
      Top             =   360
      Width           =   10335
      _Version        =   1572864
      _ExtentX        =   18230
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Acuerdos de Aplicacion de Fondos a Morosidad"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   3480
      Width           =   12855
      _Version        =   1572864
      _ExtentX        =   22675
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Lista de Acuerdo Vigentes"
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
   Begin XtremeShortcutBar.ShortcutCaption scAcuerdos 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   3840
      Width           =   2295
      _Version        =   1572864
      _ExtentX        =   4048
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "No. Acuerdos: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCO_AplFnd_Acuerdos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Identificación", 1500, vbCenter
    .Add , , "Nombre", 4000
    .Add , , "Apl.Créditos", 1500, vbCenter
    .Add , , "Apl.Obrero", 1500, vbCenter
    .Add , , "Firma Boleta", 1300, vbCenter
    .Add , , "Activo", 1200, vbCenter
    .Add , , "Notas", 3500
    .Add , , "Vence Cédula", 1300, vbCenter
    .Add , , "Usuario", 1800, vbCenter
    .Add , , "Fecha", 1800, vbCenter
End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


