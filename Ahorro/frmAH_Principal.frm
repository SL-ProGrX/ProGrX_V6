VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmAH_Principal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Principal de Patrimonio"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10140
   HelpContextID   =   2004
   Icon            =   "frmAH_Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10140
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3252
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   5736
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
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4200
      TabIndex        =   17
      Top             =   480
      Width           =   5532
      _Version        =   1441793
      _ExtentX        =   9758
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2160
      TabIndex        =   16
      Top             =   480
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkObrero 
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   6960
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aporte Obrero"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   9612
      _Version        =   1441793
      _ExtentX        =   16954
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Resumen de Patrimonio:"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtObrero 
         Height          =   312
         Left            =   1680
         TabIndex        =   22
         Top             =   480
         Width           =   1812
         _Version        =   1441793
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPatronal 
         Height          =   312
         Left            =   1680
         TabIndex        =   23
         Top             =   840
         Width           =   1812
         _Version        =   1441793
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCapitalizacion 
         Height          =   312
         Left            =   6480
         TabIndex        =   26
         Top             =   480
         Width           =   1812
         _Version        =   1441793
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCustodia 
         Height          =   312
         Left            =   6480
         TabIndex        =   27
         Top             =   840
         Width           =   1812
         _Version        =   1441793
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   312
         Left            =   6480
         TabIndex        =   30
         Top             =   1200
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   8280
         TabIndex        =   31
         Top             =   1200
         Width           =   612
         _Version        =   1441793
         _ExtentX        =   1080
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
      Begin XtremeSuiteControls.FlatEdit txtAporteCobro 
         Height          =   315
         Left            =   1680
         TabIndex        =   34
         Top             =   1200
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
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
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Aporte al Cobro?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   1335
      End
      Begin XtremeSuiteControls.Label lblFechaCustodia 
         Height          =   312
         Left            =   8400
         TabIndex        =   29
         Top             =   840
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblCapitalizado 
         Height          =   312
         Left            =   8400
         TabIndex        =   28
         Top             =   480
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblFechaPatronal 
         Height          =   312
         Left            =   3600
         TabIndex        =   25
         Top             =   840
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblFechaObrero 
         Height          =   312
         Left            =   3600
         TabIndex        =   24
         Top             =   480
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Capitalización"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   5040
         TabIndex        =   11
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ap.Pat/Custodia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   5040
         TabIndex        =   10
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   5040
         TabIndex        =   9
         Top             =   1200
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Aporte Obrero"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Aporte Patronal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Patrimonio"
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
      Appearance      =   6
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7440
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":0330
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":0654
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":0F30
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":1254
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":1B38
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":1E5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAH_Principal.frx":2178
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Index           =   1
      Left            =   1680
      TabIndex        =   2
      Top             =   3120
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Excedentes"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Histórico"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Index           =   3
      Left            =   4800
      TabIndex        =   4
      Top             =   3120
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Liquidaciones"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   3120
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Informe"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.CheckBox chkPatronal 
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   7200
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aporte Patronal"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkCapitalizacion 
      Height          =   252
      Left            =   2280
      TabIndex        =   14
      Top             =   6960
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Capitalización"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkCustodia 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   7200
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Ap.Pat/Custodia"
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
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   492
      Index           =   0
      Left            =   6600
      TabIndex        =   20
      Top             =   6960
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aportación"
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAH_Principal.frx":2494
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   492
      Index           =   1
      Left            =   8160
      TabIndex        =   21
      Top             =   6960
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Anular"
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAH_Principal.frx":2BB4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   6360
      TabIndex        =   32
      Top             =   3000
      Visible         =   0   'False
      Width           =   3735
      _Version        =   1441793
      _ExtentX        =   6588
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   8880
      TabIndex        =   33
      Top             =   3120
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Appearance      =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   384
      Left            =   0
      TabIndex        =   19
      Top             =   3120
      Width           =   12732
      _Version        =   1441793
      _ExtentX        =   22458
      _ExtentY        =   674
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   5
      Left            =   2160
      TabIndex        =   18
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12492
   End
End
Attribute VB_Name = "frmAH_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem

Dim vRA_Access As Boolean

Public Sub sbConsulta(pCedula As String)
 
lsw.ListItems.Clear

txtNombre.Text = ""
txtObrero.Text = 0
txtPatronal.Text = 0
txtCustodia.Text = 0
txtCapitalizacion.Text = 0
txtTotal.Text = 0


lblFechaObrero.Caption = ""
lblFechaPatronal.Caption = ""
lblFechaCustodia.Caption = ""
lblCapitalizado.Caption = ""


'Valida Acceso a Expediente
vRA_Access = fxSys_RA_Consulta(Trim(pCedula), glogon.Usuario)
 
If Not vRA_Access Then
    MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
    txtCedula.Text = ""
    txtNombre.Text = ""
    Exit Sub
End If


strSQL = "select *, dbo.fxPAT_Info_Aporte_Manual(CEDULA) as 'Pat_Aporte_Manual'" _
       & " from vPAT_Consolidado" _
       & " where cedula = '" & pCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  txtObrero.Text = Format(rs!Obrero, "Standard")
  txtPatronal.Text = Format(rs!Patronal, "Standard")
  txtCustodia.Text = Format(rs!Custodia, "Standard")
  txtCapitalizacion.Text = Format(rs!capitaliza, "Standard")
  
  txtTotal.Text = Format(rs!Obrero + rs!Patronal + rs!Custodia + rs!capitaliza, "Standard")
  
  txtDivisa.Text = Trim(rs!cod_Divisa)
  
  txtAporteCobro.Text = Format(rs!Pat_Aporte_Manual, "Standard")
  
  
  lblFechaObrero.Caption = IIf(IsNull(rs!fecAhorro), "", Format(rs!fecAhorro, "dd/mm/yyyy"))
  lblFechaPatronal.Caption = IIf(IsNull(rs!fecaporte), "", Format(rs!fecaporte, "dd/mm/yyyy"))
  lblFechaCustodia.Caption = IIf(IsNull(rs!fecCustodia), "", Format(rs!fecCustodia, "dd/mm/yyyy"))
  lblCapitalizado.Caption = IIf(IsNull(rs!fecCapitaliza), "", Format(rs!fecCapitaliza, "dd/mm/yyyy"))
  
 
Else

  MsgBox "No se localizó la Persona o sus registros de aportes, verifique...!", vbExclamation
  
End If

rs.Close

Call RefrescaTags(Me)

lsw.Enabled = True
chkCapitalizacion.Enabled = True
chkObrero.Enabled = True
chkPatronal.Enabled = True
chkCustodia.Enabled = True

End Sub

Private Sub btnAccion_Click(Index As Integer)
On Error GoTo vError

If txtCedula.Text = "" Or txtNombre.Text = "" Then Exit Sub

Select Case Index
    Case 0 '"registrar"
        GLOBALES.gCedulaActual = txtCedula.Text
        frmAH_RegistraAhorro.Show vbModal
        Call sbConsulta(txtCedula)
    
    Case 1 '"anular"
        GLOBALES.gCedulaActual = txtCedula.Text
        frmAH_AnulaAhorros.Show vbModal
        Call sbConsulta(txtCedula)
End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnConsulta_Click(Index As Integer)

On Error GoTo vError

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear


btnInforme.Tag = btnConsulta.Item(Index).Caption

Select Case Index
  Case 0 'Patrimonio
            With lsw
                .ColumnHeaders.Add , , "Fecha", 1550
                .ColumnHeaders.Add , , "Proceso", 1220
                .ColumnHeaders.Add , , "Tipo", 1520
                .ColumnHeaders.Add , , "Monto", 1400, 1
                .ColumnHeaders.Add , , "Movimiento", 1400
                .ColumnHeaders.Add , , "Mov.Número", 1400
                .ColumnHeaders.Add , , "Mov.Concepto", 2400
                .ColumnHeaders.Add , , "Mov.Usuario", 1400
        
        
                strSQL = "select A.*,isnull(D.descripcion,'') as TipoDoc,isnull(C.descripcion,'') as 'Concepto'" _
                       & " from ahorro_detallado A left join SIF_Documentos D on A.Tcon = D.Tipo_Documento" _
                       & " left join SIF_Conceptos C on A.cod_Concepto = C.cod_Concepto" _
                       & " where A.cedula = '" & txtCedula & "' and A.Tipo in("
                
                If chkObrero.Value = vbChecked Then
                   strSQL = strSQL & "'O',"
                End If
                If chkPatronal.Value = vbChecked Then
                   strSQL = strSQL & "'P',"
                End If
                If chkCapitalizacion.Value = vbChecked Then
                   strSQL = strSQL & "'C',"
                End If
                If chkCustodia.Value = vbChecked Then
                   strSQL = strSQL & "'X',"
                End If
                   strSQL = strSQL & "'') order by A.fecha desc,A.consec desc"
                Call OpenRecordSet(rs, strSQL)
                
                Do While Not rs.EOF
                    Set itmX = .ListItems.Add(, , Format(rs!fecha, "dd/mm/yyyy"))
                        itmX.SubItems(1) = Format(rs!FechaProc, "####-##")
                      Select Case rs!Tipo
                          Case "P"
                              itmX.SubItems(2) = "Patronal"
                          Case "O"
                              itmX.SubItems(2) = "Obrero"
                          Case "C"
                              itmX.SubItems(2) = "Capitalización"
                          Case "X"
                              itmX.SubItems(2) = "Custodia"
                      End Select
                       itmX.SubItems(3) = Format(rs!Monto, "Standard")
                       
                       If rs!TipoDoc = "" Then
                           itmX.SubItems(4) = fxTipoComprobante(rs!tcon)
                       Else
                           itmX.SubItems(4) = rs!TipoDoc
                       End If
                       itmX.SubItems(5) = rs!nCon & ""
                       itmX.SubItems(6) = rs!CONCEPTO & ""
                       itmX.SubItems(7) = rs!Usuario & ""
                       
                       'Ojo historicamente la NC se utiliza como ND y viceversa
                       If Not IsNull(rs!tcon) Then
                          If rs!tcon = "5" Or rs!tcon = "8" Or rs!tcon = "ND" Or rs!tcon = "LIQ" Then
                             itmX.ForeColor = vbRed
                          End If
                       End If
                       
                    rs.MoveNext
                Loop
                rs.Close
  
  
            End With
  
  
  Case 1 'Excedentes
  
         strSQL = " select  P.Inicio, P.CORTE , E.* " _
                & " from exc_cierre E inner join EXC_PERIODOS P on E.ID_PERIODO = P.ID_PERIODO" _
                & " where E.cedula = '" & txtCedula & "'"

         Call OpenRecordSet(rs, strSQL)
         With lsw
             .ColumnHeaders.Add , , "Desde", 1200
             .ColumnHeaders.Add , , "Hasta", 1200
             .ColumnHeaders.Add , , "Exc.Bruto", 1600, 1
             .ColumnHeaders.Add , , "Capitaliza", 1600, 1
             .ColumnHeaders.Add , , "Renta", 1600, 1
             .ColumnHeaders.Add , , "Exc.Neto", 1600, 1
           Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , Format(rs!Inicio, "dd/MM/yyyy"))
              itmX.SubItems(1) = Format(rs!Corte, "dd/MM/yyyy")
              itmX.SubItems(2) = Format(rs!excedente_bruto, "Standard")
              itmX.SubItems(3) = Format(rs!capitalizado, "Standard")
              itmX.SubItems(4) = Format(rs!Renta, "Standard")
              itmX.SubItems(5) = Format(rs!excedente_final, "Standard")
            rs.MoveNext
           Loop
         End With
         rs.Close
  
  
  Case 2 'Historico
        strSQL = "select A.*, isnull(E.Descripcion,A.EstadoActual) as 'Estado_Desc'" _
               & " from ase_per_aportes A left join AFI_ESTADOS_PERSONA E on A.estadoactual = E.cod_Estado" _
               & " where A.cedula = '" & txtCedula.Text & "' order by A.anio desc,A.mes desc"
        Call OpenRecordSet(rs, strSQL)
         With lsw
             .ColumnHeaders.Add , , "Año", 1100
             .ColumnHeaders.Add , , "Mes", 1100
             .ColumnHeaders.Add , , "Divisa", 1120, vbCenter
             .ColumnHeaders.Add , , "Obrero", 1600, 1
             .ColumnHeaders.Add , , "Patronal", 1600, 1
             .ColumnHeaders.Add , , "Custodia", 1600, 1
             .ColumnHeaders.Add , , "Capital.", 1600, 1
             .ColumnHeaders.Add , , "Estado", 1100
           Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , rs!Anio)
              itmX.SubItems(1) = rs!Mes
              itmX.SubItems(2) = rs!cod_Divisa & ""
              itmX.SubItems(3) = Format(rs!ahorro, "Standard")
              itmX.SubItems(4) = Format(rs!Aporte, "Standard")
              itmX.SubItems(5) = Format(rs!Custodia, "Standard")
              itmX.SubItems(6) = Format(rs!capitaliza, "Standard")
              itmX.SubItems(7) = rs!Estado_Desc & ""
            rs.MoveNext
           Loop
         End With
         rs.Close
  
  
  
  Case 3 'Liquidaciones"
        strSQL = "select Consec,FecLiq,Aporte_Liq,Ahorro_Liq,Extra_Liq,Capitalizado_Liq,Usuario" _
               & " from liquidacion where estado = 'P' and cedula = '" & txtCedula & "' order by fecliq desc"
        Call OpenRecordSet(rs, strSQL)
         With lsw
             .ColumnHeaders.Add , , "No. Liq.", 1200
             .ColumnHeaders.Add , , "Fecha", 1800
             .ColumnHeaders.Add , , "Obrero", 1600, 1
             .ColumnHeaders.Add , , "Patronal", 1600, 1
             .ColumnHeaders.Add , , "Capital.", 1600, 1
             .ColumnHeaders.Add , , "ExtraOrd.", 1600, 1
             .ColumnHeaders.Add , , "Usuario", 1400
           Do While Not rs.EOF
            Set itmX = .ListItems.Add(, , rs!consec)
              itmX.SubItems(1) = Format(rs!fecLiq, "dd/mm/yyyy")
              itmX.SubItems(2) = Format(rs!AHORRO_LIQ, "Standard")
              itmX.SubItems(3) = Format(rs!APORTE_LIQ, "Standard")
              itmX.SubItems(4) = Format(rs!CAPITALIZADO_LIQ, "Standard")
              itmX.SubItems(5) = Format(rs!Extra_Liq, "Standard")
              itmX.SubItems(6) = rs!Usuario & ""
            rs.MoveNext
           Loop
         End With
         rs.Close
  
  

End Select
  
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnInforme_Click()
Dim strSubtitulo As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Patrimonio"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario='Usuario : " & glogon.Usuario & "'"
 .Formulas(2) = "fxFecha='Fecha   : " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 
 
  strSQL = "{SOCIOS.CEDULA}='" & Trim(txtCedula.Text) & "'"

  Select Case btnInforme.Tag
     Case "Patrimonio"
         .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_Aportes.rpt")
     Case "Excedentes"
         .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_AportesExcedentes.rpt")
     Case "Liquidaciones"
         .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_AportesLiquidaciones.rpt")
     Case "Histórico"
         .ReportFileName = SIFGlobal.fxPathReportes("Patrimonio_AportesH.rpt")
    
  End Select
  
    .SelectionFormula = strSQL
    .PrintReport
End With

Me.MousePointer = vbDefault



End Sub

Private Sub chkCapitalizacion_Click()
 Call btnConsulta_Click(0)
End Sub

Private Sub chkExtraOrdinario_Click()
 Call btnConsulta_Click(0)
End Sub


Private Sub chkObrero_Click()
 Call btnConsulta_Click(0)
End Sub

Private Sub chkPatronal_Click()
 Call btnConsulta_Click(0)
End Sub

Public Sub sbConsulta_Externa(pCedula As String)

Call sbConsulta(pCedula)
End Sub

Private Sub Form_Load()

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture


 
'tlbPrincipal.Buttons.Item(1).Enabled = False
'tlbPrincipal.Buttons.Item(2).Enabled = False

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterna"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "cedula"
    gBusquedas.Orden = "cedula"
    frmBusquedas.Show vbModal
    
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
    
    Call sbConsulta(txtCedula)
End If

End Sub

Private Sub txtCedula_LostFocus()
 If txtCedula.Text <> "" Then
  Call sbConsulta(txtCedula)
 End If
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterna"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    
    txtCedula.Text = gBusquedas.Resultado
    txtNombre.Text = gBusquedas.Resultado2
    
    Call sbConsulta(txtCedula)
End If

End Sub
