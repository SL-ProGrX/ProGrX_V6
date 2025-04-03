VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSYS_APL_Usuarios_Permisos 
   Caption         =   "APL: Administración de Permisos"
   ClientHeight    =   8388
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   10248
   LinkTopic       =   "Form1"
   ScaleHeight     =   8388
   ScaleWidth      =   10248
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.CheckBox chkConsulta 
      Height          =   252
      Left            =   6000
      TabIndex        =   6
      Top             =   1920
      Width           =   252
      _Version        =   1245185
      _ExtentX        =   444
      _ExtentY        =   444
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   2
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6012
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   9852
      _Version        =   524288
      _ExtentX        =   17378
      _ExtentY        =   10604
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
      MaxCols         =   475
      ScrollBars      =   2
      SpreadDesigner  =   "frmSYS_APL_Usuarios_Permisos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4080
      TabIndex        =   4
      Top             =   1440
      Width           =   5772
      _Version        =   1245185
      _ExtentX        =   10181
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1800
      TabIndex        =   1
      Top             =   1440
      Width           =   2292
      _Version        =   1245185
      _ExtentX        =   4043
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDominio 
      Height          =   312
      Left            =   4080
      TabIndex        =   5
      Top             =   1080
      Width           =   5772
      _Version        =   1245185
      _ExtentX        =   10181
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkRegistro 
      Height          =   252
      Left            =   7440
      TabIndex        =   7
      Top             =   1920
      Width           =   252
      _Version        =   1245185
      _ExtentX        =   444
      _ExtentY        =   444
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.CheckBox chkAutorizador 
      Height          =   252
      Left            =   8880
      TabIndex        =   8
      Top             =   1920
      Width           =   252
      _Version        =   1245185
      _ExtentX        =   444
      _ExtentY        =   444
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Permisos del Usuario en el Dominio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   7932
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Height          =   996
      Left            =   0
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "frmSYS_APL_Usuarios_Permisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vModulo = 38

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

End Sub

