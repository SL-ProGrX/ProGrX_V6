VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSYS_APL_Usuarios 
   Caption         =   "APL_Usuarios y Roles de Dominio"
   ClientHeight    =   8100
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8280
      Top             =   1560
   End
   Begin XtremeSuiteControls.GroupBox gbCopia 
      Height          =   1452
      Left            =   240
      TabIndex        =   5
      Top             =   6360
      Width           =   8532
      _Version        =   1245185
      _ExtentX        =   15049
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Copiar Roles y Permisos:"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtCopia_Usuario_D 
         Height          =   312
         Left            =   6120
         TabIndex        =   9
         Top             =   600
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCopia_Usuario_O 
         Height          =   312
         Left            =   3840
         TabIndex        =   7
         Top             =   600
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCopia_Dominio 
         Height          =   312
         Left            =   0
         TabIndex        =   10
         Top             =   600
         Width           =   3852
         _Version        =   1245185
         _ExtentX        =   6795
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   432
         Left            =   6120
         TabIndex        =   12
         Top             =   960
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "Copiar Roles y Permisos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
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
         Caption         =   "Dominio:"
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
         Index           =   5
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Copiar a:"
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
         Left            =   6120
         TabIndex        =   8
         Top             =   360
         Width           =   1692
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
         Index           =   3
         Left            =   3840
         TabIndex        =   6
         Top             =   360
         Width           =   1692
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3972
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   8652
      _Version        =   524288
      _ExtentX        =   15261
      _ExtentY        =   7006
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
      MaxCols         =   478
      SpreadDesigner  =   "frmSYS_APL_Usuarios.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   312
      Left            =   4200
      TabIndex        =   13
      Top             =   1560
      Width           =   492
      _Version        =   1245185
      _ExtentX        =   868
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtDominioId 
      Height          =   312
      Left            =   1800
      TabIndex        =   14
      Top             =   1200
      Width           =   1092
      _Version        =   1245185
      _ExtentX        =   1926
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
   Begin XtremeSuiteControls.FlatEdit txtDominioDesc 
      Height          =   312
      Left            =   2880
      TabIndex        =   15
      Top             =   1200
      Width           =   6012
      _Version        =   1245185
      _ExtentX        =   10604
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
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
      TabIndex        =   2
      Top             =   1560
      Width           =   1692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dominio (Base):"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2172
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Roles de Usuarios en los Dominios vinculados"
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
      TabIndex        =   0
      Top             =   240
      Width           =   7932
   End
   Begin VB.Image imgBanner 
      Height          =   996
      Left            =   0
      Top             =   0
      Width           =   10200
   End
End
Attribute VB_Name = "frmSYS_APL_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
vModulo = 38

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


txtDominioId.Text = gAPL.APL_Dominio
txtDominioDesc.Text = gAPL.APL_DominioDesc

vGrid.MaxRows = 0

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String


strSQL = "exec spAPL_Dominios_Vinculados '" & txtDominioId.Text & "'"
Call sbCbo_Llena_New(cboCopia_Dominio, strSQL, True, True)
End Sub
