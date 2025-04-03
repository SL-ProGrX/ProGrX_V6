VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Begin VB.Form frmIVR_Rec_Bancos_Registro 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Transacción (Bancos, Transitoria)"
   ClientHeight    =   8508
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   11172
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8508
   ScaleWidth      =   11172
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   1932
      Left            =   0
      TabIndex        =   36
      Top             =   6000
      Width           =   11172
      _Version        =   1310720
      _ExtentX        =   19706
      _ExtentY        =   3408
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
      Appearance      =   16
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   4440
      Top             =   120
   End
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   492
      Index           =   0
      Left            =   480
      TabIndex        =   9
      Top             =   2520
      Width           =   1572
      _Version        =   1310720
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Bancos"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtInversionId 
      Height          =   492
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   1812
      _Version        =   1310720
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtInstrumento 
      Height          =   312
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   9252
      _Version        =   1310720
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtAdministrador 
      Height          =   312
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   9252
      _Version        =   1310720
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtPortafolio 
      Height          =   312
      Left            =   1800
      TabIndex        =   7
      Top             =   1560
      Width           =   9252
      _Version        =   1310720
      _ExtentX        =   16319
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   492
      Index           =   1
      Left            =   2160
      TabIndex        =   10
      Top             =   2520
      Width           =   2412
      _Version        =   1310720
      _ExtentX        =   4254
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cuenta en Transito"
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
   End
   Begin XtremeSuiteControls.GroupBox gbCuenta 
      Height          =   1212
      Left            =   480
      TabIndex        =   20
      Top             =   3360
      Width           =   10212
      _Version        =   1310720
      _ExtentX        =   18013
      _ExtentY        =   2138
      _StockProps     =   79
      Caption         =   "Cuenta en Transitoria: "
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   312
         Left            =   3600
         TabIndex        =   21
         Top             =   480
         Width           =   6012
         _Version        =   1310720
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
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   312
         Left            =   1560
         TabIndex        =   22
         Top             =   480
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaDivisa 
         Height          =   312
         Left            =   8760
         TabIndex        =   44
         Top             =   840
         Width           =   852
         _Version        =   1310720
         _ExtentX        =   1503
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaMonto 
         Height          =   312
         Left            =   6720
         TabIndex        =   23
         Top             =   840
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   36
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   41
         Left            =   5880
         TabIndex        =   24
         Top             =   840
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   372
      Index           =   0
      Left            =   7080
      TabIndex        =   26
      Top             =   5580
      Width           =   1212
      _Version        =   1310720
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nuevo"
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
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   372
      Index           =   1
      Left            =   8400
      TabIndex        =   27
      Top             =   5580
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
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
      Picture         =   "frmIVR_Bancos_Registro.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   372
      Index           =   2
      Left            =   8880
      TabIndex        =   28
      Top             =   5580
      Width           =   492
      _Version        =   1310720
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
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
      Picture         =   "frmIVR_Bancos_Registro.frx":0731
   End
   Begin XtremeSuiteControls.FlatEdit txtDivisa 
      Height          =   312
      Left            =   6240
      TabIndex        =   31
      Top             =   3000
      Width           =   852
      _Version        =   1310720
      _ExtentX        =   1503
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
      Height          =   312
      Left            =   7080
      TabIndex        =   33
      Top             =   3000
      Width           =   1212
      _Version        =   1310720
      _ExtentX        =   2138
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbBancos 
      Height          =   1212
      Left            =   480
      TabIndex        =   11
      Top             =   3360
      Width           =   10212
      _Version        =   1310720
      _ExtentX        =   18013
      _ExtentY        =   2138
      _StockProps     =   79
      Caption         =   "Bancos: "
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   1320
         TabIndex        =   12
         Top             =   480
         Width           =   4332
         _Version        =   1310720
         _ExtentX        =   7641
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboEmitir 
         Height          =   312
         Left            =   6840
         TabIndex        =   13
         Top             =   480
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3620
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboIBAN 
         Height          =   312
         Left            =   1320
         TabIndex        =   14
         Top             =   840
         Width           =   4332
         _Version        =   1310720
         _ExtentX        =   7641
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoDivisa 
         Height          =   312
         Left            =   8880
         TabIndex        =   45
         Top             =   840
         Width           =   852
         _Version        =   1310720
         _ExtentX        =   1503
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoMonto 
         Height          =   312
         Left            =   6840
         TabIndex        =   15
         Top             =   840
         Width           =   2052
         _Version        =   1310720
         _ExtentX        =   3619
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   37
         Left            =   480
         TabIndex        =   19
         Top             =   480
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   38
         Left            =   480
         TabIndex        =   18
         Top             =   840
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "IBAN:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   39
         Left            =   6000
         TabIndex        =   17
         Top             =   480
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   40
         Left            =   6000
         TabIndex        =   16
         Top             =   840
         Width           =   1332
         _Version        =   1310720
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtAd_Registrado 
      Height          =   312
      Left            =   5400
      TabIndex        =   38
      Top             =   8040
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAd_Pendiente 
      Height          =   312
      Left            =   9000
      TabIndex        =   39
      Top             =   8040
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   552
      Left            =   2040
      TabIndex        =   42
      Top             =   4800
      Width           =   8052
      _Version        =   1310720
      _ExtentX        =   14203
      _ExtentY        =   974
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
   Begin XtremeSuiteControls.FlatEdit txtImporte 
      Height          =   312
      Left            =   6240
      TabIndex        =   29
      Top             =   2640
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRequerido 
      Height          =   312
      Left            =   1800
      TabIndex        =   34
      Top             =   8040
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtImporteLocal 
      Height          =   312
      Left            =   8280
      TabIndex        =   46
      Top             =   3000
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   10
      Left            =   8280
      TabIndex        =   47
      Top             =   2760
      Width           =   1932
      _Version        =   1310720
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Importe Local:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   7
      Left            =   960
      TabIndex        =   43
      Top             =   4800
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Notas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   34
      Left            =   4080
      TabIndex        =   41
      Top             =   8040
      Width           =   1092
      _Version        =   1310720
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Registrado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   35
      Left            =   7800
      TabIndex        =   40
      Top             =   8040
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pediente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   492
      Left            =   0
      TabIndex        =   37
      Top             =   5520
      Width           =   11172
      _Version        =   1310720
      _ExtentX        =   19706
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Movimientos Registrados: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   6
      Left            =   -480
      TabIndex        =   35
      Top             =   8040
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Monto Requerido:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   5
      Left            =   4560
      TabIndex        =   32
      Top             =   3000
      Width           =   1452
      _Version        =   1310720
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Divisa:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   4
      Left            =   4560
      TabIndex        =   30
      Top             =   2640
      Width           =   1452
      _Version        =   1310720
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Importe Real:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scGestion 
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   11172
      _Version        =   1310720
      _ExtentX        =   19706
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Gestion: "
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1452
      _Version        =   1310720
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Inversión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Portafolio"
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Administrador"
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
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   2052
      _Version        =   1310720
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Instrumento"
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
   End
End
Attribute VB_Name = "frmIVR_Rec_Bancos_Registro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem, vFecha As Date
Dim vDivisaLocaL As String, vCuenta As String, vAdminCedJur As String


Private Sub btnTransac_Click(Index As Integer)

Select Case Index
    Case 0 'Nuevo
        
        Call sbInicializa
        
    Case 1 'Guardar
        Call sbGuardar
        
    Case 2 'Eliminar
        
        Dim i As Integer
        With lsw.ListItems
            For i = 1 To .Count
                If .Item(i).Checked = True Then
                    strSQL = "delete  IVR_TRANSACCIONES Where TRANSAC_ID = " & .Item(i).Text
                    Call ConectionExecute(strSQL)
                End If
            Next i
        End With
End Select

Call sbTransac_Load

End Sub



Private Sub rbTipo_Click(Index As Integer)


gbBancos.Visible = False
gbCuenta.Visible = False

Select Case Index
    Case 0 'Bancos
        gbBancos.Visible = True
    Case 1 'Cuenta Transito
        gbCuenta.Visible = True
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


On Error GoTo vError

vPaso = True


strSQL = "select COD_DIVISA  From vSys_Divisas  Where DIVISA_LOCAL = 1"
Call OpenRecordSet(rs, strSQL)
   vDivisaLocaL = Trim(rs!cod_Divisa)
rs.Close


rbTipo.Item(1).Value = True
rbTipo.Item(0).Enabled = False



txtImporte.Text = Format(gIVR_Transito.Monto, "Standard")

txtDivisa.Text = gIVR_Transito.Divisa
txtTipoCambio.Text = gIVR_Transito.TipoCambio


txtImporteLocal.Text = Format(gIVR_Transito.Monto * fxSys_Tipo_Cambio_Apl(gIVR_Transito.TipoCambio), "Standard")
txtRequerido.Text = Format(gIVR_Transito.Monto * fxSys_Tipo_Cambio_Apl(gIVR_Transito.TipoCambio), "Standard")


'strSQL = "select rtrim(COD_DIVISA) AS 'Idx', rtrim(DESCRIPCION) as 'ItmX'" _
'       & " From vSys_Divisas"
'Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

vPaso = False

Call sbConsulta(gIVR_Transito.TituloId)
Call sbInicializa
Call sbTransac_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbConsulta(pTituloId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vIVR_INVERSIONES" _
       & " Where Titulo_ID = " & pTituloId
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    txtInversionId.Text = rs!TITULO_ID
    
    txtAdministrador.Text = rs!Administrador_Desc
    txtAdministrador.Tag = rs!Cod_Administrador
    
    vAdminCedJur = rs!Administrador_Id
    
    txtInstrumento.Text = rs!Instrumento_Desc
    txtInstrumento.Tag = rs!Cod_Instrumento
    
    txtPortafolio.Text = rs!Portafolio_Desc
    txtPortafolio.Tag = rs!Cod_Portafolio
    
    txtCuenta.Text = rs!CTA_TRANSITO
    txtCuentaDesc.Text = rs!CTA_TRANSITO_DESC
    
    txtCuentaDivisa.Text = rs!CTA_TRANSITO_DIVISA & ""
    
Else
  Me.MousePointer = vbDefault
  MsgBox "No se Localizó el registro!", vbExclamation
End If
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbInicializa()

vPaso = True

txtCuentaMonto.Text = "0"
txtBancoMonto.Text = "0"

vPaso = False

End Sub


Private Sub sbGuardar()
On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pTipoDoc As String, pMonto As Currency, pDivisa As String, pTipoCambio As Currency
Dim pCuenta As String, pBancoId As Long


Select Case True
    Case rbTipo.Item(0).Value 'Bancos
        
        'TODO: Consultar Banco y Reemplazar estos datos
        pBancoId = 0
        pTipoDoc = "DP"
        pCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
        pMonto = CCur(txtBancoMonto.Text)
        pTipoCambio = CCur(txtTipoCambio.Text)
        pDivisa = txtCuentaDivisa.Text

    Case rbTipo.Item(1).Value 'Cuenta en Transito
        
        pBancoId = 0
        pTipoDoc = "CC"
        pCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
        pMonto = CCur(txtCuentaMonto.Text)
        pTipoCambio = CCur(txtTipoCambio.Text)
        pDivisa = txtCuentaDivisa.Text
        
        
        If pDivisa = vDivisaLocaL Then
           pTipoCambio = 1
        End If
        
End Select

With gIVR_Transito

strSQL = "exec spIVR_TRANSAC_REGISTRA '" & .Tipo _
       & "','" & .Codigo _
       & "','" & .Concepto _
       & "','" & .TipoMov _
       & "','" & pTipoDoc _
       & "','" & pDivisa _
       & "', " & pMonto & "," & pTipoCambio _
       & " ,'" & txtNotas.Text _
       & "','" & pCuenta _
       & "', " & pBancoId _
       & " ,'" & glogon.Usuario & "'"
End With

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbTransac_Load()

Call sbIVR_Transac_Load(lsw, gIVR_Transito.Codigo, gIVR_Transito.Tipo, gIVR_Transito.Concepto)


Dim i As Integer, pMonto As Currency

With lsw.ListItems

pMonto = 0
For i = 1 To .Count
    pMonto = pMonto + CCur(.Item(i).SubItems(3))
Next i

txtAd_Registrado.Text = Format(pMonto, "Standard")

txtAd_Pendiente.Text = Format(CCur(txtRequerido.Text) - pMonto, "Standard")

End With
 
End Sub



Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuenta = gCuenta
End If

End Sub

Private Sub txtCuenta_LostFocus()
On Error GoTo vError

   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuenta.Text, 0))
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text, 0)

vError:

End Sub


Private Sub txtCuentaMonto_GotFocus()
On Error GoTo vError
    txtCuentaMonto.Text = CCur(txtCuentaMonto.Text)
vError:
End Sub

Private Sub txtCuentaMonto_LostFocus()
On Error GoTo vError
    txtCuentaMonto.Text = Format(CCur(txtCuentaMonto.Text), "Standard")
vError:
End Sub
