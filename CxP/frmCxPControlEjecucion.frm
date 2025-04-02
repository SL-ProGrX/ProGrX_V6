VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCxPControlEjecucion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Control de Pagos : Ejecución de Pagos (Envio a Tesorería para Desembolsar)"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10995
   Icon            =   "frmCxPControlEjecucion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   10995
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2532
      Left            =   120
      TabIndex        =   23
      Top             =   2280
      Width           =   10812
      _Version        =   1572864
      _ExtentX        =   19071
      _ExtentY        =   4466
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
      UseVisualStyle  =   0   'False
   End
   Begin VB.Frame fraEjeProv 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2052
      Left            =   2490
      TabIndex        =   0
      Top             =   -30
      Visible         =   0   'False
      Width           =   5532
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo Pago"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Crédito (Dias)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   9
         Left            =   480
         TabIndex        =   9
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   10
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label lblEjeUltimoPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   7
         Top             =   240
         Width           =   1932
      End
      Begin VB.Label lblEjeProvCredito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   6
         Top             =   600
         Width           =   1932
      End
      Begin VB.Label lblEjeProvSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   1932
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cargo Flotante [M]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   11
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1932
      End
      Begin VB.Label lblEjeProvCarPerSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   3
         Top             =   1320
         Width           =   1932
      End
      Begin VB.Label lblEjeProvCarPerPorc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   2
         Top             =   1680
         Width           =   1932
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cargo Flotante [%]"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   12
         Left            =   480
         TabIndex        =   1
         Top             =   1680
         Width           =   1932
      End
   End
   Begin XtremeSuiteControls.CheckBox chkEjeTodosLsw 
      Height          =   216
      Left            =   240
      TabIndex        =   61
      Top             =   2000
      Width           =   216
      _Version        =   1572864
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.CheckBox chkPagosPendientes 
      Height          =   252
      Left            =   3120
      TabIndex        =   41
      Top             =   120
      Width           =   4452
      _Version        =   1572864
      _ExtentX        =   7853
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Proveedores con pagos pendientes al vencimiento"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   3372
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   10812
      _Version        =   1572864
      _ExtentX        =   19071
      _ExtentY        =   5948
      _StockProps     =   79
      Caption         =   "Resumen de Pago:"
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   1920
         TabIndex        =   40
         Top             =   2520
         Width           =   4572
         _Version        =   1572864
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   675
         Left            =   9120
         TabIndex        =   25
         Top             =   2520
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   1191
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmCxPControlEjecucion.frx":6852
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1920
         TabIndex        =   26
         Top             =   1800
         Width           =   4572
         _Version        =   1572864
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   312
         Left            =   1920
         TabIndex        =   27
         Top             =   2880
         Width           =   4572
         _Version        =   1572864
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipoPago 
         Height          =   312
         Left            =   3960
         TabIndex        =   28
         Top             =   2160
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkPagoTercero 
         Height          =   252
         Left            =   120
         TabIndex        =   48
         Top             =   1800
         Width           =   3132
         _Version        =   1572864
         _ExtentX        =   5524
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Pago a Tercero"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkCambioBanco 
         Height          =   252
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         Width           =   3132
         _Version        =   1572864
         _ExtentX        =   5524
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cambio de Cuenta Bancaria:"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalFactura 
         Height          =   312
         Left            =   1920
         TabIndex        =   50
         Top             =   600
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCargos 
         Height          =   312
         Left            =   1920
         TabIndex        =   51
         Top             =   960
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPagoTotal 
         Height          =   312
         Left            =   1920
         TabIndex        =   52
         Top             =   1320
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalFacturaDivisaReal 
         Height          =   312
         Left            =   4200
         TabIndex        =   53
         Top             =   600
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCargosDivisaReal 
         Height          =   312
         Left            =   4200
         TabIndex        =   54
         Top             =   960
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPagoTotalDivisaReal 
         Height          =   312
         Left            =   4200
         TabIndex        =   55
         Top             =   1320
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4043
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPagoBancos 
         Height          =   312
         Left            =   6600
         TabIndex        =   56
         Top             =   1320
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencial 
         Height          =   312
         Left            =   8760
         TabIndex        =   57
         Top             =   1320
         Width           =   1812
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoDivisa 
         Height          =   312
         Left            =   7320
         TabIndex        =   58
         Top             =   2520
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Text            =   "[Divisa]"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBancoTipoCambio 
         Height          =   312
         Left            =   7320
         TabIndex        =   59
         Top             =   2880
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Text            =   "1"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Pago de "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   13
         Left            =   120
         TabIndex        =   39
         Top             =   1320
         Width           =   1692
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6480
         TabIndex        =   38
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T.C.:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   37
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblEncabezado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Divisa Real "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   2
         Left            =   4200
         TabIndex        =   36
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label lblEncabezado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Divisa Local "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   0
         Left            =   1920
         TabIndex        =   35
         Top             =   360
         Width           =   2292
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "( - ) Cargos "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Total Facturas "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1692
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pago en Bancos:"
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
         Left            =   6960
         TabIndex        =   32
         Top             =   1080
         Width           =   1812
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencial:"
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
         Left            =   9000
         TabIndex        =   31
         Top             =   1080
         Width           =   1572
      End
      Begin VB.Label lblPagoTitulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta del Proveedor:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   30
         Top             =   2880
         Width           =   1812
      End
      Begin VB.Label lblPagoTitulo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta / Desembolso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   29
         Top             =   2520
         Width           =   1812
      End
   End
   Begin XtremeSuiteControls.RadioButton optTipo 
      Height          =   252
      Index           =   0
      Left            =   5760
      TabIndex        =   21
      Top             =   1200
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Desembolsar"
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
      Value           =   -1  'True
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9480
      TabIndex        =   13
      Top             =   840
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   492
      Left            =   7920
      TabIndex        =   17
      Top             =   1320
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmCxPControlEjecucion.frx":702A
   End
   Begin XtremeSuiteControls.PushButton btnDetalle 
      Height          =   492
      Left            =   9240
      TabIndex        =   18
      Top             =   1320
      Width           =   732
      _Version        =   1572864
      _ExtentX        =   1291
      _ExtentY        =   868
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxPControlEjecucion.frx":772A
   End
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   312
      Left            =   1200
      TabIndex        =   19
      Top             =   1200
      Width           =   4452
      _Version        =   1572864
      _ExtentX        =   7858
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboCargo 
      Height          =   312
      Left            =   1200
      TabIndex        =   20
      Top             =   1560
      Width           =   4452
      _Version        =   1572864
      _ExtentX        =   7858
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.RadioButton optTipo 
      Height          =   372
      Index           =   1
      Left            =   5760
      TabIndex        =   22
      Top             =   1440
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Cerrar / Excluir"
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
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   252
      Left            =   3120
      TabIndex        =   42
      Top             =   480
      Width           =   4452
      _Version        =   1572864
      _ExtentX        =   7853
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Facturas Generadas por TODOS los Usuarios"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   315
      Left            =   1200
      TabIndex        =   43
      Top             =   120
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   556
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtCodProv 
      Height          =   330
      Left            =   1200
      TabIndex        =   44
      Top             =   840
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtProveedor 
      Height          =   330
      Left            =   2880
      TabIndex        =   45
      Top             =   840
      Width           =   6372
      _Version        =   1572864
      _ExtentX        =   11239
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   330
      Left            =   1200
      TabIndex        =   46
      Top             =   480
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.CheckBox chkCargosVencimiento 
      Height          =   252
      Left            =   7680
      TabIndex        =   47
      Top             =   120
      Width           =   4452
      _Version        =   1572864
      _ExtentX        =   7853
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cargos al Cobro hasta el vencimiento?"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnReActivar 
      Height          =   492
      Left            =   10200
      TabIndex        =   62
      ToolTipText     =   "Re-Activa Facturas enviadas a Bancos y que se eliminó la solicitud de desembolso"
      Top             =   1320
      Width           =   732
      _Version        =   1572864
      _ExtentX        =   1291
      _ExtentY        =   868
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCxPControlEjecucion.frx":7E43
   End
   Begin XtremeShortcutBar.ShortcutCaption scTitulo 
      Height          =   372
      Left            =   120
      TabIndex        =   60
      Top             =   1920
      Width           =   10812
      _Version        =   1572864
      _ExtentX        =   19071
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Seleccione los pagos disponibles para realizar los desembolsos respectivos"
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
      Alignment       =   1
   End
   Begin VB.Label lblCargo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmCxPControlEjecucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vInicia As Boolean, vScroll As Boolean, vUnidadOmision As String, vConceptoOmision As String
Dim vDivisa As String, vTipoCambio As Currency, vPaso As Boolean, vDivisaFuncional As String
Dim vProvCedJur As String

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnDetalle_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCargosFecha As Date


If Len(txtCodProv) = 0 Then Exit Sub

fraEjeProv.Visible = IIf((fraEjeProv.Visible = True), False, True)
If Not fraEjeProv.Visible Then Exit Sub

If chkCargosVencimiento.Value = xtpChecked Then
    vCargosFecha = dtpVence.Value
Else
    vCargosFecha = fxFechaServidor
End If


'Saldo, Dias de Pagos, etc
strSQL = "select credito_plazo,ultimo_pago,saldo from cxp_proveedores" _
       & " where cod_proveedor = " & txtCodProv
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  lblEjeProvCredito.Caption = rs!credito_plazo
  lblEjeProvSaldo.Caption = Format(rs!Saldo, "Standard")
  lblEjeUltimoPago.Caption = Format((rs!ultimo_pago & ""), "yyyy/mm/dd")
End If
rs.Close

'Sacar los cargos periodicos por Montos para Deduccir hasta saldo 0
strSQL = "select dbo.fxCxP_CargoFlotanteSaldo(" & txtCodProv.Text & ",'" & Format(vCargosFecha, "yyyy/mm/dd") & " 23:59:59') as 'SaldoCargoFlotante'"
Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then
     lblEjeProvCarPerSaldo.Caption = Format(rs!SaldoCargoFlotante, "Standard")
  Else
     lblEjeProvCarPerSaldo.Caption = "0.00"
  End If
  rs.Close

'Cargos Porcentuales
strSQL = "select isnull(sum(valor),0) as valor from cxp_cargosPer where cod_proveedor = " _
       & txtCodProv & " and tipo = 'P' and vence >= '" & Format(vCargosFecha, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   lblEjeProvCarPerPorc.Caption = Format(rs!Valor, "Standard")
Else
   lblEjeProvCarPerPorc.Caption = "0.00"
End If
rs.Close



End Sub


Private Sub btnReActivar_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "update P set P.TESORERIA = Null, P.FECHA_TRASLADA = Null, P.USER_TRASLADA = Null" _
       & " from CXP_PAGOPROV P left join TES_TRANSACCIONES T on P.TESORERIA = T.NSOLICITUD" _
       & " Where IsNull(P.tesoreria, 0) > 0" _
       & " and T.NSOLICITUD is null"
Call ConectionExecute(strSQL)


Call Bitacora("Aplica", "Revisión de Pagos de Facturas en Bancos con Solicitud eliminada")

Me.MousePointer = vbDefault

MsgBox "Revisión de Pagos de Facturas en Bancos realizada satisfactoriamente!"

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBanco_Click()
Dim strSQL As String, rs As New ADODB.Recordset


If vPaso Then Exit Sub
If cboBanco.ListCount <= 0 Then Exit Sub
If Not IsNumeric(txtPagoTotal.Text) Then Exit Sub

strSQL = "select id_banco,descripcion,cod_divisa" _
       & ",dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",COD_DIVISA,dbo.MyGetdate(),'V') as 'TipoCambio'" _
       & " from Tes_Bancos where id_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)
  txtBancoDivisa.Text = Trim(rs!COD_DIVISA)
  txtBancoTipoCambio.Text = Format(rs!TipoCambio, "Standard")
rs.Close

If CCur(txtPagoTotal.Text) > 0 Then
   If txtBancoDivisa.Text = vDivisaFuncional And cboDivisa.ItemData(cboDivisa.ListIndex) <> vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotal.Text
      txtDiferencial.Text = "0"
   End If

   If txtBancoDivisa.Text = vDivisaFuncional And cboDivisa.ItemData(cboDivisa.ListIndex) = vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotal.Text
      txtDiferencial.Text = "0"
   End If

   If txtBancoDivisa.Text = cboDivisa.ItemData(cboDivisa.ListIndex) And txtBancoDivisa.Text <> vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotalDivisaReal.Text
      txtDiferencial.Text = CCur(txtPagoTotal.Text) - (CCur(txtPagoTotalDivisaReal.Text) * CCur(txtBancoTipoCambio.Text))
   End If
   
   
   If txtBancoDivisa.Text <> cboDivisa.ItemData(cboDivisa.ListIndex) And txtBancoDivisa.Text <> vDivisaFuncional Then
      txtPagoBancos.Text = CCur(txtPagoTotal.Text) / fxSys_Tipo_Cambio_Apl(CCur(txtBancoTipoCambio.Text))
      txtDiferencial.Text = "0" 'CCur(txtPagoTotal.Text) - (CCur(txtPagoTotalDivisaReal.Text) * CCur(txtBancoTipoCambio.Text))
   End If
   
End If
txtPagoBancos.Text = Format(CCur(txtPagoBancos.Text), "Standard")
txtDiferencial.Text = Format(CCur(txtDiferencial.Text), "Standard")
'spSys_Cuentas_Bancarias(@Identificacion varchar(30), @BancoId int, @DivisaCheck smallint = 0)
strSQL = "exec spSys_Cuentas_Bancarias '" & vProvCedJur & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",0"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)


End Sub



Private Sub cboDivisa_Click()
If vPaso Then Exit Sub
If cboDivisa.ListCount <= 0 Then Exit Sub

Call sbBuscar

End Sub

Private Sub chkCambioBanco_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


vPaso = True
If chkCambioBanco.Value = vbChecked Then
    cboBanco.Clear
    
    strSQL = "exec spCxP_Bancos_Autorizados"
    Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
    
    cboBanco.Enabled = True

Else
    If IsNumeric(txtCodProv.Text) Then
        Call sbBancoDefault(txtCodProv.Text)
    End If
    cboBanco.Enabled = False
End If

vPaso = False
Call cboBanco_Click


vError:

End Sub

Private Sub sbLimpiaDatos(Optional vTodo As Boolean = True)
txtPagoTotal.Text = 0
txtPagoTotalDivisaReal.Text = 0

txtTotalCargos.Text = "0"
txtTotalCargosDivisaReal.Text = "0"

txtTotalFactura.Text = "0"
txtTotalFacturaDivisaReal.Text = "0"

txtPagoBancos.Text = "0"
txtDiferencial.Text = "0"

cboCuenta.Clear

If vTodo Then
    lsw.ListItems.Clear
End If
End Sub

Private Sub chkCargosVencimiento_Click()
Call sbBuscar
End Sub

Private Sub chkEjeTodosLsw_Click()
Dim i As Integer

Call sbLimpiaDatos(False)

vPaso = True

For i = 1 To lsw.ListItems.Count
 lsw.ListItems.Item(i).Checked = chkEjeTodosLsw.Value
 If lsw.ListItems.Item(i).Checked Then
  txtTotalFactura.Text = CCur(txtTotalFactura.Text) + CCur(lsw.ListItems.Item(i).SubItems(2))
  txtTotalFacturaDivisaReal.Text = CCur(txtTotalFacturaDivisaReal.Text) + CCur(lsw.ListItems.Item(i).SubItems(11))
  
  txtTotalCargos.Text = CCur(txtTotalCargos.Text) + CCur(lsw.ListItems.Item(i).SubItems(3)) + CCur(lsw.ListItems.Item(i).SubItems(4))
  txtTotalCargosDivisaReal.Text = CCur(txtTotalCargosDivisaReal.Text) + CCur(lsw.ListItems.Item(i).SubItems(13))
  txtPagoTotal.Text = CCur(txtPagoTotal.Text) + CCur(lsw.ListItems.Item(i).SubItems(5))
  txtPagoTotalDivisaReal.Text = CCur(txtPagoTotalDivisaReal.Text) + (CCur(lsw.ListItems.Item(i).SubItems(11)) - CCur(lsw.ListItems.Item(i).SubItems(13)))
 End If
Next i

If CCur(txtPagoTotal.Text) > 0 Then
   If txtBancoDivisa.Text = vDivisaFuncional And cboDivisa.ItemData(cboDivisa.ListIndex) <> vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotal.Text
      txtDiferencial.Text = "0"
   End If

   If txtBancoDivisa.Text = vDivisaFuncional And cboDivisa.ItemData(cboDivisa.ListIndex) = vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotal.Text
      txtDiferencial.Text = "0"
   End If

   If txtBancoDivisa.Text = cboDivisa.ItemData(cboDivisa.ListIndex) And txtBancoDivisa.Text <> vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotalDivisaReal.Text
      txtDiferencial.Text = CCur(txtPagoTotal.Text) - (CCur(txtPagoTotalDivisaReal.Text) * CCur(txtBancoTipoCambio.Text))
   End If
End If

'Formato
txtPagoTotal.Text = Format(CCur(txtPagoTotal.Text), "Standard")
txtPagoTotalDivisaReal.Text = Format(CCur(txtPagoTotalDivisaReal.Text), "Standard")

txtTotalCargos.Text = Format(CCur(txtTotalCargos.Text), "Standard")
txtTotalCargosDivisaReal.Text = Format(CCur(txtTotalCargosDivisaReal.Text), "Standard")

txtTotalFactura.Text = Format(CCur(txtTotalFactura.Text), "Standard")
txtTotalFacturaDivisaReal.Text = Format(CCur(txtTotalFacturaDivisaReal.Text), "Standard")

txtPagoBancos.Text = Format(CCur(txtPagoBancos.Text), "Standard")
txtDiferencial.Text = Format(CCur(txtDiferencial.Text), "Standard")

vPaso = False

End Sub

Private Sub chkPagosPendientes_Click()
Call sbBuscar
End Sub

Private Sub chkPagoTercero_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodProv) Then Exit Sub

If chkPagoTercero.Value = vbChecked Then
 strSQL = "select * from cxp_autorizaciones where cod_proveedor = " & txtCodProv
 Call OpenRecordSet(rs, strSQL, 0)
 cbo.Clear
 Do While Not rs.EOF
  cbo.AddItem rs!Nombre
  rs.MoveNext
 Loop
 rs.Close
 cbo.Enabled = True

Else
 cbo.Enabled = False
End If

vError:

End Sub


Private Sub sbBancoDefault(vProveedor As Long)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select cod_Banco, dbo.fxSys_Cuenta_Bancos_Desc(cod_Banco) as 'Cuenta', CedJur" _
       & " from Cxp_Proveedores" _
       & " Where cod_Proveedor = " & vProveedor
Call OpenRecordSet(rs, strSQL)
  vProvCedJur = rs!CEDJUR
 
vPaso = True
Call sbCboAsignaDato(cboBanco, rs!Cuenta, True, rs!cod_banco)
vPaso = False

rs.Close

Call cboBanco_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxCtaBanco(vBanco As Long) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & vBanco
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta & ""
End If
rsX.Close
End Function


Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String _
                        , vLinea As Integer, Optional pDivisa As String = "COL", Optional pTipoCambio As Currency = 1)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad,cod_cc,cod_divisa,tipo_cambio) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidadOmision & "','','" & pDivisa & "'," & pTipoCambio & ")"
Call ConectionExecute(strSQL)

End Sub

Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Long, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date) As Long  'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,cod_unidad" _
       & ",cod_concepto,user_solicita,autoriza,fecha_autorizacion,user_autoriza,TIPO_BENEFICIARIO,tipo_cambio,cod_divisa) values(" & vBanco _
       & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','P','P','CC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S'" _
       & ",'" & vUnidadOmision & "','" & vConceptoOmision & "','" & glogon.Usuario & "','N',Null,Null" _
       & ", 3, " & CCur(txtBancoTipoCambio.Text) & ",'" & txtBancoDivisa.Text & "')"


Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "' and op = " & vOP & " and Estado = 'P'"
  Call OpenRecordSet(rsX, strSQL, 0)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function


Private Function fxTesoreria(vProveedor As Long, vMonto As Currency, Optional vTercero As String = "", Optional vMntDivReal As Currency) As Long
Dim strSQL As String, rs As New ADODB.Recordset, rsAnticipos As New ADODB.Recordset
Dim vTesoreria As Long, curAnticipos As Currency, vLinea As Integer
Dim vMontoBanco As Currency, vMontoProv As Currency
Dim vDifCuenta As String, vDifMonto As Currency

'If vMonto > 0 Then
'
Dim pTipoPago As String, pBancoId As Long, pCuentaBanco As String

Select Case cboTipoPago.Text
  Case "Transferencia"
     pTipoPago = "TE"
  Case "Cheque"
     pTipoPago = "CK"
End Select

pBancoId = cboBanco.ItemData(cboBanco.ListIndex)

If cboCuenta.ListCount > 0 Then
    pCuentaBanco = cboCuenta.ItemData(cboCuenta.ListIndex)
Else
    pCuentaBanco = ""
    pTipoPago = "CK"
End If

curAnticipos = fxMontoAncipos(vProveedor)
    
strSQL = "select P.CEDJUR,P.cod_proveedor,P.descripcion,P.cod_cuenta,P.cod_divisa,D.cod_cuenta as 'CtaDivDifIng',D.cod_cuenta_Gasto as 'CtaDivDifGst'" _
       & ",dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",P.COD_DIVISA,dbo.MyGetdate(),'V') as 'TipoCambio'" _
       & ",dbo.MyGetdate() as Fecha" _
       & " from  Cxp_Proveedores P inner join CntX_Divisas D on P.cod_divisa = D.cod_divisa" _
       & " and D.cod_contabilidad = " & GLOBALES.gEnlace _
       & " where P.cod_proveedor = " & vProveedor
    Call OpenRecordSet(rs, strSQL)
    
'    If txtBancoDivisa.Text = vDivisaFuncional Then
'       vMontoBanco = vMonto
'    Else
'       vMontoBanco = vMntDivReal
'    End If
    
    
    If CCur(txtBancoTipoCambio.Text) <> 1 Then
        vMontoBanco = vMonto / fxSys_Tipo_Cambio_Apl(CCur(txtBancoTipoCambio.Text))
    Else
        vMontoBanco = vMonto
    End If
       
    vTesoreria = fxMaestroTesoreria(pTipoPago, pBancoId, vMontoBanco, rs!CEDJUR _
                   , IIf((Len(vTercero) = 0), rs!Descripcion, vTercero), 0, "MODULO DE PROVEEDORES", 0, "PAGO AUTOMATICO" _
                   , pCuentaBanco, rs!fecha)
          
    
'    If txtBancoDivisa.Text = vDivisaFuncional Then
'       vMontoBanco = vMonto
'    Else
'       vMontoBanco = vMntDivReal * fxSys_Tipo_Cambio_Apl(CCur(txtBancoTipoCambio.Text))
'    End If
    
'    If CCur(txtBancoTipoCambio.Text) <> 1 Then
'        vMontoBanco = vMonto * fxSys_Tipo_Cambio_Apl(CCur(txtBancoTipoCambio.Text))
'    Else
'        vMontoBanco = vMonto
'    End If
    
    
    'En Divisa Funcional
    
    vMontoBanco = vMonto
    
    vLinea = 1
    Call sbCreaDetalle(vTesoreria, fxCtaBanco(pBancoId), vMontoBanco, "H", vLinea, txtBancoDivisa.Text, txtBancoTipoCambio.Text)
    
    
    Dim vTipoCambioPro As Currency
    If vDivisa = vDivisaFuncional Then
       vMontoProv = vMonto + curAnticipos
       vTipoCambioPro = rs!TipoCambio
    Else
       vMontoProv = vMonto + curAnticipos
       vTipoCambioPro = vMonto / vMntDivReal
    End If
    
    
    'Cancelación de Anticipos
    If curAnticipos > 0 Then
      With rsAnticipos
            strSQL = "select Cr.COD_CARGO, Cr.DESCRIPCION , Cr.COD_CUENTA ,  Pc.monto, Pc.COD_DIVISA" _
                    & "  from CXP_CARGOSPER Cp inner join CXP_ANTICIPOS Ca on Cp.COD_PROVEEDOR = Ca.COD_PROVEEDOR and Cp.COD_CARGO = Ca.COD_CARGO and Cp.ID = Ca.ID_CARGO" _
                    & "    inner join cxp_pagoProv Pf on Pf.COD_PROVEEDOR = Cp.COD_PROVEEDOR" _
                    & "    inner join CXP_PAGOPROVCARGOS Pc on Pf.COD_PROVEEDOR = Pc.COD_PROVEEDOR  and Pf.COD_FACTURA = Pc.COD_FACTURA and Pc.NPAGO = Pf.NPAGO and Pc.ID = Cp.ID" _
                    & "    inner join CXP_CARGOS Cr on Cp.COD_CARGO = Cr.COD_CARGO" _
                    & "  Where Cp.COD_PROVEEDOR = " & vProveedor & " and Pf.user_traslada = 'xBITxTesx'"
            .Open strSQL, glogon.Conection, adOpenStatic
            Do While Not .EOF
            
                vLinea = vLinea + 1
                Call sbCreaDetalle(vTesoreria, Trim(!cod_cuenta), !Monto, "H", vLinea, Trim(!COD_DIVISA), 1)
            
                .MoveNext
            Loop
            .Close
      End With
    End If 'CurAnticipos > 0
       
    
    
    
    
    
    vLinea = vLinea + 1
    Call sbCreaDetalle(vTesoreria, Trim(rs!cod_cuenta), vMontoProv, "D", vLinea, Trim(rs!COD_DIVISA), vTipoCambioPro)
       
       
    'Diferencial Cambiario
    vMonto = vMontoProv - (vMontoBanco + curAnticipos)
    If vMonto > 0 Then
        Call sbCreaDetalle(vTesoreria, Trim(rs!CtaDivDifIng), Abs(vMonto), "H", 3, vDivisaFuncional, 1)
    End If
    If vMonto < 0 Then
        Call sbCreaDetalle(vTesoreria, Trim(rs!CtaDivDifGst), Abs(vMonto), "D", 3, vDivisaFuncional, 1)
    End If
    
       
    rs.Close

' Else 'Monto a Girar > 0
'    vTesoreria = 0
'
'End If

fxTesoreria = vTesoreria


End Function


Private Function fxValidaPago() As Boolean
Dim vMensaje As String

vMensaje = ""

'Valida el uso de la divisa
'
'If cboDivisa.ItemData(cbodivisa.ListIndex) <> vDivisaFuncional And txtBancoDivisa.text <> vDivisaFuncional Then
'   vMensaje = vMensaje & " - La divisa de la " & vbCrLf
'End If



If Len(vMensaje) = 0 Then
   fxValidaPago = True
Else
   MsgBox vMensaje, vbExclamation
   fxValidaPago = True
End If

End Function

Private Sub chkUsuarios_Click()
If chkUsuarios.Value = vbChecked Then
   txtUsuario.Text = ""
   txtUsuario.Enabled = False
Else
   txtUsuario.Text = ""
   txtUsuario.Enabled = True
   txtUsuario.SetFocus
End If

Call sbBuscar

End Sub

Private Function fxMontoAncipos(pProveedor As Long) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(Sum(Pc.monto),0) as 'Cargos'" _
       & " from CXP_CARGOSPER Cp inner join CXP_ANTICIPOS Ca on Cp.COD_PROVEEDOR = Ca.COD_PROVEEDOR and Cp.COD_CARGO = Ca.COD_CARGO and Cp.ID = Ca.ID_CARGO" _
       & " inner join cxp_pagoProv Pf on Pf.COD_PROVEEDOR = Cp.COD_PROVEEDOR" _
       & " inner join CXP_PAGOPROVCARGOS Pc on Pf.COD_PROVEEDOR = Pc.COD_PROVEEDOR  and Pf.COD_FACTURA = Pc.COD_FACTURA and Pc.NPAGO = Pf.NPAGO and Pc.ID = Cp.ID" _
       & " Where Cp.COD_PROVEEDOR = " & pProveedor & " and Pf.user_traslada = 'xBITxTesx'"
Call OpenRecordSet(rs, strSQL)
If rs.BOF And rs.EOF Then
  fxMontoAncipos = 0
Else
  fxMontoAncipos = IIf(IsNull(rs!Cargos), 0, rs!Cargos)
End If
rs.Close


End Function



Private Sub sbReporte(vTesoreria As Long)
Dim vSQL As String, vSubTitulo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowExportBtn = True
 .WindowShowGroupTree = True
 .WindowShowPrintBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes Cuentas x Pagar"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 

  vSQL = "{CXP_PAGOPROV.TESORERIA} = " & vTesoreria
  vSubTitulo = txtProveedor.Text
 
  .Formulas(3) = "fxTitulo = 'BOLETA DE TRASLADO A BANCOS'"
  .Formulas(4) = "fxSubTitulo = 'No. Solicitud: " & vTesoreria & " -> " & UCase(vSubTitulo) & "'"
  .ReportFileName = SIFGlobal.fxPathReportes("CxP_ProgramacionListadoDetalle.rpt")
  .SelectionFormula = vSQL
 
 .PrintReport

End With
Me.MousePointer = vbDefault


Exit Sub

vError:

Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMontoNeto As Currency, curCarPerMonto As Currency, curCarPerPorc As Currency
Dim curMontoNetoDivisaReal As Currency, i  As Integer, lngTesoreria As Long
Dim vTipoCancelacion As String
Dim vCargosFecha As Date


' Restar el Saldo de la CxP del Proveedor, por el Neto. Pues todos los cargos lo Disminuyen cuando se crean
' Los cargos Flotantes, los que son por Montos reducen la CxP en el Acto,
' pero los Cargos Flotantes Porcentuales se Aplican hasta que se realiza el Pago (OJO)
'

If chkCargosVencimiento.Value = xtpChecked Then
    vCargosFecha = dtpVence.Value
Else
    vCargosFecha = fxFechaServidor
End If


'If vDivisa = vDivisaFuncional Then
'   If txtBancoDivisa.Text = vDivisaFuncional Or txtBancoDivisa.Text = cboDivisa.ItemData(cboDivisa.ListIndex) Then
'     'Pasa
'   Else
'        MsgBox "La divisa del Banco no corresponde a la divisa Funcional ni a la de la factura!", vbExclamation
'        Exit Sub
'   End If
'Else
'    If txtBancoDivisa.Text <> vDivisa Then
'      MsgBox "La divisa del Banco no corresponde a la divisa del Proveedor!", vbExclamation
'      Exit Sub
'    End If
'End If

'Tipo de Cancelacion
Select Case True
  Case optTipo.Item(0).Value 'Desembolso
    vTipoCancelacion = "D"
  Case optTipo.Item(1).Value 'Cargos
    vTipoCancelacion = "C"
  Case Else
    vTipoCancelacion = "D"
End Select


If vTipoCancelacion = "D" And CCur(txtPagoTotal.Text) = 0 Then
    MsgBox "EL monto a girar es 0, cambien el metodo de pago a Cerrar/Excluir", vbExclamation
    Exit Sub
End If

'Revisar que exista un check
vPaso = False
For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked Then vPaso = True
Next i
If Not vPaso Then
   MsgBox "Seleccione un Pago a realizar...", vbExclamation
   Exit Sub
End If

If chkPagoTercero.Value = vbChecked And Len(Trim(cbo.Text)) = 0 Then
  MsgBox "No se ha especificado el Pago a Terceros (Nombre)", vbExclamation
  Exit Sub
End If

If vTipoCancelacion = "C" And cboCargo.ListCount <= 0 Then
  MsgBox "La cancelación es por medio de registro de CARGOS y no existe ninguno disponible", vbExclamation
  Exit Sub
End If


If vTipoCancelacion = "D" And cboTipoPago.Text = "Transferencia" And cboCuenta.ListCount = 0 Then
  MsgBox "No es posible crear la solicitud de Transferencia, no se tiene cuenta bancaria!", vbExclamation
  Exit Sub
End If

Me.MousePointer = vbHourglass

On Error GoTo vError

'Registrando CARGOS DIRECTOS de CANCELACION
If vTipoCancelacion = "C" Then
    For i = 1 To lsw.ListItems.Count
     If lsw.ListItems.Item(i).Checked Then
        'Saca Monto Neto para Calculos
        curMontoNeto = CCur(lsw.ListItems.Item(i).SubItems(5))
'            spCxP_EjecucionPagos_RegistroCargos(@Proveedor int, @Factura varchar(30), @NPago int, @CodCargo varchar(10)
'                                        ,@Divisa varchar(10), @Monto dec(16,2), @TipoCambio dec(12,4),  @Usuario varchar(30))
            
            strSQL = "exec spCxP_EjecucionPagos_RegistroCargos " & lsw.ListItems.Item(i).SubItems(1) _
                   & ",'" & lsw.ListItems.Item(i).Text _
                   & "'," & lsw.ListItems.Item(i).SubItems(8) _
                   & ",'" & cboCargo.ItemData(cboCargo.ListIndex) _
                   & "','" & lsw.ListItems.Item(i).SubItems(10) _
                   & "'," & CCur(lsw.ListItems.Item(i).SubItems(5)) _
                   & "," & CCur(lsw.ListItems.Item(i).SubItems(12)) _
                   & ",'" & glogon.Usuario & "'"
            Call ConectionExecute(strSQL)
     
     End If
    Next i

End If 'vTipoCancelacion = "C"

For i = 1 To lsw.ListItems.Count
 If lsw.ListItems.Item(i).Checked Then
    'Saca Monto Neto para Calculos
    curMontoNeto = CCur(lsw.ListItems.Item(i).SubItems(2)) - CCur(lsw.ListItems.Item(i).SubItems(3))
    
    '    spCxP_EjecucionPagos_AplicaCargosFlotantes(@Proveedor int, @Factura varchar(30), @NPago int, @Disponible dec(16,2)
    '                                        , @Corte datetime, @AplicaCargos smallint ,  @Usuario varchar(30))
    strSQL = "exec spCxP_EjecucionPagos_AplicaCargosFlotantes " & lsw.ListItems.Item(i).SubItems(1) _
           & ",'" & lsw.ListItems.Item(i).Text _
           & "'," & lsw.ListItems.Item(i).SubItems(8) _
           & "," & curMontoNeto _
           & ",'" & Format(vCargosFecha, "yyyy/mm/dd") & " 23:59:59" _
           & "'," & CInt(lsw.ListItems.Item(i).SubItems(9)) _
           & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)

 End If 'Procesa Pago (Checked)
Next i


'Enviar A Tesoreria y Actualizar Seguimiento
'Trae datos de la vista de pagos, agrupados por proveedor

'1. Aplicar los Cargos Flotantes PORCENTUALES al Saldo del Proveedor / Pendientes de Pagar

strSQL = "exec spCxP_EjecucionPagos_ActualizaSaldosConCargosPorc"
Call ConectionExecute(strSQL)


'Tipo de Cancelación
lngTesoreria = 0
Select Case vTipoCancelacion
  Case "D"   'Desembolso
        curMontoNeto = 0
        curMontoNetoDivisaReal = 0
       
        
        strSQL = " select CEDJUR,cod_proveedor,sum(monto - cargos) as Neto,sum(Divisa_Real_Neto) as 'Divisa_Real_Neto'" _
               & " from vCXP_Pagos where cod_Proveedor = " & txtCodProv.Text _
               & " group by cod_proveedor,CEDJUR"
        Call OpenRecordSet(rs, strSQL)
        
        If Not rs.EOF And Not rs.BOF Then
            curMontoNeto = rs!Neto
            curMontoNetoDivisaReal = rs!Divisa_Real_Neto
        End If
        
        
        Do While Not rs.EOF
             lngTesoreria = fxTesoreria(rs!cod_Proveedor, curMontoNeto, IIf((chkPagoTercero.Value = vbChecked), cbo.Text, ""), rs!Divisa_Real_Neto)
         rs.MoveNext
        Loop
        rs.Close
        
        'Actualiza Indicadores
        If chkPagoTercero.Value = vbChecked Then
            strSQL = "update cxp_pagoprov set tesoreria = " & lngTesoreria _
                   & ",fecha_traslada = dbo.MyGetdate(),user_traslada = '" & glogon.Usuario _
                   & "',pago_tercero = '" & cbo.Text _
                   & "' where user_traslada = 'xBITxTesx'" _
                   & " and cod_proveedor = " & txtCodProv.Text
        Else
            strSQL = "update cxp_pagoprov set tesoreria = " & lngTesoreria _
                   & ",fecha_traslada = dbo.MyGetdate(),user_traslada = '" & glogon.Usuario _
                   & "',pago_tercero = '' where user_traslada = 'xBITxTesx'" _
                   & " and cod_proveedor = " & txtCodProv.Text
        End If
        Call ConectionExecute(strSQL)

        Me.MousePointer = vbDefault
        MsgBox "Pago Registrado en Tesoreria # Solicitud : " & lngTesoreria, vbInformation
    
    
'        'Cambio el 26/07/2012: Ahora el saldo del CxP se ve afectada por el spCxP_SincronizaTesoreria
'        'Actualiza el Ultimo Pago y Saldo del Proveedor
'        strSQL = "update cxp_proveedores set ultimo_pago = dbo.MyGetdate()" _
'               & ",SALDO = SALDO - " & curMontoNeto _
'               & ",SALDO_DIVISA_REAL = SALDO_DIVISA_REAL - " & curMontoNetoDivisaReal _
'               & " where cod_proveedor = " & txtCodProv
'        Call ConectionExecute(strSQL)
    
    Case "C" 'Cancelación por Cargos
        strSQL = "update cxp_pagoprov set tesoreria = 0, Tipo_Cancelacion= 'C', Tesoreria_Estado = 'E'" _
               & ",fecha_traslada = dbo.MyGetdate(),user_traslada = '" & glogon.Usuario _
               & "',pago_tercero = '',Tesoreria_Emision = dbo.MyGetdate()" _
               & " where user_traslada = 'xBITxTesx' And cod_proveedor = " & txtCodProv.Text
        Call ConectionExecute(strSQL)
    
        'TODO: Ver Asiento de los Cargos x Anticipos
        
        Me.MousePointer = vbDefault
        MsgBox "Cuenta por Pagar descontada vía Cargos!", vbInformation

End Select

'Actualiza el Detalle en Bancos con Notas de la Factura
strSQL = "exec spCxP_Tesoreria_Detalle_Update"
Call ConectionExecute(strSQL)

If lngTesoreria > 0 Then
  Call sbReporte(lngTesoreria)
End If


'Actualiza Pendientes
Call sbBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxEjeCarPerPorc(vMonto As Currency, vFecha As Date, vProveedor As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curCarPerMonto As Currency

'Saca Monto para Cargos Periodicos Porcentuales

curCarPerMonto = 0

strSQL = "select (valor / 100) as Porcentaje " _
       & " from cxp_cargosPer where cod_proveedor = " & vProveedor _
       & " and tipo = 'P' and vence >='" & Format(vFecha, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 curCarPerMonto = curCarPerMonto + (vMonto * rs!Porcentaje)
 rs.MoveNext
Loop
rs.Close

fxEjeCarPerPorc = curCarPerMonto

End Function


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curCarPerMonto As Currency
Dim curCarPerPorc As Currency, curCarPerMSaldo As Currency

Dim vCargosFecha As Date

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpiaDatos

If txtCodProv.Text = "" Then
    Me.MousePointer = vbDefault
    Exit Sub
End If

If chkCargosVencimiento.Value = xtpChecked Then
    vCargosFecha = dtpVence.Value
Else
    vCargosFecha = fxFechaServidor
End If

'Sacar los cargos periodicos por Montos para Deduccir hasta saldo 0
strSQL = "select dbo.fxCxP_CargoFlotanteSaldo(" & txtCodProv & ",'" & Format(vCargosFecha, "yyyy/mm/dd") & " 23:59:59') as 'SaldoCargoFlotante'"
Call OpenRecordSet(rs, strSQL)
    curCarPerMSaldo = rs!SaldoCargoFlotante
rs.Close

'Lista los Bancos
Call chkCambioBanco_Click

lsw.ListItems.Clear



strSQL = "exec spCxP_FacturasPendientesPago " & txtCodProv.Text & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) _
       & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & " 23:59:59',"

If chkUsuarios.Value = vbUnchecked And Len(txtUsuario.Text) > 0 Then
    strSQL = strSQL & "'" & txtUsuario.Text & "'"
Else
    strSQL = strSQL & "Null"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 
 'Saca los Cargos Periodicos Activos Porcentuales / Son Prioridad
 'Se Aplican al Monto a Desembolsas menos los cargos Directos
 curCarPerPorc = 0
 curCarPerMonto = 0
 
 If rs!apl_cargo_flotante = 1 Then
    curCarPerPorc = fxEjeCarPerPorc((rs!Monto - rs!Cargos), rs!Fecha_Vencimiento, rs!cod_Proveedor)
    'Descuenta Cargos Periodicos por Concepto de Montos (Saldos)
    If rs!Monto <= (curCarPerPorc + rs!Cargos + curCarPerMSaldo) Then
       curCarPerMonto = rs!Monto - (curCarPerPorc + rs!Cargos)
       curCarPerMSaldo = curCarPerMSaldo - curCarPerMonto
    Else
       curCarPerMonto = curCarPerMSaldo
       curCarPerMSaldo = 0
    End If
 End If
 
 Set itmX = lsw.ListItems.Add(, , rs!cod_Factura)
     itmX.SubItems(1) = rs!cod_Proveedor
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = Format(rs!Cargos, "Standard")
     itmX.SubItems(4) = Format(curCarPerPorc + curCarPerMonto, "Standard")
     itmX.SubItems(5) = Format(rs!Monto - (rs!Cargos + curCarPerPorc + curCarPerMonto), "Standard")
     itmX.SubItems(6) = Format(rs!Fecha_Vencimiento, "yyyy/mm/dd")
     itmX.SubItems(7) = rs!Proveedor
     itmX.SubItems(8) = rs!Npago
     itmX.SubItems(9) = rs!apl_cargo_flotante
     itmX.SubItems(10) = rs!COD_DIVISA
     itmX.SubItems(11) = Format(rs!Importe_divisa_real, "Standard")
     itmX.SubItems(12) = Format(rs!TIPO_CAMBIO, "###,###,##0.00###")
     itmX.SubItems(13) = Format((rs!Cargos + curCarPerPorc + curCarPerMonto) / rs!TIPO_CAMBIO, "Standard")
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Me.MousePointer = vbDefault
End Sub


Private Sub dtpVence_Change()
Call sbLimpiaDatos
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
      
  If Not IsNumeric(txtCodProv) Then
    txtCodProv.Tag = 0
  Else
    txtCodProv.Tag = txtCodProv.Text
  End If
    
  strSQL = "select Top 1 P.cod_proveedor,P.descripcion,rtrim(D.cod_divisa) as 'Cod_Divisa',P.CedJur,   rtrim(D.descripcion) as 'Divisa',P.cod_divisa" _
         & " from cxp_proveedores P inner join CntX_Divisas D on P.cod_divisa = D.cod_divisa and D.cod_contabilidad = " & GLOBALES.gEnlace _
         & " where cod_proveedor in(select cod_proveedor From cxp_PagoProv" _
         & " Where tesoreria Is Null and fecha_vencimiento <= '" _
         & Format(dtpVence.Value, "yyyy/mm/dd")
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " 23:59:59' and cod_proveedor > " & txtCodProv.Tag & " group by cod_proveedor) order by cod_proveedor asc"
    Else
       strSQL = strSQL & " 23:59:59' and cod_proveedor < " & txtCodProv.Tag & " group by cod_proveedor) order by cod_proveedor desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodProv = rs!cod_Proveedor
      txtProveedor = rs!Descripcion
      
      vProvCedJur = Trim(rs!CEDJUR & "")
      
      Call sbProveedorFusionado(rs!cod_Proveedor)
      
      vDivisa = Trim(rs!COD_DIVISA)
      Call sbCboAsignaDato(cboDivisa, rs!Divisa, True, rs!COD_DIVISA)
      
      Call chkCambioBanco_Click
      
      Call cboBanco_Click
      
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation


End Sub

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 30

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

With lsw.ColumnHeaders
  .Clear
  .Add , , "No Factura", 2400
  .Add , , "Prov.Id.", 1000, vbCenter
  .Add , , "Monto", 1300, vbRightJustify
  .Add , , "Car.Directo", 1300, vbRightJustify
  .Add , , "Car.Flotante", 1300, vbRightJustify
  .Add , , "Neto", 1300, vbRightJustify
  .Add , , "Vence", 1100, vbCenter
  .Add , , "Proveedor", 4400
  .Add , , "No Pago", 1000, vbCenter
  .Add , , "Apl.Car.Flot.", 1300, vbCenter
  .Add , , "Divisa", 900, vbCenter
  .Add , , "Importe Real", 1300, vbRightJustify
  .Add , , "Tipo Cambio", 1300, vbRightJustify
  .Add , , "Cargos Div.Real", 1300, vbRightJustify
End With

vUnidadOmision = fxCxPParametro("01")
vConceptoOmision = fxCxPParametro("02")

dtpVence.Value = fxFechaServidor

cboTipoPago.AddItem "Transferencia"
cboTipoPago.AddItem "Cheque"
cboTipoPago.Text = "Transferencia"

'Divisa Funcional
strSQL = "select COD_DIVISA" _
       & " From CNTX_DIVISAS" _
       & " Where DIVISA_LOCAL = 1 And COD_CONTABILIDAD = " & GLOBALES.gEnlace
Call OpenRecordSet(rs, strSQL)
   vDivisaFuncional = Trim(rs!COD_DIVISA)
rs.Close

 'Carga Divisas
 vPaso = True
 strSQL = "select rtrim(cod_divisa) as 'IdX',rtrim(descripcion) as ItmX" _
        & " from CntX_Divisas where cod_contabilidad = " & GLOBALES.gEnlace _
        & " order by divisa_local desc,cod_divisa"
 Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)
 vPaso = False
 
 
'Carga: Cargos
strSQL = "select rtrim(COD_CARGO) as 'IdX', RTRIM(descripcion) as 'ItmX'" _
       & " from CXP_CARGOS where ACTIVO = 1"
Call sbCbo_Llena_New(cboCargo, strSQL, False, True)
 
'Carga: Bancos
vPaso = True

strSQL = "exec spCxP_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
 
vPaso = False
 
 
Me.Width = 11088
Me.Height = 8784
 
Call optTipo_Click(0)
 
vInicia = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbProveedorFusionado(vCodPro As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select rtrim(P.descripcion) as Proveedor" _
       & " from cxp_fusiones F inner join cxp_proveedores P on F.cod_proveedor = P.cod_proveedor" _
       & " where F.cod_proveedor_fus = " & vCodPro
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   chkPagoTercero.Value = vbUnchecked
   chkPagoTercero.Enabled = True
   chkPagoTercero.Caption = "Pagar a Tercero"
Else
   chkPagoTercero.Value = vbChecked
   chkPagoTercero.Enabled = False
   cbo.Clear
   cbo.AddItem rs!Proveedor
   cbo.Text = rs!Proveedor
   chkPagoTercero.Caption = "Fusión"

End If
rs.Close

End Sub

Private Sub Form_Resize()
Dim pWidth As Long, pHeight As Long

pWidth = 11088
pHeight = 8784

If Me.Width > pWidth Then
    pWidth = Me.Width
End If

If Me.Height > pHeight Then
    pHeight = Me.Height
End If

lsw.Width = pWidth - 250
lsw.Height = pHeight - (lsw.top + gbResumen.Height + 450)

gbResumen.top = lsw.top + lsw.Height + 120
gbResumen.Width = lsw.Width


scTitulo.Width = lsw.Width




End Sub





Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

If Item.Checked Then
  txtTotalFactura.Text = CCur(txtTotalFactura.Text) + CCur(Item.SubItems(2))
  txtTotalFacturaDivisaReal.Text = CCur(txtTotalFacturaDivisaReal.Text) + CCur(Item.SubItems(11))
  
  txtTotalCargos.Text = CCur(txtTotalCargos.Text) + CCur(Item.SubItems(3)) + CCur(Item.SubItems(4))
  txtTotalCargosDivisaReal.Text = CCur(txtTotalCargosDivisaReal.Text) + CCur(Item.SubItems(13))
  
  txtPagoTotal.Text = CCur(txtPagoTotal.Text) + CCur(Item.SubItems(5))
  txtPagoTotalDivisaReal.Text = CCur(txtPagoTotalDivisaReal.Text) + (CCur(Item.SubItems(11)) - CCur(Item.SubItems(13)))
Else
  txtTotalFactura.Text = CCur(txtTotalFactura.Text) - CCur(Item.SubItems(2))
  txtTotalFacturaDivisaReal.Text = CCur(txtTotalFacturaDivisaReal.Text) - CCur(Item.SubItems(11))
  
  txtTotalCargos.Text = CCur(txtTotalCargos.Text) - (CCur(Item.SubItems(3)) + CCur(Item.SubItems(4)))
  txtTotalCargosDivisaReal.Text = CCur(txtTotalCargosDivisaReal.Text) - CCur(Item.SubItems(13))
  
  txtPagoTotal.Text = CCur(txtPagoTotal.Text) - CCur(Item.SubItems(5))
  txtPagoTotalDivisaReal.Text = CCur(txtPagoTotalDivisaReal.Text) - (CCur(Item.SubItems(11)) - CCur(Item.SubItems(13)))
End If

If CCur(txtPagoTotal.Text) > 0 Then
   If txtBancoDivisa.Text = vDivisaFuncional And cboDivisa.ItemData(cboDivisa.ListIndex) <> vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotal.Text
      txtDiferencial.Text = "0"
   End If

   If txtBancoDivisa.Text = vDivisaFuncional And cboDivisa.ItemData(cboDivisa.ListIndex) = vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotal.Text
      txtDiferencial.Text = "0"
   End If

   If txtBancoDivisa.Text = cboDivisa.ItemData(cboDivisa.ListIndex) And txtBancoDivisa.Text <> vDivisaFuncional Then
      txtPagoBancos.Text = txtPagoTotalDivisaReal.Text
      txtDiferencial.Text = CCur(txtPagoTotal.Text) - (CCur(txtPagoTotalDivisaReal.Text) * fxSys_Tipo_Cambio_Apl(CCur(txtBancoTipoCambio.Text)))
   End If
   
   If txtBancoDivisa.Text <> cboDivisa.ItemData(cboDivisa.ListIndex) Then
      txtPagoBancos.Text = CCur(txtPagoTotal.Text) / fxSys_Tipo_Cambio_Apl(CCur(txtBancoTipoCambio.Text))
      txtDiferencial.Text = "0"
    End If
Else

      txtPagoBancos.Text = "0"
      txtDiferencial.Text = "0"
   
End If

'Formato
txtPagoTotal.Text = Format(CCur(txtPagoTotal.Text), "Standard")
txtPagoTotalDivisaReal.Text = Format(CCur(txtPagoTotalDivisaReal.Text), "Standard")

txtTotalCargos.Text = Format(CCur(txtTotalCargos.Text), "Standard")
txtTotalCargosDivisaReal.Text = Format(CCur(txtTotalCargosDivisaReal.Text), "Standard")

txtTotalFactura.Text = Format(CCur(txtTotalFactura.Text), "Standard")
txtTotalFacturaDivisaReal.Text = Format(CCur(txtTotalFacturaDivisaReal.Text), "Standard")

txtPagoBancos.Text = Format(CCur(txtPagoBancos.Text), "Standard")
txtDiferencial.Text = Format(CCur(txtDiferencial.Text), "Standard")



End Sub

Private Sub optTipo_Click(Index As Integer)
Dim vValor As Boolean


Select Case True
  Case optTipo.Item(0).Value 'Cancelación por desembolso
    lblCargo.Visible = False
    cboCargo.Visible = False
    vValor = True
    
  Case optTipo.Item(1).Value 'Cancelación vía Cargo
    lblCargo.Visible = True
    cboCargo.Visible = True
    vValor = False
End Select


chkPagoTercero.Visible = vValor
cbo.Visible = vValor

chkCambioBanco.Visible = vValor
cboBanco.Visible = vValor

lblPagoTitulo(0).Visible = vValor
lblPagoTitulo(1).Visible = vValor

cboCuenta.Visible = vValor
cboTipoPago.Visible = vValor


End Sub



Private Function fxProveedor_Divisas_Desc(pProveedor As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select P.cod_proveedor,P.descripcion,rtrim(D.descripcion) as 'Divisa'" _
       & " from cxp_proveedores P inner join CntX_Divisas D on P.cod_divisa = D.cod_divisa and D.cod_contabilidad = " & GLOBALES.gEnlace _
       & " where P.cod_Proveedor = " & pProveedor
       
Call OpenRecordSet(rs, strSQL)
If rs.EOF Or rs.BOF Then
    strSQL = ""
Else
    strSQL = rs!Divisa
End If

fxProveedor_Divisas_Desc = strSQL


End Function


Private Sub txtCodProv_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus

On Error GoTo vError

',rtrim(D.cod_divisa) + ' - ' + rtrim(D.descripcion) as 'Divisa'

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion,CedJur" _
                       & " from cxp_proveedores"

  If chkPagosPendientes.Value = vbChecked Then
    gBusquedas.Filtro = " and cod_proveedor in(select cod_proveedor From cxp_PagoProv" _
                      & " Where tesoreria Is Null and fecha_vencimiento <= '" _
                      & Format(dtpVence.Value, "yyyy/mm/dd") & " 23:59:59' group by cod_proveedor)"
  Else
    gBusquedas.Filtro = ""
  End If
  
  frmBusquedas.Show vbModal
    
  If Not IsNumeric(gBusquedas.Resultado) Then Exit Sub
    
  
  txtCodProv = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
  vProvCedJur = gBusquedas.Resultado3
  
 Dim strSQL As String, rs As New ADODB.Recordset

    strSQL = "select rtrim(D.cod_Divisa) as 'Cod_Divisa', rtrim(D.descripcion) as 'Divisa'" _
           & " from cxp_proveedores P inner join CntX_Divisas D on P.cod_divisa = D.cod_divisa and D.cod_contabilidad = " & GLOBALES.gEnlace _
           & " where P.cod_Proveedor = " & txtCodProv
           
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Or Not rs.BOF Then
        vDivisa = rs!COD_DIVISA
        Call sbCboAsignaDato(cboDivisa, rs!Divisa, True, rs!COD_DIVISA)
    End If
    rs.Close
  
  
  If Len(Trim(txtCodProv)) > 0 Then Call sbProveedorFusionado(CInt(txtCodProv))

End If

vError:

End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVence.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "P.descripcion"
  gBusquedas.Orden = "P.descripcion"
  gBusquedas.Consulta = "select P.cod_proveedor,P.descripcion,rtrim(D.cod_divisa) + ' - ' + rtrim(D.descripcion) as 'Divisa'" _
                       & " from cxp_proveedores P inner join CntX_Divisas D on P.cod_divisa = D.cod_divisa and D.cod_contabilidad = " & GLOBALES.gEnlace _

  If chkPagosPendientes.Value = vbChecked Then
    gBusquedas.Filtro = " and P.cod_proveedor in(select cod_proveedor From cxp_PagoProv" _
                      & " Where tesoreria Is Null and fecha_vencimiento <= '" _
                      & Format(dtpVence.Value, "yyyy/mm/dd") & "' group by cod_proveedor)"
  Else
    gBusquedas.Filtro = ""
  End If
  
  frmBusquedas.Show vbModal
  txtCodProv = gBusquedas.Resultado
  txtProveedor = gBusquedas.Resultado2
  
  If Len(Trim(txtCodProv)) > 0 Then Call sbProveedorFusionado(CInt(txtCodProv))
  cboDivisa.Text = gBusquedas.Resultado3
  vDivisa = SIFGlobal.fxCodText(gBusquedas.Resultado3)
  
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Nombre,Descripcion from Usuarios"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtUsuario.Text = gBusquedas.Resultado
End If
End Sub
