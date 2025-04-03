VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDRetirosyLiquidaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Retiros y Liquidaciones"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4695
      Left            =   0
      TabIndex        =   30
      Top             =   3840
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   8281
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Datos del Retiro: "
      Item(0).ControlCount=   17
      Item(0).Control(0)=   "opt(0)"
      Item(0).Control(1)=   "opt(1)"
      Item(0).Control(2)=   "txtNotas"
      Item(0).Control(3)=   "txtMontoAplicar"
      Item(0).Control(4)=   "txtMontoGirar"
      Item(0).Control(5)=   "Label1(5)"
      Item(0).Control(6)=   "Label1(3)"
      Item(0).Control(7)=   "Label1(2)"
      Item(0).Control(8)=   "chkPagoTercero"
      Item(0).Control(9)=   "fraRetener"
      Item(0).Control(10)=   "cmdAplicar"
      Item(0).Control(11)=   "cboProceso"
      Item(0).Control(12)=   "fraDesembolso"
      Item(0).Control(13)=   "fraPagoTercero"
      Item(0).Control(14)=   "txtRebajos"
      Item(0).Control(15)=   "Label1(12)"
      Item(0).Control(16)=   "fraPlanDestino"
      Item(1).Caption =   "Otros Rebajos:"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "txtTotalRebajos"
      Item(1).Control(1)=   "Label2(16)"
      Item(1).Control(2)=   "vGrid"
      Begin VB.Frame fraDesembolso 
         Height          =   1455
         Left            =   600
         TabIndex        =   46
         Top             =   3240
         Width           =   6375
         Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
            Height          =   312
            Left            =   1200
            TabIndex        =   47
            Top             =   240
            Width           =   2292
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.ComboBox cboBanco 
            Height          =   312
            Left            =   1200
            TabIndex        =   48
            Top             =   600
            Width           =   4812
            _Version        =   1572864
            _ExtentX        =   8493
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
            Left            =   1200
            TabIndex        =   49
            Top             =   960
            Width           =   4812
            _Version        =   1572864
            _ExtentX        =   8493
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
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   975
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Banco"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   51
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Emitir"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame fraPlanDestino 
         Height          =   1455
         Left            =   840
         TabIndex        =   62
         Top             =   3480
         Visible         =   0   'False
         Width           =   6375
         Begin XtremeSuiteControls.ComboBox cboPlanDestino 
            Height          =   315
            Left            =   1080
            TabIndex        =   63
            Top             =   960
            Width           =   5175
            _Version        =   1572864
            _ExtentX        =   9128
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
         Begin VB.Label Label9 
            Caption         =   "Plan"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Acreditar el monto del Retiros o Liquidación a un Contrato de Fondos existente o nuevo?"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   6135
         End
      End
      Begin VB.Frame fraRetener 
         Height          =   1455
         Left            =   360
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   6375
         Begin XtremeSuiteControls.ComboBox cboRetencion 
            Height          =   312
            Left            =   1080
            TabIndex        =   41
            Top             =   1080
            Width           =   5172
            _Version        =   1572864
            _ExtentX        =   9128
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
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Caption         =   $"frmFNDRetirosyLiquidaciones.frx":0000
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   735
            Index           =   0
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   6135
         End
         Begin VB.Label Label9 
            Caption         =   "Concepto"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   42
            Top             =   1080
            Width           =   855
         End
      End
      Begin XtremeSuiteControls.PushButton opt 
         Height          =   492
         Index           =   0
         Left            =   6240
         TabIndex        =   31
         Top             =   720
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Retiro"
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
      Begin XtremeSuiteControls.PushButton opt 
         Height          =   492
         Index           =   1
         Left            =   7560
         TabIndex        =   32
         Top             =   720
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Liquidación"
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   2880
         TabIndex        =   33
         Top             =   1560
         Width           =   6012
         _Version        =   1572864
         _ExtentX        =   10604
         _ExtentY        =   1397
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
      Begin XtremeSuiteControls.FlatEdit txtMontoAplicar 
         Height          =   312
         Left            =   2880
         TabIndex        =   34
         Top             =   480
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoGirar 
         Height          =   312
         Left            =   2880
         TabIndex        =   35
         Top             =   1200
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.CheckBox chkPagoTercero 
         Height          =   372
         Left            =   2760
         TabIndex        =   39
         Top             =   2760
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Pago a Tercero?"
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
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   732
         Left            =   7320
         TabIndex        =   44
         Top             =   3240
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         Appearance      =   17
         Picture         =   "frmFNDRetirosyLiquidaciones.frx":00CF
      End
      Begin XtremeSuiteControls.ComboBox cboProceso 
         Height          =   312
         Left            =   360
         TabIndex        =   45
         Top             =   2760
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.FlatEdit txtRebajos 
         Height          =   312
         Left            =   2880
         TabIndex        =   57
         Top             =   840
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
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
      Begin VB.Frame fraPagoTercero 
         Height          =   1455
         Left            =   360
         TabIndex        =   53
         Top             =   3120
         Width           =   6375
         Begin VB.ComboBox cboPagoTercero 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   960
            Width           =   6135
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Indique a favor de Quién se emitirá el Cheque de las personas/entidades vinculadas a la persona"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   120
            Width           =   6135
         End
         Begin VB.Label Label9 
            Caption         =   "Autorizados Giro a Terceros.:"
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
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   2775
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3732
         Left            =   -69760
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   8892
         _Version        =   524288
         _ExtentX        =   15684
         _ExtentY        =   6583
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   486
         ScrollBars      =   2
         SpreadDesigner  =   "frmFNDRetirosyLiquidaciones.frx":08A7
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalRebajos 
         Height          =   312
         Left            =   -62560
         TabIndex        =   60
         Top             =   4200
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
      Begin VB.Label Label2 
         Caption         =   "Total "
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
         Index           =   16
         Left            =   -63520
         TabIndex        =   61
         Top             =   4200
         Visible         =   0   'False
         Width           =   4092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Otros Rebajos"
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
         TabIndex        =   58
         Top             =   840
         Width           =   2292
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Aplicar"
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
         Index           =   2
         Left            =   480
         TabIndex        =   38
         Top             =   480
         Width           =   2292
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto a Girar"
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
         Index           =   3
         Left            =   480
         TabIndex        =   37
         Top             =   1200
         Width           =   2292
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas del Retiro"
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
         Index           =   5
         Left            =   480
         TabIndex        =   36
         Top             =   1560
         Width           =   2292
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   9
      Top             =   8496
      Width           =   9204
      _ExtentX        =   16245
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Nivel de Autorización ..:"
            TextSave        =   "Nivel de Autorización ..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Rangos autorizados para Retiros / Liquidaciones"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
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
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   6960
      Top             =   360
   End
   Begin XtremeSuiteControls.FlatEdit txtAportes 
      Height          =   312
      Left            =   2880
      TabIndex        =   10
      Top             =   960
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtRendimientos 
      Height          =   312
      Left            =   2880
      TabIndex        =   11
      Top             =   1320
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtFechaCorte 
      Height          =   312
      Left            =   2880
      TabIndex        =   12
      Top             =   1800
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtMontoRetenido 
      Height          =   312
      Left            =   2880
      TabIndex        =   13
      Top             =   2160
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtMultaRetiro 
      Height          =   312
      Left            =   2880
      TabIndex        =   14
      Top             =   2520
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtRendPendientes 
      Height          =   312
      Left            =   2880
      TabIndex        =   15
      Top             =   3000
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtMontoDisponible 
      Height          =   312
      Left            =   2880
      TabIndex        =   16
      Top             =   3360
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.FlatEdit txtRetiros_Acumulados 
      Height          =   312
      Left            =   7080
      TabIndex        =   17
      Top             =   1800
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
   Begin XtremeSuiteControls.FlatEdit txtRenta_NoGravable 
      Height          =   312
      Left            =   7080
      TabIndex        =   19
      Top             =   2160
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
   Begin XtremeSuiteControls.FlatEdit txtRenta_MontoGravable 
      Height          =   312
      Left            =   7080
      TabIndex        =   21
      Top             =   2520
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
   Begin XtremeSuiteControls.FlatEdit txtRenta_Monto 
      Height          =   312
      Left            =   7080
      TabIndex        =   23
      Top             =   3360
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
   Begin XtremeSuiteControls.FlatEdit txtRendRetirar 
      Height          =   312
      Left            =   7080
      TabIndex        =   27
      Top             =   3000
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
   Begin XtremeSuiteControls.FlatEdit txtRenta_Porcentaje 
      Height          =   312
      Left            =   7080
      TabIndex        =   25
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
   Begin XtremeSuiteControls.CheckBox chkRentaGlobal 
      Height          =   252
      Left            =   6720
      TabIndex        =   29
      Top             =   960
      Width           =   2172
      _Version        =   1572864
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aplica Renta Global?  "
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
      Enabled         =   0   'False
      TextAlignment   =   1
      Appearance      =   16
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rendimiento a Retirar"
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
      Index           =   11
      Left            =   5040
      TabIndex        =   28
      Top             =   3000
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(%) Impuesto s/Rend."
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
      Left            =   5040
      TabIndex        =   26
      Top             =   1320
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(-) Imp. s/Rendimientos"
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
      Left            =   5040
      TabIndex        =   24
      Top             =   3360
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(i) Monto Gravable"
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
      Left            =   5040
      TabIndex        =   22
      Top             =   2520
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(i) Base No Gravable"
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
      Index           =   7
      Left            =   5040
      TabIndex        =   20
      Top             =   2160
      Width           =   2292
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(i) Retiros Acumulados"
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
      Index           =   6
      Left            =   5040
      TabIndex        =   18
      Top             =   1800
      Width           =   2292
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "[Cliente]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   7692
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(+) Rendimientos Pendientes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   7
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label lblTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aqui Tipo de Gestion (Contrato/Plan/Operadora)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   7776
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto Disponible"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Rendimientos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Aportes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(-) Multa por Retiro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(-) Retenido x Garantía"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "frmFNDRetirosyLiquidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFechaInicio As String, vOperadora As Long
Dim vPlan As String, vContrato As Long, vPaso As Boolean
Dim strCedula As String, strCliente As String
Dim mAutorizacion As Boolean, mAutoInicio As Currency, mAutoCorte As Currency
Dim mFechaServer  As Date, vPermiteLiquidar As Boolean


Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & gFondos.Cedula & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:
End Sub

Private Sub cboProceso_Click()
If vPaso Then Exit Sub

fraDesembolso.Visible = False
fraRetener.Visible = False
fraPagoTercero.Visible = False
fraPlanDestino.Visible = False

fraDesembolso.top = cboProceso.top + cboProceso.Height + 45
fraDesembolso.Left = cboProceso.Left

fraRetener.top = fraDesembolso.top
fraRetener.Left = fraDesembolso.Left

fraPagoTercero.top = fraDesembolso.top
fraPagoTercero.Left = fraDesembolso.Left
 
fraPlanDestino.top = fraDesembolso.top
fraPlanDestino.Left = fraDesembolso.Left
 

Select Case Mid(cboProceso.Text, 1, 1)

Case "D"
  chkPagoTercero.Visible = True
  
  If chkPagoTercero.Value = vbChecked Then
       fraPagoTercero.Visible = True
  Else
       fraDesembolso.Visible = True
  End If

Case "R"
   chkPagoTercero.Visible = False
   fraRetener.Visible = True

Case "F"
   chkPagoTercero.Visible = False
   fraPlanDestino.Visible = True

End Select

End Sub


Private Sub chkPagoTercero_Click()
If vPaso Then Exit Sub
Call cboProceso_Click
End Sub

Private Sub CmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vLiquidacion As Long, vTipo As String, vMonto As Currency, vTesoreria As Long
Dim vPTId As String, vPTNombre As String, vPTTipo As String, i As Long
Dim vPrimero As Boolean

On Error GoTo vError

If mAutorizacion Then
  If CCur(txtMontoAplicar.Text) < mAutoInicio Then
    MsgBox "Se Requiere Autorización -> El monto a retirar está fuera de su rango autorizado!", vbExclamation
    Exit Sub
  End If
  
  If CCur(txtMontoAplicar.Text) > mAutoCorte Then
    MsgBox "Se Requiere Autorización -> El monto a retirar está fuera de su rango autorizado!", vbExclamation
    Exit Sub
  End If
End If


If CCur(txtMontoAplicar.Text) > CCur(txtMontoDisponible.Text) Then
  MsgBox "El Monto del Retiro no puede ser mayor al disponible de Contrato!", vbExclamation
  Exit Sub
End If

If chkPagoTercero.Enabled And chkPagoTercero.Value = vbChecked And Mid(cboProceso.Text, 1, 1) = "D" Then
    If cboPagoTercero.ListCount = 0 Then
            MsgBox "Se activó el pago a terceros pero no existen personas autorizadas al giro!", vbExclamation
            Exit Sub
    Else
        If cboPagoTercero.Text = "" Then
            MsgBox "Se activó el pago a terceros y no se especificó al Beneficiario?", vbExclamation
            Exit Sub
        End If
    End If
End If

'Retiros en Cajas> Validacion
If fxTipoDocumento(cboTipoDocumento.Text) = "RC" And Mid(cboProceso.Text, 1, 1) = "D" Then
  strSQL = "select Valor from CAJAS_PARAMETROS  where cod_parametro = '15'"
  Call OpenRecordSet(rs, strSQL)
  
  If IsNumeric(rs!Valor) Then
        If rs!Valor < CCur(txtMontoAplicar.Text) Then
            MsgBox "- El Monto Máximo para Retiros de Efectivos en Cajas es de " _
                   & Format(rs!Valor, "Standard") & ", Informe a su Administrador!", vbInformation
            Exit Sub
        End If
  Else
    MsgBox "- No se ha configurado el Monto para Retiros de Efectivos en Cajas, Informe a su Administrador!", vbInformation
    Exit Sub
  End If
End If


'Validaciones Adicionales (CDPs y Otros)
If opt.Item(0).Checked Then
    vTipo = "R"
Else
    vTipo = "L"
End If
strSQL = "select dbo.fxFndRetiroValida_Notas(" & vOperadora & ", '" & vPlan & "', " & vContrato _
       & ", '" & vTipo & "', '" & glogon.Usuario & "') as 'Resultado'"
Call OpenRecordSet(rs, strSQL)
If Len(rs!Resultado) > 0 Then
    MsgBox rs!Resultado, vbExclamation
    Exit Sub
End If



Select Case True
  Case opt.Item(0).Checked 'Retiro
    vTipo = "R"
    strSQL = MsgBox("Confirma el Retiro Parcial?", vbExclamation + vbYesNo)

  Case opt.Item(1).Checked  'Liquidacion
    vTipo = "L"
    strSQL = MsgBox("Confirma La Liquidación?", vbExclamation + vbYesNo)
End Select

If strSQL = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

'    spFnd_Liquidacion_Rebajos(@Usuario varchar(30), @Contrato int, @Plan varchar(10), @Concepto varchar(10)
'                , @Documento varchar(60), @Detalle varchar(100), @Monto dec(14,2), @TipoCambio dec(10,4) = 1
'                , @Inicializa smallint = 0)

Dim pConcepto As String, pDocumento As String, pDetalle As String, pMonto As Currency, pTipoCambio As Currency
If CCur(txtRebajos.Text) > 0 Then
    vPrimero = True
    strSQL = ""
    For i = 1 To vGrid.MaxRows
        vGrid.Row = i
        vGrid.Col = 5
        If IsNumeric(vGrid.Text) Then
            If CCur(vGrid.Text) > 0 Then
                vGrid.Col = 1
                pConcepto = vGrid.Text
                vGrid.Col = 3
                pDocumento = vGrid.Text
                vGrid.Col = 4
                pDetalle = vGrid.Text
                vGrid.Col = 5
                pMonto = CCur(vGrid.Text)
                
                strSQL = strSQL & Space(10) & "exec spFnd_Liquidacion_Rebajos '" & glogon.Usuario _
                        & "'," & vContrato & ",'" & vPlan & "','" & pConcepto & "','" & pDocumento _
                        & "','" & pDetalle & "'," & pMonto & "," & pTipoCambio _
                        & "," & IIf((vPrimero = True), 1, 0)
                vPrimero = False
            End If
        End If 'Is numeric
    Next i

    'Ejecuta el Lote
    Call ConectionExecute(strSQL)
End If

Dim pRetCodigo As String, pBancoId As String, pCuentaBancaria As String

pRetCodigo = ""
pBancoId = "0"
pCuentaBancaria = ""

If cboBanco.ListCount > 0 Then
    pBancoId = cboBanco.ItemData(cboBanco.ListIndex)
End If

If cboCuenta.ListCount > 0 Then
    pCuentaBancaria = cboCuenta.ItemData(cboCuenta.ListIndex)
End If

Select Case Mid(cboProceso.Text, 1, 1)
    Case "D" 'Desembolso

    Case "R" 'Retencion
      If cboRetencion.ListCount > 0 Then
          pRetCodigo = cboRetencion.ItemData(cboRetencion.ListIndex)
      End If
    
    Case "F" 'Fondo Destino
      If cboPlanDestino.ListCount > 0 Then
          pRetCodigo = cboPlanDestino.ItemData(cboPlanDestino.ListIndex)
      End If
End Select

'Valida, aplica y envia a tesoreria (si es que aplica)
strSQL = "exec spFndRetLiqProceso " & vOperadora & ", '" & vPlan & "', " & vContrato & ",'" & strCedula & "', " & CCur(txtMontoAplicar.Text) _
       & ", '" & vTipo & "', '" & txtNotas.Text & "', '" & glogon.Usuario & "', '" & GLOBALES.gOficinaTitular _
       & "', '" & Mid(cboProceso.Text, 1, 1) & "', '" & pRetCodigo _
       & "', " & pBancoId & ",'" & fxTipoDocumento(cboTipoDocumento.Text) & "','" & pCuentaBancaria & "', '" & App.ProductName & "'"

 

If chkPagoTercero.Enabled And chkPagoTercero.Value = vbChecked And Mid(cboProceso.Text, 1, 1) = "D" Then
   vPTTipo = Mid(cboPagoTercero.Text, 1, 1)
   vPTId = Mid(SIFGlobal.fxCodText(cboPagoTercero.Text, "."), 3, 30)
   vPTNombre = Mid(cboPagoTercero.Text, Len(vPTId) + 4, 100)
   
   strSQL = strSQL & ",null,1,'" & vPTTipo & "','" & vPTId & "','" & vPTNombre & "'," & CCur(txtRebajos.Text)
Else
   strSQL = strSQL & ",null,0,'N','',''," & CCur(txtRebajos.Text)
End If

Call OpenRecordSet(rs, strSQL)
  
If glogon.error Then
    'Desplega Error desde el Login
    MsgBox "Ocurrió un problema con la liquidación, verifique!", vbExclamation
    Exit Sub
Else
  vLiquidacion = rs!Num_liq
  vMonto = rs!MontoGiro
  vTesoreria = rs!tesoreria
End If
rs.Close


Call sbTrazabilidad_Inserta("08", CStr(vLiquidacion), CStr(vLiquidacion))

 With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "FONDO DE INVERSION"
  
  .Connect = glogon.ConectRPT
  
  .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionBoleta.rpt")
  .SelectionFormula = "{FND_LIQUIDACION.CONSEC} =" & vLiquidacion
  .Formulas(0) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  .Formulas(1) = "fxCodigoBarras= '*" & vLiquidacion & "*'"
  
  .SubreportToChange = "sbAsiento"
  If GLOBALES.SysDocVersion = 1 Then
     .StoredProcParam(0) = "LI"
  Else
     .StoredProcParam(0) = "FLIQ"
  End If
  
  .StoredProcParam(1) = vLiquidacion
  .StoredProcParam(2) = 1

  .PrintReport
 End With

Me.MousePointer = vbDefault
 
Call sbSIFRegistraTags(str(vLiquidacion), "S10", "FND LIQ", "0", "FLQ")


Select Case Mid(cboProceso.Text, 1, 1)
    Case "R"
         MsgBox "Liquidación o Retiro aplicado satisfactoriamente [Se Aplica Retención al Desembolso]", vbInformation
    Case "D"
        If vTesoreria = 0 Then
           MsgBox "Liquidación o Retiro aplicado satisfactoriamente [Activado para Traslado a Tesorería]", vbInformation
        Else
           MsgBox "Liquidación o Retiro aplicado y enviado a Tesoreria [No. de solictud en Tesoreria..: " & vTesoreria & " ]", vbInformation
        End If
    Case "F"
         MsgBox "Liquidación o Retiro aplicado satisfactoriamente [Acredito a Fondo Indicado]", vbInformation
End Select

Unload Me

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbRentaGlobal()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curRndRetiro As Currency, curRetiro As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

curRndRetiro = CCur(txtRendimientos.Text) + CCur(txtRendPendientes.Text)
curRetiro = CCur(txtMontoAplicar.Text)

If curRetiro < curRndRetiro Then
   curRndRetiro = curRetiro
End If

'Consulta Renta Global
strSQL = "exec spFnd_Renta_Global '" & strCedula & "', '" & Format(mFechaServer, "yyyy/mm/dd hh:mm") _
       & "'," & curRndRetiro & ",'" & vPlan & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   chkRentaGlobal.Value = rs!RG_Aplica
   txtRenta_Porcentaje.Text = Format(rs!RG_Porcentaje, "Standard")
   txtRetiros_Acumulados.Text = Format(rs!Retiro_Acumulado, "Standard")
   txtRenta_NoGravable.Text = Format(rs!RG_MntNoGravable, "Standard")
   txtRenta_MontoGravable.Text = Format(rs!Retiro_Gravable, "Standard")
   txtRendRetirar.Text = Format(curRndRetiro, "Standard")
   
   If chkRentaGlobal.Value = xtpChecked Then
        txtRenta_Monto.Text = Format(rs!ISR_MONTO, "Standard")
   Else
        txtRenta_Monto.Text = Format(CCur(txtRendPendientes.Text) * (rs!RG_Porcentaje / 100), "Standard")
   End If
   
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 

End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


vModulo = 18 'Fondo de Inversion

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture



vOperadora = gFondos.Operadora
vPlan = gFondos.Plan
vContrato = gFondos.Contrato

lblTipo.Caption = "Plan: " & Trim(gFondos.Plan) & Space(10) & " Contrato: " & gFondos.Contrato

mFechaServer = fxFechaServidor


vPermiteLiquidar = True

vPaso = True
    cboProceso.Clear
    cboProceso.AddItem "Desembolsar"
    cboProceso.AddItem "Retener"
    cboProceso.AddItem "Fondo"
vPaso = False
cboProceso.Text = "Desembolsar"


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("RC")
cboTipoDocumento.AddItem fxTipoDocumento("FD")
cboTipoDocumento.Text = fxTipoDocumento("TE")

If fxFndParametro("01") = "S" Then
   mAutorizacion = True
   strSQL = "exec spFndSeguridadRango " & gFondos.Operadora & ",'" & gFondos.Plan & "','" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
        mAutoInicio = rs!Inicio
        mAutoCorte = rs!Corte
   rs.Close
   
   StatusBarX.Panels.Item(2).Text = Format(mAutoInicio, "Standard") & " -> " & Format(mAutoCorte, "Standard")

Else
   mAutorizacion = False
   StatusBarX.Panels.Item(2).Text = "N/A"
End If

StatusBarX.Panels.Item(3).Text = "Autorización.: " & IIf(mAutorizacion, "Sí", "No")


vPaso = True
    
    strSQL = "select B.id_banco as 'IdX',dbo.fxSys_Cuenta_Bancos_Desc(B.id_Banco) as 'ItmX'" _
           & " from tes_banco_asg T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
           & " where T.nombre = '" & glogon.Usuario & "'"
    Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
       
vPaso = False


strSQL = "select rtrim(RETENCION_CODIGO) as 'IdX', RTRIM(DESCRIPCION) as 'ItmX'" _
       & " From FND_RETENCION_CONCEPTOS  Where ACTIVO = 1" _
       & " and dbo.fxFnd_Seguridad_Acceso_Concepto('" & glogon.Usuario & "', RETENCION_CODIGO) = 1"
Call sbCbo_Llena_New(cboRetencion, strSQL, False, True)

strSQL = "exec spFndRetirosPlanesDestinos_List " & vOperadora & ", '" & vPlan & "', " & vContrato
Call sbCbo_Llena_New(cboPlanDestino, strSQL, False, True)


strSQL = "select CODIGO, DESCRIPCION, '' AS DOCUMENTO, '' AS DETALLE, 0 AS 'MONTO'" _
       & " From vFnd_Rebajos_Aplicables_List Where dbo.fxFnd_Seguridad_Acceso_Concepto('" & glogon.Usuario & "', CODIGO) = 1"
Call sbCargaGrid(vGrid, 5, strSQL, True)
If vGrid.MaxRows > 0 Then
    vGrid.MaxRows = vGrid.MaxRows - 1
End If

txtRebajos.Text = Format(0, "Standard")
txtTotalRebajos.Text = Format(0, "Standard")

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select isnull(sif_liquida,0) as 'PermiteLiquidar' from fnd_Planes where cod_Operadora = " & vOperadora & " and Cod_Plan = '" & vPlan & "'"
Call OpenRecordSet(rs, strSQL)
If rs!PermiteLiquidar Then
    vPermiteLiquidar = True
    
    strSQL = "exec spFndRetLiqConsulta " & vOperadora & ",'" & vPlan & "'," & vContrato
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
       strCedula = Trim(rs!Cedula)
       strCliente = rs!Nombre
       
       lblCliente.Caption = "Id: " & strCedula & Space(10) & strCliente
    
       
       txtAportes.Text = Format(rs!Aportes, "Standard")
       txtRendimientos.Text = Format(rs!Rendimiento, "Standard")
       txtMontoAplicar.Text = Format(rs!Aportes + rs!Rendimiento + rs!Rend_Pendiente, "Standard")
       txtFechaCorte.Text = Format(rs!fecha_corte, "yyyy/mm/dd")
       
       If rs!Plazo <= 900 And rs!plazo_Tipo = "M" Then
         txtFechaCorte.Tag = "N"
       Else
         txtFechaCorte.Tag = "S"
       End If
       
       If IsNull(rs!TIPO_PAGO) = False Then cboTipoDocumento = fxgFNDTipoPago("C", rs!TIPO_PAGO)
       txtMultaRetiro.Text = Format(rs!Multa, "Standard")
       txtRendPendientes.Text = Format(rs!Rend_Pendiente, "Standard")  'fxRendimientoHoy(fxFechaServidor)
       txtMontoRetenido.Text = Format(rs!SaldoEnGarantia, "Standard")  'Format(fxMontoRetenido(vOperadora, vPlan, vContrato), "Standard")
       
       chkPagoTercero.Value = vbUnchecked
       cboPagoTercero.Clear
       
       If rs!GiroTerceros = 1 Then
          strSQL = "exec spFndPersonaBeneficiarios " & vOperadora & ",'" & vPlan & "'," & vContrato & ",'" & strCedula & "'"
          rs.Close
          Call OpenRecordSet(rs, strSQL)
          Do While Not rs.EOF
            cboPagoTercero.AddItem rs!Tipo & "/" & Trim(rs!COD_BENEFICIARIO) & "." & rs!Nombre
            rs.MoveNext
          Loop
          
          chkPagoTercero.Enabled = True
       Else
          chkPagoTercero.Enabled = False
            
       End If
       
    End If
    rs.Close
    
    Call txtMontoAplicar_LostFocus
    
    cboTipoDocumento.Text = "Transferencia"

Else
    vPermiteLiquidar = False
End If

End Sub




Private Sub opt_Click(Index As Integer)


opt.Item(0).Checked = False
opt.Item(1).Checked = False

opt.Item(Index).Checked = True

'Disponible para Aplicar
txtMontoDisponible.Text = Format((CCur(txtAportes) + CCur(txtRendimientos) + CCur(txtRendPendientes)) - CCur(txtMontoRetenido.Text), "Standard")


Select Case Index
  
  Case 0 'Retiros
     txtMontoAplicar.Enabled = True
     txtMontoAplicar.BackColor = vbWhite
     txtMontoAplicar.Text = 0
     txtMontoGirar.Text = 0
     txtMultaRetiro.Text = 0
     txtMontoAplicar.SetFocus
  
  Case 1 'Liquidacion
     txtMultaRetiro = Format(fxgFNDCodigoMulta(vOperadora, vPlan, vContrato, (CCur(txtAportes) + CCur(txtRendimientos) + CCur(txtRendPendientes))), "Standard")
     txtMontoAplicar.Enabled = False
     txtMontoAplicar.BackColor = RGB(187, 215, 247)
     txtMontoAplicar.Text = Format((CCur(txtAportes) + CCur(txtRendimientos) + CCur(txtRendPendientes)) - CCur(txtMontoRetenido), "Standard")
     txtNotas.SetFocus
End Select

Call txtMontoAplicar_LostFocus

txtMontoGirar.BackColor = txtMontoAplicar.BackColor

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then

    Call vGrid_LeaveCell(1, 1, 1, 1, False)
    txtMontoAplicar_LostFocus
End If

End Sub

Private Sub Timer1_Timer()
   Timer1.Interval = 0
   Timer1.Enabled = False
   
   If Not vPermiteLiquidar Then
      MsgBox "Este Plan no se permite Liquidar por este medio!", vbInformation
      Unload Me
   Else
       Call opt_Click(0)
   End If
   
End Sub

Private Sub txtMontoAplicar_GotFocus()
On Error GoTo vError
txtMontoAplicar = CCur(txtMontoAplicar)
vError:
End Sub

Private Sub txtMontoAplicar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtMontoAplicar_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
    txtRenta_Monto.Text = "????"
    txtMultaRetiro.Text = "????"
    txtMontoGirar.Text = "????"
vError:
End Sub

Private Sub txtMontoAplicar_LostFocus()
On Error GoTo vError
    txtMontoAplicar.Text = Format(txtMontoAplicar, "Standard")
    txtMultaRetiro.Text = Format(fxgFNDCodigoMulta(vOperadora, vPlan, vContrato, CCur(txtMontoAplicar.Text)), "Standard")
    
    Call sbRentaGlobal
    
    txtMontoGirar.Text = Format(CCur(txtMontoAplicar) - (CCur(txtMultaRetiro) + CCur(txtRenta_Monto.Text) + CCur(txtRebajos.Text)), "Standard")
vError:
End Sub

Private Sub txtMontoRetenido_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 48 To 57, 8, 46
 Case vbKeyReturn
   
 Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim curMonto As Currency, i As Integer

On Error GoTo vError

curMonto = 0
With vGrid
  For i = 1 To .MaxRows
     .Row = i
     .Col = 5
     curMonto = curMonto + CCur(.Text)
  Next i
End With

txtTotalRebajos.Text = Format(curMonto, "Standard")
txtRebajos.Text = txtTotalRebajos.Text

Exit Sub

vError:
txtTotalRebajos.Text = Format(curMonto, "Standard")
txtRebajos.Text = txtTotalRebajos.Text
End Sub
