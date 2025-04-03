VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmIVR_Proc_Liquidacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Liquidación de Inversiones"
   ClientHeight    =   8355
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswAd 
      Height          =   1572
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   2773
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
      Left            =   4080
      Top             =   120
   End
   Begin XtremeSuiteControls.FlatEdit txtInversionId 
      Height          =   492
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAdquisicion 
      Height          =   372
      Index           =   0
      Left            =   9120
      TabIndex        =   7
      Top             =   5460
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nuevo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtAd_Pendiente 
      Height          =   312
      Left            =   5760
      TabIndex        =   8
      Top             =   7920
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAdquisicion 
      Height          =   372
      Index           =   1
      Left            =   10680
      TabIndex        =   9
      Top             =   5460
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      Picture         =   "frmIVR_Proc_Liquidacion.frx":0000
   End
   Begin XtremeSuiteControls.FlatEdit txtAd_Registrado 
      Height          =   312
      Left            =   5760
      TabIndex        =   11
      Top             =   7560
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTransacId 
      Height          =   492
      Left            =   6240
      TabIndex        =   14
      Top             =   120
      Width           =   2532
      _Version        =   1441793
      _ExtentX        =   4466
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   372
      Index           =   0
      Left            =   9480
      TabIndex        =   16
      Top             =   2100
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Nuevo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   492
      Left            =   8760
      TabIndex        =   17
      Top             =   120
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.FlatEdit txtInstrumento 
      Height          =   312
      Left            =   1920
      TabIndex        =   18
      Top             =   840
      Width           =   9252
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAdministrador 
      Height          =   312
      Left            =   1920
      TabIndex        =   19
      Top             =   1200
      Width           =   9252
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPortafolio 
      Height          =   312
      Left            =   1920
      TabIndex        =   20
      Top             =   1560
      Width           =   9252
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbMovimiento 
      Height          =   2772
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   11052
      _Version        =   1441793
      _ExtentX        =   19494
      _ExtentY        =   4890
      _StockProps     =   79
      Caption         =   "Fondos de Inversion:"
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
      BorderStyle     =   2
      Begin XtremeSuiteControls.DateTimePicker dtpFecha 
         Height          =   312
         Left            =   3120
         TabIndex        =   22
         Top             =   120
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtIntAcum 
         Height          =   312
         Left            =   3120
         TabIndex        =   23
         Top             =   2040
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
      Begin XtremeSuiteControls.FlatEdit txtLiquidacion 
         Height          =   312
         Left            =   3120
         TabIndex        =   24
         ToolTipText     =   "Indique el monto TOTAL depositado a la cuenta"
         Top             =   960
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   550
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   3120
         TabIndex        =   25
         Top             =   600
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntereses 
         Height          =   312
         Left            =   3120
         TabIndex        =   26
         Top             =   2400
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
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   5160
         TabIndex        =   27
         Top             =   960
         Width           =   852
         _Version        =   1441793
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
         Left            =   6000
         TabIndex        =   28
         Top             =   960
         Width           =   1212
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtImporteLocal 
         Height          =   312
         Left            =   7200
         TabIndex        =   29
         Top             =   960
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
         Height          =   912
         Left            =   5400
         TabIndex        =   30
         Top             =   1800
         Width           =   5412
         _Version        =   1441793
         _ExtentX        =   9546
         _ExtentY        =   1609
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
      Begin XtremeSuiteControls.FlatEdit txtPrincipal 
         Height          =   312
         Left            =   3120
         TabIndex        =   42
         Top             =   1680
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
         Index           =   7
         Left            =   1080
         TabIndex        =   43
         Top             =   1680
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Principal"
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
         Left            =   1080
         TabIndex        =   39
         Top             =   120
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fecha Liquidación"
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
         Left            =   1080
         TabIndex        =   38
         Top             =   2040
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Intereses Acumulados"
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
         Index           =   42
         Left            =   1080
         TabIndex        =   37
         Top             =   960
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto del Liquidación"
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
         Index           =   44
         Left            =   1080
         TabIndex        =   36
         Top             =   600
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "No. Documento "
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
         Index           =   9
         Left            =   1080
         TabIndex        =   35
         Top             =   2400
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Intereses"
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
         Index           =   10
         Left            =   7200
         TabIndex        =   34
         Top             =   720
         Width           =   1932
         _Version        =   1441793
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
         Index           =   5
         Left            =   6120
         TabIndex        =   33
         Top             =   720
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "T.C:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   5
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   8
         Left            =   5160
         TabIndex        =   32
         Top             =   720
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
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
         Alignment       =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   11
         Left            =   5400
         TabIndex        =   31
         Top             =   1440
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas:"
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
      Height          =   612
      Index           =   1
      Left            =   7920
      TabIndex        =   40
      Top             =   7560
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Procesar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmIVR_Proc_Liquidacion.frx":05A4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnTransac 
      Height          =   612
      Index           =   2
      Left            =   9720
      TabIndex        =   41
      Top             =   7560
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   1080
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
      Appearance      =   14
      Picture         =   "frmIVR_Proc_Liquidacion.frx":0F67
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   6
      Left            =   4560
      TabIndex        =   15
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Id Transac:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   492
      Left            =   120
      TabIndex        =   13
      Top             =   5400
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Comprobantes:"
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
      Index           =   4
      Left            =   4560
      TabIndex        =   12
      Top             =   7560
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
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
      Left            =   4560
      TabIndex        =   10
      Top             =   7920
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Pendiente"
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
      Index           =   2
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   2052
      _Version        =   1441793
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   2052
      _Version        =   1441793
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
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   2052
      _Version        =   1441793
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
      Index           =   3
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Inversión"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scGestion 
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   11172
      _Version        =   1441793
      _ExtentX        =   19706
      _ExtentY        =   868
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
   End
End
Attribute VB_Name = "frmIVR_Proc_Liquidacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem, vFecha As Date
Dim vDivisaLocaL As String


Private Sub sbAdquisicion_Load()
Call sbIVR_Transac_Load(lswAd, txtTransacId.Text, gIVR_Transito.Tipo, gIVR_Transito.Concepto)

Dim i As Integer, pMonto As Currency

With lswAd.ListItems

pMonto = 0
For i = 1 To .Count
    pMonto = pMonto + CCur(.Item(i).SubItems(3))
Next i

txtAd_Registrado.Text = Format(pMonto, "Standard")

If Not IsNumeric(txtImporteLocal.Text) Then
    txtImporteLocal.Text = "0"
End If

txtAd_Pendiente.Text = Format(CCur(txtImporteLocal.Text) - pMonto, "Standard")



End With
 
End Sub



Private Sub btnAdquisicion_Click(Index As Integer)



If Mid(txtEstado.Text, 1, 1) <> "S" Then
   Exit Sub
End If

Select Case Index
    Case 0 'Nuevo
        
        If CCur(txtImporteLocal.Text) = 0 Then
            MsgBox "No se ha indicado ningún Importe a Registrar!", vbExclamation
            Exit Sub
        End If
        
        If CLng(txtTransacId.Text) = 0 Then
           Call sbGuardar
        End If
        
        
        gIVR_Transito.Codigo = txtTransacId.Text
        gIVR_Transito.Divisa = vDivisaLocaL
        gIVR_Transito.TipoCambio = 1
        gIVR_Transito.Monto = CCur(txtAd_Pendiente.Text)
        
        frmIVR_Rec_Bancos_Registro.Show vbModal
        
    Case 1 'Eliminar
        
        Dim i As Integer
        With lswAd.ListItems
            For i = 1 To .Count
                If .Item(i).Checked = True Then
                    strSQL = "delete  IVR_TRANSACCIONES Where TRANSAC_ID = " & .Item(i).Text
                    Call ConectionExecute(strSQL)
                End If
            Next i
        End With
End Select

Call sbAdquisicion_Load

End Sub


Private Sub btnTransac_Click(Index As Integer)

On Error GoTo vError

Select Case Index
    Case 0 'Nueva Solicitud
        Call sbInicializa
    
    Case 1 'Procesar
        If Len(txtDocumento.Text) = 0 Then
            MsgBox "No ha Indicado un No. de Documento para esta transacción!", vbExclamation
            Exit Sub
        End If
    
        If dtpFecha.Value < vFecha Then
            MsgBox "La fecha de esta transacción es menor al último corte!", vbExclamation
            Exit Sub
        End If
    
       If CCur(txtIntereses.Text) < 0 Then
            MsgBox "Los intereses al corte de la liquidación no pueden ser negativos!", vbExclamation
            Exit Sub
       End If
    
        If CCur(txtImporteLocal.Text) > 0 And CCur(txtAd_Pendiente.Text) = 0 Then
    
            If CLng(txtTransacId.Text) > 0 And Mid(txtEstado.Text, 1, 1) = "S" Then
               Call sbGuardar("A")
               MsgBox "Transacción procesada Satisfactoriamente!", vbInformation
            Else
               MsgBox "Transacción ya se encuentra procesada!", vbInformation
            End If
        
        
        Else
            MsgBox "No se ha indicado ningún valor a la transacción!", vbExclamation
        End If
        
    Case 2 'Eliminar
        If CLng(txtTransacId.Text) > 0 And Mid(txtEstado.Text, 1, 1) = "S" Then
           Call sbGuardar("X")
           Call sbInicializa
        Else
            MsgBox "No se ha indicado una transacción y/o esta ya fue procesada!", vbExclamation
        End If
End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub



Private Sub Form_Load()

On Error GoTo vError

 
txtInversionId.Text = gIVR_Transito.TituloId
txtTransacId.Text = "0"
txtEstado.Text = "Solicitud"

scGestion.Caption = "Gestión:   Registro de Cupón"

strSQL = "select  isnull(max(CORTE), dbo.mygetdate())  as 'CORTE'" _
       & "  From IVR_CIERRES"
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!Corte
rs.Close

Exit Sub

vError:


End Sub



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


On Error GoTo vError

vPaso = True


strSQL = "select COD_DIVISA  From vSys_Divisas  Where DIVISA_LOCAL = 1"
Call OpenRecordSet(rs, strSQL)
   vDivisaLocaL = Trim(rs!Cod_Divisa)
rs.Close


txtDivisa.Text = gIVR_Transito.Divisa
txtTipoCambio.Text = gIVR_Transito.TipoCambio


If vDivisaLocaL = gIVR_Transito.Divisa Then
    txtTipoCambio.Text = "1"
    txtTipoCambio.Locked = True
Else
    txtTipoCambio.Locked = False
End If

vPaso = False

Call sbConsulta(gIVR_Transito.TituloId)
Call sbInicializa
Call sbAdquisicion_Load

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
    
    dtpFecha.Value = DateAdd("d", 1, vFecha)
    
    txtAdministrador.Text = rs!Administrador_Desc
    txtAdministrador.Tag = rs!Cod_Administrador
    
    txtInstrumento.Text = rs!Instrumento_Desc
    txtInstrumento.Tag = rs!Cod_Instrumento
    
    txtPortafolio.Text = rs!Portafolio_Desc
    txtPortafolio.Tag = rs!Cod_Portafolio
    
        
         
         
        
    txtLiquidacion.Text = 0
    txtIntAcum.Text = 0
    txtIntereses.Text = 0
    txtPrincipal.Text = 0
    
    If vDivisaLocaL = IIf(IsNull(rs!Cod_Divisa), gIVR_Transito.Divisa, rs!Cod_Divisa) Then
        txtTipoCambio.Text = "1"
        txtTipoCambio.Locked = True
    Else
        txtTipoCambio.Locked = False
    End If
    
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

txtTransacId.Text = "0"
txtEstado.Text = "Solicitud"

txtDocumento.Text = ""

dtpFecha.Value = DateAdd("d", 1, vFecha)
txtLiquidacion.Text = Format(0, "Standard")

txtIntAcum.Text = Format(0, "Standard")
txtIntereses.Text = Format(0, "Standard")
txtPrincipal.Text = Format(0, "Standard")



strSQL = "exec spIVR_TITULOS_L_CONSULTA " & txtInversionId.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtTransacId.Text = "0"
    
    txtDocumento.Text = ""
    
    dtpFecha.Value = rs!Fecha_Corte
    
    txtLiquidacion.Text = Format(rs!Mnt_Total, "Standard")
    
    txtIntAcum.Text = Format(rs!MNT_INTERES_APL, "Standard")
    txtIntereses.Text = Format(rs!Mnt_Interes, "Standard")
    txtPrincipal.Text = Format(rs!Mnt_Principal, "Standard")
    
    
End If

txtLiquidacion.SetFocus

txtNotas.Text = ""

lswAd.ListItems.Clear
txtAd_Pendiente.Text = Format(0, "Standard")
txtAd_Registrado.Text = Format(0, "Standard")


lswAd.ListItems.Clear

vPaso = False

dtpFecha.SetFocus

End Sub


Private Sub sbGuardar(Optional pEstado As String = "S")
On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pTipoMov As String, pMonto As Currency, pDivisa As String, pTipoCambio As Currency
Dim pCuenta As String, pBancoId As Long


With gIVR_Transito
    .Tipo = "L"
    .Concepto = "Liq"
    
    pMonto = CCur(txtLiquidacion.Text)
    pTipoMov = "L"

' spIVR_TITULOS_L_REGISTRA (@TituloId int, @TipoMov char(1), @Concepto varchar(10)
'                , @Fecha datetime, @Documento varchar(30)
'                , @Total dec(18,2), @IntAcum dec(18,2),   @Intereses dec(18,2), @Principal dec(18,2)
'                , @Divisa varchar(10), @TipoCambio dec(10,4), @Notas varchar(500)
'                , @Usuario varchar(30), @Estado char(1) = 'S', @TransacId int = 0)

strSQL = "exec spIVR_TITULOS_L_REGISTRA " & txtInversionId.Text & ", '" & pTipoMov _
       & "','" & .Concepto _
       & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") & "','" & txtDocumento.Text _
       & "', " & pMonto _
       & " , " & CCur(txtIntAcum.Text) _
       & " , " & CCur(txtIntereses.Text) _
       & " , " & CCur(txtPrincipal.Text) _
       & " ,'" & txtDivisa.Text & "' ," & CCur(txtTipoCambio.Text) _
       & " ,'" & txtNotas.Text _
       & "','" & glogon.Usuario & "', '" & pEstado & "', " & txtTransacId.Text
End With

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
    txtTransacId.Text = rs!Transac_Id
    
    If rs!Estado = "S" Then
        txtEstado.Text = "Solicitud"
    Else
        txtEstado.Text = "Procesada"
    End If
Else
     Call sbInicializa
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbMovimiento_Cal_Refresh()

On Error GoTo vError

If Not IsNumeric(txtLiquidacion.Text) Then
    txtLiquidacion.Text = "0"
End If

If Not IsNumeric(txtIntAcum.Text) Then
    txtIntAcum.Text = "0"
End If

If Not IsNumeric(txtIntereses.Text) Then
    txtIntereses.Text = "0"
End If

If Not IsNumeric(txtPrincipal.Text) Then
    txtPrincipal.Text = "0"
End If

If Not IsNumeric(txtTipoCambio.Text) Then
    txtTipoCambio.Text = "1"
End If

    
If Not IsNumeric(txtAd_Registrado.Text) Then
    txtAd_Registrado.Text = "0"
End If

If Not IsNumeric(txtAd_Pendiente.Text) Then
    txtAd_Pendiente.Text = "0"
End If
    
'Formato
txtIntAcum.Text = Format(CCur(txtIntAcum.Text), "Standard")
txtIntereses.Text = Format(CCur(txtIntereses.Text), "Standard")
txtPrincipal.Text = Format(CCur(txtPrincipal.Text), "Standard")

txtLiquidacion.Text = Format(CCur(txtLiquidacion.Text), "Standard")


txtIntereses.Text = Format(CCur(txtLiquidacion.Text) - CCur(txtIntAcum.Text) - CCur(txtPrincipal.Text), "Standard")

txtTipoCambio.Text = Format(CCur(txtTipoCambio.Text), "###,##0.000000")
txtImporteLocal.Text = Format(CCur(txtLiquidacion.Text) * fxSys_Tipo_Cambio_Apl(CCur(txtTipoCambio.Text)), "Standard")

txtAd_Pendiente.Text = Format(CCur(txtImporteLocal.Text) - CCur(txtAd_Registrado.Text), "Standard")

Exit Sub

vError:

End Sub


Private Sub txtLiquidacion_GotFocus()
On Error GoTo vError
    txtLiquidacion.Text = Format(CCur(txtLiquidacion.Text), "Standard")
vError:
End Sub

Private Sub txtLiquidacion_LostFocus()
Call sbMovimiento_Cal_Refresh
End Sub


Private Sub txtTipoCambio_GotFocus()
On Error GoTo vError
    txtTipoCambio.Text = Format(CCur(txtTipoCambio.Text), "###,##0.000000")
vError:
End Sub

Private Sub txtTipoCambio_LostFocus()
Call sbMovimiento_Cal_Refresh
End Sub



