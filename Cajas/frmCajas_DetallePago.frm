VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCajas_DetallePago 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recepción de Valores"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2415
      Left            =   120
      TabIndex        =   49
      Top             =   1560
      Width           =   3585
      _Version        =   1572864
      _ExtentX        =   6324
      _ExtentY        =   4260
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
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle de Valores Registrados..: "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   6252
      Left            =   9360
      TabIndex        =   31
      Top             =   1080
      Width           =   8892
      Begin XtremeSuiteControls.ListView lswDetalle 
         Height          =   4815
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   8535
         _Version        =   1572864
         _ExtentX        =   15055
         _ExtentY        =   8493
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
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnDET_Borrar 
         Height          =   372
         Left            =   3240
         TabIndex        =   52
         Top             =   240
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Borrar"
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
         Appearance      =   21
         Picture         =   "frmCajas_DetallePago.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnDET_Cerrar 
         Height          =   372
         Left            =   4440
         TabIndex        =   53
         Top             =   240
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Appearance      =   21
         Picture         =   "frmCajas_DetallePago.frx":05A4
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalValores 
         Height          =   315
         Left            =   6480
         TabIndex        =   76
         Top             =   5640
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
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
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Total valores..:"
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
         Left            =   4560
         TabIndex        =   32
         Top             =   5640
         Width           =   1695
      End
   End
   Begin XtremeSuiteControls.CheckBox chkSaldosFavor 
      Height          =   255
      Left            =   2160
      TabIndex        =   75
      Top             =   5400
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Aplicar Saldos a  favor"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Alignment       =   1
   End
   Begin VB.Frame fraDocumento 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   9360
      TabIndex        =   19
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.FlatEdit txtMontoDocumentos 
         Height          =   330
         Left            =   1560
         TabIndex        =   81
         Top             =   1560
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocCuenta 
         Height          =   330
         Left            =   1560
         TabIndex        =   83
         Top             =   840
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   330
         Left            =   1560
         TabIndex        =   84
         Top             =   480
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   108
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Documento:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   1
         Left            =   4080
         TabIndex        =   82
         Top             =   1560
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Doc."
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
         Height          =   315
         Index           =   25
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   210
         Index           =   17
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame fraTarjeta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   14400
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.ComboBox cboTarjetas 
         Height          =   315
         Left            =   1560
         TabIndex        =   61
         Top             =   480
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.FlatEdit txtMontoTarjetas 
         Height          =   330
         Left            =   1560
         TabIndex        =   89
         ToolTipText     =   "Digite el Monto Total del Cheque"
         Top             =   1800
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTarjeta 
         Height          =   330
         Left            =   1560
         TabIndex        =   91
         Top             =   840
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAutoriza 
         Height          =   330
         Left            =   1560
         TabIndex        =   92
         Top             =   1200
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   107
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Tarjeta Débito/Crédito:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   3
         Left            =   4080
         TabIndex        =   90
         Top             =   1800
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Autorización N°"
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
         Height          =   315
         Index           =   24
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Tarjeta"
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
         Height          =   315
         Index           =   10
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta N°"
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
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   315
         Index           =   11
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   975
      End
   End
   Begin VB.Frame fraEfectivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   9360
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.FlatEdit txtMontoEfectivo 
         Height          =   330
         Left            =   1560
         TabIndex        =   96
         ToolTipText     =   "Digite el Monto Total del Cheque"
         Top             =   960
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Efectivo:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   4
         Left            =   4080
         TabIndex        =   97
         Top             =   960
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Index           =   22
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   1212
      End
   End
   Begin VB.Frame fraCheque 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   14400
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.ComboBox cboEmisorCheque 
         Height          =   315
         Left            =   1560
         TabIndex        =   60
         Top             =   480
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.FlatEdit txtMontoCheques 
         Height          =   330
         Left            =   1560
         TabIndex        =   85
         ToolTipText     =   "Digite el Monto Total del Cheque"
         Top             =   1920
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtChequeCta 
         Height          =   330
         Left            =   1560
         TabIndex        =   87
         Top             =   960
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCheque 
         Height          =   330
         Left            =   1560
         TabIndex        =   88
         Top             =   1320
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   110
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Cheque:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   2
         Left            =   4080
         TabIndex        =   86
         Top             =   1920
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cuenta"
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
         Height          =   315
         Index           =   30
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cheque"
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
         Height          =   315
         Index           =   13
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Emisor"
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
         Height          =   315
         Index           =   14
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   210
         Index           =   18
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
   End
   Begin VB.Frame fraDepositos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   9360
      TabIndex        =   37
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.ComboBox cboDepositoBanco 
         Height          =   312
         Left            =   1560
         TabIndex        =   58
         Top             =   480
         Width           =   3012
         _Version        =   1572864
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.DateTimePicker dtpDeposito 
         Height          =   315
         Left            =   1560
         TabIndex        =   69
         Top             =   1200
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtDepositoNo 
         Height          =   330
         Left            =   1560
         TabIndex        =   95
         Top             =   840
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoDeposito 
         Height          =   330
         Left            =   1560
         TabIndex        =   98
         ToolTipText     =   "Digite el Monto Total del Cheque"
         Top             =   1680
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   109
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Depósito en Cuenta:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   5
         Left            =   4080
         TabIndex        =   99
         Top             =   1680
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Index           =   29
         Left            =   240
         TabIndex        =   41
         Top             =   1200
         Width           =   1452
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   210
         Index           =   23
         Left            =   240
         TabIndex        =   40
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
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
         Height          =   315
         Index           =   20
         Left            =   240
         TabIndex        =   39
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No.Deposito"
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
         Height          =   315
         Index           =   16
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame fraFondos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   14400
      TabIndex        =   43
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
      Begin XtremeSuiteControls.ComboBox cboFondos 
         Height          =   312
         Left            =   1560
         TabIndex        =   57
         Top             =   480
         Width           =   3012
         _Version        =   1572864
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.FlatEdit txtFondoMonto 
         Height          =   330
         Left            =   1560
         TabIndex        =   93
         Top             =   960
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFondoDisponible 
         Height          =   330
         Left            =   1560
         TabIndex        =   94
         Top             =   1320
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoFondos 
         Height          =   330
         Left            =   1560
         TabIndex        =   100
         ToolTipText     =   "Digite el Monto Total del Cheque"
         Top             =   1800
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   111
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Fondos Disponibles:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   6
         Left            =   4080
         TabIndex        =   101
         Top             =   1800
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
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
         Height          =   315
         Index           =   34
         Left            =   240
         TabIndex        =   47
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   210
         Index           =   33
         Left            =   240
         TabIndex        =   46
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   210
         Index           =   32
         Left            =   240
         TabIndex        =   45
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
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
         Height          =   210
         Index           =   31
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
   End
   Begin VB.Timer TimerInicial 
      Interval        =   10
      Left            =   240
      Top             =   240
   End
   Begin VB.TextBox txtCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   480
      Width           =   6852
   End
   Begin VB.TextBox txtServicio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
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
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   6852
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":0BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":14BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":2396
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":2492
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":25AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":26E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_DetallePago.frx":27DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   8010
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Código de Caja"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Apertura"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            TextSave        =   "24/3/2025"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            TextSave        =   "16:09"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   312
      Left            =   6480
      TabIndex        =   50
      Top             =   1100
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
   Begin XtremeSuiteControls.PushButton btnAgregar 
      Height          =   540
      Left            =   4680
      TabIndex        =   54
      Top             =   5280
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   952
      _StockProps     =   79
      Caption         =   "Agregar"
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
      Appearance      =   21
      Picture         =   "frmCajas_DetallePago.frx":28E1
   End
   Begin XtremeSuiteControls.PushButton btnDetalle 
      Height          =   540
      Left            =   6120
      TabIndex        =   55
      Top             =   5280
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   952
      _StockProps     =   79
      Caption         =   "Detalle"
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
      Appearance      =   21
      Picture         =   "frmCajas_DetallePago.frx":30D0
   End
   Begin XtremeSuiteControls.PushButton btnCerrar 
      Height          =   540
      Left            =   7560
      TabIndex        =   56
      Top             =   5280
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   952
      _StockProps     =   79
      Caption         =   "Cerrar"
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
      Appearance      =   21
      Picture         =   "frmCajas_DetallePago.frx":38AF
   End
   Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
      Height          =   315
      Left            =   6360
      TabIndex        =   62
      Top             =   6360
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   315
      Left            =   6360
      TabIndex        =   63
      Top             =   6720
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCambio 
      Height          =   315
      Left            =   6360
      TabIndex        =   64
      Top             =   7080
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   315
      Left            =   6360
      TabIndex        =   65
      Top             =   7440
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtMontoDivisaTransaccion 
      Height          =   315
      Left            =   6720
      TabIndex        =   66
      Top             =   4440
      Width           =   2055
      _Version        =   1572864
      _ExtentX        =   3625
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtMontoDivisaFuncional 
      Height          =   315
      Left            =   6720
      TabIndex        =   67
      Top             =   4080
      Width           =   2055
      _Version        =   1572864
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   1395
      Left            =   120
      TabIndex        =   68
      Top             =   6360
      Width           =   4095
      _Version        =   1572864
      _ExtentX        =   7223
      _ExtentY        =   2461
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
   Begin VB.Frame fraSaldos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   3720
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   5080
      Begin XtremeSuiteControls.ComboBox cboSaldoFavorReferencia 
         Height          =   312
         Left            =   1560
         TabIndex        =   59
         Top             =   480
         Width           =   3012
         _Version        =   1572864
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.FlatEdit txtSaldoFavorMntOrigen 
         Height          =   330
         Left            =   1560
         TabIndex        =   77
         Top             =   960
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldoFavorSaldo 
         Height          =   330
         Left            =   1560
         TabIndex        =   78
         Top             =   1320
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   582
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
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMontoSaldos 
         Height          =   330
         Left            =   1560
         TabIndex        =   79
         Top             =   1800
         Width           =   2535
         _Version        =   1572864
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scFormasPago 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   105
         Top             =   0
         Width           =   5175
         _Version        =   1572864
         _ExtentX        =   9128
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Saldos a Favor en Cajas:"
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
      Begin XtremeSuiteControls.Label lblDivisaInfo 
         Height          =   315
         Index           =   0
         Left            =   4080
         TabIndex        =   80
         Top             =   1800
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
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
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
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
         Height          =   210
         Index           =   27
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Doc."
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
         Height          =   210
         Index           =   26
         Left            =   240
         TabIndex        =   29
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Height          =   210
         Index           =   19
         Left            =   240
         TabIndex        =   24
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
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
         Height          =   315
         Index           =   15
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.ComboBox cboOrigenRecursos 
      Height          =   330
      Left            =   120
      TabIndex        =   73
      Top             =   4920
      Width           =   4095
      _Version        =   1572864
      _ExtentX        =   7223
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
      Enabled         =   0   'False
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboPagadores 
      Height          =   330
      Left            =   120
      TabIndex        =   102
      Top             =   4320
      Width           =   4095
      _Version        =   1572864
      _ExtentX        =   7223
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
      Enabled         =   0   'False
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label lblTCDivisaTransaccion 
      Height          =   315
      Left            =   7440
      TabIndex        =   104
      Top             =   4800
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "..."
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
      Alignment       =   5
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   103
      Top             =   4680
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Origen Recursos"
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
   Begin XtremeSuiteControls.Label Label4 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   74
      Top             =   4080
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Pagadores"
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
   Begin XtremeSuiteControls.Label lblTipoCambio 
      Height          =   372
      Left            =   4920
      TabIndex        =   72
      Top             =   1080
      Width           =   1452
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "1"
      ForeColor       =   4210752
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
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "T.C. :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   36
      Left            =   6600
      TabIndex        =   71
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Divisa Transacción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   450
      Index           =   35
      Left            =   4560
      TabIndex        =   48
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Divisa Local"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   12
      Left            =   3960
      TabIndex        =   36
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label lblDivisa 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8400
      TabIndex        =   35
      ToolTipText     =   "Divisa Origen de la Transacción"
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Requerido .:"
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
      Height          =   210
      Index           =   28
      Left            =   4800
      TabIndex        =   34
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto ..:"
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
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente ..:"
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
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1452
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      Index           =   0
      X1              =   120
      X2              =   8760
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Notas..:"
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
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Label lblFormaPago 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Formas de Pago:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia ..:"
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
      Height          =   210
      Index           =   3
      Left            =   4800
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cambio ..:"
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
      Height          =   210
      Index           =   2
      Left            =   4800
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Detallado .:"
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
      Height          =   210
      Index           =   5
      Left            =   4800
      TabIndex        =   0
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Totales..:"
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
      Height          =   255
      Index           =   21
      Left            =   4680
      TabIndex        =   28
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   252
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   1200
      Width           =   1452
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   380
      Left            =   0
      TabIndex        =   70
      Top             =   1080
      Width           =   9012
      _Version        =   1572864
      _ExtentX        =   15896
      _ExtentY        =   670
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
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmCajas_DetallePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mLinea As Integer, mTipoCambio As Currency, mDivisaFuncional As String
Dim mFormaPago As String, mCuenta As String, mAplSaldoFavor As Integer, mTipo As String, mFecha As Date
Dim mDP_Caso As Integer, mDP_Proceso As Integer, mAplOrigenRecursos As Integer



Private Sub sbDivisaTipoCambio()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDivisa As String, i As Integer

If vPaso Then Exit Sub

'Cargar el Tipo de Cambio
mTipoCambio = 1

vDivisa = cboDivisa.ItemData(cboDivisa.ListIndex)

strSQL = "select dbo.fxCajas_TipoCambio(" & GLOBALES.gEnlace & ",'" & vDivisa & "',dbo.MyGetdate(),'C') as 'TipoCambio'"
Call OpenRecordSet(rs, strSQL)
  mTipoCambio = rs!TipoCambio
rs.Close

lblTipoCambio.Caption = mTipoCambio

For i = 0 To lblDivisaInfo.Count - 1
   lblDivisaInfo.Item(i).Caption = vDivisa
   lblDivisaInfo.Item(i).ToolTipText = "Tipo cambio .: " & mTipoCambio
   lblDivisaInfo.Item(i).Tag = mTipoCambio
Next i


'Recalcular datos en pantalla
Select Case mTipo
   Case "E" 'Efectivo
        Call sbCalculo(txtMontoEfectivo)
   Case "F" 'Fondos
        Call sbCalculo(txtMontoFondos)
   Case "C" 'Cheques
        Call sbCalculo(txtMontoCheques)
   Case "B" 'Depositos
        Call sbCalculo(txtMontoDeposito)
   Case "D" 'Documentos
        Call sbCalculo(txtMontoDocumentos)
   Case "T" 'Tarjetas
        Call sbCalculo(txtMontoTarjetas)
   Case "S" 'Saldos a Favor
        Call sbCalculo(txtMontoSaldos)
End Select

End Sub

Private Sub btnAgregar_Click()
      If fxVerificaDatos Then
        Call sbGuardar
      Else
        Exit Sub
      End If
End Sub

Private Sub btnCerrar_Click()
 Unload Me
End Sub

Private Sub btnDET_Borrar_Click()
Dim i As Integer, strSQL As String

On Error GoTo vError


With lswDetalle.ListItems
    For i = 1 To .Count
        If .Item(i).Checked Then
           strSQL = "delete CAJAS_DESGLOCE_PAGO" _
                  & " Where Cod_Caja = '" & ModuloCajas.mCaja & "' and cod_Apertura = " & ModuloCajas.mApertura _
                  & " and Ticket = '" & ModuloCajas.mTiquete & "' and Linea = " & .Item(i).Text
           Call ConectionExecute(strSQL)
        End If
    Next i
End With
Call sbCargaTiquete

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnDET_Cerrar_Click()

fraDetalle.Visible = False
lsw.Visible = True

End Sub

Private Sub btnDetalle_Click()
     lsw.Visible = False
     fraDetalle.Visible = True
     fraDetalle.top = 960
     fraDetalle.Left = 0
End Sub

Private Sub cboDivisa_Change()
Call sbDivisaTipoCambio
End Sub

Private Sub cboDivisa_Click()
Call sbDivisaTipoCambio
End Sub

Private Sub cboEmisorCheque_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtChequeCta.SetFocus
End Sub



Private Sub cboFondos_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Or cboFondos.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_DisponibleFondos '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura _
       & ",'" & ModuloCajas.mTiquete & "','" & cboFondos.ItemData(cboFondos.ListIndex) _
       & "'," & SIFGlobal.fxCodText(cboFondos.Text)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 txtFondoMonto.Text = Format(rs!Monto, "Standard")
 txtFondoDisponible.Text = Format(rs!Disponible, "Standard")
 Call sbCboAsignaDato(cboDivisa, Trim(rs!Divisa_Desc), False)
Else
 txtFondoMonto.Text = Format(0, "Standard")
 txtFondoDisponible.Text = Format(0, "Standard")

End If
rs.Close

If CCur(txtDiferencia.Text) > 0 Then
    If CCur(txtFondoDisponible.Text) >= Abs(CCur(txtDiferencia.Text)) Then
        txtMontoFondos.Text = Abs(CCur(txtDiferencia.Text))
    Else
        txtMontoFondos.Text = txtFondoDisponible.Text
    End If
Else
    txtMontoFondos.Text = Format(0, "Standard")

End If

Call txtMontoFondos_Change

Me.MousePointer = vbDefault

End Sub

Private Sub cboSaldoFavorReferencia_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Or cboSaldoFavorReferencia.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "exec spCajas_SaldoFavor '" & ModuloCajas.mClienteId & "'," _
       & cboSaldoFavorReferencia.ItemData(cboSaldoFavorReferencia.ListIndex) _
       & ",'" & cboSaldoFavorReferencia.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 txtSaldoFavorMntOrigen.Text = Format(rs!Monto, "Standard")
 txtSaldoFavorSaldo.Text = Format(rs!Saldo, "Standard")
 Call sbCboAsignaDato(cboDivisa, rs!Divisa_Desc, False)
Else
 txtSaldoFavorMntOrigen.Text = Format(0, "Standard")
 txtSaldoFavorSaldo.Text = Format(0, "Standard")
End If
rs.Close

If CCur(txtDiferencia.Text) > 0 Then
    If CCur(txtSaldoFavorSaldo.Text) >= Abs(CCur(txtDiferencia.Text)) Then
        txtMontoSaldos.Text = Abs(CCur(txtDiferencia.Text))
    Else
        txtMontoSaldos.Text = txtSaldoFavorSaldo.Text
    End If
Else
 txtMontoSaldos.Text = Format(0, "Standard")
End If

Call txtMontoSaldos_Change


Me.MousePointer = vbDefault
End Sub


Private Sub cboSaldoFavorReferencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSaldoFavorMntOrigen.SetFocus
End Sub

Private Sub cboTarjetas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTarjeta.SetFocus
End Sub



Private Sub dtpDeposito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoDeposito.SetFocus
End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub


Private Function fxDivisaFuncional() As String

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select dbo.fxCajas_DivisaFuncional(" & GLOBALES.gEnlace & ") as 'Divisa'"
Call OpenRecordSet(rs, strSQL)

fxDivisaFuncional = Trim(rs!Divisa)

rs.Close


End Function


Private Sub Form_Load()
Dim strSQL As String

vModulo = 5

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mDivisaFuncional = fxDivisaFuncional

mFecha = fxFechaServidor
mLinea = 0
mTipoCambio = 1
mAplOrigenRecursos = 0

With lsw.ColumnHeaders
   .Clear
   .Add , , "Forma de Pago", 2500
   .Add , , "[ ID ]", 800, vbCenter
   .Add , , "", 1 'Para Reservar el Valor de Saldos a Favor
   .Add , , "", 1 'Para Reservar el Valor de Origen de Recursos
End With

With lswDetalle.ColumnHeaders
   .Clear
   .Add , , "Linea", 800
   .Add , , "Forma de Pago", 2800
   .Add , , "Monto", 1800, vbRightJustify
   .Add , , "Saldo", 1800, vbRightJustify
   .Add , , "Divisa", 1000, vbCenter
   .Add , , "Referencia", 3000
   .Add , , "Tipo de Cambio", 1800, vbRightJustify
   .Add , , "Importe", 1800, vbRightJustify
  
End With


vAseDocDeposito = ""

Me.Width = 8985

txtServicio.Text = ModuloCajas.mServicio
txtCliente.Text = ModuloCajas.mClienteId + " ...: " & ModuloCajas.mCliente

txtTotal.Text = 0
txtCambio.Text = 0
txtTotalPagar.Text = Format(ModuloCajas.mTotalAplicar, "Standard")

StatusBar.Panels.Item(1) = ModuloCajas.mCaja
StatusBar.Panels.Item(2) = ModuloCajas.mUsuario
StatusBar.Panels.Item(3) = "Apertura No.: " & ModuloCajas.mApertura

lblDivisa.Caption = ModuloCajas.mDivisa

lblTCDivisaTransaccion.Caption = fxCajasTipoCambio(lblDivisa.Caption)

End Sub

Private Sub sbCargaFrames()
Dim pTop As Long, pLeft As Long, pW As Long

On Error GoTo vError

pTop = 1560
pLeft = 3720
pW = 5080

fraDetalle.Visible = False

If lsw.ListItems.Count = 0 Or vPaso Then Exit Sub

'Formato de la forma de Pago
Select Case mTipo
   Case "E" 'Efectivo
     fraEfectivo.Visible = True
     fraEfectivo.top = pTop
     fraEfectivo.Left = pLeft
     fraEfectivo.Width = pW
     
     Call sbApagaFrames(fraEfectivo.Name)
     txtMontoEfectivo.SetFocus
     
     txtMontoEfectivo.Text = Abs(CCur(txtDiferencia.Text))
   
   Case "T" 'Tarjetas
     fraTarjeta.Visible = True
     fraTarjeta.top = pTop
     fraTarjeta.Left = pLeft
     fraTarjeta.Width = pW
     
     Call sbApagaFrames(fraTarjeta.Name)
     
     cboTarjetas.SetFocus
     txtMontoTarjetas.Text = Abs(CCur(txtDiferencia.Text))

     
   Case "B" 'Depositos
     fraDepositos.Visible = True
     fraDepositos.top = pTop
     fraDepositos.Left = pLeft
     fraDepositos.Width = pW
     
     dtpDeposito.Value = mFecha
     
     txtDepositoNo.SetFocus
     txtMontoDeposito.Text = 0 'Abs(CCur(txtDiferencia.Text))
     
     Call sbApagaFrames(fraDepositos.Name)
     Call sbDepositoCuentas
     
   Case "C" 'Cheques
     fraCheque.Visible = True
     fraCheque.top = pTop
     fraCheque.Left = pLeft
     fraCheque.Width = pW
   
     txtCheque.SetFocus
     txtMontoCheques.Text = 0 'Abs(CCur(txtDiferencia.Text))
     
     Call sbApagaFrames(fraCheque.Name)
     
   Case "D" 'Documentos
     fraDocumento.Visible = True
     fraDocumento.top = pTop
     fraDocumento.Left = pLeft
     fraDocumento.Width = pW
     
     txtMontoDocumentos.Text = Abs(CCur(txtDiferencia.Text))
     txtDocCuenta.Text = fxgCntCuentaFormato(True, mCuenta, 0)
     txtDocumento.SetFocus
     
     Call sbApagaFrames(fraDocumento.Name)

   Case "F" 'Fondos
     fraFondos.Visible = True
     fraFondos.top = pTop
     fraFondos.Left = pLeft
     fraFondos.Width = pW
     
     txtMontoFondos.Text = Abs(CCur(txtDiferencia.Text))
     cboFondos.SetFocus
     
     Call sbApagaFrames(fraFondos.Name)
   
   Case "S" 'Saldos a Favor
     fraSaldos.Visible = True
     fraSaldos.top = pTop
     fraSaldos.Left = pLeft
     fraSaldos.Width = pW
     
     cboSaldoFavorReferencia.SetFocus
     txtMontoSaldos.Text = 0
     
     Call sbApagaFrames(fraSaldos.Name)
     
     If cboSaldoFavorReferencia.ListCount > 0 Then
        Call cboSaldoFavorReferencia_Click
     End If

    Case Else
         Call sbApagaFrames("")
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbDepositoCuentas()
Dim strSQL As String, rs As New ADODB.Recordset

'Carga las cuentas bancarias asiganadas a la forma de pago
cboDepositoBanco.Clear


strSQL = "exec spCajas_DepositosCuentasBancariasAut '" & mFormaPago & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboDepositoBanco.AddItem Trim(rs!Cta) & " - " & Trim(rs!Descripcion & "")
 cboDepositoBanco.ItemData(cboDepositoBanco.ListCount - 1) = CStr(rs!Id_Banco)
 
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   Call sbCboAsignaDato(cboDepositoBanco, Trim(rs!Cta) & " - " & Trim(rs!Descripcion & ""))
End If
rs.Close

End Sub


Private Sub Form_Unload(Cancel As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

ModuloCajas.mReciboDigital = False

strSQL = "select dbo.fxCajas_ReciboDigital('" & ModuloCajas.mCaja & "', " & ModuloCajas.mApertura & ", '" & ModuloCajas.mTiquete & "') as 'Recibo'"
Call OpenRecordSet(rs, strSQL)

If rs!Recibo = 1 Then
    ModuloCajas.mReciboDigital = True
End If

vError:

End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If lsw.ListItems.Count = 0 Or vPaso Then Exit Sub

'Linea Nueva
mLinea = 0
mCuenta = Item.ListSubItems.Item(1).Tag
mFormaPago = Item.SubItems(1)
mAplSaldoFavor = Item.SubItems(2)
mAplOrigenRecursos = Item.SubItems(3)

mTipo = Item.Tag


txtMontoEfectivo.Text = 0
txtMontoCheques.Text = 0
txtMontoTarjetas.Text = 0
txtMontoDocumentos.Text = 0
txtMontoSaldos.Text = 0
txtMontoDeposito.Text = 0
txtMontoFondos.Text = 0


txtMontoDivisaFuncional.Text = 0
txtMontoDivisaTransaccion.Text = 0


chkSaldosFavor.Value = mAplSaldoFavor

Call sbCargaFrames


'Formato de la forma de Pago
Select Case mTipo
   Case "E" 'Efectivo
'     txtMontoEfectivo.Text = Format(Abs(txtDiferencia.Text), "Standard")
     Call sbCalculo(txtMontoEfectivo)
   
   Case "T" 'Tarjetas
     Call sbCalculo(txtMontoTarjetas)

   Case "B" 'Deposito
     Call sbCalculo(txtMontoDeposito)
   
   Case "C" 'Cheques
     Call sbCalculo(txtMontoCheques)
     
   Case "D" 'Documentos
     Call sbCalculo(txtMontoDocumentos)
     
   Case "F" 'Fondos
     Call sbCalculo(txtMontoFondos)
     
   Case "S" 'Saldos a Favor
     Call sbCalculo(txtMontoSaldos)

End Select


If mAplOrigenRecursos = 1 Then
   cboPagadores.Enabled = True
   cboOrigenRecursos.Enabled = True
Else
   cboPagadores.Enabled = False
   cboOrigenRecursos.Enabled = False
End If

End Sub


Private Sub lswDetalle_DblClick()
Dim strSQL As String, rs As New ADODB.Recordset

If lswDetalle.ListItems.Count = 0 Then Exit Sub

mLinea = lswDetalle.SelectedItem.Text

strSQL = "select C.*, F.DESCRIPCION as 'FormaPagoDesc',F.TIPO, rtrim(C.cod_Divisa) as 'COD_DIVISA', D.descripcion as 'Divisa'" _
       & ", Rtrim(C.cod_Tarjeta) AS 'COD_TARJETA' ,  isnull(Tj.Descripcion ,'') as 'TarjetaDesc' " _
       & ", Rtrim(C.Cheque_Emisor) AS  'CHEQUE_EMISOR' , isnull(Em.Descripcion ,'') as 'EmisorCkDesc' " _
       & ", isnull(Bn.Cta,'') as 'BancoCta', isnull(Bn.Descripcion,'') as 'BancoDesc'" _
       & " from CAJAS_DESGLOCE_PAGO C inner join SIF_FORMAS_PAGO F on  C.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
       & "  inner join CNTX_Divisas D on C.cod_Divisa = D.cod_Divisa and D.cod_Contabilidad = " & GLOBALES.gEnlace _
       & "  left  join SIF_Tarjetas Tj on C.cod_Tarjeta = Tj.Cod_Tarjeta " _
       & "  left  join sif_emisores Em on C.Cheque_Emisor = Em.Cod_Emisor" _
       & "  left  join Tes_Bancos Bn on C.DP_Banco = Bn.Id_Banco" _
       & " Where C.cod_caja = '" & ModuloCajas.mCaja & "' and C.Ticket = '" & ModuloCajas.mTiquete _
       & "' and C.Cod_Apertura = " & ModuloCajas.mApertura & " and Linea = " & mLinea
Call OpenRecordSet(rs, strSQL)
    mCuenta = rs!cod_cuenta
    mFormaPago = rs!Cod_Forma_Pago
    mAplSaldoFavor = rs!Aplica_saldo_Favor
    mTipo = rs!Tipo
    
   chkSaldosFavor.Value = mAplSaldoFavor
   'Asigna la divisa y activa el tipo de cambio
   Call sbCboAsignaDato(cboDivisa, rs!Divisa)

Call sbCargaFrames


'Formato de la forma de Pago
Select Case mTipo
   Case "E" 'Efectivo
     txtMontoEfectivo.Text = Format(rs!Monto, "Standard")
     Call sbCalculo(txtMontoEfectivo)
   
   Case "T" 'Tarjetas
     txtAutoriza.Text = rs!Tarjeta_Autorizacion & ""
     txtMontoTarjetas.Text = Format(rs!Monto, "Standard")
     Call sbCboAsignaDato(cboTarjetas, rs!TarjetaDesc)
     
     cboTarjetas.SetFocus
     txtTarjeta.Text = rs!tarjeta_Numero & ""
     Call sbCalculo(txtMontoTarjetas)

   Case "B" 'Deposito
     txtMontoDeposito.Text = Format(rs!Monto, "Standard")

     Call sbCboAsignaDato(cboDepositoBanco, (Trim(rs!BancoCta) & " - " & Trim(rs!BancoDesc & "")))
      
     dtpDeposito.Value = rs!DP_Fecha
     txtDepositoNo.SetFocus
     txtDepositoNo.Text = rs!Num_Referencia & ""
     Call sbCalculo(txtMontoDeposito)
   
   
   Case "C" 'Cheques
     txtMontoCheques.Text = Format(rs!Monto, "Standard")

     Call sbCboAsignaDato(cboEmisorCheque, rs!EmisorCkDesc)
     
     txtCheque.SetFocus
     txtCheque.Text = rs!Cheque_Numero & ""
     Call sbCalculo(txtMontoCheques)
     
   Case "D" 'Documentos
     txtMontoDocumentos.Text = Format(rs!Monto, "Standard")
     txtDocCuenta.Text = fxgCntCuentaFormato(True, mCuenta, 0)
     txtDocumento.SetFocus
     txtDocumento = rs!Num_Referencia & ""
     Call sbCalculo(txtMontoDocumentos)
     
   Case "F" 'Fondos
     cboFondos.SetFocus
     Call sbCboAsignaDato(cboFondos, Trim(rs!COD_PLAN) & " - " & rs!COD_CONTRATO)
     
     txtMontoFondos.Text = Format(rs!Monto, "Standard")
     Call sbCalculo(txtMontoFondos)
     
   Case "S" 'Saldos a Favor
     cboSaldoFavorReferencia.SetFocus
     Call sbCboAsignaDato(cboSaldoFavorReferencia, rs!Num_Referencia)
     
     txtMontoSaldos.Text = Format(rs!Monto, "Standard")
     Call sbCalculo(txtMontoSaldos)

End Select

rs.Close

End Sub




Private Function fxSaldoFavorCta(pIdSaldoFavor As Long)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select dbo.fxCajas_SaldoFavorCuenta(" & pIdSaldoFavor & ") as Cuenta"
Call OpenRecordSet(rs, strSQL)
fxSaldoFavorCta = Trim(rs!Cuenta)
rs.Close

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, pDocumento As String, pSaldoFavorId As Long
Dim pBancoCuenta As String, pFondoPlan As String, pFondoContrato As Long

Me.MousePointer = vbHourglass

On Error GoTo vError

pBancoCuenta = ""
pDocumento = ""
pSaldoFavorId = 0

curMonto = CCur(txtMontoDivisaFuncional.Text)

Select Case mTipo
  Case "E" 'Efectivo
    
  Case "C" 'Cheques
    pDocumento = txtCheque.Text
    pBancoCuenta = txtChequeCta.Text
    
  Case "B" 'Depositos
    pDocumento = Trim(txtDepositoNo.Text)
    pBancoCuenta = SIFGlobal.fxCodText(cboDepositoBanco.Text)
    
  Case "T" 'Tarjetas
    curMonto = CCur(txtMontoTarjetas.Text) * fxSys_Tipo_Cambio_Apl(mTipoCambio)
  
  Case "D" 'Documentos
    mCuenta = fxgCntCuentaFormato(False, txtDocCuenta.Text, 0)
    
    pDocumento = txtDocumento.Text
  
  Case "F" 'Fondos
    pFondoPlan = cboFondos.ItemData(cboFondos.ListIndex)
    pFondoContrato = SIFGlobal.fxCodText(cboFondos.Text)
        
    cboFondos.RemoveItem cboFondos.ListIndex
    
    pDocumento = pFondoPlan & ".." & pFondoContrato
    mCuenta = fxgFNDCuentaPlan(1, pFondoPlan)
  
  Case "S" 'Saldos a Favor
    
    pDocumento = cboSaldoFavorReferencia.Text
    pSaldoFavorId = cboSaldoFavorReferencia.ItemData(cboSaldoFavorReferencia.ListIndex)
    'Elimina Opción del Combo
    cboSaldoFavorReferencia.RemoveItem cboSaldoFavorReferencia.ListIndex
    txtSaldoFavorMntOrigen.Text = 0
    txtSaldoFavorSaldo.Text = 0
    
    mCuenta = fxSaldoFavorCta(pSaldoFavorId)
End Select


'Fix del Tamaño del Texto del Documento

pDocumento = Mid(pDocumento, 1, 30)


'No. Documento Referencia para Todo el Documento
If Len(pDocumento) > 0 Then
    vAseDocDeposito = pDocumento
End If

'Valida Cuenta Contable
If Not fxgCntCuentaValida(mCuenta) Then
    Me.MousePointer = vbDefault
    MsgBox "La cuenta Contable de la Forma de Pago no es válida!!!", vbExclamation
    Exit Sub
End If


If mLinea = 0 Then
    
    strSQL = "select isnull(max(Linea),0) + 1 as 'Linea'" _
           & " from CAJAS_DESGLOCE_PAGO" _
           & " Where Cod_Caja = '" & ModuloCajas.mCaja & "' and cod_Apertura = " & ModuloCajas.mApertura _
           & " and Ticket = '" & ModuloCajas.mTiquete & "'"
    Call OpenRecordSet(rs, strSQL)
        mLinea = rs!Linea
    rs.Close

    strSQL = "insert CAJAS_DESGLOCE_PAGO(linea,Ticket,Cod_Caja,cod_Apertura,Monto,cod_Divisa,Tipo_Cambio,registro_fecha,registro_usuario" _
           & ",Cod_Tarjeta,Tarjeta_Numero,Tarjeta_Autorizacion,Cheque_Emisor,Cheque_Numero, Cuenta_Bancaria , Num_Referencia, Cod_Cuenta" _
           & ",Aplica_Saldo_Favor,Saldo_Favor,Saldo_Favor_Id,Observaciones,cod_forma_pago,DP_Banco,DP_Fecha" _
           & ", COD_PLAN,COD_CONTRATO, COD_ENTIDAD_PAGO, COD_ORIGEN_RECURSOS)" _
           & " Values(" & mLinea & ",'" & ModuloCajas.mTiquete & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & "," & curMonto & ",'" _
           & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & mTipoCambio & ", dbo.MyGetdate(),'" & ModuloCajas.mUsuario & "'"
    
    If mTipo = "T" Then
        strSQL = strSQL & ", '" & cboTarjetas.ItemData(cboTarjetas.ListIndex) & "', '" & Trim(txtTarjeta.Text) _
                    & "', '" & Trim(txtAutoriza.Text) & "'"
    Else
        strSQL = strSQL & ",'','',''"
    End If
    
    If mTipo = "C" Then
        strSQL = strSQL & ", '" & cboEmisorCheque.ItemData(cboEmisorCheque.ListIndex) & "', '" & Trim(txtCheque.Text) & "'"
    Else
        strSQL = strSQL & ", '', Null"
    End If
    
    strSQL = strSQL & ",'" & pBancoCuenta & "', '" & pDocumento & "','" & mCuenta & "', " & mAplSaldoFavor _
           & ",0," & pSaldoFavorId & ", '" & Mid(Trim(txtNotas.Text), 1, 490) & "', '" & mFormaPago & "'"
    
    If mTipo = "B" Then
        strSQL = strSQL & ", " & cboDepositoBanco.ItemData(cboDepositoBanco.ListIndex) & ", '" & Format(dtpDeposito.Value, "yyyy/mm/dd") & "'"
    Else
        strSQL = strSQL & ",0,Null"
    End If
    
    If mTipo = "F" Then
        strSQL = strSQL & ", '" & pFondoPlan & "', " & pFondoContrato
    
    Else
        strSQL = strSQL & ",Null,Null"
    End If
    
    If cboPagadores.Enabled Then
            strSQL = strSQL & ", '" & cboPagadores.ItemData(cboPagadores.ListIndex) & "', '" & cboOrigenRecursos.ItemData(cboOrigenRecursos.ListIndex) & "')"
    Else
        strSQL = strSQL & ", Null, Null)"
    End If
    
Else
   strSQL = "update CAJAS_DESGLOCE_PAGO set Monto = " & curMonto & ",cod_divisa = '" & cboDivisa.ItemData(cboDivisa.ListIndex) _
          & "', Tipo_Cambio = " & mTipoCambio & ", cod_cuenta = '" & mCuenta & "', Observaciones = '" _
          & Mid(Trim(txtNotas.Text), 1, 490) & "', Num_Referencia = '" & pDocumento & "', Cuenta_Bancaria = '" & pBancoCuenta _
          & "', Aplica_Saldo_Favor = " & mAplSaldoFavor & ",Saldo_Favor = 0, Saldo_Favor_ID = " & pSaldoFavorId _
          & ", cod_forma_pago = '" & mFormaPago & "'"
          
   
   If mTipo = "T" Then
        strSQL = strSQL & ", Cod_Tarjeta = '" & cboTarjetas.ItemData(cboTarjetas.ListIndex) & "', Tarjeta_Numero = '" _
               & Trim(txtTarjeta.Text) & "', Tarjeta_Autorizacion = '" & Trim(txtAutoriza.Text) & "'"
   End If
   
   If mTipo = "C" Then
        strSQL = strSQL & ", Cheque_Emisor = '" & cboEmisorCheque.ItemData(cboEmisorCheque.ListIndex) _
               & "', Cheque_Numero = '" & Trim(txtCheque.Text) & "'"
   End If
   
   If mTipo = "B" Then
      strSQL = strSQL & ", DP_BANCO = " & cboEmisorCheque.ItemData(cboEmisorCheque.ListIndex) & ", DP_FECHA = '" & Format(dtpDeposito.Value, "yyyy/mm/dd") & "'"
   End If
   
   If mTipo = "F" Then
      strSQL = strSQL & ", COD_PLAN = '" & pFondoPlan & "', COD_CONTRATO = " & pFondoContrato
   End If
   
    If cboPagadores.Enabled Then
            strSQL = strSQL & ", COD_ENTIDAD_PAGO = '" & cboPagadores.ItemData(cboPagadores.ListIndex) _
                   & "', COD_ORIGEN_RECURSOS = '" & cboOrigenRecursos.ItemData(cboOrigenRecursos.ListIndex) & "'"
    End If
   
   strSQL = strSQL & " Where Cod_Caja = '" & ModuloCajas.mCaja & "' and cod_Apertura = " & ModuloCajas.mApertura _
           & " and Ticket = '" & ModuloCajas.mTiquete & "' and Linea = " & mLinea
End If

Call ConectionExecute(strSQL)

'Distribuye Saldo a Favor (Cargar Tiquete Antes)
strSQL = "exec spCajas_DistribuyeSaldoFavor '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" _
       & ModuloCajas.mTiquete & "','" & ModuloCajas.mUsuario & "'," & ModuloCajas.mTotalAplicar & ",'" & ModuloCajas.mDivisa & "'"
Call ConectionExecute(strSQL)

'Consulta totales
Call sbCargaTiquete


mTipo = ""
Call sbApagaFrames("")


Me.MousePointer = vbDefault
Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxVerificaDatos() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, curMonto As Currency
Dim pDivisa As String, dDivisa As String, pTipoCambio As Currency

vMensaje = ""

If mTipo = "" Then vMensaje = vMensaje & "No se ha seleccionado ninguna forma de pago?" & vbCrLf

pDivisa = cboDivisa.ItemData(cboDivisa.ListIndex)

Call sbSIFCleanTxtInject(txtNotas)

If CCur(txtMontoDivisaFuncional.Text) = 0 Then
    vMensaje = vMensaje & "El Monto en Divisa Funcional no es válido..." & vbCrLf
End If


Select Case mTipo
   Case "E" 'Efectivo
        If CCur(txtMontoEfectivo.Text) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        
        dDivisa = lblDivisaInfo.Item(4).Caption
        
   Case "B" 'Depositos
        If CCur(txtMontoDeposito.Text) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        If Trim(txtDepositoNo.Text) = "" Then vMensaje = vMensaje & "Debe Digitar el número de documento..." & vbCrLf
        If cboDepositoBanco.ListCount = 0 Then vMensaje = vMensaje & "No existen cuentas bancarias configuradas para realizar el registro..." & vbCrLf
         
        '-- Resultados: 0 > No Existe, 1 > Existe / No Identificado, 2 > Existe Identificado
        '--, 3 > Existe Identificado a Otra Persona, 4 > Existe deposito Registrado pero el monto no es válido
        'create function fxTes_DP_Cargado(@Banco int, @Documento varchar(30), @Cedula varchar(20) = '', @Monto dec(13,2) ) returns smallint
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxTes_DP_Cargado(" & cboDepositoBanco.ItemData(cboDepositoBanco.ListIndex) _
                    & ",'" & Trim(txtDepositoNo.Text) & "','" & ModuloCajas.mClienteId _
                    & "'," & CCur(txtMontoDeposito.Text) & ") as Resultado"
            Call OpenRecordSet(rs, strSQL)
            
            mDP_Caso = rs!Resultado
            
            Select Case rs!Resultado
               Case 0 'No Existe
                If mDP_Proceso = 1 Then
                    vMensaje = vMensaje & "Se encuentra activado el control de Depositos y este Depósito no ha sido registrado." & vbCrLf
                End If
               Case 1 'Existe No Identificado
               Case 2 'Existe Identificado
                    vMensaje = vMensaje & "Este Depósito ya fué identificado. Busquelo como Saldo a Favor del Cliente" & vbCrLf
               Case 3 'Existe Identificado a Otra Persona
                    vMensaje = vMensaje & "Este Depósito Pertenece a Otra Persona!" & vbCrLf
               Case 4 'Existe Identificado a Otra Persona
                    vMensaje = vMensaje & "Este Depósito No concuerda en MONTO con el registrado en Control de Depósitos!" & vbCrLf
            End Select
            rs.Close
        End If
        
            
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_DocumentoVerifica('" & mFormaPago & "','" _
                    & Trim(txtDepositoNo.Text) & "','" & cboDepositoBanco.ItemData(cboDepositoBanco.ListIndex) & "','') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este Depósito ya se encuentra registrado (Verifique!)" & vbCrLf
            End If
            rs.Close
        End If
            
        'Verifica que no se haya procesado ya en el detalle de pago
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FP_Registada('" & ModuloCajas.mTiquete & "','" & mFormaPago & "','" _
                    & Trim(txtDepositoNo.Text) & "','" & cboDepositoBanco.ItemData(cboDepositoBanco.ListIndex) & "','') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este Deposito fue registrado como forma de pago de esta transacción (Verifique!)"
            End If
            rs.Close
        End If
            
        dDivisa = lblDivisaInfo.Item(5).Caption
            
            
   Case "D" 'Documentos
        If CCur(txtMontoDocumentos) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        If Trim(txtDocumento.Text) = "" Then vMensaje = vMensaje & "Debe Digitar el número de documento..." & vbCrLf
        If Not fxgCntCuentaValida(txtDocCuenta.Text) Then vMensaje = vMensaje & "La cuenta contable para el registro no es válida..." & vbCrLf
        
        
        
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_DocumentoVerifica('" & mFormaPago & "','" _
                    & Trim(txtDocumento.Text) & "','','') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este VALOR DE COMPROBANTE ya se encuentra registrado (Verifique!)" & vbCrLf
            End If
            rs.Close
        End If
            
        'Verifica que no se haya procesado ya en el detalle de pago
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FP_Registada('" & ModuloCajas.mTiquete & "','" & mFormaPago & "','" _
                    & Trim(txtDocumento.Text) & "','','') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Esta VALOR DE COMPROBANTE ya fue registrado como forma de pago de esta transacción (Verifique!)"
            End If
            rs.Close
        End If
        
        
        
        dDivisa = lblDivisaInfo.Item(1).Caption
         
        
   Case "T" 'Tarjetas
        If CCur(txtMontoTarjetas) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        If cboTarjetas.ListCount = 0 Then vMensaje = vMensaje & "No existen tipos de tarjetas configuradas..." & vbCrLf
        If Len(Trim(txtTarjeta.Text)) <= 8 Then vMensaje = vMensaje & "Debe Digitar el número de tarjeta válido..." & vbCrLf
        If Len(Trim(txtAutoriza.Text)) <= 3 Then vMensaje = vMensaje & "Debe Digitar el número de autorización válido..." & vbCrLf
   
        
        
        'Verifica que no se haya procesado ya en el detalle de pago
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FP_Registada('" & ModuloCajas.mTiquete & "','" & mFormaPago & "','" _
                    & cboTarjetas.ItemData(cboTarjetas.ListIndex) & "','" & Trim(txtTarjeta.Text) & "','" & Trim(txtAutoriza.Text) & "') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Esta TARJETA ya fue registrada como forma de pago de esta transacción (Verifique!)"
            End If
            rs.Close
        End If
        
        dDivisa = lblDivisaInfo.Item(3).Caption

   Case "C" 'Cheques
        If CCur(txtMontoCheques) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        If Trim(txtCheque.Text) = "" Then vMensaje = vMensaje & "Debe Digitar el número de Cheque..." & vbCrLf
        If Len(Trim(txtChequeCta.Text)) <= 3 Then vMensaje = vMensaje & "Debe Digitar la Cuenta Bancaria del Cheque..." & vbCrLf
        If cboEmisorCheque.ListCount = 0 Then vMensaje = vMensaje & "No Existe Lista de Emisor para este Cheque..." & vbCrLf
        
        
        
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_DocumentoVerifica('" & mFormaPago & "','" _
                    & Trim(txtCheque.Text) & "','" & cboEmisorCheque.ItemData(cboEmisorCheque.ListIndex) _
                    & "','" & txtChequeCta.Text & "') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este Cheque ya se encuentra registrado (Verifique!)"
            End If
            rs.Close
        End If
   
        'Verifica que no se haya procesado ya en el detalle de pago
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FP_Registada('" & ModuloCajas.mTiquete & "','" & mFormaPago & "','" _
                    & Trim(txtCheque.Text) & "','" & cboEmisorCheque.ItemData(cboEmisorCheque.ListIndex) _
                    & "','" & txtChequeCta.Text & "') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este Cheque fue registrado como forma de pago de esta transacción (Verifique!)"
            End If
            rs.Close
        End If
        dDivisa = lblDivisaInfo.Item(2).Caption
   
   
   Case "F" 'Fondos
        If CCur(txtMontoFondos.Text) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        If cboFondos.ListCount = 0 Then vMensaje = vMensaje & " - No Existe Fondo Referenciado?" & vbCrLf
        If CCur(txtMontoFondos.Text) > CCur(txtFondoDisponible.Text) Then vMensaje = vMensaje & " - El Monto Supera el disponible del Fondo!" & vbCrLf
        
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FondosDivisa('" & cboFondos.ItemData(cboFondos.ListIndex) _
                   & "') as Resultado"
            Call OpenRecordSet(rs, strSQL)
                dDivisa = Trim(rs!Resultado)
            rs.Close
        End If

   
        'Verifica que no se haya procesado ya en el detalle de pago
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FP_Registada('" & ModuloCajas.mTiquete & "','" & mFormaPago & "','" _
                    & cboFondos.ItemData(cboFondos.ListIndex) & ".." & SIFGlobal.fxCodText(cboFondos.Text) _
                    & "','" & cboFondos.ItemData(cboFondos.ListIndex) & "','" & SIFGlobal.fxCodText(cboFondos.Text) & "') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este PLAN DE AHORROS ya fue registrado como forma de pago de esta transacción (Verifique!)"
            End If
            rs.Close
        End If
        
   
   Case "S" 'Saldos a Favor
        If CCur(txtMontoSaldos.Text) <= 0 Then vMensaje = vMensaje & "El Monto no es válido..." & vbCrLf
        If cboSaldoFavorReferencia.ListCount = 0 Then vMensaje = vMensaje & " - No Existe Referencia al Saldo a Favor!" & vbCrLf
        If CCur(txtMontoSaldos.Text) > CCur(txtSaldoFavorSaldo.Text) Then vMensaje = vMensaje & " - El Monto Supera el Saldo Disponible!" & vbCrLf
        
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_SaldoFavorDivisa(" & cboSaldoFavorReferencia.ItemData(cboSaldoFavorReferencia.ListIndex) _
                   & ") as Resultado"
            Call OpenRecordSet(rs, strSQL)
                dDivisa = rs!Resultado
            rs.Close
        End If
        
        
        'Verifica que no se haya procesado ya en el detalle de pago
        If Len(vMensaje) = 0 Then
            strSQL = "select dbo.fxCajas_FP_Registada('" & ModuloCajas.mTiquete & "','" & mFormaPago & "','" _
                    & cboSaldoFavorReferencia.Text & "','" & cboSaldoFavorReferencia.ItemData(cboSaldoFavorReferencia.ListIndex) _
                    & "','') as Resultado"
            Call OpenRecordSet(rs, strSQL)
            If rs!Resultado > 0 Then
                    vMensaje = vMensaje & "Este SALDO A FAVOR ya fue registrado como forma de pago de esta transacción (Verifique!)"
            End If
            rs.Close
        End If

End Select

If Trim(pDivisa) <> Trim(dDivisa) Then vMensaje = vMensaje & "La divisa del movimiento no coincide con la divisa de pago de la transacción!..." & vbCrLf

curMonto = CCur(txtMontoDivisaTransaccion.Text)

If mAplSaldoFavor = 0 And mTipo <> "S" Then
   If (curMonto + ModuloCajas.mTotalCambio) > ModuloCajas.mTotalAplicar Then
        vMensaje = vMensaje & "El Monto es superior al requerido y no registra saldo a favor..." & vbCrLf
   End If
End If

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  fxVerificaDatos = False
Else
  fxVerificaDatos = True
End If

End Function


Private Sub TimerInicial_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

TimerInicial.Interval = 0
TimerInicial.Enabled = False

txtMontoEfectivo.Text = 0
txtMontoCheques.Text = 0
txtMontoTarjetas.Text = 0
txtMontoDocumentos.Text = 0
txtMontoSaldos.Text = 0
txtMontoDivisaFuncional.Text = 0
txtMontoDivisaTransaccion.Text = 0

txtMontoFondos.Text = 0
txtMontoDeposito.Text = 0

'Indica valida solo depositos previamente registrados en Tesoreria para Identificacion
If fxCajasParametros("10") = "S" Then
    mDP_Proceso = 1
Else
    mDP_Proceso = 0
End If

vPaso = True

strSQL = "Select rtrim(Cod_Divisa) as 'IdX', rtrim(descripcion) as 'itmx'" _
       & " from cntx_divisas " _
       & " where cod_Contabilidad = " & GLOBALES.gEnlace _
       & " order by cod_divisa"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)


strSQL = "select rTrim(cod_emisor) as 'IdX', rtrim(descripcion) as 'itmx' from sif_emisores where activo = 1"
Call sbCbo_Llena_New(cboEmisorCheque, strSQL, False, True)


strSQL = "select rTrim(cod_tarjeta) as 'IdX', rtrim(descripcion) as 'itmx' from sif_tarjetas where activa = 1"
Call sbCbo_Llena_New(cboTarjetas, strSQL, False, True)

''Lista los Pagadores de Cheques
'strSQL = "select rtrim(COD_ENTIDAD_PAGO) as 'IdX', rtrim(descripcion) as 'itmx'" _
'       & " from SIF_ENTIDADES_PAGO where ACTIVA = 1" _
'       & " order by COD_ENTIDAD_PAGO"
'Call sbCbo_Llena_New(cboPagador, strSQL, False, True)


'Identificacion de Recursos
strSQL = "select COD_ENTIDAD_PAGO as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ENTIDADES_PAGO" _
       & " WHERE ACTIVA = 1 ORDER BY COD_ENTIDAD_PAGO"
Call sbCbo_Llena_New(cboPagadores, strSQL, False, True)

strSQL = "select COD_ORIGEN_RECURSOS as 'IdX', DESCRIPCION AS 'ItmX' from SIF_ORIGEN_RECURSOS" _
       & "  WHERE ACTIVA = 1 ORDER BY COD_ORIGEN_RECURSOS"
Call sbCbo_Llena_New(cboOrigenRecursos, strSQL, False, True)

vPaso = True

'Carga Saldos a Favor (No utilizados en este Tiquet)
cboSaldoFavorReferencia.Clear

strSQL = "exec spCajas_FormaPago_SaldoFavor '" & ModuloCajas.mClienteId & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura _
       & ",'" & ModuloCajas.mTiquete & "'"
Call sbCbo_Llena_New(cboSaldoFavorReferencia, strSQL, False, True)

txtSaldoFavorMntOrigen.Text = "0"
txtSaldoFavorSaldo.Text = "0"


If ModuloCajas.mProductoCodigo = "" Then
    ModuloCajas.mProductoNumero = 0
End If

'Lista Fondos Disponibles
strSQL = "exec spCajas_FondosDisponiblePersona '" & ModuloCajas.mClienteId & "', '" & ModuloCajas.mProductoCodigo & "', " & ModuloCajas.mProductoNumero
Call sbCbo_Llena_New(cboFondos, strSQL, False, True)

 txtFondoDisponible.Text = 0
 txtFondoMonto.Text = 0
 txtMontoFondos.Text = 0


vPaso = False

Call cboSaldoFavorReferencia_Click
Call cboFondos_Click

'
'Call OpenRecordSet(rs, strSQL)
'Do While Not rs.EOF
' cboFondos.AddItem Trim(rs!cod_Plan) & " - " & Trim(rs!COD_CONTRATO & "")
' cboFondos.ItemData(cboFondos.NewIndex) = rs!COD_CONTRATO
'
' txtFondoDisponible.Text = 0
' txtFondoMonto.Text = 0
' txtMontoFondos.Text = 0
'
' rs.MoveNext
'Loop
''If rs.RecordCount > 0 Then
''   rs.MoveFirst
''   cboFondos.Text = Trim(rs!cod_Plan) & " - " & Trim(rs!cod_Contrato & "")
''End If
'rs.Close



'Lista de Formas de Pago
Call sbCargaFormasPago

'Tiquete Registrado (Temporal)
Call sbCargaTiquete


lblDivisa.Caption = ModuloCajas.mDivisa

txtTotalPagar.Text = Format(ModuloCajas.mTotalAplicar, "Standard")
txtTotalValores.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
txtTotal.Text = Format(ModuloCajas.mTotalDetallado, "Standard")

txtCambio.Text = 0
txtDiferencia.Text = Format(ModuloCajas.mTotalDiferencia, "Standard")

fraCheque.BackColor = RGB(140, 161, 213)
fraDepositos.BackColor = RGB(140, 161, 213)
fraDocumento.BackColor = RGB(140, 161, 213)
fraEfectivo.BackColor = RGB(140, 161, 213)
fraFondos.BackColor = RGB(140, 161, 213)
fraSaldos.BackColor = RGB(140, 161, 213)
fraTarjeta.BackColor = RGB(140, 161, 213)


vPaso = False

Me.MousePointer = vbDefault

Call cboDivisa_Click
End Sub


Private Sub tlbValores_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error GoTo vError

Select Case Button.Key
  Case "Borrar"
     With lswDetalle.ListItems
         For i = 1 To .Count
             If .Item(i).Checked Then
                strSQL = "delete CAJAS_DESGLOCE_PAGO" _
                       & " Where Cod_Caja = '" & ModuloCajas.mCaja & "' and cod_Apertura = " & ModuloCajas.mApertura _
                       & " and Ticket = '" & ModuloCajas.mTiquete & "' and Linea = " & .Item(i).Text
                Call ConectionExecute(strSQL)
             End If
         Next i
     End With
     Call sbCargaTiquete
     
     
  Case "Cerrar"
     fraDetalle.Visible = False


End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtAutoriza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoTarjetas.SetFocus
End Sub

Private Sub txtCheque_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEmisorCheque.SetFocus
End Sub



Private Sub txtChequeCta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCheque.SetFocus
End Sub

Private Sub txtDepositoNo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpDeposito.SetFocus
End Sub

Private Sub txtDocCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoDocumentos.SetFocus
If KeyCode = vbKeyF4 Then
   Call sbFormsCall("frmCntX_ConsultaCuentas", vbModal, 1, 1, False, Me)
   If gCuenta <> "" Then
       txtDocCuenta.Text = fxgCntCuentaFormato(True, gCuenta)
   End If
End If
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocCuenta.SetFocus
End Sub

Private Sub txtMontoCheques_Change()

On Error GoTo vError

Call sbCalculo(txtMontoCheques)

vError:

End Sub

Private Sub txtMontoCheques_GotFocus()
On Error GoTo vError
    txtMontoCheques.Text = CCur(txtMontoCheques)
vError:
End Sub

Private Sub txtMontoCheques_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCheque.SetFocus
End Sub

Private Sub txtMontoCheques_LostFocus()
On Error GoTo vError
    
    txtMontoCheques.Text = Format(CCur(txtMontoCheques), "Standard")
    Call sbCalculo(txtMontoCheques)

vError:
End Sub

Private Sub txtMontoDeposito_Change()

On Error GoTo vError

Call sbCalculo(txtMontoDeposito)

vError:

End Sub

Private Sub txtMontoDeposito_GotFocus()
On Error GoTo vError
    txtMontoDeposito = CCur(txtMontoDeposito.Text)
vError:
End Sub

Private Sub txtMontoDeposito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDepositoNo.SetFocus
End Sub

Private Sub txtMontoDeposito_LostFocus()
On Error GoTo vError

txtMontoDeposito.Text = Format(CCur(txtMontoDeposito.Text), "Standard")
Call sbCalculo(txtMontoDeposito)

vError:

End Sub

Private Sub txtMontoDocumentos_Change()

On Error GoTo vError

Call sbCalculo(txtMontoDocumentos)

vError:

End Sub

Private Sub txtMontoDocumentos_GotFocus()

On Error GoTo vError
    txtMontoDocumentos.Text = CCur(txtMontoDocumentos)
vError:


End Sub

Private Sub txtMontoDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub txtMontoDocumentos_LostFocus()
On Error GoTo vError

txtMontoDocumentos.Text = Format(CCur(txtMontoDocumentos.Text), "Standard")
Call sbCalculo(txtMontoDocumentos)

vError:
End Sub

Private Sub txtMontoEfectivo_Change()

On Error GoTo vError

Call sbCalculo(txtMontoEfectivo)

vError:

End Sub

Private Sub txtMontoEfectivo_GotFocus()
On Error GoTo vError
    txtMontoEfectivo.Text = CCur(txtMontoEfectivo.Text)
vError:
End Sub

Private Sub txtMontoEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotal.SetFocus
End Sub

Private Sub txtMontoEfectivo_LostFocus()
On Error GoTo vError
    txtMontoEfectivo.Text = Format(CCur(txtMontoEfectivo.Text), "Standard")
    Call sbCalculo(txtMontoEfectivo)
vError:
End Sub

Private Sub txtMontoFondos_Change()

On Error GoTo vError

Call sbCalculo(txtMontoFondos)

vError:

End Sub

Private Sub txtMontoFondos_GotFocus()
On Error GoTo vError
    txtMontoFondos = CCur(txtMontoFondos.Text)
vError:
End Sub

Private Sub txtMontoSaldos_Change()

On Error GoTo vError

Call sbCalculo(txtMontoSaldos)

vError:

End Sub

Private Sub txtMontoSaldos_GotFocus()
On Error GoTo vError
    txtMontoSaldos = CCur(txtMontoSaldos.Text)
vError:
End Sub

Private Sub txtMontoSaldos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSaldoFavorMntOrigen.SetFocus
End Sub

Private Sub txtMontoSaldos_LostFocus()
On Error GoTo vError
    txtMontoSaldos.Text = Format(CCur(txtMontoSaldos.Text), "Standard")
    Call sbCalculo(txtMontoSaldos)
vError:
End Sub

Private Sub txtMontoTarjetas_Change()

On Error GoTo vError

Call sbCalculo(txtMontoTarjetas)

vError:

End Sub

Private Sub txtMontoTarjetas_GotFocus()
On Error GoTo vError
    txtMontoTarjetas.Text = CCur(txtMontoTarjetas.Text)
vError:
End Sub

Private Sub txtMontoTarjetas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTarjeta.SetFocus
End Sub


Private Sub txtMontoTarjetas_LostFocus()
On Error GoTo vError
    txtMontoTarjetas.Text = Format(CCur(txtMontoTarjetas.Text), "Standard")
    Call sbCalculo(txtMontoTarjetas)
vError:

End Sub

Private Sub sbCalculo(pTxt As FlatEdit, Optional pAjusta As Boolean = False)
Dim curTotal As Currency, pDivisa As String
Dim strSQL As String, rs   As New ADODB.Recordset
Dim pTipoCambio As Currency, curDivisaTransaccion As Currency


'pDivisa            : Divisa del Pago
'pTipoCambio        : Tipo de Cambio de conversion  a Divisa Funcional
'mTipoCambio        : Tipo de Cambio de la Divisa de Pago
'lblDivisa.Caption  : Divisa de la Transaccion
'mDivisaFuncional   : Divisa Funcional (Local - Contable)


If Not IsNumeric(pTxt) Then
    MsgBox "Debe digitar solamente números...", vbExclamation
    pTxt.Text = 0
    txtMontoDivisaFuncional.Text = 0
    txtMontoDivisaTransaccion.Text = 0
    pTxt.SetFocus
Else
    txtMontoDivisaFuncional.Text = Format(CCur(pTxt.Text) * fxSys_Tipo_Cambio_Apl(mTipoCambio), "Standard")
End If





pDivisa = cboDivisa.ItemData(cboDivisa.ListIndex)

'curTotal = (CCur(txtMontoEfectivo) + CCur(txtMontoTarjetas) + CCur(txtMontoDocumentos) _
'         + CCur(txtMontoCheques) + CCur(txtMontoSaldos.Text) + CCur(txtMontoDeposito) + CCur(txtMontoFondos.Text))

curTotal = CCur(pTxt.Text)

'Divisa Funcional - Siempre aplica el Tipo de Cambio (Del Pago)
txtMontoDivisaFuncional.Text = Format(curTotal * fxSys_Tipo_Cambio_Apl(mTipoCambio), "Standard")


Select Case True
   'Ejemplo: Pago en Colones (funcional) / Operación en Colones (Funcional)
   Case pDivisa = mDivisaFuncional _
        And lblDivisa.Caption = pDivisa
      
        'Nada

   'Ejemplo: Pago en Colones (funcional) / Operación en Dolares (Foranea)
   Case pDivisa = mDivisaFuncional _
        And lblDivisa.Caption <> mDivisaFuncional

        strSQL = "select dbo.fxCajas_TipoCambio(" & GLOBALES.gEnlace & ",'" & lblDivisa.Caption & "',dbo.MyGetdate(),'C') as 'TipoCambio'"
        Call OpenRecordSet(rs, strSQL)
          pTipoCambio = rs!TipoCambio
        rs.Close

        curTotal = curTotal / fxSys_Tipo_Cambio_Apl(pTipoCambio)

   'Ejemplo: Pago en Dolares (Foranea) / Operación en Colones (funcional)
   Case pDivisa <> mDivisaFuncional _
        And lblDivisa.Caption = mDivisaFuncional

        curTotal = curTotal * fxSys_Tipo_Cambio_Apl(mTipoCambio)


   'Ejemplo: Pago en (Foranea del mismo tipo de la Transaccion) / Operación en (Foranea)
   Case pDivisa <> mDivisaFuncional _
        And lblDivisa.Caption <> mDivisaFuncional _
        And pDivisa = lblDivisa.Caption

        'Nada
'        MsgBox ""

   'Ejemplo: Operación en Foranea / Pago en (Foranea) Distinta a la Transaccion NO PERMITIDO!
   Case pDivisa <> mDivisaFuncional _
        And lblDivisa.Caption <> mDivisaFuncional _
        And pDivisa <> lblDivisa.Caption

        curTotal = 0
        txtMontoDivisaFuncional.Text = 0
        txtMontoDivisaTransaccion.Text = 0

End Select


txtMontoDivisaTransaccion.Text = Format(curTotal, "Standard")

txtTotal.Text = Format(ModuloCajas.mTotalDetallado + curTotal, "Standard")

txtDiferencia.Text = Format(ModuloCajas.mTotalAplicar - (ModuloCajas.mTotalDetallado + curTotal), "Standard")


If CCur(txtDiferencia) < 0 Then
   txtCambio.Text = Abs(txtDiferencia.Text)
Else
   txtCambio.Text = 0
End If


'Ajusta al Total Requerido (Aplica para todas los que no generan saldo a favor
If pAjusta Then

                If (curTotal + ModuloCajas.mTotalCambio) > ModuloCajas.mTotalAplicar Then
                    curTotal = ModuloCajas.mTotalAplicar - ModuloCajas.mTotalCambio
                    pTxt.Text = curTotal
                End If

                Call sbCalculo(pTxt, False)
End If

End Sub

    
Private Sub sbCargaFormasPago()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

lsw.ListItems.Clear

vPaso = True

strSQL = "select F.*" _
       & " from CAJAS_FORMAS_PAGO C inner join SIF_FORMAS_PAGO F on  C.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
       & " Where C.cod_caja = '" & ModuloCajas.mCaja _
       & "' order by F.EFECTIVO desc, F.tipo asc, F.COD_FORMA_PAGO asc"
    
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
     itmX.SubItems(1) = rs!Cod_Forma_Pago
     itmX.ListSubItems.Item(1).Tag = rs!cod_cuenta
     itmX.SubItems(2) = rs!Aplica_saldos_Favor
     itmX.SubItems(3) = rs!OR_APLICA
     
     itmX.Tag = rs!Tipo

 rs.MoveNext
Loop
rs.Close

vPaso = False

If lsw.ListItems.Count > 0 Then
    lsw.ListItems.Item(1).Selected = True
End If
End Sub


Private Sub sbCargaTiquete()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, pTotal As Currency
Dim pTipoCambio As Currency

'Tipo de Cambio de la Divisa de la Transaccion
pTipoCambio = CCur(lblTCDivisaTransaccion.Caption)

pTotal = 0

txtMontoEfectivo.Text = 0
txtMontoCheques.Text = 0
txtMontoDocumentos.Text = 0
txtMontoDeposito.Text = 0
txtMontoFondos.Text = 0
txtMontoSaldos.Text = 0
txtMontoTarjetas.Text = 0

lswDetalle.ListItems.Clear

strSQL = "select C.*, F.DESCRIPCION as 'FormaPagoDesc',F.TIPO, D.descripcion as 'Divisa'" _
       & ", dbo.fxCajas_TipoCambio(" & GLOBALES.gEnlace & ",C.cod_divisa,dbo.MyGetdate(),'C') as 'TipoCambio'" _
       & " from CAJAS_DESGLOCE_PAGO C inner join SIF_FORMAS_PAGO F on  C.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
       & " inner join CNTX_Divisas D on C.cod_Divisa = D.cod_Divisa and D.cod_Contabilidad = " & GLOBALES.gEnlace _
       & " Where C.cod_caja = '" & ModuloCajas.mCaja & "' and C.Ticket = '" & ModuloCajas.mTiquete _
       & "' and C.Cod_Apertura = " & ModuloCajas.mApertura
    
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswDetalle.ListItems.Add(, , rs!Linea)
     itmX.SubItems(1) = rs!FormaPagoDesc
     itmX.SubItems(2) = Format(rs!Monto / fxSys_Tipo_Cambio_Apl(rs!TipoCambio), "Standard")
     itmX.SubItems(3) = Format(rs!Saldo_Favor / fxSys_Tipo_Cambio_Apl(rs!TipoCambio), "Standard")
     itmX.SubItems(4) = rs!Divisa
     itmX.SubItems(5) = "Ref .: " & rs!Num_Referencia & " ¦ CK: " & rs!Cheque_Numero & " ¦ Tarjeta: " & rs!tarjeta_Numero _
                      & "¦ Fondos:" & rs!COD_PLAN & " Cnt:" & rs!COD_CONTRATO
     itmX.SubItems(6) = rs!TipoCambio
     itmX.SubItems(7) = rs!Monto
    

    If Trim(rs!cod_Divisa) = lblDivisa.Caption Then
        pTotal = pTotal + ((rs!Monto / fxSys_Tipo_Cambio_Apl(rs!TipoCambio)) - (rs!Saldo_Favor / fxSys_Tipo_Cambio_Apl(rs!TipoCambio)))
    Else
        pTotal = pTotal + ((rs!Monto / fxSys_Tipo_Cambio_Apl(pTipoCambio)) - (rs!Saldo_Favor / fxSys_Tipo_Cambio_Apl(pTipoCambio)))
    End If
        
      
 rs.MoveNext
Loop
rs.Close

ModuloCajas.mTotalDetallado = pTotal
ModuloCajas.mTotalCambio = 0
ModuloCajas.mTotalDiferencia = ModuloCajas.mTotalAplicar - pTotal



txtTotalPagar.Text = Format(ModuloCajas.mTotalAplicar, "Standard")
txtTotalValores.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
txtTotal.Text = Format(ModuloCajas.mTotalDetallado, "Standard")

txtCambio.Text = 0
txtDiferencia.Text = Format(ModuloCajas.mTotalDiferencia, "Standard")

End Sub

Private Sub sbApagaFrames(vFrame As String)
 
 Dim Control As Object
  
  For Each Control In Me.Controls
        If TypeOf Control Is Frame Then
          If Control.Name <> vFrame Then
                Control.Visible = False
         End If
        End If
    Next
End Sub


Private Sub txtSaldoFavorMntOrigen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSaldoFavorSaldo.SetFocus
End Sub


Private Sub txtSaldoFavorSaldo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMontoSaldos.SetFocus
End Sub

Private Sub txtTarjeta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAutoriza.SetFocus
End Sub
