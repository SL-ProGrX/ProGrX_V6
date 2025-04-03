VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_AnulaAbonosNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulación de Movimientos"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "frmCR_AnulaAbonosNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10725
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2292
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   10452
      _Version        =   1572864
      _ExtentX        =   18436
      _ExtentY        =   4043
      _StockProps     =   77
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
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   210
      Left            =   240
      TabIndex        =   12
      Top             =   1840
      Width           =   210
      _Version        =   1572864
      _ExtentX        =   370
      _ExtentY        =   370
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3960
      TabIndex        =   15
      Top             =   720
      Width           =   5412
      _Version        =   1572864
      _ExtentX        =   9546
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3960
      TabIndex        =   16
      Top             =   1080
      Width           =   5412
      _Version        =   1572864
      _ExtentX        =   9546
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2160
      TabIndex        =   14
      Top             =   1080
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2160
      TabIndex        =   13
      Top             =   720
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   312
      Left            =   9360
      TabIndex        =   17
      Top             =   1080
      Width           =   1092
      _Version        =   1572864
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   312
      Left            =   9360
      TabIndex        =   18
      Top             =   720
      Width           =   1092
      _Version        =   1572864
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtIntCor 
      Height          =   312
      Left            =   1920
      TabIndex        =   22
      Top             =   4920
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtIntMor 
      Height          =   312
      Left            =   1920
      TabIndex        =   23
      Top             =   5280
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
      Height          =   312
      Left            =   1920
      TabIndex        =   24
      Top             =   5640
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   312
      Left            =   7080
      TabIndex        =   25
      Top             =   5280
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtCargos 
      Height          =   312
      Left            =   7080
      TabIndex        =   26
      Top             =   4920
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   312
      Left            =   7080
      TabIndex        =   27
      Top             =   5640
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtAbIntCor 
      Height          =   312
      Left            =   3480
      TabIndex        =   28
      Top             =   4920
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbIntMor 
      Height          =   312
      Left            =   3480
      TabIndex        =   29
      Top             =   5280
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbAmortizacion 
      Height          =   312
      Left            =   3480
      TabIndex        =   30
      Top             =   5640
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbPoliza 
      Height          =   312
      Left            =   8640
      TabIndex        =   31
      Top             =   5280
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbCargos 
      Height          =   312
      Left            =   8640
      TabIndex        =   32
      Top             =   4920
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtAbTotal 
      Height          =   312
      Left            =   8640
      TabIndex        =   33
      Top             =   5640
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbAnulacion 
      Height          =   1932
      Left            =   240
      TabIndex        =   34
      Top             =   6120
      Width           =   10332
      _Version        =   1572864
      _ExtentX        =   18224
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Datos de la anulación:"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboAccion 
         Height          =   312
         Left            =   2040
         TabIndex        =   35
         Top             =   600
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2990
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   2040
         TabIndex        =   36
         Top             =   960
         Width           =   5772
         _Version        =   1572864
         _ExtentX        =   10181
         _ExtentY        =   1397
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboUltCtaCancelada 
         Height          =   312
         Left            =   6120
         TabIndex        =   37
         Top             =   600
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2990
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton cmdAnular 
         Height          =   612
         Left            =   8040
         TabIndex        =   41
         Top             =   1080
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "&Anular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmCR_AnulaAbonosNew.frx":000C
      End
      Begin XtremeSuiteControls.CheckBox chkRecalculaCuota 
         Height          =   216
         Left            =   8040
         TabIndex        =   42
         Top             =   600
         Width           =   1776
         _Version        =   1572864
         _ExtentX        =   3133
         _ExtentY        =   381
         _StockProps     =   79
         Caption         =   "Recalcular Cuota"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.Label lblRecomendada 
         Height          =   252
         Left            =   6120
         TabIndex        =   45
         Top             =   240
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "..."
         ForeColor       =   8421504
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
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Recomendada.:"
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
         Height          =   252
         Index           =   0
         Left            =   3720
         TabIndex        =   44
         Top             =   240
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Anula y Procesar..:"
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
         Left            =   -360
         TabIndex        =   40
         Top             =   600
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   3
         Left            =   -360
         TabIndex        =   39
         Top             =   960
         Width           =   2292
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ult.Cta. Cancelada.:"
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
         Left            =   3720
         TabIndex        =   38
         Top             =   600
         Width           =   2292
      End
   End
   Begin XtremeShortcutBar.ShortcutCaption lblCapMov 
      Height          =   372
      Left            =   120
      TabIndex        =   43
      Top             =   1800
      Width           =   10452
      _Version        =   1572864
      _ExtentX        =   18436
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Movimientos Registrados"
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   2
      Left            =   840
      TabIndex        =   21
      Top             =   1080
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Línea"
      ForeColor       =   16777215
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
      Height          =   252
      Index           =   1
      Left            =   840
      TabIndex        =   20
      Top             =   720
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Identificación"
      ForeColor       =   16777215
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
      Height          =   372
      Index           =   0
      Left            =   840
      TabIndex        =   19
      Top             =   120
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Operación"
      ForeColor       =   16777215
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
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Póliza"
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
      Height          =   312
      Index           =   10
      Left            =   5760
      TabIndex        =   9
      Top             =   5280
      Width           =   1332
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Total Anulación"
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
      Height          =   312
      Index           =   8
      Left            =   5760
      TabIndex        =   8
      Top             =   5640
      Width           =   1332
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Cargos"
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
      Height          =   312
      Index           =   7
      Left            =   5760
      TabIndex        =   7
      Top             =   4920
      Width           =   1332
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Anulación"
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
      Index           =   6
      Left            =   8640
      TabIndex        =   6
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Originales"
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
      Index           =   0
      Left            =   7080
      TabIndex        =   5
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Originales"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos Anulación"
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
      Index           =   5
      Left            =   3480
      TabIndex        =   3
      Top             =   4560
      Width           =   1572
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Int.Corriente"
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
      Height          =   312
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   4920
      Width           =   1332
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Int.Morosidad"
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
      Height          =   312
      Index           =   3
      Left            =   600
      TabIndex        =   1
      Top             =   5280
      Width           =   1332
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Amortización"
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
      Height          =   312
      Index           =   4
      Left            =   600
      TabIndex        =   0
      Top             =   5640
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1575
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmCR_AnulaAbonosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vLlave As Long

Dim vOperacion As Long
Dim vInteres As Currency, vPlazo As Integer, vSaldoMes As Currency, vUltimoRecibo As Long
Dim vRetencion As Boolean, vBaseCalculo As String, vPrideduc As Long
Dim vDiasActivo As Long, vPaso As Boolean

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub



Private Sub cmdAnular_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
Dim vIntM As Currency, vIntC As Currency, vAMORTIZAm As Currency
Dim vANUIntM As Currency, vANUIntC As Currency, vANUAMORTIZAm As Currency
Dim vCuenta As String, vFecha As Date, vFP_SF As String
Dim vTipoDoc As String, vNumDoc As String
Dim lngOperacion As Long

On Error GoTo vError

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
  Exit Sub
End If


  If Not fxValidaInformacion Then
    MsgBox "La información suministrada no es válida...", vbCritical
    Exit Sub
  End If 'Validacion
  
'Valida que no se ajuste un Recaudo de Ahorros
strSQL = "select dbo.fxCrd_Operacion_Recaudo_Ahorro('" & txtCodigo.Text & "') as 'Valida'"
Call OpenRecordSet(rs, strSQL)
If rs!Valida = 0 Then
    MsgBox "No se pueden realizar este tipo de movimientos a Recaudos de Ahorros Extraordinarios, debe aplicarlos directamente al Plan de Ahorros de la persona.", vbExclamation
    rs.Close
    Exit Sub
End If
rs.Close
    
    
vTipoDoc = "ND"
vCuenta = ""
vFP_SF = ""

 Select Case cboAccion.ItemData(cboAccion.ListIndex)
  Case "C" 'Cuenta Contable
       vCuenta = Trim(fxDocumentoCuenta(vTipoDoc))
       
       If vAseDocValido = False Then
         MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
              & " válida para esta operación...", vbCritical
         Exit Sub
       End If
   
       txtNotas.Text = vAseDocDetalle
       
   Case "S" 'Saldo a favor en Cajas
       strSQL = "select COD_FORMA_PAGO, COD_CUENTA " _
              & " From SIF_FORMAS_PAGO" _
              & " where TIPO = 'S' and Activa = 1"
      Call OpenRecordSet(rs, strSQL)
        vCuenta = Trim(rs!cod_cuenta)
        vFP_SF = Trim(rs!Cod_Forma_Pago)
      rs.Close
 End Select
    
 Call sbSIFCleanTxtInject(txtNotas)
   
    
 vFecha = fxFechaServidor
 vNumDoc = fxDocumentoConsecutivo(vTipoDoc)
   
 lngOperacion = txtOperacion
   
 Me.MousePointer = vbHourglass
  
 Call sbDocumento(vTipoDoc, vNumDoc, CCur(txtABIntCor.Text), CCur(txtABIntMor.Text), CCur(txtABAmortizacion.Text) _
                            , CCur(txtABCargos.Text), CCur(txtAbPoliza.Text), vCuenta, vFP_SF, "ANULA ABONO")
    
    strSQL = "exec spCrdPlanPagoAnulaAbono " & lngOperacion & ",'CRD008','" & glogon.Usuario & "','" & vTipoDoc & "','" & vNumDoc & "',1," & CCur(txtABIntCor) _
           & "," & CCur(txtABIntMor.Text) & "," & CCur(txtABAmortizacion.Text) & "," & CCur(txtABCargos.Text) & "," & CCur(txtPoliza.Text) _
           & ",'" & Format(vFecha, "yyyy/mm/dd") & "','',1," & chkRecalculaCuota.Value _
           & "," & cboUltCtaCancelada.ItemData(cboUltCtaCancelada.ListIndex) _
           & ",'" & txtNotas.Text & "'"
    Call ConectionExecute(strSQL)

 Call Bitacora("Anula", "OP: " & txtOperacion & " Doc.:" & vNumDoc & " Total : " & CCur(txtABTotal) & " Rec.Cuota.:" & chkRecalculaCuota.Value)
 
 
 Call sbTrazabilidad_Inserta("06", vNumDoc, vNumDoc)
 
 
 Me.MousePointer = vbDefault
 
 MsgBox "Anulación Realizada ... Con Nota Debito:" & vNumDoc, vbInformation

 Call sbImprimeRecibo(vNumDoc, vTipoDoc)




Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Function fxValidaInformacion() As Boolean
 
 fxValidaInformacion = True
 
 If Len(Trim(txtABIntCor)) = 0 Then
   txtABIntCor = 0
 End If
 
 If Len(Trim(txtABCargos)) = 0 Then
   txtABCargos = 0
 End If
 
  If Len(Trim(txtAbPoliza)) = 0 Then
   txtAbPoliza = 0
 End If
 
 If Len(Trim(txtABIntMor)) = 0 Then
   txtABIntMor = 0
 End If
 
 If Len(Trim(txtABAmortizacion)) = 0 Then
   txtABAmortizacion = 0
 End If
 
 If Len(Trim(txtABTotal)) = 0 Then
   txtABTotal = 0
 End If
 
 
  If (CCur(txtABAmortizacion) + CCur(txtABIntCor) + CCur(txtABIntMor) + CCur(txtCargos) + CCur(txtPoliza)) = 0 Then
    fxValidaInformacion = False
 End If



 If Len(Trim(txtNombre.Text)) = 0 Then
    fxValidaInformacion = False
 End If

End Function


Private Sub sbDocumento(vTipoDoc As String, vNumDoc As String, curIntC As Currency, curIntM As Currency, curAmortiza As Currency, curCargo As Currency _
                                , curPoliza As Currency, vCuenta As String, vFP_SF As String, vDetalle As String)

Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuentaPoliza As String, vDivisa As String, vUnidad As String
Dim rsTmp As New ADODB.Recordset

  
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

strLinea(1) = "Saldo Actual      " & Format(rs!Saldo, "Standard")
strLinea(2) = "Interes Corriente " & Format(curIntC * -1, "Standard")
strLinea(3) = "Interes Moratorio " & Format(curIntM * -1, "Standard")
strLinea(4) = "Amortización      " & Format(curAmortiza * -1, "Standard")
strLinea(5) = "Cargos            " & Format(curCargo * -1, "Standard")
strLinea(6) = "Póliza            " & Format(curPoliza, "Standard")
strLinea(7) = "Nuevo Saldo       " & Format(rs!Saldo + curAmortiza, "Standard")
strLinea(8) = "Operación /Linea  " & txtOperacion & "_" & txtCodigo.Text & "_" & UCase(txtOpex.Text)
strLinea(9) = "Proc.Retencion    " & IIf(vRetencion, "SI", "NO")
strLinea(10) = "Usuario           " & glogon.Usuario
strLinea(11) = "Fecha Ult. Cta    " & cboUltCtaCancelada.Text

 If curPoliza > 0 Then
   strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!ID_SOLICITUD & ") as 'Cuenta'"
   Call OpenRecordSet(rsTmp, strSQL, 0)
     vCuentaPoliza = Trim(rsTmp!Cuenta)
   rsTmp.Close
 End If


strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle)" _
        & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
        & "','" & Trim(txtNombre.Text) & "','CRD008'," & curIntC + curIntM + curAmortiza + curCargo + curPoliza & ",'P','" & txtOperacion.Text _
        & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & strLinea(11) & "','" & Trim(txtNotas.Text) & vbCrLf & "Depósito..:" & vAseDocDeposito & "')"

'ASIENTO
If curIntC > 0 Then
  strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntC & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
End If

If curIntM > 0 Then
  strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntM & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
End If

If curCargo > 0 Then
  strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curCargo & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!CtaCargos _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
End If

If curPoliza > 0 Then
   strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curPoliza & ",'D','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuentaPoliza _
          & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
End If
 
If curAmortiza > 0 Then
  strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curAmortiza & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
         & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
End If

'Corte Lote No 1
Call ConectionExecute(strSQL)



'Cierre de Movimiento:
If curIntC + curIntM + curAmortiza + curCargo + curPoliza > 0 Then
        strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntC + curIntM + curCargo + curAmortiza + curPoliza & ",'C','" & rs!cod_Divisa _
               & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
               & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        Call ConectionExecute(strSQL)
                
        vDivisa = rs!cod_Divisa
        vUnidad = rs!Cod_Unidad
                
        Select Case cboAccion.ItemData(cboAccion.ListIndex)
            Case "C" 'Cuenta
                'Nada

            
            Case "S" 'Saldo a Favor
                 strSQL = "exec spCajas_SaldoFavor_Registra '" & vFP_SF & "','" & vTipoDoc & "-" & vNumDoc & "'," & curIntC + curIntM + curCargo + curAmortiza + curPoliza _
                        & ",'" & txtCedula.Text & "','" & txtNombre.Text & "','" & glogon.Usuario & "','" & rs!cod_Divisa & "'"
               
                 Call OpenRecordSet(rs, strSQL)
                 
                 'Insertar Format de Pago
                strSQL = "exec spSYS_Anulacion_Saldo_Favor '" & vTipoDoc & "','" & vNumDoc & "','" & glogon.Usuario _
                        & "','" & vFP_SF & "','" & vDivisa & "'," & curIntC + curIntM + curCargo + curAmortiza + curPoliza _
                        & ",'" & vUnidad & "','" & vCuenta & "','" & vTipoDoc & "-" & vNumDoc _
                        & "'," & rs!SF_ID
                
                Call ConectionExecute(strSQL)
        End Select
End If

rs.Close

End Sub

Private Sub Form_Activate()
 vModulo = 3
End Sub

Private Sub Form_Load()
 vModulo = 3
 
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
cboAccion.Clear
cboAccion.AddItem "Cuenta Contable"
cboAccion.ItemData(cboAccion.ListCount - 1) = "C"
cboAccion.AddItem "Saldo a Favor"
cboAccion.ItemData(cboAccion.ListCount - 1) = "S"
cboAccion.Text = "Saldo a Favor"
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1000
    .Add , , "Proceso", 1000, vbCenter
    .Add , , "Cuota", 1800, vbRightJustify
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Dias Cor.", 1100, vbCenter
    .Add , , "Dias Mor.", 1100, vbCenter

    .Add , , "Tipo Doc.", 1100, vbCenter
    .Add , , "No. Doc.", 2100
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 2100, vbCenter

End With

lsw.Checkboxes = True

 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Private Sub sbLimpia()
    
    txtOperacion = ""
    txtCedula = ""
    txtNombre.Text = ""
    txtCodigo = ""
    txtDescripcion.Text = ""
    
    txtABIntCor.Text = "0"
    txtABAmortizacion.Text = "0"
    txtABIntMor.Text = "0"
    txtABCargos.Text = "0"
    txtAbPoliza.Text = "0"
    txtABTotal.Text = "0"
    
    txtIntCor.Text = "0"
    txtIntMor.Text = "0"
    txtAmortizacion.Text = "0"
    txtCargos.Text = "0"
    txtPoliza.Text = "0"
    txtTotal.Text = "0"
    
    vOperacion = 0
    vPlazo = 0
    vInteres = 0
    vSaldoMes = 0
    
    lsw.ListItems.Clear

End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyDown(vbKeyReturn, 1)
End Sub



Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, i As Integer
Dim vTemp As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass


    
strSQL = "select R.id_solicitud,R.saldo,R.proceso" _
       & ",R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult,R.Prideduc" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas" _
       & ",S.nombre,C.descripcion,C.retencion,C.poliza,R.fechaforp,R.Base_Calculo" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado in('A','C') and R.ID_SOLICITUD = " & txtOperacion.Text
       
Call sbLimpia
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  vBaseCalculo = Trim(rs!Base_Calculo)
  vPrideduc = rs!PriDeduc
  vOperacion = rs!ID_SOLICITUD
  txtOperacion.Text = rs!ID_SOLICITUD
  vPlazo = rs!Plazo
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  
  cboUltCtaCancelada.Clear
  
  
  vTemp = fxFechaProcesoSiguiente(rs!FecUlt)
  For i = 1 To 6
        vTemp = fxFechaProcesoAnterior(vTemp)
        If vTemp >= rs!PriDeduc Then
                cboUltCtaCancelada.AddItem Format(vTemp, "####-##")
                cboUltCtaCancelada.ItemData(cboUltCtaCancelada.ListCount - 1) = CStr(vTemp)
        End If
  Next i
  Call sbCboAsignaDato(cboUltCtaCancelada, Format(rs!FecUlt, "####-##"), True, CStr(rs!FecUlt))
  
  vSaldoMes = rs!Saldo
  
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    txtCodigo.Text = rs!Codigo
    
    txtProceso.Tag = rs!Proceso
    
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
    
    
    If (rs!opex & "") = 1 Then
        txtOpex.Text = "Op.Ex."
    Else
        txtOpex.Text = "Interno"
    End If
    
    txtDescripcion.Text = rs!Descripcion
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    'Movimientos Registrados
    strSQL = "select * from CRD_OPERACION_TRANSAC" _
           & " where estado = 'C' and id_solicitud = " & rs!ID_SOLICITUD _
           & " and Tipo_Documento not in('AJ')" _
           & " and Mov_Monto > 0" _
           & " order by id_seq desc"
      
    rs.Close
    
    vPaso = True
    chkTodos.Value = vbUnchecked
    
    Call OpenRecordSet(rs, strSQL)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Num_Cuota)
          itmX.SubItems(1) = Format(rs!Fecha_Proceso, "####-##")
          itmX.SubItems(2) = Format(rs!Cuota, "Standard")
          itmX.SubItems(3) = IIf((rs!Mora_Dias > 0), "En Mora", "Al Día")
          itmX.SubItems(4) = Format(rs!Mov_IntCor, "Standard")
          itmX.SubItems(5) = Format(rs!Mov_IntMor, "Standard")
          itmX.SubItems(6) = Format(rs!Mov_Principal, "Standard")
          itmX.SubItems(7) = Format(rs!Mov_Cargos, "Standard")
          itmX.SubItems(8) = Format(rs!Mov_Poliza, "Standard")
          itmX.SubItems(9) = rs!Dias_calculo
          itmX.SubItems(10) = rs!Mora_Dias
          
          itmX.SubItems(11) = rs!Tipo_documento & ""
          itmX.SubItems(12) = rs!Num_Comprobante & ""
          itmX.SubItems(13) = rs!Mov_fecha & ""
          itmX.SubItems(14) = rs!Mov_usuario & ""
          
          
          
          
          itmX.Tag = rs!Id_seq
      rs.MoveNext
    Loop
    
    vPaso = False
Else
    MsgBox "No se encontró operación la operación...!", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbCalculaTotal()
On Error GoTo vError
  txtABTotal.Text = Format(CCur(txtABIntCor.Text) + CCur(txtABIntMor.Text) + CCur(txtABAmortizacion.Text) _
                  + CCur(txtABCargos.Text) + CCur(txtPoliza.Text), "Standard")
vError:
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

If Item.Checked Then
  txtIntCor.Text = CCur(txtIntCor.Text) + CCur(Item.SubItems(4))
  txtIntMor.Text = CCur(txtIntMor.Text) + CCur(Item.SubItems(5))
  txtAmortizacion.Text = CCur(txtAmortizacion.Text) + CCur(Item.SubItems(6))
  txtCargos.Text = CCur(txtCargos.Text) + CCur(Item.SubItems(7))
  txtPoliza.Text = CCur(txtPoliza.Text) + CCur(Item.SubItems(8))

Else
  txtIntCor.Text = CCur(txtIntCor.Text) - CCur(Item.SubItems(4))
  txtIntMor.Text = CCur(txtIntMor.Text) - CCur(Item.SubItems(5))
  txtAmortizacion.Text = CCur(txtAmortizacion.Text) - CCur(Item.SubItems(6))
  txtCargos.Text = CCur(txtCargos.Text) - CCur(Item.SubItems(7))
  txtPoliza.Text = CCur(txtPoliza.Text) - CCur(Item.SubItems(8))
End If

txtTotal.Text = CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtAmortizacion.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text)

txtIntCor.Text = Format(txtIntCor.Text, "Standard")
txtIntMor.Text = Format(txtIntMor.Text, "Standard")
txtAmortizacion.Text = Format(txtAmortizacion.Text, "Standard")
txtCargos.Text = Format(txtCargos.Text, "Standard")
txtPoliza.Text = Format(txtPoliza.Text, "Standard")

txtTotal.Text = Format(txtTotal.Text, "Standard")

txtABIntCor.Text = Format(txtIntCor.Text, "Standard")
txtABIntMor.Text = Format(txtIntMor.Text, "Standard")
txtABAmortizacion.Text = Format(txtAmortizacion.Text, "Standard")
txtABCargos.Text = Format(txtCargos.Text, "Standard")
txtAbPoliza.Text = Format(txtPoliza.Text, "Standard")
txtABTotal.Text = Format(txtTotal.Text, "Standard")


End Sub

Private Sub sbCtaRecomendada()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

lblRecomendada.Caption = "..."

If Not IsNumeric(txtOperacion.Text) Then Exit Sub

If CCur(txtABAmortizacion.Text) > 0 Then
    strSQL = "select dbo.fxCrd_Operacion_Anula_Cta_Recomendada(" & txtOperacion.Text _
           & "," & CCur(txtABAmortizacion.Text) & ")  as 'Cta'"
    
    Call OpenRecordSet(rs, strSQL)
    
    lblRecomendada.Caption = rs!Cta
    
    rs.Close
End If

Exit Sub

vError:
    lblRecomendada.Caption = "..."
End Sub

Private Sub txtAbAmortizacion_Change()

Call sbCtaRecomendada

End Sub

Private Sub txtABPoliza_GotFocus()
On Error GoTo vError
    txtAbPoliza.Text = CCur(txtAbPoliza.Text)
vError:
End Sub

Private Sub txtABPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABCargos.SetFocus
End Sub

Private Sub txtABPoliza_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABPoliza_LostFocus()
On Error GoTo vError
    txtAbPoliza.Text = Format(CCur(txtAbPoliza.Text), "Standard")
vError:
End Sub

Private Sub txtABIntCor_GotFocus()
On Error GoTo vError
    txtABIntCor.Text = CCur(txtABIntCor.Text)
vError:
End Sub

Private Sub txtABIntCor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABIntMor.SetFocus
End Sub

Private Sub txtABIntCor_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABIntCor_LostFocus()
On Error GoTo vError
    txtABIntCor.Text = Format(CCur(txtABIntCor.Text), "Standard")
vError:
End Sub


Private Sub txtABIntMor_GotFocus()
On Error GoTo vError
    txtABIntMor.Text = CCur(txtABIntMor.Text)
vError:
End Sub

Private Sub txtABIntMor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABAmortizacion.SetFocus
End Sub

Private Sub txtABIntMor_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABIntMor_LostFocus()
On Error GoTo vError
    txtABIntMor.Text = Format(CCur(txtABIntMor.Text), "Standard")
vError:
End Sub

Private Sub txtABAmortizacion_GotFocus()
On Error GoTo vError
    txtABAmortizacion.Text = CCur(txtABAmortizacion.Text)
vError:
End Sub

Private Sub txtABAmortizacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAbPoliza.SetFocus
End Sub

Private Sub txtABAmortizacion_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABAmortizacion_LostFocus()
On Error GoTo vError
    txtABAmortizacion.Text = Format(CCur(txtABAmortizacion.Text), "Standard")
vError:
End Sub

Private Sub txtABCargos_GotFocus()
On Error GoTo vError
    txtABCargos.Text = CCur(txtABCargos.Text)
vError:
End Sub

Private Sub txtABCargos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABTotal.SetFocus
End Sub

Private Sub txtABCargos_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABCargos_LostFocus()
On Error GoTo vError
    txtABCargos.Text = Format(CCur(txtABCargos.Text), "Standard")
vError:
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyReturn Then
  Call sbConsulta
End If

End Sub
