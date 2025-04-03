VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCC_Poliza_Beneficiarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Beneficiarios de Polizas"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1215
      Left            =   120
      TabIndex        =   46
      Top             =   6240
      Width           =   10455
      _Version        =   1441793
      _ExtentX        =   18441
      _ExtentY        =   2143
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnGuardar 
         Height          =   615
         Left            =   8400
         TabIndex        =   47
         Top             =   360
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmPoliza_Beneficiarios.frx":0000
         ImageAlignment  =   4
      End
   End
   Begin XtremeSuiteControls.PushButton btnClean 
      Height          =   210
      Index           =   0
      Left            =   10320
      TabIndex        =   39
      Top             =   2760
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   370
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmPoliza_Beneficiarios.frx":0731
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   600
      Top             =   120
   End
   Begin XtremeSuiteControls.ComboBox cboPoliza 
      Height          =   465
      Left            =   2520
      TabIndex        =   3
      Top             =   1800
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
      _ExtentY        =   820
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      FlatStyle       =   -1  'True
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboParentesco 
      Height          =   330
      Index           =   0
      Left            =   7560
      TabIndex        =   10
      Top             =   2760
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Index           =   0
      Left            =   1560
      TabIndex        =   11
      Top             =   2760
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   2760
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   330
      Index           =   0
      Left            =   9120
      TabIndex        =   13
      Top             =   2760
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   3240
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboParentesco 
      Height          =   330
      Index           =   1
      Left            =   7560
      TabIndex        =   15
      Top             =   3240
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Index           =   1
      Left            =   1560
      TabIndex        =   16
      Top             =   3240
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Index           =   1
      Left            =   3120
      TabIndex        =   17
      Top             =   3240
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   330
      Index           =   1
      Left            =   9120
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboParentesco 
      Height          =   330
      Index           =   2
      Left            =   7560
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Index           =   2
      Left            =   1560
      TabIndex        =   21
      Top             =   3720
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Index           =   2
      Left            =   3120
      TabIndex        =   22
      Top             =   3720
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   330
      Index           =   2
      Left            =   9120
      TabIndex        =   23
      Top             =   3720
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   4200
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboParentesco 
      Height          =   330
      Index           =   3
      Left            =   7560
      TabIndex        =   25
      Top             =   4200
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Index           =   3
      Left            =   1560
      TabIndex        =   26
      Top             =   4200
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   27
      Top             =   4200
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   330
      Index           =   3
      Left            =   9120
      TabIndex        =   28
      Top             =   4200
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Index           =   4
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboParentesco 
      Height          =   330
      Index           =   4
      Left            =   7560
      TabIndex        =   30
      Top             =   4680
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Index           =   4
      Left            =   1560
      TabIndex        =   31
      Top             =   4680
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Index           =   4
      Left            =   3120
      TabIndex        =   32
      Top             =   4680
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   330
      Index           =   4
      Left            =   9120
      TabIndex        =   33
      Top             =   4680
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   330
      Index           =   5
      Left            =   120
      TabIndex        =   34
      Top             =   5160
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboParentesco 
      Height          =   330
      Index           =   5
      Left            =   7560
      TabIndex        =   35
      Top             =   5160
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Index           =   5
      Left            =   1560
      TabIndex        =   36
      Top             =   5160
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   330
      Index           =   5
      Left            =   3120
      TabIndex        =   37
      Top             =   5160
      Width           =   4455
      _Version        =   1441793
      _ExtentX        =   7858
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   330
      Index           =   5
      Left            =   9120
      TabIndex        =   38
      Top             =   5160
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
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
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnClean 
      Height          =   210
      Index           =   1
      Left            =   10320
      TabIndex        =   40
      Top             =   3240
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   370
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmPoliza_Beneficiarios.frx":0D63
   End
   Begin XtremeSuiteControls.PushButton btnClean 
      Height          =   210
      Index           =   2
      Left            =   10320
      TabIndex        =   41
      Top             =   3720
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   370
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmPoliza_Beneficiarios.frx":1395
   End
   Begin XtremeSuiteControls.PushButton btnClean 
      Height          =   210
      Index           =   3
      Left            =   10320
      TabIndex        =   42
      Top             =   4200
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   370
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmPoliza_Beneficiarios.frx":19C7
   End
   Begin XtremeSuiteControls.PushButton btnClean 
      Height          =   210
      Index           =   4
      Left            =   10320
      TabIndex        =   43
      Top             =   4680
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   370
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmPoliza_Beneficiarios.frx":1FF9
   End
   Begin XtremeSuiteControls.PushButton btnClean 
      Height          =   210
      Index           =   5
      Left            =   10320
      TabIndex        =   44
      Top             =   5160
      Width           =   255
      _Version        =   1441793
      _ExtentX        =   450
      _ExtentY        =   370
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmPoliza_Beneficiarios.frx":262B
   End
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   450
      Left            =   9120
      TabIndex        =   48
      Top             =   5760
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   794
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   6
      Left            =   6480
      TabIndex        =   49
      Top             =   5760
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Total:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   45
      Top             =   1800
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Póliza Colectiva:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   4
      Left            =   9240
      TabIndex        =   8
      Top             =   2400
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Porcentaje"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Parentesco"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   2
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   4335
      _Version        =   1441793
      _ExtentX        =   7646
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Nombre"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   661
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Tipo de Id"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pólizas Colectivas: Registro de Beneficiarios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   7935
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   8415
      _Version        =   1441793
      _ExtentX        =   14843
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "NOMBRE_COMPLETO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   2535
      _Version        =   1441793
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "CEDULA"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.99
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCC_Poliza_Beneficiarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoJuridica As Long, mCedula As String
Dim i As Integer
Dim vPaso As Boolean

Private Sub sbLista_Clean()

For i = 0 To 5
    txtIdentificacion.Item(i).Text = ""
    txtNombre.Item(i).Text = ""
    txtPorcentaje.Item(i).Text = "0"
Next i

End Sub

Private Sub sbBeneficiario_Load()

On Error GoTo vError

Call sbLista_Clean

If cboPoliza.ListCount = 0 Then Exit Sub

i = 0

strSQL = "exec spPoliza_Persona_Beneficiarios '" & mCedula & "', '" & cboPoliza.ItemData(cboPoliza.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 
 txtIdentificacion.Item(i).Text = rs!Cedula
 txtNombre.Item(i).Text = rs!Nombre
 txtPorcentaje.Item(i).Text = Format(rs!Porcentaje, "Standard")
 
 Call sbCboAsignaDato(cboTipoId(i), rs!Tipo_Id_Desc, True, rs!Tipo_Id)
 Call sbCboAsignaDato(cboParentesco(i), rs!Parentesco_Desc, True, rs!Cod_Parentesco)
  
 i = i + 1
 rs.MoveNext
Loop
rs.Close

'Calcula totales
Call txtPorcentaje_LostFocus(0)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnClean_Click(Index As Integer)
    txtIdentificacion.Item(Index).Text = ""
    txtNombre.Item(Index).Text = ""
    txtPorcentaje.Item(Index).Text = "0"
End Sub

Private Sub btnGuardar_Click()
Dim pLinea As Integer

On Error GoTo vError

If CCur(txtTotal.Text) <> 100 Then
    MsgBox "El Porcentaje de todos los Beneficiarios es diferente al 100%", vbExclamation
    Exit Sub
End If

strSQL = ""
pLinea = 1


'spPoliza_Persona_Beneficiarios_Add(@Cedula varchar(20), @Poliza varchar(10)
'            , @Linea smallint, @Tipo_ID int, @Identificacion varchar(20), @Nombre varchar(100)
'            , @Parentesco varchar(20), @Porcentaje dec(10,2), @Usuario varchar(30))


For i = 0 To 5
   
   txtIdentificacion.Item(i).Text = fxSysCleanTxtInject(txtIdentificacion.Item(i).Text)
   txtNombre.Item(i).Text = fxSysCleanTxtInject(txtNombre.Item(i).Text)
   
   If CCur(txtPorcentaje.Item(i).Text) > 0 Then
    strSQL = strSQL & Space(10) & "exec spPoliza_Persona_Beneficiarios_Add '" & mCedula & "', '" & cboPoliza.ItemData(cboPoliza.ListIndex) _
           & "', " & pLinea & ", " & cboTipoId.Item(i).ItemData(cboTipoId.Item(i).ListIndex) _
           & ", '" & txtIdentificacion.Item(i).Text & "', '" & txtNombre.Item(i).Text _
           & "', '" & cboParentesco.Item(i).ItemData(cboParentesco.Item(i).ListIndex) _
           & "', " & CCur(txtPorcentaje.Item(i).Text) & ", '" & glogon.Usuario & "'"
       
     pLinea = pLinea + 1
   End If
Next i

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Registra", "Beneficiarios Poliza Colectiva: " & cboPoliza.ItemData(cboPoliza.ListIndex))

    MsgBox "Beneficiarios registrados satisfactoriamente!", vbInformation
End If

Call sbBeneficiario_Load

Exit Sub

vError:
 
End Sub

Private Sub cboPoliza_Click()
If vPaso Then Exit Sub


Call sbBeneficiario_Load

End Sub

Private Sub Form_Load()

vModulo = 11

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

scMain.Item(0).Caption = GLOBALES.gTag
scMain.Item(1).Caption = GLOBALES.gTag2

mCedula = GLOBALES.gTag

vTipoJuridica = 0
strSQL = "select TIPO_ID from AFI_TIPOS_IDS where Tipo_Personeria = 'J'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    vTipoJuridica = rs!Tipo_Id
End If
rs.Close


 strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
        & " order by Tipo_Id"
 Call sbCbo_Llena_New(cboTipoId(0), strSQL, False, True)

 strSQL = "select rtrim(cod_Parentesco) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from sys_Parentescos where activo = 1"
 Call sbCbo_Llena_New(cboParentesco(0), strSQL, False, True)
  
 Call sbCbo_Copia(cboParentesco(0), cboParentesco(1))
 Call sbCbo_Copia(cboParentesco(0), cboParentesco(2))
 Call sbCbo_Copia(cboParentesco(0), cboParentesco(3))
 Call sbCbo_Copia(cboParentesco(0), cboParentesco(4))
 Call sbCbo_Copia(cboParentesco(0), cboParentesco(5))

 Call sbCbo_Copia(cboTipoId(0), cboTipoId(1))
 Call sbCbo_Copia(cboTipoId(0), cboTipoId(2))
 Call sbCbo_Copia(cboTipoId(0), cboTipoId(3))
 Call sbCbo_Copia(cboTipoId(0), cboTipoId(4))
 Call sbCbo_Copia(cboTipoId(0), cboTipoId(5))


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()
 TimerX.Interval = 0
 TimerX.Enabled = False
 
 vPaso = True
 
 strSQL = "select COD_POLIZA as 'Idx', rtrim(Poliza_Desc) as 'ItmX' from vPoliza_Catalogo" _
        & " order by COD_POLIZA"
 Call sbCbo_Llena_New(cboPoliza, strSQL, False, True)
 
 vPaso = False
 
 Call cboPoliza_Click
 
End Sub


Private Sub txtIdentificacion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Consulta Padron

If txtIdentificacion(Index).Text <> "" And txtNombre(Index).Text = "" And KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    Call gBase_Padron(txtIdentificacion(Index).Text, "General", rs, "CRC")
    
    If rs.RecordCount > 0 Then
       txtNombre(Index).Text = Trim(rs!Apellido_1) & " " & Trim(rs!Apellido_2) & " " & Trim(rs!Nombre)
    End If
End If

End Sub

Private Sub txtPorcentaje_LostFocus(Index As Integer)
On Error GoTo vError
  
  txtPorcentaje.Item(Index).Text = Format(CCur(txtPorcentaje.Item(Index).Text), "Standard")
  
Dim curTotal As Currency

curTotal = 0

For i = 0 To 5
  curTotal = curTotal + CCur(txtPorcentaje.Item(i).Text)
Next i

txtTotal.Text = Format(curTotal, "Standard")
  
Exit Sub

vError:
  txtPorcentaje.Item(Index).Text = "0"
End Sub
