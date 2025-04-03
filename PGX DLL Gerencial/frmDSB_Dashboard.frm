VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Object = "{C215CB9A-0AE1-499F-A101-48B3C370D3DF}#22.1#0"; "Codejock.ChartControl.v22.1.0.ocx"
Begin VB.Form frmDSB_Dashboard 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Dashboard: Prinicipal"
   ClientHeight    =   10875
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   21990
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   21990
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox gbMenu 
      Height          =   9375
      Left            =   -4920
      TabIndex        =   63
      Top             =   1200
      Visible         =   0   'False
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
      _ExtentY        =   16536
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.ListView lswMenu 
         Height          =   8775
         Left            =   120
         TabIndex        =   64
         Top             =   480
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   15478
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         View            =   3
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   375
         Left            =   120
         TabIndex        =   65
         Top             =   120
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione una Categoría"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   16777215
      End
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   75
      Top             =   2040
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ S }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   375
      Index           =   2
      Left            =   13800
      TabIndex        =   74
      Top             =   6600
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ S }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   73
      Top             =   6600
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ S }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSave 
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   72
      Top             =   6600
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ S }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnPrint 
      Height          =   375
      Index           =   2
      Left            =   13800
      TabIndex        =   70
      Top             =   6240
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ P }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnPrint 
      Height          =   375
      Index           =   1
      Left            =   9000
      TabIndex        =   69
      Top             =   6240
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ P }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnPrint 
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   68
      Top             =   6240
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ P }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   17
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   17
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ccpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   0
         Left            =   600
         TabIndex        =   19
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.ComboBox cboChart 
      Height          =   330
      Index           =   0
      Left            =   3120
      TabIndex        =   12
      Top             =   5880
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
   Begin XtremeSuiteControls.ListView lswTop 
      Height          =   3375
      Left            =   9600
      TabIndex        =   5
      Top             =   2415
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   5953
      _StockProps     =   77
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
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   15720
      Picture         =   "frmDSB_Dashboard.frx":0000
      ScaleHeight     =   885
      ScaleWidth      =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   3360
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9240
      Top             =   120
   End
   Begin XtremeChartControl.ChartControl chartH 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   2415
      Width           =   9495
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   5948
      _StockProps     =   0
   End
   Begin XtremeChartControl.ChartControl chartC 
      Height          =   3855
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   5880
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   6794
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.ComboBox cboHFiltro 
      Height          =   315
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.ComboBox cboChart_H 
      Height          =   315
      Left            =   7440
      TabIndex        =   11
      Top             =   2070
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.ComboBox cboChart 
      Height          =   330
      Index           =   1
      Left            =   7920
      TabIndex        =   13
      Top             =   5880
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
   Begin XtremeChartControl.ChartControl chartC 
      Height          =   3855
      Index           =   1
      Left            =   4800
      TabIndex        =   14
      Top             =   5880
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   6794
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.ComboBox cboChart 
      Height          =   330
      Index           =   2
      Left            =   12720
      TabIndex        =   15
      Top             =   5880
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
   Begin XtremeChartControl.ChartControl chartC 
      Height          =   3855
      Index           =   2
      Left            =   9600
      TabIndex        =   16
      Top             =   5880
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8276
      _ExtentY        =   6794
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   1
      Left            =   3120
      TabIndex        =   20
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   21
         Tag             =   "Clpr"
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Clpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   1
         Left            =   600
         TabIndex        =   22
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   2
      Left            =   6240
      TabIndex        =   23
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   24
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ccpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   2
         Left            =   600
         TabIndex        =   25
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   3
      Left            =   9360
      TabIndex        =   26
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ccpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   3
         Left            =   600
         TabIndex        =   28
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   4
      Left            =   12480
      TabIndex        =   29
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   30
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ccpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   4
         Left            =   600
         TabIndex        =   31
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   5
      Left            =   15600
      TabIndex        =   32
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ccpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   5
         Left            =   600
         TabIndex        =   34
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.GroupBox gbKPI 
      Height          =   7695
      Left            =   14400
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   13573
      _StockProps     =   79
      Caption         =   "gbKPI"
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   10
         Top             =   480
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   600
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   37
         Top             =   1320
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   1
         Left            =   0
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   2
         Left            =   1080
         TabIndex        =   39
         Top             =   2040
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   2
         Left            =   0
         TabIndex        =   40
         Top             =   1920
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   3
         Left            =   1080
         TabIndex        =   41
         Top             =   2760
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   2640
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   4
         Left            =   1080
         TabIndex        =   43
         Top             =   3480
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   4
         Left            =   0
         TabIndex        =   44
         Top             =   3360
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   5
         Left            =   1080
         TabIndex        =   45
         Top             =   4200
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   5
         Left            =   0
         TabIndex        =   46
         Top             =   4080
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   6
         Left            =   1080
         TabIndex        =   47
         Top             =   4920
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   48
         Top             =   4800
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   7
         Left            =   1080
         TabIndex        =   49
         Top             =   5640
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   7
         Left            =   0
         TabIndex        =   50
         Top             =   5520
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   8
         Left            =   1080
         TabIndex        =   51
         Top             =   6360
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   8
         Left            =   0
         TabIndex        =   52
         Top             =   6240
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIp 
         Height          =   375
         Index           =   9
         Left            =   1080
         TabIndex        =   53
         Top             =   7080
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Text            =   "15.34%"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnKPIp 
         Height          =   615
         Index           =   9
         Left            =   0
         TabIndex        =   54
         Top             =   6960
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1931
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Tpp"
         ForeColor       =   16761024
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   0
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Indices"
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
      End
   End
   Begin XtremeSuiteControls.GroupBox gbKpiR 
      Height          =   615
      Index           =   6
      Left            =   18720
      TabIndex        =   55
      Top             =   1400
      Width           =   3135
      _Version        =   1441793
      _ExtentX        =   5530
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton btnKPIr 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   56
         Top             =   360
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ccpr"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         TextAlignment   =   1
         Appearance      =   14
         ImageAlignment  =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtKPIr 
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   57
         Top             =   0
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   4473924
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "55,706,795.38"
         BackColor       =   16777215
         Alignment       =   1
         Appearance      =   12
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
   End
   Begin XtremeSuiteControls.ComboBox cboTop 
      Height          =   330
      Left            =   12600
      TabIndex        =   58
      Top             =   2040
      Width           =   1695
      _Version        =   1441793
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
   Begin XtremeSuiteControls.ComboBox cboTopList 
      Height          =   330
      Left            =   10440
      TabIndex        =   59
      Top             =   2040
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.ComboBox cboTopCount 
      Height          =   330
      Left            =   9600
      TabIndex        =   60
      Top             =   2040
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
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
   Begin XtremeSuiteControls.PushButton btnMenu 
      Height          =   975
      Left            =   120
      TabIndex        =   61
      Top             =   120
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   1720
      _StockProps     =   79
      ForeColor       =   14737632
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      TextAlignment   =   1
      Appearance      =   2
      Picture         =   "frmDSB_Dashboard.frx":436C
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cboPalette 
      Height          =   330
      Left            =   11280
      TabIndex        =   66
      Top             =   840
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.ComboBox cboAppearance 
      Height          =   330
      Left            =   13080
      TabIndex        =   67
      Top             =   840
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.PushButton btnPrint 
      Height          =   375
      Index           =   3
      Left            =   4920
      TabIndex        =   71
      Top             =   2040
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "{ P }"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnAccess 
      Height          =   450
      Left            =   8400
      TabIndex        =   77
      Top             =   120
      Visible         =   0   'False
      Width           =   855
      _Version        =   1441793
      _ExtentX        =   1508
      _ExtentY        =   794
      _StockProps     =   79
      Caption         =   "Access"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.Label lblMenu 
      Height          =   495
      Left            =   1440
      TabIndex        =   62
      Top             =   120
      Width           =   3615
      _Version        =   1441793
      _ExtentX        =   6376
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Menú"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scHistograma 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   9495
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Filtro"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scTop 
      Height          =   375
      Left            =   9600
      TabIndex        =   35
      Top             =   2040
      Width           =   4695
      _Version        =   1441793
      _ExtentX        =   8281
      _ExtentY        =   661
      _StockProps     =   14
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
   End
   Begin XtremeSuiteControls.Label lblCorte 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Corte:"
      ForeColor       =   12582912
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblDashboard 
      Height          =   615
      Left            =   11280
      TabIndex        =   0
      Top             =   240
      Width           =   3615
      _Version        =   1441793
      _ExtentX        =   6376
      _ExtentY        =   1085
      _StockProps     =   79
      Caption         =   "DASHBOARD"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblBack 
      Height          =   2775
      Left            =   16920
      TabIndex        =   76
      Top             =   2040
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   4895
      _StockProps     =   79
      BackColor       =   16777215
   End
End
Attribute VB_Name = "frmDSB_Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mChartPallete As String, mChartLabelPosition As Integer
Dim Diagram As ChartDiagram2D
Dim Strip As ChartAxisStrip

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim mCorte As Date, vPaso As Boolean

Dim mCategoria As String


Private Declare Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hWnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long

Private Type vGraficos
    Titulo      As String
    Tema        As String
    sql         As String
    SQL_Filtro  As String
    SQL_Sp      As String
    Deciminal   As Integer
    Pattern     As String
    Cifra       As Long
    Codigo      As String
End Type

Dim mDashboard As Integer
Dim pGra_H As vGraficos, pGra_C1 As vGraficos, pGra_C2 As vGraficos, pGra_C3 As vGraficos

Public Sub sbChart_MultiSeries(pChart As Object, pTitulo As String, Optional pTema As String = "" _
                        , Optional i As Long = 1, Optional Pattern As String = "N" _
                        , Optional pDecimals As Integer = 0 _
                        , Optional pTipo As String = "Spline")
    Dim C As Currency
    
    Me.MousePointer = vbHourglass
    
    cboHFiltro.Visible = True
    cboChart_H.Visible = True
    
    
    If pChart.Content.Series.Count > 0 Then
        pChart.Content.Series.DeleteAll
    End If
    
    pChart.Content.Legend.Visible = True
    pChart.Content.Legend.HorizontalAlignment = xtpChartLegendNearOutside
    
    Dim Series As ChartSeries
            
'    If pChart.Content.Series.Count > 0 Then
'        Set Series = pChart.Content.Series(0)
'    Else
'        Set Series = pChart.Content.Series.Add("Series 1")
'    End If
'
    pChart.Content.Legend.Visible = True


    'Carga los Filtros
    Dim pSQL As String
    pSQL = pGra_H.SQL_Filtro
    
    If cboHFiltro.ListCount = 0 Then
        vPaso = True
        Call sbCbo_Llena_New(cboHFiltro, pSQL, True, True)
        vPaso = False
    End If
    
    
    'Procesa el Grafico
    Dim pSerieActual As String, pCount As Integer, x As Integer
    
    pSerieActual = ""
    pCount = 0
    Call OpenRecordSet(rs, strSQL)
    
    Do While Not rs.EOF
        If pSerieActual <> rs!Serie Then
            Set Series = pChart.Content.Series.Add(rs!Serie)
        
            pCount = pCount + 1
            pSerieActual = rs!Serie
        End If
        
        Series.Points.Add Format(rs!Corte, "yyyy-MM"), rs!Value / i
     
      rs.MoveNext
    Loop
    rs.Close
            
pChart.Content.Titles(0).Text = pTitulo
pChart.Content.Legend.Visible = True
For x = 0 To pCount - 1
    Select Case pTipo
        Case "Spline"
                Set pChart.Content.Series(x).Style = New ChartSplineSeriesStyle
        Case "Bar"
                Set pChart.Content.Series(x).Style = New ChartBarSeriesStyle
        Case "Point"
                Set pChart.Content.Series(x).Style = New ChartPointSeriesStyle
        Case "Area"
                Set pChart.Content.Series(x).Style = New ChartAreaSeriesStyle
        Case "Multi"
                Set pChart.Content.Series(x).Style = New ChartAreaSeriesStyle
'                Set pChart.Content.Series(2).Style = New ChartPointSeriesStyle
'                Set pChart.Content.Series(1).Style = New ChartBarSeriesStyle
'                Set pChart.Content.Series(0).Style = New ChartSplineSeriesStyle
    End Select
Next x

   
Dim pSeries As ChartSeries
For Each pSeries In pChart.Content.Series
    pSeries.Style.Label.Format.Category = xtpChartNumber
    pSeries.Style.Label.Format.DecimalPlaces = 2
Next

Set Diagram = pChart.Content.Series(0).Diagram

'Diagram.AxisX.Range.AutoRange = True
'Diagram.AxisY.Range.AutoRange = True

Select Case i
    Case 1
    Case 1000
        Diagram.AxisY.Title = "Monto en (miles)"
    Case 1000000
        Diagram.AxisY.Title = "Monto en (millones)"
End Select
Diagram.AxisY.Title.Visible = True
Diagram.AxisX.Title = "Meses"
Diagram.AxisX.Title.Visible = True

Diagram.AxisY.Label.Format.Category = xtpChartNumber
Diagram.AxisY.Label.Format.DecimalPlaces = 0


Me.MousePointer = vbDefault

End Sub

Public Sub sbChart_Histograma(pChart As Object, pTitulo As String, Optional pTema As String = "" _
                        , Optional i As Long = 1, Optional Pattern As String = "N" _
                        , Optional pDecimals As Integer = 0)
    Dim C As Currency
    
 On Error GoTo vError
         
Me.MousePointer = vbHourglass
         
    cboHFiltro.Visible = False
    cboChart_H.Visible = False

 
    If pChart.Content.Series.Count > 0 Then
        pChart.Content.Series.DeleteAll
    End If
    
    pChart.Content.Titles.DeleteAll
    pChart.Content.Titles.Add pTitulo
    
    pChart.Content.Legend.Visible = True
    pChart.Content.Legend.HorizontalAlignment = xtpChartLegendFar
    
    Dim Series As ChartSeries
    
    If pChart.Content.Series.Count > 0 Then
        Set Series = pChart.Content.Series(0)
    Else
        Set Series = pChart.Content.Series.Add("Series 1")
    End If
    
    pChart.Content.Legend.Visible = True

    C = 0
    
    Call OpenRecordSet(rs, strSQL)

    Do While Not rs.EOF
      Series.Points.Add Format(rs!Descripcion, "yyyy-MM"), rs!Value / i
              
      If (rs!Value / i) > C Then
        C = rs!Value / i
      End If
      
      rs.MoveNext
    Loop
    rs.Close
    
    Series.LegendVisible = False
    
    Dim LineStyle As ChartLineSeriesStyle
    Set LineStyle = New ChartLineSeriesStyle
    
    LineStyle.LineStyle.Thickness = 4
    LineStyle.LineStyle.DashStyle = xtpChartDashStyleDot
    
    Dim BarStyle As ChartBarSeriesStyle
    Set BarStyle = New ChartBarSeriesStyle
    
    Dim AreaStyle As ChartAreaSeriesStyle
    Set AreaStyle = New ChartAreaSeriesStyle
    
    
    Select Case Pattern
      Case "N" 'Numerico
            LineStyle.Label.Format.DecimalPlaces = pDecimals
            LineStyle.Label.Format.UseThousandSeparator = True
            LineStyle.Label.Format.Category = xtpChartNumber
       
       Case "P"
            LineStyle.Label.Format.Category = xtpChartPercentage
    
    End Select
    

    Set pChart.Content.Series(0).Style = LineStyle
'    Set pChart.Content.Series(0).Style = BarStyle
'    Set pChart.Content.Series(0).Style = AreaStyle
   
Me.MousePointer = vbDefault
   
   
Exit Sub

vError:
    
Me.MousePointer = vbDefault
    
End Sub


Public Sub sbChart_3D(pChart As Object, pTitulo As String, Optional pTema As String = "" _
                        , Optional p3dTipo = "3d_Pie", Optional pExpresado As Long = 1 _
                        , Optional Pattern As String = "N" _
                        , Optional pDecimals As Integer = 0)

On Error GoTo vError
    
    If pChart.Content.Series.Count > 0 Then
        pChart.Content.Series.DeleteAll
    End If
    
    Dim Series As ChartSeries
        
        
    pChart.Content.Titles.DeleteAll
    pChart.Content.Titles.Add pTitulo
    
    pChart.Content.Legend.Visible = True
    pChart.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
    
    Set Series = pChart.Content.Series.Add(pTema)
    
    
    Call OpenRecordSet(rs, strSQL)
               
    Dim i As Integer, x As Integer, C As Currency

    C = 0
    i = 0
    x = 0

    Do While Not rs.EOF
      C = rs!Value / pExpresado

      CreateSeriesPoint Series.Points, rs!Descripcion, rs!Value / pExpresado

      If rs!Value > C Then
        C = rs!Value / pExpresado
        x = i
      End If

      i = i + 1
      rs.MoveNext
    Loop
    rs.Close

    Series.Points(x).Special = True
                
                
                
    pChart.Content.Appearance.SetPalette mChartPallete
   ' pChart.Content.Series(0).Style.Label.Position mChartLabelPosition
   Select Case p3dTipo
    Case "Pie"
                        
            Dim PieStyle As ChartPieSeriesStyle
            Set PieStyle = New ChartPieSeriesStyle
            Set pChart.Content.Series(0).Style = PieStyle
            
'            PieStyle.Label.Format.Pattern = "{V} Million"
            PieStyle.Label.Format.Category = xtpChartNumber
                       
            'PieStyle.HolePercent = 40
            PieStyle.Rotation = 30
            PieStyle.Label.Visible = True
            'PieStyle.Label.ShowLines = False
            'PieStyle.ExplodedDistancePercent = 30
            
            'cmbPieLabelPosition.ListIndex = PieStyle.Label.Position
            pChart.Content.Series(0).Style.Label.Position = 1
            PieStyle.Label.Antialiasing = False
            
    Case "Pyramid"
            
            Dim PyramidStyle As ChartPyramidSeriesStyle
            Set PyramidStyle = New ChartPyramidSeriesStyle
            Set pChart.Content.Series(0).Style = PyramidStyle
            
            'PyramidStyle.Label.Format.Pattern = "{V} Million"
            PyramidStyle.Label.Format.Category = xtpChartNumber
                       
            PyramidStyle.HeightToWidthRatio = 1.25
            PyramidStyle.PointDistance = 5
            PyramidStyle.Label.Visible = True
            
            'cmbLabelPosition.ListIndex = PyramidStyle.Label.Position
            PyramidStyle.Label.Antialiasing = False
    
    Case "3d_Pie"
                  
            Dim Style3dPie As Chart3dPieSeriesStyle
            
            Set Style3dPie = New Chart3dPieSeriesStyle
            Set pChart.Content.Series(0).Style = Style3dPie

            Style3dPie.Label.Format.Category = xtpChartNumber
                       
            Dim Pie3dRotation As Chart3dRotation
            Set Pie3dRotation = New Chart3dRotation
            
            Pie3dRotation.Yaw = 60
            Pie3dRotation.Pitch = 40
            Pie3dRotation.Roll = 30
            Style3dPie.SetRotation Pie3dRotation
            Style3dPie.Label.Visible = True
            
            Style3dPie.Label.Antialiasing = False
     
     Case "3d_Doughnut"
            Dim Style3dDoughnut As Chart3dPieSeriesStyle
            Set Style3dDoughnut = New Chart3dPieSeriesStyle
            
            Set pChart.Content.Series(0).Style = Style3dDoughnut
            Style3dDoughnut.Label.Format.Category = xtpChartNumber
                       
            Dim Doughnut3dRotation As Chart3dRotation
            Set Doughnut3dRotation = New Chart3dRotation
            Doughnut3dRotation.Yaw = 10
            Doughnut3dRotation.Pitch = 20
            Doughnut3dRotation.Roll = 50
            Style3dDoughnut.SetRotation Doughnut3dRotation
            Style3dDoughnut.Label.Visible = True
            Style3dDoughnut.HolePercent = 60
            Style3dDoughnut.ExplodedDistancePercent = 20
     
     Case "3d_Pyramid"
            Dim Pyramid3dStyle As Chart3dPyramidSeriesStyle
            Set Pyramid3dStyle = New Chart3dPyramidSeriesStyle
            Set pChart.Content.Series(0).Style = Pyramid3dStyle
            
            Pyramid3dStyle.Label.Format.Category = xtpChartNumber
                       
            Dim Pyramid3dRotation As Chart3dRotation
            Set Pyramid3dRotation = New Chart3dRotation
            Pyramid3dRotation.Yaw = 70
            Pyramid3dRotation.Pitch = 20
            Pyramid3dRotation.Roll = 70
            Pyramid3dStyle.SetRotation Pyramid3dRotation
            
            Pyramid3dStyle.HeightToWidthRatio = 2
            Pyramid3dStyle.PointDistance = 2
            Pyramid3dStyle.BaseEdgeCount = 7
            Pyramid3dStyle.SmoothEdges = True
            Pyramid3dStyle.Label.Visible = True
            
     Case "3d_Torus"
            Dim Style3dTorus As Chart3dPieSeriesStyle
            Set Style3dTorus = New Chart3dPieSeriesStyle
            
            Set pChart.Content.Series(0).Style = Style3dTorus
            
            Style3dTorus.Label.Format.Category = xtpChartNumber
                       
            Dim Torus3dRotation As Chart3dRotation
            Set Torus3dRotation = New Chart3dRotation
            Torus3dRotation.Yaw = -20
            Torus3dRotation.Pitch = 0
            Torus3dRotation.Roll = 70
            Style3dTorus.SetRotation Torus3dRotation
            Style3dTorus.Label.Visible = True
            Style3dTorus.IsTorus = True
            Style3dTorus.Depth = Style3dTorus.Depth * 2
            
            Style3dTorus.Label.Antialiasing = False
            
      Case "3d_Funnel"
            'AddFunnelSeries
                        
            Dim Funnel3dStyle As Chart3dFunnelSeriesStyle
            Set Funnel3dStyle = New Chart3dFunnelSeriesStyle
            
            Set pChart.Content.Series(0).Style = Funnel3dStyle
            
            Funnel3dStyle.Label.Format.Category = xtpChartNumber
                       
            Dim Funnel3dRotation As Chart3dRotation
            Set Funnel3dRotation = New Chart3dRotation
            Funnel3dRotation.Yaw = 203
            Funnel3dRotation.Pitch = 355
            Funnel3dRotation.Roll = 79
            Funnel3dStyle.SetRotation Funnel3dRotation
            
            Funnel3dStyle.HeightToWidthRatio = 1.5
            Funnel3dStyle.BaseEdgeCount = 4
            Funnel3dStyle.SmoothEdges = True
            Funnel3dStyle.Label.Visible = True
            
            Funnel3dStyle.Label.Antialiasing = False
    
    
    End Select
                
                
                
    Select Case Pattern
      Case "N" 'Numerico
'            LineStyle.Label.Format.DecimalPlaces = pDecimals
'            LineStyle.Label.Format.UseThousandSeparator = True
'            LineStyle.Label.Format.Category = xtpChartNumber
       
            Set Diagram = pChart.Content.Series(0).Diagram
            Select Case pExpresado
                Case 1
                Case 1000
                    Diagram.AxisY.Title = "Monto en (miles)"
                Case 1000000
                    Diagram.AxisY.Title = "Monto en (millones)"
            End Select
            Diagram.AxisY.Title.Visible = True
            Diagram.AxisX.Title = "Meses"
            Diagram.AxisX.Title.Visible = True
            
            Diagram.AxisY.Label.Format.Category = xtpChartNumber
            Diagram.AxisY.Label.Format.DecimalPlaces = 0
       
       
       Case "P"
'            LineStyle.Label.Format.Category = xtpChartPercentage
       
'            Set Diagram = pChart.Content.Series(0).Diagram
'            Select Case pExpresado
'                Case 1
'                Case 1000
'                    Diagram.AxisY.Title = "Monto en (miles)"
'                Case 1000000
'                    Diagram.AxisY.Title = "Monto en (millones)"
'            End Select
'            Diagram.AxisY.Title.Visible = True
'            Diagram.AxisX.Title = "Meses"
'            Diagram.AxisX.Title.Visible = True
'
            Diagram.AxisY.Label.Format.Category = xtpChartPercentage
            Diagram.AxisY.Label.Format.DecimalPlaces = 0
       
       
    End Select
                
    
Exit Sub
    
vError:
    
End Sub

Sub CreateSeriesPoint(ByVal pPointCollection As ChartSeriesPointCollection, vArg As String, nValue As Double)
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add(vArg, nValue)
     pPoint.LabelText = Format(nValue, "Standard")
     pPoint.LegendText = vArg

End Sub



Private Sub sbCreditos_Inicial()
Dim i As Integer

On Error GoTo vError

cboHFiltro.Clear

For i = 0 To btnKPIr.Count - 1
    btnKPIr.Item(i).Visible = False
    txtKPIr.Item(i).Visible = False
Next i

For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).Visible = False
    txtKPIp.Item(i).Visible = False
Next i


'Indicadores
strSQL = "exec spDSB_Creditos_Consulta Null , 'R', 'T'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
    
    mCorte = rs!Corte
    lblCorte.Caption = Format(rs!Corte, "yyyy-mm-dd")
    
    'Indices
    btnKPIp(0).Visible = True
    txtKPIp(0).Visible = True
    btnKPIp(0).Caption = "Tpp"
    btnKPIp(0).Tag = "Tpp"
    btnKPIp(0).ToolTipText = "Tasa Promedio Ponderada"
    txtKPIp(0).Text = Format(rs!Tpp, "Standard") & "%"
    
    btnKPIp(1).Visible = True
    txtKPIp(1).Visible = True
    btnKPIp(1).Caption = "Pipp"
    btnKPIp(1).Tag = "Pipp"
    txtKPIp(1).ToolTipText = "Plazo Inicial Promedio Ponderado"
    txtKPIp(1).Text = Format(rs!Pipp, "###,###,###")
    
    btnKPIp(2).Visible = True
    txtKPIp(2).Visible = True
    btnKPIp(2).Caption = "Mora"
    btnKPIp(2).Tag = "IMora_Activa"
    btnKPIp(2).ToolTipText = "Indice Mora Activa"
    txtKPIp(2).Text = Format(rs!IMora_Activa, "Standard") & "%"
    
    btnKPIp(3).Visible = True
    txtKPIp(3).Visible = True
    btnKPIp(3).Caption = "C.Jud"
    btnKPIp(3).Tag = "ICbrJud"
    btnKPIp(3).ToolTipText = "Ind. Cobro Judicial"
    txtKPIp(3).Text = Format(rs!ICbrJud, "Standard") & "%"
    
    btnKPIp(4).Visible = True
    txtKPIp(4).Visible = True
    btnKPIp(4).Caption = "Refin"
    btnKPIp(4).Tag = "IRefinancia"
    btnKPIp(4).ToolTipText = "Ind. Saldo Refinanciado"
    txtKPIp(4).Text = Format(rs!IRefinancia, "Standard") & "%"
    
    btnKPIp(5).Visible = True
    txtKPIp(5).Visible = True
    btnKPIp(5).Caption = "Canc"
    btnKPIp(5).Tag = "ICancela"
    btnKPIp(5).ToolTipText = "Ind. Cancelación Externa"
    txtKPIp(5).Text = Format(rs!ICancela, "Standard") & "%"
    
    
 
    
    'Resultados

    btnKPIr(0).Visible = True
    txtKPIr(0).Visible = True
    btnKPIr(0).Tag = "CCPr"
    btnKPIr(0).Caption = "Cartera Corto Plazo"
    txtKPIr(0).Text = Format(rs!CCPr, "###,###,###")
    
    
    btnKPIr(1).Visible = True
    txtKPIr(1).Visible = True
    btnKPIr(1).Tag = "CLPr"
    btnKPIr(1).Caption = "Cartera Largo Plazo"
    txtKPIr(1).Text = Format(rs!CLPr, "###,###,###")
    
    btnKPIr(2).Visible = True
    txtKPIr(2).Visible = True
    btnKPIr(2).Tag = "TSaldo"
    btnKPIr(2).Caption = "Saldo al Corte"
    txtKPIr(2).Text = Format(rs!TSaldo, "###,###,###")
    
    btnKPIr(3).Visible = True
    txtKPIr(3).Visible = True
    btnKPIr(3).Tag = "Colocacion"
    btnKPIr(3).Caption = "Colocación de Cartera"
    txtKPIr(3).Text = Format(rs!Colocacion, "###,###,###")
    
    btnKPIr(4).Visible = True
    txtKPIr(4).Visible = True
    btnKPIr(4).Tag = "Operaciones"
    btnKPIr(4).Caption = "Operaciones Activas"
    txtKPIr(4).Text = Format(rs!NOperaciones, "###,###,###")
    
    btnKPIr(5).Visible = True
    txtKPIr(5).Visible = True
    btnKPIr(5).Tag = "Refinancia"
    btnKPIr(5).Caption = "Saldo Refinanciado"
    txtKPIr(5).Text = Format(rs!Refinancia, "###,###,###")
    
    btnKPIr(6).Visible = True
    txtKPIr(6).Visible = True
    btnKPIr(6).Tag = "TCbrJud"
    btnKPIr(6).Caption = "Saldo en Cobro Judicial"
    txtKPIr(6).Text = Format(rs!TCbrJud, "###,###,###")
    
    
End If
rs.Close

'Paso 2: Carga Graficos Principales


vPaso = True
    cboChart(0).Text = "3d_Pie"
vPaso = False

strSQL = "exec spDSB_Creditos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(0), "Saldo por Garantías", "Garantías", cboChart(0).Text, 1)

pGra_C1.sql = strSQL
pGra_C1.Titulo = "Saldo por Garantías"
pGra_C1.Tema = "Garantías"
pGra_C1.Deciminal = 1
pGra_C1.Pattern = "N"
pGra_C1.Cifra = 1
pGra_C1.SQL_Filtro = "exec spDSB_Creditos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C1.SQL_Sp = "exec spDSB_Creditos_Consulta"
    
    
vPaso = True
    cboChart(1).Text = "3d_Torus"
vPaso = False

'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"

strSQL = "exec spDSB_Creditos_Consulta_Morosidad '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(1), "Antiguedad de Mora", "Morosidad", cboChart(1).Text, 1)

pGra_C2.sql = strSQL
pGra_C2.Titulo = "Antiguedad de Mora"
pGra_C2.Tema = "Morosidad"
pGra_C2.Deciminal = 1
pGra_C2.Pattern = "N"
pGra_C2.Cifra = 1
pGra_C2.SQL_Filtro = "exec spDSB_Creditos_Consulta_Morosidad '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C2.SQL_Sp = "exec spDSB_Creditos_Consulta_Morosidad"


vPaso = True
    cboChart(2).Text = "3d_Doughnut"
vPaso = False

strSQL = "exec spDSB_Creditos_Consulta_Tpp_Garantia '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(2), "Tpp por Garantía", "TppGar", cboChart(2).Text, 1)

pGra_C3.sql = strSQL
pGra_C3.Titulo = "Tpp por Garantía"
pGra_C3.Tema = "Morosidad"
pGra_C3.Deciminal = 1
pGra_C3.Pattern = "N"
pGra_C3.Cifra = 1
pGra_C3.SQL_Filtro = "exec spDSB_Creditos_Consulta_Tpp_Garantia '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C3.SQL_Sp = "exec spDSB_Creditos_Consulta_Tpp_Garantia"


'Top 10
Call cboTop_Click

Exit Sub

vError:
'  MsgBox Err.Description, vbCritical
'  Resume
End Sub



Private Sub sbClientes_Inicial()
Dim i As Integer

On Error GoTo vError

cboHFiltro.Clear

For i = 0 To btnKPIr.Count - 1
    btnKPIr.Item(i).Visible = False
    txtKPIr.Item(i).Visible = False
Next i

For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).Visible = False
    txtKPIp.Item(i).Visible = False
Next i


'Indicadores
strSQL = "exec spDSB_Clientes_Consulta Null , 'R', 'T'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    mCorte = rs!Corte
    lblCorte.Caption = Format(rs!Corte, "yyyy-mm-dd")

'    'Indices
    btnKPIp(0).Visible = True
    txtKPIp(0).Visible = True
    btnKPIp(0).Caption = "Ina"
    btnKPIp(0).Tag = "INuevos"
    btnKPIp(0).ToolTipText = "Indice de Nuevos Asociados"
    txtKPIp(0).Text = Format(rs!INuevos, "Standard") & "%"


    btnKPIp(1).Visible = True
    txtKPIp(1).Visible = True
    btnKPIp(1).Caption = "Ire"
    btnKPIp(1).Tag = "IReingresos"
    btnKPIp(1).ToolTipText = "Indice de Reingresos"
    txtKPIp(1).Text = Format(rs!IReingresos, "Standard") & "%"

    btnKPIp(2).Visible = True
    txtKPIp(2).Visible = True
    btnKPIp(2).Caption = "IExA"
    btnKPIp(2).Tag = "IExAsociados"
    btnKPIp(2).ToolTipText = "Indice de Ex Asociados"
    txtKPIp(2).Text = Format(rs!IExAsociados, "Standard") & "%"

'    btnKPIp(3).Visible = True
'    txtKPIp(3).Visible = True
'    btnKPIp(3).Caption = "C.Jud"
'    btnKPIp(3).Tag = "ICbrJud"
'    btnKPIp(3).ToolTipText = "Ind. Cobro Judicial"
'    txtKPIp(3).Text = Format(rs!ICbrJud, "Standard") & "%"
'
'    btnKPIp(4).Visible = True
'    txtKPIp(4).Visible = True
'    btnKPIp(4).Caption = "Refin"
'    btnKPIp(4).Tag = "IRefinancia"
'    btnKPIp(4).ToolTipText = "Ind. Saldo Refinanciado"
'    txtKPIp(4).Text = Format(rs!IRefinancia, "Standard") & "%"
'
'    btnKPIp(5).Visible = True
'    txtKPIp(5).Visible = True
'    btnKPIp(5).Caption = "Canc"
'    btnKPIp(5).Tag = "ICancela"
'    btnKPIp(5).ToolTipText = "Ind. Cancelación Externa"
'    txtKPIp(5).Text = Format(rs!ICancela, "Standard") & "%"




    'Resultados

    btnKPIr(0).Visible = True
    txtKPIr(0).Visible = True
    btnKPIr(0).Caption = "Total de Asociados"
    btnKPIr(0).Tag = "Total"
    txtKPIr(0).Text = Format(rs!Total, "###,##0")

    btnKPIr(1).Visible = True
    txtKPIr(1).Visible = True
    btnKPIr(1).Tag = "Nuevos"
    btnKPIr(1).Caption = "Nuevos Asociados"
    txtKPIr(1).Text = Format(rs!Nuevos, "###,##0")

    btnKPIr(2).Visible = True
    txtKPIr(2).Visible = True
    btnKPIr(2).Tag = "ReIng"
    btnKPIr(2).Caption = "Reingresos"
    txtKPIr(2).Text = Format(rs!Reingresos, "###,##0")

    btnKPIr(3).Visible = True
    txtKPIr(3).Visible = True
    btnKPIr(3).Tag = "Salidas"
    btnKPIr(3).Caption = "Salidas de Asociados"
    txtKPIr(3).Text = Format(rs!salidas, "###,##0")

    btnKPIr(4).Visible = True
    txtKPIr(4).Visible = True
    btnKPIr(4).Tag = "ExAsociados"
    btnKPIr(4).Caption = "Ex Asociados"
    txtKPIr(4).Text = Format(rs!ExAsociados, "###,##0")

'    btnKPIr(5).Visible = True
'    txtKPIr(5).Visible = True
'    btnKPIr(5).Caption = "Ref"
'    btnKPIr(5).Tag = "Refinancia"
'    btnKPIr(5).ToolTipText = "Saldo Refinanciado"
'    txtKPIr(5).Text = Format(rs!Refinancia, "Standard")
'
'    btnKPIr(6).Visible = True
'    txtKPIr(6).Visible = True
'    btnKPIr(6).Caption = "Jud"
'    btnKPIr(6).Tag = "TCbrJud"
'    btnKPIr(6).ToolTipText = "Saldo en Cobro Judicial"
'    txtKPIr(6).Text = Format(rs!TCbrJud, "Standard")
'

End If
rs.Close

'Paso 2: Carga Graficos Principales

vPaso = True
    cboChart(0).Text = "Pie"
vPaso = False

strSQL = "exec spDSB_Clientes_Asociados_Edades '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(0), "Edades Asociados", "Edades", cboChart(0).Text, 1)

pGra_C1.sql = strSQL
pGra_C1.Titulo = "Edades Asociados"
pGra_C1.Tema = "Edades"
pGra_C1.Deciminal = 1
pGra_C1.Pattern = "N"
pGra_C1.Cifra = 1
pGra_C1.SQL_Filtro = "exec spDSB_Clientes_Asociados_Edades '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', ''"
pGra_C1.SQL_Sp = "exec spDSB_Clientes_Asociados_Edades"
    
'Grafico No.2
vPaso = True
    cboChart(1).Text = "3d_Pie"
vPaso = False

'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"
    
strSQL = "exec spDSB_Clientes_Asociados_Generacion '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(1), "Asociados por Generación", "Generación", cboChart(1).Text, 1, "N", 0)

pGra_C2.sql = strSQL
pGra_C2.Titulo = "Asociados por Generación"
pGra_C2.Tema = "Generación"
pGra_C2.Deciminal = 1
pGra_C2.Pattern = "N"
pGra_C2.Cifra = 1
pGra_C2.SQL_Filtro = "exec spDSB_Clientes_Asociados_Generacion '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C2.SQL_Sp = "exec spDSB_Clientes_Asociados_Generacion"
    

'Grafico No.3
vPaso = True
    cboChart(2).Text = "3d_Torus"
vPaso = False

strSQL = "exec spDSB_Clientes_Consulta_Causas '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(2), "Causas Renuncias", "Causas", cboChart(2).Text, 1)

pGra_C3.sql = strSQL
pGra_C3.Titulo = "Causas Renuncias"
pGra_C3.Tema = "Causas"
pGra_C3.Deciminal = 1
pGra_C3.Pattern = "N"
pGra_C3.Cifra = 1
pGra_C3.SQL_Filtro = "exec spDSB_Clientes_Consulta_Causas '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C3.SQL_Sp = "exec spDSB_Clientes_Consulta_Causas"

'Top 10
Call cboTop_Click

Exit Sub

vError:
'  MsgBox Err.Description, vbCritical
'  Resume
End Sub



Private Sub sbBancos_Inicial()
Dim i As Integer

On Error GoTo vError

cboHFiltro.Clear

For i = 0 To btnKPIr.Count - 1
    btnKPIr.Item(i).Visible = False
    txtKPIr.Item(i).Visible = False
Next i

For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).Visible = False
    txtKPIp.Item(i).Visible = False
Next i


Me.MousePointer = vbHourglass

'Indicadores
strSQL = "exec spDSB_Bancos_Resumen Null , 'R', 'T'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    mCorte = rs!Corte
    lblCorte.Caption = Format(rs!Corte, "yyyy-mm-dd")

'    'Indices
    btnKPIp(0).Visible = True
    txtKPIp(0).Visible = True
    btnKPIp(0).Caption = "TEs"
    btnKPIp(0).Tag = "TEc"
    btnKPIp(0).ToolTipText = "No. Transferencias"
    txtKPIp(0).Text = Format(rs!TEc, "###,##0")


    btnKPIp(1).Visible = True
    txtKPIp(1).Visible = True
    btnKPIp(1).Caption = "CKs"
    btnKPIp(1).Tag = "CKc"
    txtKPIp(1).ToolTipText = "No. Cheques"
    txtKPIp(1).Text = Format(rs!CKc, "###,##0")
    
    btnKPIp(2).Visible = True
    txtKPIp(2).Visible = True
    btnKPIp(2).Caption = "DPs"
    btnKPIp(2).Tag = "DPc"
    btnKPIp(2).ToolTipText = "No. Depositos"
    txtKPIp(2).Text = Format(rs!DPc, "###,##0")

'    btnKPIp(3).Visible = True
'    txtKPIp(3).Visible = True
'    btnKPIp(3).Caption = "C.Jud"
'    btnKPIp(3).Tag = "ICbrJud"
'    btnKPIp(3).ToolTipText = "Ind. Cobro Judicial"
'    txtKPIp(3).Text = Format(rs!ICbrJud, "Standard") & "%"
'
'    btnKPIp(4).Visible = True
'    txtKPIp(4).Visible = True
'    btnKPIp(4).Caption = "Refin"
'    btnKPIp(4).Tag = "IRefinancia"
'    btnKPIp(4).ToolTipText = "Ind. Saldo Refinanciado"
'    txtKPIp(4).Text = Format(rs!IRefinancia, "Standard") & "%"
'
'    btnKPIp(5).Visible = True
'    txtKPIp(5).Visible = True
'    btnKPIp(5).Caption = "Canc"
'    btnKPIp(5).Tag = "ICancela"
'    btnKPIp(5).ToolTipText = "Ind. Cancelación Externa"
'    txtKPIp(5).Text = Format(rs!ICancela, "Standard") & "%"




    'Resultados

    btnKPIr(0).Visible = True
    txtKPIr(0).Visible = True
    btnKPIr(0).Tag = "Sal"
    btnKPIr(0).Caption = "Total de Desembolsos"
    txtKPIr(0).Text = Format(rs!Total, "###,##0")

    btnKPIr(1).Visible = True
    txtKPIr(1).Visible = True
    btnKPIr(1).Tag = "TEm"
    btnKPIr(1).Caption = "Transferencias"
    txtKPIr(1).Text = Format(rs!Transferencias, "###,##0")

    btnKPIr(2).Visible = True
    txtKPIr(2).Visible = True
    btnKPIr(2).Tag = "CKm"
    btnKPIr(2).Caption = "Cheques"
    txtKPIr(2).Text = Format(rs!Cheques, "###,##0")

    btnKPIr(3).Visible = True
    txtKPIr(3).Visible = True
    btnKPIr(3).Tag = "DPm"
    btnKPIr(3).Caption = "Depósitos"
    txtKPIr(3).Text = Format(rs!depositos, "###,##0")

'    btnKPIr(4).Visible = True
'    txtKPIr(4).Visible = True
'    btnKPIr(4).Caption = "ExA"
'    btnKPIr(4).Tag = "ExAsociados"
'    btnKPIr(4).ToolTipText = "Ex Asociados"
'    txtKPIr(4).Text = Format(rs!ExAsociados, "###,##0")

'    btnKPIr(5).Visible = True
'    txtKPIr(5).Visible = True
'    btnKPIr(5).Caption = "Ref"
'    btnKPIr(5).Tag = "Refinancia"
'    btnKPIr(5).ToolTipText = "Saldo Refinanciado"
'    txtKPIr(5).Text = Format(rs!Refinancia, "Standard")
'
'    btnKPIr(6).Visible = True
'    txtKPIr(6).Visible = True
'    btnKPIr(6).Caption = "Jud"
'    btnKPIr(6).Tag = "TCbrJud"
'    btnKPIr(6).ToolTipText = "Saldo en Cobro Judicial"
'    txtKPIr(6).Text = Format(rs!TCbrJud, "Standard")
'

End If
rs.Close

'Paso 2: Carga Graficos Principales

vPaso = True
    cboChart(0).Text = "3d_Pie"
vPaso = False


strSQL = "exec spDSB_Bancos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'B'"
Call sbChart_3D(chartC(0), "Bancos", "Bancos", cboChart(0).Text, 1)

pGra_C1.sql = strSQL
pGra_C1.Titulo = "Salida por Banco"
pGra_C1.Tema = "Bancos"
pGra_C1.Deciminal = 1
pGra_C1.Pattern = "N"
pGra_C1.Cifra = 1
pGra_C1.SQL_Filtro = "exec spDSB_Bancos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'B'"
pGra_C1.SQL_Sp = "exec spDSB_Bancos_Consulta"
    
'Grafico No.2
vPaso = True
    cboChart(1).Text = "3d_Torus"
vPaso = False

'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"
    
strSQL = "exec spDSB_Bancos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'I'"
Call sbChart_3D(chartC(1), "Cuenta Bancaria", "Documento", cboChart(1).Text, 1, "N", 0)

pGra_C2.sql = strSQL
pGra_C2.Titulo = "Cuenta Bancaria"
pGra_C2.Tema = "Cuenta"
pGra_C2.Deciminal = 1
pGra_C2.Pattern = "N"
pGra_C2.Cifra = 1
pGra_C2.SQL_Filtro = "exec spDSB_Bancos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'I'"
pGra_C2.SQL_Sp = "exec spDSB_Bancos_Consulta"
    
    
'Grafico No.3
vPaso = True
    cboChart(2).Text = "3d_Doughnut"
vPaso = False

strSQL = "exec spDSB_Bancos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'C'"
Call sbChart_3D(chartC(2), "Transac. Concepto", "Concepto", cboChart(2).Text, 1)

pGra_C3.sql = strSQL
pGra_C3.Titulo = "Transac. Concepto"
pGra_C3.Tema = "Concepto"
pGra_C3.Deciminal = 1
pGra_C3.Pattern = "N"
pGra_C3.Cifra = 1
pGra_C3.SQL_Filtro = "exec spDSB_Bancos_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'C'"
pGra_C3.SQL_Sp = "exec spDSB_Bancos_Consulta"


'Top 10
Call cboTop_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbContabilidad_Inicial()
Dim i As Integer

On Error GoTo vError

cboHFiltro.Clear

For i = 0 To btnKPIr.Count - 1
    btnKPIr.Item(i).Visible = False
    txtKPIr.Item(i).Visible = False
Next i

For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).Visible = False
    txtKPIp.Item(i).Visible = False
Next i


Me.MousePointer = vbHourglass

'Indicadores
strSQL = "exec spDSB_Contabilidad_Resumen Null , 'R', 'T'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    mCorte = rs!Corte
    lblCorte.Caption = Format(rs!Corte, "yyyy-mm-dd")

'    'Indices
    btnKPIp(0).Visible = True
    txtKPIp(0).Visible = True
    btnKPIp(0).Caption = "ROA"
    btnKPIp(0).Tag = "ROA"
    btnKPIp(0).ToolTipText = "Rend.S/Activos"
    txtKPIp(0).Text = Format(rs!ROA, "Standard") & "%"


    btnKPIp(1).Visible = True
    txtKPIp(1).Visible = True
    btnKPIp(1).Caption = "ROE"
    btnKPIp(1).Tag = "ROE"
    btnKPIp(1).ToolTipText = "Rend.S/Patrimonio"
    txtKPIp(1).Text = Format(rs!ROE, "Standard") & "%"
    
'    btnKPIp(2).Visible = True
'    txtKPIp(2).Visible = True
'    btnKPIp(2).Caption = "DPs"
'    btnKPIp(2).Tag = "DPc"
'    btnKPIp(2).ToolTipText = "No. Depositos"
'    txtKPIp(2).Text = Format(rs!DPc, "###,##0")

'    btnKPIp(3).Visible = True
'    txtKPIp(3).Visible = True
'    btnKPIp(3).Caption = "C.Jud"
'    btnKPIp(3).Tag = "ICbrJud"
'    btnKPIp(3).ToolTipText = "Ind. Cobro Judicial"
'    txtKPIp(3).Text = Format(rs!ICbrJud, "Standard") & "%"
'
'    btnKPIp(4).Visible = True
'    txtKPIp(4).Visible = True
'    btnKPIp(4).Caption = "Refin"
'    btnKPIp(4).Tag = "IRefinancia"
'    btnKPIp(4).ToolTipText = "Ind. Saldo Refinanciado"
'    txtKPIp(4).Text = Format(rs!IRefinancia, "Standard") & "%"
'
'    btnKPIp(5).Visible = True
'    txtKPIp(5).Visible = True
'    btnKPIp(5).Caption = "Canc"
'    btnKPIp(5).Tag = "ICancela"
'    btnKPIp(5).ToolTipText = "Ind. Cancelación Externa"
'    txtKPIp(5).Text = Format(rs!ICancela, "Standard") & "%"




    'Resultados

    btnKPIr(0).Visible = True
    txtKPIr(0).Visible = True
    btnKPIr(0).Tag = "U_Mes"
    btnKPIr(0).Caption = "Utilidad Mensual"
    txtKPIr(0).Text = Format(rs!U_Mes, "###,##0")

    btnKPIr(1).Visible = True
    txtKPIr(1).Visible = True
    btnKPIr(1).Tag = "U_Acumulada"
    btnKPIr(1).Caption = "Utilidad Acumulada"
    txtKPIr(1).Text = Format(rs!U_Acumulada, "###,##0")

    btnKPIr(2).Visible = True
    txtKPIr(2).Visible = True
    btnKPIr(2).Caption = "Ingresos"
    btnKPIr(2).Tag = "I"
    btnKPIr(2).ToolTipText = "Ingresos"
    txtKPIr(2).Text = Format(rs!Ingresos, "###,##0")

    btnKPIr(3).Visible = True
    txtKPIr(3).Visible = True
    btnKPIr(3).Caption = "Gastos"
    btnKPIr(3).Tag = "G"
    btnKPIr(3).ToolTipText = "Gastos"
    txtKPIr(3).Text = Format(rs!Gastos, "###,##0")

    btnKPIr(4).Visible = True
    txtKPIr(4).Visible = True
    btnKPIr(4).Caption = "Activos"
    btnKPIr(4).Tag = "A"
    btnKPIr(4).ToolTipText = "Activos"
    txtKPIr(4).Text = Format(rs!Activos, "###,##0")

    btnKPIr(5).Visible = True
    txtKPIr(5).Visible = True
    btnKPIr(5).Caption = "Pasivos"
    btnKPIr(5).Tag = "P"
    btnKPIr(5).ToolTipText = "Pasivos"
    txtKPIr(5).Text = Format(rs!Pasivos, "###,##0")

    btnKPIr(6).Visible = True
    txtKPIr(6).Visible = True
    btnKPIr(6).Caption = "Patrimonio"
    btnKPIr(6).Tag = "C"
    btnKPIr(6).ToolTipText = "Patrimonio"
    txtKPIr(6).Text = Format(rs!Patrimonio, "###,##0")


End If
rs.Close

'Paso 2: Carga Graficos Principales

vPaso = True
    cboChart(0).Text = "3d_Torus"
vPaso = False


strSQL = "exec spDSB_Contabilidad_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(0), "Gastos", "Gastos", cboChart(0).Text, 1)

pGra_C1.sql = strSQL
pGra_C1.Titulo = "Gastos"
pGra_C1.Tema = "Gastos"
pGra_C1.Deciminal = 1
pGra_C1.Pattern = "N"
pGra_C1.Cifra = 1
pGra_C1.SQL_Filtro = "exec spDSB_Contabilidad_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C1.SQL_Sp = "exec spDSB_Contabilidad_Consulta"
    
'Grafico No.2
vPaso = True
    cboChart(1).Text = "3d_Pie"
vPaso = False

'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"
    
strSQL = "exec spDSB_Contabilidad_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Res'"
Call sbChart_3D(chartC(1), "Resultados", "Resultado", cboChart(1).Text, 1, "N", 0)

pGra_C2.sql = strSQL
pGra_C2.Titulo = "Resultados"
pGra_C2.Tema = "Res"
pGra_C2.Deciminal = 1
pGra_C2.Pattern = "N"
pGra_C2.Cifra = 1
pGra_C2.SQL_Filtro = "exec spDSB_Contabilidad_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Res'"
pGra_C2.SQL_Sp = "exec spDSB_Contabilidad_Consulta"
    
    
'Grafico No.3
vPaso = True
    cboChart(2).Text = "3d_Doughnut"
vPaso = False

strSQL = "exec spDSB_Contabilidad_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Bal'"
Call sbChart_3D(chartC(2), "Balance", "Balance", cboChart(2).Text, 1)

pGra_C3.sql = strSQL
pGra_C3.Titulo = "Balance"
pGra_C3.Tema = "Bal"
pGra_C3.Deciminal = 1
pGra_C3.Pattern = "N"
pGra_C3.Cifra = 1
pGra_C3.SQL_Filtro = "exec spDSB_Contabilidad_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Bal'"
pGra_C3.SQL_Sp = "exec spDSB_Contabilidad_Consulta"


'Top 10
Call cboTop_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbInversiones_Inicial()
Dim i As Integer

On Error GoTo vError

cboHFiltro.Clear

For i = 0 To btnKPIr.Count - 1
    btnKPIr.Item(i).Visible = False
    txtKPIr.Item(i).Visible = False
Next i

For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).Visible = False
    txtKPIp.Item(i).Visible = False
Next i


Me.MousePointer = vbHourglass

'Indicadores
strSQL = "exec spDSB_Inversiones_Resumen Null , 'R', 'T'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    mCorte = rs!Corte
    lblCorte.Caption = Format(rs!Corte, "yyyy-mm-dd")

'    'Indices
    btnKPIp(0).Visible = True
    txtKPIp(0).Visible = True
    btnKPIp(0).Caption = "TEA"
    btnKPIp(0).Tag = "TEA"
    btnKPIp(0).ToolTipText = "Tasa Efectiva Ponderada"
    txtKPIp(0).Text = Format(rs!TASA_EFECTIVA, "Standard") & "%"
'
'
'    btnKPIp(1).Visible = True
'    txtKPIp(1).Visible = True
'    btnKPIp(1).Caption = "ROE"
'    btnKPIp(1).Tag = "ROE"
'    btnKPIp(1).ToolTipText = "Rend.S/Patrimonio"
'    txtKPIp(1).Text = Format(rs!ROE, "Standard") & "%"
    
'    btnKPIp(2).Visible = True
'    txtKPIp(2).Visible = True
'    btnKPIp(2).Caption = "DPs"
'    btnKPIp(2).Tag = "DPc"
'    btnKPIp(2).ToolTipText = "No. Depositos"
'    txtKPIp(2).Text = Format(rs!DPc, "###,##0")

'    btnKPIp(3).Visible = True
'    txtKPIp(3).Visible = True
'    btnKPIp(3).Caption = "C.Jud"
'    btnKPIp(3).Tag = "ICbrJud"
'    btnKPIp(3).ToolTipText = "Ind. Cobro Judicial"
'    txtKPIp(3).Text = Format(rs!ICbrJud, "Standard") & "%"
'
'    btnKPIp(4).Visible = True
'    txtKPIp(4).Visible = True
'    btnKPIp(4).Caption = "Refin"
'    btnKPIp(4).Tag = "IRefinancia"
'    btnKPIp(4).ToolTipText = "Ind. Saldo Refinanciado"
'    txtKPIp(4).Text = Format(rs!IRefinancia, "Standard") & "%"
'
'    btnKPIp(5).Visible = True
'    txtKPIp(5).Visible = True
'    btnKPIp(5).Caption = "Canc"
'    btnKPIp(5).Tag = "ICancela"
'    btnKPIp(5).ToolTipText = "Ind. Cancelación Externa"
'    txtKPIp(5).Text = Format(rs!ICancela, "Standard") & "%"

'                    , SUM(B.CASOS) AS 'CASOS'
'                    , SUM(B.VALOR_LIBROS) AS 'VALOR_LIBROS'
'                    , SUM(B.PYD_SALDO ) AS 'PYD_SALDO'
'                    , SUM(B.INTERES_ACUM_MONTO) AS 'INTERES_ACUM_MONTO'
'                    , SUM(B.TASA_EFECTIVA) AS 'TASA_EFECTIVA'


    'Resultados

    btnKPIr(0).Visible = True
    txtKPIr(0).Visible = True
    btnKPIr(0).Tag = "VL"
    btnKPIr(0).Caption = "Valor en Libros"
    txtKPIr(0).Text = Format(rs!Valor_Libros, "###,##0")

    btnKPIr(1).Visible = True
    txtKPIr(1).Visible = True
    btnKPIr(1).Tag = "PYD"
    btnKPIr(1).Caption = "Primas y Descuentos"
    txtKPIr(1).Text = Format(rs!PYD_SALDO, "###,##0")

    btnKPIr(2).Visible = True
    txtKPIr(2).Visible = True
    btnKPIr(2).Caption = "Interés Acumulado"
    btnKPIr(2).Tag = "IA"
    txtKPIr(2).Text = Format(rs!INTERES_ACUM_MONTO, "###,##0")

    btnKPIr(3).Visible = True
    txtKPIr(3).Visible = True
    btnKPIr(3).Caption = "Interés del mes"
    btnKPIr(3).Tag = "IM"
    txtKPIr(3).Text = Format(rs!INTERES_MES, "###,##0")

'    btnKPIr(4).Visible = True
'    txtKPIr(4).Visible = True
'    btnKPIr(4).Caption = "Activos"
'    btnKPIr(4).Tag = "A"
'    btnKPIr(4).ToolTipText = "Activos"
'    txtKPIr(4).Text = Format(rs!Activos, "###,##0")
'
'    btnKPIr(5).Visible = True
'    txtKPIr(5).Visible = True
'    btnKPIr(5).Caption = "Pasivos"
'    btnKPIr(5).Tag = "P"
'    btnKPIr(5).ToolTipText = "Pasivos"
'    txtKPIr(5).Text = Format(rs!Pasivos, "###,##0")
'
'    btnKPIr(6).Visible = True
'    txtKPIr(6).Visible = True
'    btnKPIr(6).Caption = "Patrimonio"
'    btnKPIr(6).Tag = "C"
'    btnKPIr(6).ToolTipText = "Patrimonio"
'    txtKPIr(6).Text = Format(rs!Patrimonio, "###,##0")


End If
rs.Close

'Paso 2: Carga Graficos Principales

vPaso = True
    cboChart(0).Text = "3d_Doughnut"
vPaso = False


strSQL = "exec spDSB_Inversiones_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Inst'"
Call sbChart_3D(chartC(0), "Instrumentos", "Instrumentos", cboChart(0).Text, 1)

pGra_C1.sql = strSQL
pGra_C1.Titulo = "Instrumentos"
pGra_C1.Tema = "Inst"
pGra_C1.Deciminal = 1
pGra_C1.Pattern = "N"
pGra_C1.Cifra = 1
pGra_C1.SQL_Filtro = "exec spDSB_Inversiones_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Inst'"
pGra_C1.SQL_Sp = "exec spDSB_Inversiones_Consulta"
    
'Grafico No.2
vPaso = True
    cboChart(1).Text = "3d_Pie"
vPaso = False

'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"
    
strSQL = "exec spDSB_Inversiones_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Adm'"
Call sbChart_3D(chartC(1), "Administradores", "Administradores", cboChart(1).Text, 1, "N", 0)

pGra_C2.sql = strSQL
pGra_C2.Titulo = "Administradores"
pGra_C2.Tema = "Adm"
pGra_C2.Deciminal = 1
pGra_C2.Pattern = "N"
pGra_C2.Cifra = 1
pGra_C2.SQL_Filtro = "exec spDSB_Inversiones_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Adm'"
pGra_C2.SQL_Sp = "exec spDSB_Inversiones_Consulta"
    
    
'Grafico No.3
vPaso = True
    cboChart(2).Text = "3d_Torus"
vPaso = False

strSQL = "exec spDSB_Inversiones_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Cat'"
Call sbChart_3D(chartC(2), "Categorias", "Categorias", cboChart(2).Text, 1)

pGra_C3.sql = strSQL
pGra_C3.Titulo = "Categorias"
pGra_C3.Tema = "Cat"
pGra_C3.Deciminal = 1
pGra_C3.Pattern = "N"
pGra_C3.Cifra = 1
pGra_C3.SQL_Filtro = "exec spDSB_Inversiones_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Cat'"
pGra_C3.SQL_Sp = "exec spDSB_Inversiones_Consulta"

'Top 10
Call cboTop_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbCaptacion_Inicial()
Dim i As Integer

On Error GoTo vError

cboHFiltro.Clear

For i = 0 To btnKPIr.Count - 1
    btnKPIr.Item(i).Visible = False
    txtKPIr.Item(i).Visible = False
Next i

For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).Visible = False
    txtKPIp.Item(i).Visible = False
Next i


Me.MousePointer = vbHourglass

'Indicadores
strSQL = "exec spDSB_Captacion_Resumen Null , 'R', 'T'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

    mCorte = rs!Corte
    lblCorte.Caption = Format(rs!Corte, "yyyy-mm-dd")

'    'Indices
'    btnKPIp(0).Visible = True
'    txtKPIp(0).Visible = True
'    btnKPIp(0).Caption = "TEA"
'    btnKPIp(0).Tag = "TEA"
'    btnKPIp(0).ToolTipText = "Tasa Efectiva Ponderada"
'    txtKPIp(0).Text = Format(rs!TASA_EFECTIVA, "Standard") & "%"
'
'
'    btnKPIp(1).Visible = True
'    txtKPIp(1).Visible = True
'    btnKPIp(1).Caption = "ROE"
'    btnKPIp(1).Tag = "ROE"
'    btnKPIp(1).ToolTipText = "Rend.S/Patrimonio"
'    txtKPIp(1).Text = Format(rs!ROE, "Standard") & "%"
    
'    btnKPIp(2).Visible = True
'    txtKPIp(2).Visible = True
'    btnKPIp(2).Caption = "DPs"
'    btnKPIp(2).Tag = "DPc"
'    btnKPIp(2).ToolTipText = "No. Depositos"
'    txtKPIp(2).Text = Format(rs!DPc, "###,##0")

'    btnKPIp(3).Visible = True
'    txtKPIp(3).Visible = True
'    btnKPIp(3).Caption = "C.Jud"
'    btnKPIp(3).Tag = "ICbrJud"
'    btnKPIp(3).ToolTipText = "Ind. Cobro Judicial"
'    txtKPIp(3).Text = Format(rs!ICbrJud, "Standard") & "%"
'
'    btnKPIp(4).Visible = True
'    txtKPIp(4).Visible = True
'    btnKPIp(4).Caption = "Refin"
'    btnKPIp(4).Tag = "IRefinancia"
'    btnKPIp(4).ToolTipText = "Ind. Saldo Refinanciado"
'    txtKPIp(4).Text = Format(rs!IRefinancia, "Standard") & "%"
'
'    btnKPIp(5).Visible = True
'    txtKPIp(5).Visible = True
'    btnKPIp(5).Caption = "Canc"
'    btnKPIp(5).Tag = "ICancela"
'    btnKPIp(5).ToolTipText = "Ind. Cancelación Externa"
'    txtKPIp(5).Text = Format(rs!ICancela, "Standard") & "%"

'                    , SUM(B.CASOS) AS 'CASOS'
'                    , SUM(B.VALOR_LIBROS) AS 'VALOR_LIBROS'
'                    , SUM(B.PYD_SALDO ) AS 'PYD_SALDO'
'                    , SUM(B.INTERES_ACUM_MONTO) AS 'INTERES_ACUM_MONTO'
'                    , SUM(B.TASA_EFECTIVA) AS 'TASA_EFECTIVA'


    'Resultados

    btnKPIr(0).Visible = True
    txtKPIr(0).Visible = True
    btnKPIr(0).Tag = "CNT"
    btnKPIr(0).Caption = "Contratos"
    txtKPIr(0).Text = Format(rs!Casos, "###,##0")

    btnKPIr(1).Visible = True
    txtKPIr(1).Visible = True
    btnKPIr(1).Tag = "TOT"
    btnKPIr(1).Caption = "Total"
    txtKPIr(1).Text = Format(rs!Total, "###,##0")

    btnKPIr(2).Visible = True
    txtKPIr(2).Visible = True
    btnKPIr(2).Caption = "Aporte Acumulado"
    btnKPIr(2).Tag = "AP"
    txtKPIr(2).Text = Format(rs!APORTES, "###,##0")

    btnKPIr(3).Visible = True
    txtKPIr(3).Visible = True
    btnKPIr(3).Caption = "Rendimiento Acumulado"
    btnKPIr(3).Tag = "RND"
    txtKPIr(3).Text = Format(rs!Rendimiento, "###,##0")

    btnKPIr(4).Visible = True
    txtKPIr(4).Visible = True
    btnKPIr(4).Caption = "Aportaciones"
    btnKPIr(4).Tag = "Aportes"
    txtKPIr(4).Text = Format(rs!Aportaciones, "###,##0")

    btnKPIr(5).Visible = True
    txtKPIr(5).Visible = True
    btnKPIr(5).Caption = "Retiros"
    btnKPIr(5).Tag = "Retiros"
    txtKPIr(5).Text = Format(rs!Retiros, "###,##0")

'    btnKPIr(6).Visible = True
'    txtKPIr(6).Visible = True
'    btnKPIr(6).Caption = "Patrimonio"
'    btnKPIr(6).Tag = "C"
'    btnKPIr(6).ToolTipText = "Patrimonio"
'    txtKPIr(6).Text = Format(rs!Patrimonio, "###,##0")


End If
rs.Close

'Paso 2: Carga Graficos Principales

vPaso = True
    cboChart(0).Text = "3d_Doughnut"
vPaso = False


strSQL = "exec spDSB_Captacion_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Plan'"
Call sbChart_3D(chartC(0), "Planes Top 5", "¨Planes", cboChart(0).Text, 1)

pGra_C1.sql = strSQL
pGra_C1.Titulo = "Planes"
pGra_C1.Tema = "Plan"
pGra_C1.Deciminal = 1
pGra_C1.Pattern = "N"
pGra_C1.Cifra = 1
pGra_C1.SQL_Filtro = "exec spDSB_Captacion_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Plan'"
pGra_C1.SQL_Sp = "exec spDSB_Captacion_Consulta"
    
'Grafico No.2
vPaso = True
    cboChart(1).Text = "3d_Pie"
vPaso = False

'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"
    
strSQL = "exec spDSB_Captacion_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'Grupo'"
Call sbChart_3D(chartC(1), "Grupo Productos", "Grupos", cboChart(1).Text, 1, "N", 0)

pGra_C2.sql = strSQL
pGra_C2.Titulo = "Grupo Productos"
pGra_C2.Tema = "Grupo"
pGra_C2.Deciminal = 1
pGra_C2.Pattern = "N"
pGra_C2.Cifra = 1
pGra_C2.SQL_Filtro = "exec spDSB_Captacion_Consulta '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'Grupo'"
pGra_C2.SQL_Sp = "exec spDSB_Captacion_Consulta"
    
    
'Grafico No.3
vPaso = True
    cboChart(2).Text = "3d_Torus"
vPaso = False

strSQL = "exec spDSB_Clientes_Consulta_Patrimonio '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'R', 'G'"
Call sbChart_3D(chartC(2), "Patrimonio", "Rubro", cboChart(2).Text, 1)

pGra_C3.sql = strSQL
pGra_C3.Titulo = "Patrimonio"
pGra_C3.Tema = "Rubro"
pGra_C3.Deciminal = 1
pGra_C3.Pattern = "N"
pGra_C3.Cifra = 1
pGra_C3.SQL_Filtro = "exec spDSB_Clientes_Consulta_Patrimonio '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'F', 'G'"
pGra_C3.SQL_Sp = "exec spDSB_Clientes_Consulta_Patrimonio"

'Top 10
Call cboTop_Click

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub sbMenu()

On Error GoTo vError

Me.MousePointer = vbHourglass

lswMenu.ListItems.Clear
With lswMenu.ColumnHeaders
    .Clear
    .Add , , "", (lswMenu.Width - 50)
End With

strSQL = "exec spDSB_Main_Categorias_Access '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    Set itmX = lswMenu.ListItems.Add(, , rs!Descripcion)
        itmX.Tag = rs!cod_categoria
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub sbDashboard_Load()

On Error GoTo vError

Select Case mCategoria
    Case "CLI" 'Clientes
        Call sbClientes_Inicial
    
    Case "CRD" 'Credito y Cobro
        Call sbCreditos_Inicial
    
    Case "FND" 'Ahorros
        Call sbCaptacion_Inicial
    
    Case "FIN" 'Financieros
        Call sbContabilidad_Inicial
    
    
    Case "TES" 'Bancos y Cajas
        Call sbBancos_Inicial
        
    Case "IVR" 'Inversiones
        Call sbInversiones_Inicial
        
    Case "BEN" 'Beneficios
End Select

'Top List
vPaso = True
    strSQL = "exec spDSB_Main_KPI_Access '" & glogon.Usuario & "', 'T', '" & mCategoria & "'"
    Call sbCbo_Llena_New(cboTopList, strSQL, False, True)
vPaso = False

'Inicializa
Call btnKPIp_Click(0)

'Resultados del Top List
cboTopList.SetFocus
Call cboTop_Click

Me.Refresh

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub





Private Sub btnKPIp_Click(Index As Integer)

On Error GoTo vError
 
If Not btnKPIp(Index).Visible Then
  Exit Sub
End If

Select Case mCategoria
    Case "CLI" 'Clientes
        strSQL = "exec spDSB_Clientes_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIp(Index).Tag & "'"
    
        If strSQL <> "" Then
            Call sbChart_Histograma(chartH, btnKPIp.Item(Index).ToolTipText, btnKPIp.Item(Index).Caption, 100, "P")
        End If
    
    Case "CRD" 'Credito y Cobro
        strSQL = "exec spDSB_Creditos_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIp(Index).Tag & "'"
    
        If strSQL <> "" Then
            Call sbChart_Histograma(chartH, btnKPIp.Item(Index).ToolTipText, btnKPIp.Item(Index).Caption, 100, "P")
        End If
    
    Case "FND" 'Ahorros
        strSQL = "exec spDSB_Captacion_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIp(Index).Tag & "'"

        If strSQL <> "" Then
            Call sbChart_Histograma(chartH, btnKPIp.Item(Index).ToolTipText, btnKPIp.Item(Index).ToolTipText, 100, "P")
        End If
    
    
    Case "FIN" 'Financieros
        strSQL = "exec spDSB_Contabilidad_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIp(Index).Tag & "'"

        If strSQL <> "" Then
            Call sbChart_Histograma(chartH, btnKPIp.Item(Index).ToolTipText, btnKPIp.Item(Index).ToolTipText, 100, "P")
        End If

    Case "TES" 'Bancos y Cajas
        strSQL = "exec spDSB_Bancos_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIp(Index).Tag & "'"
        

        If strSQL <> "" Then
            Call sbChart_Histograma(chartH, btnKPIp.Item(Index).ToolTipText, btnKPIp.Item(Index).Caption, 1, "N")
        End If

    Case "IVR" 'Inversiones
    
        strSQL = "exec spDSB_Inversiones_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIp(Index).Tag & "'"

        If strSQL <> "" Then
            Call sbChart_Histograma(chartH, btnKPIp.Item(Index).ToolTipText, btnKPIp.Item(Index).ToolTipText, 100, "P")
        End If
    Case "BEN" 'Beneficios
End Select



 
Exit Sub
 
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnKPIr_Click(Index As Integer)

On Error GoTo vError

Select Case mCategoria
    Case "CLI" 'Clientes
        strSQL = "exec spDSB_Clientes_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIr(Index).Tag & "'"
    
    Case "CRD" 'Credito y Cobro
        strSQL = "exec spDSB_Creditos_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIr(Index).Tag & "'"
    
    Case "FND" 'Ahorros
        strSQL = "exec spDSB_Captacion_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIr(Index).Tag & "'"
    
    Case "FIN" 'Financieros
        strSQL = "exec spDSB_Contabilidad_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIr(Index).Tag & "'"
    
    Case "TES" 'Bancos y Cajas
        strSQL = "exec spDSB_Bancos_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIr(Index).Tag & "'"
        
    Case "IVR" 'Inversiones
        strSQL = "exec spDSB_Inversiones_Consulta_Histograma '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , '" & btnKPIr(Index).Tag & "'"
    
    
    Case "BEN" 'Beneficios
End Select

If strSQL <> "" Then
    Call sbChart_Histograma(chartH, btnKPIr.Item(Index).Caption, btnKPIr.Item(Index).Caption, 1, "N")
End If

Exit Sub
 
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnMenu_Click()
If gbMenu.Visible Then
    gbMenu.Visible = False
Else
    gbMenu.Left = 0
    gbMenu.top = 1080
    gbMenu.Visible = True
End If
End Sub

Private Sub btnPrint_Click(Index As Integer)


Select Case Index
 Case 0, 1, 2
    If (chartC(Index).PrintPreview) Then
        chartC(Index).PrintChart 0
    End If
    
   Case 3
    If (chartH.PrintPreview) Then
        chartH.PrintChart 0
    End If
End Select
    
End Sub

Private Sub btnSave_Click(Index As Integer)


Select Case Index
 Case 0, 1, 2
    chartC(Index).SaveAsImage SIFGlobal.DirectorioDeResultados + "\Chart.png", chartC(Index).Width, chartC(Index).Height
 Case 3
    chartH.SaveAsImage SIFGlobal.DirectorioDeResultados + "\Chart.png", chartH.Width, chartH.Height
End Select
    
    ShellExecute Me.hWnd, "open", SIFGlobal.DirectorioDeResultados + "\Chart.png", vbNullString, vbNullString, 1
    
End Sub

Private Sub cboAppearance_Click()
Dim i As Integer

'mChartPallete = cboPalette.List(cboPalette.ListIndex)

chartH.Content.Appearance.SetAppearance cboAppearance.List(cboAppearance.ListIndex)
chartC(0).Content.Appearance.SetAppearance cboAppearance.List(cboAppearance.ListIndex)
chartC(1).Content.Appearance.SetAppearance cboAppearance.List(cboAppearance.ListIndex)
chartC(2).Content.Appearance.SetAppearance cboAppearance.List(cboAppearance.ListIndex)


If cboAppearance.List(cboAppearance.ListIndex) = "Black" Then
    lswTop.BackColor = vbBlack
    lswTop.ForeColor = vbGreen
    
    cboChart(0).BackColor = vbBlack
    cboChart(0).ForeColor = vbGreen
    
    For i = 0 To btnKPIp.Count - 1
        btnKPIp.Item(i).ForeColor = vbGreen
        btnKPIp.Item(i).BackColor = vbBlack
    Next i
    
Else
    lswTop.BackColor = vbWhite
    lswTop.ForeColor = vbBlack

    cboChart(0).BackColor = vbWhite
    cboChart(0).ForeColor = vbGrayed


    For i = 0 To btnKPIp.Count - 1
        btnKPIp.Item(i).ForeColor = RGB(78, 111, 178)
        btnKPIp.Item(i).BackColor = RGB(214, 234, 248)
    Next i


End If

cboChart(1).BackColor = cboChart(0).BackColor
cboChart(1).ForeColor = cboChart(0).ForeColor
cboChart(2).BackColor = cboChart(0).BackColor
cboChart(2).ForeColor = cboChart(0).ForeColor

btnSave(0).BackColor = cboChart(0).BackColor
btnSave(0).ForeColor = cboChart(0).ForeColor
btnSave(1).BackColor = cboChart(0).BackColor
btnSave(1).ForeColor = cboChart(0).ForeColor
btnSave(2).BackColor = cboChart(0).BackColor
btnSave(2).ForeColor = cboChart(0).ForeColor


btnPrint(0).BackColor = cboChart(0).BackColor
btnPrint(0).ForeColor = cboChart(0).ForeColor
btnPrint(1).BackColor = cboChart(0).BackColor
btnPrint(1).ForeColor = cboChart(0).ForeColor
btnPrint(2).BackColor = cboChart(0).BackColor
btnPrint(2).ForeColor = cboChart(0).ForeColor

gbKPI.BackColor = cboChart(0).BackColor

lblBack.BackColor = cboChart(0).BackColor


For i = 0 To 9
    txtKPIp(i).BackColor = cboChart(0).BackColor
    txtKPIp(i).ForeColor = cboChart(0).ForeColor
Next i

End Sub

Private Sub cboChart_Click(Index As Integer)

If vPaso Then Exit Sub

Select Case Index
  Case 0
    strSQL = pGra_C1.sql
    Call sbChart_3D(chartC(Index), pGra_C1.Titulo, pGra_C1.Tema, cboChart(Index).Text)
  Case 1
    strSQL = pGra_C2.sql
    Call sbChart_3D(chartC(Index), pGra_C2.Titulo, pGra_C2.Tema, cboChart(Index).Text)
  Case 2
    strSQL = pGra_C3.sql
    Call sbChart_3D(chartC(Index), pGra_C3.Titulo, pGra_C3.Tema, cboChart(Index).Text)
End Select



End Sub

Private Sub cboChart_H_Click()
If vPaso Then Exit Sub

 On Error GoTo vError

Dim pFiltro As String

If cboHFiltro.Text = "TODOS" Or cboHFiltro.ListCount = 0 Then
    pFiltro = "Null"
Else
    pFiltro = "'" & cboHFiltro.ItemData(cboHFiltro.ListIndex) & "'"
End If

'strSQL = pGra_H.SQL_Sp & " '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'H', 'G', " & pFiltro

strSQL = Replace(pGra_H.SQL_Filtro, "'F'", "'H'") & ", " & pFiltro

Call sbChart_MultiSeries(chartH, pGra_H.Titulo, pGra_H.Tema, pGra_H.Cifra, pGra_H.Pattern, 0, cboChart_H.Text)
 
 Exit Sub
 
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboHFiltro_Click()
If vPaso Then Exit Sub

Call cboChart_H_Click

End Sub


Private Sub cboTop_Click()
Dim pFechaInicio As Date, pFechaCorte As Date

On Error GoTo vError

If cboTopList.ListCount = 0 Then Exit Sub

pFechaCorte = fxFechaServidor

Select Case cboTop.Text
    Case "  7 días"
        pFechaInicio = DateAdd("d", -7, pFechaCorte)
    Case " 15 días"
        pFechaInicio = DateAdd("d", -15, pFechaCorte)
    Case " 30 días"
        pFechaInicio = DateAdd("d", -30, pFechaCorte)
    Case " 60 días"
        pFechaInicio = DateAdd("d", -60, pFechaCorte)
    Case "120 días"
        pFechaInicio = DateAdd("d", -120, pFechaCorte)
    Case "180 días"
        pFechaInicio = DateAdd("d", -180, pFechaCorte)
    Case "365 días"
        pFechaInicio = DateAdd("d", -365, pFechaCorte)

End Select

strSQL = ""
lswTop.ListItems.Clear

strSQL = "exec spDSB_Main_Top_Executor  '" & cboTopList.ItemData(cboTopList.ListIndex) _
        & "', '" & Format(pFechaInicio, "yyyy-mm-dd") & " 00:00','" & Format(pFechaCorte, "yyyy-mm-dd") & " 23:59'" _
        & ", " & cboTopCount.ItemData(cboTopCount.ListIndex)

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswTop.ListItems.Add(, , rs!Descripcion)
     itmX.SubItems(1) = Format(rs!Dato, "###,###")
 rs.MoveNext
Loop
rs.Close


Exit Sub

vError:


End Sub

Private Sub cboTopCount_Click()
Call cboTop_Click
End Sub

Private Sub cboTopList_Click()
Call cboTop_Click
End Sub

Private Sub chartC_DblClick(Index As Integer, Button As Integer, Shift As Integer, x As Long, y As Long)
On Error GoTo vError

cboHFiltro.Clear

Select Case Index
    Case 0 'Grafico 1
        'strSQL = pGra_C1.SQL_Sp & " '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'H', 'G'"
         
        strSQL = Replace(pGra_C1.SQL_Filtro, "'F'", "'H'")
         
        pGra_H.sql = pGra_C1.sql
        pGra_H.Titulo = pGra_C1.Titulo
        pGra_H.Tema = pGra_C1.Tema
        pGra_H.Deciminal = pGra_C1.Deciminal
        pGra_H.Cifra = pGra_C1.Cifra
        pGra_H.Pattern = pGra_C1.Pattern
        pGra_H.SQL_Filtro = pGra_C1.SQL_Filtro
        pGra_H.SQL_Sp = pGra_C1.SQL_Sp
         
    Case 1 'Grafico 2
'        strSQL = pGra_C2.SQL_Sp & " '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'H', 'G'"
         strSQL = Replace(pGra_C2.SQL_Filtro, "'F'", "'H'")
        
        pGra_H.sql = pGra_C2.sql
        pGra_H.Titulo = pGra_C2.Titulo
        pGra_H.Tema = pGra_C2.Tema
        pGra_H.Deciminal = pGra_C2.Deciminal
        pGra_H.Cifra = pGra_C2.Cifra
        pGra_H.Pattern = pGra_C2.Pattern
        pGra_H.SQL_Filtro = pGra_C2.SQL_Filtro
        pGra_H.SQL_Sp = pGra_C2.SQL_Sp
         
 
    Case 2 'Grafico 3
'        strSQL = pGra_C3.SQL_Sp & " '" & Format(mCorte, "yyyy-mm-dd") & " 23:59' , 'H', 'G'"
        strSQL = Replace(pGra_C3.SQL_Filtro, "'F'", "'H'")
         
        pGra_H.sql = pGra_C3.sql
        pGra_H.Titulo = pGra_C3.Titulo
        pGra_H.Tema = pGra_C3.Tema
        pGra_H.Deciminal = pGra_C3.Deciminal
        pGra_H.Cifra = pGra_C3.Cifra
        pGra_H.Pattern = pGra_C3.Pattern
        pGra_H.SQL_Filtro = pGra_C3.SQL_Filtro
        pGra_H.SQL_Sp = pGra_C3.SQL_Sp
         
End Select

Call sbChart_MultiSeries(chartH, pGra_H.Titulo, pGra_H.Tema, pGra_H.Cifra, pGra_H.Pattern, 0, cboChart_H.Text)


Exit Sub
 
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub sbAppearance_Load(cboA As XtremeSuiteControls.ComboBox)

cboA.Clear
cboA.AddItem "Nature"
cboA.AddItem "Black"
cboA.AddItem "Gray"

cboA.Text = "Nature"

End Sub

Private Sub sbPalette_Load(cboP As XtremeSuiteControls.ComboBox)

cboP.Clear
cboP.AddItem "Victorian"
cboP.AddItem "Vibrant Pastel"
cboP.AddItem "Vibrant"
cboP.AddItem "Tropical"
cboP.AddItem "Summer"
cboP.AddItem "Spring Time"
cboP.AddItem "Rainbow"
cboP.AddItem "Purple"
cboP.AddItem "Primary Colors"
cboP.AddItem "Postmodern"
cboP.AddItem "Photodesign"
cboP.AddItem "Pastel"
cboP.AddItem "Office"
cboP.AddItem "Orange Green"
cboP.AddItem "Nature"
cboP.AddItem "Natural"
cboP.AddItem "Impresionism"
cboP.AddItem "Illustration"
cboP.AddItem "Harvest"
cboP.AddItem "Green Brown"
cboP.AddItem "Green Blue"
cboP.AddItem "Green"
cboP.AddItem "Gray"
cboP.AddItem "Four Color"
cboP.AddItem "Fire"
cboP.AddItem "Earth Tone"
cboP.AddItem "Danville"
cboP.AddItem "Caribbean"
cboP.AddItem "Cappuccino"
cboP.AddItem "Blue Gray"
cboP.AddItem "Blue"
cboP.Text = mChartPallete

End Sub

Private Sub cboPalette_Click()

mChartPallete = cboPalette.List(cboPalette.ListIndex)

chartH.Content.Appearance.SetPalette cboPalette.List(cboPalette.ListIndex)
chartC(0).Content.Appearance.SetPalette cboPalette.List(cboPalette.ListIndex)
chartC(1).Content.Appearance.SetPalette cboPalette.List(cboPalette.ListIndex)
chartC(2).Content.Appearance.SetPalette cboPalette.List(cboPalette.ListIndex)



End Sub

Private Sub Form_Load()

Dim i As Integer


vModulo = 24

Me.BackColor = RGB(78, 111, 178)

mChartPallete = "Danville"
mChartLabelPosition = 3

Call sbPalette_Load(cboPalette)
Call sbAppearance_Load(cboAppearance)

Me.BackColor = vbWhite

lblMenu.ForeColor = RGB(15, 133, 215)

lblDashboard.ForeColor = lblMenu.ForeColor
gbMenu.BackColor = lblMenu.ForeColor
gbMenu.Visible = False

lswMenu.ListItems.Clear
lswMenu.ColumnHeaders.Add , , "Descripción", lswMenu.Width - 100
lswMenu.HideColumnHeaders = True

lblCorte.ForeColor = lblMenu.ForeColor
lswMenu.ForeColor = lblMenu.ForeColor

lswTop.ListItems.Clear
lswTop.ColumnHeaders.Add , , "Descripción", 2500
lswTop.ColumnHeaders.Add , , "Resultado", lswTop.Width - 2600, vbRightJustify
lswTop.HideColumnHeaders = True



vPaso = True
    
    With cboTopCount
     .AddItem "Top 10"
     .ItemData(.ListCount - 1) = "10"
     .AddItem "Top 25"
     .ItemData(.ListCount - 1) = "25"
     .AddItem "Top 50"
     .ItemData(.ListCount - 1) = "50"
     .AddItem "Top 100"
     .ItemData(.ListCount - 1) = "100"
     
     .Text = "Top 10"
    End With
    
    
    cboTop.Clear
    cboTop.AddItem "  7 días"
    cboTop.AddItem " 15 días"
    cboTop.AddItem " 30 días"
    cboTop.AddItem " 60 días"
    cboTop.AddItem "120 días"
    cboTop.AddItem "180 días"
    cboTop.AddItem "365 días"
    cboTop.Text = " 30 días"
    
    cboChart_H.Clear
    cboChart_H.AddItem "Spline"
    cboChart_H.AddItem "Bar"
    cboChart_H.AddItem "Point"
    cboChart_H.AddItem "Area"
    cboChart_H.AddItem "Multi"
    cboChart_H.Text = "Spline"
    
    cboChart(0).Clear
    cboChart(0).AddItem "3d_Pie"
    cboChart(0).AddItem "3d_Pyramid"
    cboChart(0).AddItem "3d_Torus"
    cboChart(0).AddItem "3d_Doughnut"
    cboChart(0).AddItem "3d_Funnel"
    cboChart(0).AddItem "Pie"
    cboChart(0).AddItem "Pyramid"
    
    
    cboChart(0).Text = "3d_Torus"
    
    
    Call sbCbo_Copia(cboChart(0), cboChart(1))
    Call sbCbo_Copia(cboChart(0), cboChart(2))
    cboChart(1).Text = "3d_Pyramid"
    cboChart(2).Text = "3d_Torus"
    
vPaso = False

For i = 0 To btnKPIr.Count - 1
'    btnKPIr.Item(i).ForeColor = RGB(78, 111, 178)
'    btnKPIr.Item(i).BackColor = RGB(214, 234, 248)

    btnKPIr.Item(i).ForeColor = RGB(78, 111, 178)
    btnKPIr.Item(i).BackColor = vbWhite
Next i


For i = 0 To btnKPIp.Count - 1
    btnKPIp.Item(i).ForeColor = RGB(78, 111, 178)
    btnKPIp.Item(i).BackColor = RGB(214, 234, 248)
Next i

Call cboAppearance_Click


'Seguridad
Call Formularios(Me)





End Sub


Private Sub Form_Resize()
On Error Resume Next


gbKPI.Left = Me.Width - (gbKPI.Width + 250)

chartH.Width = Me.Width - (chartH.Left + 150 + gbKPI.Width + scTop.Width + 210)

scHistograma.Width = chartH.Width

scTop.Left = chartH.Left + chartH.Width + 60
lswTop.Left = scTop.Left


chartC(0).Left = chartH.Left
chartC(0).Width = (gbKPI.Left / 3) - 80

chartC(0).Height = Me.Height - (chartC(0).top + 650)

chartC(1).Left = chartC(0).Left + chartC(0).Width + 110
chartC(1).Width = chartC(0).Width
chartC(1).Height = chartC(0).Height

chartC(2).Left = chartC(1).Left + chartC(1).Width + 110
chartC(2).Width = chartC(0).Width
chartC(2).Height = chartC(0).Height


cboChart_H.Left = scHistograma.Width + scHistograma.Left - cboChart_H.Width

cboTop.Left = scTop.Width + scTop.Left - cboTop.Width
cboTopList.Left = cboTop.Left - cboTopList.Width
cboTopCount.Left = cboTopList.Left - cboTopCount.Width


cboChart(0).Left = chartC(0).Width + chartC(0).Left - chartC(0).Width
cboChart(1).Left = chartC(1).Width + chartC(1).Left - chartC(1).Width
cboChart(2).Left = chartC(2).Width + chartC(2).Left - chartC(2).Width

btnSave(0).Left = chartC(0).Width + chartC(0).Left - chartC(0).Width
btnSave(1).Left = chartC(1).Width + chartC(1).Left - chartC(1).Width
btnSave(2).Left = chartC(2).Width + chartC(2).Left - chartC(2).Width
 
btnPrint(0).Left = chartC(0).Width + chartC(0).Left - chartC(0).Width
btnPrint(1).Left = chartC(1).Width + chartC(1).Left - chartC(1).Width
btnPrint(2).Left = chartC(2).Width + chartC(2).Left - chartC(2).Width



picLogo.Left = Me.Width - (picLogo.Width + 150)
lblDashboard.Left = (Me.Width - lblDashboard.Width) / 2
cboPalette.Left = (Me.Width - (cboPalette.Width + cboAppearance.Width)) / 2
cboAppearance.Left = cboPalette.Left + cboPalette.Width + 10

gbKPI.Height = Me.Height - (gbKPI.top + 640)

lblBack.Left = 0
lblBack.Width = Me.Width
lblBack.Height = Me.Height

End Sub


Private Sub lswMenu_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

mCategoria = Item.Tag
lblMenu.Caption = Item.Text

gbMenu.Visible = False

Call sbDashboard_Load

End Sub


Private Sub lswTop_DblClick()
On Error GoTo vError

Me.MousePointer = vbHourglass

Call Excel_Exportar_Lsw(lswTop)


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

'Evalua Seguridad de Acceso
If btnAccess.Tag = "0" Then
   MsgBox "No cuentas con el acceso/permiso para esta opción. Contacte a su administrador!", vbExclamation
   Unload Me
   Exit Sub
End If


'Carga el Menu con las categoroias y selecciona el primero
Call sbMenu


With lswMenu.ListItems
  If .Count > 0 Then
    mCategoria = .Item(1).Tag
    lblMenu.Caption = .Item(1).Text
    
    Call sbDashboard_Load
  Else
  
    MsgBox "Su usuario no tiene vinculado ningun KPI para mostrar!", vbExclamation
    Unload Me
  End If
  
End With

End Sub

