VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmSYS_Portal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Portal: Mensajes y Estados"
   ClientHeight    =   10005
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   10980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   8415
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10695
      _Version        =   1441793
      _ExtentX        =   18865
      _ExtentY        =   14843
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
      Item(0).Caption =   "Mensajes"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "scTitulo"
      Item(0).Control(2)=   "Label2(0)"
      Item(0).Control(3)=   "Label2(1)"
      Item(0).Control(4)=   "Label2(2)"
      Item(0).Control(5)=   "txtNotaPie1"
      Item(0).Control(6)=   "gbPersonalizados"
      Item(0).Control(7)=   "btnGuardaMsj"
      Item(0).Control(8)=   "txtNotaPie2"
      Item(1).Caption =   "Portal"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "btnGuardaPortal"
      Item(1).Control(1)=   "Label2(5)"
      Item(1).Control(2)=   "Label2(11)"
      Item(1).Control(3)=   "Label2(12)"
      Item(1).Control(4)=   "Label2(13)"
      Item(1).Control(5)=   "txtLogoAlto"
      Item(1).Control(6)=   "txtLogoAncho"
      Item(1).Control(7)=   "Label2(14)"
      Item(1).Control(8)=   "btnColorPicker"
      Item(1).Control(9)=   "Label2(15)"
      Item(1).Control(10)=   "rbColorSet(0)"
      Item(1).Control(11)=   "rbColorSet(1)"
      Item(1).Control(12)=   "rbColorSet(2)"
      Item(1).Control(13)=   "rbColorSet(3)"
      Item(1).Control(14)=   "txtLogoURL"
      Begin XtremeSuiteControls.RadioButton rbColorSet 
         Height          =   252
         Index           =   0
         Left            =   -63880
         TabIndex        =   36
         Tag             =   "E93D1A"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Blue"
         ForeColor       =   16777215
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
      End
      Begin XtremeSuiteControls.GroupBox gbPersonalizados 
         Height          =   2175
         Left            =   480
         TabIndex        =   9
         Top             =   5400
         Width           =   9855
         _Version        =   1441793
         _ExtentX        =   17383
         _ExtentY        =   3836
         _StockProps     =   79
         Caption         =   "Personalizados:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.TabControl tcCuando 
            Height          =   1095
            Left            =   3240
            TabIndex        =   14
            Top             =   1440
            Width           =   5535
            _Version        =   1441793
            _ExtentX        =   9758
            _ExtentY        =   1926
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   4
            Color           =   32
            PaintManager.Position=   2
            ItemCount       =   5
            SelectedItem    =   3
            Item(0).Caption =   "Manual"
            Item(0).ControlCount=   0
            Item(1).Caption =   "Fecha"
            Item(1).ControlCount=   2
            Item(1).Control(0)=   "Label2(6)"
            Item(1).Control(1)=   "dtpFechaEspecifica"
            Item(2).Caption =   "Dia"
            Item(2).ControlCount=   2
            Item(2).Control(0)=   "Label2(7)"
            Item(2).Control(1)=   "cboDiaAlMes"
            Item(3).Caption =   "Frecuencia"
            Item(3).ControlCount=   4
            Item(3).Control(0)=   "Label2(9)"
            Item(3).Control(1)=   "txtDias"
            Item(3).Control(2)=   "Label2(10)"
            Item(3).Control(3)=   "dtpDiaInicia"
            Item(4).Caption =   "Evento"
            Item(4).ControlCount=   2
            Item(4).Control(0)=   "Label2(8)"
            Item(4).Control(1)=   "cboEvento"
            Begin XtremeSuiteControls.DateTimePicker dtpFechaEspecifica 
               Height          =   312
               Left            =   -68440
               TabIndex        =   16
               Top             =   120
               Visible         =   0   'False
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   550
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
            Begin XtremeSuiteControls.ComboBox cboDiaAlMes 
               Height          =   312
               Left            =   -68800
               TabIndex        =   18
               Top             =   120
               Visible         =   0   'False
               Width           =   1692
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
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Style           =   2
               Appearance      =   5
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cboEvento 
               Height          =   312
               Left            =   -68920
               TabIndex        =   20
               Top             =   120
               Visible         =   0   'False
               Width           =   4452
               _Version        =   1441793
               _ExtentX        =   7858
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
               Appearance      =   5
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.FlatEdit txtDias 
               Height          =   312
               Left            =   720
               TabIndex        =   22
               Top             =   120
               Width           =   1092
               _Version        =   1441793
               _ExtentX        =   1926
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
               Text            =   "30"
               Alignment       =   2
               Appearance      =   5
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.DateTimePicker dtpDiaInicia 
               Height          =   312
               Left            =   3240
               TabIndex        =   40
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   3
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   10
               Left            =   1920
               TabIndex        =   23
               Top             =   120
               Width           =   1572
               _Version        =   1441793
               _ExtentX        =   2773
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "días, Inicial el:"
               BackColor       =   -2147483633
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   9
               Left            =   120
               TabIndex        =   21
               Top             =   120
               Width           =   852
               _Version        =   1441793
               _ExtentX        =   1503
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Cada :  "
               BackColor       =   -2147483633
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   8
               Left            =   -69880
               TabIndex        =   19
               Top             =   120
               Visible         =   0   'False
               Width           =   972
               _Version        =   1441793
               _ExtentX        =   1714
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Evento: "
               BackColor       =   -2147483633
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   7
               Left            =   -69880
               TabIndex        =   17
               Top             =   120
               Visible         =   0   'False
               Width           =   1572
               _Version        =   1441793
               _ExtentX        =   2773
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Día al mes:  "
               BackColor       =   -2147483633
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
               WordWrap        =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label2 
               Height          =   252
               Index           =   6
               Left            =   -69880
               TabIndex        =   15
               Top             =   120
               Visible         =   0   'False
               Width           =   1572
               _Version        =   1441793
               _ExtentX        =   2773
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Fecha Específica:  "
               BackColor       =   -2147483633
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
               WordWrap        =   -1  'True
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtProc 
            Height          =   315
            Left            =   1440
            TabIndex        =   11
            Top             =   1080
            Width           =   8415
            _Version        =   1441793
            _ExtentX        =   14838
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
            Appearance      =   5
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboActivacion 
            Height          =   315
            Left            =   1440
            TabIndex        =   13
            Top             =   1560
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Style           =   2
            Appearance      =   5
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtImagen 
            Height          =   315
            Left            =   1440
            TabIndex        =   41
            Top             =   360
            Width           =   8415
            _Version        =   1441793
            _ExtentX        =   14838
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
            Appearance      =   5
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtImagenW 
            Height          =   315
            Left            =   3960
            TabIndex        =   45
            Top             =   720
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
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
            Text            =   "600"
            Alignment       =   2
            Appearance      =   5
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtImagenH 
            Height          =   315
            Left            =   6480
            TabIndex        =   46
            Top             =   720
            Width           =   1095
            _Version        =   1441793
            _ExtentX        =   1931
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
            Text            =   "300"
            Alignment       =   2
            Appearance      =   5
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   18
            Left            =   5280
            TabIndex        =   44
            Top             =   720
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Imagen Alto:"
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   17
            Left            =   2640
            TabIndex        =   43
            Top             =   720
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Imagen Ancho:"
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   42
            Top             =   360
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Imagen Ruta: "
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   12
            Top             =   1560
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Activación: "
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   10
            Top             =   1080
            Width           =   1335
            _Version        =   1441793
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Procedimiento: "
            BackColor       =   -2147483633
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
            WordWrap        =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   2772
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   10332
         _Version        =   524288
         _ExtentX        =   18225
         _ExtentY        =   4890
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
         MaxCols         =   498
         ScrollBars      =   2
         SpreadDesigner  =   "frmSYS_Portal.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnGuardaMsj 
         Height          =   495
         Left            =   9120
         TabIndex        =   24
         Top             =   7800
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   14
         Picture         =   "frmSYS_Portal.frx":06D7
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnGuardaPortal 
         Height          =   492
         Left            =   -62080
         TabIndex        =   25
         Top             =   3480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextAlignment   =   1
         Appearance      =   14
         Picture         =   "frmSYS_Portal.frx":0E08
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtLogoURL 
         Height          =   312
         Left            =   -68200
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   7332
         _Version        =   1441793
         _ExtentX        =   12933
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
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLogoAlto 
         Height          =   312
         Left            =   -68200
         TabIndex        =   31
         Top             =   2400
         Visible         =   0   'False
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
         Text            =   "310"
         Alignment       =   2
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLogoAncho 
         Height          =   312
         Left            =   -68200
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
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
         Text            =   "200"
         Alignment       =   2
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ColorPicker btnColorPicker 
         Height          =   612
         Left            =   -66400
         TabIndex        =   34
         Top             =   2400
         Visible         =   0   'False
         Width           =   2148
         _Version        =   1441793
         _ExtentX        =   3787
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Color Base"
         BackColor       =   -2147483633
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
         SelectedColor   =   8421504
         DefaultColor    =   8421504
      End
      Begin XtremeSuiteControls.RadioButton rbColorSet 
         Height          =   252
         Index           =   1
         Left            =   -63880
         TabIndex        =   37
         Tag             =   "FF"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Red"
         ForeColor       =   16777215
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
      End
      Begin XtremeSuiteControls.RadioButton rbColorSet 
         Height          =   252
         Index           =   2
         Left            =   -62320
         TabIndex        =   38
         Tag             =   "#33B033"
         Top             =   2400
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Green"
         ForeColor       =   16777215
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
      End
      Begin XtremeSuiteControls.RadioButton rbColorSet 
         Height          =   252
         Index           =   3
         Left            =   -62320
         TabIndex        =   39
         Tag             =   "#A99E9C"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Gray"
         ForeColor       =   16777215
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
      End
      Begin XtremeSuiteControls.FlatEdit txtNotaPie1 
         Height          =   552
         Left            =   2040
         TabIndex        =   7
         Top             =   3960
         Width           =   8412
         _Version        =   1441793
         _ExtentX        =   14838
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
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotaPie2 
         Height          =   552
         Left            =   2040
         TabIndex        =   8
         Top             =   4560
         Width           =   8412
         _Version        =   1441793
         _ExtentX        =   14838
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
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   15
         Left            =   -63880
         TabIndex        =   35
         Top             =   2040
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Recomendados: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   14
         Left            =   -66400
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Para fondos y Titulos: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   13
         Left            =   -69160
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   624
         _Version        =   1441793
         _ExtentX        =   1101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ancho"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   12
         Left            =   -69160
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   624
         _Version        =   1441793
         _ExtentX        =   1101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Alto"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   11
         Left            =   -69640
         TabIndex        =   28
         Top             =   2040
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tamaño: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   5
         Left            =   -69640
         TabIndex        =   26
         Top             =   1320
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "URL Logo: "
         BackColor       =   -2147483633
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   2
         Left            =   1680
         TabIndex        =   6
         Top             =   4680
         Width           =   384
         _Version        =   1441793
         _ExtentX        =   677
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(2)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Top             =   4080
         Width           =   384
         _Version        =   1441793
         _ExtentX        =   677
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "(1)"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   3960
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas al Pie: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Width           =   10332
         _Version        =   1441793
         _ExtentX        =   18224
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "..."
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
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Portal: Configuración para Mensajes, Estados y Notificaciones"
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
      Height          =   372
      Index           =   3
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   7812
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11052
   End
End
Attribute VB_Name = "frmSYS_Portal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, vSMTPs As String


Private Sub btnGuardaMsj_Click()
Dim pA_Dia As String, pA_CadaDia As String, pA_CadaDiaInicio As String
Dim pA_Evento As String, pA_Fecha As String

On Error GoTo vError

If scTitulo.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass

pA_Dia = 1
pA_CadaDia = 7
pA_CadaDiaInicio = "Null"
pA_Evento = "'N/A'"
pA_Fecha = "Null"

Select Case cboActivacion.Text
    Case "Manual"
        
    Case "Fecha"
        pA_Fecha = "'" & Format(dtpFechaEspecifica.Value, "yyyy/mm/dd") & "'"
    
    Case "Día del Mes"
        If cboDiaAlMes.Text = "Ult.Día" Then
            pA_Dia = "32"
        Else
            pA_Dia = cboDiaAlMes.Text
        End If
        
    Case "Cada N días"
        If Not IsNumeric(txtDias.Text) _
                Or txtDias.Text > "32" Or txtDias.Text < "1" Then
            pA_CadaDia = "7"
        End If
    
        pA_CadaDia = txtDias.Text
        pA_CadaDiaInicio = "'" & Format(dtpDiaInicia.Value, "yyyy/mm/dd") & "'"
        
    Case "Evento"
        pA_Evento = "'" & cboEvento.ItemData(cboEvento.ListIndex) & "'"
End Select
'" & Trim(txtImagen.Text) & "
strSQL = "Update SYS_NOTIFICACIONES_CFG Set " _
       & " PIE_01 = '" & txtNotaPie1.Text & "', PIE_02 = '" & txtNotaPie2.Text & "'" _
       & ", P_ACTIVACION = '" & Mid(cboActivacion.Text, 1, 1) & "', P_PROCEDIMIENTO = '" & txtProc.Text & "'" _
       & ", P_ACTIVA_FECHA = " & pA_Fecha & ", P_ACTIVA_DIA = " & pA_Dia & ", P_ACTIVA_EVENTO = " & pA_Evento _
       & ", P_ACTIVA_DIA_FREQ = " & pA_CadaDia & ", P_ACTIVA_DIA_FREQ_INICIA = " & pA_CadaDiaInicio _
       & ", MODIFICA_FECHA = dbo.mygetdate(), MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
       & ", IMAGEN_LOCATE = '" & Trim(txtImagen.Text) & "', IMAGEN_W = '" & Trim(txtImagenW.Text) & "', IMAGEN_H= '" & Trim(txtImagenH.Text) & "'" _
       & " Where Cod_Notifica = '" & scTitulo.Tag & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Portal - Notificación:  " & scTitulo.Tag)

Me.MousePointer = vbDefault

MsgBox "Notificación modificada satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxHexToDec(pHex As String) As Long

'fxHexToDec = CLng("&H" & pHex$)

Dim lngOut As Long
Dim i As Integer
Dim c As Integer

For i = 1 To Len(pHex)
  c = Asc(UCase(Mid(pHex, i, 1)))
  Select Case c
  
  Case 65 To 70
    lngOut = lngOut + ((c - 55) * 16 ^ (Len(pHex) - i))
  
  Case 48 To 57
    lngOut = lngOut + ((c - 48) * 16 ^ (Len(pHex) - i))
  
  Case Else
  
  
  End Select
Next i
fxHexToDec = lngOut


End Function

Private Sub btnGuardaPortal_Click()

Call sbPortal_Save

End Sub


Private Sub sbPortal_Load()

On Error GoTo vError

strSQL = "select LOGO_WEB_SITE, LOGO_ALTO, LOGO_ANCHO, COLOR_SET FROM SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL)

txtLogoURL.Text = rs!Logo_Web_Site
txtLogoAlto.Text = rs!Logo_Alto
txtLogoAncho = rs!Logo_Ancho
btnColorPicker.SelectedColor = fxHexToDec(Trim(rs!Color_Set))

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbPortal_Save()
Dim HexNum As String

On Error GoTo vError

HexNum = Hex(btnColorPicker.SelectedColor)


strSQL = "Update SIF_EMPRESA SET LOGO_WEB_SITE = '" & Trim(txtLogoURL.Text) _
      & "', LOGO_ALTO = " & txtLogoAlto.Text & ", LOGO_ANCHO = " & txtLogoAncho.Text _
      & ", COLOR_SET = '" & HexNum & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Actualiza", "Portal: Preferencias para Logo y Set Color")

MsgBox "Preferencias para Logo y Set Color actualizados!", vbInformation
      
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub cboActivacion_Click()

Select Case cboActivacion.Text
    Case "Manual"
        tcCuando.Item(0).Selected = True
    Case "Fecha"
        tcCuando.Item(1).Selected = True
    Case "Día del Mes"
        tcCuando.Item(2).Selected = True
    Case "Cada N días"
        tcCuando.Item(3).Selected = True
    Case "Evento"
        tcCuando.Item(4).Selected = True
End Select

End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub


Private Sub sbInicializa()
Dim i As Integer

On Error GoTo vError



For i = 1 To 30
  cboDiaAlMes.AddItem CStr(i)
Next
cboDiaAlMes.AddItem "Ult.Día"
cboDiaAlMes.Text = "Ult.Día"

For i = 0 To rbColorSet.Count - 1
    rbColorSet.Item(i).BackColor = fxHexToDec(rbColorSet.Item(i).Tag)
Next i

cboEvento.AddItem "Aprobación de Beneficio y Ayuda Social"
cboEvento.ItemData(cboEvento.ListCount - 1) = "BEN"
cboEvento.AddItem "Aprobación de Estudio de Crédito"
cboEvento.ItemData(cboEvento.ListCount - 1) = "EST"
cboEvento.AddItem "Aprobación de Crédito"
cboEvento.ItemData(cboEvento.ListCount - 1) = "CRD"
cboEvento.AddItem "Aplicación de Abonos"
cboEvento.ItemData(cboEvento.ListCount - 1) = "ABO"
cboEvento.AddItem "Aplicación de Deducciones"
cboEvento.ItemData(cboEvento.ListCount - 1) = "PLA"
cboEvento.AddItem "Emisión de Pago en Bancos"
cboEvento.ItemData(cboEvento.ListCount - 1) = "BAN"
cboEvento.AddItem "Liquidación de la Persona"
cboEvento.ItemData(cboEvento.ListCount - 1) = "LIQ"
cboEvento.AddItem "Retiros de Ahorros"
cboEvento.ItemData(cboEvento.ListCount - 1) = "RET"
cboEvento.AddItem "Registro de Cobro a Fiadores"
cboEvento.ItemData(cboEvento.ListCount - 1) = "FIA"
cboEvento.AddItem "Registro de Cobro Judicial"
cboEvento.ItemData(cboEvento.ListCount - 1) = "CBJ"
cboEvento.AddItem "Registro de Incobrables"
cboEvento.ItemData(cboEvento.ListCount - 1) = "INC"

cboEvento.Text = "Aprobación de Beneficio y Ayuda Social"

cboActivacion.Clear
cboActivacion.AddItem "Manual"
cboActivacion.AddItem "Evento"
cboActivacion.AddItem "Día del Mes"
cboActivacion.AddItem "Fecha"
cboActivacion.AddItem "Cada N días"
cboActivacion.Text = "Manual"

dtpFechaEspecifica.Value = fxFechaServidor
dtpDiaInicia.Value = dtpFechaEspecifica.Value


'Carga los SMTP Autorizados
vSMTPs = ""

strSQL = "exec spSys_SMTPs_AUT_Lista"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  
   If Len(vSMTPs) = 0 Then
    vSMTPs = Trim(rs!COD_SMTP)
  Else
    vSMTPs = vSMTPs & Chr$(9) & Trim(rs!COD_SMTP)
  End If
  
  rs.MoveNext
Loop
rs.Close

If Len(vSMTPs) = 0 Then
 vSMTPs = "CNF01"
End If

'Inicializa Parametros
strSQL = "exec spSys_Notifica_Parametros"
Call ConectionExecute(strSQL)

Call sbMensajes_Load

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbCboGridSMPT(vCol As Integer, vRow As Long, vGrid As Object)

vGrid.Col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox


vGrid.TypeComboBoxList = vSMTPs
vGrid.TypeComboBoxEditable = False

End Sub



Private Sub sbMensajes_Load()

On Error GoTo vError

scTitulo.Caption = "Seleccione un Tipo de Notificacion!"
scTitulo.Tag = ""

txtNotaPie1.Text = ""
txtNotaPie2.Text = ""

txtProc.Text = ""
cboActivacion.Text = "Manual"
Call cboActivacion_Click

vPaso = True
    strSQL = "select * from vSys_Notificaciones_Cfg" _
          & " order by COD_NOTIFICA"
    Call sbCargaGridLocal(vGrid, 6, strSQL)
vPaso = False

vGrid.MaxRows = vGrid.MaxRows + 1

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbMensaje_Load(pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vSys_Notificaciones_Cfg" _
        & " Where COD_NOTIFICA = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

scTitulo.Caption = RTrim(rs!TITULO)
scTitulo.Tag = rs!COD_NOTIFICA

txtNotaPie1.Text = rs!PIE_01
txtNotaPie2.Text = rs!PIE_02

txtProc.Text = rs!P_PROCEDIMIENTO
cboActivacion.Text = rs!ACTIVACION_DESC

Call cboActivacion_Click

Select Case rs!P_Activacion
    Case "M" 'Manual
    Case "E" 'Evento
        Call sbCboAsignaDato(cboEvento, rs!Evento_Desc, True, rs!P_ACTIVA_EVENTO)
    Case "D" 'Dia del Mes
        cboDiaAlMes.Text = rs!DIA_DESC
    Case "C" 'Cada - Frecuencia
        txtDias.Text = rs!P_ACTIVA_DIA_FREQ
        dtpDiaInicia.Value = rs!P_ACTIVA_DIA_FREQ
    Case "F" 'Fecha Especifica
        dtpFechaEspecifica.Value = rs!P_ACTIVA_FECHA
End Select

txtImagen.Text = Trim(rs!Imagen_Locate & "")
txtImagenW.Text = Trim(rs!IMAGEN_W & "")
txtImagenH.Text = Trim(rs!IMAGEN_H & "")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()

On Error GoTo vError

vModulo = 10
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub




Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, pSQL As String)
Dim i As Integer

With pGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    Call OpenRecordSet(rs, pSQL, 0)
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
      'Carga Cbo de SMPTs
      Call sbCboGridSMPT(3, .Row, vGrid)
      
      For i = 1 To MaxCol
        .Col = i
        Select Case i
          Case 1 'Codigo
            .Text = rs!COD_NOTIFICA
            .CellNote = "Registro: " & rs!Registro_Usuario & ", Fecha: " & rs!registro_Fecha & vbCrLf _
                      & "Modificado: " & rs!MODIFICA_USUARIO & ", Fecha: " & rs!MODIFICA_FECHA
          
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 2 'Titulo
            .Text = rs!TITULO
          
          Case 3 'SMPTs
            .Text = rs!SMTP_ID
          
          Case 4 'Tipo
            .Text = rs!Tipo_Desc
          Case 5 'Activo
            .Value = rs!Activa
        End Select
      Next i
      rs.MoveNext
    Loop
    rs.Close

End With

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Trim(vGrid.Text) = "" Then Exit Function

strSQL = "select isnull(count(*),0) as Existe from SYS_NOTIFICACIONES_CFG " _
       & " where COD_NOTIFICA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)
    
If rs!Existe = 0 Then 'Insertar

  strSQL = "insert into SYS_NOTIFICACIONES_CFG(COD_NOTIFICA, TITULO, SMTP_ID, TIPO, ACTIVA, PIE_01, PIE_02" _
         & ", P_ACTIVACION, P_PROCEDIMIENTO, P_ACTIVA_FECHA, P_ACTIVA_DIA, P_ACTIVA_DIA_FREQ" _
         & ", P_ACTIVA_DIA_FREQ_INICIA, P_ACTIVA_EVENTO, REGISTRO_FECHA ,REGISTRO_USUARIO" _
         & ", IMAGEN_LOCATE, IMAGEN_W, IMAGEN_H ) values('" _
         & Trim(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.Col = 3
  strSQL = strSQL & Trim(vGrid.Text) & "','"
  vGrid.Col = 4
  strSQL = strSQL & Mid(vGrid.Text, 1, 3) & "',"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ",'','','M','',Null, 1, 7, Null, 'N/A'" _
         & ",Getdate(),'" & glogon.Usuario _
         & "', '', '600', '300')"

  Call ConectionExecute(strSQL)
 
  vGrid.Col = 1
  Call Bitacora("Registra", "Portal - Notificación:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update SYS_NOTIFICACIONES_CFG set TITULO = '" & Trim(vGrid.Text) & "',SMTP_ID = '"
 vGrid.Col = 3
 strSQL = strSQL & Trim(vGrid.Text) & "', TIPO = '"
 vGrid.Col = 4
 strSQL = strSQL & Mid(vGrid.Text, 1, 3) & "', ACTIVA = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & " where COD_NOTIFICA = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Portal - Notificación:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub rbColorSet_Click(Index As Integer)

btnColorPicker.SelectedColor = fxHexToDec(rbColorSet.Item(Index).Tag)

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Mensajes
        Call sbMensajes_Load
    Case 1 'Portal
        Call sbPortal_Load
End Select
    
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicializa
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Or Col <> 6 Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1

Call sbMensaje_Load(vGrid.Text)

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    
    Call sbCboGridSMPT(3, vGrid.Row, vGrid)

  End If
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    Call sbCboGridSMPT(3, vGrid.Row, vGrid)
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete SYS_NOTIFICACIONES_CFG where COD_NOTIFICA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Portal - Notificación: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
