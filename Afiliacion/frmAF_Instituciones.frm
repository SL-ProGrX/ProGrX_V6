VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_Instituciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entidades Deductoras"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10515
   Icon            =   "frmAF_Instituciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10515
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6972
      Left            =   2040
      TabIndex        =   15
      Top             =   1200
      Width           =   8412
      _Version        =   1441793
      _ExtentX        =   14838
      _ExtentY        =   12298
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
      ItemCount       =   7
      Item(0).Caption =   "General"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "gbMain(0)"
      Item(1).Caption =   "Estado"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gbMain(1)"
      Item(2).Caption =   "Sobrantes"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gbMain(2)"
      Item(3).Caption =   "Codigos"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "gbMain(3)"
      Item(4).Caption =   "Empresas"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "gbMain(4)"
      Item(5).Caption =   "Departamentos"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "gbMain(5)"
      Item(6).Caption =   "Copia"
      Item(6).ControlCount=   1
      Item(6).Control(0)=   "gbMain(6)"
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   6732
         Index           =   4
         Left            =   -70000
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   11874
         _StockProps     =   79
         Caption         =   "Empresas Vinculadas"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswEmpresas 
            Height          =   5892
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   8172
            _Version        =   1441793
            _ExtentX        =   14414
            _ExtentY        =   10393
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
         Begin XtremeSuiteControls.RadioButton optEmpresas 
            Height          =   252
            Index           =   0
            Left            =   2160
            TabIndex        =   18
            Top             =   240
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Deduce a?"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton optEmpresas 
            Height          =   252
            Index           =   1
            Left            =   4080
            TabIndex        =   19
            Top             =   240
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Deducida por?"
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
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   6732
         Index           =   3
         Left            =   -70000
         TabIndex        =   20
         Top             =   360
         Visible         =   0   'False
         Width           =   8292
         _Version        =   1441793
         _ExtentX        =   14626
         _ExtentY        =   11874
         _StockProps     =   79
         Caption         =   "Códigos de Deducción"
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
         BorderStyle     =   2
         Begin XtremeSuiteControls.TabControl tcCodigos 
            Height          =   6372
            Left            =   120
            TabIndex        =   21
            Top             =   0
            Width           =   8172
            _Version        =   1441793
            _ExtentX        =   14414
            _ExtentY        =   11239
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
            Item(0).Caption =   "Códigos"
            Item(0).ControlCount=   2
            Item(0).Control(0)=   "GroupBox4(1)"
            Item(0).Control(1)=   "GroupBox4(0)"
            Item(1).Caption =   "Asignación"
            Item(1).ControlCount=   8
            Item(1).Control(0)=   "lswCodigos"
            Item(1).Control(1)=   "Label9(0)"
            Item(1).Control(2)=   "Label9(1)"
            Item(1).Control(3)=   "lblCodigoDeduccion"
            Item(1).Control(4)=   "lswCreditos"
            Item(1).Control(5)=   "rbCodigos(0)"
            Item(1).Control(6)=   "rbCodigos(1)"
            Item(1).Control(7)=   "rbCodigos(2)"
            Begin XtremeSuiteControls.ListView lswCodigos 
               Height          =   1212
               Left            =   -69880
               TabIndex        =   22
               Top             =   720
               Visible         =   0   'False
               Width           =   7812
               _Version        =   1441793
               _ExtentX        =   13779
               _ExtentY        =   2138
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
               Appearance      =   17
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ListView lswCreditos 
               Height          =   3492
               Left            =   -69880
               TabIndex        =   23
               Top             =   2400
               Visible         =   0   'False
               Width           =   7812
               _Version        =   1441793
               _ExtentX        =   13779
               _ExtentY        =   6159
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
            Begin XtremeSuiteControls.RadioButton rbCodigos 
               Height          =   252
               Index           =   0
               Left            =   -69760
               TabIndex        =   24
               Top             =   6000
               Visible         =   0   'False
               Width           =   1932
               _Version        =   1441793
               _ExtentX        =   3408
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Todos los Códigos"
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
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.GroupBox GroupBox4 
               Height          =   3255
               Index           =   1
               Left            =   360
               TabIndex        =   25
               Top             =   3120
               Width           =   7215
               _Version        =   1441793
               _ExtentX        =   12721
               _ExtentY        =   5736
               _StockProps     =   79
               Caption         =   "Códigos Específicos (Alternos)"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               BorderStyle     =   1
               Begin FPSpreadADO.fpSpread vGrid 
                  Height          =   2772
                  Left            =   240
                  TabIndex        =   26
                  Top             =   360
                  Width           =   6852
                  _Version        =   524288
                  _ExtentX        =   12086
                  _ExtentY        =   4890
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
                  MaxCols         =   3
                  ScrollBars      =   2
                  SpreadDesigner  =   "frmAF_Instituciones.frx":000C
                  VScrollSpecialType=   2
                  AppearanceStyle =   1
               End
            End
            Begin XtremeSuiteControls.GroupBox GroupBox4 
               Height          =   2415
               Index           =   0
               Left            =   240
               TabIndex        =   27
               Top             =   480
               Width           =   7215
               _Version        =   1441793
               _ExtentX        =   12726
               _ExtentY        =   4260
               _StockProps     =   79
               Caption         =   "Códigos Generales (Default)"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               BorderStyle     =   1
               Begin XtremeSuiteControls.FlatEdit txtCodigoDeducAportes 
                  Height          =   312
                  Left            =   3600
                  TabIndex        =   131
                  Top             =   360
                  Width           =   1572
                  _Version        =   1441793
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
                  Alignment       =   2
                  Appearance      =   6
                  UseVisualStyle  =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit txtCodigoDeducCreditos 
                  Height          =   312
                  Left            =   3600
                  TabIndex        =   132
                  Top             =   720
                  Width           =   1572
                  _Version        =   1441793
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
                  Alignment       =   2
                  Appearance      =   6
                  UseVisualStyle  =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit txtCodigoEnvAportes 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   133
                  Top             =   1200
                  Width           =   1575
                  _Version        =   1441793
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
                  Alignment       =   2
                  Appearance      =   6
                  UseVisualStyle  =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit txtCodigoEnvCreditos 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   134
                  Top             =   1560
                  Width           =   1575
                  _Version        =   1441793
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
                  Alignment       =   2
                  Appearance      =   6
                  UseVisualStyle  =   0   'False
               End
               Begin XtremeSuiteControls.FlatEdit txtCodInstDeduc 
                  Height          =   315
                  Left            =   3600
                  TabIndex        =   135
                  Top             =   2040
                  Width           =   1575
                  _Version        =   1441793
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
                  Alignment       =   2
                  Appearance      =   6
                  UseVisualStyle  =   0   'False
               End
               Begin VB.Image Image1 
                  Appearance      =   0  'Flat
                  Height          =   360
                  Index           =   2
                  Left            =   600
                  Picture         =   "frmAF_Instituciones.frx":060A
                  Top             =   1920
                  Width           =   360
               End
               Begin VB.Label Label6 
                  Caption         =   "Opcional: Solo si aplica.?"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   5400
                  TabIndex        =   34
                  Top             =   2040
                  Width           =   1935
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código de Institución"
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
                  Index           =   4
                  Left            =   1080
                  TabIndex        =   33
                  Top             =   2040
                  Width           =   2655
               End
               Begin VB.Image Image1 
                  Appearance      =   0  'Flat
                  Height          =   360
                  Index           =   1
                  Left            =   600
                  Picture         =   "frmAF_Instituciones.frx":078A
                  Top             =   1200
                  Width           =   360
               End
               Begin VB.Image Image1 
                  Appearance      =   0  'Flat
                  Height          =   360
                  Index           =   0
                  Left            =   600
                  Picture         =   "frmAF_Instituciones.frx":0906
                  Top             =   360
                  Width           =   360
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código de Recepción Créditos"
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
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   32
                  Top             =   720
                  Width           =   2655
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código de Recepción Aportes"
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
                  Index           =   0
                  Left            =   1080
                  TabIndex        =   31
                  Top             =   360
                  Width           =   2655
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código de Envio de Aportes"
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
                  Height          =   255
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   30
                  Top             =   1200
                  Width           =   2655
               End
               Begin VB.Label Label3 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Código de Envio de Créditos"
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
                  Height          =   255
                  Index           =   3
                  Left            =   1080
                  TabIndex        =   29
                  Top             =   1560
                  Width           =   2655
               End
               Begin VB.Label Label6 
                  Caption         =   "Indique NO en Mayúsculas para que el sistema no procese este tipo de deducción"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   852
                  Index           =   0
                  Left            =   5400
                  TabIndex        =   28
                  Top             =   360
                  Width           =   1932
               End
            End
            Begin XtremeSuiteControls.RadioButton rbCodigos 
               Height          =   252
               Index           =   1
               Left            =   -67240
               TabIndex        =   35
               Top             =   6000
               Visible         =   0   'False
               Width           =   1932
               _Version        =   1441793
               _ExtentX        =   3408
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Códigos Activos"
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
            End
            Begin XtremeSuiteControls.RadioButton rbCodigos 
               Height          =   252
               Index           =   2
               Left            =   -64720
               TabIndex        =   36
               Top             =   6000
               Visible         =   0   'False
               Width           =   1932
               _Version        =   1441793
               _ExtentX        =   3408
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Códigos Inactivos"
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
            End
            Begin VB.Label Label9 
               Caption         =   "Seleccione el código de deducción:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   0
               Left            =   -69880
               TabIndex        =   39
               Top             =   480
               Visible         =   0   'False
               Width           =   4932
            End
            Begin VB.Label Label9 
               Caption         =   "Indique las Líneas de Créditos y Retenciones vinculadas al código:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   1
               Left            =   -69880
               TabIndex        =   38
               Top             =   2040
               Visible         =   0   'False
               Width           =   5412
            End
            Begin VB.Label lblCodigoDeduccion 
               Alignment       =   1  'Right Justify
               Caption         =   "(Codigo)"
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
               Height          =   372
               Left            =   -65200
               TabIndex        =   37
               Top             =   2040
               Visible         =   0   'False
               Width           =   2892
            End
         End
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   6492
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   7932
         _Version        =   1441793
         _ExtentX        =   13991
         _ExtentY        =   11451
         _StockProps     =   79
         Caption         =   "General"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   2
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1092
            Index           =   0
            Left            =   0
            TabIndex        =   41
            Top             =   1560
            Width           =   8292
            _Version        =   1441793
            _ExtentX        =   14626
            _ExtentY        =   1926
            _StockProps     =   79
            Caption         =   "Formatos de Envío y Recepción de Deducciones"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboPlanillaRecibe 
               Height          =   312
               Left            =   1320
               TabIndex        =   42
               Top             =   360
               Width           =   6492
               _Version        =   1441793
               _ExtentX        =   11456
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
            End
            Begin XtremeSuiteControls.ComboBox cboPlanillaEnvio 
               Height          =   312
               Left            =   1320
               TabIndex        =   43
               Top             =   720
               Width           =   6492
               _Version        =   1441793
               _ExtentX        =   11456
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
            End
            Begin VB.Label Label1 
               Caption         =   "Envío"
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
               Left            =   120
               TabIndex        =   45
               Top             =   744
               Width           =   732
            End
            Begin VB.Label Label1 
               Caption         =   "Recepción"
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
               Left            =   120
               TabIndex        =   44
               Top             =   360
               Width           =   852
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   2172
            Index           =   1
            Left            =   0
            TabIndex        =   46
            Top             =   2760
            Width           =   8292
            _Version        =   1441793
            _ExtentX        =   14626
            _ExtentY        =   3831
            _StockProps     =   79
            Caption         =   "Cuentas Contables"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.FlatEdit txtCtaCreditoDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   105
               Top             =   360
               Width           =   4692
               _Version        =   1441793
               _ExtentX        =   8276
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
            Begin XtremeSuiteControls.FlatEdit txtCtaCredito 
               Height          =   312
               Left            =   1320
               TabIndex        =   106
               Top             =   360
               Width           =   1812
               _Version        =   1441793
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCtaObreroDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   109
               Top             =   720
               Width           =   4692
               _Version        =   1441793
               _ExtentX        =   8276
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
            Begin XtremeSuiteControls.FlatEdit txtCtaObrero 
               Height          =   312
               Left            =   1320
               TabIndex        =   110
               Top             =   720
               Width           =   1812
               _Version        =   1441793
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCtaPatronalDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   111
               Top             =   1080
               Width           =   4692
               _Version        =   1441793
               _ExtentX        =   8276
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
            Begin XtremeSuiteControls.FlatEdit txtCtaPatronal 
               Height          =   312
               Left            =   1320
               TabIndex        =   112
               Top             =   1080
               Width           =   1812
               _Version        =   1441793
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCtaFondosDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   113
               Top             =   1440
               Width           =   4692
               _Version        =   1441793
               _ExtentX        =   8276
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
            Begin XtremeSuiteControls.FlatEdit txtCtaFondos 
               Height          =   315
               Left            =   1320
               TabIndex        =   114
               Top             =   1440
               Width           =   1815
               _Version        =   1441793
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCtaIncoDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   115
               Top             =   1800
               Width           =   4692
               _Version        =   1441793
               _ExtentX        =   8276
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
            Begin XtremeSuiteControls.FlatEdit txtCtaInconsistencias 
               Height          =   315
               Left            =   1320
               TabIndex        =   116
               Top             =   1800
               Width           =   1815
               _Version        =   1441793
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
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label1 
               Caption         =   "Fondos"
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
               Left            =   120
               TabIndex        =   51
               ToolTipText     =   "(Para Planillas Directas en Fondos)"
               Top             =   1440
               Width           =   1212
            End
            Begin VB.Label Label1 
               Caption         =   "Inconsistencias"
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
               Left            =   120
               TabIndex        =   50
               Top             =   1800
               Width           =   1212
            End
            Begin VB.Label Label1 
               Caption         =   "AP.Patronal"
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
               Left            =   120
               TabIndex        =   49
               Top             =   1080
               Width           =   1212
            End
            Begin VB.Label Label1 
               Caption         =   "AP. Obrero"
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
               Index           =   4
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   852
            End
            Begin VB.Label Label1 
               Caption         =   "Créditos"
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
               Left            =   120
               TabIndex        =   47
               Top             =   360
               Width           =   732
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   1092
            Index           =   2
            Left            =   0
            TabIndex        =   52
            Top             =   5280
            Width           =   8292
            _Version        =   1441793
            _ExtentX        =   14626
            _ExtentY        =   1926
            _StockProps     =   79
            Caption         =   "Porcentajes para Registro de Patrimonio con el Patrono?"
            ForeColor       =   4210752
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboTipoAsiento 
               Height          =   312
               Left            =   6120
               TabIndex        =   53
               Top             =   720
               Width           =   1092
               _Version        =   1441793
               _ExtentX        =   1931
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
            End
            Begin XtremeSuiteControls.FlatEdit txtPorcentajeAporte 
               Height          =   312
               Left            =   4920
               TabIndex        =   107
               Top             =   360
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPorcentajeAhorro 
               Height          =   312
               Left            =   4920
               TabIndex        =   108
               Top             =   720
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label4 
               Caption         =   "(%)  Recibido del Ahorro Obrero en la Planilla"
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
               Index           =   4
               Left            =   480
               TabIndex        =   56
               Top             =   720
               Width           =   4452
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "Tipo Asiento"
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
               Left            =   6120
               TabIndex        =   55
               Top             =   360
               Width           =   1092
            End
            Begin VB.Label Label4 
               Caption         =   "(%)  Recibido del Aporte Patronal en la Planilla"
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
               TabIndex        =   54
               Top             =   360
               Width           =   4452
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   792
            Left            =   1320
            TabIndex        =   102
            Top             =   720
            Width           =   6492
            _Version        =   1441793
            _ExtentX        =   11451
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
         Begin XtremeSuiteControls.CheckBox chkDeducionPlanilla 
            Height          =   252
            Left            =   3720
            TabIndex        =   103
            Top             =   0
            Width           =   4092
            _Version        =   1441793
            _ExtentX        =   7218
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Procesa Deducción de Planilla?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkMoraAutomatica 
            Height          =   252
            Left            =   3720
            TabIndex        =   104
            Top             =   360
            Width           =   4092
            _Version        =   1441793
            _ExtentX        =   7218
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Generación de morosidad en Cierre de Mes?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.ComboBox cboTipoPago 
            Height          =   312
            Left            =   1320
            TabIndex        =   154
            Top             =   120
            Width           =   2052
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
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Pago"
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
            Index           =   15
            Left            =   120
            TabIndex        =   155
            Top             =   120
            Width           =   1092
         End
         Begin VB.Label Label1 
            Caption         =   "Dirección"
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
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   660
            Width           =   732
         End
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   6612
         Index           =   2
         Left            =   -69880
         TabIndex        =   58
         Top             =   360
         Visible         =   0   'False
         Width           =   7932
         _Version        =   1441793
         _ExtentX        =   13991
         _ExtentY        =   11663
         _StockProps     =   79
         Caption         =   "Sobrantes y Otros"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   2
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   4332
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   7932
            _Version        =   1441793
            _ExtentX        =   13991
            _ExtentY        =   7641
            _StockProps     =   79
            Caption         =   "Aplicación de Inconsistencias y Sobrantes del Proceso de Recaudación"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboDevOp 
               Height          =   312
               Left            =   1920
               TabIndex        =   60
               Top             =   720
               Width           =   5292
               _Version        =   1441793
               _ExtentX        =   9340
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
            End
            Begin XtremeSuiteControls.ComboBox cboOPSocios 
               Height          =   312
               Left            =   1920
               TabIndex        =   61
               Top             =   2520
               Width           =   5292
               _Version        =   1441793
               _ExtentX        =   9340
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
            End
            Begin XtremeSuiteControls.ComboBox cboOPExSocios 
               Height          =   312
               Left            =   1920
               TabIndex        =   62
               Top             =   3600
               Width           =   5292
               _Version        =   1441793
               _ExtentX        =   9340
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
            End
            Begin XtremeSuiteControls.FlatEdit txtDevPlanDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   117
               Top             =   1080
               Width           =   4092
               _Version        =   1441793
               _ExtentX        =   7218
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
            Begin XtremeSuiteControls.FlatEdit txtDevPlan 
               Height          =   312
               Left            =   1920
               TabIndex        =   118
               Top             =   1080
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtDevPlanPatDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   119
               Top             =   1440
               Width           =   4092
               _Version        =   1441793
               _ExtentX        =   7218
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
            Begin XtremeSuiteControls.FlatEdit txtDevPlanPat 
               Height          =   312
               Left            =   1920
               TabIndex        =   120
               Top             =   1440
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPlanSociosDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   121
               Top             =   2880
               Width           =   4092
               _Version        =   1441793
               _ExtentX        =   7218
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
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPlanSocios 
               Height          =   312
               Left            =   1920
               TabIndex        =   122
               Top             =   2880
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPlanExSociosDesc 
               Height          =   312
               Left            =   3120
               TabIndex        =   123
               Top             =   3960
               Width           =   4092
               _Version        =   1441793
               _ExtentX        =   7218
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
            Begin XtremeSuiteControls.FlatEdit txtPlanExSocios 
               Height          =   312
               Left            =   1920
               TabIndex        =   124
               Top             =   3960
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
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.CheckBox chkDevoluciones 
               Height          =   252
               Left            =   240
               TabIndex        =   127
               Top             =   360
               Width           =   8052
               _Version        =   1441793
               _ExtentX        =   14203
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[Aportes Patrimoniales] Enviar Devoluciones Al Fondo de Ahorros Siguiente"
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
            End
            Begin XtremeSuiteControls.CheckBox chkInconsistencias 
               Height          =   252
               Left            =   360
               TabIndex        =   128
               Top             =   1800
               Width           =   8052
               _Version        =   1441793
               _ExtentX        =   14203
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[Crédito] Aplicar Sobrantes a Deudas"
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
            End
            Begin XtremeSuiteControls.CheckBox chkFNDSocios 
               Height          =   252
               Left            =   720
               TabIndex        =   129
               Top             =   2160
               Width           =   8052
               _Version        =   1441793
               _ExtentX        =   14203
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[Crédito] Trasladar Devoluciones de Asociados/Clientes al Fondo Siguiente:"
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
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkFNDExSocios 
               Height          =   252
               Left            =   720
               TabIndex        =   130
               Top             =   3240
               Width           =   8052
               _Version        =   1441793
               _ExtentX        =   14203
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[Crédito] Trasladar Devoluciones de Ex-Asociados/No Socios al Fondo Siguiente:"
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
               Appearance      =   17
            End
            Begin VB.Label Label2 
               Caption         =   "Plan: Patronal"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   4
               Left            =   720
               TabIndex        =   69
               Top             =   1440
               Width           =   1212
            End
            Begin VB.Label Label2 
               Caption         =   "Operadora"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   3
               Left            =   720
               TabIndex        =   68
               Top             =   720
               Width           =   1092
            End
            Begin VB.Label Label2 
               Caption         =   "Plan: Obrero"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   2
               Left            =   720
               TabIndex        =   67
               Top             =   1080
               Width           =   1092
            End
            Begin VB.Label Label2 
               Caption         =   "Operadora"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   7
               Left            =   960
               TabIndex        =   66
               Top             =   2520
               Width           =   852
            End
            Begin VB.Label Label2 
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
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   6
               Left            =   960
               TabIndex        =   65
               Top             =   2880
               Width           =   852
            End
            Begin VB.Label Label2 
               Caption         =   "Operadora"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   5
               Left            =   960
               TabIndex        =   64
               Top             =   3600
               Width           =   852
            End
            Begin VB.Label Label2 
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
               ForeColor       =   &H00FF0000&
               Height          =   252
               Index           =   1
               Left            =   960
               TabIndex        =   63
               Top             =   3960
               Width           =   852
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox3 
            Height          =   2532
            Index           =   1
            Left            =   120
            TabIndex        =   70
            Top             =   4680
            Width           =   7692
            _Version        =   1441793
            _ExtentX        =   13568
            _ExtentY        =   4466
            _StockProps     =   79
            Caption         =   "Histórico y Cuotas (Transito y Moratorias)"
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
            Appearance      =   16
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboCuotasMora 
               Height          =   315
               Left            =   5520
               TabIndex        =   71
               Top             =   1080
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
            End
            Begin XtremeSuiteControls.ComboBox cboTransitoCompra 
               Height          =   315
               Left            =   5520
               TabIndex        =   72
               Top             =   1440
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
            End
            Begin XtremeSuiteControls.FlatEdit txtHistoricoCuotasEnviadas 
               Height          =   315
               Left            =   5520
               TabIndex        =   125
               Top             =   360
               Width           =   735
               _Version        =   1441793
               _ExtentX        =   1291
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
               Text            =   "6"
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtTransitoPlanillasMes 
               Height          =   315
               Left            =   5520
               TabIndex        =   126
               Top             =   720
               Width           =   735
               _Version        =   1441793
               _ExtentX        =   1291
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
               Text            =   "2"
               Alignment       =   2
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label8 
               Caption         =   "Número de Planillas Recibidas al mes para aplicación"
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
               TabIndex        =   77
               Top             =   720
               Width           =   5295
            End
            Begin VB.Label Label8 
               Caption         =   "Registrar Abonos en Transito conforme a la información"
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
               TabIndex        =   76
               Top             =   1440
               Width           =   5415
            End
            Begin VB.Label Label8 
               Caption         =   "Metodología de Cobro de Cuotas Atrasadas"
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
               Index           =   2
               Left            =   120
               TabIndex        =   75
               Top             =   1080
               Width           =   5295
            End
            Begin VB.Label Label8 
               Caption         =   "Mes(es)"
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
               Left            =   6360
               TabIndex        =   74
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label8 
               Caption         =   "Conservar Historial de detalle de cuotas enviadas al cobro de "
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   73
               Top             =   360
               Width           =   5535
            End
         End
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   7212
         Index           =   5
         Left            =   -69880
         TabIndex        =   78
         Top             =   360
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   12721
         _StockProps     =   79
         Caption         =   "Departamentos y Secciones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswDepartamentos 
            Height          =   2772
            Left            =   120
            TabIndex        =   80
            Top             =   480
            Width           =   8052
            _Version        =   1441793
            _ExtentX        =   14203
            _ExtentY        =   4890
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswSecciones 
            Height          =   2892
            Left            =   120
            TabIndex        =   79
            Top             =   3600
            Width           =   8052
            _Version        =   1441793
            _ExtentX        =   14203
            _ExtentY        =   5101
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnDepartamentos 
            Height          =   252
            Left            =   7680
            TabIndex        =   81
            Top             =   180
            Width           =   492
            _Version        =   1441793
            _ExtentX        =   868
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "..."
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
         End
         Begin VB.Label lblDepartmentoId 
            Caption         =   "[Departamento?]"
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
            Left            =   120
            TabIndex        =   82
            Top             =   3360
            Width           =   8172
         End
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   6732
         Index           =   6
         Left            =   -69880
         TabIndex        =   83
         Top             =   360
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   11874
         _StockProps     =   79
         Caption         =   "Crear nueva Entidad basada en una Copia de la Actual?"
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
         Begin XtremeSuiteControls.PushButton btnCopiar 
            Height          =   492
            Left            =   6120
            TabIndex        =   84
            Top             =   2940
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Copiar Entidad"
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
            Picture         =   "frmAF_Instituciones.frx":0A86
         End
         Begin XtremeSuiteControls.FlatEdit txtCopiaDesc 
            Height          =   312
            Left            =   1920
            TabIndex        =   85
            Top             =   2160
            Width           =   6012
            _Version        =   1441793
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCopiaDescCorta 
            Height          =   312
            Left            =   1920
            TabIndex        =   86
            Top             =   2520
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Desc. Corta"
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
            Index           =   9
            Left            =   120
            TabIndex        =   89
            Top             =   2508
            Width           =   1692
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Height          =   312
            Index           =   11
            Left            =   480
            TabIndex        =   88
            Top             =   2160
            Width           =   1332
         End
         Begin XtremeSuiteControls.Label Label11 
            Height          =   372
            Left            =   1800
            TabIndex        =   87
            Top             =   1560
            Width           =   6012
            _Version        =   1441793
            _ExtentX        =   10604
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Indique los datos de la nueva entidad a crear?"
            ForeColor       =   12582912
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
         End
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   6372
         Index           =   1
         Left            =   -69880
         TabIndex        =   90
         Top             =   360
         Visible         =   0   'False
         Width           =   8172
         _Version        =   1441793
         _ExtentX        =   14414
         _ExtentY        =   11239
         _StockProps     =   79
         Caption         =   "Estado de la Planilla"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   2
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1572
            Index           =   0
            Left            =   120
            TabIndex        =   91
            Top             =   120
            Width           =   7812
            _Version        =   1441793
            _ExtentX        =   13779
            _ExtentY        =   2773
            _StockProps     =   79
            Caption         =   "Proceso de la Planilla"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.CheckBox chkGeneraDeducciones 
               Height          =   252
               Left            =   240
               TabIndex        =   141
               Top             =   360
               Width           =   2172
               _Version        =   1441793
               _ExtentX        =   3831
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Genera Deducciones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkCargaDeducciones 
               Height          =   252
               Left            =   240
               TabIndex        =   142
               Top             =   600
               Width           =   2172
               _Version        =   1441793
               _ExtentX        =   3831
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Carga Deducciones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkDesgloza 
               Height          =   252
               Left            =   240
               TabIndex        =   143
               Top             =   840
               Width           =   2172
               _Version        =   1441793
               _ExtentX        =   3831
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Detalla Deducciones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkAhAplica 
               Height          =   252
               Left            =   2640
               TabIndex        =   144
               Top             =   360
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[PAT] Aplica Ahorros"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkAhInconsistencias 
               Height          =   252
               Left            =   2640
               TabIndex        =   145
               Top             =   600
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[PAT] Rep. Inconsistencias"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkAhDevoluciones 
               Height          =   252
               Left            =   2640
               TabIndex        =   146
               Top             =   840
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[PAT] Rep. Devoluciones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkCRAplica 
               Height          =   252
               Left            =   5280
               TabIndex        =   147
               Top             =   360
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[CRD] Aplicación Abonos"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkCRInconsistencias 
               Height          =   252
               Left            =   5280
               TabIndex        =   148
               Top             =   600
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[CRD] Rep. Inconsistencias"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin XtremeSuiteControls.CheckBox chkCRRecalculaMora 
               Height          =   252
               Left            =   5280
               TabIndex        =   149
               Top             =   840
               Width           =   2532
               _Version        =   1441793
               _ExtentX        =   4466
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "[CRD] Recalculo Mora"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   1572
            Index           =   1
            Left            =   120
            TabIndex        =   92
            Top             =   1800
            Width           =   7692
            _Version        =   1441793
            _ExtentX        =   13568
            _ExtentY        =   2773
            _StockProps     =   79
            Caption         =   "Ajustes de Cortes e Inicio de Operación de la Deductora"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.PushButton cmdCambiaFecha 
               Height          =   492
               Left            =   6120
               TabIndex        =   93
               Top             =   360
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   868
               _StockProps     =   79
               Caption         =   "Cambiar"
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
               Picture         =   "frmAF_Instituciones.frx":1176
               ImageAlignment  =   4
            End
            Begin XtremeSuiteControls.PushButton btnInicializaDeduccion 
               Height          =   492
               Left            =   6120
               TabIndex        =   94
               Top             =   960
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   868
               _StockProps     =   79
               Caption         =   "Inicializa"
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
               Picture         =   "frmAF_Instituciones.frx":1A22
               ImageAlignment  =   4
            End
            Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
               Height          =   315
               Left            =   3960
               TabIndex        =   100
               Top             =   480
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.DateTimePicker dtpInicializa 
               Height          =   312
               Left            =   3960
               TabIndex        =   101
               Top             =   960
               Width           =   1332
               _Version        =   1441793
               _ExtentX        =   2350
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.CheckBox chkCambiaFechaGeneral 
               Height          =   492
               Left            =   720
               TabIndex        =   150
               Top             =   360
               Width           =   3132
               _Version        =   1441793
               _ExtentX        =   5524
               _ExtentY        =   868
               _StockProps     =   79
               Caption         =   "Cambia Fecha General de Intereses para Formalizaciones"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Transparent     =   -1  'True
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Alignment       =   1
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Corte"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   312
               Index           =   6
               Left            =   5400
               TabIndex        =   96
               Top             =   480
               Width           =   732
            End
            Begin VB.Label Label3 
               Caption         =   "Estable Fecha Inicializa para deducciones de esta Institución"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   492
               Index           =   5
               Left            =   720
               TabIndex        =   95
               Top             =   960
               Width           =   3132
            End
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   2052
            Index           =   2
            Left            =   120
            TabIndex        =   97
            Top             =   3840
            Width           =   7692
            _Version        =   1441793
            _ExtentX        =   13568
            _ExtentY        =   3619
            _StockProps     =   79
            Caption         =   "Realiza Comparativos para reportar cambios en planillas?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox cboCompara 
               Height          =   312
               Left            =   4080
               TabIndex        =   136
               Top             =   360
               Width           =   3252
               _Version        =   1441793
               _ExtentX        =   5741
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
            End
            Begin XtremeSuiteControls.CheckBox chkMovimientos 
               Height          =   252
               Index           =   0
               Left            =   1920
               TabIndex        =   137
               Top             =   1440
               Width           =   1692
               _Version        =   1441793
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Inclusiones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkMovimientos 
               Height          =   252
               Index           =   1
               Left            =   1920
               TabIndex        =   138
               Top             =   1800
               Width           =   1692
               _Version        =   1441793
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Exclusiones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkMovimientos 
               Height          =   252
               Index           =   2
               Left            =   4080
               TabIndex        =   139
               Top             =   1440
               Width           =   1692
               _Version        =   1441793
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Modificaciones"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkMovimientos 
               Height          =   252
               Index           =   3
               Left            =   4080
               TabIndex        =   140
               Top             =   1800
               Width           =   1692
               _Version        =   1441793
               _ExtentX        =   2984
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Mantienen"
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
               UseVisualStyle  =   -1  'True
               Appearance      =   17
               Value           =   1
            End
            Begin XtremeSuiteControls.CheckBox chkCompara 
               Height          =   492
               Left            =   960
               TabIndex        =   151
               Top             =   240
               Width           =   2052
               _Version        =   1441793
               _ExtentX        =   3619
               _ExtentY        =   868
               _StockProps     =   79
               Caption         =   "Realiza Comparativos de Planillas"
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Transparent     =   -1  'True
               UseVisualStyle  =   -1  'True
               Appearance      =   17
            End
            Begin VB.Label Label1 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   " Tipos de Movimientos a Reportar en la Planilla de deducción"
               BeginProperty Font 
                  Name            =   "Arial Narrow"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   252
               Index           =   10
               Left            =   960
               TabIndex        =   99
               Top             =   1080
               Width           =   4332
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   "Vrs"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   98
               Top             =   480
               Width           =   612
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   7680
               X2              =   960
               Y1              =   1320
               Y2              =   1320
            End
         End
      End
   End
   Begin XtremeSuiteControls.CheckBox chkActiva 
      Height          =   252
      Left            =   8880
      TabIndex        =   7
      Top             =   840
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activa?"
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   10080
      Top             =   360
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reportes"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "lisInstituciones"
                  Text            =   "Listado de Instituciones"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "lisDepartamentos"
                  Text            =   "Listado de Departamentos"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "lisSecciones"
                  Text            =   "Listado de Secciones x Departamento"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Esquema"
                  Text            =   "Esquema General"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Xsep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "lisIng"
                  Text            =   "Listados de Ingreso vrs Instituciones"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8880
      TabIndex        =   2
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   6012
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   4
      Top             =   480
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescCorta 
      Height          =   312
      Left            =   2760
      TabIndex        =   6
      Top             =   840
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "General"
      ForeColor       =   0
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":203E
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Estado de la Planilla"
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":2817
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Sobrantes y Cuotas"
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":318B
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Códigos de Deducción"
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":3852
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   3720
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Empresas Vinculadas"
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":3F06
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Departamentos y Secciones"
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":46DE
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnAcciones 
      Height          =   492
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Copiar a Nueva Entidad"
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
      Appearance      =   6
      Picture         =   "frmAF_Instituciones.frx":4D94
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ComboBox cboDivisa 
      Height          =   312
      Left            =   6240
      TabIndex        =   153
      Top             =   840
      Width           =   2532
      _Version        =   1441793
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
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Divisa"
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
      Index           =   14
      Left            =   4200
      TabIndex        =   152
      Top             =   840
      Width           =   1692
   End
   Begin VB.Image imgBanner 
      Height          =   9516
      Left            =   0
      Picture         =   "frmAF_Instituciones.frx":555A
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2892
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desc. Corta"
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
      Index           =   13
      Left            =   960
      TabIndex        =   3
      Top             =   828
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Entidad"
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
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1332
   End
End
Attribute VB_Name = "frmAF_Instituciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean, vPaso As Boolean
Dim mFechaServer As Date


Private Sub btnCopiar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

' spAFI_Institucion_Copia(@InstOrigen int, @InstDest int = 0
'            , @Descripcion varchar(100) = '', @DescCorta varchar(15) = '', @Usuario varchar(30) = ''
'            , @Config smallint =1, @DeptSec smallint = 1, @EstadosPersona smallint = 1
'            , @Deducciones smallint = 1, @AccesoPlanilla smallint  = 1, @Empresas smallint = 1)


If Len(Trim(txtCopiaDesc.Text)) <= 5 Or Len(Trim(txtCopiaDescCorta.Text)) <= 1 Then
    MsgBox "Indique una Descripción (Completa y Corta) válida!", vbInformation
    Exit Sub
End If

strSQL = "exec spAFI_Institucion_Copia " & txtCodigo.Text & ",0,'" & txtCopiaDesc.Text _
        & "','" & txtCopiaDescCorta.Text & "','" & glogon.Usuario & "', 1, 1, 1, 1, 1, 1"
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    
    
    Call Bitacora("Aplica", "Copia de Institución: " & txtCodigo.Text & " -> Nueva -> [Inst:" & rs!cod_institucion & "]")
    
    MsgBox "Creación de Entidad vía Copia de la Actual, realizada Satisfactoriamente...", vbInformation
    
    Call sbConsulta(rs!cod_institucion)
        
    rs.Close
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnDepartamentos_Click()
Call sbFormsCall("frmAF_Departamentos", , , , False, Me)
End Sub

Private Sub btnInicializaDeduccion_Click()
Dim strSQL As String, pFrecuencia As String

On Error GoTo vError

'(@Institucion int, @Proceso int, @Usuario varchar(30))

If Mid(cboTipoPago.Text, 1, 1) = "M" Then
    pFrecuencia = ".0"
Else
   If Day(dtpInicializa.Value) > 15 Then
     pFrecuencia = ".2"
   Else
     pFrecuencia = ".1"
   End If
   
End If

strSQL = "exec spPrm_Institucion_Proceso_Inicial " & txtCodigo.Text & "," & Format(dtpInicializa.Value, "YYYYMM") & pFrecuencia _
        & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Inicializa Fecha Corte para Deducciones: " & Format(dtpInicializa.Value, "YYYYMM") & pFrecuencia & " [Inst:" & txtCodigo.Text & "]")

MsgBox "Fecha de Corte Cambiada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboPlanillaRecibe_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaCredito.SetFocus
End Sub



Private Sub cmdCambiaFecha_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "update instituciones set pr_fecha_corte = '" _
       & Format(dtpFechaCorte.Value, "yyyy/mm/dd") _
       & "' where cod_institucion = " & txtCodigo
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Cambia Fecha de Corte Formalizaciones: " & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & " [Inst:" & txtCodigo.Text & "]")

MsgBox "Fecha de Corte Cambiada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo = "" Then txtCodigo = 0

If vScroll Then
    strSQL = "select Top 1 cod_institucion from instituciones"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_institucion > " & txtCodigo & " order by cod_institucion asc"
    Else
       strSQL = strSQL & " where cod_institucion < " & txtCodigo & " order by cod_institucion desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_institucion
      Call sbConsulta(txtCodigo)
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
 vModulo = 1
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 
 vScroll = False
  FlatScrollBar.Value = 0
 vScroll = True
 
 vModulo = 1
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")

cboTipoPago.AddItem "MENSUAL"
cboTipoPago.AddItem "QUINCENAL"
cboTipoPago.Text = "MENSUAL"


With lswCodigos.ColumnHeaders
  .Add , , "Código", 1800, vbCenter
  .Add , , "Descripción", 5200
End With

lswCodigos.HideColumnHeaders = True

With lswCreditos.ColumnHeaders
  .Add , , "Código", 1800, vbCenter
  .Add , , "Descripción", 4200
  .Add , , "Tipo", 1200
End With


With lswDepartamentos.ColumnHeaders
  .Add , , "Código", 1200, vbCenter
  .Add , , "Departamento/Unidad", 5200
End With

With lswSecciones.ColumnHeaders
  .Add , , "Código", 1200, vbCenter
  .Add , , "Sección", 5200
End With

With lswEmpresas.ColumnHeaders
  .Add , , "Código", 1200, vbCenter
  .Add , , "Desc.Corta", 1200, vbCenter
  .Add , , "Empresa/Organización", 5200
End With


 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbCargaCombo(cboRef As Object, vTipo As String)
Dim strSQL As String

Select Case vTipo
  Case "A" 'Tipos de Asientos
        strSQL = "select rtrim(Tipo_Asiento) as 'IdX',rtrim(Tipo_Asiento) as 'ItmX'  from CntX_Tipos_Asientos where cod_contabilidad = " & GLOBALES.gEnlace
        
        
  Case "O" 'Operadoras
        'Carga Variables de los FONDOS
        strSQL = "SELECT COD_OPERADORA as 'IdX',rtrim(DESCRIPCION) as 'ItmX' FROM FND_OPERADORAS"
         
End Select

Call sbCbo_Llena_New(cboRef, strSQL, False, True)

End Sub

Private Sub sbFormatosLlenaCbo(cboX As Object)

cboX.Clear
cboX.AddItem "00 - Microsoft Excel"
cboX.AddItem "01 - [CCSS] Caja Costarricense Seguro Social"
cboX.AddItem "02 - [INTEGRA] Mecanizada Tesoreria Nacional"
cboX.AddItem "03 - [ASECCSS] Asociacion Solidarista Emp CCSS"
cboX.AddItem "04 - [ICE](ACOTEL)Instituto Costarricense Electricidad"
cboX.AddItem "05 - [COPECAJA] CoopeCaja RL"
cboX.AddItem "06 - [ICE] Oficinas Centrales"
cboX.AddItem "07 - [ICE] Proyectos"
cboX.AddItem "08 - [AYA] Acueductos y Alcantarillados"
cboX.AddItem "09 - [SPA] Mecanizada Tesoreria Nacional"

cboX.AddItem "10 - [SYS] Sistema Interno [F01.Indefinidos]"
cboX.AddItem "11 - [SYS] Sistema Interno [F02.Plazo definido]"

cboX.AddItem "12 - [IMAS]Institucto Mixto de Ayuda de Social"
cboX.AddItem "13 - [INA] Instituto Nacional de Apendizaje"
cboX.AddItem "14 - [MSJ] Municipalidad de San José"
cboX.AddItem "15 - [ PJ] Poder Judicial"
cboX.AddItem "16 - [StarH] PriceWaterHouseCoopers"
cboX.AddItem "17 - [UCR] Universidad de Costa Rica"
cboX.AddItem "18 - [CONAVI] Consejo Nacional de Vialidad"

cboX.AddItem "19 - [CGR]Contraloría General de la República"
cboX.AddItem "20 - [CEN-CINAI] Direccional Nacional de CEN - CINAI"
cboX.AddItem "21 - [UNATEPROT] Unión Nacional Técnicos y Profesionales en Tránsito"
cboX.AddItem "22 - [PANI] Patronato Nacional de la Infancia"
cboX.AddItem "23 - [CORREOS] Correos de Costa Rica"
cboX.AddItem "24 - [SERVICOOP] ServiCoop"
cboX.AddItem "25 - [AGH] Holcim"
cboX.AddItem "26 - [JUPEMA] Junta de Pensionados y Jubilaciones"
cboX.AddItem "27 - [RECOPE] Refinadora Costarricense de Petróleo"

cboX.AddItem "28 - [TEK Ex] TEK Experts"
cboX.AddItem "29 - [P&G] Procter & Gamble"

cboX.AddItem "30 - Excel -Formateado-"

cboX.AddItem "31 - Forza Cash Logistics"
cboX.AddItem "32 - DxC Technology Costa Rica"
cboX.AddItem "33 - DxC Technology Centroamerica"

cboX.AddItem "34 - [ASOECorr] Correos de Costa Rica"

cboX.AddItem "35 - [ProGrX] Recursos Humanos"
cboX.AddItem "36 - [INSVA] INS Valores"

cboX.Text = "00 - Microsoft Excel"

End Sub


Private Sub sbLimpiaPantalla()
Dim i As Integer

vCodigo = 0
txtCodigo.Text = ""


txtDescripcion.Text = ""
txtDescCorta.Text = ""
txtDireccion.Text = ""

chkInconsistencias.Value = vbChecked
chkMoraAutomatica.Value = vbUnchecked

chkCompara.Value = vbUnchecked
cboCompara.Text = "Ultima Planilla Enviada al Cobro"

txtCtaCredito = ""
txtCtaCreditoDesc = ""
txtCtaObrero = ""
txtCtaObreroDesc = ""
txtCtaPatronal = ""
txtCtaPatronalDesc = ""
txtCtaFondos.Text = ""
txtCtaFondosDesc.Text = ""
txtCtaInconsistencias = ""
txtCtaIncoDesc = ""

txtPorcentajeAhorro = 0
txtPorcentajeAporte = 0



dtpFechaCorte.Value = mFechaServer
dtpInicializa.Value = dtpFechaCorte.Value

chkGeneraDeducciones.Value = vbUnchecked
chkCargaDeducciones.Value = vbUnchecked
chkDesgloza.Value = vbUnchecked

chkAhAplica.Value = vbUnchecked
chkAhInconsistencias.Value = vbUnchecked
chkAhDevoluciones.Value = vbUnchecked

chkCRAplica.Value = vbUnchecked
chkCRInconsistencias.Value = vbUnchecked
chkCRRecalculaMora.Value = vbUnchecked

chkCambiaFechaGeneral.Value = vbUnchecked

txtCodigoDeducAportes = ""
txtCodigoDeducCreditos = ""
txtCodigoEnvAportes = ""
txtCodigoEnvCreditos = ""

'Tab Planilla proceso 2
chkDevoluciones.Value = vbUnchecked
chkInconsistencias.Value = vbUnchecked
chkFNDExSocios.Value = vbUnchecked
chkFNDSocios.Value = vbUnchecked

txtDevPlan = ""
txtDevPlanDesc = ""
txtDevPlanPat = ""
txtDevPlanPatDesc = ""

txtPlanExSocios = ""
txtPlanExSociosDesc = ""
txtPlanSocios = ""
txtPlanSociosDesc = ""

txtHistoricoCuotasEnviadas.Text = "6"
cboCuotasMora.Text = "Cuota de Mayor Antiguedad"


txtTransitoPlanillasMes.Text = "2"
cboTransitoCompra.Text = "Enviada"

chkMovimientos.Item(0).Value = vbChecked
chkMovimientos.Item(1).Value = vbChecked
chkMovimientos.Item(2).Value = vbChecked
chkMovimientos.Item(3).Value = vbChecked

Call btnAcciones_Click(0)

btnAcciones.Item(3).Enabled = False
btnAcciones.Item(4).Enabled = False
btnAcciones.Item(5).Enabled = False
btnAcciones.Item(6).Enabled = False

End Sub





Private Sub sbCodigos_Carga_Grid()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

tcCodigos.Item(0).Selected = True

vGrid.MaxRows = 0

strSQL = "select COD_DEDUCCION,descripcion,activo" _
       & " from AFI_INSTITUCIONES_CODIGOS" _
       & " WHERE COD_INSTITUCION = " & txtCodigo.Text _
       & " order by COD_DEDUCCION"
vPaso = True
    Call sbCargaGrid(vGrid, 3, strSQL)
vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnAcciones_Click(Index As Integer)
Dim pLeft As Long, pTop As Long, i As Integer
Dim pWidth As Long, pHeight As Long

pLeft = 2040
pTop = 1320
pWidth = 8532
pHeight = 6732

For i = 0 To 6
   If i = Index Then
      tcMain.Item(i).Visible = True
    Else
      tcMain.Item(i).Visible = False
   End If
Next i

'Codigos de Deducciónn
Select Case Index
  Case 3 'Codigos de Deduccion
    Call sbCodigos_Carga_Grid
  
  Case 4 'Empresas Vinculadas
    Call sbEmpresas_Carga
    
  Case 5 'Departamentos y Secciones
    Call sbDepartamentos_Carga
  
  Case 6 'Copia
    txtCopiaDesc.Text = ""
    txtCopiaDescCorta.Text = ""
End Select

tcMain.Item(Index).Selected = True

End Sub

Private Sub sbDepartamentos_Carga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswDepartamentos.ListItems.Clear
lswSecciones.ListItems.Clear
  
lblDepartmentoId.Caption = "[Seleccione un Departamento]"
lblDepartmentoId.Tag = ""

strSQL = "exec spAFI_Institucion_Departamentos " & txtCodigo.Text
Call OpenRecordSet(rs, strSQL)

vPaso = True
Do While Not rs.EOF
 Set itmX = lswDepartamentos.ListItems.Add(, , rs!cod_departamento)
     itmX.SubItems(1) = Trim(rs!Descripcion)
 rs.MoveNext
Loop
rs.Close

vPaso = False

  
Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbSecciones_Carga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswSecciones.ListItems.Clear
  

strSQL = "exec spAFI_Institucion_Secciones " & txtCodigo.Text & ",'" & lblDepartmentoId.Tag & "'"
Call OpenRecordSet(rs, strSQL)

vPaso = True
Do While Not rs.EOF
 Set itmX = lswSecciones.ListItems.Add(, , rs!cod_seccion)
     itmX.SubItems(1) = Trim(rs!Descripcion)
 rs.MoveNext
Loop
rs.Close

vPaso = False

  
Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbEmpresas_Carga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError


Me.MousePointer = vbHourglass

lswEmpresas.ListItems.Clear
  
strSQL = "exec spAFI_Institucion_Vinculadas " & txtCodigo.Text

Select Case True
    Case optEmpresas.Item(0).Value 'Deduce a:
        strSQL = strSQL & ",1"
    Case optEmpresas.Item(1).Value 'Deducida por:
        strSQL = strSQL & ",2"
    Case Else
        strSQL = strSQL & ",1"
End Select

Call OpenRecordSet(rs, strSQL)

vPaso = True
Do While Not rs.EOF
 Set itmX = lswEmpresas.ListItems.Add(, , rs!cod_institucion)
     itmX.SubItems(1) = Trim(rs!desc_corta & "")
     itmX.SubItems(2) = Trim(rs!Descripcion)
     
     If rs!Asignado = 1 Then
       itmX.Checked = True
     End If
     
 rs.MoveNext
Loop
rs.Close

vPaso = False

  
Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 Resume
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCodigos_Carga_Lineas_Asg()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, vEstado As String

On Error GoTo vError

lswCreditos.ListItems.Clear
If lblCodigoDeduccion.Caption = "" Then Exit Sub

Me.MousePointer = vbHourglass
  
Select Case True
    Case rbCodigos.Item(0).Value    'Todos
        vEstado = "Null"
    Case rbCodigos.Item(1).Value    'Activos
        vEstado = "1"
    Case rbCodigos.Item(2).Value    'Inactivos
        vEstado = "0"
End Select
  
strSQL = "exec spAFI_Instituciones_Codigos_Lineas " & txtCodigo.Text & ",'" & lblCodigoDeduccion.Caption & "'," & vEstado
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
 Set itmX = lswCreditos.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = Trim(rs!Descripcion)
     itmX.SubItems(2) = rs!Tipo
     
     If rs!Asignado = 1 Then
       itmX.Checked = True
     End If
     
 rs.MoveNext
Loop
rs.Close

vPaso = False

  
Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCodigos_Carga_Lista()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lblCodigoDeduccion.Caption = ""
lswCreditos.ListItems.Clear
lswCodigos.ListItems.Clear
  
strSQL = "select COD_DEDUCCION,descripcion,activo" _
       & " from AFI_INSTITUCIONES_CODIGOS" _
       & " WHERE COD_INSTITUCION = " & txtCodigo.Text & " and Activo = 1" _
       & " order by COD_DEDUCCION"
Call OpenRecordSet(rs, strSQL)

vPaso = True
Do While Not rs.EOF
 Set itmX = lswCodigos.ListItems.Add(, , rs!Cod_Deduccion)
     itmX.SubItems(1) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close
vPaso = False

  
Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lswCodigos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)

lswCodigos.SortKey = ColumnHeader.Index - 1

If lswCodigos.SortOrder = 0 Then
    lswCodigos.SortOrder = 1
Else
    lswCodigos.SortOrder = 0
End If
lswCodigos.Sorted = True
End Sub


Private Sub lswCodigos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Or lswCodigos.ListItems.Count <= 0 Then Exit Sub

lblCodigoDeduccion.Caption = Item.Text
Call sbCodigos_Carga_Lineas_Asg

End Sub

Private Sub lswCreditos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
lswCreditos.SortKey = ColumnHeader.Index - 1

If lswCreditos.SortOrder = 0 Then
    lswCreditos.SortOrder = 1
Else
    lswCreditos.SortOrder = 0
End If
lswCreditos.Sorted = True
End Sub

Private Sub lswCreditos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String, pMovimiento As String, pDetalle As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

pDetalle = "Inst. Asignación Código: " & lblCodigoDeduccion.Caption & " (Inst:" & txtCodigo.Text & ") Línea Crd:" & Item.Text

If Item.Checked = True Then
    pMovimiento = "Registra"
    strSQL = "insert AFI_INSTITUCION_ASIGNACION(cod_institucion,cod_deduccion,codigo,registro_fecha,registro_usuario)" _
           & " values(" & txtCodigo.Text & ",'" & lblCodigoDeduccion.Caption & "','" & Item.Text _
           & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
    pMovimiento = "Elimina"
    strSQL = "delete AFI_INSTITUCION_ASIGNACION where cod_institucion = " & txtCodigo.Text _
           & " and cod_deduccion = '" & lblCodigoDeduccion.Caption & "' and codigo = '" & Item.Text & "'"
End If


Call ConectionExecute(strSQL)
Call Bitacora(pMovimiento, pDetalle)
  
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswDepartamentos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
lswDepartamentos.SortKey = ColumnHeader.Index - 1

If lswDepartamentos.SortOrder = 0 Then
    lswDepartamentos.SortOrder = 1
Else
    lswDepartamentos.SortOrder = 0
End If
lswDepartamentos.Sorted = True
End Sub

Private Sub lswDepartamentos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

lblDepartmentoId.Caption = Item.SubItems(1)
lblDepartmentoId.Tag = Item.Text

Call sbSecciones_Carga

End Sub

Private Sub lswEmpresas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
lswEmpresas.SortKey = ColumnHeader.Index - 1

If lswEmpresas.SortOrder = 0 Then
    lswEmpresas.SortOrder = 1
Else
    lswEmpresas.SortOrder = 0
End If
lswEmpresas.Sorted = True
End Sub

Private Sub lswEmpresas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub


On Error GoTo vError:


Select Case True
  Case optEmpresas.Item(0).Value 'Deduce a
            If Item.Checked Then
                strSQL = "insert AFI_INSTITUCION_DEDUCTORA(COD_INSTITUCION, COD_DEDUCTORA, REGISTRO_FECHA, REGISTRO_USUARIO)" _
                       & " values(" & Item.Text & "," & txtCodigo.Text & ",dbo.Mygetdate(), '" & glogon.Usuario & "')"
            
                Call Bitacora("Aplica", "Institución : " & Item.Text & " -> Deductora: " & txtCodigo.Text)
            Else
                strSQL = "delete AFI_INSTITUCION_DEDUCTORA where cod_institucion = " & Item.Text & " and cod_deductora = " & txtCodigo.Text
                
                Call Bitacora("Elimina", "Institución : " & Item.Text & " -> Deductora: " & txtCodigo.Text)
                
            End If


  Case optEmpresas.Item(1).Value 'Deducida Por
            If Item.Checked Then
                strSQL = "insert AFI_INSTITUCION_DEDUCTORA(COD_INSTITUCION, COD_DEDUCTORA, REGISTRO_FECHA, REGISTRO_USUARIO)" _
                       & " values(" & txtCodigo.Text & "," & Item.Text & ",dbo.Mygetdate(), '" & glogon.Usuario & "')"
            
                Call Bitacora("Aplica", "Institución : " & txtCodigo.Text & " -> Deductora: " & Item.Text)
            Else
                strSQL = "delete AFI_INSTITUCION_DEDUCTORA where cod_institucion = " & txtCodigo.Text & " and cod_deductora = " & Item.Text
                
                Call Bitacora("Elimina", "Institución : " & txtCodigo.Text & " -> Deductora: " & Item.Text)
                
            End If


End Select

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub lswSecciones_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
lswSecciones.SortKey = ColumnHeader.Index - 1

If lswSecciones.SortOrder = 0 Then
    lswSecciones.SortOrder = 1
Else
    lswSecciones.SortOrder = 0
End If
lswSecciones.Sorted = True
End Sub


Private Sub optEmpresas_Click(Index As Integer)
Call sbEmpresas_Carga
End Sub

Private Sub rbCodigos_Click(Index As Integer)
If vPaso Or lswCodigos.ListItems.Count <= 0 Then Exit Sub
Call sbCodigos_Carga_Lineas_Asg
End Sub

Private Sub tcCodigos_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
  Call sbCodigos_Carga_Lista
End If

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

On Error GoTo vError

TimerX.Interval = 0
TimerX.Enabled = False

mFechaServer = fxFechaServidor

strSQL = "select rtrim(cod_Divisa) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)


Call sbCargaCombo(cboTipoAsiento, "A")

Call sbCargaCombo(cboOPSocios, "O")
Call sbCbo_Copia(cboOPSocios, cboOPExSocios)
Call sbCbo_Copia(cboOPSocios, cboDevOp)


Call sbFormatosLlenaCbo(cboPlanillaEnvio)
Call sbFormatosLlenaCbo(cboPlanillaRecibe)

cboCompara.Clear
cboCompara.AddItem "Ultima Planilla Enviada al Cobro"
cboCompara.AddItem "Ultima Planilla Recibida"


'cboCuotasMora.ItemData(cboCuotasMora.NewIndex) = 1

cboCuotasMora.Clear
cboCuotasMora.AddItem "Cuota de Mayor Antiguedad"
cboCuotasMora.ItemData(cboCuotasMora.ListCount - 1) = CStr(1)
cboCuotasMora.AddItem "Cuota de Mayor Peso"
cboCuotasMora.ItemData(cboCuotasMora.ListCount - 1) = CStr(2)
cboCuotasMora.AddItem "Todas las Cuotas"
cboCuotasMora.ItemData(cboCuotasMora.ListCount - 1) = CStr(3)
cboCuotasMora.AddItem "No Deducir Morosidad"
cboCuotasMora.ItemData(cboCuotasMora.ListCount - 1) = CStr(4)

cboTransitoCompra.Clear
cboTransitoCompra.AddItem "Enviada"
cboTransitoCompra.AddItem "Recibida"


Call sbLimpiaPantalla

Exit Sub

vError:

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_institucion,descripcion from instituciones"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
'        frmContenedor.CD.HelpContext = Me.HelpContextID
'        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Function fxTipoPlanilla(pPlanilla As String) As String
  Select Case Trim(pPlanilla)
    Case "00"
      fxTipoPlanilla = "00 - Microsoft Excel"
    Case "01"
      fxTipoPlanilla = "01 - [CCSS] Caja Costarricense Seguro Social"
    Case "02"
      fxTipoPlanilla = "02 - [INTEGRA] Mecanizada Tesoreria Nacional"
    Case "03"
      fxTipoPlanilla = "03 - [ASECCSS] Asociacion Solidarista Emp CCSS"
    Case "04"
      fxTipoPlanilla = "04 - [ICE](ACOTEL)Instituto Costarricense Electricidad"
    Case "05"
      fxTipoPlanilla = "05 - [COPECAJA] CoopeCaja RL"
    Case "06"
      fxTipoPlanilla = "06 - [ICE] Oficinas Centrales"
    Case "07"
      fxTipoPlanilla = "07 - [ICE] Proyectos"
    Case "08"
      fxTipoPlanilla = "08 - [AYA] Acueductos y Alcantarillados"
      
      
    Case "09"
      fxTipoPlanilla = "09 - [SPA] Mecanizada Tesoreria Nacional"
    Case "10"
      fxTipoPlanilla = "10 - [SIF] Sistema SIF [F01.Indefinidos]"
    Case "11"
      fxTipoPlanilla = "11 - [SIF] Sistema SIF [F02.Plazo definido]"
    Case "12"
      fxTipoPlanilla = "12 - [IMAS]Institucto Mixto de Ayuda de Social"
    Case "13"
      fxTipoPlanilla = "13 - [INA] Instituto Nacional de Apendizaje"
    Case "14"
      fxTipoPlanilla = "14 - [MSJ] Municipalidad de San José"
    Case "15"
      fxTipoPlanilla = "15 - [ PJ] Poder Judicial"
    Case "16"
      fxTipoPlanilla = "16 - [StarH] PriceWaterHouseCoopers"
    Case "17"
      fxTipoPlanilla = "17 - [UCR] Universidad de Costa Rica"
    Case "18"
      fxTipoPlanilla = "18 - [CONAVI] Consejo Nacional de Vialidad"
    Case "19"
      fxTipoPlanilla = "19 - [CGR]Contraloría General de la República"

    Case "20"
      fxTipoPlanilla = "20 - [CEN-CINAI] Direccional Nacional de CEN - CINAI"

    Case "21"
      fxTipoPlanilla = "21 - [UNATEPROT] Unión Nacional Técnicos y Profesionales en Tránsito"
    Case "22"
      fxTipoPlanilla = "22 - [PANI] Patronato Nacional de la Infancia"
    Case "23"
      fxTipoPlanilla = "23 - [CORREOS] Correos de Costa Rica"
      
    Case "24"
      fxTipoPlanilla = "24 - [SERVICOOP] ServiCoop"
    
    Case "25"
      fxTipoPlanilla = "25 - [AGH] Holcim"
      
    Case "26"
      fxTipoPlanilla = "26 - [JUPEMA] Junta de Pensionados y Jubilaciones"
      
    Case "27"
      fxTipoPlanilla = "27 - [RECOPE] Refinadora Costarricense de Petróleo"
      
    Case "28"
      fxTipoPlanilla = "28 - [TEK Ex] TEK Experts"
    
    Case "29"
      fxTipoPlanilla = "29 - [P&G] Procter & Gamble"
      
    Case "30"
      fxTipoPlanilla = "30 - Excel -Formateado-"
      
    Case "31"
      fxTipoPlanilla = "31 - Forza Cash Logistics"
      
    Case "32"
      fxTipoPlanilla = "32 - DxC Technology Costa Rica"
  
    Case "33"
      fxTipoPlanilla = "33 - DxC Technology Centroamerica"
    
    
    Case "34"
      fxTipoPlanilla = "34 - [ASOECorr] Correos de Costa Rica"
  
    Case "35"
      fxTipoPlanilla = "35 - [ProGrX] Recursos Humanos"

    Case "36"
      fxTipoPlanilla = "36 - [INSVA] INS Valores"
  End Select
  
    
End Function

Private Sub sbConsulta(xCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vAFI_Instituciones where cod_institucion = " & xCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_institucion
  txtCodigo = rs!cod_institucion
  
  txtDescripcion = rs!Descripcion & ""
  txtDireccion = rs!direccion & ""
  
  txtDescCorta.Text = Trim(rs!desc_corta & "")
      
  chkActiva.Value = rs!Activa
  chkMoraAutomatica.Value = rs!Mora_Cierres
  chkDeducionPlanilla.Value = rs!DEDUCCION_PLANILLA
  
  
  Call sbCboAsignaDato(cboPlanillaRecibe, fxTipoPlanilla(rs!planilla), True, rs!planilla)
  Call sbCboAsignaDato(cboPlanillaEnvio, fxTipoPlanilla(rs!PLANILLA_ENVIO), True, rs!PLANILLA_ENVIO)

  Call sbCboAsignaDato(cboDivisa, rs!Divisa_Desc, True, rs!COD_DIVISA & "")

  cboTipoPago.Text = rs!Frecuencia_Desc
  
  txtCtaCredito.Text = Trim(rs!Cta_Crd_Mask)
  txtCtaCreditoDesc.Text = Trim(rs!Cta_Crd_Desc)
  
  txtCtaObrero.Text = Trim(rs!CTA_OBR_MASK)
  txtCtaObreroDesc.Text = Trim(rs!CTA_OBR_DESC)
  
  txtCtaPatronal.Text = Trim(rs!CTA_PAT_MASK)
  txtCtaPatronalDesc.Text = Trim(rs!CTA_PAT_DESC)
    
  txtCtaFondos.Text = Trim(rs!Cta_Fnd_Mask)
  txtCtaFondosDesc.Text = Trim(rs!Cta_Fnd_Desc)
    
  
  txtCtaInconsistencias.Text = Trim(rs!Cta_Inc_Mask)
  txtCtaIncoDesc.Text = Trim(rs!Cta_Inc_Desc)
    
  txtPorcentajeAhorro = Format(rs!porc_ahorro, "Standard")
  txtPorcentajeAporte = Format(rs!PORC_APORTE, "Standard")
    
  cboTipoAsiento.Text = Trim(rs!TipoAsiento)
  
  
  
    dtpFechaCorte.Value = rs!pr_fecha_corte
    
    chkGeneraDeducciones.Value = rs!pr_genera
    chkCargaDeducciones.Value = rs!pr_carga
    chkDesgloza.Value = rs!pr_desgloza
    
    chkAhAplica.Value = rs!pr_apAplica
    chkAhInconsistencias.Value = rs!pr_apInco
    chkAhDevoluciones.Value = rs!pr_apDev
    
    chkCRAplica.Value = rs!pr_crAplica
    chkCRInconsistencias.Value = rs!pr_crInco
    chkCRRecalculaMora.Value = rs!pr_crMora
    
    txtCodigoDeducAportes = rs!codigo_aportes & ""
    txtCodigoDeducCreditos = rs!codigo_creditos & ""
    
    txtCodigoEnvAportes = rs!Codigo_Aportes_Env & ""
    txtCodigoEnvCreditos = rs!codigo_creditos_env & ""
    
    txtCodInstDeduc.Text = rs!codigo_inst_deduc & ""
    
    chkCambiaFechaGeneral.Value = rs!IND_CAMBIA_FECPRO
    
    chkCompara.Value = rs!Compara_Indicador
    
    If rs!compara_valor = "R" Then
      cboCompara.Text = "Ultima Planilla Recibida"
    Else
      cboCompara.Text = "Ultima Planilla Enviada al Cobro"
    End If
    
    'Tab Planilla proceso 2
    chkDevoluciones.Value = rs!fnd_ap_aplica
    chkInconsistencias.Value = rs!pr_cr_aplica_incon
    chkFNDExSocios.Value = rs!fnd_cr_exAplica
    chkFNDSocios.Value = rs!fnd_cr_soAplica
    
    
    Call sbCboAsignaDato(cboDevOp, rs!OP_AP_DESC, True, rs!fnd_ap_operadora)
    txtDevPlan.Text = rs!fnd_ap_plan
    txtDevPlanPat.Text = rs!fnd_ap_planp
    
    txtDevPlanDesc.Text = rs!PLAN_AP_OBR_DESC
    txtDevPlanPatDesc.Text = rs!PLAN_AP_PAT_DESC
    
    
    Call sbCboAsignaDato(cboOPSocios, rs!OP_CR_SOC_DESC, True, rs!fnd_cr_SoOperadora)
    txtPlanSocios.Text = rs!fnd_cr_SoPlan
    txtPlanSociosDesc.Text = rs!PLAN_CR_SOC
    
    Call sbCboAsignaDato(cboOPExSocios, rs!OP_CR_ESO_DESC, True, rs!fnd_cr_ExOperadora)
    txtPlanExSocios.Text = rs!fnd_cr_ExPlan
    txtPlanExSociosDesc.Text = rs!PLAN_CR_ESO

    
  
   txtHistoricoCuotasEnviadas.Text = rs!Historico_Cobro_envio
   Select Case rs!Tipo_Cobro_Mora
      Case 1
        cboCuotasMora.Text = "Cuota de Mayor Antiguedad"
      Case 2
        cboCuotasMora.Text = "Cuota de Mayor Peso"
      Case 3
        cboCuotasMora.Text = "Todas las Cuotas"
      Case 4
        cboCuotasMora.Text = "No Deducir Morosidad"
   End Select
   
   'Planilla en Transito
   txtTransitoPlanillasMes.Text = rs!TRANSITO_PLANILLAS_MES
   
   If rs!TRANSITO_COMPARA = "E" Then
      cboTransitoCompra.Text = "Enviada"
   Else
      cboTransitoCompra.Text = "Recibida"
   End If
   
   'Check de Tipos de Movimientos
   chkMovimientos.Item(0).Value = rs!IncInclusiones
   chkMovimientos.Item(1).Value = rs!IncExclusiones
   chkMovimientos.Item(2).Value = rs!IncModificaciones
   chkMovimientos.Item(3).Value = rs!IncMantienen
  
  
  'Habilita los tabs
  Call btnAcciones_Click(0)
  btnAcciones.Item(3).Enabled = True
  btnAcciones.Item(4).Enabled = True
  btnAcciones.Item(5).Enabled = True
  btnAcciones.Item(6).Enabled = True

Else
  
  Call sbLimpiaPantalla
  
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Ingrese el nombre de la institucion no es válido ..."
If txtDescCorta.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Indique la Descripción Corta!..."
If Len(txtDescCorta.Text) > 10 Then vMensaje = vMensaje & vbCrLf & " - La Descripción Corta no puede ser mayor a 10 caracteres..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vEdita Then
    
  strSQL = "update instituciones set descripcion = '" & Trim(txtDescripcion) & "',Desc_Corta = '" & Trim(txtDescCorta.Text) _
        & "',Activa = " & chkActiva.Value & ", Mora_Cierres = " & chkMoraAutomatica.Value & ", DEDUCCION_PLANILLA = " & chkDeducionPlanilla.Value _
         & ",planilla = '" & Mid(cboPlanillaRecibe.Text, 1, 2) & "',planilla_envio = '" & Mid(cboPlanillaEnvio.Text, 1, 2) & "'" _
         & ",cod_divisa = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "', FRECUENCIA = '" & Mid(cboTipoPago.Text, 1, 1) & "'" _
         & ",direccion = '" & Trim(txtDireccion) & "'" _
         & ",cta_credito = '" & fxgCntCuentaFormato(False, txtCtaCredito.Text) & "'" _
         & ",cta_obrero = '" & fxgCntCuentaFormato(False, txtCtaObrero.Text) & "'" _
         & ",cta_patronal = '" & fxgCntCuentaFormato(False, txtCtaPatronal.Text) & "'" _
         & ",cta_fondos = '" & fxgCntCuentaFormato(False, txtCtaFondos.Text) & "'" _
         & ",cta_inconsistencia = '" & fxgCntCuentaFormato(False, txtCtaInconsistencias.Text) & "'" _
         & ",TipoAsiento = '" & Trim(cboTipoAsiento.Text) & "'" _
         & ",codigo_aportes = '" & Trim(txtCodigoDeducAportes) & "'" _
         & ",codigo_creditos = '" & Trim(txtCodigoDeducCreditos) & "'" _
         & ",codigo_aportes_env = '" & Trim(txtCodigoEnvAportes) & "'" _
         & ",codigo_creditos_env = '" & Trim(txtCodigoEnvCreditos) & "',codigo_inst_deduc = '" & txtCodInstDeduc.Text & "'" _
         & ",porc_ahorro = " & CCur(txtPorcentajeAhorro) & ",porc_aporte = " & CCur(txtPorcentajeAporte) _
         & ",IncInclusiones = " & chkMovimientos.Item(0).Value & ",IncExclusiones = " & chkMovimientos.Item(1).Value _
         & ",IncModificaciones = " & chkMovimientos.Item(2).Value & ",IncMantienen = " & chkMovimientos.Item(3).Value _

           
  strSQL = strSQL & ",pr_genera = " & chkGeneraDeducciones.Value _
         & ",pr_carga = " & chkCargaDeducciones.Value & ",pr_desgloza = " & chkDesgloza.Value _
         & ",pr_apAplica = " & chkAhAplica.Value & ",pr_apInco = " & chkAhInconsistencias.Value _
         & ",pr_apDev = " & chkAhDevoluciones.Value & ",pr_crAplica = " & chkCRAplica.Value _
         & ",pr_crInco = " & chkCRInconsistencias.Value & ",pr_crMora = " & chkCRRecalculaMora.Value _
         & ",pr_cr_aplica_incon = " & chkInconsistencias.Value _
         & ",fnd_ap_aplica = " & chkDevoluciones.Value & ",fnd_cr_SOAplica = " & chkFNDSocios.Value _
         & ",fnd_cr_ExAplica = " & chkFNDExSocios.Value
                    
                    
  strSQL = strSQL & ",fnd_ap_plan = '" & txtDevPlan & "'" & ",fnd_ap_planp = '" & txtDevPlanPat & "'" _
         & ",fnd_cr_soPlan = '" & txtPlanSocios & "'" & ",fnd_cr_exPlan = '" & txtPlanExSocios & "'" _
         & ",fnd_ap_Operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) & ",fnd_cr_SoOperadora = " & cboOPSocios.ItemData(cboOPSocios.ListIndex) _
         & ",fnd_cr_exOperadora = " & cboOPExSocios.ItemData(cboOPExSocios.ListIndex) _
         & ",Compara_Indicador = " & chkCompara.Value _
         & ",compara_valor = '" & Mid(cboCompara.Text, 17, 1) & "'" _
         & ",Historico_Cobro_Envio = " & txtHistoricoCuotasEnviadas.Text & ",Tipo_Cobro_Mora = " & cboCuotasMora.ItemData(cboCuotasMora.ListIndex) _
         & ",TRANSITO_PLANILLAS_MES = " & txtTransitoPlanillasMes.Text & ",TRANSITO_COMPARA = '" & Mid(cboTransitoCompra.Text, 1, 1) _
         & "' where cod_institucion = " & vCodigo
  Call ConectionExecute(strSQL)
    
  Call Bitacora("Modifica", "Institución No." & vCodigo)

Else
   
   strSQL = "insert into instituciones(descripcion,desc_Corta,activa, cod_divisa, mora_cierres,DEDUCCION_PLANILLA,direccion,planilla,planilla_envio,cta_credito,cta_obrero" _
          & ",cta_patronal,cta_fondos,cta_inconsistencia,TipoAsiento,porc_ahorro,porc_aporte,pr_fecha_corte" _
          & ",pr_genera,pr_carga,pr_desgloza,pr_apAplica,pr_apDev,pr_apInco,pr_crAplica,pr_crInco" _
          & ",pr_crMora,pr_cr_aplica_incon,fnd_ap_aplica,fnd_ap_operadora,fnd_ap_plan" _
          & ",fnd_ap_planp,fnd_cr_soAplica,fnd_cr_soOperadora,fnd_cr_soPlan,fnd_cr_exAplica,fnd_cr_exOperadora" _
          & ",fnd_cr_exPlan,codigo_aportes,codigo_creditos,codigo_aportes_env,codigo_creditos_env" _
          & ",IND_CAMBIA_FECPRO,compara_indicador,compara_valor,codigo_inst_deduc,Historico_Cobro_Envio,Tipo_Cobro_Mora" _
          & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,TRANSITO_PLANILLAS_MES,TRANSITO_COMPARA, FRECUENCIA)" _
          & " values('" & Trim(txtDescripcion) & "','" & Trim(txtDescCorta.Text) & "'," & chkActiva.Value & ",'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & chkMoraAutomatica.Value _
          & "," & chkDeducionPlanilla.Value & ",'" & Trim(txtDireccion) & "','" _
          & Mid(cboPlanillaRecibe.Text, 1, 2) & "','" & Mid(cboPlanillaEnvio.Text, 1, 2) & "','" & fxgCntCuentaFormato(False, txtCtaCredito.Text) _
          & "','" & fxgCntCuentaFormato(False, txtCtaObrero.Text) & "','" & fxgCntCuentaFormato(False, txtCtaPatronal.Text) _
          & "','" & fxgCntCuentaFormato(False, txtCtaFondos.Text) & "','" & fxgCntCuentaFormato(False, txtCtaInconsistencias.Text) & "','" & Trim(cboTipoAsiento.Text) & "'," _
          & CCur(txtPorcentajeAhorro) & "," & CCur(txtPorcentajeAporte) & ",'" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") _
          & "'," & chkGeneraDeducciones.Value & "," & chkCargaDeducciones.Value & "," & chkDesgloza.Value _
          & "," & chkAhAplica.Value & "," & chkAhDevoluciones.Value & "," & chkAhInconsistencias.Value _
          & "," & chkCRAplica.Value & "," & chkCRInconsistencias.Value & "," & chkCRRecalculaMora.Value _
          & "," & chkInconsistencias.Value & "," & chkDevoluciones.Value & "," & cboDevOp.ItemData(cboDevOp.ListIndex) _
          & ",'" & txtDevPlan & "','" & txtDevPlanPat & "'," & chkFNDSocios.Value & "," & cboOPSocios.ItemData(cboOPSocios.ListIndex) _
          & ",'" & txtPlanSocios & "'," & chkFNDExSocios.Value & "," & cboOPExSocios.ItemData(cboOPExSocios.ListIndex) _
          & ",'" & txtPlanExSocios & "','" & txtCodigoDeducAportes & "','" & txtCodigoDeducCreditos & "','" & txtCodigoEnvAportes & "','" _
          & txtCodigoEnvCreditos & "'," & chkCambiaFechaGeneral.Value & "," & chkCompara.Value & ",'" & Mid(cboCompara.Text, 17, 1) _
          & "','" & txtCodInstDeduc.Text & "'," & txtHistoricoCuotasEnviadas.Text & "," & cboCuotasMora.ItemData(cboCuotasMora.ListIndex) _
          & "," & chkMovimientos.Item(0).Value & "," & chkMovimientos.Item(1).Value & "," & chkMovimientos.Item(2).Value & "," & chkMovimientos.Item(3).Value _
          & "," & txtTransitoPlanillasMes.Text & ",'" & Mid(cboTransitoCompra.Text, 1, 1) & "','" & Mid(cboTipoPago.Text, 1, 1) & "')"
     Call ConectionExecute(strSQL)
    
   'Extraer el Ultimo
   strSQL = "select isnull(max(cod_institucion),0) as Ultimo from instituciones"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo = rs!ultimo
   rs.Close
   vCodigo = txtCodigo
   
   'Inserta Departamentos y Secciones por Omision
   strSQL = "insert Afdepartamentos(cod_institucion,cod_departamento,descripcion) values(" & txtCodigo.Text & ",'','SIN IDENTIFICAR')"
   Call ConectionExecute(strSQL)
   
   
   strSQL = "insert AfSecciones(cod_institucion,cod_departamento,cod_seccion,descripcion) values(" & txtCodigo.Text & ",'','','SIN IDENTIFICAR')"
   Call ConectionExecute(strSQL)
   
   
   Call Bitacora("Registra", "Institución No." & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

'For i = 3 To ssTab.Tabs - 1
'  ssTab.TabEnabled(i) = True
'Next

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete instituciones where cod_institucion = " & vCodigo
'  Call ConectionExecute(strSQL)
  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim strSQL As String, i As Integer

If ButtonMenu.Key = "lisIng" Then

Else
  Me.MousePointer = vbHourglass
    With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Personas"
     
     .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
     Select Case ButtonMenu.Key
        Case "lisInstituciones"
          .Formulas(1) = "Titulo='Listado de Instituciones'"
          .Formulas(2) = "SubTitulo='Para Deducción de Planilla'"
          .ReportFileName = SIFGlobal.fxPathReportes("Instituciones.rpt")
        Case "lisDepartamentos"
           i = MsgBox("Si desea Mostrar los departamentos solo de la institución seleccionada" _
                    & " marque [SI/YES] si quiere todas indique [NO]", vbYesNo)
          .ReportFileName = SIFGlobal.fxPathReportes("Departamentos.rpt")
          
          .Formulas(1) = "Titulo='Listado de Departamentos'"
          .Formulas(2) = "SubTitulo='Para Deducción de Planilla'"
          
          If i = vbYes Then
            .SelectionFormula = "{INSTITUCIONES.COD_INSTITUCION} = " & vCodigo
            .Formulas(2) = "SubTitulo='" & txtDescripcion & "'"
          Else
            .Formulas(2) = "SubTitulo='Todas las Instituciones'"
          End If
        
        Case "lisSecciones"
           i = MsgBox("Si desea Mostrar los departamentos solo de la institución seleccionada" _
                    & " marque [SI/YES] si quiere todas indique [NO]", vbYesNo)
          
          .ReportFileName = SIFGlobal.fxPathReportes("EsquemaIDS.rpt")
          
          .Formulas(1) = "Titulo='Listado de Departamentos'"
          
          If i = vbYes Then
            .SelectionFormula = "{INSTITUCIONES.COD_INSTITUCION} = " & vCodigo
            .Formulas(2) = "SubTitulo='" & txtDescripcion & "'"
          Else
            .Formulas(2) = "SubTitulo='Todas las Instituciones'"
          End If
        
        Case "Esquema"
           i = MsgBox("Si desea Mostrar los departamentos solo de la institución seleccionada" _
                    & " marque [SI/YES] si quiere todas indique [NO]", vbYesNo)
          .ReportFileName = SIFGlobal.fxPathReportes("EsquemaIDS.rpt")
          
          .Formulas(1) = "Titulo='Listado de Departamentos'"
          If i = vbYes Then
            .SelectionFormula = "{INSTITUCIONES.COD_INSTITUCION} = " & vCodigo
            .Formulas(2) = "SubTitulo='" & txtDescripcion & "'"
          Else
            .Formulas(2) = "SubTitulo='Todas las Instituciones'"
          End If
          
      End Select
      
      .PrintReport
    
    End With
  Me.MousePointer = vbDefault
End If 'Buttonmenu



End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_institucion"
  gBusquedas.Orden = "cod_institucion"
  gBusquedas.Consulta = "select cod_institucion,desc_Corta, descripcion from instituciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub txtCtaCredito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaCreditoDesc.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta
  txtCtaCredito = gBusquedas.Resultado
End If
End Sub

Private Sub txtCtaCredito_LostFocus()
If txtCtaCredito <> "" Then
    txtCtaCreditoDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaCredito))
    txtCtaCredito = fxgCntCuentaFormato(True, txtCtaCredito)
End If
End Sub


Private Sub txtCtaCreditoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaObrero.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta("D")
  txtCtaCredito = gBusquedas.Resultado
  txtCtaCredito_LostFocus
End If
End Sub




Private Sub txtCtaFondos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaFondosDesc.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta
  txtCtaFondos = gBusquedas.Resultado
End If
End Sub


Private Sub txtCtaFondos_LostFocus()
txtCtaFondosDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaFondos))
txtCtaFondos = fxgCntCuentaFormato(True, txtCtaFondos)
End Sub

Private Sub txtCtaFondosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaInconsistencias.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta("D")
  txtCtaFondos = gBusquedas.Resultado
  txtCtaFondos_LostFocus
End If
End Sub

Private Sub txtCtaObrero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaObreroDesc.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta
  txtCtaObrero = gBusquedas.Resultado
End If
End Sub

Private Sub txtCtaObrero_LostFocus()
txtCtaObreroDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaObrero))
txtCtaObrero = fxgCntCuentaFormato(True, txtCtaObrero)
End Sub


Private Sub txtCtaObreroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaPatronal.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta("D")
  txtCtaObrero = gBusquedas.Resultado
  txtCtaObrero_LostFocus
End If
End Sub

Private Sub txtCtaPatronal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaPatronalDesc.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta
  txtCtaPatronal = gBusquedas.Resultado
End If
End Sub

Private Sub txtCtaPatronal_LostFocus()
txtCtaPatronalDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaPatronal))
txtCtaPatronal = fxgCntCuentaFormato(True, txtCtaPatronal)
End Sub


Private Sub txtCtaPatronalDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaFondos.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta("D")
  txtCtaPatronal = gBusquedas.Resultado
  txtCtaPatronal_LostFocus
End If
End Sub


Private Sub txtCtaInconsistencias_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaIncoDesc.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta
  txtCtaInconsistencias = gBusquedas.Resultado
End If
End Sub

Private Sub txtCtaInconsistencias_LostFocus()
txtCtaIncoDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaInconsistencias))
txtCtaInconsistencias = fxgCntCuentaFormato(True, txtCtaInconsistencias)
End Sub


Private Sub txtCtaIncoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
'  ssTab.Tab = 1
'  dtpFechaCorte.SetFocus
End If

If KeyCode = vbKeyF4 Then
  Call sbgCntCuentaConsulta("D")
  txtCtaInconsistencias = gBusquedas.Resultado
  txtCtaInconsistencias_LostFocus
End If
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_institucion,descripcion from instituciones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub txtDevPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtDevPlan = gBusquedas.Resultado
   txtDevPlanDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDevPlanDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtDevPlan = gBusquedas.Resultado
   txtDevPlanDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDevPlanPat_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtDevPlanPat = gBusquedas.Resultado
   txtDevPlanPatDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtDevPlanPatDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtDevPlanPat = gBusquedas.Resultado
   txtDevPlanPatDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDireccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPlanillaRecibe.SetFocus
End Sub

Private Sub txtPlanSociosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtPlanSocios = gBusquedas.Resultado
   txtPlanSociosDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub txtPlanSocios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtPlanSocios = gBusquedas.Resultado
   txtPlanSociosDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtPlanExSocios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtPlanExSocios = gBusquedas.Resultado
   txtPlanExSociosDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtPlanExSociosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_plan,Descripcion from fnd_planes"
   gBusquedas.Filtro = " and cod_operadora = " & cboDevOp.ItemData(cboDevOp.ListIndex) _
                     & " and cod_Moneda = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtPlanExSocios = gBusquedas.Resultado
   txtPlanExSociosDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub txtPorcentajeAhorro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoAsiento.SetFocus
End Sub

Private Sub txtPorcentajeAporte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPorcentajeAhorro.SetFocus
End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from AFI_INSTITUCIONES_CODIGOS " _
       & " where COD_INSTITUCION = " & txtCodigo.Text & " and COD_DEDUCCION = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function

  strSQL = "insert into AFI_INSTITUCIONES_CODIGOS(COD_INSTITUCION, COD_DEDUCCION,descripcion,activo,registro_fecha,registro_usuario)" _
         & " values(" & txtCodigo.Text & ",'" & vGrid.Text & "','"
  vGrid.col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 3
  strSQL = strSQL & vGrid.Value & ",Getdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  
  Call Bitacora("Registra", "Cod. Deduc.: " & vGrid.Text & " Inst.: " & txtCodigo.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update AFI_INSTITUCIONES_CODIGOS set Descripcion = '" & vGrid.Text & "',activo= "
 vGrid.col = 3
 strSQL = strSQL & vGrid.Value & " where COD_INSTITUCION = " & txtCodigo.Text & " AND COD_DEDUCCION = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Cod. Deduc.: " & vGrid.Text & " Inst.: " & txtCodigo.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 3) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete AFI_INSTITUCIONES_CODIGOS where COD_DEDUCCION = '" & vGrid.Text _
                & "' and cod_institucion = " & txtCodigo.Text
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Cod. Deduc.: " & vGrid.Text & " Inst.: " & txtCodigo.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

