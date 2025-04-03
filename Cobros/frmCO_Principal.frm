VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.TaskPanel.v24.0.0.ocx"
Begin VB.Form frmCO_Principal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Cobro"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12060
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   7.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   4001
   Icon            =   "frmCO_Principal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   12060
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Height          =   5808
      Left            =   0
      TabIndex        =   93
      Top             =   1440
      Width           =   2760
      _Version        =   1572864
      _ExtentX        =   4868
      _ExtentY        =   10245
      _StockProps     =   64
      VisualTheme     =   17
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin ComctlLib.StatusBar StatusBarX 
      Height          =   252
      Left            =   2760
      TabIndex        =   46
      Top             =   6960
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   9596
            MinWidth        =   9596
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ultima Acción"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   9596
            MinWidth        =   9596
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ultimo Seguimiento"
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5655
      Left            =   2760
      TabIndex        =   3
      Top             =   1440
      Width           =   9255
      _Version        =   1572864
      _ExtentX        =   16325
      _ExtentY        =   9975
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
      ItemCount       =   8
      Item(0).Caption =   "Estado"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gbEstado"
      Item(0).Control(1)=   "gbDeuda"
      Item(1).Caption =   "Historial"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Gestiones"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vgCobro"
      Item(3).Caption =   "Notificaciones"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "gbNotificaciones(0)"
      Item(3).Control(1)=   "gbNotificaciones(1)"
      Item(4).Caption =   "Cobro a Fiadores"
      Item(4).Tooltip =   "Operaciones en Cobro"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "tcCF"
      Item(5).Caption =   "Contacto"
      Item(5).ControlCount=   7
      Item(5).Control(0)=   "txtDirFiadores"
      Item(5).Control(1)=   "txtApartado"
      Item(5).Control(2)=   "Label7(1)"
      Item(5).Control(3)=   "Label12"
      Item(5).Control(4)=   "lswContactos"
      Item(5).Control(5)=   "lswTelefonos"
      Item(5).Control(6)=   "txtEmail"
      Item(6).Caption =   "Mora"
      Item(6).ControlCount=   5
      Item(6).Control(0)=   "lswAbonos"
      Item(6).Control(1)=   "cboTipoCuotas"
      Item(6).Control(2)=   "imgReporteCuotas"
      Item(6).Control(3)=   "Label15"
      Item(6).Control(4)=   "lblCuotas"
      Item(7).Caption =   "Deductora"
      Item(7).ControlCount=   4
      Item(7).Control(0)=   "cboDeductora"
      Item(7).Control(1)=   "Label3"
      Item(7).Control(2)=   "btnDeductora"
      Item(7).Control(3)=   "chkDeducirPlanilla"
      Begin XtremeSuiteControls.ListView lswAbonos 
         Height          =   4212
         Left            =   -69760
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   8772
         _Version        =   1572864
         _ExtentX        =   15473
         _ExtentY        =   7429
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
         View            =   3
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswContactos 
         Height          =   1692
         Left            =   -69760
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   8772
         _Version        =   1572864
         _ExtentX        =   15473
         _ExtentY        =   2984
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswTelefonos 
         Height          =   1692
         Left            =   -69760
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   5172
         _Version        =   1572864
         _ExtentX        =   9123
         _ExtentY        =   2984
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
         View            =   3
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcCF 
         Height          =   6255
         Left            =   -70000
         TabIndex        =   102
         Top             =   360
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   11033
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
         Item(0).Caption =   "Cobro a Fiadores"
         Item(0).ControlCount=   4
         Item(0).Control(0)=   "lswCF"
         Item(0).Control(1)=   "ShortcutCaption1(1)"
         Item(0).Control(2)=   "GroupBox1"
         Item(0).Control(3)=   "btnExport(0)"
         Item(1).Caption =   "Traslado de Deuda"
         Item(1).ControlCount=   5
         Item(1).Control(0)=   "lswOperacionesGeneradas"
         Item(1).Control(1)=   "gbEnCobroNp"
         Item(1).Control(2)=   "fraReversionDeTraspaso"
         Item(1).Control(3)=   "ShortcutCaption1(0)"
         Item(1).Control(4)=   "btnExport(1)"
         Begin XtremeSuiteControls.ListView lswCF 
            Height          =   2055
            Left            =   0
            TabIndex        =   128
            Top             =   720
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   3625
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
         Begin XtremeSuiteControls.ListView lswOperacionesGeneradas 
            Height          =   1935
            Left            =   -70000
            TabIndex        =   103
            Top             =   720
            Visible         =   0   'False
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   3413
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
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   360
            Index           =   0
            Left            =   8760
            TabIndex        =   149
            Top             =   360
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   635
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frmCO_Principal.frx":08CA
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   2415
            Left            =   0
            TabIndex        =   130
            Top             =   2760
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   4260
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
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.PushButton btnCF 
               Height          =   375
               Index           =   0
               Left            =   6960
               TabIndex        =   134
               ToolTipText     =   "Informe de Recaudación"
               Top             =   240
               Width           =   615
               _Version        =   1572864
               _ExtentX        =   1085
               _ExtentY        =   661
               _StockProps     =   79
               UseVisualStyle  =   -1  'True
               Appearance      =   21
               Picture         =   "frmCO_Principal.frx":0A34
            End
            Begin XtremeSuiteControls.FlatEdit txtCF_Operacion 
               Height          =   435
               Left            =   3000
               TabIndex        =   132
               Top             =   240
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3196
               _ExtentY        =   762
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCF_Codigo 
               Height          =   435
               Left            =   4800
               TabIndex        =   133
               Top             =   240
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   762
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.PushButton btnCF 
               Height          =   375
               Index           =   1
               Left            =   7560
               TabIndex        =   135
               ToolTipText     =   "Cancela el Cobro"
               Top             =   240
               Width           =   1455
               _Version        =   1572864
               _ExtentX        =   2566
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Cancela Cobro"
               UseVisualStyle  =   -1  'True
               Appearance      =   21
               Picture         =   "frmCO_Principal.frx":113B
            End
            Begin XtremeSuiteControls.FlatEdit txtCF_Cedula 
               Height          =   315
               Left            =   1200
               TabIndex        =   141
               Top             =   840
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_Inicia 
               Height          =   315
               Left            =   4200
               TabIndex        =   142
               Top             =   2040
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_UltMov 
               Height          =   315
               Left            =   7320
               TabIndex        =   143
               Top             =   2040
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_Nombre 
               Height          =   315
               Left            =   3000
               TabIndex        =   144
               Top             =   840
               Width           =   6135
               _Version        =   1572864
               _ExtentX        =   10821
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_Recaudo 
               Height          =   315
               Left            =   1200
               TabIndex        =   145
               Top             =   1680
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_Devuelto 
               Height          =   315
               Left            =   7320
               TabIndex        =   146
               Top             =   1680
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_Estado 
               Height          =   315
               Left            =   1200
               TabIndex        =   148
               Top             =   2040
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCF_Aplicado 
               Height          =   315
               Left            =   4200
               TabIndex        =   152
               Top             =   1680
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.FlatEdit txtCF_Cuota 
               Height          =   315
               Left            =   1200
               TabIndex        =   154
               Top             =   1320
               Width           =   1815
               _Version        =   1572864
               _ExtentX        =   3201
               _ExtentY        =   556
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   155
               Top             =   1320
               Width           =   975
               _Version        =   1572864
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Cuota"
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   7
               Left            =   3240
               TabIndex        =   153
               Top             =   1680
               Width           =   975
               _Version        =   1572864
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Aplicado"
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   6
               Left            =   240
               TabIndex        =   147
               Top             =   2040
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Estado"
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   5
               Left            =   6360
               TabIndex        =   140
               Top             =   1680
               Width           =   975
               _Version        =   1572864
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Devuelto"
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   4
               Left            =   240
               TabIndex        =   139
               Top             =   1680
               Width           =   975
               _Version        =   1572864
               _ExtentX        =   1720
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Recaudado"
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   3
               Left            =   6360
               TabIndex        =   138
               Top             =   2040
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Ult. Mov."
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   137
               Top             =   2040
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Inicio"
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
            Begin XtremeSuiteControls.Label Label5 
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   136
               Top             =   840
               Width           =   855
               _Version        =   1572864
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Cédula"
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
            Begin XtremeSuiteControls.Label Label4 
               Height          =   375
               Left            =   240
               TabIndex        =   131
               Top             =   240
               Width           =   2655
               _Version        =   1572864
               _ExtentX        =   4683
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Datos de la Operación de Cobro a Fiador"
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               WordWrap        =   -1  'True
            End
         End
         Begin XtremeSuiteControls.GroupBox gbEnCobroNp 
            Height          =   855
            Left            =   -69880
            TabIndex        =   104
            Top             =   2760
            Visible         =   0   'False
            Width           =   9135
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   1503
            _StockProps     =   79
            Caption         =   "[Operación Deudor Original]"
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
            Appearance      =   16
            BorderStyle     =   1
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Operación"
               ForeColor       =   &H00404040&
               Height          =   288
               Index           =   0
               Left            =   240
               TabIndex        =   114
               Top             =   360
               Width           =   1092
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Monto"
               ForeColor       =   &H00404040&
               Height          =   288
               Index           =   1
               Left            =   2400
               TabIndex        =   113
               Top             =   360
               Width           =   732
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Plazo"
               ForeColor       =   &H00404040&
               Height          =   288
               Index           =   2
               Left            =   4440
               TabIndex        =   112
               Top             =   360
               Width           =   612
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Tasa"
               ForeColor       =   &H00404040&
               Height          =   288
               Index           =   3
               Left            =   5640
               TabIndex        =   111
               Top             =   360
               Width           =   732
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Saldo"
               ForeColor       =   &H00404040&
               Height          =   288
               Index           =   4
               Left            =   6960
               TabIndex        =   110
               Top             =   360
               Width           =   492
            End
            Begin VB.Label lblMontoActualDeudor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   288
               Left            =   3120
               TabIndex        =   109
               Top             =   360
               Width           =   1332
            End
            Begin VB.Label lblPlazoActualDeudor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   288
               Left            =   5040
               TabIndex        =   108
               Top             =   360
               Width           =   612
            End
            Begin VB.Label lblInteresActualDeudor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   288
               Left            =   6360
               TabIndex        =   107
               Top             =   360
               Width           =   612
            End
            Begin VB.Label lblSaldoActualDeudor 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   288
               Left            =   7440
               TabIndex        =   106
               Top             =   360
               Width           =   1572
            End
            Begin VB.Label lblOperacionActualDeudor 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   288
               Left            =   1320
               TabIndex        =   105
               Top             =   360
               Width           =   1092
            End
         End
         Begin XtremeSuiteControls.GroupBox fraReversionDeTraspaso 
            Height          =   1215
            Left            =   -69880
            TabIndex        =   115
            Top             =   3840
            Visible         =   0   'False
            Width           =   9015
            _Version        =   1572864
            _ExtentX        =   15896
            _ExtentY        =   2138
            _StockProps     =   79
            Caption         =   "Reversión de Cobro de Fiadores/Co Deudores"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            BorderStyle     =   1
            Begin VB.TextBox txtTRAFD_MONTO 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1680
               MultiLine       =   -1  'True
               TabIndex        =   119
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtTRAFD_Plazo 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4200
               TabIndex        =   118
               Top             =   480
               Width           =   615
            End
            Begin VB.TextBox txtTRAFD_Int 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4200
               TabIndex        =   117
               Top             =   840
               Width           =   615
            End
            Begin VB.TextBox txtTRAFD_Cuota 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1680
               MultiLine       =   -1  'True
               TabIndex        =   116
               Top             =   840
               Width           =   1695
            End
            Begin XtremeSuiteControls.PushButton cmdCancelaReversionTraspaso 
               Height          =   612
               Left            =   6000
               TabIndex        =   120
               Top             =   600
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
               _ExtentY        =   1080
               _StockProps     =   79
               Caption         =   "Cancelar"
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
               Picture         =   "frmCO_Principal.frx":1851
            End
            Begin XtremeSuiteControls.PushButton cmdReversaTraspasoDeudas 
               Height          =   612
               Left            =   7440
               TabIndex        =   121
               Top             =   600
               Width           =   1452
               _Version        =   1572864
               _ExtentX        =   2561
               _ExtentY        =   1080
               _StockProps     =   79
               Caption         =   "Reversar"
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
               Picture         =   "frmCO_Principal.frx":201E
            End
            Begin VB.Label lblTasa 
               Caption         =   "..."
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
               Left            =   4920
               TabIndex        =   126
               Tag             =   "-1000"
               Top             =   840
               Width           =   1452
            End
            Begin VB.Label Label24 
               Caption         =   "Plazo"
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
               Left            =   3600
               TabIndex        =   125
               Top             =   480
               Width           =   492
            End
            Begin VB.Label Label25 
               Caption         =   "Tasa"
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
               Index           =   0
               Left            =   3600
               TabIndex        =   124
               Top             =   840
               Width           =   492
            End
            Begin VB.Label Label25 
               Caption         =   "Nuevo Monto"
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
               Left            =   360
               TabIndex        =   123
               Top             =   480
               Width           =   1092
            End
            Begin VB.Label Label25 
               Caption         =   "Nueva Cuota"
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
               Left            =   360
               TabIndex        =   122
               Top             =   840
               Width           =   1092
            End
         End
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   360
            Index           =   1
            Left            =   -61240
            TabIndex        =   150
            Top             =   360
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   635
            _StockProps     =   79
            Appearance      =   6
            Picture         =   "frmCO_Principal.frx":29AB
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   390
            Index           =   1
            Left            =   0
            TabIndex        =   129
            Top             =   360
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   688
            _StockProps     =   14
            Caption         =   "Operaciones Generadas como Cobro a Fiadores"
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
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   390
            Index           =   0
            Left            =   -70000
            TabIndex        =   127
            Top             =   360
            Visible         =   0   'False
            Width           =   9255
            _Version        =   1572864
            _ExtentX        =   16325
            _ExtentY        =   688
            _StockProps     =   14
            Caption         =   "Operaciones Generadas a Fiadores por Traspaso de Deudas"
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
      Begin XtremeSuiteControls.FlatEdit txtEmail 
         Height          =   312
         Left            =   -69760
         TabIndex        =   47
         Top             =   4200
         Visible         =   0   'False
         Width           =   5172
         _Version        =   1572864
         _ExtentX        =   9123
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
         Text            =   "..."
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbEstado 
         Height          =   4932
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   4332
         _Version        =   1572864
         _ExtentX        =   7641
         _ExtentY        =   8700
         _StockProps     =   79
         Caption         =   "Estado"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   312
            Left            =   2040
            TabIndex        =   61
            Top             =   1080
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuota 
            Height          =   312
            Left            =   2040
            TabIndex        =   62
            Top             =   2160
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAmortizado 
            Height          =   312
            Left            =   2040
            TabIndex        =   63
            Top             =   2760
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInteresPagado 
            Height          =   312
            Left            =   2040
            TabIndex        =   64
            Top             =   3120
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtGarantia 
            Height          =   312
            Left            =   2040
            TabIndex        =   65
            Top             =   3720
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDocumento 
            Height          =   312
            Left            =   2040
            TabIndex        =   66
            Top             =   4080
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEstado 
            Height          =   312
            Left            =   2040
            TabIndex        =   67
            Top             =   360
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEstadoMoroso 
            Height          =   312
            Left            =   2040
            TabIndex        =   68
            Top             =   720
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   312
            Left            =   3120
            TabIndex        =   69
            Top             =   1440
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInteresActual 
            Height          =   312
            Left            =   2400
            TabIndex        =   70
            Top             =   1800
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInteresPorcentaje 
            Height          =   312
            Left            =   3120
            TabIndex        =   71
            Top             =   1800
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtUltimoMovimiento 
            Height          =   312
            Left            =   3000
            TabIndex        =   73
            Top             =   4560
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.FlatEdit txtPrimerDeduccion 
            Height          =   312
            Left            =   2040
            TabIndex        =   72
            Top             =   4560
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
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
         Begin VB.Label Label2 
            Caption         =   "Primer /Ult. Cuota"
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
            Index           =   27
            Left            =   240
            TabIndex        =   28
            Top             =   4560
            Width           =   1452
         End
         Begin VB.Label Label2 
            Caption         =   "Documento"
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
            Index           =   26
            Left            =   240
            TabIndex        =   27
            Top             =   4080
            Width           =   1452
         End
         Begin VB.Label Label2 
            Caption         =   "Garantía"
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
            Index           =   25
            Left            =   240
            TabIndex        =   26
            Top             =   3720
            Width           =   1452
         End
         Begin VB.Label Label2 
            Caption         =   "Cuota"
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
            Index           =   7
            Left            =   240
            TabIndex        =   25
            Top             =   2160
            Width           =   1092
         End
         Begin VB.Label Label2 
            Caption         =   "Tasa % (Original)"
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
            Left            =   240
            TabIndex        =   24
            Top             =   1800
            Width           =   1932
         End
         Begin VB.Label Label2 
            Caption         =   "Plazo"
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
            Left            =   240
            TabIndex        =   23
            Top             =   1440
            Width           =   1812
         End
         Begin VB.Label Label2 
            Caption         =   "Amortizado"
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
            Left            =   240
            TabIndex        =   22
            Top             =   2760
            Width           =   1572
         End
         Begin VB.Label Label2 
            Caption         =   "Interes Pagado"
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
            Left            =   240
            TabIndex        =   21
            Top             =   3120
            Width           =   1452
         End
         Begin VB.Label Label2 
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
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   1080
            Width           =   1092
         End
         Begin VB.Label Label2 
            Caption         =   "Antiguedad"
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
            Index           =   24
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   1332
         End
         Begin VB.Label Label2 
            Caption         =   "Estado"
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
            Index           =   23
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Width           =   1332
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5055
         Left            =   -70000
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   9255
         _Version        =   524288
         _ExtentX        =   16325
         _ExtentY        =   8916
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
         MaxCols         =   496
         SpreadDesigner  =   "frmCO_Principal.frx":2B15
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox gbDeuda 
         Height          =   4932
         Left            =   4800
         TabIndex        =   6
         Top             =   480
         Width           =   4332
         _Version        =   1572864
         _ExtentX        =   7641
         _ExtentY        =   8700
         _StockProps     =   79
         Caption         =   "Deuda"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.DateTimePicker dtpCalculoIntCorte 
            Height          =   315
            Left            =   2280
            TabIndex        =   56
            Top             =   4560
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
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
         Begin XtremeSuiteControls.FlatEdit txtInteresesMoratorios 
            Height          =   315
            Left            =   2280
            TabIndex        =   74
            Top             =   1080
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSaldo 
            Height          =   315
            Left            =   2280
            TabIndex        =   75
            Top             =   360
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInteresesCorrientes 
            Height          =   315
            Left            =   2280
            TabIndex        =   76
            Top             =   720
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizasAtrasadas 
            Height          =   315
            Left            =   2280
            TabIndex        =   77
            Top             =   2160
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAmortizacionAtrasada 
            Height          =   315
            Left            =   2280
            TabIndex        =   78
            Top             =   1440
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCargosRegistrados 
            Height          =   315
            Left            =   2280
            TabIndex        =   79
            Top             =   1800
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCbrIntereses 
            Height          =   315
            Left            =   2280
            TabIndex        =   80
            Top             =   4080
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTotalMora 
            Height          =   315
            Left            =   2280
            TabIndex        =   81
            Top             =   2760
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTotalMoraLegal 
            Height          =   315
            Left            =   2280
            TabIndex        =   82
            Top             =   3120
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCbrDeuda 
            Height          =   315
            Left            =   2280
            TabIndex        =   83
            Top             =   3720
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Image imgCalculoInt 
            Height          =   252
            Index           =   1
            Left            =   3960
            Picture         =   "frmCO_Principal.frx":3335
            Stretch         =   -1  'True
            Top             =   4560
            Width           =   252
         End
         Begin VB.Image imgCalculoInt 
            Height          =   252
            Index           =   0
            Left            =   3600
            Picture         =   "frmCO_Principal.frx":3AE1
            Stretch         =   -1  'True
            Top             =   4560
            Width           =   252
         End
         Begin VB.Label Label2 
            Caption         =   "Intereses a Hoy"
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
            Index           =   18
            Left            =   480
            TabIndex        =   17
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Corte Intereses"
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
            Index           =   21
            Left            =   480
            TabIndex        =   16
            Top             =   4572
            Width           =   1812
         End
         Begin VB.Label Label2 
            Caption         =   "Total Deuda"
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
            Index           =   17
            Left            =   480
            TabIndex        =   15
            Top             =   3720
            Width           =   1932
         End
         Begin VB.Label Label2 
            Caption         =   "Mora Financiera"
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
            TabIndex        =   14
            Top             =   2760
            Width           =   1932
         End
         Begin VB.Label Label2 
            Caption         =   "Mora Legal"
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
            Index           =   20
            Left            =   480
            TabIndex        =   13
            Top             =   3120
            Width           =   2052
         End
         Begin VB.Label Label2 
            Caption         =   "Principal atrasado"
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
            Left            =   480
            TabIndex        =   12
            Top             =   1440
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Cargos registrados"
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
            Left            =   480
            TabIndex        =   11
            Top             =   1800
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Pólizas atrasadas"
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
            Index           =   19
            Left            =   480
            TabIndex        =   10
            Top             =   2160
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Interes Moratorio"
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
            TabIndex        =   9
            Top             =   1080
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Interes Corriente"
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
            Index           =   22
            Left            =   480
            TabIndex        =   8
            Top             =   720
            Width           =   1692
         End
         Begin VB.Label Label2 
            Caption         =   "Saldo"
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
            Left            =   480
            TabIndex        =   7
            Top             =   360
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.GroupBox gbNotificaciones 
         Height          =   2415
         Index           =   0
         Left            =   -69880
         TabIndex        =   29
         Top             =   3120
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   4260
         _StockProps     =   79
         Caption         =   "Notificaciones Realizadas"
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
         Begin XtremeSuiteControls.ListView lswAvisos 
            Height          =   1695
            Left            =   0
            TabIndex        =   51
            Top             =   360
            Width           =   8895
            _Version        =   1572864
            _ExtentX        =   15690
            _ExtentY        =   2990
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
            View            =   3
            Appearance      =   16
            ShowBorder      =   0   'False
         End
      End
      Begin XtremeSuiteControls.GroupBox gbNotificaciones 
         Height          =   3375
         Index           =   1
         Left            =   -69880
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   5953
         _StockProps     =   79
         Caption         =   "Notificaciones Realizadas"
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
         Begin XtremeSuiteControls.ListView lswRepOp 
            Height          =   2055
            Left            =   -120
            TabIndex        =   50
            Top             =   570
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   3625
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
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   16
            ShowBorder      =   0   'False
         End
         Begin VB.Frame fraCbrFia 
            BorderStyle     =   0  'None
            Caption         =   "Consulta Movimientos: Cobro Fiadores"
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
            Height          =   2532
            Left            =   5160
            TabIndex        =   95
            Top             =   120
            Visible         =   0   'False
            Width           =   3495
            Begin XtremeSuiteControls.PushButton PushButton1 
               Height          =   615
               Left            =   960
               TabIndex        =   96
               Top             =   1800
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "Informe"
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
               Picture         =   "frmCO_Principal.frx":445E
            End
            Begin XtremeSuiteControls.ComboBox cboRepCbrFia_Exp 
               Height          =   330
               Left            =   960
               TabIndex        =   97
               Top             =   960
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
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
            Begin XtremeSuiteControls.FlatEdit txtRepCbrFia_Estado 
               Height          =   315
               Left            =   960
               TabIndex        =   100
               Top             =   240
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
               _ExtentY        =   556
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
               Alignment       =   2
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboRepCbrFia_ExpDt 
               Height          =   330
               Left            =   960
               TabIndex        =   101
               Top             =   1320
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
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
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Estado"
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
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   99
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Expediente"
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
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   98
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.Frame fraFechas 
            BorderStyle     =   0  'None
            Caption         =   "Fechas de Corte (Para Listados)"
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
            Height          =   2532
            Left            =   5040
            TabIndex        =   36
            Top             =   120
            Visible         =   0   'False
            Width           =   3495
            Begin XtremeSuiteControls.PushButton cmdAceptarFechas 
               Height          =   615
               Left            =   720
               TabIndex        =   44
               Top             =   1800
               Width           =   2175
               _Version        =   1572864
               _ExtentX        =   3836
               _ExtentY        =   1085
               _StockProps     =   79
               Caption         =   "Informe"
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
               Picture         =   "frmCO_Principal.frx":4C1A
            End
            Begin XtremeSuiteControls.DateTimePicker dtpFechaInicio 
               Height          =   312
               Left            =   720
               TabIndex        =   57
               Top             =   360
               Width           =   1332
               _Version        =   1572864
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
               Enabled         =   0   'False
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   3
            End
            Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
               Height          =   312
               Left            =   720
               TabIndex        =   58
               Top             =   720
               Width           =   1332
               _Version        =   1572864
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
               Enabled         =   0   'False
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   3
            End
            Begin XtremeSuiteControls.ComboBox cboRepX 
               Height          =   312
               Left            =   720
               TabIndex        =   59
               Top             =   1080
               Width           =   2652
               _Version        =   1572864
               _ExtentX        =   4683
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
            Begin XtremeSuiteControls.CheckBox chkRepFechaTodas 
               Height          =   492
               Left            =   2160
               TabIndex        =   60
               Top             =   600
               Width           =   2052
               _Version        =   1572864
               _ExtentX        =   3619
               _ExtentY        =   868
               _StockProps     =   79
               Caption         =   "Todas?"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial Narrow"
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
               Value           =   1
            End
            Begin VB.Label Label22 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Filtro"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   3
               Left            =   120
               TabIndex        =   39
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Corte"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   1
               Left            =   120
               TabIndex        =   38
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label23 
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Inicio"
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   0
               Left            =   120
               TabIndex        =   37
               Top             =   360
               Width           =   615
            End
         End
         Begin XtremeSuiteControls.PushButton btnReporteOperacion 
            Height          =   492
            Left            =   6240
            TabIndex        =   45
            Top             =   2040
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Informe"
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCO_Principal.frx":53D6
         End
         Begin XtremeSuiteControls.DateTimePicker dtpCartaCorte 
            Height          =   312
            Left            =   6840
            TabIndex        =   91
            Top             =   1680
            Width           =   1332
            _Version        =   1572864
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
         Begin XtremeShortcutBar.ShortcutCaption lblRepOp 
            Height          =   315
            Left            =   0
            TabIndex        =   156
            Top             =   240
            Width           =   4815
            _Version        =   1572864
            _ExtentX        =   8493
            _ExtentY        =   556
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
            Alignment       =   1
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Corte"
            ForeColor       =   &H00FFFFFF&
            Height          =   312
            Index           =   2
            Left            =   6240
            TabIndex        =   40
            Top             =   1680
            Width           =   612
         End
      End
      Begin FPSpreadADO.fpSpread vgCobro 
         Height          =   4932
         Left            =   -70000
         TabIndex        =   41
         Top             =   480
         Visible         =   0   'False
         Width           =   9132
         _Version        =   524288
         _ExtentX        =   16108
         _ExtentY        =   8700
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
         MaxCols         =   11
         SpreadDesigner  =   "frmCO_Principal.frx":5B92
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtApartado 
         Height          =   312
         Left            =   -69760
         TabIndex        =   48
         Top             =   4920
         Visible         =   0   'False
         Width           =   5172
         _Version        =   1572864
         _ExtentX        =   9123
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
         Text            =   "..."
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDirFiadores 
         Height          =   3072
         Left            =   -64480
         TabIndex        =   49
         Top             =   2160
         Visible         =   0   'False
         Width           =   3492
         _Version        =   1572864
         _ExtentX        =   6159
         _ExtentY        =   5419
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
         Text            =   "..."
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipoCuotas 
         Height          =   312
         Left            =   -68920
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.ComboBox cboDeductora 
         Height          =   312
         Left            =   -68920
         TabIndex        =   53
         Top             =   1920
         Visible         =   0   'False
         Width           =   6852
         _Version        =   1572864
         _ExtentX        =   12091
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
      Begin XtremeSuiteControls.PushButton btnDeductora 
         Height          =   615
         Left            =   -64120
         TabIndex        =   55
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Cambiar Deductora"
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
         Picture         =   "frmCO_Principal.frx":6821
      End
      Begin XtremeSuiteControls.CheckBox chkDeducirPlanilla 
         Height          =   372
         Left            =   -64120
         TabIndex        =   92
         Top             =   720
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Deducir por Planillas?"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Left            =   -68920
         TabIndex        =   54
         Top             =   1680
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cambio de Deductora a:"
      End
      Begin VB.Label Label12 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -69760
         TabIndex        =   35
         Top             =   3960
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label7 
         Caption         =   "Apartado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69760
         TabIndex        =   34
         Top             =   4680
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "..."
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
         Height          =   252
         Left            =   -69760
         TabIndex        =   33
         Top             =   840
         Visible         =   0   'False
         Width           =   8772
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas"
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
         Left            =   -69760
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Image imgReporteCuotas 
         Height          =   252
         Left            =   -65680
         Picture         =   "frmCO_Principal.frx":6FF9
         Stretch         =   -1  'True
         ToolTipText     =   "Reporte de Movimientos Registrados"
         Top             =   480
         Visible         =   0   'False
         Width           =   252
      End
   End
   Begin MSComctlLib.ImageList imgCobro 
      Left            =   10680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":77A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":E007
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":14869
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":1B0CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":2192D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":2818F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":2E9F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":35253
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   2880
      TabIndex        =   84
      Top             =   120
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3196
      _ExtentY        =   762
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
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   432
      Left            =   4680
      TabIndex        =   85
      Top             =   120
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   762
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   432
      Left            =   6720
      TabIndex        =   86
      Top             =   120
      Width           =   1092
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   762
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2880
      TabIndex        =   87
      Top             =   600
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
      Left            =   2880
      TabIndex        =   88
      Top             =   960
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4680
      TabIndex        =   89
      Top             =   960
      Width           =   6012
      _Version        =   1572864
      _ExtentX        =   10604
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
      Left            =   4680
      TabIndex        =   90
      Top             =   600
      Width           =   6012
      _Version        =   1572864
      _ExtentX        =   10604
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
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   11280
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3BAB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3BD2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3BFC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3C146
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3C2DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3C486
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3C628
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3C7B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3C93E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3CA42
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3CCD1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3CDDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3D059
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3D11F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3D2BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_Principal.frx":3D45B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   375
      Left            =   7920
      TabIndex        =   94
      Top             =   120
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
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
      Picture         =   "frmCO_Principal.frx":3D506
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   151
      Top             =   1320
      Visible         =   0   'False
      Width           =   12135
      _Version        =   1572864
      _ExtentX        =   21405
      _ExtentY        =   238
      _StockProps     =   93
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   11
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Index           =   3
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCO_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Type vTab
 direccion As Integer   '1 y 0
 Fiadores As Integer    'Sirven Para Indicar si el Tab a sido
 Antiguedad As Integer  'Seleccionado por primera vez o no, para no repetir
 OPGeneradas As Integer 'Busquedas sobre una misma operacion
End Type

Dim vTabs As vTab, vOperacion As Boolean 'vOperacion es para almacenar el ultimo
                                         'Numero de Operacion Consultado

Dim mCurIntc As Currency, mCurIntm As Currency, mCurIVA As Currency  'Para Alm. Interes Corriente y Moratorios Totales
Dim mcurCargos As Currency, mcurPoliza As Currency, mCurAmortAtrasada As Currency
Dim mTasaPts   As Currency, mTasaLiq As Integer


Const Id_TaskItem_Expediente = 0
Const Id_TaskItem_Advertencias = 1
Const Id_TaskItem_EstadoCuenta = 2
Const Id_TaskItem_CobroFiadores = 3
Const Id_TaskItem_CobroJudicial = 4
Const Id_TaskItem_Incobrables = 5
Const Id_TaskItem_Informes = 6

Private Sub sbTaskPanel_Load()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    tpMain.VisualTheme = xtpTaskPanelThemeOffice2016
   
    Set Group = tpMain.Groups.Add(0, "Seguimiento")
    Group.ToolTip = "Información Principal del caso de cobro"
    Group.Special = True

    
    Group.Items.Add Id_TaskItem_Expediente, "Expediente", xtpTaskItemTypeLink, 4
    Group.Items.Add Id_TaskItem_Advertencias, "Advertencias", xtpTaskItemTypeLink, 8
    
    Set Group = tpMain.Groups.Add(0, "Acciones")
    Group.ToolTip = "Realizar Ejecutorias de Cobros"
    
    Group.Items.Add Id_TaskItem_CobroFiadores, "Cobro a Fiadores", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_CobroJudicial, "Cobro Judicial", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Incobrables, "Incobrables", xtpTaskItemTypeLink, 9
    
    Set Group = tpMain.Groups.Add(0, "Informes")
    Group.Items.Add Id_TaskItem_Informes, "Informes de Mora", xtpTaskItemTypeLink, 15
    Group.Items.Add Id_TaskItem_EstadoCuenta, "Estados Adjuntos", xtpTaskItemTypeLink, 15
    
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
   

End Sub


Private Sub sbTaskPanel_Accion(ItemId As Integer)
Dim i As Integer, strObservacion As String

On Error GoTo vError

If ItemId <> 6 Then
    If Not IsNumeric(txtOperacion.Text) Then Exit Sub
    GLOBALES.gTag = txtOperacion.Text
End If

Select Case ItemId
    Case Id_TaskItem_Expediente  'Gestiones
        Call vgCobro_ButtonClicked(1, 1, 1)

    Case Id_TaskItem_Advertencias  'Advertencias
        GLOBALES.gTag = txtCedula.Text
        Call sbFormsCall("frmCO_AdvertenciasRegistro", 1, , , False, Me)
 
    Case Id_TaskItem_EstadoCuenta  'Informes
        Call sbAdjuntos
    
    Case Id_TaskItem_CobroFiadores  'Cobro a Fiadores
        '1. Verificar si se realiza el pase a todos o solo a uno
        '2. Aplicar porcentaje a cada uno (Saldo + Intereses en Mora)
        'NO SE PUEDEN HACER MOVIMIENTOS DE TRASPASOS SI EL CREDITO SE ENCUENTRA EN CBR
        
        'Verificar Congelamiento
        
      Select Case UCase(txtProceso.Text)
        Case "TRASPASO DEUDAS"
                tcMain.Item(4).Selected = True
                fraReversionDeTraspaso.Visible = True
                txtTRAFD_MONTO.Text = 0
                txtTRAFD_Int.Text = txtInteresActual.Text
                txtTRAFD_Plazo.Text = 0
                txtTRAFD_Cuota.Text = 0
                For i = 1 To lswOperacionesGeneradas.ListItems.Count
                   lswOperacionesGeneradas.ListItems.Item(i).Checked = False
                Next i
        Case Else
                If fxgCongelamiento(txtCedula, "per_traspaso_deudas") Then
                  MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
                  Exit Sub
                End If
                
                
                 If txtProceso.Text = "COBRO JUDICIAL" Then
                     MsgBox "No se puede realizar traspaso de deudas porque es ya se encuentra en Cobro Judicial", vbInformation
                     Exit Sub
                 End If
                 
                 'Actualiza Contactos
                 If lswContactos.ListItems.Count = 0 Then
                    Call sbContactos
                 End If
                 
                 If lswContactos.ListItems.Count <= 1 Then
                     MsgBox "No Existen Fiadores/Co Deudores Registrados para esta Operación. Verifique...!", vbExclamation
                    Exit Sub
                 End If
                
                Call sbFormsCall("frmCO_TrasladoDeuda", 1, , , False, Me)
      End Select
        
        
    
    
    
    Case Id_TaskItem_CobroJudicial  'Cobro Judicial
      
      Select Case txtProceso.Text
        Case "NORMAL"
                If fxValidaPasoCobroJudicial Then
                  'Verificar Congelamiento
                  If fxgCongelamiento(txtCedula, "per_cobro_Judicial") Then
                    MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
                    Exit Sub
                  End If
                 
                 i = MsgBox("Esta seguro que desea enviar a cobro judicial esta Operación", vbYesNo)
                 If i = vbYes Then Call sbCobroJudicial
                
                Else 'Validacion
                 MsgBox "No se puede ejecutar el cobro judicial verifique la información", vbCritical
                End If 'Validacion
        
        Case "COBRO JUDICIAL"
            GLOBALES.gTag = txtOperacion.Text
            Call sbFormsCall("frmCO_ReversionCobroJudicial", 1, , , False, Me)
            Call sbConsulta
        End Select
      
 
    Case Id_TaskItem_Incobrables  'Incobrables
        Call sbFormsCall("frmCO_Incobrables", 1, , , False, Me)
    
    Case Id_TaskItem_Informes  'Informes de Morosidad
        Call sbFormsCall("frmCO_ReportesTransito", , , , False, Me)
      
    
End Select

Call sbConsulta

Exit Sub

vError:
    

End Sub


Private Sub sbCobro_Fiador_Informe()
Dim vSubTitulo As String

On Error GoTo vError

Me.MousePointer = vbHourglass


vSubTitulo = "Deudor: [" & txtCodigo.Text & " - " & txtOperacion.Text & "] " & "Cédula: " & txtCedula.Text & " Nombre: " & txtNombre.Text

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro"

    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "fxSubTitulo='" & Mid(vSubTitulo, 1, 250) & "'"
    .Formulas(2) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "fxUsuario='" & glogon.Usuario & "'"
             
             
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Cobro_Fiador_Boleta.rpt")
    .StoredProcParam(0) = txtCF_Operacion.Text
        
    .SubreportToChange = "sbMovimientos"
    .StoredProcParam(0) = txtCF_Operacion.Text
    
    .Action = 1
End With



Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCobro_Fiador_Cancela()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta seguro que desea Cancelar el Cobro a Fiador de esta Operación", vbYesNo)
If vbNo Then
    Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spCbr_Cobro_Fiadores_Cancela " & txtCF_Operacion.Text & ", '" & glogon.Usuario & "', ''"
Call OpenRecordSet(rs, strSQL)

If glogon.error Then
   Me.MousePointer = vbDefault
   Exit Sub
End If

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    Call Bitacora("Aplica", "Cancelación de Cobro a Fiador de la Operación: " & txtCF_Operacion.Text)
    MsgBox "Cancelación de Cobro a Fiador de la Operación: " & txtCF_Operacion.Text & ", aplicada!", vbInformation
    
    Call sbCobro_Fiadores_List
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCF_Click(Index As Integer)

If txtCF_Operacion.Text = "0" _
    Or txtCF_Operacion.Text = "" Or Not IsNumeric(txtCF_Operacion.Text) Then
    MsgBox "Consulte una Operación de Cobro a Fiador!", vbExclamation
    Exit Sub
End If

Select Case Index
    Case 0 'Informe
        Call sbCobro_Fiador_Informe
    Case 1 'Cancela
        Call sbCobro_Fiador_Cancela
End Select

End Sub

Private Sub btnDeductora_Click()
Dim strSQL As String, vDeductora As Long

If vPaso Then Exit Sub

On Error GoTo vError


If Not IsNumeric(txtOperacion) Then Exit Sub

vDeductora = cboDeductora.ItemData(cboDeductora.ListIndex)

If vDeductora <> CLng(cboDeductora.Tag) Then
 strSQL = "update reg_creditos set COD_DEDUCTORA = " & vDeductora _
        & " where id_solicitud = " & txtOperacion
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Registra", "Cambia Deductora de la OP: " & txtOperacion & ", de: " & cboDeductora.Tag & " a:" & vDeductora)
 Call sbBitacoraCredito("26", "Deductora: " & cboDeductora.Tag & " a:" & vDeductora, "C", txtOperacion.Text, txtCodigo.Text)
 
 cboDeductora.Tag = CStr(vDeductora)
 
 MsgBox "Cambio de Deductora realizado satisfactoriamente!", vbInformation

Else
 MsgBox "No ha indicado un cambio en la deductora actual?", vbExclamation

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExport_Click(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Select Case Index
  Case 0 'Cobro a Fiadores
        Call Excel_Exportar_Lsw(lswCF, ProgressBarX)
  Case 1 'Traslados de Deuda
        Call Excel_Exportar_Lsw(lswOperacionesGeneradas, ProgressBarX)
End Select

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnReporteOperacion_Click()
Call sbReporte_Operacion
End Sub

Private Sub cboTipoCuotas_Click()
 Call sbConsulta_Mora
End Sub



Private Sub chkDeducirPlanilla_Click()
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError


If Not IsNumeric(txtOperacion) Then Exit Sub

If chkDeducirPlanilla.Value = vbChecked Then
 strSQL = "update reg_creditos set IND_DEDUCE_PLANILLA='S' where id_solicitud = " & txtOperacion
 Call Bitacora("Registra", "Indica la Deducción de Planilla de la OP: " & txtOperacion)
 Call sbBitacoraCredito("17", "Activa Deducción x Planilla", "C", txtOperacion.Text, txtCodigo.Text)
 
Else
 strSQL = "update reg_creditos set IND_DEDUCE_PLANILLA='N' where id_solicitud = " & txtOperacion
 Call Bitacora("Registra", "Indica la NO Deducción de Planilla de la OP: " & txtOperacion)
 Call sbBitacoraCredito("17", "DesActiva Deducción x Planilla", "C", txtOperacion.Text, txtCodigo.Text)

End If
Call ConectionExecute(strSQL)

MsgBox "Actualización Realizada...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub chkRepFechaTodas_Click()
If chkRepFechaTodas.Value = vbChecked Then
  dtpFechaInicio.Enabled = False
Else
  dtpFechaInicio.Enabled = True
End If

dtpFechaCorte.Enabled = dtpFechaInicio.Enabled

End Sub

Private Sub cmdAceptarFechas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprimir los Reportes Generales del esta ventana
'REFERENCIAS   : fxFechaServidor - (Devuelve la fecha del Servidor)
'OBSERVACIONES : Utiliza Variables Globales
'-------------------------------------------------------------------------------------------
Dim strRuta As String, strSQL As String, vSQLx As String

Me.MousePointer = vbHourglass


vSQLx = ""
If cboRepX.Text <> "TODOS" Then
   vSQLx = " AND {SOCIOS.ESTADOACTUAL} = '" & cboRepX.ItemData(cboRepX.ListIndex) & "'"
End If



With frmContenedor.Crt
 .Reset
 
 .Connect = glogon.ConectRPT
 
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Módulo de Cobros"
    .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"




Select Case lblRepOp.Tag
  Case "REVER" 'Lista de Casos con Reversión (CJ-TD)
    
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CasosConReversion.rpt")
    
    strSQL = "{REG_CREDITOS.PROCESO} = 'N'"
    
    If chkRepFechaTodas.Value = vbUnchecked Then
        strSQL = "{REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Format(dtpFechaInicio.Value, "yyyy,mm,dd") & ")" _
                & " AND {REG_CREDITOS.FECHA_ENVIAPROCESO} <= Date(" & Format(dtpFechaCorte.Value, "yyyy,mm,dd") & ")"
    End If
    
    .SelectionFormula = strSQL & vSQLx
    
    .Formulas(2) = "SubTitulo='DE " & dtpFechaInicio.Value & " HASTA " & dtpFechaCorte.Value & " / FILTRO " & Mid(cboRepX.Text, 4, 30) & "'"
    
    
  Case "ENVCBR" 'Lista de Operaciones en Cobro Judicial
    
    vSQLx = ""
    If cboRepX.Text <> "TODOS" Then
       vSQLx = " AND {vCbrCasosCobroJudicial.EstadoActual} = '" & cboRepX.ItemData(cboRepX.ListIndex) & "'"
    End If
    
    
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CasosEnCobroJudicial.rpt")
    
    strSQL = "{vCbrCasosCobroJudicial.PROCESO} = 'J'"
    
    If chkRepFechaTodas.Value = vbUnchecked Then
        strSQL = strSQL & " AND ({vCbrCasosCobroJudicial.FECHA_ENVIAPROCESO} >= Date(" & Format(dtpFechaInicio, "yyyy,mm,dd") _
                & ") AND {vCbrCasosCobroJudicial.FECHA_ENVIAPROCESO} <= Date(" & Format(dtpFechaCorte.Value, "yyyy,mm,dd") _
                & "))"
    End If
    
    .SelectionFormula = strSQL & vSQLx
    .Formulas(2) = "SubTitulo=' DE " & dtpFechaInicio & " HASTA " & dtpFechaCorte & " / FILTRO " & cboRepX.Text & "'"
  
  Case "TRADEUD" 'Operaciones con Traspaso de Deudas
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CasosTraspasoDeudas.rpt")
    
    strSQL = "{REG_CREDITOS.PROCESO} = 'T'"
    
    If chkRepFechaTodas.Value = vbUnchecked Then
        strSQL = "{REG_CREDITOS.FECHA_ENVIAPROCESO} >= Date(" & Format(dtpFechaInicio.Value, "yyyy,mm,dd") & ")" _
                & " AND {REG_CREDITOS.FECHA_ENVIAPROCESO} <= Date(" & Format(dtpFechaCorte.Value, "yyyy,mm,dd") & ")"
    End If
    
    .SelectionFormula = strSQL & vSQLx
    
  
  Case "TRAFIA" 'Operaciones de Fiadores con TD Aplicado
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CasosTraspasoFiadores.rpt")
    
    strSQL = "{REG_CREDITOS.FECHAFORP} >= Date(" & Format(dtpFechaInicio.Value, "yyyy,mm,dd") & ")" _
            & " AND {REG_CREDITOS.FECHAFORP} <= Date(" & Format(dtpFechaCorte.Value, "yyyy,mm,dd") _
            & ") AND IsNull ({REG_CREDITOS.REFERENCIA})=FALSE"
    
    .SelectionFormula = strSQL & vSQLx
    .Formulas(2) = "SubTitulo='TRASPASO DE DEUDAS / FILTRO " & Mid(cboRepX.Text, 4, 30) & "'"

End Select
 
 .PrintReport

End With

Me.MousePointer = vbDefault
fraFechas.Visible = False

End Sub


Private Sub cmdCancelaReversionTraspaso_Click()
 fraReversionDeTraspaso.Visible = False
End Sub


Function fxValidaPasoCobroJudicial() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As Boolean

On Error GoTo vError

vResultado = False

If Not IsNumeric(txtOperacion.Text) Then
   fxValidaPasoCobroJudicial = vResultado
   Exit Function
End If


strSQL = "select count(*) as 'Existe'" _
       & " from reg_Creditos R inner join catalogo C on R.codigo = C.codigo" _
       & " and C.retencion = 'N' and C.poliza = 'N'" _
       & " where R.id_solicitud = " & txtOperacion.Text _
       & " and Proceso <> 'J'"
Call OpenRecordSet(rs, strSQL)
   If rs!Existe = 1 Then
      vResultado = True
   End If
rs.Close
 
 
vError:
    fxValidaPasoCobroJudicial = vResultado

End Function



Private Sub sbCobroJudicial()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Ejecuta el Cobro Judicial a una Operación
'REFERENCIAS   : FxFechaServidor - (Devuelve la Fecha del servidor)
'                Bitacora - (Registra el movimiento realizado)
'OBSERVACIONES : Genera Asiento
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, strCuentas As String
Dim strObservacion As String, vFecha As Date, strLinea(11) As String
Dim vOficina As String, vUnidad As String, vDivisa  As String, vCuenta As String, vCuentaCbr As String
Dim vCentroCosto As String, pTipoDoc As String, pDocumento As String, pConcepto As String
Dim curTipoCambio As Currency, curSaldo As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

strObservacion = ""
'Aqui observacion
strObservacion = InputBox("Digite las Notas de este Cobro Judicial : ", "Cobro Judicial")
If Len(Trim(strObservacion)) = 0 Then strObservacion = "NADA"

strObservacion = fxDepuraString(strObservacion, "'")

'Extrae la Cuenta de Cobro Judicial y la Fecha
strSQL = "select CtaCAmort as 'Cuenta',dbo.MyGetdate() as 'Fecha'" _
       & " from catalogo " _
       & " where codigo = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
 vCuentaCbr = Trim(rs!Cuenta)
 vFecha = rs!fecha
rs.Close

'Extrae configuración Contable de la Operación
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
 vCuenta = Trim(rs!ctaamortiza)
 vOficina = Trim(rs!cod_oficina_r)
 vUnidad = Trim(rs!Cod_Unidad)
 vDivisa = Trim(rs!cod_Divisa)
 vCentroCosto = "" 'Trim(rs!cod_Centro_Costo)
 curTipoCambio = rs!TipoCambio
rs.Close

'Otro parámetros contables
pTipoDoc = "CBJ"
pConcepto = "CBR004"
pDocumento = ""

vAseDocCuenta = ""
vAseDocDeposito = ""
vAseDocDetalle = strObservacion



'Registro Contable

        pDocumento = fxDocumentoConsecutivo(pTipoDoc)
        
        'Lineas de Comprobante
        strLinea(1) = "Saldo Actual      " & txtSaldo.Text
        strLinea(2) = "Interes Corriente " & Format(mCurIntc, "Standard")
        strLinea(3) = "Interes Atrasado  " & Format(mCurIntm, "Standard")
        strLinea(4) = "Amortización Atra." & Format(mCurAmortAtrasada, "Standard")
        strLinea(5) = "Cargos Regist.    " & Format(mcurCargos, "Standard")
        strLinea(6) = "Divisa: " & vDivisa & " / Tipo Cambio: " & curTipoCambio
        strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text
        strLinea(8) = Mid(Trim(txtDescripcion.Text), 1, 30)
        strLinea(9) = ""
        strLinea(10) = Mid("Notas: " & strObservacion, 1, 30)
        strLinea(11) = "Póliza Atradada  " & Format(mcurPoliza, "Standard")
         
      
        'Registro del Comprobante
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                 & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                 & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
                 & " values('" & pDocumento & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
                 & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & CCur(txtSaldo.Text) & ",'P','" & txtOperacion.Text _
                 & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                 & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                 & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                 & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                 & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
         Call ConectionExecute(strSQL)
         
         'ASIENTO
        
         If CCur(txtSaldo.Text) > 0 Then
           curSaldo = CCur(txtSaldo.Text) * curTipoCambio
           strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pDocumento & "'," & curSaldo & ",'C','" & vDivisa _
                  & "'," & curTipoCambio & "," & GLOBALES.gEnlace & ",'" & vUnidad & "','" & vCentroCosto & "','" & vCuenta _
                  & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
           Call ConectionExecute(strSQL)
         
           strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pDocumento & "'," & curSaldo & ",'D','" & vDivisa _
                  & "'," & curTipoCambio & "," & GLOBALES.gEnlace & ",'" & vUnidad & "','" & vCentroCosto & "','" & vCuentaCbr _
                  & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
           Call ConectionExecute(strSQL)
         
         End If
         


'Inicia Transacciones
glogon.Conection.BeginTrans

'Actualiza reg_creditos campos : Fecha_enviaproceso,observacion_proceso,proceso
strSQL = "update reg_creditos set fecha_enviaproceso = '" & Format(vFecha, "yyyy/mm/dd") _
       & "',observacion_proceso = '" & strObservacion & "',proceso = 'J'" _
       & " where id_solicitud = " & Trim(txtOperacion)
Call ConectionExecute(strSQL)

'NUEVO : Actualiza ESTADOI en Morosidades para que no Acumele mas intereses moratorios
'SE actualiza con J - VERIFICAR EL PROCESO MENSUAL:

If GLOBALES.SysPlanPagos = 0 Then
        strSQL = "update morosidad set estadoi = 'J' where estado = 'A' and id_solicitud = " & txtOperacion.Text
        Call ConectionExecute(strSQL)
End If

'Registro Historial y Expediente
Call sbCBRRegTransac("02", txtCedula, txtOperacion, strObservacion, CCur(txtSaldo), mCurIntc, mCurIntm, mcurCargos, mcurPoliza, mCurAmortAtrasada, pTipoDoc, pDocumento)

'Cierra Transacciones
glogon.Conection.CommitTrans


'Información Final
txtProceso.Text = "COBRO JUDICIAL"

Call Bitacora("Aplica", "Cobro Judicial a la Operación:" & txtOperacion)

Me.MousePointer = vbDefault

If GLOBALES.SysDocVersion = 1 Then
    MsgBox "- La operación fue enviada a Cobro Judicial" & vbCrLf & vbCrLf _
         & "- Se generó Asiento (CBR" & txtOperacion & ")", vbInformation
Else
    'Control de Documentos v2
    MsgBox "- La operación fue enviada a Cobro Judicial" & vbCrLf & vbCrLf _
         & "- Se generó la nota de cobro número: " & pDocumento, vbInformation
    Call sbImprimeRecibo(pDocumento, pTipoDoc)
End If

Call sbHistorial(txtOperacion.Text)

Exit Sub

vError:
    Me.MousePointer = vbDefault
    glogon.Conection.RollbackTrans
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Function fxValidaTextosNumericos(txt As TextBox) As Boolean

fxValidaTextosNumericos = True

End Function


Function fxOpex(strCedula As String) As Integer
Dim rsX As New ADODB.Recordset
 
rsX.Source = "select estadoactual from socios where cedula = '" & strCedula & "'"
rsX.Open , glogon.Conection, adOpenStatic
 
If rsX!EstadoActual = "S" Or rsX!EstadoActual = "N" Then
 fxOpex = 0 'Socios y no Socios Cargan la misma Cuenta
Else
 fxOpex = 1 'Ren. Asociacion y Patrono cargan la misma Cuenta
End If
rsX.Close

End Function

Private Sub AsientoTraspasoFiadorDeudorF(curMonto As Currency, curIntC As Currency _
                , curIntM As Currency, curCargo As Currency, strCedula As String, strCodigo As String, vFecha As Date)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Crea el Asiento de REVERSION de un Traspaso de Deudas para los Fiadores
'REFERENCIAS   : fxFechaServidor - (Devuelve Fecha del Servidor)
'OBSERVACIONES : Ver Reversiones de Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rs As New ADODB.Recordset, strSQL As String

 If fxOpex(strCedula) = 0 Then
  strSQL = "select ctanintc as ctaIntc, ctanintm as ctaIntm, ctanamort as ctaAmortiza "
 Else 'cuentas opex
  strSQL = "select ctaointc as ctaIntc, ctaointm as ctaIntm, ctaoamort as ctaAmortiza "
 End If
 strSQL = strSQL & "from catalogo where codigo = '" & strCodigo & "'"
 Call OpenRecordSet(rs, strSQL)
 
 If curMonto > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(Date), "00") & "','" & rs!ctaamortiza & "'," & curMonto & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  Call ConectionExecute(strSQL)
 End If
 
 If curCargo > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(Date), "00") & "','" & fxCBRParametro("23") & "'," & curCargo & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  Call ConectionExecute(strSQL)
 End If
 
 
 If curIntC > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(Date), "00") & "','" & rs!ctaintc & "'," & curIntC & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  Call ConectionExecute(strSQL)
 End If
 
 If curIntM > 0 Then
  strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
      & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
      & "','" & txtOperacion & "-FD" & Format(Day(Date), "00") & "','" & rs!ctaintm & "'," & curIntM & ",'H','" _
      & Format(vFecha, "yyyy/mm/dd") & "','P')"
  Call ConectionExecute(strSQL)
 End If
 rs.Close
 
End Sub


Private Sub AsientoTraspasoFiadorDeudor(curMonto As Currency, strCedula As String, vFecha As Date)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Crea asiento de REVERSION de Traspaso de Deudas para el Deudor
'REFERENCIAS   : fxFechaServidor - (Devuelve la fecha del Sistema)
'OBSERVACIONES : Ver Reversion de Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim rsA As New ADODB.Recordset, strSQL As String

If fxOpex(strCedula) = 0 Then
  strSQL = "select ctanamort as ctaAmortiza "
Else 'cuentas exsocios
  strSQL = "select ctaoamort as ctaAmortiza "
End If
strSQL = strSQL & "from catalogo where codigo = '" & txtCodigo & "'"
rsA.Open strSQL, glogon.Conection, adOpenStatic

If curMonto > 0 Then
    strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
        & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
        & "','" & txtOperacion & "-FD" & Format(Day(Date), "00") & "','" & rsA!ctaamortiza & "'," & curMonto & ",'D','" _
        & Format(vFecha, "yyyy/mm/dd") & "','P')"
    Call ConectionExecute(strSQL)
End If
rsA.Close
End Sub

Private Sub sbBoletaTraslado()
Dim vTipoDoc As String

Me.MousePointer = vbHourglass

If GLOBALES.SysDocVersion = 1 Then
 vTipoDoc = "4"
Else
 vTipoDoc = "TRA"
End If

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Cobro Administrativo y Judicial"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "subtitulo='BOLETA DE TRASLADO Y REVERSION DE DEUDAS'"
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_BoletaTraspasoReversion.rpt")
    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
    
'    .SubreportToChange = "MORO"
'    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = {?Pm-REG_CREDITOS.ID_SOLICITUD} and {MOROSIDAD.TCON} = '" & vTipoDoc & "' and {MOROSIDAD.ESTADO} = 'C'"
    .PrintReport
End With

Me.MousePointer = vbDefault


End Sub


Private Sub cmdReversaTraspasoDeudas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Reversar el Traspaso de Deudas
'REFERENCIAS   : AsientoTraspasoFiadorDeudorF -(Crea Lineas de Asiento de Reversion - Fiadores)
'                AsientoTraspasoFiadorDeudor -(Crea Lineas de Asiento de Reversion - Deudor)
'                fxFechaServidor -(Devuelve la Fecha del Servidor)
'                Bitacora - (Registra el movimiento realizado)
'OBSERVACIONES : Se Ejecutan los casos seleccionados, Utiliza variables globales
'-------------------------------------------------------------------------------------------

Dim itmX As ListItem, lng As Long, strSQL As String
Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
Dim curIntCor As Currency, curIntMor As Currency, lngUltimaOperacion As Long, lngPriDeduc As Long
Dim curAmortiza As Currency, curTotalInd As Currency, vFecha As Date, vPaso As Boolean
Dim strObservacion As String, curIntPendiente As Currency, curCargo As Currency, curPoliza As Currency

Dim pTipoDoc As String, pTipoDocum As String, pNumDoc As String, pConcepto As String, pCtaCargos As String
Dim pOficina As String, pUnidad As String, pCentroCosto As String, pDivisa As String, pCtaAmortiza As String
Dim strLinea(11) As String, pBaseCalculo As String, pCuota As Currency, pDiaPago As Integer, pCtaPoliza As String
Dim vTransac As Boolean

On Error GoTo vError

If CCur(txtTRAFD_MONTO) = 0 Then Exit Sub

Me.MousePointer = vbHourglass


vTransac = False
curIntCor = 0
curIntMor = 0
curAmortiza = 0
curIntPendiente = 0
curCargo = 0
curPoliza = 0

pCtaCargos = fxCBRParametro("23")

vPaso = False

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_reversiones") Then
  MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
  Exit Sub
End If


'Verifica que exista una opcion marcada
With lswOperacionesGeneradas.ListItems
 For lng = 1 To .Count
  If .Item(lng).Checked Then vPaso = True
 Next lng
End With


If Not vPaso Then
  MsgBox "No se ha marcado ningun (deuda) de fiador, para reversión verifique...?", vbExclamation
  Exit Sub
End If

vFecha = fxFechaServidor

'Cancelar Operaciones de los fiadores Marcados
If Mid(Trim(str(GLOBALES.glngFechaCR)), 5, 6) = 12 Then
 lngPriDeduc = (Val(Mid(Trim(str(GLOBALES.glngFechaCR)), 1, 4)) + 1) & "01"
Else
  lngPriDeduc = GLOBALES.glngFechaCR + 1
End If

'Aqui observacion
strObservacion = InputBox("Digite la Observación para esta Reversión : ", "Observación")
If Len(Trim(strObservacion)) = 0 Then strObservacion = "NADA"


'Configuración de la Oficina y Contabilidad
strSQL = "select O.cod_oficina,O.cod_unidad,O.cod_centro_costo,R.cod_divisa,R.dia_pago, R.Base_calculo" _
       & " from reg_creditos R left join sif_oficinas O on R.cod_oficina_R = O.cod_oficina" _
       & " where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
    pOficina = Trim(rs!COD_OFICINA & "")
    pUnidad = rs!Cod_Unidad & ""
    pCentroCosto = rs!Cod_Centro_Costo & ""
    pDivisa = IIf(IsNull(rs!cod_Divisa), pDivisa, rs!cod_Divisa)
    pDiaPago = IIf(IsNull(rs!dia_pago), 32, rs!dia_pago)
    pBaseCalculo = IIf(IsNull(rs!Base_Calculo), "01", rs!Base_Calculo)
rs.Close


'Información Base de la Operacion (Re-escribe variables de Oficinas y Centros de Costos)
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
    pOficina = rs!cod_oficina_r
    pUnidad = rs!Cod_Unidad
    pCentroCosto = rs!Cod_Centro_Costo
    pCtaAmortiza = Trim(rs!ctaamortiza)
rs.Close



'Registro Inicial en Control de Documentos
If GLOBALES.SysDocVersion = 1 Then
  'Control de Documentos v1
  pTipoDoc = "4"
  pTipoDocum = "4"
  pNumDoc = "8888"
  pConcepto = "CBR003"
Else
    'Control de Documentos v2
    pTipoDoc = "TRA"
    pTipoDocum = "TRA"
    pNumDoc = fxDocumentoConsecutivo(pTipoDoc)
    pConcepto = "CBR003"
    vAseDocDetalle = strObservacion
    vAseDocDeposito = ""
    
    
    strLinea(1) = "Saldo Anterior    " & txtSaldo.Text
    strLinea(2) = "Interes Corriente " & Format(0, "Standard")
    strLinea(3) = "Interes Moratorio " & Format(0, "Standard")
    strLinea(4) = "Cargos            " & Format(0, "Standard")
    strLinea(5) = "Amortizacion      " & txtSaldo.Text
    strLinea(6) = "Saldo Actual      " & Format(0, "Standard")
    strLinea(7) = "Operación         " & txtOperacion.Text
    strLinea(8) = "Línea             " & txtCodigo.Text
    strLinea(9) = "Proc.Retencion    " & "NO"
    strLinea(10) = "Usuario           " & glogon.Usuario
    strLinea(11) = "Póliza            " & Format(0, "Standard")
          
    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
             & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
             & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento,linea11)" _
             & " values('" & pNumDoc & "','" & pTipoDocum & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
             & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & CCur(txtTRAFD_MONTO) & ",'P','" & txtOperacion.Text _
             & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
             & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
             & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
             & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
             & vAseDocDetalle & "','" & vAseDocDeposito & "','" & strLinea(11) & "')"
     Call ConectionExecute(strSQL)
End If




'Inicia Transacciones
glogon.Conection.BeginTrans
vTransac = True

 With lswOperacionesGeneradas.ListItems
   For lng = 1 To .Count
      If .Item(lng).Checked And (CCur(.Item(lng).SubItems(4)) + CCur(.Item(lng).SubItems(8))) > 0 Then
        If GLOBALES.SysPlanPagos = 1 Then
           'Con Plan de Pagos Cancelar la Operacion
           '1. Calcular cancelación, 2. Aplicar Cancelación, 3. Asiento de Aplicacion
           
          'Actualiza Estado del Plan de Pago
          strSQL = "exec spCrdPlanPagosMoraActualizaOp " & .Item(lng).Text & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "'"
          Call ConectionExecute(strSQL)
          
          strSQL = "exec spCrdPlanPagosInfoCancelacion " & .Item(lng).Text & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "'"
          Call OpenRecordSet(rsTmp, strSQL, 0)
              curIntCor = rsTmp!IntCor
              curIntMor = rsTmp!IntMor
              curIntPendiente = 0 'Intereses a hoy
              curCargo = rsTmp!Cargos
              curPoliza = rsTmp!Poliza
              curAmortiza = rsTmp!Principal
              curTotalInd = curIntCor + curIntMor + curCargo + curPoliza + curAmortiza
          rsTmp.Close
          
          'Aplica Abono de Cancelación
          strSQL = "exec spCrdPlanPagoAbonoCancelacion " & .Item(lng).Text & ",'" & pConcepto & "','" & glogon.Usuario & "','" & pTipoDocum _
                 & "','" & pNumDoc & "'," & curTotalInd & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',''"
          Call ConectionExecute(strSQL)
         
         '2017-09-27: Pedro
         'Ajuste Operación del Deudor (Original)
         strSQL = "exec spCrdPlanPagoAnulaAbono " & txtOperacion.Text & ",'" & pConcepto & "','" & glogon.Usuario & "','" & pTipoDocum _
                & "','" & pNumDoc & "',1,0,0," & curTotalInd & ",0,0,'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") _
                & "','',1,1"
         Call ConectionExecute(strSQL)
         
        Else
           'Sin Plan de Pagos
            strSQL = "select sum(intc) as Intc, sum(intm) as Intm,sum(amortiza) as Amortiza, sum(Cargo) as Cargo" _
                   & " from morosidad " _
                   & " where estado = 'A' and id_solicitud = " & .Item(lng).Text
            Call OpenRecordSet(rs, strSQL)
                curIntCor = IIf(IsNull(rs!IntC), 0, rs!IntC)
                curIntMor = IIf(IsNull(rs!IntM), 0, rs!IntM)
                curAmortiza = IIf(IsNull(rs!Amortiza), 0, rs!Amortiza)
                curIntPendiente = CCur(.Item(lng).SubItems(8)) - (curIntCor + curIntMor)
                curCargo = IIf(IsNull(rs!Cargo), 0, rs!Cargo)
                curPoliza = 0
            rs.Close
            
            'Inserta en creditos DT Cancelación (CCur(.SelectedItem.SubItems(4)) - curAmortiza)
            curAmortiza = CCur(.Item(lng).SubItems(4)) - curAmortiza
            
            strSQL = "Update morosidad set abintc = intc, abintm = intm, abamortiza = amortiza, abCargo = Cargo, usuario = '" & glogon.Usuario _
                   & "', cod_concepto = '" & pConcepto & "', cod_caja = ''" _
                   & ", estado = 'C', tcon = '" & pTipoDoc & "',ncon = '" & pNumDoc & "',fecult = dbo.MyGetdate()" _
                   & " where estado = 'A' and id_solicitud = " & .Item(lng).Text
            
            'Deberia de realizar registro en creditos_dt
            strSQL = strSQL & Space(10) & "update reg_creditos set saldo = saldo - " & CCur(.Item(lng).SubItems(4)) & ", amortiza = amortiza + " & CCur(.Item(lng).SubItems(4)) _
                   & ", interesc = interesc + " & CCur(.Item(lng).SubItems(8)) & ", estado = 'C', Proceso = 'N'" _
                   & ",fecha_enviaproceso = dbo.MyGetdate()" _
                   & ",observacion_proceso = '" & Mid(strObservacion, 1, 255) _
                   & "' where id_solicitud = " & .Item(lng).Text
            
            
            strSQL = strSQL & Space(10) & "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
                   & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO,usuario,cod_concepto,cod_caja) values('" & .Item(lng).SubItems(1) & "'," _
                   & .Item(lng).Text & ",0," & curAmortiza + curIntPendiente _
                   & "," & curIntPendiente & "," & curAmortiza & ",dbo.MyGetdate()" _
                   & "," & GLOBALES.glngFechaCR & ",'" & pTipoDoc & "','" & pNumDoc & "','A','G','" & glogon.Usuario & "','" & pConcepto & "','')"
         
            '2017-09-27: Pedro
            'Ajuste Operación del Deudor (Original)
            strSQL = strSQL & Space(10) & "update reg_creditos set Proceso = 'N', Estado = 'A'" _
                  & ",saldo = saldo + " & CCur(.Item(lng).SubItems(4)) + CCur(.Item(lng).SubItems(8)) _
                  & " where id_solicitud = " & txtOperacion.Text
            
            strSQL = strSQL & Space(10) & "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
                   & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO,usuario,cod_concepto,cod_caja) values('" & txtCodigo.Text & "'," _
                   & txtOperacion.Text & ",0," & (CCur(.Item(lng).SubItems(4)) + CCur(.Item(lng).SubItems(8))) * -1 _
                   & ",0," & (CCur(.Item(lng).SubItems(4)) + CCur(.Item(lng).SubItems(8))) * -1 & ", dbo.MyGetdate()" _
                   & "," & GLOBALES.glngFechaCR & ",'" & pTipoDoc & "','" & pNumDoc & "','A','G','" & glogon.Usuario & "','" & pConcepto & "','')"
            
            'Aplica Lote
            Call ConectionExecute(strSQL)
            
            
        End If 'SysPlanPagos = 1
        
      End If '.item(lng).checked
   
   Next lng
 
 End With


'2017-09-27: Pedro (Elimina la Generación de una Operacion Nueva todo se realiza sobre la original)
'If Val(lblOperacionActualDeudor) = 0 Then
'  'Insertar Nueva Operacion
'   strSQL = "insert into reg_creditos(codigo,id_comite,cedula,montosol,estadosol,fechares" _
'          & ",plazo,int,interesv,montoapr,prideduc,fechaforp,fechaforf,saldo,amortiza,interesc" _
'          & ",cuota,referencia,userrec,userres,userfor,garantia,firma_deudor" _
'          & ",monto_girado,cuotas_planilla,cuotas_directas,cuotas_anuladas,Tesoreria,opex" _
'          & ",OBSERVACION_PROCESO,FECULT,TBP_PuntosAdd,LiqTasa,cod_oficina_r,cod_oficina_f) values" _
'          & "('" & txtCodigo & "',1,'" & Trim(txtCedula) & "'," & CCur(txtTRAFD_MONTO) & ",'F','" & Format(vFecha, "yyyy/mm/dd") & "'" _
'          & "," & txtTRAFD_Plazo & "," & txtTRAFD_Int & "," & txtTRAFD_Int & "," & CCur(txtTRAFD_MONTO) & "," & lngPriDeduc _
'          & ",'" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "'," & CCur(txtTRAFD_MONTO) _
'          & ",0,0," & Format(fxCalcula_Cuota(CLng(CCur(txtTRAFD_MONTO)), txtTRAFD_Plazo, txtTRAFD_Int), "##########0.00") & "," & txtOperacion _
'          & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "','F',1" _
'          & ",0,0,0,0,'" & Format(vFecha, "yyyy/mm/dd") & "'," & fxOpex(txtCedula) & ",'" & strObservacion & "'," _
'          & GLOBALES.glngFechaCR & "," & IIf((mTasaPts = -1000), "Null", mTasaPts) & "," & mTasaLiq & ",'" & pOficina _
'          & "','" & GLOBALES.gOficinaTitular & "')"
'    Call ConectionExecute(strSQL)
'
'  'Recupera la nuevo operacion
'   lngUltimaOperacion = fxUltimaOperacion(txtCedula.Text)
'   lblOperacionActualDeudor.Caption = lngUltimaOperacion
'  'Hereda Fiadores Operacion Anterior
'  strSQL = "insert into fiadores(id_solicitud,codigo,cedulaf,nombre,firma,estado,interno" _
'         & ",salario,devengado,liquidez) (select " & lngUltimaOperacion & ",codigo,cedulaf," _
'         & "nombre,firma,estado,interno,salario,devengado,liquidez from fiadores" _
'         & " where id_solicitud = " & txtOperacion & ")"
'  Call ConectionExecute(strSQL)
'
'Else
'  'Actualizar Operacion
'  strSQL = "update reg_creditos set montoapr = montoapr + " & CCur(txtTRAFD_MONTO) _
'        & ",montosol = montosol + " & CCur(txtTRAFD_MONTO) _
'        & ",saldo = saldo + " & CCur(txtTRAFD_MONTO) _
'        & ",plazo = " & CCur(txtTRAFD_Plazo) _
'        & ",interesv = " & CCur(txtTRAFD_Int) _
'        & ",cuota = " & CCur(txtTRAFD_Cuota) & " where id_solicitud = " & lblOperacionActualDeudor.Caption
'  Call ConectionExecute(strSQL)
'End If


'Cierra Transacciones
glogon.Conection.CommitTrans
vTransac = False

''Activacion del Plan de Pagos de la Nueva Operacion
'If GLOBALES.SysPlanPagos = 1 Then
'   strSQL = "exec spCrdPlanPagos " & lblOperacionActualDeudor.Caption
'   Call ConectionExecute(strSQL)
'End If


''Asientos
'strSQL = "0"
'With lswOperacionesGeneradas.ListItems
'  For lng = 1 To .Count
'     If .Item(lng).Checked And CCur(.Item(lng).SubItems(4)) > 0 Then
'        strSQL = .Item(lng) & "," & strSQL
'     End If
'  Next lng
'End With
    
strSQL = "select V.id_solicitud,V.codigo,V.cedula,V.COD_CONCEPTO,Ofi.COD_UNIDAD,Ofi.COD_CENTRO_COSTO " _
       & ",SUM(V.IntCor) as 'IntCor',SUM(V.IntMor) as 'IntMor'" _
       & ",SUM(V.Cargo) as 'Cargo', SUM(V.Poliza) as 'Poliza', SUM(V.Principal) as 'Principal'" _
       & ",SUM(V.IntCor + V.IntMor + V.Cargo + V.Poliza + V.Principal) as 'Total'" _
       & " , case when Reg.PROCESO = 'J' then Cat.CtaCIntC else Cat.CTANINTC end as 'CtaIntC'" _
       & " , case when Reg.PROCESO = 'J' then Cat.CtaCIntM else Cat.CTANINTM end as 'CtaIntM'" _
       & " , case when Reg.PROCESO = 'J' then Cat.CtaCAmort else Cat.CTANAMORT end as 'CtaAmortiza'" _
       & " from vCRDsReportesMov V inner join SIF_OFICINAS Ofi on ISNULL(V.cod_Oficina_R,'" & GLOBALES.gOficinaTitular & "') = Ofi.COD_OFICINA" _
       & "  inner join REG_CREDITOS Reg on V.id_solicitud = Reg.ID_SOLICITUD" _
       & "  inner join CATALOGO Cat on Reg.CODIGO = Cat.CODIGO" _
       & " where V.Tcon = '" & pTipoDoc & "' and V.Ncon = '" & pNumDoc _
       & "' and Reg.Id_Solicitud <> " & txtOperacion.Text _
       & " group by V.id_solicitud,V.codigo,V.cedula,V.COD_CONCEPTO,Ofi.COD_UNIDAD,Ofi.COD_CENTRO_COSTO" _
       & ",Reg.PROCESO,Cat.CtaCIntC,Cat.CTANINTC,Cat.CtaCIntM,Cat.CTANINTM,Cat.CtaCAmort,Cat.CTANAMORT"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
        
        If rs!Total > 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rs!Total & ",'D','" & pDivisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & pCtaAmortiza _
                 & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        End If
        
        If rs!Principal > 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rs!Principal & ",'C','" & pDivisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','','" & rs!ctaamortiza _
                 & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        End If
        
        
        If rs!Cargo > 0 And GLOBALES.SysPlanPagos = 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rs!Cargo & ",'C','" & pDivisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & pCtaCargos _
                 & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        End If
        
        If rs!Cargo > 0 And GLOBALES.SysPlanPagos = 1 Then
             'Detallar Cargos
              strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
              Call OpenRecordSet(rsTmp, strSQL, 0)
              Do While Not rsTmp.EOF
                    strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Mov_Monto & ",'C','" & rs!cod_Divisa _
                           & "',1," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                           & "','" & rsTmp!ID_SOLICITUD & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
                    Call ConectionExecute(strSQL)
                    rsTmp.MoveNext
              Loop
              rsTmp.Close

        End If
        
       If rs!Poliza > 0 Then
           strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!ID_SOLICITUD & ") as 'Cuenta'"
           Call OpenRecordSet(rsTmp, strSQL, 0)
             pCtaPoliza = Trim(rsTmp!Cuenta)
           rsTmp.Close
        
           strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pCtaPoliza & "'," & rs!Poliza & ",'C','" & rs!cod_Divisa _
                  & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & pCtaPoliza _
                  & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
           Call ConectionExecute(strSQL)
       End If

        
        If rs!IntCor > 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rs!IntCor & ",'C','" & pDivisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
                 & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        End If
        
        If rs!IntMor > 0 Then
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rs!IntMor & ",'C','" & pDivisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
                 & "','" & rs!ID_SOLICITUD & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        End If


 rs.MoveNext
Loop
rs.Close
    
        
'Actualiza Estado

strSQL = "update reg_creditos set PROCESO = 'N' where proceso = 'T' and estado = 'A'" _
       & "   and id_solicitud in(select referencia From reg_creditos   Where referencia = " & txtOperacion.Text _
       & "   group by referencia Having Sum(Saldo) <= 0)"
Call ConectionExecute(strSQL)


''Eliminados
'Call AsientoTraspasoFiadorDeudorF(CCur(.Item(lng).SubItems(4)), (curIntCor + curIntPendiente), curIntMor, curCargo _
'           , Trim(.Item(lng).SubItems(2)), .Item(lng).SubItems(1), vFecha)
'Call AsientoTraspasoFiadorDeudor(CCur(txtTRAFD_MONTO), txtCedula, vFecha)



'BITACORA
Call Bitacora("Reversa", "Traspaso de Deudas de la Operación:" & txtOperacion)

'Registro Historial y Expediente
Call sbCBRRegTransac("05", txtCedula, txtOperacion, strObservacion, CCur(txtTRAFD_MONTO), 0, 0, 0, 0, 0, pTipoDoc, pNumDoc)


Me.MousePointer = vbDefault


MsgBox "- Reversión de Traspaso Realizada Satisfactoriamente..." _
      & vbCrLf & vbCrLf & "- Revisar nota de Traspaso No.: " & pNumDoc, vbInformation

Call sbBoletaTraslado
fraReversionDeTraspaso.Visible = False

Call OperacionesGeneradas

Exit Sub

vError:
 Me.MousePointer = vbDefault
 If vTransac Then
     glogon.Conection.RollbackTrans
 End If
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub Form_Activate()
 vModulo = 4
 
' Call Formularios(Me)
' Call RefrescaTags(Me)
End Sub


Private Sub Form_Load()
Dim strSQL As String

 vModulo = 4
 
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call sbTaskPanel_Load

 Call Formularios(Me)
 Call RefrescaTags(Me)


 vOperacion = False 'Inicializar
 dtpFechaInicio.Value = fxFechaServidor
 dtpFechaCorte.Value = dtpFechaInicio.Value
 dtpCartaCorte.Value = dtpFechaInicio.Value
 dtpCartaCorte.MinDate = dtpFechaInicio.Value
 
 dtpCalculoIntCorte.Value = dtpFechaInicio.Value
 
 
strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as Itmx" _
       & " from afi_estados_persona"
Call sbCbo_Llena_New(cboRepX, strSQL, True, True)
'cboRepX.AddItem "X - Ex.Socios"


  
cboTipoCuotas.Clear
cboTipoCuotas.AddItem "Canceladas"
cboTipoCuotas.AddItem "Pendientes"
cboTipoCuotas.AddItem "Todas"
cboTipoCuotas.Text = "Todas"
 
With lswOperacionesGeneradas.ColumnHeaders
    .Clear
    .Add , , "Operación", 1300
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Cédula", 1400
    .Add , , "Monto", 1400, vbRightJustify
    .Add , , "Saldo", 1400, vbRightJustify
    .Add , , "Cuota", 1300, vbRightJustify
    .Add , , "Tasa", 1100, vbRightJustify
    .Add , , "Plazo", 1100, vbRightJustify
    .Add , , "Int.Pendiente", 1300, vbRightJustify
    .Add , , "Cargos", 1300, vbRightJustify
    .Add , , "Pólizas", 1300, vbRightJustify
    .Add , , "Estado", 1300, vbCenter
    .Add , , "Nombre", 3400
End With

With lswCF.ColumnHeaders
    .Clear
    .Add , , "Operación", 1300
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Cédula", 1400
    .Add , , "Nombre", 3400
    .Add , , "Cuota", 1300, vbRightJustify
    .Add , , "Recaudo", 1400, vbRightJustify
    .Add , , "Aplicado", 1400, vbRightJustify
    .Add , , "Devuelto", 1400, vbRightJustify
    .Add , , "Inicio", 1100, vbRightJustify
    .Add , , "Ult.Mov.", 1100, vbRightJustify
    .Add , , "Estado", 1300, vbRightJustify
End With
lswCF.Checkboxes = False


With lswAbonos.ColumnHeaders
    .Clear
    .Add , , "Proceso", 1200
    .Add , , "Fecha", 2400
    .Add , , "Abono", 1200, 1
    .Add , , "Int.Cor.", 1200, 1
    .Add , , "Int.Mor.", 1200, 1
    .Add , , "Cargos", 1200, 1
    .Add , , "Pólizas", 1200, 1
    .Add , , "Amortización", 1200, 1
    .Add , , "T.Comp", 1000
    .Add , , "N.Comp", 1200
    .Add , , "Usuario", 1200
    .Add , , "Concepto", 1200
End With

lswRepOp.ColumnHeaders.Clear
lswRepOp.ColumnHeaders.Add , , "Reporte", 4550

lswAvisos.ColumnHeaders.Clear
lswAvisos.ColumnHeaders.Add , , "Fecha", 1800
lswAvisos.ColumnHeaders.Add , , "Tipo", 1800



Me.MousePointer = vbDefault

Call sbLimpiaDatos
 
 
End Sub


Private Sub sbLswMovCrd()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Llena Lsw de Cuotas con el Listado de Abonos Ordinarios y Extraordinarios
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim curTotales(5) As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass


curTotales(0) = 0
curTotales(1) = 0
curTotales(2) = 0
curTotales(3) = 0
curTotales(4) = 0
curTotales(5) = 0

tcMain.Item(6).Selected = True

lswAbonos.ListItems.Clear

If txtOperacion = "" Then Exit Sub


If GLOBALES.SysPlanPagos = 1 Then
    
    Select Case Mid(cboTipoCuotas.Text, 1, 1)
        Case "P" 'Pendiente
            strSQL = "select IntCor ,IntMor, Cargos as 'Cargo',Poliza,Principal" _
                   & ",Fecha_Proceso as 'Proceso',Fecha_Pago as 'Fecha','' as 'Tipo', '' as 'Ncon','Pendiente' as 'Concepto','' as 'Usuario'" _
                   & " From Crd_Operacion_Transac" _
                   & " Where id_solicitud = " & txtOperacion.Text & " and Estado = 'A'" _
                   & " order by Fecha_Pago desc"
        Case "C" 'Canceladas
            strSQL = "select Mov_IntCor as 'IntCor',Mov_IntMor as 'IntMor', Mov_Cargos as 'Cargo',Mov_Poliza as 'Poliza',Mov_Principal as 'Principal'" _
                   & ",Fecha_Proceso as 'Proceso',Mov_Fecha as 'Fecha',Tipo_Documento as 'Tipo', NUM_COMPROBANTE as 'Ncon',COD_Concepto as 'Concepto',Mov_Usuario as 'Usuario'" _
                   & " From Crd_Operacion_Transac" _
                   & " Where id_solicitud = " & txtOperacion.Text & " and Estado in('C','N')" _
                   & " order by FECHA_PAGO desc"
        
        Case "T" 'Todas
            strSQL = "select isnull(Mov_IntCor, IntCor) as 'IntCor' , isnull(Mov_IntMor, IntMor) as 'IntMor'" _
                   & ", isnull(Mov_Cargos, Cargos) as 'Cargo',isnull(Mov_Poliza, Poliza) as 'Poliza'" _
                   & ", isnull(Mov_Principal, Principal) as 'Principal'" _
                   & ",Fecha_Proceso as 'Proceso',isnull(Mov_Fecha,Fecha_Pago) as 'Fecha', isnull(Tipo_Documento,'Pendiente') as 'Tipo', NUM_COMPROBANTE as 'Ncon',isnull(COD_Concepto,'Pendiente') as 'Concepto',Mov_Usuario as 'Usuario'" _
                   & " From Crd_Operacion_Transac Where id_solicitud = " & txtOperacion.Text _
                   & " order by Fecha_Pago desc"
    End Select
Else
    
    Select Case Mid(cboTipoCuotas.Text, 1, 1)
        Case "P" 'Pendiente
            strSQL = "select IntC as 'IntCor',IntM as 'IntMor', Cargo,0 as 'Poliza',Amortiza as 'Principal'" _
                   & ",Fechap as 'Proceso',isnull(FecUlt,dbo.MyGetdate()) as 'Fecha','' as 'Tipo', '' as 'Ncon','' as 'Concepto','' as 'Usuario'" _
                   & " From Morosidad Where id_solicitud = " & txtOperacion.Text _
                   & " and Estado = 'A'" _
                   & " order by FechaP desc"
        Case "C" 'Canceladas
            strSQL = "select IntCor,IntMor,Cargo,Poliza,Principal,Proceso,Fecha,Tipo,Ncon,Concepto,Usuario" _
                   & " From vCRDsReportesMov Where id_solicitud = " & txtOperacion.Text _
                   & " order by Fecha desc"
        
        Case "T" 'Todas
            strSQL = "select IntCor,IntMor,Cargo,Poliza,Principal,Proceso,Fecha,Tipo,Ncon,Concepto,Usuario" _
                   & " From vCRDsReportesMov Where id_solicitud = " & txtOperacion.Text _
                   & " UNION " _
                   & " select IntC as 'IntCor',IntM as 'IntMor', Cargo,0 as 'Poliza',Amortiza as 'Principal'" _
                   & ",Fechap as 'Proceso',isnull(FecUlt,dbo.MyGetdate()) as 'Fecha','' as 'Tipo', '' as 'Ncon','' as 'Concepto','' as 'Usuario'" _
                   & " From Morosidad Where id_solicitud = " & txtOperacion.Text _
                   & " and Estado = 'A'" _
                   & " order by Fecha desc"
    End Select
End If

     
Call OpenRecordSet(rs, strSQL)
 Do While Not rs.EOF
  Set itmX = lswAbonos.ListItems.Add(, , Format(IIf(IsNull(rs!Proceso), "", rs!Proceso), "####-##"))
   itmX.SubItems(1) = IIf(IsNull(rs!fecha), Date, rs!fecha)
   itmX.SubItems(2) = Format(rs!IntCor + rs!IntMor + rs!Cargo + rs!Poliza + rs!Principal, "Standard")
   itmX.SubItems(3) = Format(rs!IntCor, "Standard")
   itmX.SubItems(4) = Format(rs!IntMor, "Standard")
   itmX.SubItems(5) = Format(rs!Cargo, "Standard")
   itmX.SubItems(6) = Format(rs!Poliza, "Standard")
   itmX.SubItems(7) = Format(rs!Principal, "Standard")
   itmX.SubItems(8) = rs!Tipo & ""
   itmX.SubItems(9) = rs!nCon & ""
   itmX.SubItems(10) = rs!Usuario & ""
   itmX.SubItems(11) = rs!CONCEPTO & ""
   
   curTotales(0) = curTotales(0) + IIf(IsNull(rs!Cargo), 0, rs!Cargo)
   curTotales(1) = curTotales(1) + IIf(IsNull(rs!IntCor), 0, rs!IntCor)
   curTotales(2) = curTotales(2) + IIf(IsNull(rs!IntMor), 0, rs!IntMor)
   curTotales(3) = curTotales(3) + IIf(IsNull(rs!Poliza), 0, rs!Poliza)
   curTotales(4) = curTotales(4) + IIf(IsNull(rs!Principal), 0, rs!Principal)
   curTotales(5) = curTotales(5) + rs!Principal + rs!IntCor + rs!IntMor + rs!Cargo + rs!Poliza
   
  rs.MoveNext
 Loop
 rs.Close
 
  Set itmX = lswAbonos.ListItems.Add(, , "")
   itmX.SubItems(2) = "---------------"
   itmX.SubItems(3) = "---------------"
   itmX.SubItems(4) = "---------------"
   itmX.SubItems(5) = "---------------"
   itmX.SubItems(6) = "---------------"
   itmX.SubItems(7) = "---------------"
   
 
  Set itmX = lswAbonos.ListItems.Add(, , "Totales")
   itmX.SubItems(2) = Format(curTotales(5), "Standard")
   itmX.SubItems(3) = Format(curTotales(1), "Standard")
   itmX.SubItems(4) = Format(curTotales(2), "Standard")
   itmX.SubItems(5) = Format(curTotales(0), "Standard")
   itmX.SubItems(6) = Format(curTotales(3), "Standard")
   itmX.SubItems(7) = Format(curTotales(4), "Standard")
 
Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


Private Sub sbBusqueda(Index As Integer)

gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Index
  Case 1 'txtOperacion
    gBusquedas.Consulta = "select id_solicitud as Operacion,codigo,cedula,montoapr,saldo from reg_creditos"
    gBusquedas.Orden = "id_solicitud"
    gBusquedas.Columna = "id_solicitud"
    gBusquedas.Filtro = " and estadosol = 'F'"
    frmBusquedas.Show vbModal
    txtOperacion = gBusquedas.Resultado
    If Len(Trim(txtOperacion)) > 0 Then
      Call sbConsulta
    End If
  Case 2 'txtCodigo
   If Len(Trim(txtCedula)) > 0 Then
        gBusquedas.Consulta = "select id_solicitud as Operacion,codigo,cedula,proceso,estado from reg_creditos"
        gBusquedas.Orden = "id_solicitud"
        gBusquedas.Columna = "id_solicitud"
        gBusquedas.Filtro = " and estadosol = 'F' and cedula ='" & txtCedula & "'"
        frmBusquedas.Show vbModal
        txtOperacion = gBusquedas.Resultado
        If Len(Trim(txtOperacion)) > 0 Then
          Call sbConsulta
        End If
    Else
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
    End If
  
  Case 3 'txtCedula
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  Case 4 'txtDescripcion
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  Case 5 'txtNombre
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
        Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
'  Case 7 'txtReporteX
'        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
'        gBusquedas.Orden = "codigo"
'        gBusquedas.Columna = "codigo"
'        frmBusquedas.Show vbModal
'        txtReporteX = gBusquedas.Resultado
'        If Len(Trim(txtReporteX)) > 0 Then
'          lblXDescribe.Caption = fxDescribeCodigo(Trim(txtReporteX))
'        End If

End Select

End Sub

Function fxSaldo(lngSolicitud As Long)
Dim rsX As New ADODB.Recordset
With rsX
 .Open "select saldo from reg_creditos where id_solicitud = " & lngSolicitud, glogon.Conection, adOpenStatic
 If .EOF And .BOF Then
  fxSaldo = 0
 Else
  fxSaldo = !Saldo
 End If
 .Close
End With
End Function


Private Sub sbAdjuntos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

i = MsgBox("Desea Imprimir el Estado de Cuenta de los Fiadores?", vbYesNo)

Call sbReporteOpCtasPendientes

'Llamar el Estado de Cuenta
Call sbEstadoCuenta(txtCedula)

'Estado de Cuentas de los Fiadores
If i = vbYes Then
    strSQL = "select cedulaf from fiadores where estado = 'A' and id_solicitud = " & txtOperacion
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Call sbEstadoCuenta(rs!cedulaf)
     rs.MoveNext
    Loop
    rs.Close
End If

Exit Sub

vError:

End Sub



Private Sub imgCalculoInt_Click(Index As Integer)
Select Case Index
  Case 0
    Call sbMoraActiva(False)
  Case 1
    Call sbMoraActiva(False)
    Call sbReporteOpCtasPendientes
End Select
End Sub

Private Sub imgReporteCuotas_Click()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprime los reportes de abonos y cuotas
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor
'OBSERVACIONES : Utilizar varibles globales
'-------------------------------------------------------------------------------------------
Dim vFecha As Date

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Crédito"
    
    .Connect = glogon.ConectRPT
    
    If GLOBALES.SysPlanPagos = 0 Then
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_AbonosOperacionFull.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='ABONOS ORDINARIOS/EXTRAORDINARIOS/MORATORIOS'"
        .Formulas(2) = "Fecha='" & Format(vFecha, "dd/mm/yyyy") & "'"
        .Formulas(3) = "Titulo='MOVIMIENTOS DE LA OPERACION'"
        .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & txtOperacion.Text
        
        .SubreportToChange = "sbCorte"
        .StoredProcParam(0) = txtOperacion.Text
        .StoredProcParam(1) = Format(vFecha, "yyyy/mm/dd")
        
        .SubreportToChange = "sbMovimientos"
        
        .StoredProcParam(0) = txtOperacion.Text
        .StoredProcParam(1) = 1
        
    Else
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanPagosMov.rpt")
        
         .Formulas(0) = "fxFecha='FECHA: " & Format(vFecha, "dd/mm/yyyy  hh:mm:ss") & "'"
         .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
         .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
         .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
         
         .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & txtOperacion.Text
         
         .SubreportToChange = "sbCorte"
         .StoredProcParam(0) = txtOperacion.Text
         .StoredProcParam(1) = Format(vFecha, "yyyy/mm/dd")
    
    End If
    .PrintReport

   
End With
Me.MousePointer = vbDefault

End Sub


Private Sub sbReporte_Operacion()

'-------------------------------------------------------------------------------------------
'OBJETIVO      : Imprime Reportes sobre la operacion u operaciones
'REFERENCIAS   : fxFechaServidor - (Devuelve la Fecha del Servidor)
'OBSERVACIONES : Utiliza variables globales
'-------------------------------------------------------------------------------------------
Dim strRuta As String, strSQL As String, vMes As Integer
Dim rs As New ADODB.Recordset

strSQL = "select * from par_ahcr"
Call OpenRecordSet(rs, strSQL)
vMes = Mid(GLOBALES.glngFechaCR, 5, 2)
If rs!cr_apl = 0 Then
 If vMes = 1 Then
   vMes = 12
 Else
   vMes = vMes - 1
 End If
End If
rs.Close

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Cobro Administrativo y Judicial"

 .Connect = glogon.ConectRPT

Select Case lblRepOp.Tag
  Case "ULTEC" 'Boleta de Cobro Judicial / Boleta de Traspaso de Deudas
    Select Case txtProceso.Text
     Case "TRASPASO DEUDAS"
            .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
            .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
            .Formulas(2) = "subtitulo='BOLETA DE TRASLADO Y REVERSION DE DEUDAS'"
            .ReportFileName = SIFGlobal.fxPathReportes("Cobro_BoletaTraspasoReversion.rpt")
          
     Case "COBRO JUDICIAL"
          .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
          .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          .ReportFileName = SIFGlobal.fxPathReportes("Cobro_BoletaCobroJudicial.rpt") '******************************** No Existe
    End Select
    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
  
  Case "ECBR" 'EC CBR Resumen

  Case "ETSBR" 'Etiquetas Sobres
         .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Sobres.rpt")
         .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
  
  Case "PRIAVI" 'Primer Aviso
     'Registro Historial y Expediente
     Call sbCBRRegTransac("09", txtCedula, txtOperacion, "Carta de Primer Aviso...Corte:" & Format(dtpCartaCorte.Value, "dd/mm/yyyy"), CCur(txtSaldo), mCurIntc, mCurIntm)

  
     .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CartaPrimerAviso.rpt")
     .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
     .SelectionFormula = "{SOCIOS.CEDULA} = '" & Trim(txtCedula) & "'"
     .SubreportToChange = "Fiadores"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 And {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
  
       .SubreportToChange = "sbMora"
       .Formulas(0) = "fxCorte = 'Fecha de Corte para Cálculo de Intereses : " & Format(dtpCartaCorte.Value, "dd/mm/yyyy") & "'"
       .StoredProcParam(0) = Trim(txtCedula)
       .StoredProcParam(1) = Format(dtpCartaCorte.Value, "yyyy/mm/dd")
       
  Case "SEGAVI" 'Segundo Aviso
     'Registro Historial y Expediente
     Call sbCBRRegTransac("10", txtCedula, txtOperacion, "Carta de Segundo Aviso...Corte:" & Format(dtpCartaCorte.Value, "dd/mm/yyyy"), CCur(txtSaldo), mCurIntc, mCurIntm)
     
     .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CartaSegundoAviso.rpt")
     .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
     .SelectionFormula = "{SOCIOS.CEDULA} = '" & Trim(txtCedula) & "'"
     .SubreportToChange = "Fiadores"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 And {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
  
       .SubreportToChange = "sbMora"
       .Formulas(0) = "fxCorte = 'Fecha de Corte para Cálculo de Intereses : " & Format(dtpCartaCorte.Value, "dd/mm/yyyy") & "'"
       .StoredProcParam(0) = Trim(txtCedula)
       .StoredProcParam(1) = Format(dtpCartaCorte.Value, "yyyy/mm/dd")
     
  
  Case "TERAVI" 'Tercer Aviso
     'Registro Historial y Expediente
     Call sbCBRRegTransac("10", txtCedula, txtOperacion, "Carta de Tercer Aviso...Corte:" & Format(dtpCartaCorte.Value, "dd/mm/yyyy"), CCur(txtSaldo), mCurIntc, mCurIntm)
     
     .ReportFileName = SIFGlobal.fxPathReportes("Cobro_CartaTercerAviso.rpt")
     .Formulas(0) = "MesProceso = '" & Format(vMes, "00") & "'"
     .SelectionFormula = "{SOCIOS.CEDULA} = '" & Trim(txtCedula) & "'"
     .SubreportToChange = "Fiadores"
     .SelectionFormula = "{VISTA_MOROSIDAD.CUOTA} >= 1 And {REG_CREDITOS.CEDULA} = '" & Trim(txtCedula) & "'"
  
       .SubreportToChange = "sbMora"
       .Formulas(0) = "fxCorte = 'Fecha de Corte para Cálculo de Intereses : " & Format(dtpCartaCorte.Value, "dd/mm/yyyy") & "'"
       .StoredProcParam(0) = Trim(txtCedula)
       .StoredProcParam(1) = Format(dtpCartaCorte.Value, "yyyy/mm/dd")
  
  Case "NOTMOV" 'Notificacion del Movimiento
     .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Notificacion.rpt")
     .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
  
  Case "REVER" 'Boleta Reversion
    fraFechas.Visible = True
  Case "ENVCBR" 'Casos Enviados a Cobro Judicial
    fraFechas.Visible = True
  Case "TRADEUD" 'Casos Traspaso - Deudor
    fraFechas.Visible = True
  Case "TRAFIA" 'Casos Traspaso - Fiadores
    fraFechas.Visible = True
  
  Case "TRAREV" 'Boleta de Reversion de Traspaso de Deudas
          .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
          .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
          .Formulas(2) = "subtitulo='BOLETA DE TRASLADO Y REVERSION DE DEUDAS'"
          .ReportFileName = SIFGlobal.fxPathReportes("Cobro_BoletaTraspasoReversion.rpt")
          .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
End Select

 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbConsulta_Mora()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Llena informacion de las cuotas de la operacion
'REFERENCIAS   : LLenaAbonos - (Carga Abonos Ordinarios y Extraordinarios)
'                LlenaCuotasMorosas - (Carga Cuotas en Mora)
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------


lblCuotas.Caption = cboTipoCuotas.Text
lblCuotas.Refresh

Call sbLswMovCrd


End Sub





Private Sub sbContacto_Info(pIdentificacion As String)
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Despliega datos del fiador Seleccionado, y lo activa para posible traspaso
'                de deudas
'REFERENCIAS   : fxProvincia - (Devuelve el número o descripcion de las provincias)
'                Telefonos   - (Carga los número telefonicos del fiador)
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select S.provincia,S.canton,S.distrito,S.direccion,AF_EMAIL, APTO  " _
       & ",isnull(P.Descripcion,'') as 'ProvinciaDesc', isnull(C.Descripcion,'') as 'CantonDesc', isnull(D.Descripcion,'') as 'DistritoDesc'" _
       & " from socios S left join Provincias P on S.provincia = P.provincia" _
       & " left join Cantones C on S.Provincia = C.Provincia and S.canton = C.canton" _
       & " left join Distritos D on S.Provincia = D.Provincia and S.canton = D.canton and S.distrito= D.Distrito" _
       & " where cedula = '" & pIdentificacion & "'"
   
   
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 txtDirFiadores = "PROVINCIA: " & rs!ProvinciaDesc & vbCrLf _
    & "CANTON: " & rs!CantonDesc & vbCrLf _
    & "DISTRITO: " & rs!DistritoDesc & vbCrLf _
    & "DIRECCION: " & IIf(IsNull(rs!direccion), "", rs!direccion)
    
 txtEmail.Text = rs!AF_Email & ""
 txtApartado.Text = rs!apto & ""
End If
rs.Close



With lswTelefonos
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Tipo", 1300
    .ColumnHeaders.Add , , "Número", 1200
    .ColumnHeaders.Add , , "Ext.", 1200
    .ColumnHeaders.Add , , "Contacto", 3200
    
    
    
    strSQL = "select * from telefonos where cedula = '" & Trim(pIdentificacion) & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    
    Do While Not rs.EOF
      Set itmX = .ListItems.Add(, , fxTipoTelefono(rs!Tipo))
          itmX.SubItems(1) = IIf(IsNull(rs!Numero), "", rs!Numero)
          itmX.SubItems(2) = Trim(IIf(IsNull(rs!Ext), "", rs!Ext))
          itmX.SubItems(3) = rs!contacto & ""
      rs.MoveNext
    Loop
    
    rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCobro_Fiadores_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

tcCF.Item(0).Selected = True

txtCF_Operacion.Text = ""
txtCF_Codigo.Text = ""

txtCF_Cedula.Text = ""
txtCF_Nombre.Text = ""

txtCF_Estado.Text = ""
txtCF_Inicia.Text = ""
txtCF_UltMov.Text = ""

txtCF_Cuota.Text = "0"
txtCF_Devuelto.Text = "0"
txtCF_Recaudo.Text = "0"
txtCF_Aplicado.Text = "0"

vPaso = True

With lswCF
 .ListItems.Clear
 
 
 strSQL = "exec spCbr_Cobro_Fiadores_List " & txtOperacion.Text
 Call OpenRecordSet(rs, strSQL)
 
 Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!ID_SOLICITUD)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!Cedula
        itmX.SubItems(3) = rs!Nombre
        
        itmX.SubItems(4) = Format(rs!Cuota, "Standard")
        itmX.SubItems(5) = Format(rs!RECAUDADO, "Standard")
        itmX.SubItems(6) = Format(rs!APLICADO, "Standard")
        itmX.SubItems(7) = Format(rs!DEVUELTO, "Standard")
        itmX.SubItems(8) = Format(rs!INICIO, "yyyy-mm-dd")
        itmX.SubItems(9) = Format(rs!ULTMOV, "yyyy-mm-dd")
        
        itmX.SubItems(10) = rs!ESTADO_DESC
  
    rs.MoveNext
 Loop
 rs.Close
 
End With

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub OperacionesGeneradas()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga las operaciones de los fiadores que se generaron por un traspaso de
'                deudas.
'REFERENCIAS   : Ninguna
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------
Dim colReferencias() As Long, i As Integer, iTotal As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

tcCF.Item(1).Selected = True

lblMontoActualDeudor.Caption = 0
lblPlazoActualDeudor.Caption = 0
lblInteresActualDeudor.Caption = 0
lblSaldoActualDeudor.Caption = 0
lblOperacionActualDeudor.Caption = 0

vPaso = True

With lswOperacionesGeneradas
 .ListItems.Clear
 
 
 strSQL = "exec spCbr_Operaciones_Trasladadas " & txtOperacion.Text
 Call OpenRecordSet(rs, strSQL)
 
 ReDim colReferencias(rs.RecordCount) As Long
 
 iTotal = rs.RecordCount
 i = 1
 
 Do While Not rs.EOF
  If Trim(rs!Cedula) = Trim(txtCedula) And UCase(Trim(rs!Codigo)) = UCase(Trim(txtCodigo)) Then
    lblMontoActualDeudor.Caption = Format(rs!Monto, "Standard")
    lblPlazoActualDeudor.Caption = rs!Plazo
    lblInteresActualDeudor.Caption = rs!Tasa_Actual
    lblSaldoActualDeudor.Caption = Format(rs!Saldo, "Standard")
    lblOperacionActualDeudor.Caption = rs!ID_SOLICITUD
  Else
    
    Set itmX = .ListItems.Add(, , rs!ID_SOLICITUD)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = rs!Cedula
        itmX.SubItems(3) = Format(rs!Monto, "Standard")
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!Cuota, "Standard")
        itmX.SubItems(6) = rs!Tasa
        itmX.SubItems(7) = rs!Plazo
        
        itmX.SubItems(8) = Format(rs!C_INTERESES, "Standard")
        itmX.SubItems(9) = Format(rs!C_Cargos + rs!C_IVA, "Standard")
        itmX.SubItems(10) = Format(rs!C_Polizas, "Standard")
        
        itmX.SubItems(11) = Format(rs!Nombre, "Standard")
      
     
     colReferencias(i) = rs!ID_SOLICITUD
     itmX.Tag = itmX.Index
     i = i + 1
  
  End If
    rs.MoveNext
 
 Loop
 rs.Close
 
End With

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswCF_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

txtCF_Operacion.Text = Item.Text
txtCF_Codigo.Text = Item.SubItems(1)

txtCF_Cedula.Text = Item.SubItems(2)
txtCF_Nombre.Text = Item.SubItems(3)

txtCF_Estado.Text = Item.SubItems(10)
txtCF_Inicia.Text = Item.SubItems(8)
txtCF_UltMov.Text = Item.SubItems(9)

txtCF_Cuota.Text = Item.SubItems(4)
txtCF_Recaudo.Text = Item.SubItems(5)
txtCF_Aplicado.Text = Item.SubItems(6)
txtCF_Devuelto.Text = Item.SubItems(7)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub lswContactos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Call sbContacto_Info(Item.Text)

End Sub




Private Sub lswOperacionesGeneradas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim lng As Long

txtTRAFD_MONTO = 0

 With lswOperacionesGeneradas
   For lng = 1 To .ListItems.Count
     If .ListItems.Item(lng).Checked Then
        'Saldo + Intereses Atrasados de los Fiadores Marcados + Cargos
        txtTRAFD_MONTO = CCur(txtTRAFD_MONTO) + CCur(.ListItems.Item(lng).SubItems(4)) + CCur(.ListItems.Item(lng).SubItems(8)) _
                       + CCur(.ListItems.Item(lng).SubItems(9)) + CCur(.ListItems.Item(lng).SubItems(10))

     End If
   Next lng
 End With
 
 txtTRAFD_MONTO = Format(txtTRAFD_MONTO, "Standard")
 txtTRAFD_Int = txtInteresActual
 txtTRAFD_Plazo = fxCBRPlazoRestante(txtOperacion)
End Sub




Private Sub sbReporteOpCtasPendientes()
Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobros"
    
    .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_OperacionCtaPendientes.rpt")
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "SubTitulo='DE LA OPERACIÓN AL CORTE'"
    .Formulas(2) = "Fecha='" & Format(dtpCalculoIntCorte.Value, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='CUOTAS ATRASADAS E INTERESES VENCIDOS'"
    
    .Formulas(4) = "fxDeuda=" & CCur(txtTotalMoraLegal.Text)
    .Formulas(5) = "fxIntVenc=" & CCur(txtInteresesCorrientes.Text) + CCur(txtInteresesMoratorios.Text)
    .Formulas(6) = "fxCargos=" & CCur(txtCargosRegistrados.Text) + CCur(txtPolizasAtrasadas)
    
    
    
    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & txtOperacion.Text
    
    .SubreportToChange = "sbMovimientos"
    .StoredProcParam(0) = txtOperacion.Text
    .StoredProcParam(1) = Format(dtpCalculoIntCorte.Value, "yyyy/mm/dd")
       
     .PrintReport
   
End With
Me.MousePointer = vbDefault

End Sub


Private Sub lswRepOp_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

lblRepOp.Caption = Item.Text
lblRepOp.Tag = Item.Key

If Item.Key = "CBRFIA" Then
    fraCbrFia.Visible = True
Else
    fraCbrFia.Visible = True
End If

End Sub


Private Sub tcCF_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Cobro a Fiadores
        Call sbCobro_Fiadores_List
        
  Case 1 'Traslados de Deuda
        Call OperacionesGeneradas

End Select

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Not vOperacion Then Exit Sub

Select Case Item.Index
  Case 0 'Estado
    
  Case 1 'Historial
    Call sbHistorial(txtOperacion)
  
  Case 2 'Gestiones
    Call vgCobro_SheetChanged(1, 1)
  
  Case 3 'Notificaciones
    Call sbAvisos(txtOperacion)
  
  Case 4 'En Cobro
       Call sbCobro_Fiadores_List
  
  Case 5 'Contacto
    Call sbContactos
  
  Case 6 'Mora
   Call sbConsulta_Mora
End Select


End Sub



Private Sub sbCbrArchivoEstudio()
Dim strSQL As String, rs As New ADODB.Recordset
Dim fn, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

fn = FreeFile

Open "C:\ArchivoEstudio.txt" For Output As #fn

strSQL = "OPERACION,CODIGO,CEDULA,NOMBRE,GARANTIA,MONTO,SALDO,Mor.INTC,Mor.INTM,Mor.AMORTIZA" _
       & ",Mor.FINANCIERA,Mor.LEGAL,PRI-DED,Mor.CUOTAS,COMITE,Ult.Mov.,ESTADO,Fec.APORTE,AHORROS" _
       & ",APORTES,LIQUIDEZ,Mor.Prs.FINANCIERA,PLANILLA,PLAZO,INTERES,Mor.Prs.LEGAL,ESTADO_LABORAL"
Print #fn, strSQL
Print #fn, ""


strSQL = "select R.id_solicitud,R.codigo,R.Cedula,S.nombre,R.garantia,R.montoapr" _
       & ",R.saldo,V.intc,V.intm,V.amortiza,(V.intc+V.intm+V.amortiza) as Financiera" _
       & ",(V.intc+V.intm+R.saldo) as Legal,R.prideduc,V.cuota,C.descripcion as comite" _
       & ",R.fecult,S.estadoactual,A.fecAporte,A.ahorro+A.capitaliza as Ahorros" _
       & ",A.aporte,isnull(P.porc_liquidez,0) as Liquidez,dbo.fxCBRMoraPersona(R.cedula,'F') as MoraPersona" _
       & ",R.ind_deduce_planilla,R.plazo,R.interesv,dbo.fxCBRMoraPersona(R.cedula,'L') as MoraPersonaLegal" _
       & ",isnull(S.estadolaboral,0) as EstadoLaboral" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join vista_morosidad V on R.id_Solicitud = V.id_solicitud" _
       & " inner join ahorro_consolidado A on S.cedula = A.cedula" _
       & " inner join comites C on R.id_comite = C.id_comite" _
       & " inner join catalogo X on R.codigo = X.codigo and X.retencion = 'N'" _
       & " left join pra_principal P on R.id_solicitud = P.id_solicitud"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 strSQL = ""
 For i = 0 To rs.Fields.Count - 1
    strSQL = strSQL & rs.Fields(i).Value & ","
 Next i
 Print #fn, strSQL
 rs.MoveNext
Loop
rs.Close

Close #fn

Me.MousePointer = vbDefault
MsgBox "Se Creó el Archivo : C:\ArchivoEstudio.txt", vbInformation

Exit Sub

vError:
Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vFecha As Long

If UCase(ButtonMenu.Key) = "ARCHIVO" Then
  Call sbCbrArchivoEstudio
  Exit Sub
End If

vFecha = GLOBALES.glngFechaCR
On Error Resume Next
vFecha = CLng(InputBox("Especifique la fecha de proceso " & vbCrLf _
        & "La fecha de Proceso Actual es : " & GLOBALES.glngFechaCR, "Reportes de Mora"))

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Cobro Administrativo y Judicial"

    .Connect = glogon.ConectRPT

    Select Case UCase(ButtonMenu.Key)
      Case "REPINGRESOS"
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_IngresosAMora.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='FECHA PROCESO : " & Format(vFecha, "####-##") & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{MOROSIDAD.FECHAP}=" & vFecha
        
      Case "REPEGRESOS"
      
      Case "REPABONOS"
      
      Case "REPPLANILLA"
        .ReportFileName = SIFGlobal.fxPathReportes("Cobro_PlanillaComparativa.rpt")
        .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(1) = "SubTitulo='FECHA PROCESO : " & Format(vFecha, "####-##") & "'"
        .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
        .SelectionFormula = "{vCbrPlanillaComparativa.proceso}=" & vFecha & " and ({vCbrPlanillaComparativa.Enviado} - {vCbrPlanillaComparativa.Recibido} > 10)"
            
    End Select
    .PrintReport
End With

End Sub

Private Sub tpMain_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Call sbTaskPanel_Accion(Item.Id)
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtNombre = fxNombre(txtCedula)
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtDescripcion = fxDescribeCodigo(txtCodigo)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(5)
End Sub

Private Sub txtOperacion_Change()
 Call CambiaDatos
End Sub

Private Sub CambiaDatos()
 vOperacion = False
 

 tcMain.Item(1).Enabled = False
 tcMain.Item(2).Enabled = False
 tcMain.Item(3).Enabled = False
 tcMain.Item(4).Enabled = False
 tcMain.Item(5).Enabled = False
 tcMain.Item(6).Enabled = False
 tcMain.Item(7).Enabled = False
 
 
 vTabs.Antiguedad = 0
 vTabs.direccion = 0
 vTabs.Fiadores = 0
 vTabs.OPGeneradas = 0
'Tab Estado de Cuenta y General
 txtCodigo = ""
 txtNombre = ""
 txtDescripcion = ""
 txtCedula = ""
 txtEstado.Text = ""
 txtEstadoMoroso.Text = ""
 txtPrimerDeduccion.Text = ""
 txtUltimoMovimiento.Text = ""
 txtGarantia.Text = ""
 txtDocumento.Text = ""

 txtMonto = ""
 txtPlazo = ""
 txtSaldo = ""
 txtAmortizado = ""
 txtInteresPorcentaje = ""
 txtCuota = ""
 txtInteresPagado = ""
 
 lblTasa.Caption = ""
 
 txtCbrDeuda.Text = ""
 txtCbrIntereses.Text = ""
 
 
 txtInteresActual = ""
 txtInteresesMoratorios = ""
 txtAmortizacionAtrasada = ""
 txtPolizasAtrasadas.Text = ""
 txtCargosRegistrados.Text = ""
 txtTotalMora = ""
 txtTotalMoraLegal = ""
 
 txtProceso.Text = ""
 txtOpex.Text = ""
 
 lswAbonos.ListItems.Clear
 lblCuotas.Caption = ""

'Tab Fiadores
 txtDirFiadores = ""

 
 lswOperacionesGeneradas.ListItems.Clear
 
 End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyPress(vbKeyReturn)
End Sub


Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call sbConsulta
End Sub


Private Sub sbAvisos(lngOperacion As Long)
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListViewItem

On Error GoTo vError

strSQL = "select * from cbr_avisos where id_solicitud = " _
       & lngOperacion & " order by fecha_aviso"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAvisos.ListItems.Add(, , rs!fecha_aviso)
      Select Case rs!tipo_aviso
        Case 1
            itmX.SubItems(1) = "Primer Aviso"
        Case 2
            itmX.SubItems(1) = "Segundo Aviso"
        Case Else
            itmX.SubItems(1) = "Otro Aviso"
      End Select
  rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbHistorial(pOperacion As Long)
Dim strSQL As String

tcMain.Item(1).Selected = True

strSQL = "select fecha,case when tipo = '01' then 'Traspaso de deudas'" _
       & " when tipo = '02' then 'Cobro Judicial'" _
       & " when tipo = '03' then 'Readecuaciones'" _
       & " when tipo = '04' then 'Arreglo de Pago'" _
       & " when tipo = '05' then 'Reversión Traspaso'" _
       & " when tipo = '06' then 'Reversión Cobro Judicial'" _
       & " when tipo = '07' then 'Registro de Incobrable'" _
       & " when tipo = '08' then 'Reversión de Incobrable'" _
       & " when tipo = '09' then 'Carta Primer Aviso'" _
       & " when tipo = '10' then 'Carta 2do y 3er. Aviso' end as Tipo" _
       & ",usuario,notas,saldo,int_cor + int_mor as intereses, isnull(cargos,0), isnull(poliza,0)" _
       & ",isnull(principal,0),isnull(tipo_documento,''),isnull(cod_transaccion,''),''" _
       & " from cbr_historial where id_solicitud = " & pOperacion & " order by fecha desc"
       
vPaso = True
    Call sbCargaGrid(vGrid, 12, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
vPaso = False
End Sub


Private Sub sbDeductoras_Load(pInstitucion As Long)
Dim strSQL As String

strSQL = "select COD_DEDUCTORA AS 'IdX', DESCRIPCION AS 'ItmX'" _
       & " From vAFI_Deductoras" _
       & " Where cod_institucion = " & pInstitucion

Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)

End Sub

Private Sub sbConsulta()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Actualizar la informacion de la ventana segun la operacion seleccionada
'REFERENCIAS   : sbMoraActiva - (Carga Datos de Mora Activa de la Operacion)
'                fxDescribeCodigo - (Devuelve la descripcion de el código del crédito)
'                sbBoletaAfiliacion - (Carga los datos personales)
'OBSERVACIONES : Ver Traspaso de Deudas
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, blnContinua As Boolean

On Error GoTo vError

If Not IsNumeric(txtOperacion.Text) Then Exit Sub

Me.MousePointer = vbHourglass

blnContinua = True

 strSQL = "select R.*,isnull(R.LiqTasa,0) as 'LiqTasaX', dbo.MyGetdate() as 'FechaServer'" _
        & ",C.Descripcion as 'LineaDesc',S.nombre, G.Descripcion as 'GarantiaDesc'" _
        & ",S.cod_Institucion, Ed.Cod_Institucion as 'DeductoraCod', Ed.Descripcion as 'DeductoraDesc'" _
        & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
        & " inner join Socios S on R.cedula = S.cedula" _
        & " inner join Instituciones Ed on isnull(isnull(R.cod_deductora,S.cod_deductora),S.cod_Institucion) = Ed.cod_Institucion" _
        & "  left join Crd_Garantia_Tipos G on R.Garantia = G.Garantia" _
        & " where R.id_solicitud = " & txtOperacion
Call OpenRecordSet(rs, strSQL)
 
 If rs.EOF And rs.BOF Then
  blnContinua = False
  vOperacion = False
    tcMain.Item(1).Enabled = False
    tcMain.Item(2).Enabled = False
    tcMain.Item(3).Enabled = False
    tcMain.Item(4).Enabled = False
    tcMain.Item(5).Enabled = False
    tcMain.Item(6).Enabled = False
    tcMain.Item(7).Enabled = False
  MsgBox "No se encontró número de solicitud...", vbInformation
 
 Else
    
    vOperacion = True
    tcMain.Item(1).Enabled = True
    tcMain.Item(2).Enabled = True
    tcMain.Item(3).Enabled = True
    tcMain.Item(4).Enabled = True
    tcMain.Item(5).Enabled = True
    tcMain.Item(6).Enabled = True
    tcMain.Item(7).Enabled = True
    
    tcMain.Item(0).Selected = True
   
    txtCodigo.Text = rs!Codigo
    txtDescripcion.Text = rs!LineaDesc
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    
   
    txtEstado.Text = fxDescribeEstado(IIf(IsNull(rs!Estado), "N", rs!Estado))
    
    dtpCalculoIntCorte.Value = rs!FechaServer
    
    
    Select Case UCase(IIf(IsNull(rs!Proceso), "N", rs!Proceso))
     Case "N"
      txtProceso.Text = "NORMAL"
     Case "T"
      txtProceso.Text = "TRASPASO DEUDAS"
     Case "J"
      txtProceso.Text = "COBRO JUDICIAL"
     Case Else
      txtProceso.Text = "NORMAL"
    End Select
    
    If IIf(IsNull(rs!opex), 0, rs!opex) = 1 Then
       txtOpex.Text = "SI"
    Else
       txtOpex.Text = "NO"
    End If
    
    vPaso = True
    If Not IsNull(rs!ind_deduce_planilla) Then
        chkDeducirPlanilla.Value = IIf((rs!ind_deduce_planilla = "S"), vbChecked, vbUnchecked)
    End If
    vPaso = False
    
    txtPrimerDeduccion.Text = Format(IIf(IsNull(rs!PriDeduc), "", rs!PriDeduc), "####-##")
    txtUltimoMovimiento.Text = Format(IIf(IsNull(rs!FecUlt), "", rs!FecUlt), "####-##")
    
    txtGarantia.Text = rs!GarantiaDesc & ""
    
    txtDocumento.Text = IIf(IsNull(rs!TDOCUMENTO), "", rs!TDOCUMENTO) & "-" & IIf(IsNull(rs!nDocumento), "", rs!nDocumento)
    txtMonto = Format(IIf(IsNull(rs!montoapr), "0", rs!montoapr), "Standard")
    txtPlazo = Format(IIf(IsNull(rs!Plazo), "1", rs!Plazo), "###0")
    txtSaldo = Format(IIf(IsNull(rs!Saldo), "0", rs!Saldo), "Standard")
    txtAmortizado = Format(IIf(IsNull(rs!Amortiza), "0", rs!Amortiza), "Standard")
    txtInteresPorcentaje = Format(IIf(IsNull(rs!Int), "", rs!Int), "Standard")
    txtCuota = Format(IIf(IsNull(rs!Cuota), "0", rs!Cuota), "Standard")
    txtInteresPagado = Format(IIf(IsNull(rs!interesc), "", rs!interesc), "Standard")
    
    txtInteresActual = Format(IIf(IsNull(rs!interesv), "0", rs!interesv), "Standard")

    If Not IsNull(rs!TBP_PuntosAdd) Then
      lblTasa.Caption = "Tasa (TBP + " & rs!TBP_PuntosAdd & ")"
      mTasaPts = rs!TBP_PuntosAdd
    Else
      lblTasa.Caption = "Tasa %"
      mTasaPts = -1000 'Default para Indicar que es tasa Fija
    End If
    
    If rs!LiqTasaX = 1 Then
      lblTasa.Caption = lblTasa.Caption & " + PtsLiq"
    End If
 
    mTasaLiq = rs!LiqTasaX

    'Carga Deductoras por Institucion
    Call sbDeductoras_Load(rs!cod_institucion)
    Call sbCboAsignaDato(cboDeductora, rs!DeductoraDesc, True, rs!DeductoraCod)

    cboDeductora.Tag = CStr(rs!DeductoraCod)

   Call sbMoraActiva
 End If
 

 rs.Close

 Me.MousePointer = vbDefault

Exit Sub


vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbMoraActiva(Optional pEstadoInicial As Boolean = True)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCbrCobroJudicialInteresesHoy " & txtOperacion.Text & ",'" & Format(dtpCalculoIntCorte.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
 
mCurIntc = rs!RegIntCor
mCurIntm = rs!RegIntMor
mcurPoliza = rs!Poliza
mcurCargos = rs!Cargos
mCurAmortAtrasada = rs!RegPrincipal

txtInteresesCorrientes.Text = Format(rs!RegIntCor, "Standard")
txtInteresesMoratorios.Text = Format(rs!RegIntMor, "Standard")
txtAmortizacionAtrasada.Text = Format(rs!RegPrincipal, "Standard")
txtCargosRegistrados.Text = Format(rs!Cargos, "Standard")
txtPolizasAtrasadas.Text = Format(rs!Poliza, "Standard")

txtTotalMora = Format(rs!RegIntCor + rs!RegIntMor + rs!Cargos + rs!Poliza + rs!RegPrincipal, "Standard")
txtTotalMoraLegal.Text = Format(rs!RegIntCor + rs!RegIntMor + rs!Cargos + rs!Poliza + CCur(txtSaldo.Text), "Standard")

txtEstadoMoroso.Text = rs!Antiguedad

If pEstadoInicial Then
    txtCbrIntereses.Text = Format(rs!RegIntCor + rs!RegIntMor, "Standard")
    txtCbrDeuda.Text = Format(rs!RegIntCor + rs!RegIntMor + mcurPoliza + mcurCargos + CCur(txtSaldo.Text), "Standard")
End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbLimpiaDatos()
Dim strSQL As String, rs As New ADODB.Recordset
 
 'tabs inactivos
 tcMain.Item(1).Enabled = False
 tcMain.Item(2).Enabled = False
 tcMain.Item(3).Enabled = False
 tcMain.Item(4).Enabled = False
 tcMain.Item(5).Enabled = False
 tcMain.Item(6).Enabled = False
 tcMain.Item(0).Selected = True



mCurIntc = 0
mCurIntm = 0
mCurAmortAtrasada = 0
mcurPoliza = 0
mcurCargos = 0

 lswAvisos.ListItems.Clear

 vTabs.Antiguedad = 0
 vTabs.direccion = 0
 vTabs.Fiadores = 0
 vTabs.OPGeneradas = 0
'Tab Estado de Cuenta y General
 txtOperacion = ""
 txtCodigo = ""
 txtNombre = ""
 txtDescripcion = ""
 txtCedula = ""
 txtEstado.Text = ""
 txtEstadoMoroso.Text = ""
 txtPrimerDeduccion.Text = ""
 txtUltimoMovimiento.Text = ""
 txtGarantia.Text = ""
 txtDocumento.Text = ""

 txtMonto = ""
 txtPlazo = ""
 txtSaldo = ""
 txtAmortizado = ""
 txtInteresPorcentaje = ""
 txtCuota = ""
 txtInteresPagado = ""
 
 txtInteresActual = ""
 txtInteresesMoratorios = ""
 txtAmortizacionAtrasada = ""
 txtCargosRegistrados.Text = ""
 txtPolizasAtrasadas.Text = ""
 
 txtTotalMora = ""
 txtTotalMoraLegal = ""
 txtProceso.Text = ""
 txtOpex.Text = ""

'Tab Cuotas
 lswAbonos.ListItems.Clear
 lblCuotas.Caption = ""
 
'Tab Reportes

 'Tab Reportes
 
 With lswRepOp.ListItems
   .Clear
   .Add , "ULTEC", "Boleta [Ultimo Estado]"
'   .Add , "ECBR", "Estado de Cuenta Cobro"
   .Add , "ETSBR", "Equitetado de Sobres"
   .Add , "PRIAVI", "Carta - Primer Aviso"
   .Add , "SEGAVI", "Carta - Segundo Aviso"
   .Add , "TERAVI", "Carta - Tercer Aviso"
   .Add , "NOTMOV", "Notificación de Movimiento Realizado"
   .Add , "TRAREV", "Boleta de Traspaso y Reversión de Deudas"
   .Add , "REVER", "Operaciones con Reversión (CJ-TD)"
   .Add , "ENVCBR", "Operaciones en Cobro Judicial"
   .Add , "TRACF", "Cobros a Fiadores"
   .Add , "TRADEUD", "Operaciones con Traslado de Deudas"
   .Add , "TRAFIA", "Operaciones de Fiadores con TD Aplicado"
   .Add , "CBRFIA", "Cobro a Fiadores: Retenciones"
 
 End With

End Sub


Private Sub sbContactos()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Carga lsw con los datos de los número de teléfonos de la persona
'REFERENCIAS   : fxEmpleadoPatrono - (Devuelve 1 = si es empleado y 0 si no)
'                fxNombre - (Devuelve el nombre de la persona)
'OBSERVACIONES : Ninguna
'-------------------------------------------------------------------------------------------
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

With lswContactos
    .ListItems.Clear
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Identificación", 1300
    .ColumnHeaders.Add , , "Nombre", 3000
    .ColumnHeaders.Add , , "Calidad", 1200, vbCenter
    .ColumnHeaders.Add , , "Registro", 1400
    
    
   Set itmX = .ListItems.Add(, , txtCedula.Text)
    itmX.SubItems(1) = txtNombre.Text
    itmX.SubItems(2) = "Deudor"
    itmX.SubItems(3) = "--"

strSQL = "select F.cedulaf,F.Calidad,S.nombre,E.descripcion as 'EstadoDesc'" _
       & ", case when F.Calidad = 'F' then 'Fiador' else 'CoDeudor' end as 'CalidadDesc'" _
       & " from fiadores F inner join Socios S on F.cedulaf = S.cedula" _
       & " inner join AFI_ESTADOS_PERSONA E on S.estadoActual = E.cod_Estado" _
       & " where F.id_solicitud = " & Trim(txtOperacion) & " and F.estado = 'A'"

Call OpenRecordSet(rs, strSQL, 0)
 
Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , (rs!cedulaf))
    itmX.SubItems(1) = rs!Nombre & ""
    itmX.SubItems(2) = rs!CalidadDesc
    itmX.SubItems(3) = rs!EstadoDesc & ""
    
 rs.MoveNext
Loop
rs.Close

End With

vPaso = False
Me.MousePointer = vbDefault

Exit Sub

vError:
    vPaso = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtTRAFD_Int_Change()

If CCur(IIf((txtTRAFD_Int = ""), 0, txtTRAFD_Int)) > 0 And CCur(IIf((txtTRAFD_Plazo = ""), 0, txtTRAFD_Plazo)) > 0 _
    And CCur(IIf((txtTRAFD_MONTO = ""), 0, txtTRAFD_MONTO)) > 0 Then
 txtTRAFD_Cuota = fxCalcula_Cuota(CCur(txtTRAFD_MONTO), CCur(txtTRAFD_Plazo), CCur(txtTRAFD_Int))
 txtTRAFD_Cuota = Format(txtTRAFD_Cuota, "Standard")
End If
End Sub

Private Sub txtTRAFD_Monto_Change()
Dim x As Integer
If CCur(IIf((txtTRAFD_Int = ""), 0, txtTRAFD_Int)) > 0 And CCur(IIf((txtTRAFD_Plazo = ""), 0, txtTRAFD_Plazo)) > 0 _
    And CCur(IIf((txtTRAFD_MONTO = ""), 0, txtTRAFD_MONTO)) > 0 Then
 txtTRAFD_Cuota = fxCalcula_Cuota(CCur(txtTRAFD_MONTO), CCur(txtTRAFD_Plazo), CCur(txtTRAFD_Int))
 txtTRAFD_Cuota = Format(txtTRAFD_Cuota, "Standard")
End If
End Sub

Private Sub txtTRAFD_Plazo_Change()
If CCur(IIf((txtTRAFD_Int = ""), 0, txtTRAFD_Int)) > 0 And CCur(IIf((txtTRAFD_Plazo = ""), 0, txtTRAFD_Plazo)) > 0 _
    And CCur(IIf((txtTRAFD_MONTO = ""), 0, txtTRAFD_MONTO)) > 0 Then
 txtTRAFD_Cuota = fxCalcula_Cuota(CCur(txtTRAFD_MONTO), CCur(txtTRAFD_Plazo), CCur(txtTRAFD_Int))
 txtTRAFD_Cuota = Format(txtTRAFD_Cuota, "Standard")
End If
End Sub


Private Sub vgCobro_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frmX As Form

If vPaso Then Exit Sub

Call sbFormsCall("frmCO_ControlSeguimiento", 0, , , False, Me)

For Each frmX In Forms
   If Trim(frmX.Name) = "frmCO_ControlSeguimiento" Then
        Exit For
   End If
Next

Call frmX.sbCargaDatos(txtCedula.Text)
        
End Sub

Private Sub vgCobro_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
vPaso = True

With vgCobro
    Select Case NewSheet
      Case 1 'Gestiones
        .ActiveSheet = NewSheet
        .Sheet = NewSheet
        .MaxRows = 0
       strSQL = "select S.*, isnull(G.descripcion,'') as 'Gestion'" _
              & "   , isnull(C.DESCRIPCION,'') as 'Causa'" _
              & "   , isnull(A.descripcion,'') as 'Arreglo'" _
              & " from CBR_Seguimiento S  left join cbr_gestiones G on S.cod_gestion = G.cod_gestion" _
              & "  left join CBR_CAUSAS_MOROSIDAD C on S.COD_CAUSA = C.COD_CAUSA" _
              & "  left join CBR_TIPOS_ARREGLOS A on S.COD_ARREGLO = A.COD_ARREGLO" _
              & " where cedula = '" & txtCedula.Text & "' order by S.cod_seg desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 2 To 11
            .Col = i
            Select Case i
              Case 2 'Fecha
                .Text = Format(rs!fecha, "dd/mm/yyyy")
              Case 3 'vencimiento
                .Text = Format(DateAdd("d", rs!tiempo_resolucion, rs!fecha), "dd/mm/yyyy")
              Case 4 'Gestión
                .Text = rs!Gestion
              Case 5 ' Detalle
                .Text = rs!Notas
                .RowHeight(.Row) = .MaxTextRowHeight(.Row)

              Case 6 ' Ejecutivo
                .Text = rs!Usuario
              Case 7 ' Monto
                .Text = Format(rs!Monto, "Standard")
              Case 8 ' Dias
                .Text = CStr(rs!tiempo_resolucion)
              Case 9  'Arrelgo de Pago
                .Text = rs!Arreglo
              Case 10 'Promesa de Pago
                .Text = Format(rs!Arreglo_Vence & "", "dd/mm/yyyy")
              Case 11 'Causa de Morosidad
                .Text = rs!Causa
                
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
      Case 2 'Oficiales
      
        .ActiveSheet = NewSheet
        .Sheet = NewSheet
        .MaxRows = 0
        strSQL = "select * from cbr_asignacion_h where cedula = '" & txtCedula.Text _
               & "' order by fecha_asignacion desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 1 To 5
            .Col = i
            Select Case i
              Case 1 'Fecha
                .Text = Format(rs!fecha_asignacion, "dd/mm/yyyy")
              Case 2 'Oficial
                .Text = UCase(rs!Usuario)
              Case 3 'Mantiene
                .Value = rs!mantener
              Case 4 ' Rebajo 2x
                .Value = rs!rebajo_doble
              Case 5 ' Mora
                .Value = rs!aplica_mora
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
    End Select
End With


Me.MousePointer = vbDefault
vPaso = False
Exit Sub

vError:
 vPaso = False
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim vTipoDoc As String, vDocumento As String

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 10
vTipoDoc = vGrid.Text
vGrid.Col = 11
vDocumento = vGrid.Text

If Len(vTipoDoc) > 0 And Len(vDocumento) > 0 Then
    Call sbImprimeRecibo(vDocumento, vTipoDoc)
End If
End Sub
