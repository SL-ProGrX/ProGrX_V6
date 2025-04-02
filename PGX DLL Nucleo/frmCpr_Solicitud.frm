VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCpr_Solicitud 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Solicitud de Compra"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   12375
      _Version        =   1572864
      _ExtentX        =   21828
      _ExtentY        =   11668
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
      ItemCount       =   4
      SelectedItem    =   2
      Item(0).Caption =   "Solicitud"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "Label3(1)"
      Item(0).Control(1)=   "txtDetalle"
      Item(0).Control(2)=   "Label3(0)"
      Item(0).Control(3)=   "txtSolicita"
      Item(0).Control(4)=   "cboUEN"
      Item(0).Control(5)=   "Label3(2)"
      Item(0).Control(6)=   "Label3(5)"
      Item(0).Control(7)=   "Label3(7)"
      Item(0).Control(8)=   "txtMonto"
      Item(0).Control(9)=   "gbInfoInterna"
      Item(0).Control(10)=   "chkI_Plan_Compras"
      Item(0).Control(11)=   "Label3(9)"
      Item(0).Control(12)=   "cboValidacion"
      Item(0).Control(13)=   "chkI_Presupuesto"
      Item(0).Control(14)=   "chkI_Contrato"
      Item(0).Control(15)=   "chkI_CompraMultiple"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbResume"
      Item(1).Control(1)=   "vGrid"
      Item(2).Caption =   "Cotiza y Valoración"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tcCotiza"
      Item(3).Caption =   "Seguimiento"
      Item(3).ControlCount=   18
      Item(3).Control(0)=   "Label3(16)"
      Item(3).Control(1)=   "Label3(17)"
      Item(3).Control(2)=   "Label3(18)"
      Item(3).Control(3)=   "Label3(19)"
      Item(3).Control(4)=   "Label3(20)"
      Item(3).Control(5)=   "txtS_Reg_Usuario"
      Item(3).Control(6)=   "txtS_Reg_Fecha"
      Item(3).Control(7)=   "FlatEdit4"
      Item(3).Control(8)=   "FlatEdit5"
      Item(3).Control(9)=   "FlatEdit6"
      Item(3).Control(10)=   "FlatEdit7"
      Item(3).Control(11)=   "FlatEdit8"
      Item(3).Control(12)=   "FlatEdit9"
      Item(3).Control(13)=   "FlatEdit10"
      Item(3).Control(14)=   "FlatEdit11"
      Item(3).Control(15)=   "ShortcutCaption3"
      Item(3).Control(16)=   "ShortcutCaption4"
      Item(3).Control(17)=   "lswBitacora"
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   3735
         Left            =   -70000
         TabIndex        =   84
         Top             =   2760
         Visible         =   0   'False
         Width           =   12375
         _Version        =   1572864
         _ExtentX        =   21828
         _ExtentY        =   6588
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
      Begin XtremeSuiteControls.TabControl tcCotiza 
         Height          =   6255
         Left            =   0
         TabIndex        =   67
         Top             =   360
         Width           =   12375
         _Version        =   1572864
         _ExtentX        =   21828
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
         ItemCount       =   3
         Item(0).Caption =   "Invitación"
         Item(0).ControlCount=   5
         Item(0).Control(0)=   "txtInvitacionFiltro"
         Item(0).Control(1)=   "lswProveedores"
         Item(0).Control(2)=   "ShortcutCaption2"
         Item(0).Control(3)=   "btnProveedores"
         Item(0).Control(4)=   "btnProveedoresAdd"
         Item(1).Caption =   "Cotización"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "vGrid_Cotiza"
         Item(2).Caption =   "Adjudicación"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "FlatEdit1"
         Item(2).Control(1)=   "Label3(6)"
         Item(2).Control(2)=   "btnRecomendacion"
         Item(2).Control(3)=   "gbAdjudica"
         Begin XtremeSuiteControls.ListView lswProveedores 
            Height          =   5055
            Left            =   0
            TabIndex        =   69
            Top             =   1080
            Width           =   12375
            _Version        =   1572864
            _ExtentX        =   21828
            _ExtentY        =   8916
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
         Begin XtremeSuiteControls.GroupBox gbAdjudica 
            Height          =   3255
            Left            =   -69880
            TabIndex        =   77
            Top             =   2520
            Visible         =   0   'False
            Width           =   12015
            _Version        =   1572864
            _ExtentX        =   21193
            _ExtentY        =   5741
            _StockProps     =   79
            Caption         =   "Adjudicar"
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
            Appearance      =   17
            BorderStyle     =   1
            Begin XtremeSuiteControls.ComboBox ComboBox1 
               Height          =   330
               Left            =   2880
               TabIndex        =   78
               Top             =   360
               Width           =   6255
               _Version        =   1572864
               _ExtentX        =   11033
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
            Begin XtremeSuiteControls.PushButton btnAdjudica 
               Height          =   375
               Index           =   0
               Left            =   7920
               TabIndex        =   80
               ToolTipText     =   "Notificar al Proveedor la Adjudicación"
               Top             =   2880
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Notificar"
               BackColor       =   16777215
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
               Appearance      =   17
               Picture         =   "frmCpr_Solicitud.frx":0000
               ImageAlignment  =   6
            End
            Begin XtremeSuiteControls.PushButton btnAdjudica 
               Height          =   375
               Index           =   1
               Left            =   6600
               TabIndex        =   81
               ToolTipText     =   "Adjudicar a Este Proveedor"
               Top             =   2880
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Adjudicar"
               BackColor       =   16777215
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
               Appearance      =   17
               Picture         =   "frmCpr_Solicitud.frx":0074
               ImageAlignment  =   6
            End
            Begin XtremeSuiteControls.FlatEdit txtA_Monto 
               Height          =   330
               Left            =   7080
               TabIndex        =   89
               ToolTipText     =   "Presione F4 para consultar"
               Top             =   840
               Width           =   2055
               _Version        =   1572864
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
               Text            =   "0.00"
               BackColor       =   16777215
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtVal_Estado 
               Height          =   330
               Left            =   2880
               TabIndex        =   90
               Top             =   1680
               Width           =   2055
               _Version        =   1572864
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
               Text            =   "Pendiente"
               BackColor       =   16777215
               Alignment       =   2
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtVal_Puntaje 
               Height          =   330
               Left            =   7080
               TabIndex        =   91
               Top             =   1680
               Width           =   2055
               _Version        =   1572864
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
               Text            =   "0"
               BackColor       =   16777215
               Alignment       =   2
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtA_Divisa 
               Height          =   330
               Left            =   4920
               TabIndex        =   92
               Top             =   840
               Width           =   855
               _Version        =   1572864
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
               Text            =   "COL"
               BackColor       =   16777215
               Alignment       =   2
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtA_Importe 
               Height          =   330
               Left            =   2880
               TabIndex        =   93
               ToolTipText     =   "Presione F4 para consultar"
               Top             =   840
               Width           =   2055
               _Version        =   1572864
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
               Text            =   "0.00"
               BackColor       =   16777215
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtA_TipoCambio 
               Height          =   330
               Left            =   5760
               TabIndex        =   94
               ToolTipText     =   "Presione F4 para consultar"
               Top             =   840
               Width           =   1335
               _Version        =   1572864
               _ExtentX        =   2355
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
               Text            =   "0.00"
               BackColor       =   16777215
               Alignment       =   1
               Locked          =   -1  'True
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.Label lblVal_Puntaje 
               Height          =   255
               Left            =   7080
               TabIndex        =   95
               Top             =   2040
               Width           =   2055
               _Version        =   1572864
               _ExtentX        =   3625
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Es el Mejor Puntaje!"
               BackColor       =   16777152
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Index           =   25
               Left            =   6000
               TabIndex        =   88
               Top             =   1680
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Puntaje"
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
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Index           =   24
               Left            =   1560
               TabIndex        =   87
               Top             =   1680
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Estado"
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
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Index           =   23
               Left            =   1560
               TabIndex        =   86
               Top             =   1320
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Valoración: "
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Transparent     =   -1  'True
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Index           =   22
               Left            =   1560
               TabIndex        =   85
               Top             =   840
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Monto: "
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
            End
            Begin XtremeSuiteControls.Label Label3 
               Height          =   255
               Index           =   21
               Left            =   1560
               TabIndex        =   79
               Top             =   360
               Width           =   1095
               _Version        =   1572864
               _ExtentX        =   1926
               _ExtentY        =   444
               _StockProps     =   79
               Caption         =   "Proveedor: "
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
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtInvitacionFiltro 
            Height          =   375
            Left            =   0
            TabIndex        =   68
            Top             =   360
            Width           =   12375
            _Version        =   1572864
            _ExtentX        =   21828
            _ExtentY        =   661
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
         Begin XtremeSuiteControls.PushButton btnProveedores 
            Height          =   375
            Left            =   0
            TabIndex        =   71
            ToolTipText     =   "Agregar Nuevos Proveedores"
            Top             =   720
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   661
            _StockProps     =   79
            BackColor       =   16777215
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
            Appearance      =   17
            Picture         =   "frmCpr_Solicitud.frx":079B
            ImageAlignment  =   6
         End
         Begin XtremeSuiteControls.PushButton btnProveedoresAdd 
            Height          =   375
            Left            =   600
            TabIndex        =   72
            ToolTipText     =   "Invitar a Proveedores Seleccionados"
            Top             =   720
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Invitar"
            BackColor       =   16777215
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
            Appearance      =   17
            Picture         =   "frmCpr_Solicitud.frx":0EBB
            ImageAlignment  =   6
         End
         Begin FPSpreadADO.fpSpread vGrid_Cotiza 
            Height          =   5655
            Left            =   -69880
            TabIndex        =   73
            Top             =   480
            Visible         =   0   'False
            Width           =   11895
            _Version        =   524288
            _ExtentX        =   20981
            _ExtentY        =   9975
            _StockProps     =   64
            ArrowsExitEditMode=   -1  'True
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
            MaxCols         =   7
            ScrollBars      =   2
            SpreadDesigner  =   "frmCpr_Solicitud.frx":14DF
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   1035
            Left            =   -68320
            TabIndex        =   74
            Top             =   600
            Visible         =   0   'False
            Width           =   9495
            _Version        =   1572864
            _ExtentX        =   16748
            _ExtentY        =   1826
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
         Begin XtremeSuiteControls.PushButton btnRecomendacion 
            Height          =   375
            Left            =   -58720
            TabIndex        =   76
            ToolTipText     =   "Guardar"
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            BackColor       =   16777215
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
            Appearance      =   17
            Picture         =   "frmCpr_Solicitud.frx":1C84
            ImageAlignment  =   6
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   6
            Left            =   -69880
            TabIndex        =   75
            Top             =   600
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Recomendación: "
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
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   375
            Left            =   0
            TabIndex        =   70
            Top             =   720
            Width           =   12495
            _Version        =   1572864
            _ExtentX        =   22040
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Seleccione los proveedores a invitar"
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit11 
         Height          =   330
         Left            =   -63880
         TabIndex        =   65
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   2280
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit9 
         Height          =   330
         Left            =   -63880
         TabIndex        =   63
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit7 
         Height          =   330
         Left            =   -63880
         TabIndex        =   61
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit5 
         Height          =   330
         Left            =   -63880
         TabIndex        =   59
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkI_Plan_Compras 
         Height          =   255
         Left            =   -68560
         TabIndex        =   35
         Top             =   2880
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Se encuentra en el Plan de Compras?"
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
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.GroupBox gbInfoInterna 
         Height          =   2295
         Left            =   -69880
         TabIndex        =   31
         Top             =   4200
         Visible         =   0   'False
         Width           =   12135
         _Version        =   1572864
         _ExtentX        =   21405
         _ExtentY        =   4048
         _StockProps     =   79
         Caption         =   "Información Interna"
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.ComboBox cboI_FormaPago 
            Height          =   330
            Left            =   1560
            TabIndex        =   43
            Top             =   1320
            Width           =   4455
            _Version        =   1572864
            _ExtentX        =   7858
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
         Begin XtremeSuiteControls.FlatEdit txtI_Sitio_Entrega 
            Height          =   330
            Left            =   1560
            TabIndex        =   44
            ToolTipText     =   "Presione F4 para consultar"
            Top             =   600
            Width           =   4455
            _Version        =   1572864
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
            BackColor       =   16777215
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtI_CorreoFactura 
            Height          =   330
            Left            =   1560
            TabIndex        =   45
            ToolTipText     =   "Presione F4 para consultar"
            Top             =   960
            Width           =   4455
            _Version        =   1572864
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
            BackColor       =   16777215
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   1635
            Left            =   6120
            TabIndex        =   47
            Top             =   600
            Width           =   5895
            _Version        =   1572864
            _ExtentX        =   10398
            _ExtentY        =   2884
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   330
            Left            =   1560
            TabIndex        =   49
            Top             =   1680
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
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
            Text            =   "30"
            BackColor       =   16777215
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   15
            Left            =   2880
            TabIndex        =   50
            Top             =   1680
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "en días hábiles"
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
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   48
            Top             =   1680
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Plazo de Entrega: "
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   13
            Left            =   10440
            TabIndex        =   46
            Top             =   360
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observación: "
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
            Alignment       =   1
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   42
            Top             =   1320
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Forma de Pago: "
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
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Correo Factura: "
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
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   40
            Top             =   600
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sitio de Entrega: "
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
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   12135
            _Version        =   1572864
            _ExtentX        =   21405
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Información de Uso Interno"
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
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   1035
         Left            =   -68560
         TabIndex        =   16
         Top             =   1200
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   1826
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
      Begin XtremeSuiteControls.ComboBox cboUEN 
         Height          =   330
         Left            =   -64120
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1572864
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.FlatEdit txtSolicita 
         Height          =   330
         Left            =   -68560
         TabIndex        =   19
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   720
         Visible         =   0   'False
         Width           =   4455
         _Version        =   1572864
         _ExtentX        =   7858
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbResume 
         Height          =   975
         Left            =   -69880
         TabIndex        =   23
         Top             =   5400
         Visible         =   0   'False
         Width           =   11895
         _Version        =   1572864
         _ExtentX        =   20976
         _ExtentY        =   1714
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtSubTotal 
            Height          =   312
            Left            =   9840
            TabIndex        =   24
            Top             =   480
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label lblLineas 
            Height          =   492
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Width           =   2772
            _Version        =   1572864
            _ExtentX        =   4890
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Lineas 0"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   492
            Left            =   8880
            TabIndex        =   26
            Top             =   360
            Width           =   852
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Total"
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
         End
         Begin XtremeSuiteControls.Label lblCantidad 
            Height          =   492
            Left            =   3000
            TabIndex        =   25
            Top             =   360
            Width           =   2892
            _Version        =   1572864
            _ExtentX        =   5101
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cantidad 0"
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
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4815
         Left            =   -69880
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   12135
         _Version        =   524288
         _ExtentX        =   21405
         _ExtentY        =   8493
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
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
         MaxCols         =   487
         ScrollBars      =   2
         SpreadDesigner  =   "frmCpr_Solicitud.frx":23B5
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   330
         Left            =   -68560
         TabIndex        =   30
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
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
         Text            =   "0.00"
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkI_Presupuesto 
         Height          =   255
         Left            =   -68560
         TabIndex        =   36
         Top             =   3240
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Se encuentra Presupuestado?"
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
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkI_Contrato 
         Height          =   255
         Left            =   -68560
         TabIndex        =   37
         Top             =   3600
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Requiere Contrato?"
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
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cboValidacion 
         Height          =   330
         Left            =   -64120
         TabIndex        =   39
         Top             =   2400
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1572864
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.FlatEdit txtS_Reg_Usuario 
         Height          =   330
         Left            =   -67000
         TabIndex        =   56
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtS_Reg_Fecha 
         Height          =   330
         Left            =   -63880
         TabIndex        =   57
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit4 
         Height          =   330
         Left            =   -67000
         TabIndex        =   58
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit6 
         Height          =   330
         Left            =   -67000
         TabIndex        =   60
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit8 
         Height          =   330
         Left            =   -67000
         TabIndex        =   62
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1920
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit10 
         Height          =   330
         Left            =   -67000
         TabIndex        =   64
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   2280
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkI_CompraMultiple 
         Height          =   255
         Left            =   -64000
         TabIndex        =   96
         Top             =   2880
         Visible         =   0   'False
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Compra Multiple UENs ?"
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
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   330
         Left            =   -63880
         TabIndex        =   83
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Fecha"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   330
         Left            =   -67000
         TabIndex        =   82
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Usuario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   20
         Left            =   -68560
         TabIndex        =   55
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Adjudicación: "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   19
         Left            =   -68560
         TabIndex        =   54
         Top             =   1920
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Presupuesto: "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   18
         Left            =   -68560
         TabIndex        =   53
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Autorización: "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   17
         Left            =   -68560
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Modificación: "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   16
         Left            =   -68560
         TabIndex        =   51
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registro: "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   9
         Left            =   -65920
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Plan de Validación: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   7
         Left            =   -69760
         TabIndex        =   29
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Monto: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   5
         Left            =   -64120
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "UEN: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   -68560
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Persona: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   -69760
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Solicita: "
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
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   -69760
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Detalle: "
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
      End
   End
   Begin XtremeSuiteControls.PushButton btnAutoriza 
      Height          =   330
      Left            =   5040
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Autorización"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":2A86
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":31AD
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":37DF
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":3DDA
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":437E
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":4AAF
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Reporte"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":51AF
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   6
      Left            =   3240
      TabIndex        =   7
      ToolTipText     =   "Consultas"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCpr_Solicitud.frx":58B6
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   0
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Adjuntos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCpr_Solicitud.frx":5FB6
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8040
      TabIndex        =   9
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.DateTimePicker dtpFecha 
      Height          =   315
      Left            =   10440
      TabIndex        =   10
      Top             =   720
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3408
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
      CustomFormat    =   "dd/MM/yyyy hh:mm:ss"
      Format          =   3
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   1440
      TabIndex        =   11
      ToolTipText     =   "No. de Solicitud de Compra"
      Top             =   720
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3831
      _ExtentY        =   762
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
      Text            =   "0000000"
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   435
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "Estado de la Solicitud"
      Top             =   720
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3831
      _ExtentY        =   762
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
      Text            =   "Solicitado"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOrden 
      Height          =   330
      Left            =   10320
      TabIndex        =   33
      ToolTipText     =   "Presione F4 para consultar"
      Top             =   0
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtPresupuestoEstado 
      Height          =   435
      Left            =   5760
      TabIndex        =   66
      ToolTipText     =   "Estado del Presupuesto"
      Top             =   720
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3831
      _ExtentY        =   762
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
      Text            =   "Revisión"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   8
      Left            =   9360
      TabIndex        =   32
      Top             =   0
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Orden"
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Solicitud "
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
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   3
      Left            =   8880
      TabIndex        =   13
      Top             =   720
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fecha: "
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmCpr_Solicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdjuntos_Click()

If txtCodigo.Text <> "" Then
 gGA.Modulo = "CPR"
 gGA.Llave_01 = txtCodigo.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End If
End Sub



Private Sub btnAutoriza_Click()
Call sbFormsCall("frmCpr_Solicitud_Autoriza", vbModal, , , , Me, True)
End Sub

Private Sub Form_Load()

tcMain.Item(0).Selected = True

dtpFecha.Value = Now

With lswProveedores.ColumnHeaders
    .Clear
    .Add , , "Proveedor Id", 1200
    .Add , , "Identificación", 1500
    .Add , , "Nombre", 4500
End With
lswProveedores.Checkboxes = True


With lswBitacora.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 2100
    .Add , , "Detalle", 4500
End With

End Sub

Private Sub vGrid_Cotiza_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

Select Case Col
    Case 3
        MsgBox "Correo de Solicitud de Cotización enviada al Proveedor!", vbInformation
    Case 4
        Call sbFormsCall("frmCpr_Solicitud_Cotizaciones", vbModal, , , False, Me, True)
    Case 5
        Call sbFormsCall("frmCpr_Solicitud_Valoracion", vbModal, , , False, Me, True)
        
End Select

End Sub
