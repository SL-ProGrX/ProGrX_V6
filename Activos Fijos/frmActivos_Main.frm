VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmActivos_Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Activos Fijos"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   10485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.PushButton btnNumPlaca 
      Height          =   315
      Left            =   5400
      TabIndex        =   60
      ToolTipText     =   "Consecutivo Automático"
      Top             =   615
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   556
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmActivos_Main.frx":0000
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   5880
      TabIndex        =   59
      Top             =   645
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   10455
      _Version        =   1441793
      _ExtentX        =   18436
      _ExtentY        =   11874
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
      ItemCount       =   6
      Item(0).Caption =   "General"
      Item(0).ImageIndex=   0
      Item(0).ControlCount=   35
      Item(0).Control(0)=   "txtUDAnio"
      Item(0).Control(1)=   "txtUDProducidas"
      Item(0).Control(2)=   "cboTipo"
      Item(0).Control(3)=   "txtDescripcion"
      Item(0).Control(4)=   "txtValorHistorico"
      Item(0).Control(5)=   "txtValorRescate"
      Item(0).Control(6)=   "cbo"
      Item(0).Control(7)=   "txtVU"
      Item(0).Control(8)=   "cboVU"
      Item(0).Control(9)=   "dtpAdquisicion"
      Item(0).Control(10)=   "dtpInstalacion"
      Item(0).Control(11)=   "cboResponsable"
      Item(0).Control(12)=   "cboDepartamento"
      Item(0).Control(13)=   "cboSeccion"
      Item(0).Control(14)=   "txtNotas"
      Item(0).Control(15)=   "txtProveedor"
      Item(0).Control(16)=   "txtDocCompra"
      Item(0).Control(17)=   "Label5(0)"
      Item(0).Control(18)=   "Label5(1)"
      Item(0).Control(19)=   "Label5(2)"
      Item(0).Control(20)=   "Label5(3)"
      Item(0).Control(21)=   "Label5(4)"
      Item(0).Control(22)=   "Label5(5)"
      Item(0).Control(23)=   "Label5(6)"
      Item(0).Control(24)=   "Label5(7)"
      Item(0).Control(25)=   "Label5(8)"
      Item(0).Control(26)=   "Label5(9)"
      Item(0).Control(27)=   "Label5(10)"
      Item(0).Control(28)=   "Label5(11)"
      Item(0).Control(29)=   "Label5(12)"
      Item(0).Control(30)=   "Label5(13)"
      Item(0).Control(31)=   "Label5(14)"
      Item(0).Control(32)=   "Label5(15)"
      Item(0).Control(33)=   "cboLocaliza"
      Item(0).Control(34)=   "Label5(22)"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "txtMarca"
      Item(1).Control(1)=   "txtSerie"
      Item(1).Control(2)=   "txtModelo"
      Item(1).Control(3)=   "lsw"
      Item(1).Control(4)=   "txtOtrasSenas"
      Item(1).Control(5)=   "Label5(16)"
      Item(1).Control(6)=   "Label5(17)"
      Item(1).Control(7)=   "Label5(18)"
      Item(1).Control(8)=   "Label5(19)"
      Item(1).Control(9)=   "Label5(20)"
      Item(2).Caption =   "Modificaciones"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lswMod"
      Item(2).Control(1)=   "scSubTitulos(0)"
      Item(3).Caption =   "Histórico"
      Item(3).ControlCount=   3
      Item(3).Control(0)=   "vGrid"
      Item(3).Control(1)=   "cboHistorico"
      Item(3).Control(2)=   "scSubTitulos(1)"
      Item(4).Caption =   "Composición"
      Item(4).ControlCount=   2
      Item(4).Control(0)=   "lswCompo"
      Item(4).Control(1)=   "scSubTitulos(2)"
      Item(5).Caption =   "Pólizas"
      Item(5).ControlCount=   2
      Item(5).Control(0)=   "lswPolizas"
      Item(5).Control(1)=   "scSubTitulos(3)"
      Begin XtremeSuiteControls.ListView lswPolizas 
         Height          =   5892
         Left            =   -69880
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ListView lswCompo 
         Height          =   5892
         Left            =   -69880
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ListView lswMod 
         Height          =   5892
         Left            =   -69880
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3372
         Left            =   -69880
         TabIndex        =   26
         Top             =   3000
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
         _ExtentY        =   5948
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5895
         Left            =   -69880
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
         _ExtentY        =   10398
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
         MaxCols         =   12
         SpreadDesigner  =   "frmActivos_Main.frx":0700
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   2760
         TabIndex        =   4
         Top             =   600
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   2760
         TabIndex        =   5
         Top             =   960
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.FlatEdit txtValorHistorico 
         Height          =   312
         Left            =   2760
         TabIndex        =   6
         Top             =   1440
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
      Begin XtremeSuiteControls.FlatEdit txtValorRescate 
         Height          =   312
         Left            =   6960
         TabIndex        =   7
         Top             =   1440
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
      Begin XtremeSuiteControls.DateTimePicker dtpAdquisicion 
         Height          =   312
         Left            =   3480
         TabIndex        =   8
         Top             =   1800
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
      Begin XtremeSuiteControls.DateTimePicker dtpInstalacion 
         Height          =   312
         Left            =   7680
         TabIndex        =   9
         Top             =   1800
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   6960
         TabIndex        =   10
         Top             =   2160
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboVU 
         Height          =   312
         Left            =   3480
         TabIndex        =   11
         Top             =   2160
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtVU 
         Height          =   330
         Left            =   2760
         TabIndex        =   12
         Top             =   2160
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtUDAnio 
         Height          =   312
         Left            =   2760
         TabIndex        =   13
         Top             =   2520
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
      Begin XtremeSuiteControls.FlatEdit txtUDProducidas 
         Height          =   312
         Left            =   6960
         TabIndex        =   14
         Top             =   2520
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   912
         Left            =   2760
         TabIndex        =   15
         Top             =   3120
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
      Begin XtremeSuiteControls.ComboBox cboDepartamento 
         Height          =   312
         Left            =   2760
         TabIndex        =   16
         Top             =   4320
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.ComboBox cboSeccion 
         Height          =   312
         Left            =   2760
         TabIndex        =   17
         Top             =   4680
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.ComboBox cboResponsable 
         Height          =   312
         Left            =   2760
         TabIndex        =   18
         Top             =   5040
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   312
         Left            =   2760
         TabIndex        =   19
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   6000
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
      Begin XtremeSuiteControls.FlatEdit txtDocCompra 
         Height          =   312
         Left            =   2760
         TabIndex        =   20
         Top             =   6360
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
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
      Begin XtremeSuiteControls.FlatEdit txtModelo 
         Height          =   312
         Left            =   -69880
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
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
      Begin XtremeSuiteControls.FlatEdit txtSerie 
         Height          =   312
         Left            =   -66520
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
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
      Begin XtremeSuiteControls.FlatEdit txtMarca 
         Height          =   312
         Left            =   -63160
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
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
      Begin XtremeSuiteControls.FlatEdit txtOtrasSenas 
         Height          =   912
         Left            =   -69880
         TabIndex        =   24
         Top             =   1320
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
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
      Begin XtremeSuiteControls.ComboBox cboHistorico 
         Height          =   312
         Left            =   -61720
         TabIndex        =   25
         Top             =   504
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cboLocaliza 
         Height          =   315
         Left            =   2760
         TabIndex        =   57
         Top             =   5400
         Width           =   6255
         _Version        =   1441793
         _ExtentX        =   11033
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   22
         Left            =   840
         TabIndex        =   58
         Top             =   5400
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Localización"
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
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   372
         Index           =   3
         Left            =   -69880
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Listado de Polizas (Seguros) Asignados al Activo"
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
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   372
         Index           =   2
         Left            =   -69880
         TabIndex        =   53
         Top             =   480
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Indica las depreciaciones del Activo y sus componentes o mejoras realizadas"
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
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   372
         Index           =   1
         Left            =   -69880
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Historico de Cierres de Depreciaciones Registradas al Activo"
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
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   51
         Top             =   480
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1441793
         _ExtentX        =   17801
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Listado de Modificaciones del Activo"
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
         VisualTheme     =   3
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   20
         Left            =   -69880
         TabIndex        =   50
         Top             =   2640
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1441793
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Responsables (Histórico)..:"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   19
         Left            =   -69880
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Otras Señas"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   18
         Left            =   -63160
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Marca"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   17
         Left            =   -66520
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Serie"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   16
         Left            =   -69880
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Modelo"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   15
         Left            =   5400
         TabIndex        =   45
         Top             =   2520
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ud. a Producir"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   14
         Left            =   5400
         TabIndex        =   44
         Top             =   2160
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Depreciación"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   13
         Left            =   5400
         TabIndex        =   43
         Top             =   1800
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Instalación"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   12
         Left            =   5400
         TabIndex        =   42
         Top             =   1440
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor de rescate"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   11
         Left            =   840
         TabIndex        =   41
         Top             =   6360
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Doc.Compra"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   10
         Left            =   840
         TabIndex        =   40
         Top             =   6000
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Proveedor"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   9
         Left            =   840
         TabIndex        =   39
         Top             =   5040
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Responsable"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   8
         Left            =   840
         TabIndex        =   38
         Top             =   4680
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Sección"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   7
         Left            =   840
         TabIndex        =   37
         Top             =   4320
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Departamento"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   6
         Left            =   840
         TabIndex        =   36
         Top             =   3120
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Notas"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   5
         Left            =   840
         TabIndex        =   35
         Top             =   2520
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ud. x Año"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   4
         Left            =   840
         TabIndex        =   34
         Top             =   2160
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Vida útil"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   3
         Left            =   840
         TabIndex        =   33
         Top             =   1800
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Adquisición"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   2
         Left            =   840
         TabIndex        =   32
         Top             =   1440
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Valor Histórico"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   31
         Top             =   960
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo de Activo"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   252
         Index           =   0
         Left            =   840
         TabIndex        =   30
         Top             =   600
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
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
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   7932
      Width           =   10488
      _ExtentX        =   18494
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Usuario que Registro Activo"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   3422
            MinWidth        =   3422
            Object.ToolTipText     =   "Fecha de Registro Real"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1482
            MinWidth        =   1482
            Object.ToolTipText     =   "Ultimo Periodo Depreciado"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   2541
            MinWidth        =   2541
            Object.ToolTipText     =   "Depreciación Acumulada"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Bevel           =   0
            Object.Width           =   2187
            MinWidth        =   2187
            Object.ToolTipText     =   "Depreciación del Mes"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Main.frx":0FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Main.frx":1106
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Main.frx":1648
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Main.frx":1B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmActivos_Main.frx":1CAC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   435
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   767
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1440
      TabIndex        =   61
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1441793
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
      Picture         =   "frmActivos_Main.frx":1E06
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2520
      TabIndex        =   62
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Main.frx":2438
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   2880
      TabIndex        =   63
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Main.frx":2A33
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3480
      TabIndex        =   64
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Main.frx":2FD7
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   3840
      TabIndex        =   65
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Main.frx":3708
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   4320
      TabIndex        =   66
      ToolTipText     =   "Reporte"
      Top             =   0
      Width           =   375
      _Version        =   1441793
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
      Picture         =   "frmActivos_Main.frx":3E08
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtPlacaAlterna 
      Height          =   435
      Left            =   3360
      TabIndex        =   67
      ToolTipText     =   "Número de Placa Alterna"
      Top             =   600
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
      _ExtentY        =   767
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   255
      Left            =   7440
      TabIndex        =   56
      Top             =   600
      Width           =   2775
      _Version        =   1441793
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "xx"
      BackColor       =   -2147483633
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
   Begin XtremeSuiteControls.Label Label5 
      Height          =   375
      Index           =   21
      Left            =   120
      TabIndex        =   55
      Top             =   600
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "No. Placa"
      BackColor       =   -2147483633
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
End
Attribute VB_Name = "frmActivos_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vPaso As Boolean, vScroll As Boolean
Dim mFechaUltCierre As Date, mPermiteRegistro As Integer




Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String


Select Case Index
    Case 0 'NUEVO
        vEdita = False
        Call sbLimpiaPantalla
        txtCodigo.SetFocus
        txtCodigo.Text = fxPlaca_ID
        
        Call sbBarra_Accion("Editar")
        
    Case 1 'MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbBarra_Accion("Editar")
      
    Case 2 'BORRAR"
      Call sbBorrar
      Call sbBarra_Accion("Nuevo")
    
    Case 3 'GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    
    Case 4 'DESHACER"
      Call sbBarra_Accion("Editar")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbBarra_Accion("Nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    
    Case 5 'REPORTES
   
End Select

End Sub




Private Sub btnNumPlaca_Click()

If Not vEdita Then
      txtCodigo.Text = fxPlaca_ID
End If

End Sub

Private Sub cbo_Click()
If Mid(cbo.Text, 1, 1) = "U" Then
  txtUDProducidas.Locked = False
Else
  txtUDProducidas.Locked = True
End If

txtUDProducidas.ForeColor = IIf(txtUDProducidas.Locked, vbBlack, vbBlue)

txtUDAnio.Locked = txtUDProducidas.Locked
txtUDAnio.ForeColor = txtUDProducidas.ForeColor

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub


Private Sub cboDepartamento_Click()
Dim strSQL As String

If vPaso Then Exit Sub

If cboDepartamento.ListCount = 0 Then
    cboSeccion.Clear
    Exit Sub
End If

vPaso = True
    strSQL = "select rtrim(cod_Seccion) as 'Idx', rtrim(descripcion) as 'ItmX' from Activos_Secciones" _
           & " Where cod_departamento = '" & cboDepartamento.ItemData(cboDepartamento.ListIndex) & "' order by cod_Seccion"
    Call sbCbo_Llena_New(cboSeccion, strSQL, False, True)
vPaso = False

Call cboSeccion_Click


End Sub

Private Sub cboHistorico_Click()
If vPaso Then Exit Sub

Dim strSQL As String

     strSQL = "exec spActivos_HistoricoConsolidado '" & txtCodigo.Text & "','" & Mid(cboHistorico.Text, 1, 1) & "'"
     Call sbCargaGrid(vGrid, 12, strSQL, True)
     
End Sub

Private Sub cboSeccion_Click()
Dim strSQL As String

If vPaso Then Exit Sub

If cboSeccion.ListCount = 0 Then
   cboResponsable.Clear
   Exit Sub
End If


vPaso = True
    strSQL = "select rtrim(Identificacion) as 'IdX', rtrim(Nombre) as 'ItmX' from Activos_Personas" _
           & " Where cod_departamento = '" & cboDepartamento.ItemData(cboDepartamento.ListIndex) _
           & "' and cod_Seccion = '" & cboSeccion.ItemData(cboSeccion.ListIndex) & "' order by identificacion"
    Call sbCbo_Llena_New(cboResponsable, strSQL, False, True)
vPaso = False

End Sub

Private Sub cboTipo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub


'Llenar con valores por defecto, del tipo de activo
strSQL = "select * from Activos_tipo_activo where tipo_activo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
'  cbo.Text = fxActivos_MetodoDepreciacion(rs!met_depreciacion)
  Call sbCboAsignaDato(cbo, fxActivos_MetodoDepreciacion(rs!met_depreciacion), True, rs!met_depreciacion)
  
  txtVU.Text = rs!Vida_Util
  If rs!tipo_vida_util = "A" Then
    cboVU.Text = "Años"
  Else
    cboVU.Text = "Meses"
  End If
End If
rs.Close

End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValorHistorico.SetFocus
End Sub

Private Sub cboVU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub


Private Sub dtpAdquisicion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInstalacion.SetFocus
End Sub

Private Sub dtpInstalacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 num_placa from Activos_Principal"

    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where num_placa > '" & txtCodigo.Text & "' order by num_placa asc"
    Else
       strSQL = strSQL & " where num_placa < '" & txtCodigo.Text & "' order by num_placa desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!num_placa
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

End Sub

Private Sub Form_Activate()
vModulo = 36

End Sub

Private Function fxPermiteRegistro() As Integer

fxPermiteRegistro = 0

With glogon
 .strSQL = "select isnull(REGISTRO_PERIODO_CERRADO,0) as 'Permite' from ACTIVOS_PARAMETROS"
 Call OpenRecordSet(.Recordset, .strSQL)
 
 fxPermiteRegistro = .Recordset!Permite
 
 .Recordset.Close
End With

End Function

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError
  
vModulo = 36


With lsw.ColumnHeaders
  .Clear
  .Add , , "Identificación", 1800
  .Add , , "Nombre", 3400
  .Add , , "Registro", 2500
End With

With lswMod.ColumnHeaders
  .Clear
  .Add , , "[ID]", 800
  .Add , , "Tipo", 1500
  .Add , , "Fecha", 1800, vbCenter
  .Add , , "Monto", 1500, vbRightJustify
  .Add , , "Justificación", 3500
  .Add , , "Descripción", 3500
End With


With lswCompo.ColumnHeaders
  .Clear
  .Add , , "Tipo", 1800
  .Add , , "Placa", 1800
  .Add , , "Periodo", 1000
  .Add , , "Dep.Acumulada", 1800, vbRightJustify
  .Add , , "Dep.Mensual", 1800, vbRightJustify
  .Add , , "Adquisición", 1800, vbCenter
  .Add , , "Descripción", 2500
  .Add , , "Fecha Registro", 2500, vbCenter
End With

With lswPolizas.ColumnHeaders
  .Clear
  .Add , , "Tipo", 2040
  .Add , , "Número", 1800
  .Add , , "Documento", 1800
  .Add , , "Inicia", 1800, vbCenter
  .Add , , "Vence", 1800, vbCenter
  .Add , , "Descripción", 2500
  .Add , , "Poliza Id", 1800, vbCenter

End With


 mFechaUltCierre = fxActivos_FechaUltimoCierre
 
 mPermiteRegistro = fxPermiteRegistro
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
 vEdita = False
 
 Call sbActivos_MetodosDepreciacion(cbo)
 
 vPaso = True
  strSQL = "select rtrim(tipo_activo) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from Activos_tipo_activo order by tipo_activo"
  Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
 
  strSQL = "select rtrim(COD_LOCALIZA) as 'Idx', rtrim(descripcion) as 'ItmX'" _
       & " from ACTIVOS_LOCALIZACIONES Where Activa = 1 order by descripcion"
  Call sbCbo_Llena_New(cboLocaliza, strSQL, False, True)
 
 vPaso = False
 

 
 
 vPaso = True
   cboHistorico.AddItem "Cerrados"
   cboHistorico.AddItem "Pendientes"
   cboHistorico.Text = "Cerrados"
 
 
   cboVU.AddItem "Años"
   cboVU.AddItem "Meses"
   cboVU.Text = "Años"
   
 vPaso = False
 
 Call cboTipo_Click


  Call sbBarra_Accion("Activo")
 Call sbInicializaCombos
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 If gActivos.Placa <> "" Then
   Call sbConsultaExterna(gActivos.Placa)
 End If
 
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

  
End Sub

Private Sub sbInicializaCombos()
Dim strSQL As String

vPaso = True
    strSQL = "select rtrim(cod_departamento) as 'IdX' , rtrim(descripcion) as 'ItmX' from Activos_departamentos order by cod_departamento"
    Call sbCbo_Llena_New(cboDepartamento, strSQL, False, True)
vPaso = False

Call cboDepartamento_Click

End Sub

Private Sub sbActivaDesactiva()
Dim strSQL As String, rs As New ADODB.Recordset

'Si no hay periodos depreciados, permite la modificacion
'Total, de lo contrario, realiza bloqueo en datos de calculos
'y ubicacion

If dtpAdquisicion.Value > mFechaUltCierre Then
    cbo.BackColor = vbWhite
    cbo.Locked = False
Else
    cbo.BackColor = &HE0E0E0
    cbo.Locked = True
End If

txtVU.BackColor = cbo.BackColor
txtVU.Locked = cbo.Locked
'cboVU.BackColor = cbo.BackColor

'txtUDProducidas.BackColor = cbo.BackColor
txtUDProducidas.Locked = cbo.Locked

'txtUDAnio.BackColor = cbo.BackColor
txtUDAnio.Locked = cbo.Locked

'cboTipo.BackColor = cbo.BackColor
cboTipo.Locked = cbo.Locked

'txtValorHistorico.BackColor = cbo.BackColor
txtValorHistorico.Locked = cbo.Locked
'txtValorRescate.BackColor = cbo.BackColor
txtValorRescate.Locked = cbo.Locked

dtpAdquisicion.Enabled = IIf(cbo.Locked, False, True)
dtpInstalacion.Enabled = IIf(cbo.Locked, False, True)

cboDepartamento.Enabled = IIf(cbo.Locked, False, True)
cboSeccion.Enabled = IIf(cbo.Locked, False, True)
cboResponsable.Enabled = IIf(cbo.Locked, False, True)


strSQL = "select forzar_tipoActivo from Activos_parametros"
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
  'Nada
Else
 If rs!forzar_TipoActivo = 1 Then
   cboVU.Locked = True
   txtVU.Locked = True
   cbo.Locked = True
 End If
End If
rs.Close

End Sub

Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

txtDescripcion = ""

vCodigo = ""
txtCodigo = ""

lblEstado.Caption = ""

txtPlacaAlterna.Text = ""

txtVU = ""
cboVU.Text = "Años"

txtUDProducidas = 0
txtUDProducidas.Locked = True

txtUDAnio = 0
txtUDAnio.Locked = True


txtValorHistorico = ""
txtValorRescate = ""

dtpAdquisicion.Value = fxFechaServidor
dtpInstalacion.Value = dtpAdquisicion.Value


txtNotas = ""

txtDocCompra = ""
txtProveedor = ""
txtProveedor.Tag = ""

txtModelo = ""
txtSerie = ""
txtMarca = ""
txtOtrasSenas = ""
lsw.ListItems.Clear

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = 0
StatusBarX.Panels(4).Text = 0
StatusBarX.Panels(5).Text = 0

Call sbActivaDesactiva

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curDepAcum As Currency, curDepMes As Currency, curLibros As Currency
Dim itmX As ListViewItem, vPasoX As Boolean

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Item.Index
 Case 1 'Detalle
   lsw.ListItems.Clear
   
      strSQL = "select R.Identificacion,R.nombre,A.Registro_Fecha" _
             & " from Activos_Personas R inner join Activos_Responsables A" _
             & " on R.Identificacion = A.Identificacion" _
             & " Where A.num_placa = '" & vCodigo & "' order by A.registro_fecha desc"
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!Identificacion)
         itmX.SubItems(1) = rs!Nombre
         itmX.SubItems(2) = rs!registro_fecha & ""
     rs.MoveNext
   Loop
   rs.Close
 

 Case 2 'Modificaciones
     lswMod.ListItems.Clear
     
    
    strSQL = "select X.*,rtrim(J.cod_justificacion) + ' - ' + J.descripcion as Justifica" _
          & ",A.nombre,P.cod_proveedor,P.descripcion as Proveedor" _
          & ", case when X.Tipo = 'A' then 'Adicion/Mejora' when X.Tipo = 'M' then 'Mantenimiento' when X.Tipo = 'R' then 'Retiro'" _
          & "       when X.Tipo = 'V' then 'Revaluación'    when X.Tipo = 'D' then 'Deterioro'  else '' end as 'TipoMov'   " _
          & " from Activos_retiro_adicion X inner join Activos_Principal A on X.num_placa = A.num_placa" _
          & " inner join Activos_justificaciones J on X.cod_justificacion = J.cod_justificacion" _
          & " left join Activos_proveedores P on X.compra_proveedor = P.cod_proveedor" _
          & " where X.num_placa = '" & txtCodigo & "' order by X.id_AddRet"
    Call OpenRecordSet(rs, strSQL, 0)
    lswMod.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lswMod.ListItems.Add(, , rs!id_AddRet)
          itmX.SubItems(1) = rs!TipoMov
          itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!monto, "Standard")
          itmX.SubItems(4) = rs!Justifica
          itmX.SubItems(5) = rs!Descripcion
      rs.MoveNext
    Loop
    rs.Close
    
 
  Case 3 'Historico
     strSQL = "exec spActivos_HistoricoConsolidado '" & txtCodigo.Text & "','" & Mid(cboHistorico.Text, 1, 1) & "'"
     Call sbCargaGrid(vGrid, 12, strSQL, True)
     
    
  Case 4 'Composicion
     lswCompo.ListItems.Clear
     curDepAcum = 0
     curDepMes = 0
     strSQL = "select num_placa,'X' as Tipo,nombre as descripcion,depreciacion_periodo,depreciacion_acum" _
            & ",depreciacion_mes,fecha_adquisicion as Fecha, Valor_historico as Libros" _
            & " From Activos_Principal where num_placa = '" & txtCodigo & "'" _
            & " Union " _
            & " select num_placa + '-' + CONVERT(char(3), id_AddRet) as num_Placa,'A' as Tipo,descripcion,depreciacion_periodo,depreciacion_acum" _
            & ",depreciacion_mes,fecha as Fecha, Monto as Libros" _
            & " From Activos_retiro_Adicion where tipo = 'A' and num_placa = '" _
            & txtCodigo & "' order by fecha asc"
     Call OpenRecordSet(rs, strSQL, 0)
     Do While Not rs.EOF
       Select Case rs!Tipo
         Case "X" 'Activo
           Set itmX = lswCompo.ListItems.Add(, , "ACTIVO")
         Case "A" 'Adicion
           Set itmX = lswCompo.ListItems.Add(, , "ADICION/MEJORA")
       End Select
       
           itmX.SubItems(1) = Trim(rs!num_placa)
           itmX.SubItems(2) = rs!depreciacion_periodo & ""
           itmX.SubItems(3) = Format(rs!depreciacion_acum, "Standard")
           itmX.SubItems(4) = Format(rs!DEPRECIACION_MES, "Standard")
           itmX.SubItems(5) = Format(rs!Libros, "Standard")
           itmX.SubItems(6) = rs!Descripcion
           itmX.SubItems(7) = rs!fecha
        
        curDepAcum = curDepAcum + rs!depreciacion_acum
        curDepMes = curDepMes + rs!DEPRECIACION_MES
        curLibros = curLibros + rs!Libros
        
       rs.MoveNext
     
     Loop
     rs.Close
     
     Set itmX = lswCompo.ListItems.Add(, , "")
         itmX.SubItems(3) = "_________"
         itmX.SubItems(4) = "_________"
         itmX.SubItems(5) = "_________"

     Set itmX = lswCompo.ListItems.Add(, , "TOTAL ACTIVO")
         itmX.SubItems(3) = Format(curDepAcum, "Standard")
         itmX.SubItems(4) = Format(curDepMes, "Standard")
         itmX.SubItems(5) = Format(curLibros, "Standard")
         itmX.ForeColor = vbBlue

     Set itmX = lswCompo.ListItems.Add(, , "")
     Set itmX = lswCompo.ListItems.Add(, , "")
     Set itmX = lswCompo.ListItems.Add(, , "T.ADQUIRIDO")
         itmX.ForeColor = vbBlue
         itmX.SubItems(1) = Format(curLibros, "Standard")
     Set itmX = lswCompo.ListItems.Add(, , "T.DEPRECIADO")
         itmX.ForeColor = vbBlue
         itmX.SubItems(1) = Format(curDepAcum, "Standard")
     Set itmX = lswCompo.ListItems.Add(, , "")
         itmX.SubItems(1) = "__________"
     Set itmX = lswCompo.ListItems.Add(, , "VALOR LIBROS")
         itmX.ForeColor = vbBlue
         itmX.SubItems(1) = Format(curLibros - curDepAcum, "Standard")




  Case 5 'Polizas
     lswPolizas.ListItems.Clear
     strSQL = "select P.*,T.descripcion as DescTipo" _
            & " from Activos_polizas_tipos T inner join Activos_polizas P  on T.tipo_poliza = P.tipo_poliza" _
            & " inner join Activos_polizas_asg A on P.cod_poliza = A.cod_poliza " _
            & " and A.num_placa = '" & txtCodigo _
            & "' order by P.fecha_vence desc"
     Call OpenRecordSet(rs, strSQL, 0)
     Do While Not rs.EOF
       Set itmX = lswPolizas.ListItems.Add(, , rs!DescTipo)
           itmX.SubItems(1) = rs!num_poliza
           itmX.SubItems(2) = rs!Documento
           itmX.SubItems(3) = Format(rs!fecha_inicio, "dd/mm/yyyy")
           itmX.SubItems(4) = Format(rs!fecha_vence, "dd/mm/yyyy")
           itmX.SubItems(5) = rs!Descripcion
           itmX.SubItems(6) = rs!cod_poliza
       rs.MoveNext
     Loop
     rs.Close

End Select

vError:
Me.MousePointer = vbDefault


End Sub

Function fxPlaca_ID() As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select dbo.fxActivos_Placa_Id() as 'PLACA_ID'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   strSQL = ""
Else
   strSQL = rs!PLACA_ID
End If
rs.Close

fxPlaca_ID = strSQL

End Function


Public Sub sbConsultaExterna(pNumPlaca As String)
If pNumPlaca <> "" Then
 Call sbConsulta(pNumPlaca)
End If

gActivos.Placa = ""

End Sub


Private Sub sbConsulta(pNumPlaca As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select A.*" _
       & ",Rtrim(D.descripcion) as 'Departamento_Desc'" _
       & ",Rtrim(S.descripcion) as 'Seccion_Desc'" _
       & ",Rtrim(R.Nombre) as 'Responsable_Desc'" _
       & ",isnull(P.descripcion,'N/A') as 'Proveedor',T.descripcion as 'Tipo_Activo_Desc'" _
       & ",isnull(A.cod_Localiza,'00') as 'Localiza_Id',isnull(La.Descripcion,'No Indica') as 'Localiza_Desc'" _
       & " from Activos_Principal A" _
       & " inner join Activos_departamentos D on A.cod_departamento = D.cod_departamento" _
       & " inner join Activos_Secciones S on A.cod_departamento = S.cod_departamento and A.cod_seccion = S.cod_seccion" _
       & " inner join Activos_Personas R on A.identificacion = R.Identificacion" _
       & " inner join Activos_proveedores P on A.cod_proveedor = P.cod_proveedor" _
       & " inner join Activos_tipo_activo T on A.tipo_activo = T.tipo_activo" _
       & "  left join Activos_Localizaciones La on A.cod_localiza = La.cod_localiza" _
       & " where A.num_placa = '" & pNumPlaca & "'"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
   Call sbBarra_Accion("activo")
  vEdita = True
  vPaso = False
    
  vCodigo = rs!num_placa
  txtCodigo = rs!num_placa
  txtPlacaAlterna.Text = rs!Placa_Alterna & ""
 
  txtDescripcion = rs!Nombre
  
  Call sbCboAsignaDato(cboTipo, rs!Tipo_Activo_Desc, True, rs!tipo_activo)
   
  Call sbCboAsignaDato(cbo, fxActivos_MetodoDepreciacion(rs!met_depreciacion), True, rs!met_depreciacion)
  
  txtVU = rs!Vida_Util
  txtUDProducidas = Format(rs!ud_produccion, "Standard")
  txtUDAnio = Format(rs!ud_anio, "Standard")
  
  If rs!Estado = "A" Then
     lblEstado.Caption = "   ACTIVO VIGENTE"
  Else
     lblEstado.Caption = "   ACTIVO RETIRADO"
  End If
  
  If rs!vida_util_en = "A" Then
    cboVU.Text = "Años"
  Else
    cboVU.Text = "Meses"
  End If
      
  txtValorHistorico = Format(rs!valor_historico, "Standard")
  txtValorRescate = Format(rs!valor_desecho, "Standard")
      
  dtpAdquisicion.Value = rs!fecha_adquisicion
  dtpInstalacion.Value = rs!fecha_instalacion
  
  txtNotas = rs!Descripcion
  
  Call sbCboAsignaDato(cboDepartamento, rs!departamento_Desc, True, rs!Cod_Departamento)
  Call sbCboAsignaDato(cboSeccion, rs!seccion_Desc, True, rs!Cod_Seccion)
  Call sbCboAsignaDato(cboResponsable, rs!Responsable_Desc, True, rs!Identificacion)
  Call sbCboAsignaDato(cboLocaliza, rs!Localiza_Desc, True, rs!Localiza_Id)
  
  txtDocCompra = rs!compra_documento
  txtProveedor = rs!Proveedor
  txtProveedor.Tag = rs!COD_PROVEEDOR
  
  txtSerie = rs!NUM_SERIE
  txtModelo = rs!modelo
  txtMarca = rs!marca
  txtOtrasSenas = rs!otras_senas
  
  tcMain.Item(0).Selected = True
  
  StatusBarX.Panels(1).Text = rs!registro_usuario & ""
  StatusBarX.Panels(2).Text = rs!registro_fecha & ""
  StatusBarX.Panels(3).Text = rs!depreciacion_periodo & ""
  StatusBarX.Panels(4).Text = Format(rs!depreciacion_acum, "Standard")
  StatusBarX.Panels(5).Text = Format(rs!DEPRECIACION_MES, "Standard")
  
  Call sbActivaDesactiva
  
Else
 If vEdita Then
  MsgBox "No se encontró registro verifique...", vbInformation
 End If
End If

rs.Close
Me.MousePointer = vbDefault

Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Function fxValida() As Boolean
Dim vMensaje As String, i As Integer, x As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


x = False
vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

'Valida que la Placa no Exista
If Not vEdita Then
    strSQL = "select count(*) as 'Existe' from Activos_Principal Where Num_Placa  = '" & txtCodigo.Text & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs!Existe = 1 Then
        vMensaje = vMensaje & vbCrLf & " - El número de Placa para este activo ya Existe! ..."
    End If
End If

If Trim(txtCodigo.Text) = "" Then vMensaje = vMensaje & vbCrLf & " - Número de Activo no es válido ..."
If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripción del Activo no es válido ..."
If dtpAdquisicion.Value > dtpInstalacion.Value Then vMensaje = vMensaje & vbCrLf & " - La fecha de Adquisición no puede ser menor a la de instalación..."

If mPermiteRegistro = 0 Then
    If dtpAdquisicion.Value <= mFechaUltCierre Then vMensaje = vMensaje & vbCrLf & " - La fecha de Adquisición pertenece a un periodo cerrado..."
End If

If Not IsNumeric(txtVU) Then vMensaje = vMensaje & vbCrLf & " - Vida Util no es válida ..."
If Not IsNumeric(txtValorHistorico) Then vMensaje = vMensaje & vbCrLf & " - Valor Historico no es válido ..."
If Not IsNumeric(txtValorRescate) Then vMensaje = vMensaje & vbCrLf & " - Vida Rescate no es válido ..."
If Not IsNumeric(txtUDAnio) Or Not IsNumeric(txtUDProducidas) Then vMensaje = vMensaje & vbCrLf & " - Las unidades de producción no son válidas ..."


If IsNumeric(txtUDAnio) And IsNumeric(txtUDProducidas) Then
   If CCur(txtUDAnio) > CCur(txtUDProducidas) Then
     vMensaje = vMensaje & vbCrLf & " - Las unidades de producción Anual no pueden ser mayores a las totales..."
   End If
End If

If IsNumeric(txtValorHistorico) And IsNumeric(txtValorRescate) Then
 If CCur(txtValorHistorico) < CCur(txtValorRescate) Then vMensaje = vMensaje & vbCrLf & " - Valor Histórico no puede ser menor al valor de rescate (desecho) ..."
End If

If cboDepartamento.ListCount = 0 Then vMensaje = vMensaje & vbCrLf & " - Departamento no es válido ..."
If cboSeccion.ListCount = 0 Then vMensaje = vMensaje & vbCrLf & " - Sección no es válida ..."
If cboResponsable.ListCount = 0 Then vMensaje = vMensaje & vbCrLf & " - Responsable Inicial no es válido ..."

If txtProveedor.Tag = "" Then vMensaje = vMensaje & vbCrLf & " - Proveedor no es válido ..."

If Len(txtPlacaAlterna.Text) > 0 Then
  strSQL = "select dbo.fxActivos_Registro_Valida_Placa_Alterna('" & txtCodigo.Text & "', '" & txtPlacaAlterna.Text & "') as Resultado"
  Call OpenRecordSet(rs, strSQL)
  If rs!Resultado = 0 Then
    vMensaje = vMensaje & vbCrLf & " - El número de Placa Alterna ya está siendo utilizada por otro activo..."
  End If
  rs.Close
End If

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

  'Si el Activo no ha depreciado ningun periodo aplicar modificacion total,
  'de lo contrario, solo actualizar los datos descriptivos.
  
  strSQL = "update Activos_Principal set nombre = '" & Trim(txtDescripcion.Text) & ", Placa_Alterna = '" & txtPlacaAlterna.Text _
         & "', descripcion = '" & txtNotas & "',compra_documento = '" & txtDocCompra & "',cod_proveedor = '" & txtProveedor.Tag _
         & "', num_serie = '" & txtSerie & "',marca = '" & txtMarca & "',modelo ='" & txtModelo _
         & "', otras_senas = '" & txtOtrasSenas & "'"

  
 If CLng(StatusBarX.Panels(4).Text) = 0 Then
    strSQL = strSQL & ",tipo_activo = '" & cboTipo.ItemData(cboTipo.ListIndex) & "',met_depreciacion = '" _
           & fxActivos_MetodoDepreciacion(cbo.Text) & "',Vida_Util_en = '" _
           & Mid(cboVU.Text, 1, 1) & "',Vida_Util = " & txtVU _
           & ",UD_ANIO = " & CCur(txtUDAnio) & ",UD_PRODUCCION = " & CCur(txtUDProducidas) _
           & ",  valor_historico = " & CCur(txtValorHistorico) & ",valor_desecho = " & CCur(txtValorRescate) _
           & ",  fecha_adquisicion = '" & Format(dtpAdquisicion.Value, "yyyy/mm/dd") _
           & "', fecha_instalacion = '" & Format(dtpInstalacion.Value, "yyyy/mm/dd") _
           & "', cod_departamento = '" & cboDepartamento.ItemData(cboDepartamento.ListIndex) _
           & "', cod_seccion = '" & cboSeccion.ItemData(cboSeccion.ListIndex) _
           & "', identificacion = '" & cboResponsable.ItemData(cboResponsable.ListIndex) _
           & "', cod_Localiza = '" & cboLocaliza.ItemData(cboLocaliza.ListIndex) _
           & "', Localiza_Fecha = dbo.myGetdate(), Modifica_Fecha = getdate(), Modifica_Usuario = '" & glogon.Usuario & "'"

  End If
  
  
  
  strSQL = strSQL & " where num_placa = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
    'Vuelva a contruir tabla de depreciación
    If dtpAdquisicion.Value > mFechaUltCierre Then
        strSQL = "exec spActivos_DepreciacionTabla '" & vCodigo & "','" & glogon.Usuario & "',1"
        Call ConectionExecute(strSQL)
    End If
  
   'Registro del Responsable
   If CLng(StatusBarX.Panels(4).Text) = 0 Then
        strSQL = "exec spActivos_RegistroResponsable '" & vCodigo & "','" & cboResponsable.ItemData(cboResponsable.ListIndex) _
               & "','" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
   End If
  
  Call Bitacora("Modifica", "Activo : " & vCodigo)
  
  

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into Activos_Principal(num_placa, Placa_Alterna, nombre, tipo_activo, descripcion, met_depreciacion" _
          & ", vida_util_en, vida_util, valor_historico, valor_desecho, fecha_adquisicion, fecha_instalacion" _
          & ", cod_departamento, cod_seccion, identificacion, cod_localiza, localiza_fecha, cod_proveedor, compra_documento, num_serie, marca, modelo" _
          & ", otras_senas, estado, depreciacion_acum, depreciacion_mes, depreciacion_periodo, ud_produccion" _
          & ",ud_anio, registro_fecha, registro_usuario) " _
          & " values('" & vCodigo & "', '" & txtPlacaAlterna.Text & "', '" & Trim(txtDescripcion) & "','" & cboTipo.ItemData(cboTipo.ListIndex) & "','" & txtNotas _
          & "','" & fxActivos_MetodoDepreciacion(cbo.Text) & "','" & Mid(cboVU.Text, 1, 1) & "'," & txtVU.Text & "," & CCur(txtValorHistorico) _
          & "," & CCur(txtValorRescate) & ",'" & Format(dtpAdquisicion.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpInstalacion.Value, "yyyy/mm/dd") & "','" & cboDepartamento.ItemData(cboDepartamento.ListIndex) _
          & "','" & cboSeccion.ItemData(cboSeccion.ListIndex) & "','" & cboResponsable.ItemData(cboResponsable.ListIndex) _
          & "','" & cboLocaliza.ItemData(cboLocaliza.ListIndex) & "',dbo.myGetdate(),'" & txtProveedor.Tag & "','" & txtDocCompra _
          & "','" & txtSerie & "','" & txtMarca & "','" & txtModelo & "','" _
          & txtOtrasSenas & "','A',0,0,0," & CCur(txtUDProducidas) & "," & CCur(txtUDAnio) & ",dbo.myGetdate(),'" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)
    
   'Registro del Responsable
   strSQL = "exec spActivos_RegistroResponsable '" & vCodigo & "','" & cboResponsable.ItemData(cboResponsable.ListIndex) _
          & "','" & glogon.Usuario & "'"
   Call ConectionExecute(strSQL)
   
   'Registro Tabla de Depreciación
   strSQL = "exec spActivos_DepreciacionTabla '" & vCodigo & "','" & glogon.Usuario & "'"
   Call ConectionExecute(strSQL)


   'Registro Asiento
   strSQL = "exec spActivos_AsientoRegistroInicial '" & vCodigo & "','" & glogon.Usuario & "'"
   Call ConectionExecute(strSQL)

  Call Bitacora("Registra", "Activo: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtCodigo.Text)


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  If dtpAdquisicion.Value > mFechaUltCierre Then
    
    strSQL = "exec spActivos_EliminaActivo '" & vCodigo & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Elimina", "Activo : " & vCodigo)
    Call sbLimpiaPantalla
     Call sbBarra_Accion("nuevo")
    Call RefrescaTags(Me)
  
  End If

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  tcMain.Item(0).Selected = True
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "num_placa"
  gBusquedas.Orden = "num_placa"
  
  gBusquedas.Col1Name = "Id Placa"
  gBusquedas.Col2Name = "Id Alterna"
  gBusquedas.Col3Name = "Nombre"
  
  gBusquedas.Consulta = "select num_placa, Placa_Alterna, Nombre from Activos_Principal"
  
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select num_placa,nombre from Activos_Principal"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub

Private Sub txtDocCompra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  tcMain.Item(1).Selected = True
  txtModelo.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_FACTURA"
    gBusquedas.Orden = "COD_FACTURA"
    gBusquedas.Consulta = "select COD_FACTURA, COD_PROVEEDOR, TOTAL, FECHA, NOTAS " _
                        & " From CXP_FACTURAS"
    gBusquedas.Filtro = " AND COD_PROVEEDOR = " & txtProveedor.Tag _
                      & " AND YEAR(FECHA) = " & Year(dtpAdquisicion.Value) & " AND MONTH(FECHA) = " & Month(dtpAdquisicion.Value)
    frmBusquedas.Show vbModal
    
    txtDocCompra.Text = gBusquedas.Resultado
    
End If

End Sub



Private Sub txtMarca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOtrasSenas.SetFocus
End Sub

Private Sub txtModelo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSerie.SetFocus
End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDepartamento.SetFocus
End Sub

Private Sub txtOtrasSenas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then lsw.SetFocus
End Sub

Private Sub txtProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocCompra.SetFocus
If KeyCode = vbKeyF4 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_proveedor,descripcion from Activos_proveedores"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    If Trim(gBusquedas.Resultado) <> Trim(txtProveedor.Tag) Then
       txtProveedor.Tag = gBusquedas.Resultado
       txtProveedor = gBusquedas.Resultado2
       txtDocCompra.SetFocus
    End If
End If

End Sub


Private Sub txtSerie_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMarca.SetFocus
End Sub


Private Sub txtUDAnio_GotFocus()
On Error GoTo vError
txtUDAnio = CCur(txtUDAnio)
vError:
End Sub

Private Sub txtUDAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUDProducidas.SetFocus
End Sub

Private Sub txtUDAnio_LostFocus()
On Error GoTo vError
txtUDAnio = Format(CCur(txtUDAnio), "Standard")
vError:
End Sub

Private Sub txtUDProducidas_GotFocus()
On Error GoTo vError
txtUDProducidas = CCur(txtUDProducidas)
vError:
End Sub

Private Sub txtUDProducidas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtUDProducidas_LostFocus()
On Error GoTo vError
txtUDProducidas = Format(CCur(txtUDProducidas), "Standard")
vError:
End Sub

Private Sub txtValorHistorico_GotFocus()
On Error GoTo vError
txtValorHistorico = CCur(txtValorHistorico)
vError:
End Sub

Private Sub txtValorHistorico_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValorRescate.SetFocus
End Sub

Private Sub txtValorHistorico_LostFocus()
On Error GoTo vError
txtValorHistorico = Format(CCur(txtValorHistorico), "Standard")
vError:
End Sub

Private Sub txtValorRescate_GotFocus()
On Error GoTo vError
txtValorRescate = CCur(txtValorRescate)
vError:
End Sub

Private Sub txtValorRescate_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpAdquisicion.SetFocus
Exit Sub

vError:
  txtNotas.SetFocus
End Sub

Private Sub txtValorRescate_LostFocus()
On Error GoTo vError
txtValorRescate = Format(CCur(txtValorRescate), "Standard")
vError:
End Sub

Private Sub txtVU_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 If cboVU.Locked Then
    cbo.SetFocus
 Else
    cboVU.SetFocus
 End If
End If
End Sub


