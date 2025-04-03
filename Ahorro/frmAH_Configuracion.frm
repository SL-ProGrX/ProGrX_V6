VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmAH_Configuracion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Contables para Patrimonio y Excedentes"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   HelpContextID   =   2002
   Icon            =   "frmAH_Configuracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10470
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   120
      Top             =   1080
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   612
      Left            =   8400
      TabIndex        =   46
      Top             =   6360
      Width           =   1692
      _Version        =   1310722
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
      Appearance      =   16
      Picture         =   "frmAH_Configuracion.frx":000C
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4752
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   10212
      _Version        =   1310722
      _ExtentX        =   18013
      _ExtentY        =   8382
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
      Item(0).Caption =   "Patrimonio"
      Item(0).ControlCount=   26
      Item(0).Control(0)=   "Label5(0)"
      Item(0).Control(1)=   "txtPA_Obrero"
      Item(0).Control(2)=   "txtPA_Obrero_Desc"
      Item(0).Control(3)=   "Label7(0)"
      Item(0).Control(4)=   "Label5(1)"
      Item(0).Control(5)=   "Label5(2)"
      Item(0).Control(6)=   "Label5(3)"
      Item(0).Control(7)=   "Label5(4)"
      Item(0).Control(8)=   "txtPA_Patronal"
      Item(0).Control(9)=   "txtPA_Patronal_Desc"
      Item(0).Control(10)=   "txtPA_Custodia"
      Item(0).Control(11)=   "txtPA_Custodia_Desc"
      Item(0).Control(12)=   "txtPA_Capitalizacion"
      Item(0).Control(13)=   "txtPA_Capitalizacion_Desc"
      Item(0).Control(14)=   "txtLiqPas"
      Item(0).Control(15)=   "txtLiqPas_Desc"
      Item(0).Control(16)=   "Label5(11)"
      Item(0).Control(17)=   "txtPA_Devoluciones"
      Item(0).Control(18)=   "txtPA_Devoluciones_Desc"
      Item(0).Control(19)=   "Label5(12)"
      Item(0).Control(20)=   "txtPA_Extra"
      Item(0).Control(21)=   "txtPA_Extra_Desc"
      Item(0).Control(22)=   "txtPA_RentaCap"
      Item(0).Control(23)=   "txtPA_RentaCap_Desc"
      Item(0).Control(24)=   "cboDivisa"
      Item(0).Control(25)=   "Label1(14)"
      Item(1).Caption =   "Excedentes"
      Item(1).ControlCount=   24
      Item(1).Control(0)=   "Label5(5)"
      Item(1).Control(1)=   "Label7(1)"
      Item(1).Control(2)=   "Label5(6)"
      Item(1).Control(3)=   "Label5(7)"
      Item(1).Control(4)=   "Label5(8)"
      Item(1).Control(5)=   "Label5(9)"
      Item(1).Control(6)=   "Label5(10)"
      Item(1).Control(7)=   "txtEX_Distribuir"
      Item(1).Control(8)=   "txtEX_Distribuir_Desc"
      Item(1).Control(9)=   "txtEX_Renta"
      Item(1).Control(10)=   "txtEX_Renta_Desc"
      Item(1).Control(11)=   "txtEX_AjusteCobrar"
      Item(1).Control(12)=   "txtEX_AjustePagar"
      Item(1).Control(13)=   "txtEX_AjustePagar_Desc"
      Item(1).Control(14)=   "txtEx_Donaciones"
      Item(1).Control(15)=   "txtEx_Donaciones_Desc"
      Item(1).Control(16)=   "txtEX_NC"
      Item(1).Control(17)=   "txtEX_NC_Desc"
      Item(1).Control(18)=   "txtEX_Pagar"
      Item(1).Control(19)=   "txtEX_Pagar_Desc"
      Item(1).Control(20)=   "txtEX_AjusteCobrar_Desc"
      Item(1).Control(21)=   "txtEX_Reserva_Desc"
      Item(1).Control(22)=   "txtEX_Reserva"
      Item(1).Control(23)=   "Label5(13)"
      Begin XtremeSuiteControls.FlatEdit txtEX_Reserva 
         Height          =   315
         Left            =   -67600
         TabIndex        =   51
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_RentaCap_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   45
         Top             =   3960
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtPA_Devoluciones_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   41
         Top             =   3360
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEX_Pagar_Desc 
         Height          =   315
         Left            =   -65560
         TabIndex        =   38
         Top             =   3720
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEX_NC_Desc 
         Height          =   312
         Left            =   -65560
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEX_NC 
         Height          =   312
         Left            =   -67600
         TabIndex        =   34
         Top             =   2760
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEx_Donaciones_Desc 
         Height          =   312
         Left            =   -65560
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEx_Donaciones 
         Height          =   312
         Left            =   -67600
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_AjustePagar_Desc 
         Height          =   312
         Left            =   -65560
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEX_AjustePagar 
         Height          =   312
         Left            =   -67600
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_AjusteCobrar 
         Height          =   312
         Left            =   -67600
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_Renta_Desc 
         Height          =   312
         Left            =   -65560
         TabIndex        =   24
         Top             =   1080
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEX_Distribuir_Desc 
         Height          =   312
         Left            =   -65560
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtLiqPas_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   18
         Top             =   2880
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtLiqPas 
         Height          =   312
         Left            =   2400
         TabIndex        =   17
         Top             =   2880
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Extra_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   15
         Top             =   2280
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtPA_Extra 
         Height          =   312
         Left            =   2400
         TabIndex        =   14
         Top             =   2280
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Capitalizacion_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   12
         Top             =   1920
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtPA_Capitalizacion 
         Height          =   312
         Left            =   2400
         TabIndex        =   11
         Top             =   1920
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Custodia_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   9
         Top             =   1560
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtPA_Custodia 
         Height          =   312
         Left            =   2400
         TabIndex        =   8
         Top             =   1560
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Patronal_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   6
         Top             =   1200
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtPA_Patronal 
         Height          =   312
         Left            =   2400
         TabIndex        =   5
         Top             =   1200
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Obrero 
         Height          =   312
         Left            =   2400
         TabIndex        =   2
         Top             =   840
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Obrero_Desc 
         Height          =   312
         Left            =   4440
         TabIndex        =   3
         Top             =   840
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtEX_Distribuir 
         Height          =   312
         Left            =   -67600
         TabIndex        =   20
         Top             =   720
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_Renta 
         Height          =   312
         Left            =   -67600
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_Pagar 
         Height          =   315
         Left            =   -67600
         TabIndex        =   37
         Top             =   3720
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPA_Devoluciones 
         Height          =   312
         Left            =   2400
         TabIndex        =   40
         Top             =   3360
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_AjusteCobrar_Desc 
         Height          =   312
         Left            =   -65560
         TabIndex        =   42
         Top             =   1440
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.FlatEdit txtPA_RentaCap 
         Height          =   312
         Left            =   2400
         TabIndex        =   44
         Top             =   3960
         Width           =   2052
         _Version        =   1310722
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   312
         Left            =   7560
         TabIndex        =   48
         Top             =   480
         Width           =   2532
         _Version        =   1310722
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtEX_Reserva_Desc 
         Height          =   315
         Left            =   -65560
         TabIndex        =   50
         Top             =   3240
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1310722
         _ExtentX        =   9970
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   375
         Index           =   13
         Left            =   -69760
         TabIndex        =   52
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1310722
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reservas s/Excedentes"
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
         Left            =   5520
         TabIndex        =   49
         Top             =   480
         Width           =   1692
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   12
         Left            =   240
         TabIndex        =   43
         Top             =   3840
         Width           =   2172
         _Version        =   1310722
         _ExtentX        =   3831
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Imp.Renta Capitalización"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   11
         Left            =   240
         TabIndex        =   39
         Top             =   3240
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Otras Devoluciones"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   375
         Index           =   10
         Left            =   -69760
         TabIndex        =   36
         Top             =   3600
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Excedente por Pagar"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   9
         Left            =   -69760
         TabIndex        =   33
         Top             =   2640
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Notas de Aplicaciones"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   8
         Left            =   -69760
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Donaciones"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   7
         Left            =   -69760
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Ajustes por Pagar"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   6
         Left            =   -69760
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Ajustes por Cobrar"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   372
         Index           =   1
         Left            =   -69760
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Impuesto de Renta"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   5
         Left            =   -69760
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Excedente a Distribuir"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Transitoria Liquidaciones"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   3
         Left            =   240
         TabIndex        =   13
         Top             =   2160
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Otros Aportes"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Capitalización"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aporte Patronal Custodia"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   372
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aporte Patronal"
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
      Begin XtremeSuiteControls.Label Label5 
         Height          =   372
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aporte Obrero"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patrimonio y Excedentes"
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
      Height          =   480
      Index           =   0
      Left            =   1884
      TabIndex        =   47
      Top             =   360
      Width           =   6492
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAH_Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Dim intCaracteres As Integer 'Almacena el numero total de caracteres de la mascara

Private Sub CargaLblsDatosMED(MedX As Object, lblX As Object)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Carga la Descripción de la Cuenta, en el Label Asignado
'REFERENCIAS:   ProcedimientoErrores
'OBSERVACIONES: Ninguna
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
lblX.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, MedX.Text, 0))

MedX.Text = fxgCntCuentaFormato(True, MedX.Text, 0)

End Sub

Private Function ValidaDatos(MedX As Object, lblX As Object) As Boolean
 ValidaDatos = fxgCntCuentaValida(fxgCntCuentaFormato(False, MedX.Text, 0))
End Function

Private Function ValidaDatosGrabar(MedX As Object) As Boolean

MedX.Text = fxgCntCuentaFormato(False, MedX.Text, 0)
ValidaDatosGrabar = fxgCntCuentaValida(MedX.Text)

End Function


Private Sub cboDivisa_Click()
If vPaso Then Exit Sub

Call sbConsulta

End Sub

Private Sub cmdGuardar_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Guarda Configuración Validada, y Actualiza Etiquetas
'REFERENCIAS:   ProcedimientoErrores
'OBSERVACIONES: Ninguna
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Grabar As String, strSQL As String

On Error GoTo vError

Grabar = "S"
'valida primero todas las casillas

Select Case False
 Case ValidaDatosGrabar(txtPA_Obrero)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPA_Patronal)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPA_Extra)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPA_Capitalizacion)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPA_Custodia)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPA_Devoluciones)
    Grabar = "N"
 Case ValidaDatosGrabar(txtLiqPas)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPA_RentaCap)
    Grabar = "N"
End Select

If tcMain.Item(1).Enabled And Grabar = "S" Then
    Select Case False
     Case ValidaDatosGrabar(txtEX_Renta)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEX_Distribuir)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEX_Pagar)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEX_NC)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEX_AjusteCobrar)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEX_AjustePagar)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEx_Donaciones)
        Grabar = "N"
     Case ValidaDatosGrabar(txtEX_Reserva)
        Grabar = "N"
    End Select
End If



If Grabar = "S" Then
strSQL = "update par_afah set " _
       & "cta_obrero = '" & txtPA_Obrero.Text & "'," _
       & "cta_patronal = '" & txtPA_Patronal.Text & "'," _
       & "cta_extra = '" & txtPA_Extra.Text & "'," _
       & "cta_capitaliza = '" & txtPA_Capitalizacion.Text & "'," _
       & "cta_custodia = '" & txtPA_Custodia.Text & "'," _
       & "cta_devoluciones = '" & txtPA_Devoluciones.Text & "'," _
       & "cta_ExcDist = '" & txtEX_Distribuir.Text & "'," _
       & "cta_ExcPagar = '" & txtEX_Pagar.Text & "'," _
       & "cta_ExcAjustePagar = '" & txtEX_AjustePagar.Text & "'," _
       & "cta_ExcAjusteCobrar = '" & txtEX_AjusteCobrar.Text & "'," _
       & "cta_ExcNC = '" & txtEX_NC.Text & "'," _
       & "cta_ExcDonacion = '" & txtEx_Donaciones.Text & "'," _
       & "cta_renta = '" & txtEX_Renta.Text & "'," _
       & "cta_liqpas = '" & txtLiqPas.Text & "'," _
       & "cta_RentaCap = '" & txtPA_RentaCap.Text & "'," _
       & "CTA_EXC_RESERVA = '" & txtEX_Reserva.Text & "'" _
       & " Where Cod_Divisa = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
       
Call ConectionExecute(strSQL)

Call CargaLblsDatosMED(txtPA_Obrero, txtPA_Obrero_Desc)
Call CargaLblsDatosMED(txtPA_Patronal, txtPA_Patronal_Desc)
Call CargaLblsDatosMED(txtPA_Extra, txtPA_Extra_Desc)
Call CargaLblsDatosMED(txtPA_Capitalizacion, txtPA_Capitalizacion_Desc)
Call CargaLblsDatosMED(txtPA_Custodia, txtPA_Custodia_Desc)
Call CargaLblsDatosMED(txtPA_Devoluciones, txtPA_Devoluciones_Desc)


Call CargaLblsDatosMED(txtEX_Renta, txtEX_Renta_Desc)
Call CargaLblsDatosMED(txtEX_Distribuir, txtEX_Distribuir_Desc)
Call CargaLblsDatosMED(txtEX_Pagar, txtEX_Pagar_Desc)
Call CargaLblsDatosMED(txtEX_NC, txtEX_NC_Desc)
Call CargaLblsDatosMED(txtEX_AjusteCobrar, txtEX_AjusteCobrar_Desc)
Call CargaLblsDatosMED(txtEX_AjustePagar, txtEX_AjustePagar_Desc)
Call CargaLblsDatosMED(txtEx_Donaciones, txtEx_Donaciones_Desc)

Call CargaLblsDatosMED(txtEX_Reserva, txtEX_Reserva_Desc)

Call CargaLblsDatosMED(txtLiqPas, txtLiqPas_Desc)
Call CargaLblsDatosMED(txtPA_RentaCap, txtPA_RentaCap_Desc)


MsgBox "La Información se guardó satisfactoriamente ...", vbInformation

Else

 MsgBox "No se puede guardar la información, Verifique las cuentas ingresadas...", vbInformation

End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

vModulo = 2

Set Me.imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vPaso = True
    strSQL = "select rtrim(cod_Divisa) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
            & " from vSys_Divisas order by Divisa_Local desc, cod_divisa asc"
    Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)
vPaso = False


tcMain.Item(0).Selected = True


Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:

End Sub

Private Sub sbBusqueda(Index As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Carga la ventana de Busquedas Rápidas
'REFERENCIAS:   ProcedimientoErrores
'OBSERVACIONES: Enviar Parametros y Asignar Etiqueta para Resultados
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

On Error GoTo vError


Call sbgCntCuentaConsulta

If gBusquedas.Resultado <> "" Then
Select Case Index
    Case 0
     txtPA_RentaCap.Text = gBusquedas.Resultado
     txtPA_RentaCap.SetFocus
    Case 4
     txtPA_Obrero.Text = gBusquedas.Resultado
     txtPA_Obrero.SetFocus
    Case 5
     txtPA_Patronal.Text = gBusquedas.Resultado
     txtPA_Patronal.SetFocus
    Case 6
     txtPA_Extra.Text = gBusquedas.Resultado
     txtPA_Extra.SetFocus
    Case 7
     txtPA_Capitalizacion.Text = gBusquedas.Resultado
     txtPA_Capitalizacion.SetFocus
    Case 8
     txtEX_Renta.Text = gBusquedas.Resultado
     txtEX_Renta.SetFocus
    Case 9
     txtPA_Custodia.Text = gBusquedas.Resultado
     txtPA_Custodia.SetFocus
    Case 12
     txtLiqPas.Text = gBusquedas.Resultado
     txtLiqPas.SetFocus
    Case 13
     txtEX_Distribuir.Text = gBusquedas.Resultado
     txtEX_Distribuir.SetFocus
    Case 14
     txtEX_AjustePagar.Text = gBusquedas.Resultado
     txtEX_AjustePagar.SetFocus
    Case 15
     txtEX_AjusteCobrar.Text = gBusquedas.Resultado
     txtEX_AjusteCobrar.SetFocus
    Case 16
     txtEX_NC.Text = gBusquedas.Resultado
     txtEX_NC.SetFocus
    Case 17
     txtEX_Pagar.Text = gBusquedas.Resultado
     txtEX_Pagar.SetFocus
    Case 18
     txtEx_Donaciones.Text = gBusquedas.Resultado
     txtEx_Donaciones.SetFocus
    
    Case 19
     txtPA_Devoluciones.Text = gBusquedas.Resultado
     txtPA_Devoluciones.SetFocus
    
    
    Case 20
     txtEX_Reserva.Text = gBusquedas.Resultado
     txtEX_Reserva.SetFocus
End Select
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

tcMain.Item(0).Selected = True

'Consulta
strSQL = "exec spPAT_Parametros '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"

Call OpenRecordSet(rs, strSQL)

txtPA_Obrero.Text = rs!CTA_OBR_MASK
txtPA_Obrero_Desc.Text = rs!CTA_OBR_DESC

txtPA_Patronal.Text = rs!CTA_PAT_MASK
txtPA_Patronal_Desc.Text = rs!CTA_PAT_DESC

txtPA_Custodia.Text = rs!CTA_CST_MASK
txtPA_Custodia_Desc.Text = rs!CTA_CST_DESC

txtPA_Capitalizacion.Text = rs!CTA_CAP_MASK
txtPA_Capitalizacion_Desc.Text = rs!CTA_CAP_DESC

txtPA_Extra.Text = rs!CTA_EXT_MASK
txtPA_Extra_Desc.Text = rs!CTA_EXT_DESC

txtPA_Devoluciones.Text = rs!CTA_DEV_MASK
txtPA_Devoluciones_Desc.Text = rs!CTA_DEV_DESC

txtLiqPas.Text = rs!CTA_LIQ_MASK
txtLiqPas_Desc.Text = rs!CTA_LIQ_DESC

txtPA_RentaCap.Text = rs!CTA_RNTC_MASK
txtPA_RentaCap_Desc.Text = rs!CTA_RNTC_DESC

txtEX_Renta.Text = rs!CTA_RNT_MASK
txtEX_Renta_Desc.Text = rs!CTA_RNT_DESC

txtEX_Distribuir.Text = rs!CTA_EDst_MASK
txtEX_Distribuir_Desc.Text = rs!CTA_EDst_DESC

txtEX_Pagar.Text = rs!CTA_EPg_MASK
txtEX_Pagar_Desc.Text = rs!CTA_EPg_DESC

txtEX_AjusteCobrar.Text = rs!CTA_ECxC_MASK
txtEX_AjusteCobrar_Desc.Text = rs!CTA_ECxC_DESC

txtEX_AjustePagar.Text = rs!CTA_ECxP_MASK
txtEX_AjustePagar_Desc.Text = rs!CTA_ECxP_DESC

txtEX_NC.Text = rs!CTA_ENC_MASK
txtEX_NC_Desc.Text = rs!CTA_ENC_DESC

txtEx_Donaciones.Text = rs!CTA_EDon_MASK
txtEx_Donaciones_Desc.Text = rs!CTA_EDon_DESC


txtEX_Reserva.Text = rs!CTA_EReserva_Mask
txtEX_Reserva_Desc.Text = rs!CTA_EReserva_DESC

If rs!EXCEDENTES_CFG = 1 Then
  tcMain.Item(1).Enabled = True
Else
  tcMain.Item(1).Enabled = False
End If

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub TimerX_Timer()

On Error GoTo vError

TimerX.Interval = 0
TimerX.Enabled = False


Call sbConsulta

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtEX_AjusteCobrar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(15)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_AjusteCobrar, txtEX_AjusteCobrar_Desc) Then txtEX_NC.SetFocus
End Sub

Private Sub txtEX_AjustePagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(14)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_AjustePagar, txtEX_AjustePagar_Desc) Then txtEX_AjusteCobrar.SetFocus
End Sub

Private Sub txtEX_Distribuir_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(13)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_Distribuir, txtEX_Distribuir_Desc) Then txtEX_AjustePagar.SetFocus
End Sub

Private Sub txtEX_Donaciones_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(18)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEx_Donaciones, txtEx_Donaciones_Desc) Then txtEX_Pagar.SetFocus
End Sub


Private Sub txtEX_NC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(16)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_NC, txtEX_NC_Desc) Then txtEX_Reserva.SetFocus
End Sub

Private Sub txtEX_Pagar_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
    If KeyCode = vbKeyF4 Then Call sbBusqueda(17)
    If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_Pagar, txtEX_Pagar_Desc) Then cmdGuardar.SetFocus
vError:
End Sub

Private Sub txtEX_Renta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(8)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_Renta, txtEX_Renta_Desc) Then txtEX_Distribuir.SetFocus
End Sub



Private Sub txtEX_Reserva_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(20)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtEX_Reserva, txtEX_Reserva_Desc) Then txtEX_Pagar.SetFocus
End Sub

Private Sub txtLiqPas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF4 Then Call sbBusqueda(12)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtLiqPas, txtLiqPas_Desc) Then txtPA_RentaCap.SetFocus
End Sub

Private Sub txtPA_Capitalizacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(7)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtPA_Capitalizacion, txtPA_Capitalizacion_Desc) Then txtPA_Custodia.SetFocus
End Sub

Private Sub txtPA_Custodia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(9)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtPA_Custodia, txtPA_Custodia_Desc) Then txtPA_Devoluciones.SetFocus
End Sub


Private Sub txtPA_Devoluciones_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(19)
End Sub

Private Sub txtPA_Extra_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(6)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtPA_Extra, txtPA_Extra_Desc) Then txtPA_Capitalizacion.SetFocus
End Sub

Private Sub txtPA_Obrero_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtPA_Obrero, txtPA_Obrero_Desc) Then txtPA_Patronal.SetFocus
End Sub

Private Sub txtPA_Patronal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(5)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtPA_Patronal, txtPA_Patronal_Desc) Then txtPA_Extra.SetFocus
End Sub

Private Sub txtPA_RentaCap_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub


