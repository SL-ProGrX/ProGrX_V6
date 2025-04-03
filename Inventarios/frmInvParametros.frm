VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.ShortcutBar.v19.1.0.ocx"
Begin VB.Form frmInvParametros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventarios/POS: Parámetros Generales"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmInvParametros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1320
      Width           =   4932
   End
   Begin VB.CheckBox chkEnlaceConta 
      Alignment       =   1  'Right Justify
      Caption         =   "Enlace con Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   14
      Top             =   960
      Width           =   2532
   End
   Begin VB.CheckBox chkEnlaceSIF 
      Alignment       =   1  'Right Justify
      Caption         =   "Enlace con Crédito"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      TabIndex        =   13
      Top             =   600
      Width           =   2532
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   4212
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   8052
      _ExtentX        =   14208
      _ExtentY        =   7435
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Cuentas"
      TabPicture(0)   =   "frmInvParametros.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(14)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(16)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(17)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(18)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCtaGastos"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtCtaGastosDesc"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCtaImpVentas"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCtaImpVentasDesc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCtaImpConsumo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCtaImpConsumoDesc"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCtaComisiones"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCtaComisionesDesc"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCtaCostoVentasDesc"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCtaCostoVentas"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCtaRecibos"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCtaRecibosDesc"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCtaNotas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtCtaNotasDesc"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtCtaVentasDesc"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtCtaVentas"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "Tipos de Asiento"
      TabPicture(1)   =   "frmInvParametros.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(4)"
      Tab(1).Control(1)=   "Label1(5)"
      Tab(1).Control(2)=   "Label1(6)"
      Tab(1).Control(3)=   "Label1(7)"
      Tab(1).Control(4)=   "Label1(8)"
      Tab(1).Control(5)=   "Label1(9)"
      Tab(1).Control(6)=   "Label1(10)"
      Tab(1).Control(7)=   "Label1(11)"
      Tab(1).Control(8)=   "Label1(12)"
      Tab(1).Control(9)=   "Label1(15)"
      Tab(1).Control(10)=   "txtTA_FM"
      Tab(1).Control(11)=   "txtTA_FMDes"
      Tab(1).Control(12)=   "txtTA_FA"
      Tab(1).Control(13)=   "txtTA_FADes"
      Tab(1).Control(14)=   "txtTA_EN"
      Tab(1).Control(15)=   "txtTA_ENDes"
      Tab(1).Control(16)=   "txtTA_SA"
      Tab(1).Control(17)=   "txtTA_SADes"
      Tab(1).Control(18)=   "txtTA_TR"
      Tab(1).Control(19)=   "txtTA_TRDes"
      Tab(1).Control(20)=   "txtTA_DV"
      Tab(1).Control(21)=   "txtTA_DVDes"
      Tab(1).Control(22)=   "txtTA_RE"
      Tab(1).Control(23)=   "txtTA_REDes"
      Tab(1).Control(24)=   "txtTA_NC"
      Tab(1).Control(25)=   "txtTA_NCDes"
      Tab(1).Control(26)=   "txtTA_NDDes"
      Tab(1).Control(27)=   "txtTA_ND"
      Tab(1).Control(28)=   "txtTA_GENDes"
      Tab(1).Control(29)=   "txtTA_GEN"
      Tab(1).ControlCount=   30
      Begin VB.TextBox txtCtaVentas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   58
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   840
         Width           =   1932
      End
      Begin VB.TextBox txtCtaVentasDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   840
         Width           =   4572
      End
      Begin VB.TextBox txtCtaNotasDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3360
         Width           =   4572
      End
      Begin VB.TextBox txtCtaNotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   53
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3360
         Width           =   1932
      End
      Begin VB.TextBox txtCtaRecibosDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   3000
         Width           =   4572
      End
      Begin VB.TextBox txtCtaRecibos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   50
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3000
         Width           =   1932
      End
      Begin VB.TextBox txtTA_GEN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   48
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3840
         Width           =   1572
      End
      Begin VB.TextBox txtTA_GENDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   3840
         Width           =   4932
      End
      Begin VB.TextBox txtTA_ND 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   45
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3480
         Width           =   1572
      End
      Begin VB.TextBox txtTA_NDDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   3480
         Width           =   4932
      End
      Begin VB.TextBox txtCtaCostoVentas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   42
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2640
         Width           =   1932
      End
      Begin VB.TextBox txtCtaCostoVentasDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2640
         Width           =   4572
      End
      Begin VB.TextBox txtTA_NCDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3120
         Width           =   4932
      End
      Begin VB.TextBox txtTA_NC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   39
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3120
         Width           =   1572
      End
      Begin VB.TextBox txtTA_REDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2760
         Width           =   4932
      End
      Begin VB.TextBox txtTA_RE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   37
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2760
         Width           =   1572
      End
      Begin VB.TextBox txtTA_DVDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2400
         Width           =   4932
      End
      Begin VB.TextBox txtTA_DV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   35
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2400
         Width           =   1572
      End
      Begin VB.TextBox txtTA_TRDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   2040
         Width           =   4932
      End
      Begin VB.TextBox txtTA_TR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   33
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2040
         Width           =   1572
      End
      Begin VB.TextBox txtTA_SADes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   1680
         Width           =   4932
      End
      Begin VB.TextBox txtTA_SA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   31
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1680
         Width           =   1572
      End
      Begin VB.TextBox txtTA_ENDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1320
         Width           =   4932
      End
      Begin VB.TextBox txtTA_EN 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   29
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1320
         Width           =   1572
      End
      Begin VB.TextBox txtTA_FADes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   4932
      End
      Begin VB.TextBox txtTA_FA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   27
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   960
         Width           =   1572
      End
      Begin VB.TextBox txtTA_FMDes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   4932
      End
      Begin VB.TextBox txtTA_FM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73560
         TabIndex        =   25
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   600
         Width           =   1572
      End
      Begin VB.TextBox txtCtaComisionesDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2280
         Width           =   4572
      End
      Begin VB.TextBox txtCtaComisiones 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   23
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2280
         Width           =   1932
      End
      Begin VB.TextBox txtCtaImpConsumoDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1920
         Width           =   4572
      End
      Begin VB.TextBox txtCtaImpConsumo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   21
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1920
         Width           =   1932
      End
      Begin VB.TextBox txtCtaImpVentasDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1560
         Width           =   4572
      End
      Begin VB.TextBox txtCtaImpVentas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   19
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1560
         Width           =   1932
      End
      Begin VB.TextBox txtCtaGastosDesc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   4572
      End
      Begin VB.TextBox txtCtaGastos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   17
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   1200
         Width           =   1932
      End
      Begin VB.Label Label1 
         Caption         =   "Ventas (Ingresos)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   18
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Notas Deb/Cre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   17
         Left            =   120
         TabIndex        =   55
         Top             =   3360
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Recibos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   16
         Left            =   120
         TabIndex        =   52
         Top             =   3000
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -74880
         TabIndex        =   49
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Notas de Débito"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -74880
         TabIndex        =   46
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Costo Ventas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   14
         Left            =   120
         TabIndex        =   43
         Top             =   2640
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Notas de Crédito"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   12
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Recibos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   -74880
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Devoluciones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   -74880
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Traslados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   -74880
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Salidas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   -74880
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Entradas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -74880
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Factura Auto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -74880
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Factura Manual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -74880
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Comisiones"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Cosumo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1920
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Imp.Ventas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label1 
         Caption         =   "Gastos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.PushButton cmdGuardar 
      Height          =   492
      Left            =   6600
      TabIndex        =   56
      Top             =   6120
      Width           =   1572
      _Version        =   1245185
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Guardar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmInvParametros.frx":0342
   End
   Begin XtremeShortcutBar.ShortcutCaption lbl 
      Height          =   492
      Left            =   0
      TabIndex        =   60
      Top             =   0
      Width           =   8292
      _Version        =   1245185
      _ExtentX        =   14626
      _ExtentY        =   868
      _StockProps     =   14
      Caption         =   "Parametros Generales de Enlace con otros Sistemas"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   13
      Left            =   1800
      TabIndex        =   15
      Top             =   1320
      Width           =   1212
   End
End
Attribute VB_Name = "frmInvParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub sbCargaParametros()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCodEmpresa As Long

On Error GoTo vError

strSQL = "select * from pv_parametros_gen"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   
   vCodEmpresa = rs!cod_Empresa
   
   chkEnlaceConta.Value = IIf((rs!enlace_conta = "S"), vbChecked, vbUnchecked)
   chkEnlaceSIF = IIf((rs!enlace_sif = "S"), vbChecked, vbUnchecked)
   
   
   txtCtaVentas.Text = fxgCntCuentaFormato(True, rs!CTA_VENTAS_ING & "")
   txtCtaVentasDesc.Text = fxgCntCuentaDesc(rs!CTA_VENTAS_ING & "")
   
   txtCtaComisiones = fxgCntCuentaFormato(True, rs!cta_comisiones)
   txtCtaComisionesDesc = fxgCntCuentaDesc(rs!cta_comisiones)
   
   txtCtaGastos = fxgCntCuentaFormato(True, rs!cta_gastos)
   txtCtaGastosDesc = fxgCntCuentaDesc(rs!cta_gastos)
   
   txtCtaImpVentas = fxgCntCuentaFormato(True, rs!cta_imp_renta) 'Imp.ventas
   txtCtaImpVentasDesc = fxgCntCuentaDesc(rs!cta_imp_renta)
   
   txtCtaImpConsumo = fxgCntCuentaFormato(True, rs!cta_imp_consumo)
   txtCtaImpConsumoDesc = fxgCntCuentaDesc(rs!cta_imp_consumo)
   
   txtCtaCostoVentas = fxgCntCuentaFormato(True, rs!cta_costo_ventas)
   txtCtaCostoVentasDesc = fxgCntCuentaDesc(rs!cta_costo_ventas)
   
   txtCtaRecibos = fxgCntCuentaFormato(True, rs!cta_Recibos)
   txtCtaRecibosDesc = fxgCntCuentaDesc(rs!cta_Recibos)
   
   txtCtaNotas = fxgCntCuentaFormato(True, rs!cta_notas)
   txtCtaNotasDesc = fxgCntCuentaDesc(rs!cta_notas)
   
   
   txtTA_DV = rs!ta_devoluciones
   txtTA_DVDes = fxgCntTipoAsientoDesc(rs!ta_devoluciones)
   txtTA_EN = rs!ta_entradas
   txtTA_ENDes = fxgCntTipoAsientoDesc(rs!ta_entradas)
   txtTA_FA = rs!ta_factura_auto
   txtTA_FADes = fxgCntTipoAsientoDesc(rs!ta_factura_auto)
   txtTA_FM = rs!ta_factura_man
   txtTA_FMDes = fxgCntTipoAsientoDesc(rs!ta_factura_man)
   
   txtTA_NC = rs!ta_nc
   txtTA_NCDes = fxgCntTipoAsientoDesc(rs!ta_nc)
   
   txtTA_ND = rs!ta_nd
   txtTA_NDDes = fxgCntTipoAsientoDesc(rs!ta_nd)
   
   txtTA_GEN = rs!ta_gen
   txtTA_GENDes = fxgCntTipoAsientoDesc(rs!ta_gen)
   
   txtTA_RE = rs!ta_recibos
   txtTA_REDes = fxgCntTipoAsientoDesc(rs!ta_recibos)
   txtTA_SA = rs!ta_salidas
   txtTA_SADes = fxgCntTipoAsientoDesc(rs!ta_salidas)
   txtTA_TR = rs!ta_traslados
   txtTA_TRDes = fxgCntTipoAsientoDesc(rs!ta_traslados)
   
End If
rs.Close

strSQL = "select cod_contabilidad,nombre from CntX_Contabilidades"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  cbo.AddItem rs!Nombre
  cbo.ItemData(cbo.NewIndex) = rs!COD_CONTABILIDAD
  rs.MoveNext
Loop
rs.Close


strSQL = "select cod_contabilidad,nombre from CntX_Contabilidades  where cod_Contabilidad = " & GLOBALES.gEnlace
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  cbo.Text = rs!Nombre
End If
rs.Close

Exit Sub

vError:
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select isnull(count(*),0) as Existe from pv_parametros_gen"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   strSQL = "insert pv_parametros_gen(tipo_cambio,cod_empresa,enlace_conta,enlace_sif" _
          & ",CTA_VENTAS_ING,cta_gastos,cta_comisiones,cta_imp_renta,cta_imp_consumo,cta_costo_ventas" _
          & ",cta_recibos,cta_notas,ta_factura_man" _
          & ",ta_factura_auto,ta_entradas,ta_salidas,ta_traslados,ta_nc,ta_recibos" _
          & ",ta_devoluciones,ta_nd,ta_gen)" _
          & " values(1," & cbo.ItemData(cbo.ListIndex) _
          & ",'" & IIf((chkEnlaceConta.Value = vbChecked), "S", "N") & "','" & IIf((chkEnlaceSIF.Value = vbChecked), "S", "N") _
          & "','" & fxgCntCuentaFormato(False, txtCtaVentas) & "','" & fxgCntCuentaFormato(False, txtCtaGastos) _
          & "','" & fxgCntCuentaFormato(False, txtCtaComisiones) _
          & "','" & fxgCntCuentaFormato(False, txtCtaImpVentas) & "','" & fxgCntCuentaFormato(False, txtCtaImpConsumo) _
          & "','" & fxgCntCuentaFormato(False, txtCtaCostoVentas) & "','" & fxgCntCuentaFormato(False, txtCtaRecibos) _
          & "','" & fxgCntCuentaFormato(False, txtCtaNotas) _
          & "','" & txtTA_FM & "','" & txtTA_FA & "','" & txtTA_EN & "','" & txtTA_SA & "','" & txtTA_TR _
          & "','" & txtTA_NC & "','" & txtTA_RE & "','" & txtTA_DV & "','" & txtTA_ND & "','" & txtTA_GEN & "')"
Else
  strSQL = "update pv_parametros_gen set cod_empresa = " & cbo.ItemData(cbo.ListIndex) _
         & ",enlace_conta = '" & IIf((chkEnlaceConta.Value = vbChecked), "S", "N") _
         & "',enlace_sif = '" & IIf((chkEnlaceSIF.Value = vbChecked), "S", "N") _
         & "',CTA_VENTAS_ING = '" & fxgCntCuentaFormato(False, txtCtaVentas) _
         & "',cta_gastos = '" & fxgCntCuentaFormato(False, txtCtaGastos) _
         & "',cta_comisiones = '" & fxgCntCuentaFormato(False, txtCtaComisiones) _
         & "',cta_imp_renta = '" & fxgCntCuentaFormato(False, txtCtaImpVentas) _
         & "',cta_imp_consumo = '" & fxgCntCuentaFormato(False, txtCtaImpConsumo) _
         & "',cta_costo_ventas = '" & fxgCntCuentaFormato(False, txtCtaCostoVentas) _
         & "',cta_recibos = '" & fxgCntCuentaFormato(False, txtCtaRecibos) _
         & "',cta_notas = '" & fxgCntCuentaFormato(False, txtCtaNotas) _
         & "',ta_factura_man = '" & txtTA_FM & "',ta_factura_auto= '" & txtTA_FA _
         & "',ta_entradas = '" & txtTA_EN & "',ta_salidas = '" & txtTA_SA _
         & "',ta_traslados = '" & txtTA_TR & "',ta_nc = '" & txtTA_NC _
         & "',ta_recibos = '" & txtTA_RE & "',ta_devoluciones = '" & txtTA_DV _
         & "',ta_nd = '" & txtTA_ND & "',ta_gen = '" & txtTA_GEN & "'"
        
End If
rs.Close

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Parámetros Guardados Satisfactoriamente...", vbInformation

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
  vModulo = 34
End Sub

Private Sub Form_Load()
  vModulo = 34
  
  Call sbCargaParametros
  
  ssTab.Tab = 0

  Call Formularios(Me)
  Call RefrescaTags(Me)
End Sub

Private Sub txtCtaComisionesDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaCostoVentas.SetFocus
End Sub


Private Sub txtCtaVentas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaVentasDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaVentas = gCuenta
End If
End Sub

Private Sub txtCtaVentas_LostFocus()
txtCtaVentasDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaVentas))
txtCtaVentas.Text = fxgCntCuentaFormato(True, txtCtaVentas)
End Sub

Private Sub txtCtaCostoVentas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaCostoVentasDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaCostoVentas = gCuenta
End If
End Sub

Private Sub txtCtaCostoVentas_LostFocus()
txtCtaCostoVentasDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaCostoVentas))
txtCtaCostoVentas = fxgCntCuentaFormato(True, txtCtaCostoVentas)
End Sub

Private Sub txtCtaCostoVentasDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaRecibos.SetFocus
End Sub

Private Sub txtCtaRecibos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaRecibosDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaRecibos = gCuenta
End If
End Sub

Private Sub txtCtaRecibos_LostFocus()
txtCtaRecibosDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaRecibos))
txtCtaRecibos = fxgCntCuentaFormato(True, txtCtaRecibos)
End Sub

Private Sub txtCtaRecibosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaNotas.SetFocus
End Sub

Private Sub txtCtaNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaNotasDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaNotas = gCuenta
End If
End Sub

Private Sub txtCtaNotas_LostFocus()
txtCtaNotasDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaNotas))
txtCtaNotas = fxgCntCuentaFormato(True, txtCtaNotas)
End Sub


Private Sub txtCtaGastos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaGastosDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaGastos = gCuenta
End If
End Sub

Private Sub txtCtaGastos_LostFocus()
txtCtaGastosDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaGastos))
txtCtaGastos = fxgCntCuentaFormato(True, txtCtaGastos)
End Sub

Private Sub txtCtaComisiones_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaComisionesDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaComisiones = gCuenta
End If
End Sub

Private Sub txtCtaComisiones_LostFocus()
txtCtaComisionesDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaComisiones))
txtCtaComisiones = fxgCntCuentaFormato(True, txtCtaComisiones)
End Sub

 
Private Sub txtCtaGastosDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaImpVentas.SetFocus
End Sub

Private Sub txtCtaImpConsumo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaImpConsumoDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaImpConsumo = gCuenta
End If
End Sub

Private Sub txtCtaImpConsumo_LostFocus()
txtCtaImpConsumoDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaImpConsumo))
txtCtaImpConsumo = fxgCntCuentaFormato(True, txtCtaImpConsumo)
End Sub


Private Sub txtCtaImpConsumoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaComisiones.SetFocus
End Sub

Private Sub txtCtaImpVentas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaImpVentasDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCtaImpVentas = gCuenta
End If
End Sub

Private Sub txtCtaImpVentas_LostFocus()
txtCtaImpVentasDesc = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaImpVentas))
txtCtaImpVentas = fxgCntCuentaFormato(True, txtCtaImpVentas)
End Sub

Private Sub txtCtaImpVentasDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaImpConsumo.SetFocus
End Sub

Private Sub txtCtaNotasDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  ssTab.Tab = 1
  txtTA_FM.SetFocus
End If
End Sub

Private Sub txtTA_FM_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_FA.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_FM = gBusquedas.Resultado
   txtTA_FMDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_FM_LostFocus()
txtTA_FMDes = fxgCntTipoAsientoDesc(txtTA_FM)
End Sub


Private Sub txtTA_FA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_EN.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_FA = gBusquedas.Resultado
   txtTA_FADes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_FA_LostFocus()
txtTA_FADes = fxgCntTipoAsientoDesc(txtTA_FA)
End Sub



Private Sub txtTA_EN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_SA.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_EN = gBusquedas.Resultado
   txtTA_ENDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_EN_LostFocus()
txtTA_ENDes = fxgCntTipoAsientoDesc(txtTA_EN)
End Sub


Private Sub txtTA_SA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_TR.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_SA = gBusquedas.Resultado
   txtTA_SADes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_SA_LostFocus()
txtTA_SADes = fxgCntTipoAsientoDesc(txtTA_SA)
End Sub


Private Sub txtTA_TR_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_DV.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_TR = gBusquedas.Resultado
   txtTA_TRDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_TR_LostFocus()
txtTA_TRDes = fxgCntTipoAsientoDesc(txtTA_TR)
End Sub



Private Sub txtTA_DV_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_RE.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_DV = gBusquedas.Resultado
   txtTA_DVDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_DV_LostFocus()
txtTA_DVDes = fxgCntTipoAsientoDesc(txtTA_DV)
End Sub


Private Sub txtTA_RE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_NC.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_RE = gBusquedas.Resultado
   txtTA_REDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_RE_LostFocus()
txtTA_REDes = fxgCntTipoAsientoDesc(txtTA_RE)
End Sub


Private Sub txtTA_NC_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_ND.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_NC = gBusquedas.Resultado
   txtTA_NCDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_NC_LostFocus()
txtTA_NCDes = fxgCntTipoAsientoDesc(txtTA_NC)
End Sub

Private Sub txtTA_ND_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTA_GEN.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_ND = gBusquedas.Resultado
   txtTA_NDDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_ND_LostFocus()
txtTA_NDDes = fxgCntTipoAsientoDesc(txtTA_ND)
End Sub


Private Sub txtTA_GEN_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdGuardar.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Asiento"
   gBusquedas.Orden = "Tipo_Asiento"
   gBusquedas.Consulta = "select tipo_asiento,descripcion from CntX_Tipos_Asientos"
   gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
   frmBusquedas.Show vbModal
   txtTA_GEN = gBusquedas.Resultado
   txtTA_GENDes = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTA_GEN_LostFocus()
txtTA_GENDes = fxgCntTipoAsientoDesc(txtTA_GEN)
End Sub

