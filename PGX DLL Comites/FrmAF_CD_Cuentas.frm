VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmAF_CD_Cuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   10440
   Icon            =   "FrmAF_CD_Cuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabCuentas 
      Height          =   8850
      Left            =   13320
      TabIndex        =   0
      Top             =   1080
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   15610
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Desembolsos"
      TabPicture(0)   =   "FrmAF_CD_Cuentas.frx":3482
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label18(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label18(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tlbPrincipal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Aprobación"
      TabPicture(1)   =   "FrmAF_CD_Cuentas.frx":349E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbobancApro"
      Tab(1).Control(1)=   "vGrid"
      Tab(1).Control(2)=   "tlbAprobacion"
      Tab(1).Control(3)=   "Label9"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Consulta de Envio de Operaciones a Tesorería"
      TabPicture(2)   =   "FrmAF_CD_Cuentas.frx":34BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswinfoenvio"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Label8"
      Tab(2).ControlCount=   3
      Begin MSComctlLib.ListView lswinfoenvio 
         Height          =   6240
         Left            =   -74790
         TabIndex        =   14
         Top             =   2400
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   11007
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No.Operación"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Comite"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Cédula"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Delegado"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "No.Solicitud"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Envio Tesoreria"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Usario Envio"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Monto"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Liquidación"
            Object.Width           =   2117
         EndProperty
      End
      Begin VB.ComboBox cbobancApro 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   585
         Width           =   6120
      End
      Begin VB.Frame Frame4 
         Caption         =   "Consulta de Envio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -74790
         TabIndex        =   3
         Top             =   420
         Width           =   10260
         Begin VB.TextBox TxtSolicitud 
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
            Left            =   7410
            TabIndex        =   6
            Tag             =   "4"
            Top             =   360
            Width           =   1665
         End
         Begin VB.TextBox TxtUsuario 
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
            Left            =   1935
            TabIndex        =   5
            Tag             =   "3"
            Top             =   780
            Width           =   2625
         End
         Begin VB.TextBox txtCod_Comite 
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
            Left            =   7410
            TabIndex        =   4
            Tag             =   "4"
            Top             =   795
            Width           =   1665
         End
         Begin XtremeSuiteControls.DateTimePicker dtpEnvio 
            Height          =   330
            Left            =   1920
            TabIndex        =   23
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.DateTimePicker dtpFinalEnvio 
            Height          =   330
            Left            =   3240
            TabIndex        =   24
            Top             =   360
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   582
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
         Begin VB.Label Label20 
            Caption         =   "No. Solicitud Bancos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5490
            TabIndex        =   10
            Top             =   420
            Width           =   2355
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha de Envio"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   9
            Top             =   390
            Width           =   1290
         End
         Begin VB.Label Label24 
            Caption         =   "Usuario responsable"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label Label7 
            Caption         =   "Comité"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   5490
            TabIndex        =   7
            Top             =   870
            Width           =   1740
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   19
         Top             =   1320
         Width           =   10095
         _Version        =   524288
         _ExtentX        =   17806
         _ExtentY        =   12938
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
         MaxCols         =   8
         SpreadDesigner  =   "FrmAF_CD_Cuentas.frx":34D6
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   9240
         TabIndex        =   20
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Liquidacion"
               Object.ToolTipText     =   "Ver Liquidacion"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAprobacion 
         Height          =   330
         Left            =   -66000
         TabIndex        =   21
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Aprobar"
               Object.ToolTipText     =   "Aprobar Desembolso"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rechazar"
               Object.ToolTipText     =   "Rechaza Desembolso"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Label Label18 
         Caption         =   "Delegado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "Asociados de Ajuste"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   22
         Top             =   1290
         Width           =   1665
      End
      Begin VB.Label Label18 
         Caption         =   "Autorización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   345
         TabIndex        =   18
         Top             =   1725
         Width           =   1185
      End
      Begin VB.Label Label13 
         Caption         =   "Asociados por Comité"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   315
         TabIndex        =   17
         Top             =   1290
         Width           =   1665
      End
      Begin VB.Label Label10 
         Caption         =   "Total de Asociados"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7290
         TabIndex        =   16
         Top             =   1290
         Width           =   1755
      End
      Begin VB.Label Label3 
         Caption         =   "Comite Principal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   15
         Top             =   810
         Width           =   1200
      End
      Begin VB.Label Label9 
         Caption         =   "Bancos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74655
         TabIndex        =   13
         Top             =   615
         Width           =   870
      End
      Begin VB.Label Label8 
         Caption         =   "Información"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74775
         TabIndex        =   11
         Top             =   2010
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Desembolso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   2520
         Width           =   1515
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7815
      Left            =   0
      TabIndex        =   28
      Top             =   1560
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   13785
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
      Item(0).Caption =   "Desembolso"
      Item(0).ControlCount=   32
      Item(0).Control(0)=   "tcDetalle"
      Item(0).Control(1)=   "btnBarra(0)"
      Item(0).Control(2)=   "btnBarra(1)"
      Item(0).Control(3)=   "Label1(1)"
      Item(0).Control(4)=   "txtComiteId"
      Item(0).Control(5)=   "txtComiteDesc"
      Item(0).Control(6)=   "txtAsociados"
      Item(0).Control(7)=   "txtAjusteAsoc"
      Item(0).Control(8)=   "txtAsocTotalAjustado"
      Item(0).Control(9)=   "Label1(2)"
      Item(0).Control(10)=   "Label1(3)"
      Item(0).Control(11)=   "Label1(4)"
      Item(0).Control(12)=   "Label1(5)"
      Item(0).Control(13)=   "Label1(6)"
      Item(0).Control(14)=   "Label1(7)"
      Item(0).Control(15)=   "cboJunta"
      Item(0).Control(16)=   "cboAutorizacion"
      Item(0).Control(17)=   "cboBanco"
      Item(0).Control(18)=   "cboMiembros"
      Item(0).Control(19)=   "Label1(8)"
      Item(0).Control(20)=   "Label1(9)"
      Item(0).Control(21)=   "cboCuenta"
      Item(0).Control(22)=   "cboEmite"
      Item(0).Control(23)=   "txtFechaLiq"
      Item(0).Control(24)=   "txtFechaRegistro"
      Item(0).Control(25)=   "txtMontoPagar"
      Item(0).Control(26)=   "txtNotas"
      Item(0).Control(27)=   "lblX(4)"
      Item(0).Control(28)=   "lblX(3)"
      Item(0).Control(29)=   "lblX(2)"
      Item(0).Control(30)=   "lblX(1)"
      Item(0).Control(31)=   "btnBarra(2)"
      Item(1).Caption =   "Adjuntos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lswAdjuntos"
      Item(1).Control(1)=   "btnAdjuntos"
      Item(2).Caption =   "Histórico"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lswHistorico"
      Item(2).Control(1)=   "btnExportar"
      Begin XtremeSuiteControls.ListView lswHistorico 
         Height          =   7455
         Left            =   -70000
         TabIndex        =   72
         Top             =   360
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1572864
         _ExtentX        =   18441
         _ExtentY        =   13150
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
      End
      Begin XtremeSuiteControls.ListView lswAdjuntos 
         Height          =   7455
         Left            =   -70000
         TabIndex        =   70
         Top             =   360
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1572864
         _ExtentX        =   18441
         _ExtentY        =   13150
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
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   375
         Index           =   0
         Left            =   8160
         TabIndex        =   30
         Top             =   0
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "FrmAF_CD_Cuentas.frx":3C11
      End
      Begin XtremeSuiteControls.TabControl tcDetalle 
         Height          =   3255
         Left            =   0
         TabIndex        =   29
         Top             =   2520
         Width           =   10335
         _Version        =   1572864
         _ExtentX        =   18230
         _ExtentY        =   5741
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
         Item(0).Caption =   "Actividades"
         Item(0).ControlCount=   4
         Item(0).Control(0)=   "Label1(10)"
         Item(0).Control(1)=   "cboActividadTipo"
         Item(0).Control(2)=   "Label1(11)"
         Item(0).Control(3)=   "lswActividades"
         Item(1).Caption =   "Consultas"
         Item(1).ControlCount=   3
         Item(1).Control(0)=   "vGridRefundiciones"
         Item(1).Control(1)=   "lblRefundiciones"
         Item(1).Control(2)=   "lblX(5)"
         Item(2).Caption =   "Cargos"
         Item(2).ControlCount=   3
         Item(2).Control(0)=   "vGridCargos"
         Item(2).Control(1)=   "lblCargos"
         Item(2).Control(2)=   "lblX(6)"
         Begin XtremeSuiteControls.ListView lswActividades 
            Height          =   2295
            Left            =   2160
            TabIndex        =   68
            Top             =   840
            Width           =   8055
            _Version        =   1572864
            _ExtentX        =   14208
            _ExtentY        =   4048
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
         End
         Begin XtremeSuiteControls.ComboBox cboActividadTipo 
            Height          =   330
            Left            =   2160
            TabIndex        =   53
            Top             =   480
            Width           =   3615
            _Version        =   1572864
            _ExtentX        =   6376
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
         Begin FPSpreadADO.fpSpread vGridRefundiciones 
            Height          =   2175
            Left            =   -69880
            TabIndex        =   54
            Top             =   480
            Visible         =   0   'False
            Width           =   10215
            _Version        =   524288
            _ExtentX        =   18018
            _ExtentY        =   3836
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
            MaxCols         =   6
            ScrollBars      =   2
            SpreadDesigner  =   "FrmAF_CD_Cuentas.frx":4331
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridCargos 
            Height          =   2055
            Left            =   -69760
            TabIndex        =   57
            Top             =   480
            Visible         =   0   'False
            Width           =   9615
            _Version        =   524288
            _ExtentX        =   16960
            _ExtentY        =   3625
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
            MaxCols         =   4
            ScrollBars      =   2
            SpreadDesigner  =   "FrmAF_CD_Cuentas.frx":49AA
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   615
            Index           =   11
            Left            =   240
            TabIndex        =   69
            Top             =   840
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Selecciones las actividades"
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
         Begin VB.Label lblX 
            Caption         =   "Total Cargos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   -64600
            TabIndex        =   59
            Top             =   2760
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Label lblCargos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   315
            Left            =   -62995
            TabIndex        =   58
            Top             =   2760
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.Label lblX 
            Caption         =   "Total Refundiciones"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   -67360
            TabIndex        =   56
            Top             =   2880
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.Label lblRefundiciones 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   315
            Left            =   -65275
            TabIndex        =   55
            Top             =   2880
            Visible         =   0   'False
            Width           =   2190
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   52
            Top             =   480
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo de Actividad"
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
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   31
         Top             =   0
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
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
         Appearance      =   17
         Picture         =   "FrmAF_CD_Cuentas.frx":4FDA
      End
      Begin XtremeSuiteControls.FlatEdit txtComiteId 
         Height          =   330
         Left            =   1680
         TabIndex        =   33
         Top             =   600
         Width           =   1095
         _Version        =   1572864
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtComiteDesc 
         Height          =   330
         Left            =   2760
         TabIndex        =   34
         Top             =   600
         Width           =   7215
         _Version        =   1572864
         _ExtentX        =   12726
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsociados 
         Height          =   330
         Left            =   1680
         TabIndex        =   35
         Top             =   960
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
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
      Begin XtremeSuiteControls.FlatEdit txtAjusteAsoc 
         Height          =   330
         Left            =   5280
         TabIndex        =   36
         Top             =   960
         Width           =   1095
         _Version        =   1572864
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAsocTotalAjustado 
         Height          =   330
         Left            =   8880
         TabIndex        =   37
         Top             =   960
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1931
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
      Begin XtremeSuiteControls.ComboBox cboJunta 
         Height          =   330
         Left            =   6360
         TabIndex        =   44
         Top             =   1320
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
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
      Begin XtremeSuiteControls.ComboBox cboAutorizacion 
         Height          =   330
         Left            =   1680
         TabIndex        =   45
         Top             =   1320
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
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
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   330
         Left            =   1680
         TabIndex        =   46
         Top             =   2040
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
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
      Begin XtremeSuiteControls.ComboBox cboMiembros 
         Height          =   330
         Left            =   1680
         TabIndex        =   47
         Top             =   1680
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
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
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   330
         Left            =   6360
         TabIndex        =   50
         Top             =   2040
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
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
      Begin XtremeSuiteControls.ComboBox cboEmite 
         Height          =   330
         Left            =   6360
         TabIndex        =   51
         Top             =   1680
         Width           =   3615
         _Version        =   1572864
         _ExtentX        =   6376
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
      Begin XtremeSuiteControls.FlatEdit txtFechaLiq 
         Height          =   330
         Left            =   5280
         TabIndex        =   60
         Top             =   6000
         Width           =   1575
         _Version        =   1572864
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
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaRegistro 
         Height          =   330
         Left            =   8520
         TabIndex        =   61
         Top             =   6000
         Width           =   1695
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
            Weight          =   400
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
      Begin XtremeSuiteControls.FlatEdit txtMontoPagar 
         Height          =   330
         Left            =   2160
         TabIndex        =   62
         Top             =   6000
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1290
         Left            =   2160
         TabIndex        =   63
         Top             =   6360
         Width           =   8055
         _Version        =   1572864
         _ExtentX        =   14208
         _ExtentY        =   2275
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAdjuntos 
         Height          =   330
         Left            =   -60160
         TabIndex        =   71
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1080
         _ExtentY        =   573
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "FrmAF_CD_Cuentas.frx":570B
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Left            =   -60160
         TabIndex        =   73
         Top             =   0
         Visible         =   0   'False
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
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
         Appearance      =   17
         Picture         =   "FrmAF_CD_Cuentas.frx":5794
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   375
         Index           =   2
         Left            =   9840
         TabIndex        =   75
         ToolTipText     =   "Descartar Solicitud"
         Top             =   0
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
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
         Appearance      =   17
         Picture         =   "FrmAF_CD_Cuentas.frx":58FE
      End
      Begin VB.Label lblX 
         Caption         =   "Monto a Pagar"
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
         Left            =   240
         TabIndex        =   67
         Top             =   6000
         Width           =   1200
      End
      Begin VB.Label lblX 
         Caption         =   "Fecha Liq."
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
         Left            =   4200
         TabIndex        =   66
         Top             =   6000
         Width           =   825
      End
      Begin VB.Label lblX 
         Caption         =   "Fecha Registro"
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
         Left            =   6960
         TabIndex        =   65
         Top             =   6000
         Width           =   1200
      End
      Begin VB.Label lblX 
         Caption         =   "Observaciones"
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
         Left            =   240
         TabIndex        =   64
         Top             =   6390
         Width           =   1200
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   5400
         TabIndex        =   49
         Top             =   2040
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuenta"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   8
         Left            =   5400
         TabIndex        =   48
         Top             =   1680
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   43
         Top             =   2040
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Banco"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   42
         Top             =   1680
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Delegado"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   41
         Top             =   1320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Autorización"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   6960
         TabIndex        =   40
         Top             =   960
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total de Asociados"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   39
         Top             =   960
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asociados de Ajuste"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asociados por"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Comité"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   450
      Left            =   1845
      TabIndex        =   26
      Top             =   960
      Width           =   1695
      _Version        =   1572864
      _ExtentX        =   2990
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   74
      Top             =   1440
      Width           =   10455
      _Version        =   1572864
      _ExtentX        =   18441
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   27
      Top             =   960
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "No. Operación"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Cuentas a Comités"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   4155
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmAF_CD_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean
Dim itmX As ListViewItem
Dim vOperacion As Long


Dim Miembro As String, vCodigo As String
Dim i As Integer, x As Integer

Dim vCuentaGasto As String
Dim vMontoActividad As Double
Dim vBacCargado As Boolean
Dim vFecha As String


Function fxConseConjunto()

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(noperacion),0) as Consecutivo from afi_cd_acticonjunta "
Call OpenRecordSet(rs, strSQL)
 fxConseConjunto = rs!Consecutivo + 1
rs.Close

End Function

Function fxNomComite(vUnidad As String)
   
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select U.descripcion from uprogramatica U right join afi_cd_comites_unidades A " _
            & "on U.codigo = A.cod_comite" _
            & " where A.cod_comite = '" & vUnidad & "'"
            Call OpenRecordSet(rs, strSQL)
   If rs.EOF Then
      fxNomComite = "No existe unidad definida en Comites y Delegados"
   Else
      fxNomComite = rs!Descripcion
   End If
rs.Close
End Function

Sub sbCargaBancos()
 
 strSQL = "select id_banco,descripcion from bancos "
           rs.Open strSQL, glogon.Conection, adOpenForwardOnly
  
 While rs.EOF = False
     cboBanco.AddItem (rs!Descripcion)
     cboBanco.ItemData(cboBanco.NewIndex) = rs!ID_BANCO
 rs.MoveNext
Wend
rs.Close
End Sub

Private Sub sbCargaBan()

cbobancApro.Clear

strSQL = "select distinct C.id_banco,B.descripcion" _
       & " from bancos B inner join afi_cd_cuentas C on B.id_banco = C.id_banco"
Call OpenRecordSet(rs, strSQL)
  
vPaso = True
Do While Not rs.EOF
     cbobancApro.AddItem (rs!Descripcion)
     cbobancApro.ItemData(cbobancApro.NewIndex) = rs!ID_BANCO
 rs.MoveNext
Loop
rs.Close
vPaso = False

vBacCargado = True
vGrid.MaxRows = 0

End Sub

Private Function fxConsecutivo() As Long

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(noperacion),0) as Consecutivo from afi_cd_cuentas "
Call OpenRecordSet(rs, strSQL)
fxConsecutivo = rs!Consecutivo + 1
rs.Close

End Function

Private Sub sbConsulta(vSeleccion As Integer)
 Dim strSQL As String, rs As New ADODB.Recordset
 Dim itmX As ListItem
 Dim i As Integer
 
strSQL = "select C.noperacion,rtrim(P.cod_Comite) + ' - ' + P.descripcion as Comite,C.tesoreria_nsolicitud" _
        & ",C.tesoreria_fecha,C.tesoreria_usuario,C.liquida_fecha,C.Cedula,S.Nombre,C.Cuenta,C.id_banco " _
        & " from afi_cd_comites P right join afi_cd_cuentas C on P.cod_comite  = C.cod_comite" _
        & " left join Socios S on S.cedula = C.cedula " _
        & " Where C.estado = 'T'" _
        & " and C.tesoreria_fecha between '" & Format(dtpEnvio.Value, "yyyymmdd") & " 00:00:00' and '" & Format(dtpFinalEnvio.Value, "yyyymmdd") & " 23:59:59'"
            
If Len(Trim(TxtUsuario.Text)) > 0 Then
   strSQL = strSQL & " and C.tesoreria_usuario like '%" & TxtUsuario.Text & "%'"
End If

If Len(Trim(txtCod_Comite.Text)) > 0 Then
   strSQL = strSQL & " and C.cod_comite like '%" & txtCod_Comite.Text & "%'"
End If


lswinfoenvio.ListItems.Clear
 
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 
 While Not rs.EOF
      Set itmX = lswinfoenvio.ListItems.Add(, , rs!Noperacion)
      itmX.SubItems(1) = rs!Comite
      itmX.SubItems(2) = rs!Cedula
      itmX.SubItems(3) = rs!Nombre
      itmX.SubItems(4) = IIf(Not IsNull(rs!TESORERIA_NSOLICITUD), rs!TESORERIA_NSOLICITUD, "Sin No.Solicitud")
      itmX.SubItems(5) = IIf(Not IsNull(rs!TESORERIA_FECHA), rs!TESORERIA_FECHA, "Sin Fecha")
      itmX.SubItems(6) = IIf(Not IsNull(rs!tesoreria_usuario), rs!tesoreria_usuario, "Sin Usuario")
      itmX.SubItems(7) = 0
      itmX.SubItems(8) = rs!LIQUIDA_FECHA & ""
      rs.MoveNext
 Wend
 rs.Close

For i = 1 To lswinfoenvio.ListItems.Count
   strSQL = "select isnull(sum(Monto),0) as 'Monto' from afi_cd_cuentas_actividades where noperacion = " & lswinfoenvio.ListItems.Item(i)
   Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      lswinfoenvio.ListItems.Item(i).SubItems(7) = Format(rs!Monto, "Standard")
    End If
   rs.Close
Next i

End Sub

Sub sbLlamaComite(vComite As String)

Dim strSQL As String
Dim rs As New ADODB.Recordset
              
     cboMiembros.Clear
     Call cboActividadTipo_Click
                
     strSQL = "select N.cedula, S.nombre, N.cod_comite, C.descripcion" _
            & " from afi_cd_comites C left join afi_cd_nombramientos N on C.cod_comite = N.cod_comite" _
            & " inner join socios S on S.cedula = N.cedula " _
            & " where N.cod_comite = '" & vComite & "' and N.APL_DESEMBOLSOS = 1"
     Call OpenRecordSet(rs, strSQL)
     
     If Not rs.EOF Then
        txtComiteDesc.Text = Trim(rs!Descripcion)
     Else
        MsgBox "No se cuenta con miembro asigando el desembolso!!"
     End If

     Do While Not rs.EOF
       cboMiembros.AddItem rs!Nombre
       cboMiembros.ItemData(cboMiembros.ListCount - 1) = CStr(rs!Cedula)
       rs.MoveNext
     Loop
     rs.Close
     
End Sub

Private Sub sbMjunta()

cboJunta.Clear
strSQL = "select cod_director as 'IdX', Nombre as 'ItmX' from afi_cd_directores"
Call sbCbo_Llena_New(cboJunta, strSQL, False, True)

End Sub

Private Sub sbLimpiar()
     
     vOperacion = 0
     txtOperacion.Text = ""
     
     txtComiteId.Text = ""
     
     cboMiembros.Clear
     
     cboCuenta.Clear

     txtNotas.Text = ""
     txtComiteDesc.Text = ""
     
     txtAsociados.Text = "0"
     txtAjusteAsoc.Text = 0
     txtAsocTotalAjustado.Text = "0"
     txtMontoPagar.Text = 0
     
     lblRefundiciones.Caption = 0
     lblCargos.Caption = 0
     
     txtFechaLiq.Text = ""
     txtFechaRegistro.Text = ""
     txtNotas.Text = ""

     lblRefundiciones.Caption = 0
     lblCargos.Caption = 0

'     txtComiteId.SetFocus

End Sub


Private Sub sbTiposActividadActiva()

txtFechaLiq.Text = ""
txtMontoPagar.Text = 0

End Sub



Private Sub btnAdjuntos_Click()

'If txtOperacion.Text <> "" Then
 gGA.Modulo = "CD_01"
 gGA.Llave_01 = txtComiteId.Text
 gGA.Llave_02 = txtOperacion.Text
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
'End If

End Sub


Private Sub sbDescartar()
Dim i As Integer

On Error GoTo vError

If txtOperacion.Text = "" Or Not IsNumeric(txtOperacion.Text) Then
    MsgBox "Consulte una Cuenta Primero!", vbExclamation
    Exit Sub
End If

i = MsgBox("Esta Seguro que desea Poner como Descartada esta Solicitud de Cuenta?", vbYesNo)
If i = vbYes Then
    
   strSQL = "exec spAFI_CD_Cuenta_Descarta " & txtOperacion.Text & ", '" & glogon.Usuario & "'"
   Call OpenRecordSet(rs, strSQL)
   
   Me.MousePointer = vbDefault
   If rs!Pass = 1 Then
        MsgBox "Esta cuenta ha sido descartada!", vbInformation
        Call sbLimpiar
   Else
        MsgBox rs!Mensaje, vbExclamation
   End If

End If

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBarra_Click(Index As Integer)

Select Case Index
  Case 0
        Call sbLimpiar
  Case 1
        Call sbAplicar
        
'        Call sbLimpiar
  Case 2
        Call sbDescartar

End Select

End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lswHistorico, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboActividadTipo_Click()
If vPaso Then Exit Sub

'spAFI_CD_Actividades_List(@Tipo varchar(10) = 'T', @TotalAsoc int = 0, @Operacion int = 0, @Comite varchar(10))

tcDetalle(0).Selected = True

strSQL = "exec spAFI_CD_Actividades_List '" & cboActividadTipo.ItemData(cboActividadTipo.ListIndex) & "', " & txtAsocTotalAjustado.Text _
       & ", " & vOperacion & ", " & txtComiteId.Text
Call OpenRecordSet(rs, strSQL)

vPaso = True

With lswActividades.ListItems
    .Clear
    
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Cod_Actividad)
            itmX.SubItems(1) = RTrim(rs!Descripcion)
            itmX.SubItems(2) = Format(rs!Monto, "Standard")
            itmX.SubItems(3) = rs!Tipo
            
            If rs!Asignado = 1 Then
                itmX.Checked = True
            End If
        rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

End Sub



Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & cboMiembros.ItemData(cboMiembros.ListIndex) & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:

End Sub


Private Sub cboMiembros_Click()
 
Dim strSQL As String

On Error GoTo vError

vPaso = True

cboBanco.Clear
For i = 0 To cboMiembros.ListCount
    If cboMiembros.ListIndex = i Then
       Miembro = cboMiembros.ItemData(cboMiembros.ListIndex)
    End If
Next i


strSQL = "select Bn.ID_BANCO as 'IdX', Bn.DESCRIPCION as 'ItmX' " _
       & " from SYS_CUENTAS_BANCARIAS Cta inner join Tes_Bancos Bn on Cta.COD_BANCO = Bn.COD_GRUPO" _
       & "  Where Cta.IDENTIFICACION  = '" & Miembro & "'" _
       & "  group by Bn.ID_BANCO , Bn.DESCRIPCION"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

cboBanco.SetFocus

vPaso = False

Exit Sub

vError:
 vPaso = False
End Sub

'Private Sub CmdAplicar_Click()
' Call sbAplicar
'End Sub

Private Sub CmdImp_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String, xRemesa As String

On Error GoTo vError
Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Consulta de Envio de Operaciones a Tesorería"
 
 .Connect = glogon.ConectRPT
 .ReportFileName = SIFGlobal.fxPathReportes("Comites_ConsulaOperacion.rpt")
     
  Select Case True
   Case TxtSolicitud.Text <> ""
    strSQL = "{afi_cd_cuentas.tesoreria_nsolicitud} = " & TxtSolicitud.Text & " and "
    txtCod_Comite.Text = ""
    TxtUsuario.Text = ""
    vSubTitulo = "No. Solicitud " & TxtSolicitud.Text & ""
   Case txtCod_Comite.Text <> ""
    strSQL = "{afi_cd_cuentas.id_pricomite} = '" & txtCod_Comite.Text & "' and "
    TxtSolicitud.Text = ""
    TxtUsuario.Text = ""
    vSubTitulo = "Comité: " & txtCod_Comite.Text & ""
   Case TxtUsuario.Text <> ""
    strSQL = "{afi_cd_cuentas.tesoreria_usaurio} = " & TxtUsuario.Text & " and "
    TxtSolicitud.Text = ""
    txtCod_Comite.Text = ""
    vSubTitulo = "Usuario: " & TxtUsuario.Text & ""
  End Select
    
 strSQL = strSQL & "cdate({afi_cd_cuentas.tesoreria_fecha}) in Date(" & Format(dtpEnvio.Value, "yyyy,mm,dd")
 strSQL = strSQL & ") to Date (" & Format(dtpFinalEnvio.Value, "yyyy,mm,dd") & ")"
 .SelectionFormula = strSQL
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(vFecha, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='Consulta de Envio de Operaciones a Tesorería'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Cmdimprimir_Click()

End Sub


Private Sub cmdLimpiar_Click()
 txtComiteId.Text = ""
 TxtSolicitud.Text = ""
 TxtUsuario.Text = ""
 dtpEnvio.Value = Format(vFecha, "dd/mm/yyyy")
 dtpFinalEnvio.Value = Format(vFecha, "dd/mm/yyyy")
 lswinfoenvio.ListItems.Clear
 TxtUsuario.SetFocus
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case 48 To 57, 8
  Case 13
   
  Case Else
   KeyAscii = 0
End Select


End Sub




Private Sub DtpEnvio_Change()
 Call sbConsulta(1)
End Sub

Private Sub dtpFinalEnvio_Change()
 Call sbConsulta(1)
End Sub

Private Sub Form_Activate()
 vModulo = 40
End Sub

Private Sub Form_Load()
   
 vModulo = 40
   
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
   
 vFecha = Format(fxFechaServidor, "yyyymmdd")
 
 vPaso = True
 strSQL = "select CodTipoCuenta as 'IdX', NombreTipoCuenta as 'itmX' from AFI_CD_TIPO_CUENTA where Activo = 1"
 Call sbCbo_Llena_New(cboEmite, strSQL, False, True)
 
 strSQL = "select CodTipoActividad as 'Idx', NombreTipoActividad as 'ItmX' from AFI_CD_TIPO_ACTIVIDAD where Activo = 1"
 Call sbCbo_Llena_New(cboActividadTipo, strSQL, False, True)
  
 strSQL = "select CodTipoAprobacion as 'Idx', NombreTipoAprobacion as 'ItmX' from AFI_CD_TIPO_APROBACION where Activo = 1"
 Call sbCbo_Llena_New(cboAutorizacion, strSQL, False, True)
  
' strSQL = "select CodTipoAprobacion as 'Idx', NombreTipoAprobacion as 'ItmX' from AFI_CD_TIPO_APROBACION where Activo = 1"
' Call sbCbo_Llena_New(cboJunta, strSQL, False, True)
  
  
 vPaso = False
 
 With lswAdjuntos.ColumnHeaders
    .Clear
    .Add , , "Id Archivo", 3500
    .Add , , "Nombre", 3000
    .Add , , "Tipo", 3000
    .Add , , "Ext", 1000, vbCenter
    .Add , , "R. Fecha", 2500
    .Add , , "R. Usuario", 2500, vbCenter
 End With
 
 
 With lswHistorico.ColumnHeaders
    .Clear
    .Add , , "Id Histórico", 1500
    .Add , , "R. Fecha", 2500
    .Add , , "R. Usuario", 2500, vbCenter
    .Add , , "Estado", 2100, vbCenter
    .Add , , "Proceso", 2100, vbCenter
    .Add , , "Nota", 3000
 End With
 
 With lswActividades.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Actividad", 3200
    .Add , , "Monto", 1200, vbRightJustify
    .Add , , "Tipo", 2200
 End With
 lswActividades.Checkboxes = True
 
 vBacCargado = False
 
 SSTabCuentas.Tab = 0
 
 Call sbCargaBancos
 
 
 dtpEnvio.Value = fxFechaServidor
 dtpFinalEnvio.Value = dtpEnvio.Value
 txtFechaRegistro.Text = Format(dtpEnvio.Value, "dd/mm/yyyy")
 
' tlbPrincipal.ImageList = frmContenedor.imgToolbarIcons
' tlbPrincipal.Buttons.Item(1).Image = 2
' tlbPrincipal.Buttons.Item(3).Image = 9
' tlbPrincipal.Buttons.Item(4).Image = 4
  
' tlbAprobacion.ImageList = frmContenedor.imgToolbarIcons 'frmContenedor.imgToolbarIcons03
' tlbAprobacion.Buttons.Item(1).Image = 3
' tlbAprobacion.Buttons.Item(2).Image = 4
 

 Call sbTiposActividadActiva
 Call sbCargaCargos

 Call sbLimpiar
 
 If GLOBALES.gTag <> Empty Then
   txtComiteId.Text = GLOBALES.gTag
   Call cboMiembros.Clear
   Call sbCalMiembros
   Call sbLlamaComite(txtComiteId.Text)
   Call sbCargaLiquidaciones
   GLOBALES.gTag = Empty
 End If

End Sub

Private Sub sbCalMiembros()

On Error GoTo vError

   
strSQL = "select count(*) as 'Cantidad' from socios" _
       & " where EstadoActual = 'S' and cod_departamento in(select Codigo_UP from Afi_CD_Comites_Unidades where cod_comite = '" & txtComiteId.Text & "')"
Call OpenRecordSet(rs, strSQL)

txtAsociados.Text = rs!Cantidad
txtAsocTotalAjustado.Text = rs!Cantidad - CCur(txtAjusteAsoc.Text)

rs.Close


Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  
End Sub



Private Sub lswActividades_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub


If Item.Checked Then
    txtMontoPagar.Text = Format(CCur(txtMontoPagar.Text) + CCur(Item.SubItems(2)), "Standard")
Else
    txtMontoPagar.Text = Format(CCur(txtMontoPagar.Text) - CCur(Item.SubItems(2)), "Standard")
End If


End Sub

Private Sub sbBitacora_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_CD_Cuenta_Bitacora " & vOperacion
Call OpenRecordSet(rs, strSQL)

lswHistorico.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswHistorico.ListItems.Add(, , rs!IdRegistro)
      itmX.SubItems(1) = rs!RegistroFecha
      itmX.SubItems(2) = rs!RegistroUsuario
      itmX.SubItems(3) = rs!NombreEstado
      itmX.SubItems(4) = rs!NombreTipoProceso
      itmX.SubItems(5) = rs!Nota
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAdjuntos_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_CD_Cuenta_Adjuntos " & vOperacion
Call OpenRecordSet(rs, strSQL)

lswAdjuntos.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswAdjuntos.ListItems.Add(, , rs!IdArchivoAdjunto)
      itmX.SubItems(1) = rs!NombreArchivo
      itmX.SubItems(2) = rs!NombreTipoArchivo
      itmX.SubItems(3) = rs!Nota
      itmX.SubItems(4) = rs!RegistroFecha
      itmX.SubItems(5) = rs!RegistroUsuario
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCuenta_Load(pOperacion As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain(0).Selected = True
txtComiteId.SetFocus

strSQL = "exec spAFI_CD_Cuenta_Load " & pOperacion
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    
'T A.Noperacion, A.CEDULA, S.Nombre
'        , A.COD_COMITE , Co.DESCRIPCION as 'COMITE_DESC'
'        , A.COD_DIRECTOR, Dr.DESCRIPCION as 'DIRECTOR_DESC'
'        , A.MONTO, A.SALDO, A.ESTADO, A.PROCESO ,  A.APRUEBA, A.TIPO
'        , Te.NombreEstado as 'ESTADO_DESC'
'        , Tp.NombreTipoProceso as 'PROCESO_DESC'
'        , Td.NombreTipoCuenta as 'TIPO_DESC'
'        , ISNULL(Ta.NombreTipoAprobacion,'') as 'APROBACION_DESC'
'        , A.LIQUIDABLE
'        , A.LIQUIDA_FECHA, A.LIQUIDA_USUARIO, A.LIQUIDA_VENCE
'        , A.REGISTRO_FECHA, A.REGISTRO_USUARIO
'        , A.CUENTA, A.ID_BANCO, B.DESCRIPCION as 'BANCO_DESC'
'        , A.TESORERIA_NSOLICITUD , A.TESORERIA_FECHA, A.TESORERIA_USUARIO , A.COD_REMESA
'        , A.NOTAS, a.AJUSTE_ASOC, A.CANT_ASOCIADOS, a.MONTO_CARGOS , a.MONTO_REFUNDE
'        , A.ACTIVA_FECHA, A.ACTIVA_USUARIO
        
     vOperacion = pOperacion
     txtOperacion.Text = pOperacion
     
     txtComiteId.Text = rs!cod_comite
     txtComiteDesc.Text = rs!COMITE_DESC
     
     Call sbCboAsignaDato(cboMiembros, rs!Nombre, True, rs!Cedula)
     Call sbCboAsignaDato(cboBanco, rs!Banco_Desc, True, rs!ID_BANCO)
     Call sbCboAsignaDato(cboCuenta, rs!Cuenta_Desc, True, rs!Cuenta)
     Call sbCboAsignaDato(cboAutorizacion, rs!Aprobacion_Desc, True, rs!Aprueba)
     Call sbCboAsignaDato(cboEmite, rs!Tipo_Desc, True, rs!Tipo)

     txtNotas.Text = rs!NOTAS & ""
     
     txtAsociados.Text = rs!Cant_Asociados
     txtAjusteAsoc.Text = rs!Ajuste_Asoc
     txtAsocTotalAjustado.Text = rs!Asoc_Total
     txtMontoPagar.Text = Format(rs!Monto, "Standard")
     
     lblRefundiciones.Caption = Format(rs!MONTO_REFUNDE, "Standard")
     lblCargos.Caption = Format(rs!MONTO_CARGOS, "Standard")
     
     txtFechaLiq.Text = rs!LIQUIDA_FECHA & ""
     txtFechaRegistro.Text = rs!Registro_Fecha & ""
    
     Call cboActividadTipo_Click
Else
    Call sbLimpiar
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
  
 If Item.Index > 0 And vOperacion = 0 Then
    tcMain(0).Selected = True
 End If
  
Select Case Item.Index
  Case 0
  Case 1 'Adjuntos
    Call sbAdjuntos_Load
  Case 2 'Historico
    Call sbBitacora_Load
End Select
  
  
'  If Item.Index = 1 Then
'   If vBacCargado = False Then
'    Call sbCargaBan
'   End If
'  End If
End Sub



Private Sub txtAjusteAsoc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 txtAsocTotalAjustado.SetFocus
End If

End Sub

'Private Sub tlbConsulta_ButtonClick(ByVal Button As MSComctlLib.Button)
' Dim vTitulo As String, vSubTitulo As String, vFiltro As String
' Dim strSQL As String, vTipoUser As String, xRemesa As String
'
'
'Select Case UCase(Button.Key)
'  Case "REPORTE"
'  On Error GoTo vError
'  Me.MousePointer = vbHourglass
'
'
'            vSubTitulo = ""
'            vFiltro = ""
'            strSQL = ""
'
'
'            With frmContenedor.Crt
'             .Reset
'             .WindowShowGroupTree = True
'             .WindowShowPrintSetupBtn = True
'             .WindowShowRefreshBtn = True
'             .WindowShowSearchBtn = True
'             .WindowState = crptMaximized
'             .WindowTitle = "Consulta de Envio de Operaciones a Tesorería"
'
'             .Connect = glogon.ConectRPT
'             .ReportFileName = SIFGlobal.fxPathReportes("Comites_OperacionesEnviadasTesoreria.rpt")
'
'              Select Case True
'               Case TxtSolicitud.Text <> ""
'                strSQL = "{vAFI_CD_OperacionesEnviadasTesoreria.tesoreria_nsolicitud} = " & TxtSolicitud.Text & " "
'                txtCod_comite.Text = ""
'                TxtUsuario.Text = ""
'                vSubTitulo = "No. Solicitud " & TxtSolicitud.Text & ""
'
'               Case txtCod_comite.Text <> ""
'                strSQL = "{vAFI_CD_OperacionesEnviadasTesoreria.cod_comite} = '" & txtCod_comite.Text & "' and "
'                TxtSolicitud.Text = ""
'                TxtUsuario.Text = ""
'                vSubTitulo = "Comité: " & txtCod_comite.Text & ""
'
'               Case TxtUsuario.Text <> ""
'                strSQL = "{vAFI_CD_OperacionesEnviadasTesoreria.tesoreria_usuario} = '" & TxtUsuario.Text & "' and "
'                txtCod_comite.Text = ""
'                TxtSolicitud.Text = ""
'                vSubTitulo = "Usuario: " & TxtUsuario.Text & ""
'              End Select
'
'            If TxtSolicitud.Text = Empty Then
'             strSQL = strSQL & "cdate({vAFI_CD_OperacionesEnviadasTesoreria.tesoreria_fecha}) in Date(" & Format(dtpEnvio.Value, "yyyy/mm/dd")
'             strSQL = strSQL & " ) to Date (" & Format(dtpFinalEnvio.Value, "yyyy/mm/dd") & ")"
'            End If
'             .SelectionFormula = strSQL
'
'             .Formulas(0) = "fxFecha= '" & Format(fxFechaServidor, "dd/mm/yyyy") & "' "
'             .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
'             .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
'             .Formulas(3) = "fxTitulo='Consulta de Envio de Operaciones a Tesorería'"
'             .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
'             .Action = 1
'             '.PrintReport
'
'            End With
'
'            Me.MousePointer = vbDefault
'            Exit Sub
'
'vError:
'             Me.MousePointer = vbDefault
'             MsgBox Err.Description, vbCritical
'  Case "NUEVO"
'           txtCod_comite.Text = ""
'        TxtSolicitud.Text = ""
'        TxtUsuario.Text = ""
'        dtpEnvio.Value = Format(vFecha, "dd/mm/yyyy")
'        dtpFinalEnvio.Value = Format(vFecha, "dd/mm/yyyy")
'        lswinfoenvio.ListItems.Clear
'        TxtUsuario.SetFocus
' End Select
'End Sub

''Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
'' Dim vTitulo As String, vSubTitulo As String, vFiltro As String
'' Dim strSQL As String, vTipoUser As String, xRemesa As String
''
''
''Select Case UCase(Button.Key)
''  Case "REPORTE"
''  On Error GoTo vError
''  Me.MousePointer = vbHourglass
''
''
''            vSubTitulo = ""
''            vFiltro = ""
''            strSQL = ""
''
''
''            With frmContenedor.Crt
''             .Reset
''             .WindowShowGroupTree = True
''             .WindowShowPrintSetupBtn = True
''             .WindowShowRefreshBtn = True
''             .WindowShowSearchBtn = True
''             .WindowState = crptMaximized
''             .WindowTitle = "Consulta de Envio de Operaciones a Tesorería"
''
''             .Connect = glogon.ConectRPT
''             .ReportFileName = SIFGlobal.fxPathReportes("Comites_ConsulaOperacion.rpt")
''
''              Select Case True
''               Case TxtSolicitud.Text <> ""
''                strSQL = "{afi_cd_cuentas.tesoreria_nsolicitud} = " & TxtSolicitud.Text & " and "
''                txtCod_Comite.Text = ""
''                TxtUsuario.Text = ""
''                vSubTitulo = "No. Solicitud " & TxtSolicitud.Text & ""
''               Case txtCod_Comite.Text <> ""
''                strSQL = "{afi_cd_cuentas.id_pricomite} = '" & txtCod_Comite.Text & "' and "
''                TxtSolicitud.Text = ""
''                TxtUsuario.Text = ""
''                vSubTitulo = "Comité: " & txtCod_Comite.Text & ""
''               Case TxtUsuario.Text <> ""
''                strSQL = "{afi_cd_cuentas.tesoreria_usaurio} = " & TxtUsuario.Text & " and "
''                TxtSolicitud.Text = ""
''                txtCod_Comite.Text = ""
''                vSubTitulo = "Usuario: " & TxtUsuario.Text & ""
''              End Select
''
''             strSQL = strSQL & "cdate({afi_cd_cuentas.tesoreria_fecha}) in Date(" & Format(dtpEnvio.Value, "yyyy,mm,dd")
''             strSQL = strSQL & ") to Date (" & Format(dtpFinalEnvio.Value, "yyyy,mm,dd") & ")"
''             .SelectionFormula = strSQL
''
''             .Formulas(0) = "fxFecha='FECHA: " & Format(vFecha, "dd/mm/yyyy") & "'"
''             .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
''             .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
''             .Formulas(3) = "fxTitulo='Consulta de Envio de Operaciones a Tesorería'"
''             .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
''
''             .PrintReport
''
''            End With
''
''            Me.MousePointer = vbDefault
''            Exit Sub
''
''vError:
''             Me.MousePointer = vbDefault
''             MsgBox Err.Description, vbCritical
''  Case "NUEVO"
''           txtCod_Comite.Text = ""
''        TxtSolicitud.Text = ""
''        TxtUsuario.Text = ""
''        dtpEnvio.Value = Format(vFecha, "dd/mm/yyyy")
''        dtpFinalEnvio.Value = Format(vFecha, "dd/mm/yyyy")
''        lswinfoenvio.ListItems.Clear
''        TxtUsuario.SetFocus
'' End Select
''End Sub

''Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
''Select Case UCase(Button.Key)
''  Case "APROBAR"
''    Call sbAprobacion
''    Call sbCargPago
''End Select
''End Sub


Private Sub txtAjusteAsoc_LostFocus()

On Error GoTo vError

txtAsocTotalAjustado.Text = CCur(txtAsociados.Text) - CCur(txtAjusteAsoc.Text)
Call cboActividadTipo_Click

Exit Sub

vError:

End Sub

Private Sub txtCod_comite_KeyPress(KeyAscii As Integer)
 Select Case KeyAscii
          Case 48 To 57, 8
          Case 13
             Call sbConsulta(4)
          Case Else
           KeyAscii = 0
      End Select
End Sub

Private Sub txtComiteId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_COMITE, DESCRIPCION from AFI_CD_COMITES"
       gBusquedas.Filtro = " AND ACTIVO = 1"
       frmBusquedas.Show vbModal
       vCodigo = gBusquedas.Resultado
       txtComiteId.Text = gBusquedas.Resultado
       Call txtComiteId_KeyPress(vbKeyReturn)
End If
End Sub

Private Sub txtComiteId_KeyPress(KeyAscii As Integer)
  
 ' cargar actividades y sus caracteristICAS
  
  Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        Call cboMiembros.Clear
        Call cboActividadTipo_Click
        Call sbCalMiembros
        Call sbLlamaComite(txtComiteId.Text)
        Call sbCargaLiquidaciones
      Case Else
       KeyAscii = 0
  End Select
End Sub


Private Sub sbAplicar()

Dim vCodDir As Integer, vInfoMonto As Currency
Dim vTipo As String
Dim vNumOperacion As Integer
Dim vCuentaBanco As Integer, x As Integer
Dim vCtaDelegado As String
Dim vAprueba As String
Dim vActividad As Integer
Dim Op As Integer
Dim strLinea(10) As String
Dim vMonto As Double
Dim strComite As String
Dim vTipoDoc As String

vFecha = Format(fxFechaServidor, "yyyymmdd")

strSQL = ""

If Trim(txtComiteId) = "" Then strSQL = "Comite" & vbCrLf
If Trim(cboMiembros.Text) = "" Then strSQL = strSQL & "Necesita Agregar Miembro" & vbCrLf
If Trim(cboBanco.Text) = "" Then strSQL = strSQL & "Debe seleccionar un banco" & vbCrLf

If Trim(txtMontoPagar) = 0 Then strSQL = strSQL & "Monto" & vbCrLf
If Trim(txtNotas.Text) = "" Then strSQL = strSQL & "Necesita agregar una descripción en Observaciones" & vbCrLf
If cboEmite.ItemData(cboEmite.ListIndex) = "T" And cboCuenta.Text = "" Then strSQL = strSQL & "La cuenta del Banco vacia" & vbCrLf

If strSQL <> "" Then
   MsgBox strSQL, vbInformation, "Faltan Los Siguientes Datos:"
   Exit Sub
End If

On Error GoTo vError:

vTipo = cboEmite.ItemData(cboEmite.ListIndex)
vCuentaBanco = cboBanco.ItemData(cboBanco.ListIndex)
vCtaDelegado = cboCuenta.ItemData(cboCuenta.ListIndex)

vAprueba = cboAutorizacion.ItemData(cboAutorizacion.ListIndex)
vCodDir = cboJunta.ItemData(cboJunta.ListIndex)

strLinea(4) = "Apr: " & cboAutorizacion.Text

vNumOperacion = fxConsecutivo
txtOperacion.Text = vNumOperacion

strSQL = "insert afi_cd_cuentas(noperacion,cod_comite,cedula,registro_fecha " _
       & ", registro_usuario,estado,tipo,cuenta,id_banco,notas,aprueba,cod_director,PROCESO,AJUSTE_ASOC,MONTO,MONTO_REFUNDE, MONTO_CARGOS, SALDO, CANT_ASOCIADOS, GuidId)" _
       & " values(" & vNumOperacion & ",'" & txtComiteId.Text & "', '" & cboMiembros.ItemData(cboMiembros.ListIndex) & "'" _
       & ",getdate()" _
       & ", '" & glogon.Usuario & "', 'S','" & vTipo & "','" & vCtaDelegado & "'," & vCuentaBanco & ",'" & txtNotas.Text & "' " _
       & ", '" & vAprueba & "'," & IIf(vCodDir = 0, 1, vCodDir) & ",'T'," & txtAjusteAsoc.Text & "," & CCur(txtMontoPagar.Text) & "," & CCur(lblRefundiciones.Caption) & "," & CCur(lblCargos.Caption) & "," & CCur(txtMontoPagar.Text) & ", " & txtAsociados.Text & ", NEWID() )"
Call ConectionExecute(strSQL)


With lswActividades.ListItems
    strSQL = ""
    For i = 1 To .Count
      If .Item(i).Checked Then
           strSQL = strSQL & Space(10) & "insert into afi_cd_cuentas_actividades (COD_ACTIVIDAD, NOPERACION, MONTO) " _
                  & " values (" & .Item(i) & ", " & vNumOperacion & ", " & CCur(.Item(i).SubItems(2)) & ")"
      End If
    Next i
    
    If Len(strSQL) > 0 Then
       Call ConectionExecute(strSQL)
    End If
    
End With

txtOperacion.Text = vNumOperacion

Call sbGuardaRefundicion(vNumOperacion)
Call sbGuardaCargos(vNumOperacion)

MsgBox "Solicitud Registrada: Proceda a la Aprobación!", vbInformation, "Información"

Call sbLimpiar

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtComiteDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select distinct cod_comite,U.descripcion from afi_cd_comites_unidades A " _
                             & "left join uprogramatica U on A.cod_comite = U.codigo "
       frmBusquedas.Show vbModal
       vCodigo = gBusquedas.Resultado
       txtComiteId.Text = gBusquedas.Resultado
       Call txtComiteId_KeyPress(vbKeyReturn)
 
End If
End Sub



Private Sub TxtNotas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub




Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "No.Operación"
  gBusquedas.Col2Name = "Id Comité"
  gBusquedas.Col3Name = "Cédula"
  gBusquedas.Columna = "NOperacion"
  gBusquedas.Orden = "NOperacion"
  gBusquedas.Consulta = "select NOperacion, Cod_Comite, Cedula, Saldo from afi_cd_Cuentas"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtOperacion.Text = gBusquedas.Resultado
    Call sbCuenta_Load(txtOperacion.Text)
  End If
End If

End Sub

Public Sub sbConsulta_Externa(pOperacion As Long)

 Call sbCuenta_Load(pOperacion)
  
End Sub

Private Sub txtOperacion_LostFocus()
    If IsNumeric(txtOperacion.Text) Then
        Call sbCuenta_Load(txtOperacion.Text)
    End If
End Sub

Private Sub TxtSolicitud_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case 48 To 57, 8
     Case 13
       Call sbConsulta(3)
     Case Else
       KeyAscii = 0
End Select
End Sub

Private Sub TxtUsuario_KeyPress(KeyAscii As Integer)
  KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TxtUsuario_LostFocus()
 Call sbConsulta(2)
End Sub

Private Function fxCDParametros(vParametro) As String
On Error GoTo vError
Dim rsX As New ADODB.Recordset

With glogon
 .strSQL = "select valor from AFI_CD_PARAMETROS where cod_parametro = '" & vParametro & "'"
 rsX.Open .strSQL, .Conection, adOpenStatic
   fxCDParametros = rsX!Valor
 rsX.Close
End With

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function

Private Sub sbCargaLiquidaciones()
    
With vGridRefundiciones
   .MaxRows = 1
   For i = 1 To .MaxCols
     .Col = i
     .Text = ""
   Next i

   'Carga los datos de la liquidaciones pendientes
   strSQL = "select A.noperacion,C.notas,sum(A.monto)as Monto,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha " _
             & "from afi_cd_cuentas C inner join  afi_cd_cuentas_actividades A " _
             & "on C.noperacion = A.noperacion " _
             & "where C.cod_comite = '" & txtComiteId.Text & "' and PROCESO='T' " _
             & "group by C.notas,A.noperacion,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha"
   Call OpenRecordSet(rs, strSQL)
    
   Do While Not rs.EOF
       .Row = .MaxRows
       
       .Col = 2
       .Text = rs!Noperacion
       
       .Col = 3
       .Text = rs!NOTAS
       
       .Col = 4
       .Text = Format(rs!Monto, "Standard")
       
       .Col = 5
       .Text = IIf(IsNull(rs!TESORERIA_NSOLICITUD), 0, rs!TESORERIA_NSOLICITUD)
       
       .Col = 6
       .Text = Format(rs!LIQUIDA_FECHA, "dd/mm/yyyy")
    
       .MaxRows = .MaxRows + 1
       rs.MoveNext
   Loop
    
   rs.Close
   .MaxRows = .MaxRows - 1
   
End With
 
End Sub

Private Sub sbCargaCargos()
Dim i As Integer
On Error GoTo vError

With vGridCargos
   .MaxRows = 1
   For i = 1 To .MaxCols
     .Col = i
     .Text = ""
   Next i
        
   strSQL = "Select CODIGO, DESCRIPCION from AFI_CD_CARGOS where ESTADO = 1"
   Call OpenRecordSet(rs, strSQL)

   Do While Not rs.EOF
      .Row = .MaxRows
      .Col = 2
      .Text = rs!Codigo
       
      .Col = 3
      .Text = rs!Descripcion
               
      .MaxRows = .MaxRows + 1
      rs.MoveNext
   Loop
   rs.Close
   .MaxRows = .MaxRows - 1

End With

Exit Sub
vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub vGridCargos_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
  lblCargos.Caption = Format(fxGridSumaMarcado(vGridCargos, 4), "standard")
  txtMontoPagar.Text = 0
  txtMontoPagar.Text = Format(fxCalculaTotal, "standard")
End Sub

Public Function fxGridSumaMarcado(vGrid As Object, Columna As Long) As Double
'Este procedimiento suma las lineas marcadas
Dim suma As Double, i As Long
On Error GoTo vError

suma = 0
vGrid.Row = 1
vGrid.Col = 1
For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    If vGrid.Value = 1 Then
        vGrid.Col = Columna
        If IsNumeric(vGrid.Value) Then
            suma = suma + vGrid.Value
        End If
        vGrid.Col = 1
    End If
Next i
fxGridSumaMarcado = suma

Exit Function
    
vError:
   MsgBox Err.Description
    
End Function

Private Sub vGridRefundiciones_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
  lblRefundiciones = Format(fxGridSumaMarcado(vGridRefundiciones, 4), "standard")
  txtMontoPagar.Text = 0
  txtMontoPagar.Text = Format(fxCalculaTotal, "standard")
End Sub

Private Sub sbGuardaRefundicion(ByVal Noperacion As Integer)
Dim i As Integer
Dim nOperacionR As Integer
Dim vMontoR As Double

On Error GoTo error

With vGridRefundiciones
    For i = 1 To .MaxRows
       .Row = i
       .Col = 1
       If .Value = 1 Then
          .Col = 2
          nOperacionR = .Text
          
          .Col = 4
          vMontoR = .Text
          
       End If
    Next i
End With
   
strSQL = "Insert AFI_CD_REFUNDICIONES (NOPERACIONR, NOPERACION, MONTO, FECHA)" _
       & "values (" & nOperacionR & "," & Noperacion & "," & vMontoR & ",'" & vFecha & "')"
Call ConectionExecute(strSQL)

Exit Sub
    
error:
   MsgBox Err.Description

End Sub

Private Sub sbGuardaCargos(ByVal Noperacion As Integer)
Dim vCodigoCargo As Integer
Dim vMontoC As Double
On Error GoTo vError
 
With vGridCargos
  For i = 1 To .MaxRows
     .Row = i
     .Col = 1
     If .Value = 1 Then
        .Col = 2
        vCodigoCargo = .Text
        
        .Col = 4
        vMontoC = .Text

     End If
  Next i
End With
   
strSQL = "Insert AFI_CD_CARGOS_CUENTAS(CODIGO, NOPERACION, MONTO, FECHA)" _
       & "values (" & vCodigoCargo & "," & Noperacion & "," & vMontoC & ",'" & vFecha & "')"

Call ConectionExecute(strSQL)

Exit Sub
vError:
   MsgBox Err.Description

End Sub

Private Function fxCalculaTotal() As Double
On Error GoTo vError
   fxCalculaTotal = Format(vMontoActividad, "Standard") - (CCur(lblRefundiciones.Caption) + CCur(lblCargos.Caption))

Exit Function

vError:
  fxCalculaTotal = 0
End Function




