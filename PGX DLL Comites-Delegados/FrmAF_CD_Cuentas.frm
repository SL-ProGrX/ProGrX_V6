VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_CD_Cuentas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuentas"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   8955
   Icon            =   "FrmAF_CD_Cuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabCuentas 
      Height          =   8250
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   14552
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Desembolsos"
      TabPicture(0)   =   "FrmAF_CD_Cuentas.frx":3482
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "SSTab"
      Tab(0).Control(1)=   "txtAjusteAsoc"
      Tab(0).Control(2)=   "Frm1"
      Tab(0).Control(3)=   "txtNotas"
      Tab(0).Control(4)=   "cboBanco"
      Tab(0).Control(5)=   "cboMiembros"
      Tab(0).Control(6)=   "TxtNombreComite"
      Tab(0).Control(7)=   "txtCodComite"
      Tab(0).Control(8)=   "Frame3"
      Tab(0).Control(9)=   "tlbPrincipal"
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(11)=   "Line1(2)"
      Tab(0).Control(12)=   "Line1(1)"
      Tab(0).Control(13)=   "Line1(0)"
      Tab(0).Control(14)=   "lblX(4)"
      Tab(0).Control(15)=   "lblX(3)"
      Tab(0).Control(16)=   "lblX(2)"
      Tab(0).Control(17)=   "lblX(1)"
      Tab(0).Control(18)=   "Label18"
      Tab(0).Control(19)=   "LblScom"
      Tab(0).Control(20)=   "Label13"
      Tab(0).Control(21)=   "lblMontoPagar"
      Tab(0).Control(22)=   "LblTotal"
      Tab(0).Control(23)=   "Label10"
      Tab(0).Control(24)=   "lblLiq"
      Tab(0).Control(25)=   "Label3"
      Tab(0).Control(26)=   "lblRegistro"
      Tab(0).Control(27)=   "LblCuenta"
      Tab(0).Control(28)=   "Label6"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Aprobación"
      TabPicture(1)   =   "FrmAF_CD_Cuentas.frx":9CE4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "tlbAprobacion"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "vGrid"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cbobancApro"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Consulta de Envio de Operaciones a Tesorería"
      TabPicture(2)   =   "FrmAF_CD_Cuentas.frx":10546
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswinfoenvio"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "Label8"
      Tab(2).ControlCount=   3
      Begin TabDlg.SSTab SSTab 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   49
         Top             =   2880
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Actividades"
         TabPicture(0)   =   "FrmAF_CD_Cuentas.frx":16DA8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label17"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblX(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lswMixto"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Frame2"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cboActividades"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Refundiciones"
         TabPicture(1)   =   "FrmAF_CD_Cuentas.frx":16DC4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblX(5)"
         Tab(1).Control(1)=   "lblRefundiciones"
         Tab(1).Control(2)=   "vGridRefundiciones"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Cargos"
         TabPicture(2)   =   "FrmAF_CD_Cuentas.frx":16DE0
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lblX(6)"
         Tab(2).Control(1)=   "lblCargos"
         Tab(2).Control(2)=   "vGridCargos"
         Tab(2).ControlCount=   3
         Begin VB.ComboBox cboActividades 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1665
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   960
            Width           =   6390
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   1665
            TabIndex        =   50
            Top             =   420
            Width           =   6435
            Begin VB.OptionButton OptMixto 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "Actividades Conjuntas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   4440
               TabIndex        =   53
               Top             =   40
               Width           =   2025
            End
            Begin VB.OptionButton OptEsp 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "Actividad Especial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   300
               Left            =   2385
               TabIndex        =   52
               Top             =   40
               Width           =   1845
            End
            Begin VB.OptionButton OptDes 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               Caption         =   "Desembolso Trimestral"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   270
               Left            =   120
               TabIndex        =   51
               Top             =   40
               Width           =   2010
            End
         End
         Begin MSComctlLib.ListView lswMixto 
            Height          =   2280
            Left            =   1665
            TabIndex        =   55
            Top             =   960
            Visible         =   0   'False
            Width           =   6390
            _ExtentX        =   11271
            _ExtentY        =   4022
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Codigo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Descripción"
               Object.Width           =   5027
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Monto"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Actividad"
               Object.Width           =   1676
            EndProperty
         End
         Begin FPSpreadADO.fpSpread vGridCargos 
            Height          =   2055
            Left            =   -74640
            TabIndex        =   62
            Top             =   600
            Width           =   7695
            _Version        =   524288
            _ExtentX        =   13573
            _ExtentY        =   3625
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   4
            ScrollBars      =   2
            SpreadDesigner  =   "FrmAF_CD_Cuentas.frx":16DFC
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridRefundiciones 
            Height          =   2055
            Left            =   -74880
            TabIndex        =   63
            Top             =   480
            Width           =   8055
            _Version        =   524288
            _ExtentX        =   14208
            _ExtentY        =   3625
            _StockProps     =   64
            BackColorStyle  =   1
            BorderStyle     =   0
            EditEnterAction =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   6
            ScrollBars      =   2
            SpreadDesigner  =   "FrmAF_CD_Cuentas.frx":173F2
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin VB.Label lblCargos 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   -68355
            TabIndex        =   61
            Top             =   2880
            Width           =   1470
         End
         Begin VB.Label lblX 
            Caption         =   "Total Cargos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   -69960
            TabIndex        =   60
            Top             =   2880
            Width           =   1440
         End
         Begin VB.Label lblRefundiciones 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   -68355
            TabIndex        =   59
            Top             =   2880
            Width           =   1470
         End
         Begin VB.Label lblX 
            Caption         =   "Total Refundiciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   -69960
            TabIndex        =   58
            Top             =   2880
            Width           =   1440
         End
         Begin VB.Label lblX 
            Caption         =   "Actividad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   57
            Top             =   900
            Width           =   960
         End
         Begin VB.Label Label17 
            Caption         =   "Tipo de Actividad"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   56
            Top             =   450
            Width           =   1320
         End
      End
      Begin VB.TextBox txtAjusteAsoc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69960
         TabIndex        =   48
         Top             =   1110
         Width           =   765
      End
      Begin VB.Frame Frm1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   -72915
         TabIndex        =   35
         Top             =   1530
         Width           =   6405
         Begin VB.ComboBox cboJunta 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3705
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   40
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.OptionButton OptOficina 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Oficina de Comites"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   39
            Top             =   100
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.OptionButton OptJunta 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Junta Directiva"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   2175
            TabIndex        =   38
            Top             =   100
            Width           =   1560
         End
         Begin VB.OptionButton OptDirector 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Director de Zona"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   4080
            TabIndex        =   37
            Top             =   100
            Width           =   1530
         End
      End
      Begin MSComctlLib.ListView lswinfoenvio 
         Height          =   5880
         Left            =   -74790
         TabIndex        =   25
         Top             =   2160
         Width           =   8340
         _ExtentX        =   14711
         _ExtentY        =   10372
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
            Name            =   "Arial"
            Size            =   8.25
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   585
         Width           =   4440
      End
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   -72915
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   6990
         Width           =   6420
      End
      Begin VB.ComboBox cboBanco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72915
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2340
         Width           =   3765
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
         TabIndex        =   9
         Top             =   420
         Width           =   8340
         Begin MSComCtl2.DTPicker dtpFinalEnvio 
            Height          =   315
            Left            =   3120
            TabIndex        =   26
            Tag             =   "2"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   70909955
            CurrentDate     =   39493
         End
         Begin VB.TextBox TxtSolicitud 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6450
            TabIndex        =   13
            Tag             =   "4"
            Top             =   360
            Width           =   1665
         End
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1935
            TabIndex        =   11
            Tag             =   "3"
            Top             =   780
            Width           =   2385
         End
         Begin VB.TextBox txtCod_Comite 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   6450
            TabIndex        =   10
            Tag             =   "4"
            Top             =   795
            Width           =   1665
         End
         Begin MSComCtl2.DTPicker dtpEnvio 
            Height          =   315
            Left            =   1935
            TabIndex        =   12
            Tag             =   "1"
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   165806083
            CurrentDate     =   39189
         End
         Begin VB.Label Label20 
            Caption         =   "No. Solicitud a Tesorería"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4530
            TabIndex        =   17
            Top             =   420
            Width           =   1875
         End
         Begin VB.Label Label22 
            Caption         =   "Fecha de Envio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   390
            Width           =   1290
         End
         Begin VB.Label Label24 
            Caption         =   "Usuario responsable"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   15
            Top             =   840
            Width           =   1530
         End
         Begin VB.Label Label7 
            Caption         =   "Comité"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4530
            TabIndex        =   14
            Top             =   870
            Width           =   1140
         End
      End
      Begin VB.ComboBox cboMiembros 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -72915
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2000
         Width           =   3765
      End
      Begin VB.TextBox TxtNombreComite 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72250
         Locked          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Tecla F4 para seleccionar Comité Activos"
         Top             =   765
         Width           =   5700
      End
      Begin VB.TextBox txtCodComite 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72915
         MaxLength       =   4
         TabIndex        =   6
         ToolTipText     =   "Digite el código del Comité"
         Top             =   765
         Width           =   675
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -69120
         TabIndex        =   1
         Top             =   1965
         Width           =   2520
         Begin VB.OptionButton OptTrans 
            Caption         =   "Tranferencia"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   30
            TabIndex        =   3
            Top             =   45
            Value           =   -1  'True
            Width           =   1290
         End
         Begin VB.OptionButton OptChk 
            Caption         =   "Cheque "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1485
            TabIndex        =   2
            Top             =   45
            Width           =   945
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6975
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   8415
         _Version        =   524288
         _ExtentX        =   14843
         _ExtentY        =   12303
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         SpreadDesigner  =   "FrmAF_CD_Cuentas.frx":17A43
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   360
         Left            =   -67800
         TabIndex        =   45
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   635
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
         BorderStyle     =   1
      End
      Begin MSComctlLib.Toolbar tlbAprobacion 
         Height          =   360
         Left            =   7560
         TabIndex        =   46
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   635
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
         BorderStyle     =   1
      End
      Begin VB.Label Label2 
         Caption         =   "Asociados de Ajuste"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -71640
         TabIndex        =   47
         Top             =   1170
         Width           =   1665
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -74760
         X2              =   -66480
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74760
         X2              =   -66480
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -74760
         X2              =   -66480
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblX 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   44
         Top             =   6870
         Width           =   1200
      End
      Begin VB.Label lblX 
         Caption         =   "Fecha Registro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -69000
         TabIndex        =   43
         Top             =   6480
         Width           =   1200
      End
      Begin VB.Label lblX 
         Caption         =   "Fecha Liq."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -71160
         TabIndex        =   42
         Top             =   6480
         Width           =   825
      End
      Begin VB.Label lblX 
         Caption         =   "Monto a Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   41
         Top             =   6480
         Width           =   1200
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
         Left            =   -74655
         TabIndex        =   34
         Top             =   1485
         Width           =   1065
      End
      Begin VB.Label LblScom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72915
         TabIndex        =   33
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label13 
         Caption         =   "Asociados por Comité"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74685
         TabIndex        =   32
         Top             =   1170
         Width           =   1665
      End
      Begin VB.Label lblMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -72915
         TabIndex        =   31
         Top             =   6480
         Width           =   1470
      End
      Begin VB.Label LblTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -67320
         TabIndex        =   30
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Total de Asociados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -68790
         TabIndex        =   29
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label lblLiq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -70290
         TabIndex        =   28
         Top             =   6465
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Comite Principal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74685
         TabIndex        =   27
         Top             =   810
         Width           =   1200
      End
      Begin VB.Label Label9 
         Caption         =   "Bancos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   585
         TabIndex        =   24
         Top             =   615
         Width           =   630
      End
      Begin VB.Label lblRegistro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -67680
         TabIndex        =   22
         Top             =   6480
         Width           =   1125
      End
      Begin VB.Label LblCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -69075
         TabIndex        =   19
         Top             =   2325
         Width           =   2445
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
         Height          =   285
         Left            =   -74775
         TabIndex        =   18
         Top             =   1890
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
         Left            =   -74640
         TabIndex        =   4
         Top             =   1920
         Width           =   1515
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8760
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Caption         =   "Registro de Cuentas a Comités"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3795
   End
End
Attribute VB_Name = "frmAF_CD_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Miembro As String, vCodigo As String
Dim i As Integer, x As Integer
Dim rs As New ADODB.Recordset
Dim strSQL As String
Dim vCuentaGasto As String
Dim vMontoActividad As Double
Dim vBacCargado As Boolean, vPaso As Boolean
Dim vFecha As String

Function fxConseConjunto()

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(noperacion),0) as Consecutivo from afi_cd_acticonjunta "
rs.Open strSQL, glogon.Conection, adOpenStatic
 fxConseConjunto = rs!consecutivo + 1
rs.Close

End Function

Function FxNomComite(vUnidad As String)
   
   Dim rs As New ADODB.Recordset
   Dim strSQL As String
  
   strSQL = "select U.descripcion from uprogramatica U right join afi_cd_comites_unidades A " _
            & "on U.codigo = A.cod_comite" _
            & " where A.cod_comite = '" & vUnidad & "'"
            rs.Open strSQL, glogon.Conection, adOpenStatic
   If rs.EOF Then
      FxNomComite = "No existe unidad definida en Comites y Delegados"
   Else
      FxNomComite = rs!Descripcion
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

Private Sub sbAprobacion()

If vGrid.MaxRows = 0 Then
  MsgBox "No hay informacion para procesar", vbInformation, "Información"
  Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

For i = 1 To vGrid.MaxRows
   vGrid.Row = i
       vGrid.Col = 1
       If vGrid.Value = vbChecked Then
           vGrid.Col = 2
           
           'Activa y Registra Asiento
           strSQL = "exec spAFI_CD_AsientoCuentas '" & vGrid.Text & "', '" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular & "'"
           glogon.Conection.Execute strSQL
                      
       End If
Next i

Me.MousePointer = vbDefault
MsgBox "Aprobación Realizada", vbInformation, "Información"
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
 
End Sub

Private Sub sbRechazar()

If vGrid.MaxRows = 0 Then
  MsgBox "No hay informacion para procesar", vbInformation, "Información"
  Exit Sub
End If

On Error GoTo vError
Me.MousePointer = vbHourglass


For i = 1 To vGrid.MaxRows
   vGrid.Row = i
       vGrid.Col = 1
       If vGrid.Value = vbChecked Then
           vGrid.Col = 2
           strSQL = "update afi_cd_cuentas set estado = 'R' " _
                  & "where noperacion = '" & vGrid.Text & "'"
           glogon.Conection.Execute strSQL
       End If
Next i

Me.MousePointer = vbDefault
MsgBox "Se rechaza la Operación", vbInformation, "Información"



Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical


End Sub

Private Sub sbCargaBan()

cbobancApro.Clear

strSQL = "select distinct C.id_banco,B.descripcion" _
       & " from bancos B inner join afi_cd_cuentas C on B.id_banco = C.id_banco"
rs.Open strSQL, glogon.Conection, adOpenStatic
  
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

Sub sbActividades()
 Dim itmX As ListItem
 Dim vFiltro As String
 Dim i As Integer
 Dim vSubMonto1 As Currency, vSubMonto2 As Currency, Total As Currency
'ReDim ActMonto(LswMixto.ListItems.Count, 1) As String
 cboActividades.Clear
 
 If rs.State = 1 Then rs.Close
  
 Select Case True
   
   Case OptDes.Value = True
      vFiltro = " and tipo = 'T'"
   Case OptEsp.Value = True
      vFiltro = " and tipo = 'E'"
   Case OptMixto = True
      vFiltro = " and tipo in ('E','T')"
 End Select
   
  strSQL = "select P.fechaliq,P.cod_actividad,P.descripcion,P.tipo from afi_cd_actividades P inner join afi_cd_comites_actividades C " _
                   & "on P.cod_actividad = C.cod_actividad inner join afi_cd_comites_unidades A on A.cod_comite = C.cod_comite " _
                   & "where A.cod_comite = '" & txtCodComite.Text & "' " & vFiltro & " " _
                   & "group by P.fechaliq,P.cod_actividad,P.descripcion,P.tipo"
                   rs.Open strSQL, glogon.Conection, adOpenStatic
 
 
If OptMixto = False Then
     While Not rs.EOF
       cboActividades.AddItem rs!Descripcion
       cboActividades.ItemData(cboActividades.NewIndex) = rs!Cod_actividad
       rs.MoveNext
     Wend
Else
   
  lswMixto.ListItems.Clear
   While Not rs.EOF
      Set itmX = lswMixto.ListItems.Add(, , rs!Cod_actividad)
      itmX.SubItems(1) = Trim(rs!Descripcion)
      Select Case True
       Case rs!Tipo = "T"
        itmX.SubItems(3) = "Trimestral"
       Case rs!Tipo = "E"
        itmX.SubItems(3) = "Especial"
      End Select
     rs.MoveNext
Wend
rs.Close

For i = 1 To lswMixto.ListItems.Count
       
       strSQL = "select M.cod_actividad,A.cod_comite,M.monto,M.minimo,M.maximo " _
                & "from Afi_cd_actividades_rangos M inner join afi_cd_comites_actividades A " _
                & "on A.cod_actividad = M.cod_actividad where A.cod_comite = '" & txtCodComite.Text & "' " _
                & "and M.cod_actividad = " & lswMixto.ListItems.Item(i) & ""
                rs.Open strSQL, glogon.Conection, adOpenStatic
           
           While Not rs.EOF
            Select Case True
             Case lswMixto.ListItems.Item(i).SubItems(3) = Trim("Trimestral")
                If rs!maximo >= CInt(LblTotal.Caption) And rs!minimo <= CInt(LblTotal.Caption) Then
                    lswMixto.ListItems.Item(i).SubItems(2) = Format(rs!Monto, "Standard")
                End If
             Case lswMixto.ListItems.Item(i).SubItems(3) = Trim("Especial")
                If rs!maximo >= CInt(LblTotal.Caption) And rs!minimo <= CInt(LblTotal.Caption) Then
                  lswMixto.ListItems.Item(i).SubItems(2) = Format(rs!Monto * CInt(LblTotal.Caption), "Standard")
                End If
             End Select
           rs.MoveNext
          Wend
     rs.Close
 Next i
   


End If
End Sub

Private Function fxConsecutivo() As Long

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(noperacion),0) as Consecutivo from afi_cd_cuentas "
rs.Open strSQL, glogon.Conection, adOpenStatic
fxConsecutivo = rs!consecutivo + 1
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
            
If Len(Trim(txtUsuario.Text)) > 0 Then
   strSQL = strSQL & " and C.tesoreria_usuario like '%" & txtUsuario.Text & "%'"
End If

If Len(Trim(txtCod_Comite.Text)) > 0 Then
   strSQL = strSQL & " and C.cod_comite like '%" & txtCod_Comite.Text & "%'"
End If


lswinfoenvio.ListItems.Clear
 
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 
 While Not rs.EOF
      Set itmX = lswinfoenvio.ListItems.Add(, , rs!Noperacion)
      itmX.SubItems(1) = rs!comite
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
   rs.Open strSQL, glogon.Conection, adOpenStatic
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
     cboActividades.Clear
                
     strSQL = "select N.cedula,S.nombre,N.cod_comite,C.descripcion from afi_cd_comites C left join afi_cd_nombramientos N " _
            & "on C.cod_comite = N.cod_comite inner join socios S on S.cedula = N.cedula " _
            & "where N.cod_comite = '" & vComite & "' and N.APL_DESEMBOLSOS=1 "
     rs.Open strSQL, glogon.Conection, adOpenStatic
     
     If Not rs.EOF Then
        TxtNombreComite.Text = Trim(rs!Descripcion)
     Else
        MsgBox "No se cuenta con miembro asigando el desembolso!!"
     End If

     Do While Not rs.EOF
       cboMiembros.AddItem rs!Nombre
       cboMiembros.ItemData(cboMiembros.NewIndex) = rs!Cedula
       rs.MoveNext
     Loop
     rs.Close
     
End Sub

Sub sbMjunta()
Dim i As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim Mjunta As Integer

cboJunta.Clear

strSQL = "select * from afi_cd_directores"
         rs.Open strSQL, glogon.Conection, adOpenStatic
                   
          While Not rs.EOF
            cboJunta.AddItem rs!Nombre
            cboJunta.ItemData(cboJunta.NewIndex) = rs!cod_director
            rs.MoveNext
          Wend
          
rs.Close

End Sub
Private Sub sbLimpiar()
     txtCodComite.Text = ""
     cboMiembros.Clear
     cboActividades.Clear
     cboActividades.Visible = True
     cboBanco.Clear
     LblCuenta.Caption = ""
     txtNotas.Text = ""
     TxtNombreComite.Text = ""
     LblScom.Caption = ""
     LblTotal.Caption = ""
     lblMontoPagar.Caption = 0
     lblRefundiciones.Caption = 0
     lblCargos.Caption = 0
     lswMixto.ListItems.Clear
     lswMixto.Visible = False
     lblLiq.Caption = ""
     Frm1.Width = 6405
     cboJunta.Visible = False
     OptChk.Value = False
     OptDes.Value = False
     OptDirector.Value = False
     OptEsp.Value = False
     OptJunta.Value = False
     OptMixto.Value = False
     txtCodComite.SetFocus
End Sub


Private Sub sbTiposActividadActiva()

lblLiq.Caption = ""
lblMontoPagar.Caption = 0

cboActividades.Visible = False
lswMixto.Visible = False

Select Case True
  
  Case OptDes.Value = True
    
    cboActividades.Visible = True
  
  Case OptEsp.Value = True
    cboActividades.Visible = True
  
  Case OptMixto.Value = True
    lswMixto.Visible = True
End Select

End Sub

Private Sub cboActividades_Click()

Dim Id As Integer, i As Integer

If rs.State = 1 Then rs.Close

For i = 0 To cboActividades.ListCount
  If cboActividades.ListIndex = i Then
     Id = cboActividades.ItemData(cboActividades.ListIndex)
  End If
Next i
 
 strSQL = "select M.*,P.COD_CUENTA,P.fechaliq from afi_cd_actividades_rangos M inner join afi_cd_actividades P " _
            & "on M.cod_actividad = P.cod_actividad " _
            & "where M.cod_actividad = " & Id & ""
            rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
  lblLiq.Caption = IIf(IsNull(Format(rs!fechaliq, "mm/dd/yyyy")), fxFechaServidor, Format(rs!fechaliq, "mm/dd/yyyy"))
End If
                     
Do While Not rs.EOF
  Select Case True
      Case OptDes = True
       If rs!maximo >= CInt(LblTotal.Caption) And rs!minimo <= CInt(LblTotal.Caption) Then
        vMontoActividad = Format(rs!Monto, "standard")
        vCuentaGasto = rs!cod_cuenta
       End If
      Case OptEsp = True
       If rs!maximo >= CInt(LblTotal.Caption) And rs!minimo <= CInt(LblTotal.Caption) Then
        vMontoActividad = Format(rs!Monto * CInt(LblTotal.Caption), "Standard")
        vCuentaGasto = rs!cod_cuenta
       End If
   End Select
   lblMontoPagar = 0
   lblMontoPagar = Format(vMontoActividad, "standard")
 rs.MoveNext
Loop
  rs.Close

End Sub

Private Sub cbobancApro_Click()

If vPaso Or cbobancApro.ListCount <= 0 Then
  vGrid.MaxRows = 0
  Exit Sub
End If

  Call sbCargPago
End Sub

Private Sub cboBanco_Click()

If cboMiembros.Text <> Empty Then
    LblCuenta = fxCuentaAhorros(cboMiembros.ItemData(cboMiembros.ListIndex), cboBanco.ItemData(cboBanco.ListIndex))
End If
'Dim strSQL As String, rs As New ADODB.Recordset
'Dim Id As Integer
'
'
'For I = 0 To cboBanco.ListCount
'    If cboBanco.ListIndex = I Then
'       Id = cboBanco.ItemData(cboBanco.ListIndex)
'    End If
'Next I
'
'strSQL = "select C.cedula,C.id_banco,C.cuenta,B.descripcion " _
'         & "from cuentas_ahorros C left join bancos B " _
'         & "on C.id_banco = B.id_banco " _
'         & "where C.cedula ='" & Miembro & "' and B.id_banco = '" & Id & "'"
'         rs.Open strSQL, glogon.Conection, adOpenStatic
'          If Not rs.EOF Then
'             LblCuenta.Caption = rs!Cuenta
'            If cboBanco.Text <> "" Then
'              OptTrans.Value = True
'            End If
'          End If
'rs.Close
End Sub


Private Sub cboMiembros_Click()
 
Dim strSQL As String, rs As New ADODB.Recordset

cboBanco.Clear
For i = 0 To cboMiembros.ListCount
    If cboMiembros.ListIndex = i Then
       Miembro = cboMiembros.ItemData(cboMiembros.ListIndex)
    End If
Next i

strSQL = "select C.cedula,C.id_banco,C.cuenta,B.descripcion " _
         & "from cuentas_ahorros C left join bancos B " _
         & "on C.id_banco = B.id_banco " _
         & "where C.cedula ='" & Miembro & "'"
         rs.Open strSQL, glogon.Conection, adOpenStatic
          While Not rs.EOF
            cboBanco.AddItem rs!Descripcion
            cboBanco.ItemData(cboBanco.NewIndex) = rs!ID_BANCO
            rs.MoveNext
          Wend
rs.Close
cboBanco.SetFocus
LblCuenta.Caption = ""
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
 .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_cd_ConsulaOperacion.rpt")
     
  Select Case True
   Case TxtSolicitud.Text <> ""
    strSQL = "{afi_cd_cuentas.tesoreria_nsolicitud} = " & TxtSolicitud.Text & " and "
    txtCod_Comite.Text = ""
    txtUsuario.Text = ""
    vSubTitulo = "No. Solicitud " & TxtSolicitud.Text & ""
   Case txtCod_Comite.Text <> ""
    strSQL = "{afi_cd_cuentas.id_pricomite} = '" & txtCod_Comite.Text & "' and "
    TxtSolicitud.Text = ""
    txtUsuario.Text = ""
    vSubTitulo = "Comité: " & txtCod_Comite.Text & ""
   Case txtUsuario.Text <> ""
    strSQL = "{afi_cd_cuentas.tesoreria_usaurio} = " & txtUsuario.Text & " and "
    TxtSolicitud.Text = ""
    txtCod_Comite.Text = ""
    vSubTitulo = "Usuario: " & txtUsuario.Text & ""
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
 txtCodComite.Text = ""
 TxtSolicitud.Text = ""
 txtUsuario.Text = ""
 dtpEnvio.Value = Format(vFecha, "dd/mm/yyyy")
 dtpFinalEnvio.Value = Format(vFecha, "dd/mm/yyyy")
 lswinfoenvio.ListItems.Clear
 txtUsuario.SetFocus
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
  Case 48 To 57, 8
  Case 13
   
  Case Else
   KeyAscii = 0
End Select


End Sub

Private Sub cmdLiquidaciones_Click()
 Call sbSIFForms("frmAF_CD_Liquidaciones")
End Sub

Private Sub cmdPlan_Click()
 Call MuestraForms(frmAF_CD_Plan)
 frmAF_CD_Plan.txtComite.Caption = frmAF_CD_Cuentas.txtCodComite.Text
End Sub

Private Sub DtpEnvio_Change()
 Call sbConsulta(1)
End Sub

Private Sub dtpFinalEnvio_Change()
 Call sbConsulta(1)
End Sub

Private Sub Form_Activate()
 vModulo = 23
End Sub

Private Sub Form_Load()
   
 vModulo = 23
   
 vFecha = Format(fxFechaServidor, "yyyymmdd")
 
 vBacCargado = False
 SSTabCuentas.Tab = 0
 Call sbCargaBancos
 
 dtpEnvio.Value = fxFechaServidor
 dtpFinalEnvio.Value = dtpEnvio.Value
 lblRegistro.Caption = Format(dtpEnvio.Value, "dd/mm/yyyy")
 
 tlbPrincipal.ImageList = frmContenedor.imgToolbarIcons
 tlbPrincipal.Buttons.Item(1).Image = 2
 tlbPrincipal.Buttons.Item(3).Image = 9
 tlbPrincipal.Buttons.Item(4).Image = 4
  
 tlbAprobacion.ImageList = frmContenedor.imgToolbarIcons03
 tlbAprobacion.Buttons.Item(1).Image = 3
 tlbAprobacion.Buttons.Item(2).Image = 4
 
 
 txtAjusteAsoc.Text = 0
 lblRefundiciones.Caption = 0
 lblCargos.Caption = 0
 lblMontoPagar.Caption = 0
  
 'OptDes.Value = True
 Call sbTiposActividadActiva
 Call sbCargaCargos

 If GLOBALES.gTag <> Empty Then
   txtCodComite.Text = GLOBALES.gTag
   Call cboMiembros.Clear
   Call cboActividades.Clear
   Call sbCalMiembros
   Call sbLlamaComite(txtCodComite.Text)
   Call sbCargaLiquidaciones
   GLOBALES.gTag = Empty
 End If

End Sub

Private Sub sbCalMiembros()

On Error GoTo vError

If rs.State = 1 Then rs.Close
   
strSQL = "select count(*) as 'Cantidad' from socios" _
       & " where EstadoActual = 'S' and UP in(select Codigo_UP from Afi_CD_Comites_Unidades where cod_comite = " & txtCodComite.Text & ")"
rs.Open strSQL, glogon.Conection, adOpenStatic

LblScom.Caption = rs!cantidad
LblTotal.Caption = rs!cantidad - CCur(txtAjusteAsoc.Text)

rs.Close


Exit Sub

vError:
  MsgBox Err.Description, vbCritical
  
End Sub

Private Sub LswMixto_Click()
Dim i As Integer

vMontoActividad = 0
For i = 1 To lswMixto.ListItems.Count
    If lswMixto.ListItems.Item(i).Checked = True Then
       If lswMixto.ListItems.Item(i).SubItems(2) <> "" Then
          vMontoActividad = vMontoActividad + CCur(lswMixto.ListItems.Item(i).SubItems(2))
       Else
          MsgBox "No puede seleccionar una actividad que no se encuentra autorizada o activa para este comite", vbCritical, "Información"
          lswMixto.ListItems.Item(i).Checked = False
       End If
    End If
Next i
 
lblMontoPagar.Caption = Format(vMontoActividad, "Standard")

End Sub

Private Sub OptChk_Click()
  
  If OptChk.Value = True Then
   OptChk.Tag = 1
      
   strSQL = "select id_banco,descripcion from bancos "
   rs.Open strSQL, glogon.Conection, adOpenStatic
   
   While Not rs.EOF
     cboBanco.AddItem rs!Descripcion
     cboBanco.ItemData(cboBanco.NewIndex) = rs!ID_BANCO
    rs.MoveNext
   Wend
  
  Else
   OptChk.Tag = 0
 End If
 rs.Close
End Sub

Private Sub OptDes_Click()
 Call sbTiposActividadActiva
 Call sbActividades
End Sub

Private Sub OptDirector_Click()

If OptDirector.Value = True Then
'  Frm1.Width = 8730
  cboJunta.Visible = True
  Call sbMjunta
End If

End Sub

Private Sub OptEsp_Click()
 Call sbTiposActividadActiva
 Call sbActividades
End Sub


Private Sub OptJunta_Click()

If OptDirector.Value = False Then
  cboJunta.Visible = False
End If

End Sub

Private Sub OptMixto_Click()
 Call sbTiposActividadActiva
 Call sbActividades
End Sub


Private Sub OptOficina_Click()

If OptDirector.Value = False Then
  cboJunta.Visible = False
End If

End Sub

Private Sub OptTrans_Click()
 If OptTrans.Value = True Then
   OptTrans.Tag = 1
  Else
   OptTrans.Tag = 0
 End If
End Sub

Private Sub SSTabCuentas_Click(PreviousTab As Integer)

  If SSTabCuentas.Tab = 1 Then
   If vBacCargado = False Or OptChk.Value = True Then
    Call sbCargaBan
   End If
  End If

End Sub

'Private Sub tlbPricipal_ButtonClick(ByVal Button As ComctlLib.Button)
'Select Case UCase(Button.Key)
'  Case "LIQUIDACION"
'     Call sbSIFForms("frmAF_CD_Liquidaciones")
'  Case "NUEVO"
'     Call sbLimpiar
' Case "APLICAR"
'     Call sbAplicar
'     Call sbLimpiar
'End Select
'
'
'End Sub


Private Sub tlbAprobacion_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
  Case "APROBAR"
    Call sbAprobacion
    Call sbCargPago
  Case "Rechazado"
    Call sbRechazar
    Call sbCargPago
End Select

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
'             .ReportFileName = SIFGlobal.fxSIFPathReportes("afi_cd_OperacionesEnviadasTesoreria.rpt")
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
''             .ReportFileName = SIFGlobal.fxSIFPathReportes("Afi_cd_ConsulaOperacion.rpt")
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

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
  Case "LIQUIDACION"
    sbSIFForms ("frmAF_CD_Liquidaciones")
  Case "NUEVO"
    Call sbLimpiar
  Case "APLICAR"
    Call sbAplicar
End Select
End Sub

Private Sub txtAjusteAsoc_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Or KeyAscii = 9 Then
     LblTotal.Caption = CCur(LblScom.Caption) - CCur(txtAjusteAsoc.Text)
  End If
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

Private Sub txtCodComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select distinct cod_comite,U.descripcion from afi_cd_comites_unidades A " _
                             & "left join uprogramatica U on A.cod_comite = U.codigo "
       frmBusquedas.Show vbModal
       vCodigo = gBusquedas.Resultado
       txtCodComite.Text = gBusquedas.Resultado
       Call TxtCodComite_KeyPress(vbKeyReturn)
 
End If
End Sub

Private Sub TxtCodComite_KeyPress(KeyAscii As Integer)
  
 ' cargar actividades y sus caracteristICAS
  
  Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        Call cboMiembros.Clear
        Call cboActividades.Clear
        Call sbCalMiembros
        Call sbLlamaComite(txtCodComite.Text)
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

If Trim(txtCodComite) = "" Then strSQL = "Comite" & vbCrLf
If Trim(cboMiembros.Text) = "" Then strSQL = strSQL & "Necesita Agregar Miembro" & vbCrLf
If Trim(cboBanco.Text) = "" Then strSQL = strSQL & "Debe seleccionar un banco" & vbCrLf
If OptDirector.Value = True And cboJunta.Text = "" Then strSQL = strSQL & "Seleccionar un miembro de Junta" & vbCrLf
If OptMixto.Value = False Then
 If Trim(cboActividades.Text) = "" Then strSQL = strSQL & "Actividad" & vbCrLf
End If
If Trim(lblMontoPagar) = 0 Then strSQL = strSQL & "Monto" & vbCrLf
If Trim(txtNotas.Text) = "" Then strSQL = strSQL & "Necesita agregar una descripción en Observaciones" & vbCrLf
If OptTrans.Value = True And LblCuenta.Caption = "" Then strSQL = strSQL & "La cuenta del Banco vacia" & vbCrLf

If strSQL <> "" Then
   MsgBox strSQL, vbInformation, "Faltan Los Siguientes Datos:"
   Exit Sub
End If

On Error GoTo vError:

Select Case True
  Case OptTrans.Value = True  'Pago con Transferencia
    vTipo = "T"
    vCuentaBanco = cboBanco.ItemData(cboBanco.ListIndex)
    vCtaDelegado = LblCuenta.Caption
  Case OptChk.Value = True    'Pago con Cheque
    vTipo = "C"
    vCuentaBanco = cboBanco.ItemData(cboBanco.ListIndex)
    LblCuenta.Caption = "0"
End Select
 
Select Case True

  Case OptOficina.Value = True
      vAprueba = "O" 'Oficinas de Comites
      strLinea(4) = "Apr: Oficinas de Comites"
  Case OptJunta.Value = True
      vAprueba = "J" 'Junta Directiva
      strLinea(4) = "Apr: Junta Directiva"
      
  Case OptDirector.Value = True
      vAprueba = "D" 'Director de Zona --- si en la consulta del director que aprobo es APRUEBA = D el director es verdadero sino lo toma por alto NO ES VALIDO.
      vCodDir = cboJunta.ItemData(cboJunta.ListIndex)
      strLinea(4) = "Apr: Director de Zona"
      
End Select
 
vNumOperacion = fxConsecutivo

strSQL = "insert afi_cd_cuentas(noperacion,cod_comite,cedula,registro_fecha " _
       & ", registro_usuario,estado,tipo,cuenta,id_banco,notas,aprueba,cod_director,PROCESO,AJUSTE_ASOC,MONTO,MONTO_REFUNDE, MONTO_CARGOS, SALDO)" _
       & " values(" & vNumOperacion & ",'" & txtCodComite.Text & "', '" & cboMiembros.ItemData(cboMiembros.ListIndex) & "'" _
       & ",getdate()" _
       & ", '" & glogon.Usuario & "', 'S','" & vTipo & "','" & vCtaDelegado & "'," & vCuentaBanco & ",'" & txtNotas.Text & "' " _
       & ", '" & vAprueba & "'," & IIf(vCodDir = 0, 1, vCodDir) & ",'T'," & txtAjusteAsoc & "," & CCur(lblMontoPagar.Caption) & "," & CCur(lblRefundiciones.Caption) & "," & CCur(lblCargos.Caption) & "," & CCur(lblMontoPagar.Caption) & ")"
glogon.Conection.Execute strSQL

Select Case True
    Case OptMixto.Value = True
       For i = 1 To lswMixto.ListItems.Count
         If lswMixto.ListItems.Item(i).Checked = True Then
              strSQL = "insert into afi_cd_cuentas_actividades (COD_ACTIVIDAD, NOPERACION, MONTO) " _
                     & "values (" & lswMixto.ListItems.Item(i) & "," & vNumOperacion & "," & CCur(lswMixto.ListItems(i).SubItems(2)) & ")"
              glogon.Conection.Execute strSQL
         End If
       Next i
    
    Case OptDes.Value = True Or OptEsp.Value = True
        
        strSQL = "insert into afi_cd_cuentas_actividades(noperacion,cod_actividad,monto) " _
               & "values (" & vNumOperacion & "," & cboActividades.ItemData(cboActividades.ListIndex) & "," _
               & "" & CCur(lblMontoPagar.Caption) & ")"
        glogon.Conection.Execute strSQL
        
End Select

Call sbGuardaRefundicion(vNumOperacion)
Call sbGuardaCargos(vNumOperacion)

MsgBox "Solicitud Registrada: Proceda a la Aprobación!", vbInformation, "Información"
Call sbLimpiar
Exit Sub

vError:
Resume
  MsgBox Err.Description

End Sub


Private Sub TxtNombreComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select distinct cod_comite,U.descripcion from afi_cd_comites_unidades A " _
                             & "left join uprogramatica U on A.cod_comite = U.codigo "
       frmBusquedas.Show vbModal
       vCodigo = gBusquedas.Resultado
       txtCodComite.Text = gBusquedas.Resultado
       Call TxtCodComite_KeyPress(vbKeyReturn)
 
End If
End Sub

Private Sub TxtNotas_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub sbCargPago()

If cbobancApro.ListCount <= 0 Or cbobancApro.Text = "" Then
  vGrid.MaxRows = 0
  Exit Sub
End If

strSQL = "select distinct 0 as ValorX,C.noperacion,C.cod_comite,U.descripcion as Comite,C.cedula" _
       & ",S.nombre,C.Cuenta, coalesce(sum(M.monto),0) as Total" _
       & " from afi_cd_cuentas C left join Uprogramatica U on C.cod_comite = U.codigo" _
       & " left join Socios S on C.cedula = S.cedula " _
       & " left join afi_cd_cuentas_actividades M on C.nOperacion = M.nOperacion" _
       & " where C.id_banco = " & cbobancApro.ItemData(cbobancApro.ListIndex) & " and C.estado = 'S'" _
       & " group by C.noperacion,C.cod_comite,U.descripcion,C.cedula" _
       & ",S.nombre,C.Cuenta"
Call sbCargaGrid(vGrid, 8, strSQL)

'Elimina Linea en Blanco
vGrid.MaxRows = vGrid.MaxRows - 1

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
             & "where C.cod_comite = '" & txtCodComite.Text & "' and PROCESO='T' " _
             & "group by C.notas,A.noperacion,C.estado,C.tesoreria_nsolicitud,C.liquida_fecha"
   rs.Open strSQL, glogon.Conection, adOpenStatic
    
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
   rs.Open strSQL, glogon.Conection, adOpenStatic

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
  lblMontoPagar.Caption = 0
  lblMontoPagar.Caption = Format(fxCalculaTotal, "standard")
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
  lblMontoPagar.Caption = 0
  lblMontoPagar.Caption = Format(fxCalculaTotal, "standard")
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
glogon.Conection.Execute strSQL

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

glogon.Conection.Execute strSQL

Exit Sub
vError:
   MsgBox Err.Description

End Sub

Private Function fxCalculaTotal() As Double
   fxCalculaTotal = Format(vMontoActividad, "Standard") - (CCur(lblRefundiciones.Caption) + CCur(lblCargos.Caption))
End Function




