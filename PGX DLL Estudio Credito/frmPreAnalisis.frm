VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmPreAnalisis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PreAnalisis"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8940
   Icon            =   "frmPreAnalisis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSalida 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Salida a Impresora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   3480
      TabIndex        =   47
      Top             =   120
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   0
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtCedula 
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   480
      Width           =   5895
   End
   Begin VB.TextBox txtNombre 
      Height          =   315
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   840
      Width           =   5895
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   7853
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Deudor"
      TabPicture(0)   =   "frmPreAnalisis.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblX"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "imgObligaciones"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSalarioLiquido"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtObligacionExSaldos"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtObligacionExCuotas"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCuotaSolicitada"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtSalarioNeto"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPorcSalarioDevengado"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtLiquidez"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lsw"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdGuardarDeudor"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "ImageList1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "tlb"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtObligacionesInSaldos"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Fiadores"
      TabPicture(1)   =   "frmPreAnalisis.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdActualizaFiadores"
      Tab(1).Control(1)=   "vGrid"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Observaciones"
      TabPicture(2)   =   "frmPreAnalisis.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtObservaciones"
      Tab(2).Control(1)=   "cmdGuardaObservaciones"
      Tab(2).Control(2)=   "optObservacion(2)"
      Tab(2).Control(3)=   "optObservacion(1)"
      Tab(2).Control(4)=   "optObservacion(0)"
      Tab(2).Control(5)=   "Line3"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Reportes"
      TabPicture(3)   =   "frmPreAnalisis.frx":091E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdReporte"
      Tab(3).Control(1)=   "fraComite"
      Tab(3).Control(2)=   "fraSolicitud"
      Tab(3).Control(3)=   "cboPreAnalisis"
      Tab(3).Control(4)=   "chkPreAnalisisEspecial"
      Tab(3).Control(5)=   "chkPreAnalisis"
      Tab(3).Control(6)=   "Label3"
      Tab(3).ControlCount=   7
      Begin VB.TextBox txtObservaciones 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   1020
         Width           =   8295
      End
      Begin VB.TextBox txtObligacionesInSaldos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   3300
         Width           =   2295
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   516
         Left            =   5520
         TabIndex        =   48
         Top             =   3300
         Width           =   852
         _ExtentX        =   1503
         _ExtentY        =   910
         ButtonWidth     =   1524
         ButtonHeight    =   910
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Calcular"
               Key             =   "calcular"
               Object.ToolTipText     =   "Calcular Datos"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   240
         Top             =   2940
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreAnalisis.frx":093A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmPreAnalisis.frx":0C5E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Height          =   375
         Left            =   -70080
         TabIndex        =   46
         Top             =   3900
         Width           =   1455
      End
      Begin VB.Frame fraComite 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   -72600
         TabIndex        =   41
         Top             =   1620
         Width           =   3975
         Begin VB.TextBox txtActa 
            Height          =   315
            Left            =   2280
            Locked          =   -1  'True
            MaxLength       =   5
            TabIndex        =   43
            Top             =   720
            Width           =   1575
         End
         Begin VB.ComboBox cboComite 
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   42
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Comité"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Número de Acta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   720
            TabIndex        =   44
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame fraSolicitud 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Left            =   -72600
         TabIndex        =   36
         Top             =   2820
         Width           =   3975
         Begin VB.TextBox txtDe 
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   720
            MaxLength       =   15
            TabIndex        =   38
            ToolTipText     =   "Solicitude de Inicio"
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtA 
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   37
            ToolTipText     =   "Solicitud Final"
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Desde"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   40
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   2040
            TabIndex        =   39
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.ComboBox cboPreAnalisis 
         Height          =   288
         ItemData        =   "frmPreAnalisis.frx":1538
         Left            =   -71400
         List            =   "frmPreAnalisis.frx":1542
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1260
         Width           =   2775
      End
      Begin VB.CheckBox chkPreAnalisisEspecial 
         Appearance      =   0  'Flat
         Caption         =   "PreAnalisis Especial"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -71400
         TabIndex        =   33
         Top             =   660
         Width           =   2295
      End
      Begin VB.CheckBox chkPreAnalisis 
         Appearance      =   0  'Flat
         Caption         =   "PreAnalisis"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -72600
         TabIndex        =   32
         Top             =   660
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardarDeudor 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   7440
         TabIndex        =   31
         Top             =   4005
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizaFiadores 
         Caption         =   "..."
         Height          =   255
         Left            =   -74280
         TabIndex        =   30
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2355
         Left            =   5520
         TabIndex        =   28
         Top             =   900
         Width           =   3135
         _ExtentX        =   5525
         _ExtentY        =   4149
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#OP"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripción"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Amortiza"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Int.C"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Int.M"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cuota"
            Object.Width           =   2194
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tipo"
            Object.Width           =   970
         EndProperty
      End
      Begin VB.CommandButton cmdGuardaObservaciones 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   -67680
         TabIndex        =   27
         Top             =   3900
         Width           =   1215
      End
      Begin VB.OptionButton optObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Junta Directiva"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   -69240
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   540
         Width           =   2775
      End
      Begin VB.OptionButton optObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Resolución del Comité"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   -72000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   540
         Width           =   2775
      End
      Begin VB.OptionButton optObservacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Analistas de Crédito"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   540
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.TextBox txtLiquidez 
         Height          =   315
         Left            =   2880
         MaxLength       =   6
         TabIndex        =   23
         Top             =   2940
         Width           =   2535
      End
      Begin VB.TextBox txtPorcSalarioDevengado 
         Height          =   315
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   22
         Top             =   2580
         Width           =   2535
      End
      Begin VB.TextBox txtSalarioNeto 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2100
         Width           =   2535
      End
      Begin VB.TextBox txtCuotaSolicitada 
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1740
         Width           =   2535
      End
      Begin VB.TextBox txtObligacionExCuotas 
         Height          =   315
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   19
         Top             =   1380
         Width           =   2535
      End
      Begin VB.TextBox txtObligacionExSaldos 
         Height          =   315
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   18
         Top             =   1020
         Width           =   2535
      End
      Begin VB.TextBox txtSalarioLiquido 
         Height          =   315
         Left            =   2880
         MaxLength       =   12
         TabIndex        =   17
         Top             =   660
         Width           =   2535
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   50
         Top             =   420
         Width           =   8535
         _Version        =   524288
         _ExtentX        =   15055
         _ExtentY        =   6800
         _StockProps     =   64
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollBars      =   2
         SpreadDesigner  =   "frmPreAnalisis.frx":1566
         AppearanceStyle =   1
      End
      Begin VB.Image imgObligaciones 
         Height          =   375
         Left            =   5400
         Picture         =   "frmPreAnalisis.frx":1B69
         Stretch         =   -1  'True
         ToolTipText     =   "Aumentar/Disminuir"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Emisión"
         Height          =   255
         Left            =   -72600
         TabIndex        =   34
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74880
         X2              =   -66480
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   8640
         Y1              =   3900
         Y2              =   3900
      End
      Begin VB.Label Label2 
         Caption         =   "% Actual de Liquidez"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   2940
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "[25]% Salario Devengado"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   2580
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Salario Neto"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   14
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Cuota Operación Solicitada"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Obligaciones Externas (Cuotas)"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Obligaciones Externas (Saldos)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1020
         Width           =   2415
      End
      Begin VB.Label lblX 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Obligaciones Internas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   5520
         TabIndex        =   10
         Top             =   585
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Salario Liquido"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   660
         Width           =   2415
      End
   End
   Begin VB.Image imgPreAnalisis 
      Height          =   255
      Left            =   3000
      Picture         =   "frmPreAnalisis.frx":2433
      Stretch         =   -1  'True
      ToolTipText     =   "Imprime PreAnalisis"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmPreAnalisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, gCajaTxt As TextBox
Dim vEditar As Boolean, vObservacion(2) As String
Dim mcurLibrado As Currency, mCedula As String
Dim mCodigo As String, mMonto As Currency
Dim vMontoSol As Currency, vBusqueda As Boolean
Dim vMesesMembresia As Integer, vAhorroFiduciario As Currency
Dim vPorcentajeCancelado As Currency, vFoco As Boolean
Dim vCuotasRefundidas As Currency, mcurCuota As Currency
Dim mcurSaldo As Currency, vExpand As Boolean

Function fxAhorro() As Currency
Dim rs As New ADODB.Recordset

strSQL = "select * from par_ahcr"
With rs
   .Open strSQL, glogon.Conection, adOpenStatic
      If .EOF = False Then
         fxAhorro = !CR_POR_AHORRO / 100
      End If
   .Close
End With

End Function

Function fxCreditoAhorro(vOP As Long) As String
Dim rs As New ADODB.Recordset
Dim intMeses As Integer
Dim vFec1 As String
Dim vFec2 As String

strSQL = "Select Plazo,Prideduc,Fecult From Reg_Creditos Where Cedula='" & mCedula & "' "
strSQL = strSQL & "And Codigo='" & mCodigo & "' and Estado='A' And "
strSQL = strSQL & "id_solicitud <> " & vOP
With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   If .EOF = False Then
      vFec1 = "01/" & Mid(CStr(!Fecult), 5, 2) & "/" & Mid(CStr(!Fecult), 1, 4)
      vFec2 = "01/" & Mid(CStr(!Prideduc), 5, 2) & "/" & Mid(CStr(!Prideduc), 1, 4)
      intMeses = Abs(DateDiff("m", vFec1, vFec2))
'      intMeses = (!Fecult - !Prideduc) + 1
      If intMeses < vPorcentajeCancelado Then
         fxCreditoAhorro = "o"
      Else
         fxCreditoAhorro = "þ"
      End If
   Else
      fxCreditoAhorro = " "
   End If
 .Close
End With

End Function

Function fxCreditoAnterior(vOP As Long) As String
Dim rs As New ADODB.Recordset
Dim intMeses As Integer
Dim vFec1 As String
Dim vFec2 As String

strSQL = "Select Plazo,Prideduc,Fecult From Reg_Creditos Where Cedula='" & mCedula & "' "
strSQL = strSQL & "And Codigo='" & mCodigo & "' and Estado='A' And "
strSQL = strSQL & "id_solicitud <> " & vOP
With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   If .EOF = False Then
      vFec1 = "01/" & Mid(CStr(!Fecult), 5, 2) & "/" & Mid(CStr(!Fecult), 1, 4)
      vFec2 = "01/" & Mid(CStr(!Prideduc), 5, 2) & "/" & Mid(CStr(!Prideduc), 1, 4)
      intMeses = Abs(DateDiff("m", vFec1, vFec2))
'      intMeses = (!Fecult - !Prideduc) + 1
      If intMeses < vMesesMembresia Then
         fxCreditoAnterior = "o"
      Else
         fxCreditoAnterior = "þ"
      End If
   Else
      fxCreditoAnterior = " "
   End If
 .Close
End With

End Function

Function fxFianzasActuales(vCedula As String) As Currency
Dim rs As New ADODB.Recordset
Dim rsCreditos As New ADODB.Recordset
Dim rsRef As New ADODB.Recordset
Dim curSaldo As Currency

curSaldo = 0
mcurLibrado = 0

strSQL = "Select * from Fiadores Where cedulaf='" & vCedula & "'"
With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   Do While .EOF = False
      strSQL = "Select Saldo From Reg_Creditos Where Id_Solicitud=" & !ID_SOLICITUD
      strSQL = strSQL & " And Estado='A' And Saldo > 0 "
      rsCreditos.CursorLocation = adUseServer
      rsCreditos.Open strSQL, glogon.Conection, adOpenStatic
       If rsCreditos.EOF = False Then
          curSaldo = curSaldo + rsCreditos!Saldo
       
          strSQL = "Select * from Refundiciones Where id_solicitud=" & !ID_SOLICITUD
          rsRef.CursorLocation = adUseServer
          rsRef.Open strSQL, glogon.Conection, adOpenStatic
          If rsRef.EOF = False Then
             mcurLibrado = mcurLibrado + rsCreditos!Saldo
          End If
          rsRef.Close
       
       End If
      rsCreditos.Close
            
      .MoveNext
   Loop
 .Close
End With

fxFianzasActuales = curSaldo

End Function

Function fxGarantia(vCodigo As String) As String
Dim rs As New ADODB.Recordset

strSQL = "Select * from Pra_Codigos Where Codigo='" & vCodigo & "'"
With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   If .EOF = False Then
      Select Case !GARANTIA
        Case "V"
          fxGarantia = "F"
        Case "F"
          fxGarantia = "F"
        Case "E"
          fxGarantia = "E"
      End Select
   Else
      fxGarantia = "A"
   End If
 .Close
End With

End Function

Function fxMembresia(vGarantia As String) As Variant
Dim rs As New ADODB.Recordset
Dim intMeses As Integer

strSQL = "Select FechaIngreso From Socios Where Cedula='" & mCedula & "'"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
  If rs.EOF = False Then
     intMeses = DateDiff("m", rs!FechaIngreso, fxFechaServidor)
  End If
rs.Close

strSQL = "Select * From Pra_Membresias Where Garantia='" & vGarantia & "' "
strSQL = strSQL & "And " & intMeses & " Between Desde and Hasta"
With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       fxMembresia = intMeses & " Meses - " & !Monto
    End If
 .Close
End With

End Function

Function fxmontoOP(vOP As Long) As Currency
Dim rs As New ADODB.Recordset

fxmontoOP = 0

strSQL = "Select Codigo,Cedula,MontoSol From Reg_Creditos Where id_solicitud=" & vOP
With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   If .EOF = False Then
      fxmontoOP = !montosol
      mMonto = !montosol
      mCedula = Trim(!cedula)
      mCodigo = Trim(!Codigo)
   End If
 .Close
End With

End Function

Function fxRefundiciones(vOP As Long, Optional vTipo As String = "C") As Boolean
Dim rs As New ADODB.Recordset


If vTipo = "C" Then
strSQL = "Select * From Refundiciones Where Id_solicitudR=" & Trim(txtOperacion)
Else
strSQL = "Select * From Refunde_retencion Where Id_solicitudR=" & Trim(txtOperacion)
End If

strSQL = strSQL & " And id_solicitud=" & vOP

With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       fxRefundiciones = True
    Else
       fxRefundiciones = False
    End If
 .Close
End With

End Function

Sub sbComites()
Dim rs As New Recordset, str As String

strSQL = "select * from comites"

With rs
.ActiveConnection = glogon.Conection
.CursorType = adOpenStatic
.Source = strSQL
.Open
 Do While .EOF = False
  cboComite.AddItem !Descripcion
  cboComite.ItemData(cboComite.NewIndex) = !ID_COMITE
  .MoveNext
 Loop
.Close
End With


End Sub

Sub sbFiadores(vOP As Long)
Dim rs As New ADODB.Recordset
Dim i As Integer

strSQL = "Select * From Fiadores Where id_solicitud=" & vOP

With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       vGrid.MaxRows = .RecordCount
    Else
       vGrid.MaxRows = 1
    End If
    For i = 1 To .RecordCount
      vGrid.Row = i
      vGrid.Col = 1
      vGrid.Text = !CedulaF
      
      vGrid.Col = 2
      vGrid.Text = !nombre
      
      vGrid.Col = 3
      vGrid.Text = CCur(!Salario)
      
      vGrid.Col = 4
      vGrid.Text = CCur(!Liquidez)
      
      vGrid.Col = 5
      vGrid.Text = CCur(!Devengado)
      
      vGrid.Col = 6
      vGrid.Text = CCur(!FIA_CONSEC)
      
      .MoveNext
    Next i
 .Close
End With

End Sub

Sub sbObligaciones(vCedula As String)
Dim rs As New ADODB.Recordset
Dim itmX As ListItem

vCuotasRefundidas = 0

strSQL = "Select R.Id_Solicitud,R.Codigo,R.Saldo,R.plazo,R.Cuota,R.amortiza as Amortizado,"
strSQL = strSQL & "C.poliza,C.Descripcion,V.IntC,V.IntM,V.Amortiza,C.retencion "
strSQL = strSQL & "from Reg_Creditos R Inner Join Catalogo C On R.Codigo=C.Codigo "
strSQL = strSQL & "Left Join Vista_Morosidad V On R.id_solicitud=V.id_solicitud "
strSQL = strSQL & "Where R.Cedula='" & vCedula & "' And R.Saldo > 0 "
strSQL = strSQL & "And R.Estado='A' and R.plazo < 900"

With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   Do While Not .EOF
      Set itmX = lsw.ListItems.Add(, , !ID_SOLICITUD)
          itmX.SubItems(1) = Trim(!Codigo)
          itmX.SubItems(2) = Trim(!Descripcion)
          If !retencion = "S" Or !poliza = "S" Then
            itmX.SubItems(8) = "R"
            itmX.SubItems(3) = Format(((!PLAZO * !Cuota) - (!Amortizado + IIf(IsNull(!Amortiza), 0, !Amortiza))), "Standard")
          Else
            itmX.SubItems(8) = "C"
            itmX.SubItems(3) = Format(!Saldo, "Standard")
          End If
          itmX.SubItems(4) = Format(IIf(IsNull(!Amortiza) = True, "0", !Amortiza), "Standard")
          itmX.SubItems(5) = Format(IIf(IsNull(!IntC) = True, "0", !IntC), "Standard")
          itmX.SubItems(6) = Format(IIf(IsNull(!IntM) = True, "0", !IntM), "Standard")
          itmX.SubItems(7) = Format(!Cuota, "Standard")
          
            If fxRefundiciones(!ID_SOLICITUD, itmX.SubItems(8)) Then
               itmX.Checked = True
               vCuotasRefundidas = vCuotasRefundidas + !Cuota
            End If
      .MoveNext
   Loop
 .Close
End With

End Sub

Sub sbRegistroPreanalisis(vOP As Long)
Dim rs As New ADODB.Recordset

strSQL = "Select * From Pra_Principal where id_solicitud=" & vOP

With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       txtSalarioLiquido = Format(!SALARIO_LIQUIDO, "Standard")
       txtObligacionExSaldos = Format(!Obligaciones_Ext_Saldo, "Standard")
       txtObligacionExCuotas = Format(!Obligaciones_Ext_Cuota, "Standard")
       txtPorcSalarioDevengado = Format(!Porc_Salario_Devengado, "Standard")
       txtLiquidez = Format(!Porc_Liquidez, "Standard")
       vObservacion(0) = Trim(IIf(IsNull(!OBSERVACION_ANALISTA), "", !OBSERVACION_ANALISTA))
       vObservacion(1) = Trim(IIf(IsNull(!OBSERVACION_COMITE), "", !OBSERVACION_COMITE))
       vObservacion(2) = Trim(IIf(IsNull(!OBSERVACION_JD), "", !OBSERVACION_JD))
       Call sbSalarioNeto
       vEditar = True
    Else
       vEditar = False
    End If
 .Close
End With

End Sub

Sub sbReporteEspecial(vOP As Long, vSalida As String)
Dim rs As New ADODB.Recordset
Dim rsAhorro As New ADODB.Recordset
Dim curFianzasActuales As Currency
Dim curFianzasNuevas As Currency
Dim vGarantia As String

Me.MousePointer = vbHourglass

frmContenedor.Crt.Reset
frmContenedor.Crt.Connect = glogon.ConectRPT

frmContenedor.Crt.ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_analisisEspecial.rpt")
frmContenedor.Crt.WindowState = crptMaximized

If vSalida = "P" Then
   frmContenedor.Crt.Destination = crptToPrinter
Else
   frmContenedor.Crt.Destination = crptToWindow
End If

vGarantia = fxGarantia(mCodigo)
If vGarantia = "A" Then
 strSQL = "Select Ahorro,Capitaliza From Ahorro_Consolidado Where Cedula='" & mCedula & "'"
 rsAhorro.CursorLocation = adUseServer
 rsAhorro.Open strSQL, glogon.Conection, adOpenStatic
   If rsAhorro.EOF = False Then
      frmContenedor.Crt.Formulas(0) = "Ahorro='" & Format((rsAhorro!Ahorro + rsAhorro!Capitaliza) * fxAhorro, "Standard") & "'"
   End If
 rsAhorro.Close
 frmContenedor.Crt.Formulas(1) = "Leyenda='MONTO POR AHORRO'"
 frmContenedor.Crt.Formulas(2) = "PorcentajeCancelado='" & vPorcentajeCancelado & " PORCIENTO CANCELADO DEL CREDITO ANTERIOR'"
 frmContenedor.Crt.Formulas(3) = "PorcentajeAhorro='" & fxCreditoAhorro(vOP) & "'"
 frmContenedor.Crt.Formulas(4) = "Intereses=" & fxInteresesHastaFormalizar(vOP, fxFechaServidor)
Else
 strSQL = "Select * From Fiadores where id_solicitud=" & vOP
 With rs
  .CursorLocation = adUseServer
  .Open strSQL, glogon.Conection, adOpenStatic
     Do While .EOF = False
        curFianzasActuales = fxFianzasActuales(Trim(!CedulaF))
        curFianzasNuevas = mMonto
                  
        strSQL = "Insert into Pra_Fiadores(id_Solicitud,Cedula,Fianzas_Actuales,"
        strSQL = strSQL & "Fianzas_Libradas,Fianzas_Nuevas) Values("
        strSQL = strSQL & vOP & ",'" & Trim(!CedulaF) & "',"
        strSQL = strSQL & curFianzasActuales & "," & mcurLibrado & ","
        strSQL = strSQL & curFianzasNuevas & ")"
        Call ConectionExecute(strSQL)
                  
        .MoveNext
     Loop
  .Close
 End With
  
 frmContenedor.Crt.Formulas(0) = "Membresia='" & fxMembresia(vGarantia) & "'"
 frmContenedor.Crt.Formulas(1) = "MesesCancelados='" & vMesesMembresia & " MESES CANCELADOS DEL CREDITO ANTERIOR'"
 frmContenedor.Crt.Formulas(2) = "18Meses='" & fxCreditoAnterior(vOP) & "'"

 If vGarantia = "F" Then
  strSQL = "Select (Ahorro + Capitaliza) as Ahorro From Ahorro_Consolidado Where Cedula='" & mCedula & "'"
  rsAhorro.CursorLocation = adUseServer
  rsAhorro.Open strSQL, glogon.Conection, adOpenStatic
   If rsAhorro.EOF = False Then
      frmContenedor.Crt.Formulas(3) = "Ahorro='" & Format(rsAhorro!Ahorro * (vAhorroFiduciario / 100), "Standard") & "'"
   End If
  rsAhorro.Close
  frmContenedor.Crt.Formulas(4) = "Leyenda='MONTO " & vAhorroFiduciario & "% AHORRO'"
  frmContenedor.Crt.Formulas(5) = "PorcentajeCancelado='" & vPorcentajeCancelado & " PORCIENTO CANCELADO DEL CREDITO ANTERIOR'"
  frmContenedor.Crt.Formulas(6) = "PorcentajeAhorro='" & fxCreditoAhorro(vOP) & "'"
 End If

 frmContenedor.Crt.Formulas(7) = "Intereses=" & fxInteresesHastaFormalizar(vOP, fxFechaServidor)

End If

frmContenedor.Crt.Formulas(8) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & " - " & Format(fxFechaCalculo, "dd/mm/yyyy") & "'"

frmContenedor.Crt.Formulas(9) = "CATEGORIA='" & fxCalificacionPersona(txtCedula) & "'"
frmContenedor.Crt.SelectionFormula = "{PRA_PRINCIPAL.ID_SOLICITUD}=" & vOP
frmContenedor.Crt.PrintReport

strSQL = "Delete from pra_fiadores where id_solicitud=" & vOP
Call ConectionExecute(strSQL)

End Sub

Private Sub sbSalarioNeto()

On Error GoTo vError

txtSalarioNeto = ""


If Trim(txtSalarioLiquido) = "" Or Trim(txtObligacionExSaldos) = "" Or _
   Trim(txtObligacionExCuotas) = "" Or txtCuotaSolicitada = "" Then
Else
   txtSalarioNeto = Format((CCur(txtSalarioLiquido) + CCur(txtObligacionExCuotas) _
                    + vCuotasRefundidas) - CCur(txtCuotaSolicitada), "Standard")
End If

vError:

End Sub

Private Sub cboComite_Click()
Dim rs As New ADODB.Recordset

rs.Open "select * from comites where id_comite = " & cboComite.ItemData(cboComite.ListIndex), glogon.Conection, adOpenStatic
txtActa = IIf(IsNull(rs!Acta), 0, rs!Acta)
rs.Close
End Sub


Private Sub cboPreAnalisis_Click()
If Trim(cboPreAnalisis) = "Por Comité/Acta" Then
   fraComite.Enabled = True
   fraSolicitud.Enabled = False
   txtDe = ""
   txtA = ""
Else
   fraComite.Enabled = False
   fraSolicitud.Enabled = True
   txtActa = ""
   txtDe.SetFocus
End If

End Sub


Private Sub cmdGuardaObservaciones_Click()
 
On Error GoTo vError
 
If Trim(txtOperacion) = "" Then
  MsgBox "Faltan Datos", vbExclamation
  Exit Sub
End If



Me.MousePointer = vbHourglass

    strSQL = "Update Pra_Principal SET Observacion_Analista='" & vObservacion(0) _
           & "',Observacion_Comite='" & vObservacion(1) _
           & "',Observacion_JD='" & vObservacion(2) _
           & "' Where id_Solicitud = " & Trim(txtOperacion)
    Call ConectionExecute(strSQL)
    Call Bitacora("Modifica", "Modifico Ficha de Preanalisis a OP# " & Trim(txtOperacion))

Me.MousePointer = vbDefault

MsgBox "Datos Actualizados", vbExclamation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Sub ValidaMonto()
Dim intI As Integer
Dim strChr As String
Dim strCadena As String
Dim blnLetra As Boolean
Dim intPunto As Integer

blnLetra = False
intPunto = 0

If gCajaTxt <> "" Then
   For intI = 1 To Len(gCajaTxt)
       strChr = Mid(gCajaTxt, intI, 1)
       Select Case Asc(strChr)
          Case 45, 48 To 57
               strCadena = strCadena & strChr
          Case 46
            intPunto = intPunto + 1
            If intPunto = 1 Then
               strCadena = strCadena & strChr
            Else
               blnLetra = True
            End If
            
          Case Else
            blnLetra = True
            
       End Select
   Next intI
   
   gCajaTxt = strCadena

If blnLetra = True Then
   gCajaTxt.SelStart = Len(gCajaTxt)
End If

End If

End Sub

Private Sub cmdGuardarDeudor_Click()
Dim i As Integer


If Not (txtOperacion.Tag = "R" Or txtOperacion.Tag = "P") Then Exit Sub


If Trim(txtOperacion) = "" Or Trim(txtSalarioLiquido) = "" Or Trim(txtObligacionExSaldos) = "" _
   Or Trim(txtObligacionExCuotas) = "" Or Trim(txtPorcSalarioDevengado) = "" Or _
   Trim(txtLiquidez) = "" Then
   
  MsgBox "Faltan Datos", vbExclamation
  Exit Sub
Else
  If CCur(txtLiquidez) > 100 Then
     MsgBox "Porcentajes Deudor Incorrectos", vbExclamation
     Exit Sub
  End If
End If

Me.MousePointer = vbHourglass

If vEditar = False Then
 strSQL = "Insert into Pra_Principal(Id_Solicitud,Codigo,Salario_Liquido,"
 strSQL = strSQL & "Obligaciones_Ext_Saldo,Obligaciones_Ext_Cuota,Salario_Neto,"
 strSQL = strSQL & "Porc_Salario_Devengado,Porc_Liquidez,Total_cuota,Total_saldo) Values("
 strSQL = strSQL & Trim(txtOperacion) & ",'" & Trim(txtCodigo) & "',"
 strSQL = strSQL & CCur(txtSalarioLiquido) & "," & CCur(txtObligacionExSaldos) & ","
 strSQL = strSQL & CCur(txtObligacionExCuotas) & "," & CCur(txtSalarioNeto) & ","
 strSQL = strSQL & CCur(txtPorcSalarioDevengado) & "," & CCur(txtLiquidez) & ","
 strSQL = strSQL & mcurCuota & "," & mcurSaldo & ")"
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Registra", "Registro Ficha de Preanalisis a OP# " & Trim(txtOperacion))

 vEditar = True
Else
 strSQL = "Update Pra_Principal SET Codigo='" & Trim(txtCodigo) & "'"
 strSQL = strSQL & ",Salario_Liquido=" & CCur(txtSalarioLiquido)
 strSQL = strSQL & ",Obligaciones_Ext_Saldo=" & CCur(txtObligacionExSaldos)
 strSQL = strSQL & ",Obligaciones_Ext_Cuota=" & CCur(txtObligacionExCuotas)
 strSQL = strSQL & ",Salario_Neto=" & CCur(txtSalarioNeto)
 strSQL = strSQL & ",Porc_Salario_Devengado=" & CCur(txtPorcSalarioDevengado)
 strSQL = strSQL & ",Porc_Liquidez=" & CCur(txtLiquidez)
 strSQL = strSQL & ",Total_cuota=" & mcurCuota
 strSQL = strSQL & ",Total_saldo=" & mcurSaldo
 strSQL = strSQL & " Where id_Solicitud=" & Trim(txtOperacion)
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Modifica", "Modifico Ficha de Preanalisis a OP# " & Trim(txtOperacion))
 
End If

For i = 1 To vGrid.MaxRows
   vGrid.Row = i
   vGrid.Col = 1
   If Trim(vGrid.Text) <> "" Then
      vGrid.Col = 6
      If Trim(vGrid.Text) <> "" Then
        vGrid.Col = 1
        strSQL = "Update Fiadores Set CEDULAF='" & Trim(vGrid.Text) & "',"
        vGrid.Col = 2
        strSQL = strSQL & "NOMBRE='" & Trim(vGrid.Text) & "',"
        vGrid.Col = 3
        strSQL = strSQL & "SALARIO=" & CCur(vGrid.Text) & ","
        vGrid.Col = 4
        strSQL = strSQL & "Liquidez=" & CCur(vGrid.Text) & ","
        vGrid.Col = 5
        strSQL = strSQL & "Devengado=" & CCur(vGrid.Text) & " "
        vGrid.Col = 6
        strSQL = strSQL & "Where Fia_Consec=" & CCur(vGrid.Text)
        Call ConectionExecute(strSQL)
      Else
        vGrid.Col = 1
        strSQL = "Insert into Fiadores(ID_SOLICITUD,CODIGO,CEDULAF,NOMBRE,FIRMA,"
        strSQL = strSQL & "ESTADO,SALARIO,LIQUIDEZ,DEVENGADO)"
        strSQL = strSQL & "Values(" & Trim(txtOperacion) & ",'" & Trim(txtCodigo) & "','"
        strSQL = strSQL & Trim(vGrid.Text) & "','"
        vGrid.Col = 2
        strSQL = strSQL & Trim(vGrid.Text) & "','N','A',"
        vGrid.Col = 3
        strSQL = strSQL & CCur(vGrid.Text) & ","
        vGrid.Col = 4
        strSQL = strSQL & CCur(vGrid.Text) & ","
        vGrid.Col = 5
        strSQL = strSQL & CCur(vGrid.Text) & ")"
        Call ConectionExecute(strSQL)
      End If
   End If
Next i

vGrid.MaxRows = 0
Call sbFiadores(Trim(txtOperacion))

Me.MousePointer = vbDefault
MsgBox "Datos Actualizados", vbExclamation
End Sub

Private Sub cmdReporte_Click()
Dim i As Long
Dim rs As New ADODB.Recordset

Me.MousePointer = vbHourglass
If DatosCompletos = True Then
 
   With frmContenedor.Crt
      .Reset
      .Connect = glogon.ConectRPT
      
      .Destination = crptToPrinter
      .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_nalisis.rpt")
      
      
      .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
      Select Case Trim(cboPreAnalisis)
        Case "Por Comité/Acta"
          If chkPreAnalisis.Value = 1 Then
           .Formulas(3) = "TITULO='PREANALISIS POR ACTA'"
           .SelectionFormula = "({REG_CREDITOS.ACTA} = " & Trim(txtActa) _
             & " AND {REG_CREDITOS.ID_COMITE} = " & cboComite.ItemData(cboComite.ListIndex) & ")"
           .PrintReport
          End If
          
          If chkPreAnalisisEspecial.Value = 1 Then
           strSQL = "Select id_solicitud From reg_creditos where Acta=" & Trim(txtActa)
           strSQL = strSQL & " And Id_comite=" & cboComite.ItemData(cboComite.ListIndex)
           With rs
             .Open strSQL, glogon.Conection, adOpenStatic
               Do While .EOF = False
                  Call sbReporteEspecial(!ID_SOLICITUD, "P")
                  .MoveNext
               Loop
             .Close
           End With
          End If
          
        Case "Por Solicitud"
          If chkPreAnalisis.Value = 1 Then
           .Formulas(3) = "TITULO='PREANALISIS POR SOLICITUD'"
           .SelectionFormula = "({REG_CREDITOS.ESTADOSOL} = 'R' AND " _
            & "{REG_CREDITOS.ID_SOLICITUD} IN " & txtDe & " TO " & txtA & ")"
           .PrintReport
          End If
          
          If chkPreAnalisisEspecial.Value = 1 Then
             For i = CLng(txtDe) To CLng(txtA)
              Call sbReporteEspecial(i, "P")
             Next i
          End If
        End Select
   End With
   
 Else
   MsgBox "Faltan Datos o Datos Incorrectos", vbExclamation
 End If
 Me.MousePointer = vbDefault

End Sub


Public Function DatosCompletos() As Boolean
DatosCompletos = False


If Trim(cboPreAnalisis) = "Por Comité/Acta" Then
  If cboComite.ListIndex <> -1 And Trim(txtActa) = "" Then
     DatosCompletos = True
  ElseIf cboComite.ListIndex <> -1 And Trim(txtActa) <> "" Then
     DatosCompletos = True
  End If
ElseIf Trim(cboPreAnalisis) = "Por Solicitud" Then
  If Trim(txtDe) <> "" And Trim(txtA) <> "" Then
     If CCur(txtDe) <= CCur(txtA) Then
        DatosCompletos = True
     End If
  End If
End If

End Function
Private Sub Form_Load()
Dim rs As New ADODB.Recordset
vBusqueda = False

vModulo = 3

vExpand = False
Call Formularios(Me)
Call RefrescaTags(Me)

If Val(cmdActualizaFiadores.Tag) = 0 Then
   vGrid.Enabled = False
Else
   vGrid.Enabled = True
End If

ssTab.Tab = 0

strSQL = "Select * From Pra_Parametros"
With rs
  .Open strSQL, glogon.Conection, adOpenStatic
     If .EOF = False Then
        vMesesMembresia = !Meses_Transcurridos
        vAhorroFiduciario = !Porc_Fiduciarios
        vPorcentajeCancelado = !Porc_Cancelado
     End If
  .Close
End With

Call sbComites

If vOperacion > 0 Then
  txtOperacion = vOperacion
  txtOperacion_LostFocus
End If


End Sub

Private Sub sbBusqueda(Index As Integer)
vBusqueda = True
txtOperacion = ""

gBusquedas.Columna = "R.cedula"
gBusquedas.Consulta = "select R.id_Solicitud as Operacion,R.codigo,R.cedula,S.nombre,C.descripcion" _
          & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
          & " inner join Catalogo C on R.codigo = C.codigo"
gBusquedas.Filtro = " And Estadosol in('R','P')"


Select Case Index
  
  Case 1 'Busqueda de Cedula, Operaciones
    
    gBusquedas.Columna = "R.cedula"
    gBusquedas.Orden = "R.cedula"
    
  Case 2 'Busqueda por Nombre
    gBusquedas.Columna = "S.nombre"
    gBusquedas.Orden = "S.nombre"
 
  Case 3 'Busqueda por Código
    gBusquedas.Columna = "R.codigo"
    gBusquedas.Orden = "R.codigo"
 
  Case 4 'Busqueda por Descripcion
    gBusquedas.Columna = "C.descripcion"
    gBusquedas.Orden = "C.descripcion"
  
End Select


frmBusquedas.Show vbModal
txtOperacion = gBusquedas.Resultado
txtOperacion_LostFocus


vBusqueda = False

End Sub

Function fxFechaProcesoAnterior(lngFechaActual As Long) As Long
Dim vMes As Integer, vAnio As Long

vMes = CInt(Mid(Trim(CStr(lngFechaActual)), 5, 2))
vAnio = CLng(Mid(Trim(CStr(lngFechaActual)), 1, 4))

If vMes = 1 Then
  vMes = 12
  vAnio = vAnio - 1
Else
 vMes = vMes - 1
End If

fxFechaProcesoAnterior = vAnio & Format(vMes, "00")

End Function


Function fxCalificacionPersona(vCedula As String) As String
Dim strSQL As String, rs As New ADODB.Recordset, vFechaProceso As Long
Dim i As Integer, rsTmp As New ADODB.Recordset, vFecha As Date

'D - Cobro Judicial Activo Actualmente
strSQL = "select isnull(count(*),0) as Existe from reg_creditos R" _
       & " where R.cedula = '" & Trim(vCedula) & "' and R.estado = 'A' and R.proceso in('J')"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  rs.Close
  fxCalificacionPersona = "D"
  Exit Function
End If
rs.Close

'D - Traspaso de Deudas no Reversados
strSQL = "select id_solicitud from reg_creditos R" _
       & " where R.cedula = '" & Trim(vCedula) & "' and R.estado = 'A' and R.proceso in('T')"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 strSQL = "select isnull(count(*),0) as Existe from reg_creditos where Estado = 'A' and Referencia = " & rs!ID_SOLICITUD
 rsTmp.Open strSQL, glogon.Conection, adOpenStatic
 If Not rsTmp.EOF And Not rsTmp.BOF Then
   If rsTmp!Existe > 0 Then
      rs.Close
      rsTmp.Close
      fxCalificacionPersona = "D"
      Exit Function
   End If
 End If
 rsTmp.Close
 rs.MoveNext
Loop
rs.Close


'C - Presenta Morosidad Actualmente
strSQL = "select isnull(count(*),0) as Existe " _
       & " from reg_creditos R inner join Morosidad M on R.id_solicitud = M.id_solicitud" _
       & " where R.cedula = '" & Trim(vCedula) & "' and R.estado = 'A' and M.estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  rs.Close
  fxCalificacionPersona = "C"
  Exit Function
End If
rs.Close

rs.Open "Select *,dbo.MyGetdate() as FechaAlterna from par_ahcr", glogon.Conection, adOpenStatic
  vFecha = IIf(IsNull(rs!cr_fecha_calculo), rs!fechaAlterna, rs!cr_fecha_calculo)
  vFechaProceso = Year(vFecha) & Format(Month(vFecha), "00")
rs.Close

'Seis Periodos Atras
For i = 1 To 5
  vFechaProceso = fxFechaProcesoAnterior(vFechaProceso)
Next i

'B - Ha estado moroso en los ultimos 6 meses

strSQL = "select isnull(count(*),0) as Existe " _
       & " from reg_creditos R inner join Morosidad M on R.id_solicitud = M.id_solicitud" _
       & " where R.cedula = '" & Trim(vCedula) & "' and M.fechap >= " & vFechaProceso
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  rs.Close
  fxCalificacionPersona = "B"
  Exit Function
End If
rs.Close

'A - Está al día
fxCalificacionPersona = "A"

End Function


Private Sub imgObligaciones_Click()
On Error GoTo vError
If vExpand Then
  imgObligaciones.Left = 5520
  lblX.Left = imgObligaciones.Left
  lblX.Width = 3135
  lsw.Left = lblX.Left
  lsw.Width = lblX.Width
  txtSalarioLiquido.Visible = True
  vExpand = False
Else
  imgObligaciones.Left = 120
  lblX.Left = imgObligaciones.Left
  lblX.Width = ssTab.Width - 300
  lsw.Left = lblX.Left
  lsw.Width = lblX.Width
  txtSalarioLiquido.Visible = False
  vExpand = True
End If

salir:
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation, gMsgTitulo
    Resume salir
End Sub

Private Sub imgPreAnalisis_Click()
Dim vFechaIngreso As Date

If fxmontoOP(Trim(txtOperacion)) = 0 Then
   MsgBox "# Operación Incorrecto", vbExclamation
   Exit Sub
End If

Call sbReporteEspecial(Trim(txtOperacion), IIf(chkSalida.Value = 1, "P", "W"))

'Calcula la fecha de ingreso de esta persona, en base a su operacion y estadoactual
vFechaIngreso = fxMemFechaIngeso(CLng(txtOperacion))

With frmContenedor.Crt
    .Reset
    
    .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_nalisis.rpt")
    .WindowShowPrintBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "TITULO='PREANALISIS POR SOLICITUD'"
    .Formulas(2) = "CATEGORIA='" & fxCalificacionPersona(txtCedula) & "'"
    .Formulas(3) = "MEMBRESIA='" & UCase(fxMemCalculo(vFechaIngreso)) & "'"
    
    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD}=" & Trim(txtOperacion)
    If chkSalida.Value = 1 Then .Destination = crptToPrinter
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub optObservacion_Click(Index As Integer)

Select Case Index
 Case 0
  txtObservaciones = vObservacion(0)
 Case 1
  txtObservaciones = vObservacion(1)
 Case 2
  txtObservaciones = vObservacion(2)
End Select

txtObservaciones.SetFocus

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim curObligaciones As Currency
Dim i As Integer
Dim vFecha As Date

vCuotasRefundidas = 0
mcurCuota = 0
mcurSaldo = 0

Me.MousePointer = vbHourglass

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(i).Checked = True Then
     If Trim(lsw.ListItems.Item(i).SubItems(8)) = "C" Then
      curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(3))
      curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(5))
      curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(6))
     Else
      curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(3))
      curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(4))
     End If
     
     If curObligaciones > vMontoSol Then
        MsgBox "Los Saldos A Refundir Sobrepasan Al Monto Solicitado", vbExclamation
        Me.MousePointer = vbDefault
        Exit Sub
     End If
  End If
Next

curObligaciones = 0

vFecha = fxFechaServidor
For i = 1 To lsw.ListItems.Count
    If lsw.ListItems.Item(i).Checked = True Then
       If lsw.ListItems.Item(i).SubItems(8) = "C" Then
          curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(3))
          curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(5))
          curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(6))
          vCuotasRefundidas = vCuotasRefundidas + CCur(lsw.ListItems.Item(i).SubItems(7))
          mcurSaldo = mcurSaldo + CCur(lsw.ListItems.Item(i).SubItems(3)) - CCur(lsw.ListItems.Item(i).SubItems(4))
          mcurSaldo = mcurSaldo + CCur(lsw.ListItems.Item(i).SubItems(5))
          mcurSaldo = mcurSaldo + CCur(lsw.ListItems.Item(i).SubItems(6))
       Else
          curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(3))
          curObligaciones = curObligaciones + CCur(lsw.ListItems.Item(i).SubItems(4))
          vCuotasRefundidas = vCuotasRefundidas + CCur(lsw.ListItems.Item(i).SubItems(7))
          mcurSaldo = mcurSaldo + CCur(lsw.ListItems.Item(i).SubItems(4))
          mcurSaldo = mcurSaldo + CCur(lsw.ListItems.Item(i).SubItems(3))
       End If
       mcurCuota = mcurCuota + CCur(lsw.ListItems.Item(i).SubItems(7))
       
       If fxRefundiciones(lsw.ListItems.Item(i).Text, lsw.ListItems.Item(i).SubItems(8)) = False Then
          If lsw.ListItems.Item(i).SubItems(8) = "C" Then
             strSQL = "Insert into Refundiciones(ID_SOLICITUD,CODIGO,CODIGOR,MONTO,FECHA,"
             strSQL = strSQL & "ID_SOLICITUDR,INTCOR,INTMOR,SALDO_ANTERIOR) Values("
          Else
             strSQL = "Insert into Refunde_Retencion(ID_SOLICITUD,CODIGO,CODIGOR,MONTO,FECHA,"
             strSQL = strSQL & "ID_SOLICITUDR,MORA,SALDO_ANTERIOR) Values("
          End If
          
          strSQL = strSQL & lsw.ListItems.Item(i).Text & ",'"
          strSQL = strSQL & lsw.ListItems.Item(i).SubItems(1) & "','"
          strSQL = strSQL & Trim(txtCodigo) & "'," & CCur(lsw.ListItems.Item(i).SubItems(3))
          strSQL = strSQL & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & Trim(txtOperacion) & ","
          
          If lsw.ListItems.Item(i).SubItems(8) = "C" Then
             strSQL = strSQL & CCur(lsw.ListItems.Item(i).SubItems(5)) & ","
             strSQL = strSQL & CCur(lsw.ListItems.Item(i).SubItems(6)) & ","
             strSQL = strSQL & CCur(lsw.ListItems.Item(i).SubItems(3)) - CCur(lsw.ListItems.Item(i).SubItems(4)) & ")"
          Else
             strSQL = strSQL & CCur(lsw.ListItems.Item(i).SubItems(4)) & ","
             strSQL = strSQL & CCur(lsw.ListItems.Item(i).SubItems(3)) & ")"
          End If
                    
          Call ConectionExecute(strSQL)
          
'          Call Bitacora("Registra", "Refundicion a OP# " & lsw.ListItems.Item(i).Text & " Con OP# " & Trim(txtOperacion))

       End If
    
    Else
       If lsw.ListItems.Item(i).SubItems(8) = "C" Then
          strSQL = "Delete From Refundiciones "
       Else
          strSQL = "Delete From Refunde_Retencion "
       End If
       strSQL = strSQL & " Where id_solicitud = " & lsw.ListItems.Item(i).Text _
              & " and id_solicitudr = " & Trim(txtOperacion)
       Call ConectionExecute(strSQL)
       
'       Call Bitacora("Borra", "Refundicion a OP# " & lsw.ListItems.Item(i).Text & " Con OP# " & Trim(txtOperacion))
    
    End If
Next i

txtObligacionesInSaldos = Format(curObligaciones, "Standard")
Call sbSalarioNeto
Me.MousePointer = vbDefault

End Sub

Private Sub txtA_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
       cmdReporte.SetFocus
  Case Else
       KeyAscii = 0
End Select
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
End Sub

Private Sub txtCuotaSolicitada_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtSalarioNeto.SetFocus
End If
End Sub


Private Sub txtDe_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
       txtA.SetFocus
  Case Else
       KeyAscii = 0
End Select
End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
End Sub

Private Sub txtLiquidez_Change()
If vFoco = True Then
 Set gCajaTxt = txtLiquidez
 Call ValidaMonto
End If
End Sub

Private Sub txtLiquidez_GotFocus()
vFoco = True
End Sub


Private Sub txtLiquidez_LostFocus()
vFoco = False
If Trim(txtLiquidez) <> "" Then
   txtLiquidez = Format(txtLiquidez, "Standard")
End If
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtObligacionExCuotas_Change()
If vFoco = True Then
 Set gCajaTxt = txtObligacionExCuotas
 Call ValidaMonto
 Call sbSalarioNeto
End If
End Sub

Private Sub txtObligacionExCuotas_GotFocus()
vFoco = True
End Sub


Private Sub txtObligacionExCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtCuotaSolicitada.SetFocus
End If
End Sub


Private Sub txtObligacionExCuotas_LostFocus()
vFoco = False
If Trim(txtObligacionExCuotas) <> "" Then
   txtObligacionExCuotas = Format(txtObligacionExCuotas, "Standard")
End If
End Sub

Private Sub txtObligacionExSaldos_Change()
If vFoco = True Then
 Set gCajaTxt = txtObligacionExSaldos
 Call ValidaMonto
 Call sbSalarioNeto
End If
End Sub

Private Sub txtObligacionExSaldos_GotFocus()
vFoco = True
End Sub


Private Sub txtObligacionExSaldos_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   txtObligacionExCuotas.SetFocus
End If

End Sub


Private Sub txtObligacionExSaldos_LostFocus()
vFoco = False
If Trim(txtObligacionExSaldos) <> "" Then
   txtObligacionExSaldos = Format(txtObligacionExSaldos, "Standard")
End If
End Sub

Private Sub txtObservaciones_Change()

If optObservacion.Item(0) = True Then
   vObservacion(0) = txtObservaciones
ElseIf optObservacion.Item(1) = True Then
   vObservacion(1) = txtObservaciones
ElseIf optObservacion.Item(2) = True Then
   vObservacion(2) = txtObservaciones
End If



End Sub

Private Sub txtOperacion_Change()
vFoco = False
If Trim(txtOperacion) = "" Then
   lsw.ListItems.Clear
   txtObligacionesInSaldos = ""
   txtCodigo = ""
   txtDescripcion = ""
   txtCedula = ""
   txtNombre = ""
   
   txtSalarioLiquido = ""
   txtObligacionExSaldos = ""
   txtObligacionExCuotas = ""
   txtCuotaSolicitada = ""
   txtSalarioNeto = ""
   txtPorcSalarioDevengado = ""
   txtLiquidez = ""
   
   vGrid.MaxRows = 0
   txtObservaciones = ""
   
End If
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
     txtCodigo.SetFocus
  Case Else
End Select
End Sub


Private Sub txtOperacion_LostFocus()
Dim rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
txtObligacionesInSaldos = ""
vGrid.MaxRows = 0
txtObservaciones = ""
vObservacion(0) = ""
vObservacion(1) = ""
vObservacion(2) = ""
txtOperacion.Tag = ""

If Trim(txtOperacion) = "" Then
 txtCodigo = ""
 txtDescripcion = ""
 txtCedula = ""
 txtNombre = ""
 
 txtSalarioLiquido = ""
 txtObligacionExSaldos = ""
 txtObligacionExCuotas = ""
 txtCuotaSolicitada = ""
 txtSalarioNeto = ""
 txtPorcSalarioDevengado = ""
 txtLiquidez = ""
  
' If vBusqueda = False Then
'    txtOperacion.SetFocus
' End If

Else
 strSQL = "SELECT R.Id_solicitud,R.MontoSol,R.Codigo,R.Cedula,R.Cuota,S.Nombre,C.Descripcion,R.estadosol "
 strSQL = strSQL & "From Reg_creditos R inner join Socios S "
 strSQL = strSQL & "On R.Cedula=S.Cedula inner join Catalogo C "
 strSQL = strSQL & "On R.Codigo=C.Codigo "
 strSQL = strSQL & "Where R.Id_Solicitud=" & Trim(txtOperacion) & " "
' strSQL = strSQL & "And EstadoSol in('P','R')"

 With rs
 .CursorLocation = adUseServer
 .Open strSQL, glogon.Conection, adOpenStatic
   txtSalarioLiquido = ""
   txtObligacionExSaldos = ""
   txtObligacionExCuotas = ""
   txtCuotaSolicitada = ""
   txtSalarioNeto = ""
   txtPorcSalarioDevengado = ""
   txtLiquidez = ""
  If .EOF = False Then
      txtOperacion.Tag = !estadosol
      txtCodigo = Trim(!Codigo)
      txtDescripcion = Trim(!Descripcion)
      txtCedula = Trim(!cedula)
      txtNombre = Trim(!nombre)
      txtCuotaSolicitada = Format(!Cuota, "Standard")
      vMontoSol = !montosol
      
      Call sbObligaciones(Trim(!cedula))
      Call sbRegistroPreanalisis(!ID_SOLICITUD)
      
      Call tlb_ButtonClick(tlb.Buttons.Item(1))
      Call sbFiadores(!ID_SOLICITUD)
   Else
      txtCodigo = ""
      txtDescripcion = ""
      txtCedula = ""
      txtNombre = ""
      txtOperacion = ""
      txtOperacion.SetFocus
   End If
 .Close
 End With

End If

vError:
Me.MousePointer = vbDefault

End Sub

Private Sub txtPorcSalarioDevengado_Change()
If vFoco = True Then
 Set gCajaTxt = txtPorcSalarioDevengado
 Call ValidaMonto
End If
End Sub

Private Sub txtPorcSalarioDevengado_GotFocus()
vFoco = True
End Sub


Private Sub txtPorcSalarioDevengado_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtLiquidez.SetFocus
End If
End Sub


Private Sub txtPorcSalarioDevengado_LostFocus()
vFoco = False
If Trim(txtPorcSalarioDevengado) <> "" Then
   txtPorcSalarioDevengado = Format(txtPorcSalarioDevengado, "Standard")
End If
End Sub

Private Sub txtSalarioLiquido_Change()

If vFoco = True Then
 Set gCajaTxt = txtSalarioLiquido
' Call ValidaMonto
 Call sbSalarioNeto
End If

End Sub

Private Sub txtSalarioLiquido_GotFocus()
vFoco = True
End Sub


Private Sub txtSalarioLiquido_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtObligacionExSaldos.SetFocus
End If
End Sub


Private Sub txtSalarioLiquido_LostFocus()
vFoco = False
If Trim(txtSalarioLiquido) <> "" Then
   txtSalarioLiquido = Format(txtSalarioLiquido, "Standard")
End If
End Sub

Private Sub txtSalarioNeto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtPorcSalarioDevengado.SetFocus
End If
End Sub


Private Sub vGrid_Advance(ByVal AdvanceNext As Boolean)
Dim intI As Integer

If vGrid.ActiveRow = 1 And vGrid.MaxRows > 1 Then
   Exit Sub
End If

Select Case vGrid.ActiveCol
  Case 5
    For intI = 1 To 5
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = intI

        If Trim(vGrid.Text) = "" Then
           Exit Sub
        End If
    Next
    
    vGrid.MaxRows = vGrid.MaxRows + 1
    
    Call gsbPulsarTecla(vbKeyTab)

End Select

End Sub


