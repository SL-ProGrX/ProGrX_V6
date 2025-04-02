VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_AbonosNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Abonos"
   ClientHeight    =   7272
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9024
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_AbonosNew.frx":0000
   ScaleHeight     =   7272
   ScaleWidth      =   9024
   Begin VB.TextBox txtDocumentoExterno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6840
      TabIndex        =   70
      Top             =   4680
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   65
      Top             =   7020
      Width           =   9024
      _ExtentX        =   15917
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Linea"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Recurso"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cboDiferenciaApl 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCR_AbonosNew.frx":6852
      Left            =   6840
      List            =   "frmCR_AbonosNew.frx":685C
      Style           =   2  'Dropdown List
      TabIndex        =   63
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Frame fraCuotas 
      Caption         =   "Cuotas Activas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   52
      Top             =   840
      Width           =   8775
      Begin MSComctlLib.Toolbar tlbAjuste 
         Height          =   336
         Left            =   4320
         TabIndex        =   66
         Top             =   120
         Width           =   4332
         _ExtentX        =   7641
         _ExtentY        =   593
         ButtonWidth     =   6816
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ajustes de Condiciones en la Tabla de Pagos"
               Key             =   "Ajuste"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8160
         Top             =   360
         _ExtentX        =   995
         _ExtentY        =   995
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_AbonosNew.frx":6889
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   1455
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   8535
         _ExtentX        =   15050
         _ExtentY        =   2561
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "#"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Proceso"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fec.Pago"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Cuota"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Estado"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Int.Cor."
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Int.Mor."
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Principal"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Cargos"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Pólizas"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Dias.Cor."
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Dias.Mor."
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Corte Cta"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkMarcaTodas 
         Appearance      =   0  'Flat
         Caption         =   "Marcar Todas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   2400
         TabIndex        =   54
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.TextBox txtOperacion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   4
      TabIndex        =   37
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   36
      ToolTipText     =   "Nombre Completo del Socio (Apellidos y Nombre)"
      Top             =   960
      Width           =   5775
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   35
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   240
      TabIndex        =   24
      Top             =   1800
      Width           =   8055
      Begin VB.Label lblFecUltMovR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblCuotaR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   46
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblInteresR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   45
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblAmortizaR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblSaldoR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   43
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblFecUltMov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   34
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblAmortiza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   30
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6960
         TabIndex        =   29
         ToolTipText     =   "Si es menor a la fecha de proceso se Utiliza la Fecha de Proceso"
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5640
         TabIndex        =   28
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   48
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraAbono 
      BorderStyle     =   0  'None
      Caption         =   "Tipo de Abono"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   8655
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Adelantos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Ordinario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Cancelación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optAbono 
         BackColor       =   &H80000003&
         Caption         =   "Extra Ordinario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker dtpFechaCancelacion 
         Height          =   315
         Left            =   5040
         TabIndex        =   68
         Top             =   600
         Width           =   1215
         _ExtentX        =   2138
         _ExtentY        =   550
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   193069059
         CurrentDate     =   40310
      End
      Begin VB.Label lblFechaCancelacion 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Abono (Real) por parte del cliente para cancelación...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   69
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Abono   >"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   51
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame fraDatosAbono 
      Caption         =   "Abono"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   6255
      Begin VB.TextBox txtDiferencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   60
         Text            =   $"frmCR_AbonosNew.frx":D0EB
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtTotalCancela 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   $"frmCR_AbonosNew.frx":D0F2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cboTipoPago 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_AbonosNew.frx":D0F9
         Left            =   4320
         List            =   "frmCR_AbonosNew.frx":D109
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtCuotas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtDatosAmortiza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Text            =   "frmCR_AbonosNew.frx":D130
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cboTipo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCR_AbonosNew.frx":D134
         Left            =   4320
         List            =   "frmCR_AbonosNew.frx":D136
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtTotalPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   $"frmCR_AbonosNew.frx":D138
         Top             =   2040
         Width           =   1695
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   255
         Left            =   2520
         TabIndex        =   67
         Top             =   360
         Width           =   495
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin VB.Label lblPolizas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   61
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Cuotas de las Pólizas Asociadas..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   62
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label27 
         Caption         =   "Diferencia ...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   59
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   58
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblDatosCargos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   57
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Compromiso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   56
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label23 
         Caption         =   "Tipo - Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   19
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "Total a Pagar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label27 
         Caption         =   "# Cuotas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblDatosInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo - Doc"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblDatosAnticipo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label26 
         Caption         =   "Cargo por Cancelación Anticipada..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdReporte 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CheckBox chkRecalculaCuota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Recalcular Cuota"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton CmdAbono 
      Caption         =   "&Aplicar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7200
      Picture         =   "frmCR_AbonosNew.frx":D13F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Image imgDocumento 
      Height          =   192
      Left            =   4920
      Picture         =   "frmCR_AbonosNew.frx":D277
      ToolTipText     =   "Re-Emisión de Comprobante"
      Top             =   360
      Width           =   192
   End
   Begin VB.Label Label24 
      Caption         =   "No.Documento Externo..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   71
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label23 
      Caption         =   "Aplicar diferencias como..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   64
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Label lblOpex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   7560
      TabIndex        =   42
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   960
      TabIndex        =   41
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   40
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2640
      TabIndex        =   39
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Identifica."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   240
      TabIndex        =   38
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8880
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmCR_AbonosNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vCuotasDeducidas As Integer, vCuotasDirectas As Integer
Dim vInteres As Currency, vPlazo As Integer, vSaldoMes As Currency, vUltimoRecibo As Long
Dim vRetencion As Boolean, vBaseCalculo As String, vPrideduc As Long, vAnticipoPorc As Currency, vAnticipoMeses As Integer
Dim vDiasActivo As Long, vFechaHoy As Date, vScroll As Boolean

Private Sub cboDiferenciaApl_Click()

If Not cboDiferenciaApl.Enabled Then Exit Sub

If cboDiferenciaApl.Text = "Abono Extraordinario" Then
   chkRecalculaCuota.Enabled = True
Else
   chkRecalculaCuota.Enabled = False
End If

End Sub

Private Sub cboTipoPago_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then CmdAbono.SetFocus
End Sub


Private Sub chkMarcaTodas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkMarcaTodas.Value
Next

If lsw.ListItems.Count > 0 Then
    Call lsw_ItemCheck(lsw.ListItems.Item(1))
End If

End Sub

Private Sub chkRecalculaCuota_Click()

If vRetencion Then
   chkRecalculaCuota.Value = vbUnchecked
   MsgBox "Las retenciones no se pueden Ajustar para Recálculos, verifique...", vbExclamation
   Exit Sub
End If

Call txtTotalPagar_Change

End Sub

Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngRecibo As Long, vCuenta As String
Dim vTipo As String, vFecha As Date
Dim i As Integer, vExtraOrdinario As Boolean


Me.MousePointer = vbHourglass

On Error GoTo vError


vFecha = fxFechaServidor
vExtraOrdinario = False

vTipo = fxTipoASEDoc(cboTipo.Text)

vCuenta = Trim(fxDocumentoCuenta(vTipo))

lngRecibo = fxDocumentoConsecutivo(vTipo)

vUltimoRecibo = lngRecibo

If vAseDocValido = False Then
  MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
        & " válida para esta operación...", vbCritical
  Exit Sub
End If

'Inicia Transaccion
glogon.Conection.BeginTrans


Select Case True
  Case optAbono(0) 'Abono Ordinario
  
        If Not cboDiferenciaApl.Enabled Then
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD001','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                   & "','" & lngRecibo & "'," & CCur(txtTotalPagar.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD001','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                   & "','" & lngRecibo & "'," & CCur(txtTotalCancela.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
           
           Select Case cboDiferenciaApl.Text
             Case "Adelanto de Cuota"
                strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                       & "','" & lngRecibo & "'," & Abs(CCur(txtDiferencia.Text)) & ",'" _
                       & Format(vFecha, "yyyy/mm/dd") & "',''"
                Call ConectionExecute(strSQL)

             Case "Abono Extraordinario"
                'Calcula Datos del Abono Extraordinario (Dias, Intereses, Cargos, Principal)
                strSQL = "exec spCrdPlanPagosInfoExtraordinario " & vOperacion & "," & Abs(CCur(txtDiferencia.Text)) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
                Call OpenRecordSet(rs, strSQL)
                    'Aplica Cargo por Anticipo
                    If rs!Cargos > 0 Then
                       strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & rs!Cargos & ",'" & GLOBALES.gOficinaUnidad _
                              & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
                       Call ConectionExecute(strSQL)
                    End If
                    'Aplica Abono Extraordinario
                    strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                           & "','" & lngRecibo & "'," & rs!Dias & "," & rs!Intereses & "," & rs!Principal _
                           & "," & rs!Cargos & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
                    Call ConectionExecute(strSQL)
                rs.Close
                
                vExtraOrdinario = True
           End Select
        
        End If
  
  Case optAbono(1) 'Abono Extraordinario
        'Elimina Cuotas Activas, Registra Abono y Recalcula Plan de Pagos
        'Se Supone que solo queda una cuota activa para poder realizar un ab. extraordinario
        
        If CCur(lblDatosAnticipo.Caption) > 0 Then
           strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & CCur(lblDatosAnticipo.Caption) & ",'" & GLOBALES.gOficinaUnidad _
                  & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
           Call ConectionExecute(strSQL)
        End If
        
        strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
               & "','" & lngRecibo & "'," & vDiasActivo & "," & CCur(lblDatosInteres.Caption) & "," & CCur(txtDatosAmortiza.Text) _
               & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
        Call ConectionExecute(strSQL)

        vExtraOrdinario = True
        
  Case optAbono(2) 'Cancelacion
        'Actualiza el estado de la morosidad
'        strSQL = "exec spCrdPlanPagosMoraActualizaOp " & vOperacion & ",'" & Format(vFecha, "yyyy/mm/dd") & "'"
        
        strSQL = "exec spCrdPlanPagosMoraActualizaOp " & vOperacion & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
        Call ConectionExecute(strSQL)
        
        If CCur(lblDatosAnticipo.Caption) > 0 Then
           strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & CCur(lblDatosAnticipo.Caption) & ",'" & GLOBALES.gOficinaUnidad _
                  & "','" & GLOBALES.gOficinaCentroCosto & "','Cancelacion Anticipada','" & glogon.Usuario & "','CA','','',0"
           Call ConectionExecute(strSQL)
        End If
'        strSQL = "exec spCrdPlanPagoAbonoCancelacion " & vOperacion & ",'CRD003','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
'               & "','" & lngRecibo & "'," & CCur(txtTotalPagar.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
        
        strSQL = "exec spCrdPlanPagoAbonoCancelacion " & vOperacion & ",'CRD003','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
               & "','" & lngRecibo & "'," & CCur(txtTotalPagar.Text) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "',''"
        Call ConectionExecute(strSQL)
  
  
  Case optAbono(3) 'Adelanto de Cuotas
       'Activa Nuevas Cuotas y luego las abona
'        strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
'               & "','" & lngRecibo & "'," & CCur(txtTotalPagar.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
'        Call ConectionExecute(strSQL)
       

        If Not cboDiferenciaApl.Enabled Then
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                   & "','" & lngRecibo & "'," & CCur(txtTotalPagar.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                   & "','" & lngRecibo & "'," & CCur(txtTotalCancela.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
           
           Select Case cboDiferenciaApl.Text
             Case "Adelanto de Cuota"
                strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                       & "','" & lngRecibo & "'," & Abs(CCur(txtDiferencia.Text)) & ",'" _
                       & Format(vFecha, "yyyy/mm/dd") & "',''"
                Call ConectionExecute(strSQL)

             Case "Abono Extraordinario"
                'Calcula Datos del Abono Extraordinario (Dias, Intereses, Cargos, Principal)
                strSQL = "exec spCrdPlanPagosInfoExtraordinario " & vOperacion & "," & Abs(CCur(txtDiferencia.Text)) _
                       & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
                Call OpenRecordSet(rs, strSQL)
                    'Aplica Cargo por Anticipo
                    If rs!Cargos > 0 Then
                       strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & rs!Cargos & ",'" & GLOBALES.gOficinaUnidad _
                              & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
                       Call ConectionExecute(strSQL)
                    End If
                    'Aplica Abono Extraordinario
                    strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & fxTipoASENumero(vTipo) _
                           & "','" & lngRecibo & "'," & rs!Dias & "," & rs!Intereses & "," & rs!Principal _
                           & "," & rs!Cargos & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
                    Call ConectionExecute(strSQL)
                rs.Close
                
                vExtraOrdinario = True
           End Select
        
        End If



End Select


'Cierra Transaccion
glogon.Conection.CommitTrans

'Indica si debe reprocesar el Plan de Pagos por registro de Abonos Extraordinario
If vExtraOrdinario Then
        strSQL = "exec spCrdPlanPagos " & vOperacion
        Call ConectionExecute(strSQL)
End If

'Genera el Comprobante
Select Case True
  Case optAbono(0) 'Abono Ordinario
      Call Bitacora("Registra", "Abono Ordinario a la Operacion : " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("ABONO ORDINARIO", vTipo, CStr(lngRecibo), "CRD001", vCuenta)
  Case optAbono(1) 'Abono Extraordinario
        Call Bitacora("Registra", "Abono ExtraOrd. " & IIf((chkRecalculaCuota.Value = 1), "Con Recal.", "Sin Recal") & " a la Op.: " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("ABONO EXTRAORDINARIO", vTipo, CStr(lngRecibo), "CRD002", vCuenta)
  Case optAbono(2) 'Abono De Cancelacion
      Call Bitacora("Registra", "Cancelación de la Operacion : " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("CANCELACION DE DEUDA", vTipo, CStr(lngRecibo), "CRD003", vCuenta)
  Case optAbono(3) 'Adelanto de Cuotas
      Call Bitacora("Registra", "Adelanto de Cuotas de la Operacion : " & vOperacion)
      If uRecibos Then lngRecibo = fxDocumentoAbono("ADELANTO DE CUOTAS", vTipo, CStr(lngRecibo), "CRD004", vCuenta)
End Select



'IMPRIMIR RECIBO
If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, vTipo)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxVerifica() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String, i As Integer

vMensaje = ""

strSQL = "select cod_divisa from reg_creditos where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
If rs!cod_Divisa = "DOL" Then
   vMensaje = vMensaje & "-Esta Operación es en dólares! Debe usar el módulo de Cajas para realizar movimientos" & vbCrLf
End If
rs.Close

'Valida texto del documento
If Not fxValidaCriterio(txtDocumentoExterno.Text) Then
      vMensaje = vMensaje & "- El No.Documento Externo no es admitido!, verifique..." & vbCrLf
End If

'Verifica el proceso
If txtProceso.Tag = "J" Then
   If Not fxCrdAbonosAutorizados(txtCodigo.Text, txtProceso.Tag) Then
      vMensaje = vMensaje & "- El usuario actual no cuenta con permisos para realizar abonos a Creditos en Cobro Judicial, verifique..." & vbCrLf
   End If
End If

'Verifica que la diferencia del Monto a Cancelar no supere el Saldo
If CCur(txtDiferencia.Text) < 0 Then
 If CCur(lblSaldoR.Caption) + CCur(txtDiferencia.Text) < 0 Then
      vMensaje = vMensaje & "- La diferencia supera el saldo!, verifique..." & vbCrLf
 End If
End If

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  vMensaje = vMensaje & "- Esta Persona se encuentra CONGELADA, verifique..." & vbCrLf
End If

If vOperacion = 0 Then
  vMensaje = vMensaje & "- Número de Operacion no es válido..." & vbCrLf
End If
 
 
'Verifica Saldo Actual
If Not fxCrdSaldoVerifica(vOperacion, CCur(lblSaldo.Caption)) Then
   vMensaje = vMensaje & "- Esta Operación ha sido modificada, actualice los datos nuevamente antes de realizar el abono..." & vbCrLf
End If
 
If Not vRetencion Then
    If CCur(txtDatosAmortiza) > CCur(lblSaldo.Caption) Then
       vMensaje = vMensaje & "- La Amortización es mayor al Saldo Actual..." & vbCrLf
    End If
Else
    If CCur(txtDatosAmortiza) > ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) Then
        vMensaje = vMensaje & "- La Amortización es mayor que el Remanente a Recaudar : " _
              & ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) & vbCrLf
     End If
End If

If Not IsNumeric(txtTotalPagar.Text) Then
  vMensaje = vMensaje & "- El total a pagar no es un dato válido...verifique...!" & vbCrLf
Else
 If CCur(txtTotalPagar.Text) <= 0 Then
      vMensaje = vMensaje & "- El total a pagar no es un dato válido...verifique...!" & vbCrLf
 End If
End If


'Abono Ordinario (Verificar Secuencia de Check's)
If optAbono.Item(0).Value Then
 For i = 1 To lsw.ListItems.Count
   If i = 1 And Not lsw.ListItems.Item(i).Checked Then
      vMensaje = vMensaje & "- No se ha especificado un orden válido de aplicación de cuotas...!" & vbCrLf
      Exit For
   End If
   
   If i > 1 Then
        If lsw.ListItems.Item(i).Checked And Not lsw.ListItems.Item(i - 1).Checked Then
               vMensaje = vMensaje & "- No se ha especificado un orden válido de aplicación de cuotas...!" & vbCrLf
               Exit For
        End If
   End If
 Next
End If



If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub CmdAbono_Click()
Dim iRespuesta As Integer

If Not fxVerifica Then Exit Sub

 iRespuesta = MsgBox("Esta seguro de realizar el abono a esta Operación " & vOperacion, vbYesNo)
 If iRespuesta = vbYes Then
  
  Call sbAbono
  If vAseDocValido Then MsgBox "Abono Realizado ... " & cboTipo.Text & " #" & vUltimoRecibo, vbInformation
  Call sbConsultaOperacion
 
 Else 'Respuesta
  
  MsgBox "Transacción Cancelada...", vbInformation
 
 End If

End Sub

Private Sub sbReporte(vTitulo As String)
If vOperacion = 0 Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Módulo de Crédito"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Credito_BoletaAbono.rpt")
 
 .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "usuario='" & glogon.Usuario & "'"
 If optAbono(0).Value = True Then
  .Formulas(4) = "tipo_abono='ABONO ORDINARIO : CUOTAS: " & txtCuotas & "'"
 Else
  .Formulas(4) = "tipo_abono='ABONO EXTRAORDINARIO'"
 End If
 .Formulas(5) = "saldo_actual='" & lblSaldo.Caption & "'"
 .Formulas(6) = "amortizacion='" & Me.lblAmortiza.Caption & "'"
 .Formulas(7) = "interesc='" & lblInteres.Caption & "'"
 .Formulas(8) = "fecult='" & Format(lblFecUltMov.Caption, "####-##") & "'"

 .Formulas(9) = "saldo_res='" & lblSaldoR.Caption & "'"
 .Formulas(10) = "amortizacion_res='" & Me.lblAmortizaR.Caption & "'"
 .Formulas(11) = "interesc_res='" & lblInteresR.Caption & "'"
 .Formulas(12) = "fecult_res='" & Format(lblFecUltMovR.Caption, "####-##") & "'"
 
 .Formulas(13) = "abono_amortizacion='" & txtDatosAmortiza & "'"
 .Formulas(14) = "abono_interes='" & lblDatosInteres.Caption & "'"
 .Formulas(15) = "abono_total='" & txtTotalPagar.Text & "'"
 
 .Formulas(16) = "titulo='" & vTitulo & "'"
 .Formulas(17) = "operacion='" & vOperacion & "'"
 .Formulas(18) = "cedula='" & txtCedula & " - " & txtNombre & "'"
 .Formulas(19) = "codigo='" & txtCodigo & " - " & lblDescripcion.Caption & "'"
 
 .PrintReport
End With
Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()

Call sbReporte("ABONO A REALIZAR")

End Sub

Private Sub dtpFechaCancelacion_Change()

If dtpFechaCancelacion.Enabled Then
   'Refresca información base para Cancelación y/o Abonos Extraordinarios
   Select Case True
      Case optAbono.Item(1).Value 'Abono Extraordinario
            Call optAbono_Click(1)
      Case optAbono.Item(2).Value 'Cancelación
            Call optAbono_Click(2)
   End Select
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim vNumCuota As Integer

On Error GoTo vError


vNumCuota = txtCuotas.Text

If vScroll Then
    If FlatScrollBar.Value = 1 Then
       vNumCuota = vNumCuota + 1
    Else
       vNumCuota = vNumCuota - 1
    End If
End If

If vNumCuota <= 0 Then vNumCuota = 1

txtCuotas.Text = vNumCuota

vScroll = False
    FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxValidaCriterio(pCadena As String) As Boolean
Dim vResultado As Boolean, vMensaje As String

pCadena = UCase(pCadena)

vResultado = True
If InStr(1, pCadena, "SELECT") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "DELETE") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "UPDATE") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "INSERT") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "EXEC") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "DROP") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "CREATE") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "ALTER") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "sp_") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "'") > 0 And vResultado Then vResultado = False


If Not vResultado Then
 'Registrar en Log de Seguridad todo el criterio
 MsgBox "!Error: El criterio de busqueda contiene información o datos que pueden afectar potencialmente la integridad de la información..!", vbExclamation
End If

fxValidaCriterio = vResultado

End Function


Private Sub Form_Activate()
 vModulo = 3
 Call RefrescaTags(Me)
End Sub

Private Sub Form_Load()
Dim iDias As Integer

 vModulo = 3
 vOperacion = 0
 
 vFechaHoy = fxFechaServidor
 iDias = fxCrdParametro("32")
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbDocumentosCombo(cboTipo)
 Call sbLimpiaDatos

dtpFechaCancelacion.Value = vFechaHoy
dtpFechaCancelacion.MinDate = DateAdd("d", (iDias * -1), dtpFechaCancelacion.Value)
dtpFechaCancelacion.MaxDate = dtpFechaCancelacion.Value


End Sub

Private Sub sbConsultaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

Call sbLimpiaDatos
 
strSQL = "select R.id_solicitud,R.saldo, R.saldo - isnull(V.amortiza,0) As Saldo_mes,R.proceso" _
       & ",R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult,R.Prideduc" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas, datediff(m,R.fechaforp,dbo.MyGetdate()) as 'Meses'" _
       & ",S.nombre,C.descripcion,C.retencion,C.poliza,R.fechaforp,C.PORC_CARGO_CANCELACION,C.ANTICIPO_MESES,R.Base_Calculo" _
       & ",dbo.fxCrdPlanPagosDiasActivo(" & vOperacion & ") as 'DiasActivo', dbo.fxCrdOperacionTagReg(R.id_solicitud,'S15') as 'AutPagoAnt'" _
       & ",C.descripcion as 'LineaDesc',Ofi.descripcion as 'OficinaDesc',Pre.Descripcion as 'RecursoDesc',dbo.MyGetdate() as 'FechaServer'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join Sif_Oficinas Ofi on R.cod_Oficina_R = Ofi.cod_Oficina" _
       & " left join CATALOGO_GRUPOS Pre on R.cod_grupo = Pre.cod_grupo" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " where R.estado = 'A' and R.saldo > 0" _
       & " and R.ID_SOLICITUD = " & vOperacion
       
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  vBaseCalculo = Trim(rs!Base_Calculo)
  vPrideduc = rs!PriDeduc
  vOperacion = rs!id_solicitud
  vPlazo = rs!Plazo
  vDiasActivo = rs!DiasActivo
  
  'Indica si Aplica Cargo por Cancelacion Anticipada y no se encuentra autorizado debe de cobrarse
  If rs!Meses <= rs!ANTICIPO_MESES And rs!AutPagoAnt = 0 Then
     vAnticipoPorc = rs!PORC_CARGO_CANCELACION / 100
  Else
     vAnticipoPorc = 0
  End If
  
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  If IsNull(rs!saldo_mes) Then
    vSaldoMes = rs!Saldo
    strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!id_solicitud
    Call ConectionExecute(strSQL)
  Else
    If rs!saldo_mes = 0 Then
        vSaldoMes = rs!Saldo
        strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!id_solicitud
        Call ConectionExecute(strSQL)
    Else
       vSaldoMes = rs!saldo_mes
    End If
  
  End If
  
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     lblDatosInteres.Caption = 0
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")
    
     lblSaldo.Tag = rs!FechaForp
     lblSaldo.Caption = Format(rs!Saldo, "Standard")
     lblSaldoR.Caption = 0
    
     txtCuotas = 0
     txtOperacion = rs!id_solicitud
     cboTipoPago.Text = "Efectivo"
     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    txtCedula = rs!Cedula
    txtNombre = rs!Nombre
    txtCodigo = rs!Codigo
    
    txtProceso.Tag = rs!Proceso
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
    
    
    lblDescripcion.Caption = rs!Descripcion
    
    lblDatosAnticipo.ToolTipText = "% de Comision : " & vAnticipoPorc
    lblDatosAnticipo.Tag = vAnticipoPorc
    
   
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    'Barra de Estado
   
    StatusBarX.Panels.Item(1).Text = rs!OficinaDesc & ""
    StatusBarX.Panels.Item(2).Text = rs!LineaDesc & ""
    StatusBarX.Panels.Item(3).Text = rs!RecursoDesc & ""
        
        
       
    'Consulta Cuotas Activas
    strSQL = "select * from CRD_OPERACION_TRANSAC where estado = 'A' and id_solicitud = " & rs!id_solicitud _
           & " order by num_cuota"
      
    rs.Close
    Call OpenRecordSet(rs, strSQL)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!num_cuota)
          itmX.SubItems(1) = Format(rs!Fecha_Proceso, "####-##")
          itmX.SubItems(2) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!Cuota, "Standard")
          itmX.SubItems(4) = IIf((rs!Mora_Dias > 0), "En Mora", "Al Día")
          itmX.SubItems(5) = Format(rs!IntCor, "Standard")
          itmX.SubItems(6) = Format(rs!IntMor, "Standard")
          itmX.SubItems(7) = Format(rs!Principal, "Standard")
          itmX.SubItems(8) = Format(rs!Cargos, "Standard")
          itmX.SubItems(9) = Format(rs!Poliza, "Standard")
          itmX.SubItems(10) = rs!Dias_calculo
          itmX.SubItems(11) = rs!Mora_Dias
          itmX.SubItems(12) = Format(rs!fecha_corte, "yyyy/mm/dd")
          
          itmX.Tag = rs!Id_seq
      rs.MoveNext
    Loop
    
    
    'Activacion de Tipos de Abonos
    
    Select Case lsw.ListItems.Count
      Case Is <= 0
            optAbono.Item(0).Enabled = False 'Ordinario
            optAbono.Item(1).Enabled = True 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = True 'Adelantos
            Call optAbono_Click(1)
      
      Case Is = 1
            optAbono.Item(0).Enabled = True 'Ordinario
            optAbono.Item(1).Enabled = True 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = True 'Adelantos
            Call optAbono_Click(0)
      
      Case Is > 1
            optAbono.Item(0).Enabled = True 'Ordinario
            optAbono.Item(1).Enabled = False 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = False 'Adelantos
            Call optAbono_Click(0)
    End Select
    

Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
 lblDatosAnticipo.Caption = 0
 lblAmortiza.Caption = 0
 lblAmortizaR.Caption = 0
 lblCuota = 0
 lblCuotaR.Caption = 0
 txtDatosAmortiza = 0
 lblDatosInteres.Caption = 0
 lblFecUltMov.Caption = 0
 lblFecUltMovR.Caption = 0
 lblInteres.Caption = 0
 lblInteresR.Caption = 0
 lblDescripcion.Caption = ""
 lblOpex.Caption = ""
 lblSaldo.Caption = 0
 lblSaldoR.Caption = 0
 
 lblPolizas.Caption = 0
 
 txtCedula = ""
 txtCodigo = ""
 txtCuotas = 0
 txtNombre = ""
 txtOperacion = ""
 cboTipoPago.Text = "Efectivo"
 cboDiferenciaApl.Text = "Adelanto de Cuota"
 cboDiferenciaApl.Enabled = False
 txtTotalPagar.Text = 0
 txtTotalCancela.Text = 0
 
 txtProceso.Tag = ""
 txtProceso.Text = ""
 
 fraAbono.Enabled = False
 fraDatosAbono.Enabled = False
 
 fraCuotas.Visible = False
 lsw.ListItems.Clear
 
 chkRecalculaCuota.Value = vbUnchecked
 
 
dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False
dtpFechaCancelacion.Value = vFechaHoy

 
 
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
gBusquedas.Consulta = "Select R.id_solicitud as Operacion,R.Codigo,S.Cedula,S.Nombre,C.Descripcion" _
          & " from REG_CREDITOS R inner join SOCIOS S on R.cedula = S.cedula" _
          & " inner join Catalogo C on R.codigo = C.codigo"
gBusquedas.Columna = "R.CEDULA"
gBusquedas.Orden = "R.CEDULA"
gBusquedas.Filtro = " AND R.ESTADO = 'A'"

frmBusquedas.Show vbModal

txtOperacion = Trim(gBusquedas.Resultado)
vOperacion = txtOperacion

gBusquedas.Consulta = ""
gBusquedas.Columna = ""
gBusquedas.Orden = ""
gBusquedas.Resultado = ""
gBusquedas.Filtro = ""

Call sbConsultaOperacion

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaOperacionCodCed(vCedula As String, vCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select R.id_solicitud,R.saldo,R.saldo_mes,R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas,C.retencion,C.poliza " _
       & "from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & "where R.estado = 'A' and R.proceso <> 'N' and R.saldo > 0 " _
       & "and R.cedula = '" & txtCedula & "' and R.codigo = '" & txtCodigo & "'"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vOperacion = rs!id_solicitud
  vPlazo = rs!Plazo
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  vSaldoMes = IIf(IsNull(rs!saldo_mes), rs!Saldo, rs!saldo_mes)
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     lblDatosInteres.Caption = 0
    
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")
     lblSaldo.Caption = Format(vSaldoMes, "Standard")
     lblSaldoR.Caption = 0
     txtCuotas = 0
     txtOperacion = rs!id_solicitud
     cboTipoPago.Text = "Efectivo"
     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    
    optAbono(0).Enabled = True
    optAbono(1).Enabled = True
    
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    If Not vRetencion Then
        Select Case True
         Case optAbono(0).Value
           Call optAbono_Click(0)
         Case optAbono(1).Value
           Call optAbono_Click(1)
        End Select
    Else
           Call optAbono_Click(0)
           optAbono(1).Enabled = False
    End If
    
    
Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontrarón operaciones para abonos con esta cédula y código", vbInformation
End If
rs.Close

End Sub

Private Sub imgDocumento_Click()
  frmCR_AbonosComprobante.Show vbModal
End Sub


Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim curInteres As Currency, curPrincipal As Currency, curCargos As Currency, curPolizas As Currency
Dim i As Integer

curInteres = 0
curPrincipal = 0
curCargos = 0
curPolizas = 0

With lsw.ListItems
  For i = 1 To .Count
    If .Item(i).Checked Then
       curInteres = curInteres + CCur(.Item(i).SubItems(5)) + CCur(.Item(i).SubItems(6))
       curPrincipal = curPrincipal + CCur(.Item(i).SubItems(7))
       curCargos = curCargos + CCur(.Item(i).SubItems(8))
       curPolizas = curPolizas + CCur(.Item(i).SubItems(9))
    End If
  Next i
End With

txtDatosAmortiza.Text = Format(curPrincipal, "Standard")
lblDatosInteres.Caption = Format(curInteres, "Standard")
lblDatosCargos.Caption = Format(curCargos, "Standard")
lblPolizas.Caption = Format(curPolizas, "Standard")
lblDatosAnticipo.Caption = 0

txtTotalPagar.Text = Format(curPrincipal + curInteres + curCargos + curPolizas, "Standard")
txtTotalCancela.Text = txtTotalPagar.Text
End Sub



Private Sub tlbAjuste_ButtonClick(ByVal Button As MSComctlLib.Button)
  GLOBALES.gTag = txtOperacion.Text
  frmCR_MoraCargosAjustes.Show vbModal
  
  'Verifica si recibio modificaciones, en cuyo caso procede a actualizar datos en pantalla
  If GLOBALES.gTag2 = 1 Then
    Call sbConsultaOperacion
  End If
  
End Sub

Private Sub txtTotalPagar_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim ProcesosTmp As Currency, lngFecha As Currency, iPlazoRst As Integer, curCuota As Currency

On Error Resume Next

If chkRecalculaCuota.Value = vbChecked Then
  
    ' strSQL = "select plazo + DATEDIFF(mm,  dbo.MyGetdate(), CONVERT(DATETIME, substring(convert(varchar(6), prideduc), 1,4) + '/' + substring(convert(varchar(6), prideduc), 5,2) + '/28' )) as PlazoFaltante" _
    '       & " from reg_creditos where id_solicitud = " & txtOperacion
    ' Call OpenRecordSet(rs, strSQL)
    '    lblCuotaR.Caption = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), rs!PlazoFaltante, vInteres)
    ' rs.Close
       lngFecha = lblFecUltMovR.Caption
       If lngFecha < vPrideduc Then lngFecha = vPrideduc
      
       ProcesosTmp = vPrideduc
       iPlazoRst = 1
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       curCuota = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), iPlazoRst, vInteres)
       lblCuotaR.Caption = Format(curCuota, "Standard")
Else
  lblCuotaR.Caption = lblCuota.Caption
End If

End Sub

Private Sub optAbono_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curInteres As Currency, curIntMor As Currency, curPrincipal As Currency, curCargos As Currency
Dim vFecha As Date, vProceso As Long, i As Integer


Me.MousePointer = vbHourglass

fraCuotas.Visible = False
fraDatosAbono.Enabled = True

chkRecalculaCuota.Enabled = False
chkRecalculaCuota.Value = vbUnchecked

'&H00C0FFC0&
txtTotalPagar.BackColor = &HC0FFC0

txtTotalPagar.Locked = True

txtCuotas.Enabled = False
FlatScrollBar.Enabled = txtCuotas.Enabled

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = False
Next

dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False

Select Case Index

 Case 0 'Ordinario
   lblDatosCargos.Caption = 0
   lblDatosInteres.Caption = 0
   lblDatosAnticipo.Caption = 0
   lblPolizas.Caption = 0
   txtDatosAmortiza = 0
      
   txtTotalCancela.Text = 0
   txtTotalPagar.Text = 0
   
   txtCuotas.Text = 0 'Inicializa
   
   fraCuotas.Visible = True
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus
   
 Case 1 'Extraordinario
   txtCuotas = 0
   
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para Ab.Extraordinario:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   lblDatosInteres.Caption = 0
   lblDatosAnticipo.Caption = 0
   lblDatosCargos.Caption = 0
   lblPolizas.Caption = 0
   
  
   txtDatosAmortiza.Text = 0
   txtTotalCancela.Text = 0
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus
   
   chkRecalculaCuota.Enabled = True
   strSQL = "select dbo.fxCrdPlanPagosDiasActivoFecha( " & txtOperacion.Text & ", '" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "') as 'Dias'"
   Call OpenRecordSet(rs, strSQL)
     vDiasActivo = rs!Dias
   rs.Close
   
Case 2 'Cancelación
   
   txtDatosAmortiza.Text = 0
  
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para cancelación...:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   strSQL = "exec spCrdPlanPagosInfoCancelacion " & txtOperacion.Text & ", '" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
   Call OpenRecordSet(rs, strSQL)
    txtDatosAmortiza.Text = Format(rs!Principal, "Standard")
    lblDatosInteres.Caption = Format(rs!IntCor + rs!IntMor, "Standard")
    lblDatosCargos.Caption = Format(rs!Cargos, "Standard")
    lblDatosAnticipo.Caption = Format(rs!CargoAnticipo, "Standard")
    lblPolizas.Caption = Format(rs!Poliza, "Standard")
    txtTotalPagar.Text = Format(rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!CargoAnticipo + rs!Poliza, "Standard")
    txtTotalCancela.Text = txtTotalPagar.Text
   rs.Close
   
   If vRetencion Then
      lblDatosAnticipo.Caption = "0.00"
   End If
   


 Case 3 'Adelantos
   lblDatosAnticipo.Caption = 0
   lblDatosCargos.Caption = 0
   lblDatosInteres.Caption = 0
   lblPolizas.Caption = 0
   txtDatosAmortiza.Text = 0
   
   txtCuotas.Enabled = True
   FlatScrollBar.Enabled = txtCuotas.Enabled
   
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
   txtCuotas.SetFocus
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus

End Select

Call RefrescaTags(Me)


Me.MousePointer = vbDefault


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtNombre = fxNombre(txtCedula)
  If txtCodigo <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtCodigo.SetFocus
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtCodigo = UCase(txtCodigo)
  lblDescripcion.Caption = fxDescribeCodigo(txtCodigo)
  If txtCedula <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtOperacion.SetFocus
End If

End Sub

Private Sub sbCuotaChangeAnterior()
Dim curSaldo As Currency, curAmortiza As Currency, curInteres As Currency
Dim curTmpAmortiza As Currency, curTmpInteres As Currency, i As Integer
Dim lngFecha As Currency, lngCuotas As Long, lngCuotaMaxima As Long


Dim iDias As Integer, vFecha As Date, curCuota As Currency, iPlazoRst As Integer, ProcesosTmp As Currency

On Error Resume Next

If txtCuotas = "" Then
 lngCuotas = 0
Else
 lngCuotas = txtCuotas
End If

lngFecha = CLng(lblFecUltMov.Caption)

If Not vRetencion Then
    curSaldo = vSaldoMes
Else
  'En las retenciones hay que proyectar el saldo del mes
  curSaldo = ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption))
End If

curAmortiza = 0
curInteres = 0
curCuota = lblCuota.Caption


If lngFecha < vPrideduc Then lngFecha = vPrideduc

If vBaseCalculo = "01" Then
    For i = 1 To lngCuotas
    '360 / 360
        If curSaldo > 0 Then
          lngCuotaMaxima = i
          curTmpInteres = (curSaldo * vInteres) / 1200
          curTmpAmortiza = CCur(lblCuota.Caption) - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
        
        End If
        
        If curSaldo < 0 Then
           curAmortiza = curAmortiza + curSaldo
           curSaldo = 0
        End If
     Next i
 
 Else
   '365 / 360
   
       'Calcula el Plazo Restante
       ProcesosTmp = vPrideduc
       iPlazoRst = 0
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       
       'Saca el formato fecha del ultimo movimiento para calculo de dias
       vFecha = Mid(CStr(lngFecha), 1, 4) & "/" & Right(CStr(lngFecha), 2) & "/01"
       
       For i = 1 To lngCuotas
          lngCuotaMaxima = i

          If iPlazoRst = 1 Or iPlazoRst = vPlazo Then
            iDias = 30
          Else
            iDias = fxMesDias(Month(vFecha), Year(vFecha))
          End If
        
          curTmpInteres = curSaldo * (vInteres / 100) * iDias / 360
          curTmpAmortiza = curCuota - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
          vFecha = DateAdd("m", 1, vFecha)
          
          iPlazoRst = iPlazoRst - 1
          curCuota = fxCalcula_Cuota(CDbl(curSaldo), iPlazoRst, vInteres)
          
       Next i
    
   
 End If 'Base

lblDatosInteres.Caption = Format(curInteres, "Standard")
txtDatosAmortiza = Format(curAmortiza, "Standard")
lblFecUltMovR.Caption = lngFecha
lblCuotaR.Caption = Format(curCuota, "Standard")

If Not vRetencion Then 'El proceso nuevo de retenciones no toca los saldos
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard")
End If

lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + curAmortiza, "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + curInteres, "Standard")

If lngCuotas > lngCuotaMaxima Then txtCuotas = lngCuotaMaxima


End Sub

Private Sub txtCuotas_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngCuotas As Long

If vOperacion = 0 Then Exit Sub

On Error GoTo vError

If Not IsNumeric(txtCuotas.Text) Then
 lngCuotas = 1
Else
 lngCuotas = txtCuotas.Text
End If

If lngCuotas <= 0 Then lngCuotas = 1

strSQL = "select isnull(max(id_Seq),0) as 'SeqX', isnull(sum(IntCor + IntMor),0) as 'IntCor', isnull(sum(Principal),0) as 'Principal'" _
       & ",isnull(min(Saldo_Actual),0) as 'Saldo', isnull(max(Fecha_Proceso),0) as 'Fecha_Proceso', isnull(sum(Poliza),0) as 'Poliza'" _
       & " from CRD_OPERACION_PLAN_PAGOS where id_solicitud = " & vOperacion _
       & " and Id_Seq in(select Top " & lngCuotas & " Id_Seq from CRD_OPERACION_PLAN_PAGOS" _
       & " where estado in('A','P') and id_solicitud = " & vOperacion & " and num_cuota > 0  order by num_cuota)"
Call OpenRecordSet(rs, strSQL)
    lblDatosInteres.Caption = Format(rs!IntCor, "Standard")
    lblPolizas.Caption = Format(rs!Poliza, "Standard")
    lblFecUltMovR.Caption = rs!Fecha_Proceso
    
    If Not vRetencion Then 'El proceso nuevo de retenciones no toca los saldos
        lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - rs!Principal, "Standard")
    End If
    
    
    lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + rs!Principal, "Standard")
    lblInteresR.Caption = Format(CCur(lblInteres.Caption) + rs!IntCor, "Standard")

    'Se pone de ultimo porque activa otro sub
    txtDatosAmortiza.Text = Format(rs!Principal, "Standard")

strSQL = "select cuota from CRD_OPERACION_PLAN_PAGOS where id_seq = " & rs!SeqX & " and id_solicitud = " & vOperacion
rs.Close

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    lblCuotaR.Caption = Format(rs!Cuota, "Standard")
End If
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub txtDatosAmortiza_Change()
On Error Resume Next

If Not vRetencion Then
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - CCur(txtDatosAmortiza), "Standard")
Else
    lblSaldoR.Caption = lblCuota.Caption
End If
lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + CCur(txtDatosAmortiza), "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + CCur(lblDatosInteres), "Standard")


txtTotalPagar.Text = Format(CCur(txtDatosAmortiza) + CCur(lblDatosInteres.Caption) + CCur(lblPolizas.Caption) _
                + CCur(lblDatosAnticipo.Caption) + CCur(lblDatosCargos.Caption), "Standard")
txtTotalCancela.Text = txtTotalPagar.Text
txtDiferencia.Text = "0.00"


End Sub

Private Sub txtDatosAmortiza_GotFocus()
On Error Resume Next
txtDatosAmortiza = CCur(txtDatosAmortiza)
End Sub

Private Sub txtDatosAmortiza_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
 cboTipoPago.SetFocus
End If
End Sub

Private Sub txtDatosAmortiza_LostFocus()
On Error Resume Next
txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyPress(vbKeyReturn)
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 vOperacion = txtOperacion
 Call sbConsultaOperacion
End If
End Sub


Private Function fxDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pComprobante As String _
                                , pConcepto As String, pCuenta As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim lngRecibo As Long, strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency

vCuenta = pCuenta

lngRecibo = CLng(pComprobante)

fxDocumentoAbono = lngRecibo


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)


strSQL = "exec spCrdDocumentoAfectacion '" & fxTipoASENumero(pTipoDoc) & "','" & pComprobante & "','R'"
Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp.EOF And rsTmp.BOF Then
  curIntC = 0
  curIntM = 0
  curAmortiza = 0
  curCargo = 0
  curPoliza = 0
Else
  curIntC = rsTmp!IntCor
  curIntM = rsTmp!IntMor
  curAmortiza = rsTmp!Principal
  curCargo = rsTmp!Cargos
  curPoliza = rsTmp!Polizas
End If
rsTmp.Close


'Lineas de Comprobante
strLinea(1) = "Saldo Anterior    " & lblSaldo.Caption
strLinea(2) = "Interes Corriente " & Format(curIntC, "Standard")
strLinea(3) = "Interes Atrasado  " & Format(curIntM, "Standard")
strLinea(4) = "Amortización      " & Format(curAmortiza, "Standard")
strLinea(5) = "Cargos            " & Format(curCargo, "Standard")
strLinea(6) = "Saldo Actual      " & Format(IIf(vRetencion, lblSaldo.Caption, CCur(lblSaldo.Caption) - curAmortiza), "Standard")
strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text & " Ret.:" & IIf(vRetencion, "SI", "NO")
If cboDiferenciaApl.Enabled Then
    strLinea(8) = "Aplica Dif...: " & cboDiferenciaApl.Text
Else
    strLinea(8) = Trim(lblDescripcion.Caption)
End If
strLinea(11) = "Póliza            " & Format(curPoliza, "Standard")

strSQL = "exec spCrdOperacionFechaProxPago " & txtOperacion.Text
Call OpenRecordSet(rsTmp, strSQL, 0)
  If Not IsNull(rsTmp!Fecha_Pago) Then
       strLinea(9) = "Prox.Pago..:" & Format(rsTmp!Fecha_Pago, "dd/mm/yyyy") & " Cta.(" & rsTmp!num_cuota & ") " & Format(rsTmp!Cuota, "Standard")
  Else
       strLinea(9) = "Prox.Pago..: >> <<"
  End If
  strLinea(10) = "Notas: " & rsTmp!notas & ""
rsTmp.Close
      

If dtpFechaCancelacion.Enabled Then
  strLinea(7) = strLinea(7) & "FDoc." & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
End If
strLinea(7) = Mid(strLinea(7), 1, 80)

If pTipoDoc = "RE" Then
   vAseDocDeposito = txtDocumentoExterno.Text
End If

'Registro del Comprobante
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
         & " values('" & lngRecibo & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo + curPoliza & ",'P','" & txtOperacion.Text _
         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
         & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
 Call ConectionExecute(strSQL)
 
 'ASIENTO
 If curIntC > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
          & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curIntM > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntM & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
          & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curCargo > 0 Then
 'Detallar Cargos
   strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & lngRecibo & "'"
   Call OpenRecordSet(rsTmp, strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & rsTmp!Mov_Monto & ",'C','" & rs!cod_Divisa _
                & "',1," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
                & "','" & rsTmp!id_solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
         Call ConectionExecute(strSQL)
         rsTmp.MoveNext
   Loop
   rsTmp.Close
 End If
 
 If curPoliza > 0 Then
  
 'Detallar Poliza
   strSQL = "exec spCrdDocumentoAfectacionPolizas '" & pTipoDoc & "','" & lngRecibo & "'"
   Call OpenRecordSet(rsTmp, strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & rsTmp!Mov_Monto & ",'C','" & rs!cod_Divisa _
                & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rsTmp!cod_cuenta _
                & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
         Call ConectionExecute(strSQL)
         rsTmp.MoveNext
   Loop
   rsTmp.Close
   
 End If
 
 If curAmortiza > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza & ",'C','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
          & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
 
 If curIntC + curIntM + curAmortiza + curCargo + curPoliza > 0 Then
   strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curIntC + curIntM + curCargo + curAmortiza + curPoliza & ",'D','" & rs!cod_Divisa _
          & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!cod_centro_costo & "','" & vCuenta _
          & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
   Call ConectionExecute(strSQL)
 End If
  

rs.Close


End Function


Private Sub txtTotalPagar_GotFocus()

On Error GoTo vError
 txtTotalPagar.Text = CCur(txtTotalPagar.Text)
vError:

End Sub

Private Sub txtTotalPagar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If CmdAbono.Enabled Then
    CmdAbono.SetFocus
  Else
    cmdReporte.SetFocus
  End If
End If

End Sub

Private Sub txtTotalPagar_LostFocus()
Dim vFecha As Date, vProceso As Long
Dim curInteres As Currency, curAmortiza As Currency, curAnticipo As Currency
Dim i As Integer, vChecks As Boolean, iPlazo As Integer
 
On Error GoTo vError
 
'ExtraOrdinario
If optAbono.Item(1).Value = True Then
   'Cobra intereses desde el ultimo corte
    txtTotalCancela.Text = Format(txtTotalPagar.Text, "Standard")
    curInteres = (CCur(txtTotalPagar.Text) * vInteres / 36000) * vDiasActivo
    curAnticipo = CCur(txtTotalPagar.Text) * vAnticipoPorc
   'Se re-calculan intereses para ajustar y relacionar segun porcion amortizada
   'Previamente sobre el monto a cancelar
   
   If curInteres + curAnticipo > 0 Then
      'Hacer 10 aproximaciones
      For i = 1 To 10
            curAmortiza = CCur(txtTotalPagar.Text) - (curInteres + curAnticipo)
            curInteres = (curAmortiza * vInteres / 36000) * vDiasActivo
      Next i
   End If
   
   lblDatosInteres.Caption = Format(curInteres, "Standard")
   lblDatosAnticipo.Caption = Format(curAnticipo, "Standard")
   txtDatosAmortiza.Text = Format(CCur(txtTotalPagar.Text) - (curInteres + curAnticipo), "Standard")
End If


txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text), "Standard")

txtDiferencia.Text = Format(CCur(txtTotalCancela.Text) - CCur(txtTotalPagar.Text), "Standard")

cboDiferenciaApl.Enabled = False

If CCur(txtDiferencia.Text) < 0 Then
    Select Case True
      Case optAbono.Item(0).Value 'Abono Ordinario
           cboDiferenciaApl.Enabled = False
           'Verifica el Plazo sea menor que la ultima cuota marcada y que se hayan marcado todas con corte igual o menor a la fecha actual
           vChecks = True
           For i = 1 To lsw.ListItems.Count
             If Not lsw.ListItems.Item(i).Checked And DateDiff("d", CDate(lsw.ListItems.Item(i).SubItems(11)), vFechaHoy) >= 0 Then
               vChecks = False
             End If
           Next i
           
           'Verifica el Ultimo Plazo
           If vChecks Then
              If CCur(lblSaldoR.Caption) > CCur(txtDiferencia.Text) Then
                  cboDiferenciaApl.Enabled = True
              End If
           End If
           
      Case optAbono.Item(1).Value 'Abono ExtraOrdinario
           cboDiferenciaApl.Enabled = False
      
      Case optAbono.Item(2).Value 'Cancelacion
           cboDiferenciaApl.Enabled = False
      
      Case optAbono.Item(3).Value 'Adelanto
           'Verifica si el Saldo Resultante del Credito es mayor igual a la diferencia.
              If CCur(lblSaldoR.Caption) > CCur(txtDiferencia.Text) Then
                  cboDiferenciaApl.Enabled = True
              End If
    
    End Select
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


