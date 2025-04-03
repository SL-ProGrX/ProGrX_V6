VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_CuentasSGTRebajosInternos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abonos a Cuentas por Cobrar: Pendientes"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRefunde 
      Height          =   5652
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CheckBox chkCargoReposicion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Aplicar Cargo por Reposición"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6600
         TabIndex        =   42
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtCargoReposicion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1680
         Width           =   2175
      End
      Begin VB.OptionButton OptX 
         Appearance      =   0  'Flat
         Caption         =   "Mora"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   38
         Top             =   4080
         Width           =   2055
      End
      Begin VB.OptionButton OptX 
         Appearance      =   0  'Flat
         Caption         =   "Cancelación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   37
         Top             =   3720
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtAbono 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7080
         TabIndex        =   27
         Top             =   4080
         Width           =   2175
      End
      Begin VB.TextBox txtCargos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox txtIntMor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtIntCor 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtPrincipal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   1680
         Width           =   2055
      End
      Begin MSComctlLib.Toolbar tlbRefunde 
         Height          =   312
         Left            =   6840
         TabIndex        =   6
         Top             =   4800
         Width           =   2448
         _ExtentX        =   4313
         _ExtentY        =   556
         ButtonWidth     =   1640
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Abona"
               Key             =   "abona"
               Object.ToolTipText     =   "Abonar a esta Operación"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra Refundicion"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8520
         Top             =   1440
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
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":15172
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":2A2E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":40CA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":55E18
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":6AF8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":6E41C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":6EC0B
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Cargos por Reposición"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   14
         Left            =   4680
         TabIndex        =   41
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblDiasMora 
         Caption         =   "Dias [x]"
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
         Left            =   3480
         TabIndex        =   36
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lblDias 
         Caption         =   "Dias [x]"
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
         Left            =   3480
         TabIndex        =   35
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label lblContrato 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   4560
         TabIndex        =   34
         Top             =   1200
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Contrato"
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
         Index           =   13
         Left            =   3480
         TabIndex        =   33
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblConcepto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   4560
         TabIndex        =   32
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         Index           =   12
         Left            =   3480
         TabIndex        =   31
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Poner al día"
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
         Index           =   11
         Left            =   120
         TabIndex        =   30
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label lblMora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Abono .....:"
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
         Index           =   10
         Left            =   5880
         TabIndex        =   28
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Cancelación"
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
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblCancelacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1320
         TabIndex        =   25
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label lblDocumento 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "No.Doc."
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
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Disponible"
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
         Index           =   1
         Left            =   5880
         TabIndex        =   15
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   240
         X2              =   9120
         Y1              =   4680
         Y2              =   4680
      End
      Begin VB.Label lblDisponible 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7080
         TabIndex        =   14
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   240
         X2              =   3000
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Datos de la Operación"
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
         Index           =   7
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label2 
         Caption         =   "Principal"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Int.Moratorio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Int.Corriente"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Saldo"
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
         TabIndex        =   9
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblOperacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "# Operación"
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
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2412
      Left            =   240
      TabIndex        =   43
      Top             =   4440
      Width           =   9252
      _Version        =   1441793
      _ExtentX        =   16319
      _ExtentY        =   4254
      _StockProps     =   79
      Caption         =   "Movimientos Registrados (Programados)"
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
      BorderStyle     =   1
      Begin MSComctlLib.ListView lswRefunde 
         Height          =   2052
         Left            =   240
         TabIndex        =   44
         Top             =   360
         Width           =   8892
         _ExtentX        =   15690
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   14
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Concepto"
            Object.Width           =   1676
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Contrato"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Int.Cor."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Int.Mor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cargos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Principal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Dias"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Dias Mora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "No. Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Contrato Desc."
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   13
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   2772
      Left            =   120
      TabIndex        =   17
      Top             =   1440
      Width           =   9372
      _ExtentX        =   16536
      _ExtentY        =   4895
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Deudor"
      TabPicture(0)   =   "frmCxC_CuentasSGTRebajosInternos.frx":6F5C8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lswCuentas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Terceros"
      TabPicture(1)   =   "frmCxC_CuentasSGTRebajosInternos.frx":6F5E4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(26)"
      Tab(1).Control(1)=   "lswTerceros"
      Tab(1).Control(2)=   "txtConNombre"
      Tab(1).Control(3)=   "txtConCedula"
      Tab(1).ControlCount=   4
      Begin VB.TextBox txtConCedula 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   -74040
         TabIndex        =   19
         ToolTipText     =   "Digite la Cédula de la Persona"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtConNombre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Left            =   -72120
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Digite la Cédula de la Persona"
         Top             =   480
         Width           =   6375
      End
      Begin MSComctlLib.ListView lswCuentas 
         Height          =   2175
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Concepto"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Contrato"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Int.Cor."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Int.Mor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cargos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Principal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Dias"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Dias Mora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "No.Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Contrado Desc."
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ListView lswTerceros 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   39
         Top             =   840
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   13
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Concepto"
            Object.Width           =   1588
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Contrato"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Saldo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Int.Cor."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Int.Mor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Cargos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Principal"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Dias"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Dias Mora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "No.Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "Contrado Desc."
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cédula"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   26
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9000
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":6F600
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":72A92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnActualizar 
      Height          =   492
      Left            =   8160
      TabIndex        =   45
      Top             =   480
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Actualizar"
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
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCxC_CuentasSGTRebajosInternos.frx":7341F
      ImageAlignment  =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Abonos/Cancelación de Operaciones"
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
      Height          =   612
      Left            =   1680
      TabIndex        =   22
      Top             =   480
      Width           =   6012
   End
   Begin VB.Image imgBanner 
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "frmCxC_CuentasSGTRebajosInternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mMonto As Currency, mRebajosTotales As Currency, mIngresosTotales As Currency, mOperacion As Long, mCedula As String


Private Sub btnActualizar_Click()
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

'strSQL = "exec spCxC_TraCrdRefActualiza " & mOperacion
'Call ConectionExecute(strSQL)

Call Form_Load
Call LimpiaDatos(False)


Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkCargoReposicion_Click()
   If chkCargoReposicion.Value = vbUnchecked Then
    lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text), "Standard")
    lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text), "Standard")
   Else
    lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtCargoReposicion.Text), "Standard")
    lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtCargoReposicion.Text), "Standard")
   End If
   
  Call OptX_Click(0)
   
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


mOperacion = GLOBALES.gTag

ssTab.Tab = 0

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

strSQL = "Select isnull(dbo.fxCxC_CuentaRebajos(" & mOperacion & ",'TOT'),0) as 'Rebajos', Monto,cedula" _
       & ", isnull(dbo.fxCxC_CuentaIngresos(" & mOperacion & "),0) as 'Ingresos'" _
       & " from CxC_Cuentas Where Operacion = " & mOperacion
Call OpenRecordSet(rs, strSQL)
   mRebajosTotales = rs!Rebajos
   mIngresosTotales = rs!Ingresos
   mMonto = rs!Monto
   mCedula = Trim(rs!Cedula)
rs.Close

Me.Caption = "Operación : " & mOperacion
lblDisponible.Caption = Format(mMonto + mIngresosTotales - mRebajosTotales, "Standard")

Call sbCargaRebajos
Call sbCargaCuentas

End Sub

Private Sub sbCargaRebajos()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem

strSQL = "select R.*,X.cod_concepto,C.descripcion as 'ConceptoDesc',G.descripcion as 'ContratoDesc',X.num_Documento,isnull(X.cod_contrato,'') as 'ContratoCod'" _
       & " from CxC_Cuentas_Rebajos R inner join CxC_Cuentas X on R.Operacion_Aplicada = X.Operacion" _
       & " inner join CxC_Conceptos C on X.cod_concepto = C.cod_concepto" _
       & " left join CxC_Contratos G on X.cod_Contrato = G.cod_Contrato" _
       & " where R.Operacion = " & mOperacion
       
       
'       & " and C.PROCESO_DESCUENTO = 0"
Call OpenRecordSet(rs, strSQL)
With lswRefunde
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Operacion_Aplicada, , 5)
     itmX.SubItems(1) = rs!cod_Concepto
     itmX.SubItems(2) = rs!ContratoCod & ""
     itmX.SubItems(3) = rs!ConceptoDesc
     itmX.SubItems(4) = Format(rs!Saldo, "Standard")
     itmX.SubItems(5) = Format(IIf(IsNull(rs!Int_Cor), 0, rs!Int_Cor), "Standard")
     itmX.SubItems(6) = Format(IIf(IsNull(rs!Int_Mor), 0, rs!Int_Mor), "Standard")
     itmX.SubItems(7) = Format(IIf(IsNull(rs!Cargos), 0, rs!Cargos), "Standard")
     itmX.SubItems(8) = Format(IIf(IsNull(rs!Principal), 0, rs!Principal), "Standard")
     itmX.SubItems(9) = rs!Dias & ""
     itmX.SubItems(10) = rs!Dias_Mora & ""
     itmX.SubItems(11) = rs!Num_Documento & ""
     itmX.SubItems(12) = rs!ContratoDesc & ""
     itmX.SubItems(13) = Format(rs!Monto, "Standard")
     
     
   rs.MoveNext
  Loop
End With
rs.Close

End Sub

Private Sub sbCargaCuentas()
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem

strSQL = "exec spCxC_TraCuentasActivas '" & mCedula & "'"
Call OpenRecordSet(rs, strSQL, 0)
With lswCuentas
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Operacion, , 4)
        itmX.SubItems(1) = rs!cod_Concepto
        itmX.SubItems(2) = rs!COD_CONTRATO & ""
        itmX.SubItems(3) = rs!ConceptoDesc
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!Int_Cor, "Standard")
        itmX.SubItems(6) = Format(rs!Int_Mor, "Standard")
        itmX.SubItems(7) = Format(rs!Cargos, "Standard")
        itmX.SubItems(8) = Format(rs!Principal, "Standard")
        itmX.SubItems(9) = rs!Dias
        itmX.SubItems(10) = rs!Dias_Mora
        itmX.SubItems(11) = rs!Num_Documento
        itmX.SubItems(12) = rs!ContratoDesc
     rs.MoveNext
  Loop
End With
rs.Close

End Sub


Private Function fxExisteRefundicion(vOperacion As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from CxC_Cuentas_Rebajos" _
       & " where Operacion_Aplicada = " & vOperacion & " and Operacion = " & mOperacion
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
  fxExisteRefundicion = IIf((rs!Existe = 0), False, True)
rs.Close
End Function

Private Sub LimpiaDatos(Optional vVisible As Boolean = True)

lblOperacion.Caption = ""
lblDocumento.Caption = ""

lblConcepto.Caption = ""
lblConcepto.Tag = ""

lblContrato.Caption = ""
lblContrato.Tag = ""

lblDias.Caption = ""
lblDiasMora.Caption = ""
lblDias.Tag = 0
lblDiasMora.Tag = 0

txtSaldo.Text = "0"
txtPrincipal.Text = "0"
txtIntCor.Text = "0"
txtIntMor.Text = "0"

lblCancelacion.Caption = 0
lblMora.Caption = 0

txtAbono.Text = "0"

If vVisible Then
   fraRefunde.Visible = vVisible
   fraRefunde.top = 960
Else
   fraRefunde.Visible = vVisible
End If

End Sub



Private Sub lswCuentas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If lswCuentas.ListItems.Count <= 0 Then Exit Sub

With lswCuentas
   
   Call LimpiaDatos(True)
   
    lblOperacion.Caption = Trim(.SelectedItem.Text)
  
    lblDocumento.Caption = Trim(.SelectedItem.SubItems(11))
    
    lblConcepto.Caption = .SelectedItem.SubItems(3)
    lblConcepto.Tag = .SelectedItem.SubItems(1)
    
    lblContrato.Caption = .SelectedItem.SubItems(12)
    lblContrato.Tag = .SelectedItem.SubItems(2)
    
    lblDias.Caption = "Días Activo : " & .SelectedItem.SubItems(9)
    lblDiasMora.Caption = "Dias Atraso : " & .SelectedItem.SubItems(10)
    lblDias.Tag = .SelectedItem.SubItems(9)
    lblDiasMora.Tag = .SelectedItem.SubItems(10)
    
    txtSaldo.Text = .SelectedItem.SubItems(4)
    txtIntCor.Text = .SelectedItem.SubItems(5)
    txtIntMor.Text = .SelectedItem.SubItems(6)
    txtCargos.Text = .SelectedItem.SubItems(7)
    txtPrincipal.Text = .SelectedItem.SubItems(8)
    
    
   strSQL = "select dbo.fxCxC_CuentaCargoReposicion(" & Trim(.SelectedItem.Text) & ",null) as 'Cargo'"
   Call OpenRecordSet(rs, strSQL)
     txtCargoReposicion.Text = Format(rs!Cargo, "Standard")
   rs.Close
    
   If chkCargoReposicion.Value = vbUnchecked Then
    lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text), "Standard")
    lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text), "Standard")
   Else
    lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtCargoReposicion.Text), "Standard")
    lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text) + CCur(txtCargoReposicion.Text), "Standard")
   End If
    
    fraRefunde.Visible = True
    fraRefunde.Left = 120
    fraRefunde.top = 960
        
    Call OptX_Click(0)
   
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswRefunde_Click()
Dim strSQL As String

On Error GoTo vError

If lswRefunde.ListItems.Count <= 0 Then Exit Sub

With lswRefunde
 strSQL = "delete CxC_Cuentas_Rebajos where Operacion_Aplicada = " & .SelectedItem.Text _
        & " and Operacion = " & mOperacion
 Call ConectionExecute(strSQL)
 
 lblDisponible.Caption = CCur(lblDisponible.Caption) + CCur(.SelectedItem.SubItems(13))
 lblDisponible.Caption = Format(lblDisponible, "Standard")
End With

Call sbCargaRebajos

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValidaRefundicion() As Boolean
Dim vMensaje As String

fxValidaRefundicion = True
vMensaje = ""

If Len(vMensaje) > 0 Then
 fxValidaRefundicion = False
 MsgBox vMensaje, vbCritical
End If

End Function

Private Sub lswTerceros_Click()
On Error GoTo vError

If lswTerceros.ListItems.Count <= 0 Then Exit Sub


With lswTerceros
   
   Call LimpiaDatos(True)
   
    lblOperacion.Caption = Trim(.SelectedItem.Text)
  
    lblDocumento.Caption = Trim(.SelectedItem.SubItems(11))
    
    lblConcepto.Caption = .SelectedItem.SubItems(3)
    lblConcepto.Tag = .SelectedItem.SubItems(1)
    
    lblContrato.Caption = .SelectedItem.SubItems(12)
    lblContrato.Tag = .SelectedItem.SubItems(2)
    
    lblDias.Caption = "Días Activo : " & .SelectedItem.SubItems(9)
    lblDiasMora.Caption = "Dias Atraso : " & .SelectedItem.SubItems(10)
    lblDias.Tag = .SelectedItem.SubItems(9)
    lblDiasMora.Tag = .SelectedItem.SubItems(10)
    
    txtSaldo.Text = .SelectedItem.SubItems(4)
    txtIntCor.Text = .SelectedItem.SubItems(5)
    txtIntMor.Text = .SelectedItem.SubItems(6)
    txtCargos.Text = .SelectedItem.SubItems(7)
    txtPrincipal.Text = .SelectedItem.SubItems(8)
    
    lblCancelacion.Caption = Format(CCur(txtSaldo.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text), "Standard")
    lblMora.Caption = Format(CCur(txtPrincipal.Text) + CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtCargos.Text), "Standard")
    
    fraRefunde.Visible = True
    fraRefunde.Left = 120
    fraRefunde.top = 960
    Call OptX_Click(0)
   
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub OptX_Click(Index As Integer)
On Error GoTo vError

    Select Case True
      Case OptX.Item(0).Value 'Cancelacion
        If CCur(lblDisponible.Caption) >= CCur(lblCancelacion.Caption) Then
            txtAbono.Text = lblCancelacion.Caption
        Else
            txtAbono.Text = lblDisponible.Caption
        End If
      Case OptX.Item(1).Value 'Mora
        If CCur(lblDisponible.Caption) >= CCur(lblMora.Caption) Then
            txtAbono.Text = lblMora.Caption
        Else
            txtAbono.Text = lblDisponible.Caption
        End If
    End Select
    
'    txtAbono.SetFocus
    
Exit Sub

vError:

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)

Select Case ssTab.Tab
 Case 1
   Call sbCargaLswTerceros(txtConCedula)
 Case Else
End Select

End Sub


Private Sub sbRefunde()
Dim strSQL As String, curRefundir As Currency

On Error GoTo vError

If fxValidaRefundicion Then

curRefundir = CCur(txtAbono.Text)

If curRefundir > CCur(lblDisponible.Caption) Then
  MsgBox "El monto a refundir de la operación es mayor al disponible...", vbCritical
  Exit Sub
End If

If fxExisteRefundicion(lblOperacion.Caption) Then
  MsgBox "Esta Refundición Se encuentra Registrada VERIFIQUE...", vbInformation
  Exit Sub
Else
  
    If chkCargoReposicion.Value = vbChecked Then
        strSQL = "exec spCxC_CuentaCargoReposicion " & lblOperacion.Caption & ",'" & glogon.Usuario _
               & "','" & GLOBALES.gOficinaUnidad & "','" & GLOBALES.gOficinaCentroCosto & "',Null"
        Call ConectionExecute(strSQL)
        
        txtCargos.Text = CCur(txtCargos.Text) + CCur(txtCargoReposicion.Text)
    End If

  strSQL = "insert CxC_Cuentas_Rebajos(Operacion,Operacion_Aplicada,Monto,Saldo,Principal,Int_Cor,Int_Mor,Cargos,Dias,Dias_Mora) " _
         & "values(" & mOperacion & "," & lblOperacion.Caption & "," & curRefundir & "," & CCur(txtSaldo.Text) & "," & CCur(txtPrincipal.Text) _
         & "," & CCur(txtIntCor.Text) & "," & CCur(txtIntMor.Text) & "," & CCur(txtCargos.Text) & "," & lblDias.Tag _
         & "," & lblDiasMora.Tag & ")"
  Call ConectionExecute(strSQL)
  
  lblDisponible.Caption = CCur(lblDisponible.Caption) - curRefundir
  lblDisponible.Caption = Format(lblDisponible, "Standard")
  
  Call sbCargaRebajos
  Call LimpiaDatos(False)
  
End If

End If 'Verificacion de OPERACION

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlbRefunde_ButtonClick(ByVal Button As MSComctlLib.Button)
 
Select Case Button.Key
  Case "abona"
    Call sbRefunde
  Case "cerrar"
    Call LimpiaDatos(False)
End Select
 
End Sub

Private Sub sbCargaLswTerceros(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem

strSQL = "exec spCxC_TraCuentasActivas '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL, 0)
With lswTerceros
  .ListItems.Clear
  Do While Not rs.EOF
    Set itmX = .ListItems.Add(, , rs!Operacion, , 1)
        itmX.SubItems(1) = rs!cod_Concepto
        itmX.SubItems(2) = rs!ContratoCod
        itmX.SubItems(3) = rs!ConceptoDesc
        itmX.SubItems(4) = Format(rs!Saldo, "Standard")
        itmX.SubItems(5) = Format(rs!Int_Cor, "Standard")
        itmX.SubItems(6) = Format(rs!Int_Mor, "Standard")
        itmX.SubItems(7) = Format(rs!Cargos, "Standard")
        itmX.SubItems(8) = Format(rs!Principal, "Standard")
        itmX.SubItems(9) = rs!Dias
        itmX.SubItems(10) = rs!Dias_Mora
        itmX.SubItems(11) = rs!Num_Documento
        itmX.SubItems(12) = rs!ContratoDesc
     rs.MoveNext
  Loop
End With
rs.Close

End Sub


Private Sub txtAbono_GotFocus()
On Error GoTo vError
 
txtAbono.Text = CCur(txtAbono.Text)
vError:

End Sub

Private Sub txtAbono_LostFocus()
On Error GoTo vError
 
txtAbono.Text = Format(CCur(txtAbono.Text), "Standard")

vError:

End Sub

Private Sub txtConCedula_Change()
lswTerceros.ListItems.Clear
End Sub

Function fxPersonaNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select nombre from CxC_Personas where cedula = '" & strCedula & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxPersonaNombre = ""
Else
 fxPersonaNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close
End Function


Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtConNombre = fxPersonaNombre(txtConCedula)
    Call sbCargaLswTerceros(txtConCedula)
End If

End Sub
