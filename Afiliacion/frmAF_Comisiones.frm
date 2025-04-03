VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmAF_Comisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones de Afiliación y Renuncias"
   ClientHeight    =   7764
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10428
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7764
   ScaleWidth      =   10428
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Comisiones.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Comisiones.frx":169C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Comisiones.frx":2D384
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Comisiones.frx":424F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   7605
      Width           =   10425
      _ExtentX        =   18394
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10215
      _ExtentX        =   18013
      _ExtentY        =   11028
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Remesa"
      TabPicture(0)   =   "frmAF_Comisiones.frx":57668
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label2(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2(5)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "tlb"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lswRemesas"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtRemesa"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtUsuario"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtFecha"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtEstado"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTotal"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Generación"
      TabPicture(1)   =   "frmAF_Comisiones.frx":57684
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(6)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(7)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "tlbGenera"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lswGenera"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cboGenera"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkGenera"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Pago"
      TabPicture(2)   =   "frmAF_Comisiones.frx":576A0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(10)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(8)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Line1(7)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label2(9)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line1(8)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lswPago"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cboPago"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "tlbPago"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cboBanco"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtPagoTotal"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "chkPago"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Reportes"
      TabPicture(3)   =   "frmAF_Comisiones.frx":576BC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2(13)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label2(15)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label2(14)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label2(12)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label2(11)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label1(4)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Label1(3)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Label1(1)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Line3"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "dtpRepCorte"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "dtpRepInicio"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Frame1"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "tlbReporte"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "chkRepFechas"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "cboRepTipo"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "txtRepRemesa"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "chkRepRemesas"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "txtRepPromotor"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "chkRepPromotor"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "txtRepBanco"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "chkRepBancos"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "txtRepUsuario"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "chkRepUsuario"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "cboRepBase"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).ControlCount=   24
      Begin VB.ComboBox cboRepBase 
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1560
         Width           =   2652
      End
      Begin VB.CheckBox chkRepUsuario 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
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
         Left            =   -69840
         TabIndex        =   48
         Top             =   4920
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtRepUsuario 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   47
         Text            =   "(Presione F4 para Consultar)"
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   4920
         Width           =   3735
      End
      Begin VB.CheckBox chkRepBancos 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
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
         Left            =   -69840
         TabIndex        =   45
         Top             =   4080
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtRepBanco 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "(Presione F4 para Consultar)"
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   4080
         Width           =   3735
      End
      Begin VB.CheckBox chkRepPromotor 
         Appearance      =   0  'Flat
         Caption         =   "Todos"
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
         Left            =   -69840
         TabIndex        =   42
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtRepPromotor 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "(Presione F4 para Consultar)"
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   3240
         Width           =   3735
      End
      Begin VB.CheckBox chkRepRemesas 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
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
         Left            =   -69840
         TabIndex        =   39
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtRepRemesa 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "(Presione F4 para Consultar)"
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   2400
         Width           =   3735
      End
      Begin VB.ComboBox cboRepTipo 
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   840
         Width           =   2652
      End
      Begin VB.CheckBox chkRepFechas 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
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
         Left            =   -70920
         TabIndex        =   33
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkPago 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   31
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPagoTotal 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   5760
         Width           =   2535
      End
      Begin VB.ComboBox cboBanco 
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
         Height          =   300
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   840
         Width           =   6975
      End
      Begin MSComctlLib.Toolbar tlbPago 
         Height          =   330
         Left            =   -71880
         TabIndex        =   26
         Top             =   1320
         Width           =   2625
         _ExtentX        =   4636
         _ExtentY        =   550
         ButtonWidth     =   1736
         ButtonHeight    =   550
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "buscar"
               Object.ToolTipText     =   "Buscar Casos para Pago"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Pago"
               Key             =   "pago"
               Object.ToolTipText     =   "Crear Desembolso (Pago de Comision)"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboPago 
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
         Height          =   300
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   480
         Width           =   6975
      End
      Begin VB.CheckBox chkGenera 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   20
         Top             =   1335
         Width           =   1455
      End
      Begin VB.ComboBox cboGenera 
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
         Height          =   300
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   480
         Width           =   6975
      End
      Begin VB.TextBox txtTotal 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox txtEstado 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtFecha 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtRemesa 
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
         Left            =   3120
         TabIndex        =   10
         Top             =   480
         Width           =   2655
      End
      Begin MSComctlLib.ListView lswRemesas 
         Height          =   3255
         Left            =   3120
         TabIndex        =   3
         Top             =   2880
         Width           =   6975
         _ExtentX        =   12298
         _ExtentY        =   5736
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Remesa"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   5715
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   6000
         TabIndex        =   15
         Top             =   960
         Width           =   1908
         _ExtentX        =   3366
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswGenera 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   18
         Top             =   1560
         Width           =   9975
         _ExtentX        =   17590
         _ExtentY        =   8065
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ejecutivo"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Casos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Comisión"
            Object.Width           =   3246
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbGenera 
         Height          =   330
         Left            =   -71880
         TabIndex        =   21
         Top             =   840
         Width           =   2985
         _ExtentX        =   5271
         _ExtentY        =   550
         ButtonWidth     =   1926
         ButtonHeight    =   550
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "buscar"
               Object.ToolTipText     =   "Buscar Casos para Comision"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Promotores"
                     Text            =   "Mostrar Solo Promotores"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "mnuSep1"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Comites"
                     Text            =   "Mostrar Solo Comités"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Sep2"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Externos"
                     Text            =   "Mostrar Solo Externos"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Generar"
               Key             =   "genera"
               Object.ToolTipText     =   "Generar Comisión"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswPago 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   24
         Top             =   2040
         Width           =   9975
         _ExtentX        =   17590
         _ExtentY        =   6371
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Promotor / Comité"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Tipo"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cuenta"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Pagar a Nombre de  ?"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbReporte 
         Height          =   312
         Left            =   -66240
         TabIndex        =   51
         Top             =   5760
         Width           =   1332
         _ExtentX        =   2350
         _ExtentY        =   550
         ButtonWidth     =   1693
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reporte"
               Key             =   "Reportes"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   -68640
         TabIndex        =   52
         Top             =   2760
         Width           =   3735
         Begin VB.CheckBox chkRepSinComision 
            Appearance      =   0  'Flat
            Caption         =   "Mostar Casos sin Comisión"
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
            Left            =   720
            TabIndex        =   59
            Top             =   2400
            Value           =   1  'Checked
            Width           =   2775
         End
         Begin VB.OptionButton optReportes 
            Appearance      =   0  'Flat
            Caption         =   "Comprobante de Tesorería"
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
            Index           =   4
            Left            =   120
            TabIndex        =   58
            Top             =   2040
            Width           =   3495
         End
         Begin VB.OptionButton optReportes 
            Appearance      =   0  'Flat
            Caption         =   "Agrupado x Bancos"
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
            Index           =   3
            Left            =   120
            TabIndex        =   57
            Top             =   1680
            Width           =   3495
         End
         Begin VB.OptionButton optReportes 
            Appearance      =   0  'Flat
            Caption         =   "Agrupado x Usuarios"
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
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   1320
            Width           =   3495
         End
         Begin VB.OptionButton optReportes 
            Appearance      =   0  'Flat
            Caption         =   "Agrupado x Promotor"
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
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   3495
         End
         Begin VB.OptionButton optReportes 
            Appearance      =   0  'Flat
            Caption         =   "Listado General"
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
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Value           =   -1  'True
            Width           =   3495
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            Index           =   14
            X1              =   120
            X2              =   3600
            Y1              =   480
            Y2              =   480
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Reportes Disponibles"
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
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3495
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRepInicio 
         Height          =   312
         Left            =   -73680
         TabIndex        =   60
         Top             =   1200
         Width           =   1332
         _Version        =   1245187
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
      Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
         Height          =   312
         Left            =   -72360
         TabIndex        =   61
         Top             =   1200
         Width           =   1332
         _Version        =   1245187
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
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74880
         X2              =   -64920
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label1 
         Caption         =   "Base"
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
         Left            =   -74760
         TabIndex        =   49
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Reporte"
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
         Left            =   -74760
         TabIndex        =   36
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas"
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
         Left            =   -74760
         TabIndex        =   35
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Total...:"
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
         Left            =   -68280
         TabIndex        =   29
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   -74880
         X2              =   -72000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione los promotores o comités para Pago por Tesorería"
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
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   25
         Top             =   1800
         Width           =   9975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione los promotores o comités a Generar Comisión"
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
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   19
         Top             =   1320
         Width           =   9975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   9960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   120
         X2              =   3000
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total"
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
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   120
         X2              =   3000
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   120
         X2              =   3000
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
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
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   3000
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario"
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
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   3000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   3000
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lista de Remesas"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
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
         Index           =   6
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
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
         Index           =   8
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Banco"
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
         Index           =   10
         Left            =   -74880
         TabIndex        =   27
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Reporte de Comisiones de Afiliación"
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
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   32
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remesa"
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
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   37
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Promotor / Comité"
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
         Height          =   255
         Index           =   14
         Left            =   -74760
         TabIndex        =   40
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Banco"
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
         Height          =   255
         Index           =   15
         Left            =   -74760
         TabIndex        =   43
         Top             =   3720
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Usuario"
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
         Height          =   255
         Index           =   13
         Left            =   -74760
         TabIndex        =   46
         Top             =   4560
         Width           =   3615
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comisiones de Afiliación y Renuncias Revocadas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   0
      Left            =   1880
      TabIndex        =   2
      Top             =   360
      Width           =   9015
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_Comisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cboBanco_Click()
  lswPago.ListItems.Clear
End Sub

Private Sub cboGenera_Click()
 Call lswGenera.ListItems.Clear
End Sub


Private Sub cboPago_Click()

lswPago.ListItems.Clear

If vPaso Then Exit Sub
If cboPago.ListCount <= 0 Then Exit Sub

vPaso = True
cboBanco.Clear

'Seleccionar Tes_Bancos

strSQL = "select B.id_banco,B.descripcion" _
       & " from Afi_Comision_Pago C inner join Promotores P on C.id_promotor = P.id_promotor" _
       & " inner join Tes_Bancos B on P.cod_Banco = B.id_banco" _
       & " Where C.cod_comision = " & cboPago.ItemData(cboPago.ListIndex) & " And traslado_fecha Is Null" _
       & " group by B.id_banco,B.descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cboBanco.AddItem (Format(rs!id_Banco, "0000") & "..." & Trim(rs!Descripcion))
  cboBanco.ItemData(cboBanco.NewIndex) = rs!id_Banco
  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboBanco.Text = (Format(rs!id_Banco, "0000") & "..." & Trim(rs!Descripcion))
End If
rs.Close

vPaso = False
Call cboBanco_Click

End Sub

Private Sub chkGenera_Click()
Dim i As Integer

For i = 1 To lswGenera.ListItems.Count
  lswGenera.ListItems.Item(i).Checked = chkGenera.Value
Next i

End Sub

Private Sub chkPago_Click()
Dim i As Integer, curTotal As Currency


curTotal = 0

For i = 1 To lswPago.ListItems.Count
  lswPago.ListItems.Item(i).Checked = chkPago.Value
  
   If chkPago.Value = vbChecked Then
       curTotal = curTotal + CCur(lswPago.ListItems.Item(i).SubItems(2))
   End If
  
Next i

txtPagoTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub chkRepBancos_Click()
 If chkRepBancos.Value = vbChecked Then
    txtRepBanco.Enabled = False
 Else
    txtRepBanco.Enabled = True
 End If
End Sub

Private Sub chkRepFechas_Click()

If chkRepFechas.Value = vbChecked Then
   dtpRepInicio.Enabled = False
Else
   dtpRepInicio.Enabled = True
End If

dtpRepCorte.Enabled = dtpRepInicio.Enabled
cboRepBase.Enabled = dtpRepInicio.Enabled
End Sub

Private Sub chkRepPromotor_Click()
 If chkRepPromotor.Value = vbChecked Then
    txtRepPromotor.Enabled = False
 Else
    txtRepPromotor.Enabled = True
 End If
End Sub

Private Sub chkRepRemesas_Click()
 If chkRepRemesas.Value = vbChecked Then
    txtRepRemesa.Enabled = False
 Else
    txtRepRemesa.Enabled = True
 End If
End Sub

Private Sub chkRepUsuario_Click()
 If chkRepUsuario.Value = vbChecked Then
    txtRepUsuario.Enabled = False
 Else
    txtRepUsuario.Enabled = True
 End If
End Sub

Private Sub Form_Activate()
 vModulo = 1
End Sub

Private Sub Form_Load()
 
vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

 ssTab.Tab = 0
 Call sbToolBarIconos(tlb, False)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
 
End Sub


Private Sub sbConsulta(vComision As Long)

Call sbLimpia
  
strSQL = "select * from afi_comisiones where cod_comision = " & vComision
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = rs!Cod_Comision
  txtUsuario = rs!Usuario
  txtFecha = rs!fecha
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "G"
      txtEstado = "Remesa en Generación"
    Case "P"
      txtEstado = "Remesa Pagada"
    Case "C"
      txtEstado = "Remesa en Cola de Pago"
  End Select
  
  With glogon
    .strSQL = "select isnull(sum(monto),0) as Total from afi_comision_pago where cod_comision = " & rs!Cod_Comision
    .Recordset.Open .strSQL, .Conection, adOpenStatic
    txtTotal.Text = Format(.Recordset!Total, "Standard")
    .Recordset.Close
  End With
  
End If
rs.Close


End Sub




Private Sub lswGenera_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo vError
    
    lswGenera.SortKey = ColumnHeader.Index - 1
    
    If (lswGenera.SortOrder = lvwAscending) Then
        lswGenera.SortOrder = lvwDescending
    Else
        lswGenera.SortOrder = lvwAscending
    End If
    
    lswGenera.Sorted = True
    Exit Sub

vError:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub

Private Sub lswPago_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim curTotal As Currency

If Trim(txtPagoTotal.Text) = "" Then txtPagoTotal.Text = 0

curTotal = CCur(txtPagoTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(2))
Else
   curTotal = curTotal - CCur(Item.SubItems(2))
End If

txtPagoTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub lswPago_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo vError
    
    lswPago.SortKey = ColumnHeader.Index - 1
    
    If (lswPago.SortOrder = lvwAscending) Then
        lswPago.SortOrder = lvwDescending
    Else
        lswPago.SortOrder = lvwAscending
    End If
    
    lswPago.Sorted = True
    Exit Sub

vError:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub
Private Sub lswRemesas_Click()
If lswRemesas.ListItems.Count <= 0 Then Exit Sub
Call sbConsulta(lswRemesas.SelectedItem)
End Sub


Private Sub optReportes_Click(Index As Integer)
If Index = 4 Then
   chkRepSinComision.Value = vbUnchecked
   chkRepSinComision.Enabled = False
Else
   chkRepSinComision.Value = vbChecked
   chkRepSinComision.Enabled = True
End If

End Sub

Private Sub SSTab_Click(PreviousTab As Integer)
 Call sbLimpia
End Sub

Private Sub sbReporte()
Dim vSubTitulo As String, vFiltro As String
Dim strSQL As String

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
 .WindowTitle = "Reportes del Sub.Módulo de Comisiones de Afiliación"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionRemesas.rpt")
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
     
    strSQL = "select isnull(max(cod_comision),0) + 1 as Ultimo from afi_comisiones"
    Call OpenRecordSet(rs, strSQL)
     
    strSQL = "insert afi_comisiones(cod_comision,usuario,fecha,estado) values(" & rs!ultimo _
           & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A')"
    Call ConectionExecute(strSQL)
    
    txtRemesa = rs!ultimo
    rs.Close
    
    Call sbConsulta(txtRemesa)
    
    Call Bitacora("Registra", "Remesa Pago de Comision de Afiliacion : " & txtRemesa)
    
    Call sbLimpia
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "delete afi_comisiones where cod_comision = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Remesa : Pago de Comisiones Afiliacion : " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case "REPORTES"
     Call sbReporte

  Case "AYUDA"
'        frmContenedor.CD.HelpContext = Me.HelpContextID
'        frmContenedor.CD.ShowHelp

End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxParametrosComision(vParametro) As String
Dim rsX As New ADODB.Recordset

On Error GoTo vError

With glogon
 .strSQL = "select valor from afi_comisiones_parametros where cod_parametro = '" & vParametro & "'"
 rsX.Open .strSQL, .Conection, adOpenStatic
   fxParametrosComision = rsX!Valor
 rsX.Close

End With

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub sbLimpia()

Select Case ssTab.Tab
  Case 0 'Remesas
     txtEstado = ""
     txtFecha = ""
     txtTotal = 0
     txtUsuario = ""
     txtRemesa = ""
     
     
     strSQL = "select TOP 50 * from afi_comisiones order by fecha desc"
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!Cod_Comision)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Generacion
    'Solo busca las remesas que se encuentran abiertas o en generacion (A ó G)
    
    cboGenera.Clear
    lswGenera.ListItems.Clear
    chkGenera.Value = vbUnchecked
        
    strSQL = "select * from afi_comisiones where estado in('A','G') order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboGenera.AddItem (Format(rs!Cod_Comision, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!fecha)
      cboGenera.ItemData(cboGenera.NewIndex) = rs!Cod_Comision
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboGenera.Text = (Format(rs!Cod_Comision, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!fecha)
    End If
    
    rs.Close
    
    
  Case 2 'Pago
    'Solo busca las remesas que se encuentran en cola de Pago o Pagadas (C,G)
    
    vPaso = True
    
    cboPago.Clear
    cboBanco.Clear
    lswPago.ListItems.Clear
    chkPago.Value = vbUnchecked
        
    strSQL = "select * from afi_comisiones where estado in('G','C') order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboPago.AddItem (Format(rs!Cod_Comision, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!fecha)
      cboPago.ItemData(cboPago.NewIndex) = rs!Cod_Comision
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboPago.Text = (Format(rs!Cod_Comision, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!fecha)
    End If
    
    rs.Close

    vPaso = False
    Call cboPago_Click

  Case 3 'Reportes
    dtpRepInicio.Value = fxFechaServidor
    dtpRepCorte.Value = dtpRepInicio.Value
    chkRepFechas.Value = vbChecked
    
    cboRepBase.Clear
    cboRepBase.AddItem "Remesa"
    cboRepBase.AddItem "Pago"
    cboRepBase.Text = "Remesa"
    
    cboRepTipo.Clear
    cboRepTipo.AddItem "Detalle"
    cboRepTipo.AddItem "Resumen"
    cboRepTipo.Text = "Detalle"
    
    Call chkRepBancos_Click
    Call chkRepFechas_Click
    Call chkRepPromotor_Click
    Call chkRepRemesas_Click
    Call chkRepUsuario_Click
 End Select

End Sub


Private Sub sbGeneraBuscar(Optional pTipo As String = "P")

Me.MousePointer = vbHourglass

On Error GoTo vError

lswGenera.ListItems.Clear

strSQL = "exec spAFIComisionesConsulta '" & pTipo & "'"
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
prgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswGenera.ListItems.Add(, , rs!ID_PROMOTOR)
 Select Case rs!Tipo
    Case "A"
         itmX.SubItems(1) = "Afiliacion"
    Case "R"
          itmX.SubItems(1) = "Renuncias"
 End Select
 
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = rs!Casos
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     
     itmX.Checked = chkGenera.Value
     
 rs.MoveNext
 
 prgBar.Value = prgBar.Value + 1
 
Loop
rs.Close

prgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswGenera.ListItems.Clear

End Sub

Private Sub sbGenerar()
Dim i As Integer, vCasos As Integer

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from Afi_Comisiones" _
       & " where cod_comision = " & cboGenera.ItemData(cboGenera.ListIndex) _
       & " and estado in('A','G') "
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra en proceso de Pago...", vbExclamation
    Exit Sub
 End If
rs.Close

Me.MousePointer = vbHourglass


'Calcula los casos a procesar
vCasos = 1
For i = 1 To lswGenera.ListItems.Count
 If lswGenera.ListItems.Item(i).Checked Then
    vCasos = vCasos + 1
 End If
Next i

prgBar.Max = vCasos
prgBar.Value = 1
prgBar.Visible = True


For i = 1 To lswGenera.ListItems.Count
 If lswGenera.ListItems.Item(i).Checked Then
 
     strSQL = "exec spAFIComisionGenera " & cboGenera.ItemData(cboGenera.ListIndex) & "," _
            & lswGenera.ListItems.Item(i).Text & ",'" & Mid(lswGenera.ListItems.Item(i).SubItems(1), 1, 1) & "'"
     Call ConectionExecute(strSQL)
   
    Call Bitacora("Aplica", "Generacion Comision de Afiliacion Id." & cboGenera.ItemData(cboGenera.ListIndex) _
                    & " Prom." & lswGenera.ListItems.Item(i).Text)
     prgBar.Value = prgBar.Value + 1
    
  End If
Next i

'Actualiza el Estado de la Remesa
strSQL = "update Afi_Comisiones set estado = 'G'" _
       & " where cod_comision = " & cboGenera.ItemData(cboGenera.ListIndex)
Call ConectionExecute(strSQL)

prgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation

lswGenera.ListItems.Clear

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswGenera.ListItems.Clear

End Sub



Private Sub sbPagoBuscar()
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswPago.ListItems.Clear
curTotal = 0

strSQL = "select P.id_promotor,P.nombre,C.monto,P.tipo_documento,P.cuenta_ahorros,P.nombre_contacto" _
       & " from Afi_Comision_Pago C inner join Promotores P on C.id_promotor = P.id_promotor" _
       & " inner join Tes_Bancos B on P.cod_Banco = B.id_banco" _
       & " Where C.cod_comision = " & cboPago.ItemData(cboPago.ListIndex) _
       & " And C.traslado_fecha Is Null and B.id_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
prgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswPago.ListItems.Add(, , rs!ID_PROMOTOR)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = Trim(rs!TIPO_DOCUMENTO)
     itmX.SubItems(4) = Trim(rs!Cuenta_Ahorros)
     itmX.SubItems(5) = rs!Nombre_Contacto & ""
     
     itmX.Checked = chkPago.Value
     
     If itmX.Checked Then
        curTotal = curTotal + CCur(itmX.SubItems(2))
     End If
     
 rs.MoveNext
 
 prgBar.Value = prgBar.Value + 1
 
Loop
rs.Close

prgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswPago.ListItems.Clear

End Sub

Private Sub sbPago()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date
Dim vToken As String

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor
strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  vToken = rs!id_token
Else
  vToken = fxTesToken
End If
rs.Close


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from Afi_Comisiones" _
       & " where cod_comision = " & cboPago.ItemData(cboPago.ListIndex) _
       & " and estado in('G','C') "
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra en procesada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como Cola de Pago / Al finalizar Revisa si ya fue Totalmente Pagada
strSQL = "update Afi_Comisiones set estado = 'C'" _
       & " where cod_comision = " & cboPago.ItemData(cboPago.ListIndex)
Call ConectionExecute(strSQL)

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

'Calcula los casos a procesar y desmarca casos marcados con transferencia y sin cuenta
vCasos = 1
For i = 1 To lswPago.ListItems.Count
 If lswPago.ListItems.Item(i).Checked Then
    If Trim(lswPago.ListItems.Item(i).SubItems(3)) = "TE" Then
       If Len(Trim(lswPago.ListItems.Item(i).SubItems(4))) < 5 Then 'La cuenta no es válida
           lswPago.ListItems.Item(i).Checked = False
       Else
           vCasos = vCasos + 1
       End If
    Else
       vCasos = vCasos + 1
    End If
 End If
Next i

prgBar.Max = vCasos
prgBar.Value = 1
prgBar.Visible = True


With lswPago.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then
 
     strSQL = "exec spAFIComisionPago " & cboPago.ItemData(cboPago.ListIndex) & "," & .Item(i).Text _
            & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & vToken _
            & "'," & cboPago.ItemData(cboPago.ListIndex) & ",'AFI.COM'"
     Call ConectionExecute(strSQL)
   
   
     Call Bitacora("Aplica", "Pago Comision de Afiliacion Id." & cboPago.ItemData(cboPago.ListIndex) _
                    & " Prom." & .Item(i).Text)

    prgBar.Value = prgBar.Value + 1
  End If
Next i

End With

prgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbPagoBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswPago.ListItems.Clear

End Sub


Private Sub tlbGenera_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Key
  Case "buscar"
    If cboGenera.ListCount = 0 Then Exit Sub
    
    If Len(Button.Caption) = 7 Then
       MsgBox "Seleccione El Tipo de Busqueda y Luego Presion Buscar [x]", vbInformation
       Exit Sub
    End If
    
    Call sbGeneraBuscar(Mid(Button.Caption, 9, 1))
  
  Case "genera"
    If lswGenera.ListItems.Count = 0 Then Exit Sub
    Call sbGenerar

End Select

End Sub

Private Sub tlbGenera_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

'If lswGenera.ListItems.Count = 0 Then Exit Sub

Select Case ButtonMenu.Key
  Case "Promotores"
    tlbGenera.Buttons(1).Caption = "Buscar [Promotores]"
  
  Case "Comites"
    tlbGenera.Buttons(1).Caption = "Buscar [Comites]"
    
  Case "Externos"
    tlbGenera.Buttons(1).Caption = "Buscar [Externos]"
End Select

tlbGenera.Width = tlbGenera.ButtonWidth * tlbGenera.Buttons.Count


End Sub

Private Sub tlbPago_ButtonClick(ByVal Button As MSComctlLib.Button)

If cboPago.ListCount = 0 Then Exit Sub
If cboBanco.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "buscar"
    Call sbPagoBuscar
  
  Case "pago"
    If lswPago.ListItems.Count = 0 Then Exit Sub
    Call sbPago

End Select

End Sub

Private Sub tlbReporte_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String

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
 .WindowTitle = "Reportes del Sub.Módulo de Comisiones de Afiliación"
 
 .Connect = glogon.ConectRPT
  
 If chkRepFechas.Value = vbUnchecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    Select Case Mid(cboRepBase.Text, 1, 1)
      Case "R" 'Fecha de Creación de la Remesa
        strSQL = strSQL & "{AFI_COMISIONES.FECHA}"
        vSubTitulo = "Generadas entre " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
      Case "P" 'Fecha de Traslado a Tesoreria
        strSQL = strSQL & "{AFI_COMISION_PAGO.TRASLADO_FECHA}"
        vSubTitulo = "Pagadas entre " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
    End Select
    strSQL = strSQL & " in Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ") to date(" _
           & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
 Else
   vSubTitulo = "Historico"
 End If
 
 If chkRepRemesas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{AFI_COMISION_PAGO.COD_COMISION} = " & txtRepRemesa.Tag
   vFiltro = vFiltro & "/ REMESA : " & txtRepRemesa.Text
 Else
   vFiltro = vFiltro & "/ TODAS LAS REMESAS"
 End If
 
 
 If chkRepPromotor.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{AFI_COMISION_PAGO.ID_PROMOTOR} = " & txtRepPromotor.Tag
   vFiltro = vFiltro & "/ PROMOTOR : " & txtRepPromotor.Text
 Else
   vFiltro = vFiltro & "/ TODOS LOS PROMOTORES"
 End If
 
 If chkRepBancos.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{AFI_COMISION_PAGO.COD_BANCO} = " & txtRepBanco.Tag
   vFiltro = vFiltro & "/ BANCO : " & txtRepBanco.Text
 Else
   vFiltro = vFiltro & "/ TODOS LOS Tes_Bancos"
 End If
 
 
 If chkRepUsuario.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "

    Select Case Mid(cboRepBase.Text, 1, 1)
      Case "R" 'Fecha de Creación de la Remesa
            strSQL = strSQL & "{AFI_COMISIONES.USUARIO} = '" & txtRepUsuario.Text & "'"
            vFiltro = vFiltro & "/ USUARIO : " & txtRepUsuario.Text
      Case "P" 'Fecha de Traslado a Tesoreria
            strSQL = strSQL & "{AFI_COMISION_PAGO.TRASLADO_USER} = '" & txtRepUsuario.Text & "'"
            vFiltro = vFiltro & "/ USUARIO : " & txtRepUsuario.Text
    End Select
 
 Else
   vFiltro = vFiltro & "/ TODOS LOS USUARIOS"
 End If
 
         
If chkRepSinComision.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   
   strSQL = strSQL & "{AFI_COMISION_PAGO.MONTO} > 0"
    vFiltro = vFiltro & "/ SOLO CASOS CON MONTO > 0"
End If
 
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
 If cboRepTipo.Text = "Detalle" Then
   Select Case True
     Case optReportes.Item(0).Value 'Listado General
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionListadoGeneral.rpt")
     Case optReportes.Item(1).Value 'Agrupado x Promotor
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpPromotor.rpt")
     Case optReportes.Item(2).Value 'Agrupado x Usuario
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpUsuario.rpt")
     Case optReportes.Item(3).Value 'Agrupado x Banco
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpBanco.rpt")
     Case optReportes.Item(4).Value 'Tesoreria
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionTesoreria.rpt")
    
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{AFI_COMISION_PAGO.COD_BANCO} > 0"
        
    
    End Select
 Else
   Select Case True
     Case optReportes.Item(0).Value 'Listado General
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionListadoGeneralRsm.rpt")
     Case optReportes.Item(1).Value 'Agrupado x Promotor
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpPromotorRsm.rpt")
     Case optReportes.Item(2).Value 'Agrupado x Usuario
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpUsuarioRsm.rpt")
     Case optReportes.Item(3).Value 'Agrupado x Banco
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpBancoRsm.rpt")
     Case optReportes.Item(4).Value 'Tesoreria
         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionTesoreriaRsm.rpt")
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{AFI_COMISION_PAGO.COD_BANCO} > 0"
    
    End Select
 End If
 
 .SelectionFormula = strSQL
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If

End Sub


Private Sub txtRepBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  
  With gBusquedas
     .Resultado = ""
     .Resultado2 = ""
     .Convertir = "N"
     
     .Consulta = "select distinct B.id_banco,B.descripcion" _
               & " from Tes_Bancos B inner join afi_comision_pago C on B.id_banco = C.cod_banco"
     .Columna = "B.descripcion"
     .Orden = "B.descripcion"
     
     If chkRepRemesas.Value = vbUnchecked Then
        .Filtro = " and C.cod_comision = " & txtRepRemesa.Tag
     End If
     
     frmBusquedas.Show vbModal
     If .Resultado <> "" Then
         txtRepBanco.Text = .Resultado2
         txtRepBanco.Tag = .Resultado
     End If
  End With
End If
End Sub

Private Sub txtRepUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  
  With gBusquedas
     .Resultado = ""
     .Resultado2 = ""
     .Convertir = "N"
     
     .Consulta = "select distinct Traslado_user as Usuarios" _
               & " from afi_comision_pago"
     .Columna = "Traslado_user"
     .Orden = "Traslado_user"
     
     If chkRepRemesas.Value = vbUnchecked Then
        .Filtro = " and cod_comision = " & txtRepRemesa.Tag
     End If
     
     frmBusquedas.Show vbModal
     If .Resultado <> "" Then
         txtRepUsuario.Text = .Resultado
     End If
  End With
End If

End Sub

Private Sub txtRepRemesa_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  With gBusquedas
     .Resultado = ""
     .Resultado2 = ""
     .Convertir = "S"
     
     .Consulta = "select cod_comision,fecha,usuario from afi_comisiones"
     .Columna = "cod_comision"
     .Filtro = ""
     .Orden = "fecha desc"
     
     frmBusquedas.Show vbModal
     If .Resultado <> "" Then
         txtRepRemesa.Text = Format(.Resultado, "0000") & " - Fecha : " & .Resultado2
         txtRepRemesa.Tag = .Resultado
     End If
  End With
End If

End Sub


Private Sub txtRepPromotor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  
  With gBusquedas
     .Resultado = ""
     .Resultado2 = ""
     .Convertir = "N"
     
     .Consulta = "select P.id_promotor,P.nombre" _
               & " from promotores P inner join afi_comision_pago C on P.id_promotor = C.id_Promotor"
     .Columna = "P.nombre"
     .Orden = "P.nombre"
     
     If chkRepRemesas.Value = vbUnchecked Then
        .Filtro = " and C.cod_comision = " & txtRepRemesa.Tag
     End If
     
     frmBusquedas.Show vbModal
     If .Resultado <> "" Then
         txtRepPromotor.Text = .Resultado2
         txtRepPromotor.Tag = .Resultado
     End If
  End With
End If
End Sub
