VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmInsRegistro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Pólizas"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8670
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPoliza 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7485
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Activa"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha Activa"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario - Cierra"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Fecha Cierre"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0101
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0220
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0340
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0455
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0573
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":069D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":07C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":08DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":09DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsRegistro.frx":0AF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8670
      _ExtentX        =   15293
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   8670
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   5880
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   54
         Top             =   30
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   6075
         TabIndex        =   3
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   582
         ButtonWidth     =   1693
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Activar"
               Key             =   "Activar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "Cerrar"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6615
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Recepción"
      TabPicture(0)   =   "frmInsRegistro.frx":0C1B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(19)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraPoliza"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtNotas"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Pagos"
      TabPicture(1)   =   "frmInsRegistro.frx":747D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lswPagos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cobros"
      TabPicture(2)   =   "frmInsRegistro.frx":DCDF
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswCobros"
      Tab(2).ControlCount=   1
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
         Height          =   645
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   5820
         Width           =   7095
      End
      Begin VB.Frame fraPoliza 
         BorderStyle     =   0  'None
         Height          =   5415
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   8295
         Begin VB.TextBox txtComisionVendedor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   57
            Top             =   4680
            Width           =   2055
         End
         Begin VB.TextBox txtComisionInterna 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   56
            Top             =   5040
            Width           =   2055
         End
         Begin VB.TextBox txtEstado 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   4200
            Width           =   1455
         End
         Begin VB.TextBox txtCobroPriDeduc 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   4200
            Width           =   2055
         End
         Begin VB.TextBox txtCobroUltMov 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   3840
            Width           =   2055
         End
         Begin VB.TextBox txtCobroTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   3480
            Width           =   2055
         End
         Begin VB.TextBox txtNumCuota 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   42
            Top             =   3600
            Width           =   1455
         End
         Begin VB.TextBox txtPagoPriMov 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   2760
            Width           =   2055
         End
         Begin VB.TextBox txtPagoUltMov 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   2400
            Width           =   2055
         End
         Begin VB.TextBox txtPagoTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   38
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox txtMonto 
            Alignment       =   1  'Right Justify
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
            Left            =   960
            TabIndex        =   16
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtCuota 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   5130
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   960
            TabIndex        =   15
            Top             =   3000
            Width           =   1695
         End
         Begin VB.TextBox txtPlazo 
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
            Left            =   1920
            TabIndex        =   14
            Top             =   2640
            Width           =   735
         End
         Begin VB.TextBox txtVendedorDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   480
            Width           =   5535
         End
         Begin VB.TextBox txtVendedorCod 
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
            Left            =   1200
            TabIndex        =   12
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   11
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   120
            Width           =   5535
         End
         Begin VB.TextBox txtCedula 
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
            Left            =   960
            TabIndex        =   10
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtTipoSeguroDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   840
            Width           =   4935
         End
         Begin VB.TextBox txtTipoCuentaDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   1200
            Width           =   4935
         End
         Begin VB.TextBox txtTipoSeguroCod 
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
            Left            =   1200
            TabIndex        =   7
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtTipoCuentaCod 
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
            Left            =   1200
            TabIndex        =   6
            ToolTipText     =   "Presiones F4 para Consultar"
            Top             =   1200
            Width           =   1455
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   7440
            Top             =   0
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
                  Picture         =   "frmInsRegistro.frx":DDE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":14642
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":1AEA4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":21706
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":27F68
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":2E7CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":3502C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmInsRegistro.frx":3B88E
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoSeguro 
            Height          =   255
            Left            =   7680
            TabIndex        =   17
            Top             =   840
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1572865
         End
         Begin MSComCtl2.FlatScrollBar FlatScrollBarTipoCuenta 
            Height          =   255
            Left            =   7680
            TabIndex        =   18
            Top             =   1200
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            _Version        =   393216
            Arrows          =   65536
            Orientation     =   1572865
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00808080&
            X1              =   2880
            X2              =   2880
            Y1              =   4800
            Y2              =   5160
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   3360
            X2              =   2880
            Y1              =   5160
            Y2              =   5160
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            X1              =   3360
            X2              =   2880
            Y1              =   4800
            Y2              =   4800
         End
         Begin VB.Label Label1 
            Caption         =   "Comisión Interna..:"
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
            Index           =   18
            Left            =   3480
            TabIndex        =   59
            Top             =   5040
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Comisión Vendedor.:"
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
            Index           =   17
            Left            =   3480
            TabIndex        =   58
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Comisiones Registradas...:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   720
            TabIndex        =   55
            Top             =   4860
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Balance de Cobranza"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   15
            Left            =   120
            TabIndex        =   49
            Top             =   4200
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Realizados ...:"
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
            Index           =   14
            Left            =   3480
            TabIndex        =   45
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Ult. Mov.:"
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
            Index           =   13
            Left            =   3480
            TabIndex        =   44
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha 1er. Deduc..:"
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
            Index           =   12
            Left            =   3480
            TabIndex        =   43
            Top             =   4200
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Ultima Cuota Reportada"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   41
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Información de Cobros...:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   3120
            TabIndex        =   37
            Top             =   3120
            Width           =   2295
         End
         Begin VB.Label Label1 
            Caption         =   "Información de Pagos...:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   3120
            TabIndex        =   36
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label lblPlazo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo"
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
            Left            =   120
            TabIndex        =   29
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Estado"
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
            Index           =   10
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Vendedor"
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
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Cédula"
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
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   735
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   120
            X2              =   8160
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label1 
            Caption         =   "Monto"
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
            Left            =   120
            TabIndex        =   25
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Realizados ...:"
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
            Left            =   3480
            TabIndex        =   24
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Ult. Pago .:"
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
            Left            =   3480
            TabIndex        =   23
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha 1er. Pago ..:"
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
            Index           =   7
            Left            =   3480
            TabIndex        =   22
            Top             =   2760
            Width           =   1335
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuota"
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
            Left            =   120
            TabIndex        =   21
            Top             =   3000
            Width           =   735
         End
         Begin VB.Label lblContrato 
            Caption         =   "Tipo Seguro"
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
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lblPagador 
            Caption         =   "Tipo Cuenta"
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
            Left            =   120
            TabIndex        =   19
            Top             =   1200
            Width           =   1095
         End
      End
      Begin MSComctlLib.ListView lswPagos 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   31
         Top             =   540
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   16711680
         BackColor       =   16777215
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No.Cuota"
            Object.Width           =   1678
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Com.Interna"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Com.Vendedor"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswCobros 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   52
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10398
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Monto"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo Doc."
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Num.Doc."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Concepto"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   51
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
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
         Index           =   19
         Left            =   240
         TabIndex        =   32
         Top             =   4500
         Width           =   855
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3360
      TabIndex        =   33
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
   End
   Begin VB.Label lblNombre 
      Caption         =   "..."
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
      Left            =   4320
      TabIndex        =   35
      Top             =   480
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "No. Poliza"
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
      Index           =   0
      Left            =   240
      TabIndex        =   34
      Top             =   480
      Width           =   975
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   255
      Left            =   3960
      Top             =   480
      Width           =   255
   End
End
Attribute VB_Name = "frmInsRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean
Dim vFecha          As Date

Function fxPersonaNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select nombre from Socios where cedula = '" & strCedula & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxPersonaNombre = ""
Else
 fxPersonaNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close

End Function



Private Sub ReporteBoleta()
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

'strRuta = SIFGlobal.fxSIFPathReportes("CxC_BoletaActivacion.rpt")
'
'With frmContenedor.Crt
' .Reset
' .WindowShowRefreshBtn = True
' .WindowShowPrintSetupBtn = True
' .WindowState = crptMaximized
' .WindowShowSearchBtn = True
' .WindowTitle = "CxC...: Boleta de Activación"
' .ReportFileName = strRuta
'
' .Connect = glogon.ConectRPT
'
' .SelectionFormula = "{Ins_Polizas.num_poliza}=" & txtPoliza.Text
' .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
' .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
' .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'
' .SubreportToChange = "sbAsiento"
' .StoredProcParam(0) = "CxC_FRM"
' .StoredProcParam(1) = txtPoliza.Text
' .StoredProcParam(2) = 0
' .PrintReport
'End With

Me.MousePointer = vbDefault

End Sub



Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

fxValida = True
vMensaje = ""


If Len(txtPoliza.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- No se indicó el número de la póliza!"


If IsNumeric(txtPlazo) Then
 If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo es Inválido"
End If


If IsNumeric(txtMonto.Text) Then
 If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto NO es válido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto No es Inválido"
End If


strSQL = "select count(*) as Existe from Ins_Tipos_Seguros where Tipo_Seguro ='" & txtTipoSeguroCod.Text & "' and Activo = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de seguro no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from Ins_Tipos_Cuentas where Tipo_Cuenta ='" & txtTipoCuentaCod.Text & "' and Activo = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Cuenta no se encuentra Activa!"
rs.Close


strSQL = "select count(*) as Existe from Ins_Vendedores where cod_vendedor = " & txtVendedorCod.Text & " and Activo = 1"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El vendedor no se encuentra Activo!"
rs.Close

strSQL = "select count(*) as Existe from Socios where cedula = '" & txtCedula.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs!existe = 0 Then vMensaje = vMensaje & vbCrLf & "- La persona no existe en la base de datos!"
rs.Close


'Verificar que la persona no tenga prestamos en Cobro Judicial Activos
strSQL = "select coalesce(count(*),0) as Existe from Reg_Creditos" _
       & " where estado = 'A' and proceso = 'J' and cedula = '" & txtCedula & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!existe > 0 Then vMensaje = vMensaje & vbCrLf & "- La persona tiene créditos en Cobro Judicial"
rs.Close


If Len(vMensaje) > 0 Then fxValida = False

End Function


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtPoliza.Text = "" Then txtPoliza.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 num_poliza from Ins_Polizas"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where num_poliza > '" & txtPoliza & "' order by num_poliza asc"
Else
   strSQL = strSQL & " where num_poliza < '" & txtPoliza & "' order by num_poliza desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  txtPoliza.Text = rs!Num_Poliza
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


Private Sub FlatScrollBarTipoSeguro_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoSeguro.Tag = "" Then FlatScrollBarTipoSeguro.Tag = 0

strSQL = "select Top 1 Tipo_Seguro,Descripcion from Ins_Tipos_Seguros"

If FlatScrollBarTipoSeguro.Value > CLng(FlatScrollBarTipoSeguro.Tag) Then
   strSQL = strSQL & " where Tipo_Seguro > '" & txtTipoSeguroCod.Text & "' and Activo = 1 order by Tipo_Seguro asc"
Else
   strSQL = strSQL & " where Tipo_Seguro < '" & txtTipoSeguroCod.Text & "' and Activo = 1 order by Tipo_Seguro asc"
End If

FlatScrollBarTipoSeguro.Tag = FlatScrollBarTipoSeguro.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
        txtTipoSeguroCod.Text = rs!Tipo_Seguro
        txtTipoSeguroDesc.Text = rs!Descripcion
Else
        txtTipoSeguroCod.Text = ""
        txtTipoSeguroDesc.Text = ""
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub FlatScrollBarTipoCuenta_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBarTipoCuenta.Tag = "" Then FlatScrollBarTipoCuenta.Tag = 0

strSQL = "select Top 1 Tipo_Cuenta,Descripcion from Ins_Tipos_Cuentas"

If FlatScrollBarTipoCuenta.Value > CLng(FlatScrollBarTipoCuenta.Tag) Then
   strSQL = strSQL & " where Tipo_Cuenta > '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by Tipo_Cuenta asc"
Else
   strSQL = strSQL & " where Tipo_Cuenta < '" & txtTipoCuentaCod.Text & "' and Activo = 1 order by Tipo_Cuenta asc"
End If

FlatScrollBarTipoCuenta.Tag = FlatScrollBarTipoCuenta.Value

rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  txtTipoCuentaCod.Text = rs!Tipo_Cuenta
  txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

 vModulo = 17
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpiaPantalla


End Sub

Private Sub sbLimpiaPantalla()
 

Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(11).Picture
ImgAutorizacion.ToolTipText = "Pendiente: Consulta/Nuevo"
 
 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False

 txtCedula = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 
 txtVendedorCod.Text = ""
 txtVendedorDesc.Text = ""

 txtTipoSeguroCod.Text = ""
 txtTipoSeguroDesc.Text = ""

 txtTipoCuentaCod.Text = ""
 txtTipoCuentaDesc.Text = ""
   
 txtEstado.Text = "Pendiente"
   
 txtMonto = "0"
 txtPlazo = "60"
 txtCuota = "0"
  
 txtNumCuota.Text = 0
 txtBalance.Text = 0
 txtPagoPriMov.Text = ""
 txtPagoTotal.Text = 0
 txtPagoUltMov.Text = 0
 
 txtCobroPriDeduc.Text = ""
 txtCobroTotal.Text = 0
 txtCobroUltMov.Text = ""
 
 txtComisionInterna.Text = 0
 txtComisionVendedor.Text = 0
 
 txtNotas = ""
 
 SSTab.Tab = 0
 SSTab.TabEnabled(1) = False
 SSTab.TabEnabled(1) = False

 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 StatusBarX.Panels(4).Text = ""
 StatusBarX.Panels(5).Text = ""

End Sub



Private Sub sbConsulta()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

vPaso = True

strSQL = "select Pol.*,Ts.Descripcion as 'TipoSeguroDesc', Per.Nombre, isnull(Pol.Estado,'P') as 'Estado'" _
       & ",Ven.Nombre as 'VendedorNombre',Tc.descripcion as 'TipoCuentaDesc',dbo.MyGetdate() as 'FechaServer'" _
       & " from Ins_Polizas Pol inner join Ins_Tipos_Seguros Ts on Pol.Tipo_Seguro = Ts.Tipo_Seguro" _
       & " inner join Socios Per on Pol.cedula = Per.cedula" _
       & " left join Ins_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
       & " left join Ins_Tipos_Cuentas Tc on Pol.Tipo_Cuenta = Tc.Tipo_Cuenta" _
       & " where Pol.num_poliza = '" & txtPoliza.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic, adLockOptimistic

If Not rs.EOF And Not rs.BOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
 
 SSTab.TabEnabled(1) = True
 SSTab.TabEnabled(2) = True
 
 vFecha = rs!FechaServer
 
 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 lblNombre.Caption = txtNombre.Text
 
 txtVendedorCod.Text = rs!cod_vendedor
 txtVendedorDesc.Text = rs!VendedorNombre
 
 txtTipoSeguroCod.Text = rs!Tipo_Seguro
 txtTipoSeguroDesc.Text = rs!TipoSeguroDesc
 txtTipoCuentaCod.Text = rs!Tipo_Cuenta
 txtTipoCuentaDesc.Text = rs!TipoCuentaDesc
 
 
 txtMonto.Text = Format(IIf(IsNull(rs!MONTO), 0, rs!MONTO), "Standard")
 txtPlazo.Text = rs!Plazo
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 txtNumCuota.Text = rs!NUM_CUOTA & ""
 
 txtCobroPriDeduc.Text = rs!Cobrado_Primer_Deduc & ""
 txtCobroTotal.Text = rs!Cobrado_Total & ""
 txtCobroUltMov.Text = rs!Cobrado_Fecha_Ult & ""
 
 txtPagoPriMov.Text = rs!Pagado_Primer_Pago & ""
 txtPagoTotal.Text = Format(rs!Pagado_Total, "Standard")
 txtPagoTotal.ToolTipText = "Neto ..:" & Format(rs!Pagado_Total_Neto, "Standard")
 
 txtPagoUltMov.Text = rs!Pagado_Fecha_Ult & ""

 txtNotas = IIf(IsNull(rs!notas), "", rs!notas)


 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False

 Select Case rs!Estado
   Case "P" 'Pendiente
      txtEstado.Text = "Pendiente"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(7).Picture
      ImgAutorizacion.ToolTipText = "Activación: Pendiente"
      
      tlbAux.Buttons.Item(1).Enabled = True
      tlbAux.Buttons.Item(3).Enabled = True
   Case "A"
      txtEstado.Text = "Activa"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(5).Picture
      ImgAutorizacion.ToolTipText = "Póliza Activada!"
      tlbAux.Buttons.Item(3).Enabled = True
   Case "C"
      txtEstado.Text = "Cerrada"
      Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(6).Picture
      ImgAutorizacion.ToolTipText = "Póliza Cerrada (Inactivada)"
  End Select

 txtComisionInterna.Text = Format(rs!Comision_Interna_Total, "Standard")
 txtComisionVendedor.Text = Format(rs!Comision_Vendedor_Total, "Standard")

 StatusBarX.Panels(1).Text = rs!Registro_Usuario
 StatusBarX.Panels(2).Text = rs!registro_Fecha
 StatusBarX.Panels(3).Text = rs!Activa_Usuario & ""
 StatusBarX.Panels(4).Text = rs!Activa_Fecha & ""
 StatusBarX.Panels(5).Text = rs!Cierra_usuario & ""
 StatusBarX.Panels(6).Text = rs!Cierra_fecha & ""
 
 txtBalance.Text = Format(rs!Cobrado_Total - rs!Pagado_Total, "Standard")

Else
 If vEdita Then
    MsgBox "No existe la Póliza, verifique!", vbCritical
 End If
End If
rs.Close

vPaso = False

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub sbBorrar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "P" Then
    MsgBox "No se puede modificar esta póliza porque no se encuentra pendiente", vbExclamation
    Exit Sub
End If

strSQL = "delete Ins_Polizas where num_poliza = '" & txtPoliza.Text & "'"
glogon.Conection.Execute strSQL

MsgBox "Poliza Eliminada!", vbInformation

Call sbLimpiaPantalla

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtPoliza.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If txtPoliza.Text = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "nombre"
'       gBusquedas.Orden = "nombre"
'       gBusquedas.Consulta = "select cod_abogado,nombre from Cbr_Cj_Abogados"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = gBusquedas.Resultado
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Select Case Button.Key
  Case "Activar"
        i = MsgBox("Esta seguro que desea >> Activar << esta Póliza", vbYesNo)
        If i = vbYes Then
            
            strSQL = "exec spInsPolizaActiva '" & txtPoliza.Text & "','" & glogon.Usuario & "'"
            glogon.Conection.Execute strSQL
        
            'BITACORA
            Call Bitacora("Registra", "Activación de la Póliza: " & txtPoliza)
            
            MsgBox "Póliza Activada Satisfactoriamente!", vbInformation
            
        End If
        
  Case "Cerrar"
        i = MsgBox("Esta seguro que desea >> Cerrar << esta Póliza", vbYesNo)
        If i = vbYes Then
           Call sbPolizaCierra(txtPoliza.Text)
        End If
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
End Sub

Private Sub txtCuota_GotFocus()
On Error GoTo vError

txtCuota.Text = CCur(txtCuota.Text)

vError:
End Sub

Private Sub txtCuota_LostFocus()
On Error GoTo vError

txtCuota.Text = Format(CCur(txtCuota.Text), "Standard")

vError:

End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuota.SetFocus
End Sub

Private Sub txtPoliza_LostFocus()
  Call sbConsulta
End Sub

Private Sub txtTipoSeguroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Seguro"
   gBusquedas.Columna = "Tipo_Seguro"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Seguro,Descripcion  from Ins_Tipos_Seguros"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoSeguroCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Ins_Tipos_Seguros where Tipo_Seguro = '" & txtTipoSeguroCod.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   txtTipoSeguroCod.Text = rs!Tipo_Seguro
   txtTipoSeguroDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub



Private Sub txtTipoSeguroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Seguro,Descripcion  from Ins_Tipos_Seguros"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoSeguroCod.Text = gBusquedas.Resultado
      txtTipoSeguroDesc.Text = gBusquedas.Resultado2
      Call txtTipoSeguroCod_LostFocus
   End If
End If
End Sub


'--Tipo de Cuenta
Private Sub txtTipoCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Tipo_Cuenta"
   gBusquedas.Columna = "Tipo_Cuenta"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Cuenta,Descripcion  from Ins_Tipos_Cuentas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If

End Sub

Private Sub txtTipoCuentaCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Ins_Tipos_Cuentas where Tipo_Cuenta = '" & txtTipoCuentaCod.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   txtTipoCuentaCod.Text = rs!Tipo_Cuenta
   txtTipoCuentaDesc.Text = rs!Descripcion
End If
rs.Close


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub



Private Sub txtTipoCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Filtro = " and Activo = 1"
   gBusquedas.Consulta = "Select Tipo_Cuenta,Descripcion  from Ins_Tipos_Cuentas"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtTipoCuentaCod.Text = gBusquedas.Resultado
      txtTipoCuentaDesc.Text = gBusquedas.Resultado2
      Call txtTipoCuentaCod_LostFocus
   End If
End If
End Sub




Private Sub txtCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub


Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)


Select Case SSTab.Tab
  Case 1 'Recepcion
'    Call sbInitOpciones
  Case 2 'Historial
'    ssTabHistorial.Tab = 0
'    Call sbHistorial
  Case Else
End Select


End Sub

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vExiste As Integer

On Error GoTo vError

       
If Mid(txtEstado.Text, 1, 1) <> "P" Then
    MsgBox "No se puede modificar esta póliza porque no se encuentra pendiente", vbExclamation
    Exit Sub
End If
       
If Not vEdita Then
   strSQL = "insert Ins_Polizas(num_poliza,CEDULA,cod_vendedor,Tipo_Seguro,Tipo_Cuenta,NOTAS,MONTO,CUOTA,PLAZO" _
          & ",ESTADO,REGISTRO_FECHA,REGISTRO_USUARIO)" _
          & " VALUES('" & txtPoliza.Text & "','" & txtCedula.Text & "'," & txtVendedorCod.Text & ",'" & txtTipoSeguroCod.Text & "','" & txtTipoCuentaCod.Text _
          & "','" & txtNotas.Text & "'," & CCur(txtMonto.Text) & "," & CCur(txtMonto.Text) & "," & txtPlazo.Text & ",'P',dbo.MyGetdate(),'" & glogon.Usuario & "')"
   glogon.Conection.Execute strSQL
          
          
Else
   strSQL = "update Ins_Polizas set cod_vendedor = " & txtVendedorCod.Text & ",Tipo_Seguro = '" & txtTipoSeguroCod.Text & "',Tipo_Cuenta = '" _
          & txtTipoCuentaCod.Text & "',notas = '" & txtNotas.Text & "',Monto = " & CCur(txtMonto.Text) & ", Cuota =  " & CCur(txtCuota.Text) _
          & ", Plazo = " & txtPlazo.Text & ", cedula = '" & txtCedula.Text _
          & "' where num_poliza = '" & txtPoliza.Text & "'"
   glogon.Conection.Execute strSQL


End If

MsgBox "Póliza Registrada / Actualizada satisactoriamente!", vbInformation


Call sbConsulta

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cedula"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If


End Sub

Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


txtNombre.Text = fxPersonaNombre(txtCedula)
lblNombre.Caption = txtNombre.Text


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtCedula.Text = gBusquedas.Resultado
      txtNombre.Text = gBusquedas.Resultado2
      Call txtCedula_LostFocus
   End If
End If
End Sub

'--Vendedor
Private Sub txtVendedorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVendedorDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Cod_Vendedor"
   gBusquedas.Columna = "Cod_Vendedor"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from Ins_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If


End Sub

Private Sub txtVendedorCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select Cod_Vendedor,Nombre from Ins_Vendedores where cod_Vendedor = " & txtVendedorCod.Text
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
    txtVendedorDesc.Text = rs!Nombre
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtVendedorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoSeguroCod.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Filtro = ""
   gBusquedas.Consulta = "Select Cod_Vendedor,Nombre from Ins_Vendedores"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtVendedorCod.Text = gBusquedas.Resultado
      txtVendedorDesc.Text = gBusquedas.Resultado2
      Call txtVendedorCod_LostFocus
   End If
End If
End Sub


Public Sub sbConsultaExterna(xOpTemp As String)
 txtPoliza.Text = xOpTemp
 Call sbConsulta
End Sub


Private Sub txtPoliza_Change()
 Call sbLimpiaPantalla

End Sub

Private Sub txtPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub


