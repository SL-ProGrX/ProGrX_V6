VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPosCajasCierres 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierres de Cajas"
   ClientHeight    =   7248
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9588
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7248
   ScaleWidth      =   9588
   Begin TabDlg.SSTab ssTab 
      Height          =   4932
      Left            =   48
      TabIndex        =   0
      Top             =   2280
      Width           =   9432
      _ExtentX        =   16637
      _ExtentY        =   8700
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Movimientos"
      TabPicture(0)   =   "frmPosCajasCierres.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(9)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdMovimientos"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCierre"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdReporte"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtEA_CxCExternaInicial"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtEA_CxCInternaInicial"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtEA_EfectivoInicial"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtApertura"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lswMov"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCI_Efectivo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCI_CxCInterna"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCI_CxCExterna"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCierre"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Depósitos"
      TabPicture(1)   =   "frmPosCajasCierres.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtDepObservacion"
      Tab(1).Control(1)=   "txtDepMonto"
      Tab(1).Control(2)=   "txtDepDocumento"
      Tab(1).Control(3)=   "cmdDeposito"
      Tab(1).Control(4)=   "Label3(1)"
      Tab(1).Control(5)=   "Label3(0)"
      Tab(1).Control(6)=   "Label2(1)"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Traslado de Documentos"
      TabPicture(2)   =   "frmPosCajasCierres.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "opt(1)"
      Tab(2).Control(1)=   "opt(0)"
      Tab(2).Control(2)=   "lswDoc"
      Tab(2).Control(3)=   "cmdDocumentos"
      Tab(2).ControlCount=   4
      Begin VB.TextBox txtCierre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1160
         Width           =   1692
      End
      Begin VB.TextBox txtCI_CxCExterna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1160
         Width           =   2412
      End
      Begin VB.TextBox txtCI_CxCInterna 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1160
         Width           =   2412
      End
      Begin VB.TextBox txtCI_Efectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1160
         Width           =   2652
      End
      Begin VB.OptionButton opt 
         Caption         =   "Documentos x Cobrar Externos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   -70320
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   4572
      End
      Begin VB.OptionButton opt 
         Caption         =   "Documentos x Cobrar Internos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   480
         Value           =   -1  'True
         Width           =   4452
      End
      Begin MSComctlLib.ListView lswDoc 
         Height          =   3372
         Left            =   -74880
         TabIndex        =   16
         Top             =   840
         Width           =   9132
         _ExtentX        =   16108
         _ExtentY        =   5948
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
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "# Factura"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Documento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Forma Pago"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.TextBox txtDepObservacion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   -73320
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   1080
         Width           =   7332
      End
      Begin VB.TextBox txtDepMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68040
         TabIndex        =   13
         Top             =   720
         Width           =   2052
      End
      Begin VB.TextBox txtDepDocumento 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73320
         TabIndex        =   11
         Top             =   720
         Width           =   1935
      End
      Begin MSComctlLib.ListView lswMov 
         Height          =   2652
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   9132
         _ExtentX        =   16108
         _ExtentY        =   4678
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Origen"
            Object.Width           =   3069
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Comprobante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Detalle"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtApertura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1692
      End
      Begin VB.TextBox txtEA_EfectivoInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   2652
      End
      Begin VB.TextBox txtEA_CxCInternaInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   2412
      End
      Begin VB.TextBox txtEA_CxCExternaInicial 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   2412
      End
      Begin XtremeSuiteControls.PushButton cmdDeposito 
         Height          =   492
         Left            =   -67920
         TabIndex        =   34
         Top             =   2640
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Aplicar Depósito"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPosCajasCierres.frx":0054
      End
      Begin XtremeSuiteControls.PushButton cmdDocumentos 
         Height          =   492
         Left            =   -67680
         TabIndex        =   35
         Top             =   4320
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Traslada Documentos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPosCajasCierres.frx":082C
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   492
         Left            =   2160
         TabIndex        =   36
         Top             =   4320
         Width           =   1572
         _Version        =   1245187
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Informes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPosCajasCierres.frx":1004
      End
      Begin XtremeSuiteControls.PushButton cmdCierre 
         Height          =   492
         Left            =   5640
         TabIndex        =   37
         Top             =   4320
         Width           =   2052
         _Version        =   1245187
         _ExtentX        =   3619
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Cierre de Caja"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPosCajasCierres.frx":17C0
      End
      Begin XtremeSuiteControls.PushButton cmdMovimientos 
         Height          =   492
         Left            =   3720
         TabIndex        =   38
         Top             =   4320
         Width           =   1932
         _Version        =   1245187
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Movimientos "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPosCajasCierres.frx":21AC
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cierre"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   10
         Left            =   7560
         TabIndex        =   26
         Top             =   924
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Actual CxC Externos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   9
         Left            =   5160
         TabIndex        =   24
         Top             =   924
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Actual CxC Internos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   6
         Left            =   2760
         TabIndex        =   22
         Top             =   924
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Actual en Efectivo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   924
         Width           =   2652
      End
      Begin VB.Label Label3 
         Caption         =   "Observación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69120
         TabIndex        =   12
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "# Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74520
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apertura"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   4
         Left            =   7560
         TabIndex        =   8
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Inicial en Efectivo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2652
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Inicial CxC  Internos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   7
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Width           =   2412
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Saldo Inicial CxC Externos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   8
         Left            =   5160
         TabIndex        =   4
         Top             =   360
         Width           =   2412
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   3600
      TabIndex        =   27
      Top             =   960
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   3600
      TabIndex        =   28
      Top             =   1320
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
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
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.FlatEdit txtClave 
      Height          =   312
      Left            =   3600
      TabIndex        =   29
      Top             =   1800
      Width           =   3012
      _Version        =   1245187
      _ExtentX        =   5313
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PasswordChar    =   "*"
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Proceso de Cierre de Cajas de POS: Inicie Sesión en su Caja para proceder con el cierre"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   1200
      TabIndex        =   33
      Top             =   360
      Width           =   8172
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   2160
      TabIndex        =   32
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Caja"
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
      Height          =   312
      Index           =   1
      Left            =   2160
      TabIndex        =   31
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clave"
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
      Height          =   312
      Index           =   2
      Left            =   2160
      TabIndex        =   30
      Top             =   1800
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPosCajasCierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bInicio As Boolean

Private Function fxExisteApertura(vCaja As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from pv_cajas_ac where usuario = '" _
       & glogon.Usuario & "' and cod_caja = '" & vCaja & "' and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   fxExisteApertura = True
Else
   fxExisteApertura = False
End If
rs.Close

End Function

Private Sub cbo_Click()
Call sbLimpiaDatos
End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or vbKeyTab Then txtClave.SetFocus
End Sub

Private Sub cmdCierre_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea realizar el cierre de caja?", vbYesNo)
If i = vbNo Then Exit Sub

strSQL = "update pv_cajas_ac set ci_fecha = dbo.MyGetdate(),estado = 'C' where cod_ac = " & txtCierre _
       & " and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "' and usuario = '" & txtUsuario & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Cierre Caja: " & cbo.ItemData(cbo.ListIndex) & " US: " & txtUsuario)


MsgBox "Cierre Realizado Satisfactoriamente...", vbInformation

strSQL = "{PV_CAJAS.COD_CAJA} = '" & cbo.ItemData(cbo.ListIndex) & "' and {PV_CAJAS.USUARIO} = '" _
       & txtUsuario & "' and {PV_CAJAS_AC.COD_AC} = " & txtApertura
Call sbPosReportes("CAJAS_CIERRE", "CIERRE DE CAJAS", "APERTURA/CIERRE : " & txtApertura, strSQL)

Call sbLimpiaDatos
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdDeposito_Click()
Dim strSQL As String, vMensaje As String
Dim rs As New ADODB.Recordset, vUltimo As Integer

vMensaje = ""

If Trim(txtDepDocumento) = 0 Then vMensaje = vMensaje & vbCrLf & " - El Número de Documento no es válido"
If Trim(txtDepObservacion) = 0 Then vMensaje = vMensaje & vbCrLf & " - No se Especificó ninguna observación"
If IsNumeric(txtDepMonto) Then
    If CCur(txtDepMonto) <= 0 Then vMensaje = vMensaje & vbCrLf & " - EL Monto no es válido"
Else
    vMensaje = vMensaje & vbCrLf & " - EL Monto no es válido"
End If

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  Exit Sub
End If

strSQL = "select isnull(max(cod_deposito),0) as Ultimo from pv_cajas_depositos" _
       & " where cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "' and Usuario = '" & txtUsuario & "'"
Call OpenRecordSet(rs, strSQL)
vUltimo = rs!ultimo + 1
rs.Close

'REGISTRA EL MONTO DE DEPOSITOS
strSQL = "insert pv_cajas_depositos(cod_deposito,cod_caja,usuario,cod_ac,fecha,monto,observacion" _
       & ",documento,tipo) values(" & vUltimo & ",'" & cbo.ItemData(cbo.ListIndex) & "','" & txtUsuario _
       & "'," & txtApertura.Text & ",dbo.MyGetdate()," & CCur(txtDepMonto) & ",'" & Mid(txtDepObservacion, 1, 254) _
       & "','" & Mid(txtDepDocumento, 1, 25) & "','E')"
Call ConectionExecute(strSQL)


'Se envia indefinido la forma de pago, para que el procedimiento utilice por defecto Efectivo
Call sbPosCajaMovRegistra("DP", cbo.ItemData(cbo.ListIndex), txtUsuario, CInt(txtApertura), CCur(txtDepMonto), 99999999, CStr(vUltimo), txtDepDocumento & " - " & txtDepObservacion)

MsgBox "Depósito aplicado satisfactoriamente...", vbInformation

txtDepDocumento = ""
txtDepMonto = 0
txtDepObservacion = ""

txtDepDocumento.SetFocus

End Sub

Private Sub cmdMovimientos_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswMov.ListItems.Clear

strSQL = "select * from pv_cajas_mov where cod_caja = '" & cbo.ItemData(cbo.ListIndex) _
       & "' and usuario = '" & glogon.Usuario & "' and cod_ac = " & txtApertura
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswMov.ListItems.Add(, , rs!fecha)
     itmX.SubItems(1) = fxPosCajaTipoMov(rs!Tipo)
     itmX.SubItems(2) = fxPosCajaOrigenMov(rs!origen)
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     itmX.SubItems(4) = rs!numcom & ""
     itmX.SubItems(5) = rs!Detalle & ""
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

i = MsgBox("Desea el reporte de Caja con movimientos detallados?", vbYesNo)
 
strSQL = "{PV_CAJAS.COD_CAJA} = '" & cbo.ItemData(cbo.ListIndex) & "' and {PV_CAJAS.USUARIO} = '" _
       & txtUsuario & "' and {PV_CAJAS_AC.COD_AC} = " & txtApertura

If i = vbYes Then
    Call sbPosReportes("CAJASMOVDET", "MOVIMIENTOS DE CAJAS", "APERTURA/CIERRE : " & txtApertura, strSQL)
Else
    Call sbPosReportes("CAJASMOVRSM", "MOVIMIENTOS DE CAJAS", "APERTURA/CIERRE : " & txtApertura, strSQL)
End If

Exit Sub
vError:

End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 33

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

'Solo carga las cajas que no tengan aperturas abiertas y que no esten bloqueadas
'del usuario activo

txtUsuario = glogon.Usuario
txtClave = ""

strSQL = "select rtrim(Cd.cod_caja) as 'IdX' , rtrim(Cd.nombre) as 'ItmX'" _
       & " from pv_cajas Cd" _
       & " where Cd.estado = 'A' and Cd.usuario = '" & glogon.Usuario & "' and Cd.Bloqueo = 0" _
       & " and dbo.fxPOS_Caja_Apertura_Existe(Cd.cod_Caja, Cd.Usuario) = 1"

Call sbCbo_Llena_New(cbo, strSQL, False, True)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaDatos()
    lswMov.ListItems.Clear
    txtApertura = 0
    txtEA_CxCExternaInicial = "0.00"
    txtEA_CxCInternaInicial = "0.00"
    txtEA_EfectivoInicial = "0.00"
    
    txtCierre = 0
    txtCI_CxCExterna = "0.00"
    txtCI_CxCInterna = "0.00"
    txtCI_Efectivo = "0.00"
    
    cmdCierre.Enabled = False
    ssTab.Tab = 0
    ssTab.TabEnabled(1) = False
    ssTab.TabEnabled(2) = False
    
End Sub

Private Sub SSTab_Click(PreviousTab As Integer)

Select Case ssTab.Tab
  Case 0 'Movimientos
    Call txtClave_KeyDown(vbKeyReturn, 1)
  Case 1 'Nada
  Case 2 'Carga
End Select

End Sub


Private Sub txtClave_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

'Verifica datos nuevamente por razones de seguridad por violación
'por concurrencias.

'1. Verificar que el Estado este Activa / en el Load esta validado
'2. Que no se encuentre Bloqueada
'3. Verificar si la caja esta abierta (Apertura) y Sacar el Consecutivo
'   de la apertura.

If bInicio Then Exit Sub

If KeyCode = vbKeyReturn Then
 strSQL = "select bloqueo from pv_cajas where usuario = '" _
        & txtUsuario & "' and cod_caja = '" & cbo.ItemData(cbo.ListIndex) & "' and clave = '" _
        & fxPosEncrypta(txtClave) & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs.EOF And rs.BOF Then
   MsgBox "Caja: verifique su Usuario y Clave para Esta Caja ...", vbExclamation
 Else
  If rs!bloqueo = 0 Then
     gCajas.Caja = cbo.ItemData(cbo.ListIndex)
     gCajas.Usuario = txtUsuario
     
     'Revisa que la caja no sea nueva
     rs.Close
     strSQL = "select isnull(max(cod_ac),0) as UltCierre from pv_cajas_ac where cod_caja = '" & gCajas.Caja _
            & "' and usuario = '" & gCajas.Usuario & "' and estado = 'A'"
     Call OpenRecordSet(rs, strSQL)
     
     gCajas.Apertura = rs!ultCierre
     cmdCierre.Enabled = True
     
     If rs!ultCierre = 0 Then
        Call sbLimpiaDatos
     Else
        strSQL = "select * from pv_cajas_ac where cod_caja = '" & gCajas.Caja _
               & "' and usuario = '" & gCajas.Usuario & "' and cod_ac = " & rs!ultCierre
        rs.Close
        Call OpenRecordSet(rs, strSQL)
        
        txtApertura = rs!cod_ac
        txtEA_CxCExternaInicial = Format(rs!ap_saldo_docext, "Standard")
        txtEA_CxCInternaInicial = Format(rs!ap_saldo_docint, "Standard")
        txtEA_EfectivoInicial = Format(rs!ap_saldo_efectivo, "Standard")
        
        txtCierre = rs!cod_ac
        txtCI_CxCExterna = Format(rs!ci_saldo_docext, "Standard")
        txtCI_CxCInterna = Format(rs!ci_saldo_docint, "Standard")
        txtCI_Efectivo = Format(rs!ci_saldo_efectivo, "Standard")
        
        Call cmdMovimientos_Click
        ssTab.TabEnabled(1) = True
        ssTab.TabEnabled(2) = True
     End If
  
  Else
    MsgBox "La Caja se encuentra Bloqueada...", vbExclamation
  
  End If 'Bloqueo
 
 End If 'Select cajas
 rs.Close

End If

End Sub

Private Sub txtDepMonto_GotFocus()
On Error GoTo vError
  txtDepMonto = CCur(txtDepMonto)
  txtDepMonto.SelStart = Len(txtDepMonto)
vError:
End Sub

Private Sub txtDepMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDepObservacion.SetFocus
End Sub

Private Sub txtDepMonto_LostFocus()
On Error GoTo vError
  txtDepMonto = Format(CCur(txtDepMonto), "Standard")
vError:
End Sub


Private Sub txtDepDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDepMonto.SetFocus
End Sub

Private Sub txtDepObservacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdDeposito.SetFocus
End Sub
