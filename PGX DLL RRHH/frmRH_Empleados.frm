VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.TaskPanel.v22.1.0.ocx"
Begin VB.Form frmRH_Empleados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Empleados¦ Colaboradores"
   ClientHeight    =   8550
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11145
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Height          =   7128
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   2760
      _Version        =   1441793
      _ExtentX        =   4868
      _ExtentY        =   12573
      _StockProps     =   64
      VisualTheme     =   13
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.PushButton btnSalida 
      Height          =   330
      Left            =   9960
      TabIndex        =   129
      Top             =   0
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Salida"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":0000
      ImageAlignment  =   0
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   10440
      Top             =   600
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9840
      TabIndex        =   0
      Top             =   720
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ComboBox cboTipoId 
      Height          =   312
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4895
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   10680
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":061E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":0893
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":0B2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":0CAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":0E47
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":0FEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":1191
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":131C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":14A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":15AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":183A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":1946
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":1BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":1C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":1E28
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRH_Empleados.frx":1FC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   8295
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3245
            MinWidth        =   3245
            Object.ToolTipText     =   "Usuario Ingresa"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3599
            MinWidth        =   3599
            Object.ToolTipText     =   "Fecha Ingreso"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3245
            MinWidth        =   3245
            Object.ToolTipText     =   "Usuario Modifica"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3599
            MinWidth        =   3599
            Object.ToolTipText     =   "Fecha Modificacion"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   2760
      TabIndex        =   12
      Top             =   1440
      Width           =   8415
      _Version        =   1441793
      _ExtentX        =   14843
      _ExtentY        =   12938
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
      PaintManager.Position=   2
      ItemCount       =   4
      SelectedItem    =   3
      Item(0).Caption =   "Datos de Contacto"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "cboSexo"
      Item(0).Control(1)=   "dtpNacimiento"
      Item(0).Control(2)=   "txtApellido1"
      Item(0).Control(3)=   "txtApellido2"
      Item(0).Control(4)=   "Label2"
      Item(0).Control(5)=   "Label3"
      Item(0).Control(6)=   "Label4"
      Item(0).Control(7)=   "Label1(0)"
      Item(0).Control(8)=   "Label14"
      Item(0).Control(9)=   "Label15(0)"
      Item(0).Control(10)=   "txtNombre"
      Item(0).Control(11)=   "gbPersona(1)"
      Item(0).Control(12)=   "cboNacionalidad"
      Item(0).Control(13)=   "Label15(6)"
      Item(0).Control(14)=   "cboEstadoCivil"
      Item(0).Control(15)=   "gbBancos"
      Item(1).Caption =   "Laboral"
      Item(1).ControlCount=   11
      Item(1).Control(0)=   "FlatScroll_Laboral(3)"
      Item(1).Control(1)=   "txtProfesionCod"
      Item(1).Control(2)=   "txtProfesionDesc"
      Item(1).Control(3)=   "cboNivel"
      Item(1).Control(4)=   "Label9"
      Item(1).Control(5)=   "Label15(4)"
      Item(1).Control(6)=   "FlatScroll_Laboral(5)"
      Item(1).Control(7)=   "txtJefeCod"
      Item(1).Control(8)=   "txtJefeDesc"
      Item(1).Control(9)=   "Label21(1)"
      Item(1).Control(10)=   "tcLaboral"
      Item(2).Caption =   "Redes y Otros"
      Item(2).ControlCount=   16
      Item(2).Control(0)=   "picFoto"
      Item(2).Control(1)=   "btnFoto(0)"
      Item(2).Control(2)=   "btnFoto(1)"
      Item(2).Control(3)=   "gbPortal"
      Item(2).Control(4)=   "Label12(0)"
      Item(2).Control(5)=   "Label12(1)"
      Item(2).Control(6)=   "txtId_Tributario"
      Item(2).Control(7)=   "chkInd_Licencia"
      Item(2).Control(8)=   "chkInd_Vehiculo"
      Item(2).Control(9)=   "chkInd_Solidarista"
      Item(2).Control(10)=   "chkInd_Marcas"
      Item(2).Control(11)=   "txtNotas"
      Item(2).Control(12)=   "chkCF_Conyuge"
      Item(2).Control(13)=   "txtCF_Dependientes"
      Item(2).Control(14)=   "UpDownCf"
      Item(2).Control(15)=   "Label13"
      Item(3).Caption =   "Detalles"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "TituloOpciones"
      Item(3).Control(1)=   "lswHistorico"
      Item(3).Control(2)=   "btnEditarDetalle"
      Item(3).Control(3)=   "btnExportar"
      Begin XtremeSuiteControls.ListView lswHistorico 
         Height          =   6410
         Left            =   0
         TabIndex        =   13
         Top             =   360
         Width           =   8415
         _Version        =   1441793
         _ExtentX        =   14843
         _ExtentY        =   11307
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
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbBancos 
         Height          =   735
         Left            =   -69760
         TabIndex        =   130
         Top             =   1680
         Visible         =   0   'False
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Metodo de Pago"
         ForeColor       =   8421504
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   315
            Left            =   1440
            TabIndex        =   131
            Top             =   360
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboFormaPago 
            Height          =   330
            Left            =   4680
            TabIndex        =   133
            Top             =   360
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   132
            Top             =   360
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.TabControl tcLaboral 
         Height          =   5055
         Left            =   -70000
         TabIndex        =   79
         Top             =   1560
         Visible         =   0   'False
         Width           =   8415
         _Version        =   1441793
         _ExtentX        =   14843
         _ExtentY        =   8916
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
         SelectedItem    =   1
         Item(0).Caption =   "Actual"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "gbAccionPersonal"
         Item(1).Caption =   "Salida"
         Item(1).ControlCount=   14
         Item(1).Control(0)=   "txtSalidaFecha"
         Item(1).Control(1)=   "txtSalidaNotas"
         Item(1).Control(2)=   "txtSalidaEstado"
         Item(1).Control(3)=   "txtSalidaTipo"
         Item(1).Control(4)=   "txtSalidaTipoDesc"
         Item(1).Control(5)=   "Label16(0)"
         Item(1).Control(6)=   "Label17"
         Item(1).Control(7)=   "Label10(1)"
         Item(1).Control(8)=   "Label16(1)"
         Item(1).Control(9)=   "txtLiquidaFecha"
         Item(1).Control(10)=   "Label10(6)"
         Item(1).Control(11)=   "txtLiquidaBoleta"
         Item(1).Control(12)=   "Label10(7)"
         Item(1).Control(13)=   "btnBoletaLiq"
         Begin XtremeSuiteControls.PushButton btnBoletaLiq 
            Height          =   310
            Left            =   7200
            TabIndex        =   128
            Top             =   3000
            Width           =   310
            _Version        =   1441793
            _ExtentX        =   547
            _ExtentY        =   547
            _StockProps     =   79
            BackColor       =   -2147483633
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmRH_Empleados.frx":206F
         End
         Begin XtremeSuiteControls.GroupBox gbAccionPersonal 
            Height          =   4692
            Left            =   -69880
            TabIndex        =   80
            Top             =   360
            Visible         =   0   'False
            Width           =   7692
            _Version        =   1441793
            _ExtentX        =   13568
            _ExtentY        =   8276
            _StockProps     =   79
            ForeColor       =   4210752
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
            BorderStyle     =   2
            Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
               Height          =   255
               Index           =   0
               Left            =   7200
               TabIndex        =   81
               Top             =   120
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               Arrows          =   65536
               Orientation     =   1638401
            End
            Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
               Height          =   255
               Index           =   1
               Left            =   7200
               TabIndex        =   82
               Top             =   480
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               Arrows          =   65536
               Orientation     =   1638401
            End
            Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
               Height          =   255
               Index           =   2
               Left            =   7200
               TabIndex        =   83
               Top             =   840
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               Arrows          =   65536
               Orientation     =   1638401
            End
            Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
               Height          =   255
               Index           =   4
               Left            =   7200
               TabIndex        =   84
               Top             =   1320
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   450
               _Version        =   393216
               Arrows          =   65536
               Orientation     =   1638401
            End
            Begin XtremeSuiteControls.FlatEdit txtCentroCod 
               Height          =   315
               Left            =   1560
               TabIndex        =   85
               Top             =   120
               Width           =   735
               _Version        =   1441793
               _ExtentX        =   1291
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
               Locked          =   -1  'True
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
               Height          =   315
               Left            =   2280
               TabIndex        =   86
               Top             =   120
               Width           =   4815
               _Version        =   1441793
               _ExtentX        =   8488
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
            Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
               Height          =   315
               Left            =   1560
               TabIndex        =   87
               Top             =   480
               Width           =   735
               _Version        =   1441793
               _ExtentX        =   1291
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
               Locked          =   -1  'True
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
               Height          =   315
               Left            =   2280
               TabIndex        =   88
               Top             =   480
               Width           =   4815
               _Version        =   1441793
               _ExtentX        =   8488
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
            Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
               Height          =   315
               Left            =   1560
               TabIndex        =   89
               Top             =   840
               Width           =   735
               _Version        =   1441793
               _ExtentX        =   1291
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
               Locked          =   -1  'True
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtSecDesc 
               Height          =   315
               Left            =   2280
               TabIndex        =   90
               Top             =   840
               Width           =   4815
               _Version        =   1441793
               _ExtentX        =   8488
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
            Begin XtremeSuiteControls.FlatEdit txtPuestoCod 
               Height          =   315
               Left            =   1560
               TabIndex        =   91
               Top             =   1320
               Width           =   735
               _Version        =   1441793
               _ExtentX        =   1291
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
               Locked          =   -1  'True
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtPuestoDesc 
               Height          =   315
               Left            =   2280
               TabIndex        =   92
               Top             =   1320
               Width           =   4815
               _Version        =   1441793
               _ExtentX        =   8488
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
            Begin XtremeSuiteControls.ComboBox cboNomina 
               Height          =   315
               Left            =   1560
               TabIndex        =   93
               Top             =   2520
               Width           =   5535
               _Version        =   1441793
               _ExtentX        =   9763
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboContrato 
               Height          =   315
               Left            =   1560
               TabIndex        =   94
               Top             =   2880
               Width           =   5535
               _Version        =   1441793
               _ExtentX        =   9763
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboJornada 
               Height          =   315
               Left            =   1560
               TabIndex        =   95
               Top             =   3840
               Width           =   5535
               _Version        =   1441793
               _ExtentX        =   9763
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.ComboBox cboVacaciones 
               Height          =   315
               Left            =   1560
               TabIndex        =   96
               Top             =   4200
               Width           =   5535
               _Version        =   1441793
               _ExtentX        =   9763
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.DateTimePicker dtpIngreso 
               Height          =   315
               Left            =   1560
               TabIndex        =   97
               Top             =   2040
               Width           =   1335
               _Version        =   1441793
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
            Begin XtremeSuiteControls.FlatEdit txtSalario 
               Height          =   315
               Left            =   4680
               TabIndex        =   98
               Top             =   2040
               Width           =   2415
               _Version        =   1441793
               _ExtentX        =   4254
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
               Alignment       =   1
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.DateTimePicker dtpContrato 
               Height          =   315
               Left            =   1560
               TabIndex        =   99
               Top             =   3240
               Width           =   1335
               _Version        =   1441793
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
            Begin XtremeSuiteControls.ComboBox cboDivisa 
               Height          =   315
               Left            =   4680
               TabIndex        =   100
               Top             =   1680
               Width           =   2415
               _Version        =   1441793
               _ExtentX        =   4260
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
               Appearance      =   6
               UseVisualStyle  =   0   'False
            End
            Begin XtremeSuiteControls.FlatEdit txtVacaAcum 
               Height          =   315
               Left            =   4680
               TabIndex        =   127
               Top             =   3240
               Width           =   2415
               _Version        =   1441793
               _ExtentX        =   4254
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
               Locked          =   -1  'True
               Appearance      =   2
               UseVisualStyle  =   0   'False
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Vaca. Acum."
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3240
               TabIndex        =   126
               Top             =   3240
               Width           =   1212
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Salario:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   1
               Left            =   3240
               TabIndex        =   112
               Top             =   2040
               Width           =   1932
            End
            Begin VB.Label Label6 
               Caption         =   "Ingreso"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   0
               Left            =   240
               TabIndex        =   111
               Top             =   2040
               Width           =   732
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Régimen de Vacaciones"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   492
               Index           =   5
               Left            =   240
               TabIndex        =   110
               Top             =   4200
               Width           =   1092
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Jornada"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   3
               Left            =   240
               TabIndex        =   109
               Top             =   3840
               Width           =   1092
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
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
               Height          =   372
               Index           =   2
               Left            =   240
               TabIndex        =   108
               Top             =   2880
               Width           =   1092
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Nómina"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   372
               Index           =   1
               Left            =   240
               TabIndex        =   107
               Top             =   2520
               Width           =   1092
            End
            Begin VB.Label Label10 
               Caption         =   "Centro"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   12
               Left            =   240
               TabIndex        =   106
               Top             =   120
               Width           =   972
            End
            Begin VB.Label lblDepartamento 
               Caption         =   "Departamento"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   240
               TabIndex        =   105
               Top             =   480
               Width           =   1332
            End
            Begin VB.Label lblSeccion 
               Caption         =   "Sección"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   240
               TabIndex        =   104
               Top             =   840
               Width           =   1572
            End
            Begin VB.Label Label21 
               Caption         =   "Puesto"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   0
               Left            =   240
               TabIndex        =   103
               Top             =   1320
               Width           =   1572
            End
            Begin VB.Label lblVencimiento 
               BackStyle       =   0  'Transparent
               Caption         =   "Vencimiento"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   240
               TabIndex        =   102
               Top             =   3240
               Width           =   1212
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Divisa:"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   3
               Left            =   3240
               TabIndex        =   101
               Top             =   1680
               Width           =   1932
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtSalidaFecha 
            Height          =   312
            Left            =   1560
            TabIndex        =   113
            Top             =   600
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalidaEstado 
            Height          =   312
            Left            =   5160
            TabIndex        =   115
            Top             =   600
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalidaTipo 
            Height          =   312
            Left            =   1560
            TabIndex        =   116
            Top             =   960
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSalidaTipoDesc 
            Height          =   312
            Left            =   2280
            TabIndex        =   117
            Top             =   960
            Width           =   4812
            _Version        =   1441793
            _ExtentX        =   8488
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
         Begin XtremeSuiteControls.FlatEdit txtSalidaNotas 
            Height          =   792
            Left            =   1560
            TabIndex        =   114
            Top             =   1320
            Width           =   5532
            _Version        =   1441793
            _ExtentX        =   9758
            _ExtentY        =   1397
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtLiquidaFecha 
            Height          =   312
            Left            =   5160
            TabIndex        =   122
            Top             =   2640
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtLiquidaBoleta 
            Height          =   312
            Left            =   5160
            TabIndex        =   124
            Top             =   3000
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3408
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Liquida Boleta:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   7
            Left            =   3360
            TabIndex        =   125
            Top             =   3000
            Width           =   1692
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Liquida Fecha:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   6
            Left            =   3360
            TabIndex        =   123
            Top             =   2640
            Width           =   1692
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   121
            Top             =   1320
            Width           =   1572
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Salida"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   120
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   3960
            TabIndex        =   119
            Top             =   600
            Width           =   1332
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   118
            Top             =   960
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.UpDown UpDownCf 
         Height          =   315
         Left            =   -69040
         TabIndex        =   77
         Top             =   2880
         Visible         =   0   'False
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   556
         _StockProps     =   64
         Appearance      =   16
         UseVisualStyle  =   0   'False
         BuddyControl    =   ""
         BuddyProperty   =   ""
      End
      Begin XtremeSuiteControls.GroupBox gbPortal 
         Height          =   972
         Left            =   -69520
         TabIndex        =   57
         Top             =   5760
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1441793
         _ExtentX        =   13568
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Portal:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnPortal 
            Height          =   372
            Left            =   4200
            TabIndex        =   60
            Top             =   360
            Width           =   3492
            _Version        =   1441793
            _ExtentX        =   6159
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Enviar Clave al Correo"
            BackColor       =   -2147483633
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
            Appearance      =   17
         End
      End
      Begin XtremeSuiteControls.PushButton btnFoto 
         Height          =   372
         Index           =   0
         Left            =   -65320
         TabIndex        =   55
         Top             =   3360
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Subir"
         BackColor       =   -2147483633
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
         Appearance      =   17
      End
      Begin VB.PictureBox picFoto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3012
         Left            =   -65320
         ScaleHeight     =   2985
         ScaleWidth      =   3465
         TabIndex        =   54
         Top             =   240
         Visible         =   0   'False
         Width           =   3492
      End
      Begin XtremeSuiteControls.CheckBox chkInd_Licencia 
         Height          =   375
         Left            =   -69520
         TabIndex        =   50
         Top             =   1800
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441793
         _ExtentX        =   10181
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Posee Licencia de Conducir?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboSexo 
         Height          =   315
         Left            =   -63880
         TabIndex        =   14
         Top             =   1200
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
         Height          =   315
         Left            =   -68320
         TabIndex        =   15
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpNacimiento 
         Height          =   315
         Left            =   -63880
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.GroupBox gbPersona 
         Height          =   4335
         Index           =   1
         Left            =   -69760
         TabIndex        =   20
         Top             =   2520
         Visible         =   0   'False
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   7646
         _StockProps     =   79
         Caption         =   "Datos de Localización"
         ForeColor       =   8421504
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   312
            Left            =   1440
            TabIndex        =   21
            Top             =   3120
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2990
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   312
            Left            =   3240
            TabIndex        =   22
            Top             =   3120
            Width           =   1932
            _Version        =   1441793
            _ExtentX        =   3413
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   312
            Left            =   5280
            TabIndex        =   23
            Top             =   3120
            Width           =   2052
            _Version        =   1441793
            _ExtentX        =   3625
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail 
            Height          =   312
            Left            =   1440
            TabIndex        =   24
            Top             =   1320
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmail_02 
            Height          =   312
            Left            =   1440
            TabIndex        =   25
            Top             =   1680
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtApartado 
            Height          =   312
            Left            =   1440
            TabIndex        =   26
            Top             =   2520
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDireccion 
            Height          =   795
            Left            =   1440
            TabIndex        =   27
            Top             =   3480
            Width           =   5895
            _Version        =   1441793
            _ExtentX        =   10398
            _ExtentY        =   1402
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtWebSite 
            Height          =   312
            Left            =   1440
            TabIndex        =   58
            Top             =   2160
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelMovil 
            Height          =   312
            Left            =   1440
            TabIndex        =   30
            Top             =   480
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono1 
            Height          =   312
            Left            =   1440
            TabIndex        =   28
            Top             =   840
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelefono2 
            Height          =   312
            Left            =   5640
            TabIndex        =   29
            Top             =   480
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTelFax 
            Height          =   312
            Left            =   5640
            TabIndex        =   31
            Top             =   840
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Apto. Postal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   0
            TabIndex        =   59
            Top             =   2520
            Width           =   1332
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   0
            TabIndex        =   39
            Top             =   3120
            Width           =   732
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Email No.2"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   9
            Left            =   0
            TabIndex        =   38
            Top             =   1680
            Width           =   1092
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Site / Blog"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   0
            TabIndex        =   37
            Top             =   2160
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Email No.1"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   0
            TabIndex        =   36
            Top             =   1320
            Width           =   1092
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono Hab."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   0
            TabIndex        =   35
            Top             =   840
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono Trab."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   10
            Left            =   4200
            TabIndex        =   34
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Tel. Móvil"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   11
            Left            =   0
            TabIndex        =   33
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono/Fax"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   13
            Left            =   4200
            TabIndex        =   32
            Top             =   840
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.PushButton btnEditarDetalle 
         Height          =   345
         Left            =   6840
         TabIndex        =   40
         Top             =   0
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Editar"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmRH_Empleados.frx":2776
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.ComboBox cboNacionalidad 
         Height          =   315
         Left            =   -68320
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkInd_Vehiculo 
         Height          =   375
         Left            =   -69520
         TabIndex        =   51
         Top             =   1440
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441793
         _ExtentX        =   10181
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Posee Vehículo Propio?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkInd_Solidarista 
         Height          =   375
         Left            =   -69520
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   3375
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Es Asociado de la Asociación Solidarista?"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkInd_Marcas 
         Height          =   375
         Left            =   -69520
         TabIndex        =   53
         Top             =   720
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441793
         _ExtentX        =   10181
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Utiliza Control de Marcas?"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnFoto 
         Height          =   372
         Index           =   1
         Left            =   -63400
         TabIndex        =   56
         Top             =   3360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Eliminar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   -68200
         TabIndex        =   61
         Top             =   4320
         Visible         =   0   'False
         Width           =   6252
         _Version        =   1441793
         _ExtentX        =   11028
         _ExtentY        =   1397
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtId_Tributario 
         Height          =   312
         Left            =   -64120
         TabIndex        =   64
         Top             =   3960
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
         _ExtentY        =   556
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
      Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
         Height          =   252
         Index           =   3
         Left            =   -62680
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtProfesionCod 
         Height          =   312
         Left            =   -68320
         TabIndex        =   66
         Top             =   120
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProfesionDesc 
         Height          =   312
         Left            =   -67600
         TabIndex        =   67
         Top             =   120
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.ComboBox cboNivel 
         Height          =   312
         Left            =   -68320
         TabIndex        =   68
         Top             =   600
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1441793
         _ExtentX        =   9763
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido1 
         Height          =   315
         Left            =   -69760
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtApellido2 
         Height          =   315
         Left            =   -67480
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4043
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   315
         Left            =   -65200
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4890
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin MSComCtl2.FlatScrollBar FlatScroll_Laboral 
         Height          =   252
         Index           =   5
         Left            =   -62680
         TabIndex        =   71
         Top             =   1080
         Visible         =   0   'False
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtJefeCod 
         Height          =   312
         Left            =   -68320
         TabIndex        =   72
         Top             =   1080
         Visible         =   0   'False
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtJefeDesc 
         Height          =   312
         Left            =   -67600
         TabIndex        =   73
         Top             =   1080
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.CheckBox chkCF_Conyuge 
         Height          =   375
         Left            =   -69520
         TabIndex        =   75
         Top             =   2400
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441793
         _ExtentX        =   10181
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplica Crédito Fiscal por Cónyuge?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtCF_Dependientes 
         Height          =   315
         Left            =   -69520
         TabIndex        =   76
         Top             =   2880
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   868
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
         Text            =   "0"
         Alignment       =   2
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   345
         Left            =   8040
         TabIndex        =   140
         ToolTipText     =   "Exportar"
         Top             =   0
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   661
         _ExtentY        =   609
         _StockProps     =   79
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmRH_Empleados.frx":2D71
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.Label Label13 
         Height          =   255
         Left            =   -68680
         TabIndex        =   78
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Número de Dependientes?"
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
      End
      Begin VB.Label Label21 
         Caption         =   "Jefe/Superior"
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
         Left            =   -69640
         TabIndex        =   74
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Académico"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   4
         Left            =   -69640
         TabIndex        =   70
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Profesión"
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
         Left            =   -69640
         TabIndex        =   69
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación para efectos tributarios"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   -67720
         TabIndex        =   63
         Top             =   3960
         Visible         =   0   'False
         Width           =   3372
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69400
         TabIndex        =   62
         Top             =   4320
         Visible         =   0   'False
         Width           =   1212
      End
      Begin XtremeShortcutBar.ShortcutCaption TituloOpciones 
         Height          =   360
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   12612
         _Version        =   1441793
         _ExtentX        =   22246
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Detalles:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.1:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -69760
         TabIndex        =   48
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido No.2:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -67480
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -65200
         TabIndex        =   46
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nacimiento"
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
         Left            =   -64960
         TabIndex        =   45
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Genero"
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
         Left            =   -64960
         TabIndex        =   44
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Civil"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -69760
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -69760
         TabIndex        =   42
         Top             =   1200
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   330
      Left            =   4920
      TabIndex        =   2
      Top             =   720
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEmpleadoId 
      Height          =   330
      Left            =   2760
      TabIndex        =   3
      Top             =   720
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   330
      Left            =   7080
      TabIndex        =   11
      Top             =   720
      Width           =   2652
      _Version        =   1441793
      _ExtentX        =   4678
      _ExtentY        =   582
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   2640
      TabIndex        =   134
      ToolTipText     =   "Nuevo"
      Top             =   0
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":2EDB
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   3720
      TabIndex        =   135
      ToolTipText     =   "Editar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":350D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   4080
      TabIndex        =   136
      ToolTipText     =   "Eliminar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":3B08
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   4680
      TabIndex        =   137
      ToolTipText     =   "Guardar"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":40AC
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   5040
      TabIndex        =   138
      ToolTipText     =   "Deshacer"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":47DD
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   5520
      TabIndex        =   139
      ToolTipText     =   "Reporte"
      Top             =   0
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   656
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmRH_Empleados.frx":4EDD
      ImageAlignment  =   6
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7200
      TabIndex        =   10
      Top             =   480
      Width           =   732
   End
   Begin XtremeShortcutBar.ShortcutCaption TituloOpcion 
      Height          =   360
      Left            =   2760
      TabIndex        =   9
      Top             =   1080
      Width           =   8535
      _Version        =   1441793
      _ExtentX        =   15055
      _ExtentY        =   635
      _StockProps     =   14
      Caption         =   "Datos de Contacto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1932
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4920
      TabIndex        =   5
      Top             =   480
      Width           =   1692
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Id. Empleado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   2760
      TabIndex        =   4
      Top             =   480
      Width           =   1692
   End
End
Attribute VB_Name = "frmRH_Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vEditar As Boolean, vFechaActual As Date
Dim vCedula As String, vSeek As Integer, vScroll As Boolean, vTipoJuridica As Integer

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Const Id_TaskItem_DatosPersonales = 0
Const Id_TaskItem_RelacionLaboral = 1
Const Id_TaskItem_Otros_Add = 2

Const Id_TaskItem_Telefonos = 3
Const Id_TaskItem_Familiares = 4
Const Id_TaskItem_Cuentas = 5

Const Id_TaskItem_Boletas_Pago = 6
Const Id_TaskItem_Plan_Carrera = 7
Const Id_TaskItem_Vacaciones = 8
Const Id_TaskItem_Permisos = 9
Const Id_TaskItem_Incapacidades = 10

Const Id_TaskItem_Tarjetas = 11

Const Id_TaskItem_Vacaciones_H = 12
Const Id_TaskItem_Conceptos = 13


Const Id_TaskItem_Accion_Personal = 17



Public Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub

Private Sub btnBarra_Click(Index As Integer)

Select Case Index
  Case 0 'Nuevo
    vEditar = False
    Call sbBarra_Accion("Editar")
    Call sbClearControles
    Call sbLockControles("U")
    Call sbLimpiaDatos
    
    txtEmpleadoId.Text = fxgRH_Empleado_ID
    txtEmpleadoId.Locked = True
    txtIdentificacion.SetFocus
    
  Case 1 'Editar
    If Trim(txtIdentificacion) <> vCedula Then
     MsgBox "Vuelva a Consultar la Persona!", vbExclamation
     Exit Sub
    End If
    
    vEditar = True
    vCedula = Trim(txtIdentificacion)
    txtEmpleadoId.Locked = False
    
    
    Call sbBarra_Accion("Editar")
    Call sbLockControles("U")
    
    Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)

    If txtApellido1.Enabled Then txtApellido1.SetFocus
        
  Case 2 'Borrar
    Call sbDeleteRecord
        
  Case 3 'Guardar
    Call sbSaveRecord
    
  Case 4 'Deshacer
    vEditar = False
    Call sbBarra_Accion("nuevo")
    Call RefrescaTags(Me)
    Call sbClearControles
    Call sbLockControles("L")
    Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)
    
    txtIdentificacion.SetFocus
    
  Case 5 'Reportes
    GLOBALES.gTag = txtEmpleadoId.Text
    GLOBALES.gTag2 = txtIdentificacion.Text
    GLOBALES.gTag3 = txtApellido1.Text + " " + txtApellido2.Text + " " + txtNombre.Text
    
    frmRH_InformesConstancias.Show vbModal
    
  Case 6 'consultar
    Select Case vSeek
      Case 1, 2
       gBusquedas.Resultado = Trim(txtIdentificacion)
       txtIdentificacion = ""
       vCedula = ""
       gBusquedas.Convertir = "N"
       
       If vSeek = 1 Then
        gBusquedas.Columna = "Identificacion"
        gBusquedas.Orden = "Identificacion"
       Else
        gBusquedas.Columna = "Nombre_Completo"
        gBusquedas.Orden = "Nombre_Completo"
       End If
       
       gBusquedas.Consulta = "Select Identificacion,Empleado_Id,Nombre_Completo From Rh_Personas"
   
       frmBusquedas.Show vbModal
   
       txtIdentificacion = Trim(gBusquedas.Resultado)
       txtIdentificacion_LostFocus
       
      Case 3
       If cboProvincia.Text = "" Then Exit Sub
       gBusquedas.Resultado = ""
       gBusquedas.Resultado2 = ""
   
       gBusquedas.Convertir = "N"
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "Descripcion"
       gBusquedas.Consulta = "Select Canton,Descripcion From Cantones"
       gBusquedas.Filtro = "And Provincia =" & cboProvincia.ItemData(cboProvincia.ListIndex)
   
       frmBusquedas.Show vbModal
   
'       txtCodigoCanton = Trim(gBusquedas.Resultado)
'       txtCanton = Trim(gBusquedas.Resultado2)

    End Select
    
End Select





End Sub


Private Sub sbTaskPanel_Load()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    tpMain.VisualTheme = xtpTaskPanelThemeVisualStudio2012Light
    
  
    Set Group = tpMain.Groups.Add(0, "Registro")
    Group.ToolTip = "Información Principal para el Registro de la Persona"
    Group.Special = True

    
    Group.Items.Add Id_TaskItem_DatosPersonales, "Datos Personales", xtpTaskItemTypeLink, 4
    Group.Items.Add Id_TaskItem_RelacionLaboral, "Relación Laboral", xtpTaskItemTypeLink, 1
    Group.Items.Add Id_TaskItem_Conceptos, "Conceptos", xtpTaskItemTypeLink, 1
    Group.Items.Add Id_TaskItem_Otros_Add, "Adicionales y Portal", xtpTaskItemTypeLink, 10
    
    Set Group = tpMain.Groups.Add(0, "Detalles")
    Group.ToolTip = "Datos Complementarios"
    
    Group.Items.Add Id_TaskItem_Familiares, "Familiares & Contactos", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Cuentas, "Cuentas Bancarias", xtpTaskItemTypeLink, 9
    Group.Items.Add Id_TaskItem_Tarjetas, "Tarjetas", xtpTaskItemTypeLink, 9
   
    
    Set Group = tpMain.Groups.Add(0, "Histórico")
    Group.Expanded = False
    Group.Items.Add Id_TaskItem_Boletas_Pago, "Boletas de Pago", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Vacaciones, "Vacaciones - Boletas", xtpTaskItemTypeLink, 3
    
    Group.Items.Add Id_TaskItem_Vacaciones_H, "Vacaciones - Historico", xtpTaskItemTypeLink, 3
    
    Group.Items.Add Id_TaskItem_Incapacidades, "Incapacidades", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Permisos, "Permisos", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Accion_Personal, "Acciones de Personal", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Plan_Carrera, "Plan de Carrera", xtpTaskItemTypeLink, 3
    
   
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
    

End Sub


Private Sub sbPersona_Foto_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from RH_Personas where Empleado_Id = '" & txtEmpleadoId.Text & "'"

Set picFoto.Picture = fxImagen_Leer(strSQL, "FOTO")

picFoto.PaintPicture picFoto.Picture, 0, 0, picFoto.ScaleWidth, picFoto.ScaleHeight

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault

End Sub

Private Sub sbTaskPanel_Accion(ItemId As Integer)

Dim fraX As Frame

If Trim(txtIdentificacion.Text) = "" Then Exit Sub

On Error GoTo vError


Select Case ItemId
  Case Id_TaskItem_DatosPersonales  'Datos de Contato
    

    TituloOpcion.Caption = "Datos de Contacto"
    
    tcMain.Item(0).Selected = True
    Call cboTipoId_Click
    
'    If fraTipo.Visible Then
'       txtNombreComercial.SetFocus
'    Else
'       If txtApellido1.Enabled Then txtApellido1.SetFocus
'    End If
    
    Exit Sub
    
  Case Id_TaskItem_RelacionLaboral  'Relación Laboral
    TituloOpcion.Caption = "Datos Laborales"
    
    tcMain.Item(1).Selected = True
    txtCentroCod.SetFocus

    Exit Sub

  Case Id_TaskItem_Otros_Add 'Información Adicional
    
    TituloOpcion.Caption = "Adicionales y Portal..."
    tcMain.Item(2).Selected = True
    
    DoEvents
    
    Call sbPersona_Foto_Load
      
    Exit Sub
      
End Select

If Not vEditar Then
    MsgBox "Se encuentra en modo de Registro, guarde los datos de la persona y luego ingrese a esta opción!", vbInformation
    Exit Sub
End If

'Otras Opciones
TituloOpcion.Caption = "Otros datos:"



tcMain.Item(3).Selected = True

lswHistorico.ColumnHeaders.Clear
lswHistorico.ListItems.Clear
lswHistorico.Checkboxes = False

btnEditarDetalle.Visible = False


Select Case ItemId
  
  Case Id_TaskItem_Conceptos  'Conceptos de Nómina
        
    TituloOpciones.Caption = "Conceptos de Nomina..:"
    TituloOpciones.Tag = "Conceptos"
    
    btnEditarDetalle.Visible = True
        
    lswHistorico.ColumnHeaders.Add , , "Código", 1500
    lswHistorico.ColumnHeaders.Add , , "Descripción", 3500
    lswHistorico.ColumnHeaders.Add , , "Tipo", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Valor", 1500, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "Inicia", 2100, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Vence", 2100, vbCenter
    
    lswHistorico.ColumnHeaders.Add , , "Aplicado", 1800, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "No.Documento", 1500
    lswHistorico.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "R.Fecha", 2500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "R.Usuario", 2500, vbCenter
    
    
    strSQL = "exec spRH_Persona_Conceptos_List '" & Trim(txtEmpleadoId.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , Trim(rs!Cod_Concepto))
           itmX.SubItems(1) = rs!Descripcion
           itmX.SubItems(2) = Trim(rs!Tipo) & ""
           itmX.SubItems(3) = Format(rs!Valor, "Standard")
           itmX.SubItems(4) = rs!Fecha_Inicio
           itmX.SubItems(5) = rs!Fecha_Vence
           
           itmX.SubItems(6) = Format(rs!Monto_Acumulado, "Standard")
           itmX.SubItems(7) = rs!Documento & ""
           itmX.SubItems(8) = rs!Estado & ""
           
           itmX.SubItems(9) = rs!Registro_Fecha & ""
           itmX.SubItems(10) = rs!Registro_Usuario & ""
       
       rs.MoveNext
    Loop
    rs.Close
  
  Case Id_TaskItem_Telefonos  'Telefonos
        
    TituloOpciones.Caption = "Lista de Teléfonos..:"
    TituloOpciones.Tag = "Telefonos"
    
    btnEditarDetalle.Visible = True
        
    lswHistorico.ColumnHeaders.Add 1, , "Numero", 1500
    lswHistorico.ColumnHeaders.Add 2, , "Tipo", 1500
    lswHistorico.ColumnHeaders.Add 3, , "Extension", 1500
    lswHistorico.ColumnHeaders.Add 4, , "Contacto", 2500
    
    
    strSQL = "Select * From Telefonos where Cedula='" & Trim(txtIdentificacion) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , Trim(rs!Numero))
           itmX.SubItems(1) = (rs!Tipo)
           itmX.SubItems(2) = Trim(rs!Ext) & ""
           itmX.SubItems(3) = Trim(rs!contacto) & ""
       rs.MoveNext
    Loop
    rs.Close
    
    
  Case Id_TaskItem_Familiares 'Familiares
    btnEditarDetalle.Visible = True
 
     
    TituloOpciones.Caption = "Lista de Familiares..:"
    TituloOpciones.Tag = "Familiares"
    
    lswHistorico.ColumnHeaders.Add , , "Identificación", 1500
    lswHistorico.ColumnHeaders.Add , , "Nombre", 3500
    lswHistorico.ColumnHeaders.Add , , "Parentesco", 1100, vbCenter
    

    
    strSQL = "select Identificacion, Nombre, Parentesco_Desc from vRH_Personas_Familiares where Empleado_Id = '" & Trim(txtEmpleadoId.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!IDENTIFICACION)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!parentesco_Desc)
       
       rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Cuentas 'Cuentas Bancarias
    
    btnEditarDetalle.Visible = True
    
    
    TituloOpciones.Caption = "Cuentas bancarias..:"
    TituloOpciones.Tag = "Cuentas"
    
    lswHistorico.ColumnHeaders.Add 1, , "Cuenta", 2500
    lswHistorico.ColumnHeaders.Add 2, , "Banco", 3500
    lswHistorico.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
    lswHistorico.ColumnHeaders.Add 8, , "Fecha", 2500
    lswHistorico.ColumnHeaders.Add 9, , "Usuario", 2500

        strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & Trim(txtIdentificacion) & "'" 'and C.Modulo = 'AFI'
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!Activa = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
       rs.MoveNext
    Loop
    rs.Close
    
    
    
  
  Case Id_TaskItem_Boletas_Pago 'Boletas de Pago
  
    
    TituloOpciones.Caption = "Boletas de Pago..:"
    TituloOpciones.Tag = "BoletaPago"
    
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add , , "No. Nómina", 1200
    lswHistorico.ColumnHeaders.Add , , "No. Pago", 1000, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Nomina", 1000, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Inicio", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Corte", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Salario", 1500, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "Ingresos", 1500, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "Egresos", 1500, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "A Pagar", 1500, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "Descripción", 1500, vbRightJustify
            
    
    strSQL = "select Top 50 * from vRH_Boleta_Pago_List Where Empleado_Id = '" _
           & Trim(txtEmpleadoId.Text) & "' order by Fecha_Corte desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswHistorico.ListItems.Add(, , rs!Nomina_Num)
          itmX.SubItems(1) = rs!NPago_Mes
          itmX.SubItems(2) = rs!COD_NOMINA
          itmX.Tag = rs!COD_NOMINA
          
          itmX.SubItems(3) = Format(rs!Fecha_Inicio, "yyyy-mm-dd")
          itmX.SubItems(4) = Format(rs!Fecha_Corte, "yyyy-mm-dd")
          itmX.SubItems(5) = Format(rs!SALARIO_ORDINARIO, "Standard")
          itmX.SubItems(6) = Format(rs!Ingresos, "Standard")
          itmX.SubItems(7) = Format(rs!Egresos, "Standard")
          itmX.SubItems(8) = Format(rs!Salario_Neto, "Standard")
          itmX.SubItems(9) = rs!Nomina_Desc
          
      rs.MoveNext
    Loop
    rs.Close
  
  
  Case Id_TaskItem_Accion_Personal 'Accion
     
    TituloOpciones.Caption = "Acciones de Personal..:"
    TituloOpciones.Tag = "AccionPersonal"

    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Add , , "No. Boleta", 1500
    lswHistorico.ColumnHeaders.Add , , "Fecha", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Tipo", 2000
    lswHistorico.ColumnHeaders.Add , , "Salario", 1600, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "Salario Ant.", 1600, vbRightJustify
    lswHistorico.ColumnHeaders.Add , , "Puesto", 2100
    lswHistorico.ColumnHeaders.Add , , "Puesto Ant.", 2100
    lswHistorico.ColumnHeaders.Add , , "Centro", 2100
    lswHistorico.ColumnHeaders.Add , , "Centro Ant.", 2100
    lswHistorico.ColumnHeaders.Add , , "Departamento", 2100
    lswHistorico.ColumnHeaders.Add , , "Dept. Ant.", 2100
    lswHistorico.ColumnHeaders.Add , , "Sección", 2100
    lswHistorico.ColumnHeaders.Add , , "Sección Ant.", 2100
    
    lswHistorico.ColumnHeaders.Add , , "Nómina", 2100
    lswHistorico.ColumnHeaders.Add , , "Nómina Ant.", 2100
    
    lswHistorico.ColumnHeaders.Add , , "Estado", 2100
    lswHistorico.ColumnHeaders.Add , , "Estado Ant.", 2100
    
    lswHistorico.ColumnHeaders.Add , , "Notas", 2100
    
    lswHistorico.ColumnHeaders.Add , , "Fecha", 2500
    lswHistorico.ColumnHeaders.Add , , "Usuario", 2500


    strSQL = "select * From vRH_Accion_Personal Where Empleado_id = '" & Trim(txtEmpleadoId.Text) & "' order by cod_Accion desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!Cod_Accion)
           itmX.SubItems(1) = Format(rs!Fecha_Accion, "dd/MM/yyyy")
           itmX.SubItems(2) = rs!TipoAccionDesc
           itmX.SubItems(3) = Format(rs!Salario_Actual, "Standard")
           itmX.SubItems(4) = Format(rs!ANT_Salario, "Standard")
           itmX.SubItems(5) = rs!PuestoDesc
           itmX.SubItems(6) = rs!A_PuestoDesc
           itmX.SubItems(7) = rs!CentroDesc
           itmX.SubItems(8) = rs!A_CentroDesc
           itmX.SubItems(9) = rs!DepartamentoDesc
           itmX.SubItems(10) = rs!A_DepartamentoDesc
           itmX.SubItems(11) = rs!SeccionDesc
           itmX.SubItems(12) = rs!A_SeccionDesc
           itmX.SubItems(13) = rs!NominaDesc
           itmX.SubItems(14) = rs!NominaDesc
           itmX.SubItems(15) = rs!EstadoPersonaDesc
           itmX.SubItems(16) = rs!A_EstadoPersonaDesc
           itmX.SubItems(17) = rs!notas & ""
           itmX.SubItems(18) = rs!Registro_Fecha & ""
           itmX.SubItems(19) = Trim(rs!Registro_Usuario & "")

       rs.MoveNext
    Loop
    rs.Close

  
  Case Id_TaskItem_Plan_Carrera 'Plan de Carrera
  
    With lswHistorico
            .ListItems.Clear
            .ColumnHeaders.Clear
            .ColumnHeaders.Add , , "Id", 900
            .ColumnHeaders.Add , , "Nivel", 2100
            .ColumnHeaders.Add , , "Curso", 2100
            .ColumnHeaders.Add , , "Estado", 1500, vbCenter
            .ColumnHeaders.Add , , "Nota", 1500, vbCenter
            .ColumnHeaders.Add , , "Usuario", 1500
            .ColumnHeaders.Add , , "Fecha", 1500
            
            TituloOpciones.Caption = "Plan de Carrera..:"
            TituloOpciones.Tag = "PlanCarrera"
            
'            strSQL = "Select I.*,P.nombre as Promotor " _
'                   & " From Afi_Ingresos I left join promotores P on I.id_promotor = P.id_promotor" _
'                   & " where I.Cedula='" & Trim(txtIdentificacion) & "'"
'            Call OpenRecordSet(rs, strSQL)
'            Do While Not rs.EOF
'               Set itmX = .ListItems.Add(, , rs!consec)
'                   itmX.SubItems(1) = rs!Usuario & ""
'                   itmX.SubItems(2) = rs!fecha & ""
'                   itmX.SubItems(3) = Format(rs!Fecha_Ingreso)
'                   itmX.SubItems(4) = rs!Boleta & ""
'                   itmX.SubItems(5) = rs!promotor & ""
'               rs.MoveNext
'            Loop
'            rs.Close
    End With
  
  
  Case Id_TaskItem_Vacaciones  'Vacaciones
    
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add , , "Boleta", 1000
    lswHistorico.ColumnHeaders.Add , , "Motivo", 2100
    lswHistorico.ColumnHeaders.Add , , "Inicio", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Corte", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Días", 900, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Usuario", 1500
    lswHistorico.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Vacaciones..:"
    TituloOpciones.Tag = "Vacaciones"
    
    strSQL = "Select * from vRH_Boleta_Vacaciones" _
           & " where Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by Boleta_VAC desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!Boleta_Id)
           itmX.SubItems(1) = Trim(rs!Motivo)
           itmX.SubItems(2) = Format(rs!Fecha_Salida, "dd/mm/yyyy")
           itmX.SubItems(3) = Format(rs!Fecha_Entrada, "dd/mm/yyyy")
           itmX.SubItems(4) = rs!Dias_Disfrutados & ""
           itmX.SubItems(5) = rs!Estado_Transaccion
           itmX.SubItems(6) = rs!Registro_Usuario & ""
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lswHistorico.Enabled = True
  
  
  
  Case Id_TaskItem_Vacaciones_H   'Vacaciones - Histórico
    
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add , , "Corte", 1000
    lswHistorico.ColumnHeaders.Add , , "Anterior", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "(+) Periodo", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "(-) Disfrutadas", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Actual", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Regimen", 2500
    lswHistorico.ColumnHeaders.Add , , "Usuario", 1500
    lswHistorico.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Vacaciones Histórico..:"
    TituloOpciones.Tag = "VacacionesH"
    
    strSQL = "select P.*, Vr.DESCRIPCION" _
           & " from RH_PERSONAS_VAC_ACUM P left join RH_VACACIONES_REGIMEN Vr on P.COD_VACA_REGIMEN = Vr.COD_VACA_REGIMEN " _
           & " where P.Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by P.Corte desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , Format(rs!Corte, "yyyy-mm-dd"))
           itmX.SubItems(1) = rs!VA_Anterior
           itmX.SubItems(2) = rs!Va_Periodo
           itmX.SubItems(3) = rs!Va_Disfrutadas
           itmX.SubItems(4) = rs!Va_Actual
           itmX.SubItems(5) = rs!Descripcion
           itmX.SubItems(6) = rs!Registro_Usuario & ""
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lswHistorico.Enabled = True
  
  Case Id_TaskItem_Incapacidades 'Incapacidades
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add , , "Boleta", 1000
    lswHistorico.ColumnHeaders.Add , , "Motivo", 2100
    lswHistorico.ColumnHeaders.Add , , "Inicio", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Corte", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Días", 900, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Usuario", 1500
    lswHistorico.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Incapacidades..:"
    TituloOpciones.Tag = "Incapacidades"
    
    strSQL = "Select * from vRH_Boleta_Incapacidades" _
           & " where Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by Boleta_ID desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!Boleta_Id)
           itmX.SubItems(1) = Trim(rs!Motivo)
           itmX.SubItems(2) = Format(rs!Fecha_Salida, "dd/mm/yyyy")
           itmX.SubItems(3) = Format(rs!Fecha_Entrada, "dd/mm/yyyy")
           itmX.SubItems(4) = rs!Dias & ""
           itmX.SubItems(5) = rs!Estado_Transaccion
           itmX.SubItems(6) = rs!Registro_Usuario & ""
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lswHistorico.Enabled = True
  
  Case Id_TaskItem_Permisos 'Permisos
  
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add , , "Boleta", 1000
    lswHistorico.ColumnHeaders.Add , , "Motivo", 2100
    lswHistorico.ColumnHeaders.Add , , "Fecha/Permiso", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Hr. Inicio", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Hr. Corte", 1800, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Horas", 900, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Estado", 1500, vbCenter
    lswHistorico.ColumnHeaders.Add , , "Usuario", 1500
    lswHistorico.ColumnHeaders.Add , , "Fecha", 1500
    
    
    TituloOpciones.Caption = "Permisos..:"
    TituloOpciones.Tag = "Permisos"
    
    strSQL = "Select * from vRH_Boleta_Permisos" _
           & " where Empleado_Id = '" & Trim(txtEmpleadoId.Text) _
           & "' order by Boleta_ID desc"

    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswHistorico.ListItems.Add(, , rs!Boleta_Id)
           itmX.SubItems(1) = Trim(rs!Motivo)
           itmX.SubItems(2) = Format(rs!Hora_Inicio, "dd/mm/yyyy")
           itmX.SubItems(3) = Format(rs!Hora_Inicio, "hh:mm:ss")
           itmX.SubItems(4) = Format(rs!Hora_Corte, "hh:mm:ss")
           itmX.SubItems(5) = rs!Hrs_Total & ""
           itmX.SubItems(6) = rs!Estado_Transaccion
           itmX.SubItems(7) = rs!Registro_Usuario & ""
           itmX.SubItems(8) = rs!Registro_Fecha & ""
           
       rs.MoveNext
    Loop
    rs.Close
    
    lswHistorico.Enabled = True
    
    
  Case Id_TaskItem_Tarjetas  'Tarjetas
  
    btnEditarDetalle.Visible = True
    
    lswHistorico.ListItems.Clear
    lswHistorico.ColumnHeaders.Clear
    lswHistorico.ColumnHeaders.Add 1, , "No. Tarjeta", 2500
    lswHistorico.ColumnHeaders.Add 2, , "Tipo", 2100
    lswHistorico.ColumnHeaders.Add 3, , "Vence", 1500
    
    TituloOpciones.Caption = "Tarjetas..:"
    TituloOpciones.Tag = "Tarjetas"
  
  
            strSQL = "exec spAFI_PersonaTarjetas_Consulta " & gPortal.Empresa_Id & ",'" & txtIdentificacion.Text & "',''"
            Call OpenRecordSet(rs, strSQL)
            
            With lswHistorico.ListItems
               .Clear
               Do While Not rs.EOF
                Set itmX = .Add(, , rs!Tarjeta_Mask)
                    itmX.SubItems(1) = rs!Tarjeta_Tipo
                    itmX.SubItems(2) = Format(rs!Tarjeta_Vence, "MM/YY")
                rs.MoveNext
               Loop
               rs.Close
            End With



End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbClearControles()
Dim vControl As Control

For Each vControl In Me
  If TypeOf vControl Is TextBox Then
     vControl.Text = ""
  End If
Next

StatusBarX.Panels.Item(1) = ""
StatusBarX.Panels.Item(2) = ""
StatusBarX.Panels.Item(3) = ""
StatusBarX.Panels.Item(4) = ""


End Sub

Private Sub sbCurrentRecord(pId As String, Optional pTipo As String = "P")
Dim i As Integer, vEspacio As Integer

If pId = "" Then Exit Sub

On Error Resume Next

If Not fxSIFValidaCadena(pId) Then
   Exit Sub
End If

'Convierte Identificacion de Persona en Empleado Id
If pTipo = "P" Then
   pId = fxgRH_Persona_Empleado_ID(pId)
End If

If Not vEditar And pId <> txtEmpleadoId.Text Then
    If Not fxgRH_Empleado_Status(pId) Then
        pId = txtEmpleadoId.Text
    End If
End If

tcMain.Item(0).Selected = True

gbAccionPersonal.Enabled = fxgRH_Persona_Accion_Inicial(pId)

strSQL = "select * from vRH_Personas where EMPLEADO_ID = '" & pId & "'"

Call OpenRecordSet(rs, strSQL)
If (Not rs.EOF And Not rs.BOF) And Not glogon.error Then
      
   
   vEditar = True
   Call sbBarra_Accion("activo")
   Call RefrescaTags(Me)
   
   Call sbLockControles("L")
   Call sbLimpiaDatos 'Inicializa Datos
   
   vCedula = Trim(rs!IDENTIFICACION)
   
   txtEmpleadoId.Text = Trim(rs!Empleado_ID)
   txtIdentificacion.Text = Trim(rs!IDENTIFICACION)
   
   If Not IsNull(rs!TipoIdDesc) Then
       vPaso = True
           cboTipoId.Text = Trim(rs!TipoIdDesc)
       vPaso = False
   End If
   
   If Not IsNull(rs!Nacionalidad) Then
       vPaso = True
       Call sbCboAsignaDato(cboNacionalidad, Trim(rs!Nacionalidad), True, rs!Cod_Nacionalidad)
       vPaso = False
   End If
   
     
   txtApellido1.Text = Trim(rs!Apellido_1)
   txtApellido2.Text = Trim(rs!Apellido_2)
   txtNombre.Text = Trim(rs!Nombre)
   
   txtEstado.Text = Trim(rs!EstadoPersonaDesc)
  
   dtpNacimiento.Value = rs!fecha_nacimiento
   
   cboSexo.Text = IIf(rs!sexo = "M", "Masculino", "Femenino")
   
   Call sbCboAsignaDato(cboEstadoCivil, rs!EstadoCivilDesc & "", True, rs!EstadoCivil)  'Se activa el Click ->    Call cboProvincia_Click
     
   'Contacto
   
   txtTelefono1.Text = rs!telefono1 & ""
   txtTelefono2.Text = rs!telefono2 & ""
   txtTelMovil.Text = rs!Tel_Movil & ""
   txtTelFax.Text = rs!Fax & ""
   
   txtEmail.Text = Trim(rs!Email_01 & "")
   txtEmail_02.Text = Trim(rs!Email_02 & "")
   txtApartado.Text = Trim(rs!apto_postal & "")
   
   txtWebSite.Text = Trim(rs!WebSite & "")
     
   Call sbCboAsignaDato(cboProvincia, rs!ProvinciaDesc & "")  'Se activa el Click ->    Call cboProvincia_Click
   Call sbCboAsignaDato(cboCanton, rs!CantonDesc & "")   'Se activa el Click
   Call sbCboAsignaDato(cboDistrito, rs!DistritoDesc & "")
     
   txtDireccion = Trim(rs!Direccion) & ""
       
       
   'Laboral
   
   txtCentroCod.Text = rs!Cod_Centro
   txtCentroDesc.Text = rs!CentroDesc
  
   txtDeptCodigo.Text = Trim(rs!Cod_Departamento & "")
   txtDeptDesc.Text = Trim(rs!DepartamentoDesc & "")
   
   txtSecCodigo.Text = Trim(rs!Cod_Seccion & "")
   txtSecDesc.Text = Trim(rs!SeccionDesc & "")
   
   txtProfesionCod.Text = rs!Cod_Profesion
   txtProfesionDesc.Text = rs!ProfesionDesc
   Call sbCboAsignaDato(cboNivel, rs!NivelAcademicoDesc, True, rs!Nivel_Academico)
   
   txtPuestoCod.Text = rs!Cod_Puesto
   txtPuestoDesc.Text = rs!PuestoDesc
   
   txtJefeCod.Text = rs!Jefe_id & ""
   txtJefeDesc.Text = rs!JefeDesc
   
   
   dtpIngreso.Value = rs!FECHA_INGRESO
   dtpIngreso.Enabled = False
   
   txtSalario.Text = Format(rs!SALARIO_ORDINARIO, "Standard")
   
   txtVacaAcum.Text = rs!VACA_ACUMULADAS & ""
   
   Call sbCboAsignaDato(cboDivisa, rs!DivisaDesc, True, rs!cod_Divisa)
   
   Call sbCboAsignaDato(cboNomina, rs!NominaDesc, True, rs!COD_NOMINA)
   Call sbCboAsignaDato(cboContrato, rs!ContratoDesc, True, rs!Contrato_Tipo)
   Call sbCboAsignaDato(cboJornada, rs!JornadaDesc, True, rs!Jornada_Tipo)
   Call sbCboAsignaDato(cboVacaciones, rs!VacacionesDesc, True, rs!Cod_Vaca_Regimen)
   
   Call sbCboAsignaDato(cboBancos, rs!BancoDesc, True, rs!cod_banco)
   
   If rs!Forma_Pago = "TE" Then
    cboFormaPago.Text = "Transferencia"
   Else
    cboFormaPago.Text = "Cheque"
   
   End If
   
   If Not IsNull(rs!Contrato_Vencimiento) Then
        dtpContrato.Value = rs!Contrato_Vencimiento
        dtpContrato.Visible = True
   Else
        dtpContrato.Visible = False
   End If
   lblVencimiento.Visible = dtpContrato.Visible
   
   
   'Salida
   
   txtSalidaEstado.Tag = rs!Salida_Estado
   txtSalidaEstado.Text = rs!SALIDA_ESTADO_DESC

   txtSalidaFecha.Text = Format(rs!SALIDA_FECHA & "", "yyyy-mm-dd")
   txtSalidaTipo.Text = rs!Salida_Tipo & ""
   txtSalidaTipoDesc.Text = rs!SALIDA_TIPO_DESC
   
   txtLiquidaFecha.Text = Format(rs!Liquida_Fecha & "", "yyyy-mm-dd")
   txtLiquidaBoleta.Text = rs!Liquida_Boleta
   
   
   'Otros
   
   chkInd_Solidarista.Value = rs!APL_Solidarista
   chkInd_Marcas.Value = rs!Control_Marcas
   
   chkInd_Licencia.Value = rs!Posee_Licencia_Conducir
   chkInd_Vehiculo.Value = rs!Posee_Vehiculo
   
   txtId_Tributario.Text = rs!Id_Fiscal & ""
   txtNotas.Text = rs!notas
   
   
   txtCF_Dependientes.Text = CStr(rs!Dependientes_Numero)
   chkCF_Conyuge.Value = rs!Conyuge
  
   StatusBarX.Panels.Item(1) = rs!Registro_Usuario & ""
   StatusBarX.Panels.Item(2) = rs!Registro_Fecha & ""
 
  Else
   
   If vEditar Then
        vEditar = False
        Call sbBarra_Accion("nuevo")
        Call RefrescaTags(Me)
        Call sbClearControles
        Call sbLockControles("L")
        txtIdentificacion.SetFocus
   Else
        Call RefrescaTags(Me)
        
        Call sbLimpiaDatos
                
        txtEmpleadoId.Text = fxgRH_Empleado_ID
        
   End If

End If
rs.Close

tcMain.Item(0).Selected = True
txtApellido1.SetFocus

btnEditarDetalle.Enabled = True

End Sub


Sub sbLockControles(vModo As String)
'Dim vControl As Control
'
'For Each vControl In Me
'  If (TypeOf vControl Is TextBox And vControl.Name <> "txtIdentificacion" And vControl.Name <> _
'     "txtEstado") Or TypeOf vControl Is DTPicker Or TypeOf vControl Is ComboBox Then
'        If vModo = "L" Then
'           If vControl.Name = "txtNombre" Then
'            vControl.Locked = True
'           Else
'            vControl.Enabled = False
'           End If
'        Else
'           If vControl.Name = "txtNombre" Then
'            vControl.Locked = False
'           Else
'            vControl.Enabled = True
'           End If
'        End If
'  End If
'Next

'dtpIngreso.Enabled = False
cboTipoId.Enabled = True

End Sub


Private Sub sbDeleteRecord()
Dim i As Integer

On Error GoTo vError

If Trim(txtIdentificacion) <> vCedula Then
   MsgBox "Ha modificado la cédula", vbExclamation
   Exit Sub
End If

strSQL = "exec spAFIPersonaMovAux '" & txtIdentificacion.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Total > 0 Then
 strSQL = "Tiene Referencias en Otros Módulos..:" & vbCrLf & vbCrLf _
        & "Patrimonio...:" & rs!Patrimonio & vbCrLf _
        & "Fondos    ...:" & rs!fondos & vbCrLf _
        & "Créditos  ...:" & rs!Creditos & vbCrLf _
        & "Fianzas   ...:" & rs!FIANZAS
        
  MsgBox strSQL, vbExclamation
  rs.Close
  Exit Sub
        
 End If
rs.Close

i = MsgBox("Esta Seguro Que Desea Borrar esta Persona?", vbYesNo)
If i = vbYes Then
  vEditar = False
  strSQL = "delete Socios where Cedula='" & Trim(txtIdentificacion) & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Borra", "Persona [Identificación : " & Trim(txtIdentificacion) & "]")
  
  Call sbClearControles
  Call sbBarra_Accion("nuevo")
  Call RefrescaTags(Me)
  txtIdentificacion.SetFocus
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

'Funcion para Apellidos y Nombres que no contenga números ni otro tipo de caracter que no sean letras del alfabeto
Private Function fxValidaCadena(pCadena As String) As Boolean
Dim str As String, i As Integer, vResultado As Boolean

vResultado = True
pCadena = Trim(pCadena)
For i = 1 To Len(pCadena)
  vResultado = False
  str = Mid(pCadena, i, 1)
   If Asc(str) >= 65 And Asc(str) <= 90 Then
     vResultado = True
   End If
  
   If Asc(str) >= 97 And Asc(str) <= 122 Then
     vResultado = True
   End If
  
   If str = "á" Or str = "é" Or str = "í" Or str = "ó" Or str = "ú" Then
     vResultado = True
   End If
   
   If str = "Á" Or str = "É" Or str = "Í" Or str = "Ó" Or str = "Ú" Then
     vResultado = True
   End If
   
   If str = " " Or str = "ñ" Or str = "Ñ" Then
     vResultado = True
   End If
   
   If Not vResultado Then Exit For
Next i

fxValidaCadena = vResultado

End Function

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""

'persona fisica
If Trim(txtApellido1) = "" Then vMensaje = vMensaje & " - Falta el Apellido 1" & vbCrLf
If Trim(txtApellido2) = "" Then vMensaje = vMensaje & " - Falta el Apellido 2" & vbCrLf
If Trim(txtNombre) = "" Then vMensaje = vMensaje & " - Falta el Nombre" & vbCrLf


If Not fxValidaCadena(txtApellido1.Text) Then vMensaje = vMensaje & " - El Apellido 1: No es válido" & vbCrLf
If Not fxValidaCadena(txtApellido2.Text) Then vMensaje = vMensaje & " - El Apellido 2: No es válido" & vbCrLf
If Not fxValidaCadena(txtNombre.Text) Then vMensaje = vMensaje & " - El Nombre: No es válido" & vbCrLf

If Trim(cboSexo) = "" Then vMensaje = vMensaje & " - No se especificó el Sexo" & vbCrLf


If Trim(cboProvincia.Text) = "" Then vMensaje = vMensaje & " - No se especificó la Provincia" & vbCrLf
If Trim(cboCanton.Text) = "" Then vMensaje = vMensaje & " - No se especificó el Cantón" & vbCrLf
If Trim(txtDireccion) = "" Then vMensaje = vMensaje & " - No se especificó la Dirección" & vbCrLf


If Len(vMensaje) = 0 Then
  fxValida = True
Else
  fxValida = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub sbSaveRecord()
Dim vActiva As Boolean

Dim pNombreCompleto As String, pEstadoPersona As String
Dim pPais As String, pProvincia As String, pCanton As String, pDistrito As String

Dim pNomina As String, pNivel As String, pContrato As String, pContratoVence As String, pJefeId As String
Dim pJornada As String, pVacaciones As String, pDivisa As String, pIdFiscal As String

Dim pEstadoCivil As String, pNacionalidad As String, pSexo As String, pFormaPago As String, pBancoId As Long

On Error GoTo vError


If Not fxValida Then
  Exit Sub
End If

If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
    pEstadoCivil = "'O'"
Else
    pEstadoCivil = "'" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & "'"
End If

pEstadoPersona = "'A'"
pNombreCompleto = "'" & Trim(txtApellido1.Text) & " " & Trim(txtApellido2.Text) & " " & Trim(txtNombre.Text) & "'"
pFormaPago = "'" & cboFormaPago.ItemData(cboFormaPago.ListIndex) & "'"
pBancoId = cboBancos.ItemData(cboBancos.ListIndex)

pSexo = "'" & Mid(cboSexo.Text, 1, 1) & "'"
pPais = "'CRC'"

pProvincia = "'" & cboProvincia.ItemData(cboProvincia.ListIndex) & "'"
pCanton = "'" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
pDistrito = "'" & cboDistrito.ItemData(cboDistrito.ListIndex) & "'"
pNomina = "'" & cboNomina.ItemData(cboNomina.ListIndex) & "'"
pNivel = "'" & cboNivel.ItemData(cboNivel.ListIndex) & "'"
pJornada = "'" & cboJornada.ItemData(cboJornada.ListIndex) & "'"
pContrato = "'" & cboContrato.ItemData(cboContrato.ListIndex) & "'"
pVacaciones = "'" & cboVacaciones.ItemData(cboVacaciones.ListIndex) & "'"

pNacionalidad = "'" & cboNacionalidad.ItemData(cboNacionalidad.ListIndex) & "'"
pDivisa = "'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"

If dtpContrato.Visible Then
    pContratoVence = "'" & Format(dtpContrato.Value, "yyyy-mm-dd") & "'"
Else
    pContratoVence = "Null"
End If

If txtId_Tributario.Text = "" Then
    pIdFiscal = "'" & Trim(txtIdentificacion.Text) & "'"
Else
    pIdFiscal = "'" & Trim(txtId_Tributario.Text) & "'"
End If

If txtJefeCod.Text = "" Then
    pJefeId = "'" & Trim(txtEmpleadoId.Text) & "'"
Else
    pJefeId = "'" & Trim(txtJefeCod.Text) & "'"
End If



'Validacion de Numericos
If Not IsNumeric(txtSalario.Text) Then txtSalario.Text = "0"


If Not vEditar Then
   vActiva = True
   
   strSQL = "Insert RH_Personas(Empleado_Id, Tipo_Id, Identificacion, Id_Fiscal, Apellido_1, Apellido_2, Nombre, Nombre_Completo" _
           & ", Estado_Civil, Estado_Persona, Cod_Nacionalidad, Sexo, Fecha_Nacimiento, Pais, Provincia, Canton, Distrito, Direccion" _
           & ", Telefono1, Telefono2, Tel_Movil, Fax, Email_01, Email_02, WebSite, Apto_Postal, Notas, Dependientes_Numero,Conyuge" _
           & ", Apl_Solidarista, Control_Marcas, Posee_Vehiculo, Posee_Licencia_Conducir, Cod_Banco, Forma_Pago, Cod_Divisa" _
           & ", Cod_Centro, Cod_Departamento, Cod_Seccion, Cod_Puesto, Jefe_Id, Cod_Profesion, Nivel_Academico" _
           & ", Fecha_Ingreso, Cod_Nomina, Contrato_Tipo, Contrato_Vencimiento, Jornada_Tipo, Cod_Vaca_Regimen, Activo" _
           & ", Salario_Ordinario,Registro_Fecha, Registro_Usuario) values('" _
           & txtEmpleadoId.Text & "'," & cboTipoId.ItemData(cboTipoId.ListIndex) & ",'" & Trim(txtIdentificacion.Text) & "'," & pIdFiscal & ",'" & Trim(txtApellido1.Text) _
           & "','" & Trim(txtApellido2.Text) & "','" & Trim(txtNombre.Text) & "'," & pNombreCompleto _
           & "," & pEstadoCivil & "," & pEstadoPersona & "," & pNacionalidad & "," & pSexo & ",'" & Format(dtpNacimiento.Value, "yyyy-mm-dd") _
           & "'," & pPais & "," & pProvincia & "," & pCanton & "," & pDistrito & ",'" & Trim(txtDireccion.Text) _
           & "','" & Trim(txtTelefono1.Text) & "','" & Trim(txtTelefono2.Text) & "','" & Trim(txtTelMovil.Text) & "','" & Trim(txtTelFax.Text) _
           & "','" & Trim(txtEmail.Text) & "','" & Trim(txtEmail_02.Text) & "','" & Trim(txtWebSite.Text) & "','" & Trim(txtApartado.Text) _
           & "','" & Trim(txtNotas.Text) & "'," & txtCF_Dependientes.Text & "," & chkCF_Conyuge.Value & "," & chkInd_Solidarista.Value & "," & chkInd_Marcas.Value & "," & chkInd_Vehiculo.Value _
           & "," & chkInd_Licencia.Value & "," & pBancoId & ", " & pFormaPago & "," & pDivisa _
           & ",'" & Trim(txtCentroCod.Text) & "','" & Trim(txtDeptCodigo.Text) & "','" & Trim(txtSecCodigo.Text) & "','" & Trim(txtPuestoCod.Text) _
           & "'," & pJefeId & ",'" & Trim(txtProfesionCod.Text) & "'," & pNivel & ",'" & Format(dtpIngreso.Value, "yyyy-mm-dd") _
           & "'," & pNomina & "," & pContrato & "," & pContratoVence & "," & pJornada & "," & pVacaciones _
           & ",1," & CCur(txtSalario.Text) & ", dbo.MyGetdate() , '" & glogon.Usuario & "')"
    
   Call ConectionExecute(strSQL)
   Call Bitacora("Registra", "Persona, Identificación: " & Trim(txtIdentificacion) & ", Empleado Id: " & txtEmpleadoId.Text)
 
Else
   vActiva = False
  
    strSQL = "Update RH_Personas set Identificacion = '" & Trim(txtIdentificacion.Text) & "', Id_Fiscal = " & pIdFiscal & ", Apellido_1 = '" & Trim(txtApellido1.Text) _
           & "', Apellido_2 = '" & Trim(txtApellido2.Text) & "', Nombre = '" & Trim(txtNombre.Text) _
           & "', Nombre_Completo = " & pNombreCompleto & ", Estado_Civil = " & pEstadoCivil & ", Estado_Persona = " & pEstadoPersona _
           & ", Cod_Nacionalidad = " & pNacionalidad & ", Sexo = " & pSexo & ", Fecha_Nacimiento = '" & Format(dtpNacimiento.Value, "yyyy-mm-dd") _
           & "', Pais = " & pPais & ", Provincia = " & pProvincia & " , Canton = " & pCanton & ", Distrito = " & pDistrito _
           & ", Direccion = '" & Trim(txtDireccion.Text) & "', Telefono1 = '" & Trim(txtTelefono1.Text) & "', Telefono2 = '" & Trim(txtTelefono2.Text) _
           & "', Tel_Movil= '" & Trim(txtTelMovil.Text) & "', Fax = '" & Trim(txtTelFax.Text) & "', Email_01 = '" & Trim(txtEmail.Text) _
           & "', Email_02 = '" & Trim(txtEmail_02.Text) & "', WebSite = '" & Trim(txtWebSite.Text) & "', Apto_Postal = '" & Trim(txtApartado.Text) _
           & "', Notas = '" & Trim(txtNotas.Text) & "', Apl_Solidarista = " & chkInd_Solidarista.Value & ", Control_Marcas = " & chkInd_Marcas.Value _
           & ", Posee_Vehiculo = " & chkInd_Vehiculo.Value & ", Posee_Licencia_Conducir = " & chkInd_Licencia.Value _
           & ", Cod_Banco = " & pBancoId & ", Forma_Pago = " & pFormaPago & ", Cod_Divisa = " & pDivisa & "" _
           & ", Cod_Centro= '" & Trim(txtCentroCod.Text) & "', Cod_Departamento = '" & Trim(txtDeptCodigo.Text) & "', Cod_Seccion = '" & Trim(txtSecCodigo.Text) _
           & "', Cod_Puesto = '" & Trim(txtPuestoCod.Text) & "', Jefe_Id = " & pJefeId & ", Cod_Profesion = '" & Trim(txtProfesionCod.Text) _
           & "', Nivel_Academico = " & pNivel & ", Fecha_Ingreso = '" & Format(dtpIngreso.Value, "yyyy-mm-dd") & "', Cod_Nomina = " & pNomina _
           & ", Contrato_Tipo = " & pContrato & ", Contrato_Vencimiento = " & pContratoVence & ", Jornada_Tipo = " & pJornada _
           & ", Cod_Vaca_Regimen = " & pVacaciones & ", Salario_Ordinario = " & CCur(txtSalario.Text) _
           & ", Dependientes_Numero = " & txtCF_Dependientes.Text & ", Conyuge = " & chkCF_Conyuge.Value _
           & " Where Empleado_Id = '" & txtEmpleadoId.Text & "'"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Registra", "Persona, Identificación: " & Trim(txtIdentificacion) & ", Empleado Id: " & txtEmpleadoId.Text)
   

End If

'Registra Accion de Persona Inicial
If gbAccionPersonal.Enabled Then
    'spRH_Accion_Personal_Registro(@EmpleadoId varchar(20), @TipoAccion varchar(10), @Notas varchar(1000), @Usuario varchar(30)
    '                , @EstadoPersona varchar(10), @Puesto varchar(10), @Centro varchar(10), @Dept varchar(10), @Seccion varchar(10)
    '                , @Salario dec(12,2), @Nomina varchar(10))
   strSQL = "exec spRH_Accion_Personal_Registro '" & txtEmpleadoId.Text & "', '00', 'Registro de Primer Ingreso','" & glogon.Usuario _
           & "'," & pEstadoPersona & ",'" & Trim(txtPuestoCod.Text) & "','" & Trim(txtCentroCod.Text) & "','" & Trim(txtDeptCodigo.Text) _
           & "','" & Trim(txtSecCodigo.Text) & "'," & CCur(txtSalario.Text) & "," & pNomina
   Call ConectionExecute(strSQL)
   
End If


vCedula = Trim(txtIdentificacion)

Call sbBarra_Accion("activo")
Call RefrescaTags(Me)
Call sbLockControles("L")

MsgBox "Información guardada satisfactoriamente...", vbInformation
txtIdentificacion.SetFocus

'Abre el Marco de Contacto
Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)

vEditar = True

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnBoletaLiq_Click()
If txtLiquidaBoleta.Text <> "" Then
    Call sbBoleta_Liquidacion(txtLiquidaBoleta.Text)
End If

End Sub

Private Sub btnEditarDetalle_Click()
GLOBALES.gCedulaActual = Trim(txtIdentificacion.Text)

Select Case TituloOpciones.Tag
 Case "Familiares"
    
    GLOBALES.gCedulaActual = Trim(txtEmpleadoId.Text)
    frmRH_Empleado_Familiares.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Familiares)
 Case "Cuentas"
    GLOBALES.gTag = Trim(txtIdentificacion)
    GLOBALES.gTag2 = "RH"
    frmCC_Cuentas_Bancarias.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Cuentas)
 Case "Tarjetas"
    frmAF_PersonaTarjetas.Show vbModal
    Call sbTaskPanel_Accion(Id_TaskItem_Tarjetas)
 Case "Conceptos"
    GLOBALES.gTag = Trim(txtEmpleadoId.Text)
    frmRH_Persona_Conceptos_Fijos.Show vbModal
    
End Select

End Sub


Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass


Call Excel_Exportar_Lsw(lswHistorico)

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnFoto_Click(Index As Integer)
Dim pSQL As String, pGuardado As Boolean
Dim pArchivo As String

If Not vEditar Then Exit Sub

On Error GoTo vError

Select Case Index
  Case 0 'Subir Foto
        frmContenedor.CD.ShowOpen
        frmContenedor.CD.DialogTitle = "Buscar Imagen..."
        frmContenedor.CD.InitDir = "C:\"
        frmContenedor.CD.Filter = "*.bmp;*.gif;*.jpg"
        
        picFoto.Picture = LoadPicture(frmContenedor.CD.FileName)
        
        picFoto.PaintPicture picFoto.Picture, 0, 0, picFoto.ScaleWidth, picFoto.ScaleHeight
        
        pArchivo = SIFGlobal.DirectorioDeResultados & "\Empleado_Id_" & txtEmpleadoId.Text _
                & ".jpeg"
        
        Call sbImagen_Archivo_Guarda(picFoto, pArchivo)
        
        'Guarda al Foto
        pSQL = "select * from RH_Personas where Empleado_Id = '" & txtEmpleadoId.Text & "'"
        pGuardado = fxImagen_Guardar(pSQL, "Foto", pArchivo)
        
        
  Case 1 'Eliminar
       picFoto.Picture = Nothing
       
       pSQL = "update RH_Personas set Foto = Null where Empleado_Id = '" & txtEmpleadoId.Text & "'"
       Call ConectionExecute(pSQL)
        
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnPortal_Click()
MsgBox "Notificación de Clave para Uso de Portal enviado al Correo de la Persona!", vbInformation
End Sub

Private Sub btnSalida_Click()
If txtEmpleadoId.Text <> "" Then
   GLOBALES.gTag = txtEmpleadoId.Text
   
   Call sbFormsCall("frmRH_Salida_Empleado", vbModal, , , False, Me)
   
End If
End Sub

Private Sub cboCanton_Click()

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub

Private Sub cboContrato_Click()
If vPaso Then Exit Sub

If fxgRH_Contrato_Vence(cboContrato.ItemData(cboContrato.ListIndex)) Then
    dtpContrato.Visible = True
Else
    dtpContrato.Visible = False
End If

lblVencimiento.Visible = dtpContrato.Visible

End Sub

Private Sub cboDistrito_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDireccion.SetFocus
End Sub

Private Sub cboEstadoCivil_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboNacionalidad.SetFocus
End Sub


Private Sub chkInd_Solidarista_Click()
If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spRH_Persona_Deduccion_Solidarita '" & txtEmpleadoId.Text _
        & "','" & IIf((chkInd_Solidarista.Value = xtpChecked), "A", "E") & "','" _
        & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub FlatScroll_Laboral_Change(Index As Integer)
Dim rs As New ADODB.Recordset
Dim vCodigo As String, vColumna As String, vChar As String, vFiltroAdd As String
Dim txtCodigo As Object, txtDesc As Object

On Error GoTo vError

If Not vScroll Then Exit Sub

vChar = "'"
vFiltroAdd = ""

Select Case Index
   Case 0 'Centro
        vCodigo = txtCentroCod.Text
        vColumna = "COD_CENTRO"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_CENTRO_TRABAJO"
        
        Set txtCodigo = txtCentroCod
        Set txtDesc = txtCentroDesc
    
    Case 1 'Departamentos
        vCodigo = txtDeptCodigo.Text
        vColumna = "COD_DEPARTAMENTO"
        vFiltroAdd = " AND COD_CENTRO = '" & txtCentroCod.Text & "'"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_Departamentos"
        
        Set txtCodigo = txtDeptCodigo
        Set txtDesc = txtDeptDesc
        
        
    Case 2 'Secciones
        vCodigo = txtSecCodigo.Text
        vColumna = "COD_SECCION"
        vFiltroAdd = " AND COD_CENTRO = '" & txtCentroCod.Text & "' AND COD_DEPARTAMENTO = '" & txtDeptCodigo.Text & "'"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_SECCIONES"
        
        Set txtCodigo = txtSecCodigo
        Set txtDesc = txtSecDesc

    Case 3 'Profesion
        vCodigo = txtProfesionCod.Text
        
        vColumna = "COD_PROFESION"
        vFiltroAdd = " AND ACTIVO = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_PROFESIONES"
        
        Set txtCodigo = txtProfesionCod
        Set txtDesc = txtProfesionDesc
    
    Case 4 'Puesto
        vCodigo = txtPuestoCod.Text
        
        vColumna = "COD_PUESTO"
        vFiltroAdd = " AND ACTIVO = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',DESCRIPCION as 'Descripcion'" _
               & " from RH_PUESTOS"
        
        Set txtCodigo = txtPuestoCod
        Set txtDesc = txtPuestoDesc
    
    
    Case 5 'Jefe
        vCodigo = txtJefeCod.Text
        vColumna = "EMPLEADO_ID"
'        vFiltroAdd = " AND ACTIVA = 1"
        
        strSQL = "select Top 1 " & vColumna & " as 'Codigo',NOMBRE_COMPLETO as 'Descripcion'" _
               & " from RH_PERSONAS"
        
        Set txtCodigo = txtJefeCod
        Set txtDesc = txtJefeDesc

End Select

If vScroll Then
    
    If FlatScroll_Laboral(Index).Value = 1 Then
       strSQL = strSQL & " where " & vColumna & " > " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " asc"
    Else
       strSQL = strSQL & " where " & vColumna & " < " & vChar & vCodigo & vChar & " " & vFiltroAdd & " order by " & vColumna & " desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Codigo
      txtDesc.Text = rs!Descripcion

    End If
    rs.Close
End If



vScroll = False
FlatScroll_Laboral(Index).Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vFechaActual = Format(fxFechaServidor, "dd/mm/yyyy")

dtpIngreso.Value = vFechaActual
dtpNacimiento.Value = vFechaActual

dtpIngreso.Enabled = True

cboSexo.AddItem "Masculino"
cboSexo.AddItem "Femenino"
cboSexo.Text = "Masculino"


'Revisa cual Tipo de Identificacion es Juridica (Solo es Valido la Primera)
vTipoJuridica = 0
strSQL = "select TIPO_ID from AFI_TIPOS_IDS where Tipo_Personeria = 'J'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    vTipoJuridica = rs!Tipo_id
End If
rs.Close



strSQL = "select cod_nacionalidad as 'IdX', Descripcion as 'ItmX' from sys_nacionalidades" _
       & " where Activo = 1" _
       & " order by Omision desc, Descripcion asc"
Call sbCbo_Llena_New(cboNacionalidad, strSQL, False, True)

strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoCivil, strSQL, False, True)

'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False
Call cboTipoId_Click


vPaso = True

'Provincias
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, False, True)

'Nivel Academico
    strSQL = "select NIVEL_ACADEMICO as Idx, rtrim(Descripcion) as ItmX from RH_NIVEL_ACADEMICO"
    Call sbCbo_Llena_New(cboNivel, strSQL, False, True)

'Nomina
    strSQL = "select COD_NOMINA as Idx, rtrim(Descripcion) as ItmX from RH_NOMINAS_CATALOGO"
    Call sbCbo_Llena_New(cboNomina, strSQL, False, True)

'Divisa
    strSQL = "select COD_DIVISA as Idx, rtrim(Descripcion) as ItmX from vSys_Divisas"
    Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

'Jornada
    strSQL = "select JORNADA_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_JORNADAS_TIPOS"
    Call sbCbo_Llena_New(cboJornada, strSQL, False, True)

'Contratos
    strSQL = "Select CONTRATO_TIPO as Idx, rtrim(Descripcion) as ItmX from RH_CONTRATOS_TIPOS"
    Call sbCbo_Llena_New(cboContrato, strSQL, False, True)

'Vacaciones
    strSQL = "Select COD_VACA_REGIMEN as Idx, rtrim(Descripcion) as ItmX from RH_VACACIONES_REGIMEN"
    Call sbCbo_Llena_New(cboVacaciones, strSQL, False, True)

'Bancos Autorizados

    strSQL = "exec spRH_Bancos_Autorizados"
    Call sbCbo_Llena_New(cboBancos, strSQL, False, True)


cboFormaPago.Clear
cboFormaPago.AddItem "Transferencia"
cboFormaPago.ItemData(cboFormaPago.ListCount - 1) = "TE"
cboFormaPago.AddItem "Cheque"
cboFormaPago.ItemData(cboFormaPago.ListCount - 1) = "CK"

cboFormaPago.Text = "Transferencia"

vPaso = False

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswHistorico_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

Select Case TituloOpciones.Tag
  Case "BoletaPago"
    Call sbBoleta_Pago(txtEmpleadoId.Text, Item.Tag, Item.Text)
    
  Case "AccionPersonal"
    Call sbBoleta_Accion_Personal(Item.Text)
  
  Case "Vacaciones"
    Call sbBoleta_Vacaciones(Item.Text)
  
  Case "Incapacidades"
    Call sbBoleta_Incapacidad(Item.Text)

  Case "Permisos"
    Call sbBoleta_Permisos(Item.Text)

  Case "PlanCarrera"

  Case Else
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error GoTo vError

If Item.Index = 1 Then
    txtProfesionCod.SetFocus
    tcLaboral.Item(0).Selected = True
End If

Exit Sub
vError:

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa
End Sub

Private Sub tpMain_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)
Call sbTaskPanel_Accion(Item.Id)
End Sub


Private Sub txtCentroCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtCentroCod_LostFocus()
txtCentroDesc.Text = fxgRH_Centro_Trabajo(txtCentroCod.Text)
End Sub

Private Sub txtCentroDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select COD_CENTRO,descripcion,desc_Corta from RH_CENTRO_TRABAJO"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtCentroCod.Text = Trim(gBusquedas.Resultado)
    txtCentroDesc.Text = gBusquedas.Resultado2
  End If
End If
End Sub


Private Sub txtEmpleadoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtEmpleadoId.Text)
'   txtIdentificacion = ""
'   vCedula = ""
'
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Empleado_ID"
   gBusquedas.Orden = "Empleado_ID"
   gBusquedas.Consulta = "Select Empleado_ID,Identificacion,Nombre_Completo From Rh_Personas"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   
   txtEmpleadoId.Text = gBusquedas.Resultado
   txtEmpleadoId_LostFocus
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtEmpleadoId_LostFocus
End If

End Sub

Private Sub txtEmpleadoId_LostFocus()

Call sbCurrentRecord(txtEmpleadoId.Text, "E")

End Sub

Private Sub txtJefeCod_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtJefeDesc.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Identificacion"
   gBusquedas.Orden = "Identificacion"
   gBusquedas.Consulta = "Select Empleado_ID,Identificacion,Nombre_Completo From Rh_Personas"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   
   txtJefeCod.Text = Trim(gBusquedas.Resultado)
   txtJefeDesc.Text = gBusquedas.Resultado3
End If
End Sub

Private Sub txtJefeCod_LostFocus()
txtJefeDesc.Text = fxgRH_Empleado_Nombre(txtJefeCod.Text)
End Sub

Private Sub txtProfesionCod_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProfesionDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_profesion,descripcion from RH_Profesiones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtProfesionCod.Text = Trim(gBusquedas.Resultado)
    txtProfesionDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub


Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click

End Sub

Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub



Private Sub txtProfesionCod_LostFocus()
txtProfesionDesc.Text = fxgRH_Profesion(txtProfesionCod.Text)
End Sub

Private Sub txtProfesionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboNivel.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_profesion,descripcion from RH_Profesiones"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    txtProfesionCod.Text = Trim(gBusquedas.Resultado)
    txtProfesionDesc.Text = gBusquedas.Resultado2
  End If
End If

End Sub

Private Sub cboTipoId_Click()
If vPaso Then Exit Sub

Call sbClearControles

'If cboTipoId.ItemData(cboTipoId.ListIndex) = vTipoJuridica Then
'    fraTipo.Visible = True
'Else
'    fraTipo.Visible = False
'End If

End Sub


Private Sub FlatScrollBar_Change()
Dim rs As New ADODB.Recordset

On Error GoTo vError


If vScroll Then
    strSQL = "select Top 1 Empleado_Id from RH_PERSONAS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Empleado_Id > '" & txtEmpleadoId.Text & "' order by Empleado_Id asc"
    Else
       strSQL = strSQL & " where Empleado_Id < '" & txtEmpleadoId.Text & "' order by Empleado_Id desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbCurrentRecord(rs!Empleado_ID, "E")
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 23

Call Formularios(Me)
 
Call sbTaskPanel_Load
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
vEditar = False


Call sbBarra_Accion("nuevo")


Call sbLockControles("L")
Call RefrescaTags(Me)



tcMain.Item(0).Selected = True

End Sub




Private Sub sbLimpiaDatos()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)

gbAccionPersonal.Enabled = True

dtpIngreso.Value = vFechaActual
dtpIngreso.Enabled = True

dtpNacimiento.Value = vFechaActual
dtpContrato.Value = vFechaActual

txtVacaAcum.Text = "0"
txtIdentificacion.Text = ""
txtNombre.Text = ""
txtApellido1.Text = ""
txtApellido2.Text = ""

txtId_Tributario.Text = ""
txtEmpleadoId.Text = ""
txtEstado.Text = ""

cboSexo.Text = "Masculino"

txtTelefono1.Text = ""
txtTelefono2.Text = ""
txtTelFax.Text = ""
txtTelMovil.Text = ""

txtEmail.Text = ""
txtEmail_02.Text = ""
txtWebSite.Text = ""

txtDireccion.Text = ""
txtApartado.Text = ""

chkInd_Licencia.Value = xtpUnchecked
chkInd_Marcas.Value = xtpUnchecked
chkInd_Solidarista.Value = xtpUnchecked
chkInd_Vehiculo.Value = xtpUnchecked

chkCF_Conyuge.Value = xtpUnchecked
txtCF_Dependientes.Text = "0"


Set picFoto.Picture = Nothing

'Aplicar Default: Laboral
txtSalario.Text = Format(0, "Standard")

Call cboContrato_Click

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim lngContador As Long, strRuta As String

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Recursos Humanos: Personas"
 
 .Connect = glogon.ConectRPT
 
' Select Case ButtonMenu.Key
'
'    Case "socprov"
'       strSQL = "Select Count(*) as Registros From Socios Where EstadoActual='S'"
'       Call OpenRecordSet(rs, strSQL)
'         lngContador = rs!registros
'       rs.Close
'
'      .Formulas(0) = "Socios=" & lngContador
'      .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_SociosProvincia.rpt")
'
'    Case "exsocprov"
'       strSQL = "Select Count(*) as Registros From Socios Where EstadoActual in('A','P')"
'       Call OpenRecordSet(rs, strSQL)
'         lngContador = rs!registros
'       rs.Close
'
'      .Formulas(0) = "Socios=" & lngContador
'      .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_ExSociosProvincia.rpt")
'
'    Case "socup"
'      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_SociosPorUnidad.rpt")
'      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} = 'S'"
'
'    Case "desocup"
'      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_DetalleSociosPorUnidad.rpt")
'
'    Case "ImpBol"
'      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
'      .ReportFileName = SIFGlobal.fxPathReportes("Personas_ReimpresionBoleta.rpt")
'      .SelectionFormula = "{SOCIOS.CEDULA} ='" & Trim(txtIdentificacion) & "'"
'
'    Case "LisIng"
'         frmAf_ListadoIngreso.Show vbModal
'         Me.MousePointer = vbDefault
'         Exit Sub
'  End Select
'
' .PrintReport
End With

Me.MousePointer = vbDefault

Exit Sub

vError:

    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub txtApartado_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboProvincia.SetFocus

End Sub


Private Sub txtApellido1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtApellido2.SetFocus
End Sub

Private Sub txtApellido2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
End Sub

Private Sub txtIdentificacion_GotFocus()
vSeek = 1
End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtIdentificacion)
   txtIdentificacion = ""
   vCedula = ""
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Identificacion"
   gBusquedas.Orden = "Identificacion"
   gBusquedas.Consulta = "Select Identificacion,Empleado_ID,Nombre_Completo From Rh_Personas"
   gBusquedas.Filtro = " and Tipo_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
   frmBusquedas.Show vbModal
   
   txtIdentificacion.Text = Trim(gBusquedas.Resultado)
   txtEmpleadoId.Text = gBusquedas.Resultado2
   txtIdentificacion_LostFocus
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    txtIdentificacion_LostFocus
End If

End Sub


Private Sub txtIdentificacion_LostFocus()

If Trim(txtIdentificacion) = "" Then
  If vEditar = True Then
     vEditar = False
     Call sbBarra_Accion("nuevo")
     Call RefrescaTags(Me)
     Call sbClearControles
     Call sbLockControles("L")
  End If
Else
  If vEditar = False Or (vEditar = True And vCedula <> Trim(txtIdentificacion)) Then
     Dim pCedula As String
     
     pCedula = txtIdentificacion.Text
     
     Call sbCurrentRecord(txtIdentificacion.Text, "P")
     
     If txtIdentificacion.Text = "" Then
        txtIdentificacion.Text = pCedula
     End If
  End If
End If

End Sub


Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus

If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "cod_departamento"
    gBusquedas.Orden = "cod_departamento"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"
  
   
  
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If


End Sub

Private Sub txtDeptCodigo_LostFocus()
 txtDeptDesc.Text = fxgRH_Departamento(txtCentroCod.Text, txtDeptCodigo.Text)
End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then

    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "select cod_departamento,descripcion from RH_Departamentos"
    gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text & "'"

  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If


End Sub

Private Sub txtEmail_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtEmail_02.SetFocus
End Sub

Private Sub txtEmail_02_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtWebSite.SetFocus
End Sub

Private Sub txtNombre_GotFocus()
vSeek = 2
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstadoCivil.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = Trim(txtIdentificacion)
   txtIdentificacion = ""
   vCedula = ""
      
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre_Completo"
   gBusquedas.Orden = "Nombre_Completo"
   gBusquedas.Consulta = "Select Identificacion,Empleado_Id, Nombre_Completo From RH_Personas"
   
   frmBusquedas.Show vbModal
   
   txtIdentificacion = Trim(gBusquedas.Resultado)
   txtIdentificacion_LostFocus
End If

End Sub



Private Sub txtPuestoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDivisa.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_PUESTO"
  gBusquedas.Orden = "COD_PUESTO"
  gBusquedas.Consulta = "select COD_PUESTO,descripcion from Rh_Puestos"
  gBusquedas.Filtro = ""
        
  frmBusquedas.Show vbModal
  txtPuestoCod.Text = gBusquedas.Resultado
  txtPuestoDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtPuestoCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If txtPuestoCod.Text = "" Then Exit Sub

strSQL = "select * from RH_Puestos where cod_Puesto = '" & txtPuestoCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
   strSQL = "Salario Recomendado: " & Format(rs!Salario_Actual, "Standard") & vbCrLf _
          & "Salario Máximo     : " & Format(rs!Salario_Maximo, "Standard") & vbCrLf _
          & "Salario Mínimo     : " & Format(rs!Salario_Minimo, "Standard") & vbCrLf
   txtSalario.ToolTipText = strSQL
   
   If gbAccionPersonal.Enabled Then
    txtSalario.Text = Format(rs!Salario_Actual, "Standard")
   End If
   
End If

End Sub


Private Sub txtSalario_GotFocus()
On Error GoTo vError
txtSalario.Text = CCur(txtSalario.Text)
Exit Sub
vError:
End Sub

Private Sub txtSalario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboNomina.SetFocus
End Sub

Private Sub txtSalario_LostFocus()
On Error GoTo vError
txtSalario.Text = Format(CCur(txtSalario.Text), "Standard")
Exit Sub
vError:
End Sub

Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then

        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from Rh_Secciones"
        gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text _
                  & "' and cod_departamento = '" & txtDeptCodigo & "'"
  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtSecCodigo_LostFocus()
 txtSecDesc.Text = fxgRH_Seccion(txtCentroCod.Text, txtDeptCodigo.Text, txtSecCodigo.Text)
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPuestoCod.SetFocus

If KeyCode = vbKeyF4 Then

        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select cod_seccion,descripcion from Rh_Secciones"
        gBusquedas.Filtro = " and COD_CENTRO = '" & txtCentroCod.Text _
                  & "' and cod_departamento = '" & txtDeptCodigo & "'"

  
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub


Private Sub UpDownCf_DownClick()

Dim p As Integer

On Error GoTo vError

p = txtCF_Dependientes.Text

If p > 0 Then
    p = p - 1
    txtCF_Dependientes.Text = CStr(p)
End If

Exit Sub

vError:
  txtCF_Dependientes.Text = "0"
End Sub

Private Sub UpDownCf_UpClick()
Dim p As Integer

On Error GoTo vError

p = txtCF_Dependientes.Text

p = p + 1
txtCF_Dependientes.Text = CStr(p)

Exit Sub

vError:
  txtCF_Dependientes.Text = "0"
End Sub
