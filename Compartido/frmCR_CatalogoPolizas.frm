VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#18.6#0"; "Codejock.Controls.v18.6.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#18.6#0"; "Codejock.ShortcutBar.v18.6.0.ocx"
Begin VB.Form frmCR_CatalogoPolizas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración y Asignación de Pólizas"
   ClientHeight    =   8736
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   13068
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8736
   ScaleWidth      =   13068
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7212
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   12852
      _Version        =   1179654
      _ExtentX        =   22669
      _ExtentY        =   12721
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
      Color           =   2048
      ItemCount       =   3
      Item(0).Caption =   "Definición"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "tcAux"
      Item(0).Control(1)=   "tlb"
      Item(0).Control(2)=   "FlatScrollBar"
      Item(0).Control(3)=   "txtPoliza"
      Item(0).Control(4)=   "Label2(1)"
      Item(0).Control(5)=   "lswPolizas"
      Item(0).Control(6)=   "ShortcutCaption1"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "ArbolExp"
      Item(1).Control(2)=   "cmdActualiza"
      Item(1).Control(3)=   "lbl"
      Item(1).Control(4)=   "lblNodeLinea(0)"
      Item(1).Control(5)=   "lblNodeLinea(1)"
      Item(1).Control(6)=   "lblNodeLinea(2)"
      Item(2).Caption =   "Acreedores"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGrid"
      Begin XtremeSuiteControls.ListView lswPolizas 
         Height          =   2532
         Left            =   120
         TabIndex        =   31
         Top             =   4680
         Width           =   12612
         _Version        =   1179654
         _ExtentX        =   22246
         _ExtentY        =   4466
         _StockProps     =   77
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
         MultiSelect     =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5532
         Left            =   -64240
         TabIndex        =   62
         Top             =   840
         Visible         =   0   'False
         Width           =   6972
         _Version        =   1179654
         _ExtentX        =   12298
         _ExtentY        =   9758
         _StockProps     =   77
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "&Actualiza"
         Height          =   375
         Left            =   -58360
         TabIndex        =   57
         Top             =   6600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   3252
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   12732
         _Version        =   1179654
         _ExtentX        =   22458
         _ExtentY        =   5736
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
         Color           =   2048
         ItemCount       =   3
         Item(0).Caption =   "General"
         Item(0).ControlCount=   22
         Item(0).Control(0)=   "chkPolizaPlanPagos"
         Item(0).Control(1)=   "txtDescripcion"
         Item(0).Control(2)=   "cboAseguradora"
         Item(0).Control(3)=   "cboBase"
         Item(0).Control(4)=   "cboTipo"
         Item(0).Control(5)=   "txtCtaDesc"
         Item(0).Control(6)=   "txtCtaCodigo"
         Item(0).Control(7)=   "txtRetencionDesc"
         Item(0).Control(8)=   "txtRetencionCodigo"
         Item(0).Control(9)=   "txtCargoDesc"
         Item(0).Control(10)=   "txtCargoCodigo"
         Item(0).Control(11)=   "txtPorcentajeAplFormaliza"
         Item(0).Control(12)=   "txtPlazo"
         Item(0).Control(13)=   "txtValor"
         Item(0).Control(14)=   "Label2(5)"
         Item(0).Control(15)=   "Label2(2)"
         Item(0).Control(16)=   "Label2(3)"
         Item(0).Control(17)=   "Label2(4)"
         Item(0).Control(18)=   "Label2(6)"
         Item(0).Control(19)=   "Label2(7)"
         Item(0).Control(20)=   "Label2(12)"
         Item(0).Control(21)=   "Label2(8)"
         Item(1).Caption =   "Coberturas"
         Item(1).ControlCount=   22
         Item(1).Control(0)=   "cboVenceFrecuencia"
         Item(1).Control(1)=   "cboPolGenTipo"
         Item(1).Control(2)=   "dtpVencimiento"
         Item(1).Control(3)=   "txtDesde"
         Item(1).Control(4)=   "txtHasta"
         Item(1).Control(5)=   "txtContrato"
         Item(1).Control(6)=   "txtPolGenTipoMnt"
         Item(1).Control(7)=   "txtVenceDia"
         Item(1).Control(8)=   "chkPolizaGeneral"
         Item(1).Control(9)=   "chkPolizaRegion"
         Item(1).Control(10)=   "Label2(0)"
         Item(1).Control(11)=   "Label2(11)"
         Item(1).Control(12)=   "Label2(10)"
         Item(1).Control(13)=   "Label2(13)"
         Item(1).Control(14)=   "Label2(14)"
         Item(1).Control(15)=   "Label2(15)"
         Item(1).Control(16)=   "Label2(16)"
         Item(1).Control(17)=   "Label2(17)"
         Item(1).Control(18)=   "Label2(18)"
         Item(1).Control(19)=   "lblPolGenTipo"
         Item(1).Control(20)=   "lblPolGenTipoMnt"
         Item(1).Control(21)=   "ToolbarRegion"
         Item(2).Caption =   "Acreedores de la Póliza"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "lswAcreedores"
         Begin XtremeSuiteControls.ListView lswAcreedores 
            Height          =   2652
            Left            =   -70000
            TabIndex        =   4
            Top             =   480
            Visible         =   0   'False
            Width           =   12612
            _Version        =   1179654
            _ExtentX        =   22246
            _ExtentY        =   4678
            _StockProps     =   77
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
            Checkboxes      =   -1  'True
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   16
            ShowBorder      =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboVenceFrecuencia 
            Height          =   312
            Left            =   -65440
            TabIndex        =   5
            Top             =   1560
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboPolGenTipo 
            Height          =   312
            Left            =   -65440
            TabIndex        =   6
            Top             =   2040
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpVencimiento 
            Height          =   312
            Left            =   -65440
            TabIndex        =   7
            Top             =   1080
            Visible         =   0   'False
            Width           =   1332
            _Version        =   1179654
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtDesde 
            Height          =   312
            Left            =   -65440
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtHasta 
            Height          =   312
            Left            =   -61600
            TabIndex        =   9
            Top             =   600
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtContrato 
            Height          =   312
            Left            =   -61600
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolGenTipoMnt 
            Height          =   312
            Left            =   -61600
            TabIndex        =   11
            Top             =   2040
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtVenceDia 
            Height          =   312
            Left            =   -61600
            TabIndex        =   12
            Top             =   1560
            Visible         =   0   'False
            Width           =   972
            _Version        =   1179654
            _ExtentX        =   1714
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaGeneral 
            Height          =   372
            Left            =   -69640
            TabIndex        =   13
            Top             =   2040
            Visible         =   0   'False
            Width           =   2772
            _Version        =   1179654
            _ExtentX        =   4890
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Aplica como Póliza General "
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaRegion 
            Height          =   372
            Left            =   -69640
            TabIndex        =   14
            Top             =   2520
            Visible         =   0   'False
            Width           =   2772
            _Version        =   1179654
            _ExtentX        =   4890
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Aplica Póliza por Región "
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton ToolbarRegion 
            Height          =   372
            Left            =   -65440
            TabIndex        =   26
            Top             =   2520
            Visible         =   0   'False
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Detallar Regiones"
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
         End
         Begin XtremeSuiteControls.CheckBox chkPolizaPlanPagos 
            Height          =   972
            Left            =   10800
            TabIndex        =   33
            Top             =   600
            Width           =   1452
            _Version        =   1179654
            _ExtentX        =   2561
            _ExtentY        =   1714
            _StockProps     =   79
            Caption         =   "Aplicar dentro del Plan de Pagos"
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
         End
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   312
            Left            =   3360
            TabIndex        =   34
            Top             =   600
            Width           =   7332
            _Version        =   1179654
            _ExtentX        =   12933
            _ExtentY        =   556
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboAseguradora 
            Height          =   312
            Left            =   3360
            TabIndex        =   35
            Top             =   2760
            Width           =   7332
            _Version        =   1179654
            _ExtentX        =   12933
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboBase 
            Height          =   312
            Left            =   3360
            TabIndex        =   36
            Top             =   1320
            Width           =   2052
            _Version        =   1179654
            _ExtentX        =   3620
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboTipo 
            Height          =   312
            Left            =   5400
            TabIndex        =   37
            Top             =   1320
            Width           =   2292
            _Version        =   1179654
            _ExtentX        =   4043
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaDesc 
            Height          =   312
            Left            =   5400
            TabIndex        =   38
            Top             =   960
            Width           =   5292
            _Version        =   1179654
            _ExtentX        =   9334
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
         Begin XtremeSuiteControls.FlatEdit txtCtaCodigo 
            Height          =   312
            Left            =   3360
            TabIndex        =   39
            Top             =   960
            Width           =   2052
            _Version        =   1179654
            _ExtentX        =   3619
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRetencionDesc 
            Height          =   312
            Left            =   5400
            TabIndex        =   40
            Top             =   1680
            Width           =   5292
            _Version        =   1179654
            _ExtentX        =   9334
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
         Begin XtremeSuiteControls.FlatEdit txtRetencionCodigo 
            Height          =   312
            Left            =   3360
            TabIndex        =   41
            Top             =   1680
            Width           =   2052
            _Version        =   1179654
            _ExtentX        =   3619
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
            Height          =   312
            Left            =   5400
            TabIndex        =   42
            Top             =   2040
            Width           =   5292
            _Version        =   1179654
            _ExtentX        =   9334
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
         Begin XtremeSuiteControls.FlatEdit txtCargoCodigo 
            Height          =   312
            Left            =   3360
            TabIndex        =   43
            Top             =   2040
            Width           =   2052
            _Version        =   1179654
            _ExtentX        =   3619
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   312
            Left            =   8880
            TabIndex        =   45
            Top             =   2400
            Width           =   1812
            _Version        =   1179654
            _ExtentX        =   3196
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtValor 
            Height          =   312
            Left            =   7680
            TabIndex        =   46
            Top             =   1320
            Width           =   3012
            _Version        =   1179654
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
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPorcentajeAplFormaliza 
            Height          =   312
            Left            =   3360
            TabIndex        =   44
            Top             =   2400
            Width           =   2052
            _Version        =   1179654
            _ExtentX        =   3619
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
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo en Meses para el Cobro (Renovación)"
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
            Height          =   252
            Index           =   8
            Left            =   5520
            TabIndex        =   54
            Top             =   2400
            Width           =   3732
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Aseguradora / Beneficiario"
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
            Height          =   252
            Index           =   12
            Left            =   240
            TabIndex        =   53
            Top             =   2760
            Width           =   3012
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "% Aplicable en la Formalización"
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
            Height          =   252
            Index           =   7
            Left            =   240
            TabIndex        =   52
            Top             =   2400
            Width           =   3012
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo de Formalización Asociado"
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
            Height          =   252
            Index           =   6
            Left            =   240
            TabIndex        =   51
            Top             =   2040
            Width           =   3012
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cálculo / Distribución"
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
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   50
            Top             =   1320
            Width           =   3012
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta Contable"
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
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   49
            Top             =   960
            Width           =   3012
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   48
            Top             =   600
            Width           =   3012
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Retención Asociada"
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
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   47
            Top             =   1680
            Width           =   3012
         End
         Begin VB.Label lblPolGenTipoMnt 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Asegurado"
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
            Height          =   492
            Left            =   -62800
            TabIndex        =   25
            Top             =   1920
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label lblPolGenTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cobertura S/"
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
            Height          =   252
            Left            =   -66640
            TabIndex        =   24
            Top             =   2040
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Frecuencia"
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
            Height          =   252
            Index           =   18
            Left            =   -66640
            TabIndex        =   23
            Top             =   1560
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Día de Vencimiento"
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
            Height          =   492
            Index           =   17
            Left            =   -62800
            TabIndex        =   22
            Top             =   1440
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Contrato"
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
            Height          =   252
            Index           =   16
            Left            =   -62800
            TabIndex        =   21
            Top             =   1080
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Index           =   15
            Left            =   -66640
            TabIndex        =   20
            Top             =   1080
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Frecuencia"
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
            Height          =   252
            Index           =   14
            Left            =   -69640
            TabIndex        =   19
            Top             =   1560
            Visible         =   0   'False
            Width           =   2892
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Vencimiento de Cobertura"
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
            Height          =   252
            Index           =   13
            Left            =   -69640
            TabIndex        =   18
            Top             =   1080
            Visible         =   0   'False
            Width           =   2892
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Height          =   252
            Index           =   10
            Left            =   -66640
            TabIndex        =   17
            Top             =   600
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Height          =   252
            Index           =   11
            Left            =   -62800
            TabIndex        =   16
            Top             =   600
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rangos de Aplicación x Monto"
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
            Height          =   252
            Index           =   0
            Left            =   -69640
            TabIndex        =   15
            Top             =   600
            Visible         =   0   'False
            Width           =   2892
         End
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   4800
         TabIndex        =   27
         Top             =   480
         Width           =   3828
         _ExtentX        =   6752
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
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
               Key             =   "Reportes"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "consultar"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   252
         Left            =   4080
         TabIndex        =   28
         Top             =   480
         Width           =   492
         _ExtentX        =   868
         _ExtentY        =   445
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtPoliza 
         Height          =   312
         Left            =   1920
         TabIndex        =   29
         Top             =   480
         Width           =   2052
         _Version        =   1179654
         _ExtentX        =   3619
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6252
         Left            =   -69160
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   11172
         _Version        =   524288
         _ExtentX        =   19706
         _ExtentY        =   11028
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
         MaxCols         =   485
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_CatalogoPolizas.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   5520
         Left            =   -69880
         TabIndex        =   56
         Top             =   840
         Visible         =   0   'False
         Width           =   5532
         _ExtentX        =   9758
         _ExtentY        =   9737
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
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
      End
      Begin XtremeShortcutBar.ShortcutCaption lbl 
         Height          =   372
         Left            =   -69880
         TabIndex        =   61
         Top             =   360
         Visible         =   0   'False
         Width           =   12612
         _Version        =   1179654
         _ExtentX        =   22246
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Lista de Pólizas"
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
      End
      Begin VB.Label lblNodeLinea 
         Caption         =   "GARANTIA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -67720
         TabIndex        =   60
         ToolTipText     =   "Linea"
         Top             =   6480
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label lblNodeLinea 
         Caption         =   "DESTINO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69880
         TabIndex        =   59
         ToolTipText     =   "Linea"
         Top             =   6720
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label lblNodeLinea 
         Caption         =   "LINEA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   58
         ToolTipText     =   "Linea"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2052
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   372
         Left            =   120
         TabIndex        =   32
         Top             =   4200
         Width           =   12612
         _Version        =   1179654
         _ExtentX        =   22246
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Lista de Pólizas"
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
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Poliza"
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
         Height          =   312
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   1812
      End
   End
   Begin VB.CommandButton cmdModifica 
      Caption         =   "..."
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_CatalogoPolizas.frx":0654
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_CatalogoPolizas.frx":17016
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_CatalogoPolizas.frx":2D9D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_CatalogoPolizas.frx":42B4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_CatalogoPolizas.frx":57CBC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catálogo de Pólizas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   5172
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   13092
   End
End
Attribute VB_Name = "frmCR_CatalogoPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vConsultaActiva As Integer, vNode As Node
Dim vEditar As Boolean, vScroll As Boolean, vPaso As Boolean

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, vResulta As String
Dim vCadena As String, x As Integer

lblNodeLinea.Item(0).Tag = ""
lblNodeLinea.Item(1).Tag = ""
lblNodeLinea.Item(2).Tag = ""

lbl.Caption = Node.FullPath
lbl.Tag = Node.Key

If Right(Node.Key, 1) = "G" Then
     
   vCadena = fxIndiceCodigo(Node.Key)
   lblNodeLinea.Item(2).Tag = Right(vCadena, 1)
   x = 0
   vResulta = ""
   For i = 1 To Len(vCadena)
     If Mid(vCadena, i, 1) = "-" Then
        lblNodeLinea.Item(x).Tag = vResulta
        If x = 1 Then
          'Carta la Ultima Letra para el caso de los destinos
          lblNodeLinea.Item(x).Tag = Mid(lblNodeLinea.Item(x).Tag, 1, Len(lblNodeLinea.Item(x).Tag) - 1)
        End If
        x = x + 1
        vResulta = ""
     Else
        vResulta = vResulta & Mid(vCadena, i, 1)
     End If
   
   Next i

    Call sbCargaLswAdicional
Else
    lsw.ListItems.Clear
End If

lblNodeLinea.Item(0).Caption = "Línea   : " & lblNodeLinea.Item(0).Tag
lblNodeLinea.Item(1).Caption = "Destino : " & lblNodeLinea.Item(1).Tag
lblNodeLinea.Item(2).Caption = "Garantia: " & lblNodeLinea.Item(2).Tag


End Sub


Private Sub cboBase_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Sub cboBase_LostFocus()
    If cboBase.Text = "Avaluo" Then
        chkPolizaRegion.Visible = True
    Else
        chkPolizaRegion.Visible = False
        ToolbarRegion.Visible = False
        chkPolizaRegion.Value = vbUnchecked
    End If
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtValor.SetFocus
End Sub


Private Sub cmdReporte_Click()
With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Crédito"

   .Connect = glogon.ConectRPT

   .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
   .ReportFileName = SIFGlobal.fxPathReportes("CrdCatalogoCargos.rpt")
   
   
   .PrintReport
End With
End Sub


Private Sub cboVenceFrecuencia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtVenceDia.SetFocus
End Sub

Private Sub chkPolizaGeneral_Click()
If chkPolizaGeneral.Value = vbChecked Then
   lblPolGenTipo.Visible = True
   lblPolGenTipoMnt.Visible = True
   cboPolGenTipo.Visible = True
   txtPolGenTipoMnt.Visible = True
Else
   lblPolGenTipo.Visible = False
   lblPolGenTipoMnt.Visible = False
   cboPolGenTipo.Visible = False
   txtPolGenTipoMnt.Visible = False
End If
End Sub

Private Sub chkPolizaRegion_Click()
    If chkPolizaRegion.Value = vbChecked Then
        ToolbarRegion.Visible = True
    Else
        ToolbarRegion.Visible = False
    End If
End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_poliza from CRD_CATALOGO_POLIZAS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_poliza > '" & txtPoliza.Text & "' order by cod_poliza asc"
    Else
       strSQL = strSQL & " where cod_poliza < '" & txtPoliza.Text & "' order by cod_poliza desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtPoliza.Text = rs!cod_poliza
      Call sbConsulta(txtPoliza.Text)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 11

 vEditar = False

 tcMain.Item(0).Selected = True
 
 Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
 With lswPolizas.ColumnHeaders
    .Clear
    .Add , , "Póliza", 1200
    .Add , , "Descripción", 3000
    .Add , , "Base", 1200, vbCenter
    .Add , , "Tipo", 1600, vbCenter
    .Add , , "Valor", 1200, vbRightJustify
    .Add , , "% Apl.", 1200, vbCenter
    .Add , , "Retención", 1200, vbCenter
    .Add , , "Cargo", 1200, vbCenter
    .Add , , "Cuenta", 1800, vbCenter
 End With
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", 3000
    .Add , , "Tipo", 2100
    .Add , , "Valor", 2100, vbRightJustify
 End With
 
 With lswAcreedores.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3500
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
 End With
 
 lsw.Checkboxes = True
 lswAcreedores.Checkboxes = True
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 cboPolGenTipo.Clear
 cboPolGenTipo.AddItem "Crédito"
 cboPolGenTipo.AddItem "Asociados"
 cboPolGenTipo.Text = "Crédito"
 
 cboBase.Clear
 cboBase.AddItem "Crédito"
 cboBase.AddItem "Avaluo"
 cboBase.Text = "Crédito"


 cboVenceFrecuencia.Clear
 cboVenceFrecuencia.AddItem "Anual"
 cboVenceFrecuencia.AddItem "Mensual"
 cboVenceFrecuencia.Text = "Mensual"

 cboTipo.Clear
 cboTipo.AddItem "Porcentaje"
 cboTipo.AddItem "Monto"
 cboTipo.Text = "Porcentaje"
 
 
 txtPolGenTipoMnt.Text = Format(0, "Standard")
 
 
strSQL = "select CEDULA_JURIDICA as 'IdX',NOMBRE as 'ItmX' from CRD_POLIZAS_ASEGURADORAS where activo = 1"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)

 Call chkPolizaGeneral_Click
 Call sbLimpia
 
 Call sbPolizaLista

lsw.Enabled = cmdActualiza.Enabled

End Sub

Private Sub sbPolizaLista()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError
    
vPaso = True

lswPolizas.ListItems.Clear

strSQL = "select * from CRD_CATALOGO_POLIZAS order by cod_poliza"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  With lswPolizas.ListItems
       Set itmX = .Add(, , rs!cod_poliza)
           itmX.SubItems(1) = rs!Descripcion
           itmX.SubItems(2) = IIf((rs!base = "A"), "Avaluo", "Crédito")
           itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentaje", "Monto")
           itmX.SubItems(4) = Format(rs!Valor, "####,###,##0.0000")
           itmX.SubItems(5) = Format(rs!PORC_FORMALIZACION, "##0.0000")
           itmX.SubItems(6) = rs!Codigo_Retencion
           itmX.SubItems(7) = rs!codigo_cargo
           itmX.SubItems(8) = fxgCntCuentaFormato(True, rs!cod_cuenta)
  End With
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
    MsgBox Err.Description, vbCritical

End Sub


Private Sub sbLimpia(Optional pSoloLista As Boolean = False)
Dim strSQL As String, rs As New ADODB.Recordset

Select Case tcMain.Selected.Index
  Case 0 'Remesas
     If Not pSoloLista Then
             txtPoliza.Text = ""
             
             tcAux.Item(0).Selected = True
             
             txtDescripcion.Text = ""
             txtCtaCodigo.Text = ""
             txtCtaDesc.Text = ""
             
             cboBase.Text = "Crédito"
             cboVenceFrecuencia.Text = "Mensual"
             cboTipo.Text = "Porcentaje"
            
             chkPolizaPlanPagos.Value = vbUnchecked
            
             
             txtValor.Text = "0"
             txtDesde.Text = "0.00"
             txtHasta.Text = "999,999,999,999.99"
                  
                       
             txtRetencionCodigo.Text = ""
             txtRetencionDesc.Text = ""
             
             txtCargoCodigo.Text = ""
             txtCargoDesc.Text = ""
             
             txtPorcentajeAplFormaliza.Text = "0"
             txtPlazo.Text = "12"
             
             dtpVencimiento.Value = fxFechaServidor
             
             txtVenceDia.Text = "30"
             txtContrato.Text = ""
             
             chkPolizaGeneral.Value = vbUnchecked
     End If
     
  Case 1 'Asignacion
 End Select

End Sub


Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""
fxVerifica = True

If txtPoliza.Text = "" Then vMensaje = vMensaje & " - Especifique un código para la póliza" & vbCrLf
If txtDescripcion.Text = "" Then vMensaje = vMensaje & " - Especifique una descripción de la póliza" & vbCrLf

If txtCtaDesc.Text = "" Then vMensaje = vMensaje & " - La cuenta Contable para el registo de esta póliza" & vbCrLf
If txtRetencionDesc.Text = "" Then vMensaje = vMensaje & " - Especifique el código de Retención" & vbCrLf


If Not IsNumeric(txtPorcentajeAplFormaliza.Text) Then
    vMensaje = vMensaje & " - El porcentaje de cobro en la formalización no es válido" & vbCrLf
Else
    If CCur(txtPorcentajeAplFormaliza.Text) > 0 Then
        If txtCargoDesc.Text = "" Then vMensaje = vMensaje & " - Especifique el código de cargo para la formalización" & vbCrLf
    End If
End If

If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   fxVerifica = False
End If


End Function



Private Sub sbCargaLswAdicional()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "select R.*,A.codigo as Existe" _
       & " from CRD_CATALOGO_POLIZAS R left Join CRD_CATALOGO_POLIZAS_ASG A " _
       & " on R.cod_poliza = A.cod_poliza and A.codigo = '" & lblNodeLinea.Item(0).Tag _
       & "' and A.Cod_destino = '" & lblNodeLinea.Item(1).Tag & "' and A.Garantia = '" & lblNodeLinea.Item(2).Tag _
       & "' order by existe desc,R.cod_poliza"
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_poliza)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.SubItems(2) = IIf((rs!Tipo = "P"), "PORCENTUAL", "MONTO")
      itmX.SubItems(3) = Format(rs!Valor, "Standard")
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      
      If itmX.Checked Then itmX.ForeColor = vbBlue
      
  rs.MoveNext
Loop
rs.Close

vPaso = False


Me.MousePointer = vbDefault

End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "insert CRD_CATALOGO_POLIZAS_ASG(cod_poliza,codigo,cod_destino,garantia) values('" _
           & Item.Text & "','" & lblNodeLinea.Item(0).Tag & "','" & lblNodeLinea.Item(1).Tag _
           & "','" & lblNodeLinea.Item(2).Tag & "')"
Else
    strSQL = "delete CRD_CATALOGO_POLIZAS_ASG where cod_poliza = '" _
           & Item.Text & "' and codigo = '" & lblNodeLinea.Item(0).Tag & "' and cod_destino = '" _
           & lblNodeLinea.Item(1).Tag & "' and Garantia = '" & lblNodeLinea.Item(2).Tag & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub
vError:
 MsgBox Err.Description, vbCritical

End Sub


Private Sub sbConsulta(pPoliza As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error Resume Next

vPaso = True

strSQL = "select P.*,isnull(Ret.descripcion,'') as 'RetencionDesc',isnull(Crg.descripcion,'') as 'CargoDesc'" _
       & ", isnull(Asg.Nombre,'') as 'AseguradoraDesc'" _
       & ", isnull(Cta.Cod_Cuenta_Mask, P.cod_Cuenta) as 'CtaCodigo', isnull(Cta.Descripcion,'') as 'CtaDesc'" _
       & " from CRD_CATALOGO_POLIZAS P left join Catalogo Ret on P.codigo_Retencion = Ret.Codigo" _
       & " left join cargos_adicionales Crg on P.codigo_cargo = Crg.cod_cargo" _
       & " left join CRD_POLIZAS_ASEGURADORAS Asg on P.cod_Aseguradora = Asg.cod_Aseguradora" _
       & " left join vCNTX_CUENTAS_LOCAL Cta on P.cod_Cuenta = Cta.Cod_Cuenta" _
       & " where P.cod_poliza = '" & pPoliza & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   vEditar = True
   
   Call sbToolBar(tlb, "activo")
   
   vCodigo = Trim(rs!cod_poliza)
   
   txtPoliza.Text = rs!cod_poliza
   txtDescripcion.Text = rs!Descripcion
   
   txtCtaCodigo.Text = rs!CtaCodigo
   txtCtaDesc.Text = rs!CtaDesc
   
   If rs!base = "A" Then
      cboBase.Text = "Avaluo"
      chkPolizaRegion.Visible = True
   Else
      cboBase.Text = "Crédito"
      chkPolizaRegion.Visible = False
   End If
   
   cboTipo.Text = IIf((rs!Tipo = "P"), "Porcentaje", "Monto")
   txtValor.Text = Format(rs!Valor, "###,###,##0.00000000")
   
   chkPolizaPlanPagos.Value = rs!integra_plan_pagos
   
   txtRetencionCodigo.Text = rs!Codigo_Retencion
   txtRetencionDesc.Text = rs!RetencionDesc
   
   txtCargoCodigo.Text = rs!codigo_cargo
   txtCargoDesc.Text = rs!cargodesc
   
   txtPorcentajeAplFormaliza.Text = Format(rs!PORC_FORMALIZACION, "##0.00000000")
   txtPlazo.Text = rs!plazo_meses
   
   If Not IsNull(rs!cod_Aseguradora) Then
       Call sbCboAsignaDato(cboAseguradora, rs!AseguradoraDesc, True, rs!cod_Aseguradora)
   End If
   
   txtDesde.Text = Format(rs!COBERTURA_INICIO, "Standard")
   txtHasta.Text = Format(rs!cobertura_corte, "Standard")
             
   txtContrato.Text = rs!Contrato_Num
   txtVenceDia.Text = rs!Vence_Dia
   If rs!Vence_Frecuencia = "M" Then
     cboVenceFrecuencia.Text = "Mensual"
   Else
     cboVenceFrecuencia.Text = "Anual"
   End If
   dtpVencimiento.Value = rs!Cobertura_Vencimiento
             
   chkPolizaGeneral.Value = IIf(IsNull(rs!Poliza_general), 0, rs!Poliza_general)
       
   If rs!POLIZA_GENERAL_TIPO = "C" Then
       cboPolGenTipo.Text = "Crédito"
   Else
       cboPolGenTipo.Text = "Asociados"
   End If
   txtPolGenTipoMnt.Text = Format(rs!POLIZA_GENERAL_Monto, "Standard")
   
   
   If rs!COBERTURA_REGION = 1 Then
        chkPolizaRegion.Value = vbChecked
        ToolbarRegion.Visible = True
   Else
        chkPolizaRegion.Value = vbUnchecked
        ToolbarRegion.Visible = False
   End If
    
    Call chkPolizaGeneral_Click
    
  Else
   
   If vEditar = True Then
        vEditar = False
        Call sbToolBar(tlb, "nuevo")
        Call sbLimpia
        txtPoliza.SetFocus
   End If

End If
rs.Close

vPaso = False
'Call RefrescaTags(Me)

End Sub



Private Sub lswAcreedores_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert CRD_POLIZAS_ACREEDOR_ASG(cod_poliza,cod_acreedor,registro_fecha,registro_usuario)" _
         & " values('" & txtPoliza.Text & "','" & Item.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
  strSQL = "delete CRD_POLIZAS_ACREEDOR_ASG where cod_poliza = '" & txtPoliza.Text _
         & "' and cod_acreedor = '" & Item.Text & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox Err.Description, vbCritical


End Sub



Private Sub lswPolizas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub
Call sbConsulta(Item.Text)
End Sub



Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Lineas", "Lineas", "imgRoot")
  'Crear Arbol Inicial
  
    strSQL = "select codigo,descripcion" _
           & " from catalogo where retencion = 'N' and Poliza = 'N' and Activo = 1"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Call sbCreaNodos(vNode.Key, rs!Codigo & " - " & rs!Descripcion, "imgFolder", True, "N", "0x0" & rs!Codigo & "L")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With


End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim rsTmp As New ADODB.Recordset, vCodTmp As String


On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Lineas" Then

Select Case Right(Node.Key, 1)
        
    Case "L" 'Lineas
    
        vCodTmp = fxIndiceCodigo(Node.Key)
              
        strSQL = "select T.*" _
               & " from crd_catalogo_garantias C inner join crd_garantia_tipos T on C.garantia = T.garantia" _
               & " where C.codigo = '" & vCodTmp & "'"
        rsTmp.Open strSQL, glogon.Conection, adOpenStatic
                        
        strSQL = "select * from catalogo_destinos" _
               & " where cod_destino in (select cod_destino from CATALOGO_DESTINOSASG" _
               & " where codigo = '" & vCodTmp & "')"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          'Destinos y Garantias
          Call sbCreaNodos(Node.Key, rs!cod_destino & " - " & rs!Descripcion, "imgFolder", True, "N", "0x0" & vCodTmp & "-" & rs!cod_destino & "D")
          
          If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
          Do While Not rsTmp.EOF
             Call sbCreaNodos("0x0" & vCodTmp & "-" & rs!cod_destino & "D", rsTmp!Descripcion, "imgAsientos", False, "N", "0x0" & vCodTmp & "-" & rs!cod_destino & "D" & "-" & rsTmp!Garantia & "G")
            rsTmp.MoveNext
          Loop
          
          rs.MoveNext
        Loop
        rs.Close
        rsTmp.Close
    
    Case Else 'SubCuentas
     ''
End Select

End If

End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
               , vAcepta As String, Optional xkey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
    
End Sub




Private Sub sbBorrar()

End Sub


Private Sub sbGuardar()
Dim strSQL As String

On Error GoTo vError

'If Not fxValida Then
'  Exit Sub
'End If

If vEditar = True Then
  
 If Trim(txtPoliza) <> vCodigo Then
   MsgBox "Ha modificado el Código de la Poliza", vbExclamation
   Exit Sub
 End If
 
 If IsNumeric(txtVenceDia.Text) Then
    If CInt(txtVenceDia.Text) > 30 Or CInt(txtVenceDia.Text) < 1 Then
        MsgBox "El día de Vencimiento tienen que estar entre 1 y 30", vbExclamation
        Exit Sub
    End If
 Else
   MsgBox "El día de Vencimiento tienen que estar entre 1 y 30", vbExclamation
   Exit Sub
 End If
 
End If

If Not IsNumeric(txtPolGenTipoMnt.Text) Or txtPolGenTipoMnt.Text = "" Then
  txtPolGenTipoMnt.Text = 0
End If


If Not vEditar Then
   strSQL = "insert CRD_CATALOGO_POLIZAS(cod_poliza,descripcion,base,tipo,valor,porc_formalizacion,plazo_meses,cod_cuenta" _
          & ",codigo_retencion,codigo_cargo,cobertura_inicio,cobertura_corte,cod_aseguradora,contrato_num,cobertura_vencimiento" _
          & ",vence_frecuencia,vence_dia, Poliza_general, COBERTURA_REGION,INTEGRA_PLAN_PAGOS,POLIZA_GENERAL_TIPO,POLIZA_GENERAL_MONTO)" _
          & " values('" & Trim(txtPoliza.Text) & "','" & txtDescripcion.Text & "','" & Mid(cboBase.Text, 1, 1) _
          & "','" & Mid(cboTipo.Text, 1, 1) & "'," & CDbl(txtValor.Text) & "," & CDbl(txtPorcentajeAplFormaliza.Text) & "," & txtPlazo.Text _
          & ",'" & fxgCntCuentaFormato(False, txtCtaCodigo.Text) & "','" & txtRetencionCodigo.Text & "','" & txtCargoCodigo.Text _
          & "'," & CCur(txtDesde.Text) & "," & CCur(txtHasta.Text) & ",'" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "','" & txtContrato.Text _
          & "','" & Format(dtpVencimiento.Value, "yyyy/mm/dd") & "','" & Mid(cboVenceFrecuencia.Text, 1, 1) & "'," _
          & txtVenceDia.Text & "," & chkPolizaGeneral.Value & "," & chkPolizaRegion.Value & "," & chkPolizaPlanPagos.Value _
          & ",'" & Mid(cboPolGenTipo.Text, 1, 1) & "'," & CCur(txtPolGenTipoMnt.Text) & ")"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Registra", "Poliza (Control de Polizas) : " & Trim(txtPoliza))

Else
   strSQL = "update CRD_CATALOGO_POLIZAS set descripcion = '" & txtDescripcion.Text & "',base = '" & Mid(cboBase.Text, 1, 1) _
          & "',Tipo = '" & Mid(cboTipo.Text, 1, 1) & "',valor = " & CDbl(txtValor.Text) & ",porc_formalizacion = " & CDbl(txtPorcentajeAplFormaliza.Text) _
          & ",plazo_meses = " & txtPlazo.Text & ",cod_cuenta = '" & fxgCntCuentaFormato(False, txtCtaCodigo.Text) & "',Codigo_Retencion = '" _
          & txtRetencionCodigo.Text & "',Codigo_Cargo = '" & txtCargoCodigo.Text & "',cobertura_inicio = " & CCur(txtDesde.Text) _
          & ",cobertura_corte = " & CCur(txtHasta.Text) & ", Cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) _
          & "',contrato_num = '" & txtContrato.Text & "', cobertura_vencimiento = '" & Format(dtpVencimiento.Value, "yyyy/mm/dd") _
          & "',vence_frecuencia = '" & Mid(cboVenceFrecuencia.Text, 1, 1) & "', vence_dia = " & txtVenceDia.Text _
          & ", Poliza_General = " & chkPolizaGeneral.Value & ", COBERTURA_REGION = " & chkPolizaRegion.Value _
          & ", INTEGRA_PLAN_PAGOS = " & chkPolizaPlanPagos.Value & ", POLIZA_GENERAL_TIPO = '" & Mid(cboPolGenTipo.Text, 1, 1) _
          & "', POLIZA_GENERAL_MONTO = " & CCur(txtPolGenTipoMnt.Text) _
          & " where cod_poliza = '" & txtPoliza.Text & "'"
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Registra", "Poliza (Control de Polizas) : " & vCodigo)

End If

Call sbLimpia(True)

vCodigo = Trim(txtPoliza)
vEditar = True

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation
txtPoliza.SetFocus

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub sbPolizaListaAcreedores(pPoliza As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswAcreedores.ListItems.Clear

vPaso = True

strSQL = "select Acr.COD_ACREEDOR,Acr.IDENTIFICACION,Acr.NOMBRE,Acr.ACTIVO,Asg.registro_fecha,Asg.registro_usuario" _
       & " from CRD_POLIZAS_ACREEDORES Acr left join CRD_POLIZAS_ACREEDOR_ASG Asg" _
       & " on Acr.cod_acreedor = Asg.cod_acreedor and Asg.cod_poliza = '" & pPoliza _
       & "' Where Acr.Activo = 1  order by Asg.registro_fecha desc, Acr.COD_ACREEDOR"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAcreedores.ListItems.Add(, , rs!COD_ACREEDOR)
   itmX.SubItems(1) = rs!Identificacion
   itmX.SubItems(2) = rs!Nombre
   itmX.SubItems(3) = rs!registro_fecha & ""
   itmX.SubItems(4) = rs!registro_usuario & ""
   
   If Not IsNull(rs!registro_fecha) Then itmX.Checked = True

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

vPaso = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub


Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Or Item.Index = 1 Then Exit Sub

Call sbPolizaListaAcreedores(txtPoliza.Text)

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String

Select Case Item.Index
  Case 0 'Nada
  Case 1 'Asignación
     Me.MousePointer = vbHourglass
        vCodigo = ""
        lbl.Caption = ""
        lsw.ListItems.Clear
        
        Call sbRefrescaArbol
      Me.MousePointer = vbDefault
  
  Case 2 'Lista de Acreedores
      strSQL = "SELECT COD_ACREEDOR,IDENTIFICACION,NOMBRE,CXP_ENLACE,ACTIVO" _
             & " FROM CRD_POLIZAS_ACREEDORES ORDER BY COD_ACREEDOR"
      Call sbCargaGrid(vGrid, 5, strSQL)
End Select




End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "nuevo"
    vEditar = False
    Call sbToolBar(Me.tlb, "edicion")
    Call sbLimpia
    txtPoliza.SetFocus
    
  Case "editar"
    
    vEditar = True
    vCodigo = Trim(txtPoliza)
    Call sbToolBar(tlb, "edicion")
    txtDescripcion.SetFocus
        
  Case "borrar"
    Call sbBorrar
    Call sbPolizaLista
    
  Case "guardar"
    Call sbGuardar
    Call sbPolizaLista
    
  Case "deshacer"
    vEditar = False
    Call sbToolBar(tlb, "nuevo")
    Call RefrescaTags(Me)
    Call sbLimpia
    txtPoliza.SetFocus
    
  Case "consultar"
    Select Case vConsultaActiva
      Case 1 'Consulta Cuenta
           gCuenta = ""
           frmCntX_ConsultaCuentas.Show vbModal
           If gCuenta <> "" Then
              txtCtaDesc.Text = fxgCntCuentaDesc(gCuenta)
              txtCtaCodigo.Text = fxgCntCuentaFormato(True, gCuenta)
           End If
           
      Case 2, 3 'Consulta Retencion
          gBusquedas.Consulta = "select codigo,descripcion from catalogo"
          gBusquedas.Filtro = " and poliza = 'S'"
          gBusquedas.Resultado = ""
          gBusquedas.Resultado2 = ""
          If vConsultaActiva = 2 Then
            gBusquedas.Columna = "codigo"
            gBusquedas.Orden = "codigo"
          Else
            gBusquedas.Columna = "descripcion"
            gBusquedas.Orden = "descripcion"
          End If
          
          frmBusquedas.Show vbModal
          If gBusquedas.Resultado <> "" Then
            txtRetencionCodigo.Text = gBusquedas.Resultado
            txtRetencionDesc.Text = gBusquedas.Resultado2
          End If
          
      
      Case 4, 5 'Consulta Cargo
          gBusquedas.Consulta = "select cod_cargo,descripcion from cargos_adicionales"
          gBusquedas.Filtro = ""
          gBusquedas.Resultado = ""
          gBusquedas.Resultado2 = ""
          If vConsultaActiva = 4 Then
            gBusquedas.Columna = "cod_cargo"
            gBusquedas.Orden = "cod_cargo"
          Else
            gBusquedas.Columna = "descripcion"
            gBusquedas.Orden = "descripcion"
          End If
          
          frmBusquedas.Show vbModal
          If gBusquedas.Resultado <> "" Then
            txtCargoCodigo.Text = gBusquedas.Resultado
            txtCargoDesc.Text = gBusquedas.Resultado2
          End If
    
    
    End Select
    
End Select

End Sub

Private Sub txtAseguradora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesde.SetFocus
End Sub

Private Sub ToolbarRegion_Click()
    GLOBALES.gTag = Trim(txtPoliza.Text)
    GLOBALES.gTag2 = Trim(txtDescripcion.Text)
    frmCR_PolizasRegiones.Show vbModal
End Sub

Private Sub txtCargoCodigo_GotFocus()
vConsultaActiva = 4
End Sub

Private Sub txtCargoCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  Call tlb_ButtonClick(tlb.Buttons.Item(8))
End If
End Sub

Private Sub txtCargoCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strSQL = "select descripcion from cargos_adicionales where cod_cargo = '" _
       & txtCargoCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   txtCargoDesc.Text = ""
Else
   txtCargoDesc.Text = rs!Descripcion
End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub txtCargoDesc_GotFocus()
vConsultaActiva = 5
End Sub

Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPorcentajeAplFormaliza.SetFocus

If KeyCode = vbKeyF4 Then
  Call tlb_ButtonClick(tlb.Buttons.Item(8))
End If
End Sub

Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboVenceFrecuencia.SetFocus
End Sub

Private Sub txtCtaCodigo_GotFocus()
vConsultaActiva = 1

txtCtaCodigo.Text = fxgCntCuentaFormato(False, txtCtaCodigo.Text, 0)

End Sub

Private Sub txtCtaCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaDesc.SetFocus

If KeyCode = vbKeyF4 Then
  Call tlb_ButtonClick(tlb.Buttons.Item(8))
End If

End Sub

Private Sub txtCtaCodigo_LostFocus()
txtCtaCodigo.Text = fxgCntCuentaFormato(False, txtCtaCodigo.Text)
  txtCtaDesc.Text = fxgCntCuentaDesc(txtCtaCodigo.Text)
txtCtaCodigo.Text = fxgCntCuentaFormato(True, txtCtaCodigo.Text)
End Sub

Private Sub txtCtaDesc_GotFocus()
vConsultaActiva = 1
End Sub

Private Sub txtCtaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBase.SetFocus

If KeyCode = vbKeyF4 Then
  Call tlb_ButtonClick(tlb.Buttons.Item(8))
End If

End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaCodigo.SetFocus
End Sub



Private Sub txtDesde_GotFocus()
On Error GoTo vError
 txtDesde.Text = CCur(txtDesde.Text)
vError:
End Sub

Private Sub txtDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtHasta.SetFocus
End Sub

Private Sub txtDesde_LostFocus()
On Error GoTo vError
 txtDesde.Text = Format(CCur(txtDesde.Text), "Standard")
vError:
End Sub

Private Sub txtHasta_GotFocus()
On Error GoTo vError
 txtHasta.Text = CCur(txtHasta.Text)
vError:
End Sub

Private Sub txtHasta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpVencimiento.SetFocus
End Sub

Private Sub txtHasta_LostFocus()
On Error GoTo vError
 txtHasta.Text = Format(CCur(txtHasta.Text), "Standard")
vError:
End Sub

Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboAseguradora.SetFocus
End Sub

Private Sub txtPoliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub txtPoliza_LostFocus()
 Call sbConsulta(txtPoliza.Text)
End Sub

Private Sub txtPorcentajeAplFormaliza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub

Private Sub txtRetencionCodigo_GotFocus()
vConsultaActiva = 2
End Sub

Private Sub txtRetencionCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRetencionDesc.SetFocus

If KeyCode = vbKeyF4 Then
  Call tlb_ButtonClick(tlb.Buttons.Item(8))
End If

End Sub

Private Sub txtRetencionCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strSQL = "select descripcion from catalogo where poliza = 'S' or retencion = 'S' and codigo = '" _
       & txtRetencionCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
   txtRetencionDesc.Text = ""
Else
   txtRetencionDesc.Text = rs!Descripcion
End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub txtRetencionDesc_GotFocus()
vConsultaActiva = 3
End Sub

Private Sub txtRetencionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoCodigo.SetFocus

If KeyCode = vbKeyF4 Then
  Call tlb_ButtonClick(tlb.Buttons.Item(8))
End If
End Sub

Private Sub txtValor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRetencionCodigo.SetFocus
End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from CRD_POLIZAS_ACREEDORES " _
       & " where COD_ACREEDOR = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into CRD_POLIZAS_ACREEDORES(COD_ACREEDOR,identificacion,nombre,CXP_ENLACE,activo,registro_fecha,registro_usuario) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & ","
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Acreedor de Pólizas : " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CRD_POLIZAS_ACREEDORES set identificacion = '" & vGrid.Text & "',nombre = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "',CxP_Enlace = "
 vGrid.Col = 4
 strSQL = strSQL & vGrid.Text & ",Activo = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & " where COD_ACREEDOR = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Acreedor de Pólizas : " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CRD_POLIZAS_ACREEDORES where COD_ACREEDOR = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Acreedor de Pólizas : " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     
     End If
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub


