VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Prendas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion de la Garantia Prendaria"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   12750
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   10200
      TabIndex        =   78
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   1560
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   582
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
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCR_Prendas_New.frx":0000
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   330
      Left            =   6840
      TabIndex        =   0
      Top             =   1560
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cerrar"
            Object.ToolTipText     =   "Cierra esta ventana"
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   1560
      Width           =   4815
      _Version        =   1572864
      _ExtentX        =   8493
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
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   315
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   315
      Left            =   7680
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4920
      TabIndex        =   6
      Top             =   600
      Width           =   4935
      _Version        =   1572864
      _ExtentX        =   8705
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   12855
      _Version        =   1572864
      _ExtentX        =   22675
      _ExtentY        =   12091
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
      Color           =   128
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "General"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "Label1(1)"
      Item(0).Control(2)=   "txtCoberturaTotal"
      Item(1).Caption =   "Garantía"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "tcAux"
      Item(1).Control(1)=   "tcGarantia"
      Item(2).Caption =   "Notariado y Notas de Trámite"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "gbInfoNotarial"
      Item(2).Control(1)=   "gbTramite"
      Item(3).Caption =   "Históricos"
      Item(3).ControlCount=   7
      Item(3).Control(0)=   "Label3(37)"
      Item(3).Control(1)=   "Label3(50)"
      Item(3).Control(2)=   "lblRegistroUsuario"
      Item(3).Control(3)=   "lblRegistroUsuarioAbog"
      Item(3).Control(4)=   "lswH"
      Item(3).Control(5)=   "btnHistoricos(0)"
      Item(3).Control(6)=   "btnHistoricos(1)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   12495
         _Version        =   1572864
         _ExtentX        =   22040
         _ExtentY        =   10186
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
      Begin XtremeSuiteControls.ListView lswH 
         Height          =   5775
         Left            =   -70000
         TabIndex        =   126
         Top             =   960
         Visible         =   0   'False
         Width           =   12735
         _Version        =   1572864
         _ExtentX        =   22463
         _ExtentY        =   10186
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
         Appearance      =   21
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcGarantia 
         Height          =   3375
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   12735
         _Version        =   1572864
         _ExtentX        =   22463
         _ExtentY        =   5953
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   2
         PaintManager.Position=   3
         ItemCount       =   2
         Item(0).Caption =   "General"
         Item(0).ControlCount=   16
         Item(0).Control(0)=   "txtModelo"
         Item(0).Control(1)=   "txtSerie"
         Item(0).Control(2)=   "txtDescripcion"
         Item(0).Control(3)=   "txtMarca"
         Item(0).Control(4)=   "Label2(1)"
         Item(0).Control(5)=   "Label1(4)"
         Item(0).Control(6)=   "Label1(5)"
         Item(0).Control(7)=   "Label1(0)"
         Item(0).Control(8)=   "Label2(0)"
         Item(0).Control(9)=   "Label2(2)"
         Item(0).Control(10)=   "txtAnio"
         Item(0).Control(11)=   "txtColor"
         Item(0).Control(12)=   "txtId_01"
         Item(0).Control(13)=   "txtId_02"
         Item(0).Control(14)=   "Label4(22)"
         Item(0).Control(15)=   "Label4(23)"
         Item(1).Caption =   "Vehicular"
         Item(1).ControlCount=   35
         Item(1).Control(0)=   "Label4(0)"
         Item(1).Control(1)=   "Label4(1)"
         Item(1).Control(2)=   "Label4(2)"
         Item(1).Control(3)=   "Label4(3)"
         Item(1).Control(4)=   "Label4(4)"
         Item(1).Control(5)=   "Label4(6)"
         Item(1).Control(6)=   "Label4(7)"
         Item(1).Control(7)=   "Label4(8)"
         Item(1).Control(8)=   "Label4(9)"
         Item(1).Control(9)=   "Label4(5)"
         Item(1).Control(10)=   "Label4(10)"
         Item(1).Control(11)=   "Label4(12)"
         Item(1).Control(12)=   "Label4(14)"
         Item(1).Control(13)=   "Label4(15)"
         Item(1).Control(14)=   "Label4(17)"
         Item(1).Control(15)=   "Label4(18)"
         Item(1).Control(16)=   "txtV_PlacaRegistral"
         Item(1).Control(17)=   "txtV_PlacaProvisional"
         Item(1).Control(18)=   "txtV_Color"
         Item(1).Control(19)=   "txtV_Anio"
         Item(1).Control(20)=   "cboV_Marca"
         Item(1).Control(21)=   "cboV_Combustible"
         Item(1).Control(22)=   "txtV_Chasis"
         Item(1).Control(23)=   "txtV_Capacidad"
         Item(1).Control(24)=   "txtV_Peso"
         Item(1).Control(25)=   "txtV_Puertas"
         Item(1).Control(26)=   "txtV_Cilindraje"
         Item(1).Control(27)=   "cboV_Uso"
         Item(1).Control(28)=   "txtV_VIN"
         Item(1).Control(29)=   "cboV_Presentacion"
         Item(1).Control(30)=   "cboV_Comercializa"
         Item(1).Control(31)=   "cboUd_Capacidad"
         Item(1).Control(32)=   "cboUd_Peso"
         Item(1).Control(33)=   "cboUd_Cilindraje"
         Item(1).Control(34)=   "cboV_Modelo"
         Begin XtremeSuiteControls.FlatEdit txtModelo 
            Height          =   330
            Left            =   4320
            TabIndex        =   33
            Top             =   1920
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtSerie 
            Height          =   330
            Left            =   1920
            TabIndex        =   34
            Top             =   2880
            Width           =   4575
            _Version        =   1572864
            _ExtentX        =   8070
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
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   915
            Left            =   1920
            TabIndex        =   35
            Top             =   600
            Width           =   10335
            _Version        =   1572864
            _ExtentX        =   18230
            _ExtentY        =   1614
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMarca 
            Height          =   330
            Left            =   1920
            TabIndex        =   36
            Top             =   1920
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
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
         Begin XtremeSuiteControls.FlatEdit txtV_PlacaRegistral 
            Height          =   330
            Left            =   -68080
            TabIndex        =   57
            Top             =   120
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtV_PlacaProvisional 
            Height          =   330
            Left            =   -68080
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtV_Color 
            Height          =   330
            Left            =   -68080
            TabIndex        =   59
            Top             =   840
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtV_Anio 
            Height          =   330
            Left            =   -68080
            TabIndex        =   60
            Top             =   1200
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.ComboBox cboV_Marca 
            Height          =   330
            Left            =   -68080
            TabIndex        =   61
            Top             =   1560
            Visible         =   0   'False
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.ComboBox cboV_Presentacion 
            Height          =   330
            Left            =   -68080
            TabIndex        =   62
            Top             =   1920
            Visible         =   0   'False
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.ComboBox cboV_Combustible 
            Height          =   330
            Left            =   -68080
            TabIndex        =   63
            Top             =   2280
            Visible         =   0   'False
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.ComboBox cboV_Comercializa 
            Height          =   330
            Left            =   -68080
            TabIndex        =   64
            Top             =   2760
            Visible         =   0   'False
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.FlatEdit txtV_Chasis 
            Height          =   330
            Left            =   -62680
            TabIndex        =   65
            Top             =   480
            Visible         =   0   'False
            Width           =   3135
            _Version        =   1572864
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.FlatEdit txtV_Capacidad 
            Height          =   330
            Left            =   -62680
            TabIndex        =   66
            Top             =   1320
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtV_Puertas 
            Height          =   330
            Left            =   -62680
            TabIndex        =   68
            Top             =   2400
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtV_Cilindraje 
            Height          =   330
            Left            =   -62680
            TabIndex        =   69
            Top             =   2040
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboV_Uso 
            Height          =   330
            Left            =   -62680
            TabIndex        =   70
            Top             =   2760
            Visible         =   0   'False
            Width           =   3135
            _Version        =   1572864
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.FlatEdit txtV_VIN 
            Height          =   330
            Left            =   -62680
            TabIndex        =   71
            Top             =   840
            Visible         =   0   'False
            Width           =   3135
            _Version        =   1572864
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.ComboBox cboUd_Capacidad 
            Height          =   330
            Left            =   -61600
            TabIndex        =   79
            Top             =   1320
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.ComboBox cboUd_Peso 
            Height          =   330
            Left            =   -61600
            TabIndex        =   80
            Top             =   1680
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.ComboBox cboUd_Cilindraje 
            Height          =   330
            Left            =   -61600
            TabIndex        =   81
            Top             =   2040
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
         Begin XtremeSuiteControls.ComboBox cboV_Modelo 
            Height          =   330
            Left            =   -62680
            TabIndex        =   85
            Top             =   120
            Visible         =   0   'False
            Width           =   3135
            _Version        =   1572864
            _ExtentX        =   5530
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
         Begin XtremeSuiteControls.FlatEdit txtAnio 
            Height          =   330
            Left            =   6600
            TabIndex        =   145
            Top             =   1920
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
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
         Begin XtremeSuiteControls.FlatEdit txtColor 
            Height          =   330
            Left            =   7920
            TabIndex        =   147
            Top             =   1920
            Width           =   2175
            _Version        =   1572864
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtId_01 
            Height          =   330
            Left            =   6600
            TabIndex        =   149
            Top             =   2880
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtId_02 
            Height          =   330
            Left            =   8400
            TabIndex        =   150
            Top             =   2880
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtV_Peso 
            Height          =   330
            Left            =   -62680
            TabIndex        =   67
            Top             =   1680
            Visible         =   0   'False
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   23
            Left            =   6600
            TabIndex        =   152
            Top             =   2640
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Id Principal"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   22
            Left            =   8400
            TabIndex        =   151
            Top             =   2640
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Id Secundario"
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
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
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
            Left            =   7920
            TabIndex        =   148
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Año Fab."
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
            Left            =   6600
            TabIndex        =   146
            Top             =   1680
            Width           =   1575
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   18
            Left            =   -63760
            TabIndex        =   56
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "VIN: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   17
            Left            =   -63760
            TabIndex        =   55
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Uso: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   15
            Left            =   -63760
            TabIndex        =   54
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cilindraje: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   14
            Left            =   -63760
            TabIndex        =   53
            Top             =   2400
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Puertas: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   12
            Left            =   -63760
            TabIndex        =   52
            Top             =   1680
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Peso: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   10
            Left            =   -63760
            TabIndex        =   51
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Capacidad: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   5
            Left            =   -63760
            TabIndex        =   50
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Chasís: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   9
            Left            =   -63760
            TabIndex        =   49
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Modelo: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   8
            Left            =   -69760
            TabIndex        =   48
            Top             =   2760
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Comercializa: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   7
            Left            =   -69760
            TabIndex        =   47
            Top             =   2280
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo Combustible: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   6
            Left            =   -69760
            TabIndex        =   46
            Top             =   1920
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Presentación: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   4
            Left            =   -69760
            TabIndex        =   45
            Top             =   1560
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Marca: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   3
            Left            =   -69760
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Año Fabricación: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   2
            Left            =   -69760
            TabIndex        =   43
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Color: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   1
            Left            =   -69760
            TabIndex        =   42
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Placa Provisional:"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   0
            Left            =   -69760
            TabIndex        =   41
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Placa Registral: "
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
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Serie"
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
            Left            =   1920
            TabIndex        =   40
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion"
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
            Index           =   5
            Left            =   120
            TabIndex        =   39
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Marca"
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
            Left            =   1920
            TabIndex        =   38
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo"
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
            Left            =   4320
            TabIndex        =   37
            Top             =   1680
            Width           =   1575
         End
      End
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   3015
         Left            =   0
         TabIndex        =   13
         Top             =   3840
         Width           =   12735
         _Version        =   1572864
         _ExtentX        =   22463
         _ExtentY        =   5318
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         ItemCount       =   3
         SelectedItem    =   1
         Item(0).Caption =   "Anotaciones"
         Item(0).ControlCount=   4
         Item(0).Control(0)=   "txtObservaciones"
         Item(0).Control(1)=   "lswPolizas"
         Item(0).Control(2)=   "Label4(19)"
         Item(0).Control(3)=   "Label4(20)"
         Item(1).Caption =   "Avalúo"
         Item(1).ControlCount=   24
         Item(1).Control(0)=   "Label10(16)"
         Item(1).Control(1)=   "Label10(15)"
         Item(1).Control(2)=   "Label10(13)"
         Item(1).Control(3)=   "dtpFechaInspeccion"
         Item(1).Control(4)=   "txtAvaluo_Notas"
         Item(1).Control(5)=   "Label10(14)"
         Item(1).Control(6)=   "Label10(18)"
         Item(1).Control(7)=   "optPoliza(0)"
         Item(1).Control(8)=   "optPoliza(1)"
         Item(1).Control(9)=   "Label10(0)"
         Item(1).Control(10)=   "txtValorTotal"
         Item(1).Control(11)=   "txtExtras"
         Item(1).Control(12)=   "txtValorSExtras"
         Item(1).Control(13)=   "lblValorPrenda"
         Item(1).Control(14)=   "txtValorFiscal"
         Item(1).Control(15)=   "btnExtras"
         Item(1).Control(16)=   "txtCobertura"
         Item(1).Control(17)=   "txtCoberturaPorc"
         Item(1).Control(18)=   "Label1(2)"
         Item(1).Control(19)=   "txtPolizaFormaliza"
         Item(1).Control(20)=   "txtPolizaRstPlan"
         Item(1).Control(21)=   "Label10(1)"
         Item(1).Control(22)=   "Label10(2)"
         Item(1).Control(23)=   "btnAvaluo"
         Item(2).Caption =   "Póliza Externa"
         Item(2).ControlCount=   19
         Item(2).Control(0)=   "Label4(11)"
         Item(2).Control(1)=   "Label4(13)"
         Item(2).Control(2)=   "btnPolizaExterna"
         Item(2).Control(3)=   "cboPE_Aseguradora"
         Item(2).Control(4)=   "txtPE_Numero"
         Item(2).Control(5)=   "dtpPE_Inicia"
         Item(2).Control(6)=   "Label10(3)"
         Item(2).Control(7)=   "dtpPE_Vence"
         Item(2).Control(8)=   "Label10(4)"
         Item(2).Control(9)=   "txtPE_Prima"
         Item(2).Control(10)=   "Label4(16)"
         Item(2).Control(11)=   "cboPE_Frecuencia"
         Item(2).Control(12)=   "Label4(21)"
         Item(2).Control(13)=   "txtPE_Cobertura"
         Item(2).Control(14)=   "Label10(5)"
         Item(2).Control(15)=   "txtPE_Notas"
         Item(2).Control(16)=   "Label10(6)"
         Item(2).Control(17)=   "chkPE_Activa"
         Item(2).Control(18)=   "chkPE_Indica"
         Begin XtremeSuiteControls.CheckBox chkPE_Activa 
            Height          =   375
            Left            =   -64600
            TabIndex        =   143
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Activa ?"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.PushButton btnExtras 
            Height          =   330
            Left            =   3840
            TabIndex        =   75
            Top             =   1560
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   582
            _StockProps     =   79
            BackColor       =   14737632
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Prendas_New.frx":0089
         End
         Begin XtremeSuiteControls.ListView lswPolizas 
            Height          =   2175
            Left            =   -63760
            TabIndex        =   72
            Top             =   720
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
            _ExtentY        =   3836
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
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtObservaciones 
            Height          =   2175
            Left            =   -69880
            TabIndex        =   14
            Top             =   720
            Visible         =   0   'False
            Width           =   6015
            _Version        =   1572864
            _ExtentX        =   10610
            _ExtentY        =   3836
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtValorTotal 
            Height          =   315
            Left            =   1920
            TabIndex        =   15
            Top             =   2040
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtExtras 
            Height          =   315
            Left            =   1920
            TabIndex        =   16
            Top             =   1560
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtValorSExtras 
            Height          =   315
            Left            =   1920
            TabIndex        =   17
            Top             =   1200
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaInspeccion 
            Height          =   315
            Left            =   1920
            TabIndex        =   18
            Top             =   480
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtAvaluo_Notas 
            Height          =   1575
            Left            =   8160
            TabIndex        =   19
            Top             =   1200
            Width           =   4575
            _Version        =   1572864
            _ExtentX        =   8070
            _ExtentY        =   2778
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.RadioButton optPoliza 
            Height          =   255
            Index           =   0
            Left            =   5160
            TabIndex        =   20
            Top             =   1080
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Factor"
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
         End
         Begin XtremeSuiteControls.RadioButton optPoliza 
            Height          =   255
            Index           =   1
            Left            =   6360
            TabIndex        =   21
            Top             =   1080
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Personalizada"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtValorFiscal 
            Height          =   315
            Left            =   1920
            TabIndex        =   73
            Top             =   840
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCobertura 
            Height          =   330
            Left            =   1920
            TabIndex        =   82
            Top             =   2400
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777152
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCoberturaPorc 
            Height          =   330
            Left            =   3840
            TabIndex        =   83
            Top             =   2400
            Width           =   615
            _Version        =   1572864
            _ExtentX        =   1085
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   14737632
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaFormaliza 
            Height          =   315
            Left            =   5640
            TabIndex        =   88
            Top             =   1800
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaRstPlan 
            Height          =   315
            Left            =   5640
            TabIndex        =   89
            Top             =   2400
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   556
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboPE_Aseguradora 
            Height          =   330
            Left            =   -68080
            TabIndex        =   92
            Top             =   480
            Visible         =   0   'False
            Width           =   3255
            _Version        =   1572864
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.FlatEdit txtPE_Numero 
            Height          =   330
            Left            =   -68080
            TabIndex        =   94
            Top             =   960
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.PushButton btnAvaluo 
            Height          =   570
            Left            =   10560
            TabIndex        =   129
            Top             =   360
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1005
            _StockProps     =   79
            Caption         =   "Actualizar Avalúo"
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
            Picture         =   "frmCR_Prendas_New.frx":07A9
         End
         Begin XtremeSuiteControls.PushButton btnPolizaExterna 
            Height          =   570
            Left            =   -59440
            TabIndex        =   130
            Top             =   360
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   1005
            _StockProps     =   79
            Caption         =   "Registrar Póliza Externa"
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
            Picture         =   "frmCR_Prendas_New.frx":0EC9
         End
         Begin XtremeSuiteControls.DateTimePicker dtpPE_Inicia 
            Height          =   330
            Left            =   -68080
            TabIndex        =   131
            Top             =   2160
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.DateTimePicker dtpPE_Vence 
            Height          =   330
            Left            =   -68080
            TabIndex        =   133
            Top             =   2520
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.ComboBox cboPE_Frecuencia 
            Height          =   330
            Left            =   -68080
            TabIndex        =   137
            Top             =   1680
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
         Begin XtremeSuiteControls.FlatEdit txtPE_Cobertura 
            Height          =   735
            Left            =   -66160
            TabIndex        =   139
            Top             =   1200
            Visible         =   0   'False
            Width           =   8775
            _Version        =   1572864
            _ExtentX        =   15478
            _ExtentY        =   1296
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPE_Notas 
            Height          =   735
            Left            =   -66160
            TabIndex        =   141
            Top             =   2160
            Visible         =   0   'False
            Width           =   8775
            _Version        =   1572864
            _ExtentX        =   15478
            _ExtentY        =   1296
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkPE_Indica 
            Height          =   375
            Left            =   -62800
            TabIndex        =   144
            Top             =   480
            Visible         =   0   'False
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Utilizar Póliza Externa ?"
            BackColor       =   12640511
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.FlatEdit txtPE_Prima 
            Height          =   330
            Left            =   -68080
            TabIndex        =   135
            Top             =   1320
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
            Text            =   "0.00"
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   6
            Left            =   -66160
            TabIndex        =   142
            Top             =   1920
            Visible         =   0   'False
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaciones:"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   5
            Left            =   -66160
            TabIndex        =   140
            Top             =   960
            Visible         =   0   'False
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Detalle de la Cobertura:"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   21
            Left            =   -69760
            TabIndex        =   138
            Top             =   1680
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Frecuencia: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   16
            Left            =   -69760
            TabIndex        =   136
            Top             =   1320
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Prima: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   4
            Left            =   -69760
            TabIndex        =   134
            Top             =   2520
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Vence"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   3
            Left            =   -69760
            TabIndex        =   132
            Top             =   2160
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Inicia"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   13
            Left            =   -69760
            TabIndex        =   95
            Top             =   960
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No. Póliza: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   11
            Left            =   -69760
            TabIndex        =   93
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Aseguradora: "
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   2
            Left            =   5160
            TabIndex        =   91
            Top             =   1560
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza Formalización:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   90
            Top             =   2160
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Póliza Resto del Plan:"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cobertura"
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
            Left            =   120
            TabIndex        =   84
            Top             =   2400
            Width           =   1215
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   20
            Left            =   -63760
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Coberturas de Pólizas: "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label4 
            Height          =   255
            Index           =   19
            Left            =   -69880
            TabIndex        =   76
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Anotaciones: "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   840
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor Fiscal"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblValorPrenda 
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   2040
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor del Vehiculo"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   26
            Top             =   1560
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Extras"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Valor sin Extras"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha de Avalúo"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   14
            Left            =   8160
            TabIndex        =   23
            Top             =   960
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Observaciones:"
            ForeColor       =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label10 
            Height          =   255
            Index           =   18
            Left            =   5160
            TabIndex        =   22
            Top             =   720
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tipo de Póliza"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCoberturaTotal 
         Height          =   330
         Left            =   -59320
         TabIndex        =   86
         Top             =   6360
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbInfoNotarial 
         Height          =   1935
         Left            =   -70000
         TabIndex        =   96
         Top             =   360
         Visible         =   0   'False
         Width           =   12975
         _Version        =   1572864
         _ExtentX        =   22886
         _ExtentY        =   3413
         _StockProps     =   79
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit FlatEdit4 
            Height          =   315
            Left            =   8280
            TabIndex        =   119
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.PushButton btnNotariado 
            Height          =   375
            Left            =   5040
            TabIndex        =   103
            Top             =   1560
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Actualiza"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.FlatEdit txtNotario 
            Height          =   315
            Left            =   1320
            TabIndex        =   97
            Top             =   720
            Width           =   5295
            _Version        =   1572864
            _ExtentX        =   9340
            _ExtentY        =   556
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTomo 
            Height          =   315
            Left            =   1320
            TabIndex        =   98
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAsiento 
            Height          =   315
            Left            =   4560
            TabIndex        =   99
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit3 
            Height          =   315
            Left            =   8280
            TabIndex        =   118
            Top             =   720
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit5 
            Height          =   315
            Left            =   10320
            TabIndex        =   122
            Top             =   720
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit6 
            Height          =   315
            Left            =   10320
            TabIndex        =   123
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   13
            Left            =   7080
            TabIndex        =   125
            Top             =   720
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registro"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   12
            Left            =   6960
            TabIndex        =   124
            Top             =   1080
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Modificación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   11
            Left            =   10440
            TabIndex        =   121
            Top             =   480
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   10
            Left            =   8640
            TabIndex        =   120
            Top             =   480
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Usuario"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   1
            Left            =   0
            TabIndex        =   109
            Top             =   0
            Width           =   12855
            _Version        =   1572864
            _ExtentX        =   22675
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Información del registro Notarial:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   102
            Top             =   720
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Notario"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   101
            Top             =   1080
            Width           =   735
            _Version        =   1572864
            _ExtentX        =   1296
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Tomo"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   3240
            TabIndex        =   100
            Top             =   1080
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Folio/Asiento"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox gbTramite 
         Height          =   4815
         Left            =   -70000
         TabIndex        =   104
         Top             =   2280
         Visible         =   0   'False
         Width           =   12735
         _Version        =   1572864
         _ExtentX        =   22463
         _ExtentY        =   8493
         _StockProps     =   79
         ForeColor       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswTramite 
            Height          =   2055
            Left            =   0
            TabIndex        =   105
            Top             =   1200
            Width           =   12735
            _Version        =   1572864
            _ExtentX        =   22463
            _ExtentY        =   3625
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
            Appearance      =   21
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnNotasTramite 
            Height          =   450
            Left            =   11280
            TabIndex        =   106
            Top             =   600
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   794
            _StockProps     =   79
            Caption         =   "Notas"
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
            Picture         =   "frmCR_Prendas_New.frx":15E9
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   495
            Left            =   1320
            TabIndex        =   107
            Top             =   600
            Width           =   9735
            _Version        =   1572864
            _ExtentX        =   17171
            _ExtentY        =   873
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
            Height          =   315
            Left            =   4680
            TabIndex        =   110
            Top             =   3720
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.FlatEdit txtM_Usuario 
            Height          =   315
            Left            =   4680
            TabIndex        =   111
            Top             =   4080
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
            Height          =   315
            Left            =   6480
            TabIndex        =   112
            Top             =   3720
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.FlatEdit txtM_Fecha 
            Height          =   315
            Left            =   6480
            TabIndex        =   113
            Top             =   4080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   6
            Left            =   4800
            TabIndex        =   117
            Top             =   3480
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Usuario"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   7
            Left            =   6600
            TabIndex        =   116
            Top             =   3480
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3201
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha"
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
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   115
            Top             =   3720
            Width           =   3615
            _Version        =   1572864
            _ExtentX        =   6376
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registro General"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   9
            Left            =   1440
            TabIndex        =   114
            Top             =   4080
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Modificación General"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Index           =   0
            Left            =   0
            TabIndex        =   108
            Top             =   120
            Width           =   12855
            _Version        =   1572864
            _ExtentX        =   22675
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Observaciones del Trámite:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.PushButton btnHistoricos 
         Height          =   450
         Index           =   0
         Left            =   -70000
         TabIndex        =   127
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3413
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Avalúos"
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
         Checked         =   -1  'True
         Picture         =   "frmCR_Prendas_New.frx":1D09
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton btnHistoricos 
         Height          =   450
         Index           =   1
         Left            =   -68080
         TabIndex        =   128
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3413
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Pólizas  Externas"
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
         Picture         =   "frmCR_Prendas_New.frx":25B5
         ImageAlignment  =   0
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cobertura Total: "
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
         Left            =   -61120
         TabIndex        =   87
         Top             =   6360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeSuiteControls.Label lblRegistroUsuario 
         Height          =   315
         Left            =   -67600
         TabIndex        =   31
         Top             =   6840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblRegistroUsuarioAbog 
         Height          =   315
         Left            =   -62320
         TabIndex        =   30
         Top             =   6840
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   556
         _StockProps     =   79
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
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   37
         Left            =   -69040
         TabIndex        =   29
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   50
         Left            =   -63760
         TabIndex        =   28
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.Label lblPE_Status 
      Height          =   735
      Left            =   10080
      TabIndex        =   153
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
      _Version        =   1572864
      _ExtentX        =   4471
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Status Póliza Externa"
      ForeColor       =   16777215
      BackColor       =   8421631
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
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   1080
      Width           =   12855
      _Version        =   1572864
      _ExtentX        =   22675
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Registro de Prendas "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
      Alignment       =   1
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   9
      Top             =   240
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operación"
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   600
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identificación"
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
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   2
      Left            =   6120
      TabIndex        =   7
      Top             =   240
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Expediente"
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
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
End
Attribute VB_Name = "frmCR_Prendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mPrendaId As Long

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vPaso As Boolean
Dim vEdita As Integer, mFecha As Date



Private Sub btnAdjuntos_Click()

If mPrendaId = 0 Then
    MsgBox "Consulte una prenda para poder registrar documentos adjuntos!", vbInformation
    Exit Sub
End If

 gGA.Modulo = "CR_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = "P-" & mPrendaId
 gGA.Llave_03 = "" 'txtCodigo.Text
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnAvaluo_Click()
Dim i As Integer

On Error GoTo vError

If mPrendaId = 0 Then
    MsgBox "Consulte una Prenda primero!", vbExclamation
    Exit Sub
End If


If Not IsNumeric(txtValorFiscal.Text) Then
    MsgBox "El Dato del valor Fiscal no es correcto!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtValorSExtras.Text) Then
    MsgBox "El Dato del valor sin Extras no es correcto!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtExtras.Text) Then
    MsgBox "El Dato de las Extras no es correcto!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtValorTotal.Text) Then
    MsgBox "El Dato del valor Final de la Prenda no es correcto!", vbExclamation
    Exit Sub
End If


If txtPE_Numero.Text = "" Then
    MsgBox "Digite el número de la póliza!", vbExclamation
    Exit Sub
End If


txtAvaluo_Notas.Text = fxSysCleanTxtInject(txtAvaluo_Notas.Text)



i = MsgBox("Esta Seguro que desea actualizar el avalúo", vbYesNo)
If i = vbNo Then
    Exit Sub
End If


Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Operacion_Prenda_Avaluo " & mPrendaId & ", '', " & CCur(txtValorTotal.Text) _
       & ", " & CCur(txtCobertura.Text) & ", " & CCur(txtCoberturaPorc.Text) _
       & ", '" & txtAvaluo_Notas.Text & "', '" & Format(dtpFechaInspeccion.Value, "yyyy-mm-dd") _
       & "', " & CCur(txtValorFiscal.Text) & ", " & CCur(txtValorTotal.Text) & ", " & CCur(txtExtras.Text) _
       & ",  " & IIf((optPoliza(0).Value = True), 1, 0) & ", " & CCur(txtPolizaFormaliza.Text) _
       & ", " & CCur(txtPolizaRstPlan.Text) & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)


Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Avalúo registrado satisfactoriamente!", vbInformation
    
   Call sbGarantia_Load
    
Else
   MsgBox rs!Mensaje, vbExclamation
End If


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnNotariado_Click()
Dim i As Integer

On Error GoTo vError

If mPrendaId = 0 Then
    MsgBox "Consulte una Prenda primero!", vbExclamation
    Exit Sub
End If

If txtNotario.Text = "" Then
    MsgBox "Indique el nombre del Notario!", vbExclamation
    Exit Sub
End If

If txtAsiento.Text = "" Then
    MsgBox "Indique el Folio/Asiento de la escritura!", vbExclamation
    Exit Sub
End If

If txtTomo.Text = "" Then
    MsgBox "Indique el Tomo de la escritura", vbExclamation
    Exit Sub
End If



i = MsgBox("Esta Seguro que desea actualizar la información notarial de la prenda?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If



Me.MousePointer = vbHourglass

txtNotario.Text = fxSysCleanTxtInject(txtNotario.Text)
txtTomo.Text = fxSysCleanTxtInject(txtTomo.Text)
txtAsiento.Text = fxSysCleanTxtInject(txtAsiento.Text)

strSQL = "exec spCrd_Operacion_Prenda_Notariado " & mPrendaId & ", '" & txtNotario.Text _
       & "', '" & txtTomo.Text & "', '" & txtAsiento.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)


Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Información de notariado actualizada satisfactoriamente!", vbInformation
    
   Call sbGarantia_Load
    
Else
   MsgBox rs!Mensaje, vbExclamation
End If


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnPolizaExterna_Click()
Dim i As Integer

On Error GoTo vError

If mPrendaId = 0 Then
    MsgBox "Consulte una Prenda primero!", vbExclamation
    Exit Sub
End If

If dtpPE_Vence.Value < dtpPE_Inicia.Value Then
    MsgBox "Verifique las Fechas de Cobertura!", vbExclamation
    Exit Sub
End If

If Not IsNumeric(txtPE_Prima.Text) Then
    MsgBox "El Dato de la Prima no es correcto!", vbExclamation
    Exit Sub
End If

If txtPE_Numero.Text = "" Then
    MsgBox "Digite el número de la póliza!", vbExclamation
    Exit Sub
End If


i = MsgBox("Esta Seguro que desea actualizar la Póliza Externa?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If


If dtpPE_Vence.Value < dtpPE_Inicia.Value Then
    MsgBox "Verifique las Fechas de Cobertura!", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

txtPE_Cobertura.Text = fxSysCleanTxtInject(txtPE_Cobertura.Text)
txtPE_Notas.Text = fxSysCleanTxtInject(txtPE_Notas.Text)
txtPE_Numero.Text = fxSysCleanTxtInject(txtPE_Numero.Text)

'spCrd_Operacion_Prenda_Poliza_Externa_Registra(@PrendaId int, @AseguradoraId int, @NumeroPoliza varchar(50), @Prima dec(14,2)
'        , @Frecuencia varchar(20), @Inicio datetime, @Corte datetime, @Activa smallint, @PolizaIndica smallint
'        , @Cobertura varchar(3000), @Notas varchar(3000)
'        , @Usuario varchar(30) )

strSQL = "exec spCrd_Operacion_Prenda_Poliza_Externa_Registra " & mPrendaId & ", " & cboPE_Aseguradora.ItemData(cboPE_Aseguradora.ListIndex) _
       & ", '" & txtPE_Numero.Text & "', " & CCur(txtPE_Prima.Text) & ", '" & cboPE_Frecuencia.Text _
       & "', '" & Format(dtpPE_Inicia.Value, "yyyy-mm-dd") & "', '" & Format(dtpPE_Vence.Value, "yyyy-mm-dd") _
       & "', " & chkPE_Activa.Value & ", " & chkPE_Indica.Value _
       & ", '" & txtPE_Cobertura.Text & "', '" & txtPE_Notas.Text & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)


Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Póliza Externa " & rs!Movimiento & " satisfactoriamente!", vbInformation
    
   Call sbGarantia_Load
    
Else
   MsgBox rs!Mensaje, vbExclamation
End If


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select Formulario, Porc_Cobertura " _
       & " from crd_prendas_tipos" _
       & " Where Tipo_Prenda = '" & cboTipo.ItemData(cboTipo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)

txtCoberturaPorc.Text = Format(rs!Porc_Cobertura, "Standard")
 
If rs!Formulario = "Vehículo" Then
    lblValorPrenda.Caption = "Valor del Vehículo"
    tcGarantia.Item(1).Selected = True
Else
    tcGarantia.Item(0).Selected = True
    lblValorPrenda.Caption = "Valor del Bien"
End If

'Cargar las Polizas Acá
lswPolizas.ListItems.Clear

vPaso = True

strSQL = "exec spCrd_Prendas_Polizas_List '" & cboTipo.ItemData(cboTipo.ListIndex) & "', " & mPrendaId
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswPolizas.ListItems.Add(, , rs!Cobertura)
      itmX.Checked = IIf((rs!asignado = 1), True, False)
      itmX.Tag = rs!ID_PRENDA_COBERTURA
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault



End Sub



Private Sub cboV_Comercializa_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
 gBusquedas.Col1Name = "Modelo Id"
 gBusquedas.Col2Name = "Descripción"
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select ID_COMERCIO, DESCRIPCION from CRD_PRENDAS_COMERCIA"
 gBusquedas.Filtro = " AND ACTIVA = 1"
 frmBusquedas.Show vbModal
 If gBusquedas.Resultado <> "" Then
    Call sbCboAsignaDato(cboV_Comercializa, gBusquedas.Resultado2, True, gBusquedas.Resultado)
 End If

End If
End Sub

Private Sub cboV_Marca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Col1Name = "Marca Id"
 gBusquedas.Col2Name = "Descripción"
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select ID_MARCA, DESCRIPCION from CRD_PRENDAS_MARCAS"
 gBusquedas.Filtro = " AND ACTIVA = 1"
 frmBusquedas.Show vbModal
 If gBusquedas.Resultado <> "" Then
    Call sbCboAsignaDato(cboV_Marca, gBusquedas.Resultado2, True, gBusquedas.Resultado)
 End If
End If
End Sub

Private Sub cboV_Modelo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
 gBusquedas.Col1Name = "Modelo Id"
 gBusquedas.Col2Name = "Descripción"
 gBusquedas.Columna = "Descripcion"
 gBusquedas.Orden = "Descripcion"
 gBusquedas.Consulta = "select ID_MODELO, DESCRIPCION from CRD_PRENDAS_MODELOS"
 gBusquedas.Filtro = " AND ACTIVO = 1"
 frmBusquedas.Show vbModal
 If gBusquedas.Resultado <> "" Then
    Call sbCboAsignaDato(cboV_Modelo, gBusquedas.Resultado2, True, gBusquedas.Resultado)
 End If

End If

End Sub

Private Sub Form_Load()

vModulo = 3

scMain.Caption = "Registro de Garantías Prendarias"

txtOperacion.Text = Operacion.Operacion
txtExpediente.Text = Operacion.Expendiente

txtCedula.Text = Operacion.Cedula
txtNombre.Text = fxNombre(txtCedula.Text)


With lsw.ColumnHeaders
    .Clear
    .Add , , "Prenda Id", 1000
    .Add , , "Categoria", 2500
    .Add , , "Avaluo", 1800, vbRightJustify
    .Add , , "%", 1800, vbRightJustify
    .Add , , "Cobertura", 1800, vbRightJustify
    .Add , , "Descripción", 2500
    .Add , , "Id Principal", 1500, vbCenter
    .Add , , "Id Provisional", 1500, vbCenter
    .Add , , "Modelo", 2500
    .Add , , "Serie", 2500
    .Add , , "Marca", 2500
    .Add , , "Año", 900, vbCenter
    .Add , , "Reg.Fecha", 2100
    .Add , , "Reg.Usuario", 2100
    .Add , , "Tomo", 1400, vbCenter
    .Add , , "Folio", 1400, vbCenter
    .Add , , "Notario", 3400
    .Add , , "Notario Fecha", 2100
End With


With lswTramite.ColumnHeaders
    .Clear
    .Add , , "Tramite", 3000
    .Add , , "Notas", 3000
    .Add , , "Usuario", 1900, vbCenter
    .Add , , "Fecha", 1900, vbCenter
End With


With lswPolizas.ColumnHeaders
  .Clear
  .Add , , "Coberturas de Pólizas", lswPolizas.Width - 250
End With

Call sbToolBarIconos(tlbPrincipal, False)

With tlbPrincipal
    .Buttons(1).Enabled = True
    .Buttons(2).Enabled = False
    .Buttons(3).Enabled = False
    .Buttons(4).Enabled = False
    .Buttons(5).Enabled = False
End With

mPrendaId = 0
mFecha = fxFechaServidor


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()
 
 vPaso = True
  
 txtId_01.Text = ""
 txtId_02.Text = ""
 
 txtDescripcion.Text = ""
 txtModelo.Text = ""
 txtSerie.Text = ""
 txtMarca.Text = ""
 txtColor.Text = ""
 txtAnio.Text = Year(mFecha)
 
 txtObservaciones.Text = ""
 
 txtCoberturaPorc.Text = "0"
 txtCobertura.Text = "0"
 
 txtV_PlacaRegistral.Text = ""
 txtV_PlacaProvisional.Text = ""
  
 txtV_Chasis.Text = ""
 txtV_VIN.Text = ""
 
 txtV_Anio.Text = Year(mFecha)
 txtV_Capacidad.Text = "5"
 txtV_Cilindraje.Text = "1200"
 txtV_Peso.Text = "1000"
 txtV_Color.Text = ""
 txtV_Puertas.Text = "1"
 
 txtValorFiscal.Text = "0"
 txtValorSExtras.Text = "0"
 txtExtras.Text = "0"
 
 txtValorTotal.Text = "0"
 txtCobertura.Text = "0"
 
 dtpFechaInspeccion.Value = mFecha
  
 txtAvaluo_Notas.Text = ""
 txtPolizaFormaliza.Text = "0"
 txtPolizaRstPlan.Text = "0"
 
 
 lblPE_Status.Visible = False
 
 chkPE_Indica.Value = vbUnchecked
 dtpPE_Inicia.Value = mFecha
 dtpPE_Vence.Value = mFecha
 cboPE_Frecuencia.Text = "Mensual"
 
 chkPE_Activa.Value = vbUnchecked
    
 txtPE_Prima.Text = "0"
 txtPE_Numero.Text = ""
 txtPE_Notas.Text = ""
 txtPE_Cobertura.Text = ""
 
 mPrendaId = 0
 
 vPaso = False
 
 Call cboTipo_Click

 
End Sub

Private Sub sbPrendas_List_Load()

Dim curTotal As Currency, pOperacion As Long

curTotal = 0

If Not IsNumeric(txtOperacion.Text) Then
    pOperacion = 0
Else
    pOperacion = txtOperacion.Text
End If


tcMain.Item(0).Selected = True
strSQL = "exec spCrd_Prendas_List " & pOperacion & ", '" & txtExpediente.Text & "'"

Call OpenRecordSet(rs, strSQL)
lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Prenda_Id)
      
      itmX.SubItems(1) = rs!TIPO_PRENDA_DESC
      
      itmX.SubItems(2) = Format(rs!AVALUO, "Standard")
      itmX.SubItems(3) = Format(rs!Porc_Cobertura, "Standard")
      itmX.SubItems(4) = Format(rs!Cobertura, "Standard")
      
      itmX.SubItems(5) = rs!Descripcion & ""
      itmX.SubItems(6) = rs!ID_PRINCIPAL & ""
      itmX.SubItems(7) = rs!ID_PROVISIONAL & ""
      
      itmX.SubItems(8) = rs!Modelo & ""
      itmX.SubItems(9) = rs!Serie
      itmX.SubItems(10) = rs!Marca
      itmX.SubItems(11) = rs!Anio & ""
      itmX.SubItems(11) = rs!Registro_Fecha & ""
      itmX.SubItems(12) = rs!Registro_Usuario & ""
      itmX.SubItems(13) = rs!Tomo & ""
      itmX.SubItems(14) = rs!Folio & ""
      itmX.SubItems(13) = rs!Notario & ""
      itmX.SubItems(14) = rs!NOTARIO_REGISTRO_FECHA & ""
      
      
      curTotal = curTotal + rs!Cobertura
 rs.MoveNext
Loop
rs.Close

txtCoberturaTotal.Text = Format(curTotal, "Standard")

End Sub



Private Sub sbGarantia_Load()
Dim rs As New ADODB.Recordset

On Error GoTo vError

 strSQL = "exec spCrd_Prendas_Garantia_Load " & mPrendaId
 Call OpenRecordSet(rs, strSQL)
 
 tcMain.Item(1).Selected = True
 
 txtDescripcion.Text = rs!Descripcion & ""
 txtModelo.Text = rs!Modelo & ""
 txtSerie.Text = rs!Serie & ""
 txtMarca.Text = rs!Marca & ""
 
 txtValorFiscal.Text = Format(rs!AVALUO, "Standard")
 
 txtCoberturaPorc.Text = Format(rs!Porc_Cobertura, "Standard")
 txtCobertura.Text = Format(rs!Cobertura, "Standard")
 
 vPaso = True
     Call sbCboAsignaDato(cboTipo, rs!TIPO_PRENDA_DESC, True, Trim(rs!Tipo_Prenda))
 vPaso = False
 Call cboTipo_Click

 If Not IsNull(rs!ID_COMBUSTIBLE) Then
     Call sbCboAsignaDato(cboV_Combustible, rs!COBUSTIBLE_DESC, True, rs!ID_COMBUSTIBLE)
 End If

 If Not IsNull(rs!ID_COMERCIO) Then
     Call sbCboAsignaDato(cboV_Comercializa, rs!COMERCIALIZA_DESC, True, rs!ID_COMERCIO)
 End If

 If Not IsNull(rs!ID_MARCA) Then
     Call sbCboAsignaDato(cboV_Marca, rs!MARCA_DESC, True, rs!ID_MARCA)
 End If

 If Not IsNull(rs!ID_MODELO) Then
     Call sbCboAsignaDato(cboV_Modelo, rs!MODELO_DESC, True, rs!ID_MODELO)
 End If

 If Not IsNull(rs!ID_PRESENTACION) Then
     Call sbCboAsignaDato(cboV_Presentacion, rs!PRESENTACION_DESC, True, rs!ID_PRESENTACION)
 End If
 
 'Unidades de Medidas
 If Not IsNull(rs!CILINDRAJE_UD) Then
     Call sbCboAsignaDato(cboUd_Cilindraje, rs!CILINDRAJE_UD_DESC, True, rs!CILINDRAJE_UD)
 End If
 
 If Not IsNull(rs!PESO_UD) Then
     Call sbCboAsignaDato(cboUd_Peso, rs!PESO_UD_DESC, True, rs!PESO_UD)
 End If
 
 If Not IsNull(rs!CAPACIDAD_UD) Then
     Call sbCboAsignaDato(cboUd_Capacidad, rs!CAPACIDAD_UD_DESC, True, rs!CAPACIDAD_UD)
 End If
 
 
 
 ' Pg.CILINDRAJE_UD, Pg.PESO_UD, Pg.CAPACIDAD_UD

 txtV_Anio.Text = rs!Anio & ""
 txtV_Color.Text = rs!Color & ""
 
 txtV_PlacaRegistral.Text = rs!ID_PRINCIPAL & ""
 txtV_PlacaProvisional.Text = rs!ID_PROVISIONAL & ""
 
 txtV_Chasis.Text = rs!CHASIS_NUMERO & ""
 txtV_VIN.Text = rs!VIN_MOTOR & ""
 
 txtV_Puertas.Text = rs!PUERTAS_NUMERO & ""
 txtV_Peso.Text = rs!Peso & ""
 txtV_Capacidad.Text = rs!Capacidad & ""
 txtV_Cilindraje.Text = rs!Cilindraje & ""
 
 txtR_Fecha.Text = rs!Registro_Fecha & ""
 txtR_Usuario.Text = rs!Registro_Usuario & ""

 txtTomo.Text = rs!Tomo & ""
 txtAsiento.Text = rs!Folio & ""

 txtNotario.Text = rs!Notario & ""
 
 
 ', Pg.VALOR_MERCADO, Pg.AVALUO
 
 dtpFechaInspeccion.Value = rs!AVALUO_INSPECCION
 
 txtValorFiscal.Text = Format(rs!VALOR_FISCAL, "Standard")
 txtValorSExtras.Text = Format(rs!AVALUO - rs!Monto_Extras, "Standard")
 txtExtras.Text = Format(rs!Monto_Extras, "Standard")
 txtValorTotal.Text = Format(rs!AVALUO, "Standard")
 txtAvaluo_Notas.Text = rs!AVALUO_OBSERVACION & ""
 
 txtPolizaFormaliza.Text = Format(rs!POLIZA_MNT_FORMALIZACION, "Standard")
 txtPolizaRstPlan.Text = Format(rs!POLIZA_RST_PLAN, "Standard")
 
 'Poliza Externa
 chkPE_Indica.Value = rs!PE_INDICA
 If chkPE_Indica.Value = xtpChecked Then
    
    dtpPE_Inicia.Value = rs!PE_INICIO
    dtpPE_Vence.Value = rs!PE_VENCE
    
    cboPE_Frecuencia.Text = rs!PE_FRECUENCIA
    chkPE_Activa.Value = rs!PE_ACTIVA
    
    txtPE_Prima.Text = Format(rs!PE_PRIMA, "Standard")
    txtPE_Numero.Text = rs!PE_NUMERO
    txtPE_Notas.Text = rs!PE_NOTAS
    
    txtPE_Cobertura.Text = rs!PE_Cobertura
    
    Call sbCboAsignaDato(cboPE_Aseguradora, rs!ASEGURADORA_DESC, True, rs!ID_ASEGURADORA)
    
    lblPE_Status.Visible = True
    If rs!PE_VENCIDA = 1 Then
        lblPE_Status.BackColor = RGB(246, 95, 133)
        lblPE_Status.ForeColor = vbWhite
        lblPE_Status.Caption = "Utiliza Póliza Externa y se encuentra Vencida!"
    Else
        lblPE_Status.BackColor = RGB(236, 243, 157)
        lblPE_Status.ForeColor = RGB(246, 95, 133)
        lblPE_Status.Caption = "Utiliza Póliza Externa!"
    End If

 Else
   lblPE_Status.Visible = False
 
 End If


 rs.Close
 
With tlbPrincipal
   .Buttons(1).Enabled = False
   .Buttons(2).Enabled = True
   .Buttons(3).Enabled = True
   .Buttons(4).Enabled = False
   .Buttons(5).Enabled = False
End With


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

 mPrendaId = Item.Text
 
 Call sbGarantia_Load

End Sub



Private Sub sbInicializa()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select rtrim(tipo_prenda) as 'IdX', rtrim(descripcion) as 'ItmX'" _
       & " from crd_prendas_tipos where Activa = 1 order by descripcion "
vPaso = True
    Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
vPaso = False
 
cboPE_Frecuencia.Clear
cboPE_Frecuencia.AddItem "Mensual"
cboPE_Frecuencia.AddItem "Trimestral"
cboPE_Frecuencia.AddItem "Semestral"
cboPE_Frecuencia.AddItem "Anual"
cboPE_Frecuencia.Text = "Mensual"
 
cboV_Uso.Clear
cboV_Uso.AddItem "PERSONAL"
cboV_Uso.AddItem "TRABAJO"
cboV_Uso.Text = "PERSONAL"

strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'ASE'"
Call sbCbo_Llena_New(cboPE_Aseguradora, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'MAR'"
Call sbCbo_Llena_New(cboV_Marca, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'PRE'"
Call sbCbo_Llena_New(cboV_Presentacion, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'MOD'"
Call sbCbo_Llena_New(cboV_Modelo, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'COB'"
Call sbCbo_Llena_New(cboV_Combustible, strSQL, False, True)
 
strSQL = "exec spCrd_Prendas_Cat_List_Cbo 'COM'"
Call sbCbo_Llena_New(cboV_Comercializa, strSQL, False, True)
  
  
'Unidades
strSQL = "select rtrim(ID_Unidad) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  CRD_PRENDAS_uds Where Peso_Apl = 1 and Activa = 1 order by Descripcion"
Call sbCbo_Llena_New(cboUd_Peso, strSQL, False, True)

strSQL = "select rtrim(ID_Unidad) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  CRD_PRENDAS_uds Where Capacidad_Apl = 1 and Activa = 1 order by Descripcion"
Call sbCbo_Llena_New(cboUd_Capacidad, strSQL, False, True)

strSQL = "select rtrim(ID_Unidad) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  CRD_PRENDAS_uds Where Cilindraje_Apl = 1 and Activa = 1 order by Descripcion"
Call sbCbo_Llena_New(cboUd_Cilindraje, strSQL, False, True)
 
 
Call sbLimpia
  
Me.MousePointer = vbDefault
 
Call sbPrendas_List_Load


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswPolizas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub
If mPrendaId = 0 Then Exit Sub

On Error GoTo vError

strSQL = "exec spCrd_Prendas_Polizas_Add " & mPrendaId & ", " & Item.Tag & ", '" & glogon.Usuario _
       & "', " & IIf((Item.Checked), 1, 0)
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

On Error GoTo vError

If Operacion.GarantiaId > 0 Then
    mPrendaId = Operacion.GarantiaId
    Call sbGarantia_Load
End If

Exit Sub

vError:

End Sub


'Private Sub txtAvaluo_GotFocus()
'On Error GoTo vError
'    txtAvaluo.Text = CCur(txtAvaluo.Text)
'vError:
'
'End Sub
'
'Private Sub txtAvaluo_LostFocus()
'On Error GoTo vError
'    txtAvaluo.Text = Format(CCur(txtAvaluo.Text), "Standard")
'vError:
'End Sub
'
'
'Private Sub txtAvaluo_KeyPress(KeyAscii As Integer)
'On Error GoTo vError
'  If KeyAscii = vbKeyReturn Then txtCoberturaPorc.SetFocus
'vError:
'End Sub

Private Function fxVerificaDatos() As Boolean
Dim vMensaje As String

fxVerificaDatos = True
vMensaje = ""

'Revision de Inyección

If tcGarantia.SelectedItem = 0 Then
    txtDescripcion.Text = fxSysCleanTxtInject(txtDescripcion.Text)
    txtModelo.Text = fxSysCleanTxtInject(txtModelo.Text)
    txtSerie.Text = fxSysCleanTxtInject(txtSerie.Text)
    txtMarca.Text = fxSysCleanTxtInject(txtMarca.Text)
    
    If Len(Trim(txtDescripcion.Text)) < 10 Then vMensaje = vMensaje & vbCrLf & "- La descripción no es válida"
    If Len(Trim(txtModelo.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Modelo no es válido"
    If Len(Trim(txtSerie.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- Indique el número de serie"
    If Len(Trim(txtMarca.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- Indique la Marca"
    If Len(Trim(txtAnio.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- Indique el año de fabricación!"
    If Len(Trim(txtId_01.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Id Principal no es válido"
    If Len(Trim(txtColor.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Color no es válido"

Else
    If Len(Trim(txtV_VIN.Text)) < 10 Then vMensaje = vMensaje & vbCrLf & "- Indique el número de VIN del Motor"
    If Len(Trim(txtV_Chasis.Text)) < 10 Then vMensaje = vMensaje & vbCrLf & "- Indique El Número de Chasis"
    If Len(Trim(txtV_Anio.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- Indique el año de fabricación!"
    If Len(Trim(txtV_PlacaRegistral.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Id Principal no es válido"
    If Len(Trim(txtV_Color.Text)) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Color no es válido"

    If Not IsNumeric(txtV_Peso.Text) Then
       vMensaje = vMensaje & vbCrLf & "- El dato del Peso es erroneo!"
    End If
    
    If Not IsNumeric(txtV_Capacidad.Text) Then
       vMensaje = vMensaje & vbCrLf & "- El dato de la Capacidad es erroneo!"
    End If
    
    If Not IsNumeric(txtV_Cilindraje.Text) Then
       vMensaje = vMensaje & vbCrLf & "- El dato del cilindraje es erroneo!"
    End If

End If


If Not IsNumeric(txtValorTotal.Text) Then
   vMensaje = vMensaje & vbCrLf & "- El dato del Avalúo es erroneo!"
End If


If Len(vMensaje) > 0 Then
  fxVerificaDatos = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuarda()

On Error GoTo vError
        
'spCrd_Operacion_Prenda_Registro(@PrendaId int, @Operacion int, @Expediente varchar(30), @Cedula varchar(20), @TipoPoliza varchar(10)
'            , @IdPrincipal varchar(30), @IdProvisional varchar(30), @Descripcion varchar(1000), @Observacion varchar(1000)
'            , @Marca varchar(100), @Modelo varchar(100), @Serie varchar(100), @Color varchar(20), @Anio int
'            , @Peso dec(12,4), @Capacidad dec(12,4), @Cilindraje dec(12,4), @PuertasNo smallint, @Chasis varchar(50), @VinMotor varchar(50)
'            , @IdMarca int, @IdModelo int, @IdPresentacion int, @IdCombustible int, @IdComercio int
'            , @UdPeso varchar(10), @UdCapacidad varchar(10), @UdCilindraje varchar(10)
'            , @Avaluo dec(16,2), @CoberturaPorc dec(10,2), @Cobertura dec(16,2), @AvaluoNotas varchar(1000), @AvaluoFecha datetime
'            , @ValorFiscal dec(16,2), @ValorMercado dec(16,2), @Extras dec(16,2)
'            , @PolizaFactorApl smallint, @PolizaFormaliza dec(14,2), @PolizaPlan dec(14,2)
'            , @Usuario varchar(30) )
'
If tcGarantia.SelectedItem = 0 Then
    'General
    strSQL = "exec spCrd_Operacion_Prenda_Registro " & mPrendaId & ", " & txtOperacion.Text & ", '" & txtExpediente.Text & "', '" & txtCedula.Text _
           & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', '" & txtId_01.Text & "', '" & txtId_02.Text _
           & "', '" & txtDescripcion.Text & "', '" & txtObservaciones.Text & "', '" & txtMarca.Text & "', '" & txtModelo.Text & "', '" & txtSerie.Text & "', '" & txtColor.Text _
           & "', " & txtAnio.Text & ", 0, 0, 0 , 0, '', '', " _
           & "Null, Null, Null , Null , Null, '" & cboUd_Peso.ItemData(cboUd_Peso.ListIndex) & "', '" & cboUd_Capacidad.ItemData(cboUd_Capacidad.ListIndex) & "', '" & cboUd_Cilindraje.ItemData(cboUd_Cilindraje.ListIndex) _
           & "', " & CCur(txtValorTotal.Text) & ", " & CCur(txtCoberturaPorc.Text) & ", " & CCur(txtCobertura.Text) & ", '" & txtAvaluo_Notas.Text _
           & "', '" & Format(dtpFechaInspeccion.Value, "yyyy-mm-dd") & "', " & CCur(txtValorFiscal.Text) & ", " & CCur(txtValorTotal.Text) _
           & ", " & CCur(txtExtras.Text) & ",  " & IIf((optPoliza(0).Value = True), 1, 0) & ", " & CCur(txtPolizaFormaliza.Text) _
           & ", " & CCur(txtPolizaRstPlan.Text) & ", '" & glogon.Usuario & "'"


Else
    'Vehicular
    strSQL = "exec spCrd_Operacion_Prenda_Registro " & mPrendaId & ", " & txtOperacion.Text & ", '" & txtExpediente.Text & "', '" & txtCedula.Text _
           & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', '" & txtV_PlacaRegistral.Text & "', '" & txtV_PlacaProvisional.Text _
           & "', '" & cboV_Uso.Text & "', '" & txtObservaciones.Text & "', '" & cboV_Marca.Text & "', '" & cboV_Modelo.Text & "', '', '" & txtV_Color.Text _
           & "', " & txtV_Anio.Text & ", " & CCur(txtV_Peso.Text) & ", " & CCur(txtV_Capacidad.Text) & ", " & CCur(txtV_Cilindraje.Text) _
           & ", " & txtV_Puertas.Text & ", '" & txtV_Chasis.Text & "', '" & txtV_VIN.Text _
           & "', " & cboV_Marca.ItemData(cboV_Marca.ListIndex) & ", " & cboV_Modelo.ItemData(cboV_Modelo.ListIndex) & ", " & cboV_Presentacion.ItemData(cboV_Presentacion.ListIndex) _
           & ", " & cboV_Combustible.ItemData(cboV_Combustible.ListIndex) & ", " & cboV_Comercializa.ItemData(cboV_Comercializa.ListIndex) _
           & ", '" & cboUd_Peso.ItemData(cboUd_Peso.ListIndex) & "', '" & cboUd_Capacidad.ItemData(cboUd_Capacidad.ListIndex) & "', '" & cboUd_Cilindraje.ItemData(cboUd_Cilindraje.ListIndex) _
           & "', " & CCur(txtValorTotal.Text) & ", " & CCur(txtCoberturaPorc.Text) & ", " & CCur(txtCobertura.Text) & ", '" & txtAvaluo_Notas.Text _
           & "', '" & Format(dtpFechaInspeccion.Value, "yyyy-mm-dd") & "', " & CCur(txtValorFiscal.Text) & ", " & CCur(txtValorTotal.Text) _
           & ", " & CCur(txtExtras.Text) & ",  " & IIf((optPoliza(0).Value = True), 1, 0) & ", " & CCur(txtPolizaFormaliza.Text) _
           & ", " & CCur(txtPolizaRstPlan.Text) & ", '" & glogon.Usuario & "'"
     
End If


Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
    
   mPrendaId = rs!PrendaId

   MsgBox "Se ha " & rs!Movimiento & " satisfactoriamente, la prenda Id: " & mPrendaId, vbInformation
    
   Call sbGarantia_Load
    
Else
   MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim iRespuesta As Integer

Select Case Button.Key
  Case "insertar", "nuevo"
   vEdita = 0
   Call sbLimpia
    
   tcMain.Item(1).Selected = True
    
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
'    fra.Enabled = True
    cboTipo.SetFocus
  
  Case "editar", "modificar"
   vEdita = 1
    With tlbPrincipal
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = True
       .Buttons(5).Enabled = True
    End With
'    fra.Enabled = True
    cboTipo.SetFocus
  
  Case "borrar"
   
   If mPrendaId > 0 Then
    iRespuesta = MsgBox("Esta seguro que desea eliminar esta prenda?", vbYesNo)
    
    If iRespuesta = vbYes Then
    
      strSQL = "exec spCrd_Operacion_Prenda_Elimina " & mPrendaId & ", '" & glogon.Usuario & "'"
      Call ConectionExecute(strSQL)
      Call sbPrendas_List_Load
      Call sbLimpia
    Else
      Call sbLimpia
    End If
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
     End With
    
   End If
  
  Case "salvar", "guardar"
    If fxVerificaDatos Then
      Call sbGuarda
      
      Call sbPrendas_List_Load
      
      With tlbPrincipal
        .Buttons(1).Enabled = True
        .Buttons(2).Enabled = False
        .Buttons(3).Enabled = False
        .Buttons(4).Enabled = False
        .Buttons(5).Enabled = False
      End With
      
      Call sbLimpia
    
    Else
      MsgBox "Información Ingresada es Incorrecta por favor verifique...", vbInformation
    End If
  
  Case "deshacer"
    Call sbLimpia
    
    With tlbPrincipal
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       .Buttons(3).Enabled = False
       .Buttons(4).Enabled = False
       .Buttons(5).Enabled = False
    End With
  
  Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
        
  Case "salir", "cerrar"
    Unload Me

End Select

End Sub



Private Sub txtPE_Prima_GotFocus()
On Error GoTo vError
 txtPE_Prima.Text = CCur(txtPE_Prima.Text)
vError:
End Sub

Private Sub txtPE_Prima_LostFocus()
On Error GoTo vError
 txtPE_Prima.Text = Format(CCur(txtPE_Prima.Text), "Standard")
vError:
End Sub


Private Sub txtPolizaFormaliza_GotFocus()
On Error GoTo vError
 txtPolizaFormaliza.Text = CCur(txtPolizaFormaliza.Text)
vError:
End Sub

Private Sub txtPolizaFormaliza_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 
 
 txtPolizaFormaliza.Text = Format(CCur(txtPolizaFormaliza.Text), "Standard")

vError:
End Sub

Private Sub txtPolizaFormaliza_LostFocus()
On Error GoTo vError
 txtPolizaFormaliza.Text = Format(CCur(txtPolizaFormaliza.Text), "Standard")
vError:
End Sub

Private Sub txtPolizaRstPlan_GotFocus()
On Error GoTo vError
 txtPolizaRstPlan.Text = CCur(txtPolizaRstPlan.Text)
vError:
End Sub

Private Sub txtPolizaRstPlan_LostFocus()
On Error GoTo vError
 txtPolizaRstPlan.Text = Format(CCur(txtPolizaRstPlan.Text), "Standard")
vError:
End Sub

Private Sub txtValorFiscal_GotFocus()
On Error GoTo vError
 txtValorFiscal.Text = CCur(txtValorFiscal.Text)
vError:

End Sub

Private Sub txtValorFiscal_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

Dim pValor As Currency

 pValor = CCur(txtValorFiscal.Text) + CCur(txtExtras.Text)
 
 txtValorTotal.Text = Format(pValor, "Standard")
 txtCobertura.Text = Format(pValor * CCur(txtCoberturaPorc.Text) / 100, "Standard")
 
vError:

End Sub

Private Sub txtValorFiscal_LostFocus()
On Error GoTo vError
 txtValorFiscal.Text = Format(CCur(txtValorFiscal.Text), "Standard")
vError:

End Sub
