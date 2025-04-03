VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_PeriodosGracia 
   Caption         =   "Periodos de Gracia"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcFiltros 
      Height          =   3732
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   14292
      _Version        =   1441792
      _ExtentX        =   25209
      _ExtentY        =   6583
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
      Item(0).Caption =   "General"
      Item(0).ControlCount=   31
      Item(0).Control(0)=   "Label1(37)"
      Item(0).Control(1)=   "Label1(13)"
      Item(0).Control(2)=   "Label1(15)"
      Item(0).Control(3)=   "Label1(18)"
      Item(0).Control(4)=   "cboDestino"
      Item(0).Control(5)=   "cboRecurso"
      Item(0).Control(6)=   "cboInstitucion"
      Item(0).Control(7)=   "cboDeductora"
      Item(0).Control(8)=   "cboGarantia"
      Item(0).Control(9)=   "chkLineas"
      Item(0).Control(10)=   "txtCodigo"
      Item(0).Control(11)=   "txtDescripcion"
      Item(0).Control(12)=   "Label1(3)"
      Item(0).Control(13)=   "Label1(4)"
      Item(0).Control(14)=   "Label1(7)"
      Item(0).Control(15)=   "Label1(8)"
      Item(0).Control(16)=   "Label1(22)"
      Item(0).Control(17)=   "chkFechas"
      Item(0).Control(18)=   "dtpInicio"
      Item(0).Control(19)=   "dtpCorte"
      Item(0).Control(20)=   "Label1(1)"
      Item(0).Control(21)=   "Label1(5)"
      Item(0).Control(22)=   "Label1(6)"
      Item(0).Control(23)=   "Label1(2)"
      Item(0).Control(24)=   "cmdBuscar"
      Item(0).Control(25)=   "gbAplicar"
      Item(0).Control(26)=   "cboEstadoLaboral"
      Item(0).Control(27)=   "cboEstadoPersona"
      Item(0).Control(28)=   "cboDivisa"
      Item(0).Control(29)=   "Label1(12)"
      Item(0).Control(30)=   "btnExport"
      Item(1).Caption =   "Adicionales"
      Item(1).ControlCount=   24
      Item(1).Control(0)=   "cboCobro"
      Item(1).Control(1)=   "cboTipoOperacion"
      Item(1).Control(2)=   "cboSigno(0)"
      Item(1).Control(3)=   "cboSigno(1)"
      Item(1).Control(4)=   "txtPlazoDesde"
      Item(1).Control(5)=   "txtPlazoHasta"
      Item(1).Control(6)=   "txtTasaDesde"
      Item(1).Control(7)=   "txtTasaHasta"
      Item(1).Control(8)=   "txtUltMov"
      Item(1).Control(9)=   "txtPrideduc"
      Item(1).Control(10)=   "chkPlazos"
      Item(1).Control(11)=   "chkTasas"
      Item(1).Control(12)=   "chkPriDeduc"
      Item(1).Control(13)=   "chkUltMov"
      Item(1).Control(14)=   "Label1(34)"
      Item(1).Control(15)=   "Label1(28)"
      Item(1).Control(16)=   "Label1(29)"
      Item(1).Control(17)=   "Label1(30)"
      Item(1).Control(18)=   "Label1(32)"
      Item(1).Control(19)=   "Label1(33)"
      Item(1).Control(20)=   "Label1(36)"
      Item(1).Control(21)=   "Label1(9)"
      Item(1).Control(22)=   "txtNota"
      Item(1).Control(23)=   "Label1(14)"
      Begin XtremeSuiteControls.GroupBox gbAplicar 
         Height          =   3372
         Left            =   9240
         TabIndex        =   2
         Top             =   360
         Width           =   4812
         _Version        =   1441792
         _ExtentX        =   8488
         _ExtentY        =   5948
         _StockProps     =   79
         Caption         =   "Aplicación: "
         ForeColor       =   -2147483631
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
         Begin XtremeSuiteControls.ComboBox cboTipoAplicacion 
            Height          =   312
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2355
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpAplInicio 
            Height          =   312
            Left            =   1680
            TabIndex        =   61
            Top             =   360
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.DateTimePicker dtpAplCorte 
            Height          =   312
            Left            =   3000
            TabIndex        =   62
            Top             =   360
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.CheckBox chkAplIntereses 
            Height          =   228
            Left            =   240
            TabIndex        =   63
            Top             =   1560
            Width           =   2772
            _Version        =   1441792
            _ExtentX        =   4890
            _ExtentY        =   402
            _StockProps     =   79
            Caption         =   "Cobra Intereses"
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
            TextAlignment   =   1
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkAplCargos 
            Height          =   228
            Left            =   240
            TabIndex        =   64
            Top             =   1920
            Width           =   2772
            _Version        =   1441792
            _ExtentX        =   4890
            _ExtentY        =   402
            _StockProps     =   79
            Caption         =   "Cobra Cargos"
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
            TextAlignment   =   1
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkAplPolizas 
            Height          =   228
            Left            =   240
            TabIndex        =   65
            Top             =   2280
            Width           =   2772
            _Version        =   1441792
            _ExtentX        =   4890
            _ExtentY        =   402
            _StockProps     =   79
            Caption         =   "Cobra Pólizas"
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
            TextAlignment   =   1
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkAplAjustaPlazo 
            Height          =   228
            Left            =   1320
            TabIndex        =   66
            Top             =   2640
            Width           =   1692
            _Version        =   1441792
            _ExtentX        =   2984
            _ExtentY        =   402
            _StockProps     =   79
            Caption         =   "Ajusta Plazo "
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
            TextAlignment   =   1
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkAplRetroactivo 
            Height          =   228
            Left            =   960
            TabIndex        =   67
            Top             =   3000
            Width           =   2052
            _Version        =   1441792
            _ExtentX        =   3619
            _ExtentY        =   402
            _StockProps     =   79
            Caption         =   "Permite Retroactivo"
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
            TextAlignment   =   1
            Appearance      =   16
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton cmdAplicar 
            Height          =   615
            Left            =   3240
            TabIndex        =   70
            Top             =   2640
            Width           =   1335
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmCR_PeriodosGracia.frx":0000
            ImageAlignment  =   4
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo aplicación"
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
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   1212
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            X1              =   0
            X2              =   0
            Y1              =   360
            Y2              =   3360
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rango de Cuotas Vencimientos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Index           =   11
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1452
         End
      End
      Begin XtremeSuiteControls.ComboBox cboDestino 
         Height          =   312
         Left            =   960
         TabIndex        =   6
         Top             =   1440
         Width           =   4572
         _Version        =   1441792
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboRecurso 
         Height          =   312
         Left            =   960
         TabIndex        =   7
         Top             =   1800
         Width           =   4572
         _Version        =   1441792
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   312
         Left            =   960
         TabIndex        =   8
         Top             =   2160
         Width           =   4572
         _Version        =   1441792
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboDeductora 
         Height          =   312
         Left            =   960
         TabIndex        =   9
         Top             =   2520
         Width           =   4572
         _Version        =   1441792
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   312
         Left            =   960
         TabIndex        =   10
         Top             =   960
         Width           =   4572
         _Version        =   1441792
         _ExtentX        =   8070
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkLineas 
         Height          =   230
         Left            =   4560
         TabIndex        =   11
         Top             =   360
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1714
         _ExtentY        =   406
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1800
         TabIndex        =   13
         Top             =   600
         Width           =   3732
         _Version        =   1441792
         _ExtentX        =   6583
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
      Begin XtremeSuiteControls.ComboBox cboEstadoLaboral 
         Height          =   312
         Left            =   7080
         TabIndex        =   14
         Top             =   1800
         Width           =   1932
         _Version        =   1441792
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   228
         Left            =   8040
         TabIndex        =   15
         Top             =   360
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1714
         _ExtentY        =   406
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   7680
         TabIndex        =   16
         Top             =   600
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   7680
         TabIndex        =   17
         Top             =   960
         Width           =   1332
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.ComboBox cboEstadoPersona 
         Height          =   312
         Left            =   7080
         TabIndex        =   18
         Top             =   1440
         Width           =   1932
         _Version        =   1441792
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboCobro 
         Height          =   312
         Left            =   -68440
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441792
         _ExtentX        =   4895
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboTipoOperacion 
         Height          =   312
         Left            =   -68440
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1441792
         _ExtentX        =   4895
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboSigno 
         Height          =   312
         Index           =   0
         Left            =   -63640
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboSigno 
         Height          =   312
         Index           =   1
         Left            =   -63640
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1508
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazoDesde 
         Height          =   312
         Left            =   -63640
         TabIndex        =   23
         Top             =   840
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtPlazoHasta 
         Height          =   312
         Left            =   -62800
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtTasaDesde 
         Height          =   312
         Left            =   -63640
         TabIndex        =   25
         Top             =   1200
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtTasaHasta 
         Height          =   312
         Left            =   -62800
         TabIndex        =   26
         Top             =   1200
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtUltMov 
         Height          =   312
         Left            =   -62800
         TabIndex        =   27
         Top             =   1920
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.FlatEdit txtPrideduc 
         Height          =   312
         Left            =   -62800
         TabIndex        =   28
         Top             =   1560
         Visible         =   0   'False
         Width           =   852
         _Version        =   1441792
         _ExtentX        =   1503
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
      Begin XtremeSuiteControls.CheckBox chkPlazos 
         Height          =   252
         Left            =   -61720
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTasas 
         Height          =   252
         Left            =   -61720
         TabIndex        =   30
         Top             =   1200
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkPriDeduc 
         Height          =   252
         Left            =   -61720
         TabIndex        =   31
         Top             =   1560
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkUltMov 
         Height          =   252
         Left            =   -61720
         TabIndex        =   32
         Top             =   1920
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441792
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   312
         Left            =   7080
         TabIndex        =   33
         Top             =   2160
         Width           =   1932
         _Version        =   1441792
         _ExtentX        =   3413
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNota 
         Height          =   1392
         Left            =   -60040
         TabIndex        =   34
         Top             =   840
         Visible         =   0   'False
         Width           =   3852
         _Version        =   1441792
         _ExtentX        =   6794
         _ExtentY        =   2455
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
         Text            =   "Aplicación de Periodo de Gracia General"
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdBuscar 
         Height          =   615
         Left            =   6480
         TabIndex        =   68
         Top             =   3000
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmCR_PeriodosGracia.frx":0719
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   615
         Left            =   7800
         TabIndex        =   69
         Top             =   3000
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Picture         =   "frmCR_PeriodosGracia.frx":0E19
         ImageAlignment  =   4
      End
      Begin VB.Label Label1 
         Caption         =   "Cobro en"
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
         Left            =   -69880
         TabIndex        =   57
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso"
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
         Left            =   -68200
         TabIndex        =   56
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Institución"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   18
         Left            =   120
         TabIndex        =   55
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Recurso"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   15
         Left            =   120
         TabIndex        =   54
         Top             =   1800
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   13
         Left            =   120
         TabIndex        =   53
         Top             =   1440
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Deductora"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   37
         Left            =   120
         TabIndex        =   52
         Top             =   2520
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Linea"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   7
         Left            =   120
         TabIndex        =   51
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Garantía"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   8
         Left            =   120
         TabIndex        =   50
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Laboral"
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
         Height          =   252
         Index           =   22
         Left            =   5760
         TabIndex        =   49
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Formalizadas"
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
         Left            =   5760
         TabIndex        =   48
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   5
         Left            =   7080
         TabIndex        =   47
         Top             =   600
         Width           =   612
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   312
         Index           =   6
         Left            =   7080
         TabIndex        =   46
         Top             =   960
         Width           =   612
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Persona"
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
         Height          =   252
         Index           =   2
         Left            =   5760
         TabIndex        =   45
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Deduc."
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
         Index           =   34
         Left            =   -64960
         TabIndex        =   44
         Top             =   1560
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Operación"
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
         Index           =   28
         Left            =   -69880
         TabIndex        =   43
         Top             =   1200
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazos"
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
         Index           =   29
         Left            =   -64960
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasas"
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
         Index           =   30
         Left            =   -64960
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Desde"
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
         Height          =   252
         Index           =   32
         Left            =   -63640
         TabIndex        =   40
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Hasta"
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
         Height          =   252
         Index           =   33
         Left            =   -62800
         TabIndex        =   39
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ult.Mov."
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
         Index           =   36
         Left            =   -64960
         TabIndex        =   38
         Top             =   1920
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cobro"
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
         Left            =   -69880
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Divisa"
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
         Height          =   252
         Index           =   12
         Left            =   5760
         TabIndex        =   36
         Top             =   2160
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota del cambio de Tasas:"
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
         Index           =   14
         Left            =   -60040
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   2892
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   58
      Top             =   7890
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   59
      Top             =   8025
      Width           =   14265
      _ExtentX        =   25162
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Casos"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Cuotas Actuales"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Nuevas Cuotas "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Diferencia en Intereses"
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2892
      Left            =   0
      TabIndex        =   60
      Top             =   4920
      Width           =   14172
      _Version        =   524288
      _ExtentX        =   24998
      _ExtentY        =   5101
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
      MaxCols         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_PeriodosGracia.frx":16EA
      VScrollSpecial  =   -1  'True
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodos de Gracia"
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
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   3615
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCR_PeriodosGracia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mHeightMin As Long, mWidthMin As Long

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

tcFiltros.Item(0).Selected = True

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

dtpAplInicio.Value = dtpInicio.Value
dtpAplCorte.Value = dtpInicio.Value

chkFechas.Value = xtpChecked
chkLineas.Value = xtpChecked


cboCobro.Clear
cboCobro.AddItem "Cajas"
cboCobro.AddItem "Planilla"
cboCobro.AddItem "TODOS"
cboCobro.Text = "TODOS"

cboTipoAplicacion.Clear
cboTipoAplicacion.AddItem "TOTAL"
cboTipoAplicacion.AddItem "PARCIAL"
cboTipoAplicacion.Text = "TOTAL"

strSQL = "select rtrim(Garantia) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from crd_garantia_tipos order by descripcion"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, True)

strSQL = "select cod_divisa as 'IdX', descripcion as 'ItmX'" _
       & " From vsys_divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from afi_estados_persona order by descripcion"
Call sbCbo_Llena_New(cboEstadoPersona, strSQL, True, True)


'Instituciones
vPaso = True
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)
vPaso = False


strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX'" _
       & " from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoLaboral, strSQL, True, True)



cboTipoOperacion.Clear
cboTipoOperacion.AddItem "TODAS"
cboTipoOperacion.AddItem "Originales"
cboTipoOperacion.AddItem "Derivadas"
cboTipoOperacion.Text = "TODAS"

cboSigno(0).Clear
cboSigno(0).AddItem ">"
cboSigno(0).AddItem "<"
cboSigno(0).AddItem "="
cboSigno(0).Text = "="

cboSigno(1).Clear
cboSigno(1).AddItem ">"
cboSigno(1).AddItem "<"
cboSigno(1).AddItem "="
cboSigno(1).Text = "="

txtPrideduc.Text = GLOBALES.glngFechaCR
txtUltMov.Text = GLOBALES.glngFechaCR

txtPlazoDesde.Text = 1
txtPlazoHasta.Text = 999

txtTasaDesde.Text = 0
txtTasaHasta.Text = 100

chkTasas_Click
chkPlazos_Click
chkPriDeduc_Click
chkUltMov_Click



Call chkFechas_Click
Call chkLineas_Click
Call cboInstitucion_Click
Call cboTipoAplicacion_Click

Me.MousePointer = vbDefault

End Sub




Private Sub btnExport_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 13
    
    vHeaders.Headers(1) = "Operación"
    vHeaders.Headers(2) = "Linea"
    vHeaders.Headers(3) = "Identificación"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Monto Original"
    vHeaders.Headers(6) = "Saldo"
    vHeaders.Headers(7) = "Plazo"
    vHeaders.Headers(8) = "Tasa Actual"
    vHeaders.Headers(9) = "Cuota Actual"
    vHeaders.Headers(10) = "Pri.Deduc."
    vHeaders.Headers(11) = "Ult.Deduc."
    vHeaders.Headers(12) = "Descripcion"
    vHeaders.Headers(13) = "Deductora"
    
    
    
 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Periodos_Gracia")

End Sub

Private Sub cboDeductora_Click()
vGrid.MaxRows = 0

End Sub

Private Sub cboDestino_Click()
vGrid.MaxRows = 0

End Sub

Private Sub cboDivisa_Click()
vGrid.MaxRows = 0
End Sub

Private Sub cboEstadoLaboral_Click()
vGrid.MaxRows = 0

End Sub

Private Sub cboEstadoPersona_Click()
vGrid.MaxRows = 0

End Sub

Private Sub cboGarantia_Click()
vGrid.MaxRows = 0

End Sub

Private Sub cboRecurso_Click()
vGrid.MaxRows = 0

End Sub

Private Sub cboTipoAplicacion_Click()

If cboTipoAplicacion.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset


vGrid.MaxRows = 0

chkAplAjustaPlazo.Value = xtpChecked
chkAplRetroactivo.Value = xtpChecked

Select Case Mid(cboTipoAplicacion.Text, 1, 1)
    Case "T"
        chkAplCargos.Value = xtpUnchecked
        chkAplPolizas.Value = xtpUnchecked
        chkAplIntereses.Value = xtpUnchecked
    
        chkAplCargos.Enabled = False
        chkAplPolizas.Enabled = False
        chkAplIntereses.Enabled = False
    
    
    Case "P"
        chkAplCargos.Value = xtpChecked
        chkAplPolizas.Value = xtpChecked
        chkAplIntereses.Value = xtpChecked
    
        chkAplCargos.Enabled = True
        chkAplPolizas.Enabled = True
        chkAplIntereses.Enabled = True
End Select


End Sub



Private Sub chkPlazos_Click()
If chkPlazos.Value = vbChecked Then
 txtPlazoDesde.Enabled = False
Else
 txtPlazoDesde.Enabled = True
End If
txtPlazoHasta.Enabled = txtPlazoDesde.Enabled
End Sub

Private Sub chkPriDeduc_Click()
If chkPriDeduc.Value = vbChecked Then
   txtPrideduc.Enabled = False
Else
   txtPrideduc.Enabled = True
End If
End Sub


Private Sub chkTasas_Click()
If chkTasas.Value = vbChecked Then
 txtTasaDesde.Enabled = False
Else
 txtTasaDesde.Enabled = True
End If
txtTasaHasta.Enabled = txtTasaDesde.Enabled

End Sub

Private Sub chkUltMov_Click()
If chkUltMov.Value = vbChecked Then
   txtUltMov.Enabled = False
Else
   txtUltMov.Enabled = True
End If
End Sub

Private Sub cboInstitucion_Click()
Dim strSQL As String

If vPaso Then Exit Sub

cboDeductora.Clear

If cboInstitucion.Text = "TODOS" Then
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboDeductora, strSQL, True, True)
Else
    strSQL = "exec spAFI_Institucion_Vinculadas " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ",3"
    Call sbCbo_Llena_New(cboDeductora, strSQL, True, True)
End If

vGrid.MaxRows = 0

End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub


Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_grupos order by descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select rtrim(cod_destino) as 'IdX', rtrim(descripcion) as 'ItmX'" _
         & " from  catalogo_destinos order by descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "' order by R.descripcion"
  Call sbCbo_Llena_New(cboRecurso, strSQL, True, True)
  
  strSQL = "select (R.cod_destino) as 'IdX', rtrim(R.descripcion) as 'ItmX'" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "' order by R.Descripcion"
  Call sbCbo_Llena_New(cboDestino, strSQL, True, True)

End If

End Sub



Private Sub CmdAplicar_Click()
Dim strSQL As String, i As Long

If vGrid.MaxRows = 0 Then Exit Sub

i = MsgBox("Esta seguro que desea aplicar el periodo de gracia a las operaciones mostradas?", vbYesNo)
If i = vbNo Then
  Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = fxSQL("P")
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Periodo de Gracia Aplicado Satisfactoriamente...", vbInformation

Call sbBuscar

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxSQL(Optional pTipo As String = "C") As String
Dim pObjeto As String
Dim strSQL As String
Dim pLinea As String, pGarantia As String, pDestino As String, pRecurso As String
Dim pInstitucion As String, pDeductora As String, pDivisa As String
Dim pEstadoPersona As String, pEstadoLaboral As String
Dim pFormalizaInicio As String, pFormalizaCorte As String
Dim pPlazoRng As Integer, pPlazoInicio As String, pPlazoCorte As String
Dim pTasaRng As Integer, pTasaInicio As String, pTasaCorte As String

Dim pCobroTipo As String, pOperacionTipo As String
Dim pPriDeducFiltro As String, pPriDeduc As String, pPriDeducApl As Integer
Dim pUltDeducFiltro As String, pUltDeduc As String, pUltDeducApl As Integer


Select Case True
    Case pTipo = "C"
        pObjeto = "spCrd_Masivo_Periodo_Gracia_Consulta"
    Case pTipo = "P"
        pObjeto = "spCrd_Masivo_Periodo_Gracia"
End Select
       
If chkLineas.Value = vbUnchecked Then
   pLinea = "'" & txtCodigo.Text & "'"
Else
   pLinea = "Null"
End If
       
If cboDestino.Text <> "TODOS" Then
  pDestino = "'" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
Else
  pDestino = "Null"
End If
       
If cboRecurso.Text <> "TODOS" Then
   pRecurso = "'" & cboRecurso.ItemData(cboRecurso.ListIndex) & "'"
Else
   pRecurso = "Null"
End If
       
If cboInstitucion.Text <> "TODOS" Then
   pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
Else
   pInstitucion = "Null"
End If
       
If cboDeductora.Text <> "TODOS" Then
   pDeductora = cboDeductora.ItemData(cboDeductora.ListIndex)
Else
   pDeductora = "Null"
End If
       
If cboGarantia.Text <> "TODOS" Then
   pGarantia = "'" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
Else
   pGarantia = "Null"
End If
       
If cboDivisa.Text <> "TODOS" Then
   pDivisa = "'" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
Else
   pDivisa = "Null"
End If
       
If cboEstadoLaboral.Text <> "TODOS" Then
   pEstadoLaboral = "'" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) & "'"
Else
   pEstadoLaboral = "Null"
End If

If cboEstadoPersona.Text <> "TODOS" Then
   pEstadoPersona = "'" & cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex) & "'"
Else
   pEstadoPersona = "Null"
End If
       
If chkFechas.Value = vbUnchecked Then
   pFormalizaInicio = "'" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'"
   pFormalizaCorte = "'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Else
   pFormalizaInicio = "Null"
   pFormalizaCorte = "Null"
End If
       
     
If chkPlazos.Value = vbUnchecked Then
   pPlazoRng = 1
   pPlazoInicio = txtPlazoDesde.Text
   pPlazoCorte = txtPlazoHasta.Text
Else
   pPlazoRng = 0
   pPlazoInicio = "Null"
   pPlazoCorte = "Null"
End If
       
If chkTasas.Value = vbUnchecked Then
    pTasaRng = 1
    pTasaInicio = txtTasaDesde.Text
    pTasaCorte = txtTasaHasta.Text
Else
    pTasaRng = 0
    pTasaInicio = "Null"
    pTasaCorte = "Null"
End If
       


'Adicionales
If cboCobro.Text <> "TODOS" Then
   pCobroTipo = "'" & Mid(cboCobro.Text, 1, 1) & "'"
Else
   pCobroTipo = "Null"
End If

If cboTipoOperacion.Text <> "TODOS" Then
   pOperacionTipo = "'" & Mid(cboTipoOperacion.Text, 1, 1) & "'"
Else
   pOperacionTipo = "Null"
End If

If chkPriDeduc.Value = xtpChecked Then
    pPriDeducApl = 0
    pPriDeducFiltro = "Null"
    pPriDeduc = "Null"
Else
    pPriDeducApl = 1
    pPriDeducFiltro = "'" & cboSigno(0).Text & "'"
    pPriDeduc = txtPrideduc.Text
End If

If chkUltMov.Value = xtpChecked Then
    pUltDeducApl = 0
    pUltDeducFiltro = "Null"
    pUltDeduc = "Null"
Else
    pUltDeducApl = 1
    pUltDeducFiltro = "'" & cboSigno(1).Text & "'"
    pUltDeduc = txtUltMov.Text
End If

strSQL = "exec " & pObjeto & " " & pLinea & "," & pGarantia & "," & pDestino & "," & pRecurso _
       & "," & pInstitucion & "," & pDeductora & "," & pDivisa _
       & "," & pEstadoPersona & "," & pEstadoLaboral _
       & "," & pFormalizaInicio & "," & pFormalizaCorte _
       & ",'" & Format(dtpAplInicio.Value, "yyyy/mm/dd") & " 00:00:00','" & Format(dtpAplCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & "," & pPlazoRng & "," & pPlazoInicio & "," & pPlazoCorte _
       & "," & pTasaRng & "," & pTasaInicio & "," & pTasaCorte _
       & "," & pCobroTipo & "," & pOperacionTipo _
       & "," & pPriDeducApl & "," & pPriDeducFiltro & "," & pPriDeduc _
       & "," & pUltDeducApl & "," & pUltDeducFiltro & "," & pUltDeduc _
       & ",'" & Mid(cboTipoAplicacion.Text, 1, 1) & "'," & chkAplAjustaPlazo.Value & "," & chkAplRetroactivo.Value _
       & "," & chkAplIntereses.Value & "," & chkAplCargos.Value & "," & chkAplPolizas.Value _
       & ",'" & glogon.Usuario & "','" & txtNota.Text & "'"
       
'Return
fxSQL = strSQL

End Function


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngCasos As Long, i As Long, curCuotas As Currency, curCuotasNew As Currency

Me.MousePointer = vbHourglass

lngCasos = 0
curCuotas = 0
curCuotasNew = 0
       

strSQL = fxSQL("C")
Call OpenRecordSet(rs, strSQL)
       
vGrid.MaxRows = 0
ProgressBarX.Max = rs.RecordCount + 1
       
Do While Not rs.EOF
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  ProgressBarX.Value = vGrid.MaxRows
  
 
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
       Case 1
          vGrid.Text = CStr(rs!Id_Solicitud)
       Case 2
          vGrid.Text = CStr(rs!Codigo)
       Case 3
          vGrid.Text = CStr(rs!Cedula)
       Case 4
          vGrid.Text = CStr(rs!Nombre)
       Case 5
          vGrid.Text = Format(rs!montoapr, "Standard")
       Case 6
          vGrid.Text = Format(rs!Saldo, "Standard")
       Case 7
          vGrid.Text = CStr(rs!Plazo)
       Case 8
          vGrid.Text = CStr(rs!Tasa)
       Case 9
          vGrid.Text = Format(rs!Cuota, "Standard")
       
       Case 10
          vGrid.Text = Format(rs!PriDeduc, "####-##")
       Case 11
          vGrid.Text = Format(rs!FecUlt, "####-##")
       Case 12
          vGrid.Text = rs!Linea_Desc
       Case 13
          vGrid.Text = rs!DEDUCTORA_DESC
          
    End Select
  Next i
   
  lngCasos = lngCasos + 1
  curCuotas = curCuotas + rs!Cuota
  
  
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

stBar.Panels(1) = Format(lngCasos, "###,###,###")
stBar.Panels(2) = Format(curCuotas, "Standard")

End Sub



Private Sub cmdBuscar_Click()
Call sbBuscar
End Sub



Private Sub dtpCorte_Change()
vGrid.MaxRows = 0
End Sub

Private Sub dtpInicio_Change()
vGrid.MaxRows = 0
End Sub


Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

vGrid.MaxCols = 13
vGrid.MaxRows = 0

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

mWidthMin = 14364
mHeightMin = 8280

Me.Width = mWidthMin
Me.Height = mHeightMin

End Sub

Private Sub Form_Resize()
On Error Resume Next


If Me.Height < mHeightMin Then
    Me.Height = mHeightMin
End If

If Me.Width < mWidthMin Then
    Me.Width = mWidthMin
End If

tcFiltros.Width = Me.Width

vGrid.Width = Me.Width - 250
vGrid.Height = Me.Height - (vGrid.top + 880)

imgBanner.Width = Me.Width

End Sub



Private Sub tcFiltros_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
vGrid.MaxRows = 0

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  txtDescripcion.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then txtDescripcion.Text = fxDescribeCodigo(Trim(txtCodigo))
 Call chkLineas_Click
End Sub


