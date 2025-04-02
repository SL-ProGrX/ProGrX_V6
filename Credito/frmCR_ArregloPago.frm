VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_ArregloPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arreglos de Pago"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11415
   HelpContextID   =   3003
   Icon            =   "frmCR_ArregloPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   11415
   Begin XtremeSuiteControls.TabControl tcInfo 
      Height          =   1932
      Left            =   1440
      TabIndex        =   41
      Top             =   1560
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   3408
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
      ItemCount       =   1
      Item(0).Caption =   "Estado"
      Item(0).ControlCount=   22
      Item(0).Control(0)=   "txtIntCor"
      Item(0).Control(1)=   "txtIntMor"
      Item(0).Control(2)=   "txtCargos"
      Item(0).Control(3)=   "txtPolizas"
      Item(0).Control(4)=   "txtAmortiza"
      Item(0).Control(5)=   "txtCargosIntereses"
      Item(0).Control(6)=   "txtDeuda"
      Item(0).Control(7)=   "txtTotalPagar"
      Item(0).Control(8)=   "txtMonto"
      Item(0).Control(9)=   "txtSaldo"
      Item(0).Control(10)=   "txtUltimoMov"
      Item(0).Control(11)=   "Label4(10)"
      Item(0).Control(12)=   "Label4(9)"
      Item(0).Control(13)=   "Label4(8)"
      Item(0).Control(14)=   "Label4(6)"
      Item(0).Control(15)=   "Label2(0)"
      Item(0).Control(16)=   "Label4(0)"
      Item(0).Control(17)=   "Label4(2)"
      Item(0).Control(18)=   "Label4(3)"
      Item(0).Control(19)=   "Label7(1)"
      Item(0).Control(20)=   "Label4(4)"
      Item(0).Control(21)=   "Label4(5)"
      Begin XtremeSuiteControls.FlatEdit txtIntCor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   1440
         TabIndex        =   42
         Top             =   480
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntMor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   1440
         TabIndex        =   43
         Top             =   840
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   1440
         TabIndex        =   44
         Top             =   1200
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPolizas 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   1440
         TabIndex        =   45
         Top             =   1560
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAmortiza 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   4920
         TabIndex        =   46
         Top             =   480
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargosIntereses 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   4920
         TabIndex        =   47
         Top             =   840
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDeuda 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   4920
         TabIndex        =   48
         Top             =   1200
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   4920
         TabIndex        =   49
         Top             =   1560
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   50
         Top             =   480
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   51
         Top             =   840
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUltimoMov 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   0
         EndProperty
         Height          =   312
         Left            =   8160
         TabIndex        =   52
         Top             =   1200
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   5
         Left            =   6720
         TabIndex        =   63
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Pólizas "
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
         Left            =   120
         TabIndex        =   62
         Top             =   1560
         Width           =   2052
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Compromiso:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   3240
         TabIndex        =   61
         Top             =   1560
         Width           =   1932
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Principal Atrasado"
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
         TabIndex        =   60
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargos "
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
         Left            =   120
         TabIndex        =   59
         Top             =   1200
         Width           =   2052
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Int. Moratorio"
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
         Left            =   120
         TabIndex        =   58
         Top             =   840
         Width           =   2052
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Int. Corriente"
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
         Left            =   120
         TabIndex        =   57
         Top             =   480
         Width           =   2052
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ult. Cuota"
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
         Left            =   6720
         TabIndex        =   56
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Index           =   8
         Left            =   6720
         TabIndex        =   55
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargos + Intereses"
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
         Left            =   3240
         TabIndex        =   54
         Top             =   840
         Width           =   2052
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deuda"
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
         Left            =   3240
         TabIndex        =   53
         Top             =   1200
         Width           =   2052
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2172
      Left            =   1440
      TabIndex        =   21
      Top             =   3600
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   3831
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
      ItemCount       =   4
      SelectedItem    =   1
      Item(0).Caption =   "Capitaliza Deuda"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "chkTrasladar"
      Item(0).Control(2)=   "chkTipoIntereses"
      Item(1).Caption =   "Periodos de Gracia"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "cboTipoAplicacion"
      Item(1).Control(1)=   "dtpAplInicio"
      Item(1).Control(2)=   "dtpAplCorte"
      Item(1).Control(3)=   "Label1(10)"
      Item(1).Control(4)=   "Label1(11)"
      Item(1).Control(5)=   "chkAplIntereses"
      Item(1).Control(6)=   "chkAplCargos"
      Item(1).Control(7)=   "chkAplPolizas"
      Item(1).Control(8)=   "chkAplAjustaPlazo"
      Item(1).Control(9)=   "chkAplRetroactivo"
      Item(2).Caption =   "Vencimiento de Intereses"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lblCorte"
      Item(2).Control(1)=   "dtpCorte"
      Item(3).Caption =   "Abono Especial"
      Item(3).ControlCount=   18
      Item(3).Control(0)=   "txtAE_IntMor"
      Item(3).Control(1)=   "txtAE_IntCor"
      Item(3).Control(2)=   "txtAE_Principal"
      Item(3).Control(3)=   "txtAE_Polizas"
      Item(3).Control(4)=   "txtAE_Cargos"
      Item(3).Control(5)=   "txtAE_Total"
      Item(3).Control(6)=   "cboAE_Tipo"
      Item(3).Control(7)=   "cboAE_CuotaFecha"
      Item(3).Control(8)=   "txtAE_CuotaNum"
      Item(3).Control(9)=   "lblAE_Titulo(1)"
      Item(3).Control(10)=   "lblAE_Titulo(0)"
      Item(3).Control(11)=   "Label2(6)"
      Item(3).Control(12)=   "Label4(12)"
      Item(3).Control(13)=   "Label2(5)"
      Item(3).Control(14)=   "Label4(11)"
      Item(3).Control(15)=   "Label4(7)"
      Item(3).Control(16)=   "Label2(2)"
      Item(3).Control(17)=   "Label4(1)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1452
         Left            =   -70000
         TabIndex        =   40
         Top             =   360
         Visible         =   0   'False
         Width           =   9852
         _Version        =   1441793
         _ExtentX        =   17378
         _ExtentY        =   2561
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtAE_IntMor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -67720
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAE_IntCor 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -67720
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAE_Principal 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -67720
         TabIndex        =   24
         Top             =   1800
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAE_Polizas 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -62680
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAE_Cargos 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -62680
         TabIndex        =   26
         Top             =   1080
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAE_Total 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -62680
         TabIndex        =   27
         Top             =   1800
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Text            =   "0"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboAE_Tipo 
         Height          =   312
         Left            =   -67720
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.ComboBox cboAE_CuotaFecha 
         Height          =   312
         Left            =   -64600
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2778
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
      Begin XtremeSuiteControls.FlatEdit txtAE_CuotaNum 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   312
         Left            =   -61840
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipoAplicacion 
         Height          =   312
         Left            =   4920
         TabIndex        =   65
         Top             =   480
         Width           =   1332
         _Version        =   1441793
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
         TabIndex        =   66
         Top             =   480
         Width           =   1332
         _Version        =   1441793
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
         Left            =   1680
         TabIndex        =   67
         Top             =   840
         Width           =   1332
         _Version        =   1441793
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
         Left            =   3480
         TabIndex        =   70
         Top             =   960
         Width           =   2772
         _Version        =   1441793
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
         Left            =   3480
         TabIndex        =   71
         Top             =   1320
         Width           =   2772
         _Version        =   1441793
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
         Left            =   3480
         TabIndex        =   72
         Top             =   1680
         Width           =   2772
         _Version        =   1441793
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
         Left            =   7080
         TabIndex        =   73
         Top             =   960
         Width           =   1692
         _Version        =   1441793
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
         Left            =   6720
         TabIndex        =   74
         Top             =   1320
         Width           =   2052
         _Version        =   1441793
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -65080
         TabIndex        =   75
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.CheckBox chkTrasladar 
         Height          =   252
         Left            =   -70000
         TabIndex        =   76
         Top             =   1920
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ajustar Deuda al Plazo Restante   "
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkTipoIntereses 
         Height          =   252
         Left            =   -67120
         TabIndex        =   77
         Top             =   1920
         Visible         =   0   'False
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cálcular Intereses a hoy?   "
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
         Alignment       =   1
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
         TabIndex        =   69
         Top             =   480
         Width           =   1452
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
         Left            =   3360
         TabIndex        =   68
         Top             =   480
         Width           =   1212
      End
      Begin XtremeSuiteControls.Label lblCorte 
         Height          =   492
         Left            =   -69040
         TabIndex        =   64
         Top             =   840
         Visible         =   0   'False
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6371
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Indique la Fecha desde cuando ya no es posible cobrar intereses a cuotas vencidas:"
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
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
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
         Height          =   252
         Index           =   1
         Left            =   -69640
         TabIndex        =   39
         Top             =   1800
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Interes Corriente"
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
         Left            =   -69640
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label4 
         Caption         =   "Total:"
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
         Left            =   -64600
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label4 
         Caption         =   "Interes Moratorio"
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
         Left            =   -69640
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   2052
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
         Height          =   252
         Index           =   5
         Left            =   -64600
         TabIndex        =   35
         Top             =   1080
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label4 
         Caption         =   "Pólizas"
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
         Left            =   -64600
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Movimiento"
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
         Left            =   -69640
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label lblAE_Titulo 
         Caption         =   "Fecha Cuota"
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
         Left            =   -65800
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label lblAE_Titulo 
         Caption         =   "No. Cuota"
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
         Left            =   -62920
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   972
      End
   End
   Begin XtremeSuiteControls.PushButton cmdAcepta 
      Height          =   612
      Left            =   9720
      TabIndex        =   3
      Top             =   6720
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
      Appearance      =   16
      Picture         =   "frmCR_ArregloPago.frx":6852
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9720
      Top             =   720
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ArregloPago.frx":702A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ArregloPago.frx":714E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   612
      Left            =   1440
      TabIndex        =   4
      Top             =   5880
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   1080
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   3612
         _Version        =   1441793
         _ExtentX        =   6376
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   312
         Left            =   6360
         TabIndex        =   6
         Top             =   240
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   372
         Left            =   8280
         TabIndex        =   7
         Top             =   240
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Pago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         TextImageRelation=   1
      End
      Begin VB.Label Label3 
         Caption         =   "Total ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   5520
         TabIndex        =   9
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "Documento ..:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1452
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2400
      TabIndex        =   10
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   372
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Width           =   4812
      _Version        =   1441793
      _ExtentX        =   8488
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
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   312
      Left            =   4080
      TabIndex        =   13
      Top             =   960
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7429
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2400
      TabIndex        =   14
      Top             =   600
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2400
      TabIndex        =   15
      Top             =   960
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   312
      Left            =   8280
      TabIndex        =   16
      Top             =   960
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton isButtonMain 
      Height          =   732
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   1560
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Capitaliza Deuda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton isButtonMain 
      Height          =   732
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Periodos de Gracia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton isButtonMain 
      Height          =   732
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   3480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Vencimiento de Intereses"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton isButtonMain 
      Height          =   732
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Abono Especial"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   672
      Left            =   2760
      TabIndex        =   78
      Top             =   6720
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   1185
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
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   1560
      TabIndex        =   79
      Top             =   6720
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Height          =   312
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   600
      Width           =   1452
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Height          =   312
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
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
      Height          =   312
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   140
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12492
   End
End
Attribute VB_Name = "frmCR_ArregloPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOperacion As Long, mRetencion As Boolean, mSaldo As Currency, mSaldoActual As Currency
Dim mTasa As Currency, mPlazo As Long, mFecha As Date, mCuota As Currency
Dim pCharRelleno As String, vPaso As Boolean

Private Sub sbAbonaMorosidad(curIntCor As Currency, curIntMor As Currency, curAmortiza As Currency, curCargo As Currency _
            , curPoliza As Currency, IdMoro As Long, Proceso As Long, vTipoDoc As String, vNumDoc As String)
Dim strSQL As String

On Error GoTo vError


strSQL = "update MOROSIDAD set estado='C',abIntC = " & curIntCor & ",abIntM = " & curIntMor & ",abCargo = " & curCargo _
       & ",abAmortiza = " & curAmortiza & ",tcon = '" & vTipoDoc & "',ncon = '" & vNumDoc & "',fecAP=" & GLOBALES.glngFechaCR _
       & ",fecult = dbo.MyGetdate(), Cod_Concepto = 'CRD014',Usuario = '" & glogon.Usuario & "', Cod_Caja = ''" _
       & " where id_moro = " & IdMoro
'Call ConectionExecute(strSQL)

'Ingresa Diferencias (en caso de ser 0 no registrar)
strSQL = strSQL & Space(10) & "insert into morosidad(codigo,id_solicitud,fechap,cuota_morosa,intc,intm,cargo,amortiza,estado,fecap,estadoi,fecult,tcon,ncon,cod_concepto,usuario,cod_caja)" _
       & "(select codigo,id_solicitud,fechap, (intc+intm+cargo+amortiza) - (Abintc+ Abintm+ Abcargo+ Abamortiza), (IntC - AbIntC), (IntM - AbIntM)" _
       & ",(Cargo - AbCargo), (Amortiza - AbAmortiza),'A'," & GLOBALES.glngFechaCR & ",'A',dbo.MyGetdate(),Tcon,NCon,'CRD014','" & glogon.Usuario & "',''" _
       & " from morosidad where id_Moro = " & IdMoro & " and ((intc+intm+cargo+amortiza) - (Abintc+ Abintm+ Abcargo+ Abamortiza)) > 0)"
'Call ConectionExecute(strSQL)

strSQL = strSQL & Space(10) & "update reg_creditos set interesC = interesC + " & (curIntCor + curIntMor) _
       & ",amortiza = amortiza + " & curAmortiza & ",saldo = saldo - " & curAmortiza _
       & " where id_solicitud = " & mOperacion
       
Call ConectionExecute(strSQL)

mSaldoActual = mSaldoActual - curAmortiza

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaMora()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curIntCor As Currency, curIntMor As Currency, curCargos As Currency, curPrincipal As Currency
Dim itmX As ListViewItem, curPoliza As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass


If GLOBALES.SysPlanPagos = 1 Then
   If chkTipoIntereses.Value = vbChecked Then
        'Estado de la Operacion Calculada a Hoy
        strSQL = "exec spCrdPlanPagosProyectaCuota " & mOperacion & ", '" & Format(mFecha, "yyyy/mm/dd") & "',1"
        Call ConectionExecute(strSQL)
         'Fix Ultima Cuota
         strSQL = "update CRD_OPERACION_TRANSAC_CAL set PRINCIPAL = 0 " _
                & " where ID_SOLICITUD = " & mOperacion & " and ID_SEQ in(select max(ID_SEQ) " _
                                                & " from CRD_OPERACION_TRANSAC_CAL where ID_SOLICITUD = " & mOperacion & ")"
        
         strSQL = strSQL & Space(10) & "select Det.ID_SEQ as 'ID_Moro',Det.ID_SOLICITUD,Det.FECHA_PROCESO as 'FechaP',Det.INTCOR as 'IntC'" _
                & ",Det.INTMOR as 'IntM',Det.PRINCIPAL as 'Amortiza', Det.CARGOS as 'Cargo',Det.ESTADO" _
                & ",Det.INTCOR + Det.INTMOR + Det.PRINCIPAL + Det.Poliza + Det.CARGOS as 'Cuota_Morosa',0 as 'AbIntC',0 as 'AbIntM'" _
                & ",0 as 'AbAmortiza', 0 as 'AbCargo',Det.Poliza, 0 as 'AbPoliza'" _
                & " from CRD_OPERACION_TRANSAC_CAL Det inner join REG_CREDITOS Reg on Det.ID_SOLICITUD = Reg.ID_SOLICITUD" _
                & " where Reg.PROCESO <> 'J' and Det.ESTADO = 'A' and Det.ID_SOLICITUD = " & mOperacion
   Else
         strSQL = "select Det.ID_SEQ as 'ID_Moro',Det.ID_SOLICITUD,Det.FECHA_PROCESO as 'FechaP',Det.INTCOR as 'IntC'" _
                & ",Det.INTMOR as 'IntM',Det.PRINCIPAL as 'Amortiza', Det.CARGOS as 'Cargo',Det.ESTADO" _
                & ",Det.INTCOR + Det.INTMOR + Det.PRINCIPAL + Det.Poliza + Det.CARGOS as 'Cuota_Morosa',0 as 'AbIntC',0 as 'AbIntM'" _
                & ",0 as 'AbAmortiza', 0 as 'AbCargo',Det.Poliza, 0 as 'AbPoliza'" _
                & " from CRD_OPERACION_TRANSAC Det inner join REG_CREDITOS Reg on Det.ID_SOLICITUD = Reg.ID_SOLICITUD" _
                & " where Reg.PROCESO <> 'J' and Det.ESTADO = 'A' and Det.ID_SOLICITUD = " & mOperacion _
                & " and Det.Fecha_Corte <= '" & Format(mFecha, "yyyy/mm/dd") & "'"

   End If
Else
    strSQL = "select id_moro,id_solicitud,fechap,intc,intm,amortiza,isnull(cargo,0) as 'Cargo', 0 as 'Poliza',estado" _
           & ",cuota_morosa,abintc,abintm,isnull(abamortiza,0) as 'abAmortiza',isnull(AbCargo,0) as 'AbCargo', 0 as 'AbPoliza' from MOROSIDAD" _
           & " where id_solicitud = " & mOperacion & " and estado = 'A' order by fechap"
End If

Call OpenRecordSet(rs, strSQL)


lsw.ListItems.Clear
curIntCor = 0
curIntMor = 0
curCargos = 0
curPrincipal = 0
curPoliza = 0

Do While Not rs.EOF

  Set itmX = lsw.ListItems.Add(, , rs!id_moro)
      itmX.SubItems(1) = rs!Id_Solicitud
      itmX.SubItems(2) = Format(rs!fechap, "####-##")
      itmX.SubItems(3) = Format(rs!IntC, "Standard")
      itmX.SubItems(4) = Format(rs!IntM, "Standard")
      itmX.SubItems(5) = Format(rs!Cargo, "Standard")
      itmX.SubItems(6) = Format(rs!Poliza, "Standard")
      itmX.SubItems(7) = Format(rs!Amortiza, "Standard")
      itmX.SubItems(8) = Format(rs!IntC + rs!IntM + rs!Amortiza + rs!Cargo + rs!Poliza, "Standard")
      
      Select Case rs!Estado
       Case "A"
         itmX.SubItems(9) = "Activo"
       Case "C"
         itmX.SubItems(9) = "Cancelado"
       Case Else
         itmX.SubItems(9) = "Anulado"
      End Select
 
      itmX.SubItems(10) = Format(rs!abintc, "Standard")
      itmX.SubItems(11) = Format(rs!abintm, "Standard")
      itmX.SubItems(12) = Format(rs!AbCargo, "Standard")
      itmX.SubItems(13) = Format(rs!abPoliza, "Standard")
      itmX.SubItems(14) = Format(rs!abAmortiza, "Standard")
      itmX.SubItems(15) = Format(rs!abintc + rs!abintm + rs!abAmortiza + rs!AbCargo + rs!abPoliza, "Standard")

       curIntCor = curIntCor + rs!IntC
       curIntMor = curIntMor + rs!IntM
       curCargos = curCargos + rs!Cargo
       curPoliza = curPoliza + rs!Poliza
       curPrincipal = curPrincipal + rs!Amortiza
 rs.MoveNext

Loop
rs.Close

txtIntCor.Text = Format(curIntCor, "Standard")
txtIntMor.Text = Format(curIntMor, "Standard")
txtCargos.Text = Format(curCargos, "Standard")
txtPolizas.Text = Format(curPoliza, "Standard")
txtAmortiza.Text = Format(curPrincipal, "Standard")
txtSaldo.Text = Format(mSaldo, "Standard")
txtDeuda.Text = Format(mSaldo + curIntCor + curIntMor + curCargos + curPoliza, "Standard")


'txtTotalPagar.Text = Format(curIntCor + curIntMor + curCargos + curPoliza, "Standard")
txtTotalPagar.Text = Format(curIntCor + curIntMor + curCargos + curPoliza + curPrincipal, "Standard")
txtTotalPagar.Tag = curIntCor + curIntMor + curCargos + curPoliza + curPrincipal
txtCargosIntereses.Text = Format(curIntCor + curIntMor + curCargos + curPoliza, "Standard")


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCuotasFechas(pFecha As Date, pActual As Date, pPriDeduc As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, pNumCuota As Integer, y As Integer
Dim pFechaMin As Date, pFechaMax As Date

On Error GoTo vError

pActual = DateSerial(Year(pActual), Month(pActual) + 1, 1 - 1)

pFechaMin = DateAdd("m", -1, pFecha)
pFechaMax = DateAdd("m", 1, pActual)


cboAE_CuotaFecha.Clear

pFecha = pFechaMin

Dim pFechaPriDeduc As Date, sPrideduc As String, pAnio As Integer, pMes As Integer

sPrideduc = CStr(pPriDeduc)

pAnio = Mid(sPrideduc, 1, 4)
pMes = Mid(sPrideduc, 5, 2)

pFechaPriDeduc = DateSerial(pAnio, pMes + 1, 1 - 1)


If pFechaMin > pFechaPriDeduc Then
   pFechaMin = pFechaPriDeduc
   pFecha = pFechaMin
End If

pNumCuota = DateDiff("m", pFechaPriDeduc, pFecha) + 1

Do While pFecha < pFechaMax
    cboAE_CuotaFecha.AddItem Format(pFecha, "YYYY-MM")
    
    cboAE_CuotaFecha.ItemData(cboAE_CuotaFecha.ListCount - 1) = CStr(pNumCuota)

    pFecha = DateAdd("m", 1, pFecha)
    pNumCuota = pNumCuota + 1
Loop

pFecha = DateAdd("m", -1, pFecha)
cboAE_CuotaFecha.Text = Format(pFecha, "YYYY-MM")

Call cboAE_CuotaFecha_Click

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select R.id_Solicitud,R.Cedula,R.codigo,R.Plazo,isnull(R.Cuota,0) as 'Cuota',R.MontoApr,R.Saldo,R.Estado,R.Proceso,isnull(R.Opex,0) as 'Opex'" _
       & ",S.nombre,C.descripcion, isnull(V.Amortiza,0) as 'Amortiza',case when (C.Retencion = 'S' or C.Poliza = 'S') Then 'S' else 'N' end as 'Retencion'" _
       & ", isnull(V.intC,0) as 'IntCor' ,isnull(V.intM,0) as 'IntMor', isnull(V.cargos,0) as 'Cargos', R.PriDeduc" _
       & ",isnull(R.liqTasa,0) as LiqTasaX,dbo.fxCRDCalculoIntCorte(R.id_solicitud,dbo.MyGetdate()) as 'InteresTotal'" _
       & ",0 as 'Poliza', O.descripcion as 'OficinaDesc', R.cod_oficina_r as 'Oficina',R.cod_grupo,Pre.descripcion as 'RecursoDesc'" _
       & ",dbo.MyGetdate() as 'FechaServer',  dbo.fxSIFCorteAFecha(isnull(R.FecUlt,R.PriDeduc)) as 'FechaUltMov'" _
       & ",isnull(R.interesv,R.int) as 'Tasa',isnull(R.cod_Divisa,'COL') as 'Divisa'" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " left join sif_oficinas O on R.cod_oficina_R = O.cod_Oficina" _
       & " left join Vista_Morosidad V on R.id_solicitud = V.id_solicitud" _
       & " left join CATALOGO_GRUPOS Pre on R.cod_grupo = Pre.cod_grupo" _
       & " Where R.id_solicitud = " & txtOperacion.Text & " and R.estado = 'A'"
          
Call OpenRecordSet(rs, strSQL)

vPaso = True

Call sbLimpia

If Not rs.EOF And Not rs.BOF Then
 
tcMain.Visible = True
tcInfo.Visible = True

 If isButtonMain(0).Enabled Then isButtonMain(0).SetFocus
 
 mOperacion = rs!Id_Solicitud

 txtOperacion = rs!Id_Solicitud
 txtCedula.Text = rs!Cedula
 txtCodigo.Text = rs!Codigo
 
 txtLineaDesc.Text = rs!Descripcion
 txtNombre.Text = rs!Nombre
 

  txtOpex.Text = IIf((rs!opex = 1), "OPEX", "")
  
  txtUltimoMov.Text = Format(rs!FechaUltMov, "dd/mm/yyyy")


  ModuloCajas.mClienteId = Trim(rs!Cedula)
  ModuloCajas.mCliente = Trim(rs!Nombre)
  ModuloCajas.mTiquete = Trim(rs!Codigo) & "." & rs!Id_Solicitud & "." & Format(Time, "HH:mm:ss")
    
  ModuloCajas.mDivisa = RTrim(rs!Divisa)
  ModuloCajas.mTotalDetallado = 0
    

  mRetencion = IIf(rs!retencion = "S", True, False)
  mSaldo = rs!Saldo
  mSaldoActual = rs!Saldo
  mFecha = rs!FechaServer
  mPlazo = rs!Plazo
  mTasa = rs!Tasa
  mCuota = rs!Cuota
  
  txtAmortiza.Text = Format(rs!Amortiza, "Standard")
 
 txtMonto.Text = Format(rs!montoapr, "Standard")
 txtIntCor.Text = Format(rs!InteresTotal - rs!IntMor, "Standard")
 txtIntMor.Text = Format(rs!IntMor, "Standard")
 txtPolizas.Text = Format(rs!Poliza, "Standard")
 txtCargos.Text = Format(rs!Cargos, "Standard")
 txtSaldo.Text = Format(rs!Saldo, "Standard")
 txtAmortiza.Text = Format(rs!Amortiza, "Standard")
 
 txtCargosIntereses.Text = Format(rs!Cargos + rs!Poliza + rs!InteresTotal, "Standard")
 
 txtTotalPagar.Text = Format(rs!Amortiza + rs!InteresTotal + rs!Cargos + rs!Poliza, "Standard")


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

 Call sbCuotasFechas(rs!FechaUltMov, rs!FechaServer, rs!PriDeduc)

 If GLOBALES.SysPlanPagos = 1 Then
       strSQL = "exec spCrdPlanPagosInfoCancelacion " & txtOperacion.Text & ", '" & Format(rs!FechaServer, "yyyy/mm/dd") & "'"
       rs.Close
       Call OpenRecordSet(rs, strSQL)
        
        txtIntCor.Text = Format(rs!IntCor, "Standard")
        txtIntMor.Text = Format(rs!IntMor, "Standard")
        txtCargos.Text = Format(rs!Cargos, "Standard")
        txtPolizas.Text = Format(rs!Poliza, "Standard")
        txtAmortiza.Text = Format(rs!Principal, "Standard")
    
        
        txtTotalPagar.Text = Format(rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!Poliza, "Standard")
 End If 'GLOBALES.SysPlanPagos = 1
 
 
 Call sbCargaMora
 
Else
    MsgBox "No se encontró registro de la operación [Activa] o no es un crédito!", vbExclamation
    vPaso = False
    Exit Sub
End If

rs.Close

vPaso = False

Call isButtonMain_Click(0)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbLimpia()

tcMain.Visible = False
tcInfo.Visible = False

txtOperacion.Text = ""
txtCedula.Text = ""
txtNombre.Text = ""
txtCodigo.Text = ""
txtLineaDesc.Text = ""
txtOpex.Text = ""
txtProceso.Text = ""


txtIntCor.Text = "0.00"
txtIntMor.Text = "0.00"
txtCargos.Text = "0.00"
txtAmortiza.Text = "0.00"
txtSaldo.Text = "0.00"
txtDeuda.Text = "0.00"
txtCargosIntereses.Text = "0.00"
txtMonto.Text = "0.00"

txtTotalPagar.Text = "0.00"
txtUltimoMov.Text = ""

txtTotalCajas.Text = "0.00"
txtNotas.Text = ""

txtAE_IntCor.Text = "0.00"
txtAE_Principal.Text = "0.00"
txtAE_Total.Text = "0.00"

mRetencion = False
mSaldo = 0
mSaldoActual = 0
mTasa = 0
mPlazo = 0
mOperacion = 0

lsw.ListItems.Clear

Call isButtonMain_Click(0)



End Sub

Private Sub sbDocumentoReadecuacion(pTipoDoc As String, pNumDoc As String _
                                  , pConcepto As String, Optional pCuenta As String = "")
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String

Dim strCliente As String, vCuenta As String, vOperacion As Long
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String, curSaldo As Currency
Dim pTipoCambio As Currency


pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)

vCuenta = pCuenta


curIntC = 0
curIntM = 0
curAmortiza = 0
curCargo = 0
curPoliza = 0

vOperacion = txtOperacion.Text

If GLOBALES.SysPlanPagos = 0 Then
    strSQL = "exec spCrdDocumentoAfectacionStP '" & pTipoDoc & "','" & pNumDoc & "','R'"
Else
    strSQL = "exec spCrdDocumentoAfectacion '" & pTipoDoc & "','" & pNumDoc & "','R'"
End If

Call OpenRecordSet(rsTmp, strSQL, 0)

curIntC = rsTmp!IntCor
curIntM = rsTmp!IntMor
curAmortiza = rsTmp!IntCor + rsTmp!IntMor + rsTmp!Cargos + rsTmp!Polizas
curCargo = rsTmp!Cargos
curPoliza = rsTmp!Polizas

rsTmp.Close


strSQL = "select Saldo from reg_Creditos where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rsTmp, strSQL, 0)
  curSaldo = rsTmp!Saldo
rsTmp.Close


strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(Format(curSaldo - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(curSaldo, "Standard"), "I", pCharRelleno, 15) '
strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15)  '
strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15)  '
strLinea(5) = "Capitalización    ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza * -1, "Standard"), "I", pCharRelleno, 15) '
strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15)  '
strLinea(7) = "Pólizas           ..: " & SIFGlobal.fxStringRelleno(Format(curPoliza, "Standard"), "I", pCharRelleno, 15)  '


strLinea(8) = "Operacion/Línea   ..: " & "Op.:" & txtOperacion.Text & " L.:" & txtCodigo & "-" & UCase(txtOpex.Text)
strLinea(9) = "Descripción       ..: " & txtLineaDesc.Text
strLinea(10) = ""


If chkTrasladar.Value = vbChecked Then
 strLinea(11) = chkTrasladar.Caption
Else
 strLinea(11) = Trim(txtLineaDesc.Text)
End If

If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdOperacionFechaProxPago " & txtOperacion.Text
    Call OpenRecordSet(rsTmp, strSQL, 0)
    If Not IsNull(rsTmp!Fecha_Pago) Then
         strLinea(9) = "Prox.Pago..:" & Format(rsTmp!Fecha_Pago, "dd/mm/yyyy") & " Cta.(" & rsTmp!Num_Cuota & ") " & Format(rsTmp!Cuota, "Standard")
    Else
         strLinea(9) = "Prox.Pago..: >> <<"
    End If
    strLinea(10) = "Notas: " & rsTmp!Notas & ""
    rsTmp.Close
End If
        

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)



  'Control de Documentos v2
   
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento,Linea11)" _
                & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
                & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curAmortiza & ",'P','" & txtOperacion.Text _
                & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "','" & strLinea(11) & "')"
        
        'ASIENTO
        If curAmortiza > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza * pTipoCambio & ",'D','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If

        
        
        If curIntC > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If
        
        If curIntM > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If
        
        If GLOBALES.SysPlanPagos = 0 Then
                If curCargo > 0 Then
                  strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curCargo * pTipoCambio & ",'C','" & rs!cod_Divisa _
                         & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!CtaCargos _
                         & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
                End If
        Else
                If curCargo > 0 Then
                'Detallar Cargos
                  glogon.strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
                  Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
                  Do While Not rsTmp.EOF
                        strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Mov_Monto * pTipoCambio & ",'C','" & rs!cod_Divisa _
                               & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                               & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
'                        Call ConectionExecute(strSQL)
                        rsTmp.MoveNext
                  Loop
                  rsTmp.Close
                End If
        End If
        
        If curPoliza > 0 Then
          glogon.strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!Id_Solicitud & ") as 'Cuenta'"
          Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
            vCuentaPoliza = Trim(rsTmp!Cuenta)
          rsTmp.Close
          
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curPoliza * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuentaPoliza _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
         ' Call ConectionExecute(strSQL)
        End If
        
       'Procesa Lote
       Call ConectionExecute(strSQL)
  

rs.Close


End Sub



Private Sub btnCajas_Click()
Dim vTotalPagar As Currency

'Capitaliza Deuda
If tcMain.Item(0).Visible Then
    If Not IsNumeric(txtTotalPagar.Text) Then txtTotalPagar.Text = 0
    ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text) - CCur(txtAmortiza.Text) + CCur(txtSaldo.Text)
    vTotalPagar = ModuloCajas.mTotalAplicar
End If

'Abono Especial
If tcMain.Item(3).Visible Then
    If Not IsNumeric(txtAE_Total.Text) Then txtAE_Total.Text = 0
    ModuloCajas.mTotalAplicar = CCur(txtAE_Total.Text)
    vTotalPagar = CCur(txtAE_Total.Text)
End If


If ModuloCajas.mTotalAplicar = 0 Then
    MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
    Exit Sub
End If

ModuloCajas.mServicio = "Arreglos de Pago"

Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)

txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")


If txtTotalCajas.Text <> vTotalPagar Then
   txtTotalCajas.BackColor = vbRed
Else
   txtTotalCajas.BackColor = vbWhite
End If

End Sub


Private Sub cboAE_CuotaFecha_Click()
On Error GoTo vError

txtAE_CuotaNum.Text = cboAE_CuotaFecha.ItemData(cboAE_CuotaFecha.ListIndex)

Exit Sub
vError:
  txtAE_CuotaNum.Text = 0

End Sub

Private Sub cboAE_Tipo_Click()

If cboAE_Tipo.ListCount = 0 Or vPaso Then Exit Sub


lblAE_Titulo.Item(0).Visible = False
lblAE_Titulo.Item(1).Visible = False
cboAE_CuotaFecha.Visible = False
txtAE_CuotaNum.Visible = False

txtAE_Cargos.Text = 0
txtAE_CuotaNum.Text = 0
txtAE_IntCor.Text = 0
txtAE_IntMor.Text = 0
txtAE_Principal.Text = 0
txtAE_Polizas.Text = 0
txtAE_Total.Text = 0


txtAE_IntMor.Locked = True
txtAE_Polizas.Locked = True
txtAE_Cargos.Locked = True



If Mid(cboAE_Tipo.Text, 1, 1) = "O" Then
    lblAE_Titulo.Item(0).Visible = True
    lblAE_Titulo.Item(1).Visible = True
    cboAE_CuotaFecha.Visible = True
    txtAE_CuotaNum.Visible = True
    
    txtAE_IntMor.Locked = False
    txtAE_Polizas.Locked = False
    txtAE_Cargos.Locked = False

    Call cboAE_CuotaFecha_Click

End If

End Sub

Private Sub cboTipoAplicacion_Click()

If cboTipoAplicacion.ListCount = 0 Or vPaso Then Exit Sub


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

Private Sub chkTipoIntereses_Click()
If IsNumeric(txtOperacion.Text) Then Call sbCargaMora
End Sub


Private Sub sbCapitalizaDeuda()
Dim curIntC As Currency, curIntM As Currency, curAmortiza As Currency, curCargo As Currency, curPoliza As Currency
Dim i As Integer, IdMoro As Long, vProceso As Long, strSQL As String

Dim curTotal As Currency, curDeuda As Currency
Dim vNumArreglo As Long, vCtaVencida As Integer

Dim vTipoDoc As String, vNumDoc As String, vConcepto As String, vCuenta As String

On Error GoTo vError

curTotal = CCur(txtTotalCajas.Text)

vNumArreglo = 0
vNumDoc = "0"
vCuenta = ""
vConcepto = "CRD014"


If curTotal > 0 Then
    vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)
'    vCuenta = Trim(fxDocumentoCuenta(vTipoDoc))
    vNumDoc = fxDocumentoConsecutivo(vTipoDoc)
End If


'If vAseDocValido = False And curTotal > 0 Then
'  MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
'        & " válida para esta operación...", vbCritical
'  Exit Sub
'End If


Me.MousePointer = vbHourglass
 
  
If GLOBALES.SysPlanPagos = 1 Then
        If chkTipoIntereses.Value = vbChecked Then
           vCtaVencida = 0 'Las cuotas no son vencidas por que el cobro de intereses son a hoy
        Else
           vCtaVencida = 1 'Solo cuotas vencidas no las que estan en proceso de cobro con fecha de pago posterior a la fecha
        End If

        If curTotal > 0 Then 'Abono en Cajas: Aplicar Vertical
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & mOperacion & ",'CRD014','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & curTotal & ",'" & Format(mFecha, "yyyy/mm/dd hh:mm:ss") & "','','V'," & vCtaVencida
            Call ConectionExecute(strSQL)
        End If
        
        curDeuda = CCur(txtTotalPagar.Text) - curTotal
        
        '1. Cancela Deuda Activa (Intereses, Cargos y Pólizas)
        '2. Ajusta Saldo por el Monto Aplicado
        strSQL = ""
        If curDeuda > 0 Then
            vNumArreglo = fxDocumentoConsecutivo("REA")
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & mOperacion & ",'CBR011','" & glogon.Usuario & "','REA" _
                   & "','" & vNumArreglo & "'," & curDeuda & ",'" & Format(mFecha, "yyyy/mm/dd hh:mm:ss") & "','','V'," & vCtaVencida
            
            strSQL = strSQL & Space(10) & "exec spCrdPlanPagoAnulaAbono " & mOperacion & ",'CBR011','" & glogon.Usuario & "','REA','" & vNumArreglo & "',1,0" _
                   & ",0," & curDeuda & ",0,0,'" & Format(mFecha, "yyyy/mm/dd hh:mm:ss") & "',''"
            Call ConectionExecute(strSQL)
        End If
            
        If chkTrasladar.Value = vbChecked And curDeuda > 0 Then
           strSQL = "exec spCrdPlanPagoPrincipalTraslado " & mOperacion
           glogon.Conection.Execute strSQL
        End If
        
        
Else
          
           'Arreglo con Capitalización sin Plan de Pagos
           strSQL = "exec spCrdOperacionArreglo_Capitaliza " & mOperacion & ",'" & vTipoDoc & "','" & vNumDoc & "'," & curTotal _
                  & ",'" & glogon.Usuario & "','" & ModuloCajas.mCaja & "'," & chkTrasladar.Value
           Call OpenRecordSet(glogon.Recordset, strSQL)
              vNumArreglo = glogon.Recordset!NumDoc
           glogon.Recordset.Close
        
End If 'Plan de Pagos

'Crea el documento
If curTotal > 0 Then
  Call sbDocumentoAbono(vTipoDoc, vNumDoc, "CRD014", vCuenta)
End If

If vNumArreglo > 0 Then
'    vNumArreglo = fxDocumentoAbono("ARREGLO DE PAGO", "REA", CStr(vNumArreglo), "CBR011", vCuenta)
    Call sbDocumentoReadecuacion("REA", CStr(vNumArreglo), "CBR011", vCuenta)
End If

'Imprime Comprobante
If curTotal > 0 Then Call sbImprimeRecibo(vNumDoc, vTipoDoc)
If vNumArreglo > 0 Then Call sbImprimeRecibo(CStr(vNumArreglo), "REA")


Call Bitacora("Registra", "Arreglo de Pago Op " & mOperacion)

Me.MousePointer = vbDefault
If vNumDoc <> "" And vNumArreglo > 0 Then
    MsgBox "Arreglo de Pago Realizado... " & cboTipoDoc.Text & " # " & vNumDoc & ", y Capitalización con Nota : " & vNumArreglo, vbInformation
End If

If vNumDoc = 0 And vNumArreglo > 0 Then
    MsgBox "Arreglo de Pago Realizado con Nota : " & vNumArreglo, vbInformation
End If


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbPeriodoGracia()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spCrd_Operacion_Arreglos_Periodo_Gracia " & mOperacion _
       & ",'" & Mid(cboTipoAplicacion.Text, 1, 1) & "'," & chkAplIntereses.Value _
       & "," & chkAplCargos.Value & "," & chkAplPolizas.Value _
       & "," & chkAplRetroactivo.Value & "," & chkAplAjustaPlazo.Value _
       & ",'" & Format(dtpAplInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
       & ",'" & Format(dtpAplCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
       & ",'" & glogon.Usuario & "','" & txtNotas.Text & "'"
      
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Periodo de Gracia, Operación: " & mOperacion & " Cta Rang: " _
        & Format(dtpAplInicio.Value, "dd/mm/yyyyy") & " - " & Format(dtpAplCorte.Value, "dd/mm/yyyyy"))


Me.MousePointer = vbDefault

MsgBox "Periodo de Gracia aplicado satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbVencimientoInt()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrdOperacionArreglo_InteresVence " & mOperacion & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','" _
        & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

    Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc)

rs.Close


Call Bitacora("Registra", "Vencimiento de Intereses, Operación: " & mOperacion & " Corte: " & Format(dtpCorte.Value, "dd/mm/yyyyy"))


Me.MousePointer = vbDefault
MsgBox "Vencimiento de Intereses aplicado satisfactoriamente!", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbDocumentoAbono(vTipoDoc As String, vNumDoc As String _
                                , pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency
Dim curSaldo As Currency, pTipoCambio As Currency


vCuenta = pCuenta

pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)

If GLOBALES.SysPlanPagos = 0 Then
    strSQL = "exec spCrdDocumentoAfectacionStP '" & vTipoDoc & "','" & vNumDoc & "','R'"
Else
    strSQL = "exec spCrdDocumentoAfectacion '" & vTipoDoc & "','" & vNumDoc & "','R'"
End If

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

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
      
 

strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(txtSaldo.Text, "I", pCharRelleno, 15) '
strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(CCur(txtSaldo.Text) - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15) '
strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15) '
strLinea(5) = "Amortización      ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15) '
strLinea(7) = "Pólizas           ..: " & SIFGlobal.fxStringRelleno(Format(curPoliza, "Standard"), "I", pCharRelleno, 15) '


strLinea(8) = "Operacion/Línea   ..: " & "Op.:" & txtOperacion.Text & " L.:" & txtCodigo & "-" & UCase(txtOpex.Text)
strLinea(9) = "Descripción       ..: " & txtLineaDesc.Text
strLinea(10) = ""

  'Control de Documentos v2
   
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento,cod_caja,cod_apertura)" _
                & " values('" & vNumDoc & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
                & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
                & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" & strLinea(11) & "','" _
                & txtNotas.Text & "','" & vAseDocDeposito & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ")"
'        Call ConectionExecute(strSQL)
        
        If curIntC > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntC * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If
        
        If curIntM > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curIntM * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If
        
        
        If GLOBALES.SysPlanPagos = 0 Then
                If curCargo > 0 Then
                    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curCargo * pTipoCambio & ",'C','" & rs!cod_Divisa _
                           & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!CtaCargos _
                           & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
                End If
        Else
                If curCargo > 0 Then
                'Detallar Cargos
                  glogon.strSQL = "exec spCrdDocumentoAfectacionCargos '" & vTipoDoc & "','" & vNumDoc & "'"
                  Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
                  Do While Not rsTmp.EOF
                        strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & rsTmp!Mov_Monto * pTipoCambio & ",'C','" & rs!cod_Divisa _
                               & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                               & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
'                        Call ConectionExecute(strSQL)
                        rsTmp.MoveNext
                  Loop
                  rsTmp.Close
                End If
        End If
        
        If curPoliza > 0 Then
          glogon.strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!Id_Solicitud & ") as 'Cuenta'"
          Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
            vCuenta = Trim(rsTmp!Cuenta)
          rsTmp.Close
          
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curPoliza * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
         ' Call ConectionExecute(strSQL)
        End If
        
        
        
        If curAmortiza > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & vTipoDoc & "','" & vNumDoc & "'," & curAmortiza * pTipoCambio & ",'C','" & rs!cod_Divisa _
                 & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If

       If curIntC + curIntM + curPoliza + curCargo + curAmortiza > 0 Then
            'Procesa Formas de Pago (Registro Final / Asiento de Pago)
             strSQL = strSQL & Space(10) & "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
                     & "','" & ModuloCajas.mUsuario & "','" & vTipoDoc & "','" & vNumDoc & "','" & ModuloCajas.mUnidad _
                     & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "'"
'             Call ConectionExecute(strSQL)
       End If
       
       'Aplicación en una sola llamada
       Call ConectionExecute(strSQL)


rs.Close


End Sub




Private Sub sbAbonoEspecial()
Dim iRespuesta As Integer, strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date, vCuenta As String
Dim vTipoDoc As String, vNumDoc As String, vConcepto As String

On Error GoTo vError

Me.MousePointer = vbHourglass

If mOperacion = 0 Then
    MsgBox "Ingrese un número de operacion válido...", vbCritical
    Exit Sub
End If

If CCur(txtAE_Principal.Text) > mSaldo Then
    MsgBox "La Amortización Especificada es mayor al Saldo, verifique...", vbCritical
    Exit Sub
End If

vFecha = fxFechaServidor
vConcepto = "CRD007"
vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)


If GLOBALES.SysPlanPagos = 0 And lsw.ListItems.Count > 0 Then
      MsgBox "No se puede Aplicar Abono Especial porque esta operación se encuentra en mora...", vbCritical
      Exit Sub
End If

iRespuesta = MsgBox("Esta seguro de realizar abono especial esta operación: " & mOperacion, vbYesNo)

If iRespuesta = vbNo Then
  Me.MousePointer = vbDefault
  Exit Sub
End If

vNumDoc = fxDocumentoConsecutivo(vTipoDoc)

Dim pProceso As Long, pNumCta As Integer

If GLOBALES.SysPlanPagos = 1 Then
'spCrdPlanPagoAbonoEspecial(@Operacion int, @Concepto varchar(10), @Usuario varchar(30)
'        , @TipoCom varchar(10), @NumCom varchar(30), @Dias smallint, @IntCor dec(16,2), @IntMor dec(16,2), @Principal dec(16,2)
'        , @Cargo dec(16,2), @Poliza dec(16,2), @FechaChr datetime, @Caja varchar(10)
'        , @Proceso int = 0, @NumCuota smallint = 0,@ReCalculaCta smallint = 0, @Actualiza smallint = 1, @EliminaCargos smallint = 0)
        If Mid(cboAE_Tipo.Text, 1, 1) = "E" Then
            pProceso = 0
            pNumCta = 0
        Else
            pProceso = Replace(cboAE_CuotaFecha.Text, "-", "")
            pNumCta = txtAE_CuotaNum.Text
        End If
        
        strSQL = "exec spCrdPlanPagoAbonoEspecial " & mOperacion & ",'" & vConcepto & "','" & glogon.Usuario & "','" & vTipoDoc _
               & "','" & vNumDoc & "',0," & CCur(txtAE_IntCor.Text) & "," & CCur(txtAE_IntMor.Text) & "," & CCur(txtAE_Principal.Text) _
               & "," & CCur(txtAE_Cargos.Text) & "," & CCur(txtAE_Polizas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',''" _
               & "," & pProceso & "," & pNumCta & ",1,1,1"
        
        strSQL = strSQL & Space(10) & "exec spCrdPlanPagos " & mOperacion
        Call ConectionExecute(strSQL)
 Else
    'Sin Plan de Pagos
    strSQL = "Update reg_creditos set estado = '" & IIf((CCur(txtAE_Principal) >= mSaldo), "C", "A") & "'," _
           & "SALDO = SALDO - " & CCur(txtAE_Principal) & ",AMORTIZA=AMORTIZA + " & CCur(txtAE_Principal) _
           & ",interesc = interesc + " & CCur(txtAE_IntCor) _
           & " where id_solicitud = " & mOperacion
    
    strSQL = strSQL & Space(10) & "INSERT CREDITOS_DT(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS,FECHAP,TCON,NCON,ESTADO" _
            & ",cod_concepto,usuario,cod_caja)" _
           & " values('" & txtCodigo.Text & "'," & mOperacion & ",0," & CCur(txtAE_Principal) + CCur(txtAE_IntCor) _
           & "," & CCur(txtAE_IntCor) & "," & CCur(txtAE_Principal) & ",dbo.MyGetdate()," & GLOBALES.glngFechaCR _
           & ",'" & vTipoDoc & "'," & vNumDoc & ",'A','" & vConcepto & "','" & glogon.Usuario & "','" & ModuloCajas.mCaja & "')"
    Call ConectionExecute(strSQL)
End If


'Comprobante + Asiento
Call sbDocumentoAbono(vTipoDoc, vNumDoc, vConcepto, "")

Call sbBitacoraCredito("11", ("Int: " & txtAE_IntCor & " Amort: " & txtAE_Principal), "C", txtOperacion, txtCodigo.Text)
Call Bitacora("Aplica", "Abono Especial de la operación :" & mOperacion)

Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Me.MousePointer = vbDefault
MsgBox "Abono Especial aplicado satisfactoriamente!", vbInformation


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdAcepta_Click()

If Not IsNumeric(txtOperacion.Text) Then
      MsgBox "Indique una Operación válida!", vbExclamation
      Exit Sub
End If


'--Validación General
If Len(txtNotas.Text) < 10 Then
      MsgBox "Indique una nota válida para la transacción.", vbExclamation
      Exit Sub
End If

'Verificar Congelamiento
If CCur(txtTotalCajas.Text) > 0 Then
    If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
      MsgBox "Esta Persona se encuentra CONGELADA, No puede realizar movimientos en Cajas. Verifique!", vbExclamation
      Exit Sub
    End If
End If


'Abono Especial
If tcMain.Item(3).Visible Then
   If CCur(txtAE_Total.Text) = 0 Then
      MsgBox "No se ha especificado ningún rubro para el abono especial!", vbExclamation
      Exit Sub
   End If
   
   If CCur(txtAE_Total.Text) > CCur(txtTotalCajas.Text) Then
      MsgBox "Monto en Cajas no corresponde al monto a recaudar para el abono especial!", vbExclamation
      Exit Sub
   End If
   
   If mRetencion Then
      MsgBox "No se pueden realizar abonos especiales a Retenciones!", vbExclamation
      Exit Sub
   End If
   
End If

'Capitalización de Deuda
If tcMain.Item(0).Visible Then
    If CCur(txtTotalCajas.Text) > CCur(txtIntCor.Text) + CCur(txtCargos.Text) + CCur(txtIntMor.Text) + CCur(txtSaldo.Text) + CCur(txtPolizas.Text) Then
      MsgBox "Total en Cajas que es mayor la Deuda...", vbExclamation
      Exit Sub
    End If

   If mRetencion Then
      MsgBox "No se pueden procesar catalización de deudas a Retenciones!", vbExclamation
      Exit Sub
   End If
   
   If lsw.ListItems.Count = 0 Then
      MsgBox "Esta Operación no puede realizar una capitalización de deuda porque está al día?", vbExclamation
      Exit Sub
   End If

End If

mFecha = fxFechaServidor

Select Case True
  Case tcMain.Item(0).Visible
     Call sbCapitalizaDeuda
  
  Case tcMain.Item(1).Visible
     
    If lsw.ListItems.Count > 0 And chkAplRetroactivo.Value = xtpUnchecked Then
       MsgBox "A esta Operación no se le puede dar periodo de Gracia porque NO está al día?", vbExclamation
       Exit Sub
    End If
   
     Call sbPeriodoGracia
   
  Case tcMain.Item(2).Visible
     Call sbVencimientoInt
     
  Case tcMain.Item(3).Visible
     Call sbAbonoEspecial
  
End Select

Call sbConsulta

End Sub


Private Sub sbCajaInicial()
Dim strSQL As String

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Arreglos de Pago   ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'idX', rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)


ModuloCajas.mServicio = "Arreglos de Pago"

If IsNumeric(ModuloCajas.mRef_01) Then
    txtOperacion.Text = ModuloCajas.mRef_01
    mOperacion = txtOperacion.Text
    Call sbConsulta
End If

End Sub

Private Sub sbInicializa()

Me.MousePointer = vbHourglass

Call sbCajaInicial

If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If

'Fix de la Base de datos
glogon.strSQL = "update  REG_CREDITOS set FECULT = dbo.fxSIFPrmProcesoAnt(prideduc)" _
              & " where ESTADO = 'A' and isnull(FECULT,0) = 0"
Call ConectionExecute(glogon.strSQL)

vPaso = True

    mFecha = fxFechaServidor
    
    dtpAplInicio.Value = mFecha
    dtpAplCorte.Value = mFecha
    dtpCorte.Value = mFecha
    
    
    cboTipoAplicacion.Clear
    cboTipoAplicacion.AddItem "TOTAL"
    cboTipoAplicacion.AddItem "PARCIAL"
    cboTipoAplicacion.Text = "TOTAL"

vPaso = False

Call cboTipoAplicacion_Click

Call isButtonMain_Click(0)

Me.MousePointer = vbDefault

End Sub



Private Sub Form_Activate()
vModulo = 3
End Sub


Private Sub Form_Load()
vModulo = 3


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
    cboAE_Tipo.Clear
    cboAE_Tipo.AddItem "Extraordinario"
    cboAE_Tipo.AddItem "Ordinario"
vPaso = False

cboAE_Tipo.Text = "Extraordunario"

With lsw.ColumnHeaders
    .Clear
    .Add , , "[ID]", 900
    .Add , , "Operación", 1200
    .Add , , "Proceso", 1200, vbCenter
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Pólizas", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Estado", 1400, vbCenter
    
    .Add , , "[AB]Int.Cor.", 1800, vbRightJustify
    .Add , , "[AB]Int.Mor.", 1800, vbRightJustify
    .Add , , "[AB]Cargos", 1800, vbRightJustify
    .Add , , "[AB]Pólizas", 1800, vbRightJustify
    .Add , , "[AB]Principal", 1800, vbRightJustify
    .Add , , "[AB]Total", 1800, vbRightJustify
    
End With


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub isButtonMain_Click(Index As Integer)
Dim pTop As Integer, pLeft As Integer
Dim i As Integer

If Not IsNumeric(txtOperacion.Text) Then Exit Sub

tcMain.Visible = True
tcInfo.Visible = True

pTop = 3360
pLeft = 1560

For i = 0 To tcMain.ItemCount - 1
    tcMain.Item(i).Visible = False
Next i

tcMain.Item(Index).Visible = True
tcMain.Item(Index).Selected = True

dtpCorte.Value = mFecha

btnCajas.Visible = False

Select Case Index
  Case 0 'Capitaliza Deuda
    btnCajas.Visible = True

  Case 1 'Periodos de Gracia
     
  Case 2 'Vencimiento de Intereses
     
     dtpCorte.Value = DateAdd("yyyy", -1, dtpCorte.Value)
  
  Case 3 'Abono Especial
     
     btnCajas.Visible = True
    
     cboAE_Tipo.Text = "Extraordinario"
      
     Call cboAE_Tipo_Click
     
End Select



End Sub




Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

txtOperacion.SetFocus
Call sbInicializa

End Sub


Private Sub txtAE_Cargos_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtAE_Polizas.SetFocus
Else
    txtAE_Total.Text = Format(CCur(txtAE_IntCor.Text) + CCur(txtAE_Principal) + CCur(txtAE_IntMor.Text) _
                    + CCur(txtAE_Cargos.Text) + CCur(txtAE_Polizas.Text), "Standard")
End If

vError:

End Sub

Private Sub txtAE_Cargos_LostFocus()
On Error GoTo vError

  txtAE_Cargos.Text = Format(CCur(txtAE_Cargos.Text), "Standard")
  
vError:
End Sub

Private Sub txtAE_IntCor_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtAE_IntMor.SetFocus
Else
    txtAE_Total.Text = Format(CCur(txtAE_IntCor.Text) + CCur(txtAE_Principal) + CCur(txtAE_IntMor.Text) _
                    + CCur(txtAE_Cargos.Text) + CCur(txtAE_Polizas.Text), "Standard")
End If

vError:

End Sub



Private Sub txtAE_IntCor_LostFocus()
On Error GoTo vError

  txtAE_IntCor.Text = Format(CCur(txtAE_IntCor.Text), "Standard")
  
vError:
End Sub



Private Sub txtAE_IntMor_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtAE_Principal.SetFocus
Else
    txtAE_Total.Text = Format(CCur(txtAE_IntCor.Text) + CCur(txtAE_Principal) + CCur(txtAE_IntMor.Text) _
                    + CCur(txtAE_Cargos.Text) + CCur(txtAE_Polizas.Text), "Standard")
End If

vError:

End Sub


Private Sub txtAE_IntMor_LostFocus()
On Error GoTo vError

  txtAE_IntMor.Text = Format(CCur(txtAE_IntMor.Text), "Standard")
  
vError:
End Sub

Private Sub txtAE_Polizas_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtAE_Total.SetFocus
Else
    txtAE_Total.Text = Format(CCur(txtAE_IntCor.Text) + CCur(txtAE_Principal) + CCur(txtAE_IntMor.Text) _
                    + CCur(txtAE_Cargos.Text) + CCur(txtAE_Polizas.Text), "Standard")
End If

vError:

End Sub

Private Sub txtAE_Polizas_LostFocus()
On Error GoTo vError

  txtAE_Polizas.Text = Format(CCur(txtAE_Polizas.Text), "Standard")
  
vError:

End Sub

Private Sub txtAE_Principal_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then
    txtAE_Cargos.SetFocus
Else
    txtAE_Total.Text = Format(CCur(txtAE_IntCor.Text) + CCur(txtAE_Principal) + CCur(txtAE_IntMor.Text) _
                    + CCur(txtAE_Cargos.Text) + CCur(txtAE_Polizas.Text), "Standard")
End If

vError:

End Sub

Private Sub txtAE_Principal_LostFocus()
On Error GoTo vError

  txtAE_Principal.Text = Format(CCur(txtAE_Principal.Text), "Standard")
  
vError:
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtNombre.SetFocus

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtLineaDesc.SetFocus

End Sub

Private Sub txtLineaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtIntCor.SetFocus

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtCodigo.SetFocus

End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyTab Or KeyCode = vbKeyReturn Then txtCedula.SetFocus

End Sub

Private Sub txtOperacion_LostFocus()

If vPaso Then Exit Sub

If IsNumeric(txtOperacion.Text) Then
  Call sbConsulta
Else
  Call sbLimpia
End If
  
End Sub
