VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_Bancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bancos [Cuentas Bancarias]"
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   Icon            =   "frmTES_Bancos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   9360
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   120
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6612
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9132
      _Version        =   1441793
      _ExtentX        =   16108
      _ExtentY        =   11663
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
      Item(0).Caption =   "Información Cuenta"
      Item(0).ControlCount=   23
      Item(0).Control(0)=   "cboGrupo"
      Item(0).Control(1)=   "txtDescCorta"
      Item(0).Control(2)=   "txtCuentaBancaria"
      Item(0).Control(3)=   "Label6(4)"
      Item(0).Control(4)=   "Label6(0)"
      Item(0).Control(5)=   "Label6(1)"
      Item(0).Control(6)=   "Label6(2)"
      Item(0).Control(7)=   "txtLugarEmision"
      Item(0).Control(8)=   "Label12(0)"
      Item(0).Control(9)=   "Label7(0)"
      Item(0).Control(10)=   "Label12(2)"
      Item(0).Control(11)=   "txtSINPE_Codigo"
      Item(0).Control(12)=   "chkSINPE_CtaInterna"
      Item(0).Control(13)=   "Label7(4)"
      Item(0).Control(14)=   "GroupBox1(0)"
      Item(0).Control(15)=   "GroupBox1(1)"
      Item(0).Control(16)=   "txtCodigoCliente"
      Item(0).Control(17)=   "Label7(5)"
      Item(0).Control(18)=   "cboFormato"
      Item(0).Control(19)=   "cboFormatoN2"
      Item(0).Control(20)=   "cboEstado"
      Item(0).Control(21)=   "chkUtilizaPlan"
      Item(0).Control(22)=   "btnPlanes"
      Item(1).Caption =   "Formato y Firmas"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "GroupBox2(0)"
      Item(1).Control(1)=   "GroupBox2(1)"
      Item(2).Caption =   "Monitoreo Saldos"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "lsw"
      Item(2).Control(1)=   "Label2(2)"
      Item(2).Control(2)=   "GroupBox3"
      Item(3).Caption =   "Conciliación"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "gbConciliacion"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4332
         Left            =   -69880
         TabIndex        =   63
         Top             =   840
         Visible         =   0   'False
         Width           =   8892
         _Version        =   1441793
         _ExtentX        =   15684
         _ExtentY        =   7641
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbConciliacion 
         Height          =   6012
         Left            =   -69760
         TabIndex        =   64
         Top             =   480
         Visible         =   0   'False
         Width           =   8652
         _Version        =   1441793
         _ExtentX        =   15261
         _ExtentY        =   10604
         _StockProps     =   79
         Caption         =   "Reglas para el Proceso de Conciliación Bancaria"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.FlatEdit txtCon_ComisionSINPECtaDesc 
            Height          =   312
            Left            =   2520
            TabIndex        =   68
            Top             =   1560
            Width           =   5772
            _Version        =   1441793
            _ExtentX        =   10181
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCon_ComisionSINPECta 
            Height          =   312
            Left            =   720
            TabIndex        =   67
            Top             =   1560
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCon_ComisionSINPEMnt 
            Height          =   312
            Left            =   720
            TabIndex        =   69
            Top             =   840
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboCon_Unidad 
            Height          =   312
            Left            =   2520
            TabIndex        =   72
            Top             =   2520
            Width           =   5772
            _Version        =   1441793
            _ExtentX        =   10186
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
         Begin XtremeSuiteControls.ComboBox cboCon_Concepto 
            Height          =   312
            Left            =   2520
            TabIndex        =   73
            Top             =   4800
            Width           =   5772
            _Version        =   1441793
            _ExtentX        =   10186
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
         Begin XtremeSuiteControls.PushButton btnConciliacion 
            Height          =   372
            Left            =   6720
            TabIndex        =   74
            Top             =   5400
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Actualizar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin XtremeSuiteControls.ComboBox cboCon_Centro 
            Height          =   312
            Left            =   2520
            TabIndex        =   75
            Top             =   3360
            Width           =   5772
            _Version        =   1441793
            _ExtentX        =   10186
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
         Begin XtremeSuiteControls.ComboBox cboCon_Centro_Comision 
            Height          =   312
            Left            =   2520
            TabIndex        =   77
            Top             =   4080
            Width           =   5772
            _Version        =   1441793
            _ExtentX        =   10186
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   252
            Index           =   5
            Left            =   720
            TabIndex        =   78
            Top             =   3720
            Width           =   5052
            _Version        =   1441793
            _ExtentX        =   8911
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Centro de Costo para Comisiones de Transferencias:"
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
            Height          =   252
            Index           =   4
            Left            =   720
            TabIndex        =   76
            Top             =   3000
            Width           =   5052
            _Version        =   1441793
            _ExtentX        =   8911
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Centro de Costo por Omisión para Auto-Registro:"
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
            Height          =   252
            Index           =   3
            Left            =   720
            TabIndex        =   71
            Top             =   4440
            Width           =   5052
            _Version        =   1441793
            _ExtentX        =   8911
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Concepto por Omisión para Auto-Registro:"
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
            Height          =   252
            Index           =   2
            Left            =   720
            TabIndex        =   70
            Top             =   2160
            Width           =   5052
            _Version        =   1441793
            _ExtentX        =   8911
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Unidad por Omisión para Auto-Registro:"
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
            Height          =   252
            Index           =   1
            Left            =   720
            TabIndex        =   66
            Top             =   1200
            Width           =   7092
            _Version        =   1441793
            _ExtentX        =   12509
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta para el  Auto-Registro de Comisiones de Transferencias:"
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
            Height          =   252
            Index           =   0
            Left            =   720
            TabIndex        =   65
            Top             =   480
            Width           =   6972
            _Version        =   1441793
            _ExtentX        =   12298
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Monto para Auto-Registro de Comisiones de Transferencias:"
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
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1212
         Left            =   -69760
         TabIndex        =   25
         Top             =   5280
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1441793
         _ExtentX        =   15049
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Información de Saldos de la Cuenta Bancaria:"
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
         Begin XtremeSuiteControls.PushButton cmdCorregir 
            Height          =   372
            Left            =   7200
            TabIndex        =   26
            Top             =   480
            Width           =   1212
            _Version        =   1441793
            _ExtentX        =   2138
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Actualizar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin XtremeSuiteControls.DateTimePicker dtpFecha 
            Height          =   312
            Left            =   1800
            TabIndex        =   44
            Top             =   480
            Width           =   1332
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
            Left            =   4200
            TabIndex        =   62
            Top             =   480
            Width           =   2532
            _Version        =   1441793
            _ExtentX        =   4466
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Ultimo Corte"
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
            Left            =   360
            TabIndex        =   28
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label6 
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
            Index           =   3
            Left            =   3360
            TabIndex        =   27
            Top             =   480
            Width           =   492
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   2892
         Index           =   0
         Left            =   -69520
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   8292
         _Version        =   1441793
         _ExtentX        =   14626
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Formatos de Cheques:"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.PushButton btnArchivo_Busca 
            Height          =   315
            Index           =   0
            Left            =   7200
            TabIndex        =   56
            Top             =   1080
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.FlatEdit txtArchivoEspecial 
            Height          =   312
            Left            =   1920
            TabIndex        =   53
            Top             =   1080
            Width           =   5172
            _Version        =   1441793
            _ExtentX        =   9123
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtChequeEspecialFirma 
            Height          =   312
            Left            =   1920
            TabIndex        =   54
            Top             =   1680
            Width           =   5172
            _Version        =   1441793
            _ExtentX        =   9123
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtChequeEspecialNoFirma 
            Height          =   312
            Left            =   1920
            TabIndex        =   55
            Top             =   2280
            Width           =   5172
            _Version        =   1441793
            _ExtentX        =   9123
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnArchivo_Busca 
            Height          =   312
            Index           =   1
            Left            =   7200
            TabIndex        =   57
            Top             =   1680
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnArchivo_Busca 
            Height          =   312
            Index           =   2
            Left            =   7200
            TabIndex        =   58
            Top             =   2280
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.CheckBox chKFormatoEspecial 
            Height          =   492
            Left            =   3000
            TabIndex        =   61
            Top             =   240
            Width           =   4092
            _Version        =   1441793
            _ExtentX        =   7218
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Utilizar formatos personalizados para la impresion de documentos"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Alignment       =   1
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo Especial para Impresión de Cheques Sin Firmas"
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
            Left            =   600
            TabIndex        =   7
            Top             =   2040
            Width           =   4692
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo Especial para Impresión de Cheques Con Firmas"
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
            Left            =   600
            TabIndex        =   6
            Top             =   1440
            Width           =   4692
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo Especial para Impresión de Documento"
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
            Left            =   600
            TabIndex        =   5
            Top             =   840
            Width           =   4092
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   972
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   3240
         Width           =   8892
         _Version        =   1441793
         _ExtentX        =   15684
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Cuenta Contable:"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   330
            Left            =   6720
            TabIndex        =   35
            Top             =   240
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaContable 
            Height          =   312
            Left            =   360
            TabIndex        =   42
            Top             =   600
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaContaDesc 
            Height          =   315
            Left            =   2520
            TabIndex        =   43
            Top             =   600
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
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
            Index           =   7
            Left            =   5760
            TabIndex        =   36
            Top             =   240
            Width           =   852
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2052
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4440
         Width           =   8892
         _Version        =   1441793
         _ExtentX        =   15684
         _ExtentY        =   3619
         _StockProps     =   79
         Caption         =   "Comportamiento:"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.CheckBox chkMonitoreo 
            Height          =   252
            Left            =   4560
            TabIndex        =   37
            Top             =   360
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Monitorea Actividad de Saldos en Cuenta"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkRegional 
            Height          =   252
            Left            =   4560
            TabIndex        =   38
            Top             =   720
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta Corriente para Uso de Oficina Regional"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkCuentaBancariaPuente 
            Height          =   252
            Left            =   4560
            TabIndex        =   39
            Top             =   1080
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Esta cuenta se utiliza como puente?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkAutoGestion 
            Height          =   252
            Left            =   0
            TabIndex        =   40
            Top             =   360
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Habilitar Cuenta para AutoGestión en Apps/Web"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkSupervisa 
            Height          =   252
            Left            =   0
            TabIndex        =   41
            Top             =   720
            Width           =   4332
            _Version        =   1441793
            _ExtentX        =   7641
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Supervisar duplicidad de transacciones?"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDias 
            Height          =   312
            Left            =   8160
            TabIndex        =   52
            Top             =   1560
            Width           =   732
            _Version        =   1441793
            _ExtentX        =   1291
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
            Text            =   "0"
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox CheckBox1 
            Height          =   375
            Left            =   0
            TabIndex        =   81
            Top             =   1080
            Width           =   4335
            _Version        =   1441793
            _ExtentX        =   7646
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Habilitar SINPE ASOCIADOS en la carga de movimientos bancarios?"
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
            Alignment       =   1
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Días a validar:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   252
            Index           =   6
            Left            =   6720
            TabIndex        =   10
            Top             =   1560
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   2772
         Index           =   1
         Left            =   -69400
         TabIndex        =   11
         Top             =   3360
         Visible         =   0   'False
         Width           =   8172
         _Version        =   1441793
         _ExtentX        =   14414
         _ExtentY        =   4890
         _StockProps     =   79
         Caption         =   "Rango de Autorización de Documentos:"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.PushButton cmdActualizar 
            Height          =   372
            Left            =   5280
            TabIndex        =   23
            Top             =   2160
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Actualizar"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
         Begin XtremeSuiteControls.FlatEdit txtFirmaDesde 
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
            Left            =   3720
            TabIndex        =   59
            Top             =   1320
            Width           =   3132
            _Version        =   1441793
            _ExtentX        =   5524
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFirmaHasta 
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
            Left            =   3720
            TabIndex        =   60
            Top             =   1680
            Width           =   3132
            _Version        =   1441793
            _ExtentX        =   5524
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblDesde 
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Left            =   2640
            TabIndex        =   22
            Top             =   1320
            Width           =   852
         End
         Begin VB.Label lblHasta 
            BackStyle       =   0  'Transparent
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
            Height          =   252
            Left            =   2640
            TabIndex        =   21
            Top             =   1680
            Width           =   852
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   $"frmTES_Bancos.frx":6852
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   0
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   7332
         End
      End
      Begin XtremeSuiteControls.CheckBox chkSINPE_CtaInterna 
         Height          =   612
         Left            =   1680
         TabIndex        =   29
         Top             =   2520
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6583
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Esta es una Cuenta Interna (SINPE) Afecta Auxiliares Directamente"
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
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboGrupo 
         Height          =   312
         Left            =   1680
         TabIndex        =   31
         Top             =   600
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   7080
         TabIndex        =   32
         Top             =   600
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
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
      Begin XtremeSuiteControls.ComboBox cboFormato 
         Height          =   312
         Left            =   1680
         TabIndex        =   33
         Top             =   1320
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.ComboBox cboFormatoN2 
         Height          =   312
         Left            =   1680
         TabIndex        =   34
         Top             =   1680
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.FlatEdit txtCuentaBancaria 
         Height          =   312
         Left            =   1680
         TabIndex        =   47
         Top             =   960
         Width           =   3732
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDescCorta 
         Height          =   315
         Left            =   7080
         TabIndex        =   48
         Top             =   960
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLugarEmision 
         Height          =   315
         Left            =   7080
         TabIndex        =   49
         Top             =   1320
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigoCliente 
         Height          =   315
         Left            =   7080
         TabIndex        =   50
         Top             =   1680
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSINPE_Codigo 
         Height          =   315
         Left            =   7080
         TabIndex        =   51
         Top             =   2640
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1291
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkUtilizaPlan 
         Height          =   612
         Left            =   1680
         TabIndex        =   79
         Top             =   1920
         Width           =   3732
         _Version        =   1441793
         _ExtentX        =   6583
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Utiliza Consecutivos por Planes en el Banco"
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
         Appearance      =   2
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnPlanes 
         Height          =   375
         Left            =   7080
         TabIndex        =   80
         Top             =   2040
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Planes"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bancos.frx":68F0
         ImageAlignment  =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Formato TF No. 2"
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
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   1452
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Ultimos 30 Cortes [ver Monitoreo y Cierres]:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -69760
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   5412
      End
      Begin VB.Label Label6 
         Caption         =   "Grupo Bancario"
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
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado:"
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
         Left            =   6000
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Desc. Corta:"
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
         Left            =   5520
         TabIndex        =   17
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Cuenta Bancaria"
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
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Lugar Emisión:"
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
         Left            =   5520
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Formato TF No. 1"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   1452
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Código Cliente:"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Código SINPE :"
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
         Index           =   4
         Left            =   5520
         TabIndex        =   12
         Top             =   2640
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   264
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9864
      _ExtentX        =   17410
      _ExtentY        =   476
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta "
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repListadoGeneral"
                  Text            =   "Listado General"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8640
      TabIndex        =   2
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9480
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Bancos.frx":7009
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Bancos.frx":D86B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Bancos.frx":140CD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1200
      TabIndex        =   45
      Top             =   480
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   2040
      TabIndex        =   46
      Top             =   480
      Width           =   6492
      _Version        =   1441793
      _ExtentX        =   11451
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   972
   End
End
Attribute VB_Name = "frmTES_Bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vScroll As Boolean

Private Sub btnArchivo_Busca_Click(Index As Integer)
Dim vArchivoId As Integer

Select Case Index
    Case 0
      vArchivoId = 1
    Case 1
      vArchivoId = 2
    Case 2
      vArchivoId = 3
End Select

Call sbCargaArchivo(vArchivoId)
End Sub

Private Sub btnConciliacion_Click()
Call sbConciliacion_Update
End Sub



Private Sub btnPlanes_Click()
If IsNumeric(txtCodigo.Text) Then
    GLOBALES.gTag = txtCodigo.Text
    Call sbFormsCall("frmTES_TE_Planes", vbModal, , , False, Me)
End If
End Sub

Private Sub cboCon_Concepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Or KeyCode = vbKeyReturn And btnConciliacion.Enabled Then btnConciliacion.SetFocus
End Sub

Private Sub cboFormato_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigoCliente.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaBancaria.SetFocus
End Sub


Private Sub cboGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub



Private Sub chkSupervisa_Click()
If chkSupervisa.Value = vbChecked Then
 txtDias.Text = 5
 txtDias.Enabled = True
Else
 txtDias.Text = 0
 txtDias.Enabled = False
End If
End Sub

Private Sub chkUtilizaPlan_Click()
If chkUtilizaPlan.Value = xtpChecked Then
    btnPlanes.Visible = True
Else
    btnPlanes.Visible = False
End If
End Sub

Private Sub cmdActualizar_Click()
Dim strSQL As String

On Error GoTo vError

If vCodigo = 0 Then Exit Sub

Me.MousePointer = vbHourglass


If Trim(txtFirmaDesde) = "" Or Trim(txtFirmaHasta) = "" Then
   MsgBox "Suministre el Rango de Firmas", vbExclamation
   Me.MousePointer = vbDefault
   Exit Sub
End If

If CCur(txtFirmaDesde) > CCur(txtFirmaHasta) Then
   MsgBox "Verifique el Rango de Firmas", vbExclamation
   Me.MousePointer = vbDefault
   Exit Sub
End If

strSQL = "Update Tes_Bancos Set Firmas_Desde = " & CCur(txtFirmaDesde) & ",Firmas_Hasta=" & CCur(txtFirmaHasta) _
       & " Where ID_Banco = " & vCodigo
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Firmas Banco = " & Trim(txtDescCorta) & ", " & CCur(txtFirmaDesde) & " a " & CCur(txtFirmaHasta))
   
MsgBox "Firmas Actualizadas", vbExclamation, "Atencion!"


Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdCorregir_Click()
Dim strSQL As String

On Error GoTo vError

If vCodigo = 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "Update Tes_Bancos Set Fecha_Envia='" & Format(dtpFecha.Value, "yyyy/mm/dd") _
       & " 23:59:59',Saldo = " & CCur(txtSaldo.Text) & "  Where ID_Banco = " & vCodigo
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Cta.Id [" & vCodigo & "] Cta.Desc.: " & Trim(txtDescCorta) & ", Saldo: " & txtSaldo.Text)
   
MsgBox "Saldo y Fecha Corregidos", vbExclamation, "Atencion!"

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSaldo.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo = "" Then txtCodigo = "0"

If vScroll Then
    strSQL = "select Top 1 id_banco from Tes_Bancos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where id_banco > " & txtCodigo & " order by id_banco asc"
    Else
       strSQL = strSQL & " where id_banco < " & txtCodigo & " order by id_banco desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!Id_Banco
      Call sbConsulta(txtCodigo)
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

Private Sub Form_Activate()
vModulo = 9
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 9

On Error GoTo vError
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "[ID]", 900
    .Add , , "Inicio", 1800, vbCenter
    .Add , , "Corte", 1800, vbCenter
    .Add , , "Saldo Inicial", 2100, vbRightJustify
    .Add , , "Total Débitos", 2100, vbRightJustify
    .Add , , "Total Créditos", 2100, vbRightJustify
    .Add , , "Saldo Final", 2100, vbRightJustify
    .Add , , "Ajustes", 2100, vbRightJustify
    .Add , , "Saldo Mínimo", 2100, vbRightJustify
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Usuario", 1800, vbCenter
End With
 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
cboEstado.Text = "Activo"
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()

tcMain.Item(0).Selected = True

vCodigo = 0
txtCodigo = ""

chkRegional.Value = vbUnchecked
chkMonitoreo.Value = vbChecked
chKFormatoEspecial = vbUnchecked

txtNombre.Text = ""
txtArchivoEspecial.Text = ""
txtChequeEspecialFirma.Text = ""
txtChequeEspecialNoFirma.Text = ""

txtCuentaContable.Text = ""
txtCtaContaDesc.Text = ""

txtDescCorta.Text = ""
txtCuentaBancaria.Text = ""

txtFirmaDesde.Text = Format(0, "Standard")
txtFirmaHasta.Text = Format(0, "Standard")

txtSaldo.Text = Format(0, "Standard")
dtpFecha.Value = fxFechaServidor

chkMonitoreo.Value = vbUnchecked
chkRegional.Value = vbUnchecked
chkCuentaBancariaPuente.Value = vbUnchecked
  

chkSINPE_CtaInterna.Value = vbUnchecked
txtSINPE_Codigo.Text = ""

txtCodigoCliente.Text = ""

txtCon_ComisionSINPECtaDesc.Text = ""
txtCon_ComisionSINPECta.Text = ""
txtCon_ComisionSINPEMnt.Text = Format(0, "Standard")


chkUtilizaPlan.Value = xtpChecked
chkUtilizaPlan_Click


txtCodigo.Enabled = True

End Sub

Private Sub sbConciliacion_Update()
Dim strSQL As String

On Error GoTo vError

strSQL = "Update tes_Bancos set CONCILIA_AR_COMISION = " & CCur(txtCon_ComisionSINPEMnt.Text) _
       & ", CONCILIA_AR_COMISION_CTA = '" & fxgCntCuentaFormato(False, txtCon_ComisionSINPECta.Text, 0) _
       & "', CONCILIA_AR_UNIDAD = '" & cboCon_Unidad.ItemData(cboCon_Unidad.ListIndex) _
       & "', CONCILIA_AR_CENTRO = '" & cboCon_Centro.ItemData(cboCon_Centro.ListIndex) _
       & "', CONCILIA_AR_CENTRO_COM = '" & cboCon_Centro_Comision.ItemData(cboCon_Centro_Comision.ListIndex) _
       & "', CONCILIA_AR_CONCEPTO = '" & cboCon_Concepto.ItemData(cboCon_Concepto.ListIndex) _
       & "' Where Id_Banco = " & txtCodigo.Text
    
Call ConectionExecute(strSQL)
Call Bitacora("Modifica", "Cta.Id [" & vCodigo & "] Cta.Desc.: " & Trim(txtDescCorta) & ", Comisión: " & txtCon_ComisionSINPEMnt.Text)
   
MsgBox "Reglas de Conciliación, Actualizadas!", vbInformation
    
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Select Case Item.Index
    
    Case 2 'Monitoreo
        
        If IsNumeric(txtCodigo.Text) Then

            strSQL = "select Top 30 * from TES_BANCOS_CIERRES where id_banco = " & txtCodigo.Text & " order by corte desc"
            Call OpenRecordSet(rs, strSQL, 0)
            Do While Not rs.EOF
             Set itmX = lsw.ListItems.Add(, , rs!IdX)
                 itmX.SubItems(1) = Format(rs!Inicio, "dd/mm/yyyy")
                 itmX.SubItems(2) = Format(rs!Corte, "dd/mm/yyyy")
                 itmX.SubItems(3) = Format(rs!saldo_inicial, "Standard")
                 itmX.SubItems(4) = Format(rs!total_debitos, "Standard")
                 itmX.SubItems(5) = Format(rs!total_creditos, "Standard")
                 itmX.SubItems(6) = Format(rs!saldo_final, "Standard")
                 itmX.SubItems(7) = Format(rs!ajuste, "Standard")
                 itmX.SubItems(8) = Format(rs!saldo_minimo, "Standard")
                 itmX.SubItems(9) = Format(rs!fecha, "dd/mm/yyyy")
                 itmX.SubItems(10) = rs!Usuario
             rs.MoveNext
            Loop
            rs.Close
     
        End If

    Case 3 'Conciliación
        'Nada
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

Me.MousePointer = vbHourglass

 strSQL = "select rtrim(cod_grupo) as 'Idx',rtrim(Descripcion) as ItmX from TES_BANCOS_GRUPOS" _
        & " where Activo = 1"
 Call sbCbo_Llena_New(cboGrupo, strSQL, False, True)
 
 strSQL = "exec spSys_Divisas"
 Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)
 
 strSQL = "select rtrim(cod_Formato) as 'Idx',rtrim(Descripcion) as ItmX " _
        & " from vTes_Formatos" _
        & " where Activo = 1"
 Call sbCbo_Llena_New(cboFormato, strSQL, False, True)
 
 Call sbCbo_Copia(cboFormato, cboFormatoN2)

 
 strSQL = "select rtrim(COD_UNIDAD) AS 'IdX', rtrim(DESCRIPCION) AS 'ItmX'" _
        & " From CNTX_UNIDADES" _
        & " where COD_CONTABILIDAD in(select COD_EMPRESA_ENLACE from SIF_EMPRESA)" _
        & "  and ACTIVA = 1" _
        & " order by UNIDAD_OMISION desc, DESCRIPCION asc"
 Call sbCbo_Llena_New(cboCon_Unidad, strSQL, False, True)

 strSQL = "select COD_CENTRO_COSTO AS 'IdX', RTRIM(DESCRIPCION) AS 'ItmX'" _
        & " From CNTX_CENTRO_COSTOS" _
        & " Where Activo = 1 And COD_CONTABILIDAD = 1"
 Call sbCbo_Llena_New(cboCon_Centro, strSQL, False, True)

 Call sbCbo_Copia(cboCon_Centro, cboCon_Centro_Comision)

 Call sbCboAsignaDato(cboCon_Centro, "", True, "")
 Call sbCboAsignaDato(cboCon_Centro_Comision, "", True, "")

 strSQL = "select rtrim(COD_CONCEPTO) AS 'IdX', rtrim(DESCRIPCION) AS 'ItmX'" _
        & " From TES_CONCEPTOS" _
        & " where ESTADO = 'A'" _
        & " order by DESCRIPCION"
 Call sbCbo_Llena_New(cboCon_Concepto, strSQL, False, True)

 
tcMain.Item(0).Selected = True
 

Me.MousePointer = vbDefault

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
      Call txtCodigo_KeyDown(vbKeyF4, 1)
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(pBanco As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select B.*,rtrim(G.Descripcion) as GrupoX" _
       & ", dbo.fxTes_Formatos_Desc(B.Formato_Transferencia) as 'FormatoN1'" _
       & ", dbo.fxTes_Formatos_Desc(B.Formato_Transferencias_N2) as 'FormatoN2'" _
       & ", Dv.Descripcion as 'DivisaDesc'" _
       & ", isnull(Cb.Cod_Cuenta_Mask,'') as 'COD_CUENTA', isnull(Cb.Descripcion,'') as 'COD_CUENTA_DESC'" _
       & ", isnull(Cc.Cod_Cuenta_Mask,'') as 'COD_CUENTA_CON', isnull(Cc.Descripcion,'') as 'COD_CUENTA_CON_DESC'" _
       & ", ISNULL(Ud.COD_UNIDAD,'') AS 'UNIDAD', ISNULL(Ud.DESCRIPCION,'') AS 'UNIDAD_DESC'" _
       & ", ISNULL(Ccr.COD_CENTRO_COSTO,'') AS 'CENTRO', ISNULL(Ccr.DESCRIPCION,'') AS 'CENTRO_DESC'" _
       & ", ISNULL(Cct.COD_CENTRO_COSTO,'') AS 'CENTRO_COM', ISNULL(Cct.DESCRIPCION,'') AS 'CENTRO_COM_DESC'" _
       & ", ISNULL(Tc.COD_CONCEPTO,'') AS 'CONCEPTO', ISNULL(Tc.DESCRIPCION,'') AS 'CONCEPTO_DESC'" _
       & " from Tes_Bancos B left join TES_BANCOS_GRUPOS G on B.cod_Grupo = G.cod_Grupo" _
       & " left join CntX_Divisas Dv on B.cod_divisa = Dv.Cod_Divisa and Dv.cod_Contabilidad = " & GLOBALES.gEnlace _
       & " left join vCNTX_CUENTAS_LOCAL Cb on B.ctaConta = Cb.Cod_Cuenta" _
       & " left join vCNTX_CUENTAS_LOCAL Cc on B.CONCILIA_AR_COMISION_CTA = Cc.Cod_Cuenta" _
       & " left join CNTX_UNIDADES Ud on B.CONCILIA_AR_UNIDAD = Ud.COD_UNIDAD AND Ud.COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " left join CntX_Centro_Costos Ccr on B.CONCILIA_AR_CENTRO = Ccr.Cod_Centro_Costo AND Ccr.COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " left join CntX_Centro_Costos Cct on B.CONCILIA_AR_CENTRO_COM = Cct.Cod_Centro_Costo AND Cct.COD_CONTABILIDAD = " & GLOBALES.gEnlace _
       & " left join TES_CONCEPTOS Tc on B.CONCILIA_AR_CONCEPTO = Tc.COD_CONCEPTO" _
       & " where B.id_Banco = " & pBanco
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  
  tcMain.Item(0).Selected = True
  
  vEdita = True
  
  
  vCodigo = rs!Id_Banco
  txtCodigo = rs!Id_Banco
  
  txtNombre = rs!DESCRIPCION & ""
  txtCuentaBancaria = rs!Cta & ""
  txtDescCorta = rs!desc_corta & ""
  
  txtLugarEmision.Text = rs!Lugar_Emision & ""
  
  txtArchivoEspecial.Text = rs!ARCHIVO_ESPECIAL_CK & ""
  txtChequeEspecialFirma.Text = IIf(IsNull(rs!ARCHIVO_CHEQUES_FIRMAS), "", rs!ARCHIVO_CHEQUES_FIRMAS)
  txtChequeEspecialNoFirma.Text = IIf(IsNull(rs!ARCHIVO_CHEQUES_SIN_FIRMAS), "", rs!ARCHIVO_CHEQUES_SIN_FIRMAS)
    
  
  
  If rs!Estado = "A" Then
    cboEstado.Text = "Activo"
  Else
    cboEstado.Text = "Inactivo"
  End If
    
  chkRegional.Value = rs!cta_regional
  chkMonitoreo.Value = IIf(IsNull(rs!Monitoreo), 0, rs!Monitoreo)
  chkCuentaBancariaPuente.Value = IIf(IsNull(rs!puente), 0, rs!puente)
  chKFormatoEspecial.Value = IIf(IsNull(rs!UTILIZA_FORMATO_ESPECIAL), 0, rs!UTILIZA_FORMATO_ESPECIAL)
  chkSupervisa.Value = IIf(IsNull(rs!SUPERVISION), 0, rs!SUPERVISION)
  
  chkAutoGestion.Value = IIf(IsNull(rs!UTILIZA_AUTOGESTION), 0, rs!UTILIZA_AUTOGESTION)
  
  
  txtDias.Text = IIf(IsNull(rs!SUPERVISION_DIAS), 0, rs!SUPERVISION_DIAS)
  
  txtCuentaContable.Text = rs!cod_cuenta & ""
  txtCtaContaDesc.Text = rs!COD_CUENTA_DESC & ""
  
  If Not IsNull(rs!COD_GRUPO) Then
    Call sbCboAsignaDato(cboGrupo, rs!GrupoX, True, rs!COD_GRUPO)
  End If
  If Not IsNull(rs!Formato_Transferencia) Then
    Call sbCboAsignaDato(cboFormato, rs!FormatoN1, True, rs!Formato_Transferencia)
  End If
  If Not IsNull(rs!Formato_Transferencias_N2) Then
    Call sbCboAsignaDato(cboFormatoN2, rs!FormatoN2, True, rs!Formato_Transferencias_N2)
  End If
  
  Call sbCboAsignaDato(cboDivisa, rs!DivisaDesc, True, rs!COD_DIVISA)
  
  dtpFecha.Value = rs!fecha_envia
  txtSaldo = Format(rs!Saldo, "Standard")
  
  txtFirmaDesde.Text = IIf(IsNull(rs!firmas_desde), 0, Format(rs!firmas_desde, "Standard"))
  txtFirmaHasta.Text = IIf(IsNull(rs!firmas_hasta), 0, Format(rs!firmas_hasta, "Standard"))
  
  
  
  chkCuentaBancariaPuente.Value = rs!puente
  
  chkSINPE_CtaInterna.Value = rs!SINPE_INTERNA
  txtSINPE_Codigo.Text = rs!SINPE_EMPRESA
  
  txtCodigoCliente.Text = rs!Codigo_Cliente & ""

  'Conciliación
  txtCon_ComisionSINPECta.Text = rs!COD_CUENTA_CON & ""
  txtCon_ComisionSINPECtaDesc.Text = rs!COD_CUENTA_CON_DESC & ""
  
  txtCon_ComisionSINPEMnt.Text = Format(rs!CONCILIA_AR_COMISION, "Standard")
  
  Call sbCboAsignaDato(cboCon_Unidad, rs!Unidad_Desc, True, rs!Unidad)
  Call sbCboAsignaDato(cboCon_Centro, rs!Centro_Desc, True, rs!Centro)
  Call sbCboAsignaDato(cboCon_Centro_Comision, rs!CENTRO_COM_DESC, True, rs!Centro_COM)
  Call sbCboAsignaDato(cboCon_Concepto, rs!Concepto_Desc, True, rs!CONCEPTO)
  
  
  chkUtilizaPlan.Value = rs!UTILIZA_PLAN
  Call chkUtilizaPlan_Click
Else

  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault


Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Verifica que exista ningun otro Banco con la misma cuenta Bancaria
strSQL = "select isnull(count(*),0) as Existe from Tes_Bancos" _
       & " where id_banco not in(" & vCodigo & ") and cta = '" & Trim(txtCuentaBancaria) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
   vMensaje = vMensaje & vbCrLf & " - Existe ya un Banco registrado con la Misma Cuenta Bancaria..."
End If
rs.Close

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, txtCuentaContable, 0)) Then
   vMensaje = vMensaje & vbCrLf & " - No se especificó una cuenta contable válida..."
End If

If txtNombre.Text = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre de la Cuenta Bancaria no es válida ..."
'If txtLugarEmision.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se especificó el lugar de Emision ..."
If txtCuentaBancaria.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se indicó el número de cuenta bancaria ..."


If chKFormatoEspecial.Value = 1 Then
  If txtArchivoEspecial = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Archivo especial no puede estar en blanco ..."
  If txtChequeEspecialFirma = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Cheque especial con firma no puede estar en blanco ..."
  If txtChequeEspecialNoFirma = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre Nombre del Cheque especial sin firma  no puede estar en blanco ..."
End If

If chkSINPE_CtaInterna.Value = vbChecked Then
  If txtSINPE_Codigo.Text = "" Then vMensaje = vMensaje & vbCrLf & " - No se especificó el Código de Empresa SINPE.."
  If Not IsNumeric(txtSINPE_Codigo.Text) Or Len(txtSINPE_Codigo.Text) > 3 Then vMensaje = vMensaje & vbCrLf & " - El Código de Empresa SINPE no es válido.."
 
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


If vEdita Then
  strSQL = "update Tes_Bancos set Descripcion = '" & Trim(txtNombre.Text) & "',Puente = " & chkCuentaBancariaPuente.Value _
         & ",estado = '" & Mid(cboEstado.Text, 1, 1) & "', Utiliza_Plan = " & chkUtilizaPlan.Value _
         & " ,formato_transferencia = '" & cboFormato.ItemData(cboFormato.ListIndex) _
         & "',formato_transferencias_N2 = '" & cboFormatoN2.ItemData(cboFormatoN2.ListIndex) _
         & "',cta = '" & txtCuentaBancaria & "',CtaConta = '" & fxgCntCuentaFormato(False, txtCuentaContable) _
         & "',Desc_Corta = '" & txtDescCorta & "',cta_regional = " & chkRegional.Value & ", monitoreo = " & chkMonitoreo.Value _
         & ", cod_grupo = '" & cboGrupo.ItemData(cboGrupo.ListIndex) & "', Archivo_Especial_CK = '" & txtArchivoEspecial.Text _
         & "', archivo_cheques_firmas = '" & txtChequeEspecialFirma & "',archivo_cheques_sin_firmas = '" & txtChequeEspecialNoFirma & "'" _
         & " ,utiliza_formato_especial = " & chKFormatoEspecial.Value & ",Lugar_Emision = '" & Trim(txtLugarEmision.Text) _
         & "',SUPERVISION =" & chkSupervisa.Value & " ,SUPERVISION_DIAS = " & txtDias.Text _
         & ",SINPE_INTERNA = " & chkSINPE_CtaInterna.Value & ",SINPE_EMPRESA = '" & Trim(txtSINPE_Codigo.Text) _
         & "', CODIGO_CLIENTE = '" & Trim(txtCodigoCliente.Text) & "', cod_divisa = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'" _
         & ", UTILIZA_AUTOGESTION = " & chkAutoGestion.Value _
         & ", CONCILIA_AR_COMISION = " & CCur(txtCon_ComisionSINPEMnt.Text) _
         & ", CONCILIA_AR_COMISION_CTA = '" & fxgCntCuentaFormato(False, txtCon_ComisionSINPECta.Text, 0) _
         & "', CONCILIA_AR_UNIDAD = '" & cboCon_Unidad.ItemData(cboCon_Unidad.ListIndex) _
         & "', CONCILIA_AR_CENTRO = '" & cboCon_Centro.ItemData(cboCon_Centro.ListIndex) _
         & "', CONCILIA_AR_CENTRO_COM = '" & cboCon_Centro_Comision.ItemData(cboCon_Centro_Comision.ListIndex) _
         & "', CONCILIA_AR_CONCEPTO = '" & cboCon_Concepto.ItemData(cboCon_Concepto.ListIndex) _
         & "' Where Id_Banco = " & vCodigo
         
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Cuenta Bancaria: " & vCodigo)

Else
   strSQL = "insert Tes_Bancos(descripcion,estado,Utiliza_Plan,formato_transferencia,formato_transferencias_N2,Cta,CtaConta,Desc_Corta" _
          & ",firmas_desde,firmas_hasta,saldo,fecha_envia,cta_regional,cod_grupo,monitoreo,ARCHIVO_ESPECIAL_CK,puente" _
          & ",archivo_cheques_firmas,archivo_cheques_sin_firmas,utiliza_formato_especial,lugar_emision" _
          & ",SUPERVISION,SUPERVISION_DIAS,SINPE_INTERNA,SINPE_EMPRESA, CODIGO_CLIENTE, cod_divisa, UTILIZA_AUTOGESTION" _
          & ",CONCILIA_AR_COMISION, CONCILIA_AR_COMISION_CTA, CONCILIA_AR_UNIDAD, CONCILIA_AR_CENTRO, CONCILIA_AR_CENTRO_COM, CONCILIA_AR_CONCEPTO)" _
          & " values('" & Trim(txtNombre) & "','" & Mid(cboEstado.Text, 1, 1) & "'," & chkUtilizaPlan.Value _
          & ",'" & cboFormato.ItemData(cboFormato.ListIndex) & "','" & cboFormatoN2.ItemData(cboFormatoN2.ListIndex) _
          & "','" & txtCuentaBancaria & "','" & fxgCntCuentaFormato(False, txtCuentaContable) & "','" & txtDescCorta _
          & "',0,0,0,dbo.MyGetdate()," & chkRegional.Value & ",'" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'," & chkMonitoreo.Value _
          & ",'" & txtArchivoEspecial.Text & "'," & chkCuentaBancariaPuente.Value & "" _
          & ",'" & txtChequeEspecialFirma & "','" & txtChequeEspecialNoFirma & "'," & chKFormatoEspecial.Value & ",'" & Trim(txtLugarEmision.Text) _
          & "'," & chkSupervisa.Value & "," & txtDias.Text & "," & chkSINPE_CtaInterna.Value _
          & ",'" & Trim(txtSINPE_Codigo.Text) & "','" & Trim(txtCodigoCliente.Text) _
          & "','" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'," & chkAutoGestion.Value _
          & "," & CCur(txtCon_ComisionSINPEMnt.Text) & ",'" & fxgCntCuentaFormato(False, txtCon_ComisionSINPECta.Text, 0) _
          & "','" & cboCon_Unidad.ItemData(cboCon_Unidad.ListIndex) _
          & "','" & cboCon_Centro.ItemData(cboCon_Centro.ListIndex) _
          & "','" & cboCon_Centro_Comision.ItemData(cboCon_Centro_Comision.ListIndex) _
          & "','" & cboCon_Concepto.ItemData(cboCon_Concepto.ListIndex) & "')"
   
   Call ConectionExecute(strSQL)
    
   txtCodigo.Enabled = True
 
   strSQL = "select isnull(max(id_Banco),0) as ultimo from Tes_Bancos"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
 
   Call Bitacora("Registra", "Cuenta Bancaria: " & vCodigo)
 
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Tes_Bancos where id_banco = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Banco Cod: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlbCkFirmas_ButtonClick(ByVal Button As MSComctlLib.Button)
Call sbCargaArchivo(2)
End Sub

Private Sub tlbCkNoFirmas_ButtonClick(ByVal Button As MSComctlLib.Button)
Call sbCargaArchivo(3)
End Sub

Private Sub tlbDocumento_ButtonClick(ByVal Button As MSComctlLib.Button)
Call sbCargaArchivo(1)
End Sub

Private Sub txtArchivoEspecial_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtChequeEspecialFirma.SetFocus
End Sub

Private Sub txtChequeEspecialFirma_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtChequeEspecialNoFirma.SetFocus
End Sub


Private Sub txtChequeEspecialNoFirma_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
    tcMain.Item(0).Selected = True
End If

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Cuenta Id"
  gBusquedas.Col2Name = "Banco Desc"
  gBusquedas.Col3Name = "Cuenta"
  
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "id_banco"
  gBusquedas.Orden = "id_banco"
  gBusquedas.Consulta = "select id_banco,descripcion,cta from Tes_Bancos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub



Private Sub txtCodigoCliente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSINPE_Codigo.SetFocus
End Sub


Private Sub txtCon_ComisionSINPECta_GotFocus()
On Error GoTo vError
txtCon_ComisionSINPECta.Text = fxgCntCuentaFormato(False, txtCon_ComisionSINPECta.Text, 0)
vError:
End Sub

Private Sub txtCon_ComisionSINPECta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCon_ComisionSINPECtaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   Call sbgCntCuentaConsulta("C")
   txtCon_ComisionSINPECta.Text = gBusquedas.Resultado
   txtCon_ComisionSINPECtaDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCon_ComisionSINPECta_LostFocus()
On Error GoTo vError

txtCon_ComisionSINPECtaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCon_ComisionSINPECta.Text))
txtCon_ComisionSINPECta.Text = fxgCntCuentaFormato(True, txtCon_ComisionSINPECta.Text)

vError:
End Sub



Private Sub txtCon_ComisionSINPECtaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Or KeyCode = vbKeyReturn Then cboCon_Unidad.SetFocus
End Sub

Private Sub txtCon_ComisionSINPEMnt_GotFocus()
On Error GoTo vError
 txtCon_ComisionSINPEMnt.Text = CCur(txtCon_ComisionSINPEMnt.Text)
vError:
End Sub

Private Sub txtCon_ComisionSINPEMnt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Or KeyCode = vbKeyReturn Then txtCon_ComisionSINPECta.SetFocus
End Sub

Private Sub txtCon_ComisionSINPEMnt_LostFocus()
On Error GoTo vError
 txtCon_ComisionSINPEMnt.Text = Format(CCur(txtCon_ComisionSINPEMnt.Text), "Standard")
vError:
End Sub

Private Sub txtCtaContaDesc_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   Call sbgCntCuentaConsulta("D")
   txtCuentaContable = gBusquedas.Resultado
   txtCtaContaDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCuentaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescCorta.SetFocus
End Sub

Private Sub txtCuentaContable_GotFocus()
On Error GoTo vError
txtCuentaContable = fxgCntCuentaFormato(False, txtCuentaContable, 0)
vError:
End Sub

Private Sub txtCuentaContable_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaContaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   Call sbgCntCuentaConsulta("C")
   txtCuentaContable = gBusquedas.Resultado
   txtCtaContaDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCuentaContable_LostFocus()
On Error GoTo vError
txtCtaContaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaContable.Text))
txtCuentaContable.Text = fxgCntCuentaFormato(True, txtCuentaContable.Text)

vError:
End Sub

Private Sub txtDescCorta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtLugarEmision.SetFocus
End Sub

Private Sub txtFirmaDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFirmaHasta.SetFocus
End Sub

Private Sub txtFirmaHasta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdActualizar.SetFocus
vError:
End Sub


Private Sub txtLugarEmision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboFormato.SetFocus
End Sub

Private Sub txtSaldo_GotFocus()
On Error GoTo vError
 txtSaldo = CCur(txtSaldo)
vError:
End Sub

Private Sub txtSaldo_LostFocus()
On Error GoTo vError
 txtSaldo = Format(CCur(txtSaldo), "Standard")
vError:
End Sub


Private Sub txtFirmaDesde_GotFocus()
On Error GoTo vError
 txtFirmaDesde = CCur(txtFirmaDesde)
vError:
End Sub

Private Sub txtFirmaDesde_LostFocus()
On Error GoTo vError
 txtFirmaDesde = Format(CCur(txtFirmaDesde), "Standard")
vError:
End Sub


Private Sub txtFirmaHasta_GotFocus()
On Error GoTo vError
 txtFirmaHasta = CCur(txtFirmaHasta)
vError:
End Sub

Private Sub txtFirmaHasta_LostFocus()
On Error GoTo vError
 txtFirmaHasta = Format(CCur(txtFirmaHasta), "Standard")
vError:
End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboGrupo.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select id_Banco,descripcion from Tes_Bancos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub sbCargaArchivo(vOpcion As Integer)
'selecciona el archivo segun la opcion seleccionada
frmContenedor.CD.FileName = "*.rpt"
frmContenedor.CD.ShowOpen


If frmContenedor.CD.FileName <> "" And frmContenedor.CD.FileName <> "*.rpt" Then
    Select Case vOpcion
        Case 1
          txtArchivoEspecial = Dir(frmContenedor.CD.FileName)
        Case 2
          txtChequeEspecialFirma = Dir(frmContenedor.CD.FileName)
        Case 3
          txtChequeEspecialNoFirma = Dir(frmContenedor.CD.FileName)
   End Select
Else
   MsgBox "No selecciono ningun archivo"
End If
frmContenedor.CD.FileName = ""


End Sub


Private Sub txtSINPE_Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaContable.SetFocus
End Sub
