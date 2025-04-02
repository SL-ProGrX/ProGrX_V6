VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmSIF_Tipo_Documento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Documentos"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9660
   Icon            =   "frmSIF_Tipo_Documento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6972
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9372
      _Version        =   1441793
      _ExtentX        =   16531
      _ExtentY        =   12298
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
      Item(0).Caption =   "Detalle"
      Item(0).ImageIndex=   0
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "Label3(0)"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "Label3(2)"
      Item(0).Control(3)=   "Label3(3)"
      Item(0).Control(4)=   "chkActivo"
      Item(0).Control(5)=   "cboComprobante"
      Item(0).Control(6)=   "txtConsecutivo"
      Item(0).Control(7)=   "txtArchivoEspecial"
      Item(0).Control(8)=   "btnImagenes"
      Item(0).Control(9)=   "cboFormato"
      Item(0).Control(10)=   "chkCierreEspecial"
      Item(0).Control(11)=   "chkPermiteReversion"
      Item(0).Control(12)=   "chkRegistraImpuesto"
      Item(0).Control(13)=   "txtDiasReversion"
      Item(0).Control(14)=   "txtImpuesto"
      Item(0).Control(15)=   "Label3(7)"
      Item(0).Control(16)=   "Label3(8)"
      Item(0).Control(17)=   "gbAfectacionContable"
      Item(0).Control(18)=   "GroupBox3"
      Item(1).Caption =   "Asignación de Conceptos"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6372
         Left            =   -70000
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   9372
         _Version        =   1441793
         _ExtentX        =   16531
         _ExtentY        =   11239
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbAfectacionContable 
         Height          =   1812
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   9012
         _Version        =   1441793
         _ExtentX        =   15896
         _ExtentY        =   3196
         _StockProps     =   79
         Caption         =   "Efecto Contable"
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
         Begin XtremeSuiteControls.ComboBox cboTipoAsiento 
            Height          =   312
            Left            =   1440
            TabIndex        =   24
            Top             =   360
            Width           =   2892
            _Version        =   1441793
            _ExtentX        =   5106
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMov 
            Height          =   312
            Left            =   1440
            TabIndex        =   25
            Top             =   840
            Width           =   2892
            _Version        =   1441793
            _ExtentX        =   5106
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.CheckBox chkAsTransac 
            Height          =   252
            Left            =   4680
            TabIndex        =   28
            Top             =   360
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Genera Asientos x Transacción"
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.CheckBox chkAsFormato 
            Height          =   252
            Left            =   4680
            TabIndex        =   29
            Top             =   720
            Width           =   5892
            _Version        =   1441793
            _ExtentX        =   10393
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Establece Formato al Número de Asiento"
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.CheckBox chkAsIDModulo 
            Height          =   252
            Left            =   4680
            TabIndex        =   30
            Top             =   1440
            Width           =   4212
            _Version        =   1441793
            _ExtentX        =   7429
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Incluir el Id del Modulo en el Formato"
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
            Appearance      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtMascara 
            Height          =   312
            Left            =   5040
            TabIndex        =   31
            Top             =   1080
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   252
            Index           =   9
            Left            =   6600
            TabIndex        =   32
            Top             =   1080
            Width           =   3252
            _Version        =   1441793
            _ExtentX        =   5736
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Mascara del Formato"
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
            Height          =   372
            Index           =   6
            Left            =   0
            TabIndex        =   27
            Top             =   840
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Afectación Contable"
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
            Index           =   5
            Left            =   0
            TabIndex        =   26
            Top             =   360
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tipo Asiento"
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
      Begin XtremeSuiteControls.CheckBox chkActivo 
         Height          =   372
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Activo ?  "
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
         Transparent     =   -1  'True
         TextAlignment   =   1
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ComboBox cboComprobante 
         Height          =   312
         Left            =   1680
         TabIndex        =   11
         Top             =   840
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4048
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtConsecutivo 
         Height          =   312
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivoEspecial 
         Height          =   312
         Left            =   1680
         TabIndex        =   13
         Top             =   1560
         Width           =   6852
         _Version        =   1441793
         _ExtentX        =   12086
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
      Begin XtremeSuiteControls.PushButton btnImagenes 
         Height          =   315
         Left            =   8640
         TabIndex        =   14
         Top             =   1560
         Width           =   435
         _Version        =   1441793
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   79
         BackColor       =   -2147483643
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmSIF_Tipo_Documento.frx":0ECA
      End
      Begin XtremeSuiteControls.ComboBox cboFormato 
         Height          =   312
         Left            =   6240
         TabIndex        =   15
         Top             =   840
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4048
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
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkCierreEspecial 
         Height          =   252
         Left            =   1680
         TabIndex        =   16
         Top             =   1920
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Utilizar en Informe de Cierre Especial ?"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkPermiteReversion 
         Height          =   252
         Left            =   1680
         TabIndex        =   17
         Top             =   2280
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Permite Reversión ?"
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkRegistraImpuesto 
         Height          =   252
         Left            =   1680
         TabIndex        =   18
         Top             =   3000
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Registra Impuesto ? "
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtDiasReversion 
         Height          =   312
         Left            =   2040
         TabIndex        =   19
         Top             =   2640
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtImpuesto 
         Height          =   312
         Left            =   2040
         TabIndex        =   20
         Top             =   3360
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1452
         Left            =   240
         TabIndex        =   33
         Top             =   5640
         Width           =   9012
         _Version        =   1441793
         _ExtentX        =   15896
         _ExtentY        =   2561
         _StockProps     =   79
         Caption         =   "Cuenta Contable:"
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
         Begin XtremeSuiteControls.FlatEdit txtImpuestoCuenta 
            Height          =   312
            Left            =   1440
            TabIndex        =   38
            Top             =   840
            Width           =   2052
            _Version        =   1441793
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   34
            Top             =   480
            Width           =   5412
            _Version        =   1441793
            _ExtentX        =   9546
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
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   312
            Left            =   1440
            TabIndex        =   35
            Top             =   480
            Width           =   2052
            _Version        =   1441793
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtImpuestoCuentaDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   37
            Top             =   840
            Width           =   5412
            _Version        =   1441793
            _ExtentX        =   9546
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   252
            Index           =   11
            Left            =   0
            TabIndex        =   39
            Top             =   840
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cta. Impuesto"
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
            Index           =   10
            Left            =   0
            TabIndex        =   36
            Top             =   480
            Width           =   1572
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cuenta"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   8
         Left            =   3600
         TabIndex        =   22
         Top             =   3360
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Porcentaje de impuesto"
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
         Index           =   7
         Left            =   3600
         TabIndex        =   21
         Top             =   2640
         Width           =   3252
         _Version        =   1441793
         _ExtentX        =   5736
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Dias admisibles para reversión"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   840
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Formato Salida"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Archivo Especial"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Consecutivo"
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Comprobante"
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
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   5160
      Top             =   120
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   264
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3336
      _ExtentX        =   5874
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   9000
      TabIndex        =   1
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   330
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   6252
      _Version        =   1441793
      _ExtentX        =   11028
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   330
      Left            =   1320
      TabIndex        =   9
      Top             =   480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   582
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   4
      Left            =   -120
      TabIndex        =   10
      Top             =   480
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Documento"
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
Attribute VB_Name = "frmSIF_Tipo_Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean

Private Sub btnImagenes_Click()
With frmContenedor.CD
    .FileName = "*.rpt"
    .ShowOpen
    
    If .FileName <> "" And .FileName <> "*.rpt" Then
       txtArchivoEspecial = Dir(.FileName)
    Else
       MsgBox "No selecciono ningun archivo"
    End If
    .FileName = ""

End With
End Sub

Private Sub cboComprobante_Click()
If cboComprobante.ItemData(cboComprobante.ListIndex) = "02" Then
   btnImagenes.Enabled = True
Else
   txtArchivoEspecial.Text = Empty
   btnImagenes.Enabled = False
End If

End Sub

Private Sub chkAsFormato_Click()
If chkAsFormato.Value = vbChecked Then
   txtMascara.Enabled = True
   chkAsIDModulo.Enabled = True
Else
   txtMascara.Enabled = False
   chkAsIDModulo.Enabled = False
End If
End Sub


Private Sub chkRegistraImpuesto_Click()
If chkRegistraImpuesto.Value = vbChecked Then
   txtImpuesto.Enabled = True
Else
   txtImpuesto.Enabled = False
   txtImpuesto.Text = 0
End If
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
   
    strSQL = "select Top 1 tipo_documento from sif_documentos"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where tipo_documento > '" & txtCodigo & "' order by tipo_documento asc"
    Else
       strSQL = strSQL & " where tipo_documento < '" & txtCodigo & "' order by tipo_documento desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Tipo_Documento
      Call sbConsulta(txtCodigo.Text)
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
 vModulo = 10
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 10
 
 tcMain.Item(0).Selected = True
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1200
    .Add , , "Descripción", lsw.Width - (1500)
End With

 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpia()

tcMain.Item(0).Selected = True

vCodigo = ""
txtCodigo.Text = ""
txtDescripcion.Text = ""

cboMov.Text = "Débito"
cboComprobante.Text = "Boleta de Registro"
cboFormato.Text = "Recibo"

chkActivo.Value = vbChecked
chkCierreEspecial.Value = vbUnchecked

chkRegistraImpuesto.Value = vbChecked
txtImpuesto.Text = "0"

chkPermiteReversion.Value = vbUnchecked
txtDiasReversion.Text = "1"

chkAsFormato.Value = vbChecked
chkAsIDModulo.Value = vbChecked
chkAsTransac.Value = vbChecked

txtMascara.Text = "00000000"

txtCuenta.Text = ""
txtImpuestoCuenta.Text = ""
txtCuentaDesc.Text = ""
txtImpuestoCuentaDesc.Text = ""

End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim strSQL As String

On Error GoTo vError


If Item.Checked Then
   strSQL = "insert into sif_conceptos_documento(cod_concepto,tipo_documento,registro_fecha,registro_usuario)" _
          & " values('" & Item.Text & "','" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
   
   Item.ForeColor = vbBlue
   
Else
   strSQL = "Delete sif_conceptos_documento where cod_concepto ='" & Item.Text _
          & "' and tipo_documento = '" & vCodigo & "' "
   
   Item.ForeColor = vbRed
      
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbInicializa()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

cboMov.Clear
cboMov.AddItem "Débito"
cboMov.ItemData(cboMov.ListCount - 1) = "D"
cboMov.AddItem "Crédito"
cboMov.ItemData(cboMov.ListCount - 1) = "C"
cboMov.AddItem "Ambos"
cboMov.ItemData(cboMov.ListCount - 1) = "A"

cboComprobante.Clear
cboComprobante.AddItem "Boleta de Registro"
cboComprobante.ItemData(cboComprobante.ListCount - 1) = "00"
cboComprobante.AddItem "Recibo"
cboComprobante.ItemData(cboComprobante.ListCount - 1) = "01"
cboComprobante.AddItem "Archivo Personalizado"
cboComprobante.ItemData(cboComprobante.ListCount - 1) = "02"


cboFormato.AddItem "Recibo"
cboFormato.ItemData(cboFormato.ListCount - 1) = "01"
cboFormato.AddItem "Boleta de Registro"
cboFormato.ItemData(cboFormato.ListCount - 1) = "02"


strSQL = "select rtrim(tipo_asiento) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
       & " from Cntx_tipos_asientos where activo = 1 and cod_Contabilidad = " & GLOBALES.gEnlace
Call sbCbo_Llena_New(cboTipoAsiento, strSQL, False, True)

Call sbLimpia

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 And vCodigo <> "" Then
    Call sbLsw_Load
Else
    lsw.ListItems.Clear
End If

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpia
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpia
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select Tipo_documento,descripcion from sif_documentos"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtDescripcion.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from vSys_Documentos_Tipos where Tipo_documento = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  tcMain.Item(0).Selected = True

  vCodigo = rs!Tipo_Documento
  txtCodigo = rs!Tipo_Documento

  txtDescripcion = rs!Descripcion & ""
  txtConsecutivo.Text = CStr(IIf(IsNull(rs!Consecutivo), 0, rs!Consecutivo))
  
  chkActivo.Value = rs!Activo
  
  If rs!Tipo_Comprobante = "02" Then
     txtArchivoEspecial.Text = rs!Archivo_Per & ""
     btnImagenes.Enabled = True
  Else
     txtArchivoEspecial.Text = ""
     btnImagenes.Enabled = False
  End If

  
  'Carga Combos
  Call sbCboAsignaDato(cboMov, rs!Tipo_Movimiento_DESC, True, rs!Tipo_Movimiento)
  Call sbCboAsignaDato(cboTipoAsiento, rs!Tipo_Asiento_DESC, True, rs!Tipo_Asiento)
  Call sbCboAsignaDato(cboComprobante, rs!Tipo_Comprobante_DESC, True, rs!Tipo_Comprobante)
  Call sbCboAsignaDato(cboFormato, rs!FORMATO_SALIDA_DESC, True, rs!FORMATO_SALIDA_ID)
 
 
  If rs!asiento_modulo > 0 Then
     chkAsFormato.Value = vbUnchecked
  Else
     chkAsFormato.Value = vbUnchecked
  End If
  
  chkPermiteReversion.Value = rs!Permite_Reversion
  txtDiasReversion.Text = CStr(IIf(IsNull(rs!REVERSION_DIAS_AUTORIZADOS), 1, rs!REVERSION_DIAS_AUTORIZADOS))
  
  chkCierreEspecial.Value = IIf(IsNull(rs!APLICA_CIERRE_ESPECIAL), 0, rs!APLICA_CIERRE_ESPECIAL)
  
  
  chkAsTransac.Value = rs!Asiento_Transaccion
  chkAsIDModulo.Value = rs!asiento_modulo
  txtImpuesto.Text = Format(rs!impuesto_porcentaje, "Standard")
  txtMascara.Text = Trim(IIf(IsNull(rs!Asiento_Mascara), "", Trim(rs!Asiento_Mascara)))
  
  txtCuenta.Text = rs!Cuenta_Mask
  txtCuentaDesc.Text = rs!Cuenta_Desc
  
  txtImpuestoCuenta.Text = rs!Imp_Cuenta_Mask
  txtImpuestoCuentaDesc = rs!Imp_Cuenta_Desc
  
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
  Call sbLimpia
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
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Documento no es válido ..."
If cboComprobante.ItemData(cboComprobante.ListIndex) = "02" Then
   If txtArchivoEspecial = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Archivo especial no es válido ..."
End If

If Not IsNumeric(txtDiasReversion.Text) Then
    vMensaje = vMensaje & " - Los días autorizados para reversión de un documento no son válidos.."
End If

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, txtCuenta.Text)) Then
    vMensaje = vMensaje & " - La cuenta para cierre por omisión de los asientos de este tipo de documento no es válida.."
End If

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, txtImpuestoCuenta.Text)) Then
    vMensaje = vMensaje & " - La cuenta para Impuestos de los asientos de este tipo de documento no es válida.."
End If


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipoMov As String, vFormato As String
Dim vTipoAsiento As String, vComprobante As String

On Error GoTo vError

vTipoMov = cboMov.ItemData(cboMov.ListIndex)
vFormato = cboFormato.ItemData(cboFormato.ListIndex)
vTipoAsiento = cboTipoAsiento.ItemData(cboTipoAsiento.ListIndex)
vComprobante = cboComprobante.ItemData(cboComprobante.ListIndex)

If vEdita Then
  strSQL = "update sif_documentos set descripcion = '" & Trim(txtDescripcion.Text) & "'" _
         & ", tipo_movimiento = '" & vTipoMov & "',activo = " & chkActivo.Value _
         & ", Tipo_asiento = '" & vTipoAsiento & "', cod_cuenta = '" & fxgCntCuentaFormato(False, txtCuenta) & "'" _
         & ", asiento_transaccion = " & chkAsTransac.Value & ",asiento_mascara = '" & txtMascara.Text & "',asiento_modulo = " & chkAsIDModulo.Value & "" _
         & ", formato_salida = '" & vFormato & "',impuesto_registra = " & chkRegistraImpuesto.Value & "" _
         & ", Impuesto_porcentaje = " & CCur(IIf(txtImpuesto.Text = "", 0, txtImpuesto.Text)) _
         & ", Impuesto_cod_cuenta = '" & fxgCntCuentaFormato(False, txtImpuestoCuenta) & "'" _
         & ", tipo_comprobante = '" & vComprobante & "',archivo_per = '" & txtArchivoEspecial.Text & "' " _
         & ",Permite_Reversion = " & chkPermiteReversion.Value & ",APLICA_CIERRE_ESPECIAL = " & chkCierreEspecial.Value _
         & ",REVERSION_DIAS_AUTORIZADOS = " & txtDiasReversion.Text _
         & " where Tipo_documento = '" & vCodigo & "'"
         
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Tipo de Documento : " & vCodigo)

Else
  vCodigo = txtCodigo

   strSQL = "insert into sif_documentos(TIPO_DOCUMENTO,descripcion,consecutivo,tipo_comprobante,tipo_movimiento,tipo_asiento" _
          & ",cod_cuenta,activo,asiento_transaccion,asiento_mascara,asiento_modulo,formato_salida,impuesto_registra,Impuesto_porcentaje" _
          & " ,impuesto_cod_cuenta,registro_fecha,registro_usuario,archivo_per,Permite_Reversion,APLICA_CIERRE_ESPECIAL,REVERSION_DIAS_AUTORIZADOS)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion) & "'," & txtConsecutivo & ",'" & vComprobante & "','" & vTipoMov & "'," _
          & "'" & vTipoAsiento & "','" & fxgCntCuentaFormato(False, txtCuenta) & "'," & chkActivo.Value & "," & chkAsTransac.Value & "," _
          & "'" & txtMascara.Text & "'," & chkAsIDModulo.Value & ",'" & vFormato & "'," & chkRegistraImpuesto.Value & "," & CCur(txtImpuesto.Text) & "," _
          & "'" & fxgCntCuentaFormato(False, txtImpuestoCuenta) & "',dbo.MyGetDate(),'" & glogon.Usuario & "','" _
          & txtArchivoEspecial.Text & "'," & chkPermiteReversion.Value & "," & chkCierreEspecial.Value & "," & txtDiasReversion.Text & ")"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Tipo de Documento: " & vCodigo)

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
  strSQL = "delete sif_documentos where Tipo_documento = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Tipo de Documento : " & vCodigo)
  Call sbLimpia
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Tipo_documento"
  gBusquedas.Orden = "Tipo_documento"
  gBusquedas.Consulta = "select Tipo_documento,descripcion from sif_documentos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
'txtDescripcion = fxgConCodigos("D", txtCodigo, "Tipo de Documentos")
End Sub


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta = gCuenta
    txtCuentaDesc = ""
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuenta.SetFocus

End Sub

Private Sub txtCuenta_LostFocus()
 txtCuenta.Text = fxgCntCuentaFormato(False, txtCuenta)
 txtCuentaDesc.Text = fxgCntCuentaDesc(txtCuenta)
 txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta)
End Sub

Private Sub txtImpuestoCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtImpuestoCuenta.Text = gCuenta
  txtImpuestoCuentaDesc.Text = ""
End If
End Sub

Private Sub txtImpuestoCuenta_LostFocus()
 txtImpuestoCuenta.Text = fxgCntCuentaFormato(False, txtImpuestoCuenta)
 txtImpuestoCuentaDesc.Text = fxgCntCuentaDesc(txtImpuestoCuenta)
 txtImpuestoCuenta.Text = fxgCntCuentaFormato(True, txtImpuestoCuenta)
End Sub

Private Sub txtImpuesto_Change()
If Not IsNumeric(txtImpuesto.Text) Then
  txtImpuesto.Text = "0"
End If
End Sub

Private Sub txtImpuesto_LostFocus()
If txtImpuesto.Text > 100 Then
   MsgBox "El valor supera el 100%"
   txtImpuesto.SetFocus
End If
End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboComprobante.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select Tipo_documento,descripcion from sif_documentos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub



Private Sub sbLsw_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

vPaso = True

lsw.ListItems.Clear

'Conceptos Relacionados
strSQL = "Select C.cod_concepto,C.descripcion,X.cod_Concepto as Asignado" _
        & " from sif_conceptos C left join sif_conceptos_documento X" _
        & " on C.cod_concepto = X.cod_concepto and X.Tipo_documento = '" & vCodigo & "'" _
        & " order by X.cod_Concepto desc,C.cod_concepto asc"

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!cod_concepto)
      itmX.SubItems(1) = rs!Descripcion
      
      If Not IsNull(rs!Asignado) Then
          itmX.Checked = True
      End If
  rs.MoveNext
Loop
rs.Close
 
vPaso = False
 
Me.MousePointer = vbDefault

End Sub

