VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAF_BeneficioAsg 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Beneficios a Personas"
   ClientHeight    =   7005
   ClientLeft      =   3975
   ClientTop       =   2565
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8880
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4935
      Left            =   0
      TabIndex        =   9
      Top             =   1800
      Width           =   9375
      _Version        =   1572864
      _ExtentX        =   16536
      _ExtentY        =   8705
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
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Consulta"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lswAsignados"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gbMain"
      Item(2).Caption =   "Pago"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "frmMonetario"
      Item(2).Control(1)=   "frmProducto"
      Begin XtremeSuiteControls.ListView lswAsignados 
         Height          =   4575
         Left            =   -70000
         TabIndex        =   48
         Top             =   360
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1572864
         _ExtentX        =   16325
         _ExtentY        =   8070
         _StockProps     =   77
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox frmMonetario 
         Height          =   3975
         Left            =   -69760
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   8535
         _Version        =   1572864
         _ExtentX        =   15049
         _ExtentY        =   7006
         _StockProps     =   79
         Caption         =   "Registro para el desembolso del Beneficio o Ayuda:"
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
         Begin XtremeSuiteControls.ListView lswPago 
            Height          =   1212
            Left            =   0
            TabIndex        =   11
            Top             =   1080
            Width           =   8532
            _Version        =   1572864
            _ExtentX        =   15049
            _ExtentY        =   2138
            _StockProps     =   77
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.ComboBox cboBanco 
            Height          =   315
            Left            =   720
            TabIndex        =   12
            Top             =   3120
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
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
         Begin XtremeSuiteControls.ComboBox cboCuenta 
            Height          =   315
            Left            =   720
            TabIndex        =   13
            Top             =   3480
            Width           =   4335
            _Version        =   1572864
            _ExtentX        =   7646
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
         Begin XtremeSuiteControls.ComboBox cboEmitir 
            Height          =   315
            Left            =   6480
            TabIndex        =   14
            Top             =   3120
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   312
            Left            =   6480
            TabIndex        =   15
            Top             =   240
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMontoAsg 
            Height          =   315
            Left            =   6480
            TabIndex        =   16
            Top             =   3480
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3619
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDisponible 
            Height          =   312
            Left            =   6480
            TabIndex        =   17
            Top             =   600
            Width           =   2052
            _Version        =   1572864
            _ExtentX        =   3619
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label lblPagoCaso 
            Height          =   312
            Left            =   720
            TabIndex        =   49
            Top             =   2640
            Width           =   7812
            _Version        =   1572864
            _ExtentX        =   13779
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
         End
         Begin VB.Label Label1 
            Caption         =   "Disponible"
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
            Index           =   17
            Left            =   5280
            TabIndex        =   24
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   16
            Left            =   0
            TabIndex        =   23
            Top             =   3480
            Width           =   975
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto Girar"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   15
            Left            =   5280
            TabIndex        =   22
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Emitir"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   14
            Left            =   5280
            TabIndex        =   21
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   13
            Left            =   0
            TabIndex        =   20
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Registro"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Index           =   9
            Left            =   0
            TabIndex        =   19
            Top             =   2640
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Monto"
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
            Index           =   7
            Left            =   5280
            TabIndex        =   18
            Top             =   240
            Width           =   1092
         End
      End
      Begin XtremeSuiteControls.GroupBox frmProducto 
         Height          =   3972
         Left            =   -69760
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1572864
         _ExtentX        =   15049
         _ExtentY        =   7006
         _StockProps     =   79
         Caption         =   "Productos Asignados:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   2772
            Left            =   0
            TabIndex        =   26
            Top             =   1080
            Width           =   8412
            _Version        =   1572864
            _ExtentX        =   14838
            _ExtentY        =   4890
            _StockProps     =   77
            View            =   3
            FullRowSelect   =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtProducto 
            Height          =   312
            Left            =   0
            TabIndex        =   27
            Top             =   720
            Width           =   5652
            _Version        =   1572864
            _ExtentX        =   9970
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtProdCantidad 
            Height          =   312
            Left            =   5640
            TabIndex        =   28
            Top             =   720
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtProdCosto 
            Height          =   312
            Left            =   6600
            TabIndex        =   29
            Top             =   720
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   1
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Producto"
            ForeColor       =   &H00FFFFFF&
            Height          =   312
            Index           =   10
            Left            =   0
            TabIndex        =   32
            Top             =   480
            Width           =   5652
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cantidad"
            ForeColor       =   &H00FFFFFF&
            Height          =   312
            Index           =   11
            Left            =   5640
            TabIndex        =   31
            Top             =   480
            Width           =   972
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Costos Unidad"
            ForeColor       =   &H00FFFFFF&
            Height          =   312
            Index           =   12
            Left            =   6600
            TabIndex        =   30
            Top             =   480
            Width           =   1812
         End
      End
      Begin XtremeSuiteControls.GroupBox gbMain 
         Height          =   3975
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   8655
         _Version        =   1572864
         _ExtentX        =   15261
         _ExtentY        =   7006
         _StockProps     =   79
         Caption         =   "Datos del Beneficio o Ayuda: "
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
         Begin XtremeSuiteControls.ComboBox cbo 
            Height          =   312
            Left            =   1560
            TabIndex        =   34
            Top             =   720
            Width           =   3492
            _Version        =   1572864
            _ExtentX        =   6165
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
         Begin XtremeSuiteControls.ComboBox cboTipoBeneficio 
            Height          =   312
            Left            =   5040
            TabIndex        =   35
            Top             =   720
            Width           =   3252
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
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   312
            Left            =   1560
            TabIndex        =   36
            Top             =   2520
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4471
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
         Begin XtremeSuiteControls.FlatEdit txtNombreFallecido 
            Height          =   312
            Left            =   3240
            TabIndex        =   37
            Top             =   1080
            Width           =   5052
            _Version        =   1572864
            _ExtentX        =   8911
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCedulaFallecido 
            Height          =   312
            Left            =   1560
            TabIndex        =   38
            Top             =   1080
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   550
            _StockProps     =   77
            ForeColor       =   0
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   912
            Left            =   1560
            TabIndex        =   39
            Top             =   1440
            Width           =   6732
            _Version        =   1572864
            _ExtentX        =   11874
            _ExtentY        =   1609
            _StockProps     =   77
            ForeColor       =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label lblAutFecha 
            Height          =   312
            Left            =   1560
            TabIndex        =   51
            Top             =   3600
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
            Alignment       =   2
         End
         Begin XtremeSuiteControls.Label lblAutUser 
            Height          =   312
            Left            =   1560
            TabIndex        =   50
            Top             =   3240
            Width           =   2532
            _Version        =   1572864
            _ExtentX        =   4466
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "..."
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
            Alignment       =   2
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   47
            Top             =   3600
            Width           =   2295
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Usuario"
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   46
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Autorización:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Index           =   3
            Left            =   3000
            TabIndex        =   45
            Top             =   3000
            Width           =   1092
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Estado"
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   44
            Top             =   2520
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Notas"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   43
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Fallecido"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   42
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Beneficio"
            Height          =   252
            Index           =   8
            Left            =   5160
            TabIndex        =   41
            Top             =   480
            Width           =   852
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo"
            Height          =   252
            Index           =   21
            Left            =   1560
            TabIndex        =   40
            Top             =   480
            Width           =   852
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   1
      Top             =   6756
      Width           =   9288
      _ExtentX        =   16378
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3598
            MinWidth        =   3598
            Object.ToolTipText     =   "Usuario que Registra"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecaha y Hora"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Estado"
         EndProperty
      EndProperty
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
   Begin XtremeSuiteControls.PushButton btnNuevo 
      Height          =   325
      Left            =   6240
      TabIndex        =   2
      Top             =   1340
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   573
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmAF_BeneficioAsg.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnGuardar 
      Height          =   325
      Left            =   7560
      TabIndex        =   3
      Top             =   1340
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   573
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_BeneficioAsg.frx":0720
   End
   Begin XtremeSuiteControls.PushButton btnBoleta 
      Height          =   325
      Left            =   8160
      TabIndex        =   4
      Top             =   1340
      Width           =   612
      _Version        =   1572864
      _ExtentX        =   1080
      _ExtentY        =   573
      _StockProps     =   79
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmAF_BeneficioAsg.frx":0E51
   End
   Begin XtremeSuiteControls.FlatEdit txtBeneficioId 
      Height          =   372
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   5052
      _Version        =   1572864
      _ExtentX        =   8911
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2040
      TabIndex        =   7
      Top             =   480
      Width           =   1692
      _Version        =   1572864
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación:"
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
      Height          =   252
      Index           =   2
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   11892
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   420
      Left            =   0
      TabIndex        =   8
      Top             =   1305
      Width           =   9375
      _Version        =   1572864
      _ExtentX        =   16536
      _ExtentY        =   741
      _StockProps     =   14
      Caption         =   "Beneficio activo: "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmAF_BeneficioAsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strConsulta As String, vCedula As String, vTipo As String
Dim iBeneficiario As Integer, curTotalDif As Currency
Dim curDiferencia As Currency, curMonto As Currency, vProducto As String
Dim curDisponible As Currency, curTotal As Currency, i As Integer, vDescripcion As String
Dim bConsulta As Boolean, bAplicaParcial As Boolean
Dim cMontoBene As Currency, cMontoPagado As Currency, cMontoGrupo As Currency
Dim cMontoRealGrupo As Currency, iGrupo As Integer, bAsignado As Boolean
Dim iCantidaGrupo As Integer, cMontoAsignado As Currency, bMostroMensaje As Boolean
Dim cMontoPagar As Currency, bNuevo As Boolean
Dim bNotieneBancos As Boolean, vPaso As Boolean

Public Sub sbConsultaX(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

strSQL = "Select O.*,B.Descripcion from afi_bene_otorga O  inner join afi_beneficios B " _
       & " on O.cod_beneficio = B.cod_beneficio where O.cedula = '" & vCedula & "'"

Call OpenRecordSet(rs, strSQL)

lswAsignados.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswAsignados.ListItems.Add(, , CStr(rs!consec))

   itmX.Tag = rs!Cod_Beneficio
   itmX.SubItems(1) = rs!Descripcion
   itmX.SubItems(2) = Format(rs!MONTO, "Standard")

  Select Case rs!Estado
   Case "A"
    itmX.SubItems(3) = "Aprobado"
   Case "P"
    itmX.SubItems(3) = "Pendiente"
   Case "R"
    itmX.SubItems(3) = "Rechazado"
   Case "E"
    itmX.SubItems(3) = "Ejecutado"
   Case "S"
    itmX.SubItems(3) = "Solicitado"
  End Select

 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnBoleta_Click()
    Call sbImprimeBoleta
End Sub

Private Sub btnGuardar_Click()
   If Trim(txtCedula) <> "" Then
     Call sbGuardar
   End If
End Sub

Private Sub btnNuevo_Click()
    bNuevo = True
    Call sbInicializa
    
    tcMain.Item(1).Selected = True
    cboTipoBeneficio.Enabled = True
    cboTipoBeneficio.SetFocus
    
    cboEstado.Clear
    cboEstado.AddItem "SOLICITADO"
    cboEstado.AddItem "PENDIENTE"
    cboEstado.Text = "SOLICITADO"
    
    btnGuardar.Enabled = True
    
    lswAsignados.Enabled = False
    
    
    
End Sub

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim rs2 As New ADODB.Recordset, rs3 As New ADODB.Recordset

If vPaso Then Exit Sub

If strConsulta <> "N" And txtBeneficioId = "" And txtCedula <> "" And bNuevo = False Then
    'Consulta el beneficio
     
      strSQL = "select * from afi_beneficios where cod_beneficio = '" & cbo.ItemData(cbo.ListIndex) & "'"
      Call OpenRecordSet(rs, strSQL)
      If rs.RecordCount > 0 Then
        curMonto = fxMonto(cbo.ItemData(cbo.ListIndex))
        
        If bMostroMensaje = True Then
          Call btnNuevo_Click
          bMostroMensaje = False
          Exit Sub
        End If
        
        If curMonto = 0 Then
           tcMain.Item(1).Selected = True
           cbo.SetFocus
          Exit Sub
        End If
        
        'Llama procedimiento que verifica beneficiarios,si modifica monto,el tipo
        'si acepta beneficiarios
        bAplicaParcial = IIf(rs!aplica_parcial = 0, False, True)
        iBeneficiario = rs!aplica_beneficiarios
        curDiferencia = rs!modifica_diferencia
        
        txtMonto = Format(curMonto, "Standard")
        txtMonto.Tag = curMonto
        If rs!modifica_monto = 1 Then
            txtMonto.Locked = False
        Else
            txtMonto.Locked = True
        End If
        
        txtDisponible.Text = Format(txtMonto.Text, "Standard")
        
        If rs!tipo_monetario = 1 And rs!tipo_producto = 1 Then
           'frmAF_BeneficioTipo.Show vbModal
           cboTipoBeneficio.Enabled = True
           
        ElseIf rs!tipo_monetario = 1 Then
           cboTipoBeneficio.Text = "Monetario"
           cboTipoBeneficio.Enabled = False
        Else
           cboTipoBeneficio.Text = "Producto"
           cboTipoBeneficio.Enabled = False
        End If
        
        Call sbCargaInformacion(rs!modifica_monto, Trim(Mid(cboTipoBeneficio.Text, 1, 1)))
        rs.Close
      End If
Else
  strConsulta = "S"
  txtMontoAsg.Locked = False
End If

End Sub

Private Sub cboBanco_Click()
Dim strSQL As String

If vPaso Or cboBanco.ListCount = 0 Then Exit Sub


On Error GoTo vError

If txtCedula.Text <> "" Then
    strSQL = "exec spSys_Cuentas_Bancarias '" & lblPagoCaso.Tag & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
    Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)
End If

If lswPago.ListItems.Count > 0 Then
    lblPagoCaso.Tag = lswPago.SelectedItem.Text
    lswPago.SelectedItem.SubItems(5) = Trim(cboBanco.Text)
    lswPago.SelectedItem.SubItems(6) = cboCuenta.ItemData(cboCuenta.ListIndex)
End If



vError:


End Sub

Private Sub cboEmitir_Click()
On Error GoTo vError

If txtCedula <> "" And lswPago.ListItems.Count > 0 Then
    lswPago.SelectedItem.SubItems(4) = cboEmitir.Text
End If

vError:

End Sub


Private Sub cboTipoBeneficio_Click()
Call sbActivaFrame
End Sub

Private Sub sbGuardar()
    
If lswAsignados.Enabled = True And txtBeneficioId.Text = "" Then
   MsgBox "Necesita seleccionar la Opcion Nuevo para poder incluir un beneficio"
   
   tcMain.Item(0).Selected = True
   Exit Sub
End If

If iBeneficiario = 1 Then
  If Trim(txtCedulaFallecido) = Empty Or Trim(txtNombreFallecido) = Empty Then
     MsgBox "Verifique los datos del Fallecido"
     Exit Sub
  End If
End If

Select Case UCase(Mid(cboTipoBeneficio, 1, 1))
  Case "M"
   ' If curTotal = 0 And Not bConsulta Then
   '     MsgBox "NO se almacenó la informacion", vbInformation
   ' Exit Sub
   ' Else
     If bAplicaParcial = True Then
        If txtDisponible > 0 Then
           txtMonto = txtMontoAsg
        End If
        Call sbGuarda_Beneficio
        
     ElseIf CCur(txtDisponible) = 0 And lswPago.ListItems.Count > 0 Then
        Call sbGuarda_Beneficio
     Else
        MsgBox "No ha distribuido el disponible", vbInformation
        tcMain.Item(1).Selected = True
     Exit Sub
     End If
    'End If
  
  Case "P"
    If txtProdCantidad.Tag = 1 Then
       Call sbGuarda_Productos
    Else
       MsgBox "No se almaceno la informacion", vbInformation
       
       Exit Sub
    End If
End Select

End Sub


Private Sub Form_Activate()
vModulo = 7
txtCedula.SetFocus
End Sub



Private Sub Form_Load()
 
 vModulo = 7
 
 imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
 
 With lswAsignados.ColumnHeaders
    .Clear
    .Add , , "Nº Beneficio", 1800
    .Add , , "Beneficio", 3000
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Estado", 1300, vbCenter
 End With
 
 
 With lswPago.ColumnHeaders
    .Clear
    .Add , , "Identificación", 1800
    .Add , , "Tipo", 700, vbCenter
    .Add , , "Nombre", 3000
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Emitir", 700, vbCenter
    .Add , , "Banco", 2500
    .Add , , "Cuenta", 1800
 End With
 
 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Descripción", 3000
    .Add , , "Cantidad", 1200, vbCenter
    .Add , , "Costo/Ud", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
 End With
 
 
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

 strConsulta = "S"
 txtMontoAsg.Locked = True
 bNotieneBancos = False

End Sub


Private Sub sbInicializa()
Dim strSQL As String

On Error GoTo vError

Call sbLimpiaPantalla

bConsulta = False
txtProdCantidad.Tag = 0

tcMain.Item(1).Selected = True

cboEstado.Enabled = True
cboEstado.Clear
cboBanco.Clear
cbo.Clear
cboTipoBeneficio.Clear

strConsulta = "N"

vPaso = True

strSQL = "select rtrim(cod_Beneficio) as 'IdX',rtrim(descripcion) as 'ItmX' from afi_beneficios " _
        & " where estado = 'A' and cod_beneficio in (select cod_beneficio from AFI_BENE_GRUPOSB " _
        & " where cod_grupo in(  select cod_grupo from AFI_BENE_USERG where usuario = '" & glogon.Usuario & "'))"
Call sbCbo_Llena_New(cbo, strSQL, False, True)


'Carga Cuentas Bancarias Autorizadas
'strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"


strSQL = "select B.id_banco as 'IdX',rtrim(B.descripcion) as 'ItmX'" _
       & " from tes_banco_asg T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
       & " where T.nombre = '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

 
vPaso = False
 
 cboEmitir.Clear
 cboEmitir.AddItem fxTipoDocumento("CK")
 cboEmitir.AddItem fxTipoDocumento("TE")
 cboEmitir.Text = fxTipoDocumento("TE")
 
 cboEstado.AddItem "SOLICITADO"
 cboEstado.AddItem "PENDIENTE"

 cboTipoBeneficio.AddItem "MONETARIO"
 cboTipoBeneficio.AddItem "PRODUCTO"
 cboTipoBeneficio.Text = "MONETARIO"

 lsw.ListItems.Clear
 lswPago.ListItems.Clear
 
 txtCedulaFallecido.Text = ""
 txtNombreFallecido.Text = ""
 txtMontoAsg.Enabled = True
 txtMontoAsg.Locked = False
 
 If Not txtCedulaFallecido.Locked Then
   txtCedulaFallecido.Locked = False
   txtCedulaFallecido.BackColor = vbWhite
 End If
 
Call RefrescaTags(Me)

bNuevo = False

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaInformacion(IntModcurMonto As Integer, strTipo As String)
 Call sbCarga_Lista(strTipo)
End Sub


Private Sub sbImprimeBoleta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vBanco As String, vDescribe As String, vContaCta As String
Dim vCtaBene As String, vDescribeBene As String, vMontoBene As Currency

If txtBeneficioId.Text <> "" And txtCedula <> "" Then
    
    Me.MousePointer = vbHourglass
    
    
    
    strSQL = "select cod_banco from afi_bene_pago where  cedula = '" & txtCedula & "'" _
            & " and cod_beneficio = '" & cbo.ItemData(cbo.ListIndex) & "' and consec = " & txtBeneficioId.Text
            
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Or Not rs.BOF Then
      vBanco = IIf(Not IsNull(rs!cod_banco), rs!cod_banco, "")
    
    'If vBanco = "" Then Exit Sub
    'rs.Close
     If rs.State > 0 Then rs.Close
    strSQL = "select ctaconta,descripcion from Tes_Bancos where id_banco = '" & vBanco & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.RecordCount > 0 Then
       vContaCta = IIf(Not IsNull(rs!ctaConta), rs!ctaConta, "")
       vDescribe = IIf(Not IsNull(rs!Descripcion), rs!Descripcion, "")
    End If
    rs.Close
    
    strSQL = "select cod_cuenta,descripcion from afi_beneficios where cod_beneficio = '" & cbo.ItemData(cbo.ListIndex) & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs.RecordCount > 0 Then
       vCtaBene = IIf(Not IsNull(rs!cod_cuenta), rs!cod_cuenta, "")
       vDescribeBene = IIf(Not IsNull(rs!Descripcion), rs!Descripcion, "")
    End If
    End If
    rs.Close
    
    strSQL = "select monto from afi_bene_otorga where consec = " & txtBeneficioId.Text _
            & " and cedula = '" & txtCedula.Text & "' and cod_Beneficio = '" & cbo.ItemData(cbo.ListIndex) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    If rs.RecordCount > 0 Then
       vMontoBene = IIf(Not IsNull(rs!MONTO), Format(rs!MONTO, "Standard"), 0)
    End If
    rs.Close
    
    
    With frmContenedor.Crt
        .Reset
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowState = crptMaximized
        .WindowTitle = "Módulo de Beneficios y Ayudas Sociales"
        .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(2) = "fxCodigoBarras = '*" & txtBeneficioId.Text & "*'"
        
        .Connect = glogon.ConectRPT
        
        If Mid(cboTipoBeneficio, 1, 1) = "M" Then
          .ReportFileName = SIFGlobal.fxPathReportes("Beneficios_Boletam.rpt")
          strSQL = "{AFI_BENE_OTORGA.CONSEC} = " & txtBeneficioId.Text & " AND {AFI_BENE_OTORGA.cedula} = '" & txtCedula.Text & "'" & " AND {AFI_BENE_OTORGA.COD_BENEFICIO} = '" & cbo.ItemData(cbo.ListIndex) & "'"
          
          .SelectionFormula = strSQL
          .SubreportToChange = "Asiento"
          .Formulas(7) = "fxcuentabanco = '" & Trim(vContaCta) & "'"
          .Formulas(8) = "fxDescripcion = '" & Trim(vDescribe) & "'"
          .Formulas(9) = "fxDescribe = '" & Trim(vDescribeBene) & "'"
          .Formulas(10) = "fxcuenta = '" & Trim(vCtaBene) & "'"
          .Formulas(11) = "fxmonto = " & vMontoBene & ""
          .Formulas(12) = "fxmontobene = " & vMontoBene & ""
          
        Else
          strSQL = "{AFI_BENE_OTORGA.CONSEC} = " & txtBeneficioId.Text & " AND {AFI_BENE_OTORGA.COD_BENEFICIO} = '" & cbo.ItemData(cbo.ListIndex) & "'"
          .SelectionFormula = strSQL
          .ReportFileName = SIFGlobal.fxPathReportes("Beneficios_Boletap.rpt")
        End If
        
                
        .PrintReport
    End With
    
    Me.MousePointer = vbDefault
    
End If

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


vProducto = Item.Text
txtProducto = Item.SubItems(1)
txtProdCantidad = Item.SubItems(2)
txtProdCosto = Item.SubItems(3)
bConsulta = True
txtProdCantidad.SetFocus

End Sub

Private Sub lswAsignados_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  
strConsulta = "S"
lswPago.ListItems.Clear
cboCuenta.Clear

iGrupo = 0
bAsignado = False
bNuevo = False

txtBeneficioId.Text = Item.Text

Call sbConsulta(Trim(Item.Tag), txtBeneficioId.Text)

End Sub

Private Sub lswPago_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim vTexto As String

On Error GoTo vError

    vPaso = False

    lblPagoCaso.Tag = Item.Text
    lblPagoCaso.Caption = Item.SubItems(2)
    
    txtMontoAsg.Text = Item.SubItems(3)
    
    curTotal = txtMontoAsg.Text
    cMontoPagar = txtMontoAsg.Text
    
    vTexto = Item.SubItems(5)
    
    cboBanco.Text = IIf((vTexto = ""), cboBanco.Text, _
                        Trim(Item.SubItems(5)))
        
    vTexto = Item.SubItems(4)
    
    cboEmitir.Text = IIf((vTexto = "" Or vTexto <> cboEmitir.Text), cboEmitir.Text, _
                         Item.SubItems(4))
    
    
    If txtMontoAsg.Enabled Then
        If frmMonetario.Visible Then txtMontoAsg.SetFocus
    End If
    
    Call cboBanco_Click
    Call cboEmitir_Click
    
    Call sbCboAsignaDato(cboCuenta, Item.SubItems(6), True, Item.SubItems(6))
    
    
    
    Exit Sub

vError:
   MsgBox "No tienen el banco asignado"

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

frmProducto.Visible = False
frmMonetario.Visible = False

If Item.Index = 2 And Mid(cboTipoBeneficio.Text, 1, 1) = "M" Then
    frmMonetario.Visible = True
    Call cbo_Click
End If


If Item.Index = 2 And Mid(cboTipoBeneficio.Text, 1, 1) = "P" Then
    frmProducto.Visible = True
End If

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

 Call sbInicializa
 
 bNuevo = False
 bConsulta = True
 
 
If GLOBALES.gTag <> "" And GLOBALES.gTag2 = "" Then
   txtCedula.Text = GLOBALES.gTag
   Call txtCedula_LostFocus
End If


If IsNumeric(GLOBALES.gTag) And GLOBALES.gTag2 <> "" Then

    strConsulta = "S"
    lswPago.ListItems.Clear
    cboCuenta.Clear
    
    iGrupo = 0
    bAsignado = False
    bNuevo = False
    
    txtBeneficioId.Text = GLOBALES.gTag

    Call sbConsulta(GLOBALES.gTag2, CLng(GLOBALES.gTag))
End If


End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
bNuevo = True
If KeyCode = vbKeyF4 Then
    lswPago.ListItems.Clear
    
    txtCedula = ""
    vCedula = ""
    
    gBusquedas.Col1Name = "Cédula Colilla"
    gBusquedas.Col2Name = "Cédula Real"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "Cedula"
    gBusquedas.Orden = "Cedula"
    
    
    frmBusquedas.Show vbModal
    
    txtCedula.Text = Trim(gBusquedas.Resultado)
    txtNombre.Text = gBusquedas.Resultado2
    
    vCedula = txtCedula.Text
    If Trim(txtCedula.Text) <> "" Then Call sbConsultaX(txtCedula)
    
End If

If KeyCode = vbKeyReturn Then
    tcMain.Item(1).Selected = True
    Call cbo.SetFocus
End If

End Sub

Private Sub txtCedula_LostFocus()
txtNombre.Text = fxNombre(txtCedula)

If Trim(txtCedula.Text) <> "" Then Call sbConsultaX(txtCedula)

bNuevo = False
End Sub

Private Sub txtCedulaFallecido_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNombreFallecido.SetFocus

End Sub

Private Sub txtCedulaFallecido_LostFocus()
'Dim strSQL As String, rs As New ADODB.Recordset
'
'If Trim(txtCedulaFallecido) = "" Then Exit Sub
'
'If Trim(txtCedulaFallecido) <> "" And bConsulta = False Then
'    strSQL = "Select * from afi_bene_otorga where solicita = '" & Trim(txtCedulaFallecido) & "'"
'    Call OpenRecordSet(rs, strSQL)
'    If Not rs.EOF Or Not rs.BOF Then
'       MsgBox "El Fallecido identificado como " & rs!Nombre & " cédula " & rs!solicita & vbCrLf & _
'       "ya fue registrado con la boleta N°" & rs!consec & _
'      " Beneficio " & rs!cod_beneficio
'       txtCedulaFallecido.SetFocus
'    End If
'End If

End Sub




Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto.Text = CCur(txtMonto.Text)
vError:

End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then

   For i = 1 To lswPago.ListItems.Count
      lswPago.ListItems.Item(i).SubItems(3) = 0
   Next i
   
   txtDisponible.Text = Format(txtMonto.Text, "Standard")
   txtMontoAsg.SetFocus
   
End If
End Sub

Private Sub txtMonto_LostFocus()
Dim cDiferencia As Currency

On Error GoTo vError

   curTotalDif = CCur(txtMonto.Tag) + curDiferencia
   cDiferencia = CCur(txtMonto.Tag) - curDiferencia
  
  
   If Abs(txtMonto.Text) > Abs(curTotalDif) Or Abs(txtMonto.Text) < Abs(cDiferencia) Then
      MsgBox "se sobrepaso del monto permitido", vbInformation
      txtMonto.Text = Format(curMonto, "Standard")
      txtDisponible.Text = Format(curMonto, "Standard")
   Else
      txtMonto.Text = Format(txtMonto.Text, "Standard")
      curTotalDif = 0
   End If

vError:

End Sub


Private Sub txtMontoAsg_Change()
    
If Mid(cboEstado.Text, 1, 1) = "A" Or Mid(cboEstado.Text, 1, 1) = "E" _
   Or Mid(cboEstado.Text, 1, 1) = "R" Then Exit Sub
If Not IsNumeric(txtMontoAsg.Text) Then txtMontoAsg.Text = 0
    
If txtMontoAsg.Text > 0 Then
    If CCur(txtMontoAsg) <= CCur(txtMonto.Text) And CCur(txtMontoAsg.Text) >= 0 Then
       txtDisponible.Text = Format(txtMonto.Text - txtMontoAsg.Text, "Standard")
    Else
     MsgBox "Monto digitado no es valido"
     txtMontoAsg.Text = 0
     txtDisponible.Text = txtMonto.Text
    End If
End If

If txtMontoAsg.Text = 0 Then txtDisponible.Text = txtMonto.Text

End Sub

Private Sub txtMontoAsg_GotFocus()

On Error GoTo vError
  
  If txtMontoAsg.Text > 0 Then txtMontoAsg.Text = CCur(txtMontoAsg.Text)

vError:

End Sub

Private Sub txtMontoAsg_LostFocus()
On Error GoTo vError

If txtMontoAsg.Text = "" Then txtMontoAsg.Text = 0
txtMontoAsg.Text = CCur(txtMontoAsg.Text)   ', "Standard"
    
curDisponible = 0
    
curTotalDif = 0
If Trim(txtMontoAsg.Text) <> "" Then

    If Val(txtMonto.Text) > 0 And CCur(txtMontoAsg.Text) <> CCur(curTotal) Then
      
      If Val(txtMontoAsg) > 0 Then
        curTotal = 0
        If CCur(txtMontoAsg.Text) > CCur(txtMonto.Text) Then
           MsgBox "El monto no puede ser mayor que el disponible", vbCritical
           txtMontoAsg.Text = CCur(lswPago.SelectedItem.SubItems(3))
           Exit Sub
        End If
            
          lswPago.SelectedItem.SubItems(3) = Format(txtMontoAsg.Text, "Standard")
          curDisponible = txtDisponible.Text
          curTotal = txtMontoAsg.Text
          txtDisponible.Text = Format(CCur(txtMonto.Text) - CCur(curTotal), "Standard")
          
         Else
          
           txtDisponible.Text = Format(CCur(txtMonto.Text) + CCur(curTotal), "Standard")
           curTotal = 0
           If txtDisponible.Text > txtMonto.Text Then txtMonto.Text = Format(txtDisponible.Text, "Standard")
           If txtDisponible.Text > cMontoBene Then
            txtMonto.Text = Format(cMontoBene, "Standard")
            txtDisponible.Text = Format(txtMonto.Text, "Standard")
           End If
           If lswPago.ListItems.Count > 0 Then lswPago.SelectedItem.SubItems(3) = 0
        End If

    End If
Else
  curTotal = 0
  
    txtDisponible.Text = Format(CCur(txtMonto.Text) + CCur(curTotal), "Standard")
 
  If txtDisponible.Text > txtMonto.Text Then txtMonto.Text = Format(txtDisponible.Text, "Standard")
  
  'If txtDisponible > cMontoBene And iCantidaGrupo <= 1 Then
  If cMontoBene > cMontoAsignado Then
     txtMonto.Text = Format(cMontoBene, "Standard")
     txtDisponible.Text = Format(txtMonto.Text, "Standard")
  Else
    txtDisponible.Text = Format(cMontoGrupo - cMontoAsignado + cMontoPagar, "Standard")
    txtMonto.Text = Format(txtDisponible.Text, "Standard")
  End If
  
  
  If lswPago.ListItems.Count > 0 Then lswPago.SelectedItem.SubItems(3) = 0
           
End If
Exit Sub

vError:

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then

  lswPago.ListItems.Clear
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
  cbo.SetFocus
  If Trim(txtCedula.Text) <> "" Then Call sbConsultaX(txtCedula)
End If
End Sub

Private Sub sbLimpiaPantalla()

lswPago.ListItems.Clear

'vCedula = ""
'txtCedula = ""
lswAsignados.ListItems.Clear
txtBeneficioId.Text = Empty
txtDisponible = 0

txtMonto.Text = 0
txtMontoAsg.Text = 0
txtNotas.Text = ""
lblAutFecha.Caption = ""
lblAutUser.Caption = ""
lblPagoCaso.Caption = ""
txtProdCantidad.Text = ""
txtProdCosto.Text = ""
txtProducto.Text = ""

End Sub

Private Sub sbCarga_Lista(strListaCarga As String)
Dim itmX As ListViewItem

If strListaCarga = "M" Then 'if que define cual lista cargar sea monetario o product
   vCedula = txtCedula.Text
   lswPago.ListItems.Clear
   Set itmX = lswPago.ListItems.Add(, , txtCedula.Text)
       itmX.SubItems(1) = "S"
       itmX.SubItems(2) = txtNombre.Text
       itmX.SubItems(3) = 0
       itmX.SubItems(4) = cboEmitir.Text
       itmX.SubItems(5) = cboBanco.Text
       itmX.SubItems(6) = ""
       
       Call cboEmitir_Click
       
       If iBeneficiario = 1 Then
           txtCedulaFallecido.Locked = False
           txtNombreFallecido.Locked = False
           txtCedulaFallecido.BackColor = vbWhite
           txtNombreFallecido.BackColor = vbWhite
'           txtCedulaFallecido.SetFocus
       Else
           txtCedulaFallecido.Locked = True
           txtNombreFallecido.Locked = True
           txtCedulaFallecido.BackColor = txtNombre.BackColor
           txtNombreFallecido.BackColor = txtNombre.BackColor
       End If
 End If 'fin del if que define cual lista cargar sea monetario o producto
             
End Sub

Private Sub sbConsulta(vCodigo As String, intCodBene As Long)
Dim rs As New ADODB.Recordset, strSQL As String, rs2 As New ADODB.Recordset
Dim strTipo As String, strEstado As String, itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

    cboEstado.Clear
    cboEstado.AddItem "SOLICITADO"
    cboEstado.AddItem "PENDIENTE"
    cboEstado.AddItem "APROBADO"
    cboEstado.AddItem "RECHAZADO"
    cboEstado.Text = "SOLICITADO"

    tcMain.Item(1).Selected = True
    
    bConsulta = True
    curTotal = 0
    
    'Consulta beneficio para traer el codigo y descrpción
    strSQL = "select rtrim(cod_Beneficio) as 'IdX', rtrim(descripcion) as 'ItmX',tipo,monto" _
           & " ,modifica_diferencia,aplica_beneficiarios,aplica_parcial" _
           & " from afi_beneficios" _
           & " where Cod_Beneficio = '" & vCodigo & "'"
    Call OpenRecordSet(rs, strSQL)
            
    
    
    
    'curMonto = fxMontoBene(vCodigo)
    curMonto = fxMonto(vCodigo)
    txtMonto.Text = Format(curMonto, "Standard")
    txtMonto.Tag = Format(curMonto, "Standard")
    
    curDiferencia = rs!modifica_diferencia
    
    intCodBene = rs!aplica_beneficiarios
    If rs.RecordCount = 1 Then
    
        bAplicaParcial = IIf(rs!aplica_parcial = 0, False, True)
        
        Call sbCboAsignaDato(cbo, rs!itmX, True, rs!IdX)
    End If
    rs.Close
    
    strSQL = "select A.*,O.descripcion, S.Nombre" _
           & " from afi_bene_otorga A" _
           & "  inner join Socios S on A.cedula = S.cedula" _
           & "   left join Sif_Oficinas O on A.cod_oficina = O.cod_Oficina" _
           & " where cod_beneficio = '" & vCodigo & "' and consec = " & txtBeneficioId.Text
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.BOF And Not rs.EOF Then
        vTipo = rs!Tipo
        
        txtCedula.Text = rs!Cedula
        txtNombre.Text = rs!Nombre
        
        Select Case vTipo
         Case "M"
             cboTipoBeneficio.Text = "MONETARIO"
         Case "P"
             cboTipoBeneficio.Text = "PRODUCTO"
         End Select
        
        'Call sbActivaFrame
        
        txtProdCantidad.Tag = 1
        
        txtNotas = rs!notas & ""
        cboEstado.AddItem "EJECUTADO"
        cboEstado.Text = fxEstadoBeneficio(rs!Estado)
        
        Select Case rs!Estado
            Case "E", "R" 'Ejecutado, Rechazado y Aprobado , "A"
                btnGuardar.Enabled = False
                txtProducto.Enabled = False
                cboEstado.Enabled = False
                txtMontoAsg.Enabled = False
            
            Case "S", "P", "A"  'Solicitado y Pendiente
                'cmdGuardar.Enabled = True
                btnGuardar.Enabled = True
                txtProducto.Enabled = True
                cboEstado.Enabled = True
                txtMontoAsg.Enabled = True
        
        End Select
        
        lblAutFecha.Caption = IIf(Not IsNull(rs!Autoriza_Fecha), Format(rs!Autoriza_Fecha, "dd/mm/yyyy"), "")
        lblAutUser.Caption = IIf(Not IsNull(rs!Autoriza_user), rs!Autoriza_user, "")
        
        StatusBarX.Panels(1).Text = IIf(Not IsNull(rs!registra_user), rs!registra_user, "")
        StatusBarX.Panels(2).Text = IIf(Not IsNull(rs!Registra_Fecha), rs!Registra_Fecha, "")
        StatusBarX.Panels(3).Text = IIf(Not IsNull(rs!Descripcion), rs!Descripcion, "")
        StatusBarX.Panels(4).Text = strEstado
        
        txtNotas.Text = rs!notas
        
        txtCedulaFallecido.Text = IIf(IsNull(rs!Solicita), "", rs!Solicita)
        txtNombreFallecido.Text = IIf(IsNull(rs!Nombre), "", rs!Nombre)
        
        If txtCedulaFallecido <> "" And rs!Estado <> "A" And rs!Estado <> "E" And rs!Estado <> "R" Then
          txtCedulaFallecido.Locked = False
          txtNombreFallecido.Locked = False
          txtCedulaFallecido.BackColor = vbWhite
          txtNombreFallecido.BackColor = txtCedulaFallecido.BackColor
        Else
          txtCedulaFallecido.Locked = True
          txtNombreFallecido.Locked = True
          txtCedulaFallecido.BackColor = txtNombre.BackColor
          txtNombreFallecido.BackColor = txtCedulaFallecido.BackColor
        End If
        
        
        
        rs.Close
    
    
    
    
    Select Case vTipo
      Case "M" 'en caso de que sea Monetario
        
'               & ", dbo.fxSys_Cuenta_Bancos_Desc(COD_BANCO) as 'BancoDesc'" _
'               & ", dbo.fxSys_Cuentas_Bancarias_Desc(cedula, COD_BANCO,Cta_Bancaria) as 'CuentaDesc'" _

        strSQL = "select Bp.*, B.Descripcion as 'BancoDesc'" _
               & " from afi_bene_pago Bp" _
               & "  left join Tes_Bancos B on Bp.cod_Banco = B.id_Banco " _
               & " where Bp.consec = " & txtBeneficioId.Text & " and Bp.cod_beneficio = '" & vCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          Select Case rs!Tipo
               Case "S"
                 strTipo = "Socio"
               Case "B"
                 strTipo = "Beneficiario"
          End Select
          
          Set itmX = lswPago.ListItems.Add(, , rs!Cedula)
                 
          
          If intCodBene = 1 Then
                strSQL = "select Nombre from beneficiarios where cedulabn = '" & rs!Cedula & "' and cedula = '" & txtCedula & "'"
                rs2.Open strSQL, glogon.Conection, adOpenStatic
                If Not rs2.EOF Then
                    itmX.SubItems(2) = rs2!Nombre
                    rs2.MoveNext
                Else
                    itmX.SubItems(2) = txtNombre.Text
                End If
                rs2.Close
          Else
               itmX.SubItems(2) = txtNombre.Text
          End If
          
          itmX.SubItems(1) = strTipo
          itmX.SubItems(3) = Format(rs!MONTO, "Standard")
          itmX.SubItems(6) = rs!cta_bancaria & ""
          itmX.SubItems(5) = rs!BancoDesc & ""
          itmX.SubItems(4) = IIf(Not IsNull(rs!Tipo_Emision), fxTipoDocumento(rs!Tipo_Emision), "")
          
          lblPagoCaso.Caption = itmX.SubItems(2)
          txtMontoAsg.Text = Format(rs!MONTO, "Standard")
          
          If rs!cod_banco > 0 Then
            vPaso = True
            Call sbCboAsignaDato(cboBanco, rs!BancoDesc, True, IIf(IsNull(rs!cod_banco), 0, rs!cod_banco))
            vPaso = False
          End If
            
            'Carga Cuentas de la Persona
            Call cboBanco_Click

            'Asigna Cuenta Utilizada
            Call sbCboAsignaDato(cboCuenta, rs!cta_bancaria, True, IIf(IsNull(rs!cta_bancaria), "", rs!cta_bancaria))
'
          cboEmitir = Trim(IIf(Not IsNull(rs!Tipo_Emision), fxTipoDocumento(rs!Tipo_Emision), ""))
          rs.MoveNext
        Loop
         
         For i = 1 To lswPago.ListItems.Count
            If txtMontoAsg.Enabled = True Then curTotal = curTotal + CCur(lswPago.ListItems.Item(i).SubItems(3))
         Next i
'         If txtMonto > curTotal Then
            
            
            '' verificar aqui el monto
            
            
            If iGrupo > 0 And bAsignado = True Then
               If iCantidaGrupo > 1 Then
                  txtMonto.Text = Format(cMontoGrupo - cMontoAsignado, "Standard")
                  If CCur(txtDisponible.Text) >= curTotal Then txtDisponible.Text = Format(CCur(txtMonto.Text) - curTotal, "Standard")
               Else
                 txtMonto.Text = Format(cMontoGrupo, "Standard")
                 If CCur(txtDisponible.Text) >= curTotal Then txtDisponible.Text = Format(CCur(txtMonto.Text) - curTotal)    'Format(txtMonto, "Standard")
               End If
            End If
            
            If CCur(txtMonto.Text) = 0 Then txtDisponible.Text = 0
            
            If lswPago.ListItems.Count > 0 Then lswPago.ListItems.Item(1).Selected = True
            
            
        
      Case "P" 'En caso de que se producto
        bConsulta = True
        lsw.ListItems.Clear
       
        strSQL = "select R.*, P.Descripcion as 'ProdDesc', P.costo_unidad as 'ProdCu' " _
               & " from afi_bene_prodasg R inner join afi_bene_productos P on R.cod_Producto = P.cod_Producto " _
               & " where R.consec = " & txtBeneficioId.Text & " and R.cod_beneficio = '" & vCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
           Set itmX = lsw.ListItems.Add(, , rs!cod_producto)
               itmX.SubItems(1) = rs!ProdDesc
               itmX.SubItems(2) = rs!Cantidad
               itmX.SubItems(3) = rs!costo_unidad
               itmX.SubItems(4) = Format(rs!Cantidad * rs!costo_unidad, "Standard")
         rs.MoveNext
        Loop
      End Select
    Else
      MsgBox "No se encontró registro verifique...", vbInformation
    End If


If rs.State > 0 Then rs.Close
    
Me.MousePointer = vbDefault
cboTipoBeneficio.Enabled = False
   
Exit Sub
    
vError:
     Me.MousePointer = vbDefault
     MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGuarda_Beneficio()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iCodBanco As Integer
Dim strCedula As String, strTipo As String, strEmision As String
Dim strCta As String

Dim vBeneCod As String, vBeneConsec As Long, vBeneEstado As String

On Error GoTo vError

vBeneCod = cbo.ItemData(cbo.ListIndex)

If txtBeneficioId.Text = "" Then


    vBeneEstado = fxEstadoBeneficio(Trim(cboEstado.Text))
    vBeneConsec = fxConsec(vBeneCod)
    
    txtBeneficioId.Text = vBeneConsec
    
    strSQL = "insert afi_bene_otorga(consec,cod_beneficio,cedula,monto,modifica_monto,registra_user,registra_fecha,estado" _
           & ",notas,Solicita,nombre,tipo,cod_oficina) " _
           & " values(" & vBeneConsec & ",'" & vBeneCod & "','" & txtCedula & "'," & CCur(txtMonto.Text) & "," _
           & "'" & IIf((txtMonto.Text = txtMonto.Tag), "N", "S") & "','" & glogon.Usuario & "',dbo.MyGetdate(),'" _
           & vBeneEstado & "','" & Trim(txtNotas) & "','" & txtCedulaFallecido & "','" & UCase(txtNombreFallecido) _
           & "','" & Mid(cboTipoBeneficio, 1, 1) & "','" & GLOBALES.gOficinaTitular & "')"
    Call ConectionExecute(strSQL)
 
    Call Bitacora("Registra", "Beneficio: " & vBeneConsec & vBeneCod)
    
    For i = 1 To lswPago.ListItems.Count
        strCedula = lswPago.ListItems.Item(i).Text
        strTipo = Trim(lswPago.ListItems.Item(i).SubItems(1))
        curMonto = CCur(lswPago.ListItems.Item(i).SubItems(3))
        strEmision = fxTipoDocumento(lswPago.ListItems.Item(i).SubItems(4))
        strCta = lswPago.ListItems.Item(i).SubItems(6)
        iCodBanco = fxCodigoBanco(lswPago.ListItems.Item(i).SubItems(5))
        
        strSQL = "insert afi_bene_pago(cedula,consec,cod_beneficio,tipo,monto,cod_banco" _
               & ", tipo_emision,cta_bancaria,estado)values('" & strCedula & "'," & vBeneConsec & ",'" & vBeneCod _
               & "','" & strTipo & "'," & curMonto & "," & iCodBanco & ",'" & strEmision _
               & "','" & Trim(strCta) & "','" & vBeneEstado & "')"
        Call ConectionExecute(strSQL)
    Next i
    
    Call sbSIFRegistraTags(CStr(vBeneConsec), "S.BEN.01", "Reg. Ben", vBeneCod, "BEN")

    MsgBox "Informacion Guardada Satisfactoriamente", vbOKOnly
    
    
Else
    vBeneEstado = fxEstadoBeneficio((cboEstado.Text))
    
    strSQL = "update afi_bene_otorga set notas = '" & Trim(txtNotas) & "',estado='" & vBeneEstado & "'," _
            & "modifica_monto = '" & IIf((CCur(txtMonto.Text) = CCur(txtMonto.Tag)), "N", "S") & "'," _
            & " solicita = '" & txtCedulaFallecido & "',monto = " & CCur(txtMonto) & ",nombre = '" & txtNombreFallecido & "'," _
            & " TIPO = '" & Mid(cboTipoBeneficio, 1, 1) _
            & "' where cod_beneficio = '" & vBeneCod & "' and cedula = '" & txtCedula & "' and consec = " & txtBeneficioId & " "
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Beneficio: " & txtBeneficioId & " - " & vBeneCod)
   
   For i = 1 To lswPago.ListItems.Count
        strCedula = lswPago.ListItems.Item(i).Text
        curMonto = lswPago.ListItems.Item(i).SubItems(3)
        strEmision = fxTipoDocumento(lswPago.ListItems.Item(i).SubItems(4))
        
        strCta = lswPago.ListItems.Item(i).SubItems(6)
        
        iCodBanco = fxCodigoBanco(lswPago.ListItems.Item(i).SubItems(5))
        
         strSQL = "update afi_bene_pago set monto = " & CCur(txtMonto) & ",cod_banco = " & iCodBanco & "," _
                   & "tipo_emision = '" & strEmision & "',cta_bancaria = '" & strCta & "',estado = '" & vBeneEstado & "' " _
                   & "where cedula = '" & strCedula & "' and consec = '" & txtBeneficioId & "' " _
                   & "and cod_beneficio = '" & vBeneCod & "'  "
        Call ConectionExecute(strSQL)
   Next i
 
 MsgBox "Información Modificada satisfactoriamente", vbOKOnly
 

End If

Call txtCedula_LostFocus
lswAsignados.Enabled = True
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Function fxConsec(vCodBene As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select isnull(Max(consec),0) as consecutivo from afi_bene_otorga " _
        & "where cod_beneficio = '" & vCodBene & "'"
Call OpenRecordSet(rs, strSQL)
  fxConsec = rs!Consecutivo + 1
rs.Close

End Function


Private Function fxValida(xCodigo As String) As Boolean

' valida si ya existe el beneficio en su maximo de otorgamientos
' para un socio determinado
Dim strSQL As String, rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset, iMaxOtorga As Integer
Dim vMensaje As String
Dim strCodigoBene As String

vMensaje = ""


strSQL = "Select maximo_otorga from afi_beneficios where cod_beneficio = '" & xCodigo & "' "
Call OpenRecordSet(rs, strSQL)
  iMaxOtorga = rs!maximo_otorga

rs.Close


'hacer aqui la validacion por grupo
strSQL = "select cod_grupo from afi_grupo_beneficio where cod_beneficio = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
   iGrupo = rs!Cod_Grupo
Else
  iGrupo = 0
End If
rs.Close

bAsignado = False
 
    If iGrupo > 0 Then
       strSQL = "Select monto from afi_bene_grupos where cod_grupo  = " & iGrupo & ""
       Call OpenRecordSet(rs, strSQL)
    
       If Not rs.EOF Then
          cMontoGrupo = rs!MONTO
       End If
       rs.Close
       strSQL = " Select count(*)as cantidad,isnull(sum(B.MONTO),0) as monto from afi_bene_otorga B inner join" _
                & " afi_grupo_beneficio G ON B.cod_beneficio = G.cod_beneficio" _
                & " where B.cedula ='" & txtCedula & "'"
       Call OpenRecordSet(rs, strSQL)
    
       If Not rs.EOF Then
          iCantidaGrupo = rs!Cantidad
          cMontoAsignado = rs!MONTO
          
          bAsignado = True
          If rs!MONTO >= cMontoGrupo Then
           If strConsulta = "N" Then
             vMensaje = vMensaje & vbCrLf & " - Sobrepasa el monto asignado al grupo de beneficios " & iGrupo
           End If
          Else
          cMontoRealGrupo = cMontoGrupo - rs!MONTO
          End If
           
       End If
       rs.Close
    
    End If
 
'valida cantidad de veces que se otorga el beneficio
If bConsulta = False Then
    strSQL = "select isnull(count(*),0) as  cantidad from afi_bene_otorga where cod_beneficio = '" & xCodigo & "' and " _
           & "cedula = '" & txtCedula & "'"
           'and estado <> 'R'"
    Call OpenRecordSet(rs, strSQL)
    
    If rs!Cantidad >= iMaxOtorga Then
       vMensaje = vMensaje & vbCrLf & " - Excede el numero de veces de Otorgamientos del Beneficio"
    End If
    rs.Close
End If

If Len(vMensaje) > 0 Then
  If bMostroMensaje = False Then
     MsgBox vMensaje, vbExclamation
     bMostroMensaje = True
  Else
    bMostroMensaje = False
  End If
  cboEstado.Text = "PENDIENTE"
  cboEstado.Enabled = False
  fxValida = False
Else
  fxValida = True
End If
End Function

Private Sub sbGuarda_Productos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strCodprod As String, strDesc As String, cMonto As Currency
Dim iCantidad As Integer, iCostUni As Long, strEstado As String

On Error GoTo vError
cMonto = 0
 Select Case cboEstado.Text
      Case "PENDIENTE"
           strEstado = "P"
      Case "SOLICITADO"
           strEstado = "S"
    End Select
    

If txtBeneficioId = "" Then
     txtBeneficioId.Text = fxConsec(cbo.ItemData(cbo.ListIndex))
     
     For i = 1 To lsw.ListItems.Count
        cMonto = cMonto + CCur(lsw.ListItems.Item(i).SubItems(4))
     Next i
     
     
     strSQL = "insert afi_bene_otorga(consec,cod_beneficio,cedula,monto," _
            & "modifica_monto,registra_user,registra_fecha,estado," _
            & "notas,Solicita,nombre,tipo)values(" & txtBeneficioId.Text & ",'" & cbo.ItemData(cbo.ListIndex) & "','" & txtCedula & "'," & CCur(cMonto) & "," _
            & "'" & IIf((txtMonto.Text = txtMonto.Tag), "N", "S") & "','" & glogon.Usuario & "',dbo.MyGetdate(),'" & strEstado & "','" & Trim(txtNotas) & "','" & txtCedulaFallecido & "','" & UCase(txtNombreFallecido) & "','" & Mid(cboTipoBeneficio, 1, 1) & "')"
    
     Call ConectionExecute(strSQL)
      
     Call Bitacora("Registra", "Beneficio: " & txtBeneficioId & cbo.ItemData(cbo.ListIndex))
       
    For i = 1 To lsw.ListItems.Count
          strCodprod = lsw.ListItems.Item(i).Text
          iCantidad = lsw.ListItems.Item(i).SubItems(2)
          iCostUni = lsw.ListItems.Item(i).SubItems(3)
          strSQL = "insert afi_bene_prodasg(consec,cod_beneficio,cod_producto,cantidad,costo_unidad)" _
                 & "values('" & txtBeneficioId & "','" & cbo.ItemData(cbo.ListIndex) & "','" & strCodprod & "'," & iCantidad & "," & iCostUni & ")"
          Call ConectionExecute(strSQL)
    Next i
    
    MsgBox "Informacion almacenada satisfactoriamente", vbOKOnly

Else
     For i = 1 To lsw.ListItems.Count
        cMonto = cMonto + CCur(lsw.ListItems.Item(i).SubItems(4))
     Next i
     
     
     strSQL = "update afi_bene_otorga set notas = '" & Trim(txtNotas) & "', monto = " & CCur(cMonto) _
              & " where cod_beneficio = '" & cbo.ItemData(cbo.ListIndex) _
              & "' and cedula = '" & txtCedula & "'"
     Call ConectionExecute(strSQL)
     
     Call Bitacora("Modifica", "Beneficio: " & txtBeneficioId.Text & cbo.ItemData(cbo.ListIndex))
     
     For i = 1 To lsw.ListItems.Count
     
         strCodprod = lsw.ListItems.Item(i).Text
         iCantidad = lsw.ListItems.Item(i).SubItems(2)
         strSQL = "update afi_bene_prodasg set cantidad = " & iCantidad & " where consec = '" & txtBeneficioId & "' " _
                & "and cod_beneficio ='" & cbo.ItemData(cbo.ListIndex) & "' and cod_producto = '" & strCodprod & "'"
          Call ConectionExecute(strSQL)
     Next i
     
     MsgBox "Informacion modificada satisfactoriamente", vbOKOnly
     Call sbConsulta(cbo.ItemData(cbo.ListIndex), Val(txtBeneficioId))

End If


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtNombreFallecido_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub


Private Sub txtProdCantidad_KeyDown(KeyCode As Integer, Shift As Integer)

Dim itmX As ListViewItem, cTotal As Currency

If Not IsNumeric(txtProdCantidad.Text) Then Exit Sub

If KeyCode = vbKeyReturn And Val(txtProdCantidad) > 0 And bConsulta = False Then
    
    If fxExiste(lsw, txtProducto) And bConsulta = True Then
        MsgBox "Producto ya fue ingresado", vbOKOnly
        txtProdCantidad = ""
        txtProdCosto = ""
        txtProducto = ""
        txtProdCantidad.Tag = 0
        Exit Sub
    Else
        txtProdCantidad.Tag = 1
        Set itmX = lsw.ListItems.Add(, , vProducto)
            itmX.SubItems(1) = txtProducto.Text
            itmX.SubItems(2) = txtProdCantidad.Text
            itmX.SubItems(3) = Format(txtProdCosto, "Standard")
            itmX.SubItems(4) = Format(txtProdCosto * txtProdCantidad, "Standard")
            cTotal = (Format(txtProdCosto * txtProdCantidad, "Standard")) + cMontoPagado
            If cTotal > cMontoBene Then
               MsgBox "Excede el monto del beneficio"
               lsw.ListItems.Clear
            End If
            txtProdCantidad = ""
            txtProdCosto = ""
            txtProducto = ""
    End If

ElseIf bConsulta And KeyCode = vbKeyReturn Then
        txtProdCantidad.Tag = 1
        lsw.SelectedItem.SubItems(2) = txtProdCantidad
        lsw.SelectedItem.SubItems(4) = Format(txtProdCosto * txtProdCantidad, "Standard")
        cTotal = (Format(txtProdCosto * txtProdCantidad, "Standard")) + cMontoPagado
        If cTotal > cMontoBene Then
           MsgBox "Excede el monto del beneficio"
           lsw.ListItems.Clear
        End If
            
        txtProdCantidad = ""
        txtProdCosto = ""
        txtProducto = ""
End If

End Sub

Private Sub txtProducto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Or KeyCode = vbKeyReturn Then
    bConsulta = False
    gBusquedas.Resultado = Trim(txtProducto)
    txtProducto = ""
    gBusquedas.Convertir = "N"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Consulta = "Select cod_producto,descripcion From afi_bene_productos"
    
    frmBusquedas.Show vbModal
    vProducto = Trim(gBusquedas.Resultado)
    txtProducto = Trim(gBusquedas.Resultado2)
    txtProdCosto = Format(fxPrecio(vProducto), "Standard")
    txtProdCantidad.SetFocus
End If

End Sub

Private Function fxPrecio(cProducto As String) As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select costo_unidad,descripcion from afi_bene_productos " _
       & "where cod_producto = '" & cProducto & "' "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then

  vDescripcion = rs!Descripcion
  fxPrecio = rs!costo_unidad
  rs.Close
End If
End Function

Private Function fxExiste(lista As Object, Descripcion As String) As Boolean
Dim i As Integer
   
   For i = 1 To lista.ListItems.Count
       If Trim(lista.ListItems.Item(i).SubItems(1)) = Trim(Descripcion) Then
          fxExiste = True
          bConsulta = True
          Exit Function
       Else
          fxExiste = False
       End If
   Next i
End Function


Private Function fxMonto(vCodigo As String) As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

strSQL = "select case when estadoactual = 'S' then datediff(d,fechaingreso,dbo.MyGetdate())" _
& " else 0 end as Membresia from socios where cedula = '" & txtCedula & "'"

Call OpenRecordSet(rs, strSQL)

If rs.EOF Then
    Exit Function
End If


strSQL = "select monto from afi_beneficio_montos where" _
        & " cod_beneficio = '" & vCodigo & "' and " & rs!membresia _
        & " between inicio and corte"
rs.Close
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    cMontoBene = rs!MONTO
    If iBeneficiario = 0 Then
        strSQL = "select sum(monto) as  monto  from afi_bene_otorga where cod_beneficio = '" & vCodigo & "' and " _
               & "cedula = '" & txtCedula & "'"
    Else
        strSQL = "select sum(monto) as  monto  from afi_bene_otorga where cod_beneficio = '" & vCodigo & "' and " _
               & "cedula = '" & txtCedula & "' and solicita = '" & txtCedulaFallecido & "' "
    End If

    rs2.Open strSQL, glogon.Conection, adOpenStatic

    If Not rs2.EOF Or Not rs2.BOF Then
        cMontoPagado = IIf(IsNull(rs2!MONTO), 0, rs2!MONTO)
        If cMontoPagado >= rs!MONTO And bConsulta = False And fxValida(vCodigo) = False And bNuevo = False Then
            MsgBox "Ya le fue asignado el monto de la ayuda"
            fxMonto = 0
            rs.Close
            rs2.Close
            Exit Function
        End If
    End If

    If bMostroMensaje = True Then
        Call btnNuevo_Click
        bMostroMensaje = False
        Exit Function
    End If

    If rs!MONTO <= 0 Then
        MsgBox vbCrLf & "- No cumple con la Membresia para este beneficio"
        cboEstado.Text = "PENDIENTE"
        cboEstado.Enabled = False
        fxMonto = rs!MONTO
    ElseIf bNuevo = False Then
        fxValida (vCodigo)
        cboEstado.Enabled = True
        If iGrupo > 0 Then
            If cMontoRealGrupo >= cMontoBene And bAsignado = False Then
                fxMonto = cMontoBene
            Else
                fxMonto = cMontoRealGrupo
            End If
        Else
            fxMonto = rs!MONTO
        End If
        cboEstado.Text = "SOLICITADO"
    Else
        rs.Close
        rs2.Close
        fxMonto = 0
        Exit Function
    End If
    
Else
    MsgBox vbCrLf & "- No se encontro membresia para esta persona en este beneficio"
    cboEstado.Text = "PENDIENTE"
    cboEstado.Enabled = False
    fxMonto = 0
End If

rs.Close
rs2.Close
End Function

Private Function fxMontoBene(vCodigo As String) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select case when estadoactual = 'S' then datediff(m,fechaingreso,dbo.MyGetdate())" _
           & " else 0 end as Membresia from socios where cedula = '" & txtCedula & "'"
        Call OpenRecordSet(rs, strSQL)
        

 strSQL = "select monto from afi_beneficio_montos where" _
        & " cod_beneficio = '" & vCodigo & "' and " & rs!membresia _
        & " between inicio and corte"
rs.Close
Call OpenRecordSet(rs, strSQL)

         
 If Not rs.EOF Or Not rs.BOF Then
    fxMontoBene = IIf(IsNull(rs!MONTO), 0, rs!MONTO)
 Else
    fxMontoBene = 0
 End If
rs.Close
End Function

Private Sub sbActivaFrame()
If Mid(cboTipoBeneficio, 1, 1) = "M" Then
   frmMonetario.Visible = True
   frmProducto.Visible = False
   
Else
   frmMonetario.Visible = False
   frmProducto.Visible = True
End If

End Sub
