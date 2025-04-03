VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmTES_Bancos_Cargado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Carga de Movimientos Bancarios"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   12180
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   7695
      _Version        =   1441793
      _ExtentX        =   13573
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   11895
      _Version        =   1441793
      _ExtentX        =   20976
      _ExtentY        =   12933
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
      Item(0).Caption =   "Cargado"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "txtArchivo"
      Item(0).Control(1)=   "Label1(2)"
      Item(0).Control(2)=   "vGrid"
      Item(0).Control(3)=   "fraAccion"
      Item(0).Control(4)=   "btnArchivo(0)"
      Item(0).Control(5)=   "btnArchivo(1)"
      Item(0).Control(6)=   "btnArchivo(2)"
      Item(1).Caption =   "Registro en Bancos"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "vGridId"
      Item(1).Control(1)=   "gbRegistro"
      Item(1).Control(2)=   "gbBuscar"
      Begin XtremeSuiteControls.GroupBox gbBuscar 
         Height          =   1695
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   2990
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnRegistro_Buscar 
            Height          =   495
            Left            =   9240
            TabIndex        =   33
            Top             =   960
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Buscar"
            BackColor       =   -2147483633
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Bancos_Cargado.frx":0000
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.DateTimePicker dtpRegistroInicio 
            Height          =   315
            Left            =   1320
            TabIndex        =   34
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.DateTimePicker dtpRegistroCorte 
            Height          =   315
            Left            =   2880
            TabIndex        =   35
            Top             =   480
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   556
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
         Begin XtremeSuiteControls.ComboBox cboFechas 
            Height          =   330
            Left            =   1320
            TabIndex        =   36
            Top             =   120
            Width           =   3135
            _Version        =   1441793
            _ExtentX        =   5530
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   330
            Left            =   6120
            TabIndex        =   37
            Top             =   840
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtNumDoc 
            Height          =   315
            Left            =   6120
            TabIndex        =   38
            Top             =   120
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboTipo 
            Height          =   330
            Left            =   6120
            TabIndex        =   43
            Top             =   480
            Width           =   2895
            _Version        =   1441793
            _ExtentX        =   5106
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtMntInicio 
            Height          =   315
            Left            =   1320
            TabIndex        =   46
            Top             =   840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   556
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
            Text            =   "0"
            BackColor       =   16777215
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMntCorte 
            Height          =   315
            Left            =   2880
            TabIndex        =   47
            Top             =   840
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   556
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
            Text            =   "999,999,999,999.99"
            BackColor       =   16777215
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   315
            Left            =   1320
            TabIndex        =   52
            Top             =   1200
            Width           =   7695
            _Version        =   1441793
            _ExtentX        =   13573
            _ExtentY        =   556
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
            BackColor       =   16777215
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnReglas_Review 
            Height          =   495
            Left            =   10440
            TabIndex        =   54
            ToolTipText     =   "Revisar Reglas de Auto-Registro en Pendientes"
            Top             =   960
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Revisar Reglas"
            BackColor       =   -2147483633
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Bancos_Cargado.frx":0700
            ImageAlignment  =   4
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción"
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
            Index           =   14
            Left            =   240
            TabIndex        =   51
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Monto .:"
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
            Index           =   3
            Left            =   240
            TabIndex        =   45
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Mov.:"
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
            Left            =   4920
            TabIndex        =   44
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. Doc.:"
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
            Left            =   4920
            TabIndex        =   42
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha .:"
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
            Index           =   7
            Left            =   240
            TabIndex        =   41
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Estado.:"
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
            Index           =   8
            Left            =   4920
            TabIndex        =   40
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Base .:"
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
            Index           =   9
            Left            =   240
            TabIndex        =   39
            Top             =   120
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox gbRegistro 
         Height          =   2175
         Left            =   0
         TabIndex        =   17
         Top             =   2040
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   3836
         _StockProps     =   79
         Caption         =   "Registro: "
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.CheckBox chkMarcas 
            Height          =   255
            Left            =   9240
            TabIndex        =   18
            Top             =   1080
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Marcar"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnRegistro_Registrar 
            Height          =   495
            Left            =   9240
            TabIndex        =   19
            Top             =   1560
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Registar"
            BackColor       =   -2147483633
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Bancos_Cargado.frx":0D1C
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   315
            Left            =   1320
            TabIndex        =   20
            Top             =   960
            Width           =   1935
            _Version        =   1441793
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   315
            Left            =   3240
            TabIndex        =   21
            Top             =   960
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
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
         Begin XtremeSuiteControls.FlatEdit txtConceptoDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   23
            Top             =   600
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtConcepto 
            Height          =   330
            Left            =   1320
            TabIndex        =   24
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   600
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   26
            Top             =   1320
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtUnidad 
            Height          =   330
            Left            =   1320
            TabIndex        =   27
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1320
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.FlatEdit txtCentroDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   28
            Top             =   1680
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCentro 
            Height          =   330
            Left            =   1320
            TabIndex        =   29
            Top             =   1680
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
         Begin XtremeSuiteControls.FlatEdit txtAutoDesc 
            Height          =   330
            Left            =   3240
            TabIndex        =   49
            Top             =   240
            Width           =   5775
            _Version        =   1441793
            _ExtentX        =   10186
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAutoId 
            Height          =   330
            Left            =   1320
            TabIndex        =   50
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   240
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnExport 
            Height          =   495
            Left            =   10440
            TabIndex        =   53
            Top             =   1560
            Width           =   1215
            _Version        =   1441793
            _ExtentX        =   2143
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Exportar"
            BackColor       =   -2147483633
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Bancos_Cargado.frx":1443
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.CheckBox chkInd_Control_Depositos 
            Height          =   375
            Left            =   9240
            TabIndex        =   55
            Top             =   240
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Control de Depósito"
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
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.CheckBox chkInd_Ignora_Registro 
            Height          =   375
            Left            =   9240
            TabIndex        =   56
            Top             =   600
            Width           =   2655
            _Version        =   1441793
            _ExtentX        =   4683
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Ignora Registro en Bancos"
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
            Enabled         =   0   'False
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Auto-Reg."
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
            Index           =   13
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Centro Costo"
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
            Index           =   12
            Left            =   240
            TabIndex        =   31
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad"
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
            Index           =   6
            Left            =   240
            TabIndex        =   30
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Concepto"
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
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta"
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
            Index           =   10
            Left            =   240
            TabIndex        =   22
            Top             =   960
            Width           =   975
         End
      End
      Begin XtremeSuiteControls.GroupBox fraAccion 
         Height          =   975
         Left            =   -70000
         TabIndex        =   10
         Top             =   6480
         Visible         =   0   'False
         Width           =   11895
         _Version        =   1441793
         _ExtentX        =   20981
         _ExtentY        =   1720
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnAplicar 
            Height          =   495
            Left            =   9840
            TabIndex        =   11
            Top             =   240
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Aplicar"
            BackColor       =   -2147483633
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Bancos_Cargado.frx":1D14
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.PushButton btnCancelar 
            Height          =   495
            Left            =   8400
            TabIndex        =   12
            Top             =   240
            Width           =   1455
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cancelar"
            BackColor       =   -2147483633
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmTES_Bancos_Cargado.frx":243B
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   960
            TabIndex        =   13
            Top             =   480
            Width           =   2415
            _Version        =   1441793
            _ExtentX        =   4260
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
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCasos 
            Height          =   315
            Left            =   3360
            TabIndex        =   14
            Top             =   480
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1714
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Casos"
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
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Totales"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   855
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5055
         Left            =   -69880
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   11655
         _Version        =   524288
         _ExtentX        =   20558
         _ExtentY        =   8916
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
         MaxCols         =   5
         SpreadDesigner  =   "frmTES_Bancos_Cargado.frx":29DF
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridId 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   4320
         Width           =   11775
         _Version        =   524288
         _ExtentX        =   20770
         _ExtentY        =   5318
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
         MaxCols         =   14
         SpreadDesigner  =   "frmTES_Bancos_Cargado.frx":302B
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   672
         Left            =   -68680
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   8772
         _Version        =   1441793
         _ExtentX        =   15473
         _ExtentY        =   1185
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
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   372
         Index           =   0
         Left            =   -59680
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bancos_Cargado.frx":3A66
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   372
         Index           =   1
         Left            =   -59200
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bancos_Cargado.frx":4166
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   372
         Index           =   2
         Left            =   -58720
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_Bancos_Cargado.frx":487F
      End
      Begin VB.Label Label1 
         Caption         =   "Archivo"
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
         Index           =   2
         Left            =   -69760
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      DataField       =   "Banco"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_Bancos_Cargado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim mBanco As Long, vPaso As Boolean
Dim vHeaders As vGridHeaders, vTitulo As String

Private Sub sbLimpia()

    vGrid.MaxRows = 0
    vGridId.MaxRows = 0
   
    txtMonto.Text = 0
    txtCasos.Text = 0
    txtArchivo.Text = ""

End Sub



Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen registros cargados...[verifique!]", vbExclamation
       Exit Sub
    End If
   
    Call sbProcesar
End Sub

Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String
  
Select Case Index
  
  Case 0 'buscar
        txtArchivo.Text = ""
        Call sbArchivoBusca

  Case 1 'cargar
       Call sbArchivoCarga


  Case 2 'info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Import" & vbCrLf _
              & " 3. Columnas.: FECHA, TIPO, DOCUMENTO, IMPORTE, DESCRIPCION, SALDO"
     
     MsgBox vMensaje, vbInformation
     
     
End Select
End Sub

Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub btnExport_Click()
'Variables del Exporte
vHeaders.Columnas = vGridId.MaxCols
vTitulo = "ProGrX_TES_Bancos_Cargado_Result"
    
    vHeaders.Headers(2) = "Id Tramite"
    vHeaders.Headers(3) = "Estado"
    vHeaders.Headers(4) = "Documento"
    vHeaders.Headers(5) = "Fecha"
    vHeaders.Headers(6) = "Importe"
    vHeaders.Headers(7) = "Descripción"
    vHeaders.Headers(8) = "Registro Fecha"
    vHeaders.Headers(9) = "Registro Usuario"
    vHeaders.Headers(10) = "Reg-Bancos Fecha"
    vHeaders.Headers(11) = "Reg-Bancos Usuario"
    vHeaders.Headers(12) = "Id Tramite Bancos"
    vHeaders.Headers(13) = "Auto-Registro Id"
    vHeaders.Headers(14) = "DP Tramite Id"
   Call sbSIFGridExportar(vGridId, vHeaders, vTitulo)


End Sub


Private Sub btnRegistro_Buscar_Click()
    Call sbRegistroBuscar
End Sub


Private Sub btnRegistro_Registrar_Click()
    
If vPaso Then Exit Sub
If vGridId.MaxRows = 0 Then Exit Sub
    
    Select Case cboEstado.Text
       Case "Tramite"
            Call sbRegistroAplicar
       Case "Registrados"
            MsgBox "Los casos actuales ya fueron procesados!", vbInformation
    End Select
   
End Sub

Private Sub btnReglas_Review_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spTes_Bancos_Mov_Reglas_Update " & cboBanco.ItemData(cboBanco.ListIndex) _
        & " , '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Se actualizaron las reglas aplicables a los movimientos aun en trámite!", vbInformation

Call btnRegistro_Buscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboBanco_Click()

If vPaso Then Exit Sub

 Call sbLimpia
 
If cboBanco.ListCount = 0 Then
 mBanco = 0
Else
 mBanco = cboBanco.ItemData(cboBanco.ListIndex)
End If

End Sub




Private Sub cboFechas_Click()
If vPaso Then Exit Sub
vGridId.MaxRows = 0
End Sub

Private Sub cboEstado_Click()
If vPaso Then Exit Sub
vGridId.MaxRows = 0
End Sub

Private Sub sbBloqueoDesBloque(pTipo As String)

If pTipo = "B" Then
    txtConcepto.Enabled = False
    txtUnidad.Enabled = False
    txtCuenta.Enabled = False
    txtCentro.Enabled = False
  
    txtConcepto.BackColor = RGB(235, 245, 251)
    txtUnidad.BackColor = RGB(235, 245, 251)
    txtCuenta.BackColor = RGB(235, 245, 251)
    txtCentro.BackColor = RGB(235, 245, 251)
End If

If pTipo = "D" Then
    txtConcepto.Enabled = True
    txtUnidad.Enabled = True
    txtCuenta.Enabled = True
    txtCentro.Enabled = True
    
    txtConcepto.BackColor = vbWhite
    txtUnidad.BackColor = vbWhite
    txtCuenta.BackColor = vbWhite
    txtCentro.BackColor = vbWhite
End If


End Sub

Private Sub cboTipo_Click()
If vPaso Then Exit Sub
vGridId.MaxRows = 0



Select Case cboTipo.ItemData(cboTipo.ListIndex)
  Case "Aut-Cre", "Auto-Deb"
    Call sbBloqueoDesBloque("B")
  
  Case "Déb"
    Call sbBloqueoDesBloque("D")
  
  Case "Cré"
    Call sbBloqueoDesBloque("D")

End Select

End Sub

Private Sub chkMarcas_Click()
Dim i As Long


For i = 1 To vGridId.MaxRows
   vGridId.Row = i
   vGridId.col = 1
   vGridId.Value = chkMarcas.Value
Next i


End Sub

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim vProceso As Long

vModulo = 9
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle

vPaso = True

cboTipo.AddItem "Débitos"
cboTipo.ItemData(cboTipo.ListCount - 1) = "Déb"
cboTipo.AddItem "Créditos"
cboTipo.ItemData(cboTipo.ListCount - 1) = "Cré"
cboTipo.AddItem "Auto-Registro Débitos"
cboTipo.ItemData(cboTipo.ListCount - 1) = "Aut-Deb"
cboTipo.AddItem "Auto-Registro Crébitos"
cboTipo.ItemData(cboTipo.ListCount - 1) = "Aut-Cre"
cboTipo.Text = "Débitos"

cboFechas.AddItem "Documento"
cboFechas.AddItem "Registro"
cboFechas.Text = "Documento"

cboEstado.AddItem "Tramite"
cboEstado.AddItem "Registrados"
cboEstado.Text = "Tramite"

strSQL = "exec spTes_Cuenta_Bancaria_Acceso '" & glogon.Usuario & "','DP','SOL'"

Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

vPaso = False

dtpRegistroCorte.Value = fxFechaServidor
dtpRegistroInicio.Value = DateAdd("m", -1, dtpRegistroCorte.Value)

tcMain.Item(0).Selected = True

Call cboBanco_Click
Call cboTipo_Click
Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbArchivoCarga()
Dim rsExcel As New ADODB.Recordset

Dim i As Integer, iCampos As Integer, vExiste As Integer
Dim vFecha As Date, vDocumento As String, vImporte As Currency, vDescripcion As String, vTipo As String

Dim vCedula As String, vNombre As String, vInconsistencia As String

Dim curMonto As Currency, lCasos As Long

On Error GoTo vError
vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboBanco.ListCount <= 0 Then
    MsgBox "No existe ninguna Institución, no se puede procesar el archivo...", vbCritical
    Exit Sub
End If

Me.MousePointer = vbHourglass

vGrid.MaxRows = 0

curMonto = 0
lCasos = 0 'Total

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")

'Verifica Estructura del Archivo

iCampos = 0
For i = 0 To rsExcel.Fields.Count - 1
   Select Case UCase(rsExcel.Fields(i).Name)
      Case "DOCUMENTO", "FECHA", "IMPORTE", "DESCRIPCION", "TIPO", "SALDO"
        iCampos = iCampos + 1
      Case Else
      
   End Select
Next i

If iCampos < 6 Then
   Me.MousePointer = vbDefault
   MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "2. Los campos son Fecha, Tipo, Documento, Importe, Descripcion y Saldo", vbExclamation
   Exit Sub
End If


With vGrid
    .MaxRows = 0
    
    
    Do While Not rsExcel.EOF
         vDocumento = Trim(rsExcel!Documento & "")
         vFecha = rsExcel!fecha
         vImporte = rsExcel!Importe
         vDescripcion = rsExcel!Descripcion & ""
         vTipo = rsExcel!Tipo
         
      
         If vImporte < 0 Then
            vTipo = "D"
         End If
       
       
      If vDocumento <> "" Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            
            .col = 1
            .Text = vDocumento
            .col = 2
            .Text = Format(vImporte, "Standard")
            .col = 3
            .Text = vTipo
            .col = 4
            .Text = Format(vFecha, "yyyy-mm-dd")
            .col = 5
            .Text = vDescripcion
            
            curMonto = curMonto + vImporte
            txtCasos.Text = txtCasos.Text + 1
       
       End If
       rsExcel.MoveNext
    Loop
    rsExcel.Close
    
End With
        
'Totales
txtMonto.Text = Format(curMonto, "Standard")
Me.MousePointer = vbDefault

MsgBox "Información Pre-Cargada Satisfactoriamente", vbInformation


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia
End Sub

Private Sub sbProcesar()
Dim strSQL As String

Dim i As Long

Dim vDescripcion As String, vCuenta As String
Dim vDocumento As String, vImporte As Currency, vTipo As String, vFecha As Date

Dim vMensaje As Boolean, vCasos As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

vMensaje = False
vCasos = 0

'Inicializa
strSQL = ""
With vGrid
    For i = 1 To .MaxRows
       .Row = i
       .col = 1
       vDocumento = .Text
       .col = 2
       vImporte = CCur(.Text)
       .col = 3
       vTipo = .Text
       .col = 4
       vFecha = Format(.Text, "yyyy/mm/dd")
       .col = 5
       vDescripcion = .Text
       
        strSQL = strSQL & Space(10) & "exec spTes_Bancos_Mov_Load " & cboBanco.ItemData(cboBanco.ListIndex) & ",'" & Format(vFecha, "yyyy/mm/dd") _
               & "','" & Mid(vDocumento, 1, 30) & "','" & Mid(vTipo, 1, 1) & "'," & Abs(vImporte) & ",'" & Mid(vDescripcion, 1, 150) _
               & "',0,'" & glogon.Usuario & "'"
       
       vCasos = vCasos + 1
       
       If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
       End If
    Next i
End With

If Len(strSQL) > 0 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
End If


'Concilia y Actualiza
'strSQL = "exec spTes_Concilia_Automatica " & cboBanco.ItemData(cboBanco.ListIndex) _
'       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)
'
'strSQL = "exec spTes_Concilia_Periodo_Actualiza " & cboBanco.ItemData(cboBanco.ListIndex) _
'       & "," & feAnio.Text & "," & feMes.Text & ",'" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Carga realizada Satisfactoriamente... Registros Procesados :" & vCasos, vbInformation


Call sbLimpia

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Call sbLimpia

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

tcMain.Height = Me.Height - (tcMain.top + 650)
tcMain.Width = Me.Width - 500


gbBuscar.Width = tcMain.Width
gbRegistro.Width = tcMain.Width


vGrid.Height = tcMain.Height - (vGrid.top + fraAccion.Height + 350)
vGrid.Width = tcMain.Width - 350
fraAccion.top = vGrid.top + vGrid.Height + 100

vGridId.Height = tcMain.Height - (vGridId.top + 200)
vGridId.Width = tcMain.Width - 350

fraAccion.Width = tcMain.Width

End Sub



Private Sub sbRegistroBuscar()

Dim i As Integer

On Error GoTo vError

If cboBanco.ListCount = 0 Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

        
txtDescripcion.Text = fxSysCleanTxtInject(txtDescripcion.Text)
txtNumDoc.Text = fxSysCleanTxtInject(txtNumDoc.Text)

       
strSQL = "exec spTes_Bancos_Mov_Consulta " & cboBanco.ItemData(cboBanco.ListIndex) & ", '" & txtNumDoc.Text _
       & "', '" & cboTipo.ItemData(cboTipo.ListIndex) & "', '" & Mid(cboFechas.Text, 1, 3) & "', '" _
       & Format(dtpRegistroInicio.Value, "yyyy/mm/dd") & " 00:00:00', '" _
       & Format(dtpRegistroCorte.Value, "yyyy/mm/dd") & " 23:59:59', " _
       & CCur(txtMntInicio.Text) & ", " & CCur(txtMntCorte.Text) & ", '" & Mid(cboEstado.Text, 1, 1) _
       & "', '" & txtDescripcion.Text & "'"

Call OpenRecordSet(rs, strSQL)

vGridId.MaxRows = 0


  Do While Not rs.EOF
    vGridId.MaxRows = vGridId.MaxRows + 1
    vGridId.Row = vGridId.MaxRows
         
    vGridId.col = 1

    For i = 2 To vGridId.MaxCols
      vGridId.col = i
      Select Case i
         Case 2 'Tramite Id
            vGridId.Text = CStr(rs!Id_Linea)
         Case 3 'Estado
            vGridId.Text = CStr(rs!Estado_Desc)
            
         Case 4 'Documento
            vGridId.Text = rs!Documento
         Case 5 'Fecha del Documento
            vGridId.Text = Format(rs!fecha, "yyyy-mm-dd")
         Case 6 'Importe
            vGridId.Text = Format(rs!Importe, "Standard")
         Case 7 'Descripcion
            vGridId.Text = rs!Descripcion
         Case 8 'Registro Fecha
            vGridId.Text = rs!Registro_Fecha & ""
         Case 9 'Registro Usuario
            vGridId.Text = rs!Registro_Usuario & ""
            
         Case 10 'Procesado Fecha
            vGridId.Text = rs!PROCESADO_FECHA & ""
         Case 11 'Procesado Usuario
            vGridId.Text = rs!PROCESADO_USUARIO & ""
            
         Case 12 'Tes. Registro Id
            vGridId.Text = CStr(rs!CONCILIA_NSOLICITUD & "")
         Case 13 'Auto Registro Regla Id
            vGridId.Text = CStr(rs!AUTO_REGISTRO_ID & "")
         Case 14 'DP Tramite Id
            vGridId.Text = CStr(rs!DP_TRAMITE_ID & "")
      
      End Select
    Next i
     rs.MoveNext
   Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbRegistroAplicar()
Dim i As Long, pAutoId As Long

On Error GoTo vError


Select Case cboTipo.ItemData(cboTipo.ListIndex)
    Case "Aut-Deb", "Aut-Cre"
        'Auto Registros no Requieren Validacion
    Case Else
        If txtConcepto.Text = "" Or txtUnidad.Text = "" Or Replace(txtCuenta.Text, "-", "") = "" Then
            MsgBox "No se ha indicado un Concepto, Unidad o Cuenta Contable válida para el registro!", vbExclamation
            Exit Sub
        End If
End Select
Me.MousePointer = vbHourglass

With vGridId

strSQL = ""
For i = 1 To .MaxRows
  .Row = i
  .col = 13
  pAutoId = .Text
  
  
  .col = 1
  
  If .Value = vbChecked Then
    .col = 2
    strSQL = strSQL & Space(10) & "exec spTes_Bancos_Mov_Registro " & .Text & ", '" & glogon.Usuario & "', " & pAutoId _
           & ", '" & txtConcepto.Text & "', '" & txtUnidad.Text & "', '" & txtCentro.Text & "', '" _
           & fxgCntCuentaFormato(False, txtCuenta.Text, 0) & "'"
  End If

  If Len(strSQL) > 20000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If
  
Next i
End With

  If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
  End If

Me.MousePointer = vbDefault

MsgBox "Casos registrados en Banking satisfactoriamente!", vbInformation

Call sbRegistroBuscar


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call sbRegistroBuscar
  
End Sub


Private Sub sbArchivoBusca()


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Depósitos del Banco [Microsoft EXCEL]"
    .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
    .ShowOpen

    If .FileName = "" Then
        MsgBox "Archivo no válido...", vbExclamation
        Exit Sub
    End If

    If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
        'Ok
    Else
        MsgBox "La Extensión del Archivo no es válido...", vbExclamation
        Exit Sub
    End If

    
    txtArchivo.Text = .FileName
End With

End Sub


'------------------

Private Sub sbCodigoDescripcion(pTipo As String, pCodigo As String)

On Error GoTo vError

Dim txt As XtremeSuiteControls.FlatEdit

Select Case pTipo
  Case "Cta"
  Case "Con"
    strSQL = "select DESCRIPCION as 'ItmX' from vTes_Conceptos Where cod_concepto = '" & pCodigo & "'"
    Set txt = txtConceptoDesc
    
  Case "Ud"
    strSQL = "select DESCRIPCION as 'ItmX' from vCNTX_UNIDADES_LOCAL Where cod_Unidad = '" & pCodigo & "'"
    Set txt = txtUnidadDesc
  Case "Cc"
    strSQL = "select DESCRIPCION as 'ItmX' from vCNTX_CENTRO_COSTO_LOCAL Where cod_Centro_Costo = '" & pCodigo & "'"
    Set txt = txtCentroDesc
End Select


Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txt.Text = rs!itmX
End If
rs.Close


vError:
  
End Sub


Private Sub txtAutoId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "ID_Auto"
  gBusquedas.Orden = "ID_Auto"
  gBusquedas.Consulta = "select ID_Auto ,Descripcion from vTES_AUTO_REGISTRO"
  gBusquedas.Filtro = " and Activo = 1 and APL_TIPO_MOV = '" & Mid(cboTipo.Text, 1, 1) & "'"
  frmBusquedas.Show vbModal
  
  If IsNumeric(gBusquedas.Resultado) Then
    txtAutoId.Text = gBusquedas.Resultado
    txtAutoDesc.Text = gBusquedas.Resultado2
    
  Else
    txtAutoId.Text = ""
    txtAutoDesc.Text = ""
  
  End If

  txtAutoDesc.SetFocus

End If
End Sub

Private Sub txtAutoId_LostFocus()

If txtAutoId.Text = "" Then
    Call sbBloqueoDesBloque("D")
    
    txtAutoDesc.Text = ""
    txtConceptoDesc.Text = ""
    txtUnidadDesc.Text = ""
    txtCuentaDesc.Text = ""
    txtCentroDesc.Text = ""
    
    txtConcepto.Text = ""
    txtUnidad.Text = ""
    txtCuenta.Text = ""
    txtCentro.Text = ""
    
Else
    Call sbBloqueoDesBloque("B")
    
    strSQL = "select * from vTES_AUTO_REGISTRO Where Id_Auto = " & txtAutoId.Text
    Call OpenRecordSet(rs, strSQL)
    txtAutoDesc.Text = rs!Descripcion
    
    txtConceptoDesc.Text = rs!Concepto_Desc
    txtUnidadDesc.Text = rs!Unidad_Desc
    txtCuentaDesc.Text = rs!Cuenta_Desc
    txtCentroDesc.Text = rs!Centro_Desc
    
    txtConcepto.Text = rs!cod_Concepto
    txtUnidad.Text = rs!Cod_Unidad
    txtCuenta.Text = rs!Cod_Cuenta_Mask
    txtCentro.Text = rs!Cod_Centro_Costo
    
    chkInd_Ignora_Registro.Value = rs!IGNORA_REGISTRO_ID
    chkInd_Control_Depositos.Value = rs!DP_TRAMITE
    
    rs.Close

End If

End Sub

Private Sub txtCentro_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentroDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Centro"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = ""
  gBusquedas.Consulta = "select COD_CENTRO_COSTO, DESCRIPCION from vCNTX_CENTRO_COSTO_LOCAL"
  gBusquedas.Filtro = ""
  gBusquedas.Columna = "COD_CENTRO_COSTO"
  gBusquedas.Orden = "DESCRIPCION"
  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
     txtCentro.Text = gBusquedas.Resultado
     txtCentroDesc.Text = gBusquedas.Resultado2
  End If
  
End If

End Sub

Private Sub txtCentro_LostFocus()
Call sbCodigoDescripcion("Cc", txtCentro.Text)

End Sub


Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConceptoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Concepto"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = "Cuenta"
  gBusquedas.Consulta = "select COD_CONCEPTO, DESCRIPCION, COD_CUENTA_MASK from vTes_Conceptos"
  gBusquedas.Filtro = ""
  gBusquedas.Columna = "COD_CONCEPTO"
  gBusquedas.Orden = "DESCRIPCION"
  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
     txtConcepto.Text = gBusquedas.Resultado
     txtConceptoDesc.Text = gBusquedas.Resultado2
     
     If gBusquedas.Resultado3 <> "" Then
       txtCuenta.Text = gBusquedas.Resultado3
       Call txtCuenta_LostFocus
     End If
  End If
  
End If

End Sub

Private Sub txtConcepto_LostFocus()

Call sbCodigoDescripcion("Con", txtConcepto.Text)

End Sub


Private Sub txtConceptoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidad.SetFocus
End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuenta.Text = gCuenta
   txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta, 0)
End If

End Sub

Private Sub txtCuenta_LostFocus()
   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuenta, 0))
   txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta, 0)
End Sub


Private Sub txtMntInicio_GotFocus()
On Error GoTo vError
  
  txtMntInicio.Text = CCur(txtMntInicio.Text)

vError:
End Sub

Private Sub txtMntInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMntCorte.SetFocus

End Sub

Private Sub txtMntInicio_LostFocus()
On Error GoTo vError
  
  txtMntInicio.Text = Format(CCur(txtMntInicio.Text), "Standard")

vError:

End Sub

Private Sub txtMntCorte_GotFocus()
On Error GoTo vError
  
  txtMntCorte.Text = CCur(txtMntCorte.Text)

vError:
End Sub

Private Sub txtMntCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNumDoc.SetFocus

End Sub

Private Sub txtMntCorte_LostFocus()
On Error GoTo vError
  
  txtMntCorte.Text = Format(CCur(txtMntCorte.Text), "Standard")

vError:

End Sub


Private Sub txtUnidad_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Unidad"
  gBusquedas.Col2Name = "Descripción"
  gBusquedas.Col3Name = ""
  gBusquedas.Consulta = "select COD_UNIDAD, DESCRIPCION from vCNTX_UNIDADES_LOCAL"
  gBusquedas.Filtro = ""
  gBusquedas.Columna = "COD_UNIDAD"
  gBusquedas.Orden = "DESCRIPCION"
  frmBusquedas.Show vbModal
  
  If gBusquedas.Resultado <> "" Then
     txtUnidad.Text = gBusquedas.Resultado
     txtUnidadDesc.Text = gBusquedas.Resultado2
  End If
  
End If


End Sub

Private Sub txtUnidad_LostFocus()

Call sbCodigoDescripcion("Ud", txtUnidad.Text)

End Sub

Private Sub txtUnidadDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCentro.SetFocus
End Sub






