VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAH_ExcedentesPeriodos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excedentes: Periodos"
   ClientHeight    =   7770
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6615
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   11668
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
      ItemCount       =   5
      SelectedItem    =   1
      Item(0).Caption =   "Periodos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   8
      Item(1).Control(0)=   "Label1(2)"
      Item(1).Control(1)=   "dtpInicio"
      Item(1).Control(2)=   "Label1(3)"
      Item(1).Control(3)=   "dtpCorte"
      Item(1).Control(4)=   "GroupBox1"
      Item(1).Control(5)=   "lswRenta"
      Item(1).Control(6)=   "Label1(10)"
      Item(1).Control(7)=   "btnRecalcular"
      Item(2).Caption =   "Estado Excedentes"
      Item(2).ControlCount=   7
      Item(2).Control(0)=   "chkVisible_Web"
      Item(2).Control(1)=   "chkVisible_Sys"
      Item(2).Control(2)=   "Label2"
      Item(2).Control(3)=   "txtEstado_Nota"
      Item(2).Control(4)=   "chkVisible_Historial"
      Item(2).Control(5)=   "chkVisible_Renta_Tabla"
      Item(2).Control(6)=   "btnEstadoNota_Update"
      Item(3).Caption =   "Bitácora"
      Item(3).ControlCount=   5
      Item(3).Control(0)=   "lswBitacora"
      Item(3).Control(1)=   "rbBitacora(0)"
      Item(3).Control(2)=   "rbBitacora(2)"
      Item(3).Control(3)=   "rbBitacora(1)"
      Item(3).Control(4)=   "rbBitacora(3)"
      Item(4).Caption =   "Resumen"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "lswResumen"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6135
         Left            =   -69880
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   10815
         _Version        =   1441793
         _ExtentX        =   19076
         _ExtentY        =   10821
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lswRenta 
         Height          =   2175
         Left            =   5040
         TabIndex        =   24
         Top             =   720
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
         _ExtentY        =   3836
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ListView lswResumen 
         Height          =   6255
         Left            =   -69880
         TabIndex        =   42
         Top             =   360
         Visible         =   0   'False
         Width           =   10815
         _Version        =   1441793
         _ExtentX        =   19076
         _ExtentY        =   11033
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
      Begin XtremeSuiteControls.ListView lswBitacora 
         Height          =   5775
         Left            =   -69880
         TabIndex        =   43
         Top             =   840
         Visible         =   0   'False
         Width           =   10815
         _Version        =   1441793
         _ExtentX        =   19076
         _ExtentY        =   10186
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
      Begin XtremeSuiteControls.PushButton btnRecalcular 
         Height          =   495
         Left            =   1200
         TabIndex        =   35
         Top             =   2160
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Recalcular la Base"
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
         Picture         =   "frmAH_ExcedentesPeriodos.frx":0000
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3735
         Left            =   720
         TabIndex        =   9
         Top             =   3120
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   6588
         _StockProps     =   79
         Caption         =   "Información de aplicación:"
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
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnBoleta 
            Height          =   315
            Index           =   0
            Left            =   6480
            TabIndex        =   28
            Top             =   360
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkRentaCap 
            Height          =   255
            Left            =   5160
            TabIndex        =   22
            Top             =   3120
            Width           =   3255
            _Version        =   1441793
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Renta Incluye Capitalización?    "
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
         End
         Begin XtremeSuiteControls.FlatEdit txtNC_Mora 
            Height          =   315
            Left            =   4200
            TabIndex        =   11
            Top             =   360
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNC_OPCF 
            Height          =   315
            Left            =   4200
            TabIndex        =   13
            Top             =   840
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNC_CEXD 
            Height          =   315
            Left            =   4200
            TabIndex        =   15
            Top             =   1320
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNC_CAPIND 
            Height          =   315
            Left            =   4200
            TabIndex        =   17
            Top             =   1800
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDocExc 
            Height          =   315
            Left            =   4200
            TabIndex        =   19
            Top             =   2280
            Width           =   2172
            _Version        =   1441793
            _ExtentX        =   3831
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCapPorc 
            Height          =   312
            Left            =   4200
            TabIndex        =   21
            Top             =   3120
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnBoleta 
            Height          =   312
            Index           =   1
            Left            =   6480
            TabIndex        =   29
            Top             =   840
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnBoleta 
            Height          =   312
            Index           =   2
            Left            =   6480
            TabIndex        =   30
            Top             =   1320
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnBoleta 
            Height          =   312
            Index           =   3
            Left            =   6480
            TabIndex        =   31
            Top             =   1800
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnBoleta 
            Height          =   312
            Index           =   4
            Left            =   6480
            TabIndex        =   32
            Top             =   2280
            Width           =   372
            _Version        =   1441793
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtReservaPorc 
            Height          =   312
            Left            =   4200
            TabIndex        =   33
            Top             =   2760
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboBaseAplicacion 
            Height          =   330
            Left            =   7800
            TabIndex        =   39
            Top             =   840
            Width           =   2055
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
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnBaseAplicacionUpdate 
            Height          =   375
            Left            =   7800
            TabIndex        =   41
            Top             =   1200
            Width           =   2055
            _Version        =   1441793
            _ExtentX        =   3625
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
            Appearance      =   17
            Picture         =   "frmAH_ExcedentesPeriodos.frx":0719
            ImageAlignment  =   0
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de cálculo para aplicación mensual: "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   12
            Left            =   7320
            TabIndex        =   40
            Top             =   360
            Width           =   2535
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje de Reserva"
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
            Index           =   11
            Left            =   0
            TabIndex        =   34
            Top             =   2760
            Width           =   3972
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje de Capitalización"
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
            Index           =   9
            Left            =   0
            TabIndex        =   20
            Top             =   3120
            Width           =   3972
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No Documento: Asiento General de Excedentes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   0
            TabIndex        =   18
            Top             =   2280
            Width           =   3975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Crédito: Aplicación Capitalización Extraordinaria"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   0
            TabIndex        =   16
            Top             =   1800
            Width           =   3975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Crédito: Aplicación de Crédito en Excedentes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   0
            TabIndex        =   14
            Top             =   1320
            Width           =   3975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Crédito: Aplicación de OPCF"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   0
            TabIndex        =   12
            Top             =   840
            Width           =   3975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Crédito: Aplicación de Morosidad"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   0
            TabIndex        =   10
            Top             =   360
            Width           =   3975
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   840
         Width           =   1335
         _Version        =   1441793
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
         _Version        =   1441793
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
      Begin XtremeSuiteControls.RadioButton rbBitacora 
         Height          =   255
         Index           =   0
         Left            =   -69880
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Todos"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbBitacora 
         Height          =   255
         Index           =   2
         Left            =   -65440
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proceso de Cierre y aplicaciones"
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
      End
      Begin XtremeSuiteControls.RadioButton rbBitacora 
         Height          =   255
         Index           =   1
         Left            =   -68440
         TabIndex        =   46
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Aplicaciones mensuales"
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
      End
      Begin XtremeSuiteControls.RadioButton rbBitacora 
         Height          =   255
         Index           =   3
         Left            =   -62560
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cambios de Configuración"
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
      End
      Begin XtremeSuiteControls.CheckBox chkVisible_Web 
         Height          =   255
         Left            =   -69640
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   4695
         _Version        =   1441793
         _ExtentX        =   8281
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visible en Web/App AutoGestión de Asociados"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkVisible_Sys 
         Height          =   255
         Left            =   -69640
         TabIndex        =   49
         Top             =   1080
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6588
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visible en Sistema (Consulta Integrada)"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkVisible_Historial 
         Height          =   255
         Left            =   -69640
         TabIndex        =   50
         Top             =   1680
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visible en la sección de Historial del Estado de Excedentes"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado_Nota 
         Height          =   1215
         Left            =   -69640
         TabIndex        =   52
         Top             =   3000
         Visible         =   0   'False
         Width           =   10095
         _Version        =   1441793
         _ExtentX        =   17806
         _ExtentY        =   2143
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkVisible_Renta_Tabla 
         Height          =   255
         Left            =   -69640
         TabIndex        =   53
         Top             =   2040
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1441793
         _ExtentX        =   9340
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Visible en la sección de Tabla Renta sobre Excedentes del Periodo"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.PushButton btnEstadoNota_Update 
         Height          =   375
         Left            =   -61600
         TabIndex        =   54
         Top             =   4320
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Actualiza Nota"
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
         Picture         =   "frmAH_ExcedentesPeriodos.frx":0E40
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Left            =   -69640
         TabIndex        =   51
         Top             =   2640
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1441793
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nota para el Estado de Excedentes de este Periodo:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla de Renta del Periodo: "
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
         Index           =   10
         Left            =   5400
         TabIndex        =   25
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
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
         Index           =   3
         Left            =   960
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
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
         Index           =   2
         Left            =   960
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   372
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   2640
      TabIndex        =   23
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtEstado 
      Height          =   372
      Left            =   4080
      TabIndex        =   26
      Top             =   480
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnTablaAplicacionMensual 
      Height          =   375
      Left            =   7800
      TabIndex        =   36
      Top             =   480
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Tabla para Aplicación Mensual"
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
      Picture         =   "frmAH_ExcedentesPeriodos.frx":1571
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   6360
      TabIndex        =   37
      Top             =   480
      Width           =   1455
      _Version        =   1441793
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Picture         =   "frmAH_ExcedentesPeriodos.frx":1C79
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   38
      Top             =   960
      Visible         =   0   'False
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   238
      _StockProps     =   93
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   2760
      TabIndex        =   27
      Top             =   480
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo ID"
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
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1092
   End
End
Attribute VB_Name = "frmAH_ExcedentesPeriodos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long
Dim vScroll As Boolean, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnBaseAplicacionUpdate_Click()

On Error GoTo vError

If txtEstado.Text = "Cerrado" Then
    MsgBox " - El Periodo ya fue cerrado!", vbExclamation
    Exit Sub
End If

strSQL = "exec spExc_Periodo_Modo_Aplicacion " & vCodigo & ", '" & cboBaseAplicacion.ItemData(cboBaseAplicacion.ListIndex) & "', '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Base para distribución guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBoleta_Click(Index As Integer)
Dim vTipoDoc As String, vNumDoc As String

Select Case Index
 Case 0 'MORA
    vTipoDoc = "NC"
    vNumDoc = txtNC_Mora.Text
 
 Case 1 'OPCF
    vTipoDoc = "NC"
    vNumDoc = txtNC_OPCF.Text
 
 Case 2 'Credito de Excedente (Abono)
    vTipoDoc = "NC"
    vNumDoc = txtNC_CEXD.Text
 
 Case 3 'Capitalizacion Extraordinaria
    vTipoDoc = "FND"
    vNumDoc = txtNC_CAPIND.Text
 
 Case 4 'Asiento de Excedente
    vTipoDoc = "PLA"
    vNumDoc = txtDocExc.Text
 
 Case 1 'OPCF
    vTipoDoc = "NC"
    vNumDoc = txtDocExc.Text
 
End Select

Me.MousePointer = vbHourglass
 Call sbImprimeRecibo(vNumDoc, vTipoDoc)
Me.MousePointer = vbDefault
End Sub

Private Sub btnEstadoNota_Update_Click()

On Error GoTo vError

txtEstado_Nota.Text = fxSysCleanTxtInject(txtEstado_Nota.Text)

strSQL = "update EXC_PERIODOS set ESTADO_NOTAS = '" & txtEstado_Nota.Text _
       & "' where ID_PERIODO = " & vCodigo
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Excedentes> Periodo:  " & vCodigo & " -> Nota del Estado de Excedentes")

MsgBox "Nota del Estado de Excedentes, guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExport_Click()

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
    Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Select Case tcMain.SelectedItem
   Case 0 'Periodos
        Call Excel_Exportar_Lsw(lsw, ProgressBarX)
   Case 1 'Renta
        Call Excel_Exportar_Lsw(lswRenta, ProgressBarX)
   Case 3 'Bitácora
        Call Excel_Exportar_Lsw(lswBitacora, ProgressBarX)
   Case 4 'Resumen
        Call Excel_Exportar_Lsw(lswResumen, ProgressBarX)
End Select

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRecalcular_Click()

Dim i As Integer


On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
    Exit Sub
End If

i = MsgBox("Esta seguro que desea Recalcular la Base de Datos para Cálculo de Excedentes?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If


i = 0

Me.MousePointer = vbHourglass

strSQL = "select dbo.fxSys_FechaAnioMesToDatetime(H.ANIO, H.MES) as 'CORTE'" _
       & "  from EXC_PERIODOS E" _
       & "            inner join ASE_PER_HISTORICO H on dbo.fxSys_FechaAnioMesToDatetime(H.ANIO, H.MES)" _
       & "            Between E.INICIO And E.CORTE" _
       & "     Where ID_PERIODO = " & txtCodigo.Text & " and ESTADO = 'A'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 
 btnRecalcular.Caption = "Recalculando: " & Format(rs!Corte, "yyyy-MM-dd")
 DoEvents
 
 strSQL = "exec spSIFAuxExcedentes_WLog " & Year(rs!Corte) & ", " & Month(rs!Corte)
 Call ConectionExecute(strSQL)
 
 i = 1
 rs.MoveNext
Loop
rs.Close

If i = 1 Then
    Call Bitacora("Aplica", "Excedentes: Recalculo de la Base, Periodo Id: " & txtCodigo.Text)
End If

btnRecalcular.Caption = "Recalcular la Base"

Me.MousePointer = vbDefault

MsgBox "Se ha recalculado la base para Cálculo de Excedentes, verifique la carga mensual!", vbExclamation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnTablaAplicacionMensual_Click()
Call sbFormsCall("frmAH_Excedentes_Distribucion")
End Sub

Private Sub chkVisible_Historial_Click()
On Error GoTo vError

If vPaso Then Exit Sub

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
    Exit Sub
End If

'Revisa Permisos
If Not tlb.Buttons.Item(1).Enabled Then Exit Sub

'Si no está cerrado, se hace como edicion normal
If txtEstado.Text <> "Cerrado" Then Exit Sub

Me.MousePointer = vbHourglass

  strSQL = "update EXC_PERIODOS set MOSTRAR_EN_HISTORIAL = " & chkVisible_Historial.Value _
         & " where ID_PERIODO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Excedentes> Periodo:  " & vCodigo & " Visibilidad en Historial: " & IIf(chkVisible_Historial.Value, "Sí", "No"))


Me.MousePointer = vbDefault

MsgBox "Se ha cambiado la opción de visibilidad para la sección de Historial en el Estado de Excedentes", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkVisible_Renta_Tabla_Click()
On Error GoTo vError

If vPaso Then Exit Sub

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
    Exit Sub
End If

'Revisa Permisos
If Not tlb.Buttons.Item(1).Enabled Then Exit Sub

'Si no está cerrado, se hace como edicion normal
If txtEstado.Text <> "Cerrado" Then Exit Sub

Me.MousePointer = vbHourglass

  strSQL = "update EXC_PERIODOS set MOSTRAR_TABLA_RENTA = " & chkVisible_Renta_Tabla.Value _
         & " where ID_PERIODO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Excedentes> Periodo:  " & vCodigo & " Visibilidad en Cálculo Renta: " & IIf(chkVisible_Renta_Tabla.Value, "Sí", "No"))


Me.MousePointer = vbDefault

MsgBox "Se ha cambiado la opción de visibilidad para la sección de Cálculo de Renta en el Estado de Excedentes", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkVisible_Sys_Click()

On Error GoTo vError

If vPaso Then Exit Sub

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
    Exit Sub
End If

'Revisa Permisos
If Not tlb.Buttons.Item(1).Enabled Then Exit Sub

'Si no está cerrado, se hace como edicion normal
If txtEstado.Text <> "Cerrado" Then Exit Sub

Me.MousePointer = vbHourglass

  strSQL = "update EXC_PERIODOS set VISIBLE_SYS = " & chkVisible_Sys.Value _
         & " where ID_PERIODO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Excedentes> Periodo:  " & vCodigo & " Visibilidad en Sistema: " & IIf(chkVisible_Sys.Value, "Sí", "No"))


Me.MousePointer = vbDefault

MsgBox "Se ha cambiado la opción de visibilidad para el Sistema y consulta interna", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkVisible_Web_Click()

On Error GoTo vError

If vPaso Then Exit Sub

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
    Exit Sub
End If

'Revisa Permisos
If Not tlb.Buttons.Item(1).Enabled Then Exit Sub

'Si no está cerrado, se hace como edicion normal
If txtEstado.Text <> "Cerrado" Then Exit Sub

Me.MousePointer = vbHourglass

  strSQL = "update EXC_PERIODOS set VISIBLE_WEBAPP = " & chkVisible_Web.Value _
         & " where ID_PERIODO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Excedentes> Periodo:  " & vCodigo & " Visibilidad en Web: " & IIf(chkVisible_Web.Value, "Sí", "No"))


Me.MousePointer = vbDefault

MsgBox "Se ha cambiado la opción de visibilidad para la Web y App!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub FlatScrollBar_Change()

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then
    txtCodigo.Text = "0"
End If

If vScroll Then
    strSQL = "select Top 1 ID_PERIODO from EXC_PERIODOS"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where ID_PERIODO > " & txtCodigo.Text & " order by ID_PERIODO asc"
    Else
       strSQL = strSQL & " where ID_PERIODO < " & txtCodigo.Text & " order by ID_PERIODO desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!ID_PERIODO
      Call txtCodigo_LostFocus
    End If
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 2
End Sub

Private Sub Form_Load()

On Error GoTo vError

 vModulo = 2
 
 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
 
 With cboBaseAplicacion
    .Clear
    .AddItem "Manual"
    .ItemData(.ListCount - 1) = "M"
    .AddItem "Cargado [%]"
    .ItemData(.ListCount - 1) = "C"
    .AddItem "Real Contable"
    .ItemData(.ListCount - 1) = "R"
    .AddItem "Proyectado"
    .ItemData(.ListCount - 1) = "P"
    .AddItem "Prorrateado"
    .ItemData(.ListCount - 1) = "T"
    .Text = "Manual"
 End With

 
 With lsw.ColumnHeaders
    .Clear
    .Add , , "[ID]", 900
    .Add , , "Inicio", 2500
    .Add , , "Corte", 2500
    .Add , , "Estado", 1500, vbCenter
    
    .Add , , "[%] Reserva", 1500, vbCenter
    .Add , , "[%] Capitaliza", 1500, vbCenter
    .Add , , "Apl. Renta Cap?", 1500, vbCenter
    
    .Add , , "NC.Crédito", 1500, vbRightJustify
    .Add , , "NC.Mora", 1500, vbRightJustify
    .Add , , "NC.OPCF", 1500, vbRightJustify
    
    .Add , , "Visible Web/App", 1500, vbCenter
    .Add , , "Visible Sistema", 1500, vbCenter
    
 End With
 
 
 With lswRenta.ColumnHeaders
    .Clear
    .Add , , "Inicio", 1800, vbRightJustify
    .Add , , "Corte", 1800, vbRightJustify
    .Add , , "[%] Renta", 1100, vbRightJustify
 End With

 With lswBitacora.ColumnHeaders
    .Clear
    .Add , , "[ID]", 1100
    .Add , , "Fecha", 2100
    .Add , , "Usuario", 2500
    .Add , , "Transacción", 2500
    .Add , , "Detalle", 2500
    .Add , , "T.Doc", 900, vbCenter
    .Add , , "N.Doc", 1500, vbCenter
    .Add , , "Casos", 1500, vbRightJustify
    .Add , , "Monto", 2500, vbRightJustify
    .Add , , "Inicio", 2500, vbCenter
    .Add , , "Finaliza", 2500, vbCenter
    .Add , , "Duración (seg)", 2500, vbCenter
    
    
 End With


 With lswResumen.ColumnHeaders
    .Clear
    .Add , , "", 3500
    .Add , , "Aplicado", 2800, vbRightJustify
    .Add , , "Cargado", 2800, vbRightJustify

 End With


vEdita = False
 
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 If btnBaseAplicacionUpdate.Tag = 1 Then
 btnBaseAplicacionUpdate.Enabled = True
 End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
vCodigo = 0
txtCodigo = ""


lsw.ListItems.Clear
lswRenta.ListItems.Clear
lswBitacora.ListItems.Clear

tcMain.Item(1).Selected = True

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

txtReservaPorc.Text = 0
txtCapPorc.Text = 0
txtDocExc.Text = ""

txtNC_CAPIND.Text = ""
txtNC_CEXD.Text = ""
txtNC_Mora.Text = ""
txtNC_OPCF.Text = ""

txtEstado.Text = "Abierto"

chkRentaCap.Value = xtpUnchecked

chkVisible_Web.Value = xtpChecked
chkVisible_Sys.Value = xtpChecked
chkVisible_Historial.Value = xtpChecked
chkVisible_Renta_Tabla.Value = xtpChecked

txtEstado_Nota.Text = ""


End Sub



Private Sub lswPeriodos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Call sbConsulta(Item.Text)

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Call sbConsulta(Item.Text)

End Sub

Private Sub lswBitacora_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim vNumDoc As String, vTipoDoc As String


If Item.SubItems(5) <> "" Then
   vTipoDoc = Item.SubItems(5)
   vNumDoc = Item.SubItems(6)
   Me.MousePointer = vbHourglass
    Call sbImprimeRecibo(vNumDoc, vTipoDoc)
   Me.MousePointer = vbDefault
End If

End Sub



Private Sub rbBitacora_Click(Index As Integer)

Dim pEtapa As String

Select Case True
    Case rbBitacora(0).Value 'Todos
        pEtapa = "T"
    Case rbBitacora(1).Value 'Aplicacion Mensual
        pEtapa = "A"
    Case rbBitacora(2).Value 'Cierre
        pEtapa = "C"
    Case rbBitacora(3).Value 'Configuracion
        pEtapa = "X"
End Select

If pEtapa = "T" Then
    strSQL = "select * from vExc_Periodos_Bitacora" _
           & " where id_Periodo = " & vCodigo _
           & " order by registro_fecha desc"
Else
    strSQL = "select * from vExc_Periodos_Bitacora" _
           & " where id_Periodo = " & vCodigo & " and Etapa = '" & pEtapa & "'" _
           & " order by registro_fecha desc"
End If
Call OpenRecordSet(rs, strSQL)

lswBitacora.ListItems.Clear


Do While Not rs.EOF
  Set itmX = lswBitacora.ListItems.Add(, , rs!Linea)
      itmX.SubItems(1) = rs!Registro_Fecha
      itmX.SubItems(2) = rs!Registro_Usuario
      itmX.SubItems(3) = rs!Proceso_Desc & ""
      itmX.SubItems(4) = rs!Detalle
      itmX.SubItems(5) = rs!Tipo_Documento
      itmX.SubItems(6) = rs!Cod_Transaccion
  
      itmX.SubItems(7) = Format(rs!Casos, "###,##0")
      itmX.SubItems(8) = Format(rs!Monto, "Standard")
      itmX.SubItems(9) = rs!Time_Inicio & ""
      itmX.SubItems(10) = rs!Time_Corte & ""
      itmX.SubItems(11) = rs!Duracion & ""
  
  rs.MoveNext
Loop
rs.Close

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)


On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True


Select Case Item.Index
    Case 0 'Lista de Periodos
        strSQL = "select * from EXC_PERIODOS order by id_Periodo desc"
        Call OpenRecordSet(rs, strSQL)
        
        lsw.ListItems.Clear
        
        Do While Not rs.EOF
          Set itmX = lsw.ListItems.Add(, , rs!ID_PERIODO)
              itmX.SubItems(1) = Format(rs!INICIO, "dd/mm/yyyy")
              itmX.SubItems(2) = Format(rs!Corte, "dd/mm/yyyy")
          
          Select Case rs!ESTADO
            Case "P", "A"
                itmX.SubItems(3) = "Abierto"
            Case "C"
                itmX.SubItems(3) = "Cerrado"
          End Select
            
              itmX.SubItems(4) = Format(rs!RESERVA_PORC, "Standard")
          
              itmX.SubItems(5) = Format(rs!Capitaliza_Porc, "Standard")
              itmX.SubItems(6) = IIf((rs!Capitaliza_Renta_Aplica = 1), "Sí", "No")
          
              itmX.SubItems(7) = rs!NC_Saldos
              itmX.SubItems(8) = rs!NC_Mora
              itmX.SubItems(9) = rs!NC_OPCF
              
              itmX.SubItems(10) = IIf((rs!VISIBLE_WEBAPP = 1), "Sí", "No")
              itmX.SubItems(11) = IIf((rs!VISIBLE_SYS = 1), "Sí", "No")
              
              
          rs.MoveNext
        Loop
        rs.Close
    
    Case 1, 2 'Nada
        
    Case 3 'Bitacora
        Call rbBitacora_Click(0)
    
    
    Case 4 'Resumen

        strSQL = "select * from vExc_Periodos_Cierres_Resumen" _
               & " where id_Periodo = " & vCodigo
        Call OpenRecordSet(rs, strSQL)
        
        lswResumen.ListItems.Clear
        
        Do While Not rs.EOF
          Set itmX = lswResumen.ListItems.Add(, , "Excedente Bruto")
              itmX.SubItems(1) = Format(rs!Excedente_Bruto, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")

          Set itmX = lswResumen.ListItems.Add(, , "(-) Reserva")
              itmX.SubItems(1) = Format(rs!Reserva, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")

          Set itmX = lswResumen.ListItems.Add(, , "(-) Capitalizado")
              itmX.SubItems(1) = Format(rs!Capitalizacion, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")

          Set itmX = lswResumen.ListItems.Add(, , "(-) Renta")
              itmX.SubItems(1) = Format(rs!Renta, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")
          
          Set itmX = lswResumen.ListItems.Add(, , "Excedente Neto")
              itmX.SubItems(1) = Format(rs!Excedente_Neto, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")
              itmX.Bold = True
              
          Set itmX = lswResumen.ListItems.Add(, , "(-) Donaciones")
              itmX.SubItems(1) = Format(rs!Donacion, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")
          
          Set itmX = lswResumen.ListItems.Add(, , "(+/-) Ajustes")
              itmX.SubItems(1) = Format(rs!Ajuste_Aplicado, "Standard")
              itmX.SubItems(2) = Format(rs!Ajuste_Cargado, "Standard")
          
          
          Set itmX = lswResumen.ListItems.Add(, , "(-) Crédito s/Excedente")
              itmX.SubItems(1) = Format(rs!CEXD_Aplicado, "Standard")
              itmX.SubItems(2) = Format(rs!CEXD_Cargado, "Standard")
          
          Set itmX = lswResumen.ListItems.Add(, , "(-) Morosidad")
              itmX.SubItems(1) = Format(rs!Mora_Aplicada, "Standard")
              itmX.SubItems(2) = Format(rs!Mora_Cargada, "Standard")
          
          Set itmX = lswResumen.ListItems.Add(, , "(-) O.P.C.F.")
              itmX.SubItems(1) = Format(rs!OPCF_Aplicado, "Standard")
              itmX.SubItems(2) = Format(rs!OPCF_Cargado, "Standard")
          
          Set itmX = lswResumen.ListItems.Add(, , "Capitaliza Extradordinario")
              itmX.SubItems(1) = Format(rs!Capitalizado_Indivual, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")
          
          Set itmX = lswResumen.ListItems.Add(, , "Excedente Final")
              itmX.SubItems(1) = Format(rs!Excedente_Final, "Standard")
              itmX.SubItems(2) = Format(0, "Standard")
              itmX.Bold = True
          
          rs.MoveNext
        Loop
        rs.Close

End Select

Me.MousePointer = vbDefault

vPaso = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    
      tcMain.Item(1).Selected = True
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtEstado.SetFocus
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
       gBusquedas.Columna = "ID_PERIODO"
       gBusquedas.Orden = "ID_PERIODO"
       gBusquedas.Consulta = "select ID_PERIODO,Inicio,Corte,Estado from EXC_PERIODOS"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(pPeriodoId As Long)
Dim i As Integer

On Error GoTo vError

If Not fxSIFValidaCadena(txtCodigo.Text) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

txtEstado.SetFocus

tcMain.Item(1).Selected = True
GroupBox1.Enabled = True

For i = 0 To btnBoleta.Count - 1
    btnBoleta.Item(i).Enabled = True
Next i

vPaso = True

strSQL = "select P.*" _
       & " from vExc_Periodos_Consulta P " _
       & " where P.ID_PERIODO = " & pPeriodoId
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!ID_PERIODO
  txtCodigo.Text = rs!ID_PERIODO

    dtpInicio.Value = rs!INICIO
    dtpCorte.Value = rs!Corte
    
    
    txtReservaPorc.Text = Format(rs!RESERVA_PORC, "Standard")
    txtCapPorc.Text = Format(rs!Capitaliza_Porc, "Standard")
    txtDocExc.Text = rs!Doc_Asiento
    
    txtNC_CAPIND.Text = rs!NC_FND_Extra
    txtNC_CEXD.Text = rs!NC_Saldos
    txtNC_Mora.Text = rs!NC_Mora
    txtNC_OPCF.Text = rs!NC_OPCF
    
    Select Case rs!ESTADO
    Case "A"
        txtEstado.Text = "Abierto"
    Case "C"
        txtEstado.Text = "Cerrado"
    End Select
    
    chkRentaCap.Value = rs!Capitaliza_Renta_Aplica
    
    chkVisible_Web.Value = rs!VISIBLE_WEBAPP
    chkVisible_Sys.Value = rs!VISIBLE_SYS
    chkVisible_Historial.Value = rs!MOSTRAR_EN_HISTORIAL
    chkVisible_Renta_Tabla.Value = rs!MOSTRAR_TABLA_RENTA
    txtEstado_Nota.Text = rs!ESTADO_NOTAS & ""
    
    Call sbCboAsignaDato(cboBaseAplicacion, rs!TIPO_APL_MENSUAL_DESC, True, rs!TIPO_APL_MENSUAL)
    
    'Cargar Tabla de Renta
    strSQL = "select DESDE, HASTA, PORCENTAJE " _
           & " From EXC_RENTA_TABLA_H" _
           & " Where ID_Periodo = " & rs!ID_PERIODO _
           & " order by desde"
    Call OpenRecordSet(rs, strSQL)
    
    With lswRenta.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , Format(rs!Desde, "Standard"))
          itmX.SubItems(1) = Format(rs!Hasta, "Standard")
          itmX.SubItems(2) = Format(rs!Porcentaje, "Standard")
      rs.MoveNext
    Loop
    rs.Close
    End With
    
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

txtEstado_Nota.Text = fxSysCleanTxtInject(txtEstado_Nota.Text)

If txtEstado.Text = "Cerrado" Then vMensaje = vMensaje & vbCrLf & " - El Periodo ya fue cerrado!"

If dtpInicio.Value >= dtpCorte.Value Then vMensaje = vMensaje & vbCrLf & " - El Rango de Fechas es Erroneo!"

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()

On Error GoTo vError


If vEdita Then
  strSQL = "update EXC_PERIODOS set CAPITALIZA_PORC = " & CCur(txtCapPorc.Text) _
         & ", RESERVA_PORC = " & CCur(txtReservaPorc.Text) & ", CAPITALIZA_RENTA_APLICA = " & chkRentaCap.Value _
         & ", INICIO = '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00', CORTE = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
         & ", VISIBLE_WEBAPP = " & chkVisible_Web.Value & ", VISIBLE_SYS = " & chkVisible_Sys.Value _
         & ", MOSTRAR_EN_HISTORIAL = " & chkVisible_Historial.Value & ", MOSTRAR_TABLA_RENTA = " & chkVisible_Renta_Tabla.Value _
         & ", ESTADO_NOTAS = '" & txtEstado_Nota.Text & "'" _
         & " where ID_PERIODO = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Excedentes> Periodo:  " & vCodigo)

Else
  Dim rs As New ADODB.Recordset
  
  strSQL = "select isnull(max(id_Periodo),0) + 1 as 'Periodo' from EXC_PERIODOS"
  Call OpenRecordSet(rs, strSQL)
    vCodigo = rs!Periodo
  rs.Close

  txtCodigo.Text = CStr(vCodigo)
  
   strSQL = "insert into EXC_PERIODOS(ID_PERIODO,INICIO,CORTE, ESTADO,CAPITALIZA_PORC, RESERVA_PORC, CAPITALIZA_RENTA_APLICA" _
          & ", NC_MORA, NC_SALDOS, NC_OPCF, NC_FND_EXTRA, DOC_ASIENTO , VISIBLE_WEBAPP, VISIBLE_SYS , TIPO_APL_MENSUAL" _
          & ", MOSTRAR_EN_HISTORIAL, MOSTRAR_TABLA_RENTA, ESTADO_NOTAS" _
          & ", REGISTRO_FECHA, REGISTRO_USUARIO)" _
          & " values(" & vCodigo & ", '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00', '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59','A'" _
          & ", " & CCur(txtCapPorc.Text) & ", " & CCur(txtReservaPorc.Text) & ", " & chkRentaCap.Value _
          & ", '', '', '', '', '', " & chkVisible_Web.Value & ", " & chkVisible_Sys.Value & ", '" & cboBaseAplicacion.ItemData(cboBaseAplicacion.ListIndex) _
          & "', " & chkVisible_Historial.Value & ", " & chkVisible_Renta_Tabla.Value & ", '" & txtEstado_Nota.Text _
          & "', dbo.MyGetdate(), '" & glogon.Usuario & "')"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Excedentes> Periodo:  " & vCodigo)

End If

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete EXC_PERIODOS where ID_PERIODO = " & vCodigo
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Excedentes> Periodo:  " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "id_periodo"
  gBusquedas.Orden = "id_periodo"
  gBusquedas.Consulta = "select id_periodo,inicio,corte,estado from EXC_PERIODOS"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo.Text = gBusquedas.Resultado
  If IsNumeric(txtCodigo.Text) Then Call sbConsulta(txtCodigo.Text)
End If

End Sub


Private Sub txtCodigo_LostFocus()

If IsNumeric(txtCodigo.Text) Then
  Call sbConsulta(txtCodigo.Text)
End If
End Sub
