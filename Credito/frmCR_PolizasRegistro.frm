VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCR_PolizasRegistro 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Pólizas"
   ClientHeight    =   8040
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   9432
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6252
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   9132
      _Version        =   1310720
      _ExtentX        =   16108
      _ExtentY        =   11028
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Lista de Pólizas"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "lsw"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "txtPolizaContrato"
      Item(1).Control(1)=   "txtPolizaId"
      Item(1).Control(2)=   "tlbPrincipal"
      Item(1).Control(3)=   "tlbBeneficiarios"
      Item(1).Control(4)=   "Label1(13)"
      Item(1).Control(5)=   "Label1(17)"
      Item(1).Control(6)=   "Label1(15)"
      Item(1).Control(7)=   "cboPolizaLinea"
      Item(1).Control(8)=   "tcAux"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5772
         Left            =   -69880
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   8892
         _Version        =   1310720
         _ExtentX        =   15684
         _ExtentY        =   10181
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
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   5052
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   8772
         _Version        =   1310720
         _ExtentX        =   15473
         _ExtentY        =   8911
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
         Item(0).Caption =   "General"
         Item(0).ControlCount=   28
         Item(0).Control(0)=   "fraPolizaRetencion"
         Item(0).Control(1)=   "txtPolizaCuotaRestoPlazo"
         Item(0).Control(2)=   "chkPolizaPlazoCredito"
         Item(0).Control(3)=   "txtPolizasCtaDeduce"
         Item(0).Control(4)=   "txtPolizaPagosNum"
         Item(0).Control(5)=   "txtPolizaCoberturaMeses"
         Item(0).Control(6)=   "txtPolizaPagoMonto"
         Item(0).Control(7)=   "txtPolizaMonto"
         Item(0).Control(8)=   "cboPolizaPagoFrecuencia"
         Item(0).Control(9)=   "txtPolizaCuota"
         Item(0).Control(10)=   "cboPolizaPlan"
         Item(0).Control(11)=   "cboPolizaEstado"
         Item(0).Control(12)=   "dtpPolizaFechaPago"
         Item(0).Control(13)=   "dtpPolizaCoberturaInicio"
         Item(0).Control(14)=   "dtpPolizaCoberturaCorte"
         Item(0).Control(15)=   "Label1(39)"
         Item(0).Control(16)=   "Label1(35)"
         Item(0).Control(17)=   "Label1(34)"
         Item(0).Control(18)=   "Label1(33)"
         Item(0).Control(19)=   "Label1(26)"
         Item(0).Control(20)=   "Label1(25)"
         Item(0).Control(21)=   "Label1(24)"
         Item(0).Control(22)=   "Label1(23)"
         Item(0).Control(23)=   "Label1(22)"
         Item(0).Control(24)=   "Label1(21)"
         Item(0).Control(25)=   "Label1(18)"
         Item(0).Control(26)=   "Label1(19)"
         Item(0).Control(27)=   "Label1(20)"
         Item(1).Caption =   "Pagos"
         Item(1).ControlCount=   6
         Item(1).Control(0)=   "lswPolizaPago"
         Item(1).Control(1)=   "lblPolizaPagoProximo"
         Item(1).Control(2)=   "Label1(30)"
         Item(1).Control(3)=   "lblPolizaPagoSaldo"
         Item(1).Control(4)=   "Label1(29)"
         Item(1).Control(5)=   "Label1(27)"
         Item(2).Caption =   "Recaudación"
         Item(2).ControlCount=   6
         Item(2).Control(0)=   "Label1(36)"
         Item(2).Control(1)=   "lblPolizaRecaudadoCorte"
         Item(2).Control(2)=   "lblPolizaRecaudadoSaldo"
         Item(2).Control(3)=   "Label1(32)"
         Item(2).Control(4)=   "Label1(31)"
         Item(2).Control(5)=   "lswPolizaRecaudado"
         Item(3).Caption =   "Acreedores"
         Item(3).ControlCount=   2
         Item(3).Control(0)=   "Label1(37)"
         Item(3).Control(1)=   "lswAcreedores"
         Item(4).Caption =   "Beneficiarios"
         Item(4).ControlCount=   2
         Item(4).Control(0)=   "Label1(38)"
         Item(4).Control(1)=   "lswBeneficiarios"
         Begin XtremeSuiteControls.ListView lswPolizaRecaudado 
            Height          =   3132
            Left            =   -69760
            TabIndex        =   75
            Top             =   720
            Visible         =   0   'False
            Width           =   8532
            _Version        =   1310720
            _ExtentX        =   15049
            _ExtentY        =   5524
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
         Begin XtremeSuiteControls.ListView lswAcreedores 
            Height          =   3612
            Left            =   -69760
            TabIndex        =   76
            Top             =   720
            Visible         =   0   'False
            Width           =   8532
            _Version        =   1310720
            _ExtentX        =   15049
            _ExtentY        =   6371
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
         End
         Begin XtremeSuiteControls.ListView lswBeneficiarios 
            Height          =   3612
            Left            =   -69760
            TabIndex        =   77
            Top             =   720
            Visible         =   0   'False
            Width           =   8532
            _Version        =   1310720
            _ExtentX        =   15049
            _ExtentY        =   6371
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
         Begin XtremeSuiteControls.ListView lswPolizaPago 
            Height          =   3132
            Left            =   -69760
            TabIndex        =   74
            Top             =   720
            Visible         =   0   'False
            Width           =   8532
            _Version        =   1310720
            _ExtentX        =   15049
            _ExtentY        =   5524
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
         Begin VB.Frame fraPolizaRetencion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4092
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   8772
            Begin VB.ComboBox cboPolizaOperacion 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               ItemData        =   "frmCR_PolizasRegistro.frx":0000
               Left            =   2160
               List            =   "frmCR_PolizasRegistro.frx":0002
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   240
               Width           =   1815
            End
            Begin VB.TextBox txtPlazo 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2160
               TabIndex        =   26
               Top             =   1320
               Width           =   1695
            End
            Begin VB.TextBox txtObservaciones 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   2160
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   25
               Top             =   1680
               Width           =   6132
            End
            Begin VB.TextBox txtDocumento 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6120
               TabIndex        =   24
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox txtAnio 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   6000
               TabIndex        =   23
               Top             =   2400
               Width           =   735
            End
            Begin VB.ComboBox cboMes 
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   312
               ItemData        =   "frmCR_PolizasRegistro.frx":0004
               Left            =   6720
               List            =   "frmCR_PolizasRegistro.frx":002F
               Style           =   2  'Dropdown List
               TabIndex        =   22
               ToolTipText     =   "Mes a procesar"
               Top             =   2400
               Width           =   1335
            End
            Begin VB.TextBox txtMonto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   314
               Left            =   6120
               TabIndex        =   21
               Top             =   1320
               Width           =   1935
            End
            Begin VB.TextBox txtEstado 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   20
               Top             =   2880
               Width           =   1812
            End
            Begin VB.TextBox txtFecha 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   19
               Top             =   3240
               Width           =   1812
            End
            Begin VB.TextBox txtPlazoTrasnscurrido 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2160
               Locked          =   -1  'True
               TabIndex        =   18
               Top             =   3600
               Width           =   1812
            End
            Begin VB.TextBox txtProyectado 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   314
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   17
               Top             =   2880
               Width           =   1455
            End
            Begin VB.TextBox txtPagado 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   314
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   3240
               Width           =   1455
            End
            Begin VB.TextBox txtPendiente 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   314
               Left            =   6600
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   3600
               Width           =   1455
            End
            Begin XtremeSuiteControls.ComboBox cboDestino 
               Height          =   312
               Left            =   2160
               TabIndex        =   28
               Top             =   600
               Width           =   5892
               _Version        =   1310720
               _ExtentX        =   10393
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
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin XtremeSuiteControls.ComboBox cboGarantia 
               Height          =   312
               Left            =   2160
               TabIndex        =   29
               Top             =   960
               Width           =   1692
               _Version        =   1310720
               _ExtentX        =   2985
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
               UseVisualStyle  =   0   'False
               Text            =   "ComboBox1"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "No. Operación de Póliza"
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
               Left            =   120
               TabIndex        =   43
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Plazo"
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
               Left            =   840
               TabIndex        =   42
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Cuota"
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
               Index           =   16
               Left            =   4680
               TabIndex        =   41
               Top             =   1320
               Width           =   1092
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Observación"
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
               Left            =   840
               TabIndex        =   40
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Documento"
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
               Left            =   4680
               TabIndex        =   39
               Top             =   960
               Width           =   1092
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   12
               Left            =   1080
               TabIndex        =   38
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
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
               Height          =   255
               Index           =   28
               Left            =   1080
               TabIndex        =   37
               Top             =   600
               Width           =   975
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
               Index           =   4
               Left            =   720
               TabIndex        =   36
               Top             =   2880
               Width           =   1332
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Pagado"
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
               Index           =   5
               Left            =   5520
               TabIndex        =   35
               Top             =   3240
               Width           =   972
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
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
               Height          =   252
               Index           =   7
               Left            =   720
               TabIndex        =   34
               Top             =   3240
               Width           =   1332
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Plazo Trans."
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
               Index           =   8
               Left            =   720
               TabIndex        =   33
               Top             =   3600
               Width           =   1332
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Proyectado"
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
               Index           =   10
               Left            =   5520
               TabIndex        =   32
               Top             =   2880
               Width           =   972
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Pendiente"
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
               Left            =   5520
               TabIndex        =   31
               Top             =   3600
               Width           =   972
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               Index           =   0
               X1              =   6960
               X2              =   -360
               Y1              =   2760
               Y2              =   2760
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Primer Deducción (aaaa/mm)"
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
               Left            =   3720
               TabIndex        =   30
               Top             =   2400
               Width           =   2172
            End
         End
         Begin VB.TextBox txtPolizaCuota 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
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
            Height          =   314
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   50
            ToolTipText     =   "Cuota de la Póliza"
            Top             =   2880
            Width           =   1935
         End
         Begin VB.TextBox txtPolizaMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   314
            Left            =   2280
            TabIndex        =   49
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox txtPolizaPagoMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
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
            Height          =   314
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   3240
            Width           =   1935
         End
         Begin VB.TextBox txtPolizaCoberturaMeses 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
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
            Height          =   314
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   47
            ToolTipText     =   "Meses entre la Cobertura Inicio y Corte"
            Top             =   2520
            Width           =   1572
         End
         Begin VB.TextBox txtPolizaPagosNum 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
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
            Height          =   314
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   46
            ToolTipText     =   "No. Pagos a Realizar desde el Prox. Pago hasta la Fecha Corte de la Cobertura (según la frecuencia de pago)"
            Top             =   3240
            Width           =   1572
         End
         Begin VB.TextBox txtPolizasCtaDeduce 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
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
            Height          =   314
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2880
            Width           =   1572
         End
         Begin VB.TextBox txtPolizaCuotaRestoPlazo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.4
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   314
            Left            =   2280
            TabIndex        =   44
            ToolTipText     =   "La cuota es la Poliza entre el plazo de cobertura en meses"
            Top             =   1080
            Width           =   1935
         End
         Begin XtremeSuiteControls.ComboBox cboPolizaEstado 
            Height          =   312
            Left            =   2280
            TabIndex        =   78
            Top             =   360
            Width           =   1932
            _Version        =   1310720
            _ExtentX        =   3408
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
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboPolizaPagoFrecuencia 
            Height          =   312
            Left            =   2280
            TabIndex        =   79
            Top             =   1440
            Width           =   1932
            _Version        =   1310720
            _ExtentX        =   3408
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
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboPolizaPlan 
            Height          =   312
            Left            =   2280
            TabIndex        =   81
            Top             =   3720
            Width           =   5652
            _Version        =   1310720
            _ExtentX        =   9970
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
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.DateTimePicker dtpPolizaFechaPago 
            Height          =   312
            Left            =   2280
            TabIndex        =   82
            Top             =   1800
            Width           =   1932
            _Version        =   1310720
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.DateTimePicker dtpPolizaCoberturaInicio 
            Height          =   312
            Left            =   2280
            TabIndex        =   83
            Top             =   2160
            Width           =   1932
            _Version        =   1310720
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.DateTimePicker dtpPolizaCoberturaCorte 
            Height          =   312
            Left            =   2280
            TabIndex        =   84
            Top             =   2520
            Width           =   1932
            _Version        =   1310720
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.CheckBox chkPolizaPlazoCredito 
            Height          =   612
            Left            =   4440
            TabIndex        =   80
            Top             =   960
            Width           =   2652
            _Version        =   1310720
            _ExtentX        =   4678
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Incluir esta Cuota en el Resto del Plan?"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   16
         End
         Begin VB.Label Label1 
            Caption         =   "Lista de Beneficiarios de la Póliza..:"
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
            Index           =   38
            Left            =   -69760
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   3252
         End
         Begin VB.Label Label1 
            Caption         =   "Lista de Acreedores de la Póliza..:"
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
            Index           =   37
            Left            =   -69760
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   3252
         End
         Begin VB.Label Label1 
            Caption         =   "Cobros Realizados..:"
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
            Index           =   31
            Left            =   -69760
            TabIndex        =   73
            Top             =   360
            Visible         =   0   'False
            Width           =   1812
         End
         Begin VB.Label Label1 
            Caption         =   "Saldo ..:"
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
            Index           =   32
            Left            =   -69640
            TabIndex        =   72
            Top             =   3960
            Visible         =   0   'False
            Width           =   1212
         End
         Begin VB.Label lblPolizaRecaudadoSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "0.00"
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
            Left            =   -68200
            TabIndex        =   71
            Top             =   3960
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Label lblPolizaRecaudadoCorte 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "dd/mm/yyyy"
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
            Left            =   -63160
            TabIndex        =   70
            Top             =   3960
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Label Label1 
            Caption         =   "Ultimo Corte..:"
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
            Index           =   36
            Left            =   -64840
            TabIndex        =   69
            Top             =   3960
            Visible         =   0   'False
            Width           =   1332
         End
         Begin VB.Label Label1 
            Caption         =   "Pagos Realizados..:"
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
            Index           =   27
            Left            =   -69760
            TabIndex        =   68
            Top             =   360
            Visible         =   0   'False
            Width           =   1812
         End
         Begin VB.Label Label1 
            Caption         =   "Saldo por Pagar..:"
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
            Index           =   29
            Left            =   -69760
            TabIndex        =   67
            Top             =   3960
            Visible         =   0   'False
            Width           =   1812
         End
         Begin VB.Label lblPolizaPagoSaldo 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "0.00"
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
            Left            =   -67720
            TabIndex        =   66
            Top             =   3960
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Label Label1 
            Caption         =   "Próximo Pago..:"
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
            Index           =   30
            Left            =   -65560
            TabIndex        =   65
            Top             =   3960
            Visible         =   0   'False
            Width           =   1812
         End
         Begin VB.Label lblPolizaPagoProximo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "dd/mm/yyyy"
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
            Left            =   -63160
            TabIndex        =   64
            Top             =   3960
            Visible         =   0   'False
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   20
            Left            =   840
            TabIndex        =   63
            Top             =   360
            Width           =   1212
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Incorporar al Plan a partir de la cuota:"
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
            Index           =   19
            Left            =   120
            TabIndex        =   62
            Top             =   3600
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Mensualidad"
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
            Index           =   18
            Left            =   480
            TabIndex        =   61
            Top             =   2880
            Width           =   1572
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Frecuencia de Pago"
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
            Index           =   21
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Index           =   22
            Left            =   600
            TabIndex        =   59
            Top             =   720
            Width           =   1452
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Próximo Pago"
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
            Index           =   23
            Left            =   120
            TabIndex        =   58
            Top             =   1800
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cobertura Inicio"
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
            Index           =   24
            Left            =   120
            TabIndex        =   57
            Top             =   2160
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cobertura Corte"
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
            Index           =   25
            Left            =   120
            TabIndex        =   56
            Top             =   2520
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Monto de los Pagos"
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
            Index           =   26
            Left            =   480
            TabIndex        =   55
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cobertura (Meses)"
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
            Index           =   33
            Left            =   4200
            TabIndex        =   54
            Top             =   2520
            Width           =   1812
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "No. Pagos a Realizar"
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
            Left            =   4200
            TabIndex        =   53
            Top             =   3240
            Width           =   1932
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "No. Ctas a Deducir"
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
            Index           =   35
            Left            =   4200
            TabIndex        =   52
            Top             =   2880
            Width           =   1812
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Cuota Resto del Plazo?"
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
            Index           =   39
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   1932
         End
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   264
         Left            =   5520
         TabIndex        =   7
         Top             =   360
         Width           =   2784
         _ExtentX        =   4911
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "Imprime el listado seleccionado"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbBeneficiarios 
         Height          =   264
         Left            =   8400
         TabIndex        =   8
         Top             =   360
         Width           =   396
         _ExtentX        =   699
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Beneficiarios"
               Object.ToolTipText     =   "Registro de Benefiicarios"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.ComboBox cboPolizaLinea 
         Height          =   312
         Left            =   3720
         TabIndex        =   12
         Top             =   840
         Width           =   5172
         _Version        =   1310720
         _ExtentX        =   9123
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
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPolizaId 
         Height          =   312
         Left            =   3720
         TabIndex        =   92
         Top             =   1200
         Width           =   1812
         _Version        =   1310720
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPolizaContrato 
         Height          =   312
         Left            =   7080
         TabIndex        =   93
         Top             =   1200
         Width           =   1812
         _Version        =   1310720
         _ExtentX        =   3196
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Código de Póliza"
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
         Index           =   15
         Left            =   2280
         TabIndex        =   11
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número de Póliza"
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
         Index           =   17
         Left            =   1800
         TabIndex        =   10
         Top             =   1200
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "No. Contrato"
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
         Left            =   5160
         TabIndex        =   9
         Top             =   1200
         Width           =   1692
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasRegistro.frx":0098
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PolizasRegistro.frx":0194
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   0
      Top             =   7788
      Width           =   9432
      _ExtentX        =   16637
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3951
            MinWidth        =   3951
            Object.ToolTipText     =   "Usuario > Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha > Registro"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   1800
      TabIndex        =   85
      Top             =   240
      Width           =   2052
      _Version        =   1310720
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3960
      TabIndex        =   87
      Top             =   240
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3480
      TabIndex        =   88
      Top             =   960
      Width           =   5772
      _Version        =   1310720
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3480
      TabIndex        =   89
      Top             =   1320
      Width           =   5772
      _Version        =   1310720
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1800
      TabIndex        =   90
      Top             =   960
      Width           =   1692
      _Version        =   1310720
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   91
      Top             =   1320
      Width           =   1692
      _Version        =   1310720
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   86
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label1 
      Caption         =   "Línea"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmCR_PolizasRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String 'Envia Mensajes en Fallas de Verificacion
Dim vEdita As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso As Boolean, vScroll As Boolean
Dim vMesesPenPlan As Integer 'Indica los meses pendientes en el plan para la cobertura
Dim vPlazo As Integer, vFechaFormaliza As Date




Private Function fxVerifica() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, vPolizaLinea As String, vIntegraPP As Integer

On Error GoTo vError

fxVerifica = True
vMensaje = ""


'Verifica que exista la persona
strSQL = "select isnull(count(*),0) as Existe from socios where cedula ='" & txtCedula & "'"
Call OpenRecordSet(rsX, strSQL, 0)
 If rsX!Existe = 0 Then
   vMensaje = vMensaje & vbCrLf & "- No Existe el cliente definido (debe de Ingresarlo)..."
 End If
rsX.Close



If cboPolizaLinea.ListCount <= 0 Then
    vMensaje = vMensaje & vbCrLf & "- No Existen Polizas Configuradas para usar..."
Else

    'Informacion de Uso de la Poliza
    strSQL = "select CODIGO_RETENCION,CODIGO_CARGO,INTEGRA_PLAN_PAGOS from CRD_CATALOGO_POLIZAS " _
           & " where COD_POLIZA = '" & cboPolizaLinea.ItemData(cboPolizaLinea.ListIndex) & "'"
    Call OpenRecordSet(rsX, strSQL, 0)
      vPolizaLinea = Trim(rsX!codigo_retencion)
      vIntegraPP = rsX!integra_plan_pagos
    rsX.Close

End If 'Poliza Linea


'TODO:
' 1. Verificar si la cobertura del Plan sobre pasa el plazo de vencimiento del crédito (Revisar)

If vIntegraPP = 1 Then
        If IsNumeric(txtPolizaMonto.Text) Then
         If txtPolizaMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- La mensualidad de la Póliza no es válidad"
        Else
           vMensaje = vMensaje & vbCrLf & "- La mensualidad de la Póliza no es válidad"
        End If
        
        If IsNumeric(txtPolizaCuotaRestoPlazo.Text) Then
         If txtPolizaCuotaRestoPlazo.Text < 1 And chkPolizaPlazoCredito.Value = vbChecked Then vMensaje = vMensaje & vbCrLf & "- La mensualidad de la Póliza para el resto del Plan no es válidad"
        Else
           vMensaje = vMensaje & vbCrLf & "- La mensualidad de la Póliza para el Resto del Plan no es válidad"
        End If
        
        If cboPolizaPlan.ListCount = 0 Then
           vMensaje = vMensaje & vbCrLf & "- No existen cuotas disponibles dentro del Plan de Pagos en donde registrar la Póliza"
        End If
        
        If dtpPolizaCoberturaCorte.Value <= dtpPolizaCoberturaInicio.Value Then
           vMensaje = vMensaje & vbCrLf & "- La cobertura de la poliza no es válida verifique"
        End If
        If dtpPolizaCoberturaCorte.Value <= dtpPolizaCoberturaInicio.Value Then
           vMensaje = vMensaje & vbCrLf & "- La cobertura de la poliza no es válida verifique"
        End If
        
        
Else

        If txtDocumento.Text = "" Then vMensaje = vMensaje & vbCrLf & "- No se especificó el # Documento ? "
        
        If IsNumeric(txtPlazo) Then
         If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo definido no es válido"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
        End If
        
        If cboGarantia.Text = "" Or cboGarantia.ListCount <= 0 Then vMensaje = vMensaje & vbCrLf & "- No se especificó el tipo de garantía"
        
        If IsNumeric(txtMonto.Text) Then
         If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto de la Póliza no es válido"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado es Inválido"
        End If
        
        'Verifica la Operacion Madre
        strSQL = "select isnull(count(*),0) as Existe from catalogo where Retencion = 'N' and Poliza = 'N' and codigo ='" & txtCodigo & "'"
        Call OpenRecordSet(rsX, strSQL, 0)
         If rsX!Existe = 0 Then
           vMensaje = vMensaje & vbCrLf & "- La Línea de la Operacion madre no es un crédito o no es válido..."
         End If
        rsX.Close
        
        
        'Verifica que no existe una retencion con el mismo codigo
        strSQL = "select isnull(count(*),0) as Existe from crd_operacion_polizas where cod_poliza ='" & vPolizaLinea & "' and id_solicitud = " & txtOperacion.Text
        Call OpenRecordSet(rsX, strSQL, 0)
         If rsX!Existe > 0 Then
           vMensaje = vMensaje & vbCrLf & "- Ya existe una Póliza activa para esta operación de crédito..."
         End If
        rsX.Close
        
        strSQL = "select ctaNintC from catalogo where codigo ='" & vPolizaLinea & "'"
        Call OpenRecordSet(rsX, strSQL, 0)
        If rsX.EOF And rsX.BOF Then
           vMensaje = vMensaje & vbCrLf & "- El código de la Poliza no existe"
         Else
          If IsNull(rsX!ctaNintC) Then vMensaje = vMensaje & vbCrLf & "- El código no se encuentra codificado contablemente"
        End If
        rsX.Close
        
        If fxConvierteMES(cboMes.Text) = cboMes.Text Then vMensaje = vMensaje & vbCrLf & "- El Mes para la primer deduccion no es válido"
        
        lngPriDeduc = txtAnio.Text & Format(fxConvierteMES(cboMes.Text), "00")
        
        If lngPriDeduc <= GLOBALES.glngFechaCR Then vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque es igual o menor a la fecha de proceso actual"


End If 'Integra Plan de Pagos



If Len(vMensaje) > 0 Then fxVerifica = False


Exit Function

vError:
  vMensaje = vMensaje & vbCrLf & fxSys_Error_Handler(Err.Description)
  fxVerifica = False

End Function



Private Sub cboDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboGarantia.SetFocus
End Sub



Private Sub cboGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub


Private Sub cboPolizaLinea_Click()
Dim strSQL As String, rs As New ADODB.Recordset, vPolizaLinea As String

On Error GoTo vError

If vPaso Or cboPolizaLinea.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass


fraPolizaRetencion.BorderStyle = 0
fraPolizaRetencion.top = 360
fraPolizaRetencion.Left = 120


tcAux.Item(0).Selected = True


'Informacion de Uso de la Poliza
strSQL = "select CODIGO_RETENCION,CODIGO_CARGO,INTEGRA_PLAN_PAGOS from CRD_CATALOGO_POLIZAS " _
       & " where COD_POLIZA = '" & cboPolizaLinea.ItemData(cboPolizaLinea.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!integra_plan_pagos = 1 Then
  fraPolizaRetencion.Visible = False
'  txtPolizaContrato.Visible = True
Else
  fraPolizaRetencion.Visible = True
'  txtPolizaContrato.Visible = False
  
  vPolizaLinea = Trim(rs!codigo_retencion)

  vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, vPolizaLinea)
        Call sbSTCargaCboDestinos(cboDestino, vPolizaLinea)
  vPaso = False
  
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub cboPolizaOperacion_Click()
If vPaso Or cboPolizaOperacion.ListCount <= 0 Then Exit Sub

Call sbCargaPoliza(cboPolizaOperacion.Text)

End Sub

Private Sub sbCalculaMesesPendientePlan()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "select dbo.fxCrdPolizaMesesPendientes(" & txtOperacion.Text & "," & cboPolizaPlan.ItemData(cboPolizaPlan.ListIndex) _
       & ",'" & Format(dtpPolizaCoberturaInicio.Value, "yyyy/mm/dd") _
       & "','" & Format(dtpPolizaCoberturaCorte.Value, "yyyy/mm/dd") & "') as 'Meses'"
Call OpenRecordSet(rs, strSQL)
 vMesesPenPlan = rs!Meses
 txtPolizasCtaDeduce.Text = rs!Meses
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboPolizaPagoFrecuencia_Click()

If vPaso Then Exit Sub
Call sbCalculaMesesPendientePlan
Call sbCalculaDatosPoliza

End Sub

Private Sub cboPolizaPlan_Click()
If vPaso Then Exit Sub
If cboPolizaPlan.ListCount <= 0 Then Exit Sub

Call sbCalculaMesesPendientePlan
Call sbCalculaDatosPoliza

End Sub




Private Sub chkPolizaPlazoCredito_Click()

If chkPolizaPlazoCredito.Value = vbChecked Then
   txtPolizaCuotaRestoPlazo.Locked = False
   txtPolizaCuotaRestoPlazo.BackColor = vbWhite
Else
   txtPolizaCuotaRestoPlazo.Locked = True
   txtPolizaCuotaRestoPlazo.BackColor = txtPolizaCuota.BackColor
End If

End Sub

Private Sub dtpPolizaCoberturaCorte_Change()

Call sbCalculaMesesPendientePlan
Call sbCalculaDatosPoliza

End Sub

Private Sub dtpPolizaCoberturaInicio_Change()

Call sbCalculaMesesPendientePlan
Call sbCalculaDatosPoliza

End Sub



Private Sub dtpPolizaFechaPago_Change()

Call sbCalculaMesesPendientePlan
Call sbCalculaDatosPoliza

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtOperacion.Text = "" Then txtOperacion.Text = "0"

If vScroll Then
    strSQL = "select Top 1 R.id_solicitud from reg_creditos R inner join Catalogo C on R.codigo = C.codigo"

    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where (C.retencion = 'N' or C.poliza = 'N') and R.id_solicitud > " & txtOperacion & " order by R.id_solicitud asc"
    Else
       strSQL = strSQL & " where (C.retencion = 'N' or C.poliza = 'N') and R.id_solicitud < " & txtOperacion & " order by R.id_solicitud desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtOperacion = rs!id_solicitud
      Call sbCargaOperacion
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
Dim strSQL As String

 vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

tcMain.Item(0).Selected = True


txtPolizaId.BackColor = RGB(187, 215, 247)
txtPolizaContrato.BackColor = txtPolizaId.BackColor


With lsw.ColumnHeaders
  .Add , , "Póliza", 1200
  .Add , , "Descripción", 3200
  .Add , , "Tipo", 1200, vbCenter
  .Add , , "Referencia", 1200, vbCenter
  .Add , , "Mensualidad", 1400, vbRightJustify
  .Add , , "[]", 1400
End With


 
With lswPolizaPago.ColumnHeaders
    .Add , , "No.Pago", 1200
    .Add , , "Monto", 1200, vbRightJustify
    .Add , , "Fecha Pago", 1600
    .Add , , "Fecha Transac", 1600
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Recaudado Corte", 1600, vbRightJustify
    .Add , , "Remesa", 1400, vbCenter
End With

With lswPolizaRecaudado.ColumnHeaders
    .Add , , "Línea", 1200
    .Add , , "No.Cuota", 1200, vbCenter
    .Add , , "Monto", 1400, vbRightJustify
    .Add , , "Fecha", 1600
End With

With lswAcreedores.ColumnHeaders
    .Add , , "Código", 1200
    .Add , , "Identificación", 1450
    .Add , , "Nombre", 4000
    .Add , , "Fecha", 1400
    .Add , , "Usuario", 1400
End With

With lswBeneficiarios.ColumnHeaders
    .Add , , "Identificación", 1400
    .Add , , "Nombre", 4000
    .Add , , "Fecha Nac.", 1400
    .Add , , "Parentesco", 1400
    .Add , , "Porcentaje", 1200, vbRightJustify
End With


 Call sbToolBarIconos(tlbPrincipal, False)
 
 vPaso = True
    strSQL = "select RTRIM(cod_poliza) as 'IdX' , rtrim(DESCRIPCION) as 'ItmX' from CRD_CATALOGO_POLIZAS "
    Call sbCbo_Llena_New(cboPolizaLinea, strSQL, False, True)
 vPaso = False
 Call cboPolizaLinea_Click
 
 vScroll = False
    FlatScrollBar.Value = 0
 vScroll = True
 
 Call Formularios(Me)
 
 Call sbLimpia
 
 
 If Operacion.OperacionConsulta > 0 Then
    txtOperacion.Text = Operacion.OperacionConsulta
    Call txtOperacion_KeyPress(vbKeyReturn)
 End If
 
 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
 End With
Call RefrescaTags(Me)

End Sub


Private Sub sbCalculaDatosPoliza()
Dim curMonto As Currency, iPagos As Integer
Dim curCuota As Currency, curPago As Currency, curCuotaPlzRst As Currency
Dim iMeses As Integer, iCobertura As Integer

If vPaso Then Exit Sub

On Error GoTo vError

curMonto = CCur(txtPolizaMonto.Text)

Select Case cboPolizaPagoFrecuencia.Text
  Case "Mensual"
    iPagos = 1
  Case "Trimestral"
    iPagos = 4
  Case "Semestral"
    iPagos = 2
  Case "Anual"
    iPagos = 1
  Case "Indefinido"
    iPagos = 1
End Select

'Meses de Vigencia entre Pagos
iMeses = DateDiff("m", dtpPolizaFechaPago.Value, dtpPolizaCoberturaCorte.Value) + 1
If iMeses <= 0 Then iMeses = 1

iPagos = iMeses / iPagos
iCobertura = DateDiff("m", dtpPolizaCoberturaInicio.Value, dtpPolizaCoberturaCorte.Value) + 1

If iCobertura <= 0 Then iCobertura = 1
If iPagos <= 0 Then iPagos = 1

curPago = curMonto / iPagos
curCuota = curMonto / vMesesPenPlan
curCuotaPlzRst = curMonto / iCobertura


txtPolizaCoberturaMeses.Text = iCobertura
txtPolizaPagosNum.Text = iPagos
txtPolizaPagoMonto.Text = Format(curPago, "Standard")
txtPolizaCuota.Text = Format(curCuota, "Standard")

'Actualiza la Cuota al Resto del Plazo solo si está en 0 para no caerle encima a la definida por el usuario
If CCur(txtPolizaCuotaRestoPlazo.Text) = 0 Then
    txtPolizaCuotaRestoPlazo.Text = Format(curCuotaPlzRst, "Standard")
End If

vError:

End Sub

Private Sub sbLimpia(Optional pNuevo As Boolean = False)

vPaso = True
  
tcAux.Item(0).Selected = True



If fraPolizaRetencion.Visible Then
  cboPolizaOperacion.Visible = True
  txtObservaciones = ""
  txtPlazo.Text = "1"
  txtMonto.Text = "0"
  txtPagado.Text = "0"
  txtPendiente.Text = "0"
  txtProyectado.Text = "0"
  txtDocumento = ""
  txtFecha = ""
  txtEstado = ""
  txtPlazoTrasnscurrido = ""
  
  cboMes.Text = fxConvierteMES(Val(Mid(fxPrimerDeduccion, 5, 2)))
  txtAnio.Text = Mid(fxPrimerDeduccion, 1, 4)
    
Else
  
  
  txtPolizaId.Enabled = True
  txtPolizaId.Text = 0
  
  chkPolizaPlazoCredito.Value = vbUnchecked
  
  txtPolizaMonto.Text = 0
  txtPolizaCuota.Text = 0
  txtPolizaPagoMonto.Text = 0
  txtPolizaCuotaRestoPlazo.Text = 0
    
    
  cboPolizaEstado.Clear
  cboPolizaEstado.AddItem "Activa"
  cboPolizaEstado.AddItem "Inactiva"
  cboPolizaEstado.Text = "Activa"
  
  cboPolizaPagoFrecuencia.Clear
  cboPolizaPagoFrecuencia.AddItem "Mensual"
  cboPolizaPagoFrecuencia.AddItem "Trimestral"
  cboPolizaPagoFrecuencia.AddItem "Semestral"
  cboPolizaPagoFrecuencia.AddItem "Anual"
  cboPolizaPagoFrecuencia.AddItem "Indefinida"
  cboPolizaPagoFrecuencia.Text = "Mensual"
  
  dtpPolizaFechaPago.Value = fxFechaServidor
  dtpPolizaCoberturaInicio.Value = dtpPolizaFechaPago.Value
  dtpPolizaCoberturaCorte.Value = dtpPolizaCoberturaInicio.Value

  txtPolizaCoberturaMeses.Text = 1
  txtPolizaPagosNum.Text = 1
  txtPolizasCtaDeduce.Text = 1
  txtPolizasCtaDeduce.Tag = 1
  txtPolizasCtaDeduce.ToolTipText = "Línea de Inicio en el Plan: [x]"
  
  lswPolizaPago.ListItems.Clear
  lblPolizaPagoProximo.Caption = ""
  lblPolizaPagoSaldo.Caption = "0.00"
  
  lswPolizaRecaudado.ListItems.Clear
  lblPolizaRecaudadoSaldo.Caption = "0.00"
  lblPolizaRecaudadoCorte.Caption = ""
End If


If pNuevo Then
  If fraPolizaRetencion.Visible Then
    cboPolizaOperacion.Visible = False
  Else
    txtPolizaId.Enabled = False
  End If
End If
 
StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
 
vPaso = False
 
 
End Sub

Private Function fxOperacionDestino(vDestino As String) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select rtrim(cod_destino) + ' - ' + descripcion as ItemX from catalogo_destinos where cod_destino = '" & vDestino & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxOperacionDestino = " -"
Else
  fxOperacionDestino = rs!itemx
End If
rs.Close

End Function


Private Sub sbCargaPoliza(pOperacion As Long)

Dim strSQL As String, rs As New ADODB.Recordset
Dim vTemp As String

On Error GoTo vError

strSQL = "select R.id_solicitud,R.codigo,C.descripcion,R.cedula,S.nombre,R.cuota,R.estado" _
       & " ,R.observacion,R.fechaforp,R.plazo,R.amortiza,R.cuotas_planilla,R.cuotas_directas" _
       & " ,R.documento_referido,R.prideduc,R.userRec,R.cod_destino,R.garantia" _
       & " , RTRIM(isnull(Gt.DESCRIPCION,'')) as 'GarantiaDesc'" _
       & " , RTRIM(isnull(Cd.DESCRIPCION,'')) as 'DestinoDesc'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula " _
       & "  left join CRD_GARANTIA_TIPOS Gt on R.GARANTIA = Gt.GARANTIA " _
       & "  left join CATALOGO_DESTINOS Cd on R.COD_DESTINO = Cd.COD_DESTINO" _
       & " where R.Estado in('A','C') and C.poliza = 'S' and R.id_solicitud = " & pOperacion
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 
 Call sbCboAsignaDato(cboPolizaLinea, Trim(rs!Descripcion), True, Trim(rs!Codigo))
 
 txtMonto = Format(rs!Cuota, "Standard")
 txtPlazo = rs!Plazo
 
 txtEstado = fxEstadoCuota(rs!Estado)
 
 txtObservaciones = IIf(IsNull(rs!observacion), "", rs!observacion)
 
 txtFecha = Format((rs!FechaForp & ""), "dd/mm/yyyy")
 txtPlazoTrasnscurrido = rs!cuotas_planilla + rs!cuotas_directas
 
 txtDocumento = rs!documento_referido & ""
  
 
 txtPagado.Text = Format(rs!Amortiza, "Standard")
 If rs!Plazo >= 999 Then
    txtProyectado.Text = Format(rs!Cuota, "Standard")
    txtPendiente = Format(rs!Cuota, "Standard")
 Else
    txtProyectado.Text = Format(rs!Cuota * rs!Plazo, "Standard")
    txtPendiente = Format((rs!Cuota * rs!Plazo) - rs!Amortiza, "Standard")
 End If
 
 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = False
   .Item(3).Enabled = False
 End With
' Me.fraOperacion.Enabled = False

 If IsNull(rs!PriDeduc) Then
  cboMes.Text = fxConvierteMES(Val(Mid(fxPrimerDeduccion, 5, 2)))
  txtAnio.Text = Mid(fxPrimerDeduccion, 1, 4)
 Else
  cboMes.Text = fxConvierteMES(Val(Mid(rs!PriDeduc, 5, 2)))
  txtAnio.Text = Mid(rs!PriDeduc, 1, 4)
 End If


 'Carga Destino
 Call sbSTCargaCboDestinos(cboDestino, rs!Codigo)
 Call sbCboAsignaDato(cboDestino, rs!DestinoDesc, True, rs!cod_destino & "")


 'Carga Garantía
 vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, rs!Codigo)
        Call sbCboAsignaDato(cboGarantia, rs!GarantiaDesc, False, rs!Garantia)
 vPaso = False


 StatusBarX.Panels(1).Text = rs!userRec & ""
 StatusBarX.Panels(2).Text = Format(rs!FechaForp, "dd/mm/yyyy")

Else
 
 MsgBox "No existe esta Operación..?", vbExclamation

End If
rs.Close

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTemp As String, itmX As ListViewItem


On Error GoTo vError

Call sbLimpia

strSQL = "select R.id_solicitud,R.codigo,C.descripcion,R.cedula,S.nombre,R.cuota,R.estado,R.FechaForP, R.Plazo" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula " _
       & " where R.Estado in('A','C') and (C.retencion = 'N' or C.poliza = 'N') and R.id_solicitud = " & txtOperacion
       
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 
 txtCedula = rs!Cedula
 txtNombre = rs!Nombre
 txtCodigo = rs!Codigo
 txtDescripcion = rs!Descripcion
 
 vPlazo = rs!Plazo
 vFechaFormaliza = rs!FechaForp
 
Else
 
 MsgBox "No existe la Operación de crédito..?", vbExclamation
End If
rs.Close


'Carga Listado de Polizas Asignadas
vPaso = True
      
strSQL = "select Cat.cod_poliza,Cat.Descripcion,Cat.integra_plan_pagos,Pol.*" _
       & ",isnull(Reg.estado,'A') as 'OperacionEstado',isnull(Reg.cuota,0) as 'OperacionCuota'" _
       & " From CRD_OPERACION_POLIZAS Pol inner join Crd_Catalogo_Polizas Cat on Pol.cod_poliza = Cat.cod_poliza" _
       & " left join reg_creditos Reg on Pol.id_solicitud_poliza = Reg.id_solicitud" _
       & " Where Pol.ID_SOLICITUD = " & txtOperacion.Text
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 
 Set itmX = lsw.ListItems.Add(, , rs!cod_poliza)
     itmX.SubItems(1) = rs!Descripcion
 
 If rs!integra_plan_pagos = 0 Then
     itmX.SubItems(2) = "Retencion"
     itmX.SubItems(3) = rs!Id_Solicitud_Poliza
     itmX.SubItems(4) = Format(rs!OperacionCuota, "Standard")
     itmX.SubItems(5) = IIf(rs!OperacionEstado = "A", "Activa", "Cancelada")
     cboPolizaOperacion.AddItem rs!Id_Solicitud_Poliza
 
 Else
     itmX.SubItems(2) = "Integrado"
     itmX.SubItems(3) = rs!Num_Poliza
     itmX.SubItems(4) = Format(rs!Cuota, "Standard")
     itmX.SubItems(5) = IIf(rs!Estado = "A", "Activa", "Inactiva")

 End If
 rs.MoveNext
Loop
rs.Close

If GLOBALES.SysPlanPagos = 1 Then
        'Carga Plan de Pago pendiente
        strSQL = "select id_seq,num_cuota,Fecha_Inicio,Fecha_corte,Fecha_Pago from crd_operacion_plan_pagos where id_solicitud = " & txtOperacion.Text _
               & " and estado not in('C') and num_cuota > 0 and num_cuota_madre = 0 order by ID_SEQ"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
         cboPolizaPlan.AddItem "No.Cuota: " & rs!num_cuota & "   Fecha Pago: " & Format(rs!Fecha_Pago, "dd/mm/yyyy") & " (" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " : " & Format(rs!fecha_corte, "dd/mm/yyyy") & ")"
         cboPolizaPlan.ItemData(cboPolizaPlan.ListCount - 1) = CStr(rs!Id_seq)
         rs.MoveNext
        Loop
        
        If rs.RecordCount > 0 Then
           rs.MoveFirst
           cboPolizaPlan.Text = "No.Cuota: " & rs!num_cuota & "   Fecha Pago: " & Format(rs!Fecha_Pago, "dd/mm/yyyy") & " (" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " : " & Format(rs!fecha_corte, "dd/mm/yyyy") & ")"
        End If
        rs.Close
End If


vPaso = False

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbGuardaPolizaPlanPagos()
Dim strSQL As String, rs As New ADODB.Recordset, vFecha As Date
Dim curMonto As Currency, vFrecuencia As String, vSeqCorte As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
vSeqCorte = cboPolizaPlan.ItemData(cboPolizaPlan.ListIndex) + CInt(txtPolizasCtaDeduce.Text)

Select Case cboPolizaPagoFrecuencia.Text
  Case "Mensual"
    vFrecuencia = "M"
  Case "Trimestral"
    vFrecuencia = "T"
  Case "Semestral"
    vFrecuencia = "S"
  Case "Anual"
    vFrecuencia = "A"
  Case "Indefinida"
    vFrecuencia = "I"
  Case Else
    vFrecuencia = "M"
End Select

If txtPolizaId.Text = 0 Then
 'Registro de la Poliza Asociada
 
 strSQL = "select isnull(max(NUM_POLIZA),0) + 1 as 'Linea' from CRD_OPERACION_POLIZAS" _
        & " where id_solicitud = " & txtOperacion.Text
 Call OpenRecordSet(rs, strSQL)
 txtPolizaId.Text = rs!Linea
 rs.Close
 
 strSQL = "INSERT CRD_OPERACION_POLIZAS(id_solicitud_poliza,cod_poliza,id_solicitud,codigo,cuota,registro_fecha,registro_usuario,estado,NUM_POLIZA" _
        & ",Monto,COBERTURA_INICIO,COBERTURA_VENCE,PAGO_FRECUENCIA,PAGO_FECHA,PAGO_MONTO,PAGO_REALIZADO,PAGO_SALDO,PAGO_ULTIMO" _
        & ",RECAUDADO_MONTO,RECAUDADO_CORTE,RECAUDADO_SALDO,NUM_SEQ_INICIO,NUM_CTAS_DEDUCE,NUM_SEQ_CORTE,num_contrato" _
        & ",DEDUCE_PLAZO_CREDITO,CUOTA_RST_PLAN)" _
        & " values(0,'" & cboPolizaLinea.ItemData(cboPolizaLinea.ListIndex) & "'," & txtOperacion.Text & ",'" & txtCodigo.Text _
        & "'," & CCur(txtPolizaCuota.Text) & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & Mid(cboPolizaEstado.Text, 1, 1) _
        & "'," & txtPolizaId.Text & "," & CCur(txtPolizaMonto.Text) & ",'" & Format(dtpPolizaCoberturaInicio.Value, "yyyy/mm/dd") _
        & "','" & Format(dtpPolizaCoberturaCorte.Value, "yyyy/mm/dd") & "','" & vFrecuencia & "','" & Format(dtpPolizaFechaPago.Value, "yyyy/mm/dd") _
        & "'," & CCur(txtPolizaPagoMonto.Text) & ",0," & CCur(txtPolizaMonto.Text) & ",Null,0,dbo.MyGetdate()," & CCur(txtPolizaMonto.Text) _
        & "," & cboPolizaPlan.ItemData(cboPolizaPlan.ListIndex) & "," & txtPolizasCtaDeduce.Text & "," & vSeqCorte _
        & ",'" & txtPolizaContrato.Text & "'," & chkPolizaPlazoCredito.Value & "," & CCur(txtPolizaCuotaRestoPlazo.Text) & ")"
 Call ConectionExecute(strSQL)
Else
 strSQL = "update CRD_OPERACION_POLIZAS set estado = '" & Mid(cboPolizaEstado.Text, 1, 1) & "',cuota = " & CCur(txtPolizaCuota.Text) _
        & ", Monto = " & CCur(txtPolizaMonto.Text) & ", COBERTURA_INICIO = '" & Format(dtpPolizaCoberturaInicio.Value, "yyyy/mm/dd") _
        & "',COBERTURA_VENCE ='" & Format(dtpPolizaCoberturaCorte.Value, "yyyy/mm/dd") _
        & "',DEDUCE_PLAZO_CREDITO = " & chkPolizaPlazoCredito.Value _
        & ",CUOTA_RST_PLAN = " & CCur(txtPolizaCuotaRestoPlazo.Text) _
        & ",NUM_SEQ_INICIO = " & cboPolizaPlan.ItemData(cboPolizaPlan.ListIndex) _
        & " where id_Solicitud = " & txtOperacion.Text & " and Num_Poliza = " & txtPolizaId.Text
 Call ConectionExecute(strSQL)
      
End If

'Actualiza el Plan de Pagos con los datos de la Póliza
strSQL = "exec spCrdPolizaRegistroDetalle " & txtOperacion.Text & "," & txtPolizaId.Text & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Póliza registrada satisfactoriamente...", vbInformation
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGuardaPolizaRetencion()
Dim strSQL As String, rs As New ADODB.Recordset, vFecha As Date
Dim lngOP As Long, lngPriDeduc As Currency, vPolizaNum As Long
Dim vComite As Integer, vDestino As String, vPolizaLinea As String, vCargoCodigo As String


Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

On Error GoTo vError

lngPriDeduc = txtAnio.Text & Format(fxConvierteMES(cboMes.Text), "00")

vComite = fxCrdIdComiteLinea(txtCodigo.Text)

vDestino = cboDestino.ItemData(cboDestino.ListIndex)
If Trim(vDestino) = "" Then
  vDestino = "Null"
Else
  vDestino = "'" & vDestino & "'"
End If

'Informacion de Uso de la Poliza
strSQL = "select CODIGO_RETENCION,CODIGO_CARGO from CRD_CATALOGO_POLIZAS " _
       & " where COD_POLIZA = '" & cboPolizaLinea.ItemData(cboPolizaLinea.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
  vPolizaLinea = Trim(rs!codigo_retencion)
  vCargoCodigo = Trim(rs!codigo_cargo)
rs.Close

'Id de la Poliza
strSQL = "select isnull(max(NUM_POLIZA),0) + 1 as 'Linea' from CRD_OPERACION_POLIZAS" _
       & " where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
   vPolizaNum = rs!Linea
rs.Close
 
'Insertar la operacion
strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
       & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
       & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
       & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
       & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,documento_referido" _
       & ",cod_destino)" _
       & " values('" & UCase(vPolizaLinea) & "'," & vComite & ",'" _
       & Trim(txtCedula) & "'," & CCur(txtMonto.Text) & "," & CCur(txtMonto.Text) & ",0," & CCur(txtMonto.Text) & ",0,0," _
       & CCur(txtMonto.Text) & "," & CCur(txtMonto.Text) & ",0,0," & txtPlazo & ",'" & glogon.Usuario & "','" & glogon.Usuario _
       & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
       & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'" _
       & ",'N','OT','',0,1,0,'" & UCase(txtObservaciones) & "','A'," & lngPriDeduc _
       & "," & fxFechaProcesoAnterior(lngPriDeduc) & ",'F','" & txtDocumento & "'," & vDestino & ")"
  
 Call ConectionExecute(strSQL)
 lngOP = fxUltimaOperacion(txtCedula)
 
 
 
 'Registro de la Poliza Asociada
 strSQL = "INSERT CRD_OPERACION_POLIZAS(id_solicitud_poliza,cod_poliza, num_poliza,id_solicitud,codigo,cuota,registro_fecha,registro_usuario) values(" _
        & lngOP & ",'" & vPolizaLinea & "'," & vPolizaNum & "," & txtOperacion.Text & ",'" & txtCodigo.Text & "'," & CCur(txtMonto.Text) & ",dbo.MyGetdate(),'" _
        & glogon.Usuario & "')"
 Call ConectionExecute(strSQL)
 
 
 If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdPlanPagos " & lngOP
    Call ConectionExecute(strSQL)
 End If
 
 'Bitacora General
 Call Bitacora("Registra", "Retencion en la OP : " & lngOP)
 
 'Bitacora de Retenciones
 Call sbBitacoraCredito("08", "Op: " & lngOP & " - Monto " & CCur(txtMonto) _
        & " - Plazo: " & txtPlazo, "R", lngOP, UCase(txtCodigo))
 
 
 'TODO: Finalmente Revisar si Aplica Algun Cargo en la formalizacion para Abonarlo
 
 cboPolizaOperacion.AddItem lngOP
 
 
 Me.MousePointer = vbDefault
 
 MsgBox "Poliza Grabada Satisfactoriamente...", vbInformation

 Exit Sub
 
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub sbGuardar()
 
 If fraPolizaRetencion.Visible Then
    'Retencion de Poliza
    Call sbGuardaPolizaRetencion
 Else
    'Integrada al Plan de Pagos
    Call sbGuardaPolizaPlanPagos
 End If
End Sub





Private Sub sbPolizaListaAcreedores(pPoliza As String, pNumPoliza As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswAcreedores.ListItems.Clear
vPaso = True

strSQL = "select Acr.COD_ACREEDOR,Acr.IDENTIFICACION,Acr.NOMBRE,Acr.ACTIVO,Apl.registro_fecha,Apl.registro_usuario" _
       & " from CRD_POLIZAS_ACREEDORES Acr inner join CRD_POLIZAS_ACREEDOR_ASG Asg" _
       & " on Acr.cod_acreedor = Asg.cod_acreedor and Asg.cod_poliza = '" & pPoliza _
       & "' left join CRD_OPERACION_POLIZAS_ACREEDORES Apl on Acr.cod_acreedor = Apl.cod_acreedor" _
       & " and Apl.id_solicitud = " & txtOperacion.Text & " and Apl.num_poliza = " & txtPolizaId.Text _
       & " Where Acr.Activo = 1 order by Apl.registro_fecha desc, Acr.COD_ACREEDOR"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAcreedores.ListItems.Add(, , rs!Cod_Acreedor)
   itmX.SubItems(1) = rs!Identificacion
   itmX.SubItems(2) = rs!Nombre
   itmX.SubItems(3) = rs!Registro_Fecha & ""
   itmX.SubItems(4) = rs!Registro_Usuario & ""
   
   If Not IsNull(rs!Registro_Fecha) Then itmX.Checked = True

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

vPaso = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lsw.ListItems.Count = 0 Then Exit Sub

tcMain.Item(1).Selected = True

If Item.SubItems(2) = "Integrado" Then
    txtPolizaId.Text = Item.SubItems(3)
    Call sbPolizaPlanCarga
Else
    cboPolizaOperacion.Text = Item.SubItems(3)
End If

tlbPrincipal.Buttons(1).Enabled = True
tlbPrincipal.Buttons(2).Enabled = True
tlbPrincipal.Buttons(3).Enabled = False
tlbPrincipal.Buttons(4).Enabled = False

End Sub


Private Sub sbPolizaBeneficiarios()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswBeneficiarios.ListItems.Clear

strSQL = "Select * from CRD_OPERACION_POLIZAS_BENEFIARIOS" _
       & " Where id_Solicitud = " & txtOperacion.Text & " and num_Poliza = " & txtPolizaId.Text
Call OpenRecordSet(rs, strSQL)
     
Do While Not rs.EOF
   Set itmX = lswBeneficiarios.ListItems.Add(, , rs!Id_Beneficiario)
    itmX.SubItems(1) = rs!Nombre
    itmX.SubItems(2) = Format(rs!FechaNac, "dd/mm/yyyy")
    itmX.SubItems(3) = fxParentesco(rs!parentesco)
    itmX.SubItems(4) = rs!Porcentaje & " %"
   rs.MoveNext
Loop
     
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub lswAcreedores_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError


If Item.Checked Then
  strSQL = "insert CRD_OPERACION_POLIZAS_ACREEDORES(num_poliza,cod_acreedor,codigo,id_solicitud, registro_fecha,registro_usuario)" _
         & " values(" & txtPolizaId.Text & ",'" & Item.Text & "','" & txtCodigo.Text & "'," & txtOperacion.Text _
         & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
  strSQL = "delete CRD_OPERACION_POLIZAS_ACREEDORES where num_poliza = " & txtPolizaId.Text & " and id_solicitud = " & txtOperacion.Text _
         & " and cod_acreedor = '" & Item.Text & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Not IsNumeric(txtPolizaId.Text) Or txtPolizaId.Text = "0" Then Exit Sub

Select Case Item.Index
 Case 0 'General
 Case 1 'Pago
 Case 2 'Recaudacion
 Case 3 'Acreedores
    Call sbPolizaListaAcreedores(cboPolizaLinea.ItemData(cboPolizaLinea.ListIndex), txtPolizaId.Text)
 Case 4 'Beneficiarios
    Call sbPolizaBeneficiarios
End Select

End Sub

Private Sub tlbBeneficiarios_ButtonClick(ByVal Button As MSComctlLib.Button)
GLOBALES.gTag = txtOperacion.Text
GLOBALES.gTag2 = txtPolizaId.Text

Call sbFormsCall("frmCR_PolizasRegistroBeneficiarios", 1, , , False, Me)

End Sub


Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPlazo.SetFocus
End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto.Text = CCur(txtMonto.Text)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservaciones.SetFocus
End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
  
  Call sbLimpia(True)
  cboPolizaLinea.SetFocus
  
 Case "editar"
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
 
 
 Case "guardar"
 
  
  If fxVerifica Then
    Call sbGuardar
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    cboPolizaOperacion.Visible = True
  Else
    MsgBox vMensaje, vbCritical
  End If
  
 Case "deshacer"
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    cboPolizaOperacion.Visible = True

    If txtOperacion <> "" Then Call sbCargaOperacion
 
 Case "reportes"
    Call ReporteBoletaRetencion
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 End Select
End Sub



Private Sub ReporteBoletaRetencion()
Dim strRuta As String, rs As New ADODB.Recordset

On Error GoTo vError

strRuta = SIFGlobal.fxPathReportes("Credito_BoletaFormalizacion.rpt")
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta de Formalización / Pólizas"
   
 .Connect = glogon.ConectRPT
 
 .ReportFileName = strRuta
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD}=" & cboPolizaOperacion.Text
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 
 
 .SubreportToChange = "sbAsiento"
 .StoredProcParam(0) = "FRM"
 .StoredProcParam(1) = cboPolizaOperacion.Text
 .StoredProcParam(2) = 0
 
 
 .PrintReport
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub txtMonto_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
  If CInt(txtPlazo.Text) < 900 Then
      txtProyectado.Text = Format(CInt(txtPlazo.Text) * CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text) - CCur(txtPagado.Text), "Standard")
  Else
      txtProyectado.Text = Format(CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text), "Standard")
  End If
vError:
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
  txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
vError:
End Sub

Private Sub txtObservaciones_LostFocus()
 txtAnio.SetFocus
End Sub

Private Sub txtOperacion_Change()

 tcMain.Item(0).Selected = True
 
 txtCedula.Text = ""
 txtNombre.Text = ""
 txtCodigo.Text = ""
 txtDescripcion.Text = ""
 lsw.ListItems.Clear
 
 Call sbLimpia
  With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call sbCargaOperacion
End Sub


Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtPlazo_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
  If CInt(txtPlazo.Text) < 900 Then
      txtProyectado.Text = Format(CInt(txtPlazo.Text) * CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text) - CCur(txtPagado.Text), "Standard")
  Else
      txtProyectado.Text = Format(CCur(txtMonto.Text), "Standard")
      txtPendiente.Text = Format(CCur(txtProyectado.Text), "Standard")
  End If
vError:
End Sub


Private Sub txtPolizaCuota_GotFocus()
On Error GoTo vError
    txtPolizaCuota.Text = CCur(txtPolizaCuota.Text)
vError:
End Sub

Private Sub txtPolizaCuota_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPolizaEstado.SetFocus
End Sub

Private Sub txtPolizaCuota_LostFocus()
On Error GoTo vError
    txtPolizaCuota.Text = Format(CCur(txtPolizaCuota.Text), "Standard")
vError:
End Sub

Private Sub sbPolizaPlanCarga()
Dim strSQL As String, rs As New ADODB.Recordset

If Not IsNumeric(txtPolizaId.Text) Then Exit Sub

strSQL = "select Pol.*, rtrim(Cat.cod_poliza) as 'IdX', rtrim(Cat.descripcion) as 'ItmX'" _
       & " from CRD_OPERACION_POLIZAS Pol inner join CRD_CATALOGO_POLIZAS Cat on Pol.cod_poliza = Cat.cod_poliza" _
       & " where Pol.id_Solicitud = " & txtOperacion.Text & " and Pol.Num_Poliza = " & txtPolizaId.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   
   vPaso = True
   
   Call sbCboAsignaDato(cboPolizaLinea, rs!itmX, True, rs!idX)
   
   If rs!Estado = "A" Then
      cboPolizaEstado.Text = "Activa"
   Else
      cboPolizaEstado.Text = "Inactiva"
   End If
  
  txtPolizaContrato.Text = rs!Num_Contrato & ""
  txtPolizaMonto.Text = Format(rs!Monto, "Standard")
  txtPolizaCuota.Text = Format(rs!Cuota, "Standard")
  txtPolizaPagoMonto.Text = Format(rs!Pago_Monto, "Standard")
  
  txtPolizaCuotaRestoPlazo.Text = Format(rs!CUOTA_RST_PLAN, "Standard")
  chkPolizaPlazoCredito.Value = rs!DEDUCE_PLAZO_CREDITO
  Call chkPolizaPlazoCredito_Click
  
  Select Case rs!Pago_Frecuencia
    Case "M"
        cboPolizaPagoFrecuencia.Text = "Mensual"
    Case "T"
        cboPolizaPagoFrecuencia.Text = "Trimestral"
    Case "S"
        cboPolizaPagoFrecuencia.Text = "Semestral"
    Case "A"
        cboPolizaPagoFrecuencia.Text = "Anual"
    Case "I"
        cboPolizaPagoFrecuencia.Text = "Indefinida"
    Case Else
        cboPolizaPagoFrecuencia.Text = "Mensual"
  End Select

  
  dtpPolizaFechaPago.Value = rs!PAGO_FECHA
  dtpPolizaCoberturaInicio.Value = rs!COBERTURA_INICIO
  dtpPolizaCoberturaCorte.Value = rs!COBERTURA_VENCE

  txtPolizaCoberturaMeses.Text = DateDiff("m", rs!COBERTURA_INICIO, rs!COBERTURA_VENCE) + 1
  txtPolizaPagosNum.Text = 1
  txtPolizasCtaDeduce.Text = rs!NUM_CTAS_DEDUCE
  txtPolizasCtaDeduce.Tag = rs!NUM_SEQ_INICIO
  txtPolizasCtaDeduce.ToolTipText = "Línea de Inicio en el Plan: " & rs!NUM_SEQ_INICIO
  
  lswPolizaPago.ListItems.Clear
  lblPolizaPagoProximo.Caption = Format(rs!PAGO_FECHA, "dd/mm/yyyy")
  lblPolizaPagoSaldo.Caption = Format(rs!PAGO_SALDO, "Standard")
  
  lswPolizaRecaudado.ListItems.Clear
  lblPolizaRecaudadoSaldo.Caption = Format(rs!RECAUDADO_SALDO, "Standard")
  lblPolizaRecaudadoCorte.Caption = Format(rs!RECAUDADO_CORTE, "dd/mm/yyyy")
  
  
'  Call sbCalculaMesesPendientePlan
'  Call sbCalculaDatosPoliza

End If
rs.Close
       
vPaso = False
       
End Sub





Private Sub txtPolizaCuotaRestoPlazo_GotFocus()
On Error GoTo vError

txtPolizaCuotaRestoPlazo.Text = CCur(txtPolizaCuotaRestoPlazo.Text)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub txtPolizaCuotaRestoPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPolizaPagoFrecuencia.SetFocus
End Sub

Private Sub txtPolizaCuotaRestoPlazo_LostFocus()
On Error GoTo vError

txtPolizaCuotaRestoPlazo.Text = Format(CCur(txtPolizaCuotaRestoPlazo.Text), "Standard")

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub txtPolizaId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbPolizaPlanCarga
End Sub


Private Sub txtPolizaMonto_GotFocus()
On Error GoTo vError

txtPolizaMonto.Text = CCur(txtPolizaMonto.Text)

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub txtPolizaMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPolizaCuotaRestoPlazo.SetFocus
End Sub

Private Sub txtPolizaMonto_LostFocus()
On Error GoTo vError

txtPolizaMonto.Text = Format(CCur(txtPolizaMonto.Text), "Standard")

Call sbCalculaMesesPendientePlan
Call sbCalculaDatosPoliza

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub
