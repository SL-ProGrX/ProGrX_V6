VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_CorreccionCreditos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Corrección de Operaciones Activas"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10785
   HelpContextID   =   3013
   Icon            =   "frmCR_CorreccionCreditos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMora 
      Caption         =   "Cuotas Atrasadas"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   1440
      Width           =   10572
      Begin MSComctlLib.ListView lsw 
         Height          =   1416
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   10308
         _ExtentX        =   18177
         _ExtentY        =   2487
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha Pro"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha Sist."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Int.Cor."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Int.Mor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Amortización"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Cargos"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblMoraTexto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccione las Cuotas Morosas para Anulación y Luego Presione Aceptar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   252
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   10308
      End
   End
   Begin VB.Frame fraMov 
      Caption         =   "Movimientos:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4572
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   3252
      Begin MSComctlLib.ListView lswMov 
         Height          =   4092
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3012
         _ExtentX        =   5318
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   4657
         EndProperty
      End
   End
   Begin VB.Frame fraOpcion 
      Caption         =   "Aplicación:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4455
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   7212
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   852
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   2640
         Width           =   6972
      End
      Begin VB.Frame fraTasas 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   6375
         Begin VB.TextBox txtTasaTBPPuntos 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            Top             =   960
            Width           =   615
         End
         Begin VB.TextBox txtTasa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.CheckBox chkTasaPuntosRenuncia 
            Appearance      =   0  'Flat
            Caption         =   "Aplicar Puntos adicionales por Renuncia Interna"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   705
            Left            =   2880
            TabIndex        =   20
            Top             =   360
            Width           =   2055
         End
         Begin VB.CheckBox chkTasaIndizadaTBP 
            Appearance      =   0  'Flat
            Caption         =   "Indizada TBP + pp"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Puntos Add TBP : "
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   18
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Tasa Revisable : "
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ComboBox cboX 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1440
         Width           =   5055
      End
      Begin VB.CheckBox chkAjustePriDeduc 
         Appearance      =   0  'Flat
         Caption         =   "Ajustar Primer Deducción"
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
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   2295
      End
      Begin VB.VScrollBar vsBar 
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   720
         Width           =   270
      End
      Begin VB.TextBox txtCambio 
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
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin XtremeSuiteControls.PushButton cmdAceptar 
         Height          =   492
         Left            =   4920
         TabIndex        =   55
         Top             =   3720
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Aplicar"
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
         Picture         =   "frmCR_CorreccionCreditos.frx":000C
      End
      Begin XtremeSuiteControls.PushButton cmdAnuFormalizacion 
         Height          =   492
         Left            =   1680
         TabIndex        =   56
         Top             =   3720
         Width           =   3252
         _Version        =   1441793
         _ExtentX        =   5736
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "&Anular Formalización"
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
         Picture         =   "frmCR_CorreccionCreditos.frx":07E4
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nota del Cambio:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   54
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cambio por:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblOpcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   372
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6972
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2160
      TabIndex        =   57
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   14
      Left            =   600
      TabIndex        =   58
      Top             =   120
      Width           =   1332
   End
   Begin VB.Image imgExcluirCredito 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   4320
      Picture         =   "frmCR_CorreccionCreditos.frx":1171
      Stretch         =   -1  'True
      ToolTipText     =   "Excluye Esta operacion por Medio de Anulacion en la B.D."
      Top             =   120
      Width           =   252
   End
   Begin VB.Label lblDestino 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   312
      Left            =   6240
      TabIndex        =   52
      Top             =   1800
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   5400
      TabIndex        =   51
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblRecurso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   312
      Left            =   6240
      TabIndex        =   50
      Top             =   2160
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   5400
      TabIndex        =   49
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   10
      Left            =   5400
      TabIndex        =   48
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblOficina 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   312
      Left            =   6240
      TabIndex        =   47
      Top             =   2520
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ejecutivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   11
      Left            =   5400
      TabIndex        =   46
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblEjecutivo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   312
      Left            =   6240
      TabIndex        =   45
      Top             =   2880
      Width           =   4332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   9
      Left            =   2880
      TabIndex        =   44
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblGarantia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   3720
      TabIndex        =   43
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label txtUltMov 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   3720
      TabIndex        =   42
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   2880
      TabIndex        =   41
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "1° Deduc"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   6
      Left            =   2880
      TabIndex        =   40
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblPrideduc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   3720
      TabIndex        =   39
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   2880
      TabIndex        =   38
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   37
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   2
      Left            =   360
      TabIndex        =   36
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   35
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   34
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label lblMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   3720
      TabIndex        =   33
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblSaldo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   1200
      TabIndex        =   32
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblPlazo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   1200
      TabIndex        =   31
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblInteres 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   1200
      TabIndex        =   30
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label lblCuota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   315
      Left            =   1200
      TabIndex        =   29
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblPlazoRestante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   28
      ToolTipText     =   "Plazo Restante"
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label lblTasaOriginal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   2040
      TabIndex        =   27
      ToolTipText     =   "Tasa Original de Formalización"
      Top             =   2520
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Left            =   600
      TabIndex        =   24
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Index           =   12
      Left            =   600
      TabIndex        =   23
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label lblOpex 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   8640
      TabIndex        =   9
      Top             =   960
      Width           =   492
   End
   Begin VB.Label lblDescripcion 
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
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   3600
      TabIndex        =   1
      ToolTipText     =   "Descripción del Código"
      Top             =   960
      Width           =   5052
   End
   Begin VB.Label lblNombre 
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
      ForeColor       =   &H80000008&
      Height          =   312
      Left            =   3600
      TabIndex        =   0
      ToolTipText     =   "Nombre de la Persona"
      Top             =   600
      Width           =   5532
   End
   Begin VB.Label lblCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Código del Préstamo"
      Top             =   960
      Width           =   1452
   End
   Begin VB.Label lblCedula 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Cédula de la Persona"
      Top             =   600
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10812
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   26
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "frmCR_CorreccionCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vRetencion As Boolean

Private Function fxValidaCambio() As Boolean
Dim rs As New ADODB.Recordset

fxValidaCambio = True

If Len(txtNotas.Text) = 0 Then
    MsgBox "Especifique una Nota al movimiento...!", vbExclamation
    fxValidaCambio = False
    Exit Function
End If

If Not IsNumeric(lblOpcion.Tag) Then
    MsgBox "El movimiento no es válido...!", vbExclamation
    fxValidaCambio = False
    Exit Function
End If

Select Case CInt(lblOpcion.Tag)
  Case 1 'Tasas
    If IsNumeric(txtTasa.Text) And IsNumeric(txtTasaTBPPuntos.Text) Then
      txtTasa.Text = CCur(txtTasa.Text)
      txtTasaTBPPuntos.Text = CCur(txtTasaTBPPuntos.Text)
      
      If txtTasa.Text > 99 Then fxValidaCambio = False
    Else
    
      fxValidaCambio = False
    
    End If

  Case 2
    rs.CursorLocation = adUseServer
    rs.Open "select isnull(count(*),0) as Existe from catalogo where codigo = '" & txtCambio _
            & "'", glogon.Conection, adOpenStatic
    fxValidaCambio = IIf((rs!Existe = 1), True, False)
    txtCambio = UCase(txtCambio)
    rs.Close
  
  Case 11, 12, 13, 14, 16, 18, 19
     If cboX.Text = "" Then
        fxValidaCambio = False
     End If
  
  Case 5, 9, 10, 15, 17
    'Nada
    
  Case 8 'Abonos
'     If Not IsNumeric(txtIntereses.Text) Or Not IsNumeric(txtAmortizacion.Text) Then
'        fxValidaCambio = False
'     Else
'       If CCur(txtIntereses.Text) < 0 Or CCur(txtAmortizacion.Text) < 0 Then
'            fxValidaCambio = False
'       End If
'     End If
  
  Case Else
    If IsNumeric(txtCambio) Then
      If lblOpcion.Tag = 0 Then
        txtCambio = CLng(txtCambio)
      Else
        txtCambio = CCur(txtCambio)
      End If
      
      If lblOpcion.Tag = 1 And txtCambio > 99 Then fxValidaCambio = False
    Else
      fxValidaCambio = False
    End If
  
End Select

If vOperacion = 0 Then fxValidaCambio = False

End Function


Private Sub sbEliminaMora()
Dim strSQL As String, itmX As ListItem, lng As Long

On Error GoTo vError


With lsw.ListItems
    For lng = 1 To .Count
       If .Item(lng).Checked Then
          If GLOBALES.SysPlanPagos = 1 Then
            strSQL = "update CRD_OPERACION_TRANSAC set mora_dias = 0, intMor = 0 where ID_SEQ = " _
                   & .Item(lng).Text & " and id_solicitud = " & txtOperacion.Text
          
            strSQL = strSQL & Space(10) & "update CRD_OPERACION_PLAN_PAGOS set mora_dias = 0, intMor = 0 where ID_SEQ = " _
                   & .Item(lng).Text & " and id_solicitud = " & txtOperacion.Text
            Call ConectionExecute(strSQL)
          
          Else
            strSQL = "update morosidad set estado = 'N' where id_moro = " & .Item(lng).Text & " and Estado = 'A'"
            Call ConectionExecute(strSQL)
          End If
            
          strSQL = "Int.Mor..: " & .Item(lng).SubItems(4) & "   Dias..: " & .Item(lng).SubItems(7) & "    Notas..: " & txtNotas.Text
            
          Call sbBitacoraCredito("06", "Id..:" & .Item(lng).Text, "C", txtOperacion, lblCodigo.Caption, strSQL)
          Call Bitacora("Anula", "Morosidad OP: " & txtOperacion & " ID:" & .Item(lng).Text)
       End If
    Next lng
End With


MsgBox "Reversiones realizadas Satisfactoriamente..."
 
'optCorreccion(1).Value = True

fraMora.Visible = False


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub


Private Sub sbEliminaCargos()
Dim strSQL As String, itmX As ListItem, lng As Long

On Error GoTo vError

With lsw.ListItems
    For lng = 1 To .Count
       If .Item(lng).Checked Then
            If GLOBALES.SysPlanPagos = 1 Then
              'Con Plan de Pagos
                strSQL = "delete CRD_OPERACION_TRANSAC_CARGOS where Linea = " & .Item(lng).Text _
                       & " and id_seq = " & .Item(lng).SubItems(6) & " and id_solicitud = " & txtOperacion.Text
                
                strSQL = strSQL & Space(10) & "update CRD_OPERACION_TRANSAC set Cargos = Cargos - " & CCur(.Item(lng).SubItems(4)) _
                       & " where id_seq = " & .Item(lng).SubItems(6) & " and id_solicitud = " & txtOperacion.Text
            
                strSQL = strSQL & Space(10) & "update CRD_OPERACION_PLAN_PAGOS set Cargos = Cargos - " & CCur(.Item(lng).SubItems(4)) _
                       & " where id_seq = " & .Item(lng).SubItems(6) & " and id_solicitud = " & txtOperacion.Text
                
                Call ConectionExecute(strSQL)
            
            Else
               'Sin Plan de Pagos
                strSQL = "delete morosidad_cargos where id_cargo = " & .Item(lng).Text
                Call ConectionExecute(strSQL)
                
                strSQL = "update morosidad set Cargo = Cargo - " & CCur(.Item(lng).SubItems(4)) & " where id_moro = " & .Item(lng).SubItems(6)
                Call ConectionExecute(strSQL)
            End If
                
            Call sbBitacoraCredito("21", .Item(lng).SubItems(5), "C", txtOperacion, lblCodigo.Caption, "Monto..: " & .Item(lng).SubItems(4) + "   Id..: " & .Item(lng).Text & "    Notas..: " & txtNotas.Text)
            Call Bitacora("Elimina", "Cargos OP: " & txtOperacion & " Id:" & .Item(lng).Text & "Monto..:" & .Item(lng).SubItems(4))
       End If
    Next lng
End With
    
MsgBox "Reversión realizada Satisfactoriamente...", vbInformation
 
'optCorreccion(1).Value = True

fraMora.Visible = False


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub




Private Sub cmdAceptar_Click()
Dim iRespuesta As Integer, strSQL As String, rs As New ADODB.Recordset, vDH As String
Dim strDetalleBitacora As String, vCuenta As String, curMonto As Currency
Dim lngRecibo As Long, vTipo As String, vProceso As Long
Dim intTmp As Integer, vMontoRetencion As Currency
Dim vTipoDoc As String, vConcepto As String


If Not fxValidaCambio Then
  MsgBox "Error : El valor de cambio no es válido...", vbCritical
  Exit Sub
End If



Select Case CInt(lblOpcion.Tag)
  Case 5 'Elimina Cuotas Morosdas
      Call sbEliminaMora
      Exit Sub
  
  Case 8 'Abono Especial
'      Call sbAbonoEspecial
      Exit Sub
      
  Case 9 'Cambio de Fiadores
    If lblGarantia.Tag = "F" Then
      Operacion.Operacion = txtOperacion
      Operacion.Codigo = lblCodigo.Caption
      Call sbFormsCall("frmCR_SolicitudesFiadores", 1, , , False, Me)
      
      Exit Sub
    Else
        MsgBox "La Operación no tiene garantia Fiduciaria....verifique!", vbExclamation
        Exit Sub
    End If
  
  Case 15 'Elimina Cargos
      Call sbEliminaCargos
      Exit Sub
  
End Select


On Error GoTo vError

strDetalleBitacora = ""
iRespuesta = MsgBox("Esta seguro que desea " & lblOpcion.Caption & " de la Operación..: " & vOperacion, vbYesNo)

If iRespuesta = vbNo Or iRespuesta = vbCancel Then Exit Sub
 
 Select Case CInt(lblOpcion.Tag)
   
   Case 0 'Plazo
        
     vTipo = "OT"
     
     If chkAjustePriDeduc.Value = vbChecked Then
        'Cambia Primer Deducción; con Corre a Partir
        vProceso = fxFechaProcesoSiguiente(GLOBALES.glngFechaCR)
        
        strSQL = "update reg_Creditos set cuota_fija = 0, plazo = " & txtCambio & ",priDeduc = " & vProceso & ",cuota = " _
               & CCur(fxCalcula_Cuota(CCur(lblSaldo.Caption), txtCambio, lblInteres.Caption)) _
               & " where id_solicitud = " & vOperacion
        strDetalleBitacora = "REF:Ajusta Primer Deducción:" & vProceso
     
     Else
     
        'Calcular Meses de Ajuste
        'Se calcula el tiempo transcurrido y se le resta al nuevo plazo; para sacar tiempo de ajuste para la cuota
        If CLng(txtCambio.Text) <= (CLng(lblPlazo.Caption) - CLng(lblPlazoRestante.Caption)) Then
          MsgBox "El Plazo de Cambio, es igual o menor al tiempo (Plazo) transcurrido de la operación...", vbExclamation
          Exit Sub
        End If
        
        
        intTmp = CLng(txtCambio) - (CLng(lblPlazo.Caption) - CLng(lblPlazoRestante.Caption))
     
        strSQL = "update reg_Creditos set cuota_fija = 0,plazo = " & txtCambio & ",cuota = " _
               & CCur(fxCalcula_Cuota(CCur(lblSaldo.Caption), intTmp, lblInteres.Caption)) _
               & " where id_solicitud = " & vOperacion
        strDetalleBitacora = "REF: No Ajusta Primer Deduccción"
     End If
     Call ConectionExecute(strSQL)
     
     Call sbBitacoraCredito("01", ("De: " & lblPlazo.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
     
     strDetalleBitacora = strDetalleBitacora & "Cambia Plazo De " & lblPlazo.Caption & " a " & txtCambio & " OP " & vOperacion
     
     If GLOBALES.SysPlanPagos = 1 Then
        'Actualiza Tabla de Pagos
        strSQL = "exec spCrdPlanPagos " & vOperacion
        Call ConectionExecute(strSQL)
     End If
     
     
     
   Case 1 'Interes
     
     
     vTipo = "OT"
     strSQL = "update reg_Creditos set interesv = " & txtTasa.Text & ",cuota = " _
            & CCur(fxCalcula_Cuota(CCur(lblSaldo.Caption), lblPlazoRestante.Caption, txtTasa.Text)) _
            & ",Cuota_Fija = 0, TBP_PuntosAdd = " & IIf((chkTasaIndizadaTBP.Value = vbChecked), txtTasaTBPPuntos.Text, "Null") _
            & ",LiqTasa = " & chkTasaPuntosRenuncia.Value _
            & " where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
        
     strDetalleBitacora = "REF: Plazo Restante:" & lblPlazoRestante.Caption
     
     txtNotas.Text = "Plz.Rest: " & lblPlazoRestante.Caption & " --- Tasa : " & lblInteres.Caption & " -> " & txtTasa.Text _
                   & " --- TBP+pp : " & txtTasaTBPPuntos.Tag & " -> " & txtTasaTBPPuntos.Text & " --- Apl.Liq+pp : " & chkTasaPuntosRenuncia.Value _
                   & Space(20) & vbCrLf & "--- Nota: " & txtNotas.Text
     
     Call sbBitacoraCredito("02", ("De: " & lblInteres.Caption & " A: " & txtTasa.Text), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
     
     strDetalleBitacora = strDetalleBitacora & "Cambia Interes De " & lblInteres.Caption & " a " & txtTasa.Text & " OP " & vOperacion
     
     If GLOBALES.SysPlanPagos = 1 Then
        'Actualiza Tabla de Pagos
        strSQL = "exec spCrdPlanPagos " & vOperacion & ",1"
        Call ConectionExecute(strSQL)
     End If
     
     
   Case 2 'Línea
     
     strSQL = "exec spCrd_Operacion_Cambio_Linea " & vOperacion & ",'" & txtCambio.Text & "','" _
            & Mid(txtNotas.Text, 1, 500) & "','" & glogon.Usuario & "'"
     
     Call OpenRecordSet(rs, strSQL)
     
     If glogon.error Then
        Me.MousePointer = vbDefault
        Exit Sub
     Else
        vTipo = rs!TipoDoc
        lngRecibo = rs!NumDoc
     End If
     
     strDetalleBitacora = "Cambia Línea De " & lblCodigo.Caption & " a " & txtCambio & " OP " & vOperacion
    
   Case 3 'Monto
     'No aplica con Plan de Pagos
     If GLOBALES.SysPlanPagos = 1 Then
        Exit Sub
     End If
    
     
     If CCur(lblMonto.Caption) = CCur(txtCambio) Then Exit Sub
       
     If CCur(lblMonto.Caption) < CCur(txtCambio) Then
        curMonto = CCur(txtCambio) - CCur(lblMonto.Caption)
        vDH = "D"
     Else
        curMonto = CCur(txtCambio) - CCur(lblMonto.Caption)
        curMonto = Abs(curMonto)
        vDH = "H"
     End If
     
    
     If CCur(lblMonto.Caption) < CCur(txtCambio) Then 'ND : AUMENTA EL MONTO
       
       If Not vRetencion Then
            vTipo = "ND"
            vConcepto = "CRD012"
            vTipoDoc = "ND"
            vCuenta = Trim(fxDocumentoCuenta(vTipo))
            
            If vAseDocValido = False Then
                Me.MousePointer = vbDefault
                MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                      & " válida para esta operación...", vbCritical
                Exit Sub
            End If
            
            lngRecibo = 0
            
            strSQL = "update reg_Creditos set montoapr = " & txtCambio & ",cuota = " _
                   & CCur(fxCalcula_Cuota(txtCambio, lblPlazo.Caption, lblInteres.Caption)) _
                   & ",saldo = " & txtCambio & " - AMORTIZA" _
                   & " where id_solicitud = " & vOperacion
            Call ConectionExecute(strSQL)
     
     
            If uRecibos Then lngRecibo = fxDocumentoAbono4(curMonto, "ND", vCuenta, vDH)
            
            strDetalleBitacora = "Cambia Monto De:" & CCur(lblMonto.Caption) & " A:" _
                               & txtCambio & " OP:" & vOperacion & "-ND:" & lngRecibo
            
            strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
                   & "fechas,fechap,tcon,ncon,estado,cod_concepto,usuario,cod_caja) values('" & lblCodigo.Caption _
                   & "'," & txtOperacion & ",0,0,0," & CCur(txtCambio) - CCur(lblMonto.Caption) & ",dbo.MyGetdate()" _
                   & "," & GLOBALES.glngFechaCR & ",'" & vTipoDoc & "','" & lngRecibo & "','A','" & vConcepto & "','" & glogon.Usuario & "','')"
            Call ConectionExecute(strSQL)
        Else
          'Es una retención
           vMontoRetencion = CCur(txtCambio.Text) / CInt(lblPlazo.Caption)
           
            If CInt(lblPlazoRestante.Caption) <= 0 Then
                strSQL = "update reg_Creditos set montoapr = " & vMontoRetencion & ",cuota = " _
                       & CCur(txtCambio) & " - Amortiza, saldo = " & CCur(txtCambio) & " - Amortiza" _
                       & " where id_solicitud = " & vOperacion
            Else
                strSQL = "update reg_Creditos set montoapr = " & vMontoRetencion & ",cuota = (" _
                       & CCur(txtCambio) & " - Amortiza) / " & CInt(lblPlazoRestante.Caption) & " , saldo = " & CCur(txtCambio) & " - Amortiza" _
                       & " where id_solicitud = " & vOperacion
            End If
            Call ConectionExecute(strSQL)
     
            strDetalleBitacora = "Cambia Monto De:" & CCur(lblMonto.Caption) & " A:" _
                               & txtCambio & " OP:" & vOperacion & "-Ajuste"
          
        End If
    
     Else 'NC : DISMINUYE EL MONTO
        If Not vRetencion Then
            vTipo = "NC"
            vCuenta = Trim(fxDocumentoCuenta("NC"))
            vConcepto = "CRD012"
            If GLOBALES.SysDocVersion = 1 Then
               vTipoDoc = "7"
            Else
               vTipoDoc = "NC"
            End If
            
            If vAseDocValido = False Then
                Me.MousePointer = vbDefault
                MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                      & " válida para esta operación...", vbCritical
                Exit Sub
            End If
            
            lngRecibo = 0
            
            strSQL = "update reg_Creditos set montoapr = " & CCur(txtCambio) & ",cuota = " _
                   & CCur(fxCalcula_Cuota(txtCambio, lblPlazo.Caption, lblInteres.Caption)) _
                   & ",saldo = " & CCur(txtCambio) & "-AMORTIZA" _
                   & " where id_solicitud = " & vOperacion
            Call ConectionExecute(strSQL)
            
            
            If uRecibos Then lngRecibo = fxDocumentoAbono4(curMonto, "NC", vCuenta, vDH)
            
            strDetalleBitacora = "Cambia Monto De:" & CCur(lblMonto.Caption) & " A:" _
                               & txtCambio & " OP:" & vOperacion & "-NC:" & lngRecibo
            strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza," _
                   & "fechas,fechap,tcon,ncon,estado,cod_concepto,usuario,cod_caja) values('" & lblCodigo.Caption _
                   & "'," & txtOperacion & ",0,0,0," & CCur(lblMonto.Caption) - CCur(txtCambio) & ",dbo.MyGetdate()," & GLOBALES.glngFechaCR _
                   & ",'" & vTipoDoc & "','" & lngRecibo & "','A','" & vConcepto & "','" & glogon.Usuario & "','')"
            Call ConectionExecute(strSQL)
        
        Else
          'Es una retención
          If CCur(txtCambio.Text) >= CCur(lblSaldo.Caption) Then
                If CInt(lblPlazoRestante.Caption) <= 0 Then
                    strSQL = "update reg_Creditos set montoapr = " & vMontoRetencion & ",cuota = " _
                           & CCur(txtCambio) & " - Amortiza, saldo = " & CCur(txtCambio) & " - Amortiza" _
                           & " where id_solicitud = " & vOperacion
                Else
                    strSQL = "update reg_Creditos set montoapr = " & vMontoRetencion & ",cuota = (" _
                           & CCur(txtCambio) & " - Amortiza) / " & CInt(lblPlazoRestante.Caption) & " , saldo = " & CCur(txtCambio) & " - Amortiza" _
                           & " where id_solicitud = " & vOperacion
                End If
                Call ConectionExecute(strSQL)
           Else
              MsgBox "El nuevo monto es inferior al monto ya amortizado, esta operación no es válida para retenciones...!", vbExclamation
           End If 'Validacion del Monto a disminuir de la retencion
        End If
      
      
     End If
   
     'Actualiza Tabla de Pagos
     If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrdPlanPagos " & vOperacion
        Call ConectionExecute(strSQL)
     End If
     Call sbBitacoraCredito("09", ("De: " & lblMonto.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)

  
   Case 4 'Cuota
     
     If Not IsNumeric(txtCambio.Text) Then
        Exit Sub
     End If
     
     vTipo = "OT"
     If vRetencion Then
        strSQL = "update reg_Creditos set cuota = " & CCur(txtCambio.Text) & ", Cuota_Fija = " & CCur(txtCambio.Text) _
               & ",montoapr = " & CCur(txtCambio.Text) & ",saldo = " & CCur(txtCambio.Text) _
               & " where id_solicitud = " & vOperacion
     Else
        strSQL = "update reg_Creditos set cuota = " & CCur(txtCambio.Text) & ", Cuota_Fija = " & CCur(txtCambio.Text) _
               & " where id_solicitud = " & vOperacion
     End If
     Call ConectionExecute(strSQL)
     
     'Actualiza Tabla de Pagos
     If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrdPlanPagos " & vOperacion
        Call ConectionExecute(strSQL)
     End If
     
 
     Call sbBitacoraCredito("10", ("De: " & CCur(lblCuota.Caption) & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
 
     strDetalleBitacora = "Cambia Cuota De " & CCur(lblCuota.Caption) & " a " & txtCambio & " OP " & vOperacion
 
   Case 6 'Ultimo Abono
      
     vTipo = "OT"
      
     strSQL = "update reg_Creditos set fecult = " & txtCambio _
            & " where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
    
    If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrd_Operacion_Cambio_UltCta " & vOperacion & "," & txtCambio.Text & ",'" _
           & Mid(txtNotas.Text, 1, 500) & "','" & glogon.Usuario & "'"
        
        Call ConectionExecute(strSQL)
    End If

     Call sbBitacoraCredito("04", ("De: " & txtUltMov & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
     
     strDetalleBitacora = "Fecha Ult. Abono DE " & txtUltMov & " A " & txtCambio & " OP " & vOperacion
 
   Case 7 'Primer Deducción
      
     vTipo = "OT"
     
     lngRecibo = 0
     strSQL = "update reg_Creditos set prideduc = " & txtCambio _
            & " where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
 
     'Actualiza Tabla de Pagos
     If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrdPlanPagos " & vOperacion
        Call ConectionExecute(strSQL)
     End If
 
 
     Call sbBitacoraCredito("05", ("De: " & lblPrideduc.Caption & " A: " & txtCambio), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
 
     strDetalleBitacora = "Fecha 1er Deducción DE " & lblPrideduc.Caption & " A " & txtCambio & " OP " & vOperacion
 
   
   Case 10 'Limpia Intereses Moratorios
     vTipo = "OT"
      
     'Actualiza Tabla de Pagos
     If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "update Tra Set IntMor = 0, Mora_Dias = 0" _
               & " from CRD_OPERACION_TRANSAC Tra inner join Reg_Creditos Reg on Tra.id_solicitud = Reg.id_solicitud" _
               & " where Reg.id_solicitud = " & vOperacion & " and Reg.estado = 'A' and Reg.proceso <> 'J'" _
               & " and Tra.Estado = 'A'"
        Call ConectionExecute(strSQL)
      
        strSQL = "update Tra Set IntMor = 0, Mora_Dias = 0" _
               & " from CRD_OPERACION_PLAN_PAGOS Tra inner join Reg_Creditos Reg on Tra.id_solicitud = Reg.id_solicitud" _
               & " where Reg.id_solicitud = " & vOperacion & " and Reg.estado = 'A' and Reg.proceso <> 'J'" _
               & " and Tra.Estado = 'A'"
        Call ConectionExecute(strSQL)
      
     Else
        strSQL = "update morosidad set intm = 0 where id_solicitud = " & vOperacion & " and estado = 'A' and estadoi <> 'J'"
        Call ConectionExecute(strSQL)
     End If
     
     Call sbBitacoraCredito("12", "** Limpia Int.Mor. ***", IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
  
     strDetalleBitacora = "Elimina Intereses Moratorios OP: " & vOperacion
 
 
   Case 11 'Cambio de Garantía
      
     vTipo = "OT"
     
     lngRecibo = 0
     strSQL = "update reg_Creditos set Garantia = '" & fxGarantia(cboX.Text) _
            & "' where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
 
     Call sbBitacoraCredito("16", ("De: " & lblGarantia.Caption & " A: " & cboX.Text), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
 
     strDetalleBitacora = "Garantia : " & lblGarantia.Caption & " A " & cboX.Text & " OP " & vOperacion
 
 
   Case 12 'Cambio de Destino
      
     vTipo = "OT"
     
     lngRecibo = 0
     strSQL = "update reg_Creditos set cod_destino = '" & SIFGlobal.fxCodText(cboX.Text) _
            & "' where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
 
     Call sbBitacoraCredito("18", ("De: " & lblDestino.Tag & " A: " & SIFGlobal.fxCodText(cboX.Text)), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
 
     strDetalleBitacora = "Destino : " & lblDestino.Tag & " A " & SIFGlobal.fxCodText(cboX.Text) & " OP " & vOperacion
 
   Case 13 'Cambio de Recurso
     
     vTipo = "OT"
     
     lngRecibo = 0
     strSQL = "update reg_Creditos set cod_Grupo = '" & SIFGlobal.fxCodText(cboX.Text) _
            & "' where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
 
     Call sbBitacoraCredito("19", ("De: " & lblRecurso.Tag & " A: " & SIFGlobal.fxCodText(cboX.Text)), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
 
     strDetalleBitacora = "Recurso : " & lblRecurso.Tag & " A " & SIFGlobal.fxCodText(cboX.Text) & " OP " & vOperacion
 
   
   Case 14 'Cambio de Dia de Pago
     
     vTipo = "OT"
     
     lngRecibo = 0
     strSQL = "update reg_Creditos set Dia_Pago = " & cboX.ItemData(cboX.ListIndex) & " where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)

     'Actualiza Tabla de Pagos
     If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrdOperacionCambioDiaPago " & vOperacion & "," & cboX.ItemData(cboX.ListIndex) & ",'" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
     End If

     Call sbBitacoraCredito("20", ("De: " & lblMonto.Tag & " A: " & cboX.ItemData(cboX.ListIndex)), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)

     strDetalleBitacora = "Día de Pago : " & lblMonto.Tag & " A " & cboX.ItemData(cboX.ListIndex) & " OP " & vOperacion
 
 
 
   Case 16 'Cambio de Oficina
     
     vTipo = "ND"
     
     lngRecibo = fxDocumentoChOficina(vTipo)
     strSQL = "update reg_Creditos set cod_oficina_r = '" & SIFGlobal.fxCodText(cboX.Text) & "' where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)

     Call sbBitacoraCredito("21", ("De: " & lblMonto.Tag & " A: " & SIFGlobal.fxCodText(cboX.Text)), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)

     strDetalleBitacora = "Oficina : " & lblMonto.Tag & " A " & SIFGlobal.fxCodText(cboX.Text) & " OP " & vOperacion
 
 
   Case 17 'Ajuste de Cuota Bullet
     
     vTipo = "OT"
     
     lngRecibo = 0
     Operacion.OperacionConsulta = vOperacion
     frmCR_OperacionCtaBullet.Show vbModal
        
     strDetalleBitacora = "Ingreso para Cambio/Ajuste de Cuota Bullet...Operacion..:" & vOperacion
 
 
   Case 18 'Cambio de Actividad Economica
     
     strSQL = "update reg_Creditos set cod_actividad = '" & SIFGlobal.fxCodText(cboX.Text) & "' where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)

     Call sbBitacoraCredito("22", ("De: " & txtCambio.Text & " A: " & SIFGlobal.fxCodText(cboX.Text)), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)

     strDetalleBitacora = "Actividad : " & txtCambio.Text & " A " & SIFGlobal.fxCodText(cboX.Text) & " OP " & vOperacion
 
 
   Case 19 ' Cambio de Ejecutivo
     
     vTipo = "OT"
     
     lngRecibo = 0
     strSQL = "update reg_Creditos set id_promotor = " & cboX.ItemData(cboX.ListIndex) _
            & " where id_solicitud = " & vOperacion
     Call ConectionExecute(strSQL)
 
     Call sbBitacoraCredito("24", ("De: " & lblEjecutivo.Tag & " A: " & cboX.ItemData(cboX.ListIndex)), IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption, txtNotas.Text)
 
     strDetalleBitacora = "Ejecutivo : " & lblEjecutivo.Tag & " A " & cboX.ItemData(cboX.ListIndex) & " OP " & vOperacion
 
 End Select
 
 Call Bitacora("Aplica", strDetalleBitacora)
 
 Call sbCargaOperacion
 
 txtCambio = ""
  
Select Case vTipo
  Case "ND"
    MsgBox "Cambio Realizado Satisfactoriamente..." _
      & vbCrLf & "Se Generó Nota No. " & lngRecibo, vbInformation
      If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "ND")
  Case "NC"
    MsgBox "Cambio Realizado Satisfactoriamente..." _
      & vbCrLf & "Se Generó Nota No. " & lngRecibo, vbInformation
     If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "NC")
  Case "OT"
    MsgBox "Cambio Realizado Satisfactoriamente..."
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Function fxVerificaAnulacion(xOperacion As Long) As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, rsTmp As New ADODB.Recordset
Dim vFecha As Date

vMensaje = ""
fxVerificaAnulacion = True


'Busca y Elimina en refundiciones > Inconsistencia de Registro de Operacion = Refundicion
'Antes de continuar con la anulacion
strSQL = "delete refundiciones where (id_solicitud = id_solicitudr) and id_solicitud = " & xOperacion
Call ConectionExecute(strSQL)


'0. Verificacion base
strSQL = "select fechaforp,isnull(estado,'N') as Estado,estadosol from reg_creditos where id_solicitud = " & xOperacion
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!FechaForp
    If rs!Estado <> "A" Then
      vMensaje = vMensaje & vbCrLf & "- Esta operación no se encuentra activa..."
    End If
rs.Close

'1. Verifica Nivel de Anulacion (Usuario)

strSQL = "select isnull(count(*),0) as Existe" _
   & " from NIVEL_GRUPOS N INNER JOIN nivel_miembros A" _
   & " ON N.NV_COD_GRUPO = A.NV_COD_GRUPO INNER JOIN nivel_derechos B" _
   & " ON N.NV_COD_GRUPO = B.NV_COD_GRUPO Where A.nombre = '" & glogon.Usuario _
   & "' and B.codigo = '" & lblCodigo.Caption & "' AND N.nv_tipo = 'N'" _
   & " and (" & CCur(lblMonto.Caption) & " between nv_desde and nv_hasta)"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  vMensaje = vMensaje & vbCrLf & "- No existe nivel de anulación de este usuario para el código " & lblCodigo.Caption
End If
rs.Close

'2. Verifica que no se le registren desembolsos, Se deben de anular o eliminar
strSQL = "select isnull(count(*),0) as Existe from Tes_Transacciones where op = " & xOperacion _
       & " and estado <> 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  vMensaje = vMensaje & vbCrLf & "- Existen solicitudes o documentos emitidos (Cheques/Transferencias) en Tesorería (Proceda a Anularlos)"
End If
rs.Close

If GLOBALES.SysPlanPagos = 0 Then
    '3. Verificar si se le han realizado movimientos a la Operacion despues de su formalizacion
    strSQL = "select isnull(count(*),0) as Existe from creditos_dt where id_solicitud = " & xOperacion _
           & " and  ncon <> '" & xOperacion & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
    End If
    rs.Close
    
    'No tiene porque tener ningun registro de mora
    strSQL = "select isnull(count(*),0) as Existe from MOROSIDAD where id_solicitud = " & xOperacion
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
    End If
    rs.Close
    
    
    '3a. Verificar si se le han realizado movimientos a las refundiciones (Abonadas o Canceladas)
    strSQL = "select id_solicitud,consec from creditos_dt where tcon in('3','FRM') and ncon = '" & xOperacion & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     strSQL = "select isnull(count(*),0) as Existe from creditos_dt where id_solicitud = " _
            & rs!Id_Solicitud & " and consec > " & rs!consec
     Call OpenRecordSet(rsTmp, strSQL, 0)
        If rsTmp!Existe > 0 Then
          vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a la op:" & rs!Id_Solicitud _
                   & " posterior a su refundicion"
        End If
     rsTmp.Close
     rs.MoveNext
    Loop
    rs.Close
    
    '3a. a la fecha de formalizacion (Doble verificacion para movimientos en mora no reflejados)
    strSQL = "select isnull(count(*),0) as Existe from creditos_dt" _
            & " where fechas > '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "' and id_solicitud in(select id_solicitud" _
            & " from creditos_dt where tcon in('3','FRM') and ncon = '" & xOperacion & "' and  id_solicitud <> " & xOperacion & ")"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a refundiciones posterior a la formalizacion"
    End If
    rs.Close
    
    '3b. a la fecha de formalizacion para Morosidad
    strSQL = "select isnull(count(*),0) as Existe from morosidad" _
            & " where Estado = 'C' and fecUlt > '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "' and id_solicitud in(select id_solicitud" _
            & " from morosidad where estado = 'C' and tcon in('3','FRM') and ncon = '" & xOperacion & "'  and id_solicitud = " & xOperacion & ")"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a Mora de refundiciones posterior a la formalizacion"
    End If
    rs.Close

End If 'SysPlanPagos = 0

'4. No puede anular retenciones
strSQL = "select retencion from catalogo where codigo = '" & lblCodigo.Caption & "'"
Call OpenRecordSet(rs, strSQL)
If rs!retencion = "S" Then
   vMensaje = vMensaje & vbCrLf & "- Este es un Línea de retención No se puede Anular..."
End If
rs.Close


If Len(vMensaje) > 0 Then
 MsgBox vMensaje, vbExclamation
 fxVerificaAnulacion = False
End If

End Function


Private Sub cmdAnuFormalizacion_Click()
Dim strSQL As String, rs As New ADODB.Recordset, vMensaje As String


If lblCodigo.Caption = "" Then Exit Sub

Me.MousePointer = vbHourglass

'Verifica Anulacion
If Not fxVerificaAnulacion(txtOperacion) Then
 Me.MousePointer = vbDefault
 Exit Sub
End If


On Error GoTo vError

strSQL = "exec spCRDFormalizaAnulacion " & txtOperacion & ",'" & glogon.Usuario & "',0"
Call OpenRecordSet(rs, strSQL)
  vMensaje = rs!Mensaje & ""
rs.Close


'BITACORA
Call Bitacora("Registra", "Anulación de la OP: " & Operacion.Operacion)
Call sbBitacoraCredito("13", "Monto : " & lblMonto.Caption, IIf(vRetencion, "R", "C"), txtOperacion.Text, lblCodigo.Caption, "SGT Anula Formalización: " & txtNotas.Text)
''Tags de Seguimiento se Aplica desde el procedure.
'Call sbCrdOperacionTags(txtOperacion.Text, lblCodigo.Caption, "S09", "", "SGT Anula Formalización: " & txtNotas.Text)

Me.MousePointer = vbDefault

MsgBox "Anulación de formalización realizada satisfactoriamente...", vbInformation
If GLOBALES.SysDocVersion = 2 Then
    Call sbImprimeRecibo(txtOperacion.Text, "AFR")
End If

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer, itmX As ListItem

vOperacion = 0

On Error Resume Next

strSQL = "select R.*,S.nombre,C.descripcion,C.retencion,C.poliza,Eje.Nombre as 'Ejecutivo'" _
       & ",dbo.fxCrdPlazoRestante(R.Plazo,R.prideduc," & GLOBALES.glngFechaCR & ") as 'PlazoRestante'" _
       & ",X.descripcion as 'DestinoX',Y.descripcion as 'RecursoX',Ofi.Descripcion as 'OficinaX'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula " _
       & " left Join Catalogo_destinos X on R.cod_destino = X.cod_destino" _
       & " left Join Catalogo_grupos Y on R.cod_grupo = Y.cod_grupo" _
       & " left join SIF_Oficinas Ofi on R.cod_oficina_r = Ofi.cod_Oficina" _
       & " left join Promotores Eje on R.id_Promotor = Eje.id_Promotor" _
       & " where R.estado = 'A' and R.id_solicitud =" & txtOperacion
If GLOBALES.SysPlanPagos = 0 Then
  strSQL = strSQL & " and R.proceso <> 'J'"
End If

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 lblCedula.Caption = rs!Cedula
 lblCodigo.Caption = rs!Codigo
 lblNombre.Caption = rs!Nombre
 lblDescripcion.Caption = rs!Descripcion
 lblOpex.Caption = IIf((rs!opex = 1), "OPEX", "")

 lblMonto.Caption = Format(rs!montoapr, "Standard")
 lblMonto.Tag = rs!dia_pago
 
 lblSaldo.Caption = Format(rs!Saldo, "Standard")
 lblSaldo.Tag = rs!Base_Calculo
 lblCuota.Caption = Format(rs!Cuota, "Standard")
 txtUltMov = IIf(IsNull(rs!FecUlt), 0, rs!FecUlt)
 lblPrideduc.Caption = IIf(IsNull(rs!PriDeduc), 0, rs!PriDeduc)
 lblPlazo.Caption = rs!Plazo
 lblInteres.Caption = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  
 If rs!PlazoRestante <= 0 Then
    lblPlazoRestante.Caption = 1
 Else
    lblPlazoRestante.Caption = rs!PlazoRestante
 End If
 
 lblTasaOriginal.Caption = rs!Int & ""
 
 vOperacion = rs!Id_Solicitud
 
 lblGarantia.Tag = rs!Garantia
 lblGarantia.Caption = fxGarantia(rs!Garantia)
 
 lblDestino.Caption = rs!DestinoX & ""
 lblDestino.Tag = rs!cod_destino & ""
 
 lblRecurso.Caption = rs!RecursoX & ""
 lblRecurso.Tag = rs!Cod_Grupo & ""
 
 lblOficina.Caption = rs!OficinaX & ""
 lblOficina.Tag = rs!cod_oficina_r & ""
 
 lblEjecutivo.Caption = rs!Ejecutivo & ""
 lblEjecutivo.Tag = rs!ID_PROMOTOR & ""
 
 'Tasas
 txtTasa.Text = rs!interesv
 
 If IsNull(rs!TBP_PuntosAdd) Then
    txtTasaTBPPuntos.Text = 0
    txtTasaTBPPuntos.Tag = 0
    
    chkTasaIndizadaTBP.Value = vbUnchecked
 Else
    txtTasaTBPPuntos.Text = rs!TBP_PuntosAdd
    txtTasaTBPPuntos.Tag = rs!TBP_PuntosAdd
    
    chkTasaIndizadaTBP.Value = vbChecked
 End If
 
 If IsNull(rs!LiqTasa) Then
    chkTasaPuntosRenuncia.Value = vbUnchecked
    chkTasaPuntosRenuncia.Tag = 0
 Else
    chkTasaPuntosRenuncia.Value = rs!LiqTasa
    chkTasaPuntosRenuncia.Tag = rs!LiqTasa
 End If
 
 
 If rs!retencion = "S" Or rs!Poliza = "S" Then
   vRetencion = True
   
    With lswMov.ListItems
      .Clear
     
         Set itmX = .Add(, "i3", "Cambio de Línea")
     
     If rs!Plazo < 900 Then
         Set itmX = .Add(, "i4", "Cambio de Monto")
     
        lblSaldo.Caption = Format((rs!Cuota * rs!Plazo) - rs!Amortiza, "Standard")
        lblMonto.Caption = Format(rs!Cuota * rs!Plazo, "Standard")
     
     End If
     
         Set itmX = .Add(, "i5", "Cambio de Cuota")
         If GLOBALES.SysPlanPagos = 0 Then Set itmX = .Add(, "i6", "Elimina Cuotas en Mora")
         Set itmX = .Add(, "i7", "Cambio de último abono")
         Set itmX = .Add(, "i8", "Cambio Primer Deducción")
        
         Set itmX = .Add(, "i12", "Cambio de Garantía")
         Set itmX = .Add(, "i13", "Cambio de Destino")
         Set itmX = .Add(, "i14", "Cambio de Recurso")
         Set itmX = .Add(, "i15", "Cambio de Día de Pago")
         Set itmX = .Add(, "i17", "Cambio de Oficina")
    
    
    End With
   
   
   
 Else 'No es Retencion aplican todas las funciones
 
   vRetencion = False
    With lswMov.ListItems
         .Clear
        Set itmX = .Add(, "i1", "Cambio de Plazo")
        Set itmX = .Add(, "i2", "Cambio de Tasa")
        Set itmX = .Add(, "i3", "Cambio de Línea")
            If GLOBALES.SysPlanPagos = 0 Then Set itmX = .Add(, "i4", "Cambio de Monto") 'Opciones no soportadas en Plan de Pagos
        Set itmX = .Add(, "i5", "Cambio de Cuota")
        Set itmX = .Add(, "i6", "Elimina Cuotas en Mora") 'Opciones no soportadas en Plan de Pagos
        Set itmX = .Add(, "i7", "Cambio de último abono")
        Set itmX = .Add(, "i8", "Cambio Primer Deducción")
       ' Set itmX = .Add(, "i9", "Abonos Especiales") 'Se elimina Opción porque se traslada a Arreglos de Pago: 2016/10/19
        Set itmX = .Add(, "i10", "Cambio de Fiadores")
        Set itmX = .Add(, "i11", "Elimina Intereses Moratorios")
        Set itmX = .Add(, "i12", "Cambio de Garantía")
        Set itmX = .Add(, "i13", "Cambio de Destino")
        Set itmX = .Add(, "i14", "Cambio de Recurso")
        Set itmX = .Add(, "i15", "Cambio de Día de Pago")
        Set itmX = .Add(, "i16", "Elimina Cargos Registrados") 'Opciones no soportadas en Plan de Pagos
        Set itmX = .Add(, "i17", "Cambio de Oficina")
        If rs!Base_Calculo = "04" Then
            Set itmX = .Add(, "i18", "Ajuste de Cuota Bullet/Ballon")
        End If
        
        Set itmX = .Add(, "i19", "Cambio de Actividad")
        Set itmX = .Add(, "i20", "Cambio de Ejecutivo")
        
    End With
 
 End If
 
Else
 vOperacion = 0
 MsgBox "La operación no se encontró o está cancelada, o Pertenece a un codigo de Retencion...", vbInformation
End If

rs.Close

End Sub

Private Sub Form_Load()
Dim itmX As ListItem

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)

Call txtOperacion_Change
Call RefrescaTags(Me)

fraOpcion.Enabled = False
fraTasas.Visible = False
imgExcluirCredito.Visible = IIf((cmdAceptar.Tag = 1), True, False)



End Sub



Private Sub imgExcluirCredito_Click()
Dim iRespuesta As Integer, strSQL As String
Dim vOB As String, lngRecibo As Long, vCuenta As String
Dim vFecha As Date, vTipoDoc As String, vTipoMov As String

If vOperacion = 0 Then
    MsgBox "Ingrese un número de operacion válido...", vbInformation
    Exit Sub
End If

vOB = ""

iRespuesta = MsgBox("Esta seguro que desea Excluir del Sistema la OP " & vOperacion, vbYesNo)

lngRecibo = 0
vTipoDoc = "NC"
vTipoMov = "NC"


If iRespuesta = vbYes Then
  
vFecha = fxFechaServidor
  
  
  If Not vRetencion Then
        
        vCuenta = Trim(fxDocumentoCuenta(vTipoDoc))
        
        If vAseDocValido = False Then
            Me.MousePointer = vbDefault
            MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
                  & " válida para esta operación...", vbCritical
            Exit Sub
        End If
        
        If uRecibos Then lngRecibo = fxDocumentoAbono(CCur(lblSaldo.Caption), "NC", vCuenta)
        
        vOB = "SE EXCLUYE CON N.C. # " & lngRecibo
       
        If GLOBALES.SysPlanPagos = 1 Then
                strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD011','" & glogon.Usuario & "','NC'" _
                       & ",'" & lngRecibo & "',0,0," & CCur(lblSaldo.Caption) _
                       & ",0,'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','',1"
                Call ConectionExecute(strSQL)
        Else
            'Sin Plan de Pagos
            strSQL = "Update reg_creditos set estado = 'C',SALDO=0,AMORTIZA=MONTOAPR" _
                   & ", observacion = observacion + '" & vOB & "'" _
                   & " where id_solicitud = " & vOperacion
            Call ConectionExecute(strSQL)
            
            strSQL = "INSERT CREDITOS_DT(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS,FECHAP,TCON,NCON,ESTADO,cod_concepto,usuario,cod_Caja)" _
                   & "VALUES('" & lblCodigo.Caption & "'," & vOperacion & ",0," & CCur(lblSaldo.Caption) & ",0," & CCur(lblSaldo.Caption) _
                   & ",dbo.MyGetdate()," & GLOBALES.glngFechaCR & ",'" & vTipoMov & "'," _
                   & IIf((lngRecibo = 0), "null", lngRecibo) & ",'A','CRD011','" & glogon.Usuario & "','')"
            Call ConectionExecute(strSQL)
        End If
    
       
        
        Call sbBitacoraCredito("14", "Saldo :" & lblSaldo.Caption, IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
        Call Bitacora("Aplica", "Exclusión de la operación :" & vOperacion)
        
        If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "NC")
        
        MsgBox "Exclusión aplicada con Nota de Crédito # " & lngRecibo, vbInformation
        vOperacion = 0
  
  Else 'Es una Retencion
        
        strSQL = "Update reg_creditos set estado = 'C', MONTOAPR = AMORTIZA, SALDO = 0" _
               & " where id_solicitud = " & vOperacion
               
        strSQL = strSQL & Space(10) & "DELETE CRD_OPERACION_PLAN_PAGOS WHERE ID_SOLICITUD = " & vOperacion & " AND ESTADO IN('A','P')"
        strSQL = strSQL & Space(10) & "DELETE CRD_OPERACION_TRANSAC    WHERE ID_SOLICITUD = " & vOperacion & " AND ESTADO IN('A')"
        Call ConectionExecute(strSQL)

        Call sbBitacoraCredito("07", "Monto :" & lblSaldo.Caption, IIf(vRetencion, "R", "C"), txtOperacion, lblCodigo.Caption)
        
        Call Bitacora("Aplica", "Exclusión de la operación :" & vOperacion)
        
        MsgBox "Exclusión aplicada y Guardada en Bitacora Creditos...", vbInformation
        vOperacion = 0
  
  End If 'Retencion

End If

End Sub

Private Function fxDocumentoAbono(curAmortiza As Currency, pTipoDoc As String, vCuenta As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, strCliente As String
Dim pConcepto As String

pConcepto = "CRD011" 'Exclusion de la Operacion

lngRecibo = fxDocumentoConsecutivo(pTipoDoc)
fxDocumentoAbono = lngRecibo

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
    strLinea(1) = "Saldo Anterior    " & Format(rs!Saldo, "Standard")
    strLinea(2) = "Interes Corriente " & "0.00"
    strLinea(3) = "Interes Moratorio " & "0.00"
    strLinea(4) = "Amortizacion      " & Format(curAmortiza, "Standard")
    strLinea(5) = "Saldo Actual      " & Format(rs!Saldo - curAmortiza, "Standard")
    strLinea(6) = "Operación         " & txtOperacion & ".." & lblCodigo.Caption & ".." & UCase(lblOpex.Caption)
    strLinea(7) = "Divisa: " & rs!COD_DIVISA & " / Tipo Cambio.:" & rs!TipoCambio
    strLinea(8) = "Descripción       " & lblDescripcion.Caption
    strLinea(9) = "Usuario           " & glogon.Usuario
    strLinea(10) = "EXCLUYE   "

    curAmortiza = rs!Saldo 'Actualiza la varible por si se realizaron cambios al saldo



'Control de Documentos v2
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
        & " values('" & lngRecibo & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(lblCedula.Caption) _
        & "','" & Trim(lblNombre.Caption) & "','" & pConcepto & "'," & curAmortiza & ",'P','" & txtOperacion.Text _
        & "','" & lblCodigo.Caption & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'ASIENTO
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza * rs!TipoCambio & ",'C','" & rs!COD_DIVISA _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
         & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)

  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza * rs!TipoCambio & ",'D','" & rs!COD_DIVISA _
         & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
         & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)



rs.Close


End Function


Private Function fxDocumentoAbono2(pTipoDoc As String, vCuenta As String, vDetalle As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, strCliente As String
Dim pConcepto As String, curMonto As Currency

pConcepto = "CRD012" 'Ajuste de Condiciones
curMonto = 1

lngRecibo = fxDocumentoConsecutivo(pTipoDoc)
fxDocumentoAbono2 = lngRecibo

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

strLinea(1) = lblOpcion.Caption
strLinea(2) = vDetalle
strLinea(3) = ""
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = "Operación         " & txtOperacion & ".." & lblCodigo.Caption & ".." & UCase(lblOpex.Caption)
strLinea(7) = "Divisa.:" & rs!COD_DIVISA & " / Tipo Cambio: " & rs!TipoCambio
strLinea(8) = "Descripción       " & lblDescripcion.Caption
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""



'Control de Documentos v2
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
      & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
      & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
      & " values('" & lngRecibo & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(lblCedula.Caption) _
      & "','" & Trim(lblNombre.Caption) & "','" & pConcepto & "'," & rs!Saldo & ",'P','" & txtOperacion.Text _
      & "','" & lblCodigo.Caption & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
      & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
      & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
      & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
      & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'ASIENTO
strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curMonto * rs!TipoCambio & ",'D','" & rs!COD_DIVISA _
       & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
       & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
Call ConectionExecute(strSQL)

strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curMonto * rs!TipoCambio & ",'C','" & rs!COD_DIVISA _
       & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
       & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
Call ConectionExecute(strSQL)

rs.Close


End Function


Private Function fxDocumentoAbono4(curAmortiza As Currency, pTipoDoc As String, vCuenta, vDH As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, rsTmp As New ADODB.Recordset, strCliente As String
Dim pConcepto As String

pConcepto = "CRD012" 'Cambio de Monto

lngRecibo = fxDocumentoConsecutivo(pTipoDoc)
fxDocumentoAbono4 = lngRecibo


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

strLinea(1) = "CAMBIO DE MONTO"
strLinea(2) = "DE " & CCur(lblMonto.Caption)
strLinea(3) = "A " & CCur(txtCambio)
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = "Operación         " & txtOperacion & ".." & lblCodigo.Caption & ".." & UCase(lblOpex.Caption)
strLinea(7) = "Divisa.:" & rs!COD_DIVISA & " / Tipo Cambio: " & rs!TipoCambio
strLinea(8) = "Descripción       " & lblDescripcion.Caption
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""

'Control de Documentos v2
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
      & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
      & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
      & " values('" & lngRecibo & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(lblCedula.Caption) _
      & "','" & Trim(lblNombre.Caption) & "','" & pConcepto & "'," & curAmortiza & ",'P','" & txtOperacion.Text _
      & "','" & lblCodigo.Caption & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
      & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
      & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
      & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
      & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'ASIENTO
strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza * rs!TipoCambio & ",'" & vDH & "','" & rs!COD_DIVISA _
       & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
       & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
Call ConectionExecute(strSQL)

strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & curAmortiza * rs!TipoCambio & ",'" & IIf((vDH = "D"), "C", "D") & "','" & rs!COD_DIVISA _
       & "'," & rs!TipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
       & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
Call ConectionExecute(strSQL)


rs.Close

End Function





Private Function fxDocumentoChOficina(pTipoDoc As String) As Long
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim lngRecibo As Long, pConcepto As String, rsTmp As New ADODB.Recordset

'Cambio de Oficina / Agencia

pConcepto = "CRD012" 'Cambio de Condiciones

lngRecibo = fxDocumentoConsecutivo(pTipoDoc)
fxDocumentoChOficina = lngRecibo

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
    strLinea(1) = "Saldo          " & Format(rs!Saldo, "Standard")
    strLinea(2) = "Oficina Actual " & lblOficina.Tag & "_" & lblOficina.Caption
    strLinea(3) = "Oficina Nueva  " & cboX.Text
    strLinea(4) = ""
    strLinea(5) = ""
    strLinea(6) = "Operación      " & txtOperacion
    strLinea(7) = "Línea          " & lblCodigo.Caption & "-" & UCase(lblOpex.Caption)
    strLinea(8) = "Descripción    " & lblDescripcion.Caption
    strLinea(9) = "Usuario        " & glogon.Usuario
    strLinea(10) = "Cambia de Oficina/Agencia"

If GLOBALES.SysDocVersion = 2 Then
    'Control de Documentos v2
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
                & " values('" & lngRecibo & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(lblCedula.Caption) _
                & "','" & Trim(lblNombre.Caption) & "','" & pConcepto & "'," & rs!Saldo & ",'P','" & txtOperacion.Text _
                & "','" & lblCodigo.Caption & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "')"
        Call ConectionExecute(strSQL)
        
        'ASIENTO
        strSQL = "select cod_unidad,cod_centro_costo from sif_oficinas where cod_oficina = '" & SIFGlobal.fxCodText(cboX.Text) & "'"
        Call OpenRecordSet(rsTmp, strSQL, 0)
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & rs!Saldo & ",'D','" & rs!COD_DIVISA _
                 & "',1," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rs!ctaamortiza _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        rsTmp.Close
        
          strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngRecibo & "'," & rs!Saldo & ",'C','" & rs!COD_DIVISA _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
                 & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
          Call ConectionExecute(strSQL)
        
End If
rs.Close


End Function



Private Sub sbLlenaMorosidad()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListItem

On Error Resume Next

If txtOperacion = "" Then Exit Sub
If Not IsNumeric(txtOperacion) Then Exit Sub


lblMoraTexto.Caption = "Seleccione las Cuotas Morosas para Anulación y Luego Presione Aceptar"
fraMora.Caption = "Cuotas Atrasadas"

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear


lsw.ColumnHeaders.Add , , "ID", 900
lsw.ColumnHeaders.Add , , "Proceso", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Fecha", 1200
lsw.ColumnHeaders.Add , , "Int.Cor", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Int.Mor.", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Principal", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Cargos", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Dias", 1200, vbCenter

rs.CursorLocation = adUseServer

If GLOBALES.SysPlanPagos = 1 Then
   strSQL = "select ID_SEQ as 'ID_MORO',FECHA_PROCESO as 'FECHAP',FECHA_PAGO as 'FECULT',INTCOR AS 'INTC', INTMOR AS 'INTM'" _
          & ",CARGOS AS 'CARGO', PRINCIPAL AS 'AMORTIZA',MORA_DIAS as 'DIAS'" _
          & " From CRD_OPERACION_TRANSAC" _
          & " where MORA_DIAS > 0 AND ESTADO = 'A' AND ID_SOLICITUD = " & txtOperacion.Text

Else
    strSQL = "select *, 'N/A' as 'DIAS' from morosidad where estado = 'A' and id_solicitud = " & txtOperacion _
           & " Order by fechap desc"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!id_moro)
     itmX.SubItems(1) = Format(rs!fechap, "####-##")
     itmX.SubItems(2) = Format(rs!FecUlt, "dd/mm/yyyy")
     itmX.SubItems(3) = Format(rs!IntC, "Standard")
     itmX.SubItems(4) = Format(rs!IntM, "Standard")
     itmX.SubItems(5) = Format(rs!Amortiza, "Standard")
     itmX.SubItems(6) = Format(rs!Cargo, "Standard")
     itmX.SubItems(7) = Format(rs!Dias, "Standard")
 rs.MoveNext
Loop
rs.Close
End Sub



Private Sub sbLlenaCargos()
Dim rs As New ADODB.Recordset, strSQL As String, itmX As ListItem

On Error Resume Next

If txtOperacion = "" Then Exit Sub
If Not IsNumeric(txtOperacion) Then Exit Sub


lblMoraTexto.Caption = "Seleccione los Cargos a Eliminar y Luego Presione Aceptar"
fraMora.Caption = "Cargos Registrados"

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "ID", 900
lsw.ColumnHeaders.Add , , "Proceso", 1000, vbCenter
lsw.ColumnHeaders.Add , , "Fecha", 1200
lsw.ColumnHeaders.Add , , "Usuario", 1200
lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Detalle", 3200
lsw.ColumnHeaders.Add , , "ID.Mora", 1000


rs.CursorLocation = adUseServer
If GLOBALES.SysPlanPagos = 1 Then
   strSQL = "select M.ID_SEQ as 'ID_MORO',M.FECHA_PROCESO as 'FECHAP',FECHA_PAGO as 'FECULT',INTCOR AS 'INTC', INTMOR AS 'INTM'" _
          & ",M.CARGOS AS 'CARGO', M.PRINCIPAL AS 'AMORTIZA',C.LINEA as 'ID_CARGO',C.MONTO,C.USUARIO,C.FECHA,C.Detalle as 'DESCRIPCION'" _
          & " from CRD_OPERACION_TRANSAC M inner join CRD_OPERACION_TRANSAC_CARGOS C on M.ID_SOLICITUD = C.ID_SOLICITUD and M.ID_SEQ = C.ID_SEQ" _
          & " where M.MORA_DIAS > 0 AND M.ESTADO = 'A' AND C.MOV_MONTO = 0 AND M.ID_SOLICITUD = " & txtOperacion.Text

Else
    strSQL = "select C.*,G.DESCRIPCION,M.FechaP" _
           & " from MOROSIDAD_CARGOS C inner join CBR_GESTIONES G on C.COD_GESTION = G.COD_GESTION" _
           & " inner join Morosidad M on M.id_Moro = C.id_Moro" _
           & " where M.ESTADO = 'A' and M.ID_SOLICITUD = " & txtOperacion.Text
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Id_Cargo)
     itmX.SubItems(1) = Format(rs!fechap, "####-##")
     itmX.SubItems(2) = Format(rs!fecha, "dd/mm/yyyy")
     itmX.SubItems(3) = rs!Usuario
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = rs!Descripcion
     itmX.SubItems(6) = rs!id_moro
 rs.MoveNext
Loop
rs.Close
End Sub



Private Sub lswMov_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset, x As Byte
Dim Index As Integer


On Error GoTo vError

Index = CInt(Mid(Item.Key, 2, 2)) - 1

'Efecto Visual
For x = 1 To lswMov.ListItems.Count - 1
   lswMov.ListItems.Item(x).Bold = False
Next x
Item.Bold = True


fraMora.top = 1440
fraMora.Visible = False

fraOpcion.Enabled = True
lblOpcion.Caption = Item.Text
lblOpcion.Tag = Index

'Inicia Visualizacion
txtCambio.Visible = False
fraTasas.Visible = False

vsBar.Visible = False

chkAjustePriDeduc.Value = vbUnchecked
chkAjustePriDeduc.Visible = False

'txtIntereses.Visible = False
'txtAmortizacion.Visible = False
'lblCaptionAbono(0).Visible = False
'lblCaptionAbono(1).Visible = False

cboX.Visible = False

Select Case Index
  Case 1 'Intereses
    fraTasas.Visible = True
  
  Case 0, 2, 3, 4, 6, 7
     'Cambio con Texto
     txtCambio.Text = ""
     txtCambio.Visible = True
     txtCambio.Locked = False
     
     If Index = 0 Then chkAjustePriDeduc.Visible = True 'Cambio de Plazo
      
     If Index = 6 Then
         vsBar.Visible = True
         txtCambio.Text = txtUltMov.Caption
         vsBar.Value = 1000
         vsBar.Tag = vsBar.Value
         vsBar.Visible = True
         txtCambio.Locked = True
     End If
      
     If Index = 7 Then
         vsBar.Visible = True
         txtCambio.Text = lblPrideduc.Caption
         vsBar.Value = 1000
         vsBar.Tag = vsBar.Value
         vsBar.Visible = True
         txtCambio.Locked = True
     End If
      
      
     txtCambio.SetFocus
  
  Case 5, 9, 10, 15, 17
     'Solo Aplicar
     If Index = 5 Then
        Call sbLlenaMorosidad
        fraMora.Visible = True
     End If
     
     If Index = 15 Then
        Call sbLlenaCargos
        fraMora.Visible = True

     End If
     
  
  Case 8
     'Abonos
'     txtIntereses.Text = 0
'     txtAmortizacion.Text = 0
'
'     lblCaptionAbono(0).Top = txtCambio.Top
'     lblCaptionAbono(1).Top = lblCaptionAbono(0).Top + 360
'
'     txtIntereses.Top = lblCaptionAbono(0).Top
'     txtAmortizacion.Top = lblCaptionAbono(1).Top
'
'     txtIntereses.Visible = True
'     txtAmortizacion.Visible = True
'     lblCaptionAbono(0).Visible = True
'     lblCaptionAbono(1).Visible = True
'
'     txtIntereses.SetFocus
  

  Case 11, 12, 13, 14, 16, 18, 19 'Garantia,Destinos,Recursos, Dia de Pago, Actividad
     cboX.top = txtCambio.top
     cboX.Visible = True
     cboX.SetFocus
     If Index = 11 Then 'Garantias
          Call sbSTCargaCboGarantia(cboX, lblCodigo.Caption)
     End If
     
     If Index = 12 Then 'Destinos
          Call sbSTCargaCboDestinos(cboX, lblCodigo.Caption)
     End If
     
     If Index = 13 Then
          Call sbSTCargaCboRecursos(cboX, lblCodigo.Caption)
     End If
     
     If Index = 14 Then
          Call sbSTDiaPago(cboX, lblMonto.Tag)
     End If
     
     If Index = 16 Then
       strSQL = "select rtrim(cod_oficina) + ' - ' + rtrim(descripcion) as ItmX from sif_oficinas where cod_oficina not in('" & lblOficina.Tag & "')"
       Call sbLlenaCbo(cboX, strSQL, False, False)
     End If
     
      If Index = 18 Then '19 - 1
        strSQL = "select rtrim(cod_actividad) + ' - ' + descripcion as 'ItmX' from AFI_ACTIVIDADES_ECO where activa = 1"
        Call sbLlenaCbo(cboX, strSQL, False, False)
        
        strSQL = "select rtrim(A.cod_actividad) + ' - ' + A.descripcion as 'ItmX',A.cod_actividad" _
               & " from reg_creditos R inner join AFI_Actividades_Eco A on R.cod_Actividad = A.cod_Actividad" _
               & " Where id_Solicitud = " & txtOperacion.Text
        
        txtCambio.Text = ""
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
            cboX.Text = rs!itmX
            txtCambio.Text = Trim(rs!Cod_actividad)
        End If
        rs.Close
 
     End If 'Index = 18
                
     If Index = 19 Then 'Camibo de Ejecutivo (i20  - 1)
       strSQL = "select id_Promotor as 'Idx',rtrim(Nombre) as ItmX from Promotores where Estado = 1"
       Call sbLlenaCbo(cboX, strSQL, False, True)
     End If
                
     
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbDiaPago()

End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then cmdAceptar.SetFocus
End Sub

Private Sub txtOperacion_Change()
Dim i As Integer

 vOperacion = 0
 lblCedula.Caption = ""
 lblCodigo.Caption = ""
 lblNombre.Caption = ""
 lblDescripcion.Caption = ""
 txtUltMov = 0
 lblMonto.Caption = Format(0, "Standard")
 lblSaldo.Caption = Format(0, "Standard")
 lblCuota.Caption = Format(0, "Standard")
 lblOpex.Caption = ""
 
 lblPlazo.Caption = 0
 lblInteres.Caption = 0
 
 txtCambio = ""
 
' txtIntereses.Text = 0
' txtAmortizacion.Text = 0
 lblPlazoRestante.Caption = 0
 lblTasaOriginal.Caption = 0
 
 lblDestino.Caption = ""
 lblDestino.ToolTipText = ""
 
 lblRecurso.Caption = ""
 lblRecurso.ToolTipText = ""
 
 lblEjecutivo.Caption = ""
 lblEjecutivo.Tag = ""
 
 lswMov.ListItems.Clear
 
 fraMora.top = Me.Height + 200
 
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   Call sbCargaOperacion
End If
End Sub


Private Sub vsBar_Change()
On Error Resume Next
If vsBar.Value < Val(vsBar.Tag) Then txtCambio = fxFechaProcesoSiguiente(txtCambio)
If vsBar.Value > Val(vsBar.Tag) Then txtCambio = fxFechaProcesoAnterior(txtCambio)

vsBar.Tag = vsBar.Value

End Sub



