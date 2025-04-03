VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxC_CuentasAnulaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CxC: Movimientos > Anulaciones"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   10350
   Begin XtremeSuiteControls.PushButton btnAnular 
      Height          =   612
      Left            =   8040
      TabIndex        =   32
      Top             =   5640
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Anular"
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
      Picture         =   "frmCxC_CuentasAnulaciones.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCxC_CuentasAnulaciones.frx":0995
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtDocumento 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   31
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   240
      TabIndex        =   14
      Top             =   1605
      Width           =   1335
   End
   Begin VB.TextBox txtABIntMor 
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
      Left            =   3000
      TabIndex        =   13
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtABAmortizacion 
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
      Left            =   3000
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtABIntCor 
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
      Left            =   3000
      TabIndex        =   11
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtAmortizacion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtIntCor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      TabIndex        =   7
      ToolTipText     =   "Campo para la Cédula de Identidad"
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtConceptoCod 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      ToolTipText     =   "Código del Préstamo"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtIntMor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   4560
      Width           =   1815
   End
   Begin VB.TextBox txtCargos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtABCargos 
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
      Left            =   3000
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtABTotal 
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8040
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ComboBox cboEfecto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
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
      Height          =   312
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   3735
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2415
      Left            =   120
      TabIndex        =   15
      Top             =   1920
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Estado"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Int.Cor."
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Int.Mor."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Principal"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Cargos"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Dias.Cor."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dias.Mor."
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   10200
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   10320
      Y1              =   600
      Y2              =   600
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
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   30
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
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
      Height          =   315
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   1215
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
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   120
      TabIndex        =   28
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amortización"
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
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   27
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Int.Morosidad"
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
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   26
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Int.Corriente"
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
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label lblProceso 
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
      Left            =   8400
      TabIndex        =   24
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos Anulación"
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
      Height          =   315
      Index           =   5
      Left            =   3000
      TabIndex        =   23
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos Originales"
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
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   22
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Movimientos Registrados"
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
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   10095
   End
   Begin VB.Label lblNombre 
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
      Left            =   3000
      TabIndex        =   20
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblConceptoDesc 
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
      Left            =   3000
      TabIndex        =   19
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cargos"
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
      Height          =   315
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total Anulación"
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
      Index           =   8
      Left            =   4800
      TabIndex        =   17
      Top             =   4560
      Width           =   1452
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Efecto en el Plan"
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
      Index           =   9
      Left            =   4800
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   10080
      X2              =   4680
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "frmCxC_CuentasAnulaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long
Dim vPaso As Boolean

Private Sub btnAnular_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rs2 As New ADODB.Recordset
Dim vIntM As Currency, vIntC As Currency, vAMORTIZAm As Currency
Dim vANUIntM As Currency, vANUIntC As Currency, vANUAMORTIZAm As Currency
Dim vCuenta As String, vFecha As Date, lngRecibo As Long
Dim lngOperacion As Long

  
  If Not fxValidaInformacion Then
    MsgBox "La información suministrada no es válida...", vbCritical
    Exit Sub
  End If 'Validacion
  
  vFecha = fxFechaServidor
  vCuenta = Trim(fxDocumentoCuenta("CxC_ND"))
  
  If vAseDocValido = False Then
    MsgBox "No se puede Realizar Movimiento porque no se especificó una cuenta contable" _
         & " válida para esta operación...", vbCritical
    Exit Sub
  End If
    
  lngOperacion = txtOperacion
   
  Me.MousePointer = vbHourglass
  
    lngRecibo = 0
    If uRecibos Then lngRecibo = fxDocumentoAbono(CCur(txtABIntCor.Text), CCur(txtABIntMor.Text), CCur(txtABAmortizacion.Text) _
                            , CCur(txtABCargos.Text), "CxC_ND", vCuenta, "ANULA ABONO")
    
    strSQL = "exec spCrdPlanPagoAnulaAbono " & lngOperacion & ",'CRD008','" & glogon.Usuario & "','CxC_ND','" & lngRecibo & "',1," & CCur(txtABIntCor) _
           & "," & CCur(txtABIntMor.Text) & "," & CCur(txtABAmortizacion.Text) & "," & CCur(txtABCargos.Text) _
           & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',''"
    Call ConectionExecute(strSQL)
       
 Call Bitacora("Anula", "OP: " & txtOperacion & " Doc.:" & lngRecibo & " Total : " & CCur(txtABTotal))
 
 Me.MousePointer = vbDefault
 
 MsgBox "Anulación Realizada ... Con Nota Debito #" & lngRecibo, vbInformation

 If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "CxC_ND")

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Function fxValidaInformacion() As Boolean
 
 fxValidaInformacion = True
 
 If Len(Trim(txtABIntCor)) = 0 Then
   txtABIntCor = 0
 End If
 
 If Len(Trim(txtABCargos)) = 0 Then
   txtABCargos = 0
 End If
 
 If Len(Trim(txtABIntMor)) = 0 Then
   txtABIntMor = 0
 End If
 
 If Len(Trim(txtABAmortizacion)) = 0 Then
   txtABAmortizacion = 0
 End If
 
 If Len(Trim(txtABTotal)) = 0 Then
   txtABTotal = 0
 End If
 
 
  If (CCur(txtABAmortizacion) + CCur(txtABIntCor) + CCur(txtABIntMor) + CCur(txtCargos)) = 0 Then
    fxValidaInformacion = False
 End If

 If Len(Trim(lblNombre.Caption)) = 0 Then
    fxValidaInformacion = False
 End If

End Function


Private Function fxDocumentoAbono(curIntC As Currency, curIntM As Currency, curAmortiza As Currency, curCargo As Currency _
                                , vTipoDoc As String, vCuenta As String, vDetalle As String) As Long

Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim lngRecibo As Long, strCliente As String, vCuentaPoliza As String
Dim rsTmp As New ADODB.Recordset

lngRecibo = fxDocumentoConsecutivo(vTipoDoc)
fxDocumentoAbono = lngRecibo


  
strSQL = "exec spCxC_OperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

strLinea(1) = "Saldo Actual      " & Format(rs!Saldo, "Standard")
strLinea(2) = "Interes Corriente " & Format(curIntC * -1, "Standard")
strLinea(3) = "Interes Moratorio " & Format(curIntM * -1, "Standard")
strLinea(4) = "Amortización      " & Format(curAmortiza * -1, "Standard")
strLinea(5) = "Cargos            " & Format(curCargo * -1, "Standard")
strLinea(6) = ""
strLinea(7) = "Nuevo Saldo       " & Format(rs!Saldo + curAmortiza, "Standard")
strLinea(8) = "Operación /Linea  " & txtOperacion & "_" & txtConceptoCod.Text
strLinea(9) = ""
strLinea(10) = "Usuario           " & glogon.Usuario
strLinea(11) = "Anulación"

strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
        & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
        & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle)" _
        & " values('" & lngRecibo & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
        & "','" & Trim(lblNombre.Caption) & "','CRD008'," & curIntC + curIntM + curAmortiza + curCargo & ",'P','" & txtOperacion.Text _
        & "','" & txtConceptoCod.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
        & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
        & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
        & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
        & strLinea(11) & "','" & vAseDocDetalle & vbCrLf & "Depósito..:" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'ASIENTO
If curIntC > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & lngRecibo & "'," & curIntC & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
         & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

If curIntM > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & lngRecibo & "'," & curIntM & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
         & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

If curCargo > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & lngRecibo & "'," & curCargo & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!CtaCargos _
         & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If


If curAmortiza > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & lngRecibo & "'," & curAmortiza & ",'D','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
         & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

If curIntC + curIntM + curAmortiza + curCargo > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & vTipoDoc & "','" & lngRecibo & "'," & curIntC + curIntM + curCargo + curAmortiza & ",'C','" & rs!cod_Divisa _
         & "',1," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & vCuenta _
         & "','" & rs!Operacion & "','" & rs!cod_Concepto & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If

rs.Close

End Function

Private Sub Form_Activate()
 vModulo = 31
End Sub

Private Sub Form_Load()
 vModulo = 31
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Private Sub sbLimpia()
    
    txtOperacion = ""
    txtCedula = ""
    lblNombre.Caption = ""
    txtConceptoCod = ""
    lblConceptoDesc.Caption = ""
    
    txtABIntCor.Text = "0"
    txtABAmortizacion.Text = "0"
    txtABIntMor.Text = "0"
    txtABCargos.Text = "0"
    txtABTotal.Text = "0"
    
    txtIntCor.Text = "0"
    txtIntMor.Text = "0"
    txtAmortizacion.Text = "0"
    txtCargos.Text = "0"
    txtTotal.Text = "0"
    
    vOperacion = 0
    lsw.ListItems.Clear
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyDown(vbKeyReturn, 1)
End Sub



Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass


    
strSQL = "select R.Operacion,R.saldo,R.num_documento,R.Tasa_Corriente,R.dias_plazo,R.interesc,R.amortiza,R.Fecha_UltMov" _
       & ",R.Cod_Concepto,R.cedula,S.nombre,C.descripcion,R.Activa_Fecha,R.Tipo_Plazo,R.Proceso" _
       & " from CxC_Cuentas R inner join CxC_Conceptos C on R.Cod_Concepto = C.Cod_Concepto " _
       & " inner join CxC_Personas S on R.cedula = S.cedula" _
       & " where R.estado in('A','C') and R.Operacion = " & txtOperacion.Text
       
Call sbLimpia
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  
  vOperacion = rs!Operacion
  txtOperacion.Text = rs!Operacion
  
    txtCedula.Text = rs!Cedula
    lblNombre.Caption = rs!Nombre
    txtConceptoCod.Text = rs!cod_Concepto
    
    lblProceso.Tag = rs!Proceso
    Select Case rs!Proceso
      Case "N"
        lblProceso.Caption = "Normal"
      Case "T"
        lblProceso.Caption = "Traspaso Deuda"
      Case "J"
        lblProceso.Caption = "Cobro Judicial"
      Case "I"
        lblProceso.Caption = "Incobrable"
    End Select
    
    
    lblConceptoDesc.Caption = rs!Descripcion
    txtDocumento.Text = rs!Num_Documento & ""
        
    'Movimientos Registrados
    strSQL = "select * from CXC_CUENTAS_MOV where estado = 'C' and Operacion = " & rs!Operacion _
           & " order by linea desc"
    rs.Close
    
    vPaso = True
    chkTodos.Value = vbUnchecked
    
    Call OpenRecordSet(rs, strSQL)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Linea)
          itmX.SubItems(3) = IIf((rs!Dias_Mora > 0), "En Mora", "Al Día")
          itmX.SubItems(4) = Format(rs!Mov_IntCor, "Standard")
          itmX.SubItems(5) = Format(rs!Mov_IntMor, "Standard")
          itmX.SubItems(6) = Format(rs!Mov_Principal, "Standard")
          itmX.SubItems(7) = Format(rs!Mov_Cargos, "Standard")
          itmX.SubItems(9) = rs!Dias
          itmX.SubItems(10) = rs!Dias_Mora
          itmX.Tag = rs!Linea
      rs.MoveNext
    Loop
    
    vPaso = False
Else
    MsgBox "No se encontró operación la operación...!", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If vPaso Then Exit Sub

If Item.Checked Then
  txtIntCor.Text = CCur(txtIntCor.Text) + CCur(Item.SubItems(2))
  txtIntMor.Text = CCur(txtIntMor.Text) + CCur(Item.SubItems(3))
  txtAmortizacion.Text = CCur(txtAmortizacion.Text) + CCur(Item.SubItems(4))
  txtCargos.Text = CCur(txtCargos.Text) + CCur(Item.SubItems(5))

Else
  txtIntCor.Text = CCur(txtIntCor.Text) - CCur(Item.SubItems(2))
  txtIntMor.Text = CCur(txtIntMor.Text) - CCur(Item.SubItems(3))
  txtAmortizacion.Text = CCur(txtAmortizacion.Text) - CCur(Item.SubItems(4))
  txtCargos.Text = CCur(txtCargos.Text) - CCur(Item.SubItems(5))
End If

txtTotal.Text = CCur(txtIntCor.Text) + CCur(txtIntMor.Text) + CCur(txtAmortizacion.Text) + CCur(txtCargos.Text)

txtIntCor.Text = Format(txtIntCor.Text, "Standard")
txtIntMor.Text = Format(txtIntMor.Text, "Standard")
txtAmortizacion.Text = Format(txtAmortizacion.Text, "Standard")
txtCargos.Text = Format(txtCargos.Text, "Standard")

txtTotal.Text = Format(txtTotal.Text, "Standard")

txtABIntCor.Text = Format(txtIntCor.Text, "Standard")
txtABIntMor.Text = Format(txtIntMor.Text, "Standard")
txtABAmortizacion.Text = Format(txtAmortizacion.Text, "Standard")
txtABCargos.Text = Format(txtCargos.Text, "Standard")
txtABTotal.Text = Format(txtTotal.Text, "Standard")

 
End Sub

Private Sub sbCalculaTotal()
On Error GoTo vError
  txtABTotal.Text = Format(CCur(txtABIntCor.Text) + CCur(txtABIntMor.Text) + CCur(txtABAmortizacion.Text) _
                  + CCur(txtABCargos.Text), "Standard")
vError:
End Sub




Private Sub txtABIntCor_GotFocus()
On Error GoTo vError
    txtABIntCor.Text = CCur(txtABIntCor.Text)
vError:
End Sub

Private Sub txtABIntCor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABIntMor.SetFocus
End Sub

Private Sub txtABIntCor_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABIntCor_LostFocus()
On Error GoTo vError
    txtABIntCor.Text = Format(CCur(txtABIntCor.Text), "Standard")
vError:
End Sub


Private Sub txtABIntMor_GotFocus()
On Error GoTo vError
    txtABIntMor.Text = CCur(txtABIntMor.Text)
vError:
End Sub

Private Sub txtABIntMor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABAmortizacion.SetFocus
End Sub

Private Sub txtABIntMor_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABIntMor_LostFocus()
On Error GoTo vError
    txtABIntMor.Text = Format(CCur(txtABIntMor.Text), "Standard")
vError:
End Sub

Private Sub txtABAmortizacion_GotFocus()
On Error GoTo vError
    txtABAmortizacion.Text = CCur(txtABAmortizacion.Text)
vError:
End Sub

Private Sub txtABAmortizacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABCargos.SetFocus
End Sub

Private Sub txtABAmortizacion_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABAmortizacion_LostFocus()
On Error GoTo vError
    txtABAmortizacion.Text = Format(CCur(txtABAmortizacion.Text), "Standard")
vError:
End Sub

Private Sub txtABCargos_GotFocus()
On Error GoTo vError
    txtABCargos.Text = CCur(txtABCargos.Text)
vError:
End Sub

Private Sub txtABCargos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtABTotal.SetFocus
End Sub

Private Sub txtABCargos_KeyUp(KeyCode As Integer, Shift As Integer)
 Call sbCalculaTotal
End Sub

Private Sub txtABCargos_LostFocus()
On Error GoTo vError
    txtABCargos.Text = Format(CCur(txtABCargos.Text), "Standard")
vError:
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyReturn Then
  Call sbConsulta
End If

End Sub

