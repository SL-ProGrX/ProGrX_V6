VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmVivConsultaDesembolso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta información desembolso"
   ClientHeight    =   6825
   ClientLeft      =   -1845
   ClientTop       =   4425
   ClientWidth     =   9090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      ToolTipText     =   "Nombre"
      Top             =   480
      Width           =   3408
   End
   Begin VB.TextBox txtCedula 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1860
      TabIndex        =   7
      ToolTipText     =   "Cédula de identidad"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox TxtExpediente 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      MaxLength       =   20
      TabIndex        =   6
      ToolTipText     =   "Número de expediente"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtLinea 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7080
      TabIndex        =   5
      ToolTipText     =   "Presione (F4) para Consultarn línea de crédito"
      Top             =   480
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvwResultado 
      Height          =   5628
      Left            =   168
      TabIndex        =   0
      Top             =   936
      Width           =   8712
      _ExtentX        =   15372
      _ExtentY        =   9922
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Num. Operación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cédula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   5380
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Num.Desembolo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Disponible"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " N° operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   1860
      TabIndex        =   3
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Nombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   3408
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   1932
   End
End
Attribute VB_Name = "frmVivConsultaDesembolso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vNumOperacion As String
    

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn) Then
    Call gsbPulsarTecla(vbKeyTab)
ElseIf KeyCode = vbKeyF4 Then
    If Me.ActiveControl.Name = "txtLinea" Then
        Call sbMostraVentanBusqueda
    End If
End If

End Sub

Private Sub sbMostraVentanBusqueda()

On Error GoTo vError

gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

gBusquedas.Consulta = "select codigo,descripcion from catalogo"
gBusquedas.Orden = "codigo"
gBusquedas.Columna = "codigo"
gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N'"
frmBusquedas.Show vbModal

txtLinea.Text = gBusquedas.Resultado

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
If IsNumeric(TxtExpediente.Text) Then
    GLOBALES.gTag = TxtExpediente
Else
    GLOBALES.gTag = ""
End If
End Sub

