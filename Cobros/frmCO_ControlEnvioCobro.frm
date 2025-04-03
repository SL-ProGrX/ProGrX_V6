VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCO_ControlEnvioCobro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio al cobro de las gestiones pendientes"
   ClientHeight    =   7584
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCO_ControlEnvioCobro.frx":0000
   ScaleHeight     =   7584
   ScaleWidth      =   9720
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   975
      Left            =   8520
      Picture         =   "frmCO_ControlEnvioCobro.frx":3482
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CheckBox chkMarcas 
      Appearance      =   0  'Flat
      Caption         =   "Marcar/DesMarcar Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.CheckBox chkTodos 
      Appearance      =   0  'Flat
      Caption         =   "Todos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   3
      Top             =   240
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txtGestion 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtGestionDesc 
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Presione (F4) para Consultar"
      Top             =   240
      Width           =   5295
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   5292
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9492
      _ExtentX        =   16743
      _ExtentY        =   9335
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   1129
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cedula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   5715
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Monto"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Gestion"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Descripción"
         Object.Width           =   3775
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Usuario"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Línea"
         Object.Width           =   1658
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   9600
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   9600
      X2              =   120
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Image imgBuscar 
      Height          =   375
      Left            =   8400
      Picture         =   "frmCO_ControlEnvioCobro.frx":6904
      Stretch         =   -1  'True
      ToolTipText     =   "Buscar Gestiones Pendientes"
      Top             =   170
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Gestión"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frmCO_ControlEnvioCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMarcas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkMarcas.Value
Next i

End Sub

Private Sub chkTodos_Click()

If chkTodos.Value = vbChecked Then
   txtGestion.Enabled = False
Else
   txtGestion.Enabled = True
End If

End Sub

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()

vModulo = 4

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Sub imgBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

strSQL = "select X.*,S.nombre,G.descripcion as GestionX,G.codigo_referencia" _
       & " from Socios S inner join cbr_seguimiento X on S.cedula = X.cedula" _
       & " inner join cbr_gestiones G on X.cod_gestion = G.cod_gestion" _
       & " where X.estado = 0 and X.operacion_credito is null"
If chkTodos.Value = vbUnchecked Then
   strSQL = strSQL & " and G.cod_gestion = '" & txtGestion.Text & "'"
End If

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Cod_Seg)
      itmX.SubItems(1) = Format(rs!fecha, "dd/mm/yyyy")
      itmX.SubItems(2) = rs!Cedula
      itmX.SubItems(3) = rs!Nombre
      itmX.SubItems(4) = Format(rs!Monto, "Standard")
      itmX.SubItems(5) = rs!COD_GESTION
      itmX.SubItems(6) = rs!gestionX
      itmX.SubItems(7) = rs!Usuario
      itmX.SubItems(8) = rs!codigo_referencia
  rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtGestion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtGestionDesc.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "cod_gestion"
    gBusquedas.Orden = "cod_gestion"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtGestion = Trim(gBusquedas.Resultado)
    txtGestionDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtGestion_LostFocus()
txtGestionDesc = fxCBRControlGestion(txtGestion)
End Sub

Private Sub txtGestionDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call imgBuscar_Click

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Consulta = "Select cod_gestion,descripcion from cbr_gestiones"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtGestion = Trim(gBusquedas.Resultado)
    txtGestionDesc = Trim(gBusquedas.Resultado2)
End If

End Sub

