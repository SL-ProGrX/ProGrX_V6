VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSeguros_Usuarios_Accesos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acceso de Usuarios a Productos de Seguros"
   ClientHeight    =   6915
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   1200
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   2040
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin MSComctlLib.ListView lswUsuarios 
      Height          =   5052
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   4212
      _ExtentX        =   7435
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6068
      EndProperty
   End
   Begin MSComctlLib.ListView lswProductos 
      Height          =   5052
      Left            =   4560
      TabIndex        =   4
      Top             =   1800
      Width           =   6372
      _ExtentX        =   11245
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Aseguradora"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Producto"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique  los Productos que puede registrar el usuario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   8895
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar Usuario..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1692
   End
   Begin VB.Label lblUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario ..:"
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
      Height          =   348
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   6372
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11292
   End
End
Attribute VB_Name = "frmSeguros_Usuarios_Accesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 17

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbUsuariosLista()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select NOMBRE,DESCRIPCION" _
       & " from Usuarios" _
       & " where Estado = 'A' and (NOMBRE like '%" & Trim(txtUsuario.Text) & "%' or DESCRIPCION like '%" & Trim(txtUsuario.Text) & "%')"

vPaso = True

With lswUsuarios.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Nombre)
          itmX.SubItems(1) = rs!Descripcion
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

 
End Sub


Private Sub lswProductos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String
Dim pMov As String, pDetalle As String

If vPaso Or lswProductos.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

pDetalle = "Acceso del Producto: " & Item.Text & "¦" & Item.SubItems(1) & " al Usuario: " & lblUsuario.Tag

If Item.Checked Then
    pMov = "Aplica"
    
    strSQL = "exec spSeguros_ProductosUsuario_Registra '" & lblUsuario.Tag & "','" & Item.Text & "','" & Item.SubItems(1) & "','" & glogon.Usuario & "','A'"
Else
    pMov = "Elimina"
   
    strSQL = "exec spSeguros_ProductosUsuario_Registra '" & lblUsuario.Tag & "','" & Item.Text & "','" & Item.SubItems(1) & "','" & glogon.Usuario & "','A'"
End If

Call ConectionExecute(strSQL)

If Not glogon.error Then
  Call Bitacora(pMov, pDetalle)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
                                             
End Sub

Private Sub lswUsuarios_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If vPaso Or lswUsuarios.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

lblUsuario.Caption = "Usuario ..: " & lswUsuarios.SelectedItem.Text
lblUsuario.Tag = lswUsuarios.SelectedItem.Text


strSQL = "exec spSeguros_ProductosUsuario '" & lblUsuario.Tag & "'"

Call OpenRecordSet(rs, strSQL)

vPaso = True

With lswProductos.ListItems
  .Clear
  
Do While Not rs.EOF
  Set itmX = .Add(, , rs!cod_Aseguradora)
      itmX.SubItems(1) = rs!COD_PRODUCTO
      itmX.SubItems(2) = rs!Descripcion
      itmX.Checked = rs!Asignado

  rs.MoveNext
Loop
rs.Close

End With

vPaso = False


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbUsuariosLista

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Consulta = "Select Nombre,Descripcion from Usuarios"
    gBusquedas.Filtro = "and Estado = 'A'"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario.Text = gBusquedas.Resultado
    Call sbUsuariosLista

Else
    Call sbUsuariosLista
End If

End Sub


Private Sub txtUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyF4 Then
    Call sbUsuariosLista
End If
End Sub

