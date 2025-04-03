VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCajas_Usuarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cajas: Asignación de Usuarios"
   ClientHeight    =   7800
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10368
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   10368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAsignados 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solo usuarios asignados"
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
      Height          =   225
      Left            =   2640
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtUsuario 
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
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
   End
   Begin MSComctlLib.ImageList imgExplorer 
      Left            =   7200
      Top             =   0
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
            Picture         =   "frmCajas_Usuarios.frx":0000
            Key             =   "imgDocu"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Usuarios.frx":0EDA
            Key             =   "imgFormu"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lswCajas 
      Height          =   6135
      Left            =   5160
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
      _ExtentX        =   8911
      _ExtentY        =   10816
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
         Name            =   "Calibri"
         Size            =   9
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
   Begin MSComctlLib.ListView lswUsuarios 
      Height          =   6135
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4935
      _ExtentX        =   8700
      _ExtentY        =   10816
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
         Name            =   "Calibri"
         Size            =   9
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Asignación de Usuarios a Cajas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   6495
   End
   Begin VB.Label lblUsuario 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario ..:"
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
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Buscar Usuario..:"
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
      Index           =   5
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10455
   End
End
Attribute VB_Name = "frmCajas_Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 5
End Sub

Private Sub Form_Load()
vModulo = 5

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call sbLimpia

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpia()

vPaso = True

txtUsuario.Text = ""
lswUsuarios.ListItems.Clear
lswCajas.ListItems.Clear

vPaso = False

End Sub

Private Sub sbConsultaUsuario()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select Nombre,Descripcion " _
       & " from Usuarios" _
       & " where Estado = 'A' and Nombre like '%" & Trim(txtUsuario.Text) & "%'"

If chkAsignados.Value = vbChecked Then
   strSQL = strSQL & " and Nombre in(select usuario from Cajas_Usuarios)"
End If

vPaso = True

lblUsuario.Caption = "Usuario .: "
lblUsuario.Tag = ""

lswCajas.ListItems.Clear

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
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub sbConsultaCajas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select Def.COD_CAJA, Def.DESCRIPCION, Cus.USUARIO  " _
       & " from CAJAS_DEFINICION Def left join  CAJAS_USUARIOS Cus on Def.COD_CAJA = Cus.COD_CAJA " _
       & " and Cus.USUARIO = '" & lblUsuario.Tag & "'" _
       & " Order by Def.Cod_Caja"


vPaso = True

With lswCajas.ListItems
    .Clear
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!cod_caja)
          itmX.SubItems(1) = rs!Descripcion
      
      If Not IsNull(rs!Usuario) Then
          itmX.Checked = True
      End If
      rs.MoveNext
    Loop
    rs.Close
End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




Private Sub lswCajas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If vPaso Or lswCajas.ListItems.Count = 0 Then Exit Sub

On Error GoTo vError

If Item.Checked Then

    strSQL = "insert into cajas_usuarios(cod_caja,usuario,registro_fecha,registro_usuario,contrasena,contrasena_renovacion)" _
           & "values('" & Item.Text & "','" & lblUsuario.Tag & "',dbo.MyGetdate(),'" & glogon.Usuario & "'," _
           & "'', dbo.MyGetdate())"
    Call ConectionExecute(strSQL)
   
    strSQL = "insert into cajas_usuarios_H(Linea,cod_caja,usuario,registro_usuario,registro_fecha)" _
           & " values( (select isnull(max(Linea),0) + 1 from cajas_usuarios_h where cod_caja = '" & Item.Text & "')" _
           & ",'" & Item.Text & "','" & lblUsuario.Tag & "','" & glogon.Usuario & "',dbo.MyGetdate())"
    Call ConectionExecute(strSQL)


Else
    strSQL = "Delete cajas_usuarios where cod_caja = '" & Item.Text & "' and usuario = '" & lblUsuario.Tag & "'"
    Call ConectionExecute(strSQL)
    
    strSQL = "update cajas_usuarios_H set salida_usuario = '" & glogon.Usuario & "', salida_fecha = dbo.MyGetdate()" _
           & " where cod_Caja = '" & Item.Text & "' and Linea in(select Max(linea) from cajas_usuarios_H" _
           & " where cod_Caja = '" & Item.Text & "' and usuario = '" & lblUsuario.Tag & "')"
    Call ConectionExecute(strSQL)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswUsuarios_Click()
If vPaso Or lswUsuarios.ListItems.Count = 0 Then Exit Sub

lblUsuario.Caption = "Usuario ..: " & lswUsuarios.SelectedItem.Text
lblUsuario.Tag = lswUsuarios.SelectedItem.Text

Call sbConsultaCajas

End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Consulta = "Select nombre,descripcion from Usuarios"
    gBusquedas.Columna = "descripcion"
    gBusquedas.Orden = "descripcion"
    frmBusquedas.Show vbModal
    txtUsuario.Text = gBusquedas.Resultado
    Call sbConsultaUsuario

Else
    Call sbConsultaUsuario
End If

End Sub


Private Sub txtUsuario_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> vbKeyF4 Then
    Call sbConsultaUsuario
End If
End Sub
