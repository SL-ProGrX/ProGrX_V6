VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmInvOrdNivelAuto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizaciones / Usuarios / Asignación"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   9135
   Begin TabDlg.SSTab ssTab 
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Autorizadores"
      TabPicture(0)   =   "frmInvOrdNivelAuto.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lswAuto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Usuarios a Cargo"
      TabPicture(1)   =   "frmInvOrdNivelAuto.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbo"
      Tab(1).Control(1)=   "vGrid"
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(3)=   "lblx"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Cambios de Fecha (E/S/T)"
      TabPicture(2)   =   "frmInvOrdNivelAuto.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswCambiaFechas"
      Tab(2).Control(1)=   "ssTabAux01"
      Tab(2).Control(2)=   "Label1(2)"
      Tab(2).ControlCount=   3
      Begin VB.ComboBox cbo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73320
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   5775
      End
      Begin MSComctlLib.ListView lswAuto 
         Height          =   5532
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   8652
         _ExtentX        =   15266
         _ExtentY        =   9763
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8467
         EndProperty
      End
      Begin MSComctlLib.ListView lswCambiaFechas 
         Height          =   4812
         Left            =   -74760
         TabIndex        =   5
         Top             =   1440
         Width           =   8532
         _ExtentX        =   15055
         _ExtentY        =   8493
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
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
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   7232
         EndProperty
      End
      Begin TabDlg.SSTab ssTabAux01 
         Height          =   5412
         Left            =   -74880
         TabIndex        =   4
         Top             =   960
         Width           =   8772
         _ExtentX        =   15478
         _ExtentY        =   9551
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Entradas"
         TabPicture(0)   =   "frmInvOrdNivelAuto.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Salidas"
         TabPicture(1)   =   "frmInvOrdNivelAuto.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Traslados"
         TabPicture(2)   =   "frmInvOrdNivelAuto.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5172
         Left            =   -74880
         TabIndex        =   9
         Top             =   1200
         Width           =   8772
         _Version        =   524288
         _ExtentX        =   15473
         _ExtentY        =   9123
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmInvOrdNivelAuto.frx":00A8
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Usuarios que pueden Cambiar las fechas de las Entradas / Salidas / Traslados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   2
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   8652
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Asignación de los Usuarios que tienen a cargo los autorizadores de transacciones de  E/S/T/R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   1
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   8772
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Usuarios Autorizadores de Entradas/Salidas/Traslados y Requisiciones de Inventarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   8652
      End
      Begin VB.Label lblx 
         Caption         =   "Usuarios a Cargo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel de Autorización de los Usuarios "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   3
      Left            =   1560
      TabIndex        =   10
      Top             =   240
      Width           =   8052
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmInvOrdNivelAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vUsuario As String

On Error GoTo vError

If Not vPaso Then Exit Sub


Me.MousePointer = vbHourglass

vUsuario = fxCodigoCbo(cbo)

strSQL = "select U.nombre,U.descripcion,isnull(C.Entradas,0),isnull(C.Salidas,0),isnull(C.requisiciones,0),isnull(C.Traslados,0)" _
       & " from usuarios U left join pv_orden_autousers C on U.nombre = C.usuario_asignado" _
       & " and C.usuario = '" & fxCodigoCbo(cbo) & "' Where U.Estado = 'A'" _
       & " order by C.fecha_asignacion desc"
vPaso = False
 Call sbCargaGrid(vGrid, 6, strSQL)
vPaso = True

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()

vModulo = 32
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

ssTab.Tab = 0

Call sbLlenaLswAuto

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLlenaCboAuto()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vUltimo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = False

cbo.Clear
vUltimo = ""

strSQL = "select U.nombre,U.descripcion" _
       & " from usuarios U inner join pv_orden_autorizadores A on U.nombre = A.usuario" _
       & " Where U.Estado = 'A'" _
       & " order by U.nombre"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 cbo.AddItem Trim(rs!Nombre) & " - " & Trim(rs!Descripcion)
 vUltimo = Trim(rs!Nombre) & " - " & Trim(rs!Descripcion)
 rs.MoveNext
Loop
rs.Close

If vUltimo <> "" Then
 cbo.Text = vUltimo
 vPaso = True
 cbo_Click
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbllenalswUserFechas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswCambiaFechas.ListItems.Clear

strSQL = "select U.nombre,U.descripcion,A.tipo" _
       & " from usuarios U left join PV_INVUSRFECHAS A on U.nombre = A.usuario and A.tipo = '"
Select Case ssTabAux01.Tab
  Case 0
     strSQL = strSQL & "E'"
  Case 1
     strSQL = strSQL & "S'"
  Case 2
     strSQL = strSQL & "T'"
End Select
strSQL = strSQL & "Where U.EStado = 'A' order by A.tipo desc"

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswCambiaFechas.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!Tipo), False, True)
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswAuto_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
 strSQL = "insert pv_orden_autorizadores(usuario,fecha,estado) values('" _
        & Item.Text & "',dbo.MyGetdate(),'A')"
 Call ConectionExecute(strSQL)
Else
 strSQL = "delete pv_orden_autousers where usuario = '" & Item.Text & "'"
 Call ConectionExecute(strSQL)
  
 strSQL = "delete pv_orden_autorizadores where usuario = '" & Item.Text & "'"
 Call ConectionExecute(strSQL)
 
End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswCambiaFechas_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, vTipo As String

Select Case ssTabAux01.Tab
  Case 0
    vTipo = "E"
  Case 1
    vTipo = "S"
  Case 2
    vTipo = "T"
End Select

If Item.Checked Then
   strSQL = "insert PV_INVUSRFECHAS(usuario,tipo) values('" & Item.Text _
          & "','" & vTipo & "')"
Else
   strSQL = "delete PV_INVUSRFECHAS where usuario = '" & Item.Text _
        & "' and tipo = '" & vTipo & "'"
End If
Call ConectionExecute(strSQL)

End Sub

Private Sub sbLlenaLswAuto()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lswAuto.ListItems.Clear

strSQL = "select U.nombre,U.descripcion,A.fecha" _
       & " from usuarios U left join pv_orden_autorizadores A  on U.nombre = A.usuario" _
       & " Where U.Estado = 'A'" _
       & " order by A.fecha desc"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswAuto.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!fecha), False, True)
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub SSTab_Click(PreviousTab As Integer)
Select Case ssTab.Tab
 Case 0
   Call sbLlenaLswAuto
 Case 1
   Call sbLlenaCboAuto
 Case 2
   ssTabAux01.Tab = 0
   Call sbllenalswUserFechas
End Select
End Sub

Private Sub ssTabAux01_Click(PreviousTab As Integer)
Call sbllenalswUserFechas
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not vPaso Then Exit Sub

If col > 2 Then
  vGrid.Row = Row
  vGrid.col = 1
  strSQL = "select isnull(count(*),0) as Existe from pv_orden_autousers" _
         & " where usuario = '" & fxCodigoCbo(cbo) & "' and usuario_asignado = '" _
         & vGrid.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
    strSQL = "insert pv_orden_autoUsers(usuario,usuario_asignado,fecha_asignacion,entradas" _
           & ",salidas,requisiciones,traslados) values('" & fxCodigoCbo(cbo) & "','" & vGrid.Text _
           & "',dbo.MyGetdate(),0,0,0,0)"
    Call ConectionExecute(strSQL)
  End If
  rs.Close
    
  vGrid.col = col
  vGrid.Row = Row
  Select Case col
    Case 3 'Entradas
        strSQL = "update pv_orden_autoUsers set entradas = " & vGrid.Value
    Case 4 'Salidas
        strSQL = "update pv_orden_autoUsers set salidas = " & vGrid.Value
    Case 5 'Requisiciones
        strSQL = "update pv_orden_autoUsers set requisiciones = " & vGrid.Value
    Case 6 'Traslados
        strSQL = "update pv_orden_autoUsers set Traslados = " & vGrid.Value
  End Select
  vGrid.col = 1
  strSQL = strSQL & " where usuario = '" & fxCodigoCbo(cbo) & "' and usuario_asignado = '" _
         & vGrid.Text & "'"
  Call ConectionExecute(strSQL)
End If


Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
