VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmCprOrdNivelAuto 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Niveles de Autorización para Ordenes de Compras"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7905
   ScaleWidth      =   11550
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   11535
      _Version        =   1441792
      _ExtentX        =   20346
      _ExtentY        =   11456
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
      ItemCount       =   3
      Item(0).Caption =   "Autorizadores"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "scTitulo(0)"
      Item(0).Control(1)=   "lswAuto"
      Item(1).Caption =   "Usuarios a Cargo"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "scTitulo(1)"
      Item(1).Control(1)=   "scTitulo(2)"
      Item(1).Control(2)=   "lswAutoList"
      Item(1).Control(3)=   "lsw"
      Item(1).Control(4)=   "txtAutoFiltro"
      Item(1).Control(5)=   "txtAutorizadorSel"
      Item(2).Caption =   "Cambios de Fecha"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "scTitulo(3)"
      Item(2).Control(1)=   "lswCambiaFechas"
      Begin XtremeSuiteControls.ListView lswCambiaFechas 
         Height          =   5775
         Left            =   -70000
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   11535
         _Version        =   1441792
         _ExtentX        =   20346
         _ExtentY        =   10186
         _StockProps     =   77
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswAutoList 
         Height          =   5295
         Left            =   -70000
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
         _ExtentY        =   9340
         _StockProps     =   77
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5295
         Left            =   -64240
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
         _ExtentY        =   9340
         _StockProps     =   77
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswAuto 
         Height          =   5775
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   11535
         _Version        =   1441792
         _ExtentX        =   20346
         _ExtentY        =   10186
         _StockProps     =   77
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAutoFiltro 
         Height          =   495
         Left            =   -70000
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
         _ExtentY        =   873
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAutorizadorSel 
         Height          =   495
         Left            =   -64240
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
         _ExtentY        =   873
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   3
         Left            =   -70000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   11535
         _Version        =   1441792
         _ExtentX        =   20346
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Usuarios Autorizados a Cambiar las Fecha dela Ordenes y Compras Registradas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   2
         Left            =   -64240
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Usuarios a Cargo"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   1
         Left            =   -70000
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   5775
         _Version        =   1441792
         _ExtentX        =   10186
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Autorizadores"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   11535
         _Version        =   1441792
         _ExtentX        =   20346
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Usuarios designados como Autorizadores de Ordenes de Compra"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Niveles de Autorización para el Proceso de Compras"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   10
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmCprOrdNivelAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbLswAutorizadores_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

lswAuto.ListItems.Clear

vPaso = True

strSQL = "select U.nombre,U.descripcion,A.fecha" _
       & " from usuarios U left join cpr_orden_autorizadores A  on U.nombre = A.usuario" _
       & " order by A.fecha desc"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswAuto.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!fecha), False, True)
     If Not IsNull(rs!fecha) Then itmX.ForeColor = vbBlue
 
 rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 35
End Sub

Private Sub Form_Load()

vModulo = 35

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswAuto.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", lswAuto.Width - 2700
End With

With lsw.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", lsw.Width - 2700
End With

With lswAutoList.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", lswAutoList.Width - 2700
End With

With lswCambiaFechas.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2500
    .Add , , "Nombre", lswCambiaFechas.Width - 2700
End With


tcMain.item(0).Selected = True


Call sbLswAutorizadores_Load

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLswFechaCambio_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass


vPaso = True
lswCambiaFechas.ListItems.Clear

strSQL = "select U.nombre,U.descripcion,A.usuario" _
       & " from usuarios U left join cpr_INVUSRFECHAS A on U.nombre = A.usuario" _
       & " order by A.usuario desc"

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 Set itmX = lswCambiaFechas.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
     itmX.Checked = IIf(IsNull(rs!Usuario), False, True)
     If Not IsNull(rs!Usuario) Then itmX.ForeColor = vbBlue

 rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLswAutorizador_List()

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True
txtAutorizadorSel.Tag = ""
txtAutorizadorSel.Text = "(Seleccionar a un Autorizador)"
lsw.ListItems.Clear

txtAutoFiltro.Text = fxSysCleanTxtInject(txtAutoFiltro.Text)

With lswAutoList.ListItems

    .Clear
    
    strSQL = "select U.nombre,U.descripcion" _
           & " from usuarios U inner join cpr_orden_autorizadores A on U.nombre = A.usuario" _
           & " Where (U.nombre like '%" & txtAutoFiltro.Text & "%'" _
           & " OR U.descripcion like '%" & txtAutoFiltro.Text & "%')" _
           & " order by U.nombre"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Nombre)
            itmX.SubItems(1) = rs!Descripcion & ""
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




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal item As XtremeSuiteControls.ListViewItem)
Dim pUsuario As String

On Error GoTo vError

pUsuario = txtAutorizadorSel.Tag

If pUsuario = "" Then Exit Sub
If vPaso Then Exit Sub


If item.Checked Then
 strSQL = "insert cpr_orden_autousers(usuario,usuario_asignado,fecha_Asignacion) values('" _
        & pUsuario & "','" & item.Text & "',dbo.MyGetdate())"
 Call ConectionExecute(strSQL)

 Call Bitacora("Aplica", "Asignación de Usuario a Cargo(Aut.Ord.Cpr):" & pUsuario & "->" & item.Text)

Else
 strSQL = "delete cpr_orden_autousers where usuario = '" & pUsuario _
        & "' and usuario_asignado = '" & item.Text & "'"
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Elimina", "Asignación de Usuario a Cargo(Aut.Ord.Cpr):" & pUsuario & "->" & item.Text)

End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswAuto_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswAuto.SortKey = ColumnHeader.Index - 1
  If lswAuto.SortOrder = 0 Then lswAuto.SortOrder = 1 Else lswAuto.SortOrder = 0
  lswAuto.Sorted = True
End Sub

Private Sub lswAuto_ItemCheck(ByVal item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If item.Checked Then
 strSQL = "insert cpr_orden_autorizadores(usuario,fecha,estado) values('" _
        & item.Text & "',dbo.MyGetdate(),'A')"
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Registra", "Usuario Autorizador de Ordenes de Compra:" & item.Text)

Else
 strSQL = "delete cpr_orden_autousers where usuario = '" & item.Text & "'"
 Call ConectionExecute(strSQL)
  
 strSQL = "delete cpr_orden_autorizadores where usuario = '" & item.Text & "'"
 Call ConectionExecute(strSQL)
 
 Call Bitacora("Elimina", "Usuario Autorizador de Ordenes de Compra:" & item.Text)
 
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswAutoList_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswAutoList.SortKey = ColumnHeader.Index - 1
  If lswAutoList.SortOrder = 0 Then lswAutoList.SortOrder = 1 Else lswAutoList.SortOrder = 0
  lswAutoList.Sorted = True
End Sub

Private Sub lswAutoList_ItemClick(ByVal item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

txtAutorizadorSel.Text = item.SubItems(1)
txtAutorizadorSel.Tag = item.Text


strSQL = "select U.nombre,U.descripcion,C.fecha_asignacion" _
       & " from usuarios U left join cpr_orden_autousers C on U.nombre = C.usuario_asignado" _
       & " and C.usuario = '" & txtAutorizadorSel.Tag & "'" _
       & " order by C.fecha_asignacion desc"

vPaso = True

lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Nombre)
     itmX.SubItems(1) = rs!Descripcion
 
 If Not IsNull(rs!fecha_asignacion) Then
    itmX.Checked = True
    itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswCambiaFechas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCambiaFechas.SortKey = ColumnHeader.Index - 1
  If lswCambiaFechas.SortOrder = 0 Then lswCambiaFechas.SortOrder = 1 Else lswCambiaFechas.SortOrder = 0
  lswCambiaFechas.Sorted = True
End Sub

Private Sub lswCambiaFechas_ItemCheck(ByVal item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub

If item.Checked Then
   strSQL = "insert CPR_INVUSRFECHAS(usuario,registro_fecha,registro_usuario)" _
          & " values('" & item.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
          
   Call Bitacora("Registra", "Usuario Autorizado Cambio Fecha Compras: " & item.Text)
Else
   strSQL = "delete CPR_INVUSRFECHAS where usuario = '" & item.Text & "'"
   Call Bitacora("Elimina", "Usuario Autorizado Cambio Fecha Compras: " & item.Text)
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal item As XtremeSuiteControls.ITabControlItem)
Select Case item.Index
 Case 0
   Call sbLswAutorizadores_Load
 Case 1
   Call sbLswAutorizador_List
 Case 2
   Call sbLswFechaCambio_Load
End Select
End Sub


Private Sub txtAutoFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbLswAutorizador_List
End If
End Sub
