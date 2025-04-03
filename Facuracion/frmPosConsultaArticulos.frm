VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmPosConsultaArticulos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Articulos"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
   Icon            =   "frmPosConsultaArticulos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7185
   ScaleWidth      =   10275
   Begin XtremeSuiteControls.CheckBox chkSimilar 
      Height          =   216
      Left            =   8520
      TabIndex        =   17
      Top             =   876
      Width           =   216
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   480
      Top             =   1920
   End
   Begin VB.TextBox txtX01 
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
      Height          =   315
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Total ...:"
      Top             =   6600
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar tlbSumas 
      Height          =   1320
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   2328
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mas"
            Object.ToolTipText     =   "Aumenta la Cantidad de Articulos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Menos"
            Object.ToolTipText     =   "Disminuye la Cantidad de Articulos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Elimina"
            Object.ToolTipText     =   "Elimina Línea"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   360
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosConsultaArticulos.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosConsultaArticulos.frx":1728C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosConsultaArticulos.frx":2DC4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosConsultaArticulos.frx":42DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosConsultaArticulos.frx":57F32
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
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
      Height          =   315
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   6600
      Width           =   1815
   End
   Begin VB.TextBox txtDescripcion 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   4320
      TabIndex        =   7
      Top             =   360
      Width           =   5775
   End
   Begin VB.TextBox txtFabricante 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtBarras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   1320
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   2412
      Left            =   1080
      TabIndex        =   8
      Top             =   1320
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   4260
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   14737632
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Barras"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descripción"
         Object.Width           =   6950
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Precio"
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "# Ctd"
         Object.Width           =   1482
      EndProperty
   End
   Begin MSComctlLib.ListView lswSel 
      Height          =   2172
      Left            =   1080
      TabIndex        =   9
      Top             =   4320
      Width           =   9012
      _ExtentX        =   15901
      _ExtentY        =   3836
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   16711680
      BackColor       =   16777215
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6950
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Precio"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Cantidad"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbX 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   13
      Top             =   6855
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   582
      ButtonWidth     =   3201
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Nueva Consulta"
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Limpiar Información Actual"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Crear pedido"
            Key             =   "Pedido"
            Object.ToolTipText     =   "Crear Pedido"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Left            =   9000
      TabIndex        =   18
      Top             =   876
      Width           =   732
      _Version        =   1441793
      _ExtentX        =   1291
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Similares"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCodSim 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   2400
      TabIndex        =   15
      Top             =   876
      Width           =   5532
   End
   Begin VB.Label lblCodSel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   4440
      TabIndex        =   12
      Top             =   3876
      Width           =   5532
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Descripción"
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
      Index           =   3
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Fabricante"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Barras"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Código"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Resultados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   9972
      _Version        =   1441793
      _ExtentX        =   17590
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Articulos Seleccionados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmPosConsultaArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbBusqueda()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lblCodSim.Caption = ""
lblCodSim.Tag = ""
chkSimilar.Value = vbUnchecked

strSQL = "select cod_producto,cod_barras,cod_fabricante,descripcion,precio_regular,existencia,impuesto_ventas" _
       & " from pv_productos where estado = 'A'"

If txtCodigo.Text <> "" Then strSQL = strSQL & " and cod_producto like '" & txtCodigo.Text & "%'"
If txtDescripcion.Text <> "" Then strSQL = strSQL & " and descripcion like '%" & txtDescripcion.Text & "%'"
If txtBarras.Text <> "" Then strSQL = strSQL & " cod_barras like '" & txtBarras.Text & "%'"
If txtFabricante.Text <> "" Then strSQL = strSQL & " cod_fabricante like '" & txtFabricante.Text & "%'"

strSQL = strSQL & " order by descripcion"

lsw.ListItems.Clear
lsw.ForeColor = vbBlue
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 DoEvents
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Producto)
     itmX.SubItems(1) = rs!cod_barras
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = Format(rs!precio_regular + (rs!precio_regular * rs!impuesto_ventas / 100), "Standard")
     itmX.SubItems(4) = rs!Existencia
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault

End Sub


Private Sub sbTotales()
Dim curTotal As Currency, i As Integer

curTotal = 0

With lswSel.ListItems
 For i = 1 To .Count
      .Item(i).SubItems(4) = Format(CCur(.Item(i).SubItems(2)) * CCur(.Item(i).SubItems(3)), "Standard")
      curTotal = curTotal + CCur(.Item(i).SubItems(4))
 Next i
End With

txtTotal = Format(curTotal, "Standard")

End Sub


Private Sub chkSimilar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If chkSimilar.Value = vbUnchecked Then
  Call sbBusqueda
  Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select cod_producto,cod_barras,cod_fabricante,descripcion,precio_regular,existencia,impuesto_ventas" _
       & " from pv_productos where estado = 'A' and cod_producto <> '" & lblCodSim.Tag _
       & "' and similar in(select isnull(similar,0) from pv_productos" _
       & " where cod_producto = '" & lblCodSim.Tag & "') order by descripcion"

lsw.ListItems.Clear
lsw.ForeColor = vbBlack
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 DoEvents
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Producto)
     itmX.SubItems(1) = rs!cod_barras
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = Format(rs!precio_regular + (rs!precio_regular * rs!impuesto_ventas / 100), "Standard")
     itmX.SubItems(4) = rs!Existencia
 rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault

End Sub

Private Sub sbNuevo()

txtCodigo = ""
txtBarras = ""
txtFabricante = ""
txtDescripcion = ""

lsw.ListItems.Clear
lswSel.ListItems.Clear

lblCodSel.Tag = ""
lblCodSel.Caption = ""

txtDescripcion.SetFocus

End Sub


Private Sub lsw_Click()


If lsw.ListItems.Count = 0 Or chkSimilar.Value = vbChecked Then Exit Sub

lblCodSim.Caption = lsw.SelectedItem.SubItems(2)
lblCodSim.Tag = lsw.SelectedItem

chkSimilar.Value = vbUnchecked

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

On Error GoTo Errores
    lsw.SortKey = ColumnHeader.Index - 1
    
    If (lsw.SortOrder = lvwAscending) Then
        lsw.SortOrder = lvwDescending
    Else
        lsw.SortOrder = lvwAscending
    End If
    
    lsw.Sorted = True
Salir:
    Exit Sub
Errores:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub

Private Sub lsw_DblClick()
Dim itmX As ListItem, i As Integer
Dim vEncontro As Boolean

If lsw.ListItems.Count = 0 Then Exit Sub

vEncontro = False

With lswSel.ListItems
 For i = 1 To .Count
   If .Item(i).Text = lsw.SelectedItem Then
      vEncontro = True
      .Item(i).SubItems(3) = CCur(.Item(i).SubItems(3)) + 1
   End If
 Next i

 If Not vEncontro Then
    Set itmX = .Add(, , lsw.SelectedItem)
        itmX.SubItems(1) = lsw.SelectedItem.SubItems(2)
        itmX.SubItems(2) = lsw.SelectedItem.SubItems(3)
        itmX.SubItems(3) = 1
 End If

End With

Call sbTotales

End Sub


Private Sub lswSel_Click()
If lswSel.ListItems.Count = 0 Then Exit Sub

lblCodSel.Tag = lswSel.SelectedItem
lblCodSel.Caption = lswSel.SelectedItem.SubItems(1)
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbNuevo
End Sub

Private Sub tlbSumas_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

If lblCodSel.Tag = "" Then Exit Sub

With lswSel.ListItems
 For i = 1 To .Count
  If Trim(.Item(i).Text) = Trim(lblCodSel.Tag) Then
    Select Case Button.Key
      Case "Mas"
        .Item(i).SubItems(3) = CCur(.Item(i).SubItems(3)) + 1
      Case "Menos"
        .Item(i).SubItems(3) = CCur(.Item(i).SubItems(3)) - 1
      Case "Elimina"
        .Remove .Item(i).Index
        lblCodSel.Tag = ""
        lblCodSel.Caption = ""
        Exit For
    End Select
  End If
 Next i
End With

Call sbTotales

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
 Case "Nuevo"
   Call sbNuevo
 Case "Pedido"
End Select
End Sub

Private Sub txtBarras_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBusqueda
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBusqueda
End If
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBusqueda
End If
End Sub

Private Sub txtFabricante_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
 Call sbBusqueda
End If
End Sub
