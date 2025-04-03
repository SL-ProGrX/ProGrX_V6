VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaSubDesembolsos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraBusca 
      Caption         =   "Seleccione el Acredor"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   10680
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   10095
      Begin MSComctlLib.ListView lsw 
         Height          =   4335
         Left            =   1920
         TabIndex        =   4
         Top             =   840
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   7646
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
            Text            =   "Código"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtCriterio 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   7335
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   600
         X2              =   8880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label2 
         Caption         =   "Resultados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   600
         X2              =   1920
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Criterio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image imgCerrar 
         Height          =   480
         Left            =   9360
         Picture         =   "frmPreaSubDesembolsos.frx":0000
         ToolTipText     =   "Cerrar Busqueda"
         Top             =   360
         Width           =   480
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5295
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   9975
      _Version        =   524288
      _ExtentX        =   17595
      _ExtentY        =   9340
      _StockProps     =   64
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaSubDesembolsos.frx":09AD
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   8400
      TabIndex        =   7
      Top             =   6840
      Width           =   1455
      _Version        =   1572864
      _ExtentX        =   2566
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   315
      Left            =   7080
      TabIndex        =   8
      Top             =   6840
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Totales ..:"
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
      Left            =   5640
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desembolsos"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3972
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPreaSubDesembolsos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCol As Integer, vRow As Integer

Public mCambios As Boolean

Private Sub Form_Load()
Dim strSQL As String

mCambios = False

Me.Caption = "Expediente : " & gPreAnalisis.Expediente

vGrid.AppearanceStyle = AppearanceStyleVisualStyles

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

strSQL = "select * " _
       & " from CRD_PREA_DETALLE_DESEMBOLSOS" _
       & " where cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
Call sbCargaGridLocal(vGrid, 5, strSQL)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

mCambios = True

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
 GoTo salir
End If

If vGrid.Text = "" Then 'Insertar
 
  strSQL = "select isnull(max(Idx),0)+1 as IdxC from CRD_PREA_DETALLE_DESEMBOLSOS " _
         & " where cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
  Call OpenRecordSet(rs, strSQL)
    vGrid.Text = CStr(rs!IdxC)
  rs.Close

  strSQL = "insert into CRD_PREA_DETALLE_DESEMBOLSOS(cod_preAnalisis,IdX,cod_acredor,ordinario,descripcion" _
         & ",cuota,monto) values('" & gPreAnalisis.Expediente & "'," & vGrid.Text & ",'" & vGrid.CellTag & "',"

  vGrid.Col = 2
  strSQL = strSQL & vGrid.Value & ",'"
  
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ","
  
  vGrid.Col = 5
  strSQL = strSQL & CCur(vGrid.Text) & ")"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "PreAnalisis / Desembolso Id.: " & vGrid.Text & " Exp." & gPreAnalisis.Expediente)

Else 'Actualizar

 vGrid.Col = 1
 strSQL = "update CRD_PREA_DETALLE_DESEMBOLSOS set cod_acredor = '" & vGrid.CellTag & "', Ordinario = "
 vGrid.Col = 2
 strSQL = strSQL & vGrid.Value & ",Descripcion = '"
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Text & "', cuota = "
 vGrid.Col = 4
 strSQL = strSQL & CCur(vGrid.Text) & ", monto = "
 vGrid.Col = 5
 strSQL = strSQL & CCur(vGrid.Text) & " where cod_preAnalisis = '" & gPreAnalisis.Expediente & "' and Idx = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "PreAnalisis / Desembolso Id.: " & vGrid.Text & " Exp." & gPreAnalisis.Expediente)

End If

Call sbCalculaTotales
fxGuardar = 1

salir:
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

mCambios = True

If Not ValidaEstadoPreanalisis(gPreAnalisis.Estado) Then
 GoTo salir
End If
i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   strSQL = "delete CRD_PREA_DETALLE_DESEMBOLSOS where Idx = " & vGrid.Text _
          & " and cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
   Call ConectionExecute(strSQL)
   strSQL = vGrid.Text
   vGrid.Col = 1
   Call Bitacora("Elimina", "PreAnalisis / Desembolso Id.: " & vGrid.Text & " Exp." & gPreAnalisis.Expediente)
   
   vGrid.DeleteRows vGrid.ActiveRow, 1
   vGrid.MaxRows = vGrid.MaxRows - 1
   If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
   
   
   Call sbCalculaTotales
End If

salir:

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Unload(Cancel As Integer)
  GLOBALES.gTag = txtCuota.Text
  GLOBALES.gTag2 = txtMonto.Text
End Sub

Private Sub imgCerrar_Click()
 fraBusca.Visible = False
End Sub


Private Sub lsw_Click()

If lsw.ListItems.Count = 0 Then Exit Sub

vGrid.Row = vRow
vGrid.Col = 1
vGrid.CellTag = lsw.SelectedItem
vGrid.Col = 3
vGrid.Text = lsw.SelectedItem.SubItems(1)

vGrid.Col = 6 'MODIFICA_NOMBRE_GIRO
vGrid.Text = lsw.SelectedItem.SubItems(2)
If lsw.SelectedItem.SubItems(2) = "1" Then
    vGrid.Col = 3
    vGrid.Lock = False
Else
    vGrid.Col = 3
    vGrid.Lock = True
End If
fraBusca.Visible = False

End Sub


Private Sub txtCriterio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  Call sbConsulta
End If
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
  Call sbBorrar
End If

'Consulta
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
   
   vRow = vGrid.ActiveRow
   vCol = vGrid.ActiveCol
   
   txtCriterio.Text = ""
   
   Call sbConsulta
      
End If

End Sub

Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim x As String

If Col = 2 And NewCol = 3 And (Row = NewRow) Then
   vGrid.Row = Row
   vGrid.Col = Col
   x = vGrid.Text
   
   If x = "1" Or x = "0" Then x = ""
   
   vGrid.Col = NewCol
   If vGrid.Text = "" Then
     vGrid.Text = x
   End If


    If NewCol = 3 And vGrid.Text = "" Then
       vRow = NewRow
       vCol = NewCol
        txtCriterio.Text = ""
        Call sbConsulta
    End If

End If

End Sub



Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim i As Integer, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1 'Idx y el Tag con el Cod_Acredor
        vGrid.Text = CStr(rs!IDX)
        vGrid.CellTag = CStr(rs!cod_Acredor)
        
     Case 2
        vGrid.Text = CStr(rs!Ordinario)
     Case 3
        vGrid.Text = CStr(rs!Descripcion)
     Case 4
        vGrid.Text = CStr(rs!Cuota)
     Case 5
        vGrid.Text = CStr(rs!Monto)
    End Select
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Call sbCalculaTotales

Me.MousePointer = vbDefault

End Sub

Private Sub sbCalculaTotales()
Dim i As Integer, curCuota As Currency, curMonto As Currency

curCuota = 0
curMonto = 0

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 4
  curCuota = curCuota + IIf((vGrid.Text = ""), 0, vGrid.Text)
  vGrid.Col = 5
  curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
Next i


txtCuota.Text = Format(curCuota, "Standard")
txtMonto.Text = Format(curMonto, "Standard")

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

fraBusca.Top = vGrid.Top
fraBusca.Left = vGrid.Left
fraBusca.Width = vGrid.Width
fraBusca.Height = vGrid.Height

fraBusca.Visible = True

lsw.ListItems.Clear
txtCriterio.SetFocus

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 2

If vGrid.Value = 1 Then
   strSQL = "select cod_acredor as Codigo, Nombre as Nombre, MODIFICA_NOMBRE_GIRO from crd_prea_Acredores" _
          & " where activo = 1"
   If txtCriterio.Text <> "" Then strSQL = strSQL & " and nombre like '%" & txtCriterio.Text & "%'"

Else
   strSQL = "select cod_condeb as codigo,descripcion as nombre,0 as MODIFICA_NOMBRE_GIRO From Concepto_Desemb" _
          & " Where retiene = 1"
   
   If txtCriterio.Text <> "" Then strSQL = strSQL & " and descripcion like '%" & txtCriterio.Text & "%'"
End If

rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Nombre
     itmX.SubItems(2) = rs!MODIFICA_NOMBRE_GIRO
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
