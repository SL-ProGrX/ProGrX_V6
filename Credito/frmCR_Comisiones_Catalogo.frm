VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.ShortcutBar.v19.3.0.ocx"
Begin VB.Form frmCR_Comisiones_Catalogo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones de Colocación de Créditos"
   ClientHeight    =   8316
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   12936
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8316
   ScaleWidth      =   12936
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3492
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   12732
      _Version        =   1245187
      _ExtentX        =   22458
      _ExtentY        =   6159
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
      ItemCount       =   2
      Item(0).Caption =   "Porcentaje de Comisión"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "fpPorcentajes"
      Item(1).Caption =   "Lineas de Crédito Autorizadas"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3012
         Left            =   -67600
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   7332
         _Version        =   1245187
         _ExtentX        =   12933
         _ExtentY        =   5313
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin FPSpreadADO.fpSpread fpPorcentajes 
         Height          =   3012
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   7692
         _Version        =   524288
         _ExtentX        =   13568
         _ExtentY        =   5313
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_Comisiones_Catalogo.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3012
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   12732
      _Version        =   524288
      _ExtentX        =   22458
      _ExtentY        =   5313
      _StockProps     =   64
      BackColorStyle  =   1
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
      MaxCols         =   499
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Comisiones_Catalogo.frx":071F
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Catálogo de Comisiones de Colocación de Créditos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Index           =   2
      Left            =   2160
      TabIndex        =   0
      Top             =   420
      Width           =   10572
   End
   Begin XtremeShortcutBar.ShortcutCaption scDetalle 
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   12732
      _Version        =   1245187
      _ExtentX        =   22458
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_Comisiones_Catalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUltimaSelTipo As String
Dim vPaso As Boolean


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

Dim pCodigo As String, pCodigoDesc As String, pActivo As Integer
Dim pFechaIncio As Date, pBase As String, pCuenta As String
Dim pMovimiento As String


On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
pCodigo = Trim(vGrid.Text)
vGrid.col = 2
pCodigoDesc = Trim(vGrid.Text)
vGrid.col = 3

pFechaIncio = CDate(vGrid.Text)
vGrid.col = 4
pBase = Trim(vGrid.Text)
vGrid.col = 5
pCuenta = fxCntX_CuentaFormato(False, vGrid.Text, 0)
vGrid.col = 6
pActivo = vGrid.Value

strSQL = "exec spCrd_Comisiones_Cat_Registro '" & pCodigo & "','" & pCodigoDesc & "','" _
        & Format(pFechaIncio, "yyyy/mm/dd") & "','" & pBase & "','" & pCuenta _
       & "'," & pActivo & ",'" & glogon.Usuario & "','A'"
Call OpenRecordSet(rs, strSQL)

    vGrid.col = 1
    vGrid.Text = CStr(rs!Cod_Comision)
    
    pMovimiento = rs!Movimiento

rs.Close

Call Bitacora(pMovimiento, "Comisión de Crédito Id: " & vGrid.Text)

vGrid.col = 1
fxGuardar = 1
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

 
End Function


Private Sub Form_Activate()
vModulo = 3
End Sub



Private Sub fpPorcentajes_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim pLinea As Long, i As Integer


On Error GoTo vError

With fpPorcentajes

   .Row = .ActiveRow
   .col = 1
   If .Text = "" Then
      pLinea = 0
   Else
      pLinea = CLng(.Text)
   End If

    If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
       
       strSQL = "exec  spCrd_Comisiones_TP_Registra '" & scDetalle.Tag & "'," & pLinea
       .col = 2
       strSQL = strSQL & "," & CCur(.Text)
       .col = 3
       strSQL = strSQL & "," & CCur(.Text)
       .col = 4
       strSQL = strSQL & "," & CCur(.Text)
       .col = 5
       strSQL = strSQL & "," & CCur(.Text) & ",'" & glogon.Usuario & "', 'A'"
       
       Call OpenRecordSet(rs, strSQL)
            .col = 1
            pLinea = rs!Linea_Id
            .Text = CStr(pLinea)
       rs.Close
       
       Call Bitacora("Registra", "Comisiones, Tabla Porcentajes > Código: " & scDetalle.Tag & " > Id: " & pLinea)
       
       If .Row = .MaxRows Then
          .MaxRows = .MaxRows + 1
       End If
    
    End If


    'Inserta Linea
    If KeyCode = vbKeyInsert Then
        .MaxRows = .MaxRows + 1
        .InsertRows .ActiveRow, 1
        .Row = .ActiveRow
    End If

    'Borrar Linea
    If KeyCode = vbKeyDelete And pLinea > 0 Then
         i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
         If i = vbYes Then

            strSQL = "exec  spCrd_Comisiones_TP_Registra '" & scDetalle.Tag & "'," & pLinea _
                   & ",Null, Null, Null, Null,'" & glogon.Usuario & "', 'E'"
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Comisiones, Tabla Porcentajes > Código: " & scDetalle.Tag & " > Id: " & pLinea)
                    
            .DeleteRows .ActiveRow, 1
            If .MaxRows > 1 Then .MaxRows = .MaxRows - 1
            .Row = .ActiveRow
         End If
    End If



End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "exec spCrd_Comisiones_Lineas_Asigna_Registra '" & scDetalle.Tag & "','" & Item.Text & "','" _
        & glogon.Usuario & "','" & IIf(Item.Checked = True, "I", "E") & "'"
Call ConectionExecute(strSQL)


If Item.Checked Then
    Call Bitacora("Registra", "Comisiones Asignacion Línea Id: " & scDetalle.Tag & " .. Código: " & Item.Text)
Else
    Call Bitacora("Elimina", "Comisiones Asignacion Línea Id: " & scDetalle.Tag & " .. Código: " & Item.Text)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

If col <> 7 Then Exit Sub


vGrid.Row = Row
vGrid.col = 1

Call sbDetalle_Limpia

scDetalle.Tag = vGrid.Text
vGrid.col = 2
scDetalle.Caption = vGrid.Text

Call sbDetalle_Carga

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String
Dim i As Integer

On Error GoTo vError

If vGrid.ActiveCol >= vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 5) Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.ActiveCol
  Call sbgCntCuentaConsulta
   vGrid.Text = fxCntX_CuentaFormato(True, gBusquedas.Resultado, 0)
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        
        strSQL = "exec spCrd_Comisiones_Cat_Registro '" & vGrid.Text & "','','01/01/2020'" _
               & ",'','',0,'" & glogon.Usuario & "','E'"
        Call ConectionExecute(strSQL)

        Call Bitacora("Elimina", "Comisión de Crédito Id: " & vGrid.Text)
               
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
        
        Call sbDetalle_Limpia
     End If
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDetalle_Limpia()

scDetalle.Caption = "..."
scDetalle.Tag = ""

fpPorcentajes.MaxCols = 5
fpPorcentajes.MaxRows = 0

lsw.ListItems.Clear

tcMain.Item(0).Selected = True

End Sub

Private Sub sbDetalle_Carga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If scDetalle.Tag = "" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Comisiones_TP_Consulta '" & scDetalle.Tag & "','" & glogon.Usuario & "'"
Call sbCargaGrid(fpPorcentajes, 5, strSQL, True)


strSQL = "exec spCrd_Comisiones_Lineas_Asigna_Consulta '" & scDetalle.Tag & "'"
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
vPaso = True

Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Codigo)
       itmX.SubItems(1) = rs!Descripcion
       itmX.Checked = IIf(rs!Asignado = 1, True, False)
       
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


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Línea", 1000
    .Add , , "Descripción", 5600
End With
lsw.Checkboxes = True
    
vPaso = True
    strSQL = "select *" _
           & " from vCrd_Comisiones_Catalogo order by cod_comision"
    Call sbCargaGridLocal(vGrid, 7, strSQL)
vPaso = False

Call sbDetalle_Limpia

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strResTipo As String, vNota As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows


Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1 'Codigo de Comision
       vGrid.Text = Trim(rs!Cod_Comision)
     Case 2 'descripcion
       vGrid.Text = Trim(rs!Descripcion)
     Case 3 'Fecha Inicio
       vGrid.Text = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
     Case 4 'Base de Calculo
       vGrid.Text = Trim(rs!Base_Calculo)
     Case 5
       vGrid.Text = Trim(rs!cod_Cuenta_Mask)
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Cuenta: " & rs!Cuenta_Desc
     Case 6 'Activo
       vGrid.Value = rs!Activa
        
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Usuario " & IIf(IsNull(rs!Registro_Usuario), "...!", rs!Registro_Usuario) _
                         & vbCrLf & " Fecha " & IIf(IsNull(rs!Registro_Fecha), "...!", rs!Registro_Fecha)
     Case Else
    End Select
  Next i
  
  rs.MoveNext

  If Not rs.EOF Then
    vGrid.MaxRows = vGrid.MaxRows + 1
  End If
  

Loop

rs.Close

vGrid.MaxRows = vGrid.MaxRows + 1
vGrid.Row = vGrid.MaxRows
  
Me.MousePointer = vbDefault

End Sub
