VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCR_Comites 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comités de Resolución de Créditos"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14430
   HelpContextID   =   3006
   Icon            =   "frmCR_Comites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   14430
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2772
      Left            =   7560
      TabIndex        =   6
      Top             =   5160
      Width           =   6732
      _Version        =   1441793
      _ExtentX        =   11874
      _ExtentY        =   4890
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
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   264
      Left            =   12840
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1248
      _ExtentX        =   2196
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "insertar"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
            Object.ToolTipText     =   "Borra el registro en pantalla de la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
            Object.ToolTipText     =   "Ayuda General"
            Object.Tag             =   "1"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3012
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   14172
      _Version        =   524288
      _ExtentX        =   24998
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
      MaxCols         =   502
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Comites.frx":030A
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin FPSpreadADO.fpSpread fpGarantias 
      Height          =   2772
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   7452
      _Version        =   524288
      _ExtentX        =   13144
      _ExtentY        =   4890
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Comites.frx":0C5F
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scComite 
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   14172
      _Version        =   1441793
      _ExtentX        =   24998
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
      Height          =   372
      Left            =   7680
      TabIndex        =   4
      Top             =   4800
      Width           =   6612
      _Version        =   1441793
      _ExtentX        =   11663
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Líneas de Crédito Autorizadas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   7572
      _Version        =   1441793
      _ExtentX        =   13356
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Limites de aprobación por Garantía"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comité de Resolución"
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
      Height          =   480
      Index           =   2
      Left            =   2160
      TabIndex        =   2
      Top             =   420
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_Comites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUltimaSelTipo As String
Dim vPaso As Boolean


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

Dim pComite As Long, pComiteDesc As String, pTipo As String, pActa As Long, pActivo As Integer, pLineaFiltra As Integer
Dim pNoAprobacion As Integer, pRngInicio As Currency, pRngCorte As Currency
Dim pMovimiento As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If vGrid.Text = "" Then
    pComite = 0
    pMovimiento = "Registra"
Else
    pComite = CLng(vGrid.Text)
    pMovimiento = "Modifica"
End If

vGrid.col = 2
pComiteDesc = Trim(vGrid.Text)
vGrid.col = 3
pActa = CLng(vGrid.Text)
vGrid.col = 4
pTipo = Trim(Mid(vGrid.Text, 1, 1))
vGrid.col = 5
pNoAprobacion = CLng(vGrid.Text)
vGrid.col = 6
pRngInicio = CCur(vGrid.Text)
vGrid.col = 7
pRngCorte = CCur(vGrid.Text)
vGrid.col = 8
pLineaFiltra = vGrid.Value
vGrid.col = 9
pActivo = vGrid.Value

strSQL = "exec spCrd_Comites_Registro " & pComite & ",'" & pComiteDesc & "'," & pActa & ",'" & pTipo & "'," & pNoAprobacion _
       & "," & pRngInicio & "," & pRngCorte & "," & pLineaFiltra & "," & pActivo & ",'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

    vGrid.col = 1
    vGrid.Text = CStr(rs!id_Comite)

rs.Close

Call Bitacora(pMovimiento, "Comité Resolución Id: " & vGrid.Text)

vGrid.col = 1
fxGuardar = vGrid.Text
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

 
End Function


Private Sub Form_Activate()
vModulo = 3
End Sub



Private Sub fpGarantias_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String

On Error GoTo vError

If fpGarantias.ActiveCol = fpGarantias.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
   fpGarantias.Row = fpGarantias.ActiveRow
   fpGarantias.col = 1
   
   strSQL = "exec  spCrd_Comites_Garantias_Rangos_Registra " & scComite.Tag & ",'" & fpGarantias.Text & "'"
   fpGarantias.col = 3
   strSQL = strSQL & "," & CCur(fpGarantias.Text)
   fpGarantias.col = 4
   strSQL = strSQL & "," & CCur(fpGarantias.Text) & ",'" & glogon.Usuario & "'"
   
   Call ConectionExecute(strSQL)
End If

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

strSQL = "exec spCrd_Comites_Lineas_Asigna_Registra " & scComite.Tag & ",'" & Item.Text & "','" _
        & glogon.Usuario & "','" & IIf(Item.Checked = True, "I", "E") & "'"
Call ConectionExecute(strSQL)


If Item.Checked Then
    Call Bitacora("Registra", "Comités asignación línea, Id:" & scComite.Tag & " .. Código:" & Item.Text)
Else
    Call Bitacora("Elimina", "Comités asignación línea, Id:" & scComite.Tag & " .. Código:" & Item.Text)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

If col <> 10 Then Exit Sub


vGrid.Row = Row
vGrid.col = 1

Call sbDetalle_Limpia

scComite.Tag = vGrid.Text
vGrid.col = 2
scComite.Caption = vGrid.Text

Call sbDetalle_Carga

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGrid.ActiveCol >= vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCargaCboTipos(4, vGrid.MaxRows, vGrid)
  End If
End If

End Sub

Private Sub sbCargaCboTipos(vCol As Integer, vRow As Long, vGrid As Object)
Dim strResultado As String, rs As New ADODB.Recordset, strSQL As String

vGrid.col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

If strUltimaSelTipo = "" Then strUltimaSelTipo = "Ejecutivo"

strResultado = "Ejecutivo" & Chr$(9) & "Mancomunado"

vGrid.TypeComboBoxList = strResultado
vGrid.TypeComboBoxEditable = False
vGrid.Text = strUltimaSelTipo

End Sub

Private Sub sbDetalle_Limpia()

scComite.Caption = "..."
scComite.Tag = ""

fpGarantias.MaxCols = 4
fpGarantias.MaxRows = 0

lsw.ListItems.Clear

End Sub

Private Sub sbDetalle_Carga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If scComite.Tag = "" Then Exit Sub

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Comites_Garantias_Rangos_Consulta " & scComite.Tag & ",'" & glogon.Usuario & "'"
Call sbCargaGrid(fpGarantias, 4, strSQL, True)
fpGarantias.MaxRows = fpGarantias.MaxRows - 1


strSQL = "exec spCrd_Comites_Lineas_Asigna_Consulta " & scComite.Tag
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
    .Add , , "Línea", 900
    .Add , , "Descripción", 4900
End With
lsw.Checkboxes = True
 
Call sbToolBarIconos(tlbPrincipal, False)

If tlbPrincipal.Buttons(1).Enabled = False Then vGrid.Enabled = False
    
vPaso = True
    strSQL = "select id_comite,descripcion,acta,tipo_aprobacion,NAPROBACIONES,RNG_INICIO,RNG_CORTE,LINEA_FILTRA, ESTADO" _
           & " from comites order by id_comite"
    Call sbCargaGridLocal(vGrid, 10, strSQL)
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

strResTipo = "Ejecutivo" & Chr$(9) & "Mancomunado"


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 4
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = strResTipo
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSelTipo
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1 'id_comite
       vGrid.Text = CStr(rs!id_Comite)
     Case 2 'descripcion
       vGrid.Text = CStr(rs!Descripcion)
     Case 3 'acta
       vGrid.Text = CStr(rs!acta)
     Case 4 'Tipo
        Select Case rs!tipo_aprobacion
          Case "E" 'Simple
            vGrid.Text = "Ejecutivo"
          Case "M" 'Mancomunado
            vGrid.Text = "Mancomunado"
          Case Else
            vGrid.Text = ""
        End Select
     Case 5
       vGrid.Text = CStr(rs!NAprobaciones)
     Case 6
       vGrid.Text = Format(rs!Rng_Inicio, "Standard")
     Case 7
       vGrid.Text = Format(rs!Rng_Corte, "Standard")
       
     Case 8 'Filtra Linea
       vGrid.Value = rs!Linea_Filtra
     Case 9 'Activo
       vGrid.Value = rs!Estado
     Case Else
    End Select
  Next i
  
  rs.MoveNext

  If Not rs.EOF Then
    vGrid.MaxRows = vGrid.MaxRows + 1
  End If
  

Loop

rs.Close

vGrid.Row = vGrid.MaxRows
  
Me.MousePointer = vbDefault

End Sub


Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error GoTo vError

Select Case Button.Key
   Case "insertar"
    vGrid.MaxRows = vGrid.MaxRows + 1
    Call sbCargaCboTipos(4, vGrid.MaxRows, vGrid)
    
   Case "borrar"
   
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete comites where id_comite = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 2
        Call Bitacora("Elimina", "Comité : " & strSQL & " - " & vGrid.Text)
        vGrid.col = 1
        
        
        strSQL = "select id_comite,descripcion,acta,tipo_aprobacion,NAPROBACIONES,RNG_INICIO,RNG_CORTE,LINEA_FILTRA,estado" _
               & " from comites order by id_comite"
        Call sbCargaGridLocal(vGrid, 10, strSQL)
     End If
   
   Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

Exit Sub

vError:
  Me.MousePointer = vbDefault
End Sub


Private Function fxCodigoCboGrid(strCodigo As String) As String
Dim i As Integer, strResultado As String, blnPaso As Boolean

blnPaso = True
strResultado = ""
i = 1
Do While blnPaso
   If Mid(strCodigo, i, 1) <> "-" Then
     strResultado = strResultado & Mid(strCodigo, i, 1)
   Else
     blnPaso = False
   End If
   i = i + 1
Loop

fxCodigoCboGrid = strResultado

End Function
