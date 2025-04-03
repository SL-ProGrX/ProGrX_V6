VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Comites 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comités de Resolución de Créditos"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16140
   HelpContextID   =   3006
   Icon            =   "frmCR_Comites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   16140
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3375
      Left            =   7560
      TabIndex        =   6
      Top             =   5160
      Width           =   8535
      _Version        =   1572864
      _ExtentX        =   15055
      _ExtentY        =   5953
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
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   15975
      _Version        =   524288
      _ExtentX        =   28178
      _ExtentY        =   5318
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
      MaxCols         =   504
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Comites.frx":030A
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin FPSpreadADO.fpSpread fpGarantias 
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   7455
      _Version        =   524288
      _ExtentX        =   13150
      _ExtentY        =   5953
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
      SpreadDesigner  =   "frmCR_Comites.frx":0CC1
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption scComite 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   15975
      _Version        =   1572864
      _ExtentX        =   28178
      _ExtentY        =   661
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
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   4800
      Width           =   8415
      _Version        =   1572864
      _ExtentX        =   14843
      _ExtentY        =   661
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
      _Version        =   1572864
      _ExtentX        =   13356
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Limites de aprobación por Garantía"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.93
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
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   16815
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
Dim strSQL As String, rs As New ADODB.Recordset


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

Dim pComite As Long, pComiteDesc As String, pTipo As String, pActa As Long, pActivo As Integer, pLineaFiltra As Integer
Dim pAbreviatura As String, pOrden As String
Dim pNoAprobacion As Integer, pRngInicio As Currency, pRngCorte As Currency
Dim pMovimiento As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If vGrid.Text = "" Then
    pComite = 0
    pMovimiento = "Registra"
Else
    pComite = CLng(vGrid.Text)
    pMovimiento = "Modifica"
End If

vGrid.Col = 2
pComiteDesc = Trim(vGrid.Text)

vGrid.Col = 3
pActa = CLng(vGrid.Text)

vGrid.Col = 4
pAbreviatura = Trim(vGrid.Text)

vGrid.Col = 5
pOrden = Trim(vGrid.Text)

vGrid.Col = 6
pTipo = Trim(Mid(vGrid.Text, 1, 1))
vGrid.Col = 7
pNoAprobacion = CLng(vGrid.Text)
vGrid.Col = 8
pRngInicio = CCur(vGrid.Text)
vGrid.Col = 9
pRngCorte = CCur(vGrid.Text)
vGrid.Col = 10
pLineaFiltra = vGrid.Value
vGrid.Col = 11
pActivo = vGrid.Value

strSQL = "exec spCrd_Comites_Registro " & pComite & ",'" & pComiteDesc & "'," & pActa & ",'" & pTipo & "'," & pNoAprobacion _
       & "," & pRngInicio & "," & pRngCorte & "," & pLineaFiltra & "," & pActivo & ",'" & glogon.Usuario _
       & "', '" & pAbreviatura & "', '" & pOrden & "'"
Call OpenRecordSet(rs, strSQL)

    vGrid.Col = 1
    vGrid.Text = CStr(rs!id_Comite)

rs.Close

Call Bitacora(pMovimiento, "Comité Resolución Id: " & vGrid.Text)

MsgBox "Se " & pMovimiento & " Comité: " & pComiteDesc & " satisfactoriamente!", vbInformation

vGrid.Col = 1
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

Dim pMonto As String, pGarantia As String

If fpGarantias.ActiveCol = fpGarantias.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
   fpGarantias.Row = fpGarantias.ActiveRow
   fpGarantias.Col = 1
   pGarantia = fpGarantias.Text
   
   strSQL = "exec  spCrd_Comites_Garantias_Rangos_Registra " & scComite.Tag & ",'" & fpGarantias.Text & "'"
   
   
   fpGarantias.Col = 3
   strSQL = strSQL & "," & CCur(fpGarantias.Text)
   fpGarantias.Col = 4
   pMonto = fpGarantias.Text
   
   strSQL = strSQL & "," & CCur(fpGarantias.Text) & ",'" & glogon.Usuario & "'"
   
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Modifica", "Comité Rng Gar. " & pGarantia & ", Id Comite: " & scComite.Tag & " Mnt: " & pMonto)
   MsgBox "Se ha modificado Comité > Rango Garantía >  " & pGarantia & " > Id Comite: " & scComite.Tag & " Mnt: " & pMonto
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
    Call Bitacora("Registra", "Comités asignación línea, Id: " & scComite.Tag & " .. Código: " & Item.Text)
    MsgBox "Se ha vinculado la linea " & Item.Text & " al comité", vbInformation
Else
    Call Bitacora("Elimina", "Comités asignación línea, Id: " & scComite.Tag & " .. Código: " & Item.Text)
    MsgBox "Se ha Desvinculado la linea " & Item.Text & " al comité", vbInformation
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If vPaso Then Exit Sub

If Col <> 12 Then Exit Sub


vGrid.Row = Row
vGrid.Col = 1

Call sbDetalle_Limpia

scComite.Tag = vGrid.Text
vGrid.Col = 2
scComite.Caption = vGrid.Text

Call sbDetalle_Carga

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol >= vGrid.MaxCols - 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCargaCboTipos(6, vGrid.MaxRows, vGrid)
  End If
End If


'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este comité", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        strSQL = "exec spCrd_Comites_Elimina " & vGrid.Text & ", '" & glogon.Usuario & "'"
        Call OpenRecordSet(rs, strSQL)
        
        If rs!Pass = 1 Then
                    
            vGrid.Col = 1
            strSQL = vGrid.Text
    
            vGrid.DeleteRows vGrid.ActiveRow, 1
            vGrid.MaxRows = vGrid.MaxRows - 1
            
            If vGrid.MaxRows <= 0 Then
              vGrid.MaxRows = 1
            End If
            
            Call Bitacora("Elimina", "Comites Id: " & strSQL)
            
            MsgBox "Comité Id: " & strSQL & ", Eliminado Satisfactoriamente!", vbInformation
        Else
            MsgBox rs!Mensaje, vbExclamation
        End If


     End If
End If



End Sub

Private Sub sbCargaCboTipos(vCol As Integer, vRow As Long, vGrid As Object)
Dim strResultado As String, rs As New ADODB.Recordset, strSQL As String

vGrid.Col = vCol
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
    strSQL = "select Id_comite, Descripcion, Acta, Abreviatura, Orden, Tipo_aprobacion" _
           & ", NAPROBACIONES, RNG_INICIO,RNG_CORTE,LINEA_FILTRA, ESTADO" _
           & " From comites order by id_comite"
    Call sbCargaGridLocal(vGrid, 12, strSQL)
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
  
  vGrid.Col = 6
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = strResTipo
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = strUltimaSelTipo
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1 'id_comite
       vGrid.Text = CStr(rs!id_Comite)
     Case 2 'descripcion
       vGrid.Text = CStr(rs!Descripcion)
     Case 3 'acta
       vGrid.Text = CStr(rs!acta)
     
     Case 4 'Abreviatura
       vGrid.Text = CStr(rs!Abreviatura)
     
     Case 5 'Orden
       vGrid.Text = CStr(rs!Orden)
     
     
     Case 6 'Tipo
        Select Case rs!tipo_aprobacion
          Case "E" 'Simple
            vGrid.Text = "Ejecutivo"
          Case "M" 'Mancomunado
            vGrid.Text = "Mancomunado"
          Case Else
            vGrid.Text = ""
        End Select
        
     Case 7
       vGrid.Text = CStr(rs!NAprobaciones)
     
     Case 8
       vGrid.Text = Format(rs!Rng_Inicio, "Standard")
     
     Case 9
       vGrid.Text = Format(rs!Rng_Corte, "Standard")
       
     Case 10 'Filtra Linea
       vGrid.Value = rs!Linea_Filtra
     Case 11 'Activo
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
Dim rs As New ADODB.Recordset, strSQL As String

Dim i As Integer

On Error GoTo vError

Select Case Button.Key
   Case "insertar"
    vGrid.MaxRows = vGrid.MaxRows + 1
    Call sbCargaCboTipos(6, vGrid.MaxRows, vGrid)
    
   Case "borrar"
   
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        
        strSQL = "exec spCrd_Comites_Elimina " & vGrid.Text
        Call OpenRecordSet(rs, strSQL)
        
        If rs!Pass = 1 Then
            strSQL = vGrid.Text
            vGrid.Col = 2
            Call Bitacora("Elimina", "Comité : " & strSQL & " - " & vGrid.Text)
            vGrid.Col = 1
            
            vPaso = True
                strSQL = "select Id_comite, Descripcion, Acta, Abreviatura, Orden, Tipo_aprobacion" _
                       & ", NAPROBACIONES, RNG_INICIO,RNG_CORTE,LINEA_FILTRA, ESTADO" _
                       & " From comites order by id_comite"
                Call sbCargaGridLocal(vGrid, 12, strSQL)
            vPaso = False
        
        Else
            MsgBox rs!Mensaje, vbExclamation
        End If
        
        
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
