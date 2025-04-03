VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_Telefonos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teléfonos de la persona"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   8715
   HelpContextID   =   1012
   Icon            =   "frmAF_Telefonos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8175
      _Version        =   524288
      _ExtentX        =   14420
      _ExtentY        =   4895
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
      MaxCols         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_Telefonos.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfonos de la Persona"
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
      Index           =   0
      Left            =   1884
      TabIndex        =   0
      Top             =   360
      Width           =   6492
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_Telefonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Dim mTipos As String, mUltimoSel As String

Private Sub Form_Load()

Me.Caption = "Teléfonos [Cédula : " & GLOBALES.gCedulaActual & "]"

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

mTipos = ""
mUltimoSel = ""

strSQL = "select rtrim(nombreTipoTelefono) as 'Tipo' From AFI_TIPOS_TELEFONOS" _
       & " Where Activo = 1 ORDER BY Prioridad"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
   If mTipos = "" Then
    mUltimoSel = rs!Tipo
   End If
   
   mTipos = mTipos & Chr$(9) & rs!Tipo

   rs.MoveNext
Loop
rs.Close

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select T.Telefono, T.Tipo, T.Numero, T.Ext, T.Contacto, T.Usuario, T.Fecha, Tt.NombreTipoTelefono as 'TipoDesc'" _
       & " , dbo.MyGetdate() as FechaServidor" _
       & " from Telefonos T inner join AFI_TIPOS_TELEFONOS Tt on T.Tipo = Tt.IdTipoTelefono " _
       & " where cedula = '" & GLOBALES.gCedulaActual & "'"
Call sbCargaGridLocal(vGrid, 4, strSQL)

End Sub


Private Function fxTipoTelefono(vTipo As String) As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim pResult As Integer

On Error GoTo vError

pResult = 1

strSQL = "select IdTipoTelefono as 'Tipo' from AFI_TIPOS_TELEFONOS where NombreTipoTelefono = '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)

pResult = rs!Tipo


fxTipoTelefono = pResult
Exit Function

vError:
    fxTipoTelefono = pResult

End Function

Private Sub sbCargaGridLocal(ByRef vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  vGrid.col = 1
  vGrid.CellType = CellTypeComboBox
  
  vGrid.TypeComboBoxList = mTipos
  vGrid.TypeComboBoxEditable = False
  
  vGrid.Text = mUltimoSel
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = rs!TipoDesc & ""
        vGrid.CellTag = CStr(rs!Telefono)
     Case 2
        vGrid.Text = CStr(rs!Numero & "")
     Case 3
        vGrid.Text = CStr(rs!Ext & "")
     Case 4
        vGrid.Text = CStr(rs!contacto & "")
    End Select
  Next i
  
    vGrid.col = 4
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario " & IIf(IsNull(rs!Usuario), "...!", rs!Usuario) _
                     & vbCrLf & " Fecha " & IIf(IsNull(rs!fecha), "...!", rs!fecha)
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

  vGrid.Row = vGrid.MaxRows
  vGrid.col = 1
  vGrid.CellType = CellTypeComboBox
  
  vGrid.TypeComboBoxList = mTipos
  vGrid.TypeComboBoxEditable = False
  vGrid.Text = mUltimoSel

Me.MousePointer = vbDefault

End Sub

Private Sub sbCboTiposTelefonos(vCol As Integer, vRow As Long, vGrid As Object)

vGrid.col = vCol
vGrid.Row = vRow
vGrid.CellType = CellTypeComboBox

vGrid.TypeComboBoxList = mTipos
vGrid.TypeComboBoxEditable = False
vGrid.Text = mUltimoSel

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.Row = vGrid.MaxRows
    vGrid.col = 1
    If vGrid.CellTag <> "" Then
        vGrid.MaxRows = vGrid.MaxRows + 1
        Call sbCboTiposTelefonos(1, vGrid.MaxRows, vGrid)
    End If
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este telefono ... ", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        If vGrid.CellTag <> "" Then
            strSQL = "delete telefonos where telefono = " & vGrid.CellTag
            Call ConectionExecute(strSQL)
                    
                strSQL = "select T.Telefono, T.Tipo, T.Numero, T.Ext, T.Contacto, T.Usuario, T.Fecha, Tt.NombreTipoTelefono as 'TipoDesc'" _
                       & " , dbo.MyGetdate() as FechaServidor" _
                       & " from Telefonos T inner join AFI_TIPOS_TELEFONOS Tt on T.Tipo = Tt.IdTipoTelefono " _
                       & " where cedula = '" & GLOBALES.gCedulaActual & "'"
                Call sbCargaGridLocal(vGrid, 4, strSQL)

        
            strSQL = vGrid.CellTag
            vGrid.col = 2
            Call Bitacora("Elimina", "Número Teléfono: " & vGrid.Text & " Ced: " & GLOBALES.gCedulaActual & " ID." & strSQL)
        
        End If
     End If
  
  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp


End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        vGrid.CellTag = CStr(i)
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
          Call sbCboTiposTelefonos(1, vGrid.MaxRows, vGrid)
        End If
  End If 'Actualiza o Inserta
End If


'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    Call sbCboTiposTelefonos(1, vGrid.ActiveRow, vGrid)
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1

       If vGrid.CellTag = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
       
       strSQL = "delete telefonos where telefono = " & vGrid.CellTag
       Call ConectionExecute(strSQL)
        
        strSQL = vGrid.CellTag
        vGrid.col = 2
        Call Bitacora("Elimina", "Número Teléfono.: " & vGrid.Text & " Id.: " & strSQL)

        If vParametros.BitacoraEspecial Then
           Call sbgAFIBitacora("10", "Elimina Teléfono.: " & vGrid.Text & " Id.: " & strSQL, Trim(GLOBALES.gCedulaActual))
        End If
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxGuardar() As Long

'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

vGrid.col = 2
If Len(vGrid.Text) <> 8 Then
    MsgBox "Número de Teléfono no es válido!", vbExclamation
    Exit Function
End If

vGrid.col = 3
If Len(vGrid.Text) > 4 Then
    MsgBox "La Extensión de Teléfono no es válida!", vbExclamation
    Exit Function
End If


vGrid.col = 1



If vGrid.CellTag = "" Then
    vGrid.col = 1
    strSQL = "insert telefonos(cedula,tipo,numero,ext,contacto,usuario,fecha) values('" & GLOBALES.gCedulaActual & "','" _
           & fxTipoTelefono(vGrid.Text) & "','"
    vGrid.col = 2
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & "',"
    vGrid.col = 3
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & ",'"
    vGrid.col = 4
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & "','" & glogon.Usuario & "',dbo.MyGetdate())"
    
    Call ConectionExecute(strSQL)
  
    strSQL = "select max(telefono) as ultimo from telefonos where cedula = '" & GLOBALES.gCedulaActual & "'"
    Call OpenRecordSet(rs, strSQL)
      vGrid.col = 1
      vGrid.CellTag = CStr(rs!ultimo)
    rs.Close
   
    strSQL = vGrid.CellTag
    
    vGrid.col = 2
    Call Bitacora("Registra", "Número Teléfono.: " & vGrid.Text & " Id.: " & strSQL)
   
    If vParametros.BitacoraEspecial Then
       Call sbgAFIBitacora("09", "Registra Teléfono.: " & vGrid.Text & " Id.: " & strSQL, Trim(GLOBALES.gCedulaActual))
    End If
   
   Else 'Actualizar

    vGrid.col = 2
    strSQL = "update telefonos set numero = '" & IIf((vGrid.Text = ""), 0, vGrid.Text) & "',ext = "
    vGrid.col = 3
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & ",contacto = '"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Text & "',tipo = '"
    vGrid.col = 1
    strSQL = strSQL & fxTipoTelefono(vGrid.Text) & "',usuario = '" & glogon.Usuario & "',fecha = dbo.MyGetdate() where telefono = " & vGrid.CellTag
    Call ConectionExecute(strSQL)
    
    strSQL = vGrid.CellTag
    
    vGrid.col = 2
    Call Bitacora("Modifica", "Número Teléfono: " & vGrid.Text & " Id.: " & strSQL)
    
    If vParametros.BitacoraEspecial Then
       Call sbgAFIBitacora("11", "Modifica Teléfono.: " & vGrid.Text & " Id.: " & strSQL, Trim(GLOBALES.gCedulaActual))
    End If
    
   End If


    vGrid.col = 4
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario " & glogon.Usuario _
                     & vbCrLf & " Fecha " & Format(Date, "dd/mm/yyyy")

   vGrid.col = 1
   fxGuardar = vGrid.CellTag
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function

