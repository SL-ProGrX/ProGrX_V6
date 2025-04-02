VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmIVR_Cat_Emisores 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SCGI Emisores"
   ClientHeight    =   6645
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9855
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5175
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   9375
      _Version        =   524288
      _ExtentX        =   16531
      _ExtentY        =   9123
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmIVR_Cat_Emisores.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   6732
      _Version        =   1310722
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Emisores de Titulos"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   13932
   End
End
Attribute VB_Name = "frmIVR_Cat_Emisores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCboSectores As String


Private Sub Form_Activate()
vModulo = 22
End Sub

Private Sub sbConsulta(Optional pInicial As Integer = 0)
Dim strSQL As String, rs As New ADODB.Recordset

If pInicial = 1 Then
    strSQL = "select rtrim(COD_SECTOR) + ' - ' + rtrim(DESCRIPCION) as 'ItmX'" _
           & " FROM IVR_SECTORES ORDER BY COD_SECTOR"
    Call OpenRecordSet(rs, strSQL)
    
    mCboSectores = ""
        
    Do While Not rs.EOF
      If Len(mCboSectores) = 0 Then
        mCboSectores = Chr$(9) & rs!itmX
      Else
        mCboSectores = mCboSectores & Chr$(9) & rs!itmX
      End If
      rs.MoveNext
    Loop
    rs.Close
    
End If

 
strSQL = "select * from vIVR_EMISORES" _
      & " order by COD_EMISOR"
Call sbCargaGridLocal(vGrid, 4, strSQL)

End Sub




Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset
Dim i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
  
vGrid.Col = 3
vGrid.CellType = CellTypeComboBox
vGrid.TypeComboBoxList = mCboSectores
vGrid.TypeComboBoxEditable = False
  
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
        
  vGrid.Col = 3
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mCboSectores
  vGrid.TypeComboBoxEditable = False
 
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
       vGrid.Text = rs!COD_EMISOR
     
     Case 2
       vGrid.Text = rs!Descripcion
        
     Case 3 'Sector
        vGrid.Text = rs!Sector_ItmX
     
     Case 4 'Activa
       vGrid.Value = rs!ACTIVO
     
     Case Else
    End Select
  Next i
  
    vGrid.Col = vGrid.MaxCols
    vGrid.TextTip = TextTipFixed
    vGrid.TextTipDelay = 1000
    vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
    vGrid.CellNote = "Usuario: " & IIf(IsNull(rs!Registro_Usuario), "...!", rs!Registro_Usuario) _
                     & vbCrLf & "Fecha: " & IIf(IsNull(rs!Registro_Fecha), "...!", rs!Registro_Fecha) _
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 3
  vGrid.CellType = CellTypeComboBox
  vGrid.TypeComboBoxList = mCboSectores
  vGrid.TypeComboBoxEditable = False
  
  rs.MoveNext

Loop
rs.Close

   
Me.MousePointer = vbDefault

End Sub





Private Sub Form_Load()

vModulo = 22

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call sbConsulta(1)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from IVR_EMISORES " _
       & " where COD_EMISOR = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into IVR_EMISORES(COD_EMISOR,DESCRIPCION, COD_SECTOR, ACTIVO, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & SIFGlobal.fxCodText(vGrid.Text) & "',"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "Emisor de Titulos Valores:  " & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update IVR_EMISORES set descripcion = '" & vGrid.Text & "', COD_SECTOR = '"
  vGrid.Col = 3
  strSQL = strSQL & SIFGlobal.fxCodText(vGrid.Text) & "', ACTIVO = "
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Value & " where COD_EMISOR = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "'"
  
  Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Emisor de Titulos Valores:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  
    vGrid.Col = 3
    vGrid.CellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mCboSectores
    vGrid.TypeComboBoxEditable = False
  
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow

    vGrid.Col = 3
    vGrid.CellType = CellTypeComboBox
    vGrid.TypeComboBoxList = mCboSectores
    vGrid.TypeComboBoxEditable = False

End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete IVR_EMISORES where COD_EMISOR = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Emisor de Titulos Valores:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If

End Sub

