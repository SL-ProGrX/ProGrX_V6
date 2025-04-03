VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCO_Grupos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Cobros"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8745
   ScaleWidth      =   11310
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   4471
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
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   1
      Top             =   5640
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Gestiones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3735
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   6855
      _Version        =   524288
      _ExtentX        =   12091
      _ExtentY        =   6588
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
      MaxCols         =   483
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_Grupos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   4095
      _Version        =   1572864
      _ExtentX        =   7223
      _ExtentY        =   661
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   1
      Left            =   5280
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Arreglos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   2
      Left            =   6600
      TabIndex        =   5
      Top             =   5640
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Causas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnSel 
      Height          =   375
      Index           =   3
      Left            =   7560
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Usuarios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   492
      Left            =   1800
      TabIndex        =   8
      Top             =   240
      Width           =   6732
      _Version        =   1572864
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Grupos de Cobros"
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
   Begin XtremeShortcutBar.ShortcutCaption scTipo 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   11055
      _Version        =   1572864
      _ExtentX        =   19500
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "(Seleccione un Grupo)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCO_Grupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbConsulta_Asignacion()
Dim pTipo As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

If scTipo.Tag = "" Then
  Exit Sub
End If

vPaso = True

txtFiltro.Text = fxSysCleanTxtInject(txtFiltro.Text)

Select Case True
    Case btnSel(0).Checked  'Gestiones
        pTipo = "GES"
    Case btnSel(1).Checked 'Arreglos
        pTipo = "ARR"
    Case btnSel(2).Checked 'Causas
        pTipo = "CAU"
    Case btnSel(3).Checked 'Usuarios
        pTipo = "USU"
    Case Else
        pTipo = ""
End Select


strSQL = "exec spCbr_Grupos_List_Asignacion '" & scTipo.Tag & "', '" & pTipo & "', '" & txtFiltro.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!IdX)
      itmX.SubItems(1) = rs!itmX
      itmX.SubItems(2) = rs!REGISTRO_USUARIO & ""
      itmX.SubItems(3) = rs!REGISTRO_FECHA & ""
      itmX.Checked = IIf((rs!Asignado = 1), True, False)
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

Private Sub btnSel_Click(Index As Integer)
Dim i As Integer

If vPaso Then Exit Sub

For i = 0 To btnSel.Count - 1
    btnSel.Item(i).Checked = False
Next i

btnSel.Item(Index).Checked = True


Call sbConsulta_Asignacion
End Sub

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub sbConsulta()

vPaso = True

strSQL = "select ID_GRUPO, DESCRIPCION, convert(int,ACTIVO) as 'ACTIVO', '...'" _
      & " from CBR_GRUPOS" _
      & " order by ID_GRUPO"
Call sbCargaGrid(vGrid, 4, strSQL)

vPaso = False

End Sub

Private Sub Form_Load()

vModulo = 4

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 2100
    .Add , , "Descripción", 4000
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Fecha", 2100, vbCenter
End With

Call sbConsulta

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

If Trim(vGrid.Text) = "" Then
  
  strSQL = "insert into CBR_GRUPOS(DESCRIPCION, ACTIVO, REGISTRO_USUARIO, REGISTRO_FECHA) values('"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ", '" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  strSQL = "select isnull( MAX(ID_GRUPO) ,0) as 'ID_GRUPO' from CBR_GRUPOS "
  Call OpenRecordSet(rs, strSQL)
  vGrid.Col = 1
  vGrid.Text = CStr(rs!ID_GRUPO)
  
  Call Bitacora("Registra", "Grupo de Cobros: " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CBR_GRUPOS set descripcion = '" & vGrid.Text & "', ACTIVO = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ", MODIFICA_FECHA = dbo.mygetdate(), MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
         & " where ID_GRUPO = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Grupo de Cobros: " & vGrid.Text)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim pTipo As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case True
    Case btnSel(0).Checked  'Gestiones
        pTipo = "GES"
    Case btnSel(1).Checked 'Arreglos
        pTipo = "ARR"
    Case btnSel(2).Checked 'Causas
        pTipo = "CAU"
    Case btnSel(3).Checked 'Usuarios
        pTipo = "USU"
    Case Else
        pTipo = ""
End Select


strSQL = "exec spCbr_Grupos_List_Asignacion_Add '" & scTipo.Tag & "', '" & pTipo _
       & "', '" & Item.Text & "', '" & glogon.Usuario & "', '" & IIf((Item.Checked), "A", "E") & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub



Private Sub txtFiltro_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbConsulta_Asignacion
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

vGrid.Row = Row
vGrid.Col = 1

scTipo.Tag = vGrid.Text
vGrid.Col = 2
scTipo.Caption = vGrid.Text

Call btnSel_Click(0)

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        strSQL = "exec spCbr_Grupos_Elimina " & vGrid.Text & ", '" & glogon.Usuario & "'"
        Call OpenRecordSet(rs, strSQL)
        
        If rs!Pass = 1 Then
                    
            vGrid.Col = 1
            strSQL = vGrid.Text
    
            vGrid.DeleteRows vGrid.ActiveRow, 1
            vGrid.MaxRows = vGrid.MaxRows - 1
            
            If vGrid.MaxRows <= 0 Then
              vGrid.MaxRows = 1
            End If
            
            Call Bitacora("Elimina", "Grupo de Cobros: " & strSQL)
            
            MsgBox "Grupo de Cobros: " & strSQL & ", Eliminado Satisfactoriamente!", vbInformation
        Else
            MsgBox rs!Mensaje, vbExclamation
        End If
        
        Call sbConsulta
        
     End If
End If


End Sub

