VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCO_CausasYArreglos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Causas de Morosidad y Tipos de arreglos de pago"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   7845
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5655
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   7335
      _Version        =   524288
      _ExtentX        =   12938
      _ExtentY        =   9975
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
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_CausasYArreglos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Causas de Morosidad"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6852
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_CausasYArreglos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mSheet As Integer

Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 4
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call vGrid_SheetChanged(2, 1)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


Select Case mSheet
  Case 1 'Causas
        vTipo = "Causa de Morosidad: "
        strSQL = "select isnull(count(*),0) as Existe from CBR_CAUSAS_MOROSIDAD " _
               & " where cod_causa = '" & vGrid.Text & "'"
  Case 2 'Arreglos
        vTipo = "Tipo de Arreglo: "
        strSQL = "select isnull(count(*),0) as Existe from CBR_TIPOS_ARREGLOS " _
               & " where cod_arreglo = '" & vGrid.Text & "'"
End Select

Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
    Select Case mSheet
      Case 1 'Causas
          strSQL = "insert CBR_CAUSAS_MOROSIDAD(cod_causa,descripcion,Activa,Registro_Usuario,Registro_Fecha) values('"
      Case 2 'Arreglos
          strSQL = "insert CBR_TIPOS_ARREGLOS(cod_arreglo,descripcion,Activo,Registro_Usuario,Registro_Fecha) values('"
    End Select
         
         
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", vTipo & vGrid.Text)

Else 'Actualizar
    
    vGrid.Col = 2
    Select Case mSheet
      Case 1 'Causas
            strSQL = "update CBR_CAUSAS_MOROSIDAD set descripcion = '" & vGrid.Text & "',Activa = "
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Value & " where cod_causa = '"
          
      Case 2 'Arreglos
            strSQL = "update CBR_TIPOS_ARREGLOS set descripcion = '" & vGrid.Text & "',Activo = "
            vGrid.Col = 3
            strSQL = strSQL & vGrid.Value & " where cod_arreglo = '"
    
    End Select
 
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
    
    vGrid.Col = 1
    Call Bitacora("Modifica", vTipo & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String, vTipo As String

On Error GoTo vError

vGrid.Sheet = mSheet

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

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        Select Case mSheet
          Case 1 'Causas
                vTipo = "Causa de Morosidad: "
                strSQL = "delete CBR_CAUSAS_MOROSIDAD where cod_causa = '" & vGrid.Text & "'"
          Case 2 'Arreglos
                vTipo = "Tipo de Arreglo: "
                strSQL = "delete CBR_TIPOS_ARREGLOS where cod_arreglo = '" & vGrid.Text & "'"
        End Select
        
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", vTipo & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String

vGrid.Sheet = NewSheet
lblTitulo.Caption = vGrid.SheetName

mSheet = NewSheet
 
Select Case NewSheet
   Case 1 'Causas
        strSQL = "select cod_Causa,descripcion,Activa from CBR_CAUSAS_MOROSIDAD" _
              & " order by cod_Causa"
   Case 2 'Arreglos
        strSQL = "select Cod_Arreglo,descripcion,Activo from CBR_TIPOS_ARREGLOS" _
              & " order by Cod_Arreglo"
End Select

vGrid.Sheet = mSheet
vGrid.ActiveSheet = mSheet

Call sbCargaGrid(vGrid, 3, strSQL)


End Sub
