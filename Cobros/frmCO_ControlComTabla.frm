VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCO_ControlComTabla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla de Comisiones de Recuperación"
   ClientHeight    =   6396
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   9156
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6396
   ScaleWidth      =   9156
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8652
      _Version        =   524288
      _ExtentX        =   15261
      _ExtentY        =   8911
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
      MaxCols         =   492
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_ControlComTabla.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comisiones por Antiguedad de Cuota Recuperada"
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
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCO_ControlComTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Activate()
vModulo = 4
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 4
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select Id_Linea, Inicio, Corte, Porcentaje, Registro_Fecha, Registro_Usuario" _
       & " from Cbr_Comisiones_Tabla"
Call sbCargaGrid(vGrid, 6, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

'Guarda la información de la Id_Linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1


If Trim(vGrid.Text) = "" Then  'Insertar
  
  strSQL = "insert into Cbr_Comisiones_Tabla(inicio,corte,porcentaje,Registro_Fecha, Registro_Usuario) values("
  vGrid.col = 2
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.col = 3
  strSQL = strSQL & CLng(vGrid.Text) & ","
  vGrid.col = 4
  strSQL = strSQL & CDbl(vGrid.Text) & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
  
  Call ConectionExecute(strSQL)
    
    vGrid.col = 1
    strSQL = "select isnull(max(Id_Linea),0)  as Id_Linea from Cbr_Comisiones_Tabla"
    Call OpenRecordSet(rs, strSQL)
        vGrid.Text = rs!Id_Linea
    rs.Close
    
  Call Bitacora("Registra", "Tabla Comisión..Línea: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update Cbr_Comisiones_Tabla set inicio = " & CLng(vGrid.Text) & ",Corte = "
 vGrid.col = 3
 strSQL = strSQL & CLng(vGrid.Text) & ",Porcentaje = "
 vGrid.col = 4
 strSQL = strSQL & CDbl(vGrid.Text) & ", Registro_Fecha= dbo.MyGetdate(), Registro_Usuario = '" & glogon.Usuario _
        & "' where Id_Linea = "
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Tabla Comisión..Línea: " & vGrid.Text)

End If


vGrid.col = 5
vGrid.Text = Date
vGrid.col = 6
vGrid.Text = glogon.Usuario


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function





Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Id_Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar Id_Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete Cbr_Comisiones_Tabla where Id_Linea = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Tabla Comisión..Línea: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
