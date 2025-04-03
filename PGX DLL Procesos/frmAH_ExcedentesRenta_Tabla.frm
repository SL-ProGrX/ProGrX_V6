VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAH_ExcedentesRenta_Tabla 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Excedentes: Tabla de Renta"
   ClientHeight    =   6780
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5172
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   6972
      _Version        =   524288
      _ExtentX        =   12298
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
      MaxCols         =   482
      ScrollBars      =   2
      SpreadDesigner  =   "frmAH_ExcedentesRenta_Tabla.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabla de Renta para Excedentes"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   5172
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "frmAH_ExcedentesRenta_Tabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 2
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 2

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select ID_RENTA,DESDE,HASTA,PORCENTAJE from EXC_RENTA_TABLA" _
       & " order by ID_RENTA"
Call sbCargaGrid(vGrid, 4, strSQL)

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, vTipo As String
Dim vRentaId As Long

On Error GoTo vError

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

fxGuardar = 0


If Trim(vGrid.Text) = "" Then 'Insertar
  
    strSQL = "select isnull(max(ID_RENTA),0)+ 1 as 'RentaId' from EXC_RENTA_TABLA"
    Call OpenRecordSet(rs, strSQL)
        vRentaId = rs!RentaId
    rs.Close


  strSQL = "insert into EXC_RENTA_TABLA(ID_RENTA,Desde, Hasta, Porcentaje,registro_fecha,registro_usuario) values(" & vRentaId
  vGrid.Col = 2
  strSQL = strSQL & "," & CCur(vGrid.Text)
  vGrid.Col = 3
  strSQL = strSQL & "," & CCur(vGrid.Text)
  vGrid.Col = 4
  strSQL = strSQL & "," & CCur(vGrid.Text) & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  vGrid.Text = CStr(vRentaId)
  Call Bitacora("Registra", "Excedentes: Tabla Renta Id: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update EXC_RENTA_TABLA set Desde = " & CCur(vGrid.Text) & ", Hasta = "
    
    vGrid.Col = 3
    strSQL = strSQL & CCur(vGrid.Text) & ", Porcentaje = "
    vGrid.Col = 4
    strSQL = strSQL & CCur(vGrid.Text)
    vGrid.Col = 1
    strSQL = strSQL & " where ID_RENTA = " & vGrid.Text
    
    Call ConectionExecute(strSQL)
 
    vGrid.Col = 1
   
    Call Bitacora("Modifica", "Excedentes: Tabla Renta Id: " & vGrid.Text)
 
End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long


On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
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
        
        strSQL = "delete EXC_RENTA_TABLA where ID_RENTA = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Excedentes: Tabla Renta Id: " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




