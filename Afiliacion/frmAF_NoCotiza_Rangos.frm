VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmAF_NoCotiza_Rangos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rangos de No Cotizantes"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   12255
      _Version        =   524288
      _ExtentX        =   21616
      _ExtentY        =   11880
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
      SpreadDesigner  =   "frmAF_NoCotiza_Rangos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rangos de Tiempos para No Cotizantes"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   7812
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmAF_NoCotiza_Rangos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub sbInicial()

vPaso = True
      
strSQL = "select Linea_Id, Dia_Desde, Dia_Hasta, Descripcion, Activo, Registro_Fecha, Registro_Usuario" _
       & " from AFI_SOCIOS_SIN_APORTES_RANGOS Order by Dia_Desde asc"
Call sbCargaGrid(vGrid, 7, strSQL)

vPaso = False

End Sub


Private Sub Form_Load()

vModulo = 1

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 4
If Trim(vGrid.Text) = "" Then
    MsgBox "La descripción No es válida!", vbExclamation
    Exit Function
End If

vGrid.Col = 1
If vGrid.Text = "" Then 'Insertar
  vGrid.Col = 2
  strSQL = "insert into AFI_SOCIOS_SIN_APORTES_RANGOS(Dia_Desde, Dia_Hasta, Descripcion, Activo, Registro_Fecha, Registro_Usuario)" _
         & " values( " & CCur(vGrid.Text) & ", "
  vGrid.Col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ", '"
  vGrid.Col = 4
  strSQL = strSQL & vGrid.Text & "', "
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ", Getdate(), '" & glogon.Usuario & "')"

  Call ConectionExecute(strSQL)

  
  vGrid.Col = 4
  strSQL = "select Linea_Id, registro_Fecha, Registro_Usuario from AFI_SOCIOS_SIN_APORTES_RANGOS " _
         & " where Descripcion = '" & vGrid.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  
  
  vGrid.Col = 5
  vGrid.Text = rs!Registro_Fecha & ""
  vGrid.Col = 6
  vGrid.Text = rs!Registro_Usuario & ""
  
  
  vGrid.Col = 1
  vGrid.Text = CStr(rs!Linea_Id)
  
  Call Bitacora("Registra", "Rangos No Cotizantes Id:  " & vGrid.Text)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update AFI_SOCIOS_SIN_APORTES_RANGOS set Dia_Desde = " & CCur(vGrid.Text)
 vGrid.Col = 3
 strSQL = strSQL & ", Dia_Hasta = " & CCur(vGrid.Text)
 vGrid.Col = 4
 strSQL = strSQL & ", Descripcion = '" & vGrid.Text & "', Activo = "
 vGrid.Col = 5
 strSQL = strSQL & vGrid.Value & ", Modifica_Fecha = getdate(), Modifica_Usuario = '" _
        & glogon.Usuario & "' where Linea_Id = "
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text
 
 Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "Rangos No Cotizantes Id:  " & vGrid.Text)

End If


fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If (vGrid.ActiveCol = vGrid.MaxCols Or vGrid.ActiveCol = 5) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
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
        strSQL = "delete AFI_SOCIOS_SIN_APORTES_RANGOS where Linea_Id = " & vGrid.Text
        Call ConectionExecute(strSQL)

        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Rangos No Cotizantes Id:  " & vGrid.Text)

        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow

     End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




