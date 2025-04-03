VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmMarcasConfig 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de Marcas"
   ClientHeight    =   7860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   10335
      _Version        =   524288
      _ExtentX        =   18230
      _ExtentY        =   11245
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
      MaxCols         =   7
      ScrollBars      =   2
      SpreadDesigner  =   "frmMarcasConfig.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   612
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   6612
      _Version        =   1441793
      _ExtentX        =   11663
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Configuración de Horarios"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1116
      Left            =   0
      Top             =   0
      Width           =   10920
   End
End
Attribute VB_Name = "frmMarcasConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Activate()
vModulo = 21
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 21

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vPaso = True
    strSQL = "Select COD_HORARIO, DESCRIPCION, ENTRADA, SALALMUERZO, ENTALMUERZO, SALIDA, ESTADO" _
           & " From MARCAS_HORARIOS"
    Call sbCargaGrid(vGrid, 7, strSQL)
vPaso = False

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

strSQL = "select isnull(count(*),0) as Existe from MARCAS_HORARIOS " _
       & " where cod_horario = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)


If rs!Existe = 0 Then 'Insertar
      If Trim(vGrid.Text) = "" Then Exit Function
      
      vGrid.Col = 1
      strSQL = "insert into marcas_horarios(cod_horario,descripcion,entrada,salalmuerzo,entalmuerzo,salida" _
             & ",estado,registro_fecha, registro_usuario)" _
             & " values('" & vGrid.Text & "','"
      vGrid.Col = 2
      strSQL = strSQL & vGrid.Text & "',"
      vGrid.Col = 3
      strSQL = strSQL & vGrid.Value & ","
      vGrid.Col = 4
      strSQL = strSQL & vGrid.Value & ","
      vGrid.Col = 5
      strSQL = strSQL & vGrid.Value & ","
      vGrid.Col = 6
      strSQL = strSQL & vGrid.Value & ","
      vGrid.Col = 7
      strSQL = strSQL & vGrid.Value & ", dbo.MyGetdate(),'" & glogon.Usuario & "')"
      Call ConectionExecute(strSQL)
    
      vGrid.Col = 1
      Call Bitacora("Registra", "Horario Código: " & vGrid.Text)

Else 'Actualizar

      strSQL = "update marcas_horarios set descripcion = '"
              vGrid.Col = 2
              strSQL = strSQL & UCase(vGrid.Text) & "',entrada = "
              vGrid.Col = 3
              strSQL = strSQL & vGrid.Value & ",salalmuerzo = "
              vGrid.Col = 4
              strSQL = strSQL & vGrid.Value & ",entalmuerzo = "
              vGrid.Col = 5
              strSQL = strSQL & vGrid.Value & ", Salida = "
              vGrid.Col = 6
              strSQL = strSQL & vGrid.Value & ", Estado = "
              vGrid.Col = 7
              strSQL = strSQL & vGrid.Value & " where cod_horario = '"
              vGrid.Col = 1
              strSQL = strSQL & UCase(vGrid.Text) & "'"
              Call ConectionExecute(strSQL)
      
     vGrid.Col = 1
     Call Bitacora("Modifica", "Horario Código: " & vGrid.Text)

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
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'elimina un registro
If KeyCode = vbKeyDelete Then
  i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
       vGrid.Row = vGrid.ActiveRow
       vGrid.Col = 1
       strSQL = "delete MARCAS_HORARIOS where cod_horario = '" & vGrid.Text & "'"
       Call ConectionExecute(strSQL)
        
       If Not glogon.error Then
          Call Bitacora("Elimina", "Horario Código: " & vGrid.Text)
       End If
     End If
End If

End Sub
