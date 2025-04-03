VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Begin VB.Form frmCntX_TiposAsientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipos de Asientos"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnImportar 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   6840
      Width           =   3375
      _Version        =   1310720
      _ExtentX        =   5953
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Importar Tipos de Asientos por Omisión"
      Appearance      =   6
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5295
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   8415
      _Version        =   524288
      _ExtentX        =   14843
      _ExtentY        =   9340
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_TiposAsientos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Asientos"
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
      Height          =   372
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   7452
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_TiposAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnImportar_Click()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spCntX_Tipos_Asientos_Default " & gCntX_Parametros.CodigoConta & ", '" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Importación de Tipos de Asientos Default, Conta Id: " & gCntX_Parametros.CodigoConta)

MsgBox "Importación realizada satisfactoriamente!", vbInformation

Call Form_Load

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 20

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select tipo_asiento,descripcion,activo,consecutivo from CntX_Tipos_Asientos where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " order by descripcion"
Call sbCargaGridFps7(vGrid, 4, strSQL, False)

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

rs.Open "select isnull(count(*),0) as Total from CntX_Tipos_Asientos where tipo_asiento = '" _
        & vGrid.Text & "' and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CntX_Tipos_Asientos(tipo_asiento,COD_CONTABILIDAD,descripcion,activo,consecutivo) values('"
  vGrid.Col = 1
  strSQL = strSQL & UCase(vGrid.Text) & "'," & gCntX_Parametros.CodigoConta & ",'"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 4
  strSQL = strSQL & IIf(Not IsNumeric(vGrid.Text), 0, vGrid.Text) & ")"
  
  Call ConectionExecute(strSQL, 0)

  vGrid.Col = 2
  
  Call Bitacora("Registra", "Tipo Asiento : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
  
  fxGuardar = 1

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CntX_Tipos_Asientos set descripcion = '" & vGrid.Text & "',Activo = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & ",consecutivo = "
 vGrid.Col = 4
 strSQL = strSQL & IIf(Not IsNumeric(vGrid.Text), 0, vGrid.Text) _
        & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta & " and tipo_asiento = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL, 0)
 
 fxGuardar = 1
End If

rs.Close

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
'  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Reporte
If KeyCode = vbKeyF5 Then
    Call sbCntX_Reportes_Catalogos("Tipos_Asientos")
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
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CntX_Tipos_Asientos where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
               & " and tipo_asiento = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL, 0)
        
        Call Bitacora("Elimina", "Tipo Asiento : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


