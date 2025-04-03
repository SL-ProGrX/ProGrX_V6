VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_CausasRenuncias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Causas de Renuncia"
   ClientHeight    =   6900
   ClientLeft      =   1875
   ClientTop       =   600
   ClientWidth     =   16440
   HelpContextID   =   1004
   Icon            =   "frmAF_CausasRenuncia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   16440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   16215
      _Version        =   524288
      _ExtentX        =   28601
      _ExtentY        =   9551
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
      MaxCols         =   501
      ScrollBars      =   2
      SpreadDesigner  =   "frmAF_CausasRenuncia.frx":000C
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Causas de Renuncias"
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
      Height          =   495
      Index           =   0
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   -120
      Top             =   0
      Width           =   16695
   End
End
Attribute VB_Name = "frmAF_CausasRenuncias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

vModulo = 1

End Sub

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
If vGrid.Text = "" Then
    vGrid.Col = 2
    strSQL = "insert causas_renuncias(descripcion, Tipo_Apl, mortalidad, AJUSTE_TASAS, liq_alterna, tasa_planilla, tasa_ventanilla, institucion, COD_PLAN, Activo)" _
           & " values('" & vGrid.Text & "', '"
    vGrid.Col = 3
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', "
           
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Value & ","
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Value & ","
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Value & ","
    vGrid.Col = 7
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 8
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.Col = 9
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & ", '"
    vGrid.Col = 10
    strSQL = strSQL & vGrid.Text & "', "
    vGrid.Col = 11
    strSQL = strSQL & vGrid.Value & ")"
    
    Call ConectionExecute(strSQL)
     
  
  
    strSQL = "select max(id_causa) as ultimo from causas_renuncias"
    Call OpenRecordSet(rs, strSQL)
      vGrid.Col = 1
      vGrid.Text = CStr(rs!ultimo)
    rs.Close
    
    Call Bitacora("Registra", "Causa de Renuncia : id( " & vGrid.Text & ")")

   Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update causas_renuncias set descripcion = '" & vGrid.Text & "', Tipo_Apl = '"
    vGrid.Col = 3
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', mortalidad = "
     
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Value & ", AJUSTE_TASAS = "
    
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Value & ", liq_alterna = "
    
    
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Value & ", tasa_planilla = "
    vGrid.Col = 7
    strSQL = strSQL & CCur(vGrid.Text) & ", tasa_ventanilla = "
    vGrid.Col = 8
    strSQL = strSQL & CCur(vGrid.Text) & ", institucion = "
    vGrid.Col = 9
    strSQL = strSQL & IIf((vGrid.Text = ""), 0, vGrid.Text) & ", Cod_Plan = '"
    vGrid.Col = 10
    strSQL = strSQL & vGrid.Text & "', Activo = "
    vGrid.Col = 11
    strSQL = strSQL & vGrid.Value & " where id_causa = "
    
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Causa de Renuncia : id( " & vGrid.Text & ")")
    
    
   End If

   vGrid.Col = 1
   fxGuardar = vGrid.Text
   
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

End Sub


Private Sub Form_Load()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Cargar en el Objeto adoData Todas las causas de renuncia que existen.
'REFERENCIAS:   sbToolBarIconos - (Carga los iconos para la barra de herramientas)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String

On Error GoTo vError

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

strSQL = "select Id_Causa, Descripcion, Tipo_Apl, mortalidad, AJUSTE_TASAS, liq_alterna" _
       & ", tasa_planilla, tasa_ventanilla, institucion, cod_Plan, activo" _
       & " from vAFI_Causas_Renuncias order by id_causa"
Call sbCargaGrid(vGrid, 11, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)


Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

