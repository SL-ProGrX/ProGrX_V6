VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSIF_OficinasMetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Oficinas : Metas"
   ClientHeight    =   6780
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   8196
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8196
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   264
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1284
      _ExtentX        =   2265
      _ExtentY        =   466
      ButtonWidth     =   487
      ButtonHeight    =   466
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Crear"
            Object.ToolTipText     =   "Crea un nuevo periodo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Guardar"
            Object.ToolTipText     =   "Guarda Cambios"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Resultados"
            Object.ToolTipText     =   "Actualiza Resultados"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7080
      Top             =   1320
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_OficinasMetas.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_OficinasMetas.frx":00FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_OficinasMetas.frx":021A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboPeriodo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.ComboBox cboOficina 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1200
      Width           =   6615
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4452
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   7812
      _Version        =   524288
      _ExtentX        =   13780
      _ExtentY        =   7853
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
      MaxCols         =   5
      MaxRows         =   13
      RowHeaderDisplay=   0
      ScrollBars      =   2
      SpreadDesigner  =   "frmSIF_OficinasMetas.frx":0344
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Oficina"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficinas : Definición y Monitoreo de Metas "
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
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSIF_OficinasMetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbCargaAnioMeta(pPeriodo As String)
Dim vOficina As String, i As Integer, Row As Integer
Dim vAnio As Long, curMetaAcum As Currency, curMetaAcumAnt As Currency
Dim strSQL As String, rs As New ADODB.Recordset

vOficina = SIFGlobal.fxCodText(cboOficina.Text, "-")

vAnio = Mid(pPeriodo, 1, 4)

With vGrid
    .Sheet = 1
    .MaxCols = 5
    
    curMetaAcum = 0
    curMetaAcumAnt = 0
    Row = 0
    
    strSQL = "exec spSIFOficinaMetasPeriodo '" & vOficina & "'," & vAnio & "," & vAnio + 1 & ",'" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Row = Row + 1
     .Row = Row
     For i = 1 To .MaxCols
       .Col = i
       Select Case i
         Case 1
             .Text = rs!anio
         Case 2
             .Text = rs!mes
         Case 3
             .Text = CStr(rs!Mes_Meta_Anterior)
           
             curMetaAcumAnt = curMetaAcumAnt + rs!Mes_Meta_Anterior
         Case 4
             .Text = CStr(rs!mes_meta)
             curMetaAcum = curMetaAcum + rs!mes_meta
         Case 5
             .Text = CStr(curMetaAcum)
       End Select
     Next i
     rs.MoveNext
    Loop
    rs.Close
    
    .Row = .MaxRows
    .Col = 3
    .Text = curMetaAcumAnt
    .Col = 4
    .Text = curMetaAcum

End With

End Sub

Private Sub cboOficina_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

If cboOficina.Text = "" Then Exit Sub

vPaso = True
cboPeriodo.Clear
cboPeriodo.AddItem "<Nuevo>"

strSQL = "select * from sif_oficina_metas_periodos where cod_oficina = '" _
       & SIFGlobal.fxCodText(cboOficina.Text) & "' order by anio_Corte desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboPeriodo.AddItem rs!Anio_Inicio & "-" & rs!Anio_Corte
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
 rs.MoveFirst
 cboPeriodo.Text = rs!Anio_Inicio & "-" & rs!Anio_Corte
Else
 cboPeriodo.Text = "<Nuevo>"
End If
rs.Close

vPaso = False
Call cboPeriodo_Click



End Sub


Private Sub cboPeriodo_Click()
Dim Row As Integer, Col As Integer

If vPaso Then Exit Sub

For Row = 1 To vGrid.MaxRows
 vGrid.Row = Row
 For Col = 1 To vGrid.MaxCols
   vGrid.Col = Col
   vGrid.Text = ""
 Next Col
Next Row

If cboPeriodo.Text = "" Or cboPeriodo.Text = "<Nuevo>" Then Exit Sub
Call sbCargaAnioMeta(cboPeriodo.Text)

End Sub

Private Sub sbGuardar()
Dim strSQL As String, vOficina As String
Dim Row As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

vOficina = SIFGlobal.fxCodText(cboOficina.Text)

With vGrid
  For Row = 1 To 12
    .Row = Row
    .Col = 4
    strSQL = "update sif_oficina_metas set mes_meta = " & CCur(.Text) & ",acumulado_meta = "
    .Col = 5
    strSQL = strSQL & CCur(.Text) & ",Actualizado_Fecha = dbo.MyGetdate(), Actualizado_Usuario = '" _
            & glogon.Usuario & "' where cod_oficina = '" & vOficina & "' and Anio = "
    .Col = 1
    strSQL = strSQL & .Text & " and Mes = "
    .Col = 2
    strSQL = strSQL & .Text
    
    Call ConectionExecute(strSQL)
  Next Row
End With

Me.MousePointer = vbDefault
MsgBox "Información de Metas del Periodo : " & cboPeriodo.Text & " Actualizadas Satisfactoriamente..!", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCalculos()
Dim Row As Integer, curAcumulado As Currency


With vGrid
 .Sheet = 1
 curAcumulado = 0

For Row = 1 To .MaxRows - 1
 .Row = Row
 .Col = 4
 curAcumulado = curAcumulado + CCur(.Text)
 .Col = 5
 .Text = curAcumulado
Next Row

.Row = .MaxRows
.Col = 4
.Text = curAcumulado

End With

End Sub

Private Sub Form_Load()
Dim strSQL As String

vGrid.AppearanceStyle = fxGridStyle
imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


vPaso = True
strSQL = "select rtrim(cod_oficina) + ' - ' + descripcion as ItmX from sif_oficinas where estado = 1" _
       & " order by cod_oficina"
Call sbLlenaCbo(cboOficina, strSQL, False, False)
vPaso = False
Call cboOficina_Click


End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vPeriodo As String, vAnio As Long

On Error GoTo vError

vAnio = Year(fxFechaServidor)
vPeriodo = ""

Select Case Button.Key
  Case "Crear"
    
    If cboPeriodo.Text <> "<Nuevo>" Or vGrid.ActiveSheet = 2 Then
      MsgBox "Debe de Indicar la Opción de Nuevo en Periodos o Seleccione la Hoja de Metas!", vbExclamation
      Exit Sub
    End If
  
    vPeriodo = InputBox("Especifique el año de corte del periodo fiscal a crear " & vbCrLf _
          & "Ejemplo de especificación del periodo es " & vAnio - 1 & "-" & vAnio _
          , "Metas", vAnio - 1 & "-" & vAnio)
    If Len(vPeriodo) <> 9 Then
      MsgBox "Periodo no es válido, verifique...!", vbExclamation
      Exit Sub
    End If
    
    vPaso = True
     cboPeriodo.AddItem vPeriodo
     cboPeriodo.Text = vPeriodo
    vPaso = False
    
    Call sbCargaAnioMeta(vPeriodo)

  Case "Guardar"
    If vGrid.ActiveSheet = 1 Then
      Call sbGuardar
    End If
  Case "Resultados"


End Select


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vGrid_KeyUp(KeyCode As Integer, Shift As Integer)
If vGrid.ActiveSheet = 1 And vGrid.ActiveCol = 4 Then
    Call sbCalculos
End If
End Sub
