VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_PoliticaPago 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Politica de Pago (Programación Fecha de Pago)"
   ClientHeight    =   5844
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5844
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4452
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   7452
      _Version        =   1245185
      _ExtentX        =   13144
      _ExtentY        =   7853
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Día de Pago"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Días no hábiles [Traslados]"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "tlb"
      Item(1).Control(1)=   "Label2(2)"
      Item(1).Control(2)=   "Label2(1)"
      Item(1).Control(3)=   "lblX"
      Item(1).Control(4)=   "Label2(0)"
      Item(1).Control(5)=   "dtpInicio"
      Item(1).Control(6)=   "dtpCorte"
      Item(1).Control(7)=   "lswX"
      Item(1).Control(8)=   "cboTipo"
      Item(1).Control(9)=   "cboDiaSemana"
      Begin XtremeSuiteControls.ListView lswX 
         Height          =   2292
         Left            =   -68080
         TabIndex        =   10
         Top             =   1800
         Visible         =   0   'False
         Width           =   5364
         _Version        =   1245185
         _ExtentX        =   9462
         _ExtentY        =   4043
         _StockProps     =   77
         BackColor       =   -2147483643
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
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3612
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   6612
         _Version        =   524288
         _ExtentX        =   11663
         _ExtentY        =   6371
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
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_PoliticaPago.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   -65080
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1452
         _ExtentX        =   2561
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Nuevo"
               Object.ToolTipText     =   "Nuevo / Actualiza"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Eliminar"
               Object.ToolTipText     =   "Elimina Política"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Actualiza"
               Object.ToolTipText     =   "Actualiza Tablas de Pago"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -68080
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -66760
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   -68080
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1245185
         _ExtentX        =   4678
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboDiaSemana 
         Height          =   312
         Left            =   -68080
         TabIndex        =   12
         Top             =   1080
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1245185
         _ExtentX        =   4678
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   -69760
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label lblX 
         Caption         =   "Etiqueta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   -69760
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione la política que desea Eliminar en el listado de disponibles, y luego presione el botón de eliminar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1932
         Index           =   1
         Left            =   -69760
         TabIndex        =   5
         Top             =   1800
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Listado de Políticas disponibles : "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   -65680
         TabIndex        =   4
         Top             =   1560
         Visible         =   0   'False
         Width           =   2892
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   120
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PoliticaPago.frx":06F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PoliticaPago.frx":0802
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PoliticaPago.frx":0939
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_PoliticaPago.frx":0A66
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Política de Pagos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   6972
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_PoliticaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cboTipo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


If vPaso Then Exit Sub

cboDiaSemana.Visible = False

dtpInicio.top = cboDiaSemana.top
dtpCorte.top = cboDiaSemana.top

dtpInicio.Visible = False
dtpCorte.Visible = False

lswX.ListItems.Clear
lswX.ColumnHeaders.Clear
lswX.ColumnHeaders.Add , , "Tipo de Política", 2300

Select Case cboTipo.Text
 Case "Día de la Semana"
    cboDiaSemana.Visible = True
    lblX.Caption = "Día de la Semana"
    lswX.ColumnHeaders.Add , , "Día de la Semana", 1800
    
    strSQL = "select id_Seq,case dia_semana when 2 then 'Lunes' when 3 then 'Martes' when 4 then 'Miércoles' when 5 then 'Jueves' " _
           & " when 6 then 'Viernes' when 7 then 'Sábado' when 1 then 'Domingo' end as Dia" _
           & " from CRD_POLITICA_PAGO_TRASLADOS Where Tipo = 'DS'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswX.ListItems.Add(, , cboTipo.Text)
          itmX.Tag = rs!Id_seq
          itmX.SubItems(1) = rs!Dia
      rs.MoveNext
    Loop
    rs.Close
    
 Case "Día Recurrente"
    lblX.Caption = "Indique el Día no Hábil"
    dtpInicio.Visible = True
 
    lswX.ColumnHeaders.Add , , "Día Recurrente", 1800
 
    strSQL = "select id_Seq,Fecha_Inicio, case month(fecha_inicio) when 1 then 'Enero' when 2 then 'Febrero'" _
           & " when 3 then 'Marzo' when 4 then 'Abril' When 5 then 'Mayo' when 6 then 'Junio' when 7 then 'Julio'" _
           & " when 8 then 'Agosto' when 9 then 'Septiembre' When 10 then 'Octubre' when 11 then 'Noviembre' when 12 then 'Diciembre' end as MesDesc" _
           & " from CRD_POLITICA_PAGO_TRASLADOS Where Tipo = 'DR'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswX.ListItems.Add(, , cboTipo.Text)
          itmX.Tag = rs!Id_seq
          itmX.SubItems(1) = Day(rs!FECHA_INICIO) & " de " & rs!MesDesc
      rs.MoveNext
    Loop
    rs.Close
 
 
 Case "Fecha Específica"
    lblX.Caption = "Indique Rango de Fechas"
    dtpInicio.Visible = True
    dtpCorte.Visible = True
 
    lswX.ColumnHeaders.Add , , "Día Inicio", 1300
    lswX.ColumnHeaders.Add , , "Día Corte", 1300
 
    strSQL = "select id_Seq,Fecha_Inicio, Fecha_Corte" _
           & " from CRD_POLITICA_PAGO_TRASLADOS Where Tipo = 'FE'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lswX.ListItems.Add(, , cboTipo.Text)
          itmX.Tag = rs!Id_seq
          itmX.SubItems(1) = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
          itmX.SubItems(2) = Format(rs!fecha_corte, "dd/mm/yyyy")
      rs.MoveNext
    Loop
    rs.Close
 
End Select


End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

vGrid.AppearanceStyle = fxGridStyle


tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)

strSQL = "select id_politica,dia_inicio,dia_corte,case Politica " _
       & " when 'FOR' then 'Día de la Formalización' when 'ULT' then 'Ultimo día del Mes'" _
       & " when 'ESP' then 'Día Específico' end as PoliticaDesc,Dia_Base" _
       & " from CRD_POLITICA_PAGO" _
       & " order by dia_inicio"
Call sbCargaGrid(vGrid, 5, strSQL)

With cboDiaSemana
  .Clear
  .AddItem "Domingo"
  .ItemData(.ListCount - 1) = CStr(1)
  .AddItem "Lunes"
  .ItemData(.ListCount - 1) = CStr(2)
  .AddItem "Martes"
  .ItemData(.ListCount - 1) = CStr(3)
  .AddItem "Miércoles"
  .ItemData(.ListCount - 1) = CStr(4)
  .AddItem "Jueves"
  .ItemData(.ListCount - 1) = CStr(5)
  .AddItem "Viernes"
  .ItemData(.ListCount - 1) = CStr(6)
  .AddItem "Sábado"
  .ItemData(.ListCount - 1) = CStr(7)
  .Text = "Domingo"
End With

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

vPaso = True
 cboTipo.Clear
 cboTipo.AddItem "Día de la Semana"
 cboTipo.AddItem "Día Recurrente"
 cboTipo.AddItem "Fecha Específica"
 cboTipo.Text = "Día de la Semana"
vPaso = False

Call cboTipo_Click

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If vGrid.Text = "" Then 'Insertar
  
  strSQL = "select isnull(max(id_politica),0) + 1 as Politica from crd_politica_pago"
  Call OpenRecordSet(rs, strSQL)
      vGrid.Text = CStr(rs!Politica)
  rs.Close
  
  strSQL = "insert into CRD_POLITICA_PAGO(id_politica,dia_inicio,dia_corte,politica,dia_base) values(" _
         & vGrid.Text & ","
         
         
  vGrid.col = 2
  strSQL = strSQL & IIf(vGrid.Text = "", 1, vGrid.Text) & ","
  vGrid.col = 3
  strSQL = strSQL & IIf(vGrid.Text = "", 1, vGrid.Text) & ",'"
  vGrid.col = 4
  Select Case Trim(vGrid.Text)
    Case "Día de la Formalización"
        strSQL = strSQL & "FOR',1)"
    Case "Ultimo día del Mes"
        strSQL = strSQL & "ULT',32)"
    Case "Día Específico"
        strSQL = strSQL & "ESP',"
        vGrid.col = 5
        strSQL = strSQL & IIf(vGrid.Text = "", 1, vGrid.Text) & ")"
    Case Else
        strSQL = strSQL & "ULT',32)"
  End Select
 
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Política de Pago : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update CRD_POLITICA_PAGO set dia_inicio = " & IIf(vGrid.Text = "", 1, vGrid.Text) & ",dia_corte = "
 vGrid.col = 3
 strSQL = strSQL & IIf(vGrid.Text = "", 1, vGrid.Text) & ",Politica = '"
 vGrid.col = 4
  Select Case Trim(vGrid.Text)
    Case "Día de la Formalización"
        strSQL = strSQL & "FOR', dia_base = 1"
    Case "Ultimo día del Mes"
        strSQL = strSQL & "ULT', dia_base = 32"
    Case "Día Específico"
        strSQL = strSQL & "ESP', dia_base = "
        vGrid.col = 5
        strSQL = strSQL & IIf(vGrid.Text = "", 1, vGrid.Text)
    Case Else
        strSQL = strSQL & "ULT', dia_base = 32"
  End Select
  
 vGrid.col = 1
 strSQL = strSQL & " where id_politica = " & vGrid.Text
 
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Política de Pago : " & vGrid.Text)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError

Select Case Button.Key
  Case "Nuevo"
        
        strSQL = " select isnull(max(id_seq),0) + 1 as SeqX from  CRD_POLITICA_PAGO_TRASLADOS"
        Call OpenRecordSet(rs, strSQL)
          i = rs!SeqX
        rs.Close
        
        Select Case cboTipo.Text
         Case "Día de la Semana"
            
            strSQL = "insert CRD_POLITICA_PAGO_TRASLADOS (id_Seq, tipo,dia_semana) " _
                   & " values(" & i & ", 'DS'," & cboDiaSemana.ItemData(cboDiaSemana.ListIndex) & ")"
            
         Case "Día Recurrente"
            strSQL = "insert CRD_POLITICA_PAGO_TRASLADOS (id_Seq, tipo,Fecha_Inicio) " _
                   & " values(" & i & ", 'DR','" & Format(dtpInicio.Value, "yyyy/mm/dd") & "')"

         
         
         Case "Fecha Específica"
            strSQL = "insert CRD_POLITICA_PAGO_TRASLADOS (id_Seq, tipo,Fecha_Inicio,Fecha_Corte) " _
                   & " values(" & i & ", 'FE','" & Format(dtpInicio.Value, "yyyy/mm/dd") & "','" _
                   & Format(dtpCorte.Value, "yyyy/mm/dd") & "')"

         
        End Select
        
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Registra", "Política de Pago [No Hábil]: " & i)
        
        
        MsgBox "Politica de días hábiles registrada satisfactoriamente...!", vbInformation
        
  Case "Eliminar"
    For i = 1 To lswX.ListItems.Count
      If lswX.ListItems.Item(i).Checked Then
         strSQL = "delete CRD_POLITICA_PAGO_TRASLADOS where id_Seq = " & lswX.ListItems.Item(i).Tag
         Call ConectionExecute(strSQL)
      
         Call Bitacora("Elimina", "Política de Pago [No Hábil]: " & lswX.ListItems.Item(i).Tag)
      End If
    Next
  
    MsgBox "Politica de días hábiles Eliminada satisfactoriamente...!", vbInformation
  
  Case "Actualiza"
  
        strSQL = "exec spCrdPlanPagoDiaHabilActualiza"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Aplica", "Política de Pago [No Hábil]: Actualiza Tablas")
        MsgBox "Tablas de Pago, Actualizadas [Fecha Pago]satisfactoriamente...!", vbInformation
  
  
End Select

Call cboTipo_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

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

'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete CRD_POLITICA_PAGO where Id_Politica = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Política de Pago : " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select id_politica,dia_inicio,dia_corte,case Politica " _
               & " when 'FOR' then 'Día de la Formalización' when 'ULT' then 'Ultimo día del Mes'" _
               & " when 'ESP' then 'Día Específico' end as PoliticaDesc,Dia_Base" _
               & " from CRD_POLITICA_PAGO" _
               & " order by dia_inicio"
        Call sbCargaGrid(vGrid, 5, strSQL)
     End If
End If


End Sub

