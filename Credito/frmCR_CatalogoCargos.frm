VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_CatalogoCargos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cargos Adicionales para la Gestión del Crédito"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12930
   Icon            =   "frmCR_CatalogoCargos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   12930
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   960
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7092
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12732
      _Version        =   1310723
      _ExtentX        =   22458
      _ExtentY        =   12509
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
      ItemCount       =   3
      Item(0).Caption =   "Cargos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "lswCargos"
      Item(1).Control(1)=   "ArbolExp"
      Item(1).Control(2)=   "lbl"
      Item(1).Control(3)=   "lblNodeLinea(2)"
      Item(1).Control(4)=   "lblNodeLinea(1)"
      Item(1).Control(5)=   "lblNodeLinea(0)"
      Item(1).Control(6)=   "cmdReporte"
      Item(2).Caption =   "Tabla de Aplicación"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "Label2"
      Item(2).Control(1)=   "cboCargo"
      Item(2).Control(2)=   "vgTabla"
      Begin XtremeSuiteControls.ListView lswCargos 
         Height          =   5532
         Left            =   -63160
         TabIndex        =   11
         Top             =   720
         Visible         =   0   'False
         Width           =   5772
         _Version        =   1310723
         _ExtentX        =   10181
         _ExtentY        =   9758
         _StockProps     =   77
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   7.5
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
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6972
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   12612
         _Version        =   524288
         _ExtentX        =   22246
         _ExtentY        =   12298
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
         MaxCols         =   497
         SpreadDesigner  =   "frmCR_CatalogoCargos.frx":000C
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   5520
         Left            =   -70000
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   6732
         _ExtentX        =   11880
         _ExtentY        =   9737
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   372
         Left            =   -58720
         TabIndex        =   10
         Top             =   6360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboCargo 
         Height          =   312
         Left            =   -67000
         TabIndex        =   13
         Top             =   1080
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1310723
         _ExtentX        =   8281
         _ExtentY        =   582
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
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vgTabla 
         Height          =   5172
         Left            =   -67480
         TabIndex        =   14
         Top             =   1680
         Visible         =   0   'False
         Width           =   10092
         _Version        =   524288
         _ExtentX        =   17801
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
         MaxCols         =   7
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_CatalogoCargos.frx":0C46
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   972
         Left            =   -69760
         TabIndex        =   12
         Top             =   840
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
         _ExtentY        =   1714
         _StockProps     =   79
         Caption         =   "Cargo con Detalle de Aplicación por Tabla de Montos y Plazos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNodeLinea 
         BackStyle       =   0  'Transparent
         Caption         =   "LINEA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   7
         ToolTipText     =   "Linea"
         Top             =   6360
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label lblNodeLinea 
         BackStyle       =   0  'Transparent
         Caption         =   "DESTINO"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   -69880
         TabIndex        =   6
         ToolTipText     =   "Linea"
         Top             =   6600
         Visible         =   0   'False
         Width           =   2052
      End
      Begin VB.Label lblNodeLinea 
         BackStyle       =   0  'Transparent
         Caption         =   "GARANTIA"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -67720
         TabIndex        =   5
         ToolTipText     =   "Linea"
         Top             =   6360
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   312
         Left            =   -70000
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   12612
      End
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   372
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton cmdModifica 
      Height          =   372
      Left            =   480
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargos adicionales para la gestion de créditos"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmCR_CatalogoCargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vNode As Node
Dim vPaso As Boolean

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, vResulta As String
Dim vCadena As String, x As Integer

lblNodeLinea.Item(0).Tag = ""
lblNodeLinea.Item(1).Tag = ""
lblNodeLinea.Item(2).Tag = ""

lbl.Caption = Node.FullPath
lbl.Tag = Node.Key

If Right(Node.Key, 1) = "G" Then
     
   vCadena = fxIndiceCodigo(Node.Key)
   lblNodeLinea.Item(2).Tag = Right(vCadena, 1)
   x = 0
   vResulta = ""
   For i = 1 To Len(vCadena)
     If Mid(vCadena, i, 1) = "-" Then
        lblNodeLinea.Item(x).Tag = vResulta
        If x = 1 Then
          'Carta la Ultima Letra para el caso de los destinos
          lblNodeLinea.Item(x).Tag = Mid(lblNodeLinea.Item(x).Tag, 1, Len(lblNodeLinea.Item(x).Tag) - 1)
        End If
        x = x + 1
        vResulta = ""
     Else
        vResulta = vResulta & Mid(vCadena, i, 1)
     End If
   
   Next i

    Call sbCargaLswAdicional
Else
    lswCargos.ListItems.Clear
End If

lblNodeLinea.Item(0).Caption = "Línea   : " & lblNodeLinea.Item(0).Tag
lblNodeLinea.Item(1).Caption = "Destino : " & lblNodeLinea.Item(1).Tag
lblNodeLinea.Item(2).Caption = "Garantia: " & lblNodeLinea.Item(2).Tag


End Sub

Private Sub cboCargo_Click()
If vPaso Then Exit Sub

Dim strSQL As String


strSQL = "select ID_TABLA,MONTO_INICIO, MONTO_CORTE, PLAZO_INICIO, PLAZO_CORTE" _
       & ", CASE WHEN APL_TIPO = 'P' then 'Porcentaje' else 'Monto' END as TIPO, APL_VALOR" _
       & " FROM CRD_CARGOS_ADICIONAL_TABLA" _
       & " WHERE COD_CARGO = '" & cboCargo.ItemData(cboCargo.ListIndex) _
       & "' ORDER BY MONTO_INICIO"

Call sbCargaGrid(vgTabla, 7, strSQL, True)

End Sub

Private Sub cmdActualiza_Click()
Dim strSQL As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
'Borra la Asignación Inicial y Despues los vuelve a ingresar
strSQL = "delete cargos_asignacion where codigo = '" & vCodigo & "'"
Call ConectionExecute(strSQL)

With lswCargos
  For i = 1 To .ListItems.Count
     If .ListItems(i).Checked Then
        strSQL = "insert cargos_asignacion(cod_cargo,codigo) values('" _
               & .ListItems(i).Text & "','" & vCodigo & "')"
        Call ConectionExecute(strSQL)
     End If
  Next
End With

Call Bitacora("Modifica", "Cargos Adicionales al Código: " & vCodigo)

Me.MousePointer = vbDefault
MsgBox "Actualización de Cargos Actualizada...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer


Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
      Case 5
          Select Case CStr(rs.Fields(i - 1).Value)
            Case "C"
                vGrid.Text = "Crédito"
            Case "A"
                vGrid.Text = "Avaluo"
            Case "P"
                vGrid.Text = "Prima"
          End Select
      
      Case 6
        
          Select Case CStr(rs.Fields(i - 1).Value)
            Case "P"
                vGrid.Text = "Porcentaje"
            Case "M"
                vGrid.Text = "Monto"
            Case "T"
                vGrid.Text = "Tabla"
          End Select
      
      Case 8, 15
        vGrid.Text = fxgCntCuentaFormato(True, CStr(rs.Fields(i - 1).Value & ""), 0)
      
      Case 9 'Tipo de Deduccion
        vGrid.Text = IIf((CStr(rs.Fields(i - 1).Value) = "P"), "Periodica", "Formalización")
      
      Case 10 'Tipo de Plazo
        Select Case Trim(rs.Fields(i - 1).Value)
          Case "PZ"
            vGrid.Text = "Sin Plazo"
          Case "PF"
            vGrid.Text = "Plazo Fijo"
          Case "PC"
            vGrid.Text = "Plazo Credito"
          Case "PD"
            vGrid.Text = "Días Formaliza"
        End Select
      
      Case Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End Select
 
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub cmdReporte_Click()
With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Crédito"

   .Connect = glogon.ConectRPT

   .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
   .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoCargos.rpt")
   
   
   .PrintReport
End With
End Sub

Private Sub Form_Load()

vModulo = 3

tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 

With lswCargos.ColumnHeaders
  .Clear
  .Add , , "Código", 1440
  .Add , , "Descripción", 3600
  .Add , , "Tipo", 1100, vbCenter
  .Add , , "Valor", 2400, vbRightJustify
End With
lswCargos.Checkboxes = True



Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdModifica.Enabled
lswCargos.Enabled = cmdActualiza.Enabled

End Sub

Private Function fxExiste(vCod As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from cargos_adicionales" _
       & " where cod_cargo = '" & vCod & "'"
Call OpenRecordSet(rs, strSQL)
fxExiste = IIf((rs!Existe = 1), True, False)
rs.Close
End Function

Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""
fxVerifica = True

vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If vGrid.Text = "" Then vMensaje = vMensaje & " - Especifique un código para este cargo" & vbCrLf

vGrid.col = 5
If vGrid.Text = "" Then vMensaje = vMensaje & " - Especifique la Base de Calculo" & vbCrLf

vGrid.col = 6
If vGrid.Text = "" Then vMensaje = vMensaje & " - Especifique el Tipo de Cargo" & vbCrLf

vGrid.col = 8
If Not fxCntX_CuentaValida(vGrid.Text) Then vMensaje = vMensaje & " - Especifique una cuenta contable válida!" & vbCrLf


vGrid.col = 9
If vGrid.Text = "" Then vMensaje = vMensaje & " - Especifique el Tipo de Deducción" & vbCrLf

vGrid.col = 10
If vGrid.Text = "" Then vMensaje = vMensaje & " - Especifique el Tipo de Plazo (Para Casos de Monto no Aplica)" & vbCrLf

vGrid.col = 11
If vGrid.Text = "" Then vGrid.Text = "0"


vGrid.col = 14
If vGrid.Value = vbChecked Then
    vGrid.col = 15
    If Not fxCntX_CuentaValida(vGrid.Text) Then vMensaje = vMensaje & " - La Cuenta Contable para Diferir no es válida!" & vbCrLf
End If

If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   fxVerifica = False
End If


End Function


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

If Not fxVerifica Then Exit Function

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If Not fxExiste(vGrid.Text) Then
   vGrid.col = 1
   strSQL = "insert cargos_adicionales(cod_cargo,descripcion,automatico,AUMENTA_BASE_CRD,base,tipo,valor,cod_cuenta,tipo_deduccion" _
          & ",plazo_tipo,plazo_dias,monto_inicio,monto_corte,diferido_cargo,diferido_cod_cuenta, iva_porcentaje, activo)" _
          & " values('" & vGrid.Text & "','"
   vGrid.col = 2
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.col = 3
   strSQL = strSQL & vGrid.Value & ","
   vGrid.col = 4
   strSQL = strSQL & vGrid.Value & ",'"
   
   vGrid.col = 5
   strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "','"
   vGrid.col = 6
   strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
   vGrid.col = 7
   strSQL = strSQL & CDbl(vGrid.Text) & ",'"
   vGrid.col = 8
   strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text, 0) & "','"
   vGrid.col = 9
   strSQL = strSQL & IIf((vGrid.Text = "Periodica"), "P", "F") & "',"
   vGrid.col = 10
   Select Case Trim(vGrid.Text)
        Case "Sin Plazo"
          strSQL = strSQL & "'PZ',"
        Case "Plazo Fijo"
          strSQL = strSQL & "'PF',"
        Case "Plazo Credito"
          strSQL = strSQL & "'PC',"
        Case "Días Formaliza"
          strSQL = strSQL & "'PD',"
   End Select
   vGrid.col = 11
   strSQL = strSQL & CInt(vGrid.Text) & ","
   vGrid.col = 12
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 13
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 14
   strSQL = strSQL & vGrid.Value & ",'"
   vGrid.col = 15
   strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',"
   
   vGrid.col = 16
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 17
   strSQL = strSQL & vGrid.Value & ")"
   
   Call ConectionExecute(strSQL)
   vGrid.col = 1
   Call Bitacora("Registra", "Cargo Adicional Cod: " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.col = 2
    strSQL = "update cargos_adicionales set descripcion = '" & vGrid.Text
    vGrid.col = 3
    strSQL = strSQL & "',Automatico = " & vGrid.Value
    vGrid.col = 4
    strSQL = strSQL & ",AUMENTA_BASE_CRD = " & vGrid.Value
    
    vGrid.col = 5
    strSQL = strSQL & ",base= '" & Mid(vGrid.Text, 1, 1)
    vGrid.col = 6
    strSQL = strSQL & "',tipo = '" & Mid(vGrid.Text, 1, 1)
    vGrid.col = 7
    strSQL = strSQL & "',valor = " & CDbl(vGrid.Text) & ",cod_cuenta = '"
    vGrid.col = 8
    strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',Tipo_Deduccion = '"
    vGrid.col = 9
    strSQL = strSQL & IIf((vGrid.Text = "Periodica"), "P", "F") & "',Plazo_Tipo = "
    vGrid.col = 10
    Select Case Trim(vGrid.Text)
        Case "Sin Plazo"
          strSQL = strSQL & "'PZ',Plazo_Dias = "
        Case "Plazo Fijo"
          strSQL = strSQL & "'PF',Plazo_Dias = "
        Case "Plazo Credito"
          strSQL = strSQL & "'PC',Plazo_Dias = "
        Case "Días Formaliza"
          strSQL = strSQL & "'PD',Plazo_Dias = "
    End Select
    vGrid.col = 11
    strSQL = strSQL & CInt(vGrid.Text) & ",monto_inicio = "
    vGrid.col = 12
    strSQL = strSQL & CCur(vGrid.Text) & ",monto_corte = "
    vGrid.col = 13
    strSQL = strSQL & CCur(vGrid.Text) & ",diferido_cargo = "
    vGrid.col = 14
    strSQL = strSQL & vGrid.Value & ",diferido_cod_cuenta = '"
    vGrid.col = 15
    strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text, 0) & "', iva_porcentaje = "
    
    vGrid.col = 16
    strSQL = strSQL & CCur(vGrid.Text) & ", Activo = "
    vGrid.col = 17
    strSQL = strSQL & vGrid.Value
   
    vGrid.col = 1
    strSQL = strSQL & " where cod_cargo = '" & vGrid.Text & "'"
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Cargo Adicional Cod: " & vGrid.Text)
    
End If

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   
End Function


Private Function fxGuardarTabla() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

'If Not fxVerifica Then Exit Function

With vgTabla

fxGuardarTabla = 0
.Row = .ActiveRow
.col = 1

If .Text = "" Then
   .col = 1
   
   strSQL = "select isnull(max(ID_TABLA),0) + 1 AS 'TABLA_ID'" _
       & " FROM CRD_CARGOS_ADICIONAL_TABLA" _
       & " WHERE COD_CARGO = '" & cboCargo.ItemData(cboCargo.ListIndex) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    
    .Text = CStr(rs!Tabla_Id)
   
   strSQL = "insert CRD_CARGOS_ADICIONAL_TABLA(cod_cargo,ID_TABLA,MONTO_INICIO, MONTO_CORTE, PLAZO_INICIO, PLAZO_CORTE" _
          & ",APL_TIPO, APL_VALOR, REGISTRO_FECHA, REGISTRO_USUARIO)" _
          & " values('" & cboCargo.ItemData(cboCargo.ListIndex) & "'," & .Text & ","
   .col = 2
   strSQL = strSQL & CCur(.Text) & ","
   .col = 3
   strSQL = strSQL & CCur(.Text) & ","
   .col = 4
   strSQL = strSQL & CLng(.Text) & ","
   .col = 5
   strSQL = strSQL & CLng(.Text) & ",'"
   .col = 6
   strSQL = strSQL & Mid(.Text, 1, 1) & "',"
   .col = 7
   
   strSQL = strSQL & CCur(.Text) & ", dbo.Mygetdate(), '" & glogon.Usuario & "') "
  
   
   Call ConectionExecute(strSQL)
   
   .col = 1
   Call Bitacora("Registra", "Cargo Adicional (Tabla) Cod: " & cboCargo.ItemData(cboCargo.ListIndex) & ", Id." & .Text)
   
 Else 'Actualizar
    .col = 2
    strSQL = "update CRD_CARGOS_ADICIONAL_TABLA set MONTO_INICIO = " & CCur(.Text)
    .col = 3
    strSQL = strSQL & ", MONTO_CORTE = " & CCur(.Text)
    .col = 4
    strSQL = strSQL & ", PLAZO_INICIO = " & CLng(.Text)
    .col = 5
    strSQL = strSQL & ", PLAZO_CORTE = " & CLng(.Text)
    .col = 6
    strSQL = strSQL & ", APL_TIPO = '" & Mid(.Text, 1, 1) & "'"
    .col = 7
    strSQL = strSQL & ", APL_VALOR = " & CCur(.Text)
    
   .col = 1
    strSQL = strSQL & " WHERE COD_CARGO = '" & cboCargo.ItemData(cboCargo.ListIndex) _
            & "' AND ID_TABLA = " & .Text

   
    Call ConectionExecute(strSQL)
    
   Call Bitacora("Actualiza", "Cargo Adicional (Tabla) Cod: " & cboCargo.ItemData(cboCargo.ListIndex) & ", Id." & .Text)
    
End If

End With


Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   
End Function


Private Sub sbCargaLswAdicional()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass


vPaso = True

strSQL = "select R.*,A.codigo as Existe" _
       & " from Cargos_Adicionales R left Join CRD_CARGOS_ASG_DETALLE A " _
       & " on R.cod_cargo = A.cod_cargo and A.codigo = '" & lblNodeLinea.Item(0).Tag _
       & "' and A.Cod_destino = '" & lblNodeLinea.Item(1).Tag & "' and A.Garantia = '" & lblNodeLinea.Item(2).Tag _
       & "' order by existe desc,R.cod_cargo"
Call OpenRecordSet(rs, strSQL, 0)

lswCargos.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswCargos.ListItems.Add(, , rs!COD_CARGO)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.SubItems(2) = IIf((rs!Tipo = "P"), "PORCENTUAL", "MONTO")
      itmX.SubItems(3) = Format(rs!Valor, "Standard")
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      
      If itmX.Checked Then itmX.ForeColor = vbBlue
      
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

tcMain.Width = Me.Width - 370
tcMain.Height = Me.Height - 700

vGrid.Width = tcMain.Width - 250
vGrid.Height = tcMain.Height - (vGrid.top + 1000)

lbl.Width = tcMain.Width - 250

ArbolExp.Height = tcMain.Height - (1600 + tcMain.top)

lswCargos.Height = ArbolExp.Height
lswCargos.Width = tcMain.Width - ArbolExp.Width - 480

cmdReporte.top = lswCargos.top + lswCargos.Height + 100
cmdReporte.Left = lswCargos.Left + lswCargos.Width - cmdReporte.Width

lblNodeLinea(0).top = cmdReporte.top
lblNodeLinea(2).top = lblNodeLinea(0).top
lblNodeLinea(1).top = lblNodeLinea(0).top + 240

End Sub



Private Sub lswCargos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "insert CRD_CARGOS_ASG_DETALLE(cod_cargo,codigo,cod_destino,garantia) values('" _
           & Item.Text & "','" & lblNodeLinea.Item(0).Tag & "','" & lblNodeLinea.Item(1).Tag _
           & "','" & lblNodeLinea.Item(2).Tag & "')"
Else
    strSQL = "delete CRD_CARGOS_ASG_DETALLE where cod_cargo = '" _
           & Item.Text & "' and codigo = '" & lblNodeLinea.Item(0).Tag & "' and cod_destino = '" _
           & lblNodeLinea.Item(1).Tag & "' and Garantia = '" & lblNodeLinea.Item(2).Tag & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then Exit Sub

Dim strSQL As String

Me.MousePointer = vbHourglass

Select Case Item.Index
  Case 1 'Asignacion
        vCodigo = ""
        lbl.Caption = ""
        lswCargos.ListItems.Clear

        Call sbRefrescaArbol

  Case 2 'Tabla de Montos y Plazos
    vPaso = True
        strSQL = "select rtrim(COD_CARGO) as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
               & " FROM CARGOS_ADICIONALES WHERE TIPO = 'T' AND ACTIVO = 1"
        Call sbCbo_Llena_New(cboCargo, strSQL, False, True)
    vPaso = False
    
    Call cboCargo_Click
    
End Select
Me.MousePointer = vbDefault

End Sub

Private Sub sbConsulta()
Dim strSQL As String


tcMain.Item(0).Selected = True

strSQL = "select cod_cargo,descripcion,automatico,AUMENTA_BASE_CRD,base,tipo,valor,cod_cuenta,tipo_deduccion" _
       & ",plazo_tipo,plazo_dias,monto_inicio,monto_corte,diferido_cargo,diferido_cod_cuenta" _
       & ", iva_porcentaje, activo " _
       & " from cargos_adicionales" _
       & " order by cod_cargo"
Call sbCargaGridLocal(vGrid, 17, strSQL)

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbConsulta

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 8 Or vGrid.ActiveCol = 15) Then
   Call sbgCntCuentaConsulta("D")
   vGrid.col = vGrid.ActiveCol
   vGrid.Row = vGrid.ActiveRow
   vGrid.Text = fxCntX_CuentaFormato(True, gBusquedas.Resultado, 0)
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

End Sub




Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Lineas", "Lineas") ', "imgRoot"
  'Crear Arbol Inicial
  
    strSQL = "select codigo,descripcion" _
           & " from catalogo where retencion = 'N' and Poliza = 'N' and Activo = 1"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Call sbCreaNodos(vNode.Key, rs!Codigo & " - " & rs!Descripcion, "imgFolder", True, "N", "0x0" & rs!Codigo & "L")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With


End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim rsTmp As New ADODB.Recordset, vCodTmp As String


On Error Resume Next

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Lineas" Then

Select Case Right(Node.Key, 1)
        
    Case "L" 'Lineas
    
        vCodTmp = fxIndiceCodigo(Node.Key)
              
        strSQL = "select T.*" _
               & " from crd_catalogo_garantias C inner join crd_garantia_tipos T on C.garantia = T.garantia" _
               & " where C.codigo = '" & vCodTmp & "'"
        Call OpenRecordSet(rsTmp, strSQL, 0)
                        
        strSQL = "select * from catalogo_destinos" _
               & " where cod_destino in (select cod_destino from CATALOGO_DESTINOSASG" _
               & " where codigo = '" & vCodTmp & "')"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          'Destinos y Garantias
          Call sbCreaNodos(Node.Key, rs!cod_destino & " - " & rs!Descripcion, "imgFolder", True, "N", "0x0" & vCodTmp & "-" & rs!cod_destino & "D")
          
          If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
          Do While Not rsTmp.EOF
             Call sbCreaNodos("0x0" & vCodTmp & "-" & rs!cod_destino & "D", rsTmp!Descripcion, "imgAsientos", False, "N", "0x0" & vCodTmp & "-" & rs!cod_destino & "D" & "-" & rsTmp!Garantia & "G")
            rsTmp.MoveNext
          Loop
          
          rs.MoveNext
        Loop
        rs.Close
        rsTmp.Close
    
    Case Else 'SubCuentas
     ''
End Select

End If

End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
               , vAcepta As String, Optional xkey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
'    nodx.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
    
End Sub


Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If col = 8 Or col = 15 Then
    vGrid.Row = Row
    vGrid.col = col
    vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text, 0)
End If

End Sub


Private Sub vgTabla_KeyDown(KeyCode As Integer, Shift As Integer)

Dim i As Long

With vgTabla

If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarTabla
  .Row = .ActiveRow
  .col = 1
  If .MaxRows <= .ActiveRow Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
  End If
End If


If .ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  .col = .ActiveCol
  .Row = .ActiveRow
  .Text = .Text
End If


End With


End Sub
