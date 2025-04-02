VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_CambioInfoEstadistica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Créditos y Recaudación: Cambio Masivo de Información Relacionada"
   ClientHeight    =   7920
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10680
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDato 
      Height          =   312
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12091
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   8760
      TabIndex        =   4
      Top             =   1320
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCR_CambioInfoEstadistica.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   10452
      _Version        =   524288
      _ExtentX        =   18436
      _ExtentY        =   9335
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
      MaxCols         =   493
      SpreadDesigner  =   "frmCR_CambioInfoEstadistica.frx":0A1E
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   372
      Left            =   9240
      TabIndex        =   6
      Top             =   1320
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCR_CambioInfoEstadistica.frx":0FC2
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   372
      Left            =   9720
      TabIndex        =   7
      Top             =   1320
      Width           =   492
      _Version        =   1310723
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Appearance      =   16
      Picture         =   "frmCR_CambioInfoEstadistica.frx":1985
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   372
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   6852
      _Version        =   1310723
      _ExtentX        =   12086
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
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
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   492
      Left            =   7560
      TabIndex        =   9
      Top             =   7320
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_CambioInfoEstadistica.frx":2164
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   492
      Left            =   8880
      TabIndex        =   10
      Top             =   7320
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCR_CambioInfoEstadistica.frx":293C
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   2
      Left            =   720
      TabIndex        =   11
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmCR_CambioInfoEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vProcessId As String

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    vProcessId = ""
End Sub


Private Sub sbCarga_Listado()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim pInicial As Boolean


If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "" 'Inicializa Bloque

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")

pInicial = True

vProcessId = Mid(cboTipo, 1, 4) & "." & cboDato.ItemData(cboDato.ListIndex)
vProcessId = Mid(vProcessId, 1, 10)

strSQL = ""
Do While Not rsExcel.EOF
  If pInicial Then
    pInicial = False
    strSQL = "exec spSys_Carga_Masiva 'C','" & vProcessId & "','" & glogon.Usuario & "','" & rsExcel!Operacion & "','',1"
  Else
    strSQL = strSQL & Space(10) & "exec spSys_Carga_Masiva 'C','" & vProcessId & "','" & glogon.Usuario & "','" & rsExcel!Operacion & "','',0"
  End If
  
  If Len(strSQL) > 20000 Then
     Call ConectionExecute(strSQL)
     strSQL = ""
  End If
  
  rsExcel.MoveNext
Loop
rsExcel.Close
            
'Procesa el Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If
            
            
'Vincula y Regresa Set
            
strSQL = "exec spSys_Carga_Masiva_Credito_Ref 'C','" & vProcessId & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0

     Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .col = 1
        .Text = rs!id_solicitud
        .col = 2
        .Text = rs!Codigo
        .col = 3
        .Text = rs!Cedula
        .col = 4
        .Text = rs!Nombre
      
      rs.MoveNext
    Loop
    rs.Close
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0

End Sub


Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos cargados ...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
End Sub

Private Sub btnBuscar_Click()
txtArchivo.Text = ""

With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo [Microsoft EXCEL]"
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName
End With


End Sub

Private Sub btnCancelar_Click()
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub btnCargar_Click()
    Call sbCarga_Listado
End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: OPERACION" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub


Private Sub cboTipo_Click()
Dim strSQL As String


If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

Select Case cboTipo.Text

  Case "Deductora"
    strSQL = "select COD_INSTITUCION as 'IdX',  rtrim(DESCRIPCION) + space(10)+ '['+ rtrim(isnull(DESC_CORTA,'')) + ']' as 'ItmX'" _
           & " From INSTITUCIONES" _
           & " Where ACTIVA = 1 And DEDUCCION_PLANILLA = 1" _
           & " order by rtrim(DESC_CORTA)"
    
    Case "Destino o Plan de Inversión"
        strSQL = "select rtrim(COD_DESTINO) as 'IdX', DESCRIPCION + space(10) + '[' + rtrim(cod_destino) + ']' as 'ItmX' " _
               & "  From CATALOGO_DESTINOS" _
               & "  order by COD_DESTINO"
    
    Case "Recurso Presupuestario"
        strSQL = "select rtrim(COD_GRUPO) as 'IdX', DESCRIPCION + space(10) + '[' + rtrim(COD_GRUPO) + ']' as 'ItmX' " _
               & "  From CATALOGO_GRUPOS" _
               & "  order by COD_GRUPO"
End Select

Call sbCbo_Llena_New(cboDato, strSQL, False, True)

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()

vModulo = 1

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


vPaso = True
    cboTipo.Clear
    cboTipo.AddItem "Deductora"
    cboTipo.AddItem "Destino o Plan de Inversión"
    cboTipo.AddItem "Recurso Presupuestario"
    cboTipo.Text = "Deductora"
vPaso = False
Call cboTipo_Click

txtArchivo.Text = ""

vGrid.MaxCols = 4
vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbProcesar()
Dim strSQL As String, lng As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSys_Carga_Masiva_Credito_Procesa 'C','" & vProcessId & "','" & glogon.Usuario _
       & "','" & cboDato.ItemData(cboDato.ListIndex) & "'"

Call ConectionExecute(strSQL)
Call Bitacora("Aplica", "Cambio Masivo de " & cboTipo.Text & ", Listado de Excel: Líneas(" & vGrid.MaxRows & ")")

Me.MousePointer = vbDefault

MsgBox "Cambios Registrados Satisfactoriamente...", vbInformation

txtArchivo.Text = ""
vGrid.MaxRows = 0

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub




