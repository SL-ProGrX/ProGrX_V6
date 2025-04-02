VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_BeneficiosCargaLote 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Carga de Beneficios Masivo"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5535
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   13455
      _Version        =   1441793
      _ExtentX        =   23733
      _ExtentY        =   9763
      _StockProps     =   77
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
      ShowBorder      =   0   'False
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   12600
      Top             =   600
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   1440
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_BeneficiosCargaLote.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   1440
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_BeneficiosCargaLote.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   1440
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_BeneficiosCargaLote.frx":0E19
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   873
      _StockProps     =   77
      ForeColor       =   0
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   2640
      TabIndex        =   4
      Top             =   1080
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   495
      Left            =   9960
      TabIndex        =   9
      Top             =   8160
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmAF_BeneficiosCargaLote.frx":1532
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   8160
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextAlignment   =   1
      Appearance      =   16
      Picture         =   "frmAF_BeneficiosCargaLote.frx":1C59
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   11160
      TabIndex        =   12
      ToolTipText     =   "Exportar"
      Top             =   1440
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_BeneficiosCargaLote.frx":2359
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   9600
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   195
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   15
      Top             =   480
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Formato Estandar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Checked         =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   375
      Index           =   1
      Left            =   6120
      TabIndex        =   16
      Top             =   480
      Width           =   3255
      _Version        =   1441793
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Formato con Beneficiarios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblItems 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   8160
      Width           =   2175
      _Version        =   1441793
      _ExtentX        =   3836
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   13455
      _Version        =   1441793
      _ExtentX        =   23733
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Beneficio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13575
   End
End
Attribute VB_Name = "frmAF_BeneficiosCargaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Private Sub sbLimpia()
    lsw.ListItems.Clear
End Sub


Private Sub sbCarga_Listado()
Dim rsExcel As New ADODB.Recordset
Dim itmX As ListViewItem

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "" 'Inicializa Bloque

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
    
Dim pCedula As String, pNombre As String, pMonto As Currency
Dim pBeneId As String, pBeneNombre As String, pIBAN As String
Dim pInicia As Integer, pCodigo As String


pInicia = 1
pCodigo = cbo.ItemData(cbo.ListIndex)
    

'Cargado
With rsExcel
  Do While Not .EOF
     pCedula = !Cedula & ""
     pNombre = !Nombre & ""
     pMonto = !MONTO
     
     pBeneId = "" & ""
     pBeneNombre = "" & ""
     pIBAN = "" & ""
     
      'Formato Estandar
     If btnOpcion.Item(1).Checked Then
             pBeneId = !BENEFICIARIO_ID & ""
             pBeneNombre = !BENEFICIARIO_NOMBRE & ""
             pIBAN = !BENEFICIARIO_IBAN & ""
     End If
      
     If Trim(pCedula) <> "" Then
          strSQL = strSQL & Space(10) & "exec spBeneficio_Lote_Carga '" & pCodigo & "', '" & pCedula & "', '" & pNombre & "', " & pMonto _
                 & ", '" & glogon.Usuario & "', '" & pBeneId & "', '" & pBeneNombre & "', '" & pIBAN & "', " & pInicia
          pInicia = 0
     End If
      
      If Len(strSQL) > 20000 Then
          Call ConectionExecute(strSQL)
          strSQL = ""
      End If
    
    .MoveNext
  Loop
End With

'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If


'Valida y Presentacion

strSQL = "exec spBeneficio_Lote_Revisa '" & pCodigo & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
chkTodos.Value = xtpUnchecked

Do While Not rs.EOF
Set itmX = lsw.ListItems.Add(, , rs!Cod_Beneficio)
    
    If Len(rs!Revision) = 0 Then
        itmX.SubItems(1) = "Pass!"
    Else
        itmX.SubItems(1) = "Alerta!"
    End If
    
    itmX.SubItems(2) = rs!Cedula
    itmX.SubItems(3) = rs!Nombre
    itmX.SubItems(4) = Format(rs!MONTO, "Standard")


    
     'Formato Estandar
     If btnOpcion.Item(0).Checked Then
        itmX.SubItems(5) = rs!Revision
     Else
        itmX.SubItems(5) = rs!BENEFICIARIO_ID
        itmX.SubItems(6) = rs!BENEFICIARIO_NOMBRE
        itmX.SubItems(7) = rs!BENEFICIARIO_IBAN
        itmX.SubItems(8) = rs!Revision
     End If
    


    If Len(rs!Revision) > 0 Then
        itmX.ForeColor = vbRed
        itmX.Bold = True
        itmX.TextBackColor = RGB(250, 219, 216) 'Rojo
    End If
  
  rs.MoveNext
Loop
rs.Close

lblItems.Caption = "Total de Líneas: " & lsw.ListItems.Count
    

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    lsw.ListItems.Clear

End Sub


Private Sub btnAplicar_Click()
    If lsw.ListItems.Count = 0 Then
       MsgBox "No existen casos cargados ...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
End Sub

Private Sub btnBuscar_Click()
txtArchivo.Text = ""

With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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
    lsw.ListItems.Clear
    txtArchivo.Text = ""
End Sub

Private Sub btnCargar_Click()
    Call sbCarga_Listado
End Sub

Private Sub btnExportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lsw, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnInfo_Click()

     

MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
      & " - Columnas: CEDULA, NOMBRE, MONTO, BENEFICIARIO_ID, BENEFICIARIO_NOMBRE, BENEFICIARIO_IBAN" & vbCrLf _
      & " - Nombre de la Hoja: IMPORT" _
  , vbInformation, "Información del Archivo de Carga"


End Sub



Private Sub btnOpcion_Click(Index As Integer)


lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

txtArchivo.Text = ""

btnOpcion.Item(0).Checked = False
btnOpcion.Item(1).Checked = False


btnOpcion.Item(Index).Checked = True

Select Case Index
    Case 0 'Formato Estandar
        scMain.Caption = "Formato Estandar para Asociados"
        With lsw.ColumnHeaders
            .Add , , "Código", 1200
            .Add , , "Estado", 1200, vbCenter
            .Add , , "Identificación", 2000
            .Add , , "Nombre", 4500
            .Add , , "Monto", 2100, vbRightJustify
            .Add , , "Revisión", 5100
        End With

    
    
    Case 1 'Formato para Beneficiarios
        scMain.Caption = "Formato para Acreditar a Beneficiarios"
        With lsw.ColumnHeaders
            .Add , , "Código", 1200
            .Add , , "Estado", 1200, vbCenter
            .Add , , "Identificación", 2000
            .Add , , "Nombre", 4500
            .Add , , "Monto", 2100, vbRightJustify
            .Add , , "Id Beneficiario", 2000
            .Add , , "Beneficiario", 4500
            .Add , , "IBAN", 2500
            .Add , , "Revisión", 5100
        End With

End Select

lblItems.Caption = ""

End Sub

Private Sub chkTodos_Click()
If vPaso Then Exit Sub

Dim i As Long

For i = 1 To lsw.ListItems.Count
    lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub Form_Activate()
vModulo = 7

End Sub

Private Sub Form_Load()

vModulo = 7

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

lsw.Checkboxes = True

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbProcesar()
Dim lng As Long

Dim pCedula As String, pNombre As String, pMonto As Currency
Dim pBeneId As String, pBeneNombre As String, pIBAN As String
Dim pInicia As Integer, pCodigo As String

On Error GoTo vError

pInicia = 1
pCodigo = cbo.ItemData(cbo.ListIndex)


With lsw.ListItems

'Limpia Información solo con los casos a Procesar
For lng = 1 To .Count
 If .Item(lng).Checked Then
     pCedula = .Item(lng).SubItems(2)
     pNombre = .Item(lng).SubItems(3)
     pMonto = CCur(.Item(lng).SubItems(4))
     
     pBeneId = "" & ""
     pBeneNombre = "" & ""
     pIBAN = "" & ""
     
      'Formato Estandar
     If btnOpcion.Item(1).Checked Then
             pBeneId = .Item(lng).SubItems(5)
             pBeneNombre = .Item(lng).SubItems(6)
             pIBAN = .Item(lng).SubItems(7)
     End If


     If Trim(pCedula) <> "" Then
          strSQL = strSQL & Space(10) & "exec spBeneficio_Lote_Carga '" & pCodigo & "', '" & pCedula & "', '" & pNombre & "', " & pMonto _
                 & ", '" & glogon.Usuario & "', '" & pBeneId & "', '" & pBeneNombre & "', '" & pIBAN & "', " & pInicia
          pInicia = 0
     End If
      
      If Len(strSQL) > 20000 Then
          Call ConectionExecute(strSQL)
          If Not glogon.error Then
              strSQL = ""
          End If
      End If
    
 End If 'Checked
 
Next lng

End With

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If Not glogon.error Then
       strSQL = ""
   End If
End If


'Procesamiento

If btnOpcion.Item(0).Checked Then
    'Formato Estandar
    strSQL = "exec spBeneficio_Lote_Procesa '" & pCodigo & "', '" & glogon.Usuario & "', 'S'"
Else
    'Formato con Beneficiarios
    strSQL = "exec spBeneficio_Lote_Procesa '" & pCodigo & "', '" & glogon.Usuario & "', 'B'"
End If

Call ConectionExecute(strSQL)


Call Bitacora("Aplica", "Carga Masiva de Beneficios: " & pCodigo)


lsw.ListItems.Clear
chkTodos.Value = xtpUnchecked

Me.MousePointer = vbDefault

MsgBox "Beneficios Procesados Satisfactoriamente!", vbInformation

txtArchivo.Text = ""
lsw.ListItems.Clear

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
    txtArchivo.Text = ""
    lsw.ListItems.Clear


End Sub




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False


strSQL = "select rtrim(COD_BENEFICIO) as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
       & " From AFI_BENEFICIOS" _
       & " Where ESTADO = 'A' and TIPO_MONETARIO = 1" _
       & " order by DESCRIPCION"

Call sbCbo_Llena_New(cbo, strSQL, False, True)

Call btnOpcion_Click(0)

End Sub


