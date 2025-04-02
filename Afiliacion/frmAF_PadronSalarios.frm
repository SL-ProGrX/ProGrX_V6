VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmAF_PadronSalarios 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Carga de Padron de Empleados y Salarios"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5535
      Left            =   0
      TabIndex        =   10
      Top             =   2400
      Width           =   12855
      _Version        =   1441792
      _ExtentX        =   22675
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin VB.Timer Timerx 
      Interval        =   10
      Left            =   12000
      Top             =   480
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   240
      Width           =   3375
      _Version        =   1441792
      _ExtentX        =   5953
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Padron de Empleados"
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   1440
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronSalarios.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   375
      Left            =   10080
      TabIndex        =   1
      Top             =   1440
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronSalarios.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   1440
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronSalarios.frx":0E19
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   6855
      _Version        =   1441792
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
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   495
      Left            =   9720
      TabIndex        =   5
      Top             =   8160
      Width           =   1335
      _Version        =   1441792
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
      Picture         =   "frmAF_PadronSalarios.frx":1532
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   495
      Left            =   11040
      TabIndex        =   6
      Top             =   8160
      Width           =   1335
      _Version        =   1441792
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
      Picture         =   "frmAF_PadronSalarios.frx":1C59
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnOpcion 
      Height          =   495
      Index           =   1
      Left            =   6120
      TabIndex        =   9
      Top             =   240
      Width           =   3375
      _Version        =   1441792
      _ExtentX        =   5953
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Salarios"
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   330
      Left            =   2640
      TabIndex        =   12
      Top             =   1080
      Width           =   6855
      _Version        =   1441792
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   11160
      TabIndex        =   14
      ToolTipText     =   "Exportar"
      Top             =   1440
      Width           =   495
      _Version        =   1441792
      _ExtentX        =   873
      _ExtentY        =   661
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_PadronSalarios.frx":2359
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   9600
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3625
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin XtremeSuiteControls.Label lblItems 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   8160
      Width           =   2175
      _Version        =   1441792
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
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   2040
      Width           =   13095
      _Version        =   1441792
      _ExtentX        =   23098
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmAF_PadronSalarios"
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
    
'Padron
If btnOpcion.Item(0).Checked Then
  'Valida

  'Cargado
    
  With rsExcel
    Do While Not .EOF
      Set itmX = lsw.ListItems.Add(, , !Identificacion)
          itmX.SubItems(1) = !ID_ALTERNA & ""
          itmX.SubItems(2) = !Nombre & ""
          itmX.SubItems(3) = Format(!FECHA_INGRESO & "", "yyyy-MM-dd")
      .MoveNext
    Loop
  End With
    
End If


    
'Salarios
If btnOpcion.Item(1).Checked Then
  'Valida

  'Cargado
    
  With rsExcel
    Do While Not .EOF
      Set itmX = lsw.ListItems.Add(, , !Identificacion)
          itmX.SubItems(1) = !DIVISA & ""
          itmX.SubItems(2) = Format(!FECHA & "", "yyyy-MM-dd")
          itmX.SubItems(3) = Format(!SALARIO_BRUTO, "Standard")
          itmX.SubItems(4) = Format(!REBAJOS, "Standard")
          itmX.SubItems(5) = Format(!SALARIO_NETO, "Standard")
          itmX.SubItems(6) = !EMBARGOS
         
         If IsNumeric(!EMBARGOS) Then
          itmX.SubItems(6) = IIf(!EMBARGOS > 0, "S", "N")
         Else
            If !EMBARGOS = "S" Then
              itmX.SubItems(6) = "S"
            Else
              itmX.SubItems(6) = "N"
            End If
         End If
      
      .MoveNext
    Loop
    .Close
  End With
    
End If

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

            
If btnOpcion.Item(0).Checked Then

  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: IDENTIFICACION, ID_ALTERNA, NOMBRE, FECHA_INGRESO" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"

Else
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: IDENTIFICACION, DIVISA, FECHA, SALARIO_BRUTO, REBAJOS, SALARIO_NETO, EMBARGOS" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"

End If

End Sub



Private Sub btnOpcion_Click(Index As Integer)


lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

txtArchivo.Text = ""

btnOpcion.Item(0).Checked = False
btnOpcion.Item(1).Checked = False


btnOpcion.Item(Index).Checked = True

Select Case Index
    Case 0 'Padron
        scMain.Caption = "Listado para Carga de Padron de Empleados"
        With lsw.ColumnHeaders
            .Add , , "Identificación", 2000
            .Add , , "Id. Alterna", 2500
            .Add , , "Nombre", 4500
            .Add , , "Ingreso", 2000, vbCenter
        End With
    
    
    Case 1 'Salarios
        scMain.Caption = "Listado para Carga de Salarios"
        With lsw.ColumnHeaders
            .Add , , "Identificación", 2000
            .Add , , "Divisa", 1000, vbCenter
            .Add , , "Fecha", 1500, vbCenter
            .Add , , "Salario Bruto", 2100, vbRightJustify
            .Add , , "Rebajos Total", 2100, vbRightJustify
            .Add , , "Salario Neto", 2100, vbRightJustify
            .Add , , "Embargos", 1500, vbCenter
        End With

End Select

lblItems.Caption = ""

End Sub

Private Sub Form_Activate()
vModulo = 1

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbProcesar()
Dim lng As Long


On Error GoTo vError

With lsw.ListItems

    For lng = 1 To .Count

       'Padron
       If btnOpcion.Item(0).Checked Then
            strSQL = strSQL & Space(10) & "exec spAFI_Padron_Registro '" & .Item(lng).Text & "','" & .Item(lng).SubItems(1) _
                    & "','" & .Item(lng).SubItems(2) & "', " & cbo.ItemData(cbo.ListIndex) & ", '" & Format(.Item(lng).SubItems(3), "yyyy-MM-dd") _
                    & "', '" & glogon.Usuario & "', 'A'"
       
       End If
       

       'Salarios
       If btnOpcion.Item(1).Checked Then
            strSQL = strSQL & Space(10) & "exec spAFI_Persona_Salarios_Add '" & .Item(lng).Text & "','C','" & .Item(lng).SubItems(1) _
                    & "', '" & Format(.Item(lng).SubItems(2), "yyyy-MM-dd") & "', " & CCur(.Item(lng).SubItems(3)) _
                    & ", " & CCur(.Item(lng).SubItems(4)) & ", " & CCur(.Item(lng).SubItems(5)) & ", '" & .Item(lng).SubItems(6) _
                    & "', '" & glogon.Usuario & "', 'A'"
       End If
       

       If Len(strSQL) > 20000 Then
          Call ConectionExecute(strSQL)
          If Not glogon.error Then
              strSQL = ""
          End If
       End If

    Next lng

End With

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If Not glogon.error Then
       strSQL = ""
   End If
End If

'Call Bitacora("Aplica", "Cambio Masivo de " & cboTipo.Text & ", Listado de Excel: Líneas(" & vGrid.MaxRows & ")")

Me.MousePointer = vbDefault



MsgBox "Información Actualizada Satisfactoriamente!", vbInformation

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

Timerx.Interval = 0
Timerx.Enabled = False

strSQL = "select COD_INSTITUCION as 'Idx',  '[' + COD_DIVISA + ']  ' + DESCRIPCION as 'ItmX'" _
       & "  from INSTITUCIONES where ACTIVA = 1" _
       & "  order by COD_INSTITUCION"
Call sbCbo_Llena_New(cbo, strSQL, True, True)

Call btnOpcion_Click(0)

End Sub
