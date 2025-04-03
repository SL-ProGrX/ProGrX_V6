VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmAF_CambioInfoEstadistica 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes: Actualización Masiva de Datos de la Persona"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   11010
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   9480
      TabIndex        =   0
      Top             =   1200
      Width           =   492
      _Version        =   1310722
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_CambioInfoEstadistica.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   7932
      _Version        =   524288
      _ExtentX        =   13991
      _ExtentY        =   8911
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
      MaxCols         =   491
      SpreadDesigner  =   "frmAF_CambioInfoEstadistica.frx":0700
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CambioInfoEstadistica.frx":0BF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CambioInfoEstadistica.frx":1615
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CambioInfoEstadistica.frx":1FD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   372
      Left            =   9960
      TabIndex        =   2
      Top             =   1200
      Width           =   492
      _Version        =   1310722
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_CambioInfoEstadistica.frx":27B7
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   372
      Left            =   10440
      TabIndex        =   3
      Top             =   1200
      Width           =   492
      _Version        =   1310722
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      Appearance      =   16
      Picture         =   "frmAF_CambioInfoEstadistica.frx":2ED0
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   1200
      Width           =   6852
      _Version        =   1310722
      _ExtentX        =   12086
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   6852
      _Version        =   1310722
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.ComboBox cboDato 
      Height          =   312
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   6852
      _Version        =   1310722
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   492
      Left            =   8280
      TabIndex        =   7
      Top             =   7200
      Width           =   1332
      _Version        =   1310722
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
      Picture         =   "frmAF_CambioInfoEstadistica.frx":35E9
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   492
      Left            =   9600
      TabIndex        =   8
      Top             =   7200
      Width           =   1332
      _Version        =   1310722
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
      Picture         =   "frmAF_CambioInfoEstadistica.frx":3D10
      ImageAlignment  =   4
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
      Left            =   1080
      TabIndex        =   11
      Top             =   240
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
      Left            =   1080
      TabIndex        =   10
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
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
      Height          =   372
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
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
Attribute VB_Name = "frmAF_CambioInfoEstadistica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbLimpia()
    vGrid.MaxRows = 0
End Sub


Private Sub sbCarga_Listado()
Dim strSQL As String, rsExcel As New ADODB.Recordset
If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "" 'Inicializa Bloque

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
    
With vGrid
    .MaxRows = 0

     Do While Not rsExcel.EOF
        If Not IsNull(rsExcel!Cedula) Then
            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
            .col = 1
            .Text = rsExcel!Cedula
            .col = 2
            .Text = rsExcel!Nombre
        End If
      rsExcel.MoveNext
    Loop
    rsExcel.Close
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
    vGrid.MaxRows = 0
    txtArchivo.Text = ""
End Sub

Private Sub btnCargar_Click()
    Call sbCarga_Listado
End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: CEDULA, NOMBRE" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub

Private Sub cboTipo_Click()
Dim strSQL As String


If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

Select Case cboTipo.Text
  Case "Sectores"
    strSQL = "select cod_sector as Idx, descripcion as ItmX from afi_sectores"
  Case "Profesiones"
    strSQL = "select cod_profesion as Idx, descripcion as ItmX from afi_profesiones"
  Case "Deductora"
    strSQL = "select COD_INSTITUCION as 'IdX',  rtrim(DESCRIPCION) + space(10)+ '['+ rtrim(isnull(DESC_CORTA,'')) + ']' as 'ItmX'" _
           & " From INSTITUCIONES" _
           & " Where ACTIVA = 1 And DEDUCCION_PLANILLA = 1" _
           & " order by rtrim(DESC_CORTA)"
  Case "Institución"
    strSQL = "select COD_INSTITUCION as 'IdX',  rtrim(DESCRIPCION) + space(10)+ '['+ rtrim(isnull(DESC_CORTA,'')) + ']' as 'ItmX'" _
           & " From INSTITUCIONES" _
           & " Where ACTIVA = 1 And DEDUCCION_PLANILLA = 1" _
           & " order by rtrim(DESC_CORTA)"
    
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
    cboTipo.AddItem "Sectores"
    cboTipo.AddItem "Profesiones"
    cboTipo.AddItem "Deductora"
    cboTipo.AddItem "Institución"
    cboTipo.Text = "Sectores"
vPaso = False
Call cboTipo_Click

txtArchivo.Text = ""

vGrid.MaxCols = 2
vGrid.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbProcesar()
Dim strSQL As String, lng As Long


With vGrid
    
    For lng = 1 To .MaxRows
    
       .Row = lng
       .col = 1
       
       Select Case cboTipo.Text
         Case "Sectores"
            strSQL = strSQL & Space(10) & "update socios set cod_sector = " & cboDato.ItemData(cboDato.ListIndex) _
                   & " where cedula = '" & .Text & "'"
         Case "Profesionales"
            strSQL = strSQL & Space(10) & "update socios set cod_profesion = " & cboDato.ItemData(cboDato.ListIndex) _
                   & " where cedula = '" & .Text & "'"
        
         Case "Deductora"
            strSQL = strSQL & Space(10) & "update socios set cod_Deductora = " & cboDato.ItemData(cboDato.ListIndex) _
                   & " where cedula = '" & .Text & "'"
       
         Case "Institución"
            strSQL = strSQL & Space(10) & "update socios set cod_Institucion = " & cboDato.ItemData(cboDato.ListIndex) _
                   & " where cedula = '" & .Text & "'"
       
       End Select
       
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

Call Bitacora("Aplica", "Cambio Masivo de " & cboTipo.Text & ", Listado de Excel: Líneas(" & vGrid.MaxRows & ")")

Me.MousePointer = vbDefault



MsgBox "Cambios Registrados Satisfactoriamente...", vbInformation

txtArchivo.Text = ""
vGrid.MaxRows = 0

End Sub


