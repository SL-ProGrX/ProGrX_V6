VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAH_ConciliacionAportes 
   Caption         =   "Conciliacion de Aportes"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   Icon            =   "AH_ConciliacionAportes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtResultados 
      Height          =   315
      Left            =   9240
      TabIndex        =   6
      Text            =   "1000"
      Top             =   0
      Width           =   615
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   4800
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton cmdArchivo 
      Caption         =   "&Crear Archivo"
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   165
      Left            =   0
      TabIndex        =   1
      Top             =   3930
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4695
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cédula"
         Object.Width           =   2152
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Documento"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Fecha"
         Object.Width           =   2119
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cuenta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Movimiento"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Tipo"
         Object.Width           =   1482
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   92078083
      CurrentDate     =   36699
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   92078083
      CurrentDate     =   36699
   End
   Begin VB.Image imgReporte 
      Height          =   240
      Left            =   10560
      Picture         =   "AH_ConciliacionAportes.frx":08CA
      ToolTipText     =   "Reporte del Auxiliar de Aportes"
      Top             =   -30
      Width           =   240
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "líneas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   4
      Left            =   9840
      TabIndex        =   12
      Top             =   0
      Width           =   645
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   2880
      TabIndex        =   11
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   7730
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1130
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   1080
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Limpiar Resultados cada "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Index           =   3
      Left            =   7080
      TabIndex        =   7
      Top             =   0
      Width           =   2205
   End
End
Attribute VB_Name = "frmAH_ConciliacionAportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchivo_Click()
Dim fn, itmX As ListItem, lng As Long
Dim strLinea As String, i As Integer
On Error Resume Next

fn = FreeFile
 frmContenedor.dlg.InitDir = "C:\"
 frmContenedor.dlg.ShowSave
 
 Kill frmContenedor.dlg.FileName
 
prgBar.Value = 1
prgBar.Max = lsw.ListItems.Count + 1
Open frmContenedor.dlg.FileName For Output As #fn

'Titulos
strLinea = "Sistema ASE : Conciliación de Aportes del " & dtpDesde.Value & " hasta " & dtpHasta.Value
Print #fn, strLinea
strLinea = ""
Print #fn, strLinea

strLinea = UCase(lsw.ColumnHeaders.Item(1).Text)
For i = 2 To lsw.ColumnHeaders.Count
  strLinea = strLinea & vbTab & UCase(lsw.ColumnHeaders.Item(i).Text)
Next i
Print #fn, strLinea

'Datos
For lng = 1 To lsw.ListItems.Count
  lblEstado.Caption = " Guardando Información (" & CInt((prgBar.Value / prgBar.Max) * 100) & "%)"
  lblEstado.Refresh
  lsw.SelectedItem = lsw.ListItems(lng)
     
  strLinea = lsw.SelectedItem.Text
  For i = 1 To lsw.ColumnHeaders.Count - 1
      strLinea = strLinea & vbTab & lsw.SelectedItem.SubItems(i)
  Next i
  Print #fn, strLinea
  
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1

Next lng
 
 Close #fn
 lblEstado.Caption = ""
 MsgBox "Información Guardada en " & frmContenedor.dlg.FileName, vbInformation

End Sub

Private Sub sbCrearArchivoEnLinea(fn, Optional itmX As ListItem)
Dim strLinea As String, i As Integer
On Error Resume Next

strLinea = itmX.Text
For i = 1 To itmX.ListSubItems.Count
     strLinea = strLinea & vbTab & itmX.SubItems(i)
Next i
Print #fn, strLinea

End Sub

Private Function fxCuentaAportes(strCuenta As String, vCuentas() As String) As Boolean
Dim i As Integer

i = 1
strCuenta = Format(Trim(strCuenta), GLOBALES.gstrMascara)
fxCuentaAportes = False

For i = 1 To 10
  If strCuenta = vCuentas(i) Then
    fxCuentaAportes = True
    Exit Function
  End If
Next i


End Function


Private Function fxNombre(vCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "Select Nombre from socios where cedula = '" & Trim(vCedula) & "'", glogon.Conection, adOpenStatic
 
If rsX.EOF And rsX.BOF Then
  fxNombre = "No Encontrado"
Else
  fxNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close
End Function


Private Sub cmdBuscar_Click()
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListItem, rs2 As New ADODB.Recordset
Dim lng As Long, i As Integer, iArchivo As Integer
Dim strArchivo As String, fn, strLinea As String
Dim lngLineas As Long, vCuenta(10) As String


'1. Meter en un arreglo las cuentas formateadas, de los diferentes aportes
'2. Buscar en el detalle de ahorros, segun la fecha y sin importar el estado
'3. Buscar los ahorros liquidados en las fechas
'4. Buscar Reingresos

On Error GoTo vError

lngLineas = txtResultados

iArchivo = MsgBox("Desea Crear Archivo de Texto Automáticamente", vbYesNo)

If iArchivo = vbYes Then
    fn = FreeFile
    strArchivo = "C:\AD" & Format(Day(Date), "00") & "M" & Format(Month(Date), "00") _
            & "A" & Format(Year(Date), "0000") & ".TXT"
    
    On Error Resume Next
    Kill strArchivo
     
    Open strArchivo For Output As #fn
    strLinea = "Sistema ASE : Conciliación de Aportes del " & dtpDesde.Value & " hasta " & dtpHasta.Value
    Print #fn, strLinea
    strLinea = ""
    Print #fn, strLinea
    strLinea = UCase(lsw.ColumnHeaders.Item(1).Text)
    For i = 2 To lsw.ColumnHeaders.Count
      strLinea = strLinea & vbTab & UCase(lsw.ColumnHeaders.Item(i).Text)
    Next i
    Print #fn, strLinea
End If

On Error GoTo vError


'Carga en la Variable las cuentas de aportes
rs.Open "select * from par_afah", glogon.Conection, adOpenStatic
  vCuenta(1) = Format(Trim(rs!cta_obrero), GLOBALES.gstrMascara)
  vCuenta(2) = Format(Trim(rs!cta_patronal), GLOBALES.gstrMascara)
  vCuenta(3) = Format(Trim(rs!cta_capitaliza), GLOBALES.gstrMascara)
  vCuenta(4) = Format(Trim(rs!cta_extra), GLOBALES.gstrMascara)
  vCuenta(5) = Format(Trim(rs!cta_patronalfci), GLOBALES.gstrMascara)
  vCuenta(6) = Format(Trim(rs!cta_custodia), GLOBALES.gstrMascara)
  vCuenta(7) = Format(Trim(rs!cta_custodiafci), GLOBALES.gstrMascara)
rs.Close

lsw.ListItems.Clear

'*******************************************************************************************
'FASE I = BUSCA EN AHORRO DETALLADO

lblEstado.Caption = "Cargando FASE I ..."
lblEstado.Refresh

glogon.Conection.CommandTimeout = 2000

Me.MousePointer = vbHourglass

strSQL = "Select D.*,S.nombre" _
       & " from Ahorro_detallado D inner join Socios S on D.cedula = S.cedula" _
       & " where fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") _
       & "' and '" & Format(dtpHasta.Value, "yyyy/mm/dd") & "'"

rs.CursorLocation = adUseClient
rs.Open strSQL, glogon.Conection, adOpenStatic

lng = 1

If Not rs.EOF And Not rs.BOF Then
  prgBar.Max = rs.RecordCount + 1
  prgBar.Value = 1
End If
Do While Not rs.EOF
 If lng = 1 Then
   lblEstado.Caption = "Analizando Cola de Aportes - Fase I (Total:" & prgBar.Max _
                    & " Procesando: " & prgBar.Value & " Relativo: " _
                    & CInt((prgBar.Value / prgBar.Max) * 100) & "%)"
  lblEstado.Refresh
 End If
  
 lng = lng + 1
 If lng = 50 Then lng = 1
   
    Set itmX = lsw.ListItems.Add(, , rs!Cedula)
        itmX.Tag = itmX.Index
        itmX.SubItems(1) = IIf(IsNull(rs!Nombre), "", rs!Nombre)
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
        itmX.SubItems(3) = IIf(IsNull(rs!numcom), "", rs!numcom)
        itmX.SubItems(4) = Format(rs!Fecha, "yyyy/mm/dd")
        
        Select Case rs!Tipo
          Case Is = "O" 'Ahorro Obrero
             itmX.SubItems(5) = vCuenta(1)
            
          Case Is = "P" 'Aporte Patronal
             itmX.SubItems(5) = vCuenta(2)
          
          Case Is = "C" 'Capitalizacion
             itmX.SubItems(5) = vCuenta(3)
          
          Case Is = "E" 'Extraordinario
             itmX.SubItems(5) = vCuenta(4)
        
        End Select
          
          
        If Mid(Trim(rs!numcom), 1, 2) = "ND" Then
            itmX.SubItems(6) = "Anulación"
            itmX.SubItems(7) = "E"
        Else
            itmX.SubItems(6) = "Registro"
            itmX.SubItems(7) = "I"
        End If
        
    If iArchivo = vbYes Then Call sbCrearArchivoEnLinea(fn, itmX)
  
  If lsw.ListItems.Count > lngLineas Then lsw.ListItems.Clear
  
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  rs.MoveNext
Loop
rs.Close


'FASE II - Liquidaciones y Reingresos

'*******************************************************************************************
'FASE I = BUSCA EN AHORRO DETALLADO

lblEstado.Caption = "Cargando FASE II ..."
lblEstado.Refresh


strSQL = "Select * from asientos_tmp" _
       & " where tmp_fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") _
       & "' and '" & Format(dtpHasta.Value, "yyyy/mm/dd") & "' and tmp_tipo in('LIQ','ING')"

rs.CursorLocation = adUseClient
rs.Open strSQL, glogon.Conection, adOpenStatic

lng = 1

If Not rs.EOF And Not rs.BOF Then
  prgBar.Max = rs.RecordCount + 1
  prgBar.Value = 1
End If
Do While Not rs.EOF
 If lng = 1 Then
   lblEstado.Caption = "Analizando Cola de Asientos - Fase II (Total:" & prgBar.Max _
                    & " Procesando: " & prgBar.Value & " Relativo: " _
                    & CInt((prgBar.Value / prgBar.Max) * 100) & "%)"
  lblEstado.Refresh
 End If
  
 lng = lng + 1
 If lng = 50 Then lng = 1
     
     
  If fxCuentaAportes(rs!tmp_cuenta, vCuenta) Then
   
   Select Case rs!tmp_tipo
      
      Case Is = "LIQ"
        strSQL = "select S.cedula,S.nombre" _
               & " from liquidacion L inner join Socios S on L.cedula = S.cedula" _
               & " where consec = " & IIf(Mid(rs!tmp_caso, 1, 1) = "R", Mid(rs!tmp_caso, 2, 30), rs!tmp_caso)
       rs2.Open strSQL, glogon.Conection, adOpenStatic
       If Not rs2.EOF And Not rs2.BOF Then
        Set itmX = lsw.ListItems.Add(, , rs2!Cedula)
            itmX.Tag = itmX.Index
            itmX.SubItems(1) = IIf(IsNull(rs2!Nombre), "", rs2!Nombre)
            itmX.SubItems(6) = "Liquidación"
       Else
        Set itmX = lsw.ListItems.Add(, , "Reversada")
            itmX.Tag = itmX.Index
            itmX.SubItems(1) = "Reversada"
            itmX.SubItems(6) = "Liquidación"
       End If
       rs2.Close
        
      Case Is = "ING"
   
        Set itmX = lsw.ListItems.Add(, , rs!tmp_caso)
            itmX.Tag = itmX.Index
            itmX.SubItems(1) = fxNombre(rs!tmp_caso)
            itmX.SubItems(6) = "Reingreso"
   
   End Select
   
        itmX.SubItems(2) = Format(rs!tmp_Monto, "Standard")
        itmX.SubItems(3) = rs!tmp_tipo & "-" & IIf(IsNull(rs!tmp_caso), "", rs!tmp_caso)
        itmX.SubItems(4) = Format(rs!tmp_Fecha, "yyyy/mm/dd")
        itmX.SubItems(5) = Format(Trim(rs!tmp_cuenta), GLOBALES.gstrMascara)
          
        If rs!TMP_DEBEHABER = "D" Then
            itmX.SubItems(7) = "E"
        Else
            itmX.SubItems(7) = "I"
        End If
        
    If iArchivo = vbYes Then Call sbCrearArchivoEnLinea(fn, itmX)
  
  If lsw.ListItems.Count > lngLineas Then lsw.ListItems.Clear
  
  End If 'Validacion de Cuenta de Aporte
  
  If prgBar.Value < prgBar.Max Then prgBar.Value = prgBar.Value + 1
  rs.MoveNext
Loop
rs.Close

lblEstado.Caption = ""
Me.MousePointer = vbDefault

MsgBox "Consulta Finalizada...", vbInformation
If iArchivo = vbYes Then
 Close #fn
 MsgBox " SE CREÓ ARCHIVO EN 'C:\AD" & Format(Day(Date), "00") & "M" & Format(Month(Date), "00") _
        & "A" & Format(Year(Date), "0000") & ".TXT '", vbInformation
End If


Exit Sub


vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
 
End Sub


Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpHasta.SetFocus
End Sub

Private Sub dtpHasta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdBuscar.SetFocus
End Sub

Private Sub Form_Load()
 dtpDesde.Value = fxFechaServidor
 dtpHasta.Value = dtpDesde.Value
End Sub

Private Sub Form_Resize()
On Error Resume Next
    lsw.Height = Me.Height - 1270
    lsw.Width = Me.Width - 130
    lblEstado.Width = lsw.Width
End Sub

Private Sub imgReporte_Click()
Dim strSQL As String, strRuta As String, strInicio As String, strFinal As String

On Error GoTo vError

Me.MousePointer = vbHourglass
strRuta = SIFGlobal.fxSIFPathReportes("AuxiliarAportes.rpt")


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Auxiliar de Aportes"

.ReportFileName = strRuta
.Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
.Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
.Formulas(3) = "Subtitulo='AL : " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"

.SelectionFormula = "{SOCIOS.ESTADOACTUAL} <> 'N'"

.PrintReport

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
Me.MousePointer = vbDefault
MsgBox Err.Description, vbCritical

End Sub
