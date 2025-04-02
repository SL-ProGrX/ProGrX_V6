VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCR_ReportesCierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Cierre(Fin de mes)"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   2100
      TabIndex        =   0
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   113115139
      CurrentDate     =   37378
   End
   Begin MSComCtl2.DTPicker dtpFin 
      Height          =   315
      Left            =   2100
      TabIndex        =   1
      Top             =   1035
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   113115139
      CurrentDate     =   37378
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6862
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6980
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6AAA
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6BD0
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6CDE
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6DEB
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":6F04
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":7032
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":713F
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ReportesCierre.frx":723F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraCatalogo 
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   1095
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   4815
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   8493
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   10126
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbDocumento 
         Height          =   330
         Left            =   6360
         TabIndex        =   9
         Top             =   0
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1799
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgArbol"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aceptar"
               Key             =   "Aceptar"
               ImageIndex      =   11
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView ArbolReportes 
      Height          =   5160
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9102
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   176
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      ImageList       =   "imgArbol"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblFin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Final"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   495
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgImprime 
      Height          =   480
      Left            =   3840
      Picture         =   "frmCR_ReportesCierre.frx":DAA1
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Cierre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label LblInicio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fecha Inicial"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   5
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "frmCR_ReportesCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim vCodigos As String
'Dim strCodigos As String
Dim bSeleccionado As Boolean


Private Sub sbCargaNodos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iNodosCheck As Integer
Dim xNode As Node, i As Integer

iNodosCheck = 0
With ArbolReportes
  .Nodes.Clear
  '
  Set xNode = .Nodes.Add(, , "Pr", "Reportes")
  xNode.Bold = True
  'Reportes de Afiliacion
  .Nodes.Add "Pr", tvwChild, "Afiliacion", "Afliación Y Renuncias"
  .Nodes.Add "Afiliacion", tvwChild, "Ing", "Ingresos y Renuncias", 2, 1

  
  'Aporte obrero Patronal
  .Nodes.Add "Pr", tvwChild, "Aportes", "Aportes(Obrero/Patronal)"
  .Nodes.Add "Aportes", tvwChild, "SinAp", "Asociados que no posean aportes", 5, 1
  .Nodes.Add "Aportes", tvwChild, "Apa", "Asociados sólo poseen aporte patronal", 6, 1 ' (reafiliaciones que aún no tienen nuevos aportes) "
  .Nodes.Add "Aportes", tvwChild, "Apo", "Asociados sólo poseen aporte obrero", 7, 1 ' (asociados que están enviado su aporte a otra institución)"
  .Nodes.Add "Aportes", tvwChild, "UAfl", "Asociados que se han afiliado por única vez", 8, 1 '(están activos)"
  .Nodes.Add "Aportes", tvwChild, "LiqAfi", "Asociados que se liquidaron y en el mismo mes se afiliaron", 4, 1
  
  'Reportes de credito
  .Nodes.Add "Pr", tvwChild, "Creditos", "Reporte según estado del deudor, por fechas"
  .Nodes.Add "Creditos", tvwChild, "CrAfl-1", "Afiliación menor a 1 meses membresía (reingreso más de un mes de haberse liquidado)  ", 2, 1
  .Nodes.Add "Creditos", tvwChild, "CrAfi1", "Afiliación menor a 1 meses membresía, 1º Ingreso (nunca fueron asociados)", 3, 1
  .Nodes.Add "Creditos", tvwChild, "CrAfReing", "Afiliación menor a 1 meses membresía (reingreso inmediato a la renuncia)", 4, 1
  
  
  
  
End With

For i = 1 To ArbolReportes.Nodes.Count
   ArbolReportes.Nodes.Item(i).Expanded = True
Next i

Me.MousePointer = vbDefault

 
End Sub


Private Sub ArbolReportes_Click()
'vCodigos = ""
'chkTodos.Value = vbUnchecked
bSeleccionado = True

Select Case ArbolReportes.SelectedItem.Key

    Case "CrAfl-1"
'        fraCatalogo.Visible = True
'        imgImprime.Enabled = False
'        Call sbCargaCatalogo
        LblFin.Visible = True
        dtpFin.Visible = True
    Case "CrAfi1"
'        fraCatalogo.Visible = True
'        imgImprime.Enabled = False
'        Call sbCargaCatalogo
        LblFin.Visible = True
        dtpFin.Visible = True
    Case "CrAfReing"
'        fraCatalogo.Visible = True
'        imgImprime.Enabled = False
'        Call sbCargaCatalogo
        LblFin.Visible = True
        dtpFin.Visible = True
    Case Else
        LblFin.Visible = False
        dtpFin.Visible = False
End Select

End Sub

Private Sub chkTodos_Click()

Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub

Private Sub cmdAceptar_Click()
End Sub

Private Sub Form_Activate()
vModulo = 10
End Sub

Private Sub Form_Load()

vModulo = 10

Call sbCargaNodos
bSeleccionado = False
dtpInicio.Value = fxFechaServidor
dtpFin.Value = dtpInicio

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub imgImprime_Click()

Dim strSQL As String, rs As New ADODB.Recordset







If bSeleccionado Then
    With frmContenedor.Crt
     .Reset
     .WindowShowExportBtn = True
     .WindowShowGroupTree = True
     .WindowShowPrintBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "Reportes SIF Credito"
     
     .Connect = glogon.ConectRPT
     
     .Formulas(0) = "fxEmpresa = '" & GLOBALES.gstrNombreEmpresa & "'"
     .Formulas(1) = "fxUsuario = 'USUARIO: " & UCase(glogon.Usuario) & "'"
     .Formulas(2) = "fxFecha = 'FECHA:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
     .StoredProcParam(1) = Format(dtpFin.Value, "yyyy-MM-dd 23:59:59.000")
     Select Case ArbolReportes.SelectedItem.Key
        Case "LiqAfi"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 5
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreAportesIngresos.rpt")
             .Destination = crptToWindow
             
        Case "UAfl"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 11
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreAportesIngresos.rpt")
             .Destination = crptToWindow
             
             
        Case "Ing"
             .Formulas(3) = "fxTitulo = 'Afiliaciones del cierre del mes de ' & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 9
              .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrIngreso.rpt")
             .Destination = crptToWindow
             .PrintReport
    
             .Formulas(3) = "fxTitulo = 'Liquidaciones del cierre del mes de ' & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 10
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrRenuncias.rpt")
             .Destination = crptToWindow
             
        Case "SinAp"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 6
              .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreAportes.rpt")
             .Destination = crptToWindow
             
         Case "Apa"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 7
             .StoredProcParam(3) = "i"
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreAportes.rpt")
             .Destination = crptToWindow
             
        Case "Apo"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & '  '  & '" & Format(dtpInicio, "MMMM") & "' & ' del año ' & '" & Format(dtpInicio, "yyyy") & "'"
             .StoredProcParam(2) = 8
             .StoredProcParam(3) = "i"
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreAportes.rpt")
             .Destination = crptToWindow
             
        Case "CrAfl-1"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & ' de '  & '" & Format(dtpInicio, "dd/mm/yyyy") & "' & ' hasta ' & '" & Format(dtpFin, "dd/mm/yyyy") & "'"
             .StoredProcParam(2) = 2
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreCredito.rpt")
             .Destination = crptToWindow
        
        Case "CrAfi1"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & ' de '  & '" & Format(dtpInicio, "dd/mm/yyyy") & "' & ' hasta ' & '" & Format(dtpFin, "dd/mm/yyyy") & "'"
             .StoredProcParam(2) = 1
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreCredito.rpt")
             .Destination = crptToWindow
        
        Case "CrAfReing"
             .Formulas(3) = "fxTitulo = '" & ArbolReportes.SelectedItem.Text & "' & ' de '  & '" & Format(dtpInicio, "dd/mm/yyyy") & "' & ' hasta ' & '" & Format(dtpFin, "dd/mm/yyyy") & "'"
             .StoredProcParam(2) = 3
             .ReportFileName = SIFGlobal.fxSIFPathReportes("AfiCrCierreCredito.rpt")
             .Destination = crptToWindow
        
     End Select
     
     .PrintReport
'     .Action = 1
    End With
Else
   MsgBox "Seleccione un reporte....", vbInformation
End If
End Sub

Private Sub sbCargaCatalogo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

lsw.ListItems.Clear

strSQL = "select codigo,descripcion from catalogo where activo = 1 and retencion = 'N'"
rs.Open strSQL, glogon.Conection, adOpenStatic

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , Trim(rs!Codigo))
     itmX.SubItems(1) = rs!Descripcion
     rs.MoveNext
Loop
rs.Close
End Sub

Private Sub tlbDocumento_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim i As Integer
'vCodigos = "'"
'strCodigos = ""
'For i = 1 To lsw.ListItems.Count
'
'  If lsw.ListItems.Item(i).Checked = True Then
'    vCodigos = vCodigos & lsw.ListItems.Item(i).Text & "',"
'
'    strCodigos = strCodigos & "," & lsw.ListItems.Item(i).Text
'    vCodigos = vCodigos & "'"
'  End If
'Next i
'
'
'
'
'vCodigos = Mid(vCodigos, 1, Len(vCodigos) - 2)
'vCodigos = "[" & vCodigos & "]"
'strCodigos = Mid(strCodigos, 2, Len(strCodigos))
'fraCatalogo.Visible = False
'imgImprime.Enabled = True

End Sub
