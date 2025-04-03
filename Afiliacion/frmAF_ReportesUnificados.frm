VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAF_ReportesUnificados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Afiliación : Reportes Generales"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmAF_ReportesUnificados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Resumen de Socios x  Unidad Programatica (Todas)"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   4440
      TabIndex        =   26
      Top             =   3000
      Width           =   4095
   End
   Begin VB.CheckBox chkTodas 
      Caption         =   "&Todas"
      Height          =   255
      Left            =   5280
      TabIndex        =   25
      Top             =   3840
      Width           =   1095
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado General de Afiliaciones"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   24
      Top             =   3000
      Width           =   4095
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Socios Por Unidad Programatica (Todas)"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   4440
      TabIndex        =   23
      Top             =   2640
      Width           =   4095
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Socios Por Una Unidad Programatica"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   4440
      TabIndex        =   22
      Top             =   2280
      Width           =   3975
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Unidades De Trabajo"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   4440
      TabIndex        =   21
      Top             =   1920
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Unidades Programaticas"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   4440
      TabIndex        =   20
      Top             =   1560
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Promotores"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   4440
      TabIndex        =   19
      Top             =   1200
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Detalle De Afiliaciones Por Promotor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   4440
      TabIndex        =   18
      Top             =   840
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Resumen De Afiliaciones Por Promotor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   4440
      TabIndex        =   17
      Top             =   480
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Reimpresion De Boleta"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   4440
      TabIndex        =   16
      Top             =   120
      Width           =   4335
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado Detalle De Socios Por Unidad Programatica"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   15
      Top             =   3360
      Width           =   3975
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado Resumen De Socios Por Unidad Programatica"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   4335
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Ex Socios Por Provincia"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado De Socios Por Provincia"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txt 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   4080
      Width           =   1215
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Consulta De Acta De Afiliación"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Reporte"
      Height          =   855
      Left            =   7320
      Picture         =   "frmAF_ReportesUnificados.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1095
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado Ingresos - No Socios"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado Ingresos - Renuncia Patronal"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado Ingresos - Renuncia Asociacion"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.OptionButton OPT 
      Appearance      =   0  'Flat
      Caption         =   "Listado Ingresos - Socios Activos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20250627
      CurrentDate     =   36093
   End
   Begin MSComCtl2.DTPicker dtpDe 
      Height          =   315
      Left            =   3840
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   20250627
      CurrentDate     =   36093
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   8520
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label lbl 
      Caption         =   "# Acta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label lblDe 
      Caption         =   "Desde"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   3840
      Width           =   495
   End
   Begin VB.Label lblHasta 
      Caption         =   "Hasta"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   4200
      Width           =   495
   End
End
Attribute VB_Name = "frmAF_ReportesUnificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOPT As Integer
Private Sub cmdImprimir_Click()
Dim recUnidades As New ADODB.Recordset
Dim lngContador As Long, strRuta As String
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Afiliación"

  .Connect = glogon.ConectRPT
  
   Select Case vOPT
     Case 0
      If Trim(txt) = "" Then
       MsgBox "Suministre El Numero de Acta a Consultar", vbExclamation, "Faltan Datos"
       txt.SetFocus
       Me.MousePointer = vbDefault
       Exit Sub
      Else
       .ReportFileName = App.Path + "\Reportes\AfiImprimeActa.rpt"
       .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
       .SelectionFormula = "{SOCIOS.NACTA}=" & Trim(txt)
      End If
    
     Case 1 To 4, 17
       .ReportFileName = App.Path + "\Reportes\AfiIngresoSocios.rpt"
       .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    
       If Trim(txt) <> "" Then
          strSQL = "{SOCIOS.UP} = '" & Trim(txt) & "' And "
       Else
          strSQL = ""
       End If
    
       If vOPT = 1 Then
          strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = 'S'"
          .Formulas(1) = "Titulo='LISTADO DE SOCIOS'"
       ElseIf vOPT = 2 Then
          strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = 'A'"
          .Formulas(1) = "Titulo='LISTADO DE EX-SOCIOS (RENUNCIA ASOCIACION)'"
       ElseIf vOPT = 3 Then
          strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = 'P'"
          .Formulas(1) = "Titulo='LISTADO DE EX-SOCIOS (RENUNCIA PATRONAL)'"
       ElseIf vOPT = 4 Then
          strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = 'N'"
          .Formulas(1) = "Titulo='LISTADO DE NO-SOCIOS'"
       ElseIf vOPT = 17 Then
          strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} <> 'N'"
          .Formulas(1) = "Titulo='LISTADO DE GENERAL DE AFILIADOS'"
       End If
       .Formulas(2) = "SubTitulo='FECHA INGRESO DEL  " & Format(dtpDe, "DD/MM/YYYY") & "  AL  " & Format(dtpHasta, "DD/MM/YYYY") & "'"
       
       If Len(strSQL) = 0 Then
           strSQL = strSQL & "{SOCIOS.FECHAINGRESO} in DateTime"
       Else
           strSQL = strSQL & " And {SOCIOS.FECHAINGRESO} in DateTime"
       End If
       
       strSQL = strSQL & "(" & Year(dtpDe) & ","
       strSQL = strSQL & Month(dtpDe) & ","
       strSQL = strSQL & Day(dtpDe) & ") to DateTime "
       strSQL = strSQL & "(" & Year(dtpHasta) & ","
       strSQL = strSQL & Month(dtpHasta) & ","
       strSQL = strSQL & Day(dtpHasta) & ")"

      .SelectionFormula = strSQL
     
     Case 5
      strRuta = App.Path + "\Reportes\Afi"
      recUnidades.Source = "Select Count(*) as Registros From Socios Where EstadoActual='S'"
      recUnidades.ActiveConnection = glogon.Conection
      recUnidades.CursorLocation = adUseServer
      recUnidades.CursorType = adOpenStatic
      recUnidades.Open
     
      lngContador = recUnidades!registros
      recUnidades.Close
    
      .Formulas(0) = "Socios=" & lngContador
      .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = strRuta + "SociosProvincia.rpt"
     
     Case 6
      strRuta = App.Path + "\Reportes\Afi"
      recUnidades.Source = "Select Count(*) as Registros From Socios Where EstadoActual='A' or EstadoActual='P' "
      recUnidades.ActiveConnection = glogon.Conection
      recUnidades.CursorLocation = adUseServer
      recUnidades.CursorType = adOpenStatic
      recUnidades.Open
    
      lngContador = recUnidades!registros
      recUnidades.Close
       
      .Formulas(0) = "Socios=" & lngContador
      .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = strRuta + "ExSociosProvincia.rpt"
      
     Case 7
      strRuta = App.Path + "\Reportes\Afi"
      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = strRuta + "SociosPorUnidad.rpt"
      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} = 'S'"
     
     Case 8
      strRuta = App.Path + "\Reportes\Afi"
      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = strRuta + "DetalleSociosPorUnidad.rpt"
     
     Case 9
      If Trim(txt) = "" Then
       MsgBox "Suministre El Numero de Boleta a Imprimir", vbExclamation, "Faltan Datos"
       txt.SetFocus
       Me.MousePointer = vbDefault
       Exit Sub
      Else
       strRuta = App.Path + "\Reportes\Afi"
       .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
       .ReportFileName = strRuta + "ReimpresionBoleta.rpt"
       .SelectionFormula = "{SOCIOS.CEDULA} ='" & Trim(txt) & "'"
      End If
      
     Case 10
      GLOBALES.gstrReporte = "Resumen"
      frmAF_PromotoresReportes.Show vbModal
     
     Case 11
      GLOBALES.gstrReporte = "Detalle"
      frmAF_PromotoresReportes.Show vbModal
     
     Case 12
      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = App.Path + "\Reportes\AfiListadoPromotores.rpt"
     
     Case 13
      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = App.Path + "\Reportes\AfiListaUnidadProgramatica.rpt"
     
     Case 14
      .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
      .ReportFileName = App.Path + "\Reportes\AfiListaUnidadTrabajo.rpt"
      
     Case 15
      If Trim(txt) = "" Then
       MsgBox "Suministre El Numero Unidad Programatica", vbExclamation, "Faltan Datos"
       txt.SetFocus
       Me.MousePointer = vbDefault
       Exit Sub
      Else
       .ReportFileName = App.Path + "\Reportes\AfiDetalleSociosporUnidad.rpt"
       .SelectionFormula = "{SOCIOS.ESTADOACTUAL} ='S' And {UPROGRAMATICA.CODIGO}='" & Trim(txt) & "'"
       If chkTodas.Value = vbUnchecked Then
           .SelectionFormula = .SelectionFormula & " AND {SOCIOS.FECHAINGRESO} IN DATE(" & Format(dtpDe.Value, "yyyy,mm,dd") _
                            & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
       End If
      End If
      
     Case 16
      .ReportFileName = App.Path + "\Reportes\AfiDetalleSociosporUnidad.rpt"
      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} ='S'"
       If chkTodas.Value = vbUnchecked Then
           .SelectionFormula = .SelectionFormula & " AND {SOCIOS.FECHAINGRESO} IN DATE(" & Format(dtpDe.Value, "yyyy,mm,dd") _
                            & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
       End If
     
     Case 18
      .ReportFileName = App.Path + "\Reportes\AfiSociosPorUnidad.rpt"
      .SelectionFormula = "{SOCIOS.ESTADOACTUAL} ='S'"
       If chkTodas.Value = vbUnchecked Then
           .SelectionFormula = .SelectionFormula & " AND {SOCIOS.FECHAINGRESO} IN DATE(" & Format(dtpDe.Value, "yyyy,mm,dd") _
                            & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
       End If
     
   End Select

  .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
dtpDe = Format(fxFechaServidor, "dd/mm/yyyy")
dtpHasta = Format(fxFechaServidor, "dd/mm/yyyy")

End Sub

Private Sub opt_Click(Index As Integer)
vOPT = Index
txt = ""
Select Case Index
 Case 0
   lbl = "# Acta"
 Case 1 To 4, 15, 17
   lbl = "Unidad Programatica"
 Case 5 To 8, 10 To 14, 16
   lbl = ""
 Case 9
   lbl = "Cédula"
End Select
End Sub


