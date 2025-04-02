VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_GeneraGarantia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Generar Garantía (Letra Cambio / Pagaré)"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6030
   HelpContextID   =   3020
   Icon            =   "frmCR_GeneraGarantia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2172
      Left            =   360
      TabIndex        =   7
      Top             =   3000
      Width           =   5292
      _Version        =   1310722
      _ExtentX        =   9334
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Opciones de la Garantía"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkNombres 
         Height          =   252
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   4692
         _Version        =   1310722
         _ExtentX        =   8276
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Imprimir Nombres/Cédula (Pagaré Pre-Impreso)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboProvincia 
         Height          =   312
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   2172
         _Version        =   1310722
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.CheckBox chkCedula 
         Height          =   252
         Left            =   480
         TabIndex        =   13
         Top             =   1200
         Width           =   4692
         _Version        =   1310722
         _ExtentX        =   8276
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Utilizar Cédula Real en lugar de la ced. colilla"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkReemplazar 
         Height          =   252
         Left            =   480
         TabIndex        =   14
         Top             =   1560
         Width           =   4692
         _Version        =   1310722
         _ExtentX        =   8276
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Reemplazar Información del Pagaré Anterior"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lugar de Firma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   2
         Left            =   1560
         TabIndex        =   8
         Top             =   480
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.PushButton cmdImprimir 
      Height          =   732
      Left            =   480
      TabIndex        =   4
      Top             =   5400
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2984
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Letra de Cambio"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCR_GeneraGarantia.frx":000C
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1452
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   5292
      _Version        =   1310722
      _ExtentX        =   9334
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Rango de Operaciones: "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtDesde 
         Height          =   312
         Left            =   2160
         TabIndex        =   9
         Top             =   480
         Width           =   2292
         _Version        =   1310722
         _ExtentX        =   4043
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtHasta 
         Height          =   312
         Left            =   2160
         TabIndex        =   10
         Top             =   840
         Width           =   2292
         _Version        =   1310722
         _ExtentX        =   4043
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.PushButton cmdPagare 
      Height          =   732
      Left            =   2160
      TabIndex        =   5
      Top             =   5400
      Width           =   1692
      _Version        =   1310722
      _ExtentX        =   2984
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Pagaré"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCR_GeneraGarantia.frx":07C5
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton cmdPagarePreImpreso 
      Height          =   732
      Left            =   3840
      TabIndex        =   6
      Top             =   5400
      Width           =   1812
      _Version        =   1310722
      _ExtentX        =   3196
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Pagaré Pre-Impreso"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCR_GeneraGarantia.frx":0F7E
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnPagareEmail 
      Height          =   495
      Left            =   480
      TabIndex        =   15
      Top             =   6480
      Width           =   5175
      _Version        =   1310722
      _ExtentX        =   9128
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Pagaré Digital (Email)"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   14
      Picture         =   "frmCR_GeneraGarantia.frx":1737
      ImageAlignment  =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emisión de la Garantía"
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
      Height          =   372
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   4572
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCR_GeneraGarantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FuncionLetra(Monto As Currency, Optional pDivisa As String = "")
Dim strNum As String, strMonto As String
Dim strSQL As String, strDec As String, intI As Integer

strMonto = ""
strDec = "00"
strNum = CStr(Monto)

For intI = 1 To Len(strNum)
    If Mid(strNum, intI, 1) <> "." Then
       strMonto = strMonto & Mid(strNum, intI, 1)
    Else
       strDec = Mid(strNum, intI + 1, 2)
       Exit For
    End If
Next intI

strSQL = Conversion(strMonto)

strSQL = strSQL & " " & Trim(pDivisa) & " con " & strDec & "/100"

FuncionLetra = strSQL

End Function



Private Function fxFuncionLetraNoMoneda(Monto As Currency, Optional vPlazo As Boolean = False, Optional pDivisa As String = "")
Dim strNum As String, strMonto As String
Dim strSQL As String, strDec As String, intI As Integer

strMonto = ""
strDec = "00"
strNum = CStr(Monto)

For intI = 1 To Len(strNum)
    If Mid(strNum, intI, 1) <> "." Then
       strMonto = strMonto & Mid(strNum, intI, 1)
    Else
       strDec = Mid(strNum, intI + 1, 2)
       Exit For
    End If
Next intI

strSQL = Conversion(strMonto)
If Not vPlazo And strDec <> "00" Then strSQL = strSQL & " punto " & strDec

fxFuncionLetraNoMoneda = strSQL

End Function

Function FunMes() As String

Select Case Format(Month(fxFechaServidor), "00")
  Case "01"
       FunMes = "Enero"
  Case "02"
       FunMes = "Febrero"
  Case "03"
       FunMes = "Marzo"
  Case "04"
       FunMes = "Abril"
  Case "05"
       FunMes = "Mayo"
  Case "06"
       FunMes = "Junio"
  Case "07"
       FunMes = "Julio"
  Case "08"
       FunMes = "Agosto"
  Case "09"
       FunMes = "Setiembre"
  Case "10"
       FunMes = "Octubre"
  Case "11"
       FunMes = "Noviembre"
  Case "12"
       FunMes = "Diciembre"
End Select

End Function

Private Function fxVerificaDatosX() As Boolean
Dim vMensaje As String

vMensaje = ""

If Trim(txtDesde) = "" Then
   vMensaje = "Falta El #OP.Inicial"
ElseIf Trim(txtHasta) = "" Then
   vMensaje = vMensaje & vbCrLf & " - Falta El #OP.Final"
ElseIf CLng(txtDesde) > CLng(txtHasta) Then
   vMensaje = vMensaje & vbCrLf & " - El #OP.Inicial No Puede Ser Mayor al #OP.Final"
End If

If Len(vMensaje) > 0 Then
    MsgBox vMensaje, vbExclamation, "No Se Puede Generar"
    fxVerificaDatosX = False
Else
    fxVerificaDatosX = True
End If

End Function

Private Sub btnPagareEmail_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As Long
Dim vCedula As String, vSecciones As String

On Error GoTo vError


vOperacion = InputBox("Digite el Número de Operación : ", "Emisión de Pagaré", Operacion.Operacion)

strSQL = "exec spCrd_Pagare_Email " & vOperacion & "," & chkReemplazar.Value _
       & "," & chkCedula.Value & ",'" & cboProvincia.Text & "'"

Call OpenRecordSet(rs, strSQL)

MsgBox "Pagaré Digital enviado al correo: " & rs!Email

rs.Close


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cmdImprimir_Click()
Dim rs As New ADODB.Recordset
Dim strSQL As String, x As New clsImpresoras

On Error GoTo vError

'Imprime Letra de Cambio

If Not fxVerificaDatosX Then Exit Sub

 Me.MousePointer = vbHourglass
 
 strSQL = "Select id_solicitud,codigo,cedula,montoapr,montosol,estadosol" _
        & " From Reg_Creditos Where Id_Solicitud between " & CLng(txtDesde) _
        & " And " & CLng(txtHasta) & " And EstadoSol in ('R','P','A')"
 Call OpenRecordSet(rs, strSQL)
      
      
x.TipoImpresora = pagare
x.Reset
      
      
 Do While Not rs.EOF
    With frmContenedor.Crt
       .Reset
       .Destination = crptToPrinter
      
       .PrinterDriver = x.Controlador
       .PrinterName = x.Nombre
       .PrinterPort = x.Puerto

       .Connect = glogon.ConectRPT

       .ReportFileName = SIFGlobal.fxPathReportes("Credito_LetraCambio.rpt")
       .Formulas(0) = "LugarFecha='San José " & Format(Day(fxFechaServidor), "00") & " De " & FunMes & " De " & Year(fxFechaServidor) & "'"
       Select Case rs!estadosol
          Case "R", "P"
           .Formulas(1) = "MontoLetras='" & FuncionLetra(IIf(IsNull(rs!montosol), 0, rs!montosol)) & "'"
           .Formulas(2) = "Monto='¢ " & Format(rs!montosol, "Standard") & "'"
          Case Else
           .Formulas(1) = "MontoLetras='" & FuncionLetra(IIf(IsNull(rs!montoapr), 0, rs!montoapr)) & "'"
           .Formulas(2) = "Monto='¢ " & Format(rs!montoapr, "Standard") & "'"
       End Select
       
       .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & rs!id_solicitud
       
       .Destination = crptToPrinter
'       lblTitulo = "Imprimiendo Op. # " & rs!id_solicitud
       
       .PrintReport
    End With
    rs.MoveNext
 Loop
 rs.Close
 
' lblTitulo = "RANGO DE # OPERACION"
 
 Me.MousePointer = vbDefault
 MsgBox "# Garantias Impresas " & rs.RecordCount, vbExclamation, "Impresión Finalizada"
  
  
Exit Sub

vError:
   Me.MousePointer = vbDefault
'   lblTitulo = "RANGO DE # OPERACION"
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Function fxPagareEstadoCivil(vEstadoCivil As Variant, vSexo As String) As Variant

Select Case UCase(vEstadoCivil)
Case "S"
  fxPagareEstadoCivil = IIf((vSexo = "M"), "Soltero", "Soltera")
Case "C"
  fxPagareEstadoCivil = IIf((vSexo = "M"), "Casado", "Casada")
Case "D"
  fxPagareEstadoCivil = IIf((vSexo = "M"), "Divorciado", "Divorciada")
Case "V"
  fxPagareEstadoCivil = IIf((vSexo = "M"), "Viudo", "Viuda")
Case "U"
  fxPagareEstadoCivil = "Union libre"
Case "O"
  fxPagareEstadoCivil = "Otro"
Case Else
  fxPagareEstadoCivil = vEstadoCivil
End Select
End Function


Private Function fxPagareCalidades(pCedula As String, Optional pCedReal As Integer = 0) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCadenaX As String

strSQL = "select S.*,rtrim(Prov.Descripcion) as ProvDesc, rtrim(Cant.Descripcion) as CantonDesc, rtrim(Dist.Descripcion) as DistDesc" _
       & " from socios S " _
       & " left join Provincias Prov on S.Provincia = Prov.Provincia" _
       & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
       & " left join Distritos Dist on S.Provincia = Dist.Provincia and S.Canton = Dist.Canton and S.distrito = Dist.distrito" _
       & " where S.cedula = '" & pCedula & "'"
Call OpenRecordSet(rs, strSQL)

vCadenaX = ""
pCedula = Trim(pCedula)

If pCedReal = 1 Then
 If Not IsNull(rs!cedular) Then
   If Len(Trim(rs!cedular)) > 0 Then
      pCedula = Trim(rs!cedular)
   End If
 End If
End If


If Not rs.EOF And Not rs.BOF Then
   vCadenaX = Trim(rs!Nombre) & ", mayor, " & fxPagareEstadoCivil(rs!EstadoCivil, rs!sexo) _
            & " con cédula de identidad número "
   For i = 1 To Len(pCedula)
     Select Case i
       Case 2, 6
         vCadenaX = vCadenaX & " - " & Trim(fxFuncionLetraNoMoneda(Mid(pCedula, i, 1), True))
       Case Else
         vCadenaX = vCadenaX & " " & Trim(fxFuncionLetraNoMoneda(Mid(pCedula, i, 1), True))
     End Select
   Next i
   
   vCadenaX = vCadenaX & ", con residencia en " & rs!ProvDesc & ", " & rs!CantonDesc & "  distrito " & rs!DistDesc & ", " & Trim(rs!Direccion)
   
   
Else
   vCadenaX = "** calidades **"
End If
rs.Close

fxPagareCalidades = vCadenaX

End Function

Function fxTermina(vInicio As Long, vPlazo As Integer) As Long
Dim vFecha As Date


vFecha = Format(vInicio, "####/##") & "/01"

vFecha = DateAdd("m", vPlazo - 1, vFecha)

fxTermina = Year(vFecha) & Format(Month(vFecha), "00")

End Function


Function fxMes(iMes As Byte) As String
Dim xMes As String

xMes = ""

Select Case iMes
  Case 1
     xMes = "Enero"
  Case 2
     xMes = "Febrero"
  Case 3
     xMes = "Marzo"
  Case 4
     xMes = "Abril"
  Case 5
     xMes = "Mayo"
  Case 6
     xMes = "Junio"
  Case 7
     xMes = "Julio"
  Case 8
     xMes = "Agosto"
  Case 9
     xMes = "Septiembre"
  Case 10
     xMes = "Octubre"
  Case 11
     xMes = "Noviembre"
  Case 12
     xMes = "Diciembre"
End Select

fxMes = xMes

End Function

Private Sub cmdPagare_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As Long
Dim vCedula As String, vSecciones As String

On Error GoTo vError

'Inicializa
vSecciones = "0" 'El Pagaré utiliza las secciones

vOperacion = InputBox("Digite el Número de Operación : ", "Emisión de Pagaré", Operacion.Operacion)

strSQL = "exec spCrd_Operacion_Pagare_Registra " & vOperacion & "," & chkReemplazar.Value _
       & "," & chkCedula.Value & ",'" & cboProvincia.Text & "',0"
Call OpenRecordSet(rs, strSQL)
  vCedula = Trim(rs!Cedula)
  vSecciones = Trim(rs!Secciones)
rs.Close

'Imprimir el Reporte
With frmContenedor.Crt
  .Reset
  .WindowShowExportBtn = True
  .WindowShowPrintSetupBtn = True

  .Connect = glogon.ConectRPT
   
  .WindowState = crptMaximized
  .WindowTitle = "Emisión del Pagaré"
  .ReportFileName = SIFGlobal.fxPathReportes("Credito_Pagare.rpt")
  .Formulas(0) = "fxCedula = '" & vCedula & "'"
  .Formulas(1) = "fxSecciones = '" & vSecciones & "'"
  .Formulas(2) = "fxBarras = '*" & vOperacion & "*'"
   
  .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOperacion
  .PrintReport
End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxMoraTasa(vCodigo As String, vMonto As Currency) As Double
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select intm_soc from rangos where codigo = '" & vCodigo & "' and " & vMonto _
       & " between de and hasta"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
 fxMoraTasa = rs!intm_soc
Else
 fxMoraTasa = 0
End If
rs.Close

End Function


Private Sub cmdPagarePreImpreso_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As Long, vFiadores As Boolean
Dim rsTmp As New ADODB.Recordset, i As Byte
Dim vFormula01 As String, vFormula02 As String
Dim vFormula03 As String, vFormula04 As String, vFormula05 As String
Dim vMontoLetras As String, vPrometo As String, vMora As String

On Error GoTo vError

'Inicializa
vFiadores = False

vOperacion = InputBox("Digite el Número de Operación : ", "Emisión de Pagaré", Operacion.Operacion)

'Verificar el Estado (Aprobada y Formalizada) y la Garantía
'Si es fiduciaria procesar datos de fiadores(Calidades) tambien.
strSQL = "select codigo,cedula,estadosol,garantia,montoapr,plazo,int from reg_creditos where id_solicitud = " & vOperacion
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  If rs!estadosol = "A" Or rs!estadosol = "F" Then
     If rs!Garantia = "F" Then vFiadores = True
  Else
    MsgBox "Esta operación no se encuentra Aprobada o Formalizada, verifique..." & vOperacion, vbExclamation
  End If
Else
  MsgBox "No Existe el número de operación indicado, verifique..." & vOperacion, vbExclamation
  Exit Sub
End If

vMora = fxMoraTasa(rs!Codigo, rs!montoapr)
If vMora = 0 Then vMora = rs!Int

vFormula01 = Space(30) & fxPagareCalidades(rs!Cedula)
vFormula02 = Space(20) & fxFuncionLetraNoMoneda(rs!Plazo, True) & " cuota(s) mensual(es) iguales"
vMontoLetras = Space(75) & fxMontoLetrasX(rs!montoapr)
vPrometo = IIf(vFiadores, "            emos", "            o")
        
 '       & " pagar incondicionalmente a la orden de la Asociación Solidarista de Empleados" _
        & " de la Caja Costarricense de Seguro Social, A.S.E.C.C.S.S., cédula jurídica número" _
        & " tres-cero cero dos-cero sesenta y seis cero treinta y uno, con domicilio en San José," _
        & " trecientos cincuenta metros sur del antiguo Banco Anglo, la suma de " _
        & FuncionLetra(rs!montoapr) & " (¢ " & Format(rs!montoapr, "Standard") & ") pagadera en " _
        & fxFuncionLetraNoMoneda(rs!Plazo, True) & " cuota(s) mensual(es) iguales en las Oficinas" _
        & " de A.S.E.C.C.S.S., ubicada en Avenida 8 y 10 calle 3 la cual devenga intereses corrientes del " _
        & fxFuncionLetraNoMoneda(rs!Int) & " porciento anual sobre saldos revisable y ajustable;" _
        & " y moratorios del treinta porciento."

If vFiadores Then
  strSQL = "select cedulaf from fiadores where estado = 'A' and id_solicitud = " & vOperacion
  Call OpenRecordSet(rsTmp, strSQL, 0)
  vFormula03 = ""
  vFormula04 = ""
  vFormula05 = ""
  i = 0
  Do While Not rsTmp.EOF
   i = i + 1
   Select Case i
    Case 1
       vFormula03 = vFormula03 & fxPagareCalidades(rsTmp!cedulaf)
    Case 2
       vFormula04 = ";" & fxPagareCalidades(rsTmp!cedulaf)
    Case 3
       vFormula05 = ";" & fxPagareCalidades(rsTmp!cedulaf)
   End Select
   rsTmp.MoveNext
  Loop
  rsTmp.Close
End If

rs.Close

'Imprimir el Reporte

'Borra Registro Anterior y lo vuelve a Guardar
strSQL = "delete tmp_pagare where id_solicitud = " & vOperacion
Call ConectionExecute(strSQL)

strSQL = "insert tmp_pagare(id_solicitud,pagare,formula01) values(" & vOperacion _
       & ",'" & vFormula03 & " " & vFormula04 & " " & vFormula05 & "','" & vFormula01 & "')"
Call ConectionExecute(strSQL)

With frmContenedor.Crt
  .Reset
  .WindowShowExportBtn = True
  .WindowShowPrintSetupBtn = True

  .WindowState = crptMaximized
  .WindowTitle = "Emisión del Pagaré"
  
  .Connect = glogon.ConectRPT
  
  .ReportFileName = SIFGlobal.fxPathReportes("PagarePreImpreso.rpt")
  .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOperacion

 
'  .Formulas(0) = "fxformula01 = '" & vFormula01 & "'"
'  .Formulas(2) = "fxformula03 = '" & vFormula03 & "'"
'  .Formulas(3) = "fxformula04 = '" & vFormula04 & "'"
'  .Formulas(4) = "fxformula05 = '" & vFormula05 & "'"
  
  .Formulas(0) = "fxformula02 = '" & vFormula02 & "'"
  .Formulas(1) = "MontoLetras = '" & vMontoLetras & "'"
  .Formulas(2) = "fxPrometo = '" & vPrometo & "'"
  .Formulas(3) = "fxMora = '" & vMora & "'"
  .Formulas(4) = "fxFiadores = '" & IIf((chkNombres.Value = vbChecked), "S", "N") & "'"
  
  .PrintReport
End With

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

cboProvincia.Clear
cboProvincia.AddItem "San José"
cboProvincia.AddItem "Alajuela"
cboProvincia.AddItem "Cartago"
cboProvincia.AddItem "Heredia"
cboProvincia.AddItem "Guanacaste"
cboProvincia.AddItem "Puntarenas"
cboProvincia.AddItem "Limón"


cboProvincia.Text = "San José"

End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
KeyAscii = Validacion(KeyAscii)

If KeyAscii = vbKeyReturn Then
   txtHasta.SetFocus
End If

End Sub


Private Sub txtHasta_KeyPress(KeyAscii As Integer)
KeyAscii = Validacion(KeyAscii)

If KeyAscii = vbKeyReturn Then
   cmdImprimir.SetFocus
End If
End Sub


