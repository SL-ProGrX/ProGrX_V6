VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAF_PlanillaEnvia 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notifica Planilla de Asociados"
   ClientHeight    =   8325
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.CheckBox chkResultados 
      Height          =   372
      Left            =   5520
      TabIndex        =   9
      Top             =   1800
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Mostrar Resultados?"
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
      Alignment       =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9720
      Top             =   720
   End
   Begin XtremeSuiteControls.PushButton btnArchivo 
      Height          =   732
      Left            =   7560
      TabIndex        =   1
      Top             =   1800
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Genera Archivo"
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
      Appearance      =   16
      Picture         =   "frmAF_PlanillaEnvia.frx":0000
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   108
      Left            =   0
      TabIndex        =   2
      Top             =   8220
      Width           =   11364
      _ExtentX        =   20055
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1080
      TabIndex        =   5
      Top             =   1320
      Width           =   6372
      _Version        =   1441793
      _ExtentX        =   11245
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
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   312
      Left            =   7560
      TabIndex        =   6
      Top             =   1320
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3625
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5292
      Left            =   120
      TabIndex        =   7
      Top             =   2760
      Width           =   11052
      _Version        =   524288
      _ExtentX        =   19495
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
      MaxCols         =   13
      SpreadDesigner  =   "frmAF_PlanillaEnvia.frx":0705
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   732
      Left            =   9720
      TabIndex        =   8
      Top             =   1800
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Excel"
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
      Appearance      =   16
      Picture         =   "frmAF_PlanillaEnvia.frx":0E06
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Empresa"
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
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   732
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   5412
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Crea Archivo para Rebajo de Aportes (Obrero - Patronal)"
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
      Height          =   600
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   7212
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13572
   End
End
Attribute VB_Name = "frmAF_PlanillaEnvia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mInstitucion As Long, mFechaProceso As Currency
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cmdActualiza_Click()


On Error GoTo vError

Me.MousePointer = vbHourglass

lbl.Alignment = 0

lbl.Caption = "Sincronizando Contratos con Operaciones de Retención [Espere!]"
lbl.Refresh

strSQL = "exec spFndSincronizaContratos"
Call ConectionExecute(strSQL)

Call Bitacora("Aplica", "Sincronización de Fondos con Retenciones")

lbl.Caption = "Actualización Concluida Satisfactoriamente...."
PrgBar.Value = 1
PrgBar.Max = 1000000

Me.MousePointer = vbDefault

Exit Sub


vError:
  lbl.Caption = "Error...."
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGeneraArchivoF15_PJ()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'***********************************************************
'* Formato Poder Judicial -> Variacion del formato INTEGRA *
'***********************************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
PrgBar.Min = 1

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & mInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & cbo.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & cbo.Text & "\" & Mid(mFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & cbo.Text & "\" & Mid(mFechaProceso, 1, 4)
On Error GoTo vError


'Estandar
'vArchivo = "E-" & IIf((vCodInstitucion = ""), Format(mInstitucion, "00"), vCodInstitucion) _
'         & "_" & Format(mFechaProceso, "####-##") & " [" & Format(vFecha, "ddmmyyyy") & "-F15].txt"


vArchivo = "E-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-01.CIF"


vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


lbl = "Creando archivo a enviar"
DoEvents


'*************************************************************
' Nota: En el nuevo procedimiento de planillas de Mecaniza
' se borran las deducciones de las personas en la aplicacion
' de la nueva planilla, por lo tanto se tienen que enviar todas
' las variaciones e inclusiones, pero no las exclusiones ya que
' estas son eliminadas por si solas.
'*************************************************************

strSQL = "SELECT S.CEDULA, S.NOMBRE, 1 as 'Sector', '463020' as 'COD_DEDUCCION', 3.5 AS 'MONTO_ACTUAL' " _
    & "    FROM SOCIOS S" _
    & "    WHERE S.ESTADOACTUAL = 'S' AND S.COD_INSTITUCION = " & mInstitucion _
    & "     AND S.CEDULA NOT IN(SELECT CEDULA " _
    & "                         From AFI_CR_RENUNCIAS" _
    & "                         WHERE ESTADO = 'T' AND DATEDIFF(D, REGISTRO_FECHA, dbo.mygetdate() ) <= 30" _
    & "                        )" _
    & "    order by S.CEDULA"
       
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1


Do While Not rs.EOF
 
 'Campo 01: Cedula de 10 char, 2-4-4
 'Campo 02: Codigo de Deduccion Asignada
 'Campo 03: Monto o Valor (10 espacios)
 'Campo 04: Tipo de Aplicacion (Defecto 0 = Mensual,  parte la cuota en dos quincenas)
 
' vCadena = Mid(Format(Trim(rs!Cedula), "0000000000"), 1, 10) & vbTab
 
 vCadena = Format(Trim(rs!Cedula), "0000000000") & vbTab
 vCadena = vCadena & rs!Cod_Deduccion & vbTab & Format(rs!Monto_Actual, "############0.00") & vbTab
 vCadena = vCadena & "1"
 
 
 
 
 Print #fnFile, vCadena
 
 If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
 lbl.Caption = "Creando Archivo Reg. # " & PrgBar.Value & " de " & PrgBar.Max & "     " & Format((PrgBar.Value / PrgBar.Max) * 100, "##0") & "%"
 DoEvents
 rs.MoveNext
Loop
rs.Close

Close #fnFile
  
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 

 
lbl.Caption = "Estado..."
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbGeneraArchivoPG()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vRuta As String, vTempo As String, i As Integer
Dim fnFile, iRespuesta As Integer, vCadena As String
Dim vFile As String, vArchivo As String, vFecha As Date

Dim vTipoAporte As String, vTipoCredito As String, vPorcAhorro As Currency, vPorcAporte As Currency
Dim vMovimiento As String, vCodInstitucion As String

'********************************************
'* Formato P & G                            *
'********************************************

On Error GoTo vError

fnFile = FreeFile
vFecha = fxFechaServidor

vArchivo = ""
PrgBar.Min = 1

vGrid.MaxRows = 0

Me.MousePointer = vbHourglass

strSQL = "select planilla,codigo_aportes_env,codigo_creditos_env,porc_ahorro,codigo_inst_deduc" _
       & ",IncInclusiones,IncExclusiones,IncModificaciones,IncMantienen,porc_aporte" _
       & " from instituciones" _
       & " where cod_institucion = " & mInstitucion

Call OpenRecordSet(rs, strSQL)
  vCodInstitucion = Trim(rs!codigo_inst_deduc & "")
  vTipoAporte = Trim(rs!Codigo_Aportes_Env & "")
  vTipoCredito = Trim(rs!codigo_creditos_env & "")
  vPorcAhorro = rs!porc_ahorro
  vPorcAporte = rs!PORC_APORTE
  vMovimiento = "in('"
  If rs!IncInclusiones = 1 Then vMovimiento = vMovimiento & "I','"
  If rs!IncExclusiones = 1 Then vMovimiento = vMovimiento & "E','"
  If rs!IncModificaciones = 1 Then vMovimiento = vMovimiento & "C','"
  If rs!IncMantienen = 1 Then vMovimiento = vMovimiento & "M','"
  vMovimiento = vMovimiento & "P')"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\"
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & cbo.Text
MkDir SIFGlobal.DirectorioDeResultados & "\Planilla\" & cbo.Text & "\" & Mid(mFechaProceso, 1, 4)

vRuta = SIFGlobal.DirectorioDeResultados & "\Planilla\" & cbo.Text & "\" & Mid(mFechaProceso, 1, 4)
On Error GoTo vError



  
'------------------------------------------------------------------------------------------------------------
'           Formatos Nuevos:
'------------------------------------------------------------------------------------------------------------

lbl.Caption = "Creando archivo para las deducciones de Asociados"
DoEvents


vArchivo = "Asoc-" & vCodInstitucion & "-" & Year(vFecha) & Format(Month(vFecha), "00") _
         & Format(Day(vFecha), "00") & "-PG.csv"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Close 'Cierra todos los archivos abiertos
 Kill vTempo
End If


Open vTempo For Output As #fnFile  ' Create file name.


strSQL = "exec spPrm_Formato_PG_Soc " & mInstitucion & "," & mFechaProceso
Call OpenRecordSet(rs, strSQL)

vGrid.MaxRows = 0

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1

With vGrid

Do While Not rs.EOF
 
 vCadena = rs!Cadena
 
 Print #fnFile, vCadena
 
 If chkResultados.Value = xtpChecked Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    .col = 1
    .Text = rs!col_01
    .col = 2
    .Text = rs!col_02
    .col = 3
    .Text = rs!col_03
    .col = 4
    .Text = rs!col_04
    .col = 5
    .Text = rs!col_05
    .col = 6
    .Text = rs!col_06
    .col = 7
    .Text = rs!col_07
    .col = 8
    .Text = rs!col_08
    .col = 9
    .Text = rs!col_09
    .col = 10
    .Text = rs!col_10
    .col = 11
    .Text = rs!col_11
    
    .col = 12
    .Text = rs!Cedula
    .col = 13
    .Text = rs!Nombre
 End If
 
 If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
 lbl.Caption = "Creando Archivo Reg. # " & PrgBar.Value & " de " & PrgBar.Max & "     " & Format((PrgBar.Value / PrgBar.Max) * 100, "##0") & "%"
 rs.MoveNext
Loop
rs.Close

End With

Close #fnFile
   
Me.MousePointer = vbDefault

MsgBox "El sistema genero el siguiente archivo : " & vTempo, vbInformation
 
lbl.Caption = ""
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub





Private Sub btnArchivo_Click()

mInstitucion = cbo.ItemData(cbo.ListIndex)
mFechaProceso = cboProceso.Text


strSQL = "select PLANILLA_ENVIO From INSTITUCIONES  Where COD_INSTITUCION = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, strSQL)

Select Case rs!PLANILLA_ENVIO
    Case "29"
        Call sbGeneraArchivoPG
    Case "15"
        Call sbGeneraArchivoF15_PJ
    Case Else
        Call sbGeneraArchivoPG
End Select

End Sub

Private Sub btnExportar_Click()
Call sbExportar
End Sub

Private Sub Form_Load()

vModulo = 1 'Clientes

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture
'
'lbl.Caption = "Actualiza Operaciones de Retención que se encuentran al cobro en el " _
'            & "sistema de cuentas corrientes, actualizando la cuota al cobro o en su " _
'            & "efecto la cancelación de la misma producto de una liquidación..."
End Sub


Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 13
    vHeaders.Headers(1) = "Col_1"
    vHeaders.Headers(2) = "Col_2"
    vHeaders.Headers(3) = "Col_3"
    vHeaders.Headers(4) = "Col_4"
    vHeaders.Headers(5) = "Col_5"
    vHeaders.Headers(6) = "Col_6"
    vHeaders.Headers(7) = "Col_7"
    vHeaders.Headers(8) = "Col_8"
    vHeaders.Headers(9) = "Col_9"
    vHeaders.Headers(10) = "Col_10"
    vHeaders.Headers(11) = "Col_11"
    vHeaders.Headers(12) = "Id"
    vHeaders.Headers(13) = "Nombre"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Pla_Asoc_" & cbo.ItemData(cbo.ListIndex) & "_" & cboProceso.Text)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String, i As Integer


strSQL = "select cod_institucion as 'Idx', rtrim(descripcion) as 'itmX' " _
       & " from instituciones Where activa = 1  order by descripcion "
Call sbCbo_Llena_New(cbo, strSQL, False, True)


mFechaProceso = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
mFechaProceso = fxFechaProcesoAnterior(mFechaProceso)


cboProceso.AddItem CStr(mFechaProceso)

For i = 1 To 6
  mFechaProceso = fxFechaProcesoSiguiente(mFechaProceso)
  cboProceso.AddItem CStr(mFechaProceso)
Next i
cboProceso.Text = CStr(GLOBALES.glngFechaCR)


End Sub
