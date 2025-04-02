VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.shortcutbar.v19.2.0.ocx"
Begin VB.Form frmFNDRastreoMovDoc 
   Caption         =   "Rasteo Movimientos Documentos y Asientos"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10380
   HelpContextID   =   7002
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10380
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3732
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   8652
      _Version        =   1245186
      _ExtentX        =   15261
      _ExtentY        =   6583
      _StockProps     =   77
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   5424
      Width           =   10380
      _ExtentX        =   18309
      _ExtentY        =   487
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   1452
      _Version        =   1245186
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDRastreoMovDoc.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   1452
      _Version        =   1245186
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   1452
      _Version        =   1245186
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboResultados 
      Height          =   312
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   1452
      _Version        =   1245186
      _ExtentX        =   2561
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.PushButton cmdArchivo 
      Height          =   492
      Left            =   7800
      TabIndex        =   7
      Top             =   240
      Width           =   1452
      _Version        =   1245186
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmFNDRastreoMovDoc.frx":0A1E
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   372
      Left            =   9720
      TabIndex        =   8
      Top             =   360
      Width           =   732
      _Version        =   1245186
      _ExtentX        =   1291
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "1000"
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Líneas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   3
      Left            =   9720
      TabIndex        =   12
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   2
      Left            =   3960
      TabIndex        =   11
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   4
      Left            =   1560
      TabIndex        =   10
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Top             =   480
      Width           =   1212
   End
   Begin XtremeShortcutBar.ShortcutCaption lblEstado 
      Height          =   492
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   14412
      _Version        =   1245186
      _ExtentX        =   25421
      _ExtentY        =   868
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmFNDRastreoMovDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type vFx
  Cedula As String
  Contrato As Long
  Plan As String
  Operadora As Long
  Movimiento As String
End Type
Dim vDatosCon As vFx


Function fxNombre(strCedula As String) As String
Dim rsX As New ADODB.Recordset

rsX.Open "select nombre from socios where cedula = '" & strCedula & "'", glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxNombre = ""
Else
 fxNombre = IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close
End Function

Function fxCedulayNombre(vPlan As String, vOperadora As Long, vContrato As Long) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select S.cedula,S.nombre" _
       & " from fnd_contratos C inner join Socios S on C.cedula = S.cedula" _
       & " where C.cod_contrato = " & vContrato & " and C.cod_operadora = " _
       & vOperadora & " and C.cod_plan = '" & vPlan & "'"

Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCedulayNombre = ""
Else
 fxCedulayNombre = Trim(IIf(IsNull(rsX!Cedula), "", rsX!Cedula)) & " - " & IIf(IsNull(rsX!Nombre), "", rsX!Nombre)
End If
rsX.Close
End Function

Private Sub sbResumen()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotales(2) As Currency
Dim lngLineas As Long, vMascara As String

Me.MousePointer = vbHourglass
lsw.ListItems.Clear

vMascara = GLOBALES.gstrMascara

lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"


strSQL = "select Count(*) as Total,isnull(C.descripcion,'--> No Existe!') as 'Descripcion',A.fnd_cuenta,A.fnd_debehaber,sum(fnd_monto) as Movimiento" _
       & " from fnd_documentos D inner join fnd_asientos A on D.tipo = A.tipo" _
       & " and D.id_documento = A.id_documento and D.cod_operadora = A.cod_operadora" _
       & " left join Cntx_Cuentas C On A.fnd_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
       & " where D.fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' group by A.fnd_cuenta,A.fnd_debehaber,C.descripcion"

Call OpenRecordSet(rs, strSQL)
prg.Max = rs.RecordCount + 1
prg.Value = 1
lngLineas = 1
    
curTotales(1) = 0
curTotales(2) = 0

Do While Not rs.EOF
 If lngLineas > CLng(txtLineas) Then
   lngLineas = 1
   lsw.ListItems.Clear
 End If
     
Set itmX = lsw.ListItems.Add(, , Format(rs!fnd_cuenta, vMascara))
    itmX.SubItems(1) = rs!Descripcion
  If rs!fnd_DEBEHABER = "D" Then
     itmX.SubItems(2) = Format(rs!Movimiento, "Standard")
     itmX.SubItems(3) = "0.00"
     curTotales(1) = curTotales(1) + rs!Movimiento
  Else
     itmX.SubItems(2) = "0.00"
     itmX.SubItems(3) = Format(rs!Movimiento, "Standard")
     curTotales(2) = curTotales(2) + rs!Movimiento
  End If
  itmX.SubItems(4) = "Cola Doc."

 prg.Value = prg.Value + 1
 lblEstado.Caption = "Registros Evaluados : " & Format(rs!Total, "###,###,###,##0") _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"

 DoEvents
 
 rs.MoveNext
 lngLineas = lngLineas + 1
 
Loop
rs.Close


'Procesar la Cola de Asientos
lblEstado.Caption = vbCrLf & "****- Cargando Información Complementaria (Espere) -****"

strSQL = "select Count(*) as Total,isnull(C.descripcion,'--> No Existe!') as 'Descripcion',A.fnd_cuenta,A.fnd_debehaber,sum(A.fnd_monto) as Movimiento" _
       & " from fnd_asientos_cola A left join Cntx_Cuentas C On A.fnd_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
       & " where A.fnd_fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' group by A.fnd_cuenta,A.fnd_debehaber,C.descripcion"

Call OpenRecordSet(rs, strSQL)
prg.Max = rs.RecordCount + 1
prg.Value = 1
lngLineas = 1
    
Do While Not rs.EOF
 If lngLineas > CLng(txtLineas) Then
   lngLineas = 1
   lsw.ListItems.Clear
 End If
     
Set itmX = lsw.ListItems.Add(, , Format(rs!fnd_cuenta, vMascara))
    itmX.SubItems(1) = rs!Descripcion
  If rs!fnd_DEBEHABER = "D" Then
     itmX.SubItems(2) = Format(rs!Movimiento, "Standard")
     itmX.SubItems(3) = "0.00"
     curTotales(1) = curTotales(1) + rs!Movimiento
  Else
     itmX.SubItems(2) = "0.00"
     itmX.SubItems(3) = Format(rs!Movimiento, "Standard")
     curTotales(2) = curTotales(2) + rs!Movimiento
  End If
  itmX.SubItems(4) = "Cola Asientos"

 prg.Value = prg.Value + 1
 lblEstado = "Registros Evaluados : " & Format(rs!Total, "###,###,###,##0") _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents

 rs.MoveNext
 lngLineas = lngLineas + 1

Loop
rs.Close


'TOTALES
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(2) = "---------------------"
    itmX.SubItems(3) = "---------------------"
          
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(2) = Format(curTotales(1), "Standard")
    itmX.SubItems(3) = Format(curTotales(2), "Standard")
  
  Set itmX = lsw.ListItems.Add(, , "T")

Me.MousePointer = vbDefault


End Sub


Private Sub sbDetalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotales(2) As Currency
Dim lngLineas As Long, vMascara As String
Dim lngRegistros As Long, lngEvaluados As Long

Me.MousePointer = vbHourglass
lsw.ListItems.Clear


lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"
DoEvents

vMascara = "###-##-###-##"

strSQL = "select O.descripcion,D.concepto,D.cliente,D.usuario,D.fecha,A.*" _
       & " from fnd_documentos D inner join fnd_asientos A on D.tipo = A.tipo and D.id_documento = A.id_documento and D.cod_operadora = A.cod_operadora" _
       & " inner join fnd_operadoras O on O.cod_operadora = D.cod_operadora" _
       & " where D.fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' order by D.fecha"

Call OpenRecordSet(rs, strSQL)
lngRegistros = rs.RecordCount
lngEvaluados = 0
prg.Max = rs.RecordCount + 1
prg.Value = 1
lngLineas = 1
    
curTotales(1) = 0
curTotales(2) = 0

Do While Not rs.EOF
 lngEvaluados = lngEvaluados + 1
 If lngLineas > CLng(txtLineas) Then
   lngLineas = 1
   lsw.ListItems.Clear
 End If
     

Set itmX = lsw.ListItems.Add(, , Format(rs!fecha, "yyyy/mm/dd"))
    itmX.SubItems(1) = rs!Tipo
    itmX.SubItems(2) = rs!id_documento
    itmX.SubItems(3) = Format(rs!fnd_cuenta, vMascara)
    
  If rs!fnd_DEBEHABER = "D" Then
     itmX.SubItems(4) = Format(rs!fnd_monto, "Standard")
     itmX.SubItems(5) = "0.00"
     curTotales(1) = curTotales(1) + rs!fnd_monto
  Else
     itmX.SubItems(4) = "0.00"
     itmX.SubItems(5) = Format(rs!fnd_monto, "Standard")
     curTotales(2) = curTotales(2) + rs!fnd_monto
  End If
   
   itmX.SubItems(6) = rs!CONCEPTO & ""
   itmX.SubItems(7) = rs!Cliente & ""
   itmX.SubItems(8) = rs!Descripcion & ""
   itmX.SubItems(9) = rs!Usuario & ""
   itmX.SubItems(10) = ""
   

 prg.Value = prg.Value + 1
 lblEstado = "PASO 1/3 - Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
           & Format(lngRegistros, "###,###,###,##0") & vbCrLf _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents

 rs.MoveNext
 lngLineas = lngLineas + 1

Loop
rs.Close


'*******************************************************************************************
'FASE II = BUSCA DEDUCCIONES POR PLANILLA, conforme a los planes en el Sistema

lblEstado.Caption = vbCrLf & "****- Cargando Información Complementaria (Espere) -****"
DoEvents

strSQL = "Select P.codigo,P.id_solicitud,P.amortiza,P.opex,P.fecha,P.fecha_proceso," _
       & "S.Nombre,C.CTANAMORT,C.CTAOAMORT,P.cedula,S.nombre,Cnt.cod_plan " _
       & " from prm_creditos P Left Join Socios S on P.cedula = S.cedula" _
       & " left join catalogo C on P.codigo = C.codigo" _
       & " inner join fnd_planes F on P.codigo = F.codigo_ase" _
       & " left join fnd_contratos Cnt on P.id_solicitud = Cnt.Operacion" _
       & " where P.IND_PASO = 1 AND P.amortiza > 0 and P.fecha between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    lngRegistros = rs.RecordCount
    lngEvaluados = 0
    prg.Max = rs.RecordCount + 1
    prg.Value = 1
End If

Do While Not rs.EOF
 
    lngEvaluados = lngEvaluados + 1
    If lngLineas > CLng(txtLineas) Then
      lngLineas = 1
      lsw.ListItems.Clear
    End If
 

Set itmX = lsw.ListItems.Add(, , Format(rs!fecha, "yyyy/mm/dd"))
    itmX.SubItems(1) = "PRM"
    itmX.SubItems(2) = rs!Fecha_Proceso
    If rs!opex = 1 Then
       itmX.SubItems(3) = Format(IIf(IsNull(rs!CtaOamort), "", rs!CtaOamort), vMascara)
    Else
      itmX.SubItems(3) = Format(IIf(IsNull(rs!CtaNamort), "", rs!CtaNamort), vMascara)
    End If
    
     'Todos Son Creditos y Solo Se cargan las Amortizaciones
     ' se Excluyen el registro de intereses
     itmX.SubItems(4) = "0.00"
     itmX.SubItems(5) = Format(rs!Amortiza, "Standard")
     curTotales(2) = curTotales(2) + rs!Amortiza
   
   itmX.SubItems(6) = "Ded.Pla : " & rs!Fecha_Proceso
   itmX.SubItems(7) = Trim(rs!Cedula) & " - " & rs!Nombre & ""
   itmX.SubItems(8) = Trim(rs!Codigo) & " Op." & rs!id_solicitud & " Ex." & rs!opex
   itmX.SubItems(9) = "Ded.Pla"
   itmX.SubItems(10) = rs!cod_Plan & ""
   

   prg.Value = prg.Value + 1
   lblEstado = "PASO 2/3 - Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
             & Format(lngRegistros, "###,###,###,##0") & vbCrLf _
             & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
             & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
   DoEvents
    
   rs.MoveNext
   lngLineas = lngLineas + 1

Loop
rs.Close

'*******************************************************************************************

'Procesar la Cola de Asientos
lblEstado.Caption = vbCrLf & "****- Cargando Información Complementaria (Espere) -****"
DoEvents

' Se excluye el asiento del proceso mensual ya que este se carga en detalle
' en el paso anterior.

strSQL = "select A.*,isnull(C.descripcion,'--> No Existe!') as 'Descripcion', rtrim(Cnt.Cedula) + ' - ' + isnull(S.nombre,'--> No Existe!') as 'Nombre'" _
       & " from fnd_asientos_cola A left join CntX_Cuentas C on A.fnd_cuenta = C.cod_cuenta and C.cod_Contabilidad = " & GLOBALES.gEnlace _
       & " inner join Fnd_Contratos Cnt on A.cod_operadora = Cnt.Cod_Operadora and A.cod_Plan = Cnt.Cod_Plan and A.cod_Contrato = Cnt.Cod_Contrato" _
       & " left join Socios S on Cnt.Cedula = S.cedula" _
       & " where A.fnd_fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' and fnd_tipo not in('PRM')"

Call OpenRecordSet(rs, strSQL)
prg.Max = rs.RecordCount + 1
prg.Value = 1
lngLineas = 1
    
Do While Not rs.EOF
 If lngLineas > CLng(txtLineas) Then
   lngLineas = 1
   lsw.ListItems.Clear
 End If
     
    vDatosCon.Cedula = rs!Nombre
    vDatosCon.Contrato = rs!COD_CONTRATO
    vDatosCon.Operadora = rs!Cod_Operadora
    vDatosCon.Plan = rs!cod_Plan
     
    Select Case UCase(rs!fnd_tipo)
      Case "LI" 'LIQUIDACIONES
            vDatosCon.Movimiento = "LIQUIDACION"
      Case "RT" 'RETIROS
            vDatosCon.Movimiento = "RETIROS"
      Case "RE", "CR" 'APLICACION DE RENDIMIENTOS
            vDatosCon.Movimiento = "RENDIMIENTOS"
      Case "PM" 'PROCESOS MENSUALES N/A
            vDatosCon.Movimiento = "PROCESO MENSUAL"
      Case "PR"
            vDatosCon.Movimiento = "PROCESO PLANILLA"
      Case "RL"
            vDatosCon.Movimiento = "REV.LIQUIDACION"
      Case Else
            vDatosCon.Movimiento = "No Identificado!"
    End Select



Set itmX = lsw.ListItems.Add(, , Format(rs!fnd_Fecha, "yyyy/mm/dd"))
    itmX.SubItems(1) = rs!fnd_tipo
    itmX.SubItems(2) = "Cola de Asientos"
    itmX.SubItems(3) = Format(rs!fnd_cuenta, vMascara)
    
  If rs!fnd_DEBEHABER = "D" Then
     itmX.SubItems(4) = Format(rs!fnd_monto, "Standard")
     itmX.SubItems(5) = "0.00"
     curTotales(1) = curTotales(1) + rs!fnd_monto
  Else
     itmX.SubItems(4) = "0.00"
     itmX.SubItems(5) = Format(rs!fnd_monto, "Standard")
     curTotales(2) = curTotales(2) + rs!fnd_monto
  End If
   
   itmX.SubItems(6) = UCase(vDatosCon.Movimiento) & " - " & rs!fnd_caso
   itmX.SubItems(7) = Trim(vDatosCon.Cedula) 'Trae la cedula con el nombre
   itmX.SubItems(8) = "Operadora : " & vDatosCon.Operadora & "Plan : " & vDatosCon.Plan & " - Contrato : " & vDatosCon.Contrato
   itmX.SubItems(9) = "N/A" 'rs!fnd_Usuario
   itmX.SubItems(10) = vDatosCon.Plan


 prg.Value = prg.Value + 1
 lblEstado = "PASO 3/3 - Registros Evaluados : " & Format(prg.Value, "###,###,###,##0") _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents

 rs.MoveNext
 lngLineas = lngLineas + 1

Loop
rs.Close


'TOTALES
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(4) = "---------------------"
    itmX.SubItems(5) = "---------------------"
          
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(4) = Format(curTotales(1), "Standard")
    itmX.SubItems(5) = Format(curTotales(2), "Standard")
  

Me.MousePointer = vbDefault

End Sub



Private Sub sbArchivo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngRegistros As Long, curTotales(2) As Currency
Dim strCadena As String, fn, lng As Long, i As Integer
Dim vMascara As String, lngEvaluados As Long, vRuta As String

Me.MousePointer = vbHourglass


fn = FreeFile

vRuta = SIFGlobal.DirectorioDeResultados & "\Fondos_DetCtaDocs.txt"

Call sbTitulos

Open vRuta For Output As #fn  ' Crea Archivo.

'Imprimir Encabezados
Print #fn, "Fondos: Detalle de Movimientos a Cuentas" & vbCrLf _
            & "Del : " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy")


strCadena = ""
For lng = 1 To lsw.ColumnHeaders.Count
  strCadena = strCadena & UCase(lsw.ColumnHeaders.Item(lng).Text) & vbTab
Next lng
Print #fn, strCadena


lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"
DoEvents

vMascara = "###-##-###-##"

strSQL = "select O.descripcion,D.concepto,D.cliente,D.usuario,D.fecha,A.*" _
       & " from fnd_documentos D inner join fnd_asientos A on D.tipo = A.tipo" _
       & " and D.id_documento = A.id_documento and D.cod_operadora = A.cod_operadora" _
       & " inner join fnd_operadoras O on O.cod_operadora = D.cod_operadora" _
       & " where D.fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & "' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & "' order by D.fecha"

Call OpenRecordSet(rs, strSQL)
lngRegistros = rs.RecordCount
lngEvaluados = 0
prg.Max = rs.RecordCount + 1
prg.Value = 1
    
curTotales(1) = 0
curTotales(2) = 0

Do While Not rs.EOF
 lngEvaluados = lngEvaluados + 1

 strCadena = Format(rs!fecha, "dd/mm/yyyy") & vbTab & rs!Tipo & vbTab & rs!id_documento _
           & vbTab & Format(rs!fnd_cuenta, vMascara) & vbTab
    
  If rs!fnd_DEBEHABER = "D" Then
     strCadena = strCadena & rs!fnd_monto & vbTab & 0 & vbTab
     curTotales(1) = curTotales(1) + rs!fnd_monto
  Else
     strCadena = strCadena & 0 & vbTab & rs!fnd_monto & vbTab
     curTotales(2) = curTotales(2) + rs!fnd_monto
  End If
  strCadena = strCadena & rs!CONCEPTO & vbTab & rs!Cliente & vbTab & rs!Descripcion & vbTab & rs!Usuario & vbTab & rs!cod_Plan

 prg.Value = prg.Value + 1
 lblEstado = "Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
           & Format(lngRegistros, "###,###,###,##0") & vbCrLf _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents
 
 Print #fn, strCadena

 rs.MoveNext
Loop
rs.Close


'me falta la fase ii

'*******************************************************************************************
'FASE II = BUSCA DEDUCCIONES POR PLANILLA

lblEstado.Caption = vbCrLf & "****- Cargando Información Ded.Pla. (Espere) -****"
DoEvents

strSQL = "Select P.codigo,P.id_solicitud,P.amortiza,P.opex,P.fecha,P.fecha_proceso," _
       & "S.Nombre,C.CTANAMORT,C.CTAOAMORT,P.cedula " _
       & " from prm_creditos P Left Join Socios S on P.cedula = S.cedula" _
       & " left join catalogo C on P.codigo = C.codigo" _
       & " inner join fnd_planes F on P.codigo = F.codigo_ase" _
       & " where P.IND_PASO = 1 AND P.amortiza > 0 and P.fecha between '" _
       & Format(dtpInicio.Value, "yyyy/mm/dd") & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    lngRegistros = rs.RecordCount
    lngEvaluados = 0
    prg.Max = rs.RecordCount + 1
    prg.Value = 1
End If

Do While Not rs.EOF
 
    lngEvaluados = lngEvaluados + 1
    
    strCadena = Format(rs!fecha, "dd/mm/yyyy") & vbTab & "PRM" & vbTab & rs!Fecha_Proceso & vbTab

    If rs!opex = 1 Then
       strCadena = strCadena & Format(IIf(IsNull(rs!CtaOamort), "", rs!CtaOamort), vMascara)
    Else
       strCadena = strCadena & Format(IIf(IsNull(rs!CtaNamort), "", rs!CtaNamort), vMascara)
    End If
    
    strCadena = strCadena & vbTab & "0" & vbTab & rs!Amortiza & vbTab & "Ded.Pla : " & rs!Fecha_Proceso _
              & vbTab & Trim(rs!Cedula) & " - " & rs!Nombre & vbTab & Trim(rs!Codigo) & " Op." & rs!id_solicitud & " Ex." & rs!opex _
              & vbTab & "Ded.Pla"
    
    Print #fn, strCadena
    
    curTotales(2) = curTotales(2) + rs!Amortiza
   
    prg.Value = prg.Value + 1
    lblEstado = "Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
              & Format(lngRegistros, "###,###,###,##0") & vbCrLf _
              & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
              & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
    DoEvents

    rs.MoveNext
Loop
rs.Close



'Procesar la Cola de Asientos
lblEstado.Caption = vbCrLf & "****- Cargando Información Complementaria (Espere) -****"
DoEvents

'Excluir Asiento del Proceso Mensual, ya que este es resumen y el detalle
'ya fue cargado en el proceso anterior

strSQL = "select C.descripcion,A.*" _
       & " from fnd_asientos_cola A inner join Cuentas C" _
       & " On C.cod_cuenta = A.fnd_cuenta" _
       & " where A.fnd_fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & "' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & "' and fnd_tipo <> 'PRM'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
    prg.Max = rs.RecordCount + 1
    prg.Value = 1
    lngEvaluados = 0
End If

Do While Not rs.EOF
 lngEvaluados = lngEvaluados + 1
 
    vDatosCon.Cedula = fxCedulayNombre(rs!cod_Plan, rs!Cod_Operadora, rs!COD_CONTRATO)
    vDatosCon.Contrato = rs!COD_CONTRATO
    vDatosCon.Operadora = rs!Cod_Operadora
    vDatosCon.Plan = rs!cod_Plan
     
    Select Case UCase(rs!fnd_tipo)
      Case "LIQ" 'LIQUIDACIONES
            vDatosCon.Movimiento = "LIQUIDACION"
      Case "RET" 'RETIROS
            vDatosCon.Movimiento = "RETIROS"
      Case "REN" 'APLICACION DE RENDIMIENTOS
            vDatosCon.Movimiento = "RENDIMIENTOS"
      Case "PRM" 'PROCESOS MENSUALES N/A
            vDatosCon.Movimiento = "PROCESO MENSUAL"
    End Select

 
 strCadena = Format(rs!fnd_Fecha, "yyyy/mm/dd") & vbTab & rs!fnd_tipo & vbTab & "Cola Asientos" _
           & vbTab & Format(rs!fnd_cuenta, vMascara) & vbTab
    
  If rs!fnd_DEBEHABER = "D" Then
     strCadena = strCadena & rs!fnd_monto & vbTab & 0 & vbTab
     curTotales(1) = curTotales(1) + rs!fnd_monto
  Else
     strCadena = strCadena & 0 & vbTab & rs!fnd_monto & vbTab
     curTotales(2) = curTotales(2) + rs!fnd_monto
  End If
  
  strCadena = strCadena & UCase(vDatosCon.Movimiento) & " - " & rs!fnd_caso & vbTab & Trim(vDatosCon.Cedula) _
            & vbTab & "Operadora : " & vDatosCon.Operadora & "Plan : " & vDatosCon.Plan & " - Contrato : " & vDatosCon.Contrato _
            & vbTab & "N/A" ' rs!fnd_Usuario
     
 prg.Value = prg.Value + 1
 lblEstado = "Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
           & Format(prg.Max - 1, "###,###,###,##0") & vbCrLf _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 DoEvents
 
 Print #fn, strCadena
  
 rs.MoveNext

Loop
rs.Close


'TOTALES
  strCadena = vbTab & vbTab & vbTab & vbTab & "---------------------" & vbTab & "---------------------"
  Print #fn, strCadena
  strCadena = vbTab & vbTab & vbTab & vbTab & curTotales(1) & vbTab & curTotales(2)
  Print #fn, strCadena

'Cierra Archivo
Close #fn

Me.MousePointer = vbDefault

MsgBox "Se Creo Archivo de Texto con Tabulaciones en :C:\FNDDetCtaDoc.txt", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 On Error Resume Next
 Close #fn

End Sub




Private Sub cboResultados_Click()
lsw.ListItems.Clear
End Sub

Private Sub cboResultados_KeyDown(KeyCode As Integer, Shift As Integer)
lsw.ListItems.Clear
If KeyCode = vbKeyReturn Then cmdBuscar.SetFocus
End Sub

Private Sub cmdArchivo_Click()
Dim strTitulo As String, vRuta As String


If cboResultados.Text = "Resumen" Then
    If lsw.ListItems.Count <= 0 Then
       MsgBox "No hay información que almacenar...", vbExclamation
       Exit Sub
    End If
    
    vRuta = SIFGlobal.DirectorioDeResultados & "\Fondos_RsmCtaDocs.txt"
    strTitulo = "Resumen de Movimientos a Cuentas Registradas en Documentos FND" & vbCrLf _
              & vbCrLf & lblEstado.Caption
    Call sbgFNDCreaArchivo(lsw, vRuta, strTitulo)
    Exit Sub
End If

'Archivo detalle

Call sbArchivo



End Sub


Private Sub cmdBuscar_Click()
If dtpInicio > dtpCorte Then Exit Sub
'Encabezados
Call sbTitulos
With lsw
 If cboResultados.Text = "Resumen" Then
    Call sbResumen
 Else
    Call sbDetalle
 End If
End With
End Sub


Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboResultados.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpCorte.SetFocus
End Sub


Private Sub sbTitulos()
lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

'Encabezados
With lsw
 If cboResultados.Text = "Resumen" Then
    .ColumnHeaders.Add , , "Cuenta", 1700
    .ColumnHeaders.Add , , "Descripción", 3200
    .ColumnHeaders.Add , , "Débito", 2000, vbRightJustify
    .ColumnHeaders.Add , , "Crédito", 2000, vbRightJustify
    .ColumnHeaders.Add , , "Ubicación", 2000, vbCenter
 Else
    .ColumnHeaders.Add , , "Fecha", 1200
    .ColumnHeaders.Add , , "Tipo", 1000, vbCenter
    .ColumnHeaders.Add , , "N°Documento", 1300
    .ColumnHeaders.Add , , "Cuenta", 1700
    .ColumnHeaders.Add , , "Débito", 1800, vbRightJustify
    .ColumnHeaders.Add , , "Crédito", 1800, vbRightJustify
    .ColumnHeaders.Add , , "Concepto", 3000
    .ColumnHeaders.Add , , "Cliente", 3000
    .ColumnHeaders.Add , , "DP", 1000
    .ColumnHeaders.Add , , "Usuario", 3000
    .ColumnHeaders.Add , , "Plan", 3000
 End If
End With

End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
'Me.Icon = frmFNDOperadoras.Icon

vModulo = 18 'Fondo de Inversion

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

cboResultados.AddItem "Resumen"
cboResultados.AddItem "Detalle"
cboResultados.Text = "Resumen"

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width


lsw.Width = Me.Width - 100
lsw.Height = Me.Height - (800 + lblEstado.Height + prg.Height + lsw.top)

lblEstado.top = lsw.top + lsw.Height + 20
lblEstado.Width = lsw.Width

End Sub

