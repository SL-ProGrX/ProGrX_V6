VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmRastreoMovDoc 
   Caption         =   "Analítico de Cuentas [Contable] del Auxiliar"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   15132
   HelpContextID   =   7002
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   15132
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3492
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   5652
      _Version        =   1245185
      _ExtentX        =   9970
      _ExtentY        =   6159
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
   Begin XtremeSuiteControls.GroupBox fraCuentas 
      Height          =   1812
      Left            =   6240
      TabIndex        =   14
      Top             =   1080
      Visible         =   0   'False
      Width           =   7572
      _Version        =   1245185
      _ExtentX        =   13356
      _ExtentY        =   3196
      _StockProps     =   79
      Caption         =   "Filtros adicionales:"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkCuentas 
         Height          =   252
         Left            =   1080
         TabIndex        =   21
         Top             =   1440
         Width           =   3132
         _Version        =   1245185
         _ExtentX        =   5524
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Mostrar todas las cuentas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaInicio 
         Height          =   312
         Left            =   1080
         TabIndex        =   17
         Top             =   480
         Width           =   1932
         _Version        =   1245185
         _ExtentX        =   3408
         _ExtentY        =   550
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaCorte 
         Height          =   312
         Left            =   1080
         TabIndex        =   18
         Top             =   840
         Width           =   1932
         _Version        =   1245185
         _ExtentX        =   3408
         _ExtentY        =   550
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaInicioDes 
         Height          =   312
         Left            =   3000
         TabIndex        =   19
         Top             =   480
         Width           =   4452
         _Version        =   1245185
         _ExtentX        =   7853
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCtaCorteDes 
         Height          =   312
         Left            =   3000
         TabIndex        =   20
         Top             =   840
         Width           =   4452
         _Version        =   1245185
         _ExtentX        =   7853
         _ExtentY        =   550
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   7
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   6
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   6720
      TabIndex        =   9
      Top             =   240
      Width           =   1452
      _Version        =   1245185
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
      Picture         =   "frmRastreoMovDoc.frx":0000
   End
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   276
      Left            =   0
      TabIndex        =   0
      Top             =   5364
      Width           =   15132
      _ExtentX        =   26691
      _ExtentY        =   487
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   1452
      _Version        =   1245185
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
      TabIndex        =   4
      Top             =   480
      Width           =   1452
      _Version        =   1245185
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
      Left            =   3960
      TabIndex        =   5
      Top             =   480
      Width           =   1812
      _Version        =   1245185
      _ExtentX        =   3196
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
      Left            =   8160
      TabIndex        =   10
      Top             =   240
      Width           =   1452
      _Version        =   1245185
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Archivo"
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
      Picture         =   "frmRastreoMovDoc.frx":0A1E
   End
   Begin XtremeSuiteControls.PushButton btnFiltros 
      Height          =   492
      Left            =   6240
      TabIndex        =   11
      ToolTipText     =   "Filtros Adicionales"
      Top             =   240
      Width           =   492
      _Version        =   1245185
      _ExtentX        =   868
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "..."
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
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   372
      Left            =   10080
      TabIndex        =   12
      Top             =   360
      Width           =   732
      _Version        =   1245185
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
         Name            =   "Arial Narrow"
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
      Left            =   10080
      TabIndex        =   13
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
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
      TabIndex        =   8
      Top             =   240
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
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
      TabIndex        =   6
      Top             =   480
      Width           =   1212
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   5412
      _Version        =   1245185
      _ExtentX        =   9546
      _ExtentY        =   868
      _StockProps     =   79
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
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmRastreoMovDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type vFx
  Cedula As String
  Codigo As String
  Operacion As Long
  Movimiento As String
End Type

Dim vDatosCon As vFx



Private Sub sbResumen()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotales(2) As Currency
Dim vCuentaInicio As String, vCuentaCorte As String
Dim lngLineas As Long

Me.MousePointer = vbHourglass
lsw.ListItems.Clear

vCuentaInicio = fxgCntCuentaFormato(False, txtCtaInicio, 0)
vCuentaCorte = fxgCntCuentaFormato(False, txtCtaCorte, 0)


lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"



'Control de Documentos v2
strSQL = "select Count(*) as Total,COD_CUENTA_MASK,CUENTA_DESC,sum(MONTO_DEBITO) as 'DEBITO', sum(MONTO_CREDITO) as 'CREDITO'" _
       & ",cod_unidad,cod_Centro_Costo, Cod_Divisa,AVG(TIPO_CAMBIO) as 'Tipo_Cambio'" _
       & " From vSys_Aux_Transacciones_Cuentas" _
       & " where Registro_Fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
       & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59'"

If chkCuentas.Value = vbUnchecked Then
   strSQL = strSQL & " and cod_cuenta between '" & vCuentaInicio & "' and '" & vCuentaCorte & "'"
End If
       
strSQL = strSQL & " group by COD_CUENTA_MASK,CUENTA_DESC,cod_unidad,cod_Centro_Costo,cod_Divisa"

strSQL = strSQL & " order by COD_CUENTA_MASK"

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
     
     
     
Set itmX = lsw.ListItems.Add(, , rs!cod_Cuenta_Mask)
    itmX.SubItems(1) = rs!CUENTA_DESC
    itmX.SubItems(2) = Format(rs!Debito, "Standard")
    itmX.SubItems(3) = Format(rs!Credito, "Standard")
    
    curTotales(1) = curTotales(1) + rs!Debito
    curTotales(2) = curTotales(2) + rs!Credito
  
  itmX.SubItems(4) = "Control Doc."
  itmX.SubItems(5) = Trim(rs!cod_unidad)
  itmX.SubItems(6) = Trim(rs!cod_centro_costo)
  itmX.SubItems(7) = Trim(rs!cod_Divisa)
  itmX.SubItems(8) = rs!Tipo_Cambio
  

 prg.Value = prg.Value + 1
 lblEstado.Caption = "Registros Evaluados : " & Format(rs!total, "###,###,###,##0") _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"

 rs.MoveNext
 lngLineas = lngLineas + 1

Loop
rs.Close



'TOTALES
         
  Set itmX = lsw.ListItems.Add(, , "")
    itmX.SubItems(2) = Format(curTotales(1), "Standard")
    itmX.SubItems(3) = Format(curTotales(2), "Standard")
    
    itmX.Bold = True
    itmX.ForeColor = vbWhite
    itmX.TextBackColor = RGB(214, 234, 248)
  



Me.MousePointer = vbDefault


End Sub


Private Sub sbDetalle()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, curTotales(2) As Currency
Dim vCuentaInicio As String, vCuentaCorte As String
Dim lngLineas As Long
Dim lngRegistros As Long, lngEvaluados As Long

Me.MousePointer = vbHourglass
lsw.ListItems.Clear

vCuentaInicio = fxgCntCuentaFormato(False, txtCtaInicio, 0)
vCuentaCorte = fxgCntCuentaFormato(False, txtCtaCorte, 0)


lblEstado.Caption = vbCrLf & "****- Cargando Información (Espere) -****"

strSQL = "select Top " & txtLineas.Text & " Con.Descripcion as 'ConceptoDesc',D.Registro_Fecha,D.Registro_Usuario,D.cod_Oficina" _
         & ", rtrim(D.Cliente_Identificacion) + ' - ' + D.Cliente_Nombre as 'Cliente',D.Documento,A.*,1 as 'Tipo_Cambio'" _
         & " from SIF_TRANSACCIONES D inner join SIF_TRANSACCIONES_ASIENTO A on D.tipo_Documento = A.Tipo_Documento" _
         & " and D.cod_Transaccion = A.cod_Transaccion" _
         & " inner join SIF_Conceptos Con on D.cod_concepto = Con.Cod_Concepto" _
         & " where D.Registro_Fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
         & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59'"
         
If chkCuentas.Value = vbUnchecked Then
   strSQL = strSQL & " and A.cod_cuenta between '" & vCuentaInicio & "' and '" & vCuentaCorte & "'"
End If
        
strSQL = strSQL & " order by D.Registro_Fecha"


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
     

Set itmX = lsw.ListItems.Add(, , Format(rs!Registro_Fecha, "dd/mm/yyyy"))
    itmX.SubItems(1) = rs!Tipo_Documento
    itmX.SubItems(2) = rs!Cod_Transaccion
    itmX.SubItems(3) = Format(rs!COD_Cuenta, GLOBALES.gstrMascara)
    
  If rs!Tipo_Movimiento = "D" Then
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = "0.00"
     curTotales(1) = curTotales(1) + rs!Monto
  Else
     itmX.SubItems(4) = "0.00"
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     curTotales(2) = curTotales(2) + rs!Monto
  End If
   
   itmX.SubItems(6) = rs!ConceptoDesc & ""
   itmX.SubItems(7) = rs!Cliente & ""
   itmX.SubItems(8) = rs!Documento & ""
   itmX.SubItems(9) = rs!Registro_Usuario & ""
   
   itmX.SubItems(10) = rs!cod_unidad & ""
   itmX.SubItems(11) = rs!cod_centro_costo & ""
   itmX.SubItems(12) = rs!cod_Divisa & ""
   itmX.SubItems(13) = rs!Tipo_Cambio & ""
   
   
   itmX.SubItems(14) = rs!cod_Oficina & ""
   itmX.SubItems(15) = rs!Referencia_01 & ""
   itmX.SubItems(16) = rs!Referencia_02 & ""
   itmX.SubItems(17) = rs!Referencia_03 & ""
   
   
   

 prg.Value = prg.Value + 1
 lblEstado.Caption = "Registros Evaluados : " & Format(lngEvaluados, "###,###,###,##0") & " de " _
           & Format(lngRegistros, "###,###,###,##0") & vbCrLf _
           & " Fechas del " & Format(dtpInicio, "dd/mm/yyyy") & " al " & Format(dtpCorte, "dd/mm/yyyy") _
           & "  -  Porcentaje Procesado: " & Round((prg.Value / prg.Max) * 100, 2) & "%"
 
 If Right(CStr(prg.Value), 2) = "00" Then DoEvents

 rs.MoveNext
 lngLineas = lngLineas + 1

Loop
rs.Close


'TOTALES
         
  Set itmX = lsw.ListItems.Add(, , "T")
    itmX.SubItems(4) = Format(curTotales(1), "Standard")
    itmX.SubItems(5) = Format(curTotales(2), "Standard")
    
    itmX.Bold = True
    itmX.ForeColor = vbWhite
    itmX.TextBackColor = RGB(214, 234, 248)
  

Me.MousePointer = vbDefault

End Sub



Private Sub btnFiltros_Click()
Dim vValor As Boolean

If fraCuentas.Visible Then
   vValor = False
Else
   vValor = True
End If

fraCuentas.Visible = vValor

lsw.Visible = Not vValor


End Sub

Private Sub cboResultados_Click()
lsw.ListItems.Clear
End Sub

Private Sub cboResultados_KeyDown(KeyCode As Integer, Shift As Integer)
lsw.ListViewItems.Clear
If KeyCode = vbKeyReturn Then cmdBuscar.SetFocus
End Sub

Private Sub chkCuentas_Click()
If chkCuentas.Value = vbChecked Then
   txtCtaInicio.Enabled = False
Else
   txtCtaInicio.Enabled = True
End If

txtCtaCorte.Enabled = txtCtaInicio.Enabled

End Sub

Private Sub cmdArchivo_Click()

Call sbListViewExporFileTab(lsw)

End Sub

Private Sub cmdBuscar_Click()


If dtpInicio > dtpCorte Then Exit Sub

'Encabezados
Call sbTitulos

If cboResultados.Text = "Resumen" Then
   Call sbResumen
Else
   Call sbDetalle
End If

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
    .ColumnHeaders.Add , , "Unidad", 1200, vbCenter
    .ColumnHeaders.Add , , "Centro Costo", 1200, vbCenter
    .ColumnHeaders.Add , , "Divisa", 1200, vbCenter
    .ColumnHeaders.Add , , "Tipo Cambio", 1200, vbRightJustify
    
 
 Else
    .ColumnHeaders.Add , , "Fecha", 1200
    .ColumnHeaders.Add , , "Tipo", 1000, vbCenter
    .ColumnHeaders.Add , , "N°Documento", 1300
    .ColumnHeaders.Add , , "Cuenta", 1700
    .ColumnHeaders.Add , , "Débito", 1800, vbRightJustify
    .ColumnHeaders.Add , , "Crédito", 1800, vbRightJustify
    .ColumnHeaders.Add , , "Concepto", 3000
    .ColumnHeaders.Add , , "Cliente", 3600
    .ColumnHeaders.Add , , "DP", 1000
    .ColumnHeaders.Add , , "Usuario", 3000
    .ColumnHeaders.Add , , "Unidad", 1200, vbCenter
    .ColumnHeaders.Add , , "Centro Costo", 1200, vbCenter
    .ColumnHeaders.Add , , "Divisa", 1200, vbCenter
    .ColumnHeaders.Add , , "Tipo Cambio", 1200, vbRightJustify
 
    .ColumnHeaders.Add , , "Oficina", 1200, vbCenter
    .ColumnHeaders.Add , , "Ref_01", 1200, vbCenter
    .ColumnHeaders.Add , , "Ref_02", 1200, vbCenter
    .ColumnHeaders.Add , , "Ref_03", 1200, vbCenter
 
 
 End If
End With

End Sub


Private Sub Form_Load()

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboResultados.AddItem "Resumen"
cboResultados.AddItem "Detalle"
cboResultados.Text = "Resumen"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value


End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - 150
lsw.Height = Me.Height - (lsw.Top + lblEstado.Height + prg.Height + 480)
lblEstado.Top = lsw.Top + lsw.Height + 20
lblEstado.Width = lsw.Width

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub txtCtaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtCtaInicio.Text = fxgCntCuentaFormato(True, gCuenta)
  txtCtaInicioDes.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, gCuenta))
End If

If KeyCode = vbKeyReturn Then
  txtCtaInicioDes.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaInicio))
  txtCtaInicio.Text = fxgCntCuentaFormato(True, txtCtaInicio)
End If

End Sub

Private Sub txtCtaCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  txtCtaCorte.Text = fxgCntCuentaFormato(True, gCuenta)
  txtCtaCorteDes.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, gCuenta))
End If

If KeyCode = vbKeyReturn Then
  txtCtaCorteDes.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCtaCorte))
  txtCtaCorte.Text = fxgCntCuentaFormato(True, txtCtaCorte)
End If

End Sub

