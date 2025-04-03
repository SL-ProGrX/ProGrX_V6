VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAF_CartasNoCotizantes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Casos no Cotizantes"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   7.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAF_CartasNoCotizantes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5430
   ScaleWidth      =   9195
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3972
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   9012
      _Version        =   1441793
      _ExtentX        =   15896
      _ExtentY        =   7006
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   1
      Item(0).Caption =   "Informes"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "btnOpt(0)"
      Item(0).Control(1)=   "btnOpt(1)"
      Item(0).Control(2)=   "btnOpt(2)"
      Item(0).Control(3)=   "gbInforme"
      Item(0).Control(4)=   "Label1(3)"
      Item(0).Control(5)=   "Label1(1)"
      Item(0).Control(6)=   "Label1(0)"
      Item(0).Control(7)=   "cboMeses"
      Item(0).Control(8)=   "cboMora"
      Item(0).Control(9)=   "dtpIngreso"
      Item(0).Control(10)=   "chkCreditos"
      Item(0).Control(11)=   "txtMesesNoCotizar"
      Item(0).Control(12)=   "txtCuotaMora"
      Item(0).Control(13)=   "chkEmail"
      Item(0).Control(14)=   "rbSalida(0)"
      Item(0).Control(15)=   "rbSalida(1)"
      Begin XtremeSuiteControls.RadioButton rbSalida 
         Height          =   252
         Index           =   0
         Left            =   6360
         TabIndex        =   18
         Top             =   1920
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Visualizar En Pantalla"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   -1  'True
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkCreditos 
         Height          =   252
         Left            =   3120
         TabIndex        =   14
         Top             =   1920
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Que no posean Créditos"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox gbInforme 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   2760
         Width           =   8535
         _Version        =   1441793
         _ExtentX        =   15055
         _ExtentY        =   1720
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdReporte 
            Height          =   615
            Left            =   6960
            TabIndex        =   6
            Top             =   360
            Width           =   1575
            _Version        =   1441793
            _ExtentX        =   2778
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Reporte"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmAF_CartasNoCotizantes.frx":000C
         End
         Begin MSComctlLib.ProgressBar prgBar 
            Height          =   168
            Left            =   0
            TabIndex        =   10
            Top             =   600
            Visible         =   0   'False
            Width           =   5532
            _ExtentX        =   9763
            _ExtentY        =   291
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblEstado 
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
            Height          =   252
            Left            =   0
            TabIndex        =   20
            Top             =   240
            Width           =   2652
         End
      End
      Begin XtremeSuiteControls.PushButton btnOpt 
         Height          =   492
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Carta de Aviso"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnOpt 
         Height          =   492
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Sobres"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton btnOpt 
         Height          =   492
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Listado"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cboMeses 
         Height          =   312
         Left            =   4680
         TabIndex        =   11
         Top             =   720
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.ComboBox cboMora 
         Height          =   312
         Left            =   4680
         TabIndex        =   12
         Top             =   1080
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.DateTimePicker dtpIngreso 
         Height          =   312
         Left            =   4680
         TabIndex        =   13
         Top             =   1440
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   550
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtMesesNoCotizar 
         Height          =   312
         Left            =   6120
         TabIndex        =   15
         Top             =   720
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   550
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
         Text            =   "6"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuotaMora 
         Height          =   312
         Left            =   6120
         TabIndex        =   16
         Top             =   1080
         Width           =   732
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   550
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
         Text            =   "1"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkEmail 
         Height          =   252
         Left            =   3120
         TabIndex        =   17
         Top             =   2280
         Width           =   2892
         _Version        =   1441793
         _ExtentX        =   5101
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Enviar Carta de Aviso por Email"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.RadioButton rbSalida 
         Height          =   252
         Index           =   1
         Left            =   6360
         TabIndex        =   19
         Top             =   2280
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Enviar a la Impresora"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Alignment       =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Meses de No Cotizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   2280
         TabIndex        =   9
         Top             =   720
         Width           =   2652
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas Mora en Créditos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   2280
         TabIndex        =   8
         Top             =   1080
         Width           =   2652
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ingreso Menor a"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   2280
         TabIndex        =   7
         Top             =   1440
         Width           =   2652
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control de Casos no Cotizantes"
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
      Height          =   492
      Index           =   4
      Left            =   2004
      TabIndex        =   0
      Top             =   360
      Width           =   5412
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_CartasNoCotizantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub sbReporteListado()
Dim strMora As String, strMeses As String
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Afiliaciones"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "SubTitulo='Ingreso antes de: " & Format(dtpIngreso.Value, "yyyy/mm/dd") _
              & " ¦ Meses " & cboMeses.Text & " " & txtMesesNoCotizar _
              & " ¦ Mora " & cboMora.Text & " " & txtCuotaMora & "'"

 .Connect = glogon.ConectRPT

 Select Case Trim(cboMeses)
     Case ">="
       strMeses = Trim(cboMeses) & txtMesesNoCotizar
     Case "<="
       strMeses = Trim(cboMeses) & txtMesesNoCotizar
     Case "="
       strMeses = Trim(cboMeses) & txtMesesNoCotizar
  End Select
  
  If chkCreditos.Value = 0 Then
    Select Case Trim(cboMora)
       Case ">="
         strMora = Trim(cboMora) & txtCuotaMora
       Case "<="
         strMora = Trim(cboMora) & txtCuotaMora
       Case "="
         strMora = Trim(cboMora) & txtCuotaMora
     End Select
    strSQL = "{sp_NoCotizantes;1.Meses} " & strMeses & " and {sp_NoCotizantes;1.Cuotas} " & strMora
 Else
    strSQL = "{sp_NoCotizantes;1.Meses} " & strMeses & " and {sp_NoCotizantes;1.Saldos} = 0.00 "
 End If
  
   
   '6 and {sp_NoCotizantes;1.Saldos} = 0.00"
 .ReportFileName = SIFGlobal.fxPathReportes("Personas_NoCotizanteListado.rpt")
  
  .StoredProcParam(0) = Format(dtpIngreso.Value, "yyyy-mm-dd") & " 01:00:00.000"
  '.StoredProcParam(1) = CCur(txtCuotaMora)
  '.StoredProcParam(2) = CCur(txtMesesNoCotizar)
  .SelectionFormula = strSQL
  
 .Action = 1
 '.PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub btnOpt_Click(Index As Integer)

gbInforme.Caption = btnOpt(Index).Caption

End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vParrafo As String, vProceso As String
Dim vMes As Integer


If gbInforme.Caption = "Listado" Then
  Call sbReporteListado
 Exit Sub
End If

Me.MousePointer = vbHourglass
On Error GoTo vError


strSQL = "select * from par_ahcr"
Call OpenRecordSet(rs, strSQL)
vMes = Mid(GLOBALES.glngFechaCR, 5, 2)
If rs!cr_apl = 0 Then
 If vMes = 1 Then
   vMes = 12
 Else
   vMes = vMes - 1
 End If
End If
rs.Close


Select Case vMes
  Case 1
    vProceso = " Mes de Enero "
  Case 2
    vProceso = " Mes de Febrero "
  Case 3
    vProceso = " Mes de Marzo "
  Case 4
    vProceso = " Mes de Abril "
  Case 5
    vProceso = " Mes de Mayo "
  Case 6
    vProceso = " Mes de Junio "
  Case 7
    vProceso = " Mes de Julio "
  Case 8
    vProceso = " Mes de Agostro "
  Case 9
    vProceso = " Mes de Setiembre "
  Case 10
    vProceso = " Mes de Octubre "
  Case 11
    vProceso = " Mes de Noviembre "
  Case 12
    vProceso = " Mes de Diciembre "
End Select

vParrafo = "Sirva la presente para saludarle y a la vez informarle que de acuerdo a nuestros" _
         & " registros usted presenta más de " & txtMesesNoCotizar & " meses sin cotizar y que "
         
If CCur(txtCuotaMora) > 0 Then
  vParrafo = vParrafo & " su (s) préstamo (s) se encuentra (n) atrasado (s) al " & vProceso _
           & ", los cuales se detallan :"
Else
  vParrafo = vParrafo & " su (s) préstamo (s) se podría (n) encuentra (n) atrasado (s) al " & vProceso _
           & ", los cuales se detallan :"
End If

strSQL = "select S.cedula,S.nombre,datediff(m,A.fecahorro,dbo.MyGetdate()) as Meses" _
       & ",isnull(sum(R.saldo),0) as Saldos,isnull(sum(V.Intc),0) as IntCor" _
       & ",isnull(sum(V.IntM),0) as IntMor,isnull(sum(V.cuota),0) as Cuotas" _
       & " from Socios S inner join Ahorro_consolidado A on S.cedula = A.cedula" _
       & " and datediff(m,A.fecahorro,dbo.MyGetdate()) " & cboMeses.Text & " " & txtMesesNoCotizar _
       & " and S.estadoactual = 'S' and S.fechaingreso < '" & Format(dtpIngreso.Value, "yyyy/mm/dd") & "'" _
       & " left join Reg_Creditos R on S.cedula = R.cedula" _
       & " inner join Vista_Morosidad V on R.id_solicitud = V.id_solicitud" _
       & " inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " group by S.cedula,S.nombre,A.fecahorro" _
       & " Having isnull(Sum(V.cuota), 0) " & cboMora.Text & " " & txtCuotaMora

lblEstado.Caption = "Cargando Información Espere...."
lblEstado.Refresh

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1

PrgBar.Visible = True

Do While Not rs.EOF
  
  lblEstado.Caption = "Procesando Carta No. " & PrgBar.Value & " de " & PrgBar.Max
  lblEstado.Refresh
  

  With frmContenedor.Crt
    .Reset
    Select Case gbInforme.Caption
      Case "Carta de Avisto"
           .Formulas(0) = "fxParrafo01='" & Mid(vParrafo, 1, 250) & "'"
           .Formulas(1) = "Copia='cc. Archivo'"
           .ReportFileName = SIFGlobal.fxPathReportes("NoCotizanteCarta.rpt")
           .SelectionFormula = "{SOCIOS.CEDULA} = '" & rs!Cedula & "'"
           .SubreportToChange = "sbCreditosNoCotizantes"
           .SelectionFormula = "{REG_CREDITOS.CEDULA} = {?Pm-SOCIOS.CEDULA}" _
                             & " AND {REG_CREDITOS.SALDO} > 0 AND {REG_CREDITOS.ESTADO} = 'A'"
           
           strSQL = "insert SOCIOS_AVISOS(CEDULA,USUARIO,FECHA,MESES,MORA) values('" & Trim(rs!Cedula) _
                  & "','" & glogon.Usuario & "',dbo.MyGetdate()," & txtMesesNoCotizar & "," & txtCuotaMora & ")"
           Call ConectionExecute(strSQL)
                       
      Case "Sobres"
           .SelectionFormula = "{SOCIOS.CEDULA} = '" & rs!Cedula & "'"
           .ReportFileName = SIFGlobal.fxPathReportes("NoCotizanteSobre.rpt")
    End Select
    
    'Salida a Impresora
    If rbSalida.Item(1).Value Then
        .Destination = crptToPrinter
    End If
    
    .PrintReport
  
  End With
  
  PrgBar.Value = PrgBar.Value + 1
  rs.MoveNext
Loop
rs.Close

PrgBar.Visible = False
lblEstado.Caption = ""

Me.MousePointer = vbDefault

Exit Sub

vError:
 PrgBar.Visible = False
 lblEstado.Caption = ""
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

tcMain.Item(0).Selected = True

lblEstado.Caption = ""

cboMeses.Clear
cboMeses.AddItem "="
cboMeses.AddItem ">="
cboMeses.AddItem "<="
cboMeses.Text = ">="

cboMora.Clear
cboMora.AddItem "="
cboMora.AddItem ">="
cboMora.AddItem "<="
cboMora.Text = ">="

gbInforme.Caption = "Listado"

txtCuotaMora.Text = 1
txtMesesNoCotizar.Text = 6
dtpIngreso.Value = fxFechaServidor

End Sub


