VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Begin VB.Form frmCC_Anomalias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajustes Operativos a Cartera de Crédito"
   ClientHeight    =   6888
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   11220
   Icon            =   "frmCC_Anomalias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6888
   ScaleWidth      =   11220
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5412
      Left            =   2640
      TabIndex        =   21
      Top             =   360
      Width           =   8532
      _Version        =   1310720
      _ExtentX        =   15049
      _ExtentY        =   9546
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
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   732
      Left            =   2760
      TabIndex        =   16
      Top             =   5880
      Width           =   8292
      _Version        =   1310720
      _ExtentX        =   14626
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Indique la cuenta contable para registrar el ajuste:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.TextBox txtCuenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   360
         Width           =   1932
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   17
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   360
         Width           =   5175
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   360
         Width           =   852
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   2172
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   2412
      _Version        =   1310720
      _ExtentX        =   4254
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Acciones:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Cuotas Parciales < a 1000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1800
         Width           =   2052
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Saldos < 1000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   2052
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Saldos Negativos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   2052
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Elimina Mora < a 1000"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   2052
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1932
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   2412
      _Version        =   1310720
      _ExtentX        =   4254
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Filtros:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Linea 
         Height          =   336
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   1212
         _Version        =   1310720
         _ExtentX        =   2138
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Destino 
         Height          =   336
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   1212
         _Version        =   1310720
         _ExtentX        =   2138
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Institucion 
         Height          =   336
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1212
         _Version        =   1310720
         _ExtentX        =   2138
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.8
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
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Institución:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   972
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   120
      Top             =   5400
   End
   Begin VB.TextBox txtCasos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   1572
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   5040
      Width           =   1572
   End
   Begin XtremeSuiteControls.PushButton cmdCorrige 
      Height          =   732
      Left            =   1440
      TabIndex        =   15
      Top             =   5880
      Width           =   1092
      _Version        =   1310720
      _ExtentX        =   1926
      _ExtentY        =   1291
      _StockProps     =   79
      Caption         =   "Corregir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmCC_Anomalias.frx":6852
      TextImageRelation=   1
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Left            =   2640
      TabIndex        =   20
      Top             =   0
      Width           =   8532
      _Version        =   1310720
      _ExtentX        =   15049
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Casos Encontrados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   612
   End
   Begin VB.Label Label4 
      Caption         =   "Monto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   5040
      Width           =   852
   End
End
Attribute VB_Name = "frmCC_Anomalias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub sbCorrigeSaldoMenor()
Dim rsCta As New ADODB.Recordset
Dim lngNumero As Long, curTotalSaldos As Currency
Dim strCliente As String, strLinea(11) As String
Dim vFecha As Date, pMonto As Currency, pConcepto As String, pCuenta As String
Dim pTipoDoc As String, pTipoDocId As String, pUnidad As String, pCentroCosto As String, pDivisa As String

On Error GoTo vError

If txtCuenta = "" Then Exit Sub

Me.MousePointer = vbHourglass

pCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)

If Not fxgCntCuentaValida(pCuenta) Then
     Me.MousePointer = vbDefault
     MsgBox "Cuenta Contable no es válida, revisar!", vbExclamation
     Exit Sub
End If



pMonto = fxCrdParametro("23")
pTipoDoc = "NC"
pTipoDocId = "NC"
pConcepto = "CRD007"
curTotalSaldos = 0
lngNumero = fxDocumentoConsecutivo(pTipoDoc)

pDivisa = "COL"

strSQL = "select cod_unidad,cod_centro_costo, dbo.MyGetDate() as 'Fecha' from sif_oficinas where ESTADO = 1 and OFICINA_OMISION = 1"
Call OpenRecordSet(rs, strSQL)
    pUnidad = rs!Cod_Unidad
    pCentroCosto = rs!cod_centro_costo
    vFecha = rs!fecha
rs.Close


'Lineas de Control de Documentos
strLinea(1) = ""
strLinea(2) = "Corrige Saldos Menor a :" & pMonto
strLinea(3) = ""
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = ""
strLinea(7) = ""
strLinea(8) = ""
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""
strLinea(11) = ""

strCliente = UCase("Aplicación General")
vAseDocDetalle = ""
vAseDocDeposito = ""



pTipoDocId = "NC"

strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
      & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
      & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
      & " values('" & lngNumero & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','','" & strCliente & "','" & pConcepto & "',0,'P','" _
      & "','','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
      & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
      & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
      & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
      & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'******************************************************
'Genera el paso 3
'******************************************************

strSQL = "select C.ctanamort,C.ctaoamort,R.codigo,R.id_solicitud,R.cedula,R.saldo,R.opex,R.proceso" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado = 'A' and R.saldo between 0 and " & pMonto _
       & " and R.proceso = 'N' and C.retencion = 'N' and C.poliza = 'N' and R.cod_Divisa = '" & pDivisa _
       & "'"
       
If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If
       
strSQL = strSQL & "  order by R.codigo"
       
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF

    curTotalSaldos = curTotalSaldos + Abs(rs!Saldo)
 
    If GLOBALES.SysPlanPagos = 1 Then
            strSQL = "exec spCrdPlanPagoAbonoEC " & rs!id_solicitud & ",'" & pConcepto & "','" & glogon.Usuario & "','" & pTipoDoc & "'" _
                   & ",'" & lngNumero & "',0,0," & Abs(rs!Saldo) & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "','',1"
            Call ConectionExecute(strSQL)
            
'            strSQL = "exec spCrdPlanPagos " & rs!id_solicitud
'            Call ConectionExecute(strSQL)
     Else
        'Sin Plan de Pagos
            'Actualiza Saldos
             strSQL = "Update reg_creditos set SALDO = 0,AMORTIZA = AMORTIZA + " & rs!Saldo _
                    & ",estado = 'C' where id_solicitud = " & rs!id_solicitud
           
            'Crea Detalle en Creditos Detalle
            strSQL = strSQL & Space(10) & "INSERT CREDITOS_DT(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS,FECHAP,TCON,NCON)" _
                   & " VALUES('" & rs!Codigo & "'," & rs!id_solicitud & ",0," & Abs(rs!Saldo) _
                   & ",0," & Abs(rs!Saldo) & ",'" & Format(vFecha, "yyyy/mm/dd") _
                   & "'," & GLOBALES.glngFechaCR & ",'" & pTipoDocId & "','" & lngNumero & "')"
            Call ConectionExecute(strSQL)
    End If
  

        'Control de Documento v2
        strSQL = "exec spCrdOperacionCtas " & rs!id_solicitud
        rsCta.Open strSQL, glogon.Conection, adOpenStatic
        strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngNumero & "'," & Abs(rs!Saldo) & ",'C','" & rsCta!cod_Divisa _
               & "',1," & GLOBALES.gEnlace & ",'" & rsCta!Cod_Unidad & "','" & rsCta!cod_centro_costo & "','" & rsCta!ctaamortiza _
               & "','" & rsCta!id_solicitud & "','" & rsCta!Codigo & "','" & vAseDocDeposito & "'"
        Call ConectionExecute(strSQL)
        rsCta.Close
  
 rs.MoveNext
Loop
rs.Close


'Débito General
strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngNumero & "'," & curTotalSaldos & ",'D','" & pDivisa _
       & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','" & pCentroCosto & "','" & pCuenta _
       & "','','','" & vAseDocDeposito & "'"
Call ConectionExecute(strSQL)


'Crea Bitacora del Movimiento
Call Bitacora("Aplica", "Elimina Saldos Menores a " & pMonto & "c :" & curTotalSaldos)

Me.MousePointer = vbDefault

MsgBox "Eliminación de Saldos Menores a " & pMonto & " se realiza con Nota de Credito #" & lngNumero, vbInformation

'Imprime nota
If lngNumero > 0 Then Call sbImprimeRecibo(lngNumero, pTipoDoc)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCorrigeSaldoNegativo()
Dim rsCta As New ADODB.Recordset
Dim lngNumero As Long, curTotalSaldos As Currency
Dim strCliente As String, strLinea(11) As String
Dim vFecha As Date, pConcepto As String
Dim pTipoDoc As String, pTipoDocId As String

On Error GoTo vError

If txtCuenta = "" Then Exit Sub

Me.MousePointer = vbHourglass


pTipoDoc = "ND"
pTipoDocId = "ND"
pConcepto = "CRD008"
curTotalSaldos = 0
lngNumero = fxDocumentoConsecutivo(pTipoDoc)
vFecha = fxFechaServidor


'Lineas de Control de Documentos
strLinea(1) = ""
strLinea(2) = "Corrige Saldos Negativos"
strLinea(3) = ""
strLinea(4) = ""
strLinea(5) = ""
strLinea(6) = ""
strLinea(7) = ""
strLinea(8) = ""
strLinea(9) = "Usuario           " & glogon.Usuario
strLinea(10) = ""
strLinea(11) = ""

strCliente = UCase("Aplicación General")
vAseDocDetalle = ""
vAseDocDeposito = ""



'Control de Documentos v2
pTipoDocId = "ND"

strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
      & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
      & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
      & " values('" & lngNumero & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','','" & strCliente & "','" & pConcepto & "',0,'P','" _
      & "','','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
      & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
      & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
      & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
      & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)


'******************************************************
'Genera el paso 3
'******************************************************

strSQL = "select C.ctaCamort,C.ctanamort,C.ctaoamort,R.codigo,R.id_solicitud,R.cedula,R.saldo,R.opex,R.proceso,R.estado,R.estadosol" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo and C.Poliza = 'N' and C.retencion = 'N'" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where estado in('A','C','N') and saldo < 0 " _
       & " and C.retencion = 'N' and C.poliza = 'N' and R.cod_Divisa = 'COL'"


If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If

strSQL = strSQL & " order by R.codigo"

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF

    curTotalSaldos = curTotalSaldos + Abs(rs!Saldo)
 
    If GLOBALES.SysPlanPagos = 1 Then
            strSQL = "exec spCrdPlanPagoAnulaAbono " & rs!id_solicitud & ",'" & pConcepto & "','" & glogon.Usuario & "','" & pTipoDoc _
                   & "','" & lngNumero & "',1,0,0," & Abs(rs!Saldo) & ",0,0,'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
     Else
        'Sin Plan de Pagos
            'Actualiza Saldos
             strSQL = "Update reg_creditos set SALDO = 0,AMORTIZA = AMORTIZA + " & rs!Saldo _
                    & ",estado = 'C' where id_solicitud = " & rs!id_solicitud
            
            'Crea Detalle en Creditos Detalle
            strSQL = strSQL & Space(10) & "INSERT CREDITOS_DT(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS,FECHAP,TCON,NCON)" _
                   & " VALUES('" & rs!Codigo & "'," & rs!id_solicitud & ",0," & Abs(rs!Saldo) _
                   & ",0," & Abs(rs!Saldo) & ",'" & Format(vFecha, "yyyy/mm/dd") _
                   & "'," & GLOBALES.glngFechaCR & ",'" & pTipoDocId & "','" & lngNumero & "')"
            Call ConectionExecute(strSQL)
    End If
  

        'Control de Documento v2
        strSQL = "exec spCrdOperacionCtas " & rs!id_solicitud
        Call OpenRecordSet(rsCta, strSQL)
        strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngNumero & "'," & Abs(rs!Saldo) & ",'D','" & rsCta!cod_Divisa _
               & "',1," & GLOBALES.gEnlace & ",'" & rsCta!Cod_Unidad & "','" & rsCta!cod_centro_costo & "','" & rsCta!ctaamortiza _
               & "','" & rsCta!id_solicitud & "','" & rsCta!Codigo & "','" & vAseDocDeposito & "'"
        Call ConectionExecute(strSQL)
        rsCta.Close
  
 rs.MoveNext
Loop
rs.Close



  With GLOBALES
        strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & lngNumero & "'," & curTotalSaldos & ",'C','COL" _
               & "',1," & .gEnlace & ",'" & .gOficinaUnidad & "','" & .gOficinaCentroCosto & "','" & fxgCntCuentaFormato(False, txtCuenta.Text, 0) _
               & "','','','" & vAseDocDeposito & "'"
        Call ConectionExecute(strSQL)
  End With


'Crea Bitacora del Movimiento
Call Bitacora("Aplica", "Anulacion a Saldos Negativos total:" & curTotalSaldos)

Me.MousePointer = vbDefault

MsgBox "Saldos Negativos Anulados Satisfactoriamente con Nota de Debito #" & lngNumero, vbInformation

'Imprime nota
If lngNumero > 0 Then Call sbImprimeRecibo(lngNumero, pTipoDoc)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCorrigeMora()
Dim pMonto As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass

pMonto = fxCrdParametro("24")


'Elimina Registros de Cargos Asociados

strSQL = "delete Morosidad_Cargos where id_Moro in(select M.id_Moro" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " inner join Morosidad M on R.id_solicitud = M.id_solicitud and M.estado = 'A'" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado = 'A' and R.proceso <> 'J' and R.cod_Divisa = 'COL'" _
       & " and (M.intc + M.intm+ M.amortiza) between 0 and " & pMonto


If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If
       
strSQL = strSQL & ")"
       
Call ConectionExecute(strSQL)



'Elimina Registro de Morosidad
strSQL = "delete M" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " inner join Morosidad M on R.id_solicitud = M.id_solicitud and M.estado = 'A'" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where R.estado = 'A' and R.proceso <> 'J' and R.cod_Divisa = 'COL'" _
       & " and (M.intc + M.intm+ M.amortiza) between 0 and " & pMonto

If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If

Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault

MsgBox "Mora Menor a " & Format(pMonto, "Standard") & " Eliminada Satisfactoriamente...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdCorrige_Click()

If lsw.ListItems.Count <= 0 Then
 MsgBox "No existen datos para Ajustar, consulte nuevamente.!!!", vbExclamation
 Exit Sub
End If

Select Case True
  Case opt(0) 'Saldos < 100
'      sbCodigosDuplicados
    Call sbCorrigeSaldoMenor
  Case opt(1) 'Saldos Negativos
    Call sbCorrigeSaldoNegativo
  Case opt(2).Value 'Mora Inferior a 100c
    If GLOBALES.SysPlanPagos = 0 Then
        Call sbCorrigeMora
    Else
        MsgBox "Esta Opción No Aplica con el Modelo de Plan de Pagos!", vbInformation
    End If

  Case opt(3) 'Cuota Derivada Menor a X
    Call sbCtaDerivada_Corrige
End Select

End Sub

Private Sub FlatEdit_Destino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_DESTINO,DESCRIPCION From CATALOGO_DESTINOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Destino.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Destino.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub FlatEdit_Institucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_INSTITUCION,DESCRIPCION From INSTITUCIONES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Institucion.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Institucion.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub FlatEdit_Linea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select CODIGO,DESCRIPCION From CATALOGO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = " and LINEA_INTERNA = 1 AND RETENCION = 'N' AND POLIZA = 'N'"
    frmBusquedas.Show vbModal
    FlatEdit_Linea.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Linea.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub Form_Activate()
 vModulo = 3
End Sub

Private Sub Form_Load()
 
 vModulo = 3
 lsw.ColumnHeaders.Add , , "", 100
 
opt.Item(0).Caption = "Saldos Menores a: " & Format(fxCrdParametro("23"), "Standard")
opt.Item(2).Caption = "Mora Menor a: " & Format(fxCrdParametro("24"), "Standard")
opt.Item(3).Caption = "Cta. Derivada Menor a: " & Format(fxCrdParametro("24.1"), "Standard")
 
End Sub


Private Sub sbCodigosDuplicados()

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "SELECT A.ID_SOLICITUD,A.CODIGO,A.CEDULA,A.ESTADO,A.PROCESO,A.OPEX,A.SALDO" _
       & " FROM REG_CREDITOS A" _
       & " WHERE A.ESTADO = 'A' AND (SELECT COUNT(*) FROM REG_CREDITOS" _
       & " WHERE CEDULA = A.CEDULA AND ESTADO = 'A' AND CODIGO = A.CODIGO AND PROCESO ='N') > 1" _
       & " ORDER BY CEDULA"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "Operación", 1100
lsw.ColumnHeaders.Add , , "Código", 840
lsw.ColumnHeaders.Add , , "Cédula", 1400
lsw.ColumnHeaders.Add , , "Estado", 1140, vbCenter
lsw.ColumnHeaders.Add , , "Proceso", 1100, vbCenter
lsw.ColumnHeaders.Add , , "Opex", 840, vbCenter
lsw.ColumnHeaders.Add , , "Saldo", 1200, vbRightJustify

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!id_solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = fxDescribeEstado(rs!Estado)
     itmX.SubItems(4) = fxProcesoOperacion(rs!Proceso)
     itmX.SubItems(5) = IIf((rs!opex = 0), "NO", "SI")
     itmX.SubItems(6) = Format(rs!Saldo, "###,###,###,##0.00")
 rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbSaldosNegativos()
Dim pMonto As Currency
Dim vCasos As Long, vTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError


pMonto = 1000 'No se utiliza
vCasos = 0
vTotal = 0
       
strSQL = "select R.codigo,R.id_solicitud,R.cedula,S.nombre,R.saldo,R.opex,R.proceso,R.estado,R.estadosol" _
       & ",I.descripcion as 'Institucion', C.descripcion as 'LineaDesc',  isnull(D.descripcion,'') as 'Destino'" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
       & "  left join CATALOGO_DESTINOS D on R.cod_Destino = D.cod_Destino" _
       & " where R.estado in('A','C') and R.proceso = 'N' and R.saldo < 0" _
       & "  and R.cod_Divisa = 'COL'"

If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If

strSQL = strSQL & " Order by R.codigo, R.id_Solicitud"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "Línea", 840
lsw.ColumnHeaders.Add , , "Operación", 1100
lsw.ColumnHeaders.Add , , "Identificación", 1400
lsw.ColumnHeaders.Add , , "Nombre", 3400
lsw.ColumnHeaders.Add , , "Estado", 1140, vbCenter
lsw.ColumnHeaders.Add , , "Proceso", 1100, vbCenter
lsw.ColumnHeaders.Add , , "Opex", 840, vbCenter
lsw.ColumnHeaders.Add , , "Saldo", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Línea Desc.", 3400
lsw.ColumnHeaders.Add , , "Destino", 3400
lsw.ColumnHeaders.Add , , "Institución", 3400

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!id_solicitud
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = fxDescribeEstado(rs!Estado)
     itmX.SubItems(5) = fxProcesoOperacion(rs!Proceso)
     itmX.SubItems(6) = IIf((rs!opex = 0), "NO", "SI")
     itmX.SubItems(7) = Format(rs!Saldo, "Standard")
 
     itmX.SubItems(8) = rs!LineaDesc
     itmX.SubItems(9) = rs!Destino
     itmX.SubItems(10) = rs!Institucion
 
 vCasos = vCasos + 1
 vTotal = vTotal + rs!Saldo
 
 rs.MoveNext
Loop

rs.Close

txtMonto.Text = Format(vTotal, "Standard")
txtCasos.Text = Format(vCasos, "###,###,##0")

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbSaldoMenor()
Dim pMonto As Currency, pCuenta As String
Dim vCasos As Long, vTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError



pMonto = fxCrdParametro("23")
vCasos = 0
vTotal = 0

strSQL = "select R.codigo,R.id_solicitud,R.cedula,S.nombre,R.saldo,R.opex,R.proceso,R.estado,R.estadosol" _
       & ",I.descripcion as 'Institucion', C.descripcion as 'LineaDesc',  isnull(D.descripcion,'') as 'Destino'" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
       & "  left join CATALOGO_DESTINOS D on R.cod_Destino = D.cod_Destino" _
       & " where R.estado = 'A' and R.proceso = 'N' and R.saldo between 0 and " & pMonto _
       & "  and R.cod_Divisa = 'COL'"

If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If

strSQL = strSQL & " order by R.codigo,R.id_Solicitud"


Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "Línea", 840
lsw.ColumnHeaders.Add , , "Operación", 1100
lsw.ColumnHeaders.Add , , "Identificación", 1400
lsw.ColumnHeaders.Add , , "Nombre", 3400
lsw.ColumnHeaders.Add , , "Estado", 1140, vbCenter
lsw.ColumnHeaders.Add , , "Proceso", 1100, vbCenter
lsw.ColumnHeaders.Add , , "Opex", 840, vbCenter
lsw.ColumnHeaders.Add , , "Saldo", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Línea Desc.", 3400
lsw.ColumnHeaders.Add , , "Destino", 3400
lsw.ColumnHeaders.Add , , "Institución", 3400

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!id_solicitud
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = fxDescribeEstado(rs!Estado)
     itmX.SubItems(5) = fxProcesoOperacion(rs!Proceso)
     itmX.SubItems(6) = IIf((rs!opex = 0), "NO", "SI")
     itmX.SubItems(7) = Format(rs!Saldo, "Standard")
     
     itmX.SubItems(8) = rs!LineaDesc
     itmX.SubItems(9) = rs!Destino
     itmX.SubItems(10) = rs!Institucion
     
 vCasos = vCasos + 1
 vTotal = vTotal + rs!Saldo
 rs.MoveNext
Loop

rs.Close

txtMonto.Text = Format(vTotal, "Standard")
txtCasos.Text = Format(vCasos, "###,###,##0")

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbMoraMenor()
Dim pMonto As Currency, pCuenta As String
Dim vCasos As Long, vTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError



pMonto = fxCrdParametro("24")
vCasos = 0
vTotal = 0

strSQL = "select R.codigo,R.id_solicitud,R.cedula,S.Nombre,R.saldo,R.opex,R.proceso,R.estado,(M.intc + M.intm+ M.amortiza + M.Cargo) as MoraFinanciera" _
       & ",I.descripcion as 'Institucion', C.descripcion as 'LineaDesc',  isnull(D.descripcion,'') as 'Destino'" _
       & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
       & " inner join Morosidad M on R.id_solicitud = M.id_solicitud and M.estado = 'A' and R.cod_divisa = 'COL'" _
       & " inner join socios S on R.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
       & "  left join CATALOGO_DESTINOS D on R.cod_Destino = D.cod_Destino" _
       & " where R.estado = 'A' and R.proceso <> 'J' and (M.intc + M.intm+ M.amortiza + M.Cargo) between 0 and " & pMonto


If Len(FlatEdit_Linea.Text) > 0 Then
  strSQL = strSQL & " and R.Codigo = '" & FlatEdit_Linea.Text & "'"
End If

If Len(FlatEdit_Destino.Text) > 0 Then
  strSQL = strSQL & " and R.cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
  strSQL = strSQL & " and S.cod_institucion = " & FlatEdit_Institucion.Text
End If

strSQL = strSQL & " order by R.codigo,R.id_Solicitud"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear

lsw.ColumnHeaders.Add , , "Línea", 840
lsw.ColumnHeaders.Add , , "Operación", 1100
lsw.ColumnHeaders.Add , , "Identificación", 1400
lsw.ColumnHeaders.Add , , "Nombre", 3400
lsw.ColumnHeaders.Add , , "Estado", 1140, vbCenter
lsw.ColumnHeaders.Add , , "Proceso", 1100, vbCenter
lsw.ColumnHeaders.Add , , "Opex", 840, vbCenter
lsw.ColumnHeaders.Add , , "Mora.Línea", 1200, vbRightJustify

lsw.ColumnHeaders.Add , , "Línea Desc.", 3400
lsw.ColumnHeaders.Add , , "Destino", 3400
lsw.ColumnHeaders.Add , , "Institución", 3400


Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!id_solicitud
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = fxDescribeEstado(rs!Estado)
     itmX.SubItems(5) = fxProcesoOperacion(rs!Proceso)
     itmX.SubItems(6) = IIf((rs!opex = 0), "NO", "SI")
     itmX.SubItems(7) = Format(rs!MoraFinanciera, "Standard")
     
     itmX.SubItems(8) = rs!LineaDesc
     itmX.SubItems(9) = rs!Destino
     itmX.SubItems(10) = rs!Institucion
     
     
 vCasos = vCasos + 1
 vTotal = vTotal + rs!MoraFinanciera
 rs.MoveNext
Loop

rs.Close

txtMonto.Text = Format(vTotal, "Standard")
txtCasos.Text = Format(vCasos, "###,###,##0")

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCtaDerivada_Consulta()
Dim pMonto As Currency, pCuenta As String
Dim vCasos As Long, vTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

'--[spSys_Creditos_Clean_Ctas_Menores] 'Admin'

pMonto = fxCrdParametro("24.1")
vCasos = 0
vTotal = 0

strSQL = "select V.ID_SOLICITUD, V.CODIGO, R.CEDULA, S.NOMBRE, R.SALDO, V.NUM_CUOTA" _
       & " , (V.INTCOR+ V.INTMOR + V.CARGOS + V.POLIZA + V.PRINCIPAL) as 'Monto', C.Descripcion" _
       & " from CRD_OPERACION_TRANSAC V inner join CATALOGO C on V.CODIGO = C.CODIGO" _
       & "     and C.RETENCION = 'N' and C.POLIZA = 'N'" _
       & "     and C.LINEA_INTERNA = 1" _
       & " inner join REG_CREDITOS R ON V.ID_SOLICITUD = R.ID_SOLICITUD" _
       & " inner join CRD_GARANTIA_TIPOS Gt on R.GARANTIA = Gt.GARANTIA" _
       & " inner join Socios S on R.CEDULA = S.CEDULA" _
       & " Where V.NUM_CUOTA_MADRE > 0" _
       & " and R.ESTADO = 'A' and R.PROCESO <> 'J'" _
       & " and V.ESTADO = 'A' and V.NUM_CUOTA <> 0" _
       & " and (V.INTCOR+ V.INTMOR + V.CARGOS + V.POLIZA + V.PRINCIPAL) < " & pMonto _
       & " and R.COD_DIVISA = 'COL'" _
       & " order by V.ID_SOLICITUD"

Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear


lsw.ColumnHeaders.Add , , "Operación", 1100
lsw.ColumnHeaders.Add , , "Línea", 840, vbCenter
lsw.ColumnHeaders.Add , , "Identificación", 1400
lsw.ColumnHeaders.Add , , "Nombre", 3400
lsw.ColumnHeaders.Add , , "Saldo", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Cta.Id", 840, vbCenter
lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
lsw.ColumnHeaders.Add , , "Línea Desc.", 3400

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!id_solicitud)
     itmX.SubItems(1) = rs!Codigo
     itmX.SubItems(2) = rs!Cedula
     itmX.SubItems(3) = rs!Nombre
     itmX.SubItems(4) = Format(rs!Saldo, "Standard")
     itmX.SubItems(5) = rs!Num_Cuota
     itmX.SubItems(6) = Format(rs!Monto, "Standard")
     itmX.SubItems(7) = rs!Descripcion
     
 vCasos = vCasos + 1
 vTotal = vTotal + rs!Monto
 rs.MoveNext
Loop

rs.Close

txtMonto.Text = Format(vTotal, "Standard")
txtCasos.Text = Format(vCasos, "###,###,##0")

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCtaDerivada_Corrige()
Dim pMonto As Currency

On Error GoTo vError

If lsw.ListItems.Count <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "spSys_Creditos_Clean_Ctas_Menores '" & glogon.Usuario & "'"

Call OpenRecordSet(rs, strSQL)

    Call sbImprimeRecibo(rs!NumDoc, rs!TipoDoc)

rs.Close

Me.MousePointer = vbDefault

MsgBox "Cuotas Derivadas Aplicadas...", vbInformation

Call sbCtaDerivada_Consulta

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub opt_Click(Index As Integer)

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
txtCasos.Text = "0"
txtMonto.Text = "0.00"

'Solo para Saldos Menores y Saldos Negativos
'Poner la cuenta contable para afectación

Select Case Index
  Case 0, 3 'Saldos Menores a X
    txtCuenta.Text = fxCrdParametro("22")
  Case 1 'Saldos Negativos
    txtCuenta.Text = fxCrdParametro("21")
End Select

If Index < 3 Then
  txtDescripcion.Text = fxgCntCuentaDesc(txtCuenta.Text)
  txtCuenta.Text = fxgCntCuentaFormato(True, txtCuenta.Text, 0)
End If


Me.MousePointer = vbDefault

'Realiza la Consulta
Select Case True
  Case opt(0).Value  'Saldos Menores a X
    Call sbSaldoMenor
  Case opt(1).Value  'Saldos Negativos
    Call sbSaldosNegativos
  Case opt(2).Value  'Mora Menor a X
    Call sbMoraMenor
  Case opt(3).Value 'Cuotas Derivadas Menores a X
    Call sbCtaDerivada_Consulta
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0

Call opt_Click(0)


End Sub

Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta = gCuenta
    txtDescripcion = fxgCntCuentaDesc(gCuenta)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta = gCuenta
    txtDescripcion = fxgCntCuentaDesc(gCuenta)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
