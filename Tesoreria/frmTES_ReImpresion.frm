VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmTES_ReImpresion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ReImpresión de Cheques Continuos"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   8415
      _Version        =   1310723
      _ExtentX        =   14843
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Autorización"
      ForeColor       =   4210752
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   1560
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   360
         Width           =   2292
         _Version        =   1310723
         _ExtentX        =   4043
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
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtContraseña 
         Height          =   312
         Left            =   5880
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   2292
         _Version        =   1310723
         _ExtentX        =   4043
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
         Alignment       =   2
         PasswordChar    =   "*"
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdReImprimir 
         Height          =   648
         Left            =   5880
         TabIndex        =   16
         Top             =   840
         Width           =   2292
         _Version        =   1310723
         _ExtentX        =   4043
         _ExtentY        =   1143
         _StockProps     =   79
         Caption         =   "&Re-Imprime Documento"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmTES_ReImpresion.frx":0000
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Clave"
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
         Left            =   4800
         TabIndex        =   13
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
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
         Left            =   720
         TabIndex        =   12
         Top             =   360
         Width           =   852
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8160
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   8415
      _Version        =   1310723
      _ExtentX        =   14843
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Verifica el cambio: "
      ForeColor       =   4210752
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtVerifica 
         Height          =   1092
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   7572
         _Version        =   1310723
         _ExtentX        =   13356
         _ExtentY        =   1926
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
         Appearance      =   2
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1695
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   8415
      _Version        =   1310723
      _ExtentX        =   14843
      _ExtentY        =   2990
      _StockProps     =   79
      Caption         =   "Notas para la re-impresión:"
      ForeColor       =   4210752
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1212
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   7572
         _Version        =   1310723
         _ExtentX        =   13356
         _ExtentY        =   2138
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   2
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtSolicitud 
      Height          =   372
      Left            =   1560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   2652
      _Version        =   1310723
      _ExtentX        =   4678
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   312
      Left            =   1560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   960
      Width           =   2652
      _Version        =   1310723
      _ExtentX        =   4678
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtBanco 
      Height          =   312
      Left            =   1560
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   600
      Width           =   7092
      _Version        =   1310723
      _ExtentX        =   12509
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipo 
      Height          =   312
      Left            =   4200
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   4452
      _Version        =   1310723
      _ExtentX        =   7853
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitud"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image imgBanner 
      Height          =   1452
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmTES_ReImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdReImprimir_Click()
Dim strSQL As String, rs As New ADODB.Recordset


If txtVerifica.Tag <> "S" Then
   MsgBox "Identifique las notas de la verificación antes de ReImprimir...!!!", vbExclamation
   Exit Sub
End If

'Verificar Usuarios y Claves de Autorización
strSQL = "select isnull(count(*),0) as Existe from tes_autorizaciones where nombre = '" _
       & txtUsuario & "' and estado = 'A' and clave = '" & fxTESCifrado(txtContraseña) & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  rs.Close
  MsgBox "El usuario y clave de autorización no concuerda con ninguno de los registrados, verifique...", vbExclamation
  Exit Sub
End If
rs.Close

strSQL = MsgBox("Confirma ReImpresión?", vbExclamation + vbYesNo + vbDefaultButton2)
If strSQL = vbNo Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbReImprime

strSQL = "insert tes_ReImpresiones(nsolicitud,fecha,usuario,autoriza,notas) values(" _
       & txtSolicitud.Text & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtUsuario & "','" & Mid(Trim(txtNotas), 1, 100) & "')"
Call ConectionExecute(strSQL)

Call sbTesBitacoraEspecial(txtSolicitud.Text, "17", Mid(txtNotas.Text, 1, 150))
Call Bitacora("Aplica", "ReImpresión de Solicitud :" & txtSolicitud.Text)

MsgBox "ReImpresión Generada", vbInformation

txtUsuario = ""
txtContraseña = ""

Me.MousePointer = vbDefault

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReImprime()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vBanco As Integer, vTipo As String, x As New clsImpresoras
Dim vFirmaDesde As Currency, vFirmaHasta As Currency, vLugarEmision As String
Dim strDec As String, curMonto As Currency, vFirmas As Boolean

vBanco = txtBanco.Tag
vTipo = txtTipo.Tag

Call sbCargaArchivosEspeciales(vBanco)

strSQL = "select firmas_desde,firmas_hasta,formato_transferencia,Lugar_Emision  from Tes_Bancos where id_banco = " & vBanco
Call OpenRecordSet(rs, strSQL)
    vFirmaDesde = rs!firmas_desde
    vFirmaHasta = rs!firmas_hasta
    vLugarEmision = Trim(rs!Lugar_Emision & "")
rs.Close


strSQL = "select isnull(count(*),0) as Existe from TES_BANCO_FIRMASAUT where id_Banco = " & vBanco _
       & " and usuario = '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
    vFirmas = IIf((rs!Existe = 0), False, True)
rs.Close


strSQL = "select * from Tes_Transacciones where nsolicitud = " & txtSolicitud.Text
Call OpenRecordSet(rs, strSQL)
With frmContenedor.Crt
    .Reset
    
    x.TipoImpresora = Cheques
    x.Reset
    
    .PrinterDriver = x.Controlador
    .PrinterName = x.Nombre
    .PrinterPort = x.Puerto
    
    .Connect = glogon.ConectRPT
    
    If vLugarEmision <> "" Then
       vLugarEmision = vLugarEmision & ", "
    End If
   
    .Formulas(0) = "Fecha='" & vLugarEmision & Day(rs!Fecha_Emision) & " DE " & fxTesMesDescripcion(rs!Fecha_Emision) & " DE " & Year(rs!Fecha_Emision) & "'"
    .Formulas(1) = "Año='" & Year(rs!Fecha_Emision) & "'"
    
    '*******Codigo Nuevo para Monto en Letras 2003/03/21
    strDec = Format(rs!Monto, "##################.00")
    strDec = Trim(strDec)
    strDec = Mid(strDec, Len(strDec) - 1, 2)
    
    curMonto = Mid(Format(rs!Monto, "#################0.00"), 1, Len(Format(rs!Monto, "#################0.00")) - 3)
    .Formulas(2) = "Letras='**" & Trim(UCase(Conversion(CStr(curMonto))))
    
    If Trim(strDec) <> "00" Then
       .Formulas(2) = .Formulas(2) & UCase(" Con " & Trim(strDec) & "/100 " & fxDescDivisa(rs!cod_Divisa) & "**'")
    Else
       .Formulas(2) = .Formulas(2) & " " & UCase(fxDescDivisa(rs!cod_Divisa)) & "**'"
    End If
    '********** Fin de la Modificacion del Monto en Letras
    'strChequesFirmas = "TesDocFormat01.rpt" 'Reporte con Firmas
    'strChequesSinFirmas = "TesDocFormat02.rpt" 'Reporte sin Firmas
        
    If vFirmas Then
        If rs!Monto >= vFirmaDesde And rs!Monto <= vFirmaHasta Then
           .ReportFileName = SIFGlobal.fxPathReportes(strChequesFirmas)
        Else
           .ReportFileName = SIFGlobal.fxPathReportes(strChequesSinFirmas)
        End If
    Else
       .ReportFileName = SIFGlobal.fxPathReportes(strChequesSinFirmas)
    End If
    
    .SelectionFormula = "{CHEQUES.NSOLICITUD}=" & rs!NSolicitud
    .Destination = crptToPrinter
    .PrintReport
End With

rs.Close

                  
End Sub

Private Sub Form_Activate()
 vModulo = 9

End Sub

Private Sub Form_Load()

vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

TimerX.Interval = 0

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select C.Nsolicitud,C.tipo,C.estado,C.ndocumento,C.id_banco,B.descripcion as BancoX" _
       & ",T.descripcion as TipoDocX,C.detalle_Anulacion,C.Estado_Asiento,Y.comprobante" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
       & " inner join tes_tipos_doc T on C.tipo = T.tipo" _
       & " inner join tes_banco_docs Y on C.id_banco = Y.id_Banco and C.tipo = Y.tipo" _
       & " where C.nsolicitud = " & GLOBALES.gTag
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    'Seccion de Verificación
    txtVerifica.Tag = "N"
    txtVerifica.Text = "Este documento no es valido..."
    rs.Close
    Exit Sub

Else
    txtSolicitud.Text = rs!NSolicitud
    txtSolicitud.Tag = rs!Estado
    
    txtDocumento = rs!nDocumento & ""
    txtDocumento.Tag = rs!comprobante
    
    txtBanco.Tag = rs!ID_BANCO
    txtBanco.Text = rs!BancoX
    
    txtTipo.Tag = rs!Tipo
    txtTipo.Text = rs!TipoDocX
    
    txtNotas.Text = rs!detalle_anulacion & ""
    txtNotas.Tag = rs!estado_asiento & ""
End If
rs.Close

'Seccion de Verificación
txtVerifica.Tag = "S"


If Trim(txtDocumento.Tag) <> "01" Then
 txtVerifica = txtVerifica & vbCrLf & " - El Documento Actual no se puede ReImprimir, porque no es Cheque Continuo..."
 txtVerifica.Tag = "N"
End If


If txtSolicitud.Tag <> "I" Then
     txtVerifica = txtVerifica & vbCrLf & " - El documento no se encuentra Impreso / No se puede ReImprimir..."
     txtVerifica.Tag = "N"
End If
'Fin de Verificacion


If txtVerifica.Tag = "S" Then
   txtVerifica.Text = "----> Este Documento se puede ReImprimir"
   txtVerifica.ForeColor = vbBlue
Else
   txtVerifica.ForeColor = vbRed
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 txtVerifica.Text = fxSys_Error_Handler(Err.Description)
 txtVerifica.ForeColor = vbRed
 txtVerifica.Tag = "N"
 
End Sub

