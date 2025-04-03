VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmTES_AnulacionDoc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Anulación de Documentos"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8160
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1452
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   8292
      _Version        =   1310723
      _ExtentX        =   14626
      _ExtentY        =   2561
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
      Height          =   1692
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   3360
      Width           =   8292
      _Version        =   1310723
      _ExtentX        =   14626
      _ExtentY        =   2984
      _StockProps     =   79
      Caption         =   "Notas del cambio: "
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
      Left            =   1440
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
      Left            =   1440
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
      Left            =   1440
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
      Left            =   4080
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1092
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   5160
      Width           =   8292
      _Version        =   1310723
      _ExtentX        =   14626
      _ExtentY        =   1926
      _StockProps     =   79
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
      Begin XtremeSuiteControls.PushButton cmdBoleta 
         Height          =   645
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2355
         _ExtentY        =   1138
         _StockProps     =   79
         Caption         =   "Boleta"
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
         Picture         =   "frmTES_AnulacionDoc.frx":0000
      End
      Begin XtremeSuiteControls.PushButton cmdAnular 
         Height          =   645
         Left            =   6720
         TabIndex        =   13
         Top             =   360
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2355
         _ExtentY        =   1138
         _StockProps     =   79
         Caption         =   "Anular"
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
         Picture         =   "frmTES_AnulacionDoc.frx":07BC
      End
      Begin XtremeSuiteControls.CheckBox chkAplicaCopiaEsquema 
         Height          =   252
         Left            =   480
         TabIndex        =   14
         Top             =   480
         Width           =   2412
         _Version        =   1310723
         _ExtentX        =   4254
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Copia de Solicitud"
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
      End
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
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
      TabIndex        =   0
      Top             =   600
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1452
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmTES_AnulacionDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vFechaEmision As Date

Private Sub cmdAnular_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Anula un Documento ya emitido y actualiza saldos del Banco.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'               LimpiaObjetos - (Limpia los objetos de entrada de datos)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, pCopia As Integer, pNotas As String


If txtVerifica.Tag <> "S" Then
   MsgBox "Identifique las notas de la verificación antes de Anular...!!!", vbExclamation
   Exit Sub
End If

If Len(txtNotas.Text) = 0 Then
   MsgBox "Identifique una Nota válida para realizar el movimiento!", vbExclamation
   Exit Sub
End If

strSQL = MsgBox("Confirma Anulacion?", vbExclamation + vbYesNo + vbDefaultButton2)
If strSQL = vbNo Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


pNotas = Mid(fxSysCleanTxtInject(txtNotas.Text), 1, 500)

If chkAplicaCopiaEsquema.Enabled And chkAplicaCopiaEsquema.Value = vbChecked Then
 pCopia = 1
Else
 pCopia = 0
End If


strSQL = "exec spTES_Transaccion_Anula " & txtSolicitud.Text & ", '" & pNotas & "',  '" & glogon.Usuario & "', " & pCopia
Call ConectionExecute(strSQL)

Call Bitacora("Anula", "Anula Solicitud :" & txtSolicitud.Text)

Me.MousePointer = vbDefault
MsgBox "Anulación Generada", vbInformation

Call cmdBoleta_Click

Call TimerX_Timer

Me.MousePointer = vbDefault


Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdBoleta_Click()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_BoletaRegistro.rpt")
    .SelectionFormula = "{CHEQUES.NSOLICITUD} = " & txtSolicitud.Text
        
    .SubreportToChange = "sbDetalle"

    .StoredProcParam(0) = txtSolicitud.Text
        
    .PrintReport
End With

Me.MousePointer = vbDefault


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
       & ",T.descripcion as TipoDocX,C.detalle_Anulacion,C.Estado_Asiento,C.Fecha_emision" _
       & " from Tes_Transacciones C inner join Tes_Bancos B on C.id_banco = B.id_Banco" _
       & " inner join  tes_tipos_doc T on C.tipo = T.tipo" _
       & " where C.nsolicitud = " & GLOBALES.gTag
Call OpenRecordSet(rs, strSQL)
    txtSolicitud.Text = rs!NSolicitud
    txtSolicitud.Tag = rs!Estado
    
    txtDocumento.Text = rs!nDocumento & ""
    
    txtBanco.Tag = rs!ID_BANCO
    txtBanco.Text = rs!BancoX
    
    txtTipo.Tag = rs!Tipo
    txtTipo.Text = rs!TipoDocX
    
    txtNotas.Text = rs!detalle_anulacion & ""
    txtNotas.Tag = rs!estado_asiento & ""
    vFechaEmision = rs!Fecha_Emision
rs.Close

'Seccion de Verificación
txtVerifica.Tag = "S"

If Not fxTesTipoAccesoValida(txtBanco.Tag, glogon.Usuario, txtTipo.Tag, "N") Then
 txtVerifica = txtVerifica & vbCrLf & " - El Usuario Actual no esta Autorizado a Anular este Tipo de Documento..."
 txtVerifica.Tag = "N"
End If

Select Case txtSolicitud.Tag
   Case "P"
     txtVerifica = txtVerifica & vbCrLf & " - La solicitud no se puede anular por que se encuentra pendiente de emision..."
     txtVerifica.Tag = "N"
   Case "A"
     txtVerifica = txtVerifica & vbCrLf & " - El documento ya se encuentra anulado..."
     txtVerifica.Tag = "N"
End Select
'Fin de Verificacion


If txtVerifica.Tag = "S" Then
   txtVerifica.Text = "----> Este Documento puede ser Anulado."
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


Private Sub sbCopiaEsquemaSolicitud()

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Duplica una determinada solicitud ya ingresada a Tesoreria. Tambien duplica
'               el detalle de la misma solicitud para la nueva.
'REFERENCIAS:   Bitacora - (Registra movimientos sobre la Base de Datos)
'
'               fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim rs As New ADODB.Recordset, strSQL As String
Dim lngNewSol As Long, pNotas As String

On Error GoTo vError

Me.MousePointer = vbHourglass

pNotas = fxSysCleanTxtInject(txtNotas.Text)
    
strSQL = "exec spTES_TE_Transaccion_Copia " & txtSolicitud.Text & ", '" & pNotas & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

If Not glogon.error Then
    lngNewSol = rs!TesoreriaId
Else
    lngNewSol = 0
End If

'Bitacoras
Call Bitacora("Aplica", "Copia Solicitud : " & txtSolicitud & " A la Sol : " & lngNewSol)

Me.MousePointer = vbDefault

MsgBox "Copia Realizada, NUEVA SOLICITUD GENERADA : " & lngNewSol, vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

