VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmTES_CambioFechas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio de Fechas"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   8415
      _Version        =   1310723
      _ExtentX        =   14843
      _ExtentY        =   2566
      _StockProps     =   79
      Caption         =   "Cambio de fechas: "
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
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2052
         _Version        =   1310723
         _ExtentX        =   3625
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.DateTimePicker dtpActual 
         Height          =   312
         Left            =   2160
         TabIndex        =   8
         Top             =   720
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
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
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpNueva 
         Height          =   312
         Left            =   3600
         TabIndex        =   9
         Top             =   720
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
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
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton cmdBoleta 
         Height          =   645
         Left            =   5520
         TabIndex        =   10
         Top             =   600
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
         Picture         =   "frmTES_CambioFechas.frx":0000
      End
      Begin XtremeSuiteControls.PushButton cmdCambiar 
         Height          =   645
         Left            =   6840
         TabIndex        =   11
         Top             =   600
         Width           =   1335
         _Version        =   1310723
         _ExtentX        =   2355
         _ExtentY        =   1138
         _StockProps     =   79
         Caption         =   "Cambiar"
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
         Picture         =   "frmTES_CambioFechas.frx":07BC
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   6
         Left            =   2160
         TabIndex        =   6
         Top             =   360
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   7
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   1332
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8040
      Top             =   0
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1455
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1680
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
         TabIndex        =   14
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
      TabIndex        =   12
      Top             =   3360
      Width           =   8415
      _Version        =   1310723
      _ExtentX        =   14843
      _ExtentY        =   2990
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1212
         Left            =   600
         TabIndex        =   13
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
      TabIndex        =   15
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
      TabIndex        =   16
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
      TabIndex        =   17
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
      TabIndex        =   18
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
      Width           =   1212
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
      Width           =   1212
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
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   1572
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   12852
   End
End
Attribute VB_Name = "frmTES_CambioFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub cmdCambiar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As String


If txtVerifica.Tag <> "S" Then
   MsgBox "Identifique las notas de la verificación antes de cambiar la fecha...!!!", vbExclamation
   Exit Sub
End If

strSQL = MsgBox("Confirma Cambio de Fecha?", vbExclamation + vbYesNo + vbDefaultButton2)
If strSQL = vbNo Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

glogon.Conection.BeginTrans

Select Case Mid(cbo.Text, 1, 1)
'If Mid(cbo.Text, 1, 1) = "S" Then
 Case "S"
 'Solicitud
    strSQL = "Update Tes_Transacciones Set Fecha_Solicitud ='" & Format(dtpNueva.Value, "yyyy/mm/dd") _
           & "' Where NSolicitud = " & txtSolicitud.Text
    Call ConectionExecute(strSQL)
     
    strSQL = "Cambia Fecha Solicitud de " & Format(dtpActual.Value, "dd/mm/yyyy") _
           & " a " & Format(dtpNueva.Value, "dd/mm/yyyy") & " /Nota: " & txtNotas
     
    Call sbTesBitacoraEspecial(txtSolicitud.Text, "08", strSQL)
     
    Call Bitacora("Modifica", "Fecha Solicitud Sol:" & txtSolicitud.Text & " de " & Format(dtpActual.Value, "dd/mm/yyyy") _
                    & " a " & Format(dtpNueva.Value, "dd/mm/yyyy"))

 Case "E"
     'Emision
      strSQL = "Update Tes_Transacciones Set Fecha_Emision ='" & Format(dtpNueva.Value, "yyyy/mm/dd") _
             & "' Where NSolicitud = " & txtSolicitud.Text
      Call ConectionExecute(strSQL)
    
      strSQL = "Cambia Fecha Emisión de " & Format(dtpActual.Value, "dd/mm/yyyy") _
             & " a " & Format(dtpNueva.Value, "dd/mm/yyyy") & " /Nota: " & txtNotas
       
      Call sbTesBitacoraEspecial(txtSolicitud.Text, "08", strSQL)
    
      Call Bitacora("Modifica", "Fecha Emision Sol:" & txtSolicitud.Text & " de " & Format(dtpActual.Value, "dd/mm/yyyy") _
                      & " a " & Format(dtpNueva.Value, "dd/mm/yyyy"))
Case "A"
     'Anulado
      strSQL = "Update Tes_Transacciones Set Fecha_Anula ='" & Format(dtpNueva.Value, "yyyy/mm/dd") _
             & "' Where NSolicitud = " & txtSolicitud.Text
      Call ConectionExecute(strSQL)
    
      strSQL = "Cambia Fecha Anulacion  de " & Format(dtpActual.Value, "dd/mm/yyyy") _
             & " a " & Format(dtpNueva.Value, "dd/mm/yyyy") & " /Nota: " & txtNotas
       
      Call sbTesBitacoraEspecial(txtSolicitud.Text, "08", strSQL)
    
      Call Bitacora("Modifica", "Fecha Anulacion Sol:" & txtSolicitud.Text & " de " & Format(dtpActual.Value, "dd/mm/yyyy") _
                      & " a " & Format(dtpNueva.Value, "dd/mm/yyyy"))

End Select
'End If


glogon.Conection.CommitTrans

MsgBox "Cambio de Fechas Realizado Satisfactoriamente...", vbInformation

Call TimerX_Timer
Me.MousePointer = vbDefault

Exit Sub

vError:
   Me.MousePointer = vbDefault
   glogon.Conection.RollbackTrans
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
        
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

txtVerifica.Tag = "N"
txtVerifica.Text = ""

strSQL = "select estado,fecha_emision,fecha_solicitud,fecha_anula from Tes_Transacciones where nsolicitud = " & txtSolicitud.Text
Call OpenRecordSet(rs, strSQL)
Select Case Mid(cbo.Text, 1, 1)
   Case "S"
    'Solicitud
    If rs!Estado = "P" Then
        txtVerifica = txtVerifica & vbCrLf & " ----> Se puede cambiar la fecha de la solicitud"
        txtVerifica.Tag = "S"
    Else
        txtVerifica = txtVerifica & vbCrLf & " - La solicitud no se le puede cambiar la fecha de solicitud, porque no se encuentra solicitada..."
        txtVerifica.Tag = "N"
    End If
    
    If Not IsNull(rs!fecha_solicitud) Then
      dtpActual.Value = rs!fecha_solicitud
      dtpNueva.Value = rs!fecha_solicitud
    End If
 
Case "E"
 'Emision
     If rs!Estado = "I" Or rs!Estado = "T" Then
         txtVerifica = txtVerifica & vbCrLf & " -----> Se puede cambiar la fecha de emisión del documento"
         txtVerifica.Tag = "S"
     Else
         txtVerifica = txtVerifica & vbCrLf & " - No se puede cambiar la fecha de emisión del documento, porque no se encuentra en estado de emitido..."
         txtVerifica.Tag = "N"
     End If
    
     If Not IsNull(rs!Fecha_Emision) Then
       dtpActual.Value = rs!Fecha_Emision
       dtpNueva.Value = rs!Fecha_Emision
     End If
    
Case "A"
 'Anulación
     If rs!Estado = "A" Then
         txtVerifica = txtVerifica & vbCrLf & " -----> Se puede cambiar la fecha de anulación del documento"
         txtVerifica.Tag = "S"
     Else
         txtVerifica = txtVerifica & vbCrLf & " - No se puede cambiar la fecha de anulación del documento, porque no se encuentra en estado de anulado..."
         txtVerifica.Tag = "N"
     End If
    
     If Not IsNull(rs!Fecha_Anula) Then
       dtpActual.Value = rs!Fecha_Anula
       dtpNueva.Value = rs!Fecha_Anula
     End If
     
End Select
    

rs.Close

'Fin de Verificacion


If txtVerifica.Tag = "S" Then
   txtVerifica.ForeColor = vbBlue
Else
   txtVerifica.ForeColor = vbRed
End If

End Sub

Private Sub Form_Load()

vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

vPaso = True
    cbo.AddItem "Solicitud"
    cbo.AddItem "Emisión"
    cbo.AddItem "Anulado"
    cbo.Text = "Solicitud"
vPaso = False

End Sub



Private Sub TimerX_Timer()
Dim strSQL As String, rs As New ADODB.Recordset

TimerX.Interval = 0

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.Nsolicitud,C.tipo,C.estado,C.ndocumento,C.id_banco,B.descripcion as BancoX" _
       & ",T.descripcion as TipoDocX,C.detalle_Anulacion,C.Estado_Asiento" _
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
rs.Close

'Seccion de Verificación
Call cbo_Click

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
 txtVerifica.Text = fxSys_Error_Handler(Err.Description)
 txtVerifica.ForeColor = vbRed
 txtVerifica.Tag = "N"
 
End Sub

