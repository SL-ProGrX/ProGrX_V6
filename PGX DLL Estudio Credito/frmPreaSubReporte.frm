VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPreaSubReporte 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkDeducciones 
      Height          =   255
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Detalle de Deducciones?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      Top             =   4560
      Width           =   7575
      _Version        =   1572864
      _ExtentX        =   13361
      _ExtentY        =   2355
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   5760
         TabIndex        =   2
         Top             =   360
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Appearance      =   21
         Picture         =   "frmPreaSubReporte.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkImpresora 
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         Top             =   360
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Salida a Impresora"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Value           =   1
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.CheckBox chkResumen 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   1440
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ficha Resumen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkDetalle 
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ficha Detallada"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFichaConvenio 
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ficha para Convenios"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.CheckBox chkEstadoCuenta 
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado de Cuenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.CheckBox chkSubExpedienteResumen 
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ficha Resumen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkSubExpedienteDetalle 
      Height          =   255
      Left            =   2040
      TabIndex        =   9
      Top             =   3720
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Ficha Detallada"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.CheckBox chkSubExpedienteEstado 
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   4080
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado de Cuenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
   End
   Begin XtremeSuiteControls.CheckBox chkSubExpediente 
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   3000
      Width           =   3015
      _Version        =   1572864
      _ExtentX        =   5313
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Imprime Expedientes Asociados"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   21
      Value           =   1
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Informes"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   5292
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPreaSubReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad

Private Sub chkFichaConvenio_Click()
If chkFichaConvenio.Value = vbChecked Then
   chkDetalle.Value = vbUnchecked
   chkEstadoCuenta.Value = vbUnchecked
   chkResumen.Value = vbUnchecked
   chkSubExpediente.Value = vbUnchecked
   Call chkSubExpediente_Click
Else
   chkDetalle.Value = vbChecked
   chkEstadoCuenta.Value = vbChecked
   chkResumen.Value = vbChecked
   chkSubExpediente.Value = vbChecked
   Call chkSubExpediente_Click
End If

End Sub

Private Sub chkSubExpediente_Click()

If chkSubExpediente.Value = vbUnchecked Then
    
    chkSubExpedienteEstado.Enabled = False
    chkSubExpedienteDetalle.Enabled = False
    chkSubExpedienteResumen.Enabled = False
    chkSubExpedienteDetalle.Value = 0
    chkSubExpedienteEstado.Value = 0

Else
'    chkSubExpedienteEstado.Enabled = True
'    chkSubExpedienteDetalle.Enabled = True
    chkSubExpedienteResumen.Enabled = True
    chkSubExpedienteDetalle.Value = 0
    chkSubExpedienteEstado.Value = 0
    
End If

End Sub


Private Sub cmdReporte_Click()
Dim strSQL As String, rs As ADODB.Recordset
Dim vFechaIng As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = " select P.CEDULA, isnull(S.ESTADOACTUAL, 'N') as 'ESTADOACTUAL',  isnull(S.FECHAINGRESO, P.FECHA_CREACION) as 'FECHAINGRESO', dbo.MyGetdate() as 'FechaServer'  " _
       & ", P.COD_PREANALISIS, P.COD_PREANALISIS_REF, P.FECHA_NACIMIENTO" _
       & " from CRD_PREA_PREANALISIS P left join SOCIOS S on P.CEDULA = S.CEDULA" _
       & " where P.COD_PREANALISIS = '" & gPreAnalisis.Expediente & "'"
                    
    'Recuperar Cuotas entre el # de Fiadores de cada operacion
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
    vFechaIng = IIf(IsNull(rs!FechaIngreso), rs!FechaServer, rs!FechaIngreso)
End If

Call Imprimir(gPreAnalisis.Expediente, False, vFechaIng)



'Imprimir sub Expedientes
If (chkSubExpediente.Enabled) And (chkSubExpediente.Value = vbChecked) Then
    'Expediente maestro
                   
    strSQL = " select P.CEDULA, isnull(S.ESTADOACTUAL, 'N') as 'ESTADOACTUAL',  isnull(S.FECHAINGRESO, P.FECHA_CREACION) as 'FECHAINGRESO', dbo.MyGetdate() as 'FechaServer'  " _
           & ", P.COD_PREANALISIS, P.COD_PREANALISIS_REF, P.FECHA_NACIMIENTO" _
           & " from CRD_PREA_PREANALISIS P left join SOCIOS S on P.CEDULA = S.CEDULA" _
           & " where P.COD_PREANALISIS_REF = '" & gPreAnalisis.Expediente & "'"
                    
                    
    'Recuperar Cuotas entre el # de Fiadores de cada operacion
    Call OpenRecordSet(rs, strSQL)
    If Not glogon.error Then
       Do While Not rs.EOF
        
            vFechaIng = IIf(IsNull(rs!FechaIngreso), rs!FechaServer, rs!FechaIngreso)
            
            If rs!EstadoActual <> "S" Then
                vFechaIng = ""
            End If
            
            Call Imprimir(rs!cod_preanalisis, True, vFechaIng, rs!fecha_nacimiento)
            rs.MoveNext
        Loop
        rs.Close
    End If

End If 'SubExpedientes

salir:
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    Resume salir
End Sub


Private Sub Imprimir(ByVal pExpediente As String, ByVal pImpSubExpediente As Boolean, _
                     ByVal pFechaIngreso As String, Optional ByVal pFechaNacimiento As String)
 
Dim vEdadAnos As Integer, vEdadMax As Integer, vEdad As String, vEdadResultado As String
Dim vMembresia As String
Dim frm As Form


On Error GoTo vError


Me.MousePointer = vbHourglass

'Call sbFormActivo("frmPreaEstudio", frm)

For Each frm In forms
   If UCase(Trim(frm.Name)) = UCase(Trim("frmPreaEstudio")) Then
        Exit For
   End If
Next


If pImpSubExpediente Then
    vEdadAnos = frm.fxCalculaEdadAnos(pFechaNacimiento, "A") 'frmPreaEstudio.lblEdad.Caption
Else
    vEdadAnos = frm.fxCalculaEdadAnos(frm.dtpFecNac.Value, "A") 'frmPreaEstudio.lblEdad.Caption
End If

vEdad = frm.lblEdad.Caption


Select Case frm.fxSexoItemData(frm.cboSexo.ListIndex)
   Case "M"
        vEdadMax = GlobalEdadMaximaPermitidaHombre
   Case "F"
        vEdadMax = GlobalEdadMaximaPermitidaMujeres
End Select

If (vEdadAnos + (Val(frm.txtPlazo.Text) / 12)) >= vEdadMax Then
    vEdadResultado = " >> La edad supera el límite autorizado << "
Else
    vEdadResultado = " >> La edad es satisfactoria << "
End If


If pFechaIngreso <> "" Then
    vMembresia = fxMembresia(CDate(pFechaIngreso))
Else
    vMembresia = "NADA"
End If


With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = False
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Estudio de Crédito"
      
    .Connect = glogon.ConectRPT
      
      
    If chkImpresora.Value = vbChecked Then
     .Destination = crptToPrinter
    End If
    
    .Formulas(1) = "Fecha = 'Fecha :" & fxFechaServidor & "'"
    .Formulas(2) = "Usuario = 'Usuario :" & glogon.Usuario & "'"
    .Formulas(3) = "Empresa = ''"
    .Formulas(4) = "EdadCalculada = '" & vEdad & "'"
    .Formulas(5) = "ResultadoEdad = '" & vEdadResultado & "'"
    .Formulas(6) = "EstadoPreanalisis = '" & frm.lblEstado.Caption & "'"
    
    
    
    If Not pImpSubExpediente Then
        'Expediente del Deudor
        
        'Ficha Resumen
        If chkResumen.Value = vbChecked Then
            .SelectionFormula = "{CRD_PREA_PREANALISIS.COD_PREANALISIS} = '" & pExpediente & "'"
            
            If chkDeducciones.Value = xtpChecked Then
                .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_FichaResumenWsec.rpt")
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_FichaResumen.rpt")
            End If
            
            .PrintReport
        End If
        
        'Ficha Convenio
        If chkFichaConvenio.Value = vbChecked Then
            .Formulas(7) = "fxMembresia = '" & vMembresia & "'"
            .SelectionFormula = "{CRD_PREA_PREANALISIS.COD_PREANALISIS} = '" & pExpediente & "'"

            .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_FichaConvenio.rpt")
            
            .SubreportToChange = "DatosCreditos"
            .StoredProcParam(0) = pExpediente
            

            .PrintReport
        End If
        
        'Ficha detallada
        If chkDetalle.Value = vbChecked Then
            .SelectionFormula = ""
            .ReportFileName = ""
            .Formulas(4) = ""
            .Formulas(5) = ""
            .Formulas(7) = "Membresia = '" & vMembresia & "'"
            .StoredProcParam(0) = pExpediente
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_FichaDetalle.rpt")
                        
            DoEvents
            .PrintReport
        End If
        

        
        
        If chkEstadoCuenta.Value = vbChecked Then
            '.ReportFileName = SIFGlobal.fxPathReportes ("CrdPreaEstadoCuenta.rpt")
            '.PrintReport
        End If
        
    
    
    Else
        'Expediente de los Fiadores/CoDeudores
   
        If chkSubExpedienteResumen.Value = vbChecked Then
            .SelectionFormula = "{CRD_PREA_PREANALISIS.COD_PREANALISIS} = '" & pExpediente & "'"
            .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_FichaResumen.rpt")
            .PrintReport
        End If
        
        If chkSubExpedienteDetalle.Value = vbChecked Then
            .SelectionFormula = ""
            .ReportFileName = ""
            .Formulas(4) = ""
            .Formulas(5) = ""
            .Formulas(7) = "Membresia = '" & vMembresia & "'"
            .StoredProcParam(0) = pExpediente
             .ReportFileName = SIFGlobal.fxPathReportes("Credito_Analisis_FichaDetalle.rpt")
             DoEvents
            .PrintReport
        End If
        
        If chkSubExpedienteEstado.Value = vbChecked Then
            '.ReportFileName = SIFGlobal.fxPathReportes ("CrdPreaEstadoCuenta.rpt")
            '.PrintReport
        End If
        
    End If
End With
 
 Me.MousePointer = vbDefault
 
salir:

    Me.MousePointer = vbDefault
    frmContenedor.Crt.Formulas(1) = ""
    frmContenedor.Crt.Formulas(2) = ""
    frmContenedor.Crt.Formulas(3) = ""
    frmContenedor.Crt.Formulas(4) = ""
    frmContenedor.Crt.Formulas(5) = ""
    frmContenedor.Crt.Formulas(6) = ""
    frmContenedor.Crt.Formulas(7) = ""
    frmContenedor.Crt.SelectionFormula = ""
    
    Exit Sub
    
vError:
    MsgBox "Ocurrió un error al imprimir los reportes solicitados - " & Err.Description, vbError, gMsgTitulo
    Resume salir

End Sub


Private Sub Form_Load()

Me.Caption = "Expediente : " & gPreAnalisis.Expediente

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

If InStr(1, gPreAnalisis.Expediente, "-", vbTextCompare) = 0 Then
    chkResumen.Enabled = True
    chkDetalle.Enabled = True
    chkEstadoCuenta.Enabled = True
    chkSubExpediente.Enabled = True
    
Else
    chkSubExpediente.Enabled = False
    chkSubExpediente.Value = vbUnchecked
    chkSubExpedienteResumen.Value = vbUnchecked
    chkResumen.Value = vbUnchecked
    chkDetalle.Value = vbUnchecked
    chkEstadoCuenta.Value = vbUnchecked
    
End If

chkEstadoCuenta.Enabled = False
chkSubExpedienteDetalle.Enabled = False
chkSubExpedienteEstado.Enabled = False

Call chkSubExpediente_Click

End Sub
