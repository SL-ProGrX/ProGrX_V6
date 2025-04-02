VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
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
      Height          =   252
      Left            =   2880
      TabIndex        =   12
      Top             =   1560
      Width           =   3012
      _Version        =   1310722
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1332
      Left            =   120
      TabIndex        =   9
      Top             =   4560
      Width           =   7332
      _Version        =   1310722
      _ExtentX        =   12933
      _ExtentY        =   2350
      _StockProps     =   79
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.CheckBox chkImpresora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Salida a Impresora"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   3120
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   5760
         TabIndex        =   10
         Top             =   360
         Width           =   1452
         _Version        =   1310722
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmPreaSubReporte.frx":0000
      End
   End
   Begin VB.CheckBox chkFichaConvenio 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha para Convenios"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   5055
   End
   Begin VB.CheckBox chkSubExpedienteResumen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha Resumen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CheckBox chkResumen 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha Resumen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Value           =   1  'Checked
      Width           =   5055
   End
   Begin VB.CheckBox chkEstadoCuenta 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Estado de Cuenta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   5055
   End
   Begin VB.CheckBox chkSubExpedienteEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Estado de Cuenta"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   4080
      Width           =   3255
   End
   Begin VB.CheckBox chkSubExpedienteDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha Detallada"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   3720
      Width           =   3255
   End
   Begin VB.CheckBox chkSubExpediente 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprime Expedientes Asociados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   3000
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CheckBox chkDetalle 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficha Detallada"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Value           =   1  'Checked
      Width           =   5055
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
      TabIndex        =   7
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
            
            If rs!ESTADOACTUAL <> "S" Then
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
