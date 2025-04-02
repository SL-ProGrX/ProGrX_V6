VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_ActaAfiliacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actas de Afiliación"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   HelpContextID   =   1001
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2295
   ScaleWidth      =   7260
   Begin XtremeSuiteControls.FlatEdit txtActa 
      Height          =   512
      Left            =   5280
      TabIndex        =   1
      Top             =   240
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   903
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton cmdActas 
      Height          =   612
      Index           =   0
      Left            =   3480
      TabIndex        =   2
      Top             =   1440
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Genera Acta"
      BackColor       =   16777215
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
      Picture         =   "frmAF_ActaAfiliacion.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdActas 
      Height          =   612
      Index           =   1
      Left            =   5280
      TabIndex        =   3
      Top             =   1440
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3196
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Consulta"
      BackColor       =   16777215
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
      Picture         =   "frmAF_ActaAfiliacion.frx":09C3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Acta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAF_ActaAfiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdActas_Click(Index As Integer)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Arreglo de Controles, que permiten 1-Generar e imprimir el acta. Al generar
'               el acta se aumenta el valor de la misma en uno y se actualiza con su nuevo
'               valor el #acta para los socios que tienen el #acta pendiente y se imprime
'               listado con los socios asignados con el nuevo #acta. 2-Consultar el acta.
'               Imprime listado de socios pertenecientes al #acta suministrado.
'REFERENCIAS:   Bitacora - (Registra Movimientos Sobre la Base de Datos)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'               fxFechaServidor - (Devuelve Fecha desde el servidor)
'OBSERVACIONES: Ninguna
''''''''''''''''''''''''''''''''''''''''''''''''''

Dim recActa As New ADODB.Recordset
Dim lngNumActa As Long
Dim strSQL As String

On Error GoTo ErrorTransaccion

If Index <> 3 Then
   Me.MousePointer = vbHourglass
End If

Select Case Index
    Case 0 'Genera e imprime el acta
        
        recActa.Open "Select * From Par_AfAh", glogon.Conection, adOpenStatic
        
        If recActa.EOF = False Then
           lngNumActa = IIf(IsNull(recActa!Nacta), 1, recActa!Nacta + 1)
           strSQL = "Update Par_Afah Set Nacta=" & lngNumActa
           Call ConectionExecute(strSQL)
        End If
        
        recActa.Close
        
        strSQL = "Update Socios Set EstadoActa='I',Nacta=" & lngNumActa
        strSQL = strSQL & ",FecActa='" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
        strSQL = strSQL & " Where EstadoActa='P' and EstadoActual = 'S'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Genera", "Acta de afiliación número " & lngNumActa)
        
        With frmContenedor.Crt
            .Reset
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowState = crptMaximized
            .WindowTitle = "Reportes del Módulo de Personas"
            
            .Connect = glogon.ConectRPT
            
            .ReportFileName = SIFGlobal.fxPathReportes("Personas_ImprimeActa.rpt")
            .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
            .SelectionFormula = "{SOCIOS.NACTA}=" & lngNumActa
            .PrintReport
        End With
        
    Case 1 'consulta acta
       If Trim(txtActa) = "" Then
        MsgBox "Suministre El Numero de Acta a Consultar", vbExclamation, "Faltan Datos"
        txtActa.SetFocus
       Else
        With frmContenedor.Crt
            .Reset
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowState = crptMaximized
            .WindowTitle = "Reportes del Módulo de Personas"
            
            .Connect = glogon.ConectRPT
            
            .ReportFileName = SIFGlobal.fxPathReportes("Personas_ImprimeActa.rpt")
            .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
            .SelectionFormula = "{SOCIOS.NACTA}=" & Trim(txtActa)
            .PrintReport
        End With
       End If
       
    Case 3 'sale de la ventana
        Unload Me
End Select

If Index <> 3 Then
 Me.MousePointer = vbDefault
End If

Exit Sub



ErrorTransaccion:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub


Private Sub Form_Activate()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Verificar y establecer permisos sobre el formulario.
'REFERENCIAS:   Formularios - (Verifica los derechos que hay para el usuario en cada uno de
'               los objetos del formulario y establece respectivamente la propiedad Tag de
'               cada objeto en Uno si tiene permiso o en Cero en caso contrario)
'               RefrescaTags - (Deshabilita los objetos del formulario que tienen la
'               propiedad Tag en Cero)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

vModulo = 1
End Sub



Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub txtActa_KeyPress(KeyAscii As Integer)
On Error GoTo error

 KeyAscii = (Validacion(KeyAscii))
 
 If KeyAscii = vbKeyReturn And cmdActas(1).Enabled = True Then
    cmdActas(1).SetFocus
 End If

Exit Sub
error:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


