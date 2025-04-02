VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "codejock.controls.v19.2.0.ocx"
Begin VB.Form frmCR_SolCreacionAgenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creación de Acta"
   ClientHeight    =   3888
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   8112
   HelpContextID   =   3014
   Icon            =   "frmCR_SolCreacionAgenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3888
   ScaleWidth      =   8112
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   1320
      TabIndex        =   7
      Top             =   1920
      Width           =   5652
      _Version        =   1245186
      _ExtentX        =   9970
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Fechas"
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   1920
         TabIndex        =   11
         Top             =   360
         Width           =   1332
         _Version        =   1245186
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   1920
         TabIndex        =   12
         Top             =   720
         Width           =   1332
         _Version        =   1245186
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
      Begin XtremeSuiteControls.DateTimePicker dtpReunion 
         Height          =   312
         Left            =   1920
         TabIndex        =   13
         Top             =   1080
         Width           =   1332
         _Version        =   1245186
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
      Begin XtremeSuiteControls.PushButton btnActa 
         Height          =   852
         Left            =   3600
         TabIndex        =   14
         Top             =   480
         Width           =   1572
         _Version        =   1245186
         _ExtentX        =   2773
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Acta"
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
         Picture         =   "frmCR_SolCreacionAgenda.frx":030A
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   3
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   972
         _Version        =   1245186
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Reunión"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   972
         _Version        =   1245186
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Corte"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   252
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   360
         Width           =   972
         _Version        =   1245186
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Inicio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.CheckBox chkPreAnalisis 
      Height          =   252
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Width           =   3732
      _Version        =   1245186
      _ExtentX        =   6583
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Imprimir Estudios de Créditos asociados?"
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
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   0
      Top             =   3756
      Width           =   8112
      _ExtentX        =   14309
      _ExtentY        =   233
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.ComboBox cboComite 
      Height          =   312
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   5652
      _Version        =   1245186
      _ExtentX        =   9970
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
   Begin XtremeSuiteControls.FlatEdit txtActa 
      Height          =   312
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1692
      _Version        =   1245186
      _ExtentX        =   2984
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
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   252
      Index           =   0
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   972
      _Version        =   1245186
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "No. Acta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actas del Comité"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   5
      Left            =   2004
      TabIndex        =   2
      Top             =   360
      Width           =   6852
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   0
      TabIndex        =   1
      Top             =   3480
      Width           =   8052
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCR_SolCreacionAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Private mlngActa As Long

Private Function fxValida()

fxValida = True

If Trim(txtActa) = "" Then
  fxValida = False
End If

If cboComite.ListIndex = -1 Then
 fxValida = False
End If

End Function

Private Sub Reporte(strTitulo As String, strCampo As String, iValidaActa As Integer)
Dim str As String, strRuta As String, strInicio As String, strFinal As String

On Error GoTo vError


str = ""
strInicio = ""
strFinal = ""

strInicio = "Date(" & CStr(Year(dtpInicio.Value)) & "," & CStr(Month(dtpInicio.Value)) & "," & CStr(Day(dtpInicio.Value)) & ")"
strFinal = "Date(" & CStr(Year(dtpCorte.Value)) & "," & CStr(Month(dtpCorte.Value)) & "," & CStr(Day(dtpCorte.Value)) & ")"

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Creditos"
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_Agenda.rpt")
    
    .Connect = glogon.ConectRPT
    
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='" & strTitulo & "'"
    .Formulas(4) = "fechareunion='" & Format(dtpReunion, "yyyy/mm/dd") & "'"
    If iValidaActa = 2 Then
        .Formulas(5) = "de='" & Format(dtpInicio, "dd/mm/yyyy") & "'"
        .Formulas(6) = "a='" & Format(dtpCorte, "dd/mm/yyyy") & "'"
    End If
    .Formulas(7) = "acta='" & txtActa & "'"
    str = "{REG_CREDITOS.ACTA} = " & txtActa
    str = str & " and {REG_CREDITOS.ID_COMITE}=" & cboComite.ItemData(cboComite.ListIndex)
    If iValidaActa = 2 Then
        str = str & " and " & strCampo & " >= " & strInicio
        str = str & " and " & strCampo & " <= " & strFinal
    End If
    .SelectionFormula = str
    .PrintReport
End With

Exit Sub
vError:
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbActa_Consulta()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(acta,0) as 'Acta' from comites where id_comite=" & cboComite.ItemData(cboComite.ListIndex)
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not IsNull(rs!acta) Then
    mlngActa = rs!acta + 1
    txtActa.Text = rs!acta + 1
Else
    txtActa.Text = 1
    mlngActa = 1
End If
rs.Close
End Sub

Private Function ValidaActa()
ValidaActa = 0
If CLng(txtActa) < mlngActa Then
  ValidaActa = 1
ElseIf CLng(txtActa) = mlngActa Then
  ValidaActa = 2
ElseIf CLng(txtActa) > mlngActa Then
  ValidaActa = 3
End If
End Function


Private Sub sbEstudioCredito_Informe()

Me.MousePointer = vbHourglass
With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    .WindowTitle = "Reportes del Módulo de Creditos"
    
    .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxPathReportes("Credito_PreAnalisis.rpt")
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "TITULO='PREANALSIS POR ACTA'"
    .SelectionFormula = "({REG_CREDITOS.ACTA} = " & txtActa _
           & " AND {REG_CREDITOS.ID_COMITE} = " & cboComite.ItemData(cboComite.ListIndex) & ")"
    .PrintReport
End With
Me.MousePointer = vbDefault

End Sub

Private Sub btnActa_Click()
Call sbActa
End Sub

Private Sub cboComite_Click()
 
 If vPaso Then Exit Sub
 
 Call sbActa_Consulta
End Sub


Private Sub sbActa()
Dim iValidaActa As Integer, rs As New ADODB.Recordset, strSQL As String

Me.MousePointer = vbHourglass

If fxValida Then
    iValidaActa = ValidaActa
    
    If iValidaActa = 3 Or iValidaActa = 0 Then 'número de txtActa > COMITES.Acta
        MsgBox "Numero de Acta no puede ser mayor a numero sugerido", vbOKOnly
        Exit Sub
    
    ElseIf iValidaActa = 2 Then 'número de txtActa= COMITES.ACTA
     'Crear Acta
     Me.MousePointer = vbHourglass
     'Actualiza Comites-Acta
     
     strSQL = "Update comites set acta = " & txtActa _
            & " where id_comite = " & cboComite.ItemData(cboComite.ListIndex)
            
            
     
     Call ConectionExecute(strSQL)
     
     lblEstado.Caption = "Cargando Casos ..."
     lblEstado.Refresh
     
     Me.Height = Me.Height + prgBar.Height + lblEstado.Height + 40
     
     strSQL = "select id_solicitud from reg_creditos where acta is null and estadosol ='R'" _
            & " and fechasol between '" & Format(dtpInicio, "yyyy/mm/dd") & "'" _
            & " and '" & Format(dtpCorte, "yyyy/mm/dd") & "' and id_comite = " & cboComite.ItemData(cboComite.ListIndex)
     rs.CursorLocation = adUseServer
     Call OpenRecordSet(rs, strSQL)
     prgBar.Max = 1
     prgBar.Value = 1
     prgBar.Max = rs.RecordCount + 1
     
     lblEstado.Caption = "Generando Acta ..."
     lblEstado.Refresh
     
     Do While Not rs.EOF
       strSQL = "update reg_Creditos set acta = " & txtActa _
              & " where id_solicitud = " & rs!id_solicitud
       Call ConectionExecute(strSQL)
       rs.MoveNext
       prgBar.Value = prgBar.Value + 1
     Loop
     rs.Close

     lblEstado.Caption = "Generando Reporte ..."
     lblEstado.Refresh

'
'      glogon.conection.Execute "sp_crActualizaActa " & cboComite.ItemData(cboComite.ListIndex) _
'                                & "," & txtActa & ",'" & Format(dtpInicio, "yyyy/mm/dd") _
'                                & "','" & Format(dtpCorte, "yyyy/mm/dd") & "'"
'
 
      
      Call Reporte("AGENDA", "{REG_CREDITOS.FECHASOL}", iValidaActa)
      If chkPreAnalisis.Value = 1 Then sbEstudioCredito_Informe
      Me.MousePointer = vbDefault
      Me.Height = Me.Height - prgBar.Height - lblEstado.Height - 40
     
    ElseIf iValidaActa = 1 Then
    ' impresion de acta anterior, número en txtActa < COMITES.Acta
     Call Reporte("REIMPRESION DE AGENDA", "{REG_CREDITOS.FECHASOL}", iValidaActa)
     If chkPreAnalisis.Value = 1 Then sbEstudioCredito_Informe
    
    End If
    
Else
    MsgBox "Faltan datos", vbOKOnly
End If

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vPaso = True
    strSQL = "Select id_comite as 'IdX',descripcion as 'ItmX' from comites where estado = 1"
    Call sbCbo_Llena_New(cboComite, strSQL, False, True)
vPaso = False

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpReunion.Value = dtpInicio.Value


Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

