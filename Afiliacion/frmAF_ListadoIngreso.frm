VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAf_ListadoIngreso 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listados de Ingresos"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "frmAF_ListadoIngreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8400
      Top             =   600
   End
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2052
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   2292
      _Version        =   1441793
      _ExtentX        =   4043
      _ExtentY        =   3619
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Checkboxes      =   -1  'True
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkFecha 
      Height          =   252
      Left            =   7200
      TabIndex        =   7
      Top             =   1680
      Width           =   972
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.DateTimePicker dtpDe 
      Height          =   312
      Left            =   4440
      TabIndex        =   1
      Top             =   1680
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
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   312
      Left            =   5760
      TabIndex        =   2
      Top             =   1680
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   4440
      TabIndex        =   6
      Top             =   2040
      Width           =   5172
      _Version        =   1441793
      _ExtentX        =   9128
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.GroupBox gbInforme 
      Height          =   972
      Left            =   2520
      TabIndex        =   8
      Top             =   2760
      Width           =   7092
      _Version        =   1441793
      _ExtentX        =   12509
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   492
         Left            =   5400
         TabIndex        =   9
         Top             =   360
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   868
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
         Appearance      =   14
         Picture         =   "frmAF_ListadoIngreso.frx":030A
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Institución ¦ Empresa"
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
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas de Ingreso"
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
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Top             =   1680
      Width           =   2652
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Ingreso o registro"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
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
      TabIndex        =   3
      Top             =   360
      Width           =   5412
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estados Persona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   288
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   2292
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmAf_ListadoIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkFecha_Click()
If chkFecha.Value = vbChecked Then
   dtpDe.Enabled = False
Else
   dtpDe.Enabled = True
End If

dtpHasta.Enabled = dtpDe.Enabled

End Sub

Private Sub cmdImprimir_Click()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite reporte de Listado de Ingresos por estatus del socio.
'REFERENCIAS:   ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, i As Integer
Dim pSubTitulo As String, vCadena As String

On Error GoTo vError

If dtpDe.Value > dtpHasta.Value Then
  MsgBox "Verifique su entrada de Datos", vbExclamation, "Error en el Rango de Fechas"
  Exit Sub
End If


Me.MousePointer = vbHourglass

pSubTitulo = ""
strSQL = ""
vCadena = ""

With lsw.ListItems
    For i = 1 To .Count
      If .Item(i).Checked Then
         
         
         If Len(strSQL) = 0 Then
                strSQL = strSQL & "({SOCIOS.ESTADOACTUAL} IN ['"
         End If
                  
         If Len(pSubTitulo) > 0 Then
            pSubTitulo = pSubTitulo & ", "
         End If
         
         vCadena = vCadena & "','" & .Item(i).Tag
         
         pSubTitulo = pSubTitulo & .Item(i).Text
         
      End If
    Next i
    
    If i > 0 Then
        strSQL = strSQL & vCadena & "'])"
    End If
    
End With

If Len(strSQL) = 0 Then
  MsgBox "No especificó ningún estado de persona, marque almenos uno para el reporte...", vbExclamation
  Exit Sub
End If

 With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Personas"
    
    .Connect = glogon.ConectRPT
    
    .ReportFileName = SIFGlobal.fxPathReportes("Personas_IngresoSocios.rpt")
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Titulo='Listado General de Personas'"
    
    If pSubTitulo = "" Then
       pSubTitulo = "TODOS"
    Else
        pSubTitulo = "[" & pSubTitulo & "]"
    End If
    
    If chkFecha.Value = vbChecked Then
        .Formulas(2) = "SubTitulo='Fecha Ingreso: Todas ¦ Institución : " & cbo.Text _
                     & " ¦ Estados: " & pSubTitulo & "'"
    Else
        .Formulas(2) = "SubTitulo='Fecha Ingreso: " & Format(dtpDe.Value, "DD/MM/YYYY") _
                     & "  -  " & Format(dtpHasta.Value, "DD/MM/YYYY") & " ¦ Institución: " & cbo.Text _
                     & " ¦ Estados: " & pSubTitulo & "'"
        
        strSQL = strSQL & " And {SOCIOS.FECHAINGRESO} in DateTime(" & Format(dtpDe.Value, "yyyy,mm,dd") _
               & ") to DateTime(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
    End If
     
    If cbo.Text <> "TODOS" Then
       strSQL = strSQL & " and {SOCIOS.COD_INSTITUCION} = " & cbo.ItemData(cbo.ListIndex)
    End If
     
    .SelectionFormula = strSQL
    .PrintReport
         
 End With

 Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Descripción", 2270

End With

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select *" _
       & " from AFI_Estados_Persona"
Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
      itmX.Tag = rs!cod_estado
  rs.MoveNext
Loop
rs.Close

strSQL = "select cod_institucion as 'IdX',descripcion as 'ItmX' from instituciones"
Call sbCbo_Llena_New(cbo, strSQL, True, True)

dtpDe.Value = fxFechaServidor
dtpHasta.Value = dtpDe.Value
Call chkFecha_Click

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicializa

End Sub
