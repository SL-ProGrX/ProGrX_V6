VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmCR_ControlPolizas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Polizas"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_ControlPolizas.frx":0000
   ScaleHeight     =   4620
   ScaleWidth      =   7395
   Begin TabDlg.SSTab ssTabX 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Actualizaciones"
      TabPicture(0)   =   "frmCR_ControlPolizas.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdActualizar"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Cierres"
      TabPicture(1)   =   "frmCR_ControlPolizas.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboCierre"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdCierre"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Reportes"
      TabPicture(2)   =   "frmCR_ControlPolizas.frx":688A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Line2(0)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label3(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line2(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Line2(2)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblReporte"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdReporte"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "optReporte(3)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "optReporte(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "optReporte(1)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "optReporte(0)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "lsw"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "chkCierrePreliminar"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "chkCierreDefinitivo"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "cboPoliza"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).ControlCount=   16
      Begin VB.ComboBox cboPoliza 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1800
         Width           =   2895
      End
      Begin VB.CheckBox chkCierreDefinitivo 
         Caption         =   "&Definitivo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkCierrePreliminar 
         Caption         =   "&Preliminar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2655
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Tipo"
            Object.Width           =   2188
         EndProperty
      End
      Begin VB.ComboBox cboCierre 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1680
         Width           =   6135
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Inclusiones"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3840
         TabIndex        =   11
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Exclusiones"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3840
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "Modificaciones"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5280
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optReporte 
         Caption         =   "General "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   5280
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "Reporte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5520
         Picture         =   "frmCR_ControlPolizas.frx":68A6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton cmdCierre 
         Caption         =   "Cierre de Información para Polizas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74640
         Picture         =   "frmCR_ControlPolizas.frx":D0F8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   6135
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar Base de Datos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74640
         Picture         =   "frmCR_ControlPolizas.frx":1394A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   6135
      End
      Begin VB.Label lblReporte 
         Alignment       =   1  'Right Justify
         Caption         =   ">>> Seleccione un Reporte <<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   45
         Width           =   3135
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         Index           =   2
         X1              =   3720
         X2              =   6600
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         Index           =   1
         X1              =   3720
         X2              =   6600
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Reportes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3720
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00E0E0E0&
         Index           =   0
         X1              =   360
         X2              =   1080
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Cierres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmCR_ControlPolizas.frx":1A19C
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2775
         Index           =   1
         Left            =   -74760
         TabIndex        =   6
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   $"frmCR_ControlPolizas.frx":1A22D
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2655
         Index           =   0
         Left            =   -74760
         TabIndex        =   3
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label3 
         Caption         =   "Poliza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   4470
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      Caption         =   "Control de Polizas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmCR_ControlPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkCierreDefinitivo_Click()
Call ssTabX_Click(0)
End Sub

Private Sub chkCierrePreliminar_Click()
Call ssTabX_Click(0)
End Sub

Private Sub cmdActualizar_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCRDPolizasActualizacion"
glogon.Conection.Execute strSQL

Me.MousePointer = vbDefault
MsgBox "Polizas Actualizadas Satisfactoriamente...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub cmdCierre_Click()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCRDPolizasCierre '" & Mid(cboCierre.Text, 1, 1) & "'"
glogon.Conection.Execute strSQL
Me.MousePointer = vbDefault
MsgBox "Cierre : " & cboCierre.Text & " Realizado Satisfactoriamente...", vbInformation


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdReporte_Click()
Dim vPoliza As String


If lblReporte.Tag = "" Then
 MsgBox "Seleccione un Corte", vbExclamation
 Exit Sub
End If

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowState = crptMaximized
    
    .Connect = glogon.ConectRPT
    
    .WindowTitle = "Reportes - Control Pólizas"
    .ReportFileName = fxSIFPathReportes("CrdControlPolizas.rpt")
    
    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(1) = "Fecha='Fecha:" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "usuario='Usuario: " & glogon.Usuario & "'"

    vPoliza = fxCodigoCbo(cboPoliza)
    
    .SelectionFormula = "{CRD_POLIZAS_CONTROL.TIPO} = '" & lblReporte.Tag & "' and {CRD_POLIZAS_CONTROL.COD_POLIZA} = '" _
                      & vPoliza & "' and {CRD_POLIZAS_CONTROL.FECHA} = cdate('" & Format(CDate(lblReporte.Caption), "yyyy/mm/dd") & "')"
  
    Select Case True
      Case optReporte(0).Value 'Inclusiones
        .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(lblReporte.Tag = "P", "Preliminar", "Definitivo") & "  Corte : " & lblReporte.Caption & " Inclusiones'"
        .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR}=0 AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL}>0"
      Case optReporte(1).Value 'Exclusiones
        .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(lblReporte.Tag = "P", "Preliminar", "Definitivo") & "  Corte : " & lblReporte.Caption & " Exclusiones'"
        .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR}>0 AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL}=0"
      Case optReporte(2).Value 'Modificaciones
        .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(lblReporte.Tag = "P", "Preliminar", "Definitivo") & "  Corte : " & lblReporte.Caption & " Modificaciones'"
        .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR} <> {CRD_POLIZAS_CONTROL.MONTO_ACTUAL} AND " _
                          & " {CRD_POLIZAS_CONTROL.MONTO_ANTERIOR} > 0 AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL} > 0"
      Case optReporte(3).Value 'General
        .Formulas(3) = "subtitulo='Poliza: " & vPoliza & "   Cierre " & IIf(lblReporte.Tag = "P", "Preliminar", "Definitivo") & "  Corte : " & lblReporte.Caption & " General'"
        .SelectionFormula = .SelectionFormula & " AND {CRD_POLIZAS_CONTROL.MONTO_ACTUAL} > 0"
    End Select
    
    .PrintReport
End With

Me.MousePointer = vbDefault


End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

Me.Icon = MDIPrincipal.Icon

ssTabX.Tab = 0

cboCierre.AddItem "Preliminar"
cboCierre.AddItem "Definitivo"

cboCierre.Text = "Preliminar"


strSQL = "select Cod_Poliza + ' - ' + rtrim(descripcion) as ItmX " _
       & " From crd_catalogo_polizas"
cboPoliza.Clear
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 cboPoliza.AddItem rs!itmX
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
  rs.MoveFirst
  cboPoliza.Text = rs!itmX
End If
rs.Close

ssTabX.Tab = 0

vModulo = 3

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub lsw_Click()

If lsw.ListItems.Count = 0 Then Exit Sub

lblReporte.Caption = lsw.SelectedItem
lblReporte.Tag = Mid(lsw.SelectedItem.SubItems(1), 1, 1)

End Sub

Private Sub ssTabX_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

If ssTabX.Tab = 2 Then

    lblReporte.Caption = ">>> Seleccione un Cierre <<<"
    lblReporte.Tag = ""

   strSQL = "select * From crd_polizas_corte" _
          & " Where Tipo in("
          
   If chkCierreDefinitivo.Value = vbChecked Then
      strSQL = strSQL & "'D',"
   End If
   
   If chkCierrePreliminar.Value = vbChecked Then
      strSQL = strSQL & "'P',"
   End If
   
   strSQL = strSQL & "'') order by cod_corte desc"
   rs.Open strSQL, glogon.Conection, adOpenStatic
   
   lsw.ListItems.Clear
   
   Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!cod_corte)
         itmX.SubItems(1) = IIf((rs!Tipo = "P"), "Preliminar", "Definitivo")
     rs.MoveNext
   Loop
   rs.Close

End If

Me.MousePointer = vbDefault

End Sub

