VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmAF_CD_InformeEspecial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe Especial de Antiguedad"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lswUnidades 
      Height          =   1335
      Left            =   0
      TabIndex        =   11
      Top             =   6240
      Width           =   5775
      _Version        =   1572864
      _ExtentX        =   10186
      _ExtentY        =   2355
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
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswAntiguedad 
      Height          =   3375
      Left            =   5760
      TabIndex        =   9
      Top             =   2400
      Width           =   5775
      _Version        =   1572864
      _ExtentX        =   10186
      _ExtentY        =   5953
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
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswActividades 
      Height          =   3375
      Left            =   0
      TabIndex        =   8
      Top             =   2400
      Width           =   5775
      _Version        =   1572864
      _ExtentX        =   10186
      _ExtentY        =   5953
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
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   720
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   330
      Left            =   7080
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.ComboBox cboComite 
      Height          =   330
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7646
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.PushButton btnReporte 
      Height          =   495
      Left            =   9840
      TabIndex        =   5
      Top             =   7080
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Reporte"
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
      Picture         =   "frmAF_CD_InformeEspecial.frx":0000
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ComboBox cboZona 
      Height          =   330
      Left            =   1200
      TabIndex        =   12
      Top             =   1320
      Width           =   4335
      _Version        =   1572864
      _ExtentX        =   7646
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Zona"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   5880
      Width           =   5775
      _Version        =   1572864
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Unidades"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   1
      Left            =   5760
      TabIndex        =   7
      Top             =   2040
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10398
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Antiguedad"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   5775
      _Version        =   1572864
      _ExtentX        =   10186
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Actividades"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6120
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comite"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informes de Antiguedad de Saldos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   6
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   7485
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmAF_CD_InformeEspecial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub btnReporte_Click()
Dim vTitulo As String, vSubTitulo As String
Dim i As Integer

On Error GoTo vError

Dim pComites As String, pActividades As String, pAntiguedad As String, pZonas As String

pActividades = ""
pAntiguedad = ""
pZonas = ""

If cboComite.Text = "TODOS" Then
    pComites = "TODOS"
Else
    pComites = cboComite.ItemData(cboComite.ListIndex)
End If

If cboZona.Text = "TODOS" Then
    pZonas = "TODOS"
Else
    pZonas = cboZona.ItemData(cboZona.ListIndex)
End If



With lswActividades.ListItems
For i = 1 To .Count
   If .Item(i).Checked Then
        If pActividades = "" Then
            pActividades = .Item(i).Tag
        Else
            pActividades = pActividades & "," & .Item(i).Tag
        End If
   End If
Next i
End With


With lswAntiguedad.ListItems
For i = 1 To .Count
   If .Item(i).Checked Then
        If pAntiguedad = "" Then
            pAntiguedad = .Item(i).Tag
        Else
            pAntiguedad = pAntiguedad & "," & .Item(i).Tag
        End If
   End If
Next i
End With


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .Connect = glogon.ConectRPT
    
  vTitulo = "LIQUIDACIONES POR COMITE"
  
  If cboEstado.Text = "Activa" Then
      vSubTitulo = "LIQ. PENDIENTES [Zona: " & cboZona.Text & "..Comité: " & cboComite.Text & "..Estado: " & cboEstado.Text & "]"
  Else
      vSubTitulo = "LIQ. REALIZADAS [Zona: " & cboZona.Text & "..Comité: " & cboComite.Text & "..Estado: " & cboEstado.Text & "]"
  End If
    
    .ReportFileName = SIFGlobal.fxPathReportes("Comites_Antiguedad_Saldos.rpt")
    
    .StoredProcParam(0) = pComites
    .StoredProcParam(1) = pActividades
    .StoredProcParam(2) = pAntiguedad
    .StoredProcParam(3) = cboEstado.ItemData(cboEstado.ListIndex)
    .StoredProcParam(4) = pZonas
       
 
    .Formulas(0) = "fxTitulo = '" & vTitulo & "'"
    .Formulas(1) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "fxUsuario='USUARIO: " & glogon.Usuario & "'"
    .Formulas(4) = "fxSubtitulo = '" & vSubTitulo & "'"
    
    .Action = 1

End With

Exit Sub

vError:
      MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboComite_Click()
If vPaso Then Exit Sub
If cboComite.ListCount = 0 Then Exit Sub

If cboComite.Text = "TODOS" Then
     lswUnidades.ListItems.Clear
    Exit Sub
End If

strSQL = "select Cu.COD_COMITE, U.CODIGO, U.DESCRIPCION" _
       & "  from AFI_CD_COMITES_UNIDADES Cu inner join UPROGRAMATICA U on Cu.CODIGO_UP = U.CODIGO" _
       & " Where Cu.COD_COMITE = '" & cboComite.ItemData(cboComite.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
With lswUnidades.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!Codigo)
          itmX.SubItems(1) = rs!Descripcion & ""
      rs.MoveNext
    Loop
    rs.Close
End With

End Sub

Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "Descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_COMITE, DESCRIPCION from AFI_CD_COMITES"
       gBusquedas.Filtro = " AND ACTIVO = 1"
       frmBusquedas.Show vbModal
       If gBusquedas.Resultado <> "" Then
         Call sbCboAsignaDato(cboComite, gBusquedas.Resultado2, True, gBusquedas.Resultado)
       End If
End If
End Sub

Private Sub cboZona_Click()
If vPaso Then Exit Sub
If cboZona.ListCount < 0 Then Exit Sub

vPaso = True

If cboZona.Text = "TODOS" Then
    strSQL = "select COD_COMITE as 'IdX' , rtrim(DESCRIPCION) as 'ItmX'" _
           & " from AFI_CD_COMITES order by Descripcion"
Else
    strSQL = "select COD_COMITE as 'IdX' , rtrim(DESCRIPCION) as 'ItmX'" _
           & " from vAFI_CD_Comites_Zonas Where Cod_Zona = '" & cboZona.ItemData(cboZona.ListIndex) & "' order by Descripcion"
End If

Call sbCbo_Llena_New(cboComite, strSQL, True, True)

vPaso = False

End Sub

Private Sub Form_Load()

 vModulo = 40


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

cboEstado.Clear
cboEstado.AddItem "Activa"
cboEstado.ItemData(cboEstado.ListCount - 1) = "A"
cboEstado.AddItem "Liquidada"
cboEstado.ItemData(cboEstado.ListCount - 1) = "L"
cboEstado.Text = "Activa"

With lswActividades.ColumnHeaders
    .Clear
    .Add , , "", lswActividades.Width - 120
End With
  
With lswAntiguedad.ColumnHeaders
    .Clear
    .Add , , "", lswAntiguedad.Width - 120
End With
  
With lswUnidades.ColumnHeaders
    .Clear
    .Add , , "Unidad", 1000
    .Add , , "Descripción", lswUnidades.Width - 1120
End With
  
  
End Sub

Private Sub TimerX_Timer()
 
TimerX.Interval = 0
TimerX.Enabled = False
     
vPaso = True
    strSQL = "select COD_ZONA as 'IdX' , rtrim(DESCRIPCION) as 'ItmX'" _
           & " from AFI_ZONAS order by Descripcion"
    Call sbCbo_Llena_New(cboZona, strSQL, True, True)


    strSQL = "select COD_COMITE as 'IdX' , rtrim(DESCRIPCION) as 'ItmX'" _
           & " from AFI_CD_COMITES order by Descripcion"
    Call sbCbo_Llena_New(cboComite, strSQL, True, True)
vPaso = False
 
strSQL = "select COD_ACTIVIDAD as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
       & " from  AFI_CD_ACTIVIDADES where ACTIVA = 1 order by Descripcion"
Call OpenRecordSet(rs, strSQL)
With lswActividades.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!itmX)
          itmX.Tag = rs!IdX
      rs.MoveNext
    Loop
    rs.Close
End With

strSQL = "select COD_ANTIGUEDAD as 'IdX', Descripcion as 'ItmX'" _
       & " From CBR_ANTIGUEDAD_TIPOS order by descripcion"

Call OpenRecordSet(rs, strSQL)
With lswAntiguedad.ListItems
    .Clear
    Do While Not rs.EOF
      Set itmX = .Add(, , rs!itmX)
          itmX.Tag = rs!IdX
      rs.MoveNext
    Loop
    rs.Close
End With



End Sub
