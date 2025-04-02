VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCC_ConsultaExcedente 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta del Desgloce de Excedentes"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11145
   Icon            =   "frmCC_ConsultaExcedentes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   11145
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2772
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   10932
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   4890
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   252
      Left            =   5400
      TabIndex        =   28
      Top             =   120
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
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
      TextAlignment   =   1
      Appearance      =   16
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   2292
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   10812
      _Version        =   1441793
      _ExtentX        =   19071
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "Estado"
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
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Excedente Bruto"
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
         Height          =   252
         Index           =   1
         Left            =   480
         TabIndex        =   24
         Top             =   360
         Width           =   2172
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) Capitalizacion General"
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
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   23
         Top             =   720
         Width           =   2172
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) Impuesto de Renta"
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
         Height          =   372
         Index           =   3
         Left            =   480
         TabIndex        =   22
         Top             =   1080
         Width           =   2172
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) Mora General"
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
         Height          =   252
         Index           =   4
         Left            =   4560
         TabIndex        =   21
         Top             =   720
         Width           =   2412
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) OPCF Tras. Ded."
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
         Height          =   252
         Index           =   5
         Left            =   4560
         TabIndex        =   20
         Top             =   1080
         Width           =   2412
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) Créditos de Excedentes"
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
         Height          =   252
         Index           =   6
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   2412
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) Donaciones"
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
         Height          =   252
         Index           =   7
         Left            =   480
         TabIndex        =   18
         Top             =   1440
         Width           =   2172
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(-) Capitaliza Extraordinaria"
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
         Height          =   252
         Index           =   8
         Left            =   4560
         TabIndex        =   17
         Top             =   1440
         Width           =   2532
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Excedente a Pagar"
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
         Height          =   252
         Index           =   9
         Left            =   4560
         TabIndex        =   16
         Top             =   1800
         Width           =   2412
      End
      Begin VB.Label lblBruto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   15
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label lblRenta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   14
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label lblCapGen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   13
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label lblDonacion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   12
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label lblMoraGeneral 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   7320
         TabIndex        =   11
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label lblOPCF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   7320
         TabIndex        =   10
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label lblSaldos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   7320
         TabIndex        =   9
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label lblCapExt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   7320
         TabIndex        =   8
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label lblPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   7320
         TabIndex        =   7
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Label lblAjuste 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   2760
         TabIndex        =   6
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Label lbl 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "(+/-) Ajuste Aplicado"
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
         Height          =   252
         Index           =   11
         Left            =   480
         TabIndex        =   5
         Top             =   1800
         Width           =   2172
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3480
      TabIndex        =   26
      Top             =   480
      Width           =   5292
      _Version        =   1441793
      _ExtentX        =   9334
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1800
      TabIndex        =   25
      Top             =   480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodo 
      Height          =   312
      Left            =   1800
      TabIndex        =   29
      Top             =   120
      Width           =   3492
      _Version        =   1441793
      _ExtentX        =   6165
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   9720
      TabIndex        =   30
      Top             =   480
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCC_ConsultaExcedentes.frx":030A
   End
   Begin VB.Label lblSalida 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   312
      Left            =   0
      TabIndex        =   3
      Top             =   6240
      Width           =   11052
   End
   Begin VB.Label lblBusca 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   2772
   End
   Begin VB.Image imgReporte 
      Appearance      =   0  'Flat
      Height          =   252
      Left            =   8880
      Picture         =   "frmCC_ConsultaExcedentes.frx":0BDB
      Stretch         =   -1  'True
      Top             =   480
      Width           =   252
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
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
      Index           =   10
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1452
   End
End
Attribute VB_Name = "frmCC_ConsultaExcedente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vRA_Access As Boolean

Dim vPaso As Boolean


Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass


Call Excel_Exportar_Lsw(lsw)


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboPeriodo_Click()
If vPaso Then Exit Sub

Call sbConsulta(txtCedula.Text)

End Sub

Private Sub Form_Activate()
vModulo = 10

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMes As Integer, vFecha As Date

On Error GoTo vError
vModulo = 10


With lsw.ColumnHeaders
    .Clear
    .Add , , "Operacion", 1500
    .Add , , "Código", 1000, vbCenter
    .Add , , "Fuente", 1500, vbCenter
    .Add , , "Int.Cor.", 1800, vbRightJustify
    .Add , , "Int.Mor.", 1800, vbRightJustify
    .Add , , "Cargos", 1800, vbRightJustify
    .Add , , "Póliza", 1800, vbRightJustify
    .Add , , "Principal", 1800, vbRightJustify
    .Add , , "Total", 1800, vbRightJustify
    .Add , , "Tipo Doc.", 1500, vbCenter
    .Add , , "Num. Doc.", 2100
    .Add , , "Descripción", 4000
End With

vPaso = True
    strSQL = "select Idx, ItmX  From vExc_Periodos where ESTADO = 'C' order by IdX desc"
    Call sbCbo_Llena_New(cboPeriodo, strSQL, False, True)
vPaso = False
vError:

End Sub


Private Sub imgReporte_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long

Me.MousePointer = vbHourglass

On Error GoTo vError


If chkTodos.Value = vbUnchecked Then
    Call sbEstadoExcedentes(txtCedula.Text, cboPeriodo.ItemData(cboPeriodo.ListIndex))
Else
   strSQL = "select S.cedula from exc_cierre E inner join Socios S on E.cedula = S.cedula" _
          & " where E.id_Periodo = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) _
          & " order by S.nombre"
   Call OpenRecordSet(rs, strSQL)
   i = 1
   Do While Not rs.EOF
     lblBusca.Caption = "Procesando Registros: " & i & " de " & rs.RecordCount
     lblBusca.Refresh
    
        Call sbEstadoExcedentes(txtCedula.Text, cboPeriodo.ItemData(cboPeriodo.ListIndex))
     
     i = i + 1
     rs.MoveNext
   Loop
   rs.Close
   
   lblBusca.Caption = ""
   
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault

End Sub




Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

On Error Resume Next

If KeyCode = vbKeyReturn Then
  strSQL = "select nombre from socios where cedula = '" & txtCedula & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then
     txtNombre = rs!Nombre & ""
     Call sbConsulta(txtCedula)
  Else
     MsgBox "No se encontró registro de la persona...", vbInformation
  End If
End If

If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Identificación"
   gBusquedas.Col2Name = "Id Alterno"
   gBusquedas.Col3Name = "Nombre"
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "Select Cedula,CedulaR, Nombre From Socios"
   
   frmBusquedas.Show vbModal
   
  
  txtCedula.Text = gBusquedas.Resultado
  txtNombre.Text = gBusquedas.Resultado2
   
  Call sbConsulta(txtCedula)

End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

If KeyCode = vbKeyF4 Then
  Call sbLimpiaDatos
  
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from socios"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  
  frmBusquedas.Show vbModal
  
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
  
  Call sbConsulta(txtCedula)
End If

End Sub


Private Sub sbLimpiaDatos()

lblBruto.Caption = ""
lblBusca.Caption = ""
lblCapExt.Caption = ""
lblDonacion.Caption = ""
lblMoraGeneral.Caption = ""
lblOPCF.Caption = ""
lblPagar.Caption = ""
lblRenta.Caption = ""
lblSaldos.Caption = ""
lblCapGen.Caption = ""
lblAjuste.Caption = ""

lblSalida.Caption = ""

lsw.ListItems.Clear

End Sub


Private Function fxSalida(vSalida As String)

Select Case vSalida
  Case "01"
     fxSalida = "SALARIOS - BANCO NACIONAL"
  Case "02"
     fxSalida = "SALARIOS - BANCO POPULAR"
  Case "03"
     fxSalida = "SALARIOS - BANCO COSTA RICA"
  Case "04"
     fxSalida = "SALARIOS - COCIQUE"
  Case "05"
     fxSalida = "SALARIOS - COOPE ALIANZA"
  Case "06"
     fxSalida = "SALARIOS - BANCO SAN JOSE"
  Case "07"
     fxSalida = "SALARIOS - MUTUAL ALAJUELA"
  Case "08"
     fxSalida = "SALARIOS - MUCAP"
  Case "09"
     fxSalida = "SALARIOS - BANCO CREDITO AGRICOLA"
  Case "10"
     fxSalida = "SALARIOS - BANCO BANEX"
  Case "11"
     fxSalida = "SALARIOS - MUTUAL HEREDIA"
  Case "12"
     fxSalida = "SALARIOS - COOPE SAN RAMON"
  Case "13"
     fxSalida = "SALARIOS - MUTUAL CARTAGO"
  Case "EF"
     fxSalida = "TESORERIA - EF.FONDO"
  Case "CK"
     fxSalida = "TESORERIA - CHEQUES"
  Case "FX", "FY", "F0", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "F10", "C1", "C2", "C3", "C4"
     fxSalida = "FONDOS DE DIVIDENDOS - CASOS ESPECIALES"
  Case "BK"
     fxSalida = "CASOS BLOQUEADOS"
  Case "BN"
     fxSalida = "TESORERIA - BANCO NACIONAL"
  Case "BC"
     fxSalida = "TESORERIA - BANCO DE COSTA RICA"
  Case "PO"
     fxSalida = "TESORERIA - BANCO POPULAR"
  Case Else
     fxSalida = "NO IDENTIFICADO"
End Select

End Function

Private Sub sbConsulta(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNC_Mora As String, vNC_Saldos As String, vNC_OPCF As String
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

Call sbLimpiaDatos


'Valida Acceso a Expediente
vRA_Access = fxSys_RA_Consulta(Trim(vCedula), glogon.Usuario)
 
If Not vRA_Access Then
    MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
    txtCedula.Text = ""
    txtNombre.Text = ""
    Exit Sub
End If
    

strSQL = "select * from Exc_Periodos where id_periodo = " & cboPeriodo.ItemData(cboPeriodo.ListIndex)
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "El periodo indicado no existe...", vbCritical
  Exit Sub
End If

vNC_Mora = IIf(IsNull(rs!nc_mora), 0, rs!nc_mora)
vNC_OPCF = IIf(IsNull(rs!nc_opcf), 0, rs!nc_opcf)
vNC_Saldos = IIf(IsNull(rs!nc_saldos), 0, rs!nc_saldos)

strSQL = "select E.*,isnull(S.DESCRIPCION,'No Identificada') as 'SalidaDesc'" _
       & " from exc_cierre E left join EXC_TIPOS_SALIDAS S on E.SALIDA_CODIGO = S.COD_SALIDA" _
       & " where E.id_periodo = " & rs!Id_Periodo _
       & " and E.cedula = '" & vCedula & "'"
rs.Close

Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  rs.Close
  MsgBox "No se encontró registro de excedentes para esta persona en el periodo indicado", vbCritical
  Exit Sub
End If

lblBruto.Caption = Format(rs!excedente_bruto, "Standard")
lblCapGen.Caption = Format(rs!capitalizado, "Standard")
lblRenta.Caption = Format(rs!Renta, "Standard")
lblDonacion.Caption = Format(rs!donacion, "Standard")
lblAjuste.Caption = Format(rs!ajuste_aplicado, "Standard")
lblMoraGeneral.Caption = Format(rs!mora_aplicada, "Standard")
lblOPCF.Caption = Format(rs!moraopcf_aplicada, "Standard")
lblSaldos.Caption = Format(rs!saldos_ase_aplicados, "Standard")
lblPagar.Caption = Format(rs!excedente_final, "Standard")
lblCapExt.Caption = Format(rs!capitalizado_individual, "Standard")
lblSalida.Caption = rs!SalidaDesc ' fxSalida(IIf(IsNull(rs!SALIDA), "", rs!SALIDA))
rs.Close

'Busca Notas
If vNC_Mora <> "" And CCur(lblMoraGeneral.Caption) > 0 Then
   strSQL = "select M.*" _
          & " from vSIFAuxCreditosMovDetalle M " _
          & " where M.tcon in('7','NC') and M.ncon = '" & vNC_Mora & "' and M.cedula = '" & vCedula _
          & "'"
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!Id_Solicitud)
         itmX.SubItems(1) = rs!Codigo & ""
         itmX.SubItems(2) = rs!CONCEPTO
         itmX.SubItems(3) = Format(rs!IntCor, "Standard")
         itmX.SubItems(4) = Format(rs!IntMor, "Standard")
         itmX.SubItems(5) = Format(rs!Cargo, "Standard")
         itmX.SubItems(6) = Format(rs!Poliza, "Standard")
         itmX.SubItems(7) = Format(rs!Principal, "Standard")
         itmX.SubItems(8) = Format(rs!Total_Mov, "Standard")
         itmX.SubItems(9) = "NC"
         itmX.SubItems(10) = rs!nCon & ""
         itmX.SubItems(11) = rs!Antiguedad
     rs.MoveNext
   Loop
   rs.Close
End If

If vNC_OPCF <> "" And CCur(lblOPCF.Caption) > 0 Then
 
  strSQL = "select * from vSIFAuxCreditosMovDetalle where tcon in('7','NC')  and ncon = '" _
         & vNC_OPCF & "' and id_solicitud in(" _
         & "(select id_solicitud from reg_creditos where referencia in(" _
         & "select id_solicitud from reg_creditos where cedula = '" _
         & vCedula & "' and garantia = 'F')))"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!Id_Solicitud)
         itmX.SubItems(1) = rs!Codigo & ""
         itmX.SubItems(2) = "OPCF: " & rs!CONCEPTO
         itmX.SubItems(3) = Format(rs!IntCor, "Standard")
         itmX.SubItems(4) = Format(rs!IntMor, "Standard")
         itmX.SubItems(5) = Format(rs!Cargo, "Standard")
         itmX.SubItems(6) = Format(rs!Poliza, "Standard")
         itmX.SubItems(7) = Format(rs!Principal, "Standard")
         itmX.SubItems(8) = Format(rs!Total_Mov, "Standard")
         itmX.SubItems(9) = "NC"
         itmX.SubItems(10) = rs!nCon & ""
         itmX.SubItems(11) = rs!Antiguedad
     rs.MoveNext
   Loop
   rs.Close
End If

If vNC_Saldos <> "" And CCur(lblSaldos.Caption) > 0 Then
   strSQL = "select C.*" _
          & " from vSIFAuxCreditosMovDetalle C" _
          & " where C.tcon in('7','NC')  and C.ncon = '" & vNC_Saldos & "' and C.cedula = '" & vCedula & "'"
   Call OpenRecordSet(rs, strSQL)
   Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!Id_Solicitud)
         itmX.SubItems(1) = rs!Codigo & ""
         itmX.SubItems(2) = "CEXD: " & rs!CONCEPTO
         itmX.SubItems(3) = Format(rs!IntCor, "Standard")
         itmX.SubItems(4) = Format(rs!IntMor, "Standard")
         itmX.SubItems(5) = Format(rs!Cargo, "Standard")
         itmX.SubItems(6) = Format(rs!Poliza, "Standard")
         itmX.SubItems(7) = Format(rs!Principal, "Standard")
         itmX.SubItems(8) = Format(rs!Total_Mov, "Standard")
         itmX.SubItems(9) = "NC"
         itmX.SubItems(10) = rs!nCon & ""
         itmX.SubItems(11) = rs!Antiguedad
     rs.MoveNext
   Loop
   rs.Close
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
