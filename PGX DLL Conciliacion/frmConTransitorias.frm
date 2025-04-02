VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmConTransitorias 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   Caption         =   "Conciliación de Cuentas Transitorias"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13830
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9970
      _ExtentY        =   6159
      _StockProps     =   77
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
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   4920
      TabIndex        =   9
      Top             =   240
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   9975
      _Version        =   1441792
      _ExtentX        =   17595
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   240
      Top             =   240
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   495
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmConTransitorias.frx":0000
   End
   Begin XtremeSuiteControls.ComboBox cboOrigen 
      Height          =   330
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      _Version        =   1441792
      _ExtentX        =   6800
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
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2561
      _ExtentY        =   868
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
      Picture         =   "frmConTransitorias.frx":0700
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   330
      Left            =   6360
      TabIndex        =   6
      Top             =   240
      Width           =   1455
      _Version        =   1441792
      _ExtentX        =   2566
      _ExtentY        =   582
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   6360
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmConTransitorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Private Sub btnExport_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

prgBar.Visible = True

Call Excel_Exportar_Lsw(lsw, prgBar)

prgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnBuscar_Click()

On Error GoTo vError


Me.MousePointer = vbHourglass

lsw.ListItems.Clear



strSQL = "exec spSys_Cuentas_Transitorias '" & cboOrigen.ItemData(cboOrigen.ListIndex) _
        & "', '" & Format(dtpInicio.Value, "yyyy-MM-dd") & "','" & Format(dtpCorte.Value, "yyyy-MM-dd") & " 23:59:59'"

Call OpenRecordSet(rs, strSQL)

prgBar.Visible = True

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
     
     Set itmX = lsw.ListItems.Add(, , rs!Cedula)
       itmX.SubItems(1) = Trim(rs!Nombre)
       
     Select Case cboOrigen.ItemData(cboOrigen.ListIndex)
        Case "CRD"
            itmX.SubItems(2) = Trim(rs!Operacion & "")
            itmX.SubItems(3) = Trim(rs!Codigo & "")
            itmX.SubItems(4) = Format(rs!Monto, "Standard")
            itmX.SubItems(5) = rs!Cuenta & ""
            itmX.SubItems(6) = rs!Cuenta_Desc & ""
            itmX.SubItems(7) = rs!Fecha_Registro & ""
            itmX.SubItems(9) = rs!Tesoreria_Id & ""
            itmX.SubItems(9) = rs!Fecha_Liquida & ""
           
        Case "DSM"
            
            itmX.SubItems(2) = Trim(rs!Operacion & "")
            itmX.SubItems(3) = Trim(rs!Codigo & "")
            
            itmX.SubItems(4) = rs!Concepto & ""
            itmX.SubItems(5) = Format(rs!Monto, "Standard")
            
            itmX.SubItems(6) = rs!Cuenta & ""
            itmX.SubItems(7) = rs!Cuenta_Desc & ""
            itmX.SubItems(8) = rs!Fecha_Registro & ""
            itmX.SubItems(9) = rs!Tesoreria_Id & ""
            itmX.SubItems(10) = rs!Fecha_Liquida & ""
        
        
        Case "FND"
            itmX.SubItems(2) = Trim(rs!Num_Liq & "")
            itmX.SubItems(3) = Trim(rs!cod_Plan & "")
            itmX.SubItems(4) = rs!Descripcion & ""
            itmX.SubItems(5) = rs!cod_Contrato & ""
            
            itmX.SubItems(6) = Format(rs!Monto, "Standard")
            itmX.SubItems(7) = rs!Cuenta & ""
            itmX.SubItems(8) = rs!Cuenta_Desc & ""
            itmX.SubItems(9) = rs!Fecha_Registro & ""
            itmX.SubItems(10) = rs!Tesoreria_Id & ""
            itmX.SubItems(11) = rs!Fecha_Liquida & ""
        
        
        Case "LIQ"
            itmX.SubItems(2) = Trim(rs!Num_Liq & "")
           
            itmX.SubItems(3) = Format(rs!Monto, "Standard")
            itmX.SubItems(4) = rs!Cuenta & ""
            itmX.SubItems(5) = rs!Cuenta_Desc & ""
            itmX.SubItems(6) = rs!Fecha_Registro & ""
            itmX.SubItems(7) = rs!Tesoreria_Id & ""
            itmX.SubItems(8) = rs!Fecha_Liquida & ""
    
    End Select

 rs.MoveNext

 prgBar.Value = prgBar.Value + 1
  
Loop
rs.Close

prgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    prgBar.Visible = False
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cboOrigen_Click()

lsw.ListItems.Clear


With lsw.ColumnHeaders
       .Clear
       
 Select Case cboOrigen.ItemData(cboOrigen.ListIndex)
    Case "CRD"
       .Add , , "Identificación", 2100, vbCenter
       .Add , , "Nombre", 3900
       .Add , , "Operación", 1900, vbCenter
       .Add , , "Código", 1200, vbCenter
       .Add , , "Monto", 2100, vbRightJustify
       .Add , , "Cuenta", 2100, vbCenter
       .Add , , "Cuenta Desc", 3900
       .Add , , "Fecha Registro", 2100, vbCenter
       .Add , , "Tesoreria Id", 2100, vbCenter
       .Add , , "Fecha Liquida", 2100, vbCenter

    Case "DSM"
       .Add , , "Identificación", 2100, vbCenter
       .Add , , "Nombre", 3900
       .Add , , "Operación", 1900, vbCenter
       .Add , , "Código", 1200, vbCenter
       .Add , , "Beneficiario", 3900
       .Add , , "Monto", 2100, vbRightJustify
       .Add , , "Cuenta", 2100, vbCenter
       .Add , , "Cuenta Desc", 3900
       .Add , , "Fecha Registro", 2100, vbCenter
       .Add , , "Tesoreria Id", 2100, vbCenter
       .Add , , "Fecha Liquida", 2100, vbCenter

    Case "FND"
       .Add , , "Identificación", 2100, vbCenter
       .Add , , "Nombre", 3900
       .Add , , "No. Liq.", 1900, vbCenter
       .Add , , "Plan", 1200, vbCenter
       .Add , , "Plan Desc.", 3900
       .Add , , "No. Contrato", 1900, vbCenter
       .Add , , "Monto", 2100, vbRightJustify
       .Add , , "Cuenta", 2100, vbCenter
       .Add , , "Cuenta Desc", 3900
       .Add , , "Fecha Registro", 2100, vbCenter
       .Add , , "Tesoreria Id", 2100, vbCenter
       .Add , , "Fecha Liquida", 2100, vbCenter

    Case "LIQ"
       .Add , , "Identificación", 2100, vbCenter
       .Add , , "Nombre", 3900
       .Add , , "No. Liq.", 1900, vbCenter
       .Add , , "Monto", 2100, vbRightJustify
       .Add , , "Cuenta", 2100, vbCenter
       .Add , , "Cuenta Desc", 3900
       .Add , , "Fecha Registro", 2100, vbCenter
       .Add , , "Tesoreria Id", 2100, vbCenter
       .Add , , "Fecha Liquida", 2100, vbCenter

End Select

End With

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - 150
lsw.Height = Me.Height - (lsw.Top + prgBar.Height + 480)

prgBar.Left = 0
prgBar.Width = Me.Width

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub Form_Load()

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture


cboOrigen.AddItem "Desembolsos de Créditos"
cboOrigen.ItemData(cboOrigen.ListCount - 1) = "CRD"

cboOrigen.AddItem "Desembolsos Créditos a Terceros"
cboOrigen.ItemData(cboOrigen.ListCount - 1) = "DSM"

cboOrigen.AddItem "Liq/Retiros de Ahorros"
cboOrigen.ItemData(cboOrigen.ListCount - 1) = "FND"

cboOrigen.AddItem "Liquidación de Asociados"
cboOrigen.ItemData(cboOrigen.ListCount - 1) = "LIQ"

cboOrigen.AddItem "Hipotecarios"
cboOrigen.ItemData(cboOrigen.ListCount - 1) = "HIP"


Me.BackColor = RGB(214, 234, 248)


End Sub



Private Sub sbInicial()

On Error GoTo vError

Me.MousePointer = vbHourglass

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = dtpCorte.Value


cboOrigen.Text = "Desembolsos de Créditos"
Call cboOrigen_Click

Me.MousePointer = vbDefault


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub


