VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCR_GarantiasPatrimoniales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Garantías Patrimoniales"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2892
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   9800
      _Version        =   1310723
      _ExtentX        =   17286
      _ExtentY        =   5101
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
   End
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   372
      Left            =   8760
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   372
      _Version        =   1310723
      _ExtentX        =   656
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox gbGarantias 
      Height          =   1932
      Left            =   960
      TabIndex        =   6
      Top             =   5040
      Width           =   8412
      _Version        =   1310723
      _ExtentX        =   14838
      _ExtentY        =   3408
      _StockProps     =   79
      Caption         =   "Registro"
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
      Begin XtremeSuiteControls.CheckBox chkPatrimonio 
         Height          =   312
         Left            =   6600
         TabIndex        =   16
         Top             =   1080
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Patrimonio?   "
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
         Enabled         =   0   'False
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnMov 
         Height          =   312
         Index           =   0
         Left            =   2760
         TabIndex        =   14
         Top             =   1440
         Width           =   372
         _Version        =   1310723
         _ExtentX        =   656
         _ExtentY        =   556
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmCR_GarantiasPatrimoniales.frx":0000
      End
      Begin XtremeSuiteControls.ComboBox cboOperadora 
         Height          =   312
         Left            =   1680
         TabIndex        =   11
         Top             =   360
         Width           =   6612
         _Version        =   1310723
         _ExtentX        =   11668
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
      Begin XtremeSuiteControls.FlatEdit txtPlan 
         Height          =   312
         Left            =   1680
         TabIndex        =   10
         Top             =   720
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   550
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   312
         Left            =   1680
         TabIndex        =   12
         Top             =   1440
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   550
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlanDesc 
         Height          =   312
         Left            =   2640
         TabIndex        =   13
         Top             =   720
         Width           =   5652
         _Version        =   1310723
         _ExtentX        =   9970
         _ExtentY        =   550
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnMov 
         Height          =   312
         Index           =   1
         Left            =   3120
         TabIndex        =   15
         Top             =   1440
         Width           =   372
         _Version        =   1310723
         _ExtentX        =   656
         _ExtentY        =   556
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmCR_GarantiasPatrimoniales.frx":0720
      End
      Begin XtremeSuiteControls.FlatEdit txtLinea 
         Height          =   312
         Left            =   5520
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   550
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMemInicio 
         Height          =   312
         Left            =   1680
         TabIndex        =   20
         Top             =   1080
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   550
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMemCorte 
         Height          =   312
         Left            =   2640
         TabIndex        =   21
         Top             =   1080
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
         _ExtentY        =   550
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rango en Días"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   372
         Index           =   6
         Left            =   3720
         TabIndex        =   22
         Top             =   1080
         Width           =   3252
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Membresía"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   5
         Left            =   600
         TabIndex        =   19
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Plan"
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
         Index           =   2
         Left            =   600
         TabIndex        =   9
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Operadora"
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
         Index           =   4
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   972
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   8760
      Top             =   480
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   4212
      _Version        =   1310723
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   4212
      _Version        =   1310723
      _ExtentX        =   7435
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado de la Persona"
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
      Index           =   1
      Left            =   5280
      TabIndex        =   2
      Top             =   1320
      Width           =   2412
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Garantía"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1320
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Garantías Patrimoniales + Mixtas"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6972
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCR_GarantiasPatrimoniales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbConsulta()
Dim strSQL As String, rs  As New ADODB.Recordset
Dim vGarantia As String, vEstado As String
Dim itmX As ListViewItem

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

If cbo.ListCount = 0 Or cboEstado.ListCount = 0 Then Exit Sub
If cbo.Text = "" Or cboEstado.Text = "" Then Exit Sub

txtPlan.Text = ""
txtPlanDesc.Text = ""
txtPorcentaje.Text = Format(0, "Standard")
txtMemInicio.Text = "0"
txtMemCorte.Text = "99999"
txtLinea.Text = "0"
chkPatrimonio.Value = xtpUnchecked


vGarantia = cbo.ItemData(cbo.ListIndex)
vEstado = cboEstado.ItemData(cboEstado.ListIndex)

strSQL = "exec spCrd_Garantia_Ahorros_Consulta '" & vGarantia & "','" & vEstado & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Linea_Id)
      itmX.SubItems(1) = rs!cod_Operadora
      itmX.SubItems(2) = rs!cod_Plan
      itmX.SubItems(3) = rs!Descripcion
      itmX.SubItems(4) = rs!MEMBRESIA_INICIO
      itmX.SubItems(5) = rs!MEMBRESIA_CORTE
      itmX.SubItems(6) = Format(rs!Porcentaje, "Standard")
      itmX.SubItems(7) = IIf(rs!Patrimonio = 1, "Sí", "No")

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub btnMov_Click(Index As Integer)
Dim strSQL As String, strBitacora As String
Dim pOperadora As Integer, pPlan As String, pGarantia As String, pEstado As String

On Error GoTo vError


If Not IsNumeric(txtPorcentaje.Text) Then
   MsgBox "Porcentaje no es válido", vbExclamation
   Exit Sub
Else
  If CCur(txtPorcentaje.Text) < 0 Or CCur(txtPorcentaje.Text) > 999 Then
     MsgBox "Porcentaje no es válido", vbExclamation
     Exit Sub
  End If
End If

If Not IsNumeric(txtMemInicio.Text) Then
   MsgBox "Rango de Inicio la membresía no es válido", vbExclamation
   Exit Sub
End If

If Not IsNumeric(txtMemCorte.Text) Then
   MsgBox "Rango de Corte la membresía no es válido", vbExclamation
   Exit Sub
End If


pOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
pPlan = txtPlan.Text
pGarantia = cbo.ItemData(cbo.ListIndex)
pEstado = cboEstado.ItemData(cboEstado.ListIndex)

strBitacora = "Garantia de s/Ahorros, Linea: " & txtLinea.Text & ", Gar: " & pGarantia & ", Est: " & pEstado _
         & " Plan : " & pPlan & " Porcentaje : " & CCur(txtPorcentaje.Text) _
         & ", Mem.I: " & txtMemInicio.Text & ", Mem.C: " & txtMemCorte.Text

strSQL = "exec spCrd_Garantia_Ahorros_Registro '" & pGarantia & "','" & pEstado & "'," & txtLinea _
       & "," & txtMemInicio.Text & "," & txtMemCorte.Text & "," & chkPatrimonio.Value _
       & "," & pOperadora & ",'" & pPlan & "'," & CCur(txtPorcentaje.Text) & ",'" & glogon.Usuario & "'"

Select Case Index
 Case 0 'Agregar / Modificar
       
    strSQL = strSQL & ",'A'"
    Call ConectionExecute(strSQL)
       
    Call Bitacora("Registra", strBitacora)
    
 
 Case 1 'Elimina

    strSQL = strSQL & ",'E'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Elimina", strBitacora)
    
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_Click()
If vPaso Then Exit Sub

If cbo.ListCount = 0 Or cboEstado.ListCount = 0 Then Exit Sub
If cbo.Text = "" Or cboEstado.Text = "" Then Exit Sub

Call sbConsulta

End Sub


Private Sub cboEstado_Click()
If vPaso Then Exit Sub

If cbo.ListCount = 0 Or cboEstado.ListCount = 0 Then Exit Sub
If cbo.Text = "" Or cboEstado.Text = "" Then Exit Sub

Call sbConsulta

End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_Load()

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Inicio", 1200, vbCenter
    .Add , , "Corte", 1200, vbCenter
    .Add , , "Porcentaje", 1200, vbRightJustify
    .Add , , "Patrimonio", 1100, vbCenter
End With


Call Formularios(Me)
Call RefrescaTags(Me)


btnMov(0).Enabled = cmdAplicar.Enabled
btnMov(1).Enabled = cmdAplicar.Enabled

End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

fxValida = True

vMensaje = ""


If Len(vMensaje) > 0 Then
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function



Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

txtLinea.Text = Item.Text
txtPlan.Text = Item.SubItems(2)
txtPlanDesc.Text = Item.SubItems(3)
txtPorcentaje.Text = Item.SubItems(6)

txtMemInicio.Text = Item.SubItems(4)
txtMemCorte.Text = Item.SubItems(5)

chkPatrimonio.Value = IIf(Mid(Item.SubItems(7), 1, 1) = "S", xtpChecked, xtpUnchecked)

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

vPaso = True

strSQL = "select Garantia as 'IdX', rtrim(Descripcion) as ItmX" _
       & " from CRD_GARANTIA_TIPOS where formulario = 'F01'"
Call sbCbo_Llena_New(cbo, strSQL, False, True)


strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  afi_Estados_Persona"
Call sbCbo_Llena_New(cboEstado, strSQL, False, True)


strSQL = "select rtrim(cod_Operadora) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  fnd_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)


vPaso = False

Call cbo_Click


End Sub


