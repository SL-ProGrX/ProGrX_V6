VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCxPFusion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fusion de Proveedores"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   8550
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   8292
      _Version        =   1245185
      _ExtentX        =   14626
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Indique el Proveedor que Fusiona a los otros:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   38
         TabIndex        =   6
         ToolTipText     =   "Presione F4 Para Consultar"
         Top             =   360
         Width           =   5292
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         DataField       =   "e"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Presione F4 Para Consultar"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chkTrasladarCxP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Trasladar Facturas x Pagar al Nuevo Proveedor"
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   4
         Top             =   840
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox chkTrasladarCargos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Trasladar Cargos Flotantes al Nuevo Proveedor"
         Enabled         =   0   'False
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   612
         Left            =   6120
         TabIndex        =   8
         Top             =   960
         Width           =   2052
         _Version        =   1245185
         _ExtentX        =   3619
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aplicar la Fusión"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmCxPFusion.frx":0000
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
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
         Left            =   720
         TabIndex        =   7
         Top             =   360
         Width           =   1092
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00E0E0E0&
         X1              =   8280
         X2              =   600
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   7080
      Top             =   240
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3972
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   8292
      _ExtentX        =   14631
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   9596
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1658
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmCxPFusion.frx":07DE
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   7932
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   852
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCxPFusion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdAplicar_Click()
Dim strSQL As String, i As Byte, x As Integer

On Error GoTo vError


i = MsgBox("Esta seguro que desea fusionar los proveedores seleccionado en el nuevo...", vbYesNo)
If i = vbNo Then Exit Sub

If Len(Trim(txtCodigo)) = 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = ""

With lsw.ListItems
 For x = 1 To .Count
  If .Item(x).Checked Then
    strSQL = strSQL & Space(10) & "update cxp_proveedores set estado = 'I',fusion = dbo.MyGetdate()" _
           & " where cod_proveedor = " & .Item(x).SubItems(1)
   
    strSQL = strSQL & Space(10) & "insert cxp_fusiones(cod_proveedor,cod_proveedor_fus) values(" _
           & txtCodigo & "," & .Item(x).SubItems(1) & ")"
  End If
 Next x
End With

'Ejecuta el Lote
Call ConectionExecute(strSQL)



Me.MousePointer = vbDefault
MsgBox "Proveedores Fusionados Satisfactoriamente...", vbInformation

txtCodigo = ""
txtNombre = ""
Call Timer1_Timer

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub

Private Sub Form_Load()
vModulo = 30

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Timer1_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Timer1.Interval = 0

strSQL = "select cod_proveedor,descripcion from cxp_proveedores" _
       & " where estado = 'A' and fusion is null order by descripcion"
Call OpenRecordSet(rs, strSQL, 0)
lsw.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Descripcion)
     itmX.SubItems(1) = rs!cod_proveedor
 rs.MoveNext
Loop
rs.Close

vError:
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdAplicar.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub



