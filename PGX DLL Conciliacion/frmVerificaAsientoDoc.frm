VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmVerificaAsientoDoc 
   Caption         =   "Detección de Asientos Desbalanceados en el Auxiliar"
   ClientHeight    =   6195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   HelpContextID   =   7006
   LinkTopic       =   "Form1"
   ScaleHeight     =   6195
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   3492
      Left            =   0
      TabIndex        =   7
      Top             =   1080
      Width           =   5652
      _Version        =   1310723
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
   Begin MSComctlLib.ProgressBar prg 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5925
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton cmdBuscar 
      Height          =   492
      Left            =   5400
      TabIndex        =   1
      Top             =   240
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmVerificaAsientoDoc.frx":0000
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
      Left            =   2280
      TabIndex        =   3
      Top             =   480
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   550
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.PushButton cmdArchivo 
      Height          =   492
      Left            =   6840
      TabIndex        =   4
      Top             =   240
      Width           =   1452
      _Version        =   1310723
      _ExtentX        =   2561
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Archivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmVerificaAsientoDoc.frx":0A1E
   End
   Begin XtremeSuiteControls.Label lblEstado 
      Height          =   492
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   5412
      _Version        =   1310723
      _ExtentX        =   9546
      _ExtentY        =   868
      _StockProps     =   79
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
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   5
      Left            =   1560
      TabIndex        =   6
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   4
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   1212
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmVerificaAsientoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdArchivo_Click()

Call sbListViewExporFileTab(lsw)

End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem, lngRegistros As Long

Me.MousePointer = vbHourglass
lsw.ListItems.Clear

lblEstado.Caption = "Cargando Información (Espere)..."


    strSQL = "select A.Cod_Transaccion,A.Tipo_Documento, sum (case " _
           & " when A.Tipo_Movimiento = 'D' then  A.Monto" _
           & " when A.Tipo_Movimiento = 'C' then  0 end) as Debitos" _
           & " , sum (case " _
           & "    when A.Tipo_Movimiento = 'D' then  0" _
           & "    when A.Tipo_Movimiento = 'C' then  A.Monto end) as Creditos,Con.Descripcion as 'ConceptoDesc'" _
           & ",D.Registro_Usuario,D.Registro_Fecha" _
           & " From SIF_Transacciones D inner join SIF_Transacciones_Asiento A on D.Tipo_Documento = A.Tipo_Documento" _
           & " and D.Cod_Transaccion = A.Cod_Transaccion" _
           & " inner join SIF_Conceptos Con on D.cod_Concepto = Con.Cod_Concepto" _
           & " where D.Registro_Fecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" _
           & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59'" _
           & " group by A.Tipo_Documento,A.Cod_Transaccion,Con.Descripcion,D.Registro_Usuario,D.Registro_Fecha" _
           & " Having sum(case when A.Tipo_Movimiento = 'D' then A.Monto else 0 end) " _
           & " <> sum(case when A.Tipo_Movimiento <> 'D' then A.Monto else 0 end) "
Call OpenRecordSet(rs, strSQL)

lngRegistros = rs.RecordCount
prg.Max = rs.RecordCount + 1
prg.Value = 1

Do While Not rs.EOF
    If rs!Debitos <> rs!Creditos Then
       Set itmX = lsw.ListItems.Add(, , Format(rs!Registro_Fecha, "dd/mm/yyyy"))
           itmX.SubItems(1) = rs!Tipo_Documento
           itmX.SubItems(2) = rs!Cod_Transaccion
           itmX.SubItems(3) = Format(rs!Debitos, "Standard")
           itmX.SubItems(4) = Format(rs!Creditos, "Standard")
           itmX.SubItems(5) = Format(rs!Debitos - rs!Creditos, "Standard")
           itmX.SubItems(6) = rs!ConceptoDesc
           itmX.SubItems(7) = rs!Registro_Usuario & ""
    End If
    
  prg.Value = prg.Value + 1
  lblEstado.Caption = "Procesados  " & Format(prg.Value, "###,###,###,##0") _
          & "  De  " & Format(lngRegistros, "###,###,###,##0") & vbCrLf & " Porcentaje: " _
          & Round((prg.Value / lngRegistros) * 100, 2) & "%"
  
  If Right(CStr(prg.Value), 2) = "00" Then DoEvents
  
  rs.MoveNext

Loop
rs.Close

lblEstado.Caption = ""

Me.MousePointer = vbDefault

If lsw.ListItems.Count = 0 Then
 MsgBox "No se encontraron diferencias en Asientos x Documentos", vbInformation
End If

End Sub


Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdBuscar.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpCorte.SetFocus
End Sub

Private Sub Form_Load()

Set Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

With lsw.ColumnHeaders
    .Clear
    .Add , , "Fecha", 2100
    .Add , , "Tipo Doc.", 1800, vbCenter
    .Add , , "Transacción Id", 2200, vbCenter
    .Add , , "Débitos", 2100, vbRightJustify
    .Add , , "Crébitos", 2100, vbRightJustify
    .Add , , "Diferencia", 2100, vbRightJustify
    .Add , , "Concepto", 3100
    .Add , , "Usuario", 2100, vbCenter
    
End With

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

lsw.Width = Me.Width - 150
lsw.Height = Me.Height - (lsw.Top + lblEstado.Height + prg.Height + 480)
lblEstado.Top = lsw.Top + lsw.Height + 20
lblEstado.Width = lsw.Width

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub
