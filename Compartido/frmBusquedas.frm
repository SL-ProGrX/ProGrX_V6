VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmBusquedas 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar por:"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   10200
   DrawWidth       =   2
   HelpContextID   =   9004
   Icon            =   "frmBusquedas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   5292
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   10092
      _Version        =   1441793
      _ExtentX        =   17801
      _ExtentY        =   9334
      _StockProps     =   77
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      GridLines       =   -1  'True
      FullRowSelect   =   -1  'True
      Appearance      =   16
      ShowBorder      =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCriterio 
      Height          =   372
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCriterio 
      Height          =   372
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCriterio 
      Height          =   372
      Index           =   2
      Left            =   6600
      TabIndex        =   3
      Top             =   840
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   720
      Top             =   6120
   End
   Begin XtremeSuiteControls.FlatEdit txtLineas 
      Height          =   312
      Left            =   9000
      TabIndex        =   8
      Top             =   6840
      Width           =   852
      _Version        =   1441793
      _ExtentX        =   1503
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "30"
      Alignment       =   2
      Appearance      =   2
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Left            =   6840
      TabIndex        =   9
      Top             =   6840
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Líneas de Resultado:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblCriterio 
      Height          =   252
      Index           =   2
      Left            =   6600
      TabIndex        =   7
      Top             =   600
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Criterio No. 3"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblCriterio 
      Height          =   252
      Index           =   1
      Left            =   3360
      TabIndex        =   6
      Top             =   600
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Criterio No. 2"
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
      Transparent     =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblCriterio 
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   3252
      _Version        =   1441793
      _ExtentX        =   5736
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Criterio No. 1"
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
      Transparent     =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption scSubTitulos 
      Height          =   1452
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10212
      _Version        =   1441793
      _ExtentX        =   18013
      _ExtentY        =   2561
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.98
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      VisualTheme     =   3
      Alignment       =   1
   End
End
Attribute VB_Name = "frmBusquedas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rs As New ADODB.Recordset, ColIndex As Integer


Private Sub Form_Load()

lblCriterio.Item(0).Visible = False
txtCriterio.Item(0).Visible = False
lblCriterio.Item(1).Visible = False
txtCriterio.Item(1).Visible = False
lblCriterio.Item(2).Visible = False
txtCriterio.Item(2).Visible = False

gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
gBusquedas.Resultado3 = ""

ColIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
 gBusquedas.Orden = ""
 gBusquedas.Columna = ""
 gBusquedas.Consulta = ""
 gBusquedas.Filtro = ""
 gBusquedas.Convertir = "S"
 gBusquedas.Mascara = ""
 
 gBusquedas.Col1Name = ""
 gBusquedas.Col2Name = ""
 gBusquedas.Col3Name = ""
 
End Sub

Private Function fxValidaCriterio(pCadena As String) As Boolean
Dim vResultado As Boolean, vMensaje As String

pCadena = UCase(pCadena)

vResultado = True
If pCadena <> "" Then
    If InStr(1, "SELECT", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "DELETE", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "UPDATE", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "INSERT", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "EXEC", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "DROP", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "CREATE", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "ALTER", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "SP_", pCadena) > 0 And vResultado Then vResultado = False
    If InStr(1, "'", pCadena) > 0 And vResultado Then vResultado = False
End If

If Not vResultado Then
 'Registrar en Log de Seguridad todo el criterio
 MsgBox "!Error: El criterio de busqueda contiene información o datos que pueden afectar potencialmente la integridad de la información..!", vbExclamation
End If

fxValidaCriterio = vResultado

End Function

Private Sub sbBuscar()
Dim strSQL As String, bWhere As Boolean


On Error GoTo vError

If Not fxValidaCriterio(txtCriterio.Item(0).Text) Then
   Exit Sub
End If
If Not fxValidaCriterio(txtCriterio.Item(1).Text) Then
   Exit Sub
End If
If Not fxValidaCriterio(txtCriterio.Item(2).Text) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

bWhere = False

strSQL = gBusquedas.Consulta

'Busqueda con Filtrado Inicial
If lblCriterio.Item(0).Visible = False Then
        If UCase(gBusquedas.Convertir) = "S" Or gBusquedas.Convertir = "" Then
            strSQL = strSQL & " Where CONVERT(varchar(200)," _
                   & gBusquedas.Columna & ")"
        Else
             strSQL = strSQL & " Where " & gBusquedas.Columna
        End If
        
        strSQL = strSQL & " like '%" & Format(txtCriterio.Item(0).Text, gBusquedas.Mascara) & "%'"
End If



If txtCriterio.Item(0).Visible And Len(txtCriterio.Item(0).Text) > 0 Then
  If bWhere Then
      strSQL = strSQL & " AND "
  Else
      strSQL = strSQL & " WHERE "
      bWhere = True
  End If
    
    If UCase(gBusquedas.Convertir) = "S" Or gBusquedas.Convertir = "" Then
        strSQL = strSQL & " CONVERT(varchar(200)," _
               & lblCriterio.Item(0).Tag & ") like '%" & txtCriterio.Item(0).Text & "%'"
    Else
        strSQL = strSQL & "isnull(" & lblCriterio.Item(0).Tag & ",'') like '%" & txtCriterio.Item(0).Text & "%'"
    End If
    
End If



If txtCriterio.Item(1).Visible Then 'And Len(txtCriterio.Item(1).Text) > 0
  If bWhere Then
      strSQL = strSQL & " AND "
  Else
      strSQL = strSQL & " WHERE "
      bWhere = True
  End If
    strSQL = strSQL & "isnull(" & lblCriterio.Item(1).Tag & ",'') like '%" & txtCriterio.Item(1).Text & "%'"
End If

If txtCriterio.Item(2).Visible And Len(txtCriterio.Item(2).Text) > 0 Then
  If bWhere Then
      strSQL = strSQL & " AND "
  Else
      strSQL = strSQL & " WHERE "
      bWhere = True
  End If
  strSQL = strSQL & "isnull(" & lblCriterio.Item(2).Tag & ",'') like '%" & txtCriterio.Item(2).Text & "%'"
End If



If Len(Trim(gBusquedas.Filtro)) > 0 Then strSQL = strSQL & " " & gBusquedas.Filtro

strSQL = strSQL & " Order by " & gBusquedas.Orden

Call sbCargaLsw(strSQL)

If txtCriterio.Item(ColIndex).Enabled And txtCriterio.Item(ColIndex).Visible Then txtCriterio.Item(ColIndex).SetFocus


Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault


End Sub

Private Sub sbFixTamañoEtiquetas(pNumero As Integer)
Dim vWidth As Long

If pNumero > 3 Then pNumero = 3

vWidth = (lsw.Width - 120) / pNumero


lblCriterio.Item(0).Visible = True
txtCriterio.Item(0).Visible = True
lblCriterio.Item(1).Visible = False
txtCriterio.Item(1).Visible = False
lblCriterio.Item(2).Visible = False
txtCriterio.Item(2).Visible = False


lblCriterio.Item(0).Width = vWidth
txtCriterio.Item(0).Width = vWidth

lblCriterio.Item(1).Width = vWidth
txtCriterio.Item(1).Width = vWidth

lblCriterio.Item(2).Width = vWidth
txtCriterio.Item(2).Width = vWidth

Select Case pNumero
  Case 2
        lblCriterio.Item(1).Visible = True
        txtCriterio.Item(1).Visible = True
        
        lblCriterio.Item(1).Left = vWidth + 105
        txtCriterio.Item(1).Left = vWidth + 105
        
  Case 3
        lblCriterio.Item(1).Visible = True
        txtCriterio.Item(1).Visible = True
        lblCriterio.Item(2).Visible = True
        txtCriterio.Item(2).Visible = True

        lblCriterio.Item(1).Left = vWidth + 105
        txtCriterio.Item(1).Left = vWidth + 105

        lblCriterio.Item(2).Left = (vWidth * 2) + 105
        txtCriterio.Item(2).Left = (vWidth * 2) + 105

End Select


End Sub

Private Sub sbCargaLsw(strSQL As String)
Dim i As Integer, itmX As ListViewItem
Dim x As Integer, y As Integer, IconX As Integer

On Error GoTo vError

lsw.ColumnHeaders.Clear
lsw.ListItems.Clear

strSQL = UCase(strSQL)

If IsNumeric(txtLineas.Text) Then
    strSQL = "SELECT TOP " & txtLineas.Text & " " & Trim(Mid(strSQL, 7, 2000))
Else
    strSQL = "SELECT TOP 30 " & Trim(Mid(strSQL, 7, 2000))
End If

Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then

    With lsw
     'Carga Titulos
     x = 0
     For i = 0 To (rs.Fields.Count - 1)
        Select Case i
            Case 0
               If gBusquedas.Col1Name = "" Then
                    .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
                    lblCriterio.Item(0).Caption = UCase(rs.Fields(i).Name)
               Else
                    .ColumnHeaders.Add (i + 1), , UCase(gBusquedas.Col1Name)
                    lblCriterio.Item(0).Caption = UCase(gBusquedas.Col1Name)
               End If
               lblCriterio.Item(0).Tag = UCase(rs.Fields(i).Name)
        
            Case 1
               If gBusquedas.Col2Name = "" Then
                    .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
                    lblCriterio.Item(1).Caption = UCase(rs.Fields(i).Name)
               Else
                    .ColumnHeaders.Add (i + 1), , UCase(gBusquedas.Col2Name)
                    lblCriterio.Item(1).Caption = UCase(gBusquedas.Col2Name)
               End If
               lblCriterio.Item(1).Tag = UCase(rs.Fields(i).Name)
        
            Case 2
               If gBusquedas.Col3Name = "" Then
                    .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
                    lblCriterio.Item(2).Caption = UCase(rs.Fields(i).Name)
               Else
                    .ColumnHeaders.Add (i + 1), , UCase(gBusquedas.Col3Name)
                    lblCriterio.Item(2).Caption = UCase(gBusquedas.Col3Name)
               End If
               lblCriterio.Item(2).Tag = UCase(rs.Fields(i).Name)
               
            Case Else
               .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
            
        End Select
        
        'Si el Texto de la Columna tiene mas caracteres que los valores
        'Que esta recibe, entonces utilizar el nombre de columna
        If Len(rs.Fields(i).Name) > rs.Fields(i).DefinedSize Then
            .ColumnHeaders.Item(i + 1).Width = Len(rs.Fields(i).Name) * 140
        Else
            .ColumnHeaders.Item(i + 1).Width = rs.Fields(i).DefinedSize * 95
        End If
        
        If i > 0 Then
            Select Case rs.Fields(i).Type
               Case adCurrency, adDecimal, adDouble, adNumeric
                  .ColumnHeaders.Item(i + 1).Alignment = lvwColumnRight
               Case Else
            End Select
        End If
        
        
     Next i
     
     'Activa/Rediseña y visualiza Criterios
     Call sbFixTamañoEtiquetas(lsw.ColumnHeaders.Count)
     
      Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs.Fields(0).Value)
        For i = 1 To (rs.Fields.Count - 1)
            itmX.SubItems(i) = rs.Fields(i).Value & ""
        Next i
        
      
        rs.MoveNext
      Loop
    End With

End If 'inicio y fin de tabla en true
rs.Close

Exit Sub

vError:
 If Err.Number = 3705 Then
    rs.Close
    Call sbCargaLsw(strSQL)
 Else
    MsgBox "No se encontraron datos para esta busqueda!", vbInformation
 End If

End Sub


Private Sub lsw_Click()
On Error Resume Next

If lsw.ListItems.Count <= 0 Then
    gBusquedas.Resultado = ""
    gBusquedas.Resultado2 = ""
    gBusquedas.Resultado3 = ""
Else
    gBusquedas.Resultado = lsw.SelectedItem.Text
End If

If lsw.ColumnHeaders.Count >= 2 And lsw.ListItems.Count > 0 Then gBusquedas.Resultado2 = lsw.SelectedItem.SubItems(1)
If lsw.ColumnHeaders.Count >= 3 And lsw.ListItems.Count > 0 Then gBusquedas.Resultado3 = lsw.SelectedItem.SubItems(2)

Unload Me

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error Resume Next

If lsw.ListItems.Count > 0 Then gBusquedas.Resultado = Item.Text
If lsw.ColumnHeaders.Count >= 2 Then gBusquedas.Resultado2 = Item.SubItems(1)
If lsw.ColumnHeaders.Count >= 3 Then gBusquedas.Resultado3 = Item.SubItems(2)
Unload Me


End Sub

Private Sub lsw_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyReturn And lsw.ListItems.Count > 0 Then
    gBusquedas.Resultado = lsw.SelectedItem.Text
    If lsw.ColumnHeaders.Count >= 2 Then gBusquedas.Resultado2 = lsw.SelectedItem.SubItems(1)
    If lsw.ColumnHeaders.Count >= 3 Then gBusquedas.Resultado3 = lsw.SelectedItem.SubItems(2)
    Unload Me
End If

End Sub



Private Sub TimerX_Timer()
 TimerX.Interval = 0
 TimerX.Enabled = False
 Call sbBuscar
End Sub

Private Sub txtCriterio_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ColIndex = Index
    Call sbBuscar
End If
End Sub

Private Sub txtLineas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 txtCriterio.Item(ColIndex).SetFocus
 Call sbBuscar
End If

vError:
End Sub

Private Sub txtLineas_LostFocus()
If Not IsNumeric(txtLineas.Text) Then
    txtLineas.Text = 30
End If
End Sub
