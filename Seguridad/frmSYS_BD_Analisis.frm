VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_BD_Analisis 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Analiza Estructura de Base de Datos"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11265
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.ListView lswEstructura 
      Height          =   3855
      Left            =   3600
      TabIndex        =   6
      Top             =   4080
      Width           =   7575
      _Version        =   1441793
      _ExtentX        =   13361
      _ExtentY        =   6800
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswResultado 
      Height          =   3375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   7575
      _Version        =   1441793
      _ExtentX        =   13361
      _ExtentY        =   5953
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ListView lswTabla 
      Height          =   6975
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   12303
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      View            =   3
      FullRowSelect   =   -1  'True
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   2880
      Top             =   0
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3372
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.FlatEdit txtTabla 
      Height          =   372
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   3012
      _Version        =   1441793
      _ExtentX        =   5313
      _ExtentY        =   656
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtSQL 
      Height          =   372
      Left            =   6720
      TabIndex        =   2
      Top             =   120
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7853
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.FlatEdit txtFiltro 
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3372
      _Version        =   1441793
      _ExtentX        =   5948
      _ExtentY        =   656
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
End
Attribute VB_Name = "frmSYS_BD_Analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pServer As String, pDataBase As String, pUser As String, pKey As String
Dim mCon As New ADODB.Connection, pApp As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub cbo_Click()

On Error Resume Next

pApp = cbo.Text

Select Case cbo.Text
  Case "CODEAS"
    pDataBase = "CODEAS_Migra"
    pServer = "progrx.centralus.cloudapp.azure.com"
    pUser = "Sys_Migracion"
    pKey = "/f0rDymK3yL0g1n."
  Case "CINGE"
    pDataBase = "CINGE_Migra"
    pServer = "progrx.centralus.cloudapp.azure.com"
    pUser = "Sys_Migracion"
    pKey = "/f0rDymK3yL0g1n."
  Case "OPTISOFT"
    pDataBase = "CODEAS_Migra"
    pServer = "progrx.centralus.cloudapp.azure.com"
    pUser = "Sys_Migracion"
    pKey = "/f0rDymK3yL0g1n."
  Case "SIBU"
    pDataBase = "CODEAS_Migra"
    pServer = "progrx.centralus.cloudapp.azure.com"
    pUser = "Sys_Migracion"
    pKey = "/f0rDymK3yL0g1n."
  Case "AXAPTA"
    pDataBase = "ASECCSSProd"
    pServer = "10.10.1.36"
    pUser = "soporte_systemlogic"
    pKey = "H$#0D$Cju0xAei(-V298"
End Select




vPaso = False
strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=" & pServer _
       & ";Database=" & pDataBase & ";APP=PGX_Portal_Admin;tcp:" & pServer _
       & "," & SIFGlobal.PuertosDisponibles & ";"

With mCon
  .Close
  .CommandTimeout = 15
  .Mode = adModeReadWrite
  .CursorLocation = adUseClient
  
  .Open strSQL, pUser, pKey
  .CommandTimeout = 360
End With

Call TimerX_Timer

End Sub

Private Sub Form_Load()

Me.BackColor = RGB(214, 234, 248)

cbo.Clear
cbo.AddItem "CODEAS"
cbo.AddItem "OPTISOFT"
cbo.AddItem "SIBU"
cbo.AddItem "CINGE"
cbo.AddItem "AXAPTA"

cbo.Text = "CODEAS"

Call cbo_Click

End Sub


Private Sub sbCarga_Tablas()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "  select name  from sys.objects " _
       & " where type = 'U'" _
       & " and name like '%" & Trim(txtFiltro.Text) & "%'" _
       & " order by name"

vPaso = True

lswTabla.ListItems.Clear
lswTabla.ColumnHeaders.Clear
lswTabla.ColumnHeaders.Add , , "Objeto", 3000

rs.Open strSQL, mCon, adOpenStatic
Do While Not rs.EOF
   Set itmX = lswTabla.ListItems.Add(, , rs!Name)
  rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Exit Sub
    
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbCargaResultados(pObjeto As String)
Dim i As Integer
Dim x As Integer, y As Integer, IconX As Integer

On Error GoTo vError

If vPaso Then Exit Sub

Me.MousePointer = vbHourglass

lswResultado.ColumnHeaders.Clear
lswResultado.ListItems.Clear


strSQL = "SELECT TOP 50 * from " & pObjeto
txtTabla.Text = pObjeto
txtSQL.Text = strSQL

rs.Open strSQL, mCon, adOpenStatic
If Not rs.BOF And Not rs.EOF Then

    With lswResultado
     'Carga Titulos
     x = 0
     For i = 0 To (rs.Fields.Count - 1)
        .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
        'Si el Texto de la Columna tiene mas caracteres que los valores
        'Que esta recibe, entonces utilizar el nombre de columna
'        If Len(rs.Fields(i).Name) > rs.Fields(i).DefinedSize Then
'            .ColumnHeaders.Item(i + 1).Width = Len(rs.Fields(i).Name) * 140
'        Else
'            .ColumnHeaders.Item(i + 1).Width = rs.Fields(i).DefinedSize * 95
'        End If
        
        
        .ColumnHeaders.Item(i + 1).Width = 1200
        If i > 0 Then
            Select Case rs.Fields(i).Type
               Case adCurrency, adDecimal, adDouble, adNumeric
                  .ColumnHeaders.Item(i + 1).Alignment = lvwColumnRight
               Case Else
            End Select
        End If
     Next i
     
      Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs.Fields(0).Value & "")
        For i = 1 To (rs.Fields.Count - 1)
            itmX.SubItems(i) = rs.Fields(i).Value & ""
        Next i
        rs.MoveNext
      Loop
    
    End With

End If 'inicio y fin de tabla en true
rs.Close


'Estructura
lswEstructura.ColumnHeaders.Clear
lswEstructura.ListItems.Clear

strSQL = "  select COLUMN_NAME AS 'COLUMNA', DATA_TYPE AS 'TIPO_DATO', IS_NULLABLE AS 'NULOS', isnull(CHARACTER_MAXIMUM_LENGTH,'') as 'TAMANO'" _
       & " from INFORMATION_SCHEMA.columns where table_name='" & pObjeto & "'"

rs.Open strSQL, mCon, adOpenStatic
If Not rs.BOF And Not rs.EOF Then

    With lswEstructura
     'Carga Titulos
     x = 0
     For i = 0 To (rs.Fields.Count - 1)
        .ColumnHeaders.Add (i + 1), , UCase(rs.Fields(i).Name)
        'Si el Texto de la Columna tiene mas caracteres que los valores
        'Que esta recibe, entonces utilizar el nombre de columna
'        If Len(rs.Fields(i).Name) > rs.Fields(i).DefinedSize Then
'            .ColumnHeaders.Item(i + 1).Width = Len(rs.Fields(i).Name) * 140
'        Else
'            .ColumnHeaders.Item(i + 1).Width = rs.Fields(i).DefinedSize * 95
'        End If
        
        
        
        
        .ColumnHeaders.Item(i + 1).Width = 1200
        
        If i > 0 Then
            Select Case rs.Fields(i).Type
               Case adCurrency, adDecimal, adDouble, adNumeric
                  .ColumnHeaders.Item(i + 1).Alignment = lvwColumnRight
               Case Else
            End Select
        End If
     Next i
     
      Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , rs.Fields(0).Value & "")
        For i = 1 To (rs.Fields.Count - 1)
            itmX.SubItems(i) = rs.Fields(i).Value & ""
        Next i
        rs.MoveNext
      Loop
    
    End With

End If 'inicio y fin de tabla en true
rs.Close

Me.MousePointer = vbDefault


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub Form_Resize()
On Error Resume Next

lswTabla.Height = Me.Height - (lswTabla.Top + 550)

lswResultado.Height = (lswTabla.Height / 2) - 100
lswResultado.Width = Me.Width - (lswTabla.Left + lswTabla.Width + 400)

txtSQL.Width = Me.Width - (txtSQL.Left + 400)

lswEstructura.Width = lswResultado.Width
lswEstructura.Height = lswResultado.Height + 50
lswEstructura.Top = lswResultado.Top + lswResultado.Height + 100


End Sub




Private Sub lswTabla_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Me.Caption = "Visualizando Tabla: " & UCase(Item.Text)

Dim i As Long, vItem As String

With lswTabla.ListItems
    For i = 1 To .Count
        .Item(i).Bold = False
    Next i
End With

Item.Bold = True
Call sbCargaResultados(Item.Text)

End Sub


Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0
Call sbCarga_Tablas

End Sub


Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Call sbCarga_Tablas
End If

End Sub
