VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmUS_Aplicaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cartera de Aplicaciones"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   10215
      _Version        =   1441792
      _ExtentX        =   18018
      _ExtentY        =   11245
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Aplicaciones"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Bloqueos"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "Label1(0)"
      Item(1).Control(2)=   "txtVersion"
      Item(1).Control(3)=   "btnBloqueo_Registro"
      Item(1).Control(4)=   "btnBloqueo_Elimina"
      Item(1).Control(5)=   "Label1(4)"
      Item(1).Control(6)=   "dtpBloqueoFecha"
      Item(2).Caption =   "Actualizaciones"
      Item(2).ControlCount=   9
      Item(2).Control(0)=   "lswUpdates"
      Item(2).Control(1)=   "Label1(1)"
      Item(2).Control(2)=   "Label1(2)"
      Item(2).Control(3)=   "Label1(3)"
      Item(2).Control(4)=   "dtpUpdate"
      Item(2).Control(5)=   "btnUpdate"
      Item(2).Control(6)=   "txtUpdateVersion"
      Item(2).Control(7)=   "txtUpdateNotas"
      Item(2).Control(8)=   "btnUpdateEliimina"
      Begin XtremeSuiteControls.ListView lswUpdates 
         Height          =   3855
         Left            =   -69160
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1441792
         _ExtentX        =   16325
         _ExtentY        =   6800
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
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5895
         Left            =   -70000
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   6015
         _Version        =   1441792
         _ExtentX        =   10610
         _ExtentY        =   10398
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
      Begin XtremeSuiteControls.DateTimePicker dtpBloqueoFecha 
         Height          =   375
         Left            =   -63400
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441792
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5775
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   9495
         _Version        =   524288
         _ExtentX        =   16748
         _ExtentY        =   10186
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   486
         ScrollBars      =   2
         SpreadDesigner  =   "frmUS_Aplicaciones.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtVersion 
         Height          =   375
         Left            =   -63400
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
         _ExtentY        =   661
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
      Begin XtremeSuiteControls.FlatEdit txtUpdateVersion 
         Height          =   375
         Left            =   -69160
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441792
         _ExtentX        =   5106
         _ExtentY        =   661
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
      Begin XtremeSuiteControls.DateTimePicker dtpUpdate 
         Height          =   375
         Left            =   -65200
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1441792
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin XtremeSuiteControls.PushButton btnUpdate 
         Height          =   495
         Left            =   -63400
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441792
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Actualización"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmUS_Aplicaciones.frx":059B
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnUpdateEliimina 
         Height          =   495
         Left            =   -61720
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3201
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Elimina Selección"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmUS_Aplicaciones.frx":0CBB
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtUpdateNotas 
         Height          =   1215
         Left            =   -69160
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1441792
         _ExtentX        =   16325
         _ExtentY        =   2143
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBloqueo_Registro 
         Height          =   495
         Left            =   -63400
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Bloquear Versión"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmUS_Aplicaciones.frx":125F
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBloqueo_Elimina 
         Height          =   495
         Left            =   -63400
         TabIndex        =   17
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1441792
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Elimina Selección"
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
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmUS_Aplicaciones.frx":196B
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   4
         Left            =   -63400
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   -66040
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   135
         Index           =   2
         Left            =   -69880
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "Notas:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   -69880
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Versión:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   135
         Index           =   0
         Left            =   -63400
         TabIndex        =   5
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2143
         _ExtentY        =   238
         _StockProps     =   79
         Caption         =   "Versión:"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   345
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   7335
      _Version        =   1441792
      _ExtentX        =   12938
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicación"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10575
   End
End
Attribute VB_Name = "frmUS_Aplicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub btnBloqueo_Elimina_Click()
Dim i As Integer
Dim vRemoved As Boolean

On Error GoTo vError

vRemoved = False

With lsw.ListItems
For i = 1 To .Count
  If .Item(i).Checked Then
    strSQL = "delete US_APP_BLOCK where COD_LINEA = " & .Item(i).Text _
           & " and COD_APP = '" & cbo.ItemData(cbo.ListIndex) & "'"
    Call ConectionExecute(strSQL)

    vRemoved = True
  End If
  
Next i
     
End With

If vRemoved Then
    MsgBox "Bloqueos Eliminados Satisfactoriamente...", vbInformation
    Call cbo_Click
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnBloqueo_Registro_Click()
Dim i As Integer

On Error GoTo vError

If txtVersion.Text = "" Then Exit Sub
If cbo.ListCount = 0 Then Exit Sub

strSQL = "select isnull(max(cod_Linea),0) + 1 as 'Linea' from US_APP_BLOCK where cod_app = '" & cbo.ItemData(cbo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
i = rs!Linea
rs.Close

strSQL = "insert US_APP_BLOCK(COD_LINEA,COD_APP,FECHA_BLOQUEO,VERSION_BLOQUEADA,REGISTRO_FECHA,REGISTRO_USUARIO)" _
       & " values(" & i & ",'" & cbo.ItemData(cbo.ListIndex) & "','" & Format(dtpBloqueoFecha.Value, "yyyy/mm/dd") _
       & "','" & txtVersion.Text & "',getdate(),'" & glogon.Usuario & "')"
Call ConectionExecute(strSQL)

MsgBox "Bloqueo Realizado Satisfactoriamente...", vbInformation

Call cbo_Click

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnUpdate_Click()
On Error GoTo vError

If txtUpdateVersion.Text = "" Then Exit Sub
If cbo.ListCount = 0 Then Exit Sub

strSQL = "insert US_APP_UPDATES(COD_APP,VERSION,NOTAS_DESCARGA,FECHA_LIBERA,REGISTRO_FECHA,REGISTRO_USUARIO)" _
       & " values('" & cbo.ItemData(cbo.ListIndex) & "','" & Trim(txtUpdateVersion.Text) & "','" & Trim(txtUpdateNotas.Text) _
       & "','" & Format(dtpUpdate.Value, "yyyy/mm/dd") & "',getdate(),'" & glogon.Usuario & "')"
Call ConectionExecute(strSQL)

MsgBox "Registro de Actualización realizado Satisfactoriamente...", vbInformation

Call cbo_Click



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnUpdateEliimina_Click()
Dim i As Integer
Dim vRemoved As Boolean

On Error GoTo vError

vRemoved = False

With lswUpdates.ListItems
For i = 1 To .Count
  If .Item(i).Checked Then
    strSQL = "delete US_APP_UPDATES where COD_APP = '" & cbo.ItemData(cbo.ListIndex) & "' and VERSION = '" & .Item(i).Text & "'"
    Call ConectionExecute(strSQL)

    vRemoved = True
  End If
  
Next i
     
End With

If vRemoved Then
    MsgBox "Actualizaciones Eliminadas Satisfactoriamente...", vbInformation
    Call cbo_Click
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub
If cbo.ListCount <= 0 Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear
lswUpdates.ListItems.Clear

txtVersion.Text = ""
dtpBloqueoFecha.Value = fxFechaServidor
dtpUpdate.Value = dtpBloqueoFecha.Value

strSQL = "select TOP 100 * from US_APP_BLOCK where COD_APP = '" & cbo.ItemData(cbo.ListIndex) & "' order by Cod_Linea desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cod_Linea)
     itmX.SubItems(1) = Format(rs!Fecha_Bloqueo, "yyyy-mm-dd")
     itmX.SubItems(2) = rs!Version_Bloqueada
 rs.MoveNext
Loop
rs.Close

strSQL = "select TOP 100 * from US_APP_UPDATES where COD_APP = '" & cbo.ItemData(cbo.ListIndex) & "' order by Version desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswUpdates.ListItems.Add(, , rs!Version)
     itmX.SubItems(1) = Format(rs!Fecha_Libera, "yyyy-mm-dd")
     itmX.SubItems(2) = rs!Notas_Descarga
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

vModulo = 13

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 1200
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Version", lsw.Width - (1200 + 1500 + 250)
End With

With lswUpdates.ColumnHeaders
    .Clear
    .Add , , "Versión", 3500
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Notas", lswUpdates.Width - (5000 + 250)
End With

tcMain.Item(0).Selected = True

'Llena Grid
strSQL = "select COD_APP,descripcion,Activa from US_APP_BANK" _
      & " order by COD_APP"
Call sbCargaGrid(vGrid, 3, strSQL)

'Carga el Combo de Aplicativos
strSQL = "select rtrim(COD_APP) as 'IdX', rtrim(descripcion) as 'ItmX' from US_APP_BANK where Activa = 1" _
      & " order by COD_APP"
vPaso = True
    Call sbCbo_Llena_New(cbo, strSQL, False, True)
vPaso = False


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim vNuevo As Boolean
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0


On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

    strSQL = "select count(*) as Existe from US_APP_BANK where cod_app = '" & vGrid.Text & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe = 0 Then
      vNuevo = True
    Else
      vNuevo = False
    End If
    rs.Close


If vNuevo Then 'Insertar
  
  strSQL = "insert into US_APP_BANK(COD_APP,descripcion,Activa,Registro_Usuario,Registro_Fecha) values('" _
         & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',getdate())"
  
  Call ConectionExecute(strSQL)

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update US_APP_BANK set descripcion = '" & vGrid.Text & "',Activa = "
 vGrid.Col = 3
 strSQL = strSQL & vGrid.Value & " where COD_APP = '"

 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

If cbo.ListCount > 0 Then
    Call cbo_Click
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        strSQL = "delete US_APP_BANK where COD_APP = " & vGrid.Text
        Call ConectionExecute(strSQL)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




