VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmPres_Modelo_Cuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   144
      Left            =   0
      TabIndex        =   20
      Top             =   7428
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6372
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   10212
      _Version        =   1572864
      _ExtentX        =   18013
      _ExtentY        =   11239
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
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Cuentas"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "cboUnidad"
      Item(0).Control(2)=   "cboCentroCosto"
      Item(0).Control(3)=   "btnBuscar"
      Item(0).Control(4)=   "Label2(5)"
      Item(0).Control(5)=   "Label2(6)"
      Item(1).Caption =   "Importar"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gbImport"
      Item(1).Control(1)=   "tcImport"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4812
         Left            =   -69880
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   10092
         _Version        =   1572864
         _ExtentX        =   17801
         _ExtentY        =   8488
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbImport 
         Height          =   1332
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   9972
         _Version        =   1572864
         _ExtentX        =   17590
         _ExtentY        =   2350
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   1
         Begin XtremeSuiteControls.RadioButton rbModo 
            Height          =   264
            Index           =   0
            Left            =   720
            TabIndex        =   17
            Top             =   960
            Width           =   2292
            _Version        =   1572864
            _ExtentX        =   4043
            _ExtentY        =   466
            _StockProps     =   79
            Caption         =   "Modo 1: Vertical + Cortes"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnImportar 
            Height          =   495
            Left            =   8280
            TabIndex        =   12
            Top             =   360
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Importar"
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
            Picture         =   "frmPres_Modelo_Cuentas.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   492
            Left            =   4800
            TabIndex        =   13
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
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
            Picture         =   "frmPres_Modelo_Cuentas.frx":07C2
         End
         Begin XtremeSuiteControls.PushButton btnCargar 
            Height          =   492
            Left            =   6120
            TabIndex        =   14
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Cargar"
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
            Picture         =   "frmPres_Modelo_Cuentas.frx":11E0
         End
         Begin XtremeSuiteControls.FlatEdit txtArchivo 
            Height          =   492
            Left            =   720
            TabIndex        =   15
            Top             =   360
            Width           =   3972
            _Version        =   1572864
            _ExtentX        =   7006
            _ExtentY        =   868
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.RadioButton rbModo 
            Height          =   264
            Index           =   1
            Left            =   3120
            TabIndex        =   18
            Top             =   960
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   466
            _StockProps     =   79
            Caption         =   "Modo 2: Horizontal"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.PushButton btnRevisar 
            Height          =   495
            Left            =   7665
            TabIndex        =   24
            ToolTipText     =   "Revisar el listado a cargar"
            Top             =   360
            Width           =   630
            _Version        =   1572864
            _ExtentX        =   1111
            _ExtentY        =   873
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
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmPres_Modelo_Cuentas.frx":1BA3
         End
         Begin VB.Label lblStatus 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Left            =   5280
            TabIndex        =   19
            Top             =   960
            Width           =   4212
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Archivo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   480
            Width           =   612
         End
      End
      Begin XtremeSuiteControls.ComboBox cboUnidad 
         Height          =   312
         Left            =   -68200
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   3972
         _Version        =   1572864
         _ExtentX        =   7011
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
      Begin XtremeSuiteControls.ComboBox cboCentroCosto 
         Height          =   312
         Left            =   -68200
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   3972
         _Version        =   1572864
         _ExtentX        =   7011
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
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   612
         Left            =   -61960
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
         _ExtentY        =   1080
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
         Picture         =   "frmPres_Modelo_Cuentas.frx":2381
      End
      Begin XtremeSuiteControls.TabControl tcImport 
         Height          =   4584
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   9972
         _Version        =   1572864
         _ExtentX        =   17590
         _ExtentY        =   8086
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
         ItemCount       =   2
         Item(0).Caption =   "Datos"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lswImport"
         Item(1).Caption =   "Inconsistencias"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswInco"
         Begin XtremeSuiteControls.ListView lswImport 
            Height          =   4095
            Left            =   0
            TabIndex        =   22
            Top             =   360
            Width           =   9975
            _Version        =   1572864
            _ExtentX        =   17595
            _ExtentY        =   7223
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
         Begin XtremeSuiteControls.ListView lswInco 
            Height          =   4095
            Left            =   -70000
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   9975
            _Version        =   1572864
            _ExtentX        =   17595
            _ExtentY        =   7223
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
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Centro de Costo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   6
         Left            =   -69880
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad de Negocio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   -69880
         TabIndex        =   9
         Top             =   480
         Visible         =   0   'False
         Width           =   1692
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   960
      Top             =   360
   End
   Begin XtremeSuiteControls.ComboBox cboModelo 
      Height          =   312
      Left            =   3480
      TabIndex        =   0
      Top             =   600
      Width           =   6492
      _Version        =   1572864
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.ComboBox cboContabilidad 
      Height          =   312
      Left            =   3480
      TabIndex        =   1
      Top             =   240
      Width           =   6492
      _Version        =   1572864
      _ExtentX        =   11456
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
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
      Height          =   252
      Index           =   0
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
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
      Height          =   252
      Index           =   10
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmPres_Modelo_Cuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub btnArchivo_Click()

With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo para Importar Presupuesto [Microsoft EXCEL]"
    .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
    .ShowOpen

    If .FileName = "" Then
        MsgBox "Archivo no válido...", vbExclamation
        Exit Sub
    End If

    If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
        'Ok
    Else
        MsgBox "La Extensión del Archivo no es válido...", vbExclamation
        Exit Sub
    End If

    
    txtArchivo.Text = .FileName
End With


End Sub

Private Sub btnBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim pUnidad As String, pCentroCosto As String

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

lsw.ListItems.Clear

pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

If cboCentroCosto.ListCount = 0 Then
  pCentroCosto = ""
Else
  pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If

strSQL = "exec spPres_CuentasCatalogo " & cboContabilidad.ItemData(cboContabilidad.ListIndex) & ",'" _
                & cboModelo.ItemData(cboModelo.ListIndex) _
                & "','" & pUnidad & "','" & pCentroCosto & "'"
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Cuenta)
        itmX.Tag = rs!cod_cuenta
        itmX.SubItems(1) = rs!Descripcion
        itmX.Checked = rs!Asignada
  
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

Private Sub btnCargar_Click()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim itmX As ListViewItem

Dim i As Integer, iCampos As Integer, vExiste As Integer
Dim vCuenta As String, pCamposTotal As Integer

On Error GoTo vError

tcImport.Item(0).Selected = True

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass

lswImport.ListItems.Clear

Set rsExcel = Excel_Load(txtArchivo.Text, "Presupuesto")

'Verifica Estructura del Archivo



iCampos = 0
Select Case True
 Case rbModo.Item(0).Value 'Formato Mensual por Cortes
    pCamposTotal = 6

    For i = 0 To rsExcel.Fields.Count - 1
       Select Case UCase(rsExcel.Fields(i).Name)
          Case "CUENTA", "DESCRIPCION", "UNIDAD", "CENTRO", "MONTO", "CORTE"
            iCampos = iCampos + 1
          Case Else
          
       End Select
    Next i
    
    If iCampos < pCamposTotal Then
'       rsExcel.Close
       Me.MousePointer = vbDefault
       MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
              "2. Los campos son Cuenta, Descripcion, Unidad, Centro, Monto, Corte" & vbCrLf & _
              "3. Nombre de la hoja: PRESUPUESTO", vbExclamation
       Exit Sub
    End If

 
 Case rbModo.Item(1).Value 'Formato Horizontal Anual
    pCamposTotal = 16

    For i = 0 To rsExcel.Fields.Count - 1
       Select Case UCase(rsExcel.Fields(i).Name)
          Case "CUENTA", "DESCRIPCION", "UNIDAD", "CENTRO", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
            iCampos = iCampos + 1
          Case Else
          
       End Select
    Next i
    
    If iCampos < pCamposTotal Then
'       rsExcel.Close
       Me.MousePointer = vbDefault
       MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
              "2. Los campos son Cuenta, Descripcion, Unidad, Centro, Enero, Febrero, Marzo...Setiembre, Octubre, Noviembre,Diciembre" & vbCrLf & _
              "3. Nombre de la hoja: PRESUPUESTO", vbExclamation
        
       Exit Sub
    End If

End Select


lblStatus.Caption = "Cargando..."
DoEvents

With lswImport.ListItems

Do While Not rsExcel.EOF

  If Not IsNull(rsExcel!Cuenta) Then
    Set itmX = .Add(, , rsExcel!Cuenta & "")
        itmX.SubItems(1) = rsExcel!Descripcion & ""
        itmX.SubItems(2) = rsExcel!Unidad & ""
        itmX.SubItems(3) = Format(rsExcel!Centro, "00") & ""
    
    Select Case True
        Case rbModo.Item(0).Value 'Cortes
            itmX.SubItems(4) = Format(rsExcel!Monto, "Standard")
            itmX.SubItems(5) = Format(rsExcel!Corte, "yyyy-mm-dd")
        
        Case rbModo.Item(1).Value 'Horizontal
            itmX.SubItems(4) = Format(rsExcel!Enero, "Standard")
            itmX.SubItems(5) = Format(rsExcel!Febrero, "Standard")
            itmX.SubItems(6) = Format(rsExcel!Marzo, "Standard")
            itmX.SubItems(7) = Format(rsExcel!Abril, "Standard")
            itmX.SubItems(8) = Format(rsExcel!Mayo, "Standard")
            itmX.SubItems(9) = Format(rsExcel!Junio, "Standard")
            itmX.SubItems(10) = Format(rsExcel!Julio, "Standard")
            itmX.SubItems(11) = Format(rsExcel!Agosto, "Standard")
            itmX.SubItems(12) = Format(rsExcel!Setiembre, "Standard")
            itmX.SubItems(13) = Format(rsExcel!Octubre, "Standard")
            itmX.SubItems(14) = Format(rsExcel!Noviembre, "Standard")
            itmX.SubItems(15) = Format(rsExcel!Diciembre, "Standard")
    
    End Select

  End If
   
  rsExcel.MoveNext
Loop
rsExcel.Close

End With
        

lblStatus.Caption = ""
Me.MousePointer = vbDefault


MsgBox "Información Cargada Satisfactoriamente", vbInformation


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbImportar_Carga_Main()
Dim i As Long, x As Integer, strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String
Dim pUnidad As String, pCentroCosto As String, pCuenta As String

Dim pCorte As Date, pMonto As Currency, pInicializa As Integer

Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)


Dim cEnero As Date, cFebrero As Date, cMarzo As Date, cAbril As Date, cMayo As Date
Dim cJunio As Date, cJulio As Date, cAgosto As Date, cSetiembre As Date
Dim cOctubre As Date, cNoviembre As Date, cDiciembre As Date


'Coloca los meses segun el periodo fiscal
strSQL = "exec spCntX_Periodo_Fiscal_Meses " & pContabilidad & ",0,'" & pModelo & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Select Case rs!Mes
    Case 1
        cEnero = rs!Corte
    Case 2
        cFebrero = rs!Corte
    Case 3
        cMarzo = rs!Corte
    Case 4
        cAbril = rs!Corte
    Case 5
        cMayo = rs!Corte
    Case 6
        cJunio = rs!Corte
    Case 7
        cJulio = rs!Corte
    Case 8
        cAgosto = rs!Corte
    Case 9
        cSetiembre = rs!Corte
    Case 10
        cOctubre = rs!Corte
    Case 11
        cNoviembre = rs!Corte
    Case 12
        cDiciembre = rs!Corte
  End Select
  rs.MoveNext
Loop
rs.Close

With lswImport.ListItems

 ProgressBarX.Visible = True
 ProgressBarX.Max = .Count + 1


lblStatus.Caption = "Importando en ESPEJO el Presupuesto..."
DoEvents

strSQL = ""
For i = 1 To .Count
    ProgressBarX.Value = i
    DoEvents
    
    If i = 1 Then
      pInicializa = 1
    Else
      pInicializa = 0
    End If

    pCuenta = .Item(i).Text
    pUnidad = .Item(i).SubItems(2)
    pCentroCosto = .Item(i).SubItems(3)
    
    If rbModo.Item(0).Value Then     'Subida por Cortes
        pMonto = CCur(.Item(i).SubItems(4))
        pCorte = CDate(.Item(i).SubItems(5))
        
        'Registra la Cuenta
        strSQL = strSQL & Space(10) & "exec spPres_Presupuesto_Import_Load '" & pModelo & "'," _
               & pContabilidad & ",'" & pCuenta & "','" & pUnidad & "','" & pCentroCosto _
               & "','" & Format(pCorte, "yyyy/mm/dd") & "'," & pMonto & ",'" & glogon.Usuario & "'," & pInicializa
    End If
    
    
    If rbModo.Item(1).Value Then     'Subida Horizontal
        
        For x = 4 To lswImport.ColumnHeaders.Count - 1
          If IsNumeric(.Item(i).SubItems(x)) Then
            pMonto = CCur(.Item(i).SubItems(x))
          Else
            pMonto = 0
          End If
          
          Select Case UCase(lswImport.ColumnHeaders.Item(x + 1).Text)
                Case "ENERO"
                    pCorte = cEnero
                Case "FEBRERO"
                    pCorte = cFebrero
                Case "MARZO"
                    pCorte = cMarzo
                Case "ABRIL"
                    pCorte = cAbril
                Case "MAYO"
                    pCorte = cMayo
                Case "JUNIO"
                    pCorte = cJunio
                Case "JULIO"
                    pCorte = cJulio
                Case "AGOSTO"
                    pCorte = cAgosto
                Case "SETIEMBRE", "SEPTIEMBRE"
                    pCorte = cSetiembre
                Case "OCTUBRE"
                    pCorte = cOctubre
                Case "NOVIEMBRE"
                    pCorte = cNoviembre
                Case "DICIEMBRE"
                    pCorte = cDiciembre
          End Select

        strSQL = strSQL & Space(10) & "exec spPres_Presupuesto_Import_Load '" & pModelo & "'," _
               & pContabilidad & ",'" & pCuenta & "','" & pUnidad & "','" & pCentroCosto _
               & "','" & Format(pCorte, "yyyy/mm/dd") & "'," & pMonto & ",'" & glogon.Usuario & "'," & pInicializa
        
        If pInicializa = 1 Then
            pInicializa = 0
        End If
        
        Next x
        
    End If
    
    
    If Len(strSQL) > 20000 Then
          Call ConectionExecute(strSQL)
          strSQL = ""
    End If

Next i

'Lote Final
If Len(strSQL) > 0 Then
      Call ConectionExecute(strSQL)
      strSQL = ""
End If
End With



ProgressBarX.Value = 0
ProgressBarX.Visible = False

lblStatus.Caption = ""

Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbImportar_Carga_Revisa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String

Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass


pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)


'Sube el Presupuesto
Call sbImportar_Carga_Main


lblStatus.Caption = "Revisando el Presupuesto..."
DoEvents

'Procesa Revision de la Carga
strSQL = "exec spPres_Presupuesto_Import_Revisa '" & pModelo & "'," _
           & pContabilidad & ",'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

With lswInco.ListItems
    .Clear
    Do While Not rs.EOF
    
        Set itmX = .Add(, , rs!cod_cuenta)
            itmX.SubItems(1) = rs!Descripcion
            itmX.SubItems(2) = rs!Cod_Unidad
            itmX.SubItems(3) = rs!cod_Centro_Costo
            itmX.SubItems(4) = Format(rs!Presupuesto, "Standard")
            itmX.SubItems(5) = Format(rs!Corte, "yyyy-mm-dd")
            itmX.SubItems(6) = rs!Detalle
        
        rs.MoveNext
    Loop
    rs.Close
    
End With


lblStatus.Caption = ""

Me.MousePointer = vbDefault
MsgBox "Revisión Finalizada!", vbInformation

tcImport.Item(1).Selected = True


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnImportar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pContabilidad As Long, pModelo As String

'Dim i As Long, x As Integer
'Dim pUnidad As String, pCentroCosto As String, pCuenta As String
'Dim iMes As Integer, iAnio As Integer, cMes As Integer, cAnio As Integer
'Dim pCorte As Date, pMonto As Currency, pInicializa As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass


pContabilidad = cboContabilidad.ItemData(cboContabilidad.ListIndex)
pModelo = cboModelo.ItemData(cboModelo.ListIndex)


'Sube el Presupuesto
Call sbImportar_Carga_Main


lblStatus.Caption = "Mapeando Cuentas..."
DoEvents

'Creando Mapeo
strSQL = "exec spPres_Presupuesto_Import_Mapeo '" & pModelo & "'," & pContabilidad & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


lblStatus.Caption = "Procesando Importación del Presupuesto..."
DoEvents

ProgressBarX.Visible = True

strSQL = "exec spPres_Presupuesto_Import_Procesa '" & pModelo & "'," & pContabilidad & ",'" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
Do While rs!Pendientes > 0
    

    ProgressBarX.Max = rs!Pendientes + rs!Procesados
    ProgressBarX.Value = rs!Procesados
    
    lblStatus.Caption = "Importando [" & rs!Procesados & " de " & rs!Pendientes + rs!Procesados & "]"
    DoEvents

    strSQL = "exec spPres_Presupuesto_Import_Procesa '" & pModelo & "'," & pContabilidad & ",'" & glogon.Usuario & "'"
    Call OpenRecordSet(rs, strSQL)
Loop
rs.Close


lblStatus.Caption = "Mapeando Cuentas sin Centro de Costo..."
DoEvents

strSQL = "exec spPres_MapeaCuentasSinCentroCosto '" & pModelo & "'," & pContabilidad & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


ProgressBarX.Value = 0
ProgressBarX.Visible = False

lblStatus.Caption = ""

Me.MousePointer = vbDefault
MsgBox "Importación del Presupuesto realizado satisfactoriamente!", vbInformation

lswImport.ListItems.Clear
txtArchivo.Text = ""

Call btnBuscar_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRevisar_Click()
Call sbImportar_Carga_Revisa
End Sub

Private Sub cboCentroCosto_Click()
If vPaso Then Exit Sub
lsw.ListItems.Clear

End Sub

Private Sub cboContabilidad_Click()
If vPaso Then Exit Sub


Dim strSQL As String

lsw.ListItems.Clear


vPaso = True

strSQL = "select P.cod_modelo as 'IdX' , P.DESCRIPCION as 'ItmX'" _
       & " From PRES_MODELOS P INNER JOIN PRES_MODELOS_USUARIOS Pmu on P.cod_Contabilidad = Pmu.cod_contabilidad" _
       & "  and P.cod_Modelo = Pmu.cod_Modelo and Pmu.Usuario = '" & glogon.Usuario & "'" _
       & " inner join CNTX_CIERRES Cc on P.cod_Contabilidad = Cc.cod_Contabilidad and P.ID_CIERRE = Cc.ID_CIERRE " _
       & " Where P.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex) _
       & " order by Cc.Inicio_Anio desc"
Call sbCbo_Llena_New(cboModelo, strSQL, False, True)

strSQL = "select Cu.Cod_Unidad as 'IdX' , Cu.DESCRIPCION as 'ItmX'" _
       & " From CNTX_UNIDADES Cu INNER JOIN PRES_USUARIOS_NIVEL Pun on Cu.cod_Contabilidad = Pun.cod_contabilidad" _
       & "  and Cu.cod_Unidad = Pun.Cod_Unidad and Pun.Usuario = '" & glogon.Usuario & "'" _
       & " Where Cu.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)
Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)

vPaso = False

End Sub

Private Sub cboModelo_Click()
If vPaso Then Exit Sub
lsw.ListItems.Clear

End Sub

Private Sub cboUnidad_Click()
If vPaso Then Exit Sub


Dim strSQL As String


vPaso = True

strSQL = "select Cc.COD_CENTRO_COSTO as 'IdX'  , Cc.DESCRIPCION as 'ItmX'" _
       & " from CNTX_CENTRO_COSTOS Cc inner join CNTX_UNIDADES_CC Uc on Cc.COD_CONTABILIDAD = Cc.COD_CONTABILIDAD" _
       & "     and Uc.COD_UNIDAD = '" & cboUnidad.ItemData(cboUnidad.ListIndex) & "'" _
       & "     and Cc.COD_CENTRO_COSTO = Uc.COD_CENTRO_COSTO" _
       & " Where Cc.COD_CONTABILIDAD = " & cboContabilidad.ItemData(cboContabilidad.ListIndex)

Call sbCbo_Llena_New(cboCentroCosto, strSQL, False, True)

lsw.ListItems.Clear

vPaso = False

Call btnBuscar_Click

End Sub

Private Sub Form_Load()
vModulo = 12

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

lsw.ColumnHeaders.Add , , "Cuenta", 2400
lsw.ColumnHeaders.Add , , "Descripcion", 5400

tcMain.Item(0).Selected = True

End Sub

Private Sub Form_Resize()
On Error Resume Next

tcMain.Height = Me.Height - (tcMain.Top + ProgressBarX.Height + 760)
tcMain.Width = Me.Width - 350

lsw.Width = tcMain.Width - 150
lsw.Height = tcMain.Height - (lsw.Top + 150)

tcImport.Width = tcMain.Width - 150
tcImport.Height = tcMain.Height - (lsw.Top + 150)

lswImport.Width = tcImport.Width - 150
lswImport.Height = tcImport.Height - (lswImport.Top + 150)

lswInco.Width = lswImport.Width
lswInco.Height = lswImport.Height

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

Dim strSQL As String, pCentroCosto As String

On Error GoTo vError

If cboCentroCosto.ListCount = 0 Then
  pCentroCosto = ""
Else
  pCentroCosto = cboCentroCosto.ItemData(cboCentroCosto.ListIndex)
End If

If Item.Checked = True Then
    strSQL = "exec spPres_PresupuestoInicialCrea '" & cboModelo.ItemData(cboModelo.ListIndex) & "'," _
        & cboContabilidad.ItemData(cboContabilidad.ListIndex) & ",'" & cboUnidad.ItemData(cboUnidad.ListIndex) _
        & "','" & pCentroCosto & "','" & Item.Tag & "',0,'" & glogon.Usuario & "'"
Else
    strSQL = "exec spPres_CuentasExcluye '" & cboModelo.ItemData(cboModelo.ListIndex) & "'," _
        & cboContabilidad.ItemData(cboContabilidad.ListIndex) & ",'" & cboUnidad.ItemData(cboUnidad.ListIndex) _
        & "','" & pCentroCosto & "','" & Item.Tag & "'"

End If
Call ConectionExecute(strSQL)

Exit Sub
                                    
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
                                    
End Sub

Private Sub lswImport_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswImport.SortKey = ColumnHeader.Index - 1
  If lswImport.SortOrder = 0 Then lswImport.SortOrder = 1 Else lswImport.SortOrder = 0
  lswImport.Sorted = True
End Sub

Private Sub rbModo_Click(Index As Integer)


With lswImport.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2100
    .Add , , "Descripción", 4100
    .Add , , "Unidad", 1400, vbCenter
    .Add , , "Centro", 1400, vbCenter
    
    Select Case True
        Case rbModo(0).Value
            .Add , , "Monto", 2100, vbRightJustify
            .Add , , "Corte", 2100, vbCenter
        
        Case rbModo(1).Value
            .Add , , "Enero", 2100, vbRightJustify
            .Add , , "Febrero", 2100, vbRightJustify
            .Add , , "Marzo", 2100, vbRightJustify
            .Add , , "Abril", 2100, vbRightJustify
            .Add , , "Mayo", 2100, vbRightJustify
            .Add , , "Junio", 2100, vbRightJustify
            .Add , , "Julio", 2100, vbRightJustify
            .Add , , "Agosto", 2100, vbRightJustify
            .Add , , "Setiembre", 2100, vbRightJustify
            .Add , , "Octubre", 2100, vbRightJustify
            .Add , , "Noviembre", 2100, vbRightJustify
            .Add , , "Diciembre", 2100, vbRightJustify
    End Select

End With

With lswInco.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2100
    .Add , , "Descripción", 4100
    .Add , , "Unidad", 1400, vbCenter
    .Add , , "Centro", 1400, vbCenter
    .Add , , "Monto", 2100, vbRightJustify
    .Add , , "Corte", 2100, vbCenter
    .Add , , "Detalle", 3100
End With


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)


Select Case Item.Index
Case 0 'Cuentas
  If cboModelo.ListCount > 0 Then
      Call btnBuscar_Click
  End If
  
Case 1 'Importar
  txtArchivo.Text = ""
  Call rbModo_Click(0)
  
End Select

Call Form_Resize

End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

 vPaso = True
    strSQL = "select cod_contabilidad as 'IdX', Nombre as 'ItmX' from CNTX_Contabilidades" _
           & " order by cod_contabilidad"
    Call sbCbo_Llena_New(cboContabilidad, strSQL, False, True)
 vPaso = False

Call cboContabilidad_Click

End Sub

