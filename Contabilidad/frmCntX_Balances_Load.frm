VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCntX_Balances_Load 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Balances: Cargado para Empresa Consolidadora"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14085
   LinkTopic       =   "Form6"
   ScaleHeight     =   9240
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13815
      _Version        =   1572864
      _ExtentX        =   24368
      _ExtentY        =   13361
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
      Item(0).Caption =   "Carga de Balance"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gbArchivo"
      Item(0).Control(1)=   "vGrid"
      Item(1).Caption =   "Histórico"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "gHistorico"
      Item(1).Control(1)=   "cboHistorico"
      Item(1).Control(2)=   "Label2(1)"
      Item(1).Control(3)=   "btnHistorico"
      Begin XtremeSuiteControls.GroupBox gbArchivo 
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   13575
         _Version        =   1572864
         _ExtentX        =   23945
         _ExtentY        =   2566
         _StockProps     =   79
         Caption         =   "Carga de Archivo de Balance Contable"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   432
            Index           =   0
            Left            =   8400
            TabIndex        =   2
            ToolTipText     =   "Busca Archivo de Carga"
            Top             =   480
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   762
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCntX_Balances_Load.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   432
            Index           =   1
            Left            =   8880
            TabIndex        =   3
            ToolTipText     =   "Carga Archivo"
            Top             =   480
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   762
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCntX_Balances_Load.frx":0700
         End
         Begin XtremeSuiteControls.PushButton btnArchivo 
            Height          =   432
            Index           =   2
            Left            =   9360
            TabIndex        =   4
            ToolTipText     =   "Información del Archivo a Cargar"
            Top             =   480
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   762
            _StockProps     =   79
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCntX_Balances_Load.frx":0E19
         End
         Begin XtremeSuiteControls.PushButton btnInicializa 
            Height          =   525
            Left            =   11880
            TabIndex        =   12
            Top             =   480
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Inicializa"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCntX_Balances_Load.frx":1532
         End
         Begin XtremeSuiteControls.PushButton btnImportar 
            Height          =   525
            Left            =   10440
            TabIndex        =   13
            Top             =   480
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Importar"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCntX_Balances_Load.frx":1EBF
         End
         Begin XtremeSuiteControls.FlatEdit txtArchivo 
            Height          =   555
            Left            =   1920
            TabIndex        =   5
            Top             =   480
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11245
            _ExtentY        =   979
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
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption scStatus 
            Height          =   375
            Left            =   1920
            TabIndex        =   7
            Top             =   1080
            Visible         =   0   'False
            Width           =   11415
            _Version        =   1572864
            _ExtentX        =   20135
            _ExtentY        =   661
            _StockProps     =   14
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Seleccione el Archivo:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Index           =   10
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1572
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5295
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   13575
         _Version        =   524288
         _ExtentX        =   23945
         _ExtentY        =   9340
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_Balances_Load.frx":2681
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gHistorico 
         Height          =   6375
         Left            =   -69880
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   13575
         _Version        =   524288
         _ExtentX        =   23945
         _ExtentY        =   11245
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_Balances_Load.frx":2EEC
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboHistorico 
         Height          =   330
         Left            =   -68440
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
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
      Begin XtremeSuiteControls.PushButton btnHistorico 
         Height          =   435
         Left            =   -58240
         TabIndex        =   18
         ToolTipText     =   "Busca Archivo de Carga"
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Consultar"
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
         Picture         =   "frmCntX_Balances_Load.frx":3757
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   -69760
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Histórico"
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
   Begin XtremeSuiteControls.ComboBox cboUnidad 
      Height          =   330
      Left            =   1680
      TabIndex        =   9
      Top             =   1080
      Width           =   6615
      _Version        =   1572864
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
   Begin XtremeSuiteControls.PushButton btnImport 
      Height          =   435
      Left            =   8400
      TabIndex        =   19
      ToolTipText     =   "Importar el Balance Directamente de la Contabilidad Base"
      Top             =   1040
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   767
      _StockProps     =   79
      Caption         =   "Importar Conta-Base"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Balances_Load.frx":3E57
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   435
      Left            =   9960
      TabIndex        =   20
      ToolTipText     =   "Importar el Balance Directamente de la Contabilidad Base"
      Top             =   1040
      Width           =   615
      _Version        =   1572864
      _ExtentX        =   1085
      _ExtentY        =   767
      _StockProps     =   79
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Balances_Load.frx":455F
   End
   Begin XtremeSuiteControls.Label lblPeriodo 
      Height          =   375
      Left            =   10680
      TabIndex        =   11
      Top             =   1080
      Width           =   3255
      _Version        =   1572864
      _ExtentX        =   5741
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Setiembre 2025"
      BackColor       =   16777152
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Unidad"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carga de Balances de Unidades a la Consolidadora"
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
      Height          =   375
      Index           =   0
      Left            =   1800
      TabIndex        =   8
      Top             =   360
      Width           =   9855
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCntX_Balances_Load"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean

Dim mContabilidad As Long, mAnio As Long, mMes As Integer

Private Sub sbArchivoBusca()


With frmContenedor.CD
    .InitDir = "C:\"
    .DialogTitle = "Localice Archivo de Balance Contable [Microsoft EXCEL]"
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


Private Sub sbArchivoCarga()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset

Dim i As Integer, iCampos As Integer, vExiste As Integer

Dim pUnidad As String
Dim pCuenta As String, pConsolidadora As String, pDescripcion As String
Dim pSI As Currency, pDebitos As Currency, pCreditos As Currency, pSF As Currency, pTC As Currency


On Error GoTo vError

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If


Me.MousePointer = vbHourglass



scStatus.Visible = True
scStatus.Caption = "Cargado archivo, espere!"

Set rsExcel = Excel_Load(txtArchivo.Text, "Import")

'Verifica Estructura del Archivo

iCampos = 0
For i = 0 To rsExcel.Fields.Count - 1
   Select Case UCase(rsExcel.Fields(i).Name)
      Case "CUENTA", "CONSOLIDADORA", "DESCRIPCION", "SALDO_INICIAL", "DEBITOS", "CREDITOS", "SALDO_FINAL", "TC"
        iCampos = iCampos + 1
      Case Else

   End Select
Next i

If iCampos < 8 Then
   scStatus.Visible = False

   Me.MousePointer = vbDefault
   MsgBox "1. No coincide la estructura del archivo a cargar..." & vbCrLf & _
          "2. Los campos son CUENTA, CONSOLIDADORA, DESCRIPCION, SALDO_INICIAL, DEBITOS, CREDITOS, SALDO_FINAL, TC", vbExclamation
   Exit Sub
End If



Dim vCount As Long

vCount = 0
pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

Do While Not rsExcel.EOF
 pCuenta = Trim(rsExcel!Cuenta & "")
 pConsolidadora = Trim(rsExcel!CONSOLIDADORA & "")
 pDescripcion = Trim(rsExcel!DESCRIPCION & "")
 pSI = rsExcel!SALDO_INICIAL
 pDebitos = rsExcel!DEBITOS
 pCreditos = rsExcel!CREDITOS
 pSF = rsExcel!SALDO_FINAL
 pTC = rsExcel!TC
 vCount = vCount + 1

 scStatus.Caption = "Cargando archivo...Registro No." & vCount
 DoEvents

    strSQL = strSQL & Space(10) & "exec spCntX_Consolida_Balance_Importa_Cargado " & mContabilidad & ", '" & pUnidad & "', " & mAnio & ", " & mMes _
           & ", '" & pCuenta & "', '" & pConsolidadora & "', '" & pDescripcion _
           & "', " & pSI & ", " & pDebitos & ", " & pCreditos & ", " & pSF & ", " & pTC _
           & ", '" & glogon.Usuario & "', " & vCount
 
 
 'Inserta Valores
 If Len(strSQL) > 25000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
 End If

rsExcel.MoveNext
Loop
rsExcel.Close

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


scStatus.Caption = "Realizando la Auto Mapeo y Validación, espere!"
DoEvents

'Concilia y Actualiza
strSQL = "exec spCntX_Consolida_Balance_Importa_Mapeo " & mContabilidad & ", '" & pUnidad & "', " & mAnio & ", " & mMes _
       & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

scStatus.Caption = "Cargando Resultados, espere!"
DoEvents

'Carga Los Resultaods
strSQL = "exec spCntX_Consolida_Balance_Importa_Resultados " & mContabilidad & ", '" & pUnidad & "', " & mAnio & ", " & mMes _
       & ", '" & glogon.Usuario & "'"
Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)


scStatus.Visible = False

Me.MousePointer = vbDefault
MsgBox "Información Cargada Satisfactoriamente", vbInformation



Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub btnArchivo_Click(Index As Integer)
Dim vMensaje As String
  


Select Case Index
  
  Case 0 'buscar
        txtArchivo.Text = ""
        
        Call sbArchivoBusca

  Case 1 'Carga
       Call sbArchivoCarga
       
  Case 2 'Info
     vMensaje = "-> FORMATO DEL ARCHIVO DE CARGA <-" & vbCrLf & vbCrLf _
              & " 1. Microsoft Excel" & vbCrLf _
              & " 2. Nombre de la Hoja.: Import" & vbCrLf _
              & " 3. Columnas.: CUENTA, CONSOLIDADORA, DESCRIPCION, SALDO_INICIAL, DEBITOS, CREDITOS, SALDO_FINAL, TC"
     
     MsgBox vMensaje, vbInformation
     
     
End Select

End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 10
    vHeaders.Headers(1) = "Cuenta"
    vHeaders.Headers(2) = "Cta. Consolida"
    vHeaders.Headers(3) = "Descripción"
    vHeaders.Headers(4) = "Saldo Inicial"
    vHeaders.Headers(5) = "Total Débitos"
    vHeaders.Headers(6) = "Total Créditos"
    vHeaders.Headers(7) = "Saldo Final"
    vHeaders.Headers(8) = "Validación"
    vHeaders.Headers(9) = "Divisa"
    vHeaders.Headers(10) = "Tipo Cambio"
    
   If tcMain.SelectedItem = 0 Then
     Call sbSIFGridExportar(vGrid, vHeaders, "Contabilidad_Balances_" & cboUnidad.ItemData(cboUnidad.ListIndex) & " " & lblPeriodo.Caption)
   Else
     Call sbSIFGridExportar(gHistorico, vHeaders, "Contabilidad_Balances_" & cboUnidad.ItemData(cboUnidad.ListIndex) _
            & " " & lblPeriodo.Caption & "_H:" & cboHistorico.ItemData(cboHistorico.ListIndex))
   End If
End Sub

Private Sub btnHistorico_Click()

If cboHistorico.ListCount <= 0 Then
    Exit Sub
End If

Call sbHistorico_Consulta

End Sub

Private Sub btnImport_Click()
Dim i As Integer

Dim strSQL As String


i = MsgBox("Esta Seguro que desea Importar el Balance de la Contabilidad Base para Este Periodo" _
            & ", este proceso Reemplazará el actual?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If

On Error GoTo vError
Me.MousePointer = vbHourglass
'spCntX_Consolida_Importa_Conta_Base(@Consolidadora int, @Usuario varchar(30), @Anio int, @Mes smallint)
strSQL = "exec spCntX_Consolida_Importa_Conta_Base " & mContabilidad & ", '" & glogon.Usuario _
       & "', " & mAnio & ", " & mMes
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Importación de Balance desde la Contabilidad Base realizado satisfactoriamente!", vbInformation
   Call Bitacora("Aplica", "Importación del Balance de la Contabilidad Base de: " & mContabilidad _
            & "  " & lblPeriodo.Caption)
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnImportar_Click()
Dim i As Long, pUnidad As String

On Error GoTo vError

pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

strSQL = "exec spCntX_Consolida_Balance_Importa_Valida " & mContabilidad & ", '" & pUnidad & "', " & mAnio & ", " & mMes _
       & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Casos_Erroneos > 0 Then
    MsgBox "Existen " & rs!Casos_Erroneos & " Líneas Erroneas, verifiquelas primero antes de importarlas", vbExclamation
    Exit Sub
End If


i = MsgBox("Esta Seguro que desea Importar el Balance para este Periodo" _
            & ", este proceso Reemplazará el actual?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If


Me.MousePointer = vbHourglass

' spCntX_Consolida_Balance_Importa(@Consolidadora int, @Unidad varchar(10), @Anio int, @Mes smallint, @Usuario varchar(30))
strSQL = "exec spCntX_Consolida_Balance_Importa " & mContabilidad & ", '" & pUnidad _
       & "', " & mAnio & ", " & mMes & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Importación de Balance realizado satisfactoriamente!", vbInformation
   
   vGrid.MaxRows = 0
   txtArchivo.Text = ""
   
   Call Bitacora("Aplica", "Importación del Balance de la Contabilidad Id: [" & mContabilidad _
            & "]  " & lblPeriodo.Caption & " Unidad: " & pUnidad)
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnInicializa_Click()
Dim i As Long

On Error GoTo vError


i = MsgBox("Esta Seguro que desea Inicializar el Balance para este Periodo" _
            & ", este proceso Eliminará el actual?", vbYesNo)
If i = vbNo Then
    Exit Sub
End If


Me.MousePointer = vbHourglass

'spCntX_Consolida_Balance_Inicializa(@Consolidadora int, @Unidad varchar(10), @Anio int, @Mes smallint, @Usuario varchar(30))
strSQL = "exec spCntX_Consolida_Balance_Inicializa " & mContabilidad & ", '" & cboUnidad.ItemData(cboUnidad.ListIndex) _
       & "', " & mAnio & ", " & mMes & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

Me.MousePointer = vbDefault

If rs!Pass = 1 Then
   MsgBox "Balance Inicializado satisfactoriamente!", vbInformation
   
   vGrid.MaxRows = 0
   txtArchivo.Text = ""
   
   Call Bitacora("Aplica", "Inicialización del Balance de la Contabilidad Id: [" & mContabilidad _
            & "]  " & lblPeriodo.Caption & ", Unidad: " & cboUnidad.ItemData(cboUnidad.ListIndex))
Else
    MsgBox rs!Mensaje, vbExclamation
End If

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboUnidad_Click()
If vPaso Then Exit Sub

tcMain.Item(0).Selected = True
vGrid.MaxRows = 0

txtArchivo.Text = ""
scStatus.Caption = ""

Dim pUnidad As String

pUnidad = cboUnidad.ItemData(cboUnidad.ListIndex)

strSQL = "select isnull(I_CONSOLIDADORA, 0) as 'Consolida_Ind'" _
       & ", isnull(CONSOLIDA_CONTA_BASE, 0) as 'Consolida_Conta', isnull(CONSOLIDA_UNIDAD_BASE, '') as 'Consolida_Unidad'" _
       & " from CntX_Contabilidades where cod_contabilidad = " & mContabilidad
Call OpenRecordSet(rs, strSQL)

If rs!Consolida_Ind = 1 And rs!Consolida_Conta > 0 And pUnidad = rs!Consolida_Unidad Then
    btnImport.Visible = True
Else
    btnImport.Visible = False
End If

End Sub

Private Sub Form_Load()

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


mContabilidad = gCntX_Parametros.CodigoConta
mAnio = gCntX_Parametros.PeriodoAnio
mMes = gCntX_Parametros.PeriodoMes

lblPeriodo.Caption = fxCntX_PeriodoDesc(mAnio, mMes)


strSQL = "select Cod_Unidad as 'IdX', Descripcion as 'ItmX' from CntX_Unidades" _
      & "  where cod_Contabilidad = " & mContabilidad & " and Activa = 1"

vPaso = True
    Call sbCbo_Llena_New(cboUnidad, strSQL, False, True)
vPaso = False


Call cboUnidad_Click

tcMain.Item(0).Selected = True
vGrid.MaxRows = 0


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

tcMain.Width = Me.Width - 350
tcMain.Height = Me.Height - (tcMain.Top + 600)

vGrid.Width = tcMain.Width - 250
vGrid.Height = tcMain.Height - (vGrid.Top + 250)

gHistorico.Width = vGrid.Width
gHistorico.Height = tcMain.Height - (gHistorico.Top + 250)


End Sub

Private Sub sbHistorico_List()
Dim strSQL As String

On Error GoTo vError

gHistorico.MaxRows = 0

If cboUnidad.ListCount <= 0 Then
    Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spCntX_Balance_Cargado_Historico " & mContabilidad & ", '" & cboUnidad.ItemData(cboUnidad.ListIndex) _
       & "', " & mAnio & ", " & mMes
       
Call sbCbo_Llena_New(cboHistorico, strSQL, False, True)

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbHistorico_Consulta()
Dim strSQL As String

On Error GoTo vError
Me.MousePointer = vbHourglass

gHistorico.MaxRows = 0

strSQL = "exec spCntX_Balance_Cargado_Historico_Consulta " & cboHistorico.ItemData(cboHistorico.ListIndex)
Call sbCargaGrid(gHistorico, gHistorico.MaxCols, strSQL)

If gHistorico.MaxRows > 0 Then
    gHistorico.MaxRows = gHistorico.MaxRows - 1
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 1 Then
  Call sbHistorico_List
End If

End Sub
