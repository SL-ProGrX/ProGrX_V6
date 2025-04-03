VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_Polizas_CargaLote 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pólizas: Cargado en Lote"
   ClientHeight    =   7875
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11415
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   9480
      TabIndex        =   0
      Top             =   1320
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Polizas_CargaLote.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4212
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   11172
      _Version        =   524288
      _ExtentX        =   19706
      _ExtentY        =   7430
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
      MaxCols         =   495
      SpreadDesigner  =   "frmCR_Polizas_CargaLote.frx":0700
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton btnCargar 
      Height          =   372
      Left            =   9960
      TabIndex        =   2
      Top             =   1320
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Polizas_CargaLote.frx":0DE3
   End
   Begin XtremeSuiteControls.PushButton btnInfo 
      Height          =   372
      Left            =   10440
      TabIndex        =   3
      Top             =   1320
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Polizas_CargaLote.frx":14FC
   End
   Begin XtremeSuiteControls.FlatEdit txtArchivo 
      Height          =   372
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12086
      _ExtentY        =   656
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
   Begin XtremeSuiteControls.ComboBox cboPrideduc 
      Height          =   312
      Left            =   4200
      TabIndex        =   5
      Top             =   1800
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cboCliente 
      Height          =   312
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   312
      Left            =   2520
      TabIndex        =   11
      Top             =   840
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   4200
      TabIndex        =   15
      Top             =   7080
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboCuenta 
      Height          =   312
      Left            =   4200
      TabIndex        =   16
      Top             =   7440
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
      Height          =   312
      Left            =   6600
      TabIndex        =   17
      Top             =   6720
      Width           =   1812
      _Version        =   1441793
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.ComboBox cboConfirma 
      Height          =   312
      Left            =   2520
      TabIndex        =   22
      Top             =   480
      Width           =   6852
      _Version        =   1441793
      _ExtentX        =   12091
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
   Begin XtremeSuiteControls.CheckBox chkExcel 
      Height          =   255
      Left            =   7680
      TabIndex        =   23
      Top             =   1800
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Archivo Excel"
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
      Value           =   1
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnAplicar 
      Height          =   495
      Left            =   8640
      TabIndex        =   24
      Top             =   7200
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Polizas_CargaLote.frx":1C15
   End
   Begin XtremeSuiteControls.PushButton btnCancelar 
      Height          =   495
      Left            =   9960
      TabIndex        =   25
      Top             =   7200
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2350
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Cancelar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCR_Polizas_CargaLote.frx":23ED
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1560
      TabIndex        =   27
      Top             =   6720
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   556
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtComision 
      Height          =   315
      Left            =   1560
      TabIndex        =   28
      Top             =   7080
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   556
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNeto 
      Height          =   315
      Left            =   1560
      TabIndex        =   26
      Top             =   7440
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   556
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar"
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
      Height          =   372
      Index           =   3
      Left            =   960
      TabIndex        =   21
      Top             =   480
      Width           =   1332
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Height          =   372
      Index           =   0
      Left            =   960
      TabIndex        =   10
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
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
      Index           =   15
      Left            =   3360
      TabIndex        =   20
      Top             =   7080
      Width           =   492
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      Index           =   14
      Left            =   3360
      TabIndex        =   19
      Top             =   7440
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Emitir"
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
      Index           =   13
      Left            =   5640
      TabIndex        =   18
      Top             =   6720
      Width           =   852
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Neto"
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
      Index           =   2
      Left            =   360
      TabIndex        =   14
      Top             =   7440
      Width           =   1092
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comisión"
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
      Index           =   1
      Left            =   360
      TabIndex        =   13
      Top             =   7080
      Width           =   1092
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Aseguradora"
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
      Height          =   372
      Index           =   1
      Left            =   960
      TabIndex        =   12
      Top             =   840
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   360
      TabIndex        =   8
      Top             =   6720
      Width           =   1092
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Primer deducción"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12735
   End
End
Attribute VB_Name = "frmCR_Polizas_CargaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mAseguradoraId As String

Private Sub sbLimpia()
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtComision.Text = 0
    txtNeto.Text = 0

End Sub


Private Sub btnAplicar_Click()
    If vGrid.MaxRows = 0 Then
       MsgBox "No existen casos para procesar...[verifique!]", vbExclamation
       Exit Sub
    End If
    Call sbProcesar
End Sub

Private Sub btnBuscar_Click()
        txtArchivo.Text = ""
        
        With frmContenedor.CD
         If chkExcel.Value = vbChecked Then
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]"
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
        
                
         Else
                .InitDir = "C:\"
                .DialogTitle = "Localice Archivo de Deducciones [Texto]..."
                .Filter = "*.txt"
                .ShowOpen
                
                If .FileName = "" Then
                  MsgBox "Archivo no válido...", vbExclamation
                  Exit Sub
                End If
                
                If UCase(Right(.FileName, 3)) <> "TXT" Then
                  MsgBox "La Extensión del Archivo no es válido...", vbExclamation
                  Exit Sub
                End If
         End If
        
         txtArchivo.Text = .FileName
        
        End With

End Sub

Private Sub btnCancelar_Click()
    txtArchivo.Text = ""
    Call sbLimpia
End Sub

Private Sub btnCargar_Click()
    Call sbCargaArchivo
End Sub

Private Sub btnInfo_Click()
  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
        & " - Columnas: CEDULA, NOMBRE, MONTO, TASA, PLAZO, COMISION" & vbCrLf _
        & " - Nombre de la Hoja: IMPORT" _
    , vbInformation, "Información del Archivo de Carga"
End Sub

Private Sub cboAseguradora_Click()

If vPaso Or cboAseguradora.ListCount = 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select CEDULA_Juridica from CRD_POLIZAS_ASEGURADORAS" _
       & " where cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
    mAseguradoraId = rs!Cedula_Juridica
rs.Close

Exit Sub

vError:


End Sub

Private Sub cboBanco_Click()
If vPaso Or cboBanco.ListCount = 0 Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & mAseguradoraId & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:
End Sub


Private Sub cboCliente_Click()
If vPaso Or cboCliente.ListCount = 0 Then Exit Sub
 Call sbLimpia
End Sub


Private Sub cboPrideduc_Click()
If vPaso Or cboPrideduc.ListCount = 0 Then Exit Sub
 Call sbLimpia
End Sub

Private Sub chkExcel_Click()
 Call sbLimpia
End Sub

Function fxFechaProcesoSiguiente(lngFecha As Long) As Long
Dim strMes As String, strAnio As String, strFecha As String
Dim iMes As Integer, iAnio As Integer
strFecha = Trim(CStr(lngFecha))
     strAnio = Mid(strFecha, 1, 4)
     strMes = Mid(strFecha, 5, 2)
     iAnio = CInt(strAnio)
     iMes = CInt(strMes)
     If CInt(strMes) = 12 Then
         iAnio = iAnio + 1
         strAnio = Trim(str(iAnio))
         strMes = "01"
     Else
       Select Case iMes
       Case 1, 2, 3, 4, 5, 6, 7, 8
         iMes = iMes + 1
         strMes = "0" & Trim(str(iMes))
       Case 9, 10, 11
         iMes = iMes + 1
         strMes = Trim(str(iMes))
       End Select
     End If
     fxFechaProcesoSiguiente = CLng(Trim(strAnio) & Trim(strMes))
End Function

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer
Dim vProceso As Long

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


vPaso = True


cboCliente.Clear
cboConfirma.Clear

strSQL = "select rtrim(codigo) as 'IdX' , rtrim(descripcion) + '  ['  + rtrim(codigo) + ']' as 'ItmX'" _
       & " from catalogo where retencion = 'N' and activo = 1" _
       & " and codigo not in(select codigo_ase from fnd_planes)"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 cboCliente.AddItem rs!itmX & ""
 cboCliente.ItemData(cboCliente.ListCount - 1) = CStr(rs!IdX)
 
 cboConfirma.AddItem rs!itmX & ""
 cboConfirma.ItemData(cboConfirma.ListCount - 1) = CStr(rs!IdX)
 
 
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboCliente.Text = rs!itmX & ""
End If
rs.Close


strSQL = "select cod_Aseguradora as 'IdX',NOMBRE as 'ItmX' from CRD_POLIZAS_ASEGURADORAS where activo = 1"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)

cboTipoDocumento.Clear
cboTipoDocumento.AddItem "CHEQUE"
cboTipoDocumento.AddItem "TRANSFERENCIA"
cboTipoDocumento.Text = "TRANSFERENCIA"


strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

txtArchivo.Text = ""

vGrid.MaxCols = 7
vGrid.MaxRows = 0

vProceso = GLOBALES.glngFechaCR
cboPrideduc.AddItem vProceso

For i = 1 To 6
  vProceso = fxFechaProcesoSiguiente(vProceso)
  cboPrideduc.AddItem vProceso
Next i
cboPrideduc.Text = GLOBALES.glngFechaCR

vPaso = False

Call cboAseguradora_Click

End Sub

Private Sub sbCargaArchivo()
Dim strSQL As String, rs As New ADODB.Recordset, rsExcel As New ADODB.Recordset
Dim strCadena As String, curMonto As Currency, curComision As Currency, iLinea As Long

Dim pCliente As String, pProceso As Long, pComision As Currency, pNeto As Currency
Dim pCedula As String, pNombre As String, pAseguradora As String
Dim pMonto As Currency, pPlazo As Integer, pTasa As Currency, pCuota As Currency


Dim fn, Casos(4) As Long


If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If

If cboCliente.ListCount <= 0 Then Exit Sub
If cboAseguradora.ListCount <= 0 Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

curMonto = 0
curComision = 0
iLinea = 0

pProceso = cboPrideduc.Text
pAseguradora = cboAseguradora.ItemData(cboAseguradora.ListIndex)
pCliente = cboCliente.ItemData(cboCliente.ListIndex)

'La llave es solo codigo y proceso
'& " and cod_aseguradora = '" & pAseguradora & "'"

strSQL = "delete CRD_CREDITOS_CARGADO_H where codigo = '" & pCliente _
       & "' and PROCESO = " & pProceso
       
Call ConectionExecute(strSQL)

strSQL = "" 'Inicializa Bloque



Set rsExcel = Excel_Load(txtArchivo.Text, "Import")
iLinea = 0

Do While Not rsExcel.EOF

    iLinea = iLinea + 1
    
    pCedula = Trim(CStr(rsExcel!Cedula & ""))
    
    
     If pCedula <> "" Then
                
            pNombre = Trim(CStr(rsExcel!Nombre))
            pMonto = CCur(IIf(IsNull(rsExcel!Monto), 0, rsExcel!Monto))
            pComision = CCur(IIf(IsNull(rsExcel!Comision), 0, rsExcel!Comision))
            pNeto = pMonto - pComision
            
            
            
            curMonto = curMonto + pMonto
            curComision = curComision + pComision
            pPlazo = rsExcel!Plazo
            pTasa = rsExcel!Tasa
            pCuota = 0
                
                strSQL = strSQL & Space(10) & "Insert CRD_CREDITOS_CARGADO_H(LINEA,CODIGO,cod_aseguradora,PROCESO,CEDULA,MONTO,NOMBRE,TIPO, PLAZO, TASA, CUOTA, COMISION)" _
                        & " VALUES(" & iLinea & ",'" & pCliente & "','" & pAseguradora & "'," & pProceso & ",'" & pCedula & "'," & pMonto & ",'" & pNombre _
                        & "','D'," & pPlazo & "," & pCuota & "," & pTasa & "," & pComision & ")"
     End If
  
     
     If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
     End If
  
  rsExcel.MoveNext
Loop



'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


'Procesa Revisión de la Carga de Datos
curMonto = 0
curComision = 0

strSQL = "exec spCrd_Creditos_Cargado_Revisado '" & pCliente & "','" & pAseguradora & "'," & pProceso
Call OpenRecordSet(rs, strSQL)

With vGrid
    .MaxRows = 0
    Do While Not rs.EOF
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 1
        .Text = rs!Cedula
        .Col = 2
        .Text = rs!Nombre
        .Col = 3
        .Text = CStr(rs!Monto)
        .Col = 4
        .Text = CStr(rs!Plazo)
        .Col = 5
        .Text = CStr(rs!Tasa)
        .Col = 6
        .Text = CStr(rs!Cuota)
        .Col = 7
        .Text = CStr(rs!Comision)
        
        curMonto = curMonto + rs!Monto
        curComision = curComision + rs!Comision
        
        rs.MoveNext
    Loop
    rs.Close
End With


'Totales
txtMonto.Text = Format(curMonto, "Standard")
txtComision.Text = Format(curComision, "Standard")
txtNeto.Text = Format(curMonto - curComision, "Standard")

Me.MousePointer = vbDefault

MsgBox "Información Cargada Satisfactoriamente", vbInformation

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    vGrid.MaxRows = 0
    txtMonto.Text = 0
    txtComision.Text = 0
    txtNeto.Text = 0
End Sub

Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String, vConcepto As String) As Long                                  'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,tipo_Beneficiario,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & vConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "',5,'" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','Pol','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL, 0)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL, 0)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  rsX.CursorLocation = adUseServer
  Call OpenRecordSet(rsX, strSQL, 0)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Function fxCtaPuente(pCodigo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CtaPuente from Catalogo where codigo  ='" & pCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX.EOF And rsX.BOF Then
     fxCtaPuente = ""
Else
     fxCtaPuente = rsX!CtaPuente
End If

rsX.Close

End Function



Private Sub sbProcesar()
Dim strSQL As String, pCedula As String, i As Long
Dim pClienteId As String, pAseguradora As String, pProceso As Long
Dim pTesoreriaId As Long, vFecha As Date, pConfirma As String

Dim pCuenta As String, pUnidad As String, pConcepto As String, pTipo As String

On Error GoTo vError



pClienteId = cboCliente.ItemData(cboCliente.ListIndex)
pConfirma = cboConfirma.ItemData(cboConfirma.ListIndex)

If pClienteId <> pConfirma Then
   MsgBox "La confirmación de la línea/cliente ha fallado, revise!", vbExclamation
   Exit Sub
End If

pAseguradora = cboAseguradora.ItemData(cboAseguradora.ListIndex)
pProceso = cboPrideduc.Text



pUnidad = "OC"
pConcepto = "CAR"

vFecha = fxFechaServidor

If cboTipoDocumento.Text = "CHEQUE" Then
    pTipo = "CK"
Else
    pTipo = "TE"
End If



Me.MousePointer = vbHourglass


'Procesa Lote
strSQL = "exec spCrd_Creditos_Cargado_Procesa '" & cboCliente.ItemData(cboCliente.ListIndex) _
       & "'," & cboPrideduc.Text & ",'" & cboAseguradora.ItemData(cboAseguradora.ListIndex) _
       & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'TESORERIA


pTesoreriaId = fxMaestroTesoreria(pTipo, cboBanco.ItemData(cboBanco.ListIndex), CCur(txtNeto.Text) _
            , Trim(mAseguradoraId) _
            , cboAseguradora.Text, 0, "Ops:" & vGrid.MaxRows & " Cp:" & cboCliente.ItemData(cboCliente.ListIndex), 0 _
            , ("Docs:" & Trim("")) _
            , cboCuenta.ItemData(cboCuenta.ListIndex), vFecha, pUnidad, pConcepto)
            

pCuenta = fxCtaBanco(cboBanco.ItemData(cboBanco.ListIndex))
Call sbCreaDetalle(pTesoreriaId, pCuenta, CCur(txtNeto.Text), "H", 1, pUnidad)

pCuenta = fxCtaPuente(cboCliente.ItemData(cboCliente.ListIndex))
Call sbCreaDetalle(pTesoreriaId, pCuenta, CCur(txtNeto.Text), "D", 2, pUnidad)
    

Call sbReportes(pTesoreriaId)

txtArchivo.Text = ""
Call sbLimpia

Me.MousePointer = vbDefault

MsgBox "Cargado y Registro de Solicitud en Bancos realizada satisfactoriamente!", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbReportes(pTesoreria As Long)

If pTesoreria = 0 Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    .Formulas(1) = "fxCodigoBarras = '*" & pTesoreria & "*'"
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_BoletaRegistro.rpt")
    .SelectionFormula = "{CHEQUES.NSOLICITUD} = " & pTesoreria
    
    .SubreportToChange = "sbDetalle"

    .StoredProcParam(0) = pTesoreria

    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub



Private Sub sbGuardaBk()
'Dim i As Long, vCadena As String, vTempo As String
'Dim vFile As String, vArchivo As String, vRuta As String, vFecha As Date
'Dim fnFile, vFechaProceso As Long
'
'
'vFecha = fxFechaServidor
'fnFile = FreeFile
'
'vFechaProceso = cboPrideduc.Text
'
'
''Crea Directorios
'
'On Error Resume Next
'
'MkDir SIFGlobal.DirectorioDeResultados
'MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex)
'MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\Cargado"
'MkDir SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\Cargado\" & vFechaProceso
'
'
'vRuta = SIFGlobal.DirectorioDeResultados & "\" & cboCliente.ItemData(cboCliente.ListIndex) & "\Cargado\" & vFechaProceso
'
'
'vArchivo = vFechaProceso & " [Cargado] " & cboCliente.ItemData(cboCliente.ListIndex) & " - " & cboAseguradora.ItemData(cboAseguradora.ListIndex) _
'          & " [" & glogon.Usuario & "].txt"
'
'
'vTempo = vRuta & "\" & vArchivo
'
'vFile = Dir(vTempo, vbArchive)
'
'If vFile = vArchivo Then  'El archivo existe
' Kill vTempo
'End If
'
'
'On Error GoTo vError
'
'Dim strSQL As String
'Dim vIdCliente As String, vInstitucion As Integer
'Dim vCedula As String, vNombre As String
'Dim vMonto As Currency, vInstExiste As String, vMovimiento As String
'
'
'
'vIdCliente = cboCliente.ItemData(cboCliente.ListIndex)
'vInstitucion = cboAseguradora.ItemData(cboAseguradora.ListIndex)
'vFechaProceso = cboPrideduc.Text
'
'strSQL = "delete CRD_CREDITOS_CARGADO_H where codigo = '" & vIdCliente _
'       & "' and PROCESO = '" & vFechaProceso & "' and cod_institucion = " & vInstitucion
'
'Call ConectionExecute(strSQL)
'
'
'
'Open vTempo For Output As #fnFile  ' Create file name.
'
'For i = 1 To vGrid.MaxRows
' vGrid.Row = i
' vGrid.Col = 1
' vCedula = Trim(vGrid.Text)
' vCadena = SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 15)
'
' vGrid.Col = 2
' vNombre = Trim(vGrid.Text)
' vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "D", " ", 50)
'
' vGrid.Col = 3
' vMonto = CCur(vGrid.Text)
' vCadena = vCadena & Format(vGrid.Text, "000000000.00")
'
' vGrid.Col = 4
' vMovimiento = Trim(vGrid.Text)
' vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "I", " ", 10)
'
' vGrid.Col = 5
' vInstExiste = Trim(vGrid.Text)
' vCadena = vCadena & SIFGlobal.fxStringRelleno(vGrid.Text, "I", " ", 10)
'
'
' strSQL = "Insert CRD_CREDITOS_CARGADO_H(LINEA,CODIGO,COD_INSTITUCION,PROCESO,CEDULA,MONTO,NOMBRE,MOVIMIENTO,TIPO, EXISTE_INST)" _
'         & " VALUES(" & i & ",'" & vIdCliente & "'," & vInstitucion & "," & vFechaProceso & ",'" & vCedula & "'," & vMonto & ",'" & vNombre _
'         & "','" & Mid(vMovimiento, 1, 1) & "','I','" & vInstExiste & "')"
'
' If vMovimiento <> "Error" Then
'    Call ConectionExecute(strSQL)
'End If
'
' Print #fnFile, vCadena
'Next i
'
'Close #fnFile
'
'Me.MousePointer = vbDefault
'
'Exit Sub
'
'vError:
'  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Function fxRevisaInst(pCedula As String) As String
Dim Resultado As String


Resultado = "Ok"



fxRevisaInst = Resultado
End Function



