VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmPreaSubFianzas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3972
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   11652
      _Version        =   524288
      _ExtentX        =   20553
      _ExtentY        =   7006
      _StockProps     =   64
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmPreaSubFianzas.frx":0000
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdActualizar_Fianzas 
      Height          =   612
      Left            =   9840
      TabIndex        =   2
      Top             =   5760
      Width           =   1572
      _Version        =   1310722
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Actualizar"
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
      Appearance      =   16
      Picture         =   "frmPreaSubFianzas.frx":0958
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
      _Version        =   1310722
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCuota 
      Height          =   315
      Left            =   3120
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
      _Version        =   1310722
      _ExtentX        =   2566
      _ExtentY        =   556
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
      Text            =   "0.00"
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Totales ..:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fianzas"
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
      Height          =   615
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaSubFianzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Public mCambios As Boolean

Private Sub cmdActualizar_Fianzas_Click()

    If (MsgBox("Está seguro que desea actualizar las fianzas, deberá volver a selecionar los créditos que desea aplicar", vbQuestion + vbYesNo)) = vbYes Then
            
            mCambios = True
    
            'Actualizar Fianzas
            glogon.strSQL = "spCRDPreaFianzas " & fxFormatearValor(gPreAnalisis.Expediente, Caracter) & "," & "'I'"
        
            If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                MsgBox "Ocurrió un error al inicializar fianzas.", vbInformation, gMsgTitulo
            End If
    
            Call CargarGrid
            vGrid.MaxRows = vGrid.MaxRows - 1
    End If
End Sub


Private Sub Form_Load()


Me.Caption = "Expediente : " & gPreAnalisis.Expediente

Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture
vGrid.AppearanceStyle = AppearanceStyleVisualStyles
    
    Call CargarGrid
    
    mCambios = False

    ' Activa el botón de actualizar si el estado es R
    cmdActualizar_Fianzas.Visible = False
    If fxValidaEstado(gPreAnalisis.Expediente) = True Then
        cmdActualizar_Fianzas.Visible = True
    Else
        cmdActualizar_Fianzas.Visible = False
    End If

vGrid.MaxRows = vGrid.MaxRows - 1

End Sub


Private Sub CargarGrid()

On Error GoTo vError

    Dim strSQL As String

    vPaso = True
    strSQL = "select F.id_solicitud,F.saldo,F.cuota,F.nfiadores,F.Mora_Cuotas,F.Mora_Monto,F.aplica,F.Cancela_Mora" _
           & ", R.MontoApr, ((R.MontoApr-F.Saldo) /R.MontoApr) * 100  as Porcentaje" _
           & " from CRD_PREA_DETALLE_FIANZAS F inner join reg_creditos R on F.id_Solicitud = R.id_solicitud" _
           & " where F.cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
    vGrid.MaxRows = 0
    Call sbCargaGrid(vGrid, 10, strSQL)
    
    
    Call sbCalculaTotales
    vPaso = False
    
    Exit Sub
    
vError:
    MsgBox "Ocurrió un error al cargar grid. " & "-" & Err.Description, vbExclamation


End Sub

Public Function fxValidaEstado(mExpediente As String) As Boolean
On Error GoTo vError
    
    '' Esta función verifica el estado del preanalisis
    
    Dim rs As New ADODB.Recordset, strSQL As String
    
        strSQL = "select ESTADO from CRD_PREA_PREANALISIS where COD_PREANALISIS = '" & Trim(mExpediente) & "'"
        
        Call OpenRecordSet(rs, strSQL)
        
        If Not rs.EOF Then
            If rs.Fields(0) = "R" Then
                fxValidaEstado = True
            Else
                fxValidaEstado = False
            End If
        Else
            fxValidaEstado = False
        End If
        
        rs.Close
        
        Exit Function
vError:
    MsgBox "Ocurrió un error al validar el estado del expediente. " & "-" & Err.Description, vbExclamation

End Function



Private Sub sbCalculaTotales()
Dim i As Integer, curCuota As Currency, curMonto As Currency
Dim vNumFia As Integer

curCuota = 0
curMonto = 0


vPaso = True

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  'Si se marca solo cancelación de mora (desmarca automáticamente el Aplica)
  
    vGrid.Col = 4 'Fiadores
    vNumFia = IIf((vGrid.Text = ""), 0, vGrid.Text)
    
    If vNumFia = 0 Then vNumFia = 1
  
  
  vGrid.Col = 7
  If vGrid.Value = vbChecked Then
    vGrid.Col = 2 'Saldo
    curMonto = curMonto + (IIf((vGrid.Text = ""), 0, vGrid.Text) / vNumFia)
    vGrid.Col = 3 'Cuota
    curCuota = curCuota + (IIf((vGrid.Text = ""), 0, vGrid.Text) / vNumFia)
  End If


  vGrid.Col = 8
  If vGrid.Value = vbChecked Then
    vGrid.Col = 6 'Monto en Mora
    curMonto = curMonto + (IIf((vGrid.Text = ""), 0, vGrid.Text) / vNumFia)
  End If


Next i
vPaso = False

txtCuota.Text = Format(curCuota, "Standard")
txtMonto.Text = Format(curMonto, "Standard")

End Sub

Private Sub Form_Unload(Cancel As Integer)
  GLOBALES.gTag = txtCuota.Text
  GLOBALES.gTag2 = txtMonto.Text
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError


If Col = 7 Or Col = 8 Then
 
    Call sbCalculaTotales
  
    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
        Exit Sub
    End If
    
    mCambios = True

   vGrid.Row = Row
   vGrid.Col = 7
   strSQL = "update CRD_PREA_DETALLE_FIANZAS set Aplica = " & vGrid.Value
   vGrid.Col = 8
   strSQL = strSQL & ", Cancela_Mora = " & vGrid.Value & " where cod_PreAnalisis = '" _
          & gPreAnalisis.Expediente & "' and id_solicitud = "
   vGrid.Col = 1
   strSQL = strSQL & vGrid.Text
   
   Call ConectionExecute(strSQL)

End If


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
