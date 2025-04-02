VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPreaSubRefundiciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Expediente : xx"
   ClientHeight    =   5976
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5976
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCuota 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   5400
      Width           =   1455
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4092
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   12132
      _Version        =   524288
      _ExtentX        =   21400
      _ExtentY        =   7218
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
      SpreadDesigner  =   "frmPreaSubRefundiciones.frx":0000
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.PushButton cmdActualizar_Refundiciones 
      Height          =   612
      Left            =   10560
      TabIndex        =   5
      Top             =   5280
      Width           =   1572
      _Version        =   1245187
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Actualizar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmPreaSubRefundiciones.frx":097E
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Totales ..:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Refundiciones"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frmPreaSubRefundiciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Private clsEntidad As New ProGrX_EstudioCrd.clsEntidad
Public mCambios As Boolean



Private Sub cmdActualizar_Refundiciones_Click()

    If (MsgBox("Está seguro que desea actualizar las refundiciones, deberá volver a selecionar los créditos que desea refundir", vbQuestion + vbYesNo)) = vbYes Then
            'Actualizar Refundiciones
            
            mCambios = True
            
            If fxValidaEstado(gPreAnalisis.Expediente) = True Then
                glogon.strSQL = "exec spCRDPreaRefundiciones '" & gPreAnalisis.Expediente & "','A'"
                If Not clsEntidad.fxEjecutaSQL(glogon.strSQL) Then
                    MsgBox "Ocurrió un error al inicializar fianzas.", vbInformation, gMsgTitulo
                End If
            End If
    
            Call CargarGrid
    End If

End Sub

Private Sub CargarGrid()

On Error GoTo vError

Dim strSQL As String

vPaso = True

strSQL = "select R.id_solicitud,X.descripcion,R.saldo,R.cuota,mora_principal,Mora_intereses" _
    & ",Aplica,Apl_Mora, C.garantia, G.descripcion, ((C.MontoApr-R.Saldo) /C.MontoApr) * 100 as 'Porcentaje' ,C.MontoApr" _
    & " from CRD_PREA_REFUNDICIONES R inner join Reg_Creditos C on R.id_solicitud = C.id_solicitud" _
    & " inner join Catalogo X on C.codigo = X.codigo " _
    & " inner join crd_garantia_tipos G on G.garantia = C.garantia " _
    & " where R.cod_PreAnalisis = '" & gPreAnalisis.Expediente & "'"
Call sbCargaGridLocal(vGrid, 11, strSQL)

'vGrid.MaxRows = vGrid.MaxRows - 1

Call sbCalculaTotales

vPaso = False


Exit Sub

vError:
    MsgBox "Ocurrió un error al cargar grid. " & "-" & Err.Description, vbExclamation

End Sub

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    
    Select Case i
      Case 9
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 700
        vGrid.CellNote = "Garantía: " & Chr(13) & rs!Descripcion
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
      Case 10 'Porcentaje
        vGrid.Text = Format(rs!Porcentaje, "##0.00")
      Case 11 'Monto
        vGrid.Text = Format(rs!montoapr, "Standard")
      Case Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
      
    End Select
    
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext
Loop
rs.Close

Call sbCalculaTotales

End Sub

Private Sub Form_Load()
Dim strSQL As String

    Me.Caption = "Expediente : " & gPreAnalisis.Expediente
    
    mCambios = False
    Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture
    
    Call CargarGrid

    vGrid.MaxRows = vGrid.MaxRows - 1
    
    ' Activa el botón de actualizar si el estado es R
    cmdActualizar_Refundiciones.Visible = False
    If fxValidaEstado(gPreAnalisis.Expediente) = True Then
        cmdActualizar_Refundiciones.Visible = True
    Else
        cmdActualizar_Refundiciones.Visible = False
    End If

End Sub


Private Sub sbCalculaTotales()
Dim i As Integer, curCuota As Currency, curMonto As Currency


vPaso = True

curCuota = 0
curMonto = 0

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  'Revisa que si aplica solo morosidad desmarca automáticamente la refundicion total
'  vGrid.Col = 8
'  If vGrid.Value = vbChecked Then
'     vGrid.Col = 7
'     vGrid.Value = vbUnchecked
'  End If
  
  
  vGrid.Col = 7
  If vGrid.Value = vbChecked Then
    vGrid.Col = 4 'Cuota
    curCuota = curCuota + IIf((vGrid.Text = ""), 0, vGrid.Text)
    vGrid.Col = 3 'Saldo
    curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
    vGrid.Col = 6 'Saldo + Itereses Moratorios
    curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
  End If
  
  vGrid.Col = 8
  If vGrid.Value = vbChecked Then
'En Cancelación de Morosidad no Reconoce la cuota para cálculo de liquidez
'    vGrid.Col = 4 'Cuota
'    curCuota = curCuota + IIf((vGrid.Text = ""), 0, vGrid.Text)
    vGrid.Col = 5 'Principal Atrasado
    curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
    vGrid.Col = 6 'Principal Atrasado + Itereses Moratorios
    curMonto = curMonto + IIf((vGrid.Text = ""), 0, vGrid.Text)
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

vGrid.Row = Row
vGrid.Col = Col

mCambios = True

Select Case Col
 Case 7 'Todo
   If vGrid.Value = 1 Then
      vGrid.Col = 8
      vGrid.Value = 0
   End If
   
 Case 8 'Mora
   If vGrid.Value = 1 Then
      vGrid.Col = 7
      vGrid.Value = 0
   End If
   
End Select


If Col = 7 Or Col = 8 Then
    
    If Not ValidaEstadoPreanalisis(gPreAnalisis.ESTADO) Then
        Exit Sub
    End If

   vGrid.Row = Row
   vGrid.Col = 7
   strSQL = "update CRD_PREA_REFUNDICIONES set Aplica = " & vGrid.Value
   vGrid.Col = 8
   strSQL = strSQL & ", Apl_Mora = " & vGrid.Value & " where cod_PreAnalisis = '" _
          & gPreAnalisis.Expediente & "' and id_solicitud = "
   vGrid.Col = 1
   strSQL = strSQL & vGrid.Text
   Call ConectionExecute(strSQL)
   
   Call sbCalculaTotales
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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


