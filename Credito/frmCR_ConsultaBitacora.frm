VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_ConsultaBitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Movimientos por Persona"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   10800
      Top             =   120
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4335
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   11415
      _Version        =   524288
      _ExtentX        =   20135
      _ExtentY        =   7646
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      SpreadDesigner  =   "frmCR_ConsultaBitacora.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
      TabEnhancedShape=   0
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   2520
      TabIndex        =   3
      Top             =   1080
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1200
      TabIndex        =   4
      Top             =   1080
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
      _ExtentY        =   550
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
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   656
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
      Picture         =   "frmCR_ConsultaBitacora.frx":1B78
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmCR_ConsultaBitacora.frx":2278
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas...:"
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   972
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "frmCR_ConsultaBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCedula As String, mFecha As Date


Private Sub btnConsulta_Click()

On Error GoTo vError

Call vGrid_SheetChanged(1, vGrid.ActiveSheet)

vError:
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders


Select Case vGrid.ActiveSheet
  Case 1 'Registro
        vHeaders.Columnas = vGrid.MaxCols
        vHeaders.Headers(1) = "No. Transacción"
        vHeaders.Headers(2) = "Tipo Transac."
        vHeaders.Headers(3) = "No. Documento"
        vHeaders.Headers(4) = "Fecha"
        vHeaders.Headers(5) = "Usuario"
        vHeaders.Headers(6) = "Monto"
        vHeaders.Headers(7) = "Concepto"
        vHeaders.Headers(8) = "Detalle"
        vHeaders.Headers(9) = "Referencia"
        vHeaders.Headers(10) = "Sistema"

  Case 2 'Creditos
        vHeaders.Columnas = vGrid.MaxCols
        vHeaders.Headers(1) = "No. Operacion"
        vHeaders.Headers(2) = "Linea"
        vHeaders.Headers(3) = "Descripción"
        vHeaders.Headers(4) = "Fecha Proceso"
        vHeaders.Headers(5) = "Concepto"
        vHeaders.Headers(6) = "Fecha"
        vHeaders.Headers(7) = "Usuario"
        vHeaders.Headers(8) = "Interés Corriente"
        vHeaders.Headers(9) = "Interés Moratorio"
        vHeaders.Headers(10) = "Cargos"
        vHeaders.Headers(11) = "Pólizas"
        vHeaders.Headers(12) = "Principal"
        vHeaders.Headers(13) = "Total Mov."
        vHeaders.Headers(14) = "Tipo Documento"
        vHeaders.Headers(15) = "Num. Comprobante"
        vHeaders.Headers(16) = "Caja"
        vHeaders.Headers(17) = "Garantía"

  Case 3 'Fondos
        vHeaders.Columnas = vGrid.MaxCols
        vHeaders.Headers(1) = "Plan"
        vHeaders.Headers(2) = "Contrato"
        vHeaders.Headers(3) = "Descripción"
        vHeaders.Headers(4) = "Monto"
        vHeaders.Headers(5) = "Fecha"
        vHeaders.Headers(6) = "Usuario"
        vHeaders.Headers(7) = "Concepto"
        vHeaders.Headers(8) = "Tipo Documento"
        vHeaders.Headers(9) = "Num. Comprobante"
        vHeaders.Headers(10) = "Caja"
  
  Case 4 'Patrimonio
        vHeaders.Columnas = vGrid.MaxCols
        vHeaders.Headers(1) = "Rubro/Plan"
        vHeaders.Headers(2) = "Monto"
        vHeaders.Headers(3) = "Fecha"
        vHeaders.Headers(4) = "Usuario"
        vHeaders.Headers(5) = "Concepto"
        vHeaders.Headers(6) = "Tipo Documento"
        vHeaders.Headers(7) = "Num. Comprobante"
        vHeaders.Headers(8) = "Caja"
     
  Case 5 'Bancos
        vHeaders.Columnas = vGrid.MaxCols
        vHeaders.Headers(1) = "Banco"
        vHeaders.Headers(2) = "Cuenta"
        vHeaders.Headers(3) = "Tipo Transac."
        vHeaders.Headers(4) = "Tesoreria Id"
        vHeaders.Headers(5) = "Documento"
        vHeaders.Headers(6) = "Lote"
        vHeaders.Headers(7) = "Monto"
        vHeaders.Headers(8) = "Fecha"
        vHeaders.Headers(9) = "Usuario"
        vHeaders.Headers(10) = "Divisa"
        vHeaders.Headers(11) = "Ref 01"
        vHeaders.Headers(12) = "Ref 02"
        vHeaders.Headers(13) = "Ref 03"
        vHeaders.Headers(14) = "Concepto"
        vHeaders.Headers(15) = "Detalle"

End Select

Call sbSIFGridExportar(vGrid, vHeaders, "ProGRX_Persona_MovLog_" & mCedula & "_" & vGrid.SheetName _
        & "_" & Format(dtpInicio.Value, "yyyy-mm-dd") & " - " & Format(dtpCorte.Value, "yyyy-mm-dd"))

End Sub

Private Sub Form_Activate()
vModulo = 3

End Sub

Private Sub Form_Load()
vModulo = 3

imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mCedula = GLOBALES.gTag

Call sbInicializa

End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strSQL = "select cedula,nombre,dbo.MyGetdate() as Fecha from socios where cedula = '" _
       & mCedula & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  lblCliente.Caption = Trim(rs!Cedula) & " - " & Trim(rs!Nombre)
  mFecha = rs!fecha
End If
rs.Close


dtpInicio.Value = DateAdd("d", -7, mFecha)
dtpCorte.Value = mFecha

vGrid.Sheet = 1
vGrid.MaxRows = 0

MousePointer = vbDefault

End Sub



Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 650
vGrid.Height = Me.Height - (vGrid.top + 800)
imgBanner.Width = Me.Width

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

vGrid.Sheet = 1
Call vGrid_SheetChanged(1, 1)

End Sub


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

With vGrid
    .Sheet = NewSheet
    
   Select Case NewSheet
        Case 1 'Registro de Historial
   
            strSQL = "exec spSIFPersonaMovimientos '" & mCedula _
                    & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
                    & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
            Call OpenRecordSet(rs, strSQL)
             .MaxRows = 0
             Do While Not rs.EOF
                 .MaxRows = .MaxRows + 1
                 .Row = .MaxRows
                 For i = 1 To .MaxCols
                   .Col = i
                   Select Case i
                     Case 1 'NTransaccion
                        .Text = CStr(rs!NTransaccion)
                     Case 2 'Tipo
                        .Text = CStr(rs!TDOCUMENTO)
                     Case 3 'Numero de Doc.
                        .Text = CStr(rs!nDocumento & "")
                     Case 4 'Fecha
                        .Text = rs!fecha & ""
                     Case 5 'Usuario
                        .Text = rs!Usuario & ""
                     Case 6 'Monto
                        .Text = Format(rs!Monto, "Standard")
                     Case 7 'Concepto
                        .Text = rs!CONCEPTO & ""
                     Case 8 'Detalle
                        .Text = rs!Detalle & ""
                     Case 9 'Referencia
                        .Text = rs!Referencia & ""
                     Case 10 'Cod.App
                        .Text = rs!CodApp & ""
                   End Select
                 Next i
                
                .RowHeight(.Row) = .MaxTextRowHeight(.Row)
                rs.MoveNext
             
             Loop
            rs.Close


         
        Case 2 'Movimientos a Creditos
            strSQL = "exec spCrdPersonaMovimientos '" & mCedula _
                    & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
                    & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
            Call OpenRecordSet(rs, strSQL)
             .MaxRows = 0
             Do While Not rs.EOF
                 .MaxRows = .MaxRows + 1
                 .Row = .MaxRows
                 For i = 1 To .MaxCols
                   .Col = i
                   Select Case i
                     Case 1 'Operacion
                        .Text = CStr(rs!ID_SOLICITUD)
                     Case 2 'Linea
                        .Text = CStr(rs!Codigo)
                     Case 3 'Descripcion
                        .Text = CStr(rs!LineaX)
                     Case 4 'Proceso
                        .Text = Format(rs!Proceso, "####-##")
                     Case 5 'Concepto
                        .Text = rs!CONCEPTO
                     Case 6 'Fecha
                        .Text = Format(rs!fecha, "yyyy-mm-dd hh:mm:ss")
                     Case 7 'Usuario
                        .Text = rs!Usuario & ""
                     
                     Case 8 'Int.Cor.
                        .Text = Format(rs!IntCor, "Standard")
                     Case 9 'Int.Mor.
                        .Text = Format(rs!IntMor, "Standard")
                     Case 10 'Cargos
                        .Text = Format(rs!Cargo, "Standard")
                     Case 11 'Poliza (Prevista)
                        .Text = Format(rs!Poliza, "Standard")
                     Case 12 'Amortiza
                        .Text = Format(rs!Principal, "Standard")
                     Case 13 'Total
                        .Text = Format(rs!IntCor + rs!IntMor + rs!Cargo + rs!Poliza + rs!Principal, "Standard")
                     Case 14 'Tipo Doc
                        .Text = rs!Tipo & ""
                     Case 15 'NDocumento
                        .Text = rs!nCon & ""
                     Case 16 'Caja
                        .Text = rs!COD_CAJA & ""
                     Case 17 'Garantia
                        .Text = CStr(rs!GarantiaDesc)
                   End Select
                 Next i
                 rs.MoveNext
             Loop
            rs.Close
    
        
        
        Case 3 'Movimientos a Fondos
           
            strSQL = "select C.COD_OPERADORA, C.COD_PLAN, C.COD_CONTRATO, C.CEDULA, S.NOMBRE" _
                   & "    , P.DESCRIPCION as 'PLAN', M.MONTO, M.FECHA, M.USUARIO" _
                   & "    , isnull(Cc.DESCRIPCION,'') as 'ConceptoDesc'" _
                   & "    , isnull(D.DESCRIPCION,'') as 'TipoDocDesc'" _
                   & "    , M.NCON , M.COD_CAJA" _
                   & " from FND_CONTRATOS C" _
                   & "    inner join SOCIOS S on C.CEDULA = S.CEDULA" _
                   & "    inner join FND_PLANES P on C.COD_OPERADORA = P.COD_OPERADORA and  C.COD_PLAN = P.COD_PLAN" _
                   & "    inner join FND_CONTRATOS_DETALLE M on C.COD_OPERADORA = M.COD_OPERADORA" _
                   & "        and  C.COD_PLAN = M.COD_PLAN and C.COD_CONTRATO = M.COD_CONTRATO" _
                   & "    left join SIF_CONCEPTOS Cc on M.COD_CONCEPTO = Cc.COD_CONCEPTO" _
                   & "    left join SIF_DOCUMENTOS D on M.TCON = D.TIPO_DOCUMENTO " _
                   & " Where C.CEDULA = '" & mCedula & "'" _
                   & "   and M.FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
                   & " order by M.FECHA"

            Call OpenRecordSet(rs, strSQL)
             .MaxRows = 0
        
             Do While Not rs.EOF
                 .MaxRows = .MaxRows + 1
                 .Row = .MaxRows
                 For i = 1 To .MaxCols
                   .Col = i
                   Select Case i
                     Case 1 'Id Plan
                        .Text = rs!COD_PLAN
                     Case 2 'Contrato
                        .Text = rs!COD_CONTRATO
                     Case 3 'Descripcion
                        .Text = rs!Plan
                     Case 4 'Monto
                        .Text = Format(rs!Monto, "Standard")
                     Case 5 'Fecha
                        .Text = Format(rs!fecha, "yyyy-mm-dd hh:mm:ss")
                     Case 6 'Usuario
                        .Text = rs!Usuario
                     Case 7 'Concepto
                        .Text = rs!ConceptoDesc
                     Case 8 'TipoDoc
                        .Text = rs!TipoDocDesc
                     Case 9 'Num.Doc
                        .Text = rs!nCon & ""
                     Case 10  'Caja
                        .Text = rs!COD_CAJA & ""
                   End Select
                 Next i
                 rs.MoveNext
             Loop
            rs.Close
    
        
        
        Case 4 'Movimientos a Patrimonio
           
            strSQL = "select M.DESCRIPCION as 'PLAN', M.MONTO, M.FECHA, M.USUARIO" _
                   & "     , isnull(C.DESCRIPCION,'') as 'ConceptoDesc'" _
                   & "     , isnull(D.DESCRIPCION,'') as 'TipoDocDesc'" _
                   & "     , M.NCON , M.COD_CAJA" _
                   & "  from vPAT_Movimientos M" _
                   & "     left join SIF_CONCEPTOS C on M.COD_CONCEPTO = C.COD_CONCEPTO" _
                   & "     left join SIF_DOCUMENTOS D on M.TCON = D.TIPO_DOCUMENTO" _
                   & "  Where CEDULA = '" & mCedula & "'" _
                   & "   and FECHA between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'" _
                   & " order by FECHA"

            Call OpenRecordSet(rs, strSQL)
             .MaxRows = 0
        
             Do While Not rs.EOF
                 .MaxRows = .MaxRows + 1
                 .Row = .MaxRows
                 For i = 1 To .MaxCols
                   .Col = i
                   Select Case i
                     Case 1 'Rubro
                        .Text = rs!Plan
                     Case 2 'Monto
                        .Text = Format(rs!Monto, "Standard")
                     Case 3 'Fecha
                        .Text = Format(rs!fecha, "yyyy-mm-dd hh:mm:ss")
                     Case 4 'Usuario
                        .Text = rs!Usuario
                     Case 5 'Concepto
                        .Text = rs!ConceptoDesc
                     Case 6 'TipoDoc
                        .Text = rs!TipoDocDesc
                     Case 7 'Num.Doc
                        .Text = rs!nCon & ""
                     Case 8  'Caja
                        .Text = rs!COD_CAJA & ""
                   End Select
                 Next i
                 rs.MoveNext
             Loop
            rs.Close
    
    
    
        Case 5 'Movimientos en Bancos
           
            strSQL = " select Gb.DESCRIPCION as 'BancoDesc', B.CTA as 'CtaId', B.DESCRIPCION as 'CtaDesc'" _
                   & "     , T.MONTO , T.FECHA_EMISION as 'FECHA', T.USER_SOLICITA as 'USUARIO', T.COD_DIVISA" _
                   & "     , T.REF_01, T.REF_02, T.REF_03" _
                   & "     , Tc.DESCRIPCION as 'ConceptoDesc', Td.Descripcion as 'TipoDoc'" _
                   & "     , T.DETALLE1 + ' ' + T.DETALLE2 + ' ' + isnull(T.DETALLE3,'')" _
                   & "       + isnull(T.DETALLE4,'') + ' ' +  isnull(T.DETALLE5,'') as 'Detalle'" _
                   & "     , T.NSOLICITUD, T.NDOCUMENTO, T.DOCUMENTO_BASE" _
                   & "  from TES_TRANSACCIONES T" _
                   & "     inner join TES_BANCOS B on T.ID_BANCO = B.ID_BANCO" _
                   & "     inner join TES_BANCOS_GRUPOS Gb on B.COD_GRUPO = Gb.COD_GRUPO" _
                   & "     inner join TES_TIPOS_DOC Td on T.TIPO = Td.TIPO" _
                   & "     inner join TES_CONCEPTOS Tc on T.COD_CONCEPTO = Tc.COD_CONCEPTO" _
                   & " Where T.CODIGO = '" & mCedula & " '" _
                   & "   and T.ESTADO in('E','T', 'I')" _
                   & "   and T.FECHA_EMISION between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"

            Call OpenRecordSet(rs, strSQL)
             .MaxRows = 0
        
             Do While Not rs.EOF
                 .MaxRows = .MaxRows + 1
                 .Row = .MaxRows
                 For i = 1 To .MaxCols
                   .Col = i
                   Select Case i
                     Case 1 'Banco
                        .Text = rs!BancoDesc
                     Case 2 'Cta Desc
                        .Text = rs!CtaDesc
                        
                     Case 3 'Tipo Doc
                        .Text = rs!TipoDoc
                        
                     Case 4 'Tesoreria Id
                        .Text = rs!NSolicitud
                     Case 5 'No.Documento
                        .Text = rs!nDocumento & ""
                     Case 6 'No. DocBase
                        .Text = rs!DOCUMENTO_BASE & ""
                        
                        
                        
                     Case 7 'Monto
                        .Text = Format(rs!Monto, "Standard")
                     Case 8 'Fecha
                        .Text = Format(rs!fecha, "yyyy-mm-dd hh:mm:ss")
                     Case 9 'Usuario
                        .Text = rs!Usuario & ""
                     Case 10 'Divisa
                        .Text = rs!cod_Divisa & ""
                     
                     Case 11 'Ref 1
                        .Text = rs!ref_01 & ""
                     Case 12 'Ref 2
                        .Text = rs!Ref_02 & ""
                     Case 13 'Ref 3
                        .Text = rs!Ref_03 & ""
                     Case 14 'Concepto
                        .Text = rs!ConceptoDesc & ""
                     Case 15 'Detalle
                        .Text = rs!Detalle & ""
                   End Select
                 Next i
                 rs.MoveNext
             Loop
            rs.Close
    

    
    End Select
End With

MousePointer = vbDefault


Exit Sub

vError:
 MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
