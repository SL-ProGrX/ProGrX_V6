VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_TransferenciaRepControl 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Control de Transferencias"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   Icon            =   "frmTES_TransferenciaRepControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   7965
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1572
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   7332
      _Version        =   1441793
      _ExtentX        =   12933
      _ExtentY        =   2773
      _StockProps     =   79
      Caption         =   "Consulta de Transferencia"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboDoc 
         Height          =   312
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   5172
         _Version        =   1441793
         _ExtentX        =   9128
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboFormato 
         Height          =   312
         Left            =   1800
         TabIndex        =   10
         Top             =   1080
         Width           =   3372
         _Version        =   1441793
         _ExtentX        =   5953
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboPlan 
         Height          =   312
         Left            =   5160
         TabIndex        =   13
         Top             =   1080
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formato"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   600
         TabIndex        =   11
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   852
      End
   End
   Begin XtremeSuiteControls.PushButton btnOpciones 
      Height          =   648
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   3480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   1143
      _StockProps     =   79
      Caption         =   "Carta al Banco"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_TransferenciaRepControl.frx":6852
   End
   Begin XtremeSuiteControls.PushButton btnOpciones 
      Height          =   648
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   3480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   1143
      _StockProps     =   79
      Caption         =   "Detalle de la Transferencia"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_TransferenciaRepControl.frx":700E
   End
   Begin XtremeSuiteControls.PushButton btnOpciones 
      Height          =   648
      Index           =   2
      Left            =   5640
      TabIndex        =   6
      Top             =   3480
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   1143
      _StockProps     =   79
      Caption         =   "&Genera Archivo"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmTES_TransferenciaRepControl.frx":77CA
   End
   Begin XtremeSuiteControls.FlatEdit txtNTran 
      Height          =   492
      Left            =   5160
      TabIndex        =   12
      Top             =   2880
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   868
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Informe de Transferencias"
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
      Height          =   492
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   240
      Width           =   6252
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Transferencia:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   1
      Left            =   2760
      TabIndex        =   0
      Top             =   3000
      Width           =   2292
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmTES_TransferenciaRepControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub sbNTrasnferencia()
Dim strSQL As String, rs As New ADODB.Recordset

txtNTran = fxTesTipoDocConsec(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "/", cboPlan.ItemData(cboPlan.ListIndex))

End Sub

Private Sub btnOpciones_Click(Index As Integer)

If Not IsNumeric(txtNTran.Text) Then Exit Sub

Select Case Index
  Case 0 'Carta
         Call sbTesReporteTransferencia(cbo.ItemData(cbo.ListIndex), txtNTran, "C", "TE", cboPlan.ItemData(cboPlan.ListIndex))
  Case 1 'Informe
         Call sbTesReporteTransferencia(cbo.ItemData(cbo.ListIndex), txtNTran, "D", "TE", cboPlan.ItemData(cboPlan.ListIndex))
  Case 2 'Archivo
     Call sbArchivo
End Select

End Sub

Private Sub cbo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub

vPaso = True

strSQL = "exec spTes_Formatos_Bancos " & cbo.ItemData(cbo.ListIndex)
Call sbCbo_Llena_New(cboFormato, strSQL, False, True)


strSQL = "select rtrim(T.tipo) as 'IdX', rtrim(T.descripcion) as ItmX" _
       & " from tes_banco_docs D inner join tes_tipos_doc T on D.tipo = T.tipo" _
       & " where D.comprobante = '04' and D.id_Banco = " & cbo.ItemData(cbo.ListIndex)
Call sbCbo_Llena_New(cboDoc, strSQL, False, True)

strSQL = "select Bp.COD_PLAN as 'IdX', Bp.COD_PLAN as 'ItmX'" _
       & " from TES_BANCOS B inner join TES_BANCO_PLANES_TE Bp on B.ID_BANCO = Bp.ID_BANCO" _
       & " Where B.ID_BANCO = " & cbo.ItemData(cbo.ListIndex) & " And B.UTILIZA_PLAN = 1" _
       & " order by Bp.COD_PLAN  asc"
Call sbCbo_Llena_New(cboPlan, strSQL, False, True)
If cboPlan.ListCount = 0 Then
   cboPlan.AddItem "Sin Plan"
   cboPlan.ItemData(cboPlan.ListCount - 1) = "-sp-"
   cboPlan.Text = "Sin Plan"
End If

Call cboDoc_Click

vPaso = False

Call cboDoc_Click

End Sub

Private Sub cboDoc_Click()
If vPaso Then Exit Sub
Call sbNTrasnferencia
End Sub


Private Sub cboPlan_Click()
If vPaso Then Exit Sub
Call sbNTrasnferencia
End Sub

Private Sub Form_Load()
 vModulo = 9

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vPaso = True
 Call sbTesBancoCargaCboAccesoGeneral(cbo)
vPaso = False

Call cbo_Click


End Sub



Private Sub sbTeFormatoEstandar(pFormato As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Generacion con Formatos Estandares de Transferencias Bancarias
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strSQL = "select  REPLACE(cedula_juridica,'-','') as 'Cedula_Juridica',NOMBRE" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL, 0)
    vNumNegocio = Trim(rs!cedula_juridica)
    vCedulaReg = Trim(rs!cedula_juridica)
    vRazon = "TRANSFERENCIAS " & rs!Nombre
rs.Close


Dim vExtension As String, vProcedimiento As String

strSQL = "select Procedimiento,Extension from vTes_Formatos where cod_formato = '" & pFormato & "'"
Call OpenRecordSet(rs, strSQL, 0)
    vExtension = Trim(rs!Extension)
    vProcedimiento = Trim(rs!Procedimiento)
rs.Close


    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     
     gTesGlobal.BancoPlan = cboPlan.ItemData(cboPlan.ListIndex)
     
     gTesGlobal.BancoConsec = txtNTran.Text
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     i = 1
    strSQL = "select documento_base,count(*) From Tes_Transacciones" _
         & " where id_banco = " & gTesGlobal.BancoID & " and fecha_emision = '" _
         & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     i = i + 1
     rs.MoveNext
    Loop
    rs.Close
    vConArchivo = Format(i, "000")
     
     
     strSQL = "select dbo.fxTesCantidadTEDiarias('" & Format(vFecha, "yyyy/mm/dd") & "' ," & gTesGlobal.BancoID & ") as 'Cantidad'"
     Call OpenRecordSet(rs, strSQL)
         iLineInicio = rs!Cantidad
     rs.Close
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & "." & vExtension
     
     Open strArchivo For Output As #1

'     'REGISTRO DE CONTROL
'     i = 1
'
'     strCadena = "000"                                                              'Estado 3
'     strCadena = strCadena & SIFGlobal.fxStringRelleno(vCedulaReg, "I", "0", 12)    'Cedula Juridica 12
'     strCadena = strCadena & vConArchivo                                            'Consecutivo Archivo 3
'     strCadena = strCadena & Format(vFecha, "ddmmyyyy")                             'Fecha Aplicacion 8
'     strCadena = strCadena & "000000000000"                                         'Cedula de Registro 12
'     strCadena = strCadena & "000000000000"                                         '12 Filler con 0
'     strCadena = strCadena & "000000"                                               '6 Hora Estado Se rellena con ceros
'     strCadena = strCadena & SIFGlobal.fxStringRelleno("", "D", "0", 138)           '138 Filler con 0
'
'     Print #1, strCadena
     
     'LINEA CONTROL
     strSQL = "exec " & vProcedimiento & "_Archivo 1," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
            & "'," & gTesGlobal.BancoConsec & "," & 100000 & ",'" & gTesGlobal.BancoPlan & "'"
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        If rs!Linea1 <> "" Then
            Print #1, rs!Linea1 'Linea Control
        End If
        rs.MoveNext
     Loop
     rs.Close
 
  
     'DEBITOS
     strSQL = "exec " & vProcedimiento & "_Archivo 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
            & "'," & gTesGlobal.BancoConsec & "," & 100000 & ",'" & gTesGlobal.BancoPlan & "'"
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        If rs!Linea2 <> "" Then
            Print #1, rs!Linea2 'Debitos
        End If
        rs.MoveNext
     Loop
     rs.Close
   
     'CREDITOS
     strSQL = "exec " & vProcedimiento & "_Archivo 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
            & "'," & gTesGlobal.BancoConsec & "," & 100000 & ",'" & gTesGlobal.BancoPlan & "'"
     Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        If rs!Linea3 <> "" Then
            Print #1, rs!Linea3 'Creditos
        End If
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
    
     MsgBox strArchivo, vbInformation, "Archivo de Transferencia Generado..."
     
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbArchivo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vBanco As Integer, vTipo As String, vFecha As Date, vFormatoTe As String
Dim vConsecutivo As String


vBanco = cbo.ItemData(cbo.ListIndex)
vTipo = cboDoc.ItemData(cboDoc.ListIndex)
vFormatoTe = cboFormato.ItemData(cboFormato.ListIndex)
vFecha = fxFechaServidor
vConsecutivo = 0

Select Case vFormatoTe
  Case "A" 'Banco Nacional
        
        strSQL = "Select sum(monto) as 'Monto' From Tes_Transacciones Where Estado = 'T' And Tipo = '" _
               & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and documento_base = " & txtNTran.Text
        Call OpenRecordSet(rs, strSQL)
        'Consulta del Detalle
        strSQL = "Select * From Tes_Transacciones Where Estado = 'T' And Tipo = '" _
               & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and documento_base = " & txtNTran.Text _
               & " Order by Nsolicitud"
         Call sbTeBancoNacional(strSQL, rs!Monto)
        rs.Close
       
  Case "B" 'Banco Popular
        strSQL = "Select * From Tes_Transacciones Where Estado = 'T' And Tipo = '" _
               & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and documento_base = " & txtNTran.Text _
               & " Order by Nsolicitud"
        Call sbTeBancoPopular(strSQL)
       
  Case "C" 'BCR
        strSQL = "select sum(dbo.fxTESBCRTestkey(cta_ahorros,monto)) as TestKeyX, sum(Monto) as Monto" _
               & " From Tes_Transacciones " _
               & " Where Estado = 'T' And Tipo = '" & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and documento_base = " & txtNTran.Text
        Call OpenRecordSet(rs, strSQL)
         
         strSQL = "Select * From Tes_Transacciones Where Estado = 'T' And Tipo = '" _
               & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and documento_base = " & txtNTran.Text _
               & " Order by Nsolicitud"
               
        Dim xTestKey As Long
       
        If rs!TestKeyX > 2147483468 Then
                xTestKey = 2147483468
        Else
                xTestKey = rs!TestKeyX
        End If
        
        Call sbTeBCR(strSQL, xTestKey, rs!Monto)
        rs.Close
       
       
      Case "D" 'D - BCR. Empresas

         Call sbTeBCR_Empresarial
    
      Case "E" 'E - BCT. Enlace
          strSQL = "Select * From Tes_Transacciones Where Estado = 'T' And Tipo = '" _
                 & vTipo & "' And ID_Banco=" & vBanco & " And Autoriza='S' and documento_base = " & txtNTran.Text _
                 & " Order by Nsolicitud"
            Call sbTeBCT_Enlace
      
      Case "F" 'F - BCR. Comercial

         Call sbTeBCR_Comercial
         
      Case "G" 'G - BNCR SINPE
         Call sbTeBNCR_Sinpe
         
      Case "DV1", "DV2"
         Call sbTeFormatoEstandar(vFormatoTe)
         
      Case "S" 'SINPE
                
      Case Else
         Call sbTeFormatoEstandar(vFormatoTe)
      
End Select


End Sub


Private Sub sbTeBNCR_Sinpe()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato Empresarial para el BCR. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim strSQL As String, fn, vTesKeyCh As String
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     
     gTesGlobal.BancoConsec = txtNTran.Text
     
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".tef"
     
     Open strArchivo For Output As #1

     'REGISTRO DE CONTROL
     i = 1
    

    'ENCABEZADO: LINEA 1
    strSQL = "exec spTES_BNCR_SINPE_Archivo 1," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
           & "'," & gTesGlobal.BancoConsec & "," & 0
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea1 'Encabezado
        rs.MoveNext
     Loop
     rs.Close
     
  
    'DEBITOS
    strSQL = "exec spTES_BNCR_SINPE_Archivo 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
           & "'," & gTesGlobal.BancoConsec & "," & 0
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea2 'Debitos
        rs.MoveNext
     Loop
     rs.Close
   
    'CREDITOS
    strSQL = "exec spTES_BNCR_SINPE_Archivo 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
           & "'," & gTesGlobal.BancoConsec & "," & 0
    Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        Print #1, rs!Linea3 'Creditos
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     MsgBox strArchivo, vbInformation, "Archivo de Transferencia Generado..."
     
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub




''Procedimiento para crear el nuevo archivo del BCR, Banca Empresarial
Private Sub sbTeBCT_Enlace()
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor


i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = txtNTran.Text
     gTesGlobal.BancoNombre = cbo.Text
     
     iLineInicio = 1
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     
     Open strArchivo For Output As #1

     '******************************
     ' DETALLE DE LA TRANSFERENCIA *
     '******************************

    strSQL = "exec spTES_BCT_Enlace_ArchivoLog " & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc _
            & "','" & gTesGlobal.BancoConsec & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
        strCadena = rs!Linea
        Print #1, strCadena
     
     rs.MoveNext
    Loop
    rs.Close
   
     Close #1   ' Close file.
     
     MsgBox strArchivo, vbInformation, "Archivo de Transferencia Generado..."

Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


''Procedimiento para crear el nuevo archivo del BCR, Banca Empresarial
Private Sub sbTeBCR_Empresarial()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato Empresarial para el BCR. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strSQL = "select  REPLACE(cedula_juridica,'-','') as 'Cedula_Juridica',NOMBRE" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL, 0)
    vNumNegocio = Trim(rs!cedula_juridica)
    vCedulaReg = Trim(rs!cedula_juridica)
    vRazon = "TRANSFERENCIAS " & rs!Nombre
rs.Close

     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     
     gTesGlobal.BancoConsec = txtNTran.Text
     
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     i = 1
    strSQL = "select documento_base,count(*) From Tes_Transacciones" _
         & " where id_banco = " & gTesGlobal.BancoID & " and fecha_emision = '" _
         & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     i = i + 1
     rs.MoveNext
    Loop
    rs.Close
    vConArchivo = Format(i, "000")
     
     
     strSQL = "select dbo.fxTesCantidadTEDiarias('" & Format(vFecha, "yyyy/mm/dd") & "' ," & gTesGlobal.BancoID & ") as 'Cantidad'"
     Call OpenRecordSet(rs, strSQL)
         iLineInicio = rs!Cantidad
     rs.Close
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     
     Open strArchivo For Output As #1

     'REGISTRO DE CONTROL
     i = 1
    
     strCadena = "000"                                                              'Estado 3
     strCadena = strCadena & SIFGlobal.fxStringRelleno(vCedulaReg, "I", "0", 12)    'Cedula Juridica 12
     strCadena = strCadena & vConArchivo                                            'Consecutivo Archivo 3
     strCadena = strCadena & Format(vFecha, "ddmmyyyy")                             'Fecha Aplicacion 8
     strCadena = strCadena & "000000000000"                                         'Cedula de Registro 12
     strCadena = strCadena & "000000000000"                                         '12 TestKey  no se genera, se rellena con ceros
     strCadena = strCadena & "000000"                                               '6 Hora Estado Se rellena con ceros
     strCadena = strCadena & Space(6)                                               'filler 6 espacios en blanco
     strCadena = strCadena & "TLB"                                                  'Tipo de archivo
     strCadena = strCadena & Space(128)                                             'filler 128 espacios en blanco
     strCadena = strCadena & "D"                                                    'Tipo de movinento Debido
    
     Print #1, strCadena
     
   
    'DEBITOS
    strSQL = "exec spTES_BCR_Empresarial_Archivo 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & 100000
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea2 'Debitos
        rs.MoveNext
     Loop
   rs.Close
   
    'CREDITOS
    strSQL = "exec spTES_BCR_Empresarial_Archivo 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & 100000
    Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        Print #1, rs!Linea3 'Creditos
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     MsgBox "Archivo Generado Satisfactoriamente!", vbInformation
     
     
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub




''Procedimiento para crear el nuevo archivo del BCR, Banca Empresarial
Private Sub sbTeBCR_Comercial()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato Empresarial para el BCR. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim strSQL As String, fn, vTesKeyCh As String
Dim strCedJuridica As String
Dim iLineInicio As Integer 'variable para la linea con la que inicia el detalle de las transferencias
Dim strLinea As String

On Error GoTo vError



vPaso = False
vFecha = fxFechaServidor

' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strSQL = "select  REPLACE(cedula_juridica,'-','') as 'Cedula_Juridica',NOMBRE" _
       & " From SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL, 0)
    vNumNegocio = Trim(rs!cedula_juridica)
    vCedulaReg = Trim(rs!cedula_juridica)
    vRazon = "TRANSFERENCIAS " & rs!Nombre
rs.Close

    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     
     gTesGlobal.BancoConsec = txtNTran.Text
     
     gTesGlobal.BancoNombre = cbo.Text
     
     vPaso = True
     i = 1
    strSQL = "select documento_base,count(*) From Tes_Transacciones" _
         & " where id_banco = " & gTesGlobal.BancoID & " and fecha_emision = '" _
         & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     i = i + 1
     rs.MoveNext
    Loop
    rs.Close
    vConArchivo = Format(i, "000")
     
     
     strSQL = "select dbo.fxTesCantidadTEDiarias('" & Format(vFecha, "yyyy/mm/dd") & "' ," & gTesGlobal.BancoID & ") as 'Cantidad'"
     Call OpenRecordSet(rs, strSQL)
         iLineInicio = rs!Cantidad
     rs.Close
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".txt"
     
     Open strArchivo For Output As #1

     'REGISTRO DE CONTROL
     i = 1
    
     strCadena = "000"                                                              'Estado 3
     strCadena = strCadena & SIFGlobal.fxStringRelleno(vCedulaReg, "I", "0", 12)    'Cedula Juridica 12
     strCadena = strCadena & vConArchivo                                            'Consecutivo Archivo 3
     strCadena = strCadena & Format(vFecha, "ddmmyyyy")                             'Fecha Aplicacion 8
     strCadena = strCadena & "000000000000"                                         'Cedula de Registro 12
     strCadena = strCadena & "000000000000"                                         '12 Filler con 0
     strCadena = strCadena & "000000"                                               '6 Hora Estado Se rellena con ceros
     strCadena = strCadena & SIFGlobal.fxStringRelleno("", "D", "0", 138)           '138 Filler con 0
   
     Print #1, strCadena
     
  
    'DEBITOS
    strSQL = "exec spTES_BCR_Comercial_Archivo 2," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & 100000
    Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
        Print #1, rs!Linea2 'Debitos
        rs.MoveNext
     Loop
   rs.Close
   
    'CREDITOS
    strSQL = "exec spTES_BCR_Comercial_Archivo 3," & gTesGlobal.BancoID & ",'" & gTesGlobal.BancoTDoc & "','" & vNumNegocio _
           & "'," & gTesGlobal.BancoConsec & "," & 100000
    Call OpenRecordSet(rs, strSQL)
 
     Do While Not rs.EOF
        Print #1, rs!Linea3 'Creditos
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     MsgBox strArchivo, vbInformation, "Archivo de Transferencia Generado..."
     
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbTeBCR(strSQL As String, Optional vTestKey As Long = 0, Optional vMontoTotal As Currency = 0)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato para el Banco Nacional. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset, strArchivo As String
Dim strPath As String, strCadena As String, i As Integer
Dim vCuentaBanco As String, vFecha As Date, vPaso As Boolean
Dim vConArchivo As String, vCedulaReg As String, vNumNegocio As String, vRazon As String
Dim ySQL As String, fn, vTesKeyCh As String

On Error GoTo vError

gstrQuery = strSQL

vPaso = False
vFecha = fxFechaServidor

'Leer Archivo de Texto : BCRFormat.ini
' Linea 1 -> Numero de Negocio (Registrado en el Sistema del BCR, SCIC)
' Linea 2 -> Cedula de Registro
' Linea 3 -> Razon o Detalle del Pago
i = 1

fn = FreeFile

vRazon = ""
vNumNegocio = ""
vCedulaReg = ""

strArchivo = SIFGlobal.DirectorioDeResultados & "\Configuracion\BCRFormat.ini"
If Dir(strArchivo, vbArchive) = "" Then
  strArchivo = App.Path & "\BCRFormat.ini"
End If

Open strArchivo For Input As #fn
 Do While Not EOF(fn)
   Input #fn, ySQL
   Select Case i
     Case 1
       vNumNegocio = ySQL
     Case 2
       vCedulaReg = ySQL
     Case 3
       vRazon = ySQL
   End Select
   i = i + 1
 Loop
Close #fn   ' Close file.


For i = Len(vRazon) To 30
  vRazon = vRazon & " "
Next i

'Calcular el Numero de Archivo , Numero de la Transferencia en el Dia
i = 1
ySQL = "select documento_base,count(*) From Tes_Transacciones" _
     & " where id_banco = " & cbo.ItemData(cbo.ListIndex) & " and fecha_emision = '" _
     & Format(vFecha, "yyyy/mm/dd") & "' and estado = 'T' group by documento_base"
rs.Open ySQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 i = i + 1
 rs.MoveNext
Loop
rs.Close
vConArchivo = Format(i, "000")
    
    
'Crear y Sacar la cuenta de Tes_Bancos, se Asume que esta cuenta tiene el digito verificador
ySQL = "select Cta from Tes_Bancos where id_Banco = " & cbo.ItemData(cbo.ListIndex)
rs.Open ySQL, glogon.Conection, adOpenStatic
 'Se indica la oficina 001 de apertura por Omision
 vCuentaBanco = "001" & Format(Trim(rs!Cta), "00000000")
rs.Close
    
'Calcular TestKey Complementario (de la primera Linea)
ySQL = "select dbo.fxTESBCRTestkey('" & vCuentaBanco & "'," & vMontoTotal & ") as TestKey"
rs.Open ySQL, glogon.Conection, adOpenStatic
If vTestKey + rs!TestKey > 2147483468 Then
        vTestKey = 2147483468
Else
        vTestKey = vTestKey + rs!TestKey
End If
rs.Close

'Validando Largo del TestKey  = 12
vTesKeyCh = Trim(CStr(vTestKey))
If Len(vTesKeyCh) > 12 Then
  vTestKey = Right(vTesKeyCh, 12)
End If
    
    
     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = txtNTran.Text
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     strArchivo = strArchivo & "\" & gTesGlobal.BancoConsec & ".BCR"
     
     Open strArchivo For Output As #1

     '*****************************************
     'ENCABEZADO DEL FORMATO DE TRANSFERENCIA *
     '*****************************************

     strCadena = "000" 'Estado
     strCadena = strCadena & vNumNegocio            '12 char
     strCadena = strCadena & vConArchivo            '3 char
     strCadena = strCadena & "000000"               '6 Filler
     strCadena = strCadena & vCedulaReg             '12 char
     strCadena = strCadena & Format(vTestKey, "000000000000") '12 TestKey ** Generarlo **
     strCadena = strCadena & "000000"               '6 Hora
     strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
     strCadena = strCadena & Space(21)              'filler 21 char
     strCadena = strCadena & "Y"                    'Señal de Y2k
    
     Print #1, strCadena
     
     
     
     '******************************
     ' DETALLE DE LA TRANSFERENCIA *
     '******************************
     
     'Linea 1 es la de Debito cuenta Bancaria
     
     i = 1
          
        strCadena = "000"                       'Estado Relleno con Ceros
        strCadena = strCadena & "1"             'Concepto 1 = Cuenta Corriente / 2 Cuenta Ahorro
        strCadena = strCadena & "00000"         'Filler 5
        strCadena = strCadena & Mid(Trim(vCuentaBanco), 1, 11) 'Oficina -> 3c, Cuenta -> 7 + 1 Digito verificador
        strCadena = strCadena & "1"             'Moneda  1 = Colones, 2 = Dolares
        strCadena = strCadena & "4"             '2 -> Credito, 4 -> Debito
        strCadena = strCadena & "0000"          'Codigo de Causa
        strCadena = strCadena & Format(gTesGlobal.BancoConsec, "0000") & Format(i, "0000") 'Numero de Documento 8
        strCadena = strCadena & Format((vMontoTotal * 100), "000000000000") '12 Sin Decimales
        strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
        strCadena = strCadena & "0"             'Filler 1
        strCadena = strCadena & vRazon          'Razon de Transferencia (Detalle) 30
     
        Print #1, strCadena
     
     
     Call OpenRecordSet(rs, strSQL)
     
     
     Do While Not rs.EOF
        i = i + 1
      
        strCadena = "000"                       'Estado Relleno con Ceros
        strCadena = strCadena & "2"             'Concepto 1 = Cuenta Corriente / 2 Cuenta Ahorro
        strCadena = strCadena & "00000"         'Filler 5
        strCadena = strCadena & Mid(Trim(rs!Cta_Ahorros), 1, 11) 'Oficina -> 3c, Cuenta -> 7 + 1 Digito verificador
        strCadena = strCadena & "1"             'Moneda  1 = Colones, 2 = Dolares
        strCadena = strCadena & "2"             '2 -> Credito, 4 -> Debito
        strCadena = strCadena & "0000"          'Codigo de Causa
        strCadena = strCadena & Format(gTesGlobal.BancoConsec, "0000") & Format(i, "0000") 'Numero de Documento 8
        strCadena = strCadena & Format((rs!Monto * 100), "000000000000") '12 Sin Decimales
        strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
        strCadena = strCadena & "0"             'Filler 1
        strCadena = strCadena & vRazon          'Razon de Transferencia (Detalle) 30
        
        Print #1, strCadena
        
        rs.MoveNext
     Loop
     rs.Close
     
     Close #1   ' Close file.
     
     MsgBox strArchivo, vbInformation, "Archivo de Transferencia Generado..."
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
   
End Sub



Private Sub sbTeBancoNacional(strSQL As String, curPlanilla As Currency)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato para el Banco Nacional. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim rs As New ADODB.Recordset
Dim strArchivo As String, strCedula As String, strMonto As String
Dim strPath As String, strCadena As String, i As Integer
Dim curMonto1 As Currency, curMonto2 As Currency
Dim curCuentas As Currency, vFecha As Date, vPaso As Boolean
Dim vConcepto As String, vCuentaEmpresa As String

Dim vNumCliente As String, pSQL As String, vCuenta As String

On Error GoTo vError

gstrQuery = strSQL

vPaso = False
vFecha = Format(fxFechaServidor, "dd/mm/yyyy")
curMonto1 = IIf(IsNull(curPlanilla), 0, curPlanilla)
strMonto = CStr(Format(IIf(IsNull(curPlanilla), 0, curPlanilla), "0000000000.00"))
strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)



vConcepto = SIFGlobal.fxStringRelleno("TF " & gPortal.Empresa_Name, "I", " ", 30)  'GLOBALES.gstrNombreEmpresa

pSQL = "select Cta,codigo_Cliente from tes_Bancos" _
       & " Where id_Banco = " & cbo.ItemData(cbo.ListIndex)
Call OpenRecordSet(rs, pSQL)
 vCuentaEmpresa = fxDepuraString(rs!Cta, "-")
 vNumCliente = Trim(rs!Codigo_Cliente & "")
rs.Close

vNumCliente = SIFGlobal.fxStringRelleno(vNumCliente, "I", "0", 6)

     '*****************************************
     'VERIFICA EXISTENCIA DEL DIR. Y ARCHIVO  *
     '*****************************************
     
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir (strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir (strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir (strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir (strArchivo)
        Else
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(strArchivo, vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir (strArchivo)
           End If
        End If
     End If
     
     ChDir (strArchivo)
          
     'Inicializa Variables Globales de Tes_Bancos y Consecutivo
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = txtNTran.Text
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     
     '*****************************************
     'CONFIRMA REALIZACION DE LA TRANSFERECIA *
     '*****************************************
     
     Open strArchivo & "\" & gTesGlobal.BancoConsec & ".ENV " For Output As #1

     '*****************************************
     'ENCABEZADO DEL FORMATO DE TRANSFERENCIA *
     '*****************************************

     strCadena = "1"
     strCadena = strCadena & vNumCliente
     strCadena = strCadena & Format(Day(vFecha), "00") & Format(Month(vFecha), "00") & Format(Year(vFecha), "0000")
     'strCadena = strCadena & "0A645500000010000000000000000000000000000000000000000"
     
     strCadena = strCadena & Format(gTesGlobal.BancoID, "000000000000")
     strCadena = strCadena & "1" & "0000"
     strCadena = strCadena & strMonto
     strCadena = strCadena & "000000000000000000000000"
     
     Print #1, strCadena
     
     
     
     '******************************
     ' DETALLE DE LA TRANSFERENCIA *
     '******************************
     
     Call OpenRecordSet(rs, strSQL)
     
     i = 0
     
     Do While Not rs.EOF
        i = i + 1
        
        ' x[1,2]00,01,XXX,XXXXXX,X
        vCuenta = Replace(Trim(rs!Cta_Ahorros), "-", "")
        
        strCadena = "3" 'Credito
        strCadena = strCadena & Mid(vCuenta, 6, 3) & Mid(vCuenta, 1, 3) & "01"  '  "20001"        '  "000" 'Oficina de Apertura
        strCadena = strCadena & Right(vCuenta, 7) ' & Mid(Trim(rs!Cta_Ahorros), 9, 7)      ' trim(RS!cta_ahorros)
                                'Incluye: 100 o 200 -> Tipo de Cuenta (Corriente-Ahorros)
                                '          01 -> Tipo Moneda (Colones)
                                '      000000 -> Cuenta de la Persona
                                '           0 -> Digito Verificador

        'Suma las Cuentas para Registro de Totales
        ' - Solo tiene que la cuenta de la persona
            
            curCuentas = curCuentas + CCur(Mid(Right(Trim(vCuenta), 7), 1, 6))   'Sin Verificador
            curMonto2 = curMonto2 + rs!Monto
        
        'Fin del Calculo de las cuentas y del Monto de acreditaciones
        
        strCadena = strCadena & Format(i, "00000000") '8d Numero Comprobante (Consecutivo Interno)
                
        strMonto = CStr(Format(rs!Monto, "0000000000.00")) '12d Monto sin el punto decimal
        strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
        
        strCadena = strCadena & strMonto                       'corte
        strCadena = strCadena & SIFGlobal.fxStringRelleno(vConcepto, "D", " ", 30)  '30d Concepto de Pago
        strCadena = strCadena & "00" 'Fin de Linea
                    
        Print #1, strCadena
        
        rs.MoveNext
     Loop
     rs.Close
     
     '*********************************************************
     'CREA ULTIMA LINEA DE DETALLE CON EL DEBITO A LA EMPRESA *
     '*********************************************************
     strCadena = "2" & Mid(Trim(vCuentaEmpresa), 1, 3)   'Movimiento de Debito, y 000 Sucursal de Apertura
     strCadena = strCadena & "10001" 'Cuenta Corriente y Moneda en Colones
     strCadena = strCadena & Right(Trim(vCuentaEmpresa), 7) 'Cuenta de la Empresa  + Digito Verificador
     strCadena = strCadena & Format(i + 1, "00000000") 'Numero Comprobante
        
      strMonto = CStr(Format(curMonto2, "0000000000.00")) '12d Monto sin el punto decimal
      strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
        
     strCadena = strCadena & strMonto 'Total de los Creditos para Debitar a esta cuenta
     strCadena = strCadena & vConcepto '30d Concepto de Pago
     strCadena = strCadena & "00" 'Fin de Linea
    
     Print #1, strCadena
     curCuentas = curCuentas + CCur(Mid(Right(Trim(vCuentaEmpresa), 7), 1, 6))   'Sin Verificador
     


     '**************************************************
     'REGISTRO DE CONTROL DEL ARCHIVO DE TRANSFERENCIA *
     '**************************************************
     
     strCadena = "4" 'Codigo de Control de registro
     strMonto = CStr(Format(curMonto1 + curMonto2, "0000000000000.00")) 'Suma Debitos y Creditos de la Transferencia
     strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
     strCadena = strCadena & strMonto
     
     strMonto = CStr(Format(curCuentas, "0000000000")) 'Sumatoria de Cuentas
     strCadena = strCadena & strMonto
     strCadena = strCadena & "0000000000"
     strCadena = strCadena & "000000000000"
     strCadena = strCadena & "000000000000"
     strCadena = strCadena & "00000000"
     
     Print #1, strCadena
     
     Close #1   ' Close file.
     
     MsgBox strArchivo & "\" & gTesGlobal.BancoConsec & ".ENV", vbInformation, "Archivo de Transferencia Generado..."
     
     
Exit Sub

vError:
       
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbTeBancoPopular(strSQL As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Emite la Transferencia en formato para el Banco Popular. Genera archivo de
'               texto en la direccion "C:\NombreEmpresa\Banco\Fecha\ConsecutivoDeposito.txt"
'               y finalmente despliega el formulario de control de transferencias.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim recCheques As New ADODB.Recordset
Dim strArchivo As String, strNombre As String
Dim strMonto As String, strCuenta As String
Dim strPath As String, strCadena As String
Dim strSelf As String, strProducto As String
Dim strEstado As String, strTipo As String
Dim strFecha As String, intI As Integer
Dim vFecha As Date, vPaso As Boolean


On Error GoTo vError

vFecha = Format(fxFechaServidor, "dd/mm/yyyy")

'Cada Global, para ser utilizada en el modulo de Ejecucion de la Transferencia
gstrQuery = strSQL
vPaso = False

With recCheques
     .Open strSQL, glogon.Conection, adOpenStatic
   
     strArchivo = SIFGlobal.DirectorioDeResultados & "\Transferencias"
     strPath = Dir(strArchivo, vbDirectory)
     
     If strPath = "" Then
        ChDir ("C:\")
        MkDir Trim(strArchivo)
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        MkDir Trim(strArchivo)
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
        MkDir Trim(strArchivo)
     Else
        strArchivo = strArchivo & "\" & Trim(cbo.Text)
        strPath = Dir(strArchivo, vbDirectory)
        
        If strPath = "" Then
           ChDir ("C:\")
           MkDir Trim(strArchivo)
           strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           MkDir Trim(strArchivo)
        Else
        strArchivo = strArchivo & "\" & Format(vFecha, "yyyy.mm.dd")
           strPath = Dir(Trim(strArchivo), vbDirectory)
           
           If strPath = "" Then
              ChDir ("C:\")
              MkDir Trim(strArchivo)
           End If
        End If
     End If
     
     ChDir Trim(strArchivo)
          
     'Inicializa Variables Globales y Consecutivos
     gTesGlobal.BancoID = cbo.ItemData(cbo.ListIndex)
     gTesGlobal.BancoTDoc = cboDoc.ItemData(cboDoc.ListIndex)
     gTesGlobal.BancoConsec = txtNTran.Text
     gTesGlobal.BancoNombre = cbo.Text
     vPaso = True
     
     Open strArchivo & "\" & gTesGlobal.BancoConsec & ".txt" For Output As #1
     
     Do While Not .EOF
        
        
        Select Case Len(Trim(!Codigo))
           Case 8
               strCadena = "0" & Mid(Trim(!Codigo), 1, 1) & "0" & Mid(Trim(!Codigo), 2, 7)
           Case 9
               strCadena = "0" & Trim(!Codigo)
           Case Is < 8
               strCadena = Format(!Codigo, "0000000000")
           Case Is > 10
               strCadena = Mid(Trim(!Codigo), 1, 4) & "0" & Mid(Trim(!Codigo), 6, 5)
           Case Else
               strCadena = Trim(!Codigo)
        End Select
        
        strNombre = Trim(!Beneficiario)
                
        If Len(strNombre) > 30 Then
         strNombre = Mid(strNombre, 1, 30)
        Else
         Do Until Len(strNombre) = 30
           strNombre = strNombre & " "
         Loop
        End If
        
        strCadena = strCadena & strNombre
        
        strCuenta = IIf(IsNull(!Cta_Ahorros), "0", Trim(!Cta_Ahorros))
        
        If Len(strCuenta) > 13 Then
           strCuenta = Mid(strCuenta, 1, 13)
        Else
         Do Until Len(strCuenta) = 13
            strCuenta = "0" & strCuenta
         Loop
        End If
        
        strCadena = strCadena & strCuenta
        
        strSelf = " "
        strCadena = strCadena & strSelf
        
        strMonto = CStr(Format(!Monto, "000000000.00"))
        strMonto = Mid(strMonto, 1, Len(strMonto) - 3) & Mid(strMonto, Len(strMonto) - 1, 2)
        
        strCadena = strCadena & strMonto
                
        strFecha = Format(Day(vFecha), "00")
        strFecha = strFecha & Format(Month(vFecha), "00")
        strFecha = strFecha & Format(Year(vFecha), "0000")
        
        strCadena = strCadena & strFecha
        
        strTipo = "A"
        strCadena = strCadena & strTipo
        
        strProducto = "06"
        strCadena = strCadena & strProducto
        
        strEstado = "P"
        strCadena = strCadena & strEstado
        
        strCadena = strCadena & strFecha
        strCadena = strCadena & strMonto
                
        Print #1, strCadena
 
        .MoveNext
     Loop
     Close #1   ' Close file.
     
     .Close
         
End With

MsgBox strArchivo & "\" & gTesGlobal.BancoConsec & ".txt", vbInformation, "Archivo de Transferencia Generado..."
     
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
     
End Sub



