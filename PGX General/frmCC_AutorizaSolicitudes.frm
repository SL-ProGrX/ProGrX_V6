VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCC_AutorizaSolicitudes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorización de tramites con Traslado a Tesorería (Con advertencia de duplicidad!)"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7020
   ScaleWidth      =   10575
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   4572
      Left            =   0
      TabIndex        =   5
      Top             =   1800
      Width           =   10572
      _Version        =   1441793
      _ExtentX        =   18648
      _ExtentY        =   8064
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_AutorizaSolicitudes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCC_AutorizaSolicitudes.frx":6862
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   4
      Top             =   6888
      Visible         =   0   'False
      Width           =   10572
      _ExtentX        =   18653
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.ComboBox cboModulo 
      Height          =   312
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7646
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   3840
      TabIndex        =   9
      Top             =   840
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   492
      Index           =   0
      Left            =   7320
      TabIndex        =   10
      Top             =   720
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
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
      Appearance      =   16
      Picture         =   "frmCC_AutorizaSolicitudes.frx":D0C4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnTool 
      Height          =   492
      Index           =   1
      Left            =   8880
      TabIndex        =   11
      Top             =   720
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Autorizar"
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
      Appearance      =   16
      Picture         =   "frmCC_AutorizaSolicitudes.frx":D7C4
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   204
      Left            =   120
      TabIndex        =   12
      Top             =   1524
      Width           =   204
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
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
      Alignment       =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtCargaTotal 
      Height          =   312
      Left            =   8280
      TabIndex        =   14
      Top             =   6480
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   550
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeShortcutBar.ShortcutCaption lblX 
      Height          =   372
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   10572
      _Version        =   1441793
      _ExtentX        =   18648
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "                                Seleccione los tramites pendientes para autorización de la transacción"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   8
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   1452
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total...:"
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
      Height          =   252
      Index           =   2
      Left            =   6960
      TabIndex        =   2
      Top             =   6480
      Width           =   1212
   End
   Begin VB.Label lblTipoCaso 
      BackStyle       =   0  'Transparent
      Caption         =   "Módulo"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1452
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13212
   End
End
Attribute VB_Name = "frmCC_AutorizaSolicitudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub btnTool_Click(Index As Integer)
Select Case Index
  Case 0 'Buscar
    Call sbBuscar(cboModulo.ItemData(cboModulo.ListIndex))
  Case 1 'Autorizar
    Call sbAutoriza(cboModulo.ItemData(cboModulo.ListIndex))
End Select
End Sub

Private Sub cboModulo_Click()
Call sbEncabezadolista(cboModulo.ItemData(cboModulo.ListIndex))
End Sub

Private Sub chkTodos_Click()
Dim i As Integer, curTotal As Currency
Dim x As Integer

On Error GoTo vError

x = 4
If cboModulo.ItemData(cboModulo.ListIndex) = 2 Then x = 5

txtCargaTotal.Text = 0
vPaso = True

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkTodos.Value
  
   If chkTodos.Value = vbChecked And x > 0 Then
       txtCargaTotal.Text = Format(CCur(txtCargaTotal.Text) + CCur(lsw.ListItems.Item(i).SubItems(x)), "Standard")
   End If
  
Next i

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Load()
Dim strSQL As String

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

lsw.ListItems.Clear

cboModulo.Clear
cboModulo.AddItem "Credito"
cboModulo.ItemData(cboModulo.ListCount - 1) = CStr(1)

cboModulo.AddItem "Fondos"
cboModulo.ItemData(cboModulo.ListCount - 1) = CStr(2)
cboModulo.AddItem "Liquidación"
cboModulo.ItemData(cboModulo.ListCount - 1) = CStr(3)
cboModulo.AddItem "Beneficios"
cboModulo.ItemData(cboModulo.ListCount - 1) = CStr(4)


cboModulo.AddItem "Hipotecario Desembolsos"
cboModulo.ItemData(cboModulo.ListCount - 1) = CStr(5)


cboModulo.Text = "Credito"

dtpInicio.Value = Format(fxFechaServidor, "dd/mm/yyyy")
dtpCorte.Value = dtpInicio.Value

strSQL = "select ID_BANCO as 'IdX' ,rtrim(DESCRIPCION) as 'itmx'from TES_BANCOS where ESTADO = 'A' and supervision = 1"

Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
        




End Sub


Private Sub sbEncabezadolista(vModulo As Integer)
lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
Select Case vModulo
    Case 1
       lsw.ColumnHeaders.Add , , "Operación", 1440
       lsw.ColumnHeaders.Add , , "Código", 1440
       lsw.ColumnHeaders.Add , , "Identificación", 2140
       lsw.ColumnHeaders.Add , , "Nombre", 4500
       lsw.ColumnHeaders.Add , , "Monto", 2440, vbRightJustify
    Case 2
       lsw.ColumnHeaders.Add , , "Liq.Id.", 1440
       lsw.ColumnHeaders.Add , , "Identificación", 1840
       lsw.ColumnHeaders.Add , , "Nombre", 4500
       lsw.ColumnHeaders.Add , , "Contrato", 1440
       lsw.ColumnHeaders.Add , , "Plan", 1440
       lsw.ColumnHeaders.Add , , "Monto", 2440, vbRightJustify
    Case 3
       lsw.ColumnHeaders.Add , , "Liq.Id.", 1440
       lsw.ColumnHeaders.Add , , "Tipo", 1840
       lsw.ColumnHeaders.Add , , "Identificación", 1840
       lsw.ColumnHeaders.Add , , "Nombre", 4500
       lsw.ColumnHeaders.Add , , "Monto", 2440, vbRightJustify
    Case 4
       lsw.ColumnHeaders.Add , , "Bene.Id.", 1440
       lsw.ColumnHeaders.Add , , "Tipo", 1840
       lsw.ColumnHeaders.Add , , "Identificación", 1840
       lsw.ColumnHeaders.Add , , "Nombre", 4500
       lsw.ColumnHeaders.Add , , "Monto", 2440, vbRightJustify
       
    Case 5 'Hipotecario
       lsw.ColumnHeaders.Add , , "Id Desembolso", 1440
       lsw.ColumnHeaders.Add , , "Operación", 1440
       lsw.ColumnHeaders.Add , , "Beneficiario", 4500
       lsw.ColumnHeaders.Add , , "Monto", 2440, vbRightJustify
       lsw.ColumnHeaders.Add , , "D.Cedula", 1500
       lsw.ColumnHeaders.Add , , "D.Nombre", 2500

End Select

End Sub



Private Sub sbBuscar(vModulo As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError


vPaso = True

txtCargaTotal.Text = "0"
chkTodos.Value = xtpUnchecked


Select Case vModulo
    Case 1
        strSQL = "select R.id_solicitud,R.codigo,S.cedula,S.nombre,R.monto_girado" _
               & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
               & " inner join Catalogo C on R.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
               & " where R.estadosol='F' and R.fechaforp between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59'" _
               & " and R.tesoreria is null and R.estado in('A','C') and id_solicitud not in(select id_solicitud from CRD_REMESAS_TES_DETALLE)" _
               & " and dbo.fxTesSupervisa(S.cedula,S.nombre,R.monto_girado,0,'C') = 1 and R.TES_SUPERVISION_FECHA is null"
        If cboBanco.Text <> "TODOS" Then
           strSQL = strSQL & " And R.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
        End If
    
            Call OpenRecordSet(rs, strSQL)
            
            PrgBar.Max = rs.RecordCount + 1
            PrgBar.Value = 1
            PrgBar.Visible = True
            
            With lsw
             .ListItems.Clear
             Do While Not rs.EOF
               Set itmX = .ListItems.Add(, , rs!Id_Solicitud)
                   
                   itmX.SubItems(1) = rs!Codigo
                   itmX.SubItems(2) = rs!Cedula
                   itmX.SubItems(3) = rs!Nombre
                   itmX.SubItems(4) = Format(rs!monto_girado, "Standard")
                   itmX.Checked = chkTodos.Value
                     
                   If itmX.Checked Then
                        txtCargaTotal.Text = txtCargaTotal.Text + CCur(itmX.SubItems(4))
                   End If
                    
                   rs.MoveNext
                    
                   PrgBar.Value = PrgBar.Value + 1
             Loop
            End With


    Case 2
        strSQL = "Select L.Consec,C.Cedula,S.nombre,L.Cod_Plan,L.Cod_Contrato" _
               & ",case when L.Total_Girar is null then L.Aportes_Liq+L.Rendi_Liq - isnull(L.multa_retiro,0) else L.Total_Girar end as 'Total_Girar'" _
               & " From Fnd_Liquidacion L inner join Fnd_Contratos C on L.Cod_Operadora=C.Cod_Operadora " _
               & " and L.Cod_Plan = C.Cod_Plan and L.Cod_Contrato = C.Cod_Contrato" _
               & " inner join Socios S on C.cedula = S.cedula" _
               & " Where L.Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
               & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59' And L.Traspaso_tesoreria is Null and L.TES_SUPERVISION_FECHA is null" _
               & " and  dbo.fxTesSupervisa(C.cedula,S.nombre,isnull(L.Total_Girar,L.Aportes_Liq+L.Rendi_Liq - isnull(L.multa_retiro,0)),0,'C') = 1"
        If cboBanco.Text <> "TODOS" Then
             strSQL = strSQL & " And L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
        End If
                
         Call OpenRecordSet(rs, strSQL)
         
         PrgBar.Max = rs.RecordCount + 1
         PrgBar.Value = 1
         PrgBar.Visible = True

        With lsw
         .ListItems.Clear
         Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , rs!consec)
               
               itmX.SubItems(1) = rs!Cedula
               itmX.SubItems(2) = rs!Nombre
               itmX.SubItems(3) = rs!COD_CONTRATO
               itmX.SubItems(4) = rs!cod_Plan
               itmX.SubItems(5) = Format(rs!TOTAL_GIRAR, "Standard")
               itmX.Checked = chkTodos.Value
                 
               If itmX.Checked Then
                    txtCargaTotal.Text = txtCargaTotal.Text + CCur(itmX.SubItems(5))
               End If
                
                rs.MoveNext
                
                PrgBar.Value = PrgBar.Value + 1
         Loop
        End With

    Case 3

        strSQL = "Select L.consec,S.cedula,S.nombre,L.TNeto" _
               & ",case when L.EstadoActLiq = 'A' then 'Ren.Asociación' when  L.EstadoActLiq = 'P' then 'Ren.Patronal' end as 'Tipo'" _
               & " from Liquidacion L inner join Socios S on L.cedula = S.cedula" _
               & " where L.FecLiq between '" & Format(dtpInicio, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte, "yyyy/mm/dd") & " 23:59:59' and L.Ubicacion='T'" _
               & " and L.Estado = 'P' and L.TES_SUPERVISION_FECHA is null and dbo.fxTesSupervisa(S.cedula,S.nombre,L.TNeto,0,'L') = 1"
        
        If cboBanco.Text <> "TODOS" Then
           strSQL = strSQL & " And L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
        End If
    
        Call OpenRecordSet(rs, strSQL)
        
        PrgBar.Max = rs.RecordCount + 1
        PrgBar.Value = 1
        PrgBar.Visible = True
        
        With lsw
         .ListItems.Clear
         Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , rs!consec)
               
               itmX.SubItems(1) = rs!Tipo
               itmX.SubItems(2) = rs!Cedula
               itmX.SubItems(3) = rs!Nombre
               itmX.SubItems(4) = Format(rs!TNETO, "Standard")
               itmX.Checked = chkTodos.Value
                 
               If itmX.Checked Then
                    txtCargaTotal.Text = txtCargaTotal.Text + CCur(itmX.SubItems(4))
               End If
                
                rs.MoveNext
                
               PrgBar.Value = PrgBar.Value + 1
         Loop
        End With
    Case 4
        strSQL = "Select B.Cedula,B.consec,B.cod_beneficio,S.Nombre,B.monto" _
             & " from afi_bene_pago B inner join socios S on B.cedula = S.cedula" _
             & " inner join afi_bene_otorga O on B.cod_beneficio = O.cod_beneficio and B.consec = O.consec" _
             & " inner join Afi_Estados_Persona E on S.EstadoActual = E.Cod_Estado" _
             & " inner join Tes_Bancos Ban on B.cod_Banco = Ban.id_Banco" _
             & " where O.cod_remesa is null and B.TES_SUPERVISION_FECHA is null" _
             & "   and O.registra_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd 00:00:00") & "' and '" & Format(dtpCorte.Value, "yyyy/mm/dd 23:59:59") & "'" _
             & "   and B.ESTADO = 'S' and B.tesoreria is null and dbo.fxTesSupervisa(B.cedula,S.nombre,B.monto,0,'C') = 1"
            
            If cboBanco.Text <> "TODOS" Then
              strSQL = strSQL & " And B.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
            End If
            Call OpenRecordSet(rs, strSQL)
            
            PrgBar.Max = rs.RecordCount + 1
            PrgBar.Value = 1
            PrgBar.Visible = True
        
            With lsw
             .ListItems.Clear
             Do While Not rs.EOF
               Set itmX = .ListItems.Add(, , rs!consec)
                   
                   itmX.SubItems(1) = rs!cod_beneficio
                   itmX.SubItems(2) = rs!Cedula
                   itmX.SubItems(3) = rs!Nombre
                   itmX.SubItems(4) = Format(rs!Monto, "Standard")
                   itmX.Checked = chkTodos.Value
                     
                   If itmX.Checked Then
                        txtCargaTotal.Text = txtCargaTotal.Text + CCur(itmX.SubItems(4))
                   End If
                    
                   rs.MoveNext
                    
                   PrgBar.Value = PrgBar.Value + 1
             Loop
            End With




    Case 5 'Hipotecario
        strSQL = "select D.CodigoDesembolso,D.NumeroOperacion,D.Beneficiario,D.Monto,D.RegistroFecha,D.RegistroUsuario" _
               & ",S.cedula,S.nombre,R.codigo,D.TES_SUPERVISION_FECHA  " _
               & " From ViviendaDesembolsos D inner join Reg_Creditos R on D.numeroOperacion = R.id_solicitud" _
               & " inner join Socios S on R.cedula = S.cedula" _
               & " where D.TesoreriaRemesa is null and D.TES_SUPERVISION_FECHA is null" _
               & " and D.RegistroFecha between '" & Format(dtpInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(dtpCorte, "yyyy/mm/dd") _
               & " 23:59:59'" _
               & " and dbo.fxTesSupervisa(D.Identificacion,D.Beneficiario,D.Monto,0,'V') = 1 --as 'Duplicado'"
            
            If cboBanco.Text <> "TODOS" Then
              strSQL = strSQL & " And B.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
            End If
            Call OpenRecordSet(rs, strSQL)
            
            PrgBar.Max = rs.RecordCount + 1
            PrgBar.Value = 1
            PrgBar.Visible = True
        
            With lsw
             .ListItems.Clear
             Do While Not rs.EOF
               Set itmX = .ListItems.Add(, , rs!CodigoDesembolso)
                   
                   itmX.SubItems(1) = rs!NumeroOperacion
                   itmX.SubItems(2) = rs!Beneficiario
                   itmX.SubItems(3) = Format(rs!Monto, "Standard")
                   
                   
                   itmX.SubItems(4) = rs!Cedula
                   itmX.SubItems(5) = rs!Nombre
                   
                   itmX.Checked = chkTodos.Value
                     
                     
                     
                   If itmX.Checked Then
                        txtCargaTotal.Text = txtCargaTotal.Text + CCur(itmX.SubItems(4))
                   End If
                    
                   rs.MoveNext
                    
                   PrgBar.Value = PrgBar.Value + 1
             Loop
            End With


End Select

vPaso = False

PrgBar.Visible = False

If rs.State = 1 Then rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lsw.ListItems.Clear
 If rs.State = 1 Then rs.Close
 
End Sub

Private Sub sbAutoriza(vModulo As Integer)
Dim strSQL As String
Dim i As Integer



Select Case vModulo
    Case 1
            
            With lsw.ListItems
                PrgBar.Max = .Count
                PrgBar.Value = 1
                PrgBar.Visible = True
                For i = 1 To .Count
                    If .Item(i).Checked Then
                    strSQL = "update REG_CREDITOS SET TES_SUPERVISION_USUARIO = '" & glogon.Usuario & "' , TES_SUPERVISION_FECHA  = dbo.MyGetdate()" _
                          & " where id_solicitud = " & .Item(i).Text
                    Call ConectionExecute(strSQL)
                    End If
                    If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
                   
                Next i
            End With


    Case 2

        With lsw.ListItems
            PrgBar.Max = .Count
            PrgBar.Value = 1
            PrgBar.Visible = True
            For i = 1 To .Count
                If .Item(i).Checked Then
                strSQL = "update Fnd_Liquidacion SET TES_SUPERVISION_USUARIO = '" & glogon.Usuario & "' , TES_SUPERVISION_FECHA  = dbo.MyGetdate()" _
                      & " where consec = " & .Item(i).Text
                Call ConectionExecute(strSQL)
                End If
                If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
            Next i
        End With
            
    Case 3
    
        With lsw.ListItems
            PrgBar.Max = .Count
            PrgBar.Value = 1
            PrgBar.Visible = True
            For i = 1 To .Count
                If .Item(i).Checked Then
                strSQL = "update Liquidacion SET TES_SUPERVISION_USUARIO = '" & glogon.Usuario & "' , TES_SUPERVISION_FECHA  = dbo.MyGetdate()" _
                      & " where consec = " & .Item(i).Text
                Call ConectionExecute(strSQL)
                End If
                If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
            Next i
         End With

    Case 4
        With lsw.ListItems
             PrgBar.Max = .Count
             PrgBar.Value = 1
             PrgBar.Visible = True
             For i = 1 To .Count
                If .Item(i).Checked Then
                    strSQL = "update afi_bene_pago SET TES_SUPERVISION_USUARIO = '" & glogon.Usuario & "' , TES_SUPERVISION_FECHA  = dbo.MyGetdate()" _
                           & " where consec = " & .Item(i).Text _
                           & " and cod_beneficio = '" & Trim(.Item(i).SubItems(1)) & "'"
                    Call ConectionExecute(strSQL)
                End If
                If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
                   
                Next i
        End With
            


    Case 5 'Hipotecario
            
            With lsw.ListItems
                PrgBar.Max = .Count
                PrgBar.Value = 1
                PrgBar.Visible = True
                For i = 1 To .Count
                    If .Item(i).Checked Then
                    strSQL = "update ViviendaDesembolsos SET TES_SUPERVISION_USUARIO = '" & glogon.Usuario & "' , TES_SUPERVISION_FECHA  = dbo.MyGetdate()" _
                          & " where CodigoDesembolso = " & .Item(i).Text
                    Call ConectionExecute(strSQL)
                    End If
                    If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
                   
                Next i
            End With


End Select
PrgBar.Visible = False
Call sbBuscar(cboModulo.ItemData(cboModulo.ListIndex))
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim i As Integer

If vPaso Then Exit Sub

i = 4
If cboModulo.ItemData(cboModulo.ListIndex) = 2 Then i = 5

If Not IsNumeric(txtCargaTotal.Text) Then
    txtCargaTotal.Text = 0
End If

If Item.Checked Then
    txtCargaTotal.Text = Format(CCur(txtCargaTotal.Text) + CCur(Item.SubItems(i)), "Standard")
Else
    txtCargaTotal.Text = Format(CCur(txtCargaTotal.Text) - CCur(Item.SubItems(i)), "Standard")
End If

End Sub


