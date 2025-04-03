VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_GestionPago 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguros: Gestion de Pago"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraVencimientos 
      Height          =   3735
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   9495
      Begin VB.OptionButton optVencimientos 
         Appearance      =   0  'Flat
         Caption         =   "Desembolsado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   3
         Left            =   6480
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optVencimientos 
         Appearance      =   0  'Flat
         Caption         =   "Generado / Remesado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   2
         Left            =   4200
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton optVencimientos 
         Appearance      =   0  'Flat
         Caption         =   "Generados / No Remesado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   2415
      End
      Begin VB.OptionButton optVencimientos 
         Appearance      =   0  'Flat
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComctlLib.ListView lswVencimientos 
         Height          =   3015
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Vencimiento"
            Object.Width           =   3775
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Casos"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Remesa"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Comision.Rem."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Comision Monto"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbVence 
         Height          =   312
         Left            =   8280
         TabIndex        =   6
         Top             =   156
         Width           =   1176
         _ExtentX        =   2064
         _ExtentY        =   556
         ButtonWidth     =   1746
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ocultar"
               Key             =   "Ocultar"
               Object.ToolTipText     =   "Cierra ventana"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11520
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_GestionPago.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_GestionPago.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_GestionPago.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_GestionPago.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12735
      _Version        =   524288
      _ExtentX        =   22463
      _ExtentY        =   11456
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
      MaxCols         =   497
      SpreadDesigner  =   "frmSeguros_GestionPago.frx":1A188
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.Toolbar tlbX 
      Height          =   336
      Left            =   5400
      TabIndex        =   3
      Top             =   720
      Width           =   5724
      _ExtentX        =   10107
      _ExtentY        =   582
      ButtonWidth     =   3069
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Casos en Riesgo"
            Key             =   "Buscar"
            Object.ToolTipText     =   "Buscar archivos"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Actualiza Cobros"
            Key             =   "Actualiza"
            Object.ToolTipText     =   "Actualiza Info. Cobros"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Vencimientos ?"
            Key             =   "Vencimientos"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbProceso 
      Height          =   336
      Left            =   2040
      TabIndex        =   4
      Top             =   7920
      Width           =   8736
      _ExtentX        =   15399
      _ExtentY        =   582
      ButtonWidth     =   7038
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar (Autorización de Comisiones / Pagos)"
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Aplicar Calculo de Pago"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar (Pólizas en Riesgo de Cobranza)"
            Key             =   "Cerrar"
            Object.ToolTipText     =   "Cerrar Pólizas Selecionadas"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.DateTimePicker dtpVence 
      Height          =   330
      Left            =   3240
      TabIndex        =   12
      Top             =   240
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   582
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   5400
      TabIndex        =   13
      Top             =   240
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
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
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   12840
      X2              =   0
      Y1              =   700
      Y2              =   700
   End
   Begin VB.Label Label1 
      Caption         =   "Casos con diferencias en la balanza de cobros..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha vencimiento de Pago..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmSeguros_GestionPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Private Sub Form_Load()

 vModulo = 17
 
 vGrid.MaxRows = 0
 vGrid.MaxCols = 9
 
 dtpVence.Value = fxFechaServidor
 
strSQL = "select cod_aseguradora as 'IdX', rtrim(nombre) as 'ItmX' from seguros_Aseguradoras"
Call sbCbo_Llena_New(cboAseguradora, strSQL, False, True)

 
End Sub

Private Sub optVencimientos_Click(Index As Integer)

Me.MousePointer = vbHourglass

strSQL = "select FECHA_VENCE, count(*) as 'Casos'" _
        & " , isnull(sum(monto_pago),0) as 'MontoPago'  , isnull(sum(COMISION_VENDEDOR),0) as 'ComisionVendedor'" _
        & " , isnull(max(COD_REMESA),0) as 'RemesaPago', isnull(max(COD_REMESA_COMISION),0) as 'RemesaComision'" _
        & " From SEGUROS_PAGOS " _
        & " Where NUM_CUOTA > 0  and cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
 
 
Select Case True
   Case optVencimientos.Item(0).Value 'Pendientes
        strSQL = strSQL & " and cod_remesa is null and Estado_Pago = 'P' and comision_monto_base is null"

   Case optVencimientos.Item(1).Value 'Generadas / No Remesadas
        strSQL = strSQL & " and cod_remesa is null and Estado_Pago = 'A'"
        
   Case optVencimientos.Item(2).Value 'Generadas / Remesadas
        strSQL = strSQL & " and cod_remesa is not null and Estado_Pago = 'A'"
        
   Case optVencimientos.Item(3).Value 'Desembolsadas
        strSQL = strSQL & " and cod_remesa is not null and Estado_Pago not in('P', 'A') and tesoreria_Solicitud is not null"
   
End Select
 
strSQL = strSQL & " group by fecha_vence order by Fecha_Vence desc"

lswVencimientos.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswVencimientos.ListItems.Add(, , Format(rs!Fecha_Vence, "dd/mm/yyyy"))
      itmX.SubItems(1) = rs!Casos
      itmX.SubItems(2) = rs!RemesaPago
      itmX.SubItems(3) = Format(rs!MontoPago, "Standard")
      itmX.SubItems(4) = rs!RemesaComision
      itmX.SubItems(5) = Format(rs!ComisionVendedor, "Standard")
      
  rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

End Sub

Private Sub tlbProceso_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, i As Long
Dim vMensaje As String


On Error GoTo vError

vMensaje = ""

Select Case Button.Key
  Case "Aplicar"
        i = MsgBox("Esta seguro que desea >> Activar << los pagos de póliza y comisiones enlazadas?", vbYesNo)
        If i = vbYes Then
            Me.MousePointer = vbHourglass
        
            strSQL = "exec spSeguros_PagoComisionActivacion '" & Format(dtpVence.Value, "yyyy/mm/dd") & " 23:59:59'"
            Call ConectionExecute(strSQL)
            
            vMensaje = "Pagos de Pólizas Activado Satisfactoriamente...!"
          
        End If 'i = vbYes
  
  Case "Cerrar"
        i = MsgBox("Esta seguro que desea >> Cerrar << las polizas en Riegos de Cobranza?", vbYesNo)
        If i = vbYes Then
            
            Me.MousePointer = vbHourglass
            
            For i = 1 To vGrid.MaxRows
               vGrid.Row = i
               vGrid.Col = 9
               
               If vGrid.Value = vbChecked Then
                 vGrid.Col = 5
'                 Call sbPolizaCierra(cboAseguradora.ItemData (cboAseguradora.ListIndex), vGrid.Text)
               End If
            
            Next i
            
            vMensaje = "Pólizas Cerradas Satisfactoriamente...!"
        End If 'i = vbYes
      
End Select

Me.MousePointer = vbDefault

If Len(vMensaje) > 0 Then
    MsgBox vMensaje, vbInformation
End If

'Vuelve a Consultar
Call tlbX_ButtonClick(tlbX.Buttons.Item(1))

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tlbVence_ButtonClick(ByVal Button As MSComctlLib.Button)
 fraVencimientos.Visible = False
End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Me.MousePointer = vbHourglass

Select Case Button.Key
  Case "Buscar"
      strSQL = "select P.cedula,S.nombre,P.num_Cuota,P.Cuota, P.Num_Poliza,P.COD_PRODUCTO,P.Pagado_Total,P.Cobrado_Total,0" _
             & " from SEGUROS_REGISTRO P inner join Socios S on P.cedula = S.cedula" _
             & " Where P.ESTADO = 'A' and P.PAGADO_TOTAL - P.COBRADO_TOTAL > CUOTA " _
             & " and P.cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
      Call sbCargaGrid(vGrid, 9, strSQL, True)
      
      vGrid.MaxRows = vGrid.MaxRows - 1
  
  Case "Actualiza"
      strSQL = "exec spSeguros_CobrosActualiza '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
      Call ConectionExecute(strSQL)
              
      'Vuelve a Consultar
      Call tlbX_ButtonClick(tlbX.Buttons.Item(1))
     
     
  Case "Vencimientos"
      fraVencimientos.Visible = True
      optVencimientos.Item(0).Value = True
      Call optVencimientos_Click(0)
    
      
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
