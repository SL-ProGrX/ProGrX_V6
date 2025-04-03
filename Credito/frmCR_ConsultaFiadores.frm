VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_ConsultaOpcionDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de la Operación"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmCR_ConsultaFiadores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   8010
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaFiadores.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lswConsulta 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7223
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblConsulta 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "xx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmCR_ConsultaOpcionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lngOperacion As Long, iOpcion As Integer
Dim strSQL As String

Private Sub Form_Load()
 lngOperacion = frmCR_ConsultaDetalle.lblId_Solicitud
 Call CargaDatos
End Sub


Sub CargaDatos()
Dim rs As New ADODB.Recordset

Call ConfiguraLista

Me.Refresh

Select Case GLOBALES.gOpcion
 Case 0 'fiadores
   lblConsulta.Caption = "Consulta Fiadores de la Operación"
   strSQL = "select * from fiadores where estado = 'A' and id_solicitud = " & lngOperacion
 Case 1 'cuotas Morosas
   lblConsulta.Caption = "Consulta Cuotas Morosas Activas"
   strSQL = "select * from morosidad where estado = 'A' and id_solicitud = " & lngOperacion
 Case 2 'Refundiciones
   lblConsulta.Caption = "Consulta Refundiciones de la Operación"
   strSQL = "select * from refundiciones where id_solicitudr = " & lngOperacion
 Case 3 'Desembolosos
   lblConsulta.Caption = "Consulta Desembolsos de la Operación"
   strSQL = "select * from desembolsos where id_solicitud = " & lngOperacion
 Case 4 'Abonos
   lblConsulta.Caption = "Consulta Abonos Ordinarios y ExtraOrdinarios"
   strSQL = "select * from creditos_DT where estado = 'A' and id_solicitud = " _
          & lngOperacion & " Order by fechas"
 Case Else
   lblConsulta.Caption = "Consulta Fiadores de la Operación"
   strSQL = "select * from fiadores where estado = 'A' and id_solicitud = " & lngOperacion
End Select

lblConsulta.Refresh
Call CargaListaDetalle
End Sub

Sub ConfiguraLista()
lswConsulta.ListItems.Clear
lswConsulta.ColumnHeaders.Clear

With lswConsulta.ColumnHeaders
 Select Case GLOBALES.gOpcion
   Case 0 'fiadores
     .Add , , "Operación"
     .Add , , "Código", 800
     .Add , , "Céd. Fiador"
     .Add , , "Nombre Fiador", 2500
   Case 1 'cuotas morosas
     .Add , , "Fec.Pro"
     .Add , , "Int.Cor", , 1
     .Add , , "Int.Mor", , 1
     .Add , , "Amortización", 1200, 1
     .Add , , "Total", 1
     .Add , , "Cuota Origen", 1200, 1
     .Add , , "Tipo Comp"
     .Add , , "Num.Comp"
   Case 2 'Refundiciones
     .Add , , "Operación"
     .Add , , "Código", 800
     .Add , , "Ab.Saldo", 1300, 1
     .Add , , "Ab.Int.Cor", 1300, 1
     .Add , , "Ab.Int.Mor", 1300, 1
     .Add , , "Ab.Fecha", 1200
   Case 3 'Desembolsos
     .Add , , "Concepto", 4000
     .Add , , "Monto", 1000, 1
     .Add , , "Cuenta", 1200
   Case 4 'Abonos
     .Add , , "Fec.Proc"
     .Add , , "Fecha"
     .Add , , "Abono", 1200, 1
     .Add , , "Int.Cor", 1200, 1
     .Add , , "Amortización", 1200, 1
     .Add , , "Tipo Comp"
     .Add , , "Num.Comp"
 End Select
End With
End Sub

Sub CargaListaDetalle()
Dim rs As New ADODB.Recordset, itmX As ListItem
On Error Resume Next
Me.MousePointer = 11

With rs
  .CursorLocation = adUseServer
  .Open strSQL, GLOBALES.gConDatos, adOpenStatic
  Select Case GLOBALES.gOpcion
    Case 0 'Fiadores
      Do While .EOF = False
       Set itmX = lswConsulta.ListItems.Add(lswConsulta.ListItems.Count + 1, , lngOperacion, , 1)
         itmX.SubItems(1) = !Codigo
         itmX.SubItems(2) = IIf(IsNull(!cedulaf), "", !cedulaf)
         itmX.SubItems(3) = IIf((IsNull(!Nombre) Or Len(Trim(!Nombre)) = 0), fxNombre(!cedulaf), !Nombre)
       .MoveNext
      Loop
    Case 1 'Cuotas Morosas
      Do While .EOF = False
       Set itmX = lswConsulta.ListItems.Add(lswConsulta.ListItems.Count + 1, , Format(!FechaP, "####-##"), , 1)
         itmX.SubItems(1) = IIf(IsNull(!intc), 0, Format(!intc, "###,###,###,##0.00"))
         itmX.SubItems(2) = IIf(IsNull(!intm), 0, Format(!intm, "###,###,###,##0.00"))
         itmX.SubItems(3) = IIf(IsNull(!Amortiza), 0, Format(!Amortiza, "###,###,###,##0.00"))
         itmX.SubItems(4) = Format(IIf(IsNull(!intc), 0, !intc) + IIf(IsNull(!intm), 0, !intm) + IIf(IsNull(!Amortiza), 0, !Amortiza), "###,###,###,##0.00")
         itmX.SubItems(5) = IIf(IsNull(!cuota_morosa), 0, Format(!cuota_morosa, "###,###,###,##0.00"))
         itmX.SubItems(6) = fxTipoComprobante(!tcon)
         itmX.SubItems(7) = IIf(IsNull(!ncon), 0, !ncon)
       .MoveNext
      Loop
    
    Case 2 'Refundiciones
      Do While .EOF = False
       Set itmX = lswConsulta.ListItems.Add(lswConsulta.ListItems.Count + 1, , !Id_solicitud, , 1)
         itmX.SubItems(1) = !Codigo
         itmX.SubItems(2) = IIf(IsNull(!Monto), 0, Format(!Monto, "###,###,###,##0.00"))
         itmX.SubItems(3) = IIf(IsNull(!IntCor), 0, Format(!IntCor, "###,###,###,##0.00"))
         itmX.SubItems(4) = IIf(IsNull(!IntMor), 0, Format(!IntMor, "###,###,###,##0.00"))
         itmX.SubItems(5) = Format(IIf(IsNull(!Fecha), Date, !Fecha), "dd/mm/yyyy")
       
       .MoveNext
      Loop
    Case 3 'Desembolsos
      Do While .EOF = False
       Set itmX = lswConsulta.ListItems.Add(lswConsulta.ListItems.Count + 1, , !concepto, , 1)
         itmX.SubItems(1) = IIf(IsNull(!Monto), 0, Format(!Monto, "###,###,###,##0.00"))
         itmX.SubItems(2) = Format(IIf(IsNull(!cuenta_conta), 0, !cuenta_conta), GLOBALES.gstrMascara)
       .MoveNext
      Loop
    Case 4 'Abonos
      Do While .EOF = False
       Set itmX = lswConsulta.ListItems.Add(lswConsulta.ListItems.Count + 1, , Format(!FechaP, "####-##"), , 1)
         itmX.SubItems(1) = Format(IIf(IsNull(!fechas), Date, !fechas), "dd/mm/yyyy")
         itmX.SubItems(2) = IIf(IsNull(!Abono), 0, Format(!Abono, "###,###,###,##0.00"))
         itmX.SubItems(3) = IIf(IsNull(!intcp), 0, Format(!intcp, "###,###,###,##0.00"))
         itmX.SubItems(4) = IIf(IsNull(!Amortiza), 0, Format(!Amortiza, "###,###,###,##0.00"))
         itmX.SubItems(5) = fxTipoComprobante(rs!tcon)
         itmX.SubItems(6) = IIf(IsNull(!ncon), 0, !ncon)
       .MoveNext
      Loop
      
  End Select
  
  .Close

End With

Me.MousePointer = 1

End Sub

