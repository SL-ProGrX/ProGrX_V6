VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFNDCorreccionesBas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Basurero Para Correcciones"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optX 
      Caption         =   "Actualiza Mov. Excedentes (Capexc)"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   5295
   End
   Begin VB.OptionButton optX 
      Caption         =   "Actualiza Mov. Incon. Planillas (IncoASE)"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   5295
   End
   Begin VB.ComboBox cbo 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtContrato 
      Height          =   315
      Left            =   2280
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.OptionButton optX 
      Caption         =   "Identifica Rendimientos x Historicos"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   5295
   End
   Begin VB.OptionButton optX 
      Caption         =   "Actualiza Movimientos detalle de Planillas"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   5295
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.OptionButton optX 
      Caption         =   "Actualiza Movimientos detalle de RE, ND, NC"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   5295
   End
   Begin VB.OptionButton optX 
      Caption         =   "Actualiza Movimientos detalle de las liquidaciones "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5295
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   0
      Top             =   5040
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   7560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      Caption         =   "Contrato a Revisar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Plan a Revisar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   240
      X2              =   7680
      Y1              =   1080
      Y2              =   1080
   End
End
Attribute VB_Name = "frmFNDCorreccionesBas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub cmdAplicar_Click()

Select Case True
 Case optX.Item(0) 'Mov. Liquidacion
   Call sbActualizaMovLiquidaciones
 Case optX.Item(1) 'Mov. Documentos
   Call sbActualizaMovDoc
 Case optX.Item(2) 'Mov. Planillas
   Call sbActualizaMovPlanilla
   
   
 Case optX.Item(3) 'Mov. Rendimientos
   
 Case optX.Item(4) 'Mov. Planillas IncoASE
    Call sbActualizaMovPlanillaFondo
 Case optX.Item(5) 'Mov. Excedentes Capexc
    Call sbActualizaMovExc
End Select

End Sub

Private Sub sbActualizaMovLiquidaciones()

'Revisa las Liquidaciones / Valida los contratos

strSQL = "select * from fnd_liquidacion"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   strSQL = "update fnd_contratos_detalle set cod_contrato = " & rs!cod_contrato _
        & " where tcon = 5 and ncon = " & rs!consec & " and cod_contrato <> " & rs!cod_contrato
   glogon.Conection.Execute strSQL
   rs.MoveNext
   prgBar.Value = prgBar.Value + 1
Loop
rs.Close

MsgBox "Fin...", vbInformation

End Sub



Private Sub sbActualizaMovDoc()
Dim vTcon As Long, rsTmp As New ADODB.Recordset
Dim vContrato As Long, vConcepto As String
Dim i As Integer, vResultado As String

'Revisa las Liquidaciones / Valida los contratos

strSQL = "select * from fnd_documentos where tipo in('RE','NC','ND')"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   Select Case rs!Tipo
     Case "RE"
       vTcon = 2
     Case "ND"
       vTcon = 8
     Case "NC"
       vTcon = 7
   End Select

   vConcepto = Right(rs!concepto, 5)
   vResultado = ""
   
   For i = 1 To Len(vConcepto)
     If IsNumeric(Mid(vConcepto, i, 1)) Then
       vResultado = vResultado & Mid(vConcepto, i, 1)
     End If
   Next i
   
   If IsNumeric(vResultado) Then
       vContrato = vResultado
    
       strSQL = "update fnd_contratos_detalle set cod_contrato = " & vContrato _
            & " where tcon = " & vTcon & " and ncon = " & rs!id_documento & " and cod_contrato <> " & vContrato
       glogon.Conection.Execute strSQL

   End If
   
   rs.MoveNext
   prgBar.Value = prgBar.Value + 1
Loop
rs.Close

MsgBox "Fin...", vbInformation

End Sub


Private Sub sbActualizaMovPlanilla()
Dim vTcon As Long, rsTmp As New ADODB.Recordset, rsSet As New ADODB.Recordset
Dim vContrato As Long, vConcepto As String
Dim i As Integer, vResultado As String


If cbo.Text = "" Or txtContrato.Text = "" Or Not IsNumeric(txtContrato.Text) Then
   MsgBox "Especifique el Plan y Contrato Duplicado"
   Exit Sub
End If

'Revisa las Liquidaciones / Valida los contratos

strSQL = "select C.cod_plan,cod_contrato,cedula,fecha_inicio,operacion,liq_fecha,P.codigo_ase " _
       & " from fnd_contratos C inner join fnd_planes P on C.cod_plan = P.cod_plan and P.codigo_ase <> 'FGEN'" _
       & " where C.operacion is not null and C.cod_plan = '" & cbo.Text & "' and C.cod_contrato <> " & txtContrato.Text
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   vTcon = 1
   
   'Pregunta si existe el mov. en el expediente
   strSQL = "select * from creditos_dt where id_solicitud = " & rs!Operacion _
          & " and tcon = 1 and fechas not in(select fecha" _
          & " from fnd_contratos_detalle where cod_plan = '" & rs!cod_plan & "' and cod_contrato = " _
          & rs!cod_contrato & "  and tcon = 1)"

   rsTmp.Open strSQL, glogon.Conection, adOpenStatic
   Do While Not rsTmp.EOF
     'Buscan en el otro contrato
     strSQL = "select Top 1 cod_fnd_detalle from fnd_contratos_detalle where cod_plan = '" & cbo.Text & "' and cod_contrato = " & txtContrato.Text _
            & " and tcon = 1 and monto = " & rsTmp!abono & " and fecha = '" & Format(rsTmp!fechas, "yyyy/mm/dd") & "'"
     

     rsSet.Open strSQL, glogon.Conection, adOpenStatic
     If Not rsSet.EOF And Not rsSet.BOF Then
     
            strSQL = "update fnd_contratos_detalle set cod_contrato = " & rs!cod_contrato _
                 & " where cod_fnd_detalle = " & rsSet!cod_fnd_detalle
            glogon.Conection.Execute strSQL
     
     End If
     rsSet.Close
      
     rsTmp.MoveNext
   Loop
   rsTmp.Close

  
   rs.MoveNext
   prgBar.Value = prgBar.Value + 1
Loop
rs.Close

MsgBox "Fin...", vbInformation


End Sub


Private Sub sbActualizaMovPlanillaFondo()
Dim vTcon As Long, rsTmp As New ADODB.Recordset
Dim vContrato As Long, vConcepto As String
Dim i As Integer, vResultado As String


If UCase(cbo.Text) <> UCase("Incoase") Or txtContrato.Text = "" Or Not IsNumeric(txtContrato.Text) Then
   MsgBox "Especifique el Plan INCOASE y Contrato Duplicado"
   Exit Sub
End If

'Revisa el historico de transacciones de fondos x inconsistencias de planilla / Valida los contratos

strSQL = "select * From prm_fondo where cod_plan = 'IncoASE'"
rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   
   strSQL = "select count(*) as existe from fnd_contratos where cedula = '" & rs!Cedula & "' and cod_plan = '" & cbo.Text & "'"
   rsTmp.Open strSQL, glogon.Conection, adOpenStatic
   If rsTmp.EOF And rsTmp.BOF Then
      vContrato = 0
   Else
      If rsTmp!existe = 1 Then
        'Si tiene un contrato
         rsTmp.Close
         strSQL = "select cod_contrato from fnd_contratos where cedula = '" & rs!Cedula & "' and cod_plan = '" & cbo.Text & "'"
         rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         vContrato = rsTmp!cod_contrato
      Else
         'Si tiene mas de un contrato
         rsTmp.Close
         strSQL = "select cod_contrato from fnd_contratos where cedula = '" & rs!Cedula & "' and cod_plan = '" & cbo.Text & "'" _
                & " and '" & Format(rs!Fecha, "yyyy/mm/dd") & "' between dateadd(d, -1,fecha_inicio) and isnull(liq_fecha,getdate())"
         rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         If rsTmp.EOF And rsTmp.BOF Then
             vContrato = 0
         Else
             vContrato = rsTmp!cod_contrato
         End If
      End If
   End If
   
   rsTmp.Close
   
    'Revisa si tiene el mov. registrado
    
     strSQL = "select count(*) as Existe from fnd_contratos_detalle where cod_plan = '" & cbo.Text & "' and cod_contrato = " & vContrato _
            & " and tcon = 1 and monto = " & rs!Monto & " and fecha between '" & Format(rs!Fecha, "yyyy/mm/dd") & " 00:00:00' and '" _
            & Format(rs!Fecha, "yyyy/mm/dd") & " 23:59:59'"
     rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         If rsTmp!existe >= 1 Then
           vContrato = 0
         End If
     rsTmp.Close
    
    If vContrato > 0 Then
    
            'Localiza el movimiento en el contrato a revisar
             strSQL = "select Top 1 cod_fnd_detalle from fnd_contratos_detalle where cod_plan = '" & cbo.Text & "' and cod_contrato = " & txtContrato.Text _
                    & " and tcon = 1 and monto = " & rs!Monto & " and fecha between '" & Format(rs!Fecha, "yyyy/mm/dd") & " 00:00:00' and '" _
                    & Format(rs!Fecha, "yyyy/mm/dd") & " 23:59:59'"
             rsTmp.Open strSQL, glogon.Conection, adOpenStatic
             If Not rsTmp.EOF And Not rsTmp.BOF Then
             
                    strSQL = "update fnd_contratos_detalle set cod_contrato = " & vContrato _
                         & " where cod_fnd_detalle = " & rsTmp!cod_fnd_detalle
                    glogon.Conection.Execute strSQL
             
             End If
             rsTmp.Close
    End If ' vContrato
    
   rs.MoveNext
   prgBar.Value = prgBar.Value + 1
Loop
rs.Close

MsgBox "Fin...", vbInformation


End Sub


Private Sub sbActualizaMovExc()
Dim vTcon As Long, rsTmp As New ADODB.Recordset
Dim vContrato As Long, vConcepto As String
Dim i As Integer, vResultado As String


If UCase(cbo.Text) <> UCase("capexc") Or txtContrato.Text = "" Or Not IsNumeric(txtContrato.Text) Then
   MsgBox "Especifique el Plan CAPEXC y Contrato Duplicado"
   Exit Sub
End If

'Revisa el historico Excedentes (Capitalizaciones) / Valida los contratos

strSQL = "select cedula,capitalizado_individual as Monto, 'CAPEXC' as Cod_Plan From exc_cierre" _
       & " Where capitalizado_individual > 0"

rs.Open strSQL, glogon.Conection, adOpenStatic

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1

Do While Not rs.EOF
   
   strSQL = "select count(*) as existe from fnd_contratos where cedula = '" & rs!Cedula & "' and cod_plan = '" & cbo.Text & "'"
   rsTmp.Open strSQL, glogon.Conection, adOpenStatic
   If rsTmp.EOF And rsTmp.BOF Then
      vContrato = 0
   Else
      If rsTmp!existe = 1 Then
        'Si tiene un contrato
         rsTmp.Close
         strSQL = "select cod_contrato from fnd_contratos where cedula = '" & rs!Cedula & "' and cod_plan = '" & cbo.Text & "'"
         rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         vContrato = rsTmp!cod_contrato
      Else
         'Si tiene mas de un contrato
'         rsTmp.Close
'         strSQL = "select cod_contrato from fnd_contratos where cedula = '" & rs!Cedula & "' and cod_plan = '" & cbo.Text & "'" _
'                & " and '" & Format(rs!Fecha, "yyyy/mm/dd") & "' between dateadd(d, -1,fecha_inicio) and isnull(liq_fecha,getdate())"
'         rsTmp.Open strSQL, glogon.Conection, adOpenStatic
'         If rsTmp.EOF And rsTmp.BOF Then
             vContrato = 0
'         Else
'             vContrato = rsTmp!cod_contrato
'         End If
      End If
   End If
   
   rsTmp.Close
   
    'Revisa si tiene el mov. registrado
    
     strSQL = "select count(*) as Existe from fnd_contratos_detalle where cod_plan = '" & cbo.Text & "' and cod_contrato = " & vContrato _
            & " and tcon = 1 and monto = " & rs!Monto
     rsTmp.Open strSQL, glogon.Conection, adOpenStatic
         If rsTmp!existe >= 1 Then
           vContrato = 0
         End If
     rsTmp.Close
    
    If vContrato > 0 Then
    
            'Localiza el movimiento en el contrato a revisar
             strSQL = "select Top 1 cod_fnd_detalle from fnd_contratos_detalle where cod_plan = '" & cbo.Text & "' and cod_contrato = " & txtContrato.Text _
                    & " and tcon = 1 and monto = " & rs!Monto
             rsTmp.Open strSQL, glogon.Conection, adOpenStatic
             If Not rsTmp.EOF And Not rsTmp.BOF Then
             
                    strSQL = "update fnd_contratos_detalle set cod_contrato = " & vContrato _
                         & " where cod_fnd_detalle = " & rsTmp!cod_fnd_detalle
                    glogon.Conection.Execute strSQL
             
             End If
             rsTmp.Close
    End If ' vContrato
    
   rs.MoveNext
   prgBar.Value = prgBar.Value + 1
Loop
rs.Close

MsgBox "Fin...", vbInformation


End Sub




Private Sub Form_Load()

strSQL = "select cod_plan from fnd_planes"
cbo.Clear
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 cbo.AddItem Trim(rs!cod_plan)
 rs.MoveNext
Loop
rs.Close

'select cod_plan,cod_contrato,count(*)
'From fnd_contratos_detalle
'where cod_plan = 'incoase'
'group by cod_plan,cod_contrato
'having count(*) > 10
'
'CAPEXC -> 286
'ANAV   ->4186
'ANAV 05-> 1635
'incoase = 172 y 595
'
'
'select * from fnd_documentos
End Sub


