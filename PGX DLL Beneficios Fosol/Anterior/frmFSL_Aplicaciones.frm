VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Aplicaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FOSOL: Aplicaciones de FOSOL"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEstado 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   8
      ToolTipText     =   "Expediente"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtCedula 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   315
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   2
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   840
      Width           =   2070
   End
   Begin VB.TextBox txtExpediente 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   330
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   1
      ToolTipText     =   "Expediente"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   4125
   End
   Begin FPSpreadADO.fpSpread vgGrid 
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   7935
      _Version        =   524288
      _ExtentX        =   13996
      _ExtentY        =   5741
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      SpreadDesigner  =   "frmFSL_Aplicaciones.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Height          =   330
      Left            =   6720
      TabIndex        =   6
      Top             =   5520
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   960
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":0714
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":15886
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":2A9F8
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":3FB6A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":54CDC
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":555B6
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":6A728
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":7F89A
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Aplicaciones.frx":80174
            Key             =   "Print"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Estado :"
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
      Left            =   4320
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicación Montos FOSOL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   240
      Width           =   3135
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   240
      Picture         =   "frmFSL_Aplicaciones.frx":80A4E
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Asociado"
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
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Expediente"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7920
      Y1              =   1680
      Y2              =   1680
   End
End
Attribute VB_Name = "frmFSL_Aplicaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vEstado As String, vFecha As String, vProceso As Integer
Dim vAplicado As String, vTotalMonto As Currency

Private Sub Form_Activate()
 vModulo = 22
End Sub

Private Sub Form_Load()

 vModulo = 22
 txtCedula.Text = GLOBALES.gTag
 txtNombre.Text = GLOBALES.gTag2
 txtExpediente.Text = GLOBALES.gTag3
 vFecha = fxFechaServidor
 vProceso = Year(vFecha) & Format(Month(vFecha))
 
 Call sbTraeDatosAprobacion
 tlbMenu.Enabled = True
 
End Sub

'Aplica el monto calculado para el FOSOL
Private Sub sbAplicaMonto()
Dim i As Integer, vOperacion As Long, vAbono As Currency
Dim vSobrante As Currency, vTipoDoc As String, vNumDoc As String
Dim CuentaContable As String
  
On Error GoTo vError
  
'Trae la cuenta contable
strSQL = "select valor from Fsl_parametros where cod_parametro = '01'"
rs.Open strSQL, glogon.Conection, adOpenStatic
 
 CuentaContable = rs!Valor

rs.Close
  
vTipoDoc = "NC"
vNumDoc = fxDocumentoConsecutivo(vTipoDoc)
 
With vgGrid
    For i = 1 To .MaxRows - 1
     .Row = i
     .Col = 1
       
     If .Text <> "" Then
       
       vOperacion = .Text
       
       .Col = 4
       vAbono = .Text
       
       
       If vAbono > 0 Then
       
        vTotalMonto = vTotalMonto + CCur(vAbono)
              
         'Ejecuta aplicacion desde el sp si el monto a aplicar es mayor que 0
         strSQL = "exec spFSL_AplicaMonto " & vOperacion & " ," & vAbono & "" _
                & ",'" & glogon.Usuario & "','" & vTipoDoc & "','" & vNumDoc & "','" & vProceso & "','" & txtExpediente.Text & "'"
         glogon.Conection.Execute strSQL

       End If
       
     End If
       
    Next i
End With
        
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
         & " values('" & vNumDoc & "','" & vTipoDoc & "',getdate(),'" & glogon.Usuario & "','" _
         & txtCedula.Text & "','" & fxNombre(txtCedula.Text) & "','AFL004'," & vTotalMonto & ",'P','" _
         & txtCedula.Text & "','" & txtExpediente.Text & "','','" & GLOBALES.gOficinaTitular & "', " _
         & "'Aplicación Fosol','Exp.: ' + '" & txtExpediente.Text & "','','','" _
         & "','','','','','','Aplicación de Fondo Solidario','" & vNumDoc & "')"
glogon.Conection.Execute strSQL
        
            
'Asiento General de la Aplicación
strSQL = "exec spCrdMovAsientoCredito '" & vTipoDoc & "','" & vNumDoc & "','" & CuentaContable & "'" _
       & ",'" & vOperacion & "','" & txtExpediente.Text & "',''"
glogon.Conection.Execute strSQL
  
Call sbImprimeRecibo(vNumDoc, vTipoDoc)
  
'Actualiza estado expediente
strSQL = " update FSL_EXPEDIENTES set APLICADO = 'S' " _
       & " where COD_EXPEDIENTE = " & txtExpediente.Text & " "
glogon.Conection.Execute strSQL

'Refresca grid
Call sbTraeDatosAprobacion

  
'       strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
'              & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
'
'       glogon.Conection.Execute strSQL

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
Resume
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
  If txtEstado.Text = "APLICADO" Then Exit Sub
  Call sbAplicaMonto
End Sub

Private Sub sbTraeDatosAprobacion()
Dim i As Integer, vDisponible As Currency, vSobrante As Currency, vBase As Currency, vTSobrante As Currency

vTSobrante = 0

With vgGrid

    If txtExpediente.Text = Empty Then Exit Sub
       strSQL = "Select ED.COD_EXPEDIENTE, ED.PRIMERA_DEDUCCION,ED.ID_SOLICITUD " _
           & ", ED.TOTAL_DEUDA_P, ED.MONTO_LIQUIDACION, ED.PORCENTAJE, ED.MONTO_FOSOL" _
           & ", ED.TIPO_APLICACION_FOSOL, ED.MONTO_FORMALIZADO, E.ESTADO,Reg.SALDO as 'SaldoActual',E.APLICADO" _
           & ", isnull(Vm.IntC + Vm.IntM + Vm.Cargos + Vm.Poliza,0) + Reg.Saldo as 'TDeudaActual'" _
           & " from FSL_EXPEDIENTES E inner join FSL_EXPEDIENTES_DETALLE ED on ED.COD_EXPEDIENTE = E.COD_EXPEDIENTE" _
           & "  inner join reg_creditos Reg on Ed.id_Solicitud = Reg.Id_Solicitud" _
           & "  left join vista_morosidad Vm on Reg.id_solicitud = Vm.id_Solicitud" _
           & " where E.COD_EXPEDIENTE = '" & txtExpediente.Text & "' "
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    vEstado = rs!Estado
    vAplicado = rs!APLICADO
    
    If vAplicado = "S" Then
      txtEstado.Text = "APLICADO"
      tlbMenu.Enabled = False
    Else
        Select Case vEstado
           Case "REC"
             txtEstado.Text = "RECHAZADO"
             
           Case "APR"
              txtEstado.Text = "APROBADO"
        
           Case "PEN"
              txtEstado.Text = "PENDIENTE"
                 
           Case "APL"
              txtEstado.Text = "APELACION"
        
        End Select
    End If
    
    .MaxRows = 1
    
    Do While Not rs.EOF
      .Row = .MaxRows
            
      If rs!Tipo_Aplicacion_Fosol = "M" Then
         vBase = rs!MONTO_FORMALIZADO
      Else
         vBase = rs!TOTAL_DEUDA_P
      End If
      
      vDisponible = vBase * rs!Porcentaje / 100
      vSobrante = rs!TDeudaActual - vDisponible
      vTSobrante = vTSobrante + vSobrante
      
      .Col = 1 'Operacion
      .Text = CStr(rs!ID_SOLICITUD)
                      
      .Col = 2 'Total Deuda en el Momento de la Presentacion
      .Text = Format(rs!TOTAL_DEUDA_P, "Standard")
      
      .Col = 3  'Saldo al día
      .Text = Format(rs!TDeudaActual, "Standard")
      
      .Col = 4  'Disponible
      .Text = Format(vDisponible, "Standard")
           
      .Col = 5  'Sobrante
      .Text = Format(vSobrante, "Standard")
       
     .MaxRows = .MaxRows + 1
     rs.MoveNext
    Loop
    rs.Close
    
End With
  
End Sub


