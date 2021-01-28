VERSION 5.00
Begin VB.Form frmDDFeAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DDF-e API"
   ClientHeight    =   9735
   ClientLeft      =   765
   ClientTop       =   -330
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   10845
   Begin VB.PictureBox SSTab1 
      Height          =   9735
      Left            =   0
      ScaleHeight     =   9675
      ScaleWidth      =   10755
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txtCaminhoManif 
         Height          =   345
         Left            =   -74340
         TabIndex        =   52
         Text            =   "C:\Notas\"
         Top             =   4860
         Width           =   9255
      End
      Begin VB.CheckBox chkRetornoSimples 
         Caption         =   "Retorno somente chaves baixadas"
         Height          =   255
         Left            =   3840
         TabIndex        =   51
         Top             =   2160
         Width           =   2775
      End
      Begin VB.CheckBox chkApenasPendLote 
         Caption         =   "Carregar apenas docs pendentes"
         Height          =   255
         Left            =   7320
         TabIndex        =   48
         Top             =   2160
         Width           =   2655
      End
      Begin VB.CheckBox chkComEventosLote 
         Caption         =   "Incluir XMLs dos eventos"
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CheckBox chkApenasComXmlLote 
         Caption         =   "Apenas XMLs autorizados"
         Height          =   255
         Left            =   7680
         TabIndex        =   44
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CheckBox chkIncluirPdfLote 
         Caption         =   "Incluir PDF das Notas"
         Height          =   255
         Left            =   840
         TabIndex        =   43
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtUltNSULote 
         Height          =   285
         Left            =   7200
         TabIndex        =   42
         Text            =   "0"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.ComboBox cbModeloLote 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":0000
         Left            =   6120
         List            =   "frmDDFeAPI.frx":000D
         TabIndex        =   41
         Text            =   "55"
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox cbTpAmbLote 
         Height          =   315
         Left            =   4440
         TabIndex        =   39
         Text            =   "2"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtCNPJIntLote 
         Height          =   285
         Left            =   840
         TabIndex        =   37
         Text            =   "12378462000126"
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtCaminhoLote 
         Height          =   285
         Left            =   960
         TabIndex        =   35
         Text            =   "C:\Notas\"
         Top             =   4800
         Width           =   9135
      End
      Begin VB.CommandButton btnDownloadLote 
         Caption         =   "Enviar Documento para Processamento >>>>>>"
         Height          =   375
         Left            =   360
         TabIndex        =   34
         Top             =   5520
         Width           =   10215
      End
      Begin VB.TextBox txtResultLote 
         Height          =   3015
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   6600
         Width           =   10215
      End
      Begin VB.TextBox txtCaminhoUniq 
         Height          =   285
         Left            =   -74280
         TabIndex        =   32
         Text            =   "C:\Notas\"
         Top             =   4800
         Width           =   9135
      End
      Begin VB.CheckBox chkIncluirPdfUniq 
         Caption         =   "Incluir PDF das Notas"
         Height          =   255
         Left            =   -74280
         TabIndex        =   30
         Top             =   1920
         Width           =   1935
      End
      Begin VB.CheckBox chkComEventosUniq 
         Caption         =   "Incluir XMLs dos eventos"
         Height          =   255
         Left            =   -68400
         TabIndex        =   29
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CheckBox chkApenasComXmlUniq 
         Caption         =   "Apenas XMLs autorizados"
         Height          =   255
         Left            =   -73080
         TabIndex        =   28
         Top             =   3600
         Width           =   2175
      End
      Begin VB.ComboBox cbModeloUniq 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":001D
         Left            =   -69120
         List            =   "frmDDFeAPI.frx":002A
         TabIndex        =   27
         Text            =   "55"
         Top             =   1440
         Width           =   615
      End
      Begin VB.ComboBox cbTpAmbUniq 
         Height          =   315
         Left            =   -70920
         TabIndex        =   26
         Text            =   "2"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cbTpDocUniq 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":003A
         Left            =   -67920
         List            =   "frmDDFeAPI.frx":0044
         TabIndex        =   25
         Text            =   "nsu"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtCNPJIntUniq 
         Height          =   285
         Left            =   -74280
         TabIndex        =   24
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtNumDocUniq 
         Height          =   285
         Left            =   -67920
         TabIndex        =   23
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtResultUniq 
         Height          =   3015
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   6480
         Width           =   10215
      End
      Begin VB.TextBox txtXJustManif 
         Height          =   1095
         Left            =   -67920
         TabIndex        =   16
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox cbTpDocManif 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":0054
         Left            =   -68520
         List            =   "frmDDFeAPI.frx":005E
         TabIndex        =   13
         Text            =   "nsu"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtNumDocManif 
         Height          =   375
         Left            =   -68520
         TabIndex        =   12
         Top             =   1320
         Width           =   3375
      End
      Begin VB.ComboBox cbTpAmbManif 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":006E
         Left            =   -74280
         List            =   "frmDDFeAPI.frx":0070
         TabIndex        =   11
         Text            =   "2"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.ComboBox cbTpEventoManif 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":0072
         Left            =   -74280
         List            =   "frmDDFeAPI.frx":0082
         TabIndex        =   5
         Text            =   "210200"
         Top             =   3600
         Width           =   3735
      End
      Begin VB.TextBox txtResultManif 
         Height          =   3015
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   6480
         Width           =   10215
      End
      Begin VB.CommandButton btnManif 
         Caption         =   "Enviar Documento para Processamento >>>>>>"
         Height          =   375
         Left            =   -74760
         TabIndex        =   3
         Top             =   5520
         Width           =   10215
      End
      Begin VB.CommandButton btnDownUnico 
         Caption         =   "Enviar Documento para Processamento >>>>>>"
         Height          =   375
         Left            =   -74760
         TabIndex        =   2
         Top             =   5520
         Width           =   10215
      End
      Begin VB.TextBox txtCNPJInteressadoManif 
         Height          =   405
         Left            =   -74280
         TabIndex        =   1
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label19 
         Caption         =   "Salvar em:"
         Height          =   255
         Left            =   5040
         TabIndex        =   50
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "Ultimo NSU"
         Height          =   255
         Left            =   7200
         TabIndex        =   49
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbUltNSU 
         Caption         =   "0"
         Height          =   255
         Left            =   6120
         TabIndex        =   47
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Numero do ultimo NSU Disponivel:"
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label16 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   6120
         TabIndex        =   40
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo de Ambiente"
         Height          =   255
         Left            =   4440
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "CNPJ Interessado"
         Height          =   255
         Left            =   960
         TabIndex        =   36
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Se selecionar chave ao inves de NSU, você pode utilizar as seguintes funções::"
         Height          =   255
         Left            =   -72480
         TabIndex        =   31
         Top             =   3120
         Width           =   5775
      End
      Begin VB.Label Label11 
         Caption         =   "Resposta do Servidor"
         Height          =   255
         Left            =   -74760
         TabIndex        =   22
         Top             =   6240
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Modelo"
         Height          =   255
         Left            =   -69000
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Salvar em:"
         Height          =   255
         Left            =   -70080
         TabIndex        =   20
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo de Ambiente"
         Height          =   255
         Left            =   -70920
         TabIndex        =   19
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "CNPJ Interessado"
         Height          =   255
         Left            =   -74280
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Justificação:"
         Height          =   255
         Left            =   -68880
         TabIndex        =   15
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Informar o xJust somente quando evento for 210240"
         Height          =   255
         Left            =   -68880
         TabIndex        =   14
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Salvar em:"
         Height          =   195
         Left            =   -70320
         TabIndex        =   10
         Top             =   4560
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Ambiente:"
         Height          =   195
         Left            =   -74280
         TabIndex        =   9
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo de Download:"
         Height          =   255
         Left            =   -74280
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Resposta do Servidor"
         Height          =   195
         Left            =   -74760
         TabIndex        =   7
         Top             =   6240
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ Interessado"
         Height          =   195
         Left            =   -74280
         TabIndex        =   6
         Top             =   1080
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmDDFeAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnDownloadLote_Click()
    On Error GoTo SAI
    Dim retorno As String
    
    If (txtCNPJIntLote.Text <> "") And (cbTpAmbLote.Text <> "") And (cbModeloLote.Text <> "") And (txtUltNSULote.Text <> "") And (txtCaminhoLote.Text <> "") Then
        txtResultLote.Text = ""
        retorno = downloadLote(txtCaminhoLote.Text, txtCNPJIntLote.Text, cbTpAmbLote.Text, cbModeloLote.Text, txtUltNSULote.Text, chkIncluirPdfLote.Value, chkApenasComXmlLote.Value, chkComEventosLote.Value, chkApenasPendLote, chkRetornoSimples.Value)
        txtResultLote.Text = retorno
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleDDFeAPI
End Sub

Private Sub btnDownUnico_Click()
    On Error GoTo SAI
    Dim retorno As String
    
    If (txtCNPJIntUniq.Text <> "") And (cbTpAmbUniq.Text <> "") And (cbModeloUniq.Text <> "") And (cbTpDocUniq.Text <> "") And (txtNumDocUniq.Text <> "") Then
        txtResultUniq.Text = ""
        If (cbTpDocUniq.Text = "nsu") Then
            retorno = downloadUnico(txtCaminhoUniq.Text, txtCNPJIntUniq.Text, cbTpAmbUniq.Text, cbModeloUniq.Text, txtNumDocUniq.Text, "", chkIncluirPdfUniq.Value)
        Else
            retorno = downloadUnico(txtCaminhoUniq.Text, txtCNPJIntUniq.Text, cbTpAmbUniq.Text, cbModeloUniq.Text, "", txtNumDocUniq.Text, chkIncluirPdfUniq.Value, chkApenasComXmlUniq.Value, chkComEventosUniq.Value)
        End If
        
        txtResultUniq.Text = retorno
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleDDFeAPI
End Sub

Private Sub btnManif_Click()
    On Error GoTo SAI
    Dim retorno As String
    
    If (txtCaminhoManif.Text <> "") And (txtCNPJInteressadoManif.Text <> "") And (txtNumDocManif.Text <> "") And (cbTpEventoManif.Text <> "") And (cbTpAmbManif.Text <> "") Then
        txtResultManif.Text = ""
        If (cbTpDocManif.Text = "nsu") Then
            retorno = manifestacao(txtCaminhoManif.Text, txtCNPJInteressadoManif.Text, cbTpEventoManif.Text, cbTpAmbManif, txtNumDocManif.Text, "", txtXJustManif.Text)
        Else
            retorno = manifestacao(txtCaminhoManif.Text, txtCNPJInteressadoManif.Text, cbTpEventoManif.Text, cbTpAmbManif, "", txtNumDocManif.Text, txtXJustManif.Text)
        End If
        
        txtResultManif.Text = retorno
    Else
        MsgBox ("Todos os campos devem ser preenchidos")
    End If
    
    Exit Sub
SAI:
    MsgBox ("Problemas ao Requisitar emissão ao servidor" & vbNewLine & Err.Description), vbInformation, titleDDFeAPI

End Sub

Private Sub Form_Load()
    cbTpAmbManif.List(0) = 2
    cbTpAmbManif.List(1) = 1
    cbTpAmbUniq.List(0) = 2
    cbTpAmbUniq.List(1) = 1
    cbTpAmbLote.List(0) = 2
    cbTpAmbLote.List(1) = 1
End Sub

