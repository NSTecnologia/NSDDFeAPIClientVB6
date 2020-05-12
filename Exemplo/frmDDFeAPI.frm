VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmDDFeAPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DDF-e API"
   ClientHeight    =   9735
   ClientLeft      =   6810
   ClientTop       =   990
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   10845
   Begin TabDlg.SSTab SSTab1 
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   17171
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Manisfestação"
      TabPicture(0)   =   "frmDDFeAPI.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCNPJInteressadoManif"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "btnManif"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtResultManif"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cbTpEventoManif"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbTpAmbManif"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNumDocManif"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbTpDocManif"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtXJustManif"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCaminhoManif"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Download Único"
      TabPicture(1)   =   "frmDDFeAPI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label7"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label9"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label12"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "btnDownUnico"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txtResultUniq"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtNumDocUniq"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txtCNPJIntUniq"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cbTpDocUniq"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cbTpAmbUniq"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cbModeloUniq"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chkApenasComXmlUniq"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkComEventosUniq"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "chkIncluirPdfUniq"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtCaminhoUniq"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Download Lote"
      TabPicture(2)   =   "frmDDFeAPI.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label14"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label16"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label17"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lbUltNSU"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label18"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label19"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtResultLote"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "btnDownloadLote"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtCaminhoLote"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtCNPJIntLote"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cbTpAmbLote"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cbModeloLote"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "txtUltNSULote"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "chkIncluirPdfLote"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "chkApenasComXmlLote"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "chkComEventosLote"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "chkApenasPendLote"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "chkRetornoSimples"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).ControlCount=   19
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
         ItemData        =   "frmDDFeAPI.frx":0054
         Left            =   6120
         List            =   "frmDDFeAPI.frx":0061
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
         ItemData        =   "frmDDFeAPI.frx":0071
         Left            =   -69120
         List            =   "frmDDFeAPI.frx":007E
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
         ItemData        =   "frmDDFeAPI.frx":008E
         Left            =   -67920
         List            =   "frmDDFeAPI.frx":0098
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
         ItemData        =   "frmDDFeAPI.frx":00A8
         Left            =   -68520
         List            =   "frmDDFeAPI.frx":00B2
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
         ItemData        =   "frmDDFeAPI.frx":00C2
         Left            =   -74280
         List            =   "frmDDFeAPI.frx":00C4
         TabIndex        =   11
         Text            =   "2"
         Top             =   2520
         Width           =   3735
      End
      Begin VB.ComboBox cbTpEventoManif 
         Height          =   315
         ItemData        =   "frmDDFeAPI.frx":00C6
         Left            =   -74280
         List            =   "frmDDFeAPI.frx":00D6
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
         Caption         =   "Numero do ultimo NSU baixado:"
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

