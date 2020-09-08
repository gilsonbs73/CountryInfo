VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form CountryInfoWithISOCodeA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Country Information With ISOCode equal A"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFiltro 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1860
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "A"
      Top             =   900
      Width           =   645
   End
   Begin RichTextLib.RichTextBox rText 
      Height          =   7125
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   12568
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"CountryInfoWithISOCodeA.frx":0000
   End
   Begin VB.CommandButton cmdSalvarDados 
      Caption         =   "Salvar Dados"
      Height          =   825
      Left            =   9120
      TabIndex        =   2
      Top             =   540
      Width           =   2025
   End
   Begin VB.CommandButton cmdBaixarDados 
      Caption         =   "Baixar Dados"
      Height          =   825
      Left            =   4710
      TabIndex        =   0
      Top             =   540
      Width           =   2025
   End
   Begin VB.Label Label1 
      Caption         =   "Filtrar Paises que começa com a Letra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      TabIndex        =   4
      Top             =   930
      Width           =   1785
   End
   Begin VB.Label lblConnectionState 
      Caption         =   "-"
      Height          =   375
      Left            =   -60
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "CountryInfoWithISOCodeA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Reference to:
'
'   Microsoft WinHTTP Services, version 5.1

Private Req As WinHttp.WinHttpRequest
Private strXML As String
Private countryControl As CountryInfoControl.CountryInfoISOCode

Private Sub cmdBaixarDados_Click()

    Set countryControl = New CountryInfoControl.CountryInfoISOCode

    'rText.Text = countryControl.GetCountryIsoCodeXML("", txtFiltro.Text)
    rText.Text = countryControl.GetCountryIsoCode("", txtFiltro.Text)
    
End Sub

Private Sub cmdSalvarDados_Click()
    
    countryControl.SaveFilteredData

End Sub
Private Sub Form_Load()
        Set Req = New WinHttp.WinHttpRequest
End Sub

Private Sub txtFiltro_Change()
    If IsNumeric(txtFiltro.Text) Then
        MsgBox "Informar apenas letra para filtrar o Pais !!", vbOKOnly + vbInformation, "Aviso"
    End If
    txtFiltro.Text = UCase(txtFiltro.Text)
    CountryInfoWithISOCodeA.Caption = "Country Information With ISOCode equal " + txtFiltro.Text
End Sub
