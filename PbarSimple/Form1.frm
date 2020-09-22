VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBProgressBar 1.0"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   2280
      Top             =   1080
   End
   Begin VB.PictureBox picProgress 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   98
      ScaleHeight     =   390
      ScaleWidth      =   5985
      TabIndex        =   0
      Top             =   240
      Width           =   6045
      Begin VB.PictureBox picProgressSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         ScaleHeight     =   345
         ScaleWidth      =   2040
         TabIndex        =   1
         Top             =   0
         Width           =   2040
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "PLeaSe VoTe FoR Me! - Visit my site: http://members.xoom.it/vbwork/index.html"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "VBProgressBar 1.0 - by d@w conti"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Contatore As Integer

Public Sub StartAndStopProgress()
    '<< Proprietà della ProgressBar, colore, font, ecc..
    With picProgress
        .Cls
        '<< Colore del testo
        .ForeColor = vbHighlight
        '<< Colore della ProgressBar
        .BackColor = vbRed
        .Visible = False
        '<< Impostazioni font del testo all'interno della ProgressBar
        .FontName = "Tahoma"
        .FontSize = 8
        .FontBold = True
    End With
    
    With picProgressSlide
        .Cls
        '<< Colore della ProgressBarSlide che avanza all'interno della ProgressBar
        .BackColor = vbGreen
        '<< Colore del testo dopo il superamento
        .ForeColor = vbWhite
        .Move 0, 0, 1, picProgress.ScaleHeight
        .Visible = False
        '<< Impostazioni font del testo all'interno della ProgressBar
        .FontName = "Tahoma"
        .FontSize = 8
        .FontBold = True
    End With
End Sub

Public Sub PercentualeAvanzamento(ByVal Percentage As Single)

Dim lTextTop As Long, lTextLeft As Long
    '<< Visualizza la ProgressBar
    picProgress.Visible = True
    picProgressSlide.Visible = True
    '<< Formatta i valori precedenti
    picProgress.Cls
    picProgressSlide.Cls
    '<< Calcolo della percentuale di avanzamento della ProgressBar
    picProgressSlide.Width = picProgress.ScaleWidth * Percentage
    '<< Impostazioni della posizione del testo all'interno della ProgressBar
    lTextTop = (picProgress.ScaleHeight - picProgress.TextHeight(Percentage * 100 & " %")) / 2
    lTextLeft = (picProgress.ScaleWidth - picProgress.TextWidth(Percentage * 100 & " %")) / 2
    picProgress.CurrentX = lTextLeft
    picProgress.CurrentY = lTextTop
    picProgressSlide.CurrentX = lTextLeft
    picProgressSlide.CurrentY = lTextTop
    '<< Scrive il valore in percentuale nel testo
    picProgress.Print Percentage * 100 & " %"
    picProgressSlide.Print Percentage * 100 & " %"
    '<< Aggiorna le ProgressBar
    picProgress.Refresh
    picProgressSlide.Refresh
End Sub

Private Sub Form_Load()
    StartAndStopProgress
    tmrUpdate.Enabled = True
    Me.Show
End Sub

Private Sub tmrUpdate_Timer()
    Randomize
    tmrUpdate.Interval = 200 '<< Intervallo di Tempo
    '<< Contatore di avanzamento della progressBar
    Contatore = Contatore + 1
    picProgressSlide.BackColor = vbGreen
    
    If Contatore = 100 Then
        StartAndStopProgress
        tmrUpdate.Enabled = False
        Contatore = 0
        End
        Exit Sub
    End If
    PercentualeAvanzamento (Contatore / 100)
End Sub

