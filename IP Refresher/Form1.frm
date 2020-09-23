VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shannons IP Refresher"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReNu 
      Caption         =   "Re&Nu IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdRelease 
      Caption         =   "&Release IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Only 4 http://www.pscode.com
'by Shannon
'My first attempt at this I know its chitty, but sorry
'i'm sooo very nu at this.. I hate being a nubie
'it took me like 30 min just 2 get this thingy working
'please guyz don't b 2 harsh on me ;)
'Shannon Loves all u guyz
Private Sub cmdRelease_Click()
'Loads up ipconfig.exe in DOS mode and then releases
'ur ip addy..
'vbhide hides the dos prompt thing (eyez know u all know that)
Shell ("C:\windows\system32\ipconfig.exe /release"), vbHide
Label1.Caption = "ur IP Addy is now Released"
'Displays ur IP Addy is now Released in the label caption
End Sub

Private Sub cmdReNu_Click()
'Loads up ipconfig.exe in DOS mode and then renu's
'ur ip addy.. I know..lol.. Renewed not re-nude ;)
'vbhide hides the dos prompt thing (eyez know u all know that)
Shell ("C:\windows\system32\ipconfig.exe /renew"), vbHide
Label1.Caption = "ur IP Addy is now ReNued"
'Displays ur IP Addy is now Released in the label caption
End Sub

Private Sub Form_Terminate()
End
'Once program is closed out, it leaves your memory
'I think thats right
End Sub
