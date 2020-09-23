VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Original"
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   1080
      MaxLength       =   4
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "New Res Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "New Res X:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Current "
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current "
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is the change resolution box!
Private Sub Command1_Click()
'sets new resolution
With Form1
    .ImgEdit1.ImageResolutionX = txtWidth.Text
    .ImgEdit1.ImageResolutionY = txtHeight.Text
End With
Unload Me
End Sub

Private Sub Command2_Click()
'Set original resolution
With Form1
    .ImgEdit1.ImageResolutionX = Form1.OriginalResX
    .ImgEdit1.ImageResolutionY = Form1.OriginalResY
End With
Unload Me
End Sub

Private Sub Command3_Click()
'cancel
Unload Me
End Sub

Private Sub Form_Load()
Form1.Enabled = False 'makes sure this window will close so it wont float
                      'around behind things
'shows current reolution
Label1.Caption = "Current Res X = " & Form1.ImgEdit1.ImageResolutionX
Label2.Caption = "Current Rex Y = " & Form1.ImgEdit1.ImageResolutionY
End Sub

Private Sub Form_Unload(Cancel As Integer)
'turn form1 back on
Form1.Enabled = True
End Sub
