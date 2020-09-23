VERSION 5.00
Object = "{6D940288-9F11-11CE-83FD-02608C3EC08A}#2.1#0"; "IMGEDIT.OCX"
Object = "{009541A3-3B81-101C-92F3-040224009C02}#2.0#0"; "IMGADMIN.OCX"
Object = "{84926CA3-2941-101C-816F-0E6013114B7F}#1.0#0"; "IMGSCAN.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8955
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8955
   ScaleWidth      =   9315
   StartUpPosition =   3  'Windows Default
   Begin ScanLibCtl.ImgScan ImgScan1 
      Left            =   7560
      Top             =   5880
      _Version        =   65536
      _ExtentX        =   1085
      _ExtentY        =   873
      _StockProps     =   0
      DestImageControl=   "ImgEdit1"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5880
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8400
      Top             =   5880
   End
   Begin ImgeditLibCtl.ImgEdit ImgEdit1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _Version        =   131073
      _ExtentX        =   13996
      _ExtentY        =   5530
      _StockProps     =   96
      BorderStyle     =   1
      ImageControl    =   "ImgEdit1"
      AnnotationBackColor=   0
      BeginProperty AnnotationFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRefresh     =   -1  'True
      UndoBufferSize  =   80909568
      OcrZoneVisibility=   -4044
      AnnotationOcrType=   127
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Quick Image"
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000C000&
         Height          =   255
         Left            =   0
         OLEDropMode     =   1  'Manual
         TabIndex        =   1
         Top             =   0
         Width           =   960
      End
   End
   Begin AdminLibCtl.ImgAdmin ImgAdmin1 
      Left            =   6600
      Top             =   5880
      _Version        =   131072
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
      PrintStartPage  =   0
      PrintEndPage    =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnubrk4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileScan 
         Caption         =   "Scan"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnubrk1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnubrk7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy Image"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut Image"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnubrk6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMouse 
         Caption         =   "Mouse Position"
         Begin VB.Menu mnuEditMouseOn 
            Caption         =   "On"
         End
         Begin VB.Menu mnuEditMouseOff 
            Caption         =   "Off"
         End
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "&Image"
      Begin VB.Menu mnuImageRes 
         Caption         =   "Change Resolution"
      End
      Begin VB.Menu mnuImageRotate 
         Caption         =   "Rotate Image"
         Begin VB.Menu mnuImageRotate90cw 
            Caption         =   "90 Clockwise"
         End
         Begin VB.Menu mnuImageRotate90ccw 
            Caption         =   "90 Counter Clockwise"
         End
         Begin VB.Menu mnuImageRotate180 
            Caption         =   "180"
         End
      End
      Begin VB.Menu mnubrk3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageZoom 
         Caption         =   "Zoom"
         Begin VB.Menu mnuImageZoomOriginal 
            Caption         =   "Original"
         End
         Begin VB.Menu mnuImageZoomTwenty 
            Caption         =   "1/20"
         End
         Begin VB.Menu mnuImageZoomTenth 
            Caption         =   "1/10"
         End
         Begin VB.Menu mnuImageZoomFifth 
            Caption         =   "1/5"
         End
         Begin VB.Menu mnuImageZoomQuartar 
            Caption         =   "1/4"
         End
         Begin VB.Menu mnuImageZoomHalf 
            Caption         =   "1/2"
         End
         Begin VB.Menu mnuImageZoomX2 
            Caption         =   "X2"
         End
         Begin VB.Menu mnuImageZoomX4 
            Caption         =   "X4"
         End
         Begin VB.Menu mnuImageZoomX5 
            Caption         =   "X5"
         End
         Begin VB.Menu mnuImageZoomX10 
            Caption         =   "X10"
         End
         Begin VB.Menu mnuImageZoomX20 
            Caption         =   "X20"
         End
      End
      Begin VB.Menu mnuImageZoomtoFit 
         Caption         =   "Fit to screen"
      End
      Begin VB.Menu mnuImageZoomtoSel 
         Caption         =   "Zoom to Selection"
      End
      Begin VB.Menu mnubrk2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageColor 
         Caption         =   "Color"
         Begin VB.Menu mnuImageColorRGB 
            Caption         =   "Grey Scale"
         End
         Begin VB.Menu mnuImageColorGrey 
            Caption         =   "RGB Scale"
         End
      End
      Begin VB.Menu mnubrk5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImageText 
         Caption         =   "Add Text"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MPOS As Boolean 'shows and hides mouse position label1
Public colour As Boolean 'set at open
Public OriginalResX As Integer 'set at open
Public OriginalResY As Integer 'set at open
Public MX As Integer 'Mouse Position X
Public MY As Integer 'Mouse position Y
Public FileOpen As Boolean 'set at open and close
Const ErrCancel = 32755 'Cancel is pressed
Const AnnoStraightLine = 4 'stuff for ImgEdit1
Const AnnoText = 7 'stuff for ImgEdit1
Const BestFit = 0 ''stuff for ImgEdit1 zoom
Const Tiff = 1 'for save
Const Awd = 2 'for save
Const Bmp = 3 'for save

Private Sub Form_Load()
Form1.Top = 0 'keeps form at top of screen
'Turns off menu command so there will be no errors untill a file is opened
mnuEdit.Enabled = False
mnuImage.Enabled = False
mnuFileSave.Enabled = False
mnuFileClose.Enabled = False
mnuFilePrint.Enabled = False
mnuEditMouseOn.Checked = True 'sets mouse position on
mnuEditMouseOff.Checked = False
FileOpen = False 'no file open yet
'sets position of label1
With Label1
    .FontBold = True
    .Top = 10
    .Left = 10
End With
End Sub

Private Sub Form_Resize()
'keeps the ImgEdit1 control same size as the form window
With ImgEdit1
    .Left = Form1.ScaleLeft
    .Top = Form1.ScaleTop
    .Width = Form1.ScaleWidth
    .Height = Form1.ScaleHeight
End With
End Sub

Private Sub mnuEditCopy_Click()
'copies image to clipboard
ImgEdit1.ClipboardCopy 0, 0, ImgAdmin1.ImageWidth, ImgAdmin1.ImageHeight
End Sub

Private Sub mnuEditCut_Click()
'cuts image to clipboard
ImgEdit1.ClipboardCut 0, 0, ImgAdmin1.ImageWidth, ImgAdmin1.ImageHeight
End Sub

Private Sub mnuEditMouseOff_Click()
'Turns Mouse Pos label1 off
mnuEditMouseOff.Checked = True
mnuEditMouseOn.Checked = False
MPOS = False
Label1.Visible = False
End Sub

Private Sub mnuEditMouseOn_Click()
'sets the mouse position label1 to on
mnuEditMouseOn.Checked = True
mnuEditMouseOff.Checked = False
MPOS = True
Label1.Visible = True
End Sub

Private Sub mnuEditPaste_Click()
'pastes clipboard to ImgEdit1
ImgEdit1.ClipboardPaste 0, 0
End Sub

Private Sub mnuFileClose_Click()
'Closes file
ImgEdit1.ClearDisplay
FileOpen = False
End Sub

Private Sub mnuFileExit_Click()
'exit
Unload Me
End Sub

Private Sub mnuFileOpen_Click()

On Error Resume Next 'error handling

Dim Image As String 'image name and path

ImgAdmin1.ShowFileDialog 0, Form1.hWnd 'show open dialog

If Err = ErrCancel Then Exit Sub 'if they click cancel
If ImgAdmin1.StatusCode <> 0 Then 'check for errors
    MsgBox Err.Description + " Code = " + Hex(ImgAdmin1.StatusCode), 16 'only on errors
    Exit Sub
End If

Image = ImgAdmin1.Image 'sets image file

With ImgEdit1
    .Image = Image 'loads image file
    .Display    'shows image file
End With

'Displays filename and widhtxheight in titlebar
Form1.Caption = "Quick Image - " & Image & " " & ImgEdit1.ImageWidth & "x" & ImgEdit1.ImageHeight

'Set original resolutions
OriginalResX = ImgEdit1.ImageResolutionX
OriginalResY = ImgEdit1.ImageResolutionY

'enable menu command
mnuEdit.Enabled = True
mnuImage.Enabled = True
mnuFileSave.Enabled = True
mnuFileClose.Enabled = True
mnuFilePrint.Enabled = True

'check Color or Grey scale
If colour = True Then
    mnuImageColorRGB.Checked = True
    mnuImageColorGrey.Checked = False
ElseIf colour = False Then
    mnuImageColorRGB.Checked = False
    mnuImageColorGrey.Checked = True
End If
'positions Mouse Pos label1
With Label1
    .Top = ImgEdit1.Height - 400
    .Left = 20
End With

FileOpen = True

End Sub

Private Sub mnuFilePrint_Click()
On Error GoTo PrintErr
'show printer properties, you can change settings
CommonDialog1.ShowPrinter
'this will print the image!
ImgEdit1.PrintImage 1, 1

PrintErr:
    If Err.Number = 32755 Then
        Exit Sub
    Else
        MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Error!"
        Exit Sub
    End If
End Sub

Private Sub mnuFileSave_Click()
Dim FileType As Integer

On Error Resume Next 'error handling

'sets filetypes for save
With ImgAdmin1
    .Filter = "TIFF files (*.tif)|*.tif|BMP files (*.bmp)|*.bmp|AWD files(*.awd)|*.awd|"
    .ShowFileDialog 1, Form1.hWnd
End With

If Err = ErrCancel Then 'Cancel pressed
    Exit Sub
End If

If ImgAdmin1.Image = ImgEdit1.Image Then 'Save
    ImgEdit1.Save False
Else 'Save as
    If ImgAdmin1.FilterIndex = 1 Then
        FileType = Tiff
    ElseIf ImgAdmin1.FilterIndex = 2 Then
        FileType = Bmp
    Else
        FileType = Awd 'awd is a fax document! saves in black and white, bad quality
    End If
    ImgEdit1.SaveAs ImgAdmin1.Image, FileType 'saves
    ImgEdit1.Image = ImgAdmin1.Image 'sets currently saved name for save later
    ImgAdmin1.Image = ImgEdit1.Image 'this will refresh the properties in the Admin control
    
End If
With ImgAdmin1
    .FilterIndex = 0
    .Filter = ""
End With
If ImgEdit1.StatusCode <> 0 Then 'on errors in ImgEdit1
    MsgBox Err.Description + " Code = " + Hex(ImgEdit1.StatusCode), 16
    Exit Sub
End If
End Sub

Private Sub mnuFileScan_Click()
On Error GoTo ScanErr
'Scan an Image
If ImgScan1.ScannerAvailable = True Then
With ImgScan1
    .OpenScanner
    .StartScan
End With

ImgEdit1.Image = ImgScan1.Image 'sets imgedit1 to show scanned image

Form1.Caption = "Quick Image - Scanned Image " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY

ImgScan1.CloseScanner 'close scanner
Else
    MsgBox "Scanner Not available!", vbExclamation, "Error!"
End If
ScanErr:
    If Err.Number = 1002 Then 'Exited without scanning
        Exit Sub
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuImageColorGrey_Click()
'sets image to grey scale
If colour = True Then
    ImgEdit1.DisplayScaleAlgorithm = 0
    ImgEdit1.Refresh
    colour = False
    mnuImageColorRGB.Checked = False
    mnuImageColorGrey.Checked = True
    Exit Sub
ElseIf colour = False Then
    Exit Sub
End If
End Sub

Private Sub mnuImageColorRGB_Click()
'sets image to RGB scale, Color
If colour = False Then
    ImgEdit1.DisplayScaleAlgorithm = 2
    ImgEdit1.Refresh
    colour = True
    mnuImageColorRGB.Checked = True
    mnuImageColorGrey.Checked = False
    Exit Sub
ElseIf colour = True Then
    Exit Sub
End If
End Sub

Private Sub mnuImageRes_Click()
'this changes the RESOLUTION!!! Not the height and width!!
Form2.Visible = True
End Sub

Private Sub mnuImageRotate180_Click()
'flip
ImgEdit1.Flip
End Sub

Private Sub mnuImageRotate90ccw_Click()
'turns it counterclockwise 90 deg
ImgEdit1.RotateLeft
End Sub

Private Sub mnuImageRotate90cw_Click()
'turns it clockwise 90 deg
ImgEdit1.RotateRight
End Sub

Private Sub mnuImageText_Click()
'allows you to write text on the image
ImgEdit1.SelectTool AnnoText
End Sub

Private Sub mnuImageZoomFifth_Click()
'zoom out
ImgEdit1.Zoom = 20
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom 1/5"
End Sub

Private Sub mnuImageZoomHalf_Click()
'zoom out
ImgEdit1.Zoom = 50
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom 1/2"
End Sub

Private Sub mnuImageZoomOriginal_Click()
'zoom to original size
ImgEdit1.Zoom = 100
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY
End Sub

Private Sub mnuImageZoomQuartar_Click()
'zoom out
ImgEdit1.Zoom = 25
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom 1/4"
End Sub

Private Sub mnuImageZoomTenth_Click()
'zoom out
ImgEdit1.Zoom = 10
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom 1/10"
End Sub

Private Sub mnuImageZoomtoFit_Click()
'Fits image to application window
ImgEdit1.FitTo (BestFit)
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Fit to Screen"
End Sub

Private Sub mnuImageZoomtoSel_Click()
'Zoom to a selected area
On Error GoTo ZoomErr
ImgEdit1.ZoomToSelection

ZoomErr:
    If Err.Number = 1007 Then 'nothing selected
        MsgBox "You Must Select An Area On The Image!", vbExclamation, "HEY!"
    Else
        Exit Sub
    End If
End Sub

Private Sub mnuImageZoomTwenty_Click()
'zoom out
ImgEdit1.Zoom = 5
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom 1/20"
End Sub

Private Sub mnuImageZoomX10_Click()
'zoom in
ImgEdit1.Zoom = 1000
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom X10"
End Sub

Private Sub mnuImageZoomX2_Click()
'zoom in
ImgEdit1.Zoom = 200
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom X2"
End Sub

Private Sub mnuImageZoomX20_Click()
'zoom in
ImgEdit1.Zoom = 2000
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom X20"
End Sub

Private Sub mnuImageZoomX4_Click()
'zoom in
ImgEdit1.Zoom = 400
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom X4"
End Sub

Private Sub mnuImageZoomX5_Click()
'zoom in
ImgEdit1.Zoom = 500
Form1.Caption = "Quick Image - " & ImgEdit1.Image & " " & ImgEdit1.ImageResolutionX & "x" & ImgEdit1.ImageResolutionY & " Zoom X5"
End Sub


Private Sub Timer1_Timer()
'Shows mouse position in label1
Call Get_XY
Label1.Caption = "X: " & MX & " Y: " & MY
End Sub
