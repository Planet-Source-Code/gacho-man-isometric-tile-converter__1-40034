VERSION 5.00
Begin VB.Form MyForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daniel Marin's Isometric Converter"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6840
   Icon            =   "MyForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   393
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MyPicture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   5
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton MyCommand 
      Caption         =   "Transform to Isometric"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   4
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Frame MyFrame 
      Caption         =   "Source Image"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.DriveListBox MyDrive 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C8FF&
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   3495
      End
      Begin VB.FileListBox MyFile 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C8FF&
         Height          =   1650
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.DirListBox MyDir 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C8FF&
         Height          =   1665
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox MySource 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   4200
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame MyOtherFrame 
      Caption         =   "Options"
      Height          =   1095
      Left            =   3960
      TabIndex        =   7
      Top             =   3360
      Width           =   2775
      Begin VB.OptionButton MyOption 
         Caption         =   "Right Wall"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton MyOption 
         Caption         =   "Left Wall"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.OptionButton MyOption 
         Caption         =   "Horizontal Tile"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame MyLastFrameIPromess 
      Caption         =   "BackColor"
      Height          =   1335
      Left            =   3960
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
      Begin VB.HScrollBar MyScroll 
         Height          =   255
         Index           =   2
         LargeChange     =   15
         Left            =   120
         Max             =   255
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.HScrollBar MyScroll 
         Height          =   255
         Index           =   1
         LargeChange     =   15
         Left            =   120
         Max             =   255
         TabIndex        =   16
         Top             =   600
         Width           =   2055
      End
      Begin VB.HScrollBar MyScroll 
         Height          =   255
         Index           =   0
         LargeChange     =   15
         Left            =   120
         Max             =   255
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label MyColorLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H0000C8FF&
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.Label MyColorLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H0000C8FF&
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   17
         Top             =   600
         Width           =   375
      End
      Begin VB.Label MyColorLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H0000C8FF&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Image MyBuffer 
      Height          =   495
      Left            =   4200
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image MyImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   2400
      Left            =   4080
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label MyLabel 
      Caption         =   "Source:"
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label MyLabel 
      Caption         =   "Result:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   615
   End
   Begin VB.Menu MyMenu 
      Caption         =   "&Commands"
      Begin VB.Menu MyItem 
         Caption         =   "&Copy Result to Clipboard"
         Index           =   0
         Shortcut        =   ^C
      End
      Begin VB.Menu MyItem 
         Caption         =   "&Paste Source from Clipboard"
         Index           =   1
         Shortcut        =   ^V
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu EndProgram 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "MyForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub EndProgram_Click()
End
End Sub

Private Sub MyCommand_Click(Index As Integer)
Select Case Index
Case 0
    Dim k As Byte
    If Not MyCommand(0).Caption = "Cancel" Then
        MyCommand(0).Caption = "Cancel"
        For k = 0 To 2
            If MyOption(k).Value Then Iso_Transform k: Exit For
        Next k
    Else
        MyCommand(0).Caption = "Transform To Isometric"
        Caption = "Isometric Cancelled"
    End If
End Select
End Sub

Sub Iso_Transform(ByVal Action As Byte)
'variable 'Action'
    '+---+-------------+
    '| # | Description |
    '+---+-------------+
    '| 0 | Floor       |
    '| 1 | Ceiling     |
    '| 2 | Left Wall   |
    '| 3 | Right Wall  |
    '+---+-------------+
'Coordinates of the drawing point in the image
    Dim x As Long, y As Long
'Isometric Coordinates of that point or "real ones"
    Dim Rx As Long, Ry As Long
'For best performance we hide and clear 'MyPicture'
    MyPicture.Visible = False
    MyPicture.Cls
'Resize it for the isometric result, we use X and Y as temporary Width and Height
    If Action = 0 Then
        'Dimensions of the result of an horizontal tile
        x = MyBuffer.Width + MyBuffer.Height
        y = Int(x / 2)
    Else
        'Dimensions of the result of walls
        x = MyBuffer.Width
        y = MyBuffer.Height + Fix(x / 2)
    End If
        MyPicture.Move MyPicture.Left, MyPicture.Top, x, y
    Select Case Action
 'Creating an Horizontal Tile
    Case 0
        For x = 0 To MyBuffer.Width - 1
            For y = 0 To MyBuffer.Height - 1
            'Real Coordinates
                Rx = x + y
                Ry = Fix(MyBuffer.Width / 2) - Int(x / 2) + Fix(y / 2)
            'Paint the isometric points
                MyPicture.PSet (Rx, Ry), MySource.Point(x, y)
                MyPicture.PSet (Rx + 1, Ry), MySource.Point(x, y)
                DoEvents
                If Not MyCommand(0).Caption = "Cancel" Then Exit Sub
            Next y
            Caption = "Transforming..." & Int(100 * x / MyBuffer.Width) & "%"
        Next x
'Creating Walls 1=Left, 2=Right
    Case 1, 2
        For x = 0 To MyBuffer.Width - 1
            For y = 0 To MyBuffer.Height - 1
            'Real Coordinates
                Rx = x
                If Action = 1 Then
                    Ry = y + Fix(x / 2)
                Else
                    Ry = Fix(MyBuffer.Width / 2) - Fix(x / 2) + y
                End If
            'Paint the isometric points
                MyPicture.PSet (Rx, Ry), MySource.Point(x, y)
                'MyPicture.PSet (Rx + 1, Ry), MySource.Point(x, y)
                DoEvents
                If Not MyCommand(0).Caption = "Cancel" Then Exit Sub
            Next y
            Caption = "Transforming..." & Int(100 * x / MyBuffer.Width) & "%"
        Next x
    End Select
    MyCommand(0).Caption = "Transform To Isometric"
    Caption = "Isometric Ready"
    MyPicture.Visible = True
'BEEP! BEEP! I'VE FINISHED MAN
    Beep
End Sub

Private Sub MyDir_Change()
MyFile.Path = MyDir.Path
End Sub

Private Sub Form_Load()
MyDrive.Drive = Left$(App.Path, 2)
End Sub

Private Sub MyDrive_Change()
MyDir.Path = Left$(MyDrive.Drive, 2) + "\"
End Sub

Private Sub MyFile_Click()
On Error GoTo ErrorHandler
MyBuffer.Picture = LoadPicture(MyDir.Path & "\" & MyFile.FileName)
MyImage.Picture = MyBuffer.Picture
Caption = MyBuffer.Width & "x" & MyBuffer.Height
MySource.Cls
MySource.Width = MyBuffer.Width
MySource.Height = MyBuffer.Height
MySource.Picture = MyBuffer.Picture
Exit Sub

ErrorHandler:
Select Case Err.Number
Case 481 'The user selected a non-image-recognized file
    Beep
Case Else
    MsgBox "An error #" & Err.Number & " has ocurred" & Chr$(13) & "'" & Err.Description & "'", vbCritical, "ISO_ERROR"
    End
End Select
End Sub

Private Sub MyItem_Click(Index As Integer)
Select Case Index
Case 1
    If Clipboard.GetFormat(2) Then
        MyImage.Picture = Clipboard.GetData(2)
        MyBuffer.Picture = Clipboard.GetData(2)
        Caption = MyBuffer.Width & "x" & MyBuffer.Height
        MySource.Cls
        MySource.Width = MyBuffer.Width
        MySource.Height = MyBuffer.Height
        MySource.Picture = MyBuffer.Picture
    Else
        Beep
    End If
Case 0
    'Dim lastW As Integer, lastH As Integer
    'lastW = MyPicture.Width
    'lastH = MyPicture.Height
    Clipboard.Clear
    Clipboard.SetData MyPicture.Image
    'MyPicture.Move MyPicture.Top, MyPicture.Top, lastW, lastH
End Select
End Sub

Private Sub MyScroll_Change(Index As Integer)
MyColorLabel(Index).Caption = MyScroll(Index).Value
MyPicture.BackColor = RGB(MyScroll(0).Value, MyScroll(1).Value, MyScroll(2).Value)
End Sub

Private Sub MyScroll_Scroll(Index As Integer)
MyColorLabel(Index).Caption = MyScroll(Index).Value
MyPicture.BackColor = RGB(MyScroll(0).Value, MyScroll(1).Value, MyScroll(2).Value)
End Sub
