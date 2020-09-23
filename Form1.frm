VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cut up an image"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox r2c2 
      Height          =   2175
      Left            =   6960
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
   End
   Begin VB.PictureBox r2c1 
      Height          =   2175
      Left            =   4080
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
   Begin VB.PictureBox r1c2 
      Height          =   2175
      Left            =   6960
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox r1c1 
      Height          =   2175
      Left            =   4080
      ScaleHeight     =   141
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   181
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cut into little pieces"
      Height          =   555
      Left            =   480
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   3800
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   250
      ScaleMode       =   0  'User
      ScaleWidth      =   251.976
      TabIndex        =   0
      Top             =   120
      Width           =   3825
   End
   Begin VB.Label Label3 
      Caption         =   "by John Sheridan"
      Height          =   255
      Left            =   2400
      TabIndex        =   8
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Just use SavePicture() to save these images."
      Height          =   255
      Left            =   5160
      TabIndex        =   7
      Top             =   4680
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Original ^"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   3960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'this code what asked about on psc, so i thought
'i'd show an easy way to do it
'please vote for me!

r1c1.Cls
r1c2.Cls
r2c1.Cls
r2c2.Cls 'clear all the pictures


'row 1, column 1
For x = 0 To img.ScaleWidth / 2
 For y = 0 To img.ScaleHeight / 2
  r1c1.PSet (x, y), img.Point(x, y)
 Next y
Next x

'row2, column 1
For x = 0 To img.ScaleWidth / 2
 For y = img.ScaleHeight / 2 To img.ScaleHeight
  r2c1.PSet (x, y - (img.ScaleHeight / 2)), img.Point(x, y)
 Next y
Next x

'row1, column 2
For x = img.ScaleWidth / 2 To img.ScaleWidth
 For y = 0 To img.ScaleHeight / 2
  r1c2.PSet (x - (img.ScaleWidth / 2), y), img.Point(x, y)
 Next y
Next x

'row2, column 2
For x = img.ScaleWidth / 2 To img.ScaleWidth
 For y = img.ScaleHeight / 2 To img.ScaleHeight
  r2c2.PSet (x - (img.ScaleWidth / 2), y - (img.ScaleHeight / 2)), img.Point(x, y)
 Next y
Next x

'please vote!
End Sub
