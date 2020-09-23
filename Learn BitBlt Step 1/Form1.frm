VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn BitBlt - Step 1                                          :)     Exit program ->"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   266
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start Engine (Main loop)"
      Height          =   615
      Left            =   5040
      TabIndex        =   4
      Top             =   3000
      Width           =   1335
   End
   Begin VB.PictureBox Sprite1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   4320
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   900
      ScaleWidth      =   450
      TabIndex        =   2
      Top             =   3000
      Width           =   450
   End
   Begin VB.PictureBox MainField 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      Picture         =   "Form1.frx":240D
      ScaleHeight     =   247
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "<- Recomended that you put properties - ScaleMode=Pixel       (for easier sprite moving)"
      Height          =   735
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "<- If you want to draw something in the picture box or on form you must make in his properties- AutoRedraw=True"
      Height          =   855
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XX As Integer, YY As Integer

Private Sub Command1_Click()
'Give him starting position
XX = Int(Rnd * MainField.Width) 'Same as left=100 (I just made a random position somewhere on the MainField)
YY = -30                        'Same as top=-30

Command1.Caption = "Engine is ON"
MainLoop 'It starts Main Loop who will stop automaticly when little house comes to the end of the screen/MainField
End Sub

Public Sub MainLoop()
Do 'Starting loop
MainField.Cls

'First you must draw him a Mask (Mask is needed
'to create a transparen color around the picture
BitBlt MainField.hDC, XX, YY, 30, 30, Sprite1.hDC, 0, 30, SRCAND
'Now draw the picture over Mask
BitBlt MainField.hDC, XX, YY, 30, 30, Sprite1.hDC, 0, 0, SRCPAINT

'Now, move the object one pixel to down (small house)
YY = YY + 1

'If it get outside the screen to end this loop
If YY > MainField.Height Then
    Command1.Caption = "Start Engine (Main loop)"
    Exit Sub
End If

DoEvents
Sleep 1 'If you have good PC, to slow down application - to pause all processes on 1 milisecond (1sec = 1000 ms)

Loop 'Contionue loop
End Sub

Private Sub Form_Unload(Cancel As Integer)
End 'Confirm exit program
End Sub
