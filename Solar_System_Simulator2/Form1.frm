VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   2445
   ClientTop       =   1935
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   8805
   Begin VB.CommandButton Command1 
      Caption         =   "Code Under Here"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   6360
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim X As Long
    Dim Y As Long
    Dim dOffset As Double
    Dim dHeight As Double
    
    Dim dMustIncludeX As Double
    Dim dMustIncludeY As Double
    Dim dBroadness As Double
    
    Me.ScaleMode = 1 ' set to user scale
    
    ' to make 0,0 in the middle of the screen
    Me.ScaleLeft = -500      'Max left is -500
    Me.ScaleWidth = 1000     'Max right is 500
    Me.ScaleTop = 500        'Max top is 500
    Me.ScaleHeight = -1000    'Max bottom is -500
    
    
    'Set default curve values
    dOffset = 150
    dHeight = 320
    
    ' The curve must pass through this point
    dMustIncludeX = 300
    dMustIncludeY = -10
    
    ' Calculate the broadness coefficient
    dBroadness = (dMustIncludeY - dHeight) _
    / ((dMustIncludeX - dOffset) * (dMustIncludeX - dOffset))
    
    ' Clear the screen
    Me.Cls
    
    'Draw a Red box 10 inside our form to prove we have the co-ordinates correct
    Me.Line (-490, -490)-(490, 490), RGB(255, 0, 0), B
    
    X = 0
    Y = 0
    'Draw the origin
    Me.Line (-10, 0)-(10, 0), RGB(255, 255, 255) ' Draw Origin
    Me.Line (0, -10)-(0, 10), RGB(255, 255, 255) ' Draw Origin

    ' Calculate the curve and print it to screen
    For X = -500 To 500
        ' Calculate the Y position
        Y = dBroadness * ((X - dOffset) * (X - dOffset)) + dHeight
        ' Draw the pixel on screen
        Me.PSet (X, Y), RGB(0, 0, 255) ' Draw Blue Pixel
    Next
    
    
End Sub
