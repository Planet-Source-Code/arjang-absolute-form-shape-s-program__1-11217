VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form Shapes"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Q"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&S"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   375
   End
   Begin VB.OptionButton optCircle 
      Caption         =   "Circle"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.OptionButton optDialog 
      Caption         =   "Dialog"
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.OptionButton optTriangle 
      Caption         =   "Triangle"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton optDiamond 
      Caption         =   "Diamond"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.OptionButton optStar 
      Caption         =   "Star"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "For more info feel free to email me at arjang7@hotmail.com"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   4125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
   End
End Sub

Private Sub cmdShow_Click()
Dim X(2) As pointapi
Dim Diamond(3) As pointapi
Dim Star(5) As pointapi
Dim lRegion As Long
Dim lRegion1 As Long
Dim lRegion2 As Long
Dim lResult As Long

If optCircle.Value = True Then
   Unload frmTest
   Load frmTest
   frmTest.Width = 500 * Screen.TwipsPerPixelX
   frmTest.Height = 500 * Screen.TwipsPerPixelY
   lRegion = CreateEllipticRgn(0, 0, 250, 250)
   lResult = SetWindowRgn(frmTest.hWnd, lRegion, True)
   frmTest.Show
ElseIf optDialog.Value = True Then
   Unload frmTest
   Load frmTest
   frmTest.Width = 500 * Screen.TwipsPerPixelX
   frmTest.Height = 500 * Screen.TwipsPerPixelY
   X(0).X = 0
   X(0).Y = 0
   X(1).X = 100
   X(1).Y = 0
   X(2).X = 0
   X(2).Y = 100
   lRegion = CreatePolygonRgn(X(0), 3, alternate)
   X(0).X = 100
   X(0).Y = 200
   X(1).X = 110
   X(1).Y = 200
   X(2).X = 100
   X(2).Y = 220
   lRegion1 = CreatePolygonRgn(X(0), 3, alternate)
   lRegion2 = CreateRoundRectRgn(0, 0, 202, 202, 30, 30)
   lResult = CombineRgn(lRegion, lRegion1, lRegion2, rgn_or)
   DeleteObject lRegion1
   DeleteObject lRegion2
   lResult = SetWindowRgn(frmTest.hWnd, lRegion, True)
   frmTest.Show
 ElseIf optDiamond.Value = True Then
   Unload frmTest
   Load frmTest
   frmTest.Width = 500 * Screen.TwipsPerPixelX
   frmTest.Height = 500 * Screen.TwipsPerPixelY
   Diamond(0).X = 100
   Diamond(0).Y = 0
   Diamond(1).X = 0
   Diamond(1).Y = 150
   Diamond(2).X = 200
   Diamond(2).Y = 150
   Diamond(3).X = 100
   Diamond(3).Y = 300
   lRegion = CreatePolygonRgn(Diamond(0), 4, alternate)
   lResult = SetWindowRgn(frmTest.hWnd, lRegion, True)
   frmTest.Show
ElseIf optStar.Value = True Then
   Unload frmTest
   Load frmTest
   frmTest.Width = 500 * Screen.TwipsPerPixelX
   frmTest.Height = 500 * Screen.TwipsPerPixelY
   Star(0).X = 231
   Star(0).Y = 12
   Star(1).X = 220
   Star(1).Y = 57
   Star(2).X = 259
   Star(2).Y = 31
   Star(3).X = 209
   Star(3).Y = 31
   Star(4).X = 245
   Star(4).Y = 57
   Star(5).X = 231
   Star(5).Y = 12
   lRegion = CreatePolygonRgn(Star(0), 6, winding)
   lResult = SetWindowRgn(frmTest.hWnd, lRegion, True)
   frmTest.Show
ElseIf optTriangle.Value = True Then
   Unload frmTest
   Load frmTest
   frmTest.Width = 500 * Screen.TwipsPerPixelX
   frmTest.Height = 500 * Screen.TwipsPerPixelY
   X(0).X = 100
   X(0).Y = 0
   X(1).X = 0
   X(1).Y = 150
   X(2).X = 200
   X(2).Y = 150
   lRegion = CreatePolygonRgn(X(0), 3, alternate)
   lResult = SetWindowRgn(frmTest.hWnd, lRegion, True)
   frmTest.Show
End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.Hide
End Sub

