VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   6285
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   4560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   4560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show2"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show1"
      Height          =   495
      Left            =   4920
      TabIndex        =   0
      Top             =   4320
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''<!--******************************************************************************************-->
''
'' Program Name: Draw a Taichie Oracle Demo
''
'' AUTHOR: Wen-Kai Chung
''
'' 2002/8/4 New Create
''
'' AUTHOR'S E-mail: wkchung@anet.net.tw
''
'' COMMENTS:None
''
''<!--******************************************************************************************-->
''<HR>

Option Explicit

Public i As Integer
Public j As Integer
Const PI = 3.1415926

Dim bangle As Single
Dim eangle As Single
Dim bangle1 As Single
Dim eangle1 As Single
Dim secX As Single
Dim secY As Single
Dim secangle As Single


Private Sub Command1_Click()
  
  Timer2.Enabled = False
  Timer1.Interval = 1000 / 27 'draw 27 pics per sec
  Timer1.Enabled = Not Timer1.Enabled
    
End Sub


Private Sub Command2_Click()
  
  Timer1.Enabled = False
  Timer2.Interval = 1000 / 27 'draw 27 pics per sec
  Timer2.Enabled = Not Timer2.Enabled
    
End Sub


Private Sub form_activate()

i = 90 'Given Show2 begin angle from 90 degree
j = 90 'Given Shoo1 begin angle from 90 degree

FillStyle = 0

bangle = 90 * PI / 180
eangle = 270 * PI / 180
FillColor = vbBlack 'half of Big black circle
Circle (3000, 2500), 2000, vbBlack, -bangle, -eangle

bangle = 270 * PI / 180
eangle = 90 * PI / 180
FillColor = vbWhite 'half of Big white circle
Circle (3000, 2500), 2000, vbWhite, -bangle, -eangle


bangle = 90 * PI / 180
eangle = 270 * PI / 180
FillColor = vbWhite 'half of Small white circle
Circle (3000, 3500), 1000, vbWhite, -bangle, -eangle

bangle = 270 * PI / 180
eangle = 90 * PI / 180
FillColor = vbBlack  'Half of Small black circle
Circle (3000, 1500), 1000, vbBlack, -bangle, -eangle

FillColor = vbBlack  'Small black circle
Circle (3000, 3500), 500, vbBlack

FillColor = vbWhite  'Small white circle
Circle (3000, 1500), 500, vbWhite

End Sub


Private Sub Timer1_Timer()  'Show1

bangle1 = 90
eangle1 = 90

j = j + 3 'every time steps 3 degree

bangle1 = j
eangle1 = j

FillStyle = 0

FillColor = vbBlack
bangle = bangle1 * PI / 180
eangle = (eangle1 + 180) * PI / 180

bangle = ((bangle / 2 / PI) - Int(bangle / 2 / PI)) * 2 * PI
eangle = ((eangle / 2 / PI) - Int(eangle / 2 / PI)) * 2 * PI

Circle (3000, 2500), 2000, vbBlack, -bangle, -eangle

bangle = (bangle1 + 180) * PI / 180
eangle = eangle1 * PI / 180
FillColor = vbWhite
bangle = ((bangle / 2 / PI) - Int(bangle / 2 / PI)) * 2 * PI
eangle = ((eangle / 2 / PI) - Int(eangle / 2 / PI)) * 2 * PI
Circle (3000, 2500), 2000, vbWhite, -bangle, -eangle

FillColor = vbWhite
bangle = bangle1 * PI / 180
eangle = (eangle1 + 180) * PI / 180
bangle = ((bangle / 2 / PI) - Int(bangle / 2 / PI)) * 2 * PI
eangle = ((eangle / 2 / PI) - Int(eangle / 2 / PI)) * 2 * PI
Circle (3000 - eangle, 3500 + bangle), 1000, vbWhite, -bangle, -eangle

bangle = (bangle1 + 180) * PI / 180
eangle = bangle1 * PI / 180
FillColor = vbBlack
bangle = ((bangle / 2 / PI) - Int(bangle / 2 / PI)) * 2 * PI
eangle = ((eangle / 2 / PI) - Int(eangle / 2 / PI)) * 2 * PI
Circle (3000 + eangle, 1500 - bangle), 1000, vbBlack, -bangle, -eangle

FillColor = vbBlack
Circle (3000, 3500), 500, vbBlack

FillColor = vbWhite
Circle (3000, 1500), 500, vbWhite

If j >= 450 Then j = 90

End Sub


Private Sub Timer2_Timer()  'Show2

i = i + 3  'every time steps 3 degree

   
FillStyle = 0

bangle = i * PI / 180
If i > 360 Then bangle = (i - 360) * PI / 180

eangle = (180 + i) * PI / 180
If (180 + i) > 360 Then eangle = (180 + i - 360) * PI / 180

FillColor = vbBlack 'half of Big black circle
Circle (3000, 2500), 2000, vbBlack, -bangle, -eangle

bangle = (180 + i) * PI / 180
If (180 + i) > 360 Then bangle = (180 + i - 360) * PI / 180

eangle = i * PI / 180
If i > 360 Then eangle = (i - 360) * PI / 180

FillColor = vbWhite 'half of Big white circle
Circle (3000, 2500), 2000, vbWhite, -bangle, -eangle


secangle = 2 * PI * ((180 - i) / 360)
secX = Cos(secangle) * 1000
secY = Sin(secangle) * 1000

'bangle = i * PI / 180
'If i > 360 Then bangle = (i - 360) * PI / 180

'eangle = (180 + i) * PI / 180
'If (180 + i) > 360 Then eangle = (180 + i - 360) * PI / 180

FillColor = vbWhite 'half of small white circle,but we draw a clrcle here
Circle (3000 + secX, 2500 + secY), 1000, vbWhite ', -bangle, -eangle
FillColor = vbBlack 'small black circle
Circle (3000 + secX, 2500 + secY), 500, vbBlack


secangle = 2 * PI * ((360 - i) / 360)
secX = Cos(secangle) * 1000
secY = Sin(secangle) * 1000

'bangle = (180 + i) * PI / 180
'If (180 + i) > 360 Then bangle = (180 + i - 360) * PI / 180

'eangle = i * PI / 180
'If i > 360 Then eangle = (i - 360) * PI / 180

FillColor = vbBlack 'half of Small black circle,but we draw a circle here
Circle (3000 + secX, 2500 + secY), 1000, vbBlack ', -bangle, -eangle
FillColor = vbWhite 'small white circle
Circle (3000 + secX, 2500 + secY), 500, vbWhite


If i >= 450 Then i = 90


End Sub
