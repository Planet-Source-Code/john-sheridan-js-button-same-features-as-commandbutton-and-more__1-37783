VERSION 5.00
Begin VB.UserControl JSbutton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   ScaleHeight     =   915
   ScaleWidth      =   1965
   Begin VB.Shape bord 
      BorderColor     =   &H0000FF00&
      Height          =   255
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label but 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label txt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Button1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape l 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "JSbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''
'Button OCX control by John Sheridan
'I'm sure this will help everyone.
'You can also improve this code by
'adding new properties or making a cooler fade.
'
'Please vote and if i messed anything up, tell
'me.     (to vote, just do a search for
'        "john sheridan" in vb, and
'        my submissions will popup)
'
'I hope this helps ppl! :-)
''''''''''''''''''''''''''''''''''''''''''


'declare events
Event clickMe()
Event mouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event mouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event resizeMe()
Dim fade As Integer

'fade the button to 3d
Private Function threeDize()
UserControl.ScaleMode = vbPixels
Dim x, y, clp

clp = fade / UserControl.ScaleHeight

  For y = 0 To UserControl.ScaleHeight

   UserControl.Line (0, y)-(UserControl.Width, y), RGB(255 - (y * clp), 255 - (y * clp), 255 - (y * clp))
    
  Next y

UserControl.ScaleMode = vbTwips
End Function

'declare object properties
Public Property Get caption() As String
caption = txt.caption  'return button caption
End Property

Public Property Let caption(ByVal newCap As String)
txt.caption() = newCap 'set button caption
End Property


Public Property Get fadeAmount() As Integer
 fadeAmount = fade  'return fade amount
End Property

Public Property Let fadeAmount(ByVal newFade As Integer)
If newFade < 1 Or newFade > 255 Then
MsgBox "Invalid fade amount: " & newFade, vbInformation, "Button OCX"
Exit Property
End If
fade = newFade  'set fade amount
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor 'return bgcolor
End Property

Public Property Let BackColor(ByVal newBack As OLE_COLOR)
    UserControl.BackColor() = newBack 'set bgcolor
End Property


Private Sub but_Click()
'RaiseEvent calls what the user put in as
'code in his/her vb app
RaiseEvent clickMe
End Sub

Private Sub but_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
threeDize
l.BorderColor = RGB(0, 0, 255)
l.BorderWidth = 2

RaiseEvent mouseDown(Button, Shift, x, y)
'call event
End Sub


Private Sub but_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

RaiseEvent mouseMove(Button, Shift, x, y)
'call event
End Sub

Private Sub but_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

UserControl.Cls   'get rid of 3d effect

l.BorderColor = RGB(0, 0, 0)
l.BorderWidth = 1

RaiseEvent mouseUp(Button, Shift, x, y)
'call event
End Sub


Private Sub UserControl_GotFocus()
bord.Visible = True
'make green box appear
End Sub

Private Sub UserControl_Initialize()
fade = 200
'make the fade (by default) 200
End Sub

Private Sub UserControl_LostFocus()
bord.Visible = False
'make the green box disappear
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
but.Move 0, 0, UserControl.Width, UserControl.Height * 100
l.Move 0, 0, UserControl.Width, UserControl.Height
txt.Move 0, (UserControl.Height / 2) - (txt.Height / 2), UserControl.Width, 255
bord.Move 120, 120, l.Width - 240, l.Height - 240
'move all the objects into correct positions

RaiseEvent resizeMe
'call the event
End Sub
