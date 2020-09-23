VERSION 5.00
Begin VB.UserControl SwooshBox 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ScaleHeight     =   315
   ScaleWidth      =   480
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   435
   End
End
Attribute VB_Name = "SwooshBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim sStartWidth As Integer
Dim sEndWidth As Integer
Dim ScrollOut As Boolean

Event ScrollBack()
Event ScrollOut()
Event Click()

Private Sub Text1_Click()
    RaiseEvent Click
End Sub

Private Sub Text1_GotFocus()
    Dim X As Integer
    ScrollOut = True
    Do Until Width >= sEndWidth
        Width = Width + 70
        X = 0
        Do Until X = 1000
            If ScrollOut = False Then Exit Sub
            DoEvents
            X = X + 1
        Loop
    Loop
    RaiseEvent ScrollOut
End Sub

Private Sub Text1_LostFocus()
    Dim X As Integer
        ScrollOut = False
        Do Until X = 10000
            DoEvents
            X = X + 1
        Loop
        
    Do Until Width <= sStartWidth
        Width = Width - 70
        X = 0
        Do Until X = 1000
            If ScrollOut = True Then Exit Sub
            DoEvents
            X = X + 1
        Loop
    Loop
    RaiseEvent ScrollBack
End Sub

Private Sub UserControl_InitProperties()
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    Height = 285
    'Width = sStartWidth
    Text1.Width = Width
End Sub


Public Property Get StartWidth() As Integer
    StartWidth = sStartWidth
End Property
Public Property Let StartWidth(NewSWidth As Integer)
    sStartWidth = NewSWidth
    PropertyChanged "StartWidth"
    Width = sStartWidth
End Property

Public Property Get EndWidth() As Integer
    EndWidth = sEndWidth
End Property
Public Property Let EndWidth(NewEWidth As Integer)
    sEndWidth = NewEWidth
    PropertyChanged "EndWidth"
End Property

Public Property Get ForeColor() As OLE_COLOR
    EndWidth = Text1.ForeColor
End Property
Public Property Let ForeColor(NewForeColor As OLE_COLOR)
    Text1.ForeColor = NewForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Text() As String
    Text = Text1
End Property
Public Property Let Text(newText As String)
    Text1 = newText
    PropertyChanged "Text"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Text1 = PropBag.ReadProperty("Text", "")
    sEndWidth = PropBag.ReadProperty("EndWidth", 1575)
    sStartWidth = PropBag.ReadProperty("StartWidth", 435)
    Text1.ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    Text1.Width = sStartWidth
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Text", Text1.Text, "")
    Call PropBag.WriteProperty("EndWidth", sEndWidth, 1575)
    Call PropBag.WriteProperty("StartWidth", sStartWidth, 435)
    Call PropBag.WriteProperty("ForeColor", Text1.ForeColor, vbBlack)
End Sub
