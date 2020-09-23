VERSION 5.00
Begin VB.UserControl Fire 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   193
   ToolboxBitmap   =   "ucFire.ctx":0000
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   61
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox PictureBig 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   120
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   240
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1380
      Top             =   -60
   End
End
Attribute VB_Name = "Fire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------------------------------------------------------------------
'
'                 « in the name of Allah »
'
' this control is a good fire. for use it you must
' add UcFire.ctl and UcFireAPI.Bas Module
'
' a) Some Properties:
' DX,DY: Real Size Of Fire -> these properties act on Resolution
' ColDecrease: Rate Of decreasing fire from bottom to top
' MouseSensitive: Fire go up and down by Mouse movement
' BackColor: Fire BackColor
' Text: the string if exist appears on the fire.
' DisplayAtDesign: determine if fire is active in design mode or not.
' ToolTipTextString: Use this instead of tooltiptext
' TimeInterval: interval of timer in msec (speed)
' FirePicObject: Read Only RunTime Access Object -> you can use this picture box object's Properies And Methods such as Print Or Image
'
' b) Some useful Info:
' this code first set colors of the last line in bottom
' of fire in random form and then  find upper pixels by
' a formula as a function of bottom and previous pixels.
' i use setpixel API function for writing pixels instead
' of pset vb method for more speed (almost 7 times).
' the 'DoFire' Sub is the main sub for drawing fire.
' you can modify it as well as you want and add some
' favorite effects to it. for example you can add some
' code in it for drawing on fire.
'
' by: Saeed Serpooshan - Iran - 2001 (1380sun)
' EMail: SSerpooshan@Yahoo.com
' WebPage: http://www.JamAcademic.com/vb.html
'------------------------------------------------------------------------------------------------------------------------------------------

Public Event Click()
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

DefInt A-Z
Dim flame(400, 200) As Integer
Dim ch13 As String * 2
Dim cols(255 * 3) As Long
Dim ColDecreaseVal As Integer
Dim ColDecMS As Integer
Dim dxVal As Integer, dyVal As Integer
Dim xText As Long, yText As Long

'
'Default Property Values:
Const m_def_Text = ""
Const m_def_DispAtDesign = 0
Const m_def_DY = 50
Const m_def_DX = 80
Const m_def_MouseSensitive = True
Const m_def_BackColor = &H40&
Const m_def_A_Creator = "By: S.Serpooshan (2001 Iran)"
'Property Variables:
Dim m_Text As String
Dim m_FirePicObject As Object
'Dim m_FirePicObject As Object
Dim m_DispAtDesign As Boolean
Dim m_DY As Long
Dim m_DX As Long
Dim m_MouseSensitive As Boolean
Dim m_BackColor As OLE_COLOR
Dim m_A_Creator As Variant
'end General

Private Sub PictureBig_Click()
 RaiseEvent Click
End Sub

 Sub PictureBig_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseMove(Button, Shift, x, y)
 If m_MouseSensitive Then
  ColDecMS = InterpolateLinear(0, -ColDecreaseVal, ScaleHeight - 2, ColDecreaseVal * 4, y)
 End If
End Sub

Private Sub Timer1_Timer()
 DoFire
End Sub

Private Sub Picture2_Click(Index As Integer)
 RaiseEvent Click
End Sub

Private Sub Picture2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
 RaiseEvent MouseMove(Button, Shift, x, y)
End Sub


Public Sub UserControl_Initialize()
ch13 = Chr$(13) + Chr$(10)
DX = 50: DY = 40
ColDecreaseVal = 10
With Picture1
 .ScaleMode = 3
 .AutoRedraw = True
 '.Visible = False
 .BackColor = 0   '&HFFFFFF
 .Width = DX * 1
 .Height = DY * 1
End With
PictureBig.Move 0, 0, ScaleWidth, ScaleHeight

End Sub

Public Property Get TimeInterVal() As Variant
 TimeInterVal = Timer1.Interval
End Property

Public Property Let TimeInterVal(ByVal vNewValue As Variant)
 Timer1.Interval = vNewValue
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
 Timer1.Interval = PropBag.ReadProperty("TimeInterval", 100)
 PictureBig.ToolTipText = PropBag.ReadProperty("ToolTipTextString", "")
 BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
 ColDecrease = PropBag.ReadProperty("ColDecrease", 10)
 DX = PropBag.ReadProperty("DX", m_def_DX)
 DY = PropBag.ReadProperty("DY", m_def_DY)
 m_Text = PropBag.ReadProperty("Text", m_def_Text)
 Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
 PictureBig.MousePointer = PropBag.ReadProperty("MousePointer", 0)
 m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
 UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
 m_A_Creator = PropBag.ReadProperty("A_Creator", m_def_A_Creator)
 m_MouseSensitive = PropBag.ReadProperty("MouseSensitive", m_def_MouseSensitive)
 m_DispAtDesign = PropBag.ReadProperty("DispAtDesign", m_def_DispAtDesign)
 Set m_FirePicObject = PropBag.ReadProperty("FirePicObject", Nothing)
 Set PictureBig.Font = PropBag.ReadProperty("Font", Ambient.Font)
 PictureBig.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
If Err Then MsgBox "ReadProp-Err:" + Err.Description

End Sub

Private Sub UserControl_Resize()
PictureBig.Move 0, 0, ScaleWidth, ScaleHeight
Call SetXYText
DoFire
'UserControl.Width = UserControl.ScaleX(DX, 3, 1)
End Sub

Private Sub UserControl_Show()
On Error Resume Next
 PictureBig.BackColor = BackColor
 Call SetXYText
 If Ambient.UserMode Or DispAtDesign Then
  Timer1.Enabled = True
 Else
  Timer1.Enabled = False
  With PictureBig
   .CurrentX = 0: .CurrentY = 0
   PictureBig.Print "UcFire  -  Design Mode"
  End With
 End If
If Err Then MsgBox "Init-Err:" + Err.Description
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 PropBag.WriteProperty "TimeInterval", Timer1.Interval
 PropBag.WriteProperty "BackColor", Picture1.BackColor
 PropBag.WriteProperty "ColDecrease", ColDecreaseVal
 PropBag.WriteProperty "DX", DX, m_def_DX
 PropBag.WriteProperty "DY", DY, m_def_DY
 PropBag.WriteProperty "ToolTipTextString", PictureBig.ToolTipText
 
 Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
 Call PropBag.WriteProperty("MousePointer", PictureBig.MousePointer, 0)
 Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
 Call PropBag.WriteProperty("A_Creator", m_A_Creator, m_def_A_Creator)
 Call PropBag.WriteProperty("MouseSensitive", m_MouseSensitive, m_def_MouseSensitive)
 Call PropBag.WriteProperty("DispAtDesign", m_DispAtDesign, m_def_DispAtDesign)
 Call PropBag.WriteProperty("FirePicObject", m_FirePicObject, Nothing)
 Call PropBag.WriteProperty("Text", m_Text, m_def_Text)
 Call PropBag.WriteProperty("Font", PictureBig.Font, Ambient.Font)
 Call PropBag.WriteProperty("ForeColor", PictureBig.ForeColor, &H80000012)
End Sub

Public Property Get ColDecrease() As Variant
 ColDecrease = ColDecreaseVal
End Property

Public Property Let ColDecrease(ByVal vNewValue As Variant)
 ColDecreaseVal = vNewValue
End Property


Sub SetColors()

 Dim c As Long
 'k = 255 / 63
 c = m_BackColor
 Picture1.BackColor = c
 r1 = c And &HFF
 g1 = (c Mod 65536) \ 256
 b1 = c \ 65536
 cols(0) = c
 For b = 1 To 255
  r = (255 - r1) * (b - 1&) / 254 + r1
  gg = g1 - g1 * (b - 1&) / 254
  bb = b1 - b1 * (b - 1&) / 254
  cols(b) = RGB(r, gg, bb)         ' BackColor -> Red
  cols(b + 255) = RGB(255, b, 0) ' Red       -> Yellow
  cols(b + 510) = RGB(255, 255, 0)  ' Yellow
  'Picture1.Line (0, b)-Step(200, 0), cols(b)
 Next b

End Sub

Public Sub DoFire()

On Error Resume Next

Dim hdc1 As Long
hdc1 = Picture1.hdc
Picture1.FillStyle = 0


XO = 0: YO = 0 'dyVal
Static secoundCall As Integer

If Not secoundCall Then
 secoundCall = True
 Randomize Timer
 Static k As Single, colMax As Integer
 Call SetColors
 colMax = 255 * 3
 Picture1.PSet (0, 0), Picture1.Point(0, 0)
End If


  For x = 1 To dxVal - 1
    a = 200 + Rnd * colMax: If a > colMax Then a = colMax
    flame(x, dyVal + 1) = a
  Next x
   
   
  For x = 1 To dxVal
    For y = dyVal To 0 Step -1
      a = (flame(x - 1, y) + flame(x, y + 1) + flame(x + 1, y + 1)) / 3 - (ColDecreaseVal + ColDecMS) '* 3 ' '3.5
      If a < 0 Then a = 0 'a = Rnd * 400
      flame(x, y) = a
    Next y
  Next x

  'Picture1.Line (XO, YO)-Step(dxVal - 1, dyVal - 1), 0, BF
  yy = YO + dyVal - 1
  For y = dyVal - 1 To 0 Step -1
    xx = XO
    For x = 1 To dxVal
      Col = flame(x, y)
      SetPixel hdc1, xx, yy, cols(Col)
      xx = xx + 1 '1
    Next x
    yy = yy - 1 '1
  Next y

        
If Picture1.Visible Then Picture1.Refresh
PictureBig.PaintPicture Picture1.Image, 0, 0, ScaleWidth, ScaleHeight, 0, 0, dxVal, dyVal
If Err Then MsgBox "DoFire-4-Err:" + Err.Description: Timer1.Enabled = False

Call PrintText


Exit Sub '--------------------------------------------

DoFireReverse:

ner = 252
Do

  For x = 1 To dxVal - 1
    flame(x, dyVal + 1) = CInt(Rnd * ner)
  Next x

  ner = ner - 10
  If ner <= 0 Then ner = 0

  For x = 1 To dxVal
    For y = dyVal To 0 Step -1
      flame(x, y) = (flame(x - 1, y) + flame(x, y + 1) + flame(x + 1, y + 1)) / 3 - 2.2
    Next y
  Next x

  For y = 0 To dyVal - 1
    For x = 0 To dxVal
      Col = flame(x, y)
      If Col < 0 Then Col = 0
      'Line (XO + X * 1, YO + Y * 1)-Step(1 - 1, 1 - 1), cols(col), BF
    Next x
  Next y

  f = 0
  For x = 1 To dxVal
    For y = dyVal To 0 Step -1
      If flame(x, y) <= 0 Then f = f + 1
    Next y
  Next x

Loop Until f > 2400

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PictureBig,PictureBig,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = PictureBig.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set PictureBig.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PictureBig,PictureBig,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = PictureBig.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    PictureBig.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,Image
Public Property Get Image() As Picture
Attribute Image.VB_Description = "Returns a handle, provided by Microsoft Windows, to a persistent bitmap."
    Set Image = Picture1.Image
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_A_Creator = m_def_A_Creator
    m_MouseSensitive = m_def_MouseSensitive
    m_DY = m_def_DY
    m_DX = m_def_DX
    m_DY = m_def_DY
    m_DispAtDesign = m_def_DispAtDesign
    m_Text = m_def_Text
End Sub

Private Function InterpolateLinear(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x As Single) As Single
Dim y As Single
If x2 <> x1 Then
 y = (y2 - y1) * (x - x1) / (x2 - x1) + y1
Else
 y = (y1 + y2) / 2
End If
InterpolateLinear = y
End Function

Sub PrintText()
If m_Text <> "" Then
 With PictureBig
 .CurrentX = xText: .CurrentY = yText
 PictureBig.Print m_Text;
 End With
End If
End Sub

Public Property Let ToolTipTextString(ByVal vNewValue As String)
PictureBig.ToolTipText = vNewValue
End Property

Public Property Get ToolTipTextString() As String
 ToolTipTextString = PictureBig.ToolTipText
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    Call SetColors
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,1,0,0
Public Property Get A_Creator() As Variant
Attribute A_Creator.VB_Description = "This Control Is Created By: S.Serpooshan\r\n(IRAN-2001) JamAcademic.com/vb/vb.html"
    A_Creator = m_A_Creator
End Property

Public Property Let A_Creator(ByVal New_A_Creator As Variant)
    Err.Raise 382
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get MouseSensitive() As Boolean
Attribute MouseSensitive.VB_Description = "set is mouse sensitive or not"
    MouseSensitive = m_MouseSensitive
End Property

Public Property Let MouseSensitive(ByVal New_MouseSensitive As Boolean)
    m_MouseSensitive = New_MouseSensitive
    PropertyChanged "MouseSensitive"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get DX() As Long
Attribute DX.VB_Description = "real width (act on x resolution)"
    DX = m_DX
End Property

Public Property Let DX(ByVal New_DX As Long)
    m_DX = New_DX
    dxVal = m_DX
    Picture1.Width = dxVal
    Erase flame
    PropertyChanged "DX"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get DY() As Long
Attribute DY.VB_Description = "Real Height pixels (act on y resolution)"
    DY = m_DY
End Property

Public Property Let DY(ByVal New_DY As Long)
    m_DY = New_DY
    dyVal = m_DY
    Picture1.Height = dyVal
    Erase flame
    PropertyChanged "DY"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get DispAtDesign() As Boolean
Attribute DispAtDesign.VB_Description = "Determine if fire is active in design mode"
    DispAtDesign = m_DispAtDesign
End Property

Public Property Let DispAtDesign(ByVal New_DispAtDesign As Boolean)
    m_DispAtDesign = New_DispAtDesign
    If m_DispAtDesign = True Then Timer1.Enabled = True Else If Ambient.UserMode = False Then Timer1.Enabled = False
    PropertyChanged "DispAtDesign"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,3,1,0
Public Property Get FirePicObject() As Object
Attribute FirePicObject.VB_Description = "Fire Picture Object."
    If Ambient.UserMode Then Err.Raise 393
    Set FirePicObject = PictureBig
End Property

Public Property Set FirePicObject(ByVal New_FirePicObject As Object)
    If Ambient.UserMode = False Then Err.Raise 383
    If Ambient.UserMode Then Err.Raise 382
    'Set m_FirePicObject = New_FirePicObject
    'PropertyChanged "FirePicObject"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get Text() As String
Attribute Text.VB_Description = "the String written on Fire"
    Text = m_Text
End Property

Public Property Let Text(ByVal New_Text As String)
    m_Text = New_Text
    Call SetXYText
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PictureBig,PictureBig,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = PictureBig.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set PictureBig.Font = New_Font
    Call SetXYText
    If m_Text <> "" Then DoFire
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=PictureBig,PictureBig,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = PictureBig.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    PictureBig.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Sub SetXYText()
 xText = (ScaleWidth - PictureBig.TextWidth(m_Text)) \ 2
 yText = (ScaleHeight - PictureBig.TextHeight(m_Text)) \ 2
End Sub
