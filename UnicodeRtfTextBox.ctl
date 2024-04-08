VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.UserControl UnicodeRtfTextBox 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   2340
   ScaleWidth      =   3990
   ToolboxBitmap   =   "UnicodeRtfTextBox.ctx":0000
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   180
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   540
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1085
      _Version        =   393217
      BorderStyle     =   0
      HideSelection   =   0   'False
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"UnicodeRtfTextBox.ctx":0532
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RtbFix 
      Height          =   615
      Left            =   600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1380
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1085
      _Version        =   393217
      Enabled         =   0   'False
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"UnicodeRtfTextBox.ctx":05B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "UnicodeRtfTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' ALL Methods, Events, & Properties passed through with the only exception of the data binding properties.
' There are no methods or events associated with the data binding.
' The properties not passed through are:
'
'           DataBindings
'           DataField
'           DataFormat
'           DataMembers
'           DataSource
'
Option Explicit
'
Public Enum AboutConst
    [Double Click Me] = 1
End Enum
Public Enum MultiLineConst
    [True (read only)] = True
End Enum
Public Enum ScrollBarsConst
    [Both (read only)] = rtfBoth
End Enum
Private Type GETTEXTEX
    cb As Long
    flags As Long
    codepage As Long
    lpDefaultChar As Long
    lpUsedDefChar As Long
End Type
Private Type GETTEXTLENGTHEX
    flags As Long
    codepage As Long
End Type
Private Type SETTEXTEX
    flags As Long
    codepage As Long
End Type
'
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SendMessageWLng Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'
Private Const WS_HSCROLL As Long = &H100000
Private Const WS_VSCROLL As Long = &H200000
Private Const ES_AUTOHSCROLL As Long = &H80&
Private Const ES_AUTOVSCROLL As Long = &H40&
Private Const WS_BORDER  As Long = &H800000
Private Const WSEX_3D As Long = &H200&
'
Private Const GWL_STYLE As Long = &HFFFFFFF0
Private Const GWL_EXSTYLE As Long = &HFFFFFFEC
Private Const SWP_NOSIZE  As Long = &H1&
Private Const SWP_NOMOVE  As Long = &H2&
Private Const SWP_NOZORDER  As Long = &H4&
Private Const SWP_FRAMECHANGED  As Long = &H20&
'
Private Const EM_SETTEXTEX As Long = &H461
Private Const RTBC_DEFAULT As Long = 0&
Private Const CP_UNICODE As Long = 1200&
Private Const EM_GETTEXTEX As Long = &H45E
Private Const EM_GETTEXTLENGTHEX As Long = &H45F
Private Const GTL_USECRLF As Long = 1&
Private Const GTL_PRECISE As Long = 2&
Private Const GTL_NUMCHARS As Long = 8&
Private Const GT_USECRLF As Long = 1&
'
Private Const CF_TEXT = 1
Private Const CF_UNICODETEXT = 13
Private Const CF_RICHTEXT = 49448
Private Const CF_RICHTEXTNOOBJECTS = 49453
'
Private EditingInIde As Boolean
Private PatchingUnicodeClipboard As Boolean
Private SecondsSinceLastEditActivity As Single
Private TabStaysInsideOfControl As Boolean
'

'*********************************
' Events passed through.
'*********************************

Event Click()
Event Change()
Event DblClick()
'Event DragDrop(source As Control, X As Single, Y As Single) ' Done by user control.
'Event DragOver(source As Control, X As Single, Y As Single, State As Integer) ' Done by user control.
'Event GotFocus() ' Done by user control.
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
'Event LostFocus() ' Done by user control.
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
Event SelChange()
Event Validate(Cancel As Boolean)
'

Private Sub RTB_Click()
    RaiseEvent Click
End Sub

Private Sub RTB_Change()
    RaiseEvent Change
End Sub

Private Sub RTB_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
    If (KeyCode = vbKeyV) And (Shift = vbCtrlMask) Then CheckUnicodeFormat ' Let KeyCode pass through on this one.
    '
    If (KeyCode = vbKeyB) And (Shift = vbCtrlMask) Then SelBold = Not SelBold: KeyCode = 0
    If (KeyCode = vbKeyI) And (Shift = vbCtrlMask) Then SelItalic = Not SelItalic: KeyCode = 0
    If (KeyCode = vbKeyU) And (Shift = vbCtrlMask) Then SelUnderline = Not SelUnderline: KeyCode = 0
    If (KeyCode = vbKeyUp) And (Shift = vbCtrlMask) Then SelFontSize = SelFontSize + 1: KeyCode = 0
    If (KeyCode = vbKeyDown) And (Shift = vbCtrlMask) Then SelFontSize = SelFontSize - 1: KeyCode = 0
    If (KeyCode = vbKeyTab) And (Shift = 0) Then SelText = vbTab: KeyCode = 0
    If (KeyCode = vbKeyB) And (Shift = vbAltMask) Then SelBullet = Not SelBullet: KeyCode = 0
    If (KeyCode = vbKeyI) And (Shift = vbAltMask) Then SelIndent = SelIndent + 45: KeyCode = 0
    If (KeyCode = vbKeyI) And (Shift = (vbAltMask + vbShiftMask)) Then SelIndent = SelIndent - 45: KeyCode = 0
    If (KeyCode = vbKeyH) And (Shift = vbAltMask) Then SelHangingIndent = SelHangingIndent + 45: KeyCode = 0
    If (KeyCode = vbKeyH) And (Shift = (vbAltMask + vbShiftMask)) Then SelHangingIndent = SelHangingIndent - 45: KeyCode = 0
    If (KeyCode = vbKeyS) And (Shift = vbCtrlMask) Then SelCharOffset = SelCharOffset + 20: KeyCode = 0
    If (KeyCode = vbKeyS) And (Shift = (vbCtrlMask + vbShiftMask)) Then SelCharOffset = SelCharOffset - 20: KeyCode = 0
End Sub

Private Sub RTB_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub RTB_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub RTB_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CheckUnicodeFormat
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub RTB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub RTB_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub RTB_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub RTB_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub RTB_OLEDragOver(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub RTB_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub RTB_OLESetData(Data As RichTextLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub RTB_OLEStartDrag(Data As RichTextLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub RTB_SelChange()
    SecondsSinceLastEditActivity = Timer
    RaiseEvent SelChange
End Sub

Private Sub RTB_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'*********************************
' Methods passed through.
'*********************************

Public Function Find(sString As String, Optional iStart As Variant, Optional iEnd As Variant, Optional Options As RichTextLib.FindConstants = 0)
    Find = RTB.Find(sString, iStart, iEnd, Options)
End Function

Public Function GetLineFromChar(iCharPos As Long)
    GetLineFromChar = RTB.GetLineFromChar(iCharPos)
End Function

Public Sub LoadFile(sPathName As String, Optional sFileType As RichTextLib.LoadSaveConstants = rtfRTF)
    RTB.LoadFile sPathName, sFileType
End Sub

Public Sub Refresh()
    RTB.Refresh
End Sub

Public Sub SaveFile(sPathName As String, Optional sFileType As RichTextLib.LoadSaveConstants = rtfRTF)
    RTB.SaveFile sPathName, sFileType
End Sub

Public Sub SelPrint(hDC As Long)
    RTB.SelPrint hDC
End Sub

Public Sub Span(sCharacterSet As String, Optional bForward As Variant, Optional bNegate As Variant)
    RTB.Span sCharacterSet, bForward, bNegate
End Sub

Public Sub Upto(sCharacterSet As String, Optional bForward As Variant, Optional bNegate As Variant)
    RTB.Upto sCharacterSet, bForward, bNegate
End Sub

'*********************************
' Deal with Properties stuff.
'*********************************

Private Sub UserControl_InitProperties()
    RichTextBoxUnicodeText(RTB.hWnd) = sHelloWorldRussian
    PropertyChanged "TextRTF"
    TabStaysInsideOfControl = True ' It's the default so no need for notification.
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' Just as a note, not all of the properties actually go into the property bag.
    ' There is no need to put runtime-only properties in the property bag.
    ' This is also true of the properties that just change while in 'Edit' mode.
    Set RTB.MouseIcon = PropBag.ReadProperty("MouseIcon", LoadPicture())
    RTB.MousePointer = PropBag.ReadProperty("MousePointer", rtfDefault)
    RTB.OleDragMode = PropBag.ReadProperty("OleDragMode", rtfOLEDragAutomatic)
    RTB.OleDropMode = PropBag.ReadProperty("OleDropMode", rtfOLEDropAutomatic)
    RTB.HideSelection = PropBag.ReadProperty("HideSelection", False)
    TabStaysInsideOfControl = PropBag.ReadProperty("TabStaysInside", True)
    RTB.CausesValidation = PropBag.ReadProperty("CausesValidation", False)
    Set RTB.Font = PropBag.ReadProperty("Font", RTB.Font) ' Be sure to do this before text is set.
    RTB.TextRTF = PropBag.ReadProperty("TextRTF", "") ' This way we save the formatting as well as the Unicode.
    RTB.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    RTB.RightMargin = PropBag.ReadProperty("RightMargin", 0)
    RTB.BulletIndent = PropBag.ReadProperty("BulletIndent", 0)
    RTB.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", True)
    RTB.FileName = PropBag.ReadProperty("FileName", "")
    RTB.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    RTB.Enabled = PropBag.ReadProperty("Enabled", True)
    RTB.Locked = PropBag.ReadProperty("Locked", False)
    RtbBorderStyle = PropBag.ReadProperty("BorderStyle", rtfFixedSingle)
    RtbAppearance = PropBag.ReadProperty("Appearance", rtfThreeD)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    DesignTimeEditMode = False ' If we close the form while editing, this will clean it up.  (Lost focus may not fire.)
    '
    ' Just as a note, not all of the properties actually go into the property bag.
    ' There is no need to put runtime-only properties in the property bag.
    ' This is also true of the properties that just change while in 'Edit' mode.
    PropBag.WriteProperty "MouseIcon", RTB.MouseIcon, LoadPicture()
    PropBag.WriteProperty "MousePointer", RTB.MousePointer, rtfDefault
    PropBag.WriteProperty "OleDragMode", RTB.OleDragMode, rtfOLEDragAutomatic
    PropBag.WriteProperty "OleDropMode", RTB.OleDropMode, rtfOLEDropAutomatic
    PropBag.WriteProperty "HideSelection", RTB.HideSelection, False
    PropBag.WriteProperty "TabStaysInside", TabStaysInsideOfControl, True
    PropBag.WriteProperty "CausesValidation", RTB.CausesValidation, False
    PropBag.WriteProperty "Font", RTB.Font
    PropBag.WriteProperty "TextRTF", RTB.TextRTF, "" ' This way we save the formatting as well as the Unicode.
    PropBag.WriteProperty "MaxLength", RTB.MaxLength, 0
    PropBag.WriteProperty "RightMargin", RTB.RightMargin, 0
    PropBag.WriteProperty "BulletIndent", RTB.BulletIndent, 0
    PropBag.WriteProperty "AutoVerbMenu", RTB.AutoVerbMenu, True
    PropBag.WriteProperty "FileName", RTB.FileName, ""
    PropBag.WriteProperty "BackColor", RTB.BackColor, &H8000000F
    PropBag.WriteProperty "Enabled", RTB.Enabled, True
    PropBag.WriteProperty "Locked", RTB.Locked, False
    PropBag.WriteProperty "BorderStyle", RtbBorderStyle, rtfFixedSingle
    PropBag.WriteProperty "Appearance", RtbAppearance, rtfThreeD
End Sub

'*********************************
' The PropertyBag Properties.
'*********************************

Public Property Get About() As AboutConst
    If RunTime Then Exit Property
    About = [Double Click Me]
End Property

Public Property Let About(prop As AboutConst)
    If RunTime Then Exit Property ' Only at design time.
    MsgBox "GPLv3 licensed.  Right-click for 'Edit' mode." & vbCrLf & vbCrLf & _
           "MultiLine always 'True', ScrollBars always 'both'." & vbCrLf & vbCrLf & _
           "For BorderStyle & Appearance changes to show," & vbCrLf & _
           "close and re-open form (or run project)." & vbCrLf & vbCrLf & _
           "Sel... properties are for 'selected' text." & vbCrLf & _
           "Works at either design or runtime." & vbCrLf & vbCrLf & _
           "At runtime, use TextUnicode property to set Unicode." & vbCrLf & _
           "In 'Edit' mode at design time just paste." & vbCrLf & vbCrLf & _
           "SelRTF & TextRTF can also be used, but they must be" & vbCrLf & _
           "correctly formatted RTF strings." & vbCrLf & vbCrLf & _
           "Some shortcut (text editing) keys:" & vbCrLf & _
           "    Ctrl-B (toggle bold), Ctrl-I (toggle italic), " & vbCrLf & _
           "    Ctrl-U (toggle underline), Alt-B (toggle bullet)," & vbCrLf & _
           "    Ctrl-Up (increase font), Ctrl-Down (decrease font)," & vbCrLf & _
           "    Alt-I (increase indent), Alt-Shift-I (decrease)," & vbCrLf & _
           "    Alt-H (increase hang), Alt-Shift-H (decrease)," & vbCrLf & _
           "    Ctrl-S (more superscript), Ctrl-Shift-S (less, sub)", _
           vbInformation, "UnicodeTextBox"
End Property

Public Property Get MouseIcon() As Picture
    Set MouseIcon = RTB.MouseIcon
End Property

Public Property Set MouseIcon(prop As Picture)
    Set RTB.MouseIcon = prop
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As RichTextLib.MousePointerConstants
    OleDragMode = RTB.MousePointer
End Property

Public Property Let MousePointer(prop As RichTextLib.MousePointerConstants)
    RTB.MousePointer = prop
    PropertyChanged "MousePointer"
End Property

Public Property Get OleDragMode() As RichTextLib.OLEDragConstants
    OleDragMode = RTB.OleDragMode
End Property

Public Property Let OleDragMode(prop As RichTextLib.OLEDragConstants)
    RTB.OleDragMode = prop
    PropertyChanged "OleDragMode"
End Property

Public Property Get OleDropMode() As RichTextLib.OLEDropConstants
    OleDropMode = RTB.OleDropMode
End Property

Public Property Let OleDropMode(prop As RichTextLib.OLEDropConstants)
    RTB.OleDropMode = prop
    PropertyChanged "OleDropMode"
End Property

Public Property Get HideSelection() As Boolean
    HideSelection = RTB.HideSelection
End Property

Public Property Let HideSelection(prop As Boolean)
    RTB.HideSelection = prop
    PropertyChanged "HideSelection"
End Property

Public Property Get TabStaysInside() As Boolean
    TabStaysInside = TabStaysInsideOfControl
End Property

Public Property Let TabStaysInside(prop As Boolean)
    TabStaysInsideOfControl = prop
    PropertyChanged "TabStaysInside"
End Property

Public Property Get CausesValidation() As Boolean
    CausesValidation = RTB.CausesValidation
End Property

Public Property Let CausesValidation(prop As Boolean)
    RTB.CausesValidation = prop
    PropertyChanged "CausesValidation"
End Property

Public Property Get Font() As Font
    Set Font = RTB.Font
End Property

Public Property Set Font(prop As Font)
    Dim s As String
    '
    ' Not sure why, but changing the font sometimes messes up the Unicode, so we save and restore it.
    's = RichTextBoxUnicodeText(RTB.hWnd)
    s = RTB.TextRTF
    Set RTB.Font = prop
    'RichTextBoxUnicodeText(RTB.hWnd) = s
    RTB.TextRTF = s
    PropertyChanged "Font"
End Property

Public Property Get TextUnicode() As String
Attribute TextUnicode.VB_Description = "Unicode text.  Use 'Edit' to change this property.  At runtime, just assign a string to it with Unicode and it'll show as Unicode."
Attribute TextUnicode.VB_MemberFlags = "200"
    If RunTime Then
        TextUnicode = RichTextBoxUnicodeText(RTB.hWnd)
    Else
        TextUnicode = "(Unicode, right-click control & edit, read-write at runtime)" ' Just show this in the Properties Window.
    End If
End Property

Public Property Let TextUnicode(prop As String)
    If Not RunTime Then Exit Property ' Only allowed in runtime, not from Properties Window.  See next procedure.
    RichTextBoxUnicodeText(RTB.hWnd) = prop
End Property

Private Sub RTB_LostFocus() ' Not technically a property.
    If RunTime Then Exit Sub ' Don't need to worry about PropertyBag at runtime.
    If PatchingUnicodeClipboard Then Exit Sub
    ' It's this one that takes care of saving the contents of RTB into the "TextRTF" property.
    ' If "Edit" was selected in design time, this will take care of saving changes.
    ' Don't need to call RichTextBoxUnicodeText because user is changing it by typing.
    DesignTimeEditMode = False ' This takes care of property changes.
End Sub

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Read only."
    MaxLength = RTB.MaxLength
End Property

Public Property Let MaxLength(prop As Long)
    RTB.MaxLength = prop
    PropertyChanged "MaxLength"
End Property

Public Property Get RightMargin() As Long
    RightMargin = RTB.RightMargin
End Property

Public Property Let RightMargin(prop As Long)
    RTB.RightMargin = prop
    PropertyChanged "RightMargin"
End Property

Public Property Get BulletIndent() As Long
Attribute BulletIndent.VB_Description = "Returns or sets the amount of indent used in the control when SelBullet is set to True."
    BulletIndent = RTB.BulletIndent
End Property

Public Property Let BulletIndent(prop As Long)
    RTB.BulletIndent = prop
    PropertyChanged "BulletIndent"
End Property

Public Property Get AutoVerbMenu() As Boolean
    AutoVerbMenu = RTB.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(prop As Boolean)
    RTB.AutoVerbMenu = prop
    PropertyChanged "AutoVerbMenu"
End Property

Public Property Get FileName() As String
Attribute FileName.VB_Description = "Must be a fully specified filename."
    FileName = RTB.FileName
End Property

Public Property Let FileName(prop As String)
    On Error Resume Next ' Must be a valid filename (and path).  Just ignore errors.
    RTB.FileName = prop
    PropertyChanged "FileName"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = RTB.BackColor
End Property

Public Property Let BackColor(prop As OLE_COLOR)
    RTB.BackColor = prop
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean
    Enabled = RTB.Enabled
End Property

Public Property Let Enabled(prop As Boolean)
    RTB.Enabled = prop
    PropertyChanged "Enabled"
End Property

Public Property Get Locked() As Boolean
    Locked = RTB.Locked
End Property

Public Property Let Locked(prop As Boolean)
    RTB.Locked = prop
    PropertyChanged "Locked"
End Property

Public Property Get MultiLine() As MultiLineConst
    MultiLine = True
End Property

Public Property Let MultiLine(prop As MultiLineConst)
    If RunTime Then MsgBox "Multiline is read-only.", vbCritical, "UnicodeTextbox": Exit Property ' Read only.
    ' This is in the PropertyBag but they never change.
End Property

Public Property Get ScrollBars() As ScrollBarsConst
Attribute ScrollBars.VB_Description = "Read only."
    ScrollBars = rtfBoth
End Property

Public Property Let ScrollBars(prop As ScrollBarsConst)
    If RunTime Then MsgBox "ScrollBars is read-only.", vbCritical, "UnicodeTextbox": Exit Property ' Read only at runtime.
    ' This is in the PropertyBag but they never change.
End Property

Public Property Get BorderStyle() As RichTextLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "For BorderStyle & Appearance changes to show, you must close and re-open the form (or run the project)."
    BorderStyle = RtbBorderStyle
End Property

Public Property Let BorderStyle(prop As RichTextLib.BorderStyleConstants)
    If RunTime Then MsgBox "BorderStyle read-only at runtime.", vbCritical, "UnicodeTextbox": Exit Property ' Read only at runtime.
    RtbBorderStyle = prop
    PropertyChanged "BorderStyle"
End Property

Public Property Get Appearance() As RichTextLib.AppearanceConstants
Attribute Appearance.VB_Description = "For BorderStyle & Appearance changes to show, you must close and re-open the form (or run the project)."
    Appearance = RtbAppearance
End Property

Public Property Let Appearance(prop As RichTextLib.AppearanceConstants)
    If RunTime Then MsgBox "Appearance read-only at runtime.", vbCritical, "UnicodeTextbox": Exit Property ' Read only at runtime.
    RtbAppearance = prop
    PropertyChanged "Appearance"
End Property

'*********************************
' The Non-PropertyBag Properties.
'*********************************

Public Property Get Text() As String
    Text = RTB.Text
End Property

Public Property Let Text(prop As String)
    RTB.Text = prop
End Property

Public Property Get SelTextUnicode() As String
    If Not RunTime Then
        SelTextUnicode = "(Runtime only, just edit & paste in design mode)"
    Else
        ' We must convert it to RTF before it can be correctly returned.
        RtbFix.Enabled = True
        RtbFix.Locked = False
        On Error Resume Next ' In case there's not a selection point.
        RtbFix.TextRTF = RTB.SelRTF ' Stuff it in the "Fix" RTB.
        SelTextUnicode = RichTextBoxUnicodeText(RtbFix.hWnd)
        On Error GoTo 0
        RtbFix.Enabled = False
        RtbFix.Locked = True
    End If
End Property

Public Property Let SelTextUnicode(prop As String)
    If Not RunTime Then Exit Property
    ' We must convert it to RTF before it can be correctly inserted.
    RtbFix.Enabled = True
    RtbFix.Locked = False
    RichTextBoxUnicodeText(RtbFix.hWnd) = prop ' Stuff it in the "Fix" RTB.
    ' Strip off the EOL terminator.
    On Error Resume Next ' In case there's not a selection point.
    RtbFix.SelStart = 0
    RtbFix.SelLength = Len(prop)
    RTB.SelRTF = RtbFix.SelRTF
    On Error GoTo 0
    RtbFix.Enabled = False
    RtbFix.Locked = True
End Property

Public Property Get DisableNoScroll() As Boolean ' Not implemented
    DisableNoScroll = False
End Property

Public Property Let DisableNoScroll(prop As Boolean)
    ' Not implemented.
End Property

Public Property Get SelTabsString() As String
    Dim i As Integer
    Dim s As String
    '
    On Error Resume Next ' Possibly no selection point or bad index.
    If IsNull(RTB.SelTabCount) Then Exit Property
    If RTB.SelTabCount = 0 Then Exit Property
    s = Format$(RTB.SelTabs(0))
    For i = 1 To RTB.SelTabCount - 1
        s = s & "," & Format$(RTB.SelTabs(i))
    Next i
    SelTabsString = s
End Property

Public Property Let SelTabsString(prop As String)
    Dim iCnt As Integer
    Dim index As Long
    Dim iStart As Long
    Dim iPtr As Long
    Dim k As Long
    '
    On Error Resume Next ' Possibly no selection point or bad index.
    If Len(Trim$(prop)) = 0 Then
        RTB.SelTabCount = 0
    Else
        iStart = 1
        Do ' Count how many tabs we're setting.
            iCnt = iCnt + 1
            iPtr = InStr(iStart, prop, ",")
            If iPtr = 0 Then Exit Do
            iStart = iPtr + 1
        Loop
        RTB.SelTabCount = iCnt ' Make room for tabs.
        iStart = 1
        For index = 1 To iCnt - 1 ' Set the tabs.
            iPtr = InStr(iStart, prop, ",")
            RTB.SelTabs(index - 1) = CInt(Mid$(prop, iStart, iPtr - iStart))
            iStart = iPtr + 1
        Next index
        RTB.SelTabs(iCnt - 1) = CInt(Mid$(prop, iStart))
    End If
End Property

Public Property Get SelTabs(index As Integer) As Single
    On Error Resume Next ' Possibly no selection point or bad index.
    If IsNull(RTB.SelTabs(index)) Then SelTabs = 0 Else SelTabs = RTB.SelTabs(index)
End Property

Public Property Let SelTabs(index As Integer, prop As Single)
    On Error Resume Next ' Possibly no selection point or bad index.
    RTB.SelTabs(index) = prop
End Property

Public Property Get SelTabCount() As Integer
    If IsNull(RTB.SelTabCount) Then SelTabCount = 0 Else SelTabCount = RTB.SelTabCount
End Property

Public Property Let SelTabCount(prop As Integer)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelTabCount = prop
End Property

Public Property Get SelProtected() As Boolean
    If IsNull(RTB.SelProtected) Then SelProtected = False Else SelProtected = RTB.SelProtected
End Property

Public Property Let SelProtected(prop As Boolean)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelProtected = prop
End Property

Public Property Get SelText() As String
    If IsNull(RTB.SelText) Then SelText = "" Else SelText = RTB.SelText
End Property

Public Property Let SelText(prop As String)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelText = prop
End Property

Public Property Get SelLength() As Long
    If IsNull(RTB.SelLength) Then SelLength = 0 Else SelLength = RTB.SelLength
End Property

Public Property Let SelLength(prop As Long)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelLength = prop
End Property

Public Property Get SelStart() As Long
    If IsNull(RTB.SelStart) Then SelStart = 0 Else SelStart = RTB.SelStart
End Property

Public Property Let SelStart(prop As Long)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelStart = prop
End Property

Public Property Get SelHangingIndent() As Integer
    If IsNull(RTB.SelHangingIndent) Then SelHangingIndent = 0 Else SelHangingIndent = RTB.SelHangingIndent
End Property

Public Property Let SelHangingIndent(prop As Integer)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelHangingIndent = prop
End Property

Public Property Get SelIndent() As Integer
    If IsNull(RTB.SelIndent) Then SelIndent = 0 Else SelIndent = RTB.SelIndent
End Property

Public Property Let SelIndent(prop As Integer)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelIndent = prop
End Property

Public Property Get SelRightIndent() As Integer
    If IsNull(RTB.SelRightIndent) Then SelRightIndent = 0 Else SelRightIndent = RTB.SelRightIndent
End Property

Public Property Let SelRightIndent(prop As Integer)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelRightIndent = prop
End Property

Public Property Get SelFontSize() As Integer
    If IsNull(RTB.SelFontSize) Then SelFontSize = 0 Else SelFontSize = RTB.SelFontSize
End Property

Public Property Let SelFontSize(prop As Integer)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelFontSize = prop
End Property

Public Property Get SelFontName() As String
    If IsNull(RTB.SelFontName) Then SelFontName = RTB.Font.Name Else SelFontName = RTB.SelFontName
End Property

Public Property Let SelFontName(prop As String)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelFontName = prop
End Property

Public Property Get SelColor() As OLE_COLOR
    If IsNull(RTB.SelColor) Then SelColor = vbBlack Else SelColor = RTB.SelColor
End Property

Public Property Let SelColor(prop As OLE_COLOR)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelColor = prop
End Property

Public Property Get SelAlignment() As RichTextLib.SelAlignmentConstants
Attribute SelAlignment.VB_Description = "Returns or sets a value that controls the alignment of the paragraphs in a control."
    If IsNull(RTB.SelAlignment) Then SelAlignment = rtfLeft Else SelAlignment = RTB.SelAlignment
End Property

Public Property Let SelAlignment(prop As RichTextLib.SelAlignmentConstants)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelAlignment = prop
End Property

Public Property Get SelBold() As Boolean
    If IsNull(RTB.SelBold) Then SelBold = False Else SelBold = RTB.SelBold
End Property

Public Property Let SelBold(prop As Boolean)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelBold = prop
End Property

Public Property Get SelItalic() As Boolean
    If IsNull(RTB.SelItalic) Then SelItalic = False Else SelItalic = RTB.SelItalic
End Property

Public Property Let SelItalic(prop As Boolean)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelItalic = prop
End Property

Public Property Get SelStrikethru() As Boolean
    If IsNull(RTB.SelStrikethru) Then SelStrikethru = False Else SelStrikethru = RTB.SelStrikethru
End Property

Public Property Let SelStrikethru(prop As Boolean)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelStrikethru = prop
End Property

Public Property Get SelUnderline() As Boolean
    If IsNull(RTB.SelUnderline) Then SelUnderline = False Else SelUnderline = RTB.SelUnderline
End Property

Public Property Let SelUnderline(prop As Boolean)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelUnderline = prop
End Property

Public Property Get SelBullet() As Boolean
Attribute SelBullet.VB_Description = "Returns or sets a value that determines if a paragraph in the control containing the current selection or insertion point has the bullet style."
    If IsNull(RTB.SelBullet) Then SelBullet = False Else SelBullet = RTB.SelBullet
End Property

Public Property Let SelBullet(prop As Boolean)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelBullet = prop
End Property

Public Property Get SelCharOffset() As Integer
Attribute SelCharOffset.VB_Description = "Returns or sets a value that determines whether text in the control appears on the baseline (normal), as a superscript above the baseline, or as a subscript below the baseline."
    If IsNull(RTB.SelCharOffset) Then SelCharOffset = 0 Else SelCharOffset = RTB.SelCharOffset
End Property

Public Property Let SelCharOffset(prop As Integer)
    On Error Resume Next ' Possibly no selection point.
    RTB.SelCharOffset = prop
End Property

Public Property Get TextRTF() As String
Attribute TextRTF.VB_Description = "SelRTF & TextRTF are only for backwards compatability.  It's better to 'Edit' or paste Unicode directly into control, or use Text property at runtime."
    If Not EditingInIde Then TextRTF = RTB.TextRTF
End Property

Public Property Let TextRTF(prop As String)
    RTB.TextRTF = prop
End Property

Public Property Get SelRTF() As String
Attribute SelRTF.VB_Description = "SelRTF & TextRTF are only for backwards compatability.  It's better to 'Edit' or paste Unicode directly into control, or use Text property at runtime."
    If Not EditingInIde Then SelRTF = RTB.SelRTF
End Property

Public Property Let SelRTF(prop As String)
    RTB.SelRTF = prop
End Property

'*********************************
' Other support procedures.
'*********************************

Private Sub UserControl_Resize() ' Keep sub-controls the same size as the UserControl.
    RTB.Top = 0: RTB.Left = 0: RTB.Width = Width: RTB.Height = Height
End Sub

Private Property Let DesignTimeEditMode(b As Boolean)
    Static RealBackColor As OLE_COLOR
    If b = EditingInIde Then Exit Property
    '
    EditingInIde = b
    If EditingInIde Then
        RealBackColor = RTB.BackColor
        RTB.BackColor = RGB(&HFF, &HD0, &HD0)
        tmr.Enabled = True
        PropertyChanged "" ' Make sure UserControl_WriteProperties fires in case the form is directly closed.
    Else
        RTB.BackColor = RealBackColor
        '
        ' Now make the property window re-read the changed properties.
        tmr.Enabled = False
        PropertyChanged "SelRTF"
        PropertyChanged "TextRTF"
    End If
End Property

Private Sub UserControl_EnterFocus()
    If RunTime Then Exit Sub
    DesignTimeEditMode = True ' The only time this fires in design time is when editing.
End Sub

Private Function RunTime() As Boolean
    RunTime = Ambient.UserMode
End Function

Private Property Let RichTextBoxUnicodeText(hWndRtb As Long, sText As String)
    Dim stUnicode As SETTEXTEX
    '
    stUnicode.flags = RTBC_DEFAULT ' This could be otherwise.
    stUnicode.codepage = CP_UNICODE
    SendMessageWLng hWndRtb, EM_SETTEXTEX, VarPtr(stUnicode), StrPtr(sText)
End Property

Private Property Get RichTextBoxUnicodeText(hWndRtb As Long) As String
    Dim uGTL As GETTEXTLENGTHEX
    Dim uGT As GETTEXTEX
    Dim iChars As Long
    '
    uGTL.flags = GTL_USECRLF Or GTL_PRECISE Or GTL_NUMCHARS
    uGTL.codepage = CP_UNICODE
    iChars = SendMessageWLng(hWndRtb, EM_GETTEXTLENGTHEX, VarPtr(uGTL), 0&)
    '
    uGT.cb = (iChars + 1) * 2
    uGT.flags = GT_USECRLF
    uGT.codepage = CP_UNICODE
    RichTextBoxUnicodeText = String$(iChars, 0&)
    SendMessageWLng hWndRtb, EM_GETTEXTEX, VarPtr(uGT), StrPtr(RichTextBoxUnicodeText)
End Property

Private Function sHelloWorldRussian()
    Static s As String
    If Len(s) = 0 Then s = ChrW$(&H43F) & ChrW$(&H440) & ChrW$(&H438) & ChrW$(&H432) & ChrW$(&H435) & ChrW$(&H442) & " " & ChrW$(&H43C) & ChrW$(&H438) & ChrW$(&H440)
    sHelloWorldRussian = s
End Function

Private Property Get RtbAppearance() As RichTextLib.AppearanceConstants
    Dim lExStyle As Long
    '
    lExStyle = GetWindowLong(RTB.hWnd, GWL_EXSTYLE)
    Select Case True
    Case (WSEX_3D And lExStyle) <> 0
        RtbAppearance = rtfThreeD
    Case Else
        RtbAppearance = rtfFlat
    End Select
End Property

Private Property Let RtbAppearance(iAppearance As RichTextLib.AppearanceConstants)
    Dim lExStyle As Long
    Dim lNew As Long
    Dim lTemp As Long
    '
    lExStyle = GetWindowLong(RTB.hWnd, GWL_EXSTYLE)
    If iAppearance = rtfThreeD Then lNew = lNew Or WSEX_3D
    lTemp = Not (WSEX_3D)    ' All "other" styles.
    lExStyle = lExStyle And lTemp  ' Remove old styles.
    lExStyle = lExStyle Or lNew    ' Add new styles.
    SetWindowLong RTB.hWnd, GWL_EXSTYLE, lExStyle
End Property

Private Property Get RtbBorderStyle() As RichTextLib.BorderStyleConstants
    Dim lStyle As Long
    '
    lStyle = GetWindowLong(RTB.hWnd, GWL_STYLE)
    Select Case True
    Case (WS_BORDER And lStyle) <> 0
        RtbBorderStyle = rtfFixedSingle
    Case Else
        RtbBorderStyle = rtfNoBorder
    End Select
End Property

Private Property Let RtbBorderStyle(iBorderStyle As RichTextLib.BorderStyleConstants)
    Dim lStyle As Long
    Dim lNew As Long
    Dim lTemp As Long
    '
    lStyle = GetWindowLong(RTB.hWnd, GWL_STYLE)
    If iBorderStyle = rtfFixedSingle Then lNew = lNew Or WS_BORDER
    lTemp = Not (WS_BORDER)    ' All "other" styles.
    lStyle = lStyle And lTemp  ' Remove old styles.
    lStyle = lStyle Or lNew    ' Add new styles.
    SetWindowLong RTB.hWnd, GWL_STYLE, lStyle
End Property

Private Sub CheckUnicodeFormat()
    ' The RTB can have RTF pasted into it or Unicode pasted into it, but not both.
    ' It's set for RTF.  Therefore, if pure unicode (with not RTF format) attempts to
    ' get pasted into it, it just goes in as text.  This procedure forces that
    ' Unicode into another RTB and puts it back into the clipboard with an RTF format.
    Dim s As String
    Const VK_CONTROL As Byte = &H11
    Const VK_C_KEY As Byte = &H43
    Const KEYEVENTF_EXTENDEDKEY = &H1
    Const KEYEVENTF_KEYDOWN = &H0
    Const KEYEVENTF_KEYUP = &H2
    '
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        If (IsClipboardFormatAvailable(CF_RICHTEXT) Or IsClipboardFormatAvailable(CF_RICHTEXTNOOBJECTS)) = 0 Then
            PatchingUnicodeClipboard = True
            ' Get Unicode from clipboard.
            s = UniClipboard
            Clipboard.Clear
            RtbFix.Enabled = True
            RtbFix.Visible = True
            RtbFix.Locked = False
            RichTextBoxUnicodeText(RtbFix.hWnd) = s ' Stuff it in the "Fix" RTB.
            ' Now select all text and simulate a ctrl-c.
            RtbFix.SelStart = 0
            RtbFix.SelLength = Len(s)
            RtbFix.SetFocus: DoEvents
            keybd_event VK_CONTROL, 0, KEYEVENTF_KEYDOWN, 0: DoEvents
            keybd_event VK_C_KEY, 0, KEYEVENTF_KEYDOWN, 0: DoEvents
            keybd_event VK_C_KEY, 0, KEYEVENTF_KEYUP, 0: DoEvents
            keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0: DoEvents
            ' Clean up and get out.
            RtbFix.Enabled = False
            RtbFix.Visible = False
            RtbFix.Locked = True
            RTB.SetFocus
            PatchingUnicodeClipboard = False
        End If
    End If
End Sub

Public Property Get UniClipboard() As String
    ' Gets a UNICODE string from the clipboard and puts it in a standard VB string (which is UNICODE).
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
    Dim sUniText As String
    '
    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
            lstrcpy StrPtr(sUniText), iLock
            GlobalUnlock iStrPtr
        End If
        UniClipboard = sUniText
    End If
    CloseClipboard
End Property

Private Sub tmr_Timer()
    If Abs(SecondsSinceLastEditActivity - Timer) > 0.5 Then PropertyChanged "" ' This seems to get it done.
End Sub

