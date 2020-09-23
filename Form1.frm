VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Btn2 
      Caption         =   "Clear document"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox ChcUpper 
      Caption         =   "Uppercase HTML tags"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   0
      Width           =   1335
   End
   Begin MSComctlLib.TreeView Tvw1 
      Height          =   4815
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8493
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.CommandButton Btn1 
      Caption         =   "Whole document in tree"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin RichTextLib.RichTextBox RTbox 
      Height          =   4815
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   8493
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCaretPos Lib _
  "user32" (lpPoint As POINTAPI) As Long

Dim bIntag As Boolean 'working in a tag
Dim ThisTag As String 'this one
Dim Tagstart As Long 'startposition of it
Dim Suggest As String 'our suggestion
Dim bUcaseTags As Boolean 'uppercase or lowercase preference

Dim Alltags As Collection 'all available opening tags (close suggestions are generated)
Dim AllShorts As Collection 'short versions and info whether they require closing

'**********************************************************************
'**********************************************************************
'          FORM CODE:                                          *
'                                                                   *
'                                                                   *
'**********************************************************************
'**********************************************************************


Private Sub Form_Load()
Call GetTags
ChcUpper.Value = 0
Call ChcUpper_Click
RTbox.LoadFile App.Path & "\Demo.htm"
End Sub

Private Sub Form_Resize()
On Error Resume Next
RTbox.Height = Me.Height - RTbox.Top - 400
Tvw1.Height = Me.Height - Tvw1.Top - 400
RTbox.Width = Me.Width - RTbox.Left - 120

End Sub


Private Function GetTCursX() As Long
Dim pt As POINTAPI
GetCaretPos pt
GetTCursX = pt.x
End Function

Private Function GetTCursY() As Long
Dim pt As POINTAPI
GetCaretPos pt
GetTCursY = pt.y
End Function


Private Sub ChcUpper_Click()
Dim L As Long
If bIntag = True Then Call AcceptSuggest
If ChcUpper.Value = 0 Then
    bUcaseTags = False
    For L = 1 To Alltags.Count
        Alltags(L).FullTag = LCase$(Alltags(L).FullTag)
    Next
Else
    bUcaseTags = True
    For L = 1 To Alltags.Count
        Alltags(L).FullTag = UCase$(Alltags(L).FullTag)
    Next
End If
End Sub

'**********************************************************************
'**********************************************************************
'          Read tags from text file and fill the collections:      *
'  This requires Closing : "<HEAD>"                                 *
'  This doesn't "<BR>N"  -there's a character after the ">" in the text file *
'**********************************************************************
'**********************************************************************


Private Sub GetTags()
Dim sTag As String
Dim NewTag As cHtmlTag
Dim NewShort As cShort
Set Alltags = Nothing
Set Alltags = New Collection
Set AllShorts = Nothing
Set AllShorts = New Collection

Open App.Path & "\Tags.txt" For Input Shared As #1
While Not EOF(1)
    Line Input #1, sTag
    sTag = Trim(sTag)
    If sTag <> "" Then
        Set NewTag = New cHtmlTag
        If Right(sTag, 1) = ">" Then
            NewTag.FullTag = sTag
             NewTag.CloseRequired = True
        Else
            NewTag.FullTag = Left(sTag, InStr(sTag, ">"))
             NewTag.CloseRequired = False
        End If
        If bUcaseTags = True Then
            NewTag.FullTag = UCase$(NewTag.FullTag)
        Else
            NewTag.FullTag = LCase$(NewTag.FullTag)
        End If
        NewTag.ShortTag = Shortened(NewTag.FullTag, False)
        Alltags.Add NewTag
        On Error Resume Next
        If AllShorts(NewTag.ShortTag) Is Nothing Then
            Set NewShort = New cShort
            NewShort.ShortTag = NewTag.ShortTag
            NewShort.CloseRequired = NewTag.CloseRequired
            AllShorts.Add NewShort, NewShort.ShortTag
            Set NewShort = Nothing
        End If
        On Error GoTo 0
        Set NewTag = Nothing
    End If
Wend
Close #1
End Sub

'show doc structure in tree:
Private Sub Btn1_Click()
Call PrevOpenTag(False)
End Sub

'clear document:
Private Sub Btn2_Click()
If bIntag = True Then Me.UseTag = ""
RTbox.Text = ""
RTbox.SetFocus
End Sub

'get done with automatics before clicking nodes:
Private Sub Tvw1_GotFocus()
If bIntag = True Then Call AcceptSuggest
End Sub

'This one selects the opening node when clicked in the threeview:

Private Sub Tvw1_NodeClick(ByVal Node As MSComctlLib.Node)
'meCaption = Node.Key & " - " & Node.Tag
RTbox.SelStart = Val(Mid(Node.Key, 2))
RTbox.SelLength = Val(Mid(Node.Key, InStr(Node.Key, "L") + 1))
End Sub

'**********************************************************************
'**********************************************************************
'          Rich textbox code:                                       *
'                                                                   *
'                                                                   *
'**********************************************************************
'**********************************************************************



Private Sub RTbox_KeyDown(KeyCode As Integer, Shift As Integer)
'Decide what to do when a key is pressed:
With RTbox
Select Case KeyCode

Case 8 'backspace
    If bIntag = False Then Exit Sub
    If .SelStart = Tagstart Then
        .SelStart = Tagstart - 1
        .SelLength = Len(Suggest) + 1
        .SelText = ""
        Unload FormSugest
        bIntag = False
        KeyCode = 0
    Else
        .SelText = ""
    End If
Case 27, 46 'escape, delete
    If bIntag = False Then Exit Sub
    'We are not having a tag here, remove the suggestion:
    .SelStart = Tagstart - 1
    .SelLength = Len(Suggest) + 1
    .SelText = ""
    Unload FormSugest
    bIntag = False
    KeyCode = 0
Case 40 'down arrow
    If bIntag = False Then Exit Sub
'    arrow down the suggestion list; select #1 and set focus to it
    If FormSugest.List.ListCount > 0 Then
        KeyCode = 0
        FormSugest.List.ListIndex = FormSugest.List.ListIndex + 1
        FormSugest.SetFocus
    End If
Case 9, 13, 39 'tab , R, enter, arrow
    If bIntag = False Then Exit Sub
    'accept suggested tag, complete it
    Call AcceptSuggest
    KeyCode = 0
Case 226 ' < >
    If Shift = 0 Then '<
        'START A NEW TAG
        RTbox.SelText = ""
        bIntag = True 'OK, we're finally working on a tag
        Tagstart = .SelStart + 1 'the tag starts here, remember this
        'position the Suggestion list where the caret is:
        FormSugest.Left = Me.Left + .Left + ScaleX(GetTCursX, vbPixels, vbTwips) + 220
        FormSugest.Top = Me.Top + .Top + ScaleY(GetTCursY, vbPixels, vbTwips) + 600
        FormSugest.Show vbModeless, Me
        'and return the focus to the rich text box:
        Me.SetFocus
    ElseIf Shift = 1 Then '>
        'Tag completed as written
        Unload FormSugest
        bIntag = False
    End If

Case Else

End Select
End With
End Sub

'somewhere -probably else- is clicked, so accept/complete the suggested tag:

Private Sub RTbox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If bIntag = False Then Exit Sub
Call AcceptSuggest
End Sub

'when in tag, convert key entries to upper- or lowercase:

Private Sub RTbox_KeyPress(KeyAscii As Integer)
If bIntag = False Then Exit Sub
Select Case KeyAscii
Case 65 To 90
    If bUcaseTags = False Then KeyAscii = KeyAscii + 32
Case 97 To 122
    If bUcaseTags = True Then KeyAscii = KeyAscii - 32
Case Else
End Select
End Sub

'entry is done, see if there are matching suggestions:

Private Sub RTbox_KeyUp(KeyCode As Integer, Shift As Integer)
Dim S As Long
If bIntag = False Then Exit Sub
S = RTbox.SelStart 'compared to Tagstart, this shows the length of entry
ThisTag = Mid(RTbox.Text, Tagstart, S - Tagstart + 1)
Call FindMatch
If Suggest <> "" Then Call ShowSuggest(S) 'S says "and select from this position"
End Sub


'**********************************************************************
'**********************************************************************
'          Code to handle the Tag suggestions for autocomplete:     *
'                                                                   *
'                                                                   *
'**********************************************************************
'**********************************************************************'

Private Sub FindMatch()
Dim L As Long 'Counter for tag collection Alltags
Dim i As Integer 'length of currently entered tag

FormSugest.List.Visible = False
Screen.MousePointer = vbHourglass
i = Len(ThisTag)
Suggest = ""
FormSugest.List.Clear

'First see if there's an existing tag in the document that might close here:
If i = 1 Then Suggest = PrevOpenTag(True)
If Suggest <> "" Then FormSugest.List.AddItem Suggest

'Then loop tag collection for matches
For L = 1 To Alltags.Count
    If UCase$(Alltags(L).FullTag) Like UCase$(ThisTag) & "*" Then
        If Suggest = "" Then Suggest = Alltags(L).FullTag
        FormSugest.List.AddItem Alltags(L).FullTag
    End If
Next

'hide or show/resize the suggestion list:
FormSugest.List.Visible = True
If FormSugest.List.ListCount < 2 Then 'one suggestion only, no need for a list
    FormSugest.Hide
    Me.SetFocus
Else 'multiple suggestions, show the list
    If FormSugest.List.ListCount > 8 Then '1785
        FormSugest.Height = 1800
        FormSugest.List.Height = 1800
    Else
        FormSugest.Height = FormSugest.List.ListCount * 230
        FormSugest.List.Height = FormSugest.List.ListCount * 230
    End If
    FormSugest.Show
    Me.SetFocus
End If
Screen.MousePointer = vbDefault
End Sub

'This one receives tags when the suggestion list is clicked:
Public Property Let SuggestTag(strTag As String)
Suggest = strTag
Call ShowSuggest(0)
End Property

'This one receives the chosen tag from the suggestion list
'on Enter, tab, space, rightarrow or dblclick there:

Public Property Let UseTag(strTag As String)
    Unload FormSugest
    With RTbox
    .SelStart = Tagstart - 1
    .SelLength = Len(Suggest)
    .SelText = strTag
    .SelStart = .SelStart + .SelLength
    End With
    bIntag = False
End Property

'Display the suggestion in the Rich Textbox,
'optional select from current position and to end of it:

Private Sub ShowSuggest(FromPos As Long)
Dim L As Long
With RTbox
If FromPos > 0 Then
    .SelText = Mid(Suggest, Len(ThisTag) + 1)
    .SelStart = FromPos
    .SelLength = Len(Suggest) - Len(ThisTag)
Else
    L = .SelStart + .SelLength
    L = L - Tagstart + 1
    .SelStart = Tagstart - 1
    .SelLength = L
    .SelText = Suggest
End If
End With
End Sub

'Use the suggestion, close the automatics and suggestions
'and place caret after the tag:

Public Sub AcceptSuggest()
Unload FormSugest
With RTbox
.SelStart = Tagstart - 1
.SelLength = Len(Suggest)
.SelText = Suggest
.SelStart = .SelStart + .SelLength
bIntag = False
End With
End Sub

'This puts all tags into the tree in a hierarchy
'and decides which one that should be closed at (or after) the caret position

Function PrevOpenTag(blnToHere As Boolean) As String

Dim T As Long
Dim Stmp As String
Dim L As Long
Dim LS As Integer

Dim ThisTag As String
Dim ShortTag As String

Dim Nodx As Node
Dim CurNode As Node
Tvw1.Nodes.Clear
Tvw1.Visible = False
With RTbox
T = 0
L = -1
Stmp = .Text & " "
If blnToHere = True Then Stmp = Left(.Text, Tagstart - 1)
'loop the string and analyse everything within < and > 's:
Do
    T = InStr(Stmp, "<")
    If T = 0 Then Exit Do
    Stmp = Mid(Stmp, T + 1)
    LS = InStr(Stmp, ">") + 1
    L = L + T
    ThisTag = Mid(.Text, L + 1, LS)
    ShortTag = UCase$(Shortened(ThisTag, False))
    If Gotit(ShortTag) Then
    
        If CurNode Is Nothing Then 'tag is not inside another open tag:
            Set Nodx = Tvw1.Nodes.Add(, , "P" & L & "L" & LS, ShortTag)
            Nodx.Tag = "O"
            Nodx.ForeColor = vbRed
            Set CurNode = Nodx
            Set Nodx = Nothing
        Else
            If Left(ThisTag, 2) = "</" Then 'Closing tag
                If CurNode.Text <> ShortTag Then 'not last open tag that's closed
                    'so mark it visible
                    Set Nodx = CurNode
                    Nodx.EnsureVisible
                    'and go search for the real one:
                    Do
                        If Nodx.Parent Is Nothing Then Exit Do
                        Set Nodx = Nodx.Parent
                    Loop Until Nodx.Text = ShortTag And Nodx.Tag = "O"
                    If Not Nodx Is Nothing Then
                        Set CurNode = Nodx
                        CurNode.Tag = "C"
                        CurNode.ForeColor = vbBlue
                    End If
                    Set Nodx = Nothing
                Else
                    CurNode.Tag = "C"
                    CurNode.ForeColor = vbBlue
                    If CurNode.Parent Is Nothing Then
                        Set CurNode = Nothing
                    Else
                        Set CurNode = CurNode.Parent
                    End If
                End If
            Else 'new opening tag, add under current:
                Set Nodx = Tvw1.Nodes.Add(CurNode.Key, tvwChild, _
                    "P" & L & "L" & LS, ShortTag)
                'Mark it open if it requires closing:
                If ShouldCloseIt(ShortTag) = True Then
                    Nodx.Tag = "O"
                    Nodx.ForeColor = vbRed
                    Set CurNode = Nodx
                Else
                    Nodx.ForeColor = vbBlue 'so mark it closed
                    Nodx.Tag = "C"
                    Set Nodx = Nothing
                End If
            End If
        End If
    End If
Loop Until T = 0
If Not CurNode Is Nothing Then
    CurNode.EnsureVisible
Else
    GoTo none
End If
End With

If CurNode.Tag = "C" Then
    On Error Resume Next
    Set Nodx = CurNode
    Do
        Set Nodx = CurNode.Parent
    Loop Until Nodx.Tag = "C"
    PrevOpenTag = Shortened(Nodx.Text, True)
Else
    PrevOpenTag = Shortened(CurNode.Text, True)
End If
If bUcaseTags = True Then
    PrevOpenTag = UCase$(PrevOpenTag) 'the returned suggestion for caret position
Else
    PrevOpenTag = LCase$(PrevOpenTag) 'the returned suggestion for caret position
End If
Tvw1.Visible = True
Exit Function
none:
If bUcaseTags = True Then
    PrevOpenTag = "<HTML>"
Else
    PrevOpenTag = "<html>"
End If
Tvw1.Visible = True
End Function

'This creates a short version, like
'<P> from a lengthy <P ALIGN=CENTER>
'and optionally closes it like </P>

Function Shortened(strTag As String, _
    blnCloseit As Boolean) As String
Shortened = Replace(strTag, " ", ">")
'Shortened = UCase(Shortened)
Shortened = Left(Shortened, InStr(Shortened, ">"))
If Left(Shortened, 2) = "</" Then
    If blnCloseit = False Then _
        Shortened = Left(Shortened, 1) & Mid(Shortened, 3)
Else
    If blnCloseit = True Then _
        Shortened = Left(Shortened, 1) & "/" & Mid(Shortened, 2)
End If
End Function

'This reads a short version and decides
'whether it needs a closing tag:

Private Function ShouldCloseIt(ShortTag As String) As Boolean
On Error Resume Next
ShouldCloseIt = AllShorts(ShortTag).CloseRequired
End Function

'Is this short tag in our tag collection at all ?
'(we skip silly stuff not in out textfile, like <whatever>)

Private Function Gotit(ShortTag As String) As Boolean
On Error Resume Next
Gotit = Len(AllShorts(ShortTag).ShortTag)
End Function

