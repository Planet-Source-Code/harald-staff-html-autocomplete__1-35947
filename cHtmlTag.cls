VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHtmlTag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarFullTag As String
Private mvarCloseRequired As Boolean
Private mvarShortTag As String

Public Property Let ShortTag(ByVal vData As String)
mvarShortTag = UCase$(vData)
End Property

Public Property Get ShortTag() As String
ShortTag = mvarShortTag
End Property

Public Property Let CloseRequired(ByVal vData As Boolean)
mvarCloseRequired = vData
End Property

Public Property Get CloseRequired() As Boolean
CloseRequired = mvarCloseRequired
End Property

Public Property Let FullTag(ByVal vData As String)
mvarFullTag = vData
End Property

Public Property Get FullTag() As String
FullTag = mvarFullTag
End Property



