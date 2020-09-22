VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarPageName As String 'local copy
Private mvarPageUrl As String 'local copy
Private mvarPageDetails As String 'local copy

Public Property Let PageDetails(ByVal vData As String)
    mvarPageDetails = vData
End Property

Public Property Get PageDetails() As String
    PageDetails = mvarPageDetails
End Property

Public Property Let PageUrl(ByVal vData As String)
    mvarPageUrl = vData
End Property

Public Property Get PageUrl() As String
    PageUrl = mvarPageUrl
End Property

Public Property Let PageName(ByVal vData As String)
    mvarPageName = vData
End Property

Public Property Get PageName() As String
    PageName = mvarPageName
End Property



