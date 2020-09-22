VERSION 5.00
Object = "{52E535C3-131A-49DA-8069-D2856765301C}#10.0#0"; "googler.ocx"
Begin VB.Form Form1 
   Caption         =   "Google Search Control Example"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   90
      TabIndex        =   7
      Top             =   1845
      Width           =   7395
   End
   Begin Googler.Google Google1 
      Height          =   1650
      Left            =   3285
      Top             =   135
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   2910
   End
   Begin VB.TextBox txtMaxResults 
      Height          =   330
      Left            =   1125
      TabIndex        =   2
      Text            =   "50"
      Top             =   495
      Width           =   645
   End
   Begin VB.TextBox txtSearch 
      Height          =   330
      Left            =   1125
      TabIndex        =   1
      Text            =   "Batman"
      Top             =   90
      Width           =   2040
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   330
      Left            =   1935
      TabIndex        =   0
      Top             =   495
      Width           =   1230
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status:"
      Height          =   285
      Left            =   180
      TabIndex        =   6
      Top             =   5805
      Width           =   7260
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   825
      Left            =   135
      TabIndex        =   5
      Top             =   900
      Width           =   2985
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Max Results:"
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   585
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search Term:"
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSearch_Click()
  
  If txtSearch.Text <> "" Then
      List1.Clear
      List1.AddItem "Please Wait! Searching Google For " & txtSearch.Text & "..."
  
      'search google with search item and number of max results desired
      'Note: Maximum results allowed is 100
       Google1.QueryGoogle txtSearch.Text, CInt(txtMaxResults.Text)
  Else
      MsgBox "Please Enter An Item To Search For!", vbCritical
  End If
  
End Sub

Private Sub Google1_QueryComplete(Results As Googler.GoogleResults)
  
  Dim Page As Googler.Result
  
  lblStatus = "Search Complete"
  List1.Clear
  If Results.Count > 0 Then 'found anything?
      
    'yes, display results...
    List1.AddItem "Found " & CStr(Results.Count) & " Pages..."
    List1.AddItem ""
    
    'loop thru collection and display
    For Each Page In Results
      List1.AddItem Page.PageName
      List1.AddItem Page.PageUrl
      List1.AddItem ""
    Next
  
  Else
   
   'nothing found...
    List1.AddItem "No Results Found For " & txtSearch.Text & "..."
  
  End If


End Sub

Private Sub Google1_QueryError(ByVal ErrNumber As Integer, ByVal ErrDescription As String)
   
   lblStatus = "Search Error!"
   MsgBox "[" & CStr(ErrNumber) & "] " & ErrDescription, vbCritical
   
End Sub

Private Sub Google1_QueryProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusNumber As AsyncStatusCodeConstants, ByVal StatusText As String)
   On Error Resume Next
   If ProgressMax >= Progress Then
        lblStatus = "Progress: " & CStr(100 \ Int(ProgressMax / Progress)) & "% of " & CStr(ProgressMax)
   Else
        lblStatus = "Progress: " & CStr(Progress) & "%"
   End If
End Sub

Private Sub List1_Click()
   On Error Resume Next
   'code to click on urls to open them in a browser
   If List1.Text <> "" Then
     If Left$(List1.Text, 7) = "http://" Then
         Shell "explorer.exe " & List1.Text, vbNormalFocus
     End If
   End If
End Sub
