VERSION 5.00
Begin VB.UserControl Google 
   CanGetFocus     =   0   'False
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   ClipControls    =   0   'False
   HasDC           =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4410
   ToolboxBitmap   =   "Google.ctx":0000
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   0
      Picture         =   "Google.ctx":0312
      Top             =   0
      Width           =   4140
   End
End
Attribute VB_Name = "Google"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Event QueryComplete(Results As GoogleResults)
Event QueryError(ByVal ErrNumber As Integer, ByVal ErrDescription As String)
Event QueryProgress(ByVal Progress As Long, ByVal ProgressMax As Long, ByVal StatusNumber As AsyncStatusCodeConstants, ByVal StatusText As String)

Dim strGoogleSearch As String
Dim strGoogleBuffer As String

'Default Property Values:
Const m_def_MaxResults As Integer = 50

Dim m_MaxResults     As Integer

Sub CancelQuery()

    On Error Resume Next
        UserControl.CancelAsyncRead vbNullString

End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

    On Error Resume Next
      Dim Codestring As String

        With AsyncProp
            If .Status = "" Then
                Codestring = GetStatusCode(.StatusCode)
              Else
                Codestring = .Status
            End If

            If .StatusCode = vbAsyncStatusCodeError Then
                RaiseEvent QueryError(.StatusCode, Codestring)
              Else
                RaiseEvent QueryProgress(.BytesRead, .BytesMax, .StatusCode, Codestring)
            End If
        End With

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

    On Error Resume Next
      Dim strHeader As String, strHtml As String
      Dim lngPlace As Long, strCodestring As String

        With AsyncProp

            If .StatusCode = vbAsyncStatusCodeError Then
                If .Status = "" Then
                    strCodestring = GetStatusCode(.StatusCode)
                  Else
                    strCodestring = .Status
                End If
                RaiseEvent QueryError(.StatusCode, strCodestring)
                Exit Sub
              Else
                strGoogleBuffer = StrConv(.Value, vbUnicode)
            End If
        End With
        
        'look for end of page
        lngPlace = InStr(UCase$(strGoogleBuffer), "</HTML>")
        
        'if found, seperate header from page
        If lngPlace > 0 Then
            lngPlace = InStr(strGoogleBuffer, vbCrLf & vbCrLf)
            If lngPlace Then
                strHeader = Mid$(strGoogleBuffer, 1, lngPlace - 1)
                strHtml = Mid$(strGoogleBuffer, lngPlace + 3)
                ParseGoogle strHtml 'parse urls
            Else
                If LenB(strGoogleBuffer) > 0 Then
                    ParseGoogle strGoogleBuffer
                Else
                    GoTo ErrHandler
                End If
            End If
        Else
            GoTo ErrHandler
        End If

Exit Sub
ErrHandler:
        RaiseEvent QueryError(744, "Search Text Not Found")


End Sub

Private Sub UserControl_Resize()

    UserControl.Width = Image1.Width
    UserControl.Height = Image1.Height

End Sub

'this is the call you make from your application
Sub QueryGoogle(ByVal SearchItem As String, Optional ByVal MaximumResults As Integer = -1)

    On Error GoTo Errhandle

        If MaximumResults <> -1 Then
            MaxResults = MaximumResults
        End If
        
        'save search string for later use
        strGoogleSearch = Replace$(SearchItem, " ", "+")
        strGoogleBuffer = ""
        
        'start connection to google
        UserControl.AsyncRead "http://www.google.com/search?num=" & CStr(MaxResults) & "&hl=en&lr=&ie=UTF-8&oe=UTF-8&q=" & Trim$(strGoogleSearch), vbAsyncTypeByteArray, vbNullString, vbAsyncReadForceUpdate

Exit Sub
Errhandle:
       If Err.Source = UserControl.Name Then
           RaiseEvent QueryError(Err.Number + vbObjectError, Err.Description)
       Else
           Resume Next
       End If
           
End Sub


'this is a parser, that looks thru the html and grabs
'all the urls from it
Private Sub ParseGoogle(ByVal Packet As String)

    On Error Resume Next

      Dim lngPlace As Long, lngLength As Long, X As Long
      Dim strResults As String, strPage As String, strUrl As String
      Dim Reslist As GoogleResults

        'No pages were found containing "jkgfkgfkgfkhgf".
        'No standard web pages containing all your search terms were found

        Set Reslist = New GoogleResults

        lngPlace = InStr(1, Packet, "NO PAGES WERE FOUND CONTAINING " & Chr$(34) & UCase$(strGoogleSearch) & Chr$(34), vbTextCompare)
        If lngPlace = 0 Then
            lngPlace = InStr(1, Packet, "NO STANDARD WEB PAGES CONTAINING", vbTextCompare)
        End If

        If lngPlace = 0 Then
            
            'Results <b>1</b> - <b>10</b> of about <b>1,320</b>.   Search took <b>0.81</b> seconds.</font>
            'lngPlace = InStr(1, Packet, "Results <B>", vbTextCompare)
            'lngLength = InStr(lngPlace, Packet, "</FONT", vbTextCompare) - lngPlace
            'strResults = Mid$(Packet, lngPlace, lngLength)
            'strResults = Replace$(strResults, "<B>", "", , , vbTextCompare)
            'strGoogleResult = Replace$(strResults, "</B>", "", , , vbTextCompare)
            
            lngPlace = InStr(lngPlace + 1, Packet, "<p class=g>", vbTextCompare)
            'lngPlace = InStr(lngPlace + 1, Packet, "<p class=g>", vbTextCompare)
            If lngPlace Then
                Do While (lngPlace > 0)
                    lngPlace = InStr(lngPlace, Packet, "<a href=", vbTextCompare)
                    If lngPlace Then
                        lngPlace = lngPlace + 8
                        lngLength = InStr(lngPlace + 1, Packet, ">", vbTextCompare)
                        lngLength = lngLength - lngPlace
                        strUrl = RemoveFlack(Mid$(Packet, lngPlace, lngLength)) 'grab url
                        lngPlace = lngPlace + lngLength + 1
                        lngLength = InStr(lngPlace, Packet, "</a>", vbTextCompare)
                        lngLength = lngLength - lngPlace
                        strPage = Mid$(Packet, lngPlace, lngLength) 'grab page
                        strPage = Replace$(strPage, "<b>", "", , , vbTextCompare)
                        strPage = Replace$(strPage, "</b>", "", , , vbTextCompare)
                        Reslist.Add strPage, strUrl, "" 'add both to list
                        lngPlace = lngPlace + lngLength
                    End If
                    X = X + 1
                    lngPlace = InStr(lngPlace + 1, Packet, "<p class=g>", vbTextCompare)
                Loop
            End If
            RaiseEvent QueryComplete(Reslist) 'raise event
          Else
            GoTo NoResults
        End If

    Exit Sub

NoResults:
        RaiseEvent QueryError(744, "Search Text Not Found")

End Sub


'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_MaxResults = m_def_MaxResults

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_MaxResults = PropBag.ReadProperty("MaxResults", m_def_MaxResults)

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("MaxResults", m_MaxResults, m_def_MaxResults)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,10
Public Property Get MaxResults() As Integer

    MaxResults = m_MaxResults

End Property

Public Property Let MaxResults(ByVal New_MaxResults As Integer)

    If New_MaxResults < 10 Or New_MaxResults > 100 Then
        New_MaxResults = m_def_MaxResults
    End If

    m_MaxResults = New_MaxResults
    PropertyChanged "MaxResults"

End Property

Public Property Get GoogleHtml() As String

    GoogleHtml = strGoogleBuffer

End Property

Function GetStatusCode(ByVal Code As AsyncStatusCodeConstants) As String

    If Code = vbAsyncStatusCodeBeginDownloadData Then
        GetStatusCode = "Download Initialized."
      ElseIf Code = vbAsyncStatusCodeBeginSyncOperation Then
        GetStatusCode = "Synchronous Download Has Started."
      ElseIf Code = vbAsyncStatusCodeCacheFileNameAvailable Then
        GetStatusCode = "Local Cache File Is Available."
      ElseIf Code = vbAsyncStatusCodeConnecting Then
        GetStatusCode = "Connecting To Resource."
      ElseIf Code = vbAsyncStatusCodeDownloadingData Then
        GetStatusCode = "Download In Progress."
      ElseIf Code = vbAsyncStatusCodeEndDownloadData Then
        GetStatusCode = "Download Complete."
      ElseIf Code = vbAsyncStatusCodeEndSyncOperation Then
        GetStatusCode = "Synchronous Download Complete."
      ElseIf Code = vbAsyncStatusCodeError Then
        GetStatusCode = "An Error Has Occurred."
      ElseIf Code = vbAsyncStatusCodeFindingResource Then
        GetStatusCode = "Finding Resource."
      ElseIf Code = vbAsyncStatusCodeMIMETypeAvailable Then
        GetStatusCode = "MIME Type Is Available."
      ElseIf Code = vbAsyncStatusCodeRedirecting Then
        GetStatusCode = "Redirecting."
      ElseIf Code = vbAsyncStatusCodeSendingRequest Then
        GetStatusCode = "Sending Request."
      ElseIf Code = vbAsyncStatusCodeUsingCachedCopy Then
        GetStatusCode = "Using Cached Copy."
      Else
        GetStatusCode = "Unknown."
    End If

End Function

Function RemoveFlack(ByVal strUrl As String)

  Dim X As Long
  For X = 255 To 1 Step -1
    Do While InStr(1, strUrl, "%" & Right$("0" & Hex$(X), 2)) > 0
      strUrl = Replace$(strUrl, "%" & Right$("0" & Hex$(X), 2), Chr$("&H" & Right$("0" & Hex$(X), 2)))
    Loop
  Next
  If InStr(strUrl, "/ ") Then
     strUrl = Left$(strUrl, InStr(strUrl, "/ ") - 1)
  End If
  If InStr(strUrl, " ") Then
     strUrl = Left$(strUrl, InStr(strUrl, " ") - 1)
  End If
  If InStr(strUrl, " target=nw") Then
     strUrl = Replace$(strUrl, " target=nw", "", , , vbTextCompare)
  End If
  RemoveFlack = Replace$(Replace$(strUrl, "+", " "), "  ", " ")
  
End Function
