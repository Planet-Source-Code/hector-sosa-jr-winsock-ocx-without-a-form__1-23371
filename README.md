<div align="center">

## Winsock ocx without a form


</div>

### Description

This class wraps the winsock.ocx methods and properties. This allows you to use winsock functions without putting the ocx on a form. This class can be compiled into a dll if wanted.

It was said on the usenet that it could not be done, but I did it. I hope this saves people from headaches and upset stomachs ;-)

Enjoy!
 
### More Info
 
Make sure that the toolbar does not have the winsock control on it, otherwise this code will not work.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Hector Sosa, Jr](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/hector-sosa-jr.md)
**Level**          |Intermediate
**User Rating**    |3.7 (11 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/hector-sosa-jr-winsock-ocx-without-a-form__1-23371/archive/master.zip)





### Source Code

```
'------------------------------------------------------------------------
'
' Class Module clsWinsock
' File: clsWinsock.cls
' Author: Hector
' Date: 5/10/01
' Purpose: This class allows to use the winsock functions without having
'     to put a winsock control on a form. Make sure to have a
'     reference to the winsock.ocx in the project where you're going
'     to use this class or this won't work.
'
'------------------------------------------------------------------------
Option Explicit
Private WithEvents objSocket As Winsock
Public Event DataInStream(ByVal lngSocketNumber As Long, ByVal strData As String)
Public Event SocketClosed(ByVal lngSocketNumber As Long)
Public Event ConnectionRequested(ByVal lngSocketNumber As Long)
Public Event AcceptedSocket(ByVal lngSocketNumber As Long)
Private mvarPortNumber As Long
Private mvarCurrDataStream As String
Private mvarCurrentID As Long
Private blnSoftReturn As Boolean
'*****************************************************************************************
'* Property  : CurrentSocketID
'* Notes    : Returns the current socket number.
'*****************************************************************************************
Public Property Get CurrentSocketID() As Long
  CurrentSocketID = mvarCurrentID
End Property
'*****************************************************************************************
'* Property  : CurrDataStream
'* Notes    : Returns the raw input from the current socket.
'*****************************************************************************************
Private Property Let CurrDataStream(ByVal vData As String)
  mvarCurrDataStream = vData
End Property
Public Property Get CurrDataStream() As String
  CurrDataStream = mvarCurrDataStream
End Property
'*****************************************************************************************
'* Property  : LocalPort
'* Notes    : Returns/Sets the port where the socket will be listening on.
'*****************************************************************************************
Public Property Let LocalPort(ByVal vData As Long)
  mvarPortNumber = vData
  objSocket.LocalPort = vData
End Property
Public Property Get LocalPort() As Long
  LocalPort = mvarPortNumber
End Property
Private Sub Class_Initialize()
Set objSocket = New Winsock
End Sub
Private Sub Class_Terminate()
  If objSocket.State <> sckClosed Then objSocket.Close
  Set objSocket = Nothing
End Sub
'-----------------------------------------------------------------------
'
' Procedure objSocket_ConnectionRequest
' Author: Hector
' Date: 5/16/01
' Purpose: Handles connection requests.
' Result:
' Input parameters: requestID
'
' Output parameters:
'
'------------------------------------------------------------------------
Private Sub objSocket_ConnectionRequest(ByVal requestID As Long)
  If objSocket.State <> sckClosed Then objSocket.Close
  mvarCurrentID = requestID
  RaiseEvent ConnectionRequested(requestID)
End Sub
'-----------------------------------------------------------------------
'
' Procedure objSocket_DataArrival
' Author: Hector
' Date: 5/16/01
' Purpose: Handles data arriving to the socket.
' Result:
' Input parameters: bytesTotal
'
' Output parameters:
'
' Last Modification:
' 5/22/01 - Finished the handling of broken packets (consecutive streams).
'------------------------------------------------------------------------
Private Sub objSocket_DataArrival(ByVal bytesTotal As Long)
  Dim strIncoming As String
  Static strInputBuffer As String
  Dim strOutBuffer As String
  Dim intPos As Integer
  objSocket.GetData strIncoming
  CurrDataStream = strIncoming
  mvarCurrentID = objSocket.SocketHandle
  ' Replace Carriage Returns/Line Feeds or just Line Feeds with
  ' a Carriage Return for consistant handling.
  strIncoming = Replace$(strIncoming, vbCrLf, vbCr)
  strIncoming = Replace$(strIncoming, vbLf, vbCr)
  ' Check for Carriage Returns in the incoming stream, and mark
  ' the position where it's found, if any.
  intPos = InStr(1, strIncoming, vbCr)
  ' Make sure that the Carriage Return is not at the beginning of the stream.
  ' If the Carriage Return is at position 1 then it means that it belongs to the
  ' previous stream.
  If intPos > 1 Then
    ' Grab a string including the Carriage Return for processing.
    strOutBuffer = Left$(strIncoming, intPos)
    strOutBuffer = StripCRLF(strIncoming)
    RaiseEvent DataInStream(mvarCurrentID, strOutBuffer)
    ' Flush the buffers so that data won't get added to the next stream.
    strOutBuffer = ""
    strInputBuffer = ""
  Else
    ' Add to the input buffer if there is no Carriage Return.
    strInputBuffer = strInputBuffer & strIncoming
  End If
  ' The code below handles broken packets, meaning that all the data did not
  ' come in one stream.
  '******************************************************************************
  If Right$(strIncoming, 1) = vbCr Then  'check last character
    blnSoftReturn = True
  End If
  If blnSoftReturn = True Then
    If Left$(strIncoming, 1) = vbCr Then
      strOutBuffer = Mid$(strInputBuffer, 1)
      strOutBuffer = StripCRLF(strOutBuffer)
      RaiseEvent DataInStream(mvarCurrentID, strOutBuffer)
      ' Flush the buffers so that data won't get added to the next stream.
      strOutBuffer = ""
      strInputBuffer = ""
    End If
    blnSoftReturn = False
  End If
  '*******************************************************************************
End Sub
Private Sub objSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  ' Lame error handling. If you want something better go ahead and put it here. When there is
  ' and error, it usually means that the socked was disconnected.
  If objSocket.State <> sckClosed Then objSocket.Close
  MsgBox "Something happened to socket #" & CStr(Number)
End Sub
'-----------------------------------------------------------------------
'
' Procedure StripCRLF
' Author: Hector
' Date: 5/16/01
' Purpose: Removes carriage returns and line feeds from incoming data.
' Result:
' Input parameters: strData
'
' Output parameters:
'
'------------------------------------------------------------------------
Private Function StripCRLF(strData As String)
  Dim strHold As String
  strHold = Replace(strData, vbCr, "")
  strHold = Replace(strHold, vbLf, "")
  StripCRLF = strHold
End Function
'-----------------------------------------------------------------------
'
' Procedure SocketListen
' Author: Hector
' Date: 5/16/01
' Purpose: Allows the socket to listen to incoming transmitions.
' Result:
' Input parameters: None
'
' Output parameters:
'
'------------------------------------------------------------------------
Public Sub SocketListen()
  objSocket.Listen
End Sub
'-----------------------------------------------------------------------
'
' Procedure SocketClose
' Author: Hector
' Date: 5/16/01
' Purpose: Stops the socket from listening to any more requests or data
'     arrivals.
' Result:
' Input parameters: None
'
' Output parameters:
'
'------------------------------------------------------------------------
Public Sub SocketClose()
  objSocket.Close
End Sub
'-----------------------------------------------------------------------
'
' Procedure AcceptRequest
' Author: Hector
' Date: 5/16/01
' Purpose: Accepts a request to connect.
' Result:
' Input parameters: lngSocketNumber
'
' Output parameters:
'
'------------------------------------------------------------------------
Public Sub AcceptRequest(ByVal lngSocketNumber As Long)
  objSocket.Accept lngSocketNumber
  RaiseEvent AcceptedSocket(lngSocketNumber)
End Sub
'-----------------------------------------------------------------------
'
' Procedure SendData
' Author: Hector
' Date: 5/17/01
' Purpose: Sends data to the user connected to this socket.
' Result:
' Input parameters: sDataToSend
'
' Output parameters:
'
'------------------------------------------------------------------------
Public Sub SendData(ByVal sDataToSend As String)
  objSocket.SendData sDataToSend
End Sub
```

