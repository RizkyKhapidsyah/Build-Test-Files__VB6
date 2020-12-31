Attribute VB_Name = "basMain"
Option Explicit

Public g_blnStopProcessing  As Boolean
Public g_strVersion         As String
Public cRnd                 As clsRndData
Public cFSO                 As clsFSO

Private Const BASE_LENGTH   As Long = 32768

Sub Main()

' -----------------------------------------------------------------------------
' Set up the path where all of the mail processing will take place.
' ---------------------------------------------------------------------------
  ChDrive App.Path
  ChDir App.Path
      
' ---------------------------------------------------------------------------
' See if there is another instance of this program running
' ---------------------------------------------------------------------------
  If App.PrevInstance Then
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Intialize global variables
' ---------------------------------------------------------------------------
  g_strVersion = "Build Test Files v" & CStr(App.Major) & "." & CStr(App.Minor)
  Set cRnd = New clsRndData
  Set cFSO = New clsFSO
  
' ---------------------------------------------------------------------------
' set processing switches
' ---------------------------------------------------------------------------
  g_blnStopProcessing = False
  cFSO.CancelProcessing = False
  cRnd.CancelProcessing = False
  
' ---------------------------------------------------------------------------
' Load the main form
' ---------------------------------------------------------------------------
  Load frmBldRnd
  
End Sub

Public Sub StopApplication()

' ---------------------------------------------------------------------------
' Unload all forms then terminate application
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 10-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------
  Set cFSO = Nothing    ' Free class objects from memory
  Set cRnd = Nothing
  
  Unload_All_Forms      ' free form objects from memory
  End
  
End Sub

Public Sub Unload_All_Forms()

' ---------------------------------------------------------------------------
' Unload all forms before terminating an application.  The calling module
' will call this routine and usually executes END when it returns.
' ===========================================================================
'    DATE      NAME                      DESCRIPTION
' -----------  ------------------------  ------------------------------------
' 10-DEC-2000  Kenneth Ives              Written by kenaso@home.com
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim frm As Form
  
' ---------------------------------------------------------
' As we find a form, we will unload it and free memory.
' ---------------------------------------------------------
  For Each frm In Forms
      If TypeOf frm Is Form Then
          frm.Hide                ' Hide the form
          Unload frm              ' Deactivate the form object
          Set frm = Nothing       ' Free form object from memory
      End If
  Next
  
End Sub
Public Function Build_Fixed(strFilename As String, _
                                 ByVal lngNumberOfBytes As Long, _
                                 ByVal blnHexConvert As Boolean, _
                                 ByVal blnUseKeyboardChars As Boolean) As Boolean

' ***************************************************************************
' Routine:       Build_Fixed
'
' Description:   This routine will build a file with fixed length records.
'                These will be 80 bytes in length.  Seventy-eight printable
'                characters and two hidden (carriage return and linefeed).
'
' Return Values: The newly created record.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' ***************************************************************************

  On Error GoTo Cancel_Processing
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex     As Long
  Dim lngByteCnt   As Long
  Dim lngRecLen    As Long
  Dim lngAmtLeft   As Long
  Dim intChar      As Integer
  Dim hFile        As Integer
  Dim strOutRec    As String
  Dim strTmp       As String
  Dim strTmpRnd    As String
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  Close     ' close all open files
  lngRecLen = 78
  lngByteCnt = 0
  strOutRec = ""
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  If lngNumberOfBytes < lngRecLen Then
      lngRecLen = lngNumberOfBytes - 2
  End If

' ---------------------------------------------------------------------------
' Build a 32kb random data string once.  Since this is only for building a
' test file, we do not have to worry about security.
' ---------------------------------------------------------------------------
  strTmpRnd = cRnd.Build_Random_Data(BASE_LENGTH, blnHexConvert, blnUseKeyboardChars)
      
' ---------------------------------------------------------------------------
' Open the new test file as output
' ---------------------------------------------------------------------------
  hFile = FreeFile
  Open strFilename For Output As #hFile

' ---------------------------------------------------------------------------
' Start building the fixed length file here.
' ---------------------------------------------------------------------------
  Do
      ' if the Stop Button is pressed then exit
      DoEvents
      If g_blnStopProcessing Then
          Exit Do
      End If
      
      ' reload temp data string
      strTmp = strTmpRnd
      
      Do
          ' if the Stop Button is pressed then exit
          DoEvents
          If g_blnStopProcessing Then
              Exit Do
          End If

          strOutRec = Left$(strTmp, lngRecLen)  ' capture data
          strTmp = Mid$(strTmp, lngRecLen + 1)  ' resize the random data
            
          ' no trailing semi-colon forces a hard return at the
          ' end of each record thus adding 2 characters (0x0D & 0x0A)
          ' while a semi-colon creates a contiguous record.
          Print #hFile, strOutRec    ' <-- No semi-colon
                                       
          ' calculate how much is left
          lngByteCnt = lngByteCnt + Len(strOutRec) + 2
          lngAmtLeft = lngNumberOfBytes - lngByteCnt
          
          If lngAmtLeft = 2 Then
              Print #1, ""             ' write a blank line
              lngByteCnt = lngByteCnt + 2  ' increment counter
              lngRecLen = 0              ' reset record length to 0
              Exit Do                 ' exit FOR loop
          ElseIf lngAmtLeft < 2 Then
              lngByteCnt = lngNumberOfBytes   ' WE are finished
              lngRecLen = 0              ' reset record length to 0
              Exit Do
          ElseIf lngAmtLeft >= (lngRecLen + 2) Then
              lngRecLen = 78             ' reset record length to max
          Else
              lngRecLen = (lngAmtLeft - 2)  ' reset to what is left minus 2
          End If
          
      Loop Until Len(strTmp) = 0
      
  Loop Until lngByteCnt >= lngNumberOfBytes
  
' ---------------------------------------------------------------------------
' see if the Stop Button is pressed then exit
' ---------------------------------------------------------------------------
  DoEvents
  If g_blnStopProcessing Then
      GoTo Cancel_Processing
  End If
    
' ---------------------------------------------------------------------------
' Good finish
' ---------------------------------------------------------------------------
  Close #hFile
  Build_Fixed = True
  Exit Function
  
' ---------------------------------------------------------------------------
' User pressed the STOP button
' ---------------------------------------------------------------------------
Cancel_Processing:
  Close #hFile
  
  If cFSO.File_Exist(strFilename) Then
      Kill strFilename
  End If
              
  Build_Fixed = False

End Function

Public Function Build_Continuous(strFilename As String, _
                                 ByVal lngNumberOfBytes As Long, _
                                 ByVal blnHexConvert As Boolean, _
                                 ByVal blnUseKeyboardChars As Boolean) As Boolean

' ***************************************************************************
' Routine:       Build_Continuous
'
' Description:   This routine will build a file with one contiguous record.
'                Sometimes referred to a variable length record.
'
' Return Values: The newly created record.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' 30-SEP-2000  Kenneth Ives     Corrected the calculation of the last record
'                               so as to accurately display the output
' ***************************************************************************

  On Error GoTo Cancel_Processing
' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim lngIndex     As Long
  Dim lngStart     As Long
  Dim lngByteCnt   As Long
  Dim lngRecLen    As Long
  Dim lngAmtLeft   As Long
  Dim intChar      As Integer
  Dim hFile        As Integer
  Dim strOutRec    As String
  Dim strTmp       As String
  Dim strTmpRnd    As String

  Const REC_LENGTH As Long = 4096
  
' ---------------------------------------------------------------------------
' Initialize local variables
' ---------------------------------------------------------------------------
  Close     ' close all open files
  lngByteCnt = 0
  strOutRec = ""
  
' ---------------------------------------------------------------------------
' Build a 32kb random data string once.  Since this is only for building a
' test file, we do not have to worry about security.
' ---------------------------------------------------------------------------
  strTmpRnd = cRnd.Build_Random_Data(BASE_LENGTH, blnHexConvert, blnUseKeyboardChars)
      
' ---------------------------------------------------------------------------
' Open the new test file as output
' ---------------------------------------------------------------------------
  hFile = FreeFile
  Open strFilename For Output As #hFile

' ---------------------------------------------------------------------------
' Start building the fixed length file here.
' ---------------------------------------------------------------------------
  Do
      ' if the Stop Button is pressed then exit
      DoEvents
      If g_blnStopProcessing Then
          Exit Do
      End If
      
      ' initialize variables
      lngStart = 1
      lngRecLen = REC_LENGTH
      
      ' Test record length
      If lngRecLen > lngNumberOfBytes Then
          lngRecLen = lngNumberOfBytes
      ElseIf lngRecLen >= lngAmtLeft Then
          lngRecLen = lngAmtLeft
      End If
  
      strTmp = strTmpRnd  ' Copy random data to temp variable
      
      Do
          ' if the Stop Button is pressed then exit
          DoEvents
          If g_blnStopProcessing Then
              Exit Do
          End If

          strOutRec = Left$(strTmp, lngRecLen) ' capture data
          strTmp = Mid$(strTmp, lngRecLen + 1)
            
          ' no trailing semi-colon forces a hard return at the
          ' end of each record thus adding 2 characters (0x0D & 0x0A)
          ' while a semi-colon creates a contiguous record.
          Print #hFile, strOutRec;   ' <-- semi-colon
          
          ' calculate how much is left
          lngByteCnt = lngByteCnt + Len(strOutRec)
          lngAmtLeft = lngNumberOfBytes - lngByteCnt
          
          If lngAmtLeft <= 0 Then
              Exit Do
          End If
          
          If lngAmtLeft >= BASE_LENGTH Then
              lngRecLen = REC_LENGTH        ' reset record length to max
          Else
              lngRecLen = lngAmtLeft        ' Save just what is left to write
          End If
      Loop Until Len(strTmp) = 0
      
      If lngByteCnt >= lngNumberOfBytes Then
          Exit Do
      End If
      
  Loop
  
' ---------------------------------------------------------------------------
' see if the Stop Button is pressed then exit
' ---------------------------------------------------------------------------
  DoEvents
  If g_blnStopProcessing Then
      GoTo Cancel_Processing
  End If
    
' ---------------------------------------------------------------------------
' Good finish
' ---------------------------------------------------------------------------
  Close #hFile
  Build_Continuous = True
  Exit Function
  
' ---------------------------------------------------------------------------
' User pressed the STOP button
' ---------------------------------------------------------------------------
Cancel_Processing:
  Close #hFile
  
  If cFSO.File_Exist(strFilename) Then
      Kill strFilename
  End If
              
  Build_Continuous = False
  
End Function

