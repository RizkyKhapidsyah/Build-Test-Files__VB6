VERSION 5.00
Begin VB.Form frmBldRnd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   3885
   ClientTop       =   2685
   ClientWidth     =   5520
   Icon            =   "BldRnd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5520
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   1
      Left            =   4875
      Picture         =   "BldRnd.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   28
      Top             =   75
      Width           =   510
   End
   Begin VB.ComboBox cboDrive 
      Height          =   315
      Left            =   90
      TabIndex        =   26
      Text            =   "cboDrive"
      Top             =   1605
      Width           =   975
   End
   Begin VB.Frame fraHexConv 
      Caption         =   "Convert to hex"
      Height          =   825
      Left            =   1170
      TabIndex        =   23
      Top             =   1965
      Width           =   2025
      Begin VB.OptionButton optHex 
         Caption         =   "Yes (Two Char)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optHex 
         Caption         =   "No (Single Char)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   480
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.Frame fraChars 
      Caption         =   "Type of characters to use"
      Height          =   825
      Left            =   3390
      TabIndex        =   20
      Top             =   1005
      Width           =   2025
      Begin VB.OptionButton optChars 
         Caption         =   "All  (0-255)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optChars 
         Caption         =   "Keyboard (33-126)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fraPredefined 
      Caption         =   "Predefined record sizes"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   2805
      Width           =   5295
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   960
         TabIndex        =   17
         Text            =   "cboSize"
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Record definition"
      Height          =   825
      Left            =   3390
      TabIndex        =   11
      Top             =   1965
      Width           =   2025
      Begin VB.OptionButton optRecLength 
         Caption         =   "Customize length"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optRecLength 
         Caption         =   "Predefined length"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.Frame fraCustom 
      Caption         =   "Define record length"
      Height          =   1095
      Left            =   120
      TabIndex        =   10
      Top             =   2805
      Visible         =   0   'False
      Width           =   5295
      Begin VB.OptionButton optSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   720
         Width           =   3465
      End
      Begin VB.OptionButton optSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   14
         Top             =   480
         Width           =   3465
      End
      Begin VB.OptionButton optSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   3465
      End
      Begin VB.TextBox txtLength 
         Height          =   285
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type of record"
      Height          =   825
      Left            =   1170
      TabIndex        =   7
      Top             =   1005
      Width           =   2025
      Begin VB.OptionButton optType 
         Caption         =   "Continuous"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optType 
         Caption         =   "Fixed Length"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   510
      Index           =   0
      Left            =   90
      Picture         =   "BldRnd.frx":0614
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   75
      Width           =   510
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4020
      TabIndex        =   3
      Top             =   4950
      Width           =   1365
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   2
      Top             =   4950
      Width           =   1290
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Drive"
      Height          =   375
      Left            =   90
      TabIndex        =   27
      Top             =   1125
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1650
      TabIndex        =   6
      Top             =   4950
      Width           =   2175
   End
   Begin VB.Label lblByteCnt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   645
      Width           =   5295
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Height          =   930
      Left            =   225
      TabIndex        =   1
      Top             =   3975
      Width           =   5070
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BldRnd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1410
      TabIndex        =   0
      Top             =   0
      Width           =   2625
   End
End
Attribute VB_Name = "frmBldRnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ***************************************************************************
' Project:       BldRnd
'
' Module:        frmBldRnd
'
' Description:   This is the main screen to determine how the user wants to
'                create their test file.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-DEC-1999  Kenneth Ives     Application created by kenaso@home.com
'
' 02-APR-2001  Kenneth Ives     Remove two classes.  Added class for access
'                               to Scripting.FileSystemObject (scrrun.dll) and
'                               updated random data generation class.  Added
'                               shutdown switches that stop processing almost
'                               immediately.  Created single data string
'                               of 32768 bytes of random data in which to build
'                               test files.  1 mb test file is now created in under
'                               15 seconds on a 500 mhz machine.
' ***************************************************************************

' ---------------------------------------------------------------------------
' Define module level variables
' ---------------------------------------------------------------------------
  Private m_strFilename          As String
  Private m_strRecSize           As String
  Private m_strDriveLtr          As String
  Private m_intSizeIndex         As Integer
  Private m_lngNumberOfBytes     As Long
  Private m_blnHexConvert        As Boolean
  Private m_blnUseKeyboardChars  As Boolean
  Private m_blnFixed             As Boolean
  Private m_blnBytes             As Boolean
  Private m_blnKilobytes         As Boolean
  Private m_blnMegabytes         As Boolean
  Private m_blnLoading           As Boolean
  Private m_blnPredefined        As Boolean
  Private g_blnStopWasPressed    As Boolean

  Private Const m_strByteCaption As String = "Do you have space available on destination drive?"
  
Private Sub Load_Combo_Boxes()

' ***************************************************************************
' Routine:       Load_Combo_Boxes
'
' Description:   This routine will preload the combo box with the most
'                common selections.
'
' ===========================================================================
'    DATE      NAME             DESCRIPTION
' -----------  ---------------  ---------------------------------------------
' 23-DEC-1999  Kenneth Ives     Routine created by kenaso@home.com
' 30 SEP 2000  Kenneth Ives     Added the available drive letters
' ***************************************************************************

' ---------------------------------------------------------------------------
' define local variables
' ---------------------------------------------------------------------------
  Dim intIndex    As Integer
  Dim intPointer  As Integer
  Dim arDrive()   As String
  
' ---------------------------------------------------------------------------
' Empty the combo boxes
' ---------------------------------------------------------------------------
  cboSize.Clear
  cboDrive.Clear
  intPointer = 0
  
' ---------------------------------------------------------------------------
' Load the Size combo box
' ---------------------------------------------------------------------------
  With cboSize
       .AddItem "1 kb (1024 bytes)", 0
       .AddItem "2 kb (2048 bytes)", 1
       .AddItem "4 kb (4096 bytes)", 2
       .AddItem "6 kb (6144 bytes)", 3
       .AddItem "8 kb (8192 bytes)", 4
       .AddItem "10 kb (10,240 bytes)", 5
       .AddItem "16 kb (16,384 bytes)", 6
       .AddItem "32 kb (32,768 bytes)", 7
       .AddItem "64 kb (65,536 bytes)", 8
       .AddItem "128 kb (131,072 bytes)", 9
       .AddItem "256 kb (262,144 bytes)", 10
       .AddItem "512 kb (524,288 bytes)", 11
       .AddItem "1 mb (1,024,000 bytes)", 12
       .AddItem "2 mb (2,048,000 bytes)", 13
       .AddItem "3 mb (3,072,000 bytes)", 14
       .AddItem "4 mb (4,096,000 bytes)", 15
       .AddItem "5 mb (5,120,000 bytes)", 16
       .AddItem "6 mb (6,144,000 bytes)", 17
       .AddItem "7 mb (7,168,000 bytes)", 18
       .AddItem "8 mb (8,192,000 bytes)", 19
       .AddItem "9 mb (9,216,000 bytes)", 20
       .AddItem "10 mb (10,240,000 bytes)", 21
       .AddItem "1.44 mb (1,457,664 bytes)", 22
       .AddItem "720 kb (730,112 bytes)", 23
  End With
  
' ---------------------------------------------------------------------------
' Load the available drives combo box
' ---------------------------------------------------------------------------
  arDrive = cFSO.Available_Drives  ' get available drive letters
  
' ---------------------------------------------------------------------------
' unload the array into the combo box.
' ---------------------------------------------------------------------------
  For intIndex = 0 To UBound(arDrive)
      cboDrive.AddItem StrConv(arDrive(intIndex), vbUpperCase), intIndex
      ' look for drive C:
      If InStr(1, arDrive(intIndex), "C") > 0 Then
          intPointer = intIndex
      End If
  Next
  
' ---------------------------------------------------------------------------
' Set the combo box displays to the first item and erase the array
' ---------------------------------------------------------------------------
  cboSize.ListIndex = 0
  cboDrive.ListIndex = intPointer  ' drive C: index in combo box
  Erase arDrive
  
End Sub

Private Sub cboDrive_Click()

' ---------------------------------------------------------------------------
' If this is the initial form load then exit
' ---------------------------------------------------------------------------
  If m_blnLoading Then
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Capture the visible drive letter on the list
' ---------------------------------------------------------------------------
  m_strDriveLtr = cboDrive.Text
  
End Sub

Private Sub cboSize_Click()

' ---------------------------------------------------------------------------
' If this is the initial form load then exit
' ---------------------------------------------------------------------------
  If m_blnLoading Then
      m_intSizeIndex = 0
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Capture the index and title of item selected
' ---------------------------------------------------------------------------
  m_intSizeIndex = cboSize.ListIndex
  m_strRecSize = cboSize.Text
  m_blnPredefined = True
  lblByteCnt.Caption = m_strByteCaption
  
' ---------------------------------------------------------------------------
' Based on the index selection, get the file length and
' filename.
' ---------------------------------------------------------------------------
  Select Case m_intSizeIndex
         Case 0:    m_strFilename = "T_1KB":   m_lngNumberOfBytes = 1024
         Case 1:    m_strFilename = "T_2KB":   m_lngNumberOfBytes = 2048
         Case 2:    m_strFilename = "T_4KB":   m_lngNumberOfBytes = 4096
         Case 3:    m_strFilename = "T_6KB":   m_lngNumberOfBytes = 6144
         Case 4:    m_strFilename = "T_8KB":   m_lngNumberOfBytes = 8192
         Case 5:    m_strFilename = "T_10KB":  m_lngNumberOfBytes = 10240
         Case 6:    m_strFilename = "T_16KB":  m_lngNumberOfBytes = 16384
         Case 7:    m_strFilename = "T_32KB":  m_lngNumberOfBytes = 32768
         Case 8:    m_strFilename = "T_64KB":  m_lngNumberOfBytes = 65536
         Case 9:    m_strFilename = "T_128KB": m_lngNumberOfBytes = 131072
         Case 10:   m_strFilename = "T_256KB": m_lngNumberOfBytes = 262144
         Case 11:   m_strFilename = "T_512KB": m_lngNumberOfBytes = 524288
         Case 12:   m_strFilename = "T_1MB":   m_lngNumberOfBytes = 1024000
         Case 13:   m_strFilename = "T_2MB":   m_lngNumberOfBytes = 2048000
         Case 14:   m_strFilename = "T_3MB":   m_lngNumberOfBytes = 3072000
         Case 15:   m_strFilename = "T_4MB":   m_lngNumberOfBytes = 4096000
         Case 16:   m_strFilename = "T_5MB":   m_lngNumberOfBytes = 5120000
         Case 17:   m_strFilename = "T_6MB":   m_lngNumberOfBytes = 6144000
         Case 18:   m_strFilename = "T_7MB":   m_lngNumberOfBytes = 7168000
         Case 19:   m_strFilename = "T_8MB":   m_lngNumberOfBytes = 8192000
         Case 20:   m_strFilename = "T_9MB":   m_lngNumberOfBytes = 9216000
         Case 21:   m_strFilename = "T_10MB":  m_lngNumberOfBytes = 10240000
         Case 22:   m_strFilename = "T_144MB": m_lngNumberOfBytes = 1457664
         Case 23:   m_strFilename = "T_720KB": m_lngNumberOfBytes = 730112
         Case Else: m_strFilename = "":        m_lngNumberOfBytes = 0
  End Select

End Sub

Private Sub cmdExit_Click()

' ---------------------------------------------------------------------------
' set processing switches
' ---------------------------------------------------------------------------
  g_blnStopProcessing = True
  cFSO.CancelProcessing = True
  cRnd.CancelProcessing = True
  DoEvents

' ---------------------------------------------------------------------------
' Unload this form.  Now go to Form_QueryUnload() event
' ---------------------------------------------------------------------------
  Unload Me
  
End Sub

Private Sub cmdStart_Click()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim dblRecordLength    As Double
  Dim dblTotalDiskspace  As Double
  Dim dblFreeSpace       As Double
  Dim dblUsedSpace       As Double
  Dim lngDriveType       As Long
  Dim strMsg             As String
  
  Const MAX_FILE_LENGTH  As Long = 1024000000  ' 1 gb
  
' ---------------------------------------------------------------------------
' Are we starting or stopping?
' ---------------------------------------------------------------------------
  If cmdStart.Caption = "Start" Then
      cmdStart.Caption = "Stop"
      
      ' set processing switches
      g_blnStopProcessing = False
      cFSO.CancelProcessing = False
      cRnd.CancelProcessing = False
      DoEvents
  Else
      cmdStart.Caption = "Start"
      
      ' set processing switches
      g_blnStopProcessing = True
      cFSO.CancelProcessing = True
      cRnd.CancelProcessing = True
      DoEvents
      Exit Sub
  End If

' ---------------------------------------------------------------------------
' Initialize variables
' ---------------------------------------------------------------------------
  m_lngNumberOfBytes = 0
  dblTotalDiskspace = 0
  dblFreeSpace = 0
  dblUsedSpace = 0
  strMsg = ""                            ' clear the message area
  
' ---------------------------------------------------------------------------
' Verify we have a file size selected
' ---------------------------------------------------------------------------
  If m_blnPredefined Then
      cboSize_Click  ' verify we have data
  End If
  
' ---------------------------------------------------------------------------
' Do we have to convert to 2-char hex representation
' ---------------------------------------------------------------------------
  If optHex(0).Value Then
      m_blnHexConvert = True   ' convert to 2 char hex codes
  Else
      m_blnHexConvert = False  ' Do not convert to 2 char hex codes
  End If

' ---------------------------------------------------------------------------
' Test record sizes
' ---------------------------------------------------------------------------
  If m_blnPredefined Then
      ' drop thru
  Else
      If Len(Trim$(txtLength.Text)) = 0 Or txtLength.Text = "0" Then
             MsgBox "Will not build a zero byte file.     ", _
                     vbExclamation + vbOKOnly, "Invalid file size"
          GoTo Normal_Exit
      End If

      ' Test record sizes for values
      If Val(txtLength.Text) < 4 And optSize(0).Value = True Then
             MsgBox "Minimum size file is 4 bytes.     ", _
                    vbExclamation + vbOKOnly, "Invalid file size"
          GoTo Normal_Exit
      End If

      ' See if a valid record length has been entered
      ' Test input if user opted to custimize the file length
      ' calculate the file size
      If optSize(0).Value = True Then                          ' bytes
          dblRecordLength = CDbl(txtLength.Text)
          m_strFilename = "T_Bytes"
      ElseIf optSize(1).Value = True Then                      ' kilobytes
          dblRecordLength = CDbl(txtLength.Text) * 1024
          m_strFilename = "T_" & Trim$(txtLength.Text) & "Kb"
      Else                                                     ' megabytes
          dblRecordLength = (CDbl(txtLength.Text) * 1024) * 1000
          m_strFilename = "T_" & Trim$(txtLength.Text) & "Mb"
      End If
      
      m_lngNumberOfBytes = CLng(dblRecordLength)
      m_strRecSize = Format$(m_lngNumberOfBytes, "#,0") & " byte file"
  
      ' if file size exceeds 1gb, display a message
      ' then exit this routine
      If dblRecordLength > MAX_FILE_LENGTH Then
          strMsg = "You requested " & Format$(dblRecordLength, "#,0") & " bytes."
          strMsg = strMsg & vbCrLf & vbCrLf
          strMsg = strMsg & "Cannot build a file greater than " & _
          strMsg = strMsg & Format$(MAX_FILE_LENGTH, "#,0") & " bytes."
          MsgBox strMsg, vbExclamation + vbOKOnly, "Invalid file size"
          GoTo Normal_Exit
      End If
  End If

' ---------------------------------------------------------------------------
' Make sure the destination drive is available.
' ---------------------------------------------------------------------------
  If Not cFSO.Drive_Exist(m_strDriveLtr, lngDriveType) Then
      strMsg = "Drive " & m_strDriveLtr & " is not available at this time."
      MsgBox strMsg, vbExclamation + vbOKOnly, "Invalid drive"
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Test the drive type
' ---------------------------------------------------------------------------
  If lngDriveType < 1 Or lngDriveType > 4 Then
      strMsg = "Drive " & m_strDriveLtr & " is not available."
      MsgBox strMsg, vbExclamation + vbOKOnly, "Invalid drive"
      GoTo Normal_Exit
  End If
    
' ---------------------------------------------------------------------------
' Get the available disk space of destination drive
' ---------------------------------------------------------------------------
  If Not cFSO.Get_Disk_Space(m_strDriveLtr, dblTotalDiskspace, dblFreeSpace, dblUsedSpace) Then
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Test available disk space against the file size requested
' ---------------------------------------------------------------------------
  If m_lngNumberOfBytes > dblFreeSpace Then
      strMsg = "Your requested file is " & Format$(m_lngNumberOfBytes, "#,0") & " bytes.  "
      strMsg = strMsg & vbCrLf & "Drive " & m_strDriveLtr & " only has "
      strMsg = strMsg & Format$(dblFreeSpace, "#,0") & " bytes available.  "
      strMsg = strMsg & vbCrLf & vbCrLf & "Cleanup drive " & m_strDriveLtr
      strMsg = strMsg & " of any unwanted files first."
      MsgBox strMsg, vbExclamation + vbOKOnly, "Not enough space"
      GoTo Normal_Exit
  End If
  
' ---------------------------------------------------------------------------
' Is this area restricted?
' ---------------------------------------------------------------------------
  If cFSO.IsThisRestricted(m_strDriveLtr) Then
      strMsg = "Drive " & m_strDriveLtr & " is not available."
      MsgBox strMsg, vbExclamation + vbOKOnly, "Invalid drive"
      GoTo Normal_Exit
  End If
    
' ---------------------------------------------------------------------------
' Format the filename to point to the root directory of the destination drive
' ---------------------------------------------------------------------------
  m_strFilename = m_strDriveLtr & m_strFilename & ".dat"
  lblByteCnt.Caption = "Working on " & m_strRecSize
  Screen.MousePointer = vbHourglass
 
' ---------------------------------------------------------------------------
' Build the test file
' ---------------------------------------------------------------------------
  If m_blnFixed Then
      ' create fixed length records
      If Not Build_Fixed(m_strFilename, m_lngNumberOfBytes, _
                      m_blnHexConvert, m_blnUseKeyboardChars) Then
          GoTo Clean_Up
      End If
  Else
      ' build one contiguous record
      If Not Build_Continuous(m_strFilename, m_lngNumberOfBytes, _
                          m_blnHexConvert, m_blnUseKeyboardChars) Then
          GoTo Clean_Up
      End If
  End If
  
' ---------------------------------------------------------------------------
' Reset mouse pointer and display message
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbNormal
  lblByteCnt.Caption = m_strFilename & " = " & Format$(m_lngNumberOfBytes, "#,#") & " bytes"
  MsgBox vbCrLf & "Finished building " & vbCrLf & vbCrLf & m_strFilename

Normal_Exit:
' ---------------------------------------------------------------------------
' Change the caption on the Stop/Start button
' ---------------------------------------------------------------------------
  If cmdStart.Caption = "Start" Then
      cmdStart.Caption = "Stop"
  Else
      cmdStart.Caption = "Start"
  End If

Clean_Up:
' ---------------------------------------------------------------------------
' Reenable exit Button
' ---------------------------------------------------------------------------
  Screen.MousePointer = vbNormal
  
End Sub

Private Sub Form_Initialize()

' ---------------------------------------------------------------------------
' Center form on the screen.  I use this statement here because of a bug in
' the Form property "Startup Position".  In the VB IDE, under
' Tools\Options\Advanced, when you place a checkmark in the SDI Development
' Environment check box and set the form property to startup in the center
' of the screen, it works while in the IDE.  Whenever you leave the IDE, the
' property reverts back to the default [0-Manual].  This is a known bug with
' Microsoft.
' ---------------------------------------------------------------------------
  Me.Move (Screen.Width - Me.Width) \ 2, (Screen.Height - Me.Height) \ 2

End Sub

Private Sub Form_Load()

' ---------------------------------------------------------------------------
' Define local variables
' ---------------------------------------------------------------------------
  Dim strLblInfo  As String

' ---------------------------------------------------------------------------
' Load the combo box
' ---------------------------------------------------------------------------
  m_blnLoading = True
  Load_Combo_Boxes
  m_blnLoading = False
  
' ---------------------------------------------------------------------------
' initialize variables
' ---------------------------------------------------------------------------
  m_intSizeIndex = 0
  strLblInfo = ""
  m_strFilename = "T_1KB"
  m_strRecSize = "1 kb (1024 bytes)"
  m_blnFixed = True
  m_blnPredefined = True
  m_blnUseKeyboardChars = True
  m_blnHexConvert = False
  g_blnStopWasPressed = False
  
  strLblInfo = strLblInfo & "Create an ASCII text test file built with a series of random "
  strLblInfo = strLblInfo & "generated characters.  Choose between an 80-byte fixed record "
  strLblInfo = strLblInfo & "length or one continuous record.  You can also customize the "
  strLblInfo = strLblInfo & "size of the file and the character base.  Maximum file size is "
  strLblInfo = strLblInfo & "one gigabyte (1,024,000,000)."
  
' ---------------------------------------------------------------------------
' Initialize screen
' ---------------------------------------------------------------------------
  With frmBldRnd
      .Caption = g_strVersion
      .Label2.Caption = "Freeware by Kenneth Ives" & vbCrLf & "kenaso@home.com"
      .lblInfo.Caption = strLblInfo
      .lblByteCnt.Caption = m_strByteCaption
      '
      .optType(0).Value = True
      .optType(1).Value = False
      '
      .optChars(0).Value = False
      .optChars(1).Value = True
      '
      .optHex(0).Value = False
      .optHex(1).Value = True
      '
      .optRecLength(0).Value = True
      .optRecLength(1).Value = False
      '
      .optSize(0).Value = True
      .optSize(1).Value = False
      .optSize(2).Value = False
      '
      .optSize(0).Caption = "Bytes" & Space$(16) & "(Minimum 4 bytes)"
      .optSize(1).Caption = "Kilobytes" & Space$(11) & "(Input * 1024)"
      .optSize(2).Caption = "Megabytes" & Space$(8) & "(Input * 1024) * 1000"
      '
      .fraCustom.Visible = False
      .fraCustom.Enabled = False
      '
      .fraPredefined.Visible = True
      .fraPredefined.Enabled = True
      '
      cboDrive_Click
      cboSize_Click
      '
      .Show vbModeless               ' display the screen with no flicker
      .Refresh
  End With
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

' ---------------------------------------------------------------------------
' Based on the the unload code the system passes, we determine what to do.
'
' Unloadmode codes
'     0 - Close from the control-menu box or Upper right "X"
'     1 - Unload method from code elsewhere in the application
'     2 - Windows Session is ending
'     3 - Task Manager is closing the application
'     4 - MDI Parent is closing
' ---------------------------------------------------------------------------
  Select Case UnloadMode
         Case 0: StopApplication
         Case Else: ' Fall thru. Something else is shutting us down.
  End Select

End Sub

Private Sub optChars_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Set the appropriate switch based on option choice
' ---------------------------------------------------------------------------
  If optChars(0).Value = True Then
      m_blnUseKeyboardChars = False    ' use all ASCII codes (0 to 255)
  Else
      m_blnUseKeyboardChars = True     ' use only values 33 to 126
  End If

End Sub

Private Sub optHex_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Set the appropriate switch based on option choice
' ---------------------------------------------------------------------------
  If optHex(0).Value = True Then
      m_blnHexConvert = True   ' convert to 2 char hex codes
  Else
      m_blnHexConvert = False  ' Do not convert to 2 char hex codes
  End If

End Sub

Private Sub optRecLength_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Set the appropriate switch based on option choice
' ---------------------------------------------------------------------------
  If optRecLength(0).Value = True Then
      m_blnPredefined = True              ' use the combo box
      fraCustom.Visible = False           ' hide the custom frame
      fraCustom.Enabled = False
      fraPredefined.Visible = True        ' display the combo box frame
      fraPredefined.Enabled = True
  Else
      m_blnPredefined = False             ' do not use the combo box
      fraCustom.Visible = True            ' display the custom frame
      fraCustom.Enabled = True
      fraPredefined.Visible = False       ' hide the combo box frame
      fraPredefined.Enabled = False
  End If
  
  lblByteCnt.Caption = m_strByteCaption
  
End Sub

Private Sub optSize_Click(Index As Integer)

' ---------------------------------------------------------------------------
' see what was selected
' ---------------------------------------------------------------------------
  Select Case Index
         Case 0
              optSize(0).Value = True
              optSize(1).Value = False
              optSize(2).Value = False
         
         Case 1
              optSize(0).Value = False
              optSize(1).Value = True
              optSize(2).Value = False
         
         Case 2
              optSize(0).Value = False
              optSize(1).Value = False
              optSize(2).Value = True
  End Select
  
End Sub

Private Sub optType_Click(Index As Integer)

' ---------------------------------------------------------------------------
' Set the appropriate switch based on option choice
' ---------------------------------------------------------------------------
  If optType(0).Value = True Then
      m_blnFixed = True      ' use an 80-char fixed length record
  Else
      m_blnFixed = False     ' use one contiguous record
  End If
  
End Sub

Private Sub txtLength_KeyPress(KeyAscii As Integer)

' ---------------------------------------------------------------------------
' if ENTER or the TAB key is pressed then nullify the
' keystroke, tab to the next control and exit this routine.
' ---------------------------------------------------------------------------
  If KeyAscii = 13 Or KeyAscii = 9 Then
      KeyAscii = 0
      SendKeys "{TAB}"
      Exit Sub
  End If
  
' ---------------------------------------------------------------------------
' Save only numeric and the backspace character
' ---------------------------------------------------------------------------
  Select Case KeyAscii
         Case 8, 48 To 57: ' valid entry
         Case Else: KeyAscii = 0
  End Select
  
End Sub

