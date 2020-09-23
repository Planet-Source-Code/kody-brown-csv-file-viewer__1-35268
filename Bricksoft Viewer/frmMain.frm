VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "CSV Parser/Viewer"
   ClientHeight    =   5190
   ClientLeft      =   975
   ClientTop       =   4035
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   7905
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6960
      Top             =   510
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   6900
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtf 
      CausesValidation=   0   'False
      Height          =   4545
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   8017
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      MaxLength       =   6400000
      RightMargin     =   6.00000e5
      TextRTF         =   $"frmMain.frx":0000
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
   Begin VB.CommandButton Command1 
      Caption         =   "invisible"
      Height          =   225
      Left            =   6750
      TabIndex        =   1
      Top             =   1200
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenFile 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileDash100 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Op&tions"
      End
      Begin VB.Menu mnuFileDash150 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------]
' Bricksoft CSV Viewer -                                             ]
'--------------------------------------------------------------------]
' December 12, 2001                                                  ]
'              Version 1.0.00                                        ]
'              Created By Kody Brown                                 ]
'              December 12, 2001                                     ]
'              Copyright 2001, 2002 Kody Brown. All Rights Reserved. ]
'                                                                    ]
' Copyright Notice: You are free to use and change this code in any  ]
'              way, provided that you agree and abide by the         ]
'              following terms:                                      ]
'              a) use completely at your own risk. author is not     ]
'                 liable in any way.                                 ]
'              b) the original author's name must be included with   ]
'                 any changes.                                       ]
'                                                                    ]
' Description: This file/application will display (read-only) any    ]
'              CSV file. This application has replaced the need      ]
'              for Excel in many instances (and does not change      ]
'              the formatting of date columns for instance).         ]
'                                                                    ]
'--------------------------------------------------------------------]

'        1         2         3         4         5         6         7 _
1234567890123456789012345678901234567890123456789012345678901234567890
Option Explicit

Private mFilePtr As Integer
Private mFileCol As Collection ' used to hold the entire file (in mLineAr's)
Private mColumnLengths() As Long ' used to hold the max length of each column
Private mLineAr() As Variant ' used to hold a single row/record
Private mColumnCount As Long

Private mFileName As String
Private mFileType As EFileTypes

Private mLeftPos As Long
Private mTopPos As Long
Private mWidth As Long
Private mHeight As Long

Private MAXLINES As Long

Private Sub LoadFile()
   If Dir(mFileName, vbNormal) = "" And mFileName <> "" Then
      MsgBox "Invalid file name or path", vbCritical + vbOKOnly, "File not found"
      Exit Sub
   End If
   
   If mFileName <> "" Then
      Me.Caption = "Bricksoft Viewer - " & mFileName
   End If
   
   mFileType = CSVFileType
   
   ' load each line into a collection (of arrays)
   Dim Count As Long
   
   OpenFile
   
   Set mFileCol = New Collection
   
   Do While nextRow
      ' add each line (in array form - mLineAr) to
      ' the collection of lines
      mFileCol.Add mLineAr
      
      ' You can stop processing at MAXLINES, because
      ' it will really slow down on large files,
      ' or comment out this If statement to always
      ' read every line.
      Count = Count + 1
      If Count >= MAXLINES Then
         ReDim mLineAr(1)
         mLineAr(1) = "Stopped at " & MAXLINES
         mFileCol.Add mLineAr
         Exit Do
      End If
   Loop
   
   CloseFile
   
   Dim index As Long
   Dim obj As Variant
   Dim sLine As String
   
   ReDim mColumnLengths(1 To mColumnCount) As Long
   
   ' here we're going to figure out the maximum
   ' number of characters there are in each column
   ' and store it for padding each column
   For index = 1 To mColumnCount
      mColumnLengths(index) = GetMaxColumnLength(index)
   Next index
   
   ' set the focus to the hidden command button
   ' to prevent the rtf control from flickering
   ' while adding data.
   rtf.Text = ""
   On Error Resume Next
   Command1.SetFocus
   Err.Clear
   On Error GoTo 0
   
   ' add column number headers.
   sLine = ""
   For index = 1 To mColumnCount
      sLine = sLine & index & String((mColumnLengths(index) - Len(CStr(index))) + 3, " ") & "| "
   Next index
   rtf.SelText = sLine & vbCrLf
   
   ' add column header separator.
   sLine = ""
   For index = 1 To mColumnCount
      sLine = sLine & String(mColumnLengths(index) + 3, "-") & "| "
   Next index
   rtf.SelText = sLine & vbCrLf
   
   ' display each row, padding each column
   ' to the mColumnLengths()
   For Each obj In mFileCol
      sLine = ""
      For index = 1 To mColumnCount
         If index <= UBound(obj) Then
            sLine = sLine & obj(index) & String((mColumnLengths(index) - Len(obj(index))) + 3, " ") & "| "
         End If
      Next index
      rtf.SelText = sLine & vbCrLf
   Next
End Sub

Private Function GetMaxColumnLength(index As Long) As Long
   Dim i As Long
   Dim length As Long
   Dim obj As Variant
   
   length = 0
   
   ' Simply save the largest columns character count
   If index <= mColumnCount Then
      For Each obj In mFileCol
         If index <= UBound(obj) Then
            If Len(CStr(obj(index))) > length Then
               length = Len(obj(index))
            End If
         End If
      Next
   End If
   
   GetMaxColumnLength = length
End Function

Public Sub OpenFile()
   ' simple wrapper to the visual basic file access
   If mFileName = "" Or Dir(mFileName, vbNormal) = "" Then
      Err.Raise 20002, "FileLoader.OpenFile()", "You must set the FileName property before trying to open."
      Exit Sub
   End If
   
   If mFileType = 0 Then
      Err.Raise 20004, "FileLoader.OpenFile()", "You must set the FileType property before trying to open."
      Exit Sub
   End If
   
   mFilePtr = FreeFile()
   Open mFileName For Input As #mFilePtr
End Sub

Public Function EndOfFIle() As Boolean
   EndOfFIle = EOF(mFilePtr)
End Function

Public Sub CloseFile()
   Close #mFilePtr
End Sub

Public Function nextRow() As Boolean
   ' This is a simple wrapper for getting the next row
   ' or line from a file, without caring about the file
   ' type.
   
   ' I am currently adding support for TAB files as well.
   If mFileType = CSVFileType Then
      nextRow = CSVnextRow()
   'ElseIf mFileType = TABFileType Then
   '   nextRow = TABnextRow()
   Else
      nextRow = False
      Err.Raise 21001, "FileLoader.nextRow()", "This FileType is not currently supported"
   End If
End Function

Private Function CSVnextRow() As Boolean
   If EndOfFIle() Then
      CSVnextRow = False
      Exit Function
   End If
   '
   ' DOE,JOHN,9123 SOUTH PARK DRIVE,SANDY,UT,84092,8015551212,,1996,DODGE,STRATUS ES,13XXJ56XTXN1DD686,102712,34841,11292001,205,C,1,106,LOF,"LUBE, OIL AND FILTER",109.8,1,11141995,0,,0,2763,7041956,9602940,,8015554255
   '
   ' Updated Dec 28 2001 by Kody Brown.
   '
   ' I am adding support for columns that contain
   ' a double quote somewhere in the column data,
   ' but not the first position, as this would
   ' indicate the entire column to be enclosed in
   ' double quotes.
   '
   Dim sLine As String
   Dim sChar As String
   Dim sLastChar As String
   Dim bInString As Boolean
   Dim output As String
   Dim Count As Integer
   
   Count = 0
   
   Line Input #mFilePtr, sLine
   
   sChar = ELeft(sLine, 1)
   
   ' This is where a comma-separated line is parsed
   ' into an array (mLineAr). It is all done by hand
   ' because visual basic's intrinsic methods are
   ' inadaquate. For instance you can't tell when
   ' you've reached the end of a line, and there's no
   ' support for finding out how many solumns there
   ' are in that line. There are also bugs in their
   ' code in dealing with multiple quotes in a column,
   ' etc.
   
   Do While sChar <> ""
      Select Case sChar
         Case ","
            If bInString Then
               output = output & sChar
            Else
               Count = Count + 1
               ReDim Preserve mLineAr(Count)
               mLineAr(Count) = output
               output = ""
            End If
         Case """"
            If bInString And Left$(sLine, 1) <> """" Then
               Count = Count + 1
               ReDim Preserve mLineAr(Count)
               mLineAr(Count) = output
               output = ""
               bInString = False
               ' dump the comma, that follows
               sLastChar = sChar
               sChar = ELeft(sLine, 1)
            ElseIf bInString = False Then
               bInString = True ' do not insert the quotes into the output
            Else
               output = output & sChar
               sLastChar = sChar ' remove the next double quote
               sChar = ELeft(sLine, 1)
            End If
         Case Else
            output = output & sChar
            
      End Select
      
      sLastChar = sChar
      sChar = ELeft(sLine, 1)
   Loop
   
   ' catch the last entry of the line (if it exists)
   If output <> "" Then
      Count = Count + 1
      ReDim Preserve mLineAr(Count)
      mLineAr(Count) = output
      output = ""
   End If
   
   ' set the module-level column count variable
   If Count > mColumnCount Then
      mColumnCount = Count
   End If
   
   CSVnextRow = True
End Function

Public Function ELeft(sIn As String, lLength As Long) As String
'--------------------------------------------------------------------]
' ELeft                                                              ]
'--------------------------------------------------------------------]
' May 8, 1998                                                        ]
'              Created By Kody Brown                                 ]
'              of Bricksoft                                          ]
'              Copyright 1998 Bricksoft. All Rights reserved.        ]
'                                                                    ]
' Purpose:     This will return lLengthReturn of sIn, and remove     ]
'              it from the input string (sIn).                       ]
'                                                                    ]
' Effects:     This WILL CHANGE the input string (sIn).              ]
'                                                                    ]
' Returns:     String containing the lLengthReturn of sIn.           ]
'                                                                    ]
' Example:     Msg = "Hello World!"                                  ]
'              sReturn = ELeft$(Msg, 5)                              ]
'                                                                    ]
'              Msg = " World!" AND sReturn = "Hello"                 ]
'                                                                    ]
' June 8, 1998                                                       ]
'              I removed the '$' from the procedure names..          ]
'                                                                    ]
' December 22, 1998                                                  ]
'              I added some error checking on the lengths.           ]
'              I re-made the banner, and fixed the example.          ]
'              I cleaned up the code a little bit, creating a        ]
'              single exit point and clearer flow.                   ]
'                                                                    ]
'--------------------------------------------------------------------]
   On Error GoTo ExitNormal
   ELeft = ""
   
   If Len(sIn) = 0 Or lLength = 0 Then
      GoTo ExitNormal
   End If
   
   If lLength > Len(sIn) Then
      'I can not return more of sIn than there is!
      GoTo ExitNormal
   End If
   
   ELeft = Left$(sIn, lLength)
   sIn = Right$(sIn, Len(sIn) - lLength)
   
ExitNormal:
End Function

Private Sub Form_Load()
   Me.Caption = "Bricksoft Viewer (ver " & App.Major & "." & App.Minor & "." & App.Revision & ")"
   
   mFileName = GetSetting(App.Title, P_SETTINGS, P_FILENAME, "")
   
   ' support a command-line argument for a single
   ' file, with or without double-quotes (for paths
   ' that include spaces in them).
   Dim args
   args = Command()
   If args <> "" Then
      If Left$(args, 1) = """" Then
         args = Right$(args, Len(args) - 1)
      End If
      If Right$(args, 1) = """" Then
         args = Left$(args, Len(args) - 1)
      End If
      mFileName = args
      LoadFile
   End If
   
   ' Great care has been taken in ensuring a professional
   ' appearance/interface including resizing and maximizing.
   ' Hence, the following code.
   
   ' We'll use model-level variables for the coordinates.
   mLeftPos = GetSetting(App.Title, P_SETTINGS, P_LEFT, Me.Left)
   Me.Left = mLeftPos
   
   mTopPos = GetSetting(App.Title, P_SETTINGS, P_TOP, Me.Top)
   Me.Top = mTopPos
   
   mWidth = GetSetting(App.Title, P_SETTINGS, P_WIDTH, Me.Width)
   Me.Width = mWidth
   
   mHeight = GetSetting(App.Title, P_SETTINGS, P_HEIGHT, Me.Height)
   Me.Height = mHeight
   
   Me.WindowState = GetSetting(App.Title, P_SETTINGS, "WindowState", Me.WindowState)
   
   ' Load the font properties from the registry.
   rtf.Font.Bold = GetSetting(App.Title, P_SETTINGS, P_FONTBOLD, False)
   rtf.Font.Italic = GetSetting(App.Title, P_SETTINGS, P_FONTITALIC, False)
   rtf.Font.Name = GetSetting(App.Title, P_SETTINGS, P_FONTNAME, "Courier New")
   rtf.Font.Size = GetSetting(App.Title, P_SETTINGS, P_FONTSIZE, 9)
   
   ' Load the maximum number of lines to read
   MAXLINES = GetSetting(App.Title, P_SETTINGS, P_MAXLINES, 1000)
End Sub

Private Sub Form_Resize()
   If Me.WindowState <> vbMinimized Then
      ' make sure the window is large enough for basic
      ' resizing of the text box.
      If Me.Width < 4000 Then
         Me.Width = 4000
      End If
      If Me.Height < 4000 Then
         Me.Height = 4000
      End If
      
      Const separator As Long = 15
      rtf.Move separator, separator, Me.ScaleWidth - (separator + separator)
      rtf.Height = Me.ScaleHeight - (separator + rtf.Top)
      
      ' Don't update the coordinate variables when maximized
      If Me.WindowState <> vbMaximized Then
         mLeftPos = Me.Left
         mTopPos = Me.Top
         mWidth = Me.Width
         mHeight = Me.Height
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   SaveSetting App.Title, P_SETTINGS, P_FILENAME, mFileName
   
   ' save the coordinate variables, not the window position,
   ' this prevents saving the windows position when it is
   ' maximized.
   SaveSetting App.Title, P_SETTINGS, P_LEFT, mLeftPos
   SaveSetting App.Title, P_SETTINGS, P_TOP, mTopPos
   SaveSetting App.Title, P_SETTINGS, P_WIDTH, mWidth
   SaveSetting App.Title, P_SETTINGS, P_HEIGHT, mHeight
   
   SaveSetting App.Title, P_SETTINGS, "WindowState", Me.WindowState
   
   ' NOTE: we don't have to save the font properties
   ' and MAXLINES because they are saved the registry
   ' at the only place they can be changed, in frmOptions.
End Sub

Private Sub mnuClose_Click()
   Unload Me
End Sub

Private Sub mnuOpenFile_Click()
   SelectFile
End Sub

Private Sub SelectFile()
   ' basic file-open procedure
   With dlg
      .CancelError = True
      .DialogTitle = "View CSV File"
      
      If mFileName <> "" And Dir(mFileName, vbNormal) <> "" Then
         .FileName = mFileName
      Else
         .FileName = "*.CSV"
      End If
      
      .Filter = "Comma-Separated Files (*.CSV)|*.CSV"
      .FilterIndex = 0
      .Flags = cdlOFNFileMustExist + cdlOFNExplorer
      
      On Error Resume Next
      .ShowOpen
      If Err.Number = 0 Then
         mFileName = .FileName
         LoadFile
      End If
      On Error GoTo 0
   End With
End Sub

Private Sub mnuOptions_Click()
   Dim f As frmOptions
   Set f = New frmOptions
   
   f.Show vbModal
   
   If f.NotCanceled Then
      rtf.Font.Bold = GetSetting(App.Title, P_SETTINGS, P_FONTBOLD, False)
      rtf.Font.Italic = GetSetting(App.Title, P_SETTINGS, P_FONTITALIC, False)
      rtf.Font.Name = GetSetting(App.Title, P_SETTINGS, P_FONTNAME, "Courier New")
      rtf.Font.Size = GetSetting(App.Title, P_SETTINGS, P_FONTSIZE, 9)
      
      MAXLINES = GetSetting(App.Title, P_SETTINGS, P_MAXLINES, 1000)
   End If
   
   Set f = Nothing
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   ' When the window is not maximized and not minimized
   ' save its position every second to the coordinate
   ' variables. This ensures that when the form is moved
   ' we capture the new position.
   If Me.WindowState = vbNormal Then
      mLeftPos = Me.Left
      mTopPos = Me.Top
      mWidth = Me.Width
      mHeight = Me.Height
   End If
End Sub
