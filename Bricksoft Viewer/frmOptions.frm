VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Viewer Options"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1980
      Top             =   450
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSelectFont 
      Caption         =   "Select &Font"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   600
      Width           =   1065
   End
   Begin VB.TextBox txtMaxLines 
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   270
      Width           =   705
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   2850
      TabIndex        =   1
      Top             =   510
      Width           =   825
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Default         =   -1  'True
      Height          =   465
      Left            =   2850
      TabIndex        =   0
      Top             =   30
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum number of Lines to Read"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   2460
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bNotCanceled As Boolean
Private mFontName As String
Private mFontSize As Long
Private mFontBold As Boolean
Private mFontItalic As Boolean

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdSelectFont_Click()
   ' basic select-font procedure
   With dlg
      .CancelError = True
      .DialogTitle = "Select Display Font"
      
      .FontBold = mFontBold
      .FontItalic = mFontItalic
      .FontName = mFontName
      .FontSize = mFontSize
      
      .Flags = cdlCFBoth
      
      On Error Resume Next
      .ShowFont
      If Err.Number = 0 Then
         mFontBold = .FontBold
         mFontItalic = .FontItalic
         mFontName = .FontName
         mFontSize = .FontSize
      End If
      On Error GoTo 0
   End With
End Sub

Private Sub cmdUpdate_Click()
   ' save the font properties
   SaveSetting App.Title, P_SETTINGS, P_FONTBOLD, mFontBold
   SaveSetting App.Title, P_SETTINGS, P_FONTITALIC, mFontItalic
   SaveSetting App.Title, P_SETTINGS, P_FONTNAME, mFontName
   SaveSetting App.Title, P_SETTINGS, P_FONTSIZE, mFontSize
   
   SaveSetting App.Title, P_SETTINGS, P_MAXLINES, txtMaxLines.Text
   
   bNotCanceled = True
   Unload Me
End Sub

Public Property Get NotCanceled() As String
   NotCanceled = bNotCanceled
End Property

Private Sub Form_Load()
   mFontBold = GetSetting(App.Title, P_SETTINGS, P_FONTBOLD, False)
   mFontItalic = GetSetting(App.Title, P_SETTINGS, P_FONTITALIC, False)
   mFontName = GetSetting(App.Title, P_SETTINGS, P_FONTNAME, "Courier New")
   mFontSize = GetSetting(App.Title, P_SETTINGS, P_FONTSIZE, 9)
   
   txtMaxLines.Text = GetSetting(App.Title, P_SETTINGS, P_MAXLINES, 1000)
End Sub
