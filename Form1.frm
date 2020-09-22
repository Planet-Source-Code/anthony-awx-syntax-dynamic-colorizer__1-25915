VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Colorize Project"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "View File"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6570
      TabIndex        =   3
      Top             =   1005
      Width           =   1125
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   4245
      Left            =   240
      TabIndex        =   0
      Top             =   1410
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7488
      _Version        =   393217
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Filename that contains Keywords to be colorized: KEYWORDS.TXT"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   285
      TabIndex        =   2
      Top             =   1065
      Width           =   6300
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":00D0
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   285
      TabIndex        =   1
      Top             =   135
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// ALL VARIABLES MUST BE DECLARED                          //
    Option Explicit
    

Private Sub Command1_Click()

'// CALL NOTEPAD TO SHOW KEYWORDS TEXT FILE                 //
    Dim Dummy
    Dummy = ShellExecute(Me.hwnd, vbNullString, _
                         CurDir & "\keywords.txt", _
                         vbNullString, "c:\", 1)
    
End Sub

Private Sub Form_Load()

'// LOAD KEYWORDS FROM TEXT FILE                            //
'// YOU CAN ADD YOUR OWN WORDS/MAKE YOUR OWN                //
'// KEYWORDS FILE. JUST ENTER ONE KEYWORD PER LINE          //
    doGetScriptKeywords
    
End Sub

Private Sub RTF1_KeyUp(KeyCode As Integer, Shift As Integer)
    
'// SETUP LOCAL VARIABLES                                   //
    Dim CommentColor As Long
    Dim StringColor As Long
    Dim KeysColor As Long
    
'// ELIMINATE REFRESH BLINK                                 //
    LockWindowUpdate Me.hwnd
    
'// PERFORM "COLORIZE" ROUTINE ON SPECIFIC KEYPRESSES       //
'// (THIS EXAMPLE USES "ENTER" and "SPACE"                  //
    Select Case KeyCode
    
    Case 13, 32
    '// SET UP COLORS FOR KEYWORDS, COMMENTS, AND NORMAL TEXT/
    '// BTW, THESE ARE COLORS VB USES BY DEFAULT FOR ITS IDE /
        CommentColor = RGB(0, 128, 0)       '// DARK GREEN  //
        StringColor = RGB(0, 0, 0)          '// BLACK       //
        KeysColor = RGB(0, 0, 128)          '// DARK BLUE   //
        
    '// COLORIZE THE TEXT                                   //
    '// THIS ROUTINE USES THE "Current Line"                //
    '// TO KEEP UPDATE FAST, ESPECIALLY WHEN USING LARGE    //
    '// KEYWORD FILES                                       //
    '//                                                     //
        Colorize RTF1, CommentColor, StringColor, _
                 KeysColor, KeyCode
            
    '// SET COLOR OR CURRENT CURSOR POSITION BACK TO NORMAL //
    '// SO NEXT TYPED CHARACTER IS NORMAL COLOR AND NOT     //
    '// LAST COLORIZED COLOR                                //
            RTF1.SelColor = StringColor
        
    End Select

'// UNLOCK UPDATE TO SHOW COLORIZED RESULTS                 //
    LockWindowUpdate 0&
        
End Sub

Private Sub RichTextBox1_Change()

End Sub
