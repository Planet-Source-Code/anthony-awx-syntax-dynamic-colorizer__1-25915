Attribute VB_Name = "modCodeColorize"
'// THIS API CALL IS UED ONLY FOR PURPOSES OF ALLOWING      //
'// YOU TO VIEW THE TEXT FILE IN NOTEPAD, AND IS NOT        //
'// NECESSARY OTHERWISE                                     //
Declare Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hwnd As Long, _
         ByVal lpOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) _
         As Long


'// ALL VARIABLES MUST BE DECLARED                          //
    Option Explicit
    
'// SETUP PUBLIC VARIABLES                                  //
    Public KeyWords
    Public Declare Function LockWindowUpdate Lib "user32" _
           (ByVal hwndLock As Long) As Long
           
                                
Public Sub Colorize(RTFBox As RichTextBox, CommentColor, _
                    StringColor, KeysColor, KeyCode)

'// SETUP LOCAL VARIABLES                                   //
    Dim lTextSelPos As Long
    Dim lTextSelLen As Long
    Dim thisLine As Integer
    Dim cStart As Integer
    Dim cEnd As Integer
    Dim i As Long
    Dim sBuffer As String
    Dim lBufferLen As Long
    Dim lSelPos As Long
    Dim lSelLen As Long
    Dim sTempBuffer As String
    Dim sSearchChar As String
    Dim lSearchCharLen As Long
    Dim StartText As Integer
    Dim RepText As String
    
    
'// HANDLE ERRORS                                           //
    On Error GoTo ErrHandler

'// Save the cursor position                                //
    lTextSelPos = RTFBox.SelStart
    lTextSelLen = RTFBox.SelLength


'// MAKE ENTIRE TEXT BLACK (OR STRINGCOLOR DEFINED)         //
'// SO WHEN USER CHANGES A KEYWORD TO A NON-KEYWORD,        //
'// IT WILL NOT STILL BE BLUE (OR DEFINED KEYWORD COLOR)    //
    With RTFBox
        
    '// ONLY CHANGE CHARS ON THIS LINE                      //
        cStart% = .SelStart     ' CURRENT POSITION OF CURSOR//
        cEnd% = .SelStart       ' AGAIN CURR POS OF CURSOR  //
        
    '// SET "thisLine%" TO THE LINE CURSOR IS ON            //
        thisLine = .GetLineFromChar(.SelStart)
        
    '// IF ENTER WAS KEY PRESSED, COLORIZE LINE ABOVE SINCE //
    '// OUR LOCATION IS NOW THE NEXT LINE                   //
    '// (KEYCODE 13 = [ENTER] KEY)                          //
        If KeyCode = 13 Then
            thisLine = thisLine - 1
            cStart% = cStart% - 1
        End If
        
    '// DETERMINE "cStart" or STARTING CHARACTER TO         //
    '// EVALUATE FOR COLORIZATION PROCESS                   //
    '// WE ARE GOING TO DO THIS BY COUNTING FROM THE        //
    '// CURRENT CURSOR POSITION BACKWARDS TO BEGINNING OF   //
    '// THE CURRENT LINE , OR TO THE BEGINNING OF THE FILE, //
    '// WHICHEVER COMES FIRST                               //
        Do Until .GetLineFromChar(cStart%) <> thisLine
            cStart% = cStart% - 1
            If cStart% < 0 Then
                cStart% = 0
                Exit Do
            End If
        Loop
    '// NOW WE ARE GOING TO DETERMINE THE "cEnd" OR ENDING  //
    '// CHARACTER OF OUR EVALUATION STRING TO COLORIZE.     //
    '// WE DO THIS BY COUNTING FROM CURSOR POSITION TO THE  //
    '// END OF CURRENT LINE OR END OF THE FILE, WHICHEVER   //
    '// COMES FIRST.                                        //
    '//                                                     //
    '// THIS ROUTINE IS NECESSARY SINCE WE MAY BE INSERING  //
    '// TEXT IN THE MIDDLE OF A LINE, BUT WE STILL WANT TO  //
    '// EVALUATE ENTIRE LINE IN CASE WE ARE CHANGING A      //
    '// KEYWORD TO A NON-KEYWORD, OR VICE VERSA             //
        Do Until .GetLineFromChar(cEnd%) <> thisLine
            cEnd% = cEnd% + 1
            If cEnd% > Len(.Text) Then
                cEnd = Len(.Text)
                Exit Do
            End If
        Loop
        
    '// SET COLOR OF TEXT WE ARE WORKING WITH BACK TO       //
    '// ORIGINAL COLOR FOR NOW SINCE IT MAY BE A KEYWORK    //
    '// THAT IS GETTING CHANGED TO A NON-KEYWORD. THE       //
    '// NEXT ROUTINE WILL COLORIZE IT IF IT FINDS KEYWORDS  //
        .SelStart = cStart%
        .SelLength = cEnd% - cStart%
        .SelColor = StringColor
        .SelLength = 0
        
    End With
    


'// BEGIN EVALUAINTG AND CHANGING COLORS OF WORDS           //
    With RTFBox
    '// INSURE "WHOLE WORDS" ARE COLORIZED AND NOT          //
    '// PARTIAL WORDS (EG: the word "If", and not "Gift"    //
    '// WHERE "IF" EXISTS IN "GIFT"                         //
        sBuffer = .Text & " "
        lBufferLen = Len(sBuffer)
        sTempBuffer = ""
        If cStart = 0 Then cStart = 1
        
    '// LOOP THROUGH CHARACTERS USING RANGE WE DEFINED      //
    '// EARLIER IN THIS SUB                                 //
        For i = cStart% To cEnd%
        
        Select Case Asc(Mid(sBuffer, i, 1))
        
    '// COMMENTS - ENTIRE LINE IS COLORIZED REGARDLESS OF   //
    '// CONTENT. COMMENT PREFIXES ARE HARD-CODED HERE, BUT  //
    '// YOU CAN MODIFY/ADD/REMOVE FROM HERE IF YOU WANT     //
    '// BY FIRST INCLUDING PREFIX ASC CODE IN THIS "CASE"   //
    '// STATEMENT, AND THEN WRITING AN ElseIf EVALUATION    //
    '// AGAINST THE CHARACTER(S) THAT MAKE UP YOUR REMARK   //
    '// PREFIX                                              //
    '// (CHR$(47) = "/", AND CHR$(39) = "'"                 //
        Case 47, 39
          
    '// C/C++ STYLE COMMENT                                 //
            If Mid(sBuffer, i, 2) = "//" Then
        '// COLORIZE FROM PREFIX TO "sSearchChar"           //
                sSearchChar = vbCrLf
                lSearchCharLen = 0
    '// VISUAL BASIC STYLE COMMENT                          //
            ElseIf Mid(sBuffer, i, 1) = "'" Then
        '// COLORIZE FROM PREFIX TO "sSearchChar"           //
                sSearchChar = vbCrLf
                lSearchCharLen = 0
    '// IF NOT A COMMENT, GOTO THE "EXITCOMMENT" ROUTINE    //
    '// TO BYPASS COMMENT COLORIZATION ROUTINES             //
            Else
                GoTo ExitComment
            End If
          
    '// SET TEMPBUFFER (sTempBuffer" to NOTHING             //
            sTempBuffer = ""
          
    '// COLORIZE THE COMMENT STRING                         //
    '// "i" IS CURRENT COUNT OF LOOP                        //
            .SelStart = i - 1
            lSelLen = InStr(i, sBuffer, sSearchChar) _
                    + lSearchCharLen
                
            If lSelLen <> lSearchCharLen Then   '// FileEnd?//
                lSelLen = lSelLen - i
            Else
                lSelLen = lBufferLen - i
            End If
                
            .SelLength = lSelLen
            .SelColor = CommentColor
            i = .SelStart + .SelLength
          
'// QUOTE COLORIZE ROUTINE                                  //
ExitComment:
        Case 34
            If Mid(sBuffer, i, 1) = Chr$(34) Then
        '// COLORIZE FROM PREFIX TO "sSearchChar"           //
                sSearchChar = Chr$(34)
                lSearchCharLen = 0
            Else
                GoTo ExitQuote
            End If
          
    '// SET TEMPBUFFER (sTempBuffer" to NOTHING             //
            sTempBuffer = ""
          
    '// COLORIZE THE QUOTE STRING                           //
    '// "i" IS CURRENT COUNT OF LOOP                        //
            .SelStart = i - 1
            lSelLen = InStr(i + 1, sBuffer, sSearchChar) _
                    + lSearchCharLen
                
            If lSelLen <> lSearchCharLen Then   '// FileEnd?//
                lSelLen = lSelLen - i
            ElseIf lSelLen < 1 Then
            '// SET CUR POSITION BACK AND DONT COLORIZE     //
            '// ANYTHING SINCE "END QUOTE" HAS NOT BEEN     //
            '// ENTERED YET                                 //
                GoTo ErrHandler
            Else
                lSelLen = lBufferLen - i
            End If
                
            .SelLength = lSelLen
            .SelColor = StringColor
            i = .SelStart + .SelLength

'// COLORIZE KEYWORDS ROUTINE                               //
ExitQuote:

        '// THE FOLLOWING "CASE" STATEMENT SETS THESE       //
        '// CHARACTERS AS VALID PARTS OF A COLORIZATION     //
        '// STRING. IN OTHER WORDS, ANY KEYWORDS YOU DEFINE //
        '// CAN HAVE THESE ASCII CHARACTERS, AND IF THE     //
        '// DONT, THEY WILL NOT QUALIFY.                    //
        '// EXAMPLE: IF YOUR KEYWORD IS SOMETHING LIKE      //
        '// THIS:  My_ROUTINE (with the underscore), YOU    //
        '// NEED TO MAKE SURE THIS CASE STATEMENT INCLUDES  //
        '// THE ASCII CODE FOR THAT CHARACTER (UNDERSCORE)  //
        '// AS WELL AS ALL ALPHANUMERICK CHARACTERS         //
        '// ASCII 33 = "!", 35 to 38 = #, $, %, &           //
        '// ASCII 46 = . (dot)                              //
        '// ASCII 60 = "<" and 62 = ">"                     //
        '// ASCII 49 to 57 = Numbers 1,2,3,4,5,6,7,8,9,0    //
        '// ASCII 97 to 122 = lowercase a to z              //
        '// ASCII 65 to 90 = UPPERCASE A to Z               //
             Case 33, 35 To 38, 46, 60, 62, _
                  49 To 57, 97 To 122, 65 To 90
                  
                If sTempBuffer = "" Then lSelPos = i
                sTempBuffer = sTempBuffer & Mid(sBuffer, i, 1)
             
             Case Else
                
                If Trim(sTempBuffer) <> "" Then
                    .SelStart = lSelPos - 1
                    .SelLength = Len(sTempBuffer)
                    StartText% = InStr(1, KeyWords, _
                                 "|" & sTempBuffer & "|", 1)
                    If StartText% <> 0 Then
                '// ALTER COLOR                             //
                        .SelColor = KeysColor
                        
                
                '// CHANGE FOUND MATCH TO BE THE SAME       //
                '// CASE AS WORD IN LIBRARY                 //
                '// (EG: "print" would change to "Print"    //
                '//  with the CAPITAL "P")                  //
                        RepText$ = _
                        Mid$(KeyWords, StartText% + 1, _
                        Len(sTempBuffer))
                        .SelText = RepText$
                    End If
                
                End If
                
                sTempBuffer = ""
        
        
        End Select
      
        Next

        End With

ErrHandler:

    '// Set the Cursor to the old position                  //
    RTFBox.SelStart = lTextSelPos
    RTFBox.SelLength = lTextSelLen

End Sub

Public Sub doGetScriptKeywords()

'// THIS ROUTINE GRABS KEYWORDS FROM A FILE. EXPECTED       //
'// IS ONE KEYWORD ON EACH LINE, EACH FOLLOWED (AND         //
'// THEREFORE DELIMITED) BY A HARD RETRUN (VbCrLf)          //
'//                                                         //
'// SEE INCLUDED SAMPLE FILE FOR EXAMPLE                    //
'// FILE NAMED "KEYWORDS.TXT"                               //

'// SET UP LOCAL VARIABLES
    Dim DataFile As String
    Dim sWords As String
    Dim LineData As String
    Dim ff As Integer
        
'// DEFINE FILENAME WHICH HOLDS KEYWORDS                    //
    DataFile$ = CurDir & "\keywords.txt"
    
    
'// DEFINE START OF "sWords", WHICH IS THE TEMPORARY        //
'// LOCAL VARIABLE WE USE TO BUILD THE MEMORY-RESIDENT      //
'// STRING WHICH HOLDS ALL KEYWORDS. THIS SYMBOL IS WHAT    //
'// WE USE AS A DELIMITER FOR OUR MEMORY-RESIDENT STRING    //
'// WHICH SEPARATES EACH KEYWORD                            //
        sWords$ = "|"
    
'// OPEN FILE AND GRAB KEYWORDS                             //
        ff = FreeFile
        Open DataFile$ For Input As ff
        Do Until EOF(ff)
        Line Input #ff, LineData$
        sWords$ = sWords$ & LineData$ & "|"
        Loop
        Close ff
'// "KeyWords" IS A PUBLIC VARIABLE WHICH HOLDS STRING      //
'// CONTAINING ALL KEYWORDS, DELIMITED BY A "|" CHARACTER   //
        KeyWords = sWords$
        
'// FREE VARIABLES                                          //
        sWords$ = ""
        LineData$ = ""

End Sub
