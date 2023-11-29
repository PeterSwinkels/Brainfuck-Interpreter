Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API functions used by this program.
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long

'The constants used by this program.
Private Const MAX_STRING As Long = 65535    'The maximum length for a string buffer.

'This procedure loads the specified code or returns any currently loaded query.
Public Function Code(Optional Path As String = vbNullString) As String
On Error GoTo ErrorTrap
Dim FileHandle As Long
Static CurrentCode As String
Static CurrentPath As String

   If Not Path = vbNullString Then
      CurrentCode = vbNullString
      If Left$(Path, 1) = """" Then Path = Mid$(Path, 2)
      If Right$(Path, 1) = """" Then Path = Left$(Path, Len(Path) - 1)
      CurrentPath = Path
      
      FileHandle = FreeFile()
      Open CurrentPath For Input Lock Read Write As FileHandle: Close FileHandle
      Open CurrentPath For Binary Lock Read Write As FileHandle
         CurrentCode = Input$(LOF(FileHandle), FileHandle)
      Close FileHandle
   End If
   
EndRoutine:
   Code = CurrentCode
   Path = CurrentPath
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function
'This procedure adds text to the specified text box.
Private Sub Display(Text As String, TextBox As TextBox)
On Error GoTo ErrorTrap

   If InStr(Text, vbCrLf) = 0 Then
      If InStr(Text, vbCr) Then
         Text = Replace(Text, vbCr, vbCrLf)
      ElseIf InStr(Text, vbLf) Then
         Text = Replace(Text, vbLf, vbCrLf)
      End If
   End If
   
   Text = Escape(Text)
   
   With TextBox
      If Len(.Text & Text) > MAX_STRING Then
         .SelStart = 0
         .SelLength = Len(Text)
         .SelText = vbNullString
      End If
      .SelLength = 0
      .SelStart = Len(.Text)
      .SelText = .SelText & Text
   End With
   
EndRoutine:
   DoEvents
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure converts non-displayable characters in the specified text to escape sequences.
Public Function Escape(Text As String, Optional EscapeCharacter As String = "/", Optional EscapeLineBreaks As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Escaped As String
Dim Index As Long
Dim NextCharacter As String

   Escaped = vbNullString
   Index = 1
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         Escaped = Escaped & String$(2, EscapeCharacter)
      ElseIf Character = vbTab Or Character >= " " Then
         Escaped = Escaped & Character
      ElseIf Character & NextCharacter = vbCrLf And Not EscapeLineBreaks Then
         Escaped = Escaped & vbCrLf
         Index = Index + 1
      Else
         Escaped = Escaped & EscapeCharacter & String$(2 - Len(Hex$(Asc(Character))), "0") & Hex$(Asc(Character))
      End If
      Index = Index + 1
   Loop
   
EndRoutine:
   Escape = Escaped
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function



'This procedure checks whether the return value for escape sequence procedures indicates an error.
Public Function EscapeSequenceError(ErrorAt As Long) As Boolean
On Error GoTo ErrorTrap
Dim EscapeError As Boolean

EscapeError = (ErrorAt > 0)
   If EscapeError Then MsgBox "Bad escape sequence at character #" & CStr(ErrorAt) & ".", vbExclamation

EndRoutine:
   EscapeSequenceError = EscapeError
Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure executes the specified code.
Public Function Execute(Optional ExecuteCode As String = vbNullString, Optional NewOutputBox As TextBox = Nothing, Optional Abort As Boolean = False, Optional WarnUser As Boolean = False) As Boolean
On Error GoTo ErrorTrap
Dim Character As String
Dim InstructionP As Long
Dim Memory(&H1& To &H8000&) As Byte
Dim MemoryP As Long
Static Executing As Boolean
Static OutputBox As TextBox

   If WarnUser And Executing Then
      MsgBox "Code is currently being executed.", vbExclamation
   Else
      If (Not ((ExecuteCode = vbNullString) Or Executing)) Or Abort Then
         InputStream , Flush:=True
         Executing = Not Abort
      
         Loops , ExecuteCode
         If Not NewOutputBox Is Nothing Then Set OutputBox = NewOutputBox
         
         InstructionP = 1
         MemoryP = LBound(Memory())
         Do
            If Not Executing Then Exit Do
      
            Select Case Mid$(ExecuteCode, InstructionP, 1)
               Case ">"
                  If MemoryP = UBound(Memory()) Then MemoryP = LBound(Memory()) Else MemoryP = MemoryP + 1
               Case "<"
                  If MemoryP = LBound(Memory()) Then MemoryP = UBound(Memory()) Else MemoryP = MemoryP - 1
               Case "+"
                  If Memory(MemoryP) = &HFF& Then Memory(MemoryP) = &H0& Else Memory(MemoryP) = Memory(MemoryP) + &H1&
               Case "-"
                  If Memory(MemoryP) = &H0& Then Memory(MemoryP) = &HFF& Else Memory(MemoryP) = Memory(MemoryP) - &H1&
               Case "."
                  Display Chr$(Memory(MemoryP)), OutputBox
               Case ","
                  Do
                     If Not Executing Then Exit Do
      
                     Character = InputStream()
                     DoEvents
                  Loop While Character = vbNullString And Forms.Count > 0
                  
                  If Not Character = vbNullString Then Memory(MemoryP) = Asc(Character)
               Case "["
                  If Memory(MemoryP) = &H0& Then InstructionP = Loops(InstructionP)
               Case "]"
                  If Not Memory(MemoryP) = &H0& Then InstructionP = Loops(InstructionP)
            End Select
         
            InstructionP = InstructionP + 1
         Loop While InstructionP > 0 And InstructionP <= Len(ExecuteCode) And Forms.Count > 0

         Executing = False
         If Abort Then MsgBox "Finished execution.", vbInformation
      End If
   End If

EndRoutine:
   Execute = Executing
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number
   Err.Clear
   
   On Error Resume Next
   
   Description = "Error: " & CStr(ErrorCode) & vbCr & Description
   MsgBox Description, vbExclamation
End Sub

'This procedure manages the input line break.
Public Function InputLineBreak(Optional NewInputLineBreak As String = vbNullString, Optional Remove As Boolean = False) As String
On Error GoTo ErrorTrap
   Static CurrentInputLineBreak As String
   
   If Not NewInputLineBreak = vbNullString Then
      CurrentInputLineBreak = NewInputLineBreak
   ElseIf Remove Then
      CurrentInputLineBreak = vbNullString
   End If

EndRoutine:
   InputLineBreak = CurrentInputLineBreak
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the input stream.
Public Function InputStream(Optional NewInput As String = vbNullString, Optional Flush As Boolean = False) As String
On Error GoTo ErrorTrap
Dim Character As String
Static CurrentInput As String

   Character = vbNullString
   If Flush Then
      CurrentInput = vbNullString
   ElseIf Not NewInput = vbNullString Then
      CurrentInput = NewInput & CurrentInput
   ElseIf Not CurrentInput = vbNullString Then
      Character = Left$(CurrentInput, 1)
      CurrentInput = Mid$(CurrentInput, 2)
   End If
   
EndRoutine:
   InputStream = Character
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the user's most recent input.
Public Function LastInput(Optional NewInput As String) As String
On Error GoTo ErrorTrap
Static CurrentInput As String

   If Not NewInput = vbNullString Then CurrentInput = NewInput
   
EndRoutine:
   LastInput = CurrentInput
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the list of loop start/end addresses.
Private Function Loops(Optional LoopInstructionP As Long = -1, Optional Code As String = vbNullString) As Long
On Error GoTo ErrorTrap
Dim InstructionP As Long
Dim LoopStack() As Long
Dim NewInstructionP As Long
Static LoopList() As Long

   NewInstructionP = -1
   If Not Code = vbNullString Then
      ReDim LoopList(1 To Len(Code))
      
      For InstructionP = 1 To Len(Code)
         Select Case Mid$(Code, InstructionP, 1)
            Case "["
               If SafeArrayGetDim(LoopStack()) = 0 Then
                  ReDim LoopStack(0 To 0) As Long
               Else
                  ReDim Preserve LoopStack(LBound(LoopStack) To UBound(LoopStack) + 1) As Long
               End If
               
               LoopStack(UBound(LoopStack)) = InstructionP
            Case "]"
               If SafeArrayGetDim(LoopStack()) = 0 Then
                  MsgBox "End of loop without start.", vbExclamation
                  Exit For
               Else
                  LoopList(LoopStack(UBound(LoopStack))) = InstructionP
                  LoopList(InstructionP) = LoopStack(UBound(LoopStack))
                  If UBound(LoopStack()) < 1 Then
                     Erase LoopStack()
                  Else
                     ReDim Preserve LoopStack(LBound(LoopStack) To UBound(LoopStack) - 1) As Long
                  End If
               End If
         End Select
      Next InstructionP
      
      If Not SafeArrayGetDim(LoopStack()) = 0 Then MsgBox "Loop without end.", vbExclamation
   ElseIf Not LoopInstructionP = -1 Then
      If Not SafeArrayGetDim(LoopList) = 0 Then NewInstructionP = LoopList(LoopInstructionP)
   End If
   
EndRoutine:
   Loops = NewInstructionP
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   InputLineBreak NewInputLineBreak:=Unescape("/0D")
   LastInput NewInput:=vbNullString

   If Not Trim$(Command$()) = vbNullString Then ToBeExecuted NewToBeExecuted:=Command$()

   InterfaceWindow.Show
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure manages the name of the file to be loaded and executed.
Public Function ToBeExecuted(Optional NewToBeExecuted As String = vbNullString, Optional Remove As Boolean = False) As String
On Error GoTo ErrorTrap
   Static CurrentToBeExecuted As String
   
   If Not NewToBeExecuted = vbNullString Then
      CurrentToBeExecuted = NewToBeExecuted
   ElseIf Remove Then
      CurrentToBeExecuted = vbNullString
   End If
   
EndRoutine:
   ToBeExecuted = CurrentToBeExecuted
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure converts any escape sequences in the specified text to characters.
Public Function Unescape(Text As String, Optional EscapeCharacter As String = "/", Optional ErrorAt As Long = 0) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Hexadecimals As String
Dim Index As Long
Dim NextCharacter As String
Dim Unescaped As String

   ErrorAt = 0
   Index = 1
   Unescaped = vbNullString
   Do Until Index > Len(Text)
      Character = Mid$(Text, Index, 1)
      NextCharacter = Mid$(Text, Index + 1, 1)
   
      If Character = EscapeCharacter Then
         If NextCharacter = EscapeCharacter Then
            Unescaped = Unescaped & Character
            Index = Index + 1
         Else
            Hexadecimals = UCase$(Mid$(Text, Index + 1, 2))
            If Len(Hexadecimals) = 2 Then
               If Left$(Hexadecimals, 1) = "0" Then Hexadecimals = Right$(Hexadecimals, 1)
      
               If UCase$(Hex$(CLng(Val("&H" & Hexadecimals & "&")))) = Hexadecimals Then
                  Unescaped = Unescaped & Chr$(CLng(Val("&H" & Hexadecimals & "&")))
                  Index = Index + 2
               Else
                  ErrorAt = Index
                  Exit Do
               End If
            Else
               ErrorAt = Index
               Exit Do
            End If
         End If
      Else
         Unescaped = Unescaped & Character
      End If
      Index = Index + 1
   Loop
   
EndRoutine:
   Unescape = Unescaped
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


