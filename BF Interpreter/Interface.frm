VERSION 5.00
Begin VB.Form InterfaceWindow 
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   636
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ScaleHeight     =   13.5
   ScaleMode       =   4  'Character
   ScaleWidth      =   39
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer DropWatcher 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton EnterButton 
      Caption         =   "&Enter"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox ProgramInputBox 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox ProgramOutputBox 
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   7.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   120
      Width           =   4455
   End
   Begin VB.Menu ProgramMainMenu 
      Caption         =   "&Program"
      Begin VB.Menu InformationMenu 
         Caption         =   "&Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu ProgramMainMenuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu QuitMenu 
         Caption         =   "&Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu CodeMainMenu 
      Caption         =   "&Code"
      Begin VB.Menu AbortExecutionMenu 
         Caption         =   "&Abort Execution"
         Shortcut        =   ^A
      End
      Begin VB.Menu LoadCodeMenu 
         Caption         =   "&Load Code"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu OtherMainMenu 
      Caption         =   "&Other"
      Begin VB.Menu ClearOutputMenu 
         Caption         =   "&Clear Output"
         Shortcut        =   {F1}
      End
      Begin VB.Menu FlushInputMenu 
         Caption         =   "&Flush Input"
         Shortcut        =   {F2}
      End
      Begin VB.Menu InputLineBreakMenu 
         Caption         =   "Input &Linebreak"
         Shortcut        =   {F3}
      End
      Begin VB.Menu RepeatInputMenu 
         Caption         =   "&Repeat Input"
         Shortcut        =   {F4}
      End
   End
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains the main interface window.
Option Explicit




'This procedure aborts the execution of the current code.
Private Sub AbortExecutionMenu_Click()
On Error GoTo ErrorTrap
   If Execute() Then Execute , , Abort:=True
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure clears the output box.
Private Sub ClearOutputMenu_Click()
On Error GoTo ErrorTrap
   ProgramOutputBox.Text = vbNullString
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure checks whether there are any files waiting to be executed.
Private Sub DropWatcher_Timer()
On Error GoTo ErrorTrap
   If Not ToBeExecuted() = vbNullString Then
      Execute Code(ToBeExecuted()), NewOutputBox:=ProgramOutputBox
      ToBeExecuted , Remove:=True
   End If

   DropWatcher.Enabled = False
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to send the users input to the program being executed.
Private Sub EnterButton_Click()
On Error GoTo ErrorTrap
Dim ErrorAt As Long

   If Not ProgramInputBox.Text = vbNullString Then
      LastInput NewInput:=Unescape(ProgramInputBox.Text, , ErrorAt)
      If Not EscapeSequenceError(ErrorAt) Then
         If Not Code() = vbNullString Then
            ProgramInputBox.Text = vbNullString
            InputStream NewInput:=LastInput() & InputLineBreak()
         Else
            MsgBox "Load code first.", vbExclamation
         End If
      End If
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to flush the input stream.
Private Sub FlushInputMenu_Click()
On Error GoTo ErrorTrap
   InputStream , Flush:=True
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap
   With App
      Me.Caption = .Title & " - by: " & .CompanyName & ", v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision)
   End With
   
   Me.Width = Screen.Width / 1.5
   Me.Height = Screen.Height / 1.5
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure closes this program when this window is closed.
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo ErrorTrap
   End
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure requests the user to specify a line break to be appended to the input.
Private Sub InputLineBreakMenu_Click()
On Error GoTo ErrorTrap
Dim ErrorAt As Long
Dim NewInputLineBreak As String
   
   NewInputLineBreak = InputBox("Input line break:", , Escape(InputLineBreak()))
   If Not StrPtr(NewInputLineBreak) = 0 Then
      NewInputLineBreak = Unescape(NewInputLineBreak, , ErrorAt)
      If Not EscapeSequenceError(ErrorAt) Then
          InputLineBreak NewInputLineBreak:=NewInputLineBreak, Remove:=(NewInputLineBreak = vbNullString)
      End If
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure opens the code file dropped into the outputbox.
Private Sub ProgramOutputBox_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorTrap

   If Data.Files.Count > 0 And Not Execute(, , , WarnUser:=True) Then
      ToBeExecuted NewToBeExecuted:=Data.Files.Item(1)
      DropWatcher.Enabled = True
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure adjusts the size and position of the interface object's to the window's new size.
Private Sub Form_Resize()
On Error Resume Next
   ProgramOutputBox.Height = Me.ScaleHeight - ProgramInputBox.Height
   ProgramOutputBox.Left = 0
   ProgramOutputBox.Top = 0
   ProgramOutputBox.Width = Me.ScaleWidth
   
   ProgramInputBox.Left = 0
   ProgramInputBox.Width = Me.ScaleWidth - EnterButton.Width - 2
   ProgramInputBox.Top = Me.ScaleHeight - ProgramInputBox.Height
   
   EnterButton.Left = ProgramInputBox.Width + 1
   EnterButton.Top = (Me.ScaleHeight - (ProgramInputBox.Height / 2)) - (EnterButton.Height / 2)
End Sub

'This procedure displays information about this program.
Private Sub InformationMenu_Click()
On Error GoTo ErrorTrap
   MsgBox App.Comments, vbInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure gives the command to load and execute the specified code.
Private Sub LoadCodeMenu_Click()
On Error GoTo ErrorTrap
Dim Path As String

   If Not Execute(, , , WarnUser:=True) Then
      Path = InputBox$("Code file:")
      If Not Path = vbNullString Then Execute Code(Path), ProgramOutputBox
   End If
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub


'This procedure gives the command to close this program.
Private Sub QuitMenu_Click()
On Error GoTo ErrorTrap
   Unload Me
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure repeats the previous user console input.
Private Sub RepeatInputMenu_Click()
On Error GoTo ErrorTrap
   ProgramInputBox.Text = LastInput()
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

