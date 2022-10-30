Attribute VB_Name = "ProcedureSorterModule"
'This module contains this program's interface and core procedures.
Option Base 0
Option Compare Binary
Option Explicit

'The Microsoft Windows API constants, functions, and structures used by this program.

Private Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustomFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Private Const ERROR_SUCCESS As Long = 0
Private Const MAX_STRING As Long = 65535
Private Const OFN_EXPLORER As Long = &H80000
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_LONGNAMES As Long = &H200000
Private Const OFN_PATHMUSTEXIST As Long = &H800

Private Declare Function CommDlgExtendedError Lib "Comdlg32.dll" () As Long
Private Declare Function GetOpenFileNameA Lib "Comdlg32.dll" (lpofn As OPENFILENAME) As Long
Private Declare Function SafeArrayGetDim Lib "Oleaut32.dll" (ByRef saArray() As Any) As Long


'The constants, enumerations, structures, and variables used by this program.

'This enumeration lists the possible nullable boolean values.
Private Enum NullableBooleanE
   NBNull    'Defines null.
   NBFalse   'Defines false.
   NBTrue    'Defines true.
End Enum

'This structure defines the command line arguments.
Private Type CommandLineArgumentsStr
   CheckForBinaryFiles As Boolean         'Indicates whether to check for binary format files.
   CurrentFile As String                  'Defines the path selected by the user.
   DeleteEmptyProcedures As Boolean       'Indicates whether empty procedures are deleted from a module.
   SortUnderScoresSeparately As Boolean   'Indicates whether procedures with an underscore in their name are sorted separately.
End Type

'This structure defines the current file names.
Private Type CurrentFilesStr
   ModuleFile As String                   'Defines a module's path.
   ProjectFile As String                  'Defines a project's path.
   ProjectGroupFile As String             'Defines a project group's path.
End Type

'This structure defines a module's code.
Private Type ModuleStr
   HeaderCode As String                   'Defines a module's header.
   ProcedureCode() As String              'Defines a module's procedures.
   ProcedureEmpty() As Boolean            'Defines the list that indicates which procedures are empty.
   ProcedureNames() As String             'Defines a module's procedure names.
End Type

'This structure defines the sorting's status.
Private Type SortingStatusStr
   BinaryFilesDetected As Long            'Defines the number of detected binary files.
   EmptyProcedures As Long                'Defines the number of empty procedures found.
   ProcedureCount As Long                 'Defines the total number of procedures found.
   ProceduresSorted As Long               'Defines the number of procedures that were sorted.
   Success As Boolean                     'Indicates whether the sorting process succeeded.
   UnderScoreProcedureCount As Long       'Defines the number of procedures with an underscore in their name found.
End Type

Private CommentStatements() As Variant        'Contains the list of comment statements.
Private ModuleExtensions() As Variant         'Contains the list of supported module file name extensions.
Private ModuleTypes() As Variant              'Contains the list of supported module types.
Private ProcedureEndStatements() As Variant   'Contains the list of supported procedure end statements.
Private ProcedureModifiers() As Variant       'Contains the list of supported procedure modifiers.
Private ProcedureScopes() As Variant          'Contains the list of supported procedure scope statements.
Private ProcedureStatements() As Variant      'Contains the list of supported procedure start statements.
Private ProjectExtensions() As Variant        'Contains the list of supported project file name extensions.
Private ProjectTypes() As Variant             'Contains the list of supported project types.
Private ProjectGroupExtensions() As Variant   'Contains the list of supported project group file name extensions.
Private ProjectGroupTypes() As Variant        'Contains the list of supported project types.
Private PropertyTypes() As Variant            'Contains the list of supported property procedure type statements.

Private Const ARGUMENT_DELIMITER As String = "/"           'Defines the command line argument delimiter.
Private Const MODULE_PATH_DELIMITER As String = ";"        'Defines the delimiter for module names and module paths.
Private Const MODULE_PROPERTIES_DELIMITER = "="            'Defines the delimiter for module properties in project files.
Private Const PARAMETER_LIST As String = "("               'Defines the first character for a procedure's parameter list.
Private Const PROJECT_PROPERTIES_DELIMITER = "="           'Defines the delimiter for project properties in project group files.
Private Const PROPERTY_STATEMENT As String = "Property "   'Defines the property statement.
Private Const UNDERSCORE_MARKER As String = "~"            'Indicates procedure names containing underscores.

'This procedure appends a line and if necessary a line break to the specified text.
Private Sub AppendLine(ByRef Text As String, NewLine As String)
On Error GoTo ErrorTrap
   If Not Text = vbNullString Then Text = Text & vbCrLf
   Text = Text & NewLine
EndRoutine:
   Exit Sub

ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure manages and returns the command line arguments
Private Function CommandLineArguments(Optional NewArguments As String = vbNullString) As CommandLineArgumentsStr
On Error GoTo ErrorTrap
Dim Arguments As String
Dim Position As Long
Static CurrentArguments As CommandLineArgumentsStr

   If Not NewArguments = vbNullString Then
      With CurrentArguments
         .CheckForBinaryFiles = False
         .CurrentFile = vbNullString
         .DeleteEmptyProcedures = False
         .SortUnderScoresSeparately = False
         
         Arguments = NewArguments
         Position = InStr(Arguments, ARGUMENT_DELIMITER)
         
         If Position < 0 Then
            .CurrentFile = Arguments
         Else
            If Position > 1 Then
               .CurrentFile = Trim$(Left$(Arguments, Position - 1))
               Arguments = UCase$(Trim$(Mid$(Arguments, Position - 1))) & " "
            End If
            
            Arguments = UCase$(Trim$(Arguments)) & " "
            
            .CheckForBinaryFiles = (InStr(Arguments, ARGUMENT_DELIMITER & "CFB ") > 0)
            .DeleteEmptyProcedures = (InStr(Arguments, ARGUMENT_DELIMITER & "DEP ") > 0)
            .SortUnderScoresSeparately = (InStr(Arguments, ARGUMENT_DELIMITER & "SUS ") > 0)
         End If
      End With
   End If

EndRoutine:
   CommandLineArguments = CurrentArguments
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure counts the number or procedures to be sorted.
Private Sub CountProceduresToSort(ModuleCode As ModuleStr, SortingStatus As SortingStatusStr)
Dim Index As Long
   
   With ModuleCode
      If Not SafeArrayGetDim(.ProcedureNames) = 0 Then
         For Index = LBound(.ProcedureNames()) To UBound(.ProcedureNames()) - 1
            If LCase$(.ProcedureNames(Index)) > LCase$(.ProcedureNames(Index + 1)) Then
               If Not (.ProcedureNames(Index) = vbNullString Or .ProcedureNames(Index + 1) = vbNullString) Then
                  SortingStatus.ProceduresSorted = SortingStatus.ProceduresSorted + 1
               End If
            End If
         Next Index
      End If
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure displays the results of the sorting processs.
Private Sub DisplayResults(SortingStatus As SortingStatusStr)
On Error GoTo ErrorTrap
Dim Results As String
  
   Results = vbNullString
   With SortingStatus
      If CommandLineArguments().CheckForBinaryFiles Then AppendLine Results, "Binary files detected: " & CStr(.BinaryFilesDetected)
      If CommandLineArguments().DeleteEmptyProcedures Then AppendLine Results, "Empty procedures deleted: " & CStr(.EmptyProcedures)
      If CommandLineArguments().SortUnderScoresSeparately Then AppendLine Results, "Underscore procedures found: " & CStr(.UnderScoreProcedureCount)
      AppendLine Results, "Procedures found: " & CStr(.ProcedureCount)
      AppendLine Results, "Procedures sorted: " & CStr(.ProceduresSorted)
   End With

   MsgBox Results, vbOKOnly Or vbInformation
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure returns the specified path without any file name that might be present.
Private Function GetDirectory(Path As String) As String
On Error GoTo ErrorTrap
Dim Directory As String
Dim Position As Long

   Directory = "."
   Position = InStrRev(Path, "\")
   If Position > 0 Then Directory = Left$(Path, Position - 1)
EndRoutine:
   GetDirectory = Directory
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns the drive in the specified path.
Private Function GetDrive(Path As String) As String
On Error GoTo ErrorTrap
Dim Drive As String

   Drive = Left$(Path, InStr(Path, ":"))

EndRoutine:
   GetDrive = Drive
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns the file name without any directories that might be present in the specified path.
Private Function GetFileName(Path As String) As String
On Error GoTo ErrorTrap
Dim FileName As String
Dim Position As Long

   FileName = vbNullString
   Position = InStrRev(Path, "\")
   If Position > 0 Then FileName = Mid$(Path, Position + 1)
EndRoutine:
   GetFileName = FileName
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns the module code for the specified file.
Private Function GetModuleCode(CurrentFiles As CurrentFilesStr, SortingStatus As SortingStatusStr) As ModuleStr
On Error GoTo ErrorTrap
Dim Code As String
Dim CommentBlock As String
Dim CurrentProcedureName As String
Dim FileHandle As Long
Dim ModuleCode As ModuleStr
Dim ProcedureBody As String
   
   With ModuleCode
      ReDim .ProcedureCode(0 To 0) As String
      ReDim .ProcedureEmpty(0 To 0) As Boolean
      ReDim .ProcedureNames(0 To 0) As String

      CommentBlock = vbNullString
      CurrentProcedureName = vbNullString
      ProcedureBody = vbNullString

      .HeaderCode = vbNullString
      
      FileHandle = FreeFile()
      Open CurrentFiles.ModuleFile For Input Lock Read Write As FileHandle
         Do Until EOF(FileHandle)
            Line Input #FileHandle, Code
            
            If IsProcedureStart(Code, CurrentProcedureName, SortingStatus) Then
               .ProcedureCode(UBound(.ProcedureCode())) = .ProcedureCode(UBound(.ProcedureCode())) & CommentBlock & Code & vbCrLf
               .ProcedureEmpty(UBound(.ProcedureEmpty())) = False
               .ProcedureNames(UBound(.ProcedureNames())) = CurrentProcedureName
               CommentBlock = vbNullString
               Exit Do
            ElseIf IsLineComment(Code) Then
               CommentBlock = CommentBlock & Code & vbCrLf
            Else
               If Not .HeaderCode = vbNullString Then .HeaderCode = .HeaderCode & vbCrLf
               .HeaderCode = .HeaderCode & CommentBlock & Code
               CommentBlock = vbNullString
            End If
         Loop
         
         Do Until EOF(FileHandle)
            Line Input #FileHandle, Code
            If Not (Trim$(Code) = vbNullString And CurrentProcedureName = vbNullString) Then Code = Code & vbCrLf
            If IsProcedureStart(Code, CurrentProcedureName, SortingStatus) Then
               .ProcedureCode(UBound(.ProcedureCode())) = .ProcedureCode(UBound(.ProcedureCode())) & CommentBlock & Code
               .ProcedureEmpty(UBound(.ProcedureEmpty())) = False
               .ProcedureNames(UBound(.ProcedureNames())) = CurrentProcedureName
               CommentBlock = vbNullString
               ProcedureBody = vbNullString
            ElseIf IsProcedureEnd(Code, CurrentProcedureName) Then
               .ProcedureCode(UBound(.ProcedureCode())) = .ProcedureCode(UBound(.ProcedureCode())) & CommentBlock & Code
               
               .ProcedureEmpty(UBound(.ProcedureEmpty())) = (RemoveLineBreaks(ProcedureBody) = vbNullString)
               If .ProcedureEmpty(UBound(.ProcedureEmpty())) Then SortingStatus.EmptyProcedures = SortingStatus.EmptyProcedures + 1
               
               ReDim Preserve .ProcedureCode(LBound(.ProcedureCode()) To UBound(.ProcedureCode()) + 1) As String
               ReDim Preserve .ProcedureEmpty(LBound(.ProcedureEmpty()) To UBound(.ProcedureEmpty()) + 1) As Boolean
               ReDim Preserve .ProcedureNames(LBound(.ProcedureNames()) To UBound(.ProcedureNames()) + 1) As String
               ProcedureBody = vbNullString
               
               SortingStatus.ProcedureCount = SortingStatus.ProcedureCount + 1
            ElseIf IsLineComment(Code) Then
               CommentBlock = CommentBlock & Code
               ProcedureBody = ProcedureBody & Trim$(CommentBlock) & Trim$(Code)
            Else
               ProcedureBody = ProcedureBody & Trim$(CommentBlock) & Trim$(Code)
               .ProcedureCode(UBound(.ProcedureCode())) = .ProcedureCode(UBound(.ProcedureCode())) & CommentBlock & Code
               CommentBlock = vbNullString
            End If
         Loop
      Close FileHandle
   End With
EndRoutine:
   GetModuleCode = ModuleCode
   Exit Function
   
ErrorTrap:
   Select Case HandleError(CurrentPath:=CurrentFiles.ModuleFile)
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure handles any errors that occur and returns the action selected by the user.
Private Function HandleError(Optional CurrentPath As String = vbNullString) As Long
Dim Choice As Long
Dim ErrorCode As Long
Dim ErrorText As String
Dim Message As String

   ErrorCode = Err.Number
   ErrorText = Trim$(Err.Description)
   
   On Error GoTo ErrorTrap
   
   If Not Right$(ErrorText, 1) = "." Then ErrorText = ErrorText & "."
   Message = "Error Code: " & ErrorCode & vbCr
   Message = Message & ErrorText
   If Not CurrentPath = vbNullString Then Message = Message & vbCr & "Current path: """ & CurrentPath & """"
   
   Choice = MsgBox(Message, vbAbortRetryIgnore Or vbExclamation)
   
   If Choice = vbAbort Then
      Reset
      Resume EndProgram
   End If

   HandleError = Choice
   Exit Function
   
EndProgram:
   End

ErrorTrap:
   Resume EndProgram
End Function

'This procedure returns whether the specified text contains a property type statement.
Private Function HasPropertyType(Text As String) As Boolean
On Error GoTo ErrorTrap
Dim HasType As Boolean
Dim Index As Long

   HasType = False
   For Index = LBound(PropertyTypes()) To UBound(PropertyTypes())
      If LCase$(Left$(Text, Len(PropertyTypes(Index)))) = LCase$(PropertyTypes(Index)) Then
         HasType = True
         Exit For
      End If
   Next Index
EndRoutine:
   HasPropertyType = HasType
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure initializes this program.
Private Sub Initialize(SortingStatus As SortingStatusStr)
On Error GoTo ErrorTrap
   CommentStatements = Array("'", "Rem ")
   ModuleExtensions = Array(".bas", ".cls", ".ctl", ".dob", ".frm", ".pag")
   ModuleTypes = Array("Code Module", "Class Module", "User Control", "Designer Object", "Form", "Property Page")
   ProcedureEndStatements = Array("End Function", "End Property", "End Sub")
   ProcedureModifiers = Array(vbNullString, "Static ")
   ProcedureStatements = Array("Function ", "Property ", "Sub ")
   ProjectExtensions = Array(".mak", ".vbp")
   ProjectTypes = Array("Project", "Visual Basic Project")
   ProjectGroupExtensions = Array(".vbg")
   ProjectGroupTypes = Array("Visual Basic Project Group")
   PropertyTypes = Array("Get ", "Let ", "Set ")
   ProcedureScopes = Array(vbNullString, "Friend ", "Private ", "Public ")
   
   With SortingStatus
      .BinaryFilesDetected = 0
      .EmptyProcedures = 0
      .ProcedureCount = 0
      .ProceduresSorted = 0
      .Success = True
      .UnderScoreProcedureCount = 0
   End With
      
   CommandLineArguments NewArguments:=Command$()
EndRoutine:
   Exit Sub

ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure returns whether the specified file is considered to be stored in binary format.
Private Function IsBinaryFormat(FileName As String, SortingStatus As SortingStatusStr) As Boolean
On Error GoTo ErrorTrap
Dim Data As String
Dim FileHandle As Long
Dim IsBinary As Boolean

   IsBinary = False
   FileHandle = FreeFile()
   Open FileName For Input Lock Read Write As FileHandle: Close FileHandle
   Open FileName For Binary Lock Read Write As FileHandle
      If Loc(FileHandle) <= LOF(FileHandle) Then
         IsBinary = (Asc(Input$(1, FileHandle)) >= &HFC&)
      End If
   Close FileHandle

   If IsBinary Then
      SortingStatus.BinaryFilesDetected = SortingStatus.BinaryFilesDetected + 1
   End If

EndRoutine:
   IsBinaryFormat = IsBinary
   Exit Function
 
ErrorTrap:
   Select Case HandleError(CurrentPath:=FileName)
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns whether the line of code being processed is a comment.
Private Function IsLineComment(Code As String) As Boolean
On Error GoTo ErrorTrap
Dim Index As Long
Dim IsComment As Boolean

   IsComment = False
   For Index = LBound(CommentStatements()) To UBound(CommentStatements())
      If LCase$(Left$(Trim$(Code), Len(CommentStatements(Index)))) = LCase$(CommentStatements(Index)) Then
         IsComment = True
         Exit For
      End If
   Next Index
EndRoutine:
   IsLineComment = IsComment
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function


'This procedure returns whether the specified file is a module file.
Private Function IsModuleFile(FileName As String) As Boolean
On Error GoTo ErrorTrap
Dim Index As Long
Dim IsModule As Boolean

   IsModule = False
   For Index = LBound(ModuleExtensions()) To UBound(ModuleExtensions())
      If LCase$(Right$(Trim$(FileName), Len(ModuleExtensions(Index)))) = LCase$(ModuleExtensions(Index)) Then
         IsModule = True
         Exit For
      End If
   Next Index
EndRoutine:
   IsModuleFile = IsModule
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns whether the end of a procedure has been reached.
Private Function IsProcedureEnd(Code As String, ByRef CurrentProcedureName As String) As Boolean
On Error GoTo ErrorTrap
Dim Index As Long
Dim IsEnd As Boolean

   IsEnd = False
   For Index = LBound(ProcedureEndStatements()) To UBound(ProcedureEndStatements())
      If LCase$(Left$(Trim$(Code), Len(ProcedureEndStatements(Index)))) = LCase$(ProcedureEndStatements(Index)) Then
         IsEnd = True
         CurrentProcedureName = vbNullString
         Exit For
      End If
   Next Index
EndRoutine:
   IsProcedureEnd = IsEnd
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns whether the start of a procedure has been reached.
Private Function IsProcedureStart(Code As String, ByRef CurrentProcedureName As String, SortingStatus As SortingStatusStr) As Boolean
On Error GoTo ErrorTrap
Dim ActualNamePosition As Long
Dim ArgumentsPosition As Long
Dim Index1 As Long
Dim Index2 As Long
Dim Index3 As Long
Dim IsStart As NullableBooleanE
Dim NamePosition As Long
Dim ProcedureName As String

   IsStart = NBNull
   For Index1 = LBound(ProcedureScopes()) To UBound(ProcedureScopes())
      For Index2 = LBound(ProcedureModifiers()) To UBound(ProcedureModifiers())
         For Index3 = LBound(ProcedureStatements()) To UBound(ProcedureStatements())
            NamePosition = Len(ProcedureScopes(Index1) & ProcedureModifiers(Index2) & ProcedureStatements(Index3))
            If LCase$(Left$(Trim$(Code), NamePosition)) = LCase$(ProcedureScopes(Index1) & ProcedureModifiers(Index2) & ProcedureStatements(Index3)) Then
               ArgumentsPosition = InStr(NamePosition, Code, PARAMETER_LIST)
               If ArgumentsPosition = 0 Then ArgumentsPosition = Len(Code)
               ProcedureName = Trim$(Mid$(Code, NamePosition, ArgumentsPosition - NamePosition))
               If InStr(CurrentProcedureName, " ") > 0 Then
                  ActualNamePosition = InStr(ProcedureName, " ") + 1
                  ProcedureName = Mid$(ProcedureName, ActualNamePosition) & " " & Left$(ProcedureName, ActualNamePosition - 1)
               End If
               If LCase$(ProcedureStatements(Index3)) = LCase$(PROPERTY_STATEMENT) Then
                  If HasPropertyType(ProcedureName) Then
                     ProcedureName = SwapPropertyTypeName(ProcedureName)
                  Else
                     IsStart = NBFalse
                     Exit For
                  End If
               End If
               If CommandLineArguments().SortUnderScoresSeparately And InStr(ProcedureName, "_") > 0 Then
                  ProcedureName = UNDERSCORE_MARKER & ProcedureName
                  SortingStatus.UnderScoreProcedureCount = SortingStatus.UnderScoreProcedureCount + 1
               End If
               CurrentProcedureName = ProcedureName
               IsStart = NBTrue
               Exit For
            End If
         Next Index3
         If Not IsStart = NBNull Then Exit For
      Next Index2
      If Not IsStart = NBNull Then Exit For
   Next Index1
   
EndRoutine:
   IsProcedureStart = IsStart
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns whether the specified file is a project file.
Private Function IsProjectFile(FileName As String) As Boolean
On Error GoTo ErrorTrap
Dim Index As Long
Dim IsProject As Boolean

   IsProject = False
   For Index = LBound(ProjectExtensions()) To UBound(ProjectExtensions())
      If LCase$(Right$(Trim$(FileName), Len(ProjectExtensions(Index)))) = LCase$(ProjectExtensions(Index)) Then
         IsProject = True
         Exit For
      End If
   Next Index
EndRoutine:
   IsProjectFile = IsProject
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure returns whether the specified file is a project group file.
Private Function IsProjectGroupFile(FileName As String) As Boolean
On Error GoTo ErrorTrap
Dim Index As Long
Dim IsProjectGroup As Boolean

   IsProjectGroup = False
   For Index = LBound(ProjectGroupExtensions()) To UBound(ProjectGroupExtensions())
      If Right$(Trim$(FileName), Len(ProjectGroupExtensions(Index))) = ProjectGroupExtensions(Index) Then
         IsProjectGroup = True
         Exit For
      End If
   Next Index
EndRoutine:
   IsProjectGroupFile = IsProjectGroup
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
Dim Arguments As String
Dim CurrentFiles As CurrentFilesStr
Dim SortingStatus As SortingStatusStr

   If App.PrevInstance Then
      MsgBox App.Title & " is already running.", vbOKOnly Or vbExclamation
   Else
      Initialize SortingStatus
      
      With CommandLineArguments()
         If Left$(.CurrentFile, 1) = """" Then .CurrentFile = Mid$(.CurrentFile, 2)
         If Right$(.CurrentFile, 1) = """" Then .CurrentFile = Left$(.CurrentFile, Len(.CurrentFile) - 1)
          
         If .CurrentFile = vbNullString Then
            .CurrentFile = RequestFilePath()
            If .CurrentFile = vbNullString Then
               MsgBox "Must select a file to process.", vbOKOnly Or vbExclamation
            Else
               Arguments = Trim$(InputBox$("Command line arguments:"))
            End If
         End If
      
         If Not .CurrentFile = vbNullString Then
            If Not Arguments = vbNullString Then
               CommandLineArguments NewArguments:=Arguments
            End If
            
            If IsModuleFile(.CurrentFile) Then
               CurrentFiles.ModuleFile = .CurrentFile
               ProcessModule CurrentFiles, SortingStatus
            ElseIf IsProjectFile(.CurrentFile) Then
               CurrentFiles.ProjectFile = .CurrentFile
               ProcessProject CurrentFiles, SortingStatus
            ElseIf IsProjectGroupFile(.CurrentFile) Then
               CurrentFiles.ProjectGroupFile = .CurrentFile
               ProcessProjectGroup CurrentFiles, SortingStatus
            Else
               MsgBox "This type of file is not recognized: """ & .CurrentFile & """.", vbOKOnly Or vbExclamation
               SortingStatus.Success = False
            End If
            
            If SortingStatus.Success Then
               DisplayResults SortingStatus
            End If
         End If
      End With
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure gives the command to process the specified module.
Private Sub ProcessModule(CurrentFiles As CurrentFilesStr, SortingStatus As SortingStatusStr)
On Error GoTo ErrorTrap
Dim ModuleCode As ModuleStr

   If CommandLineArguments().CheckForBinaryFiles Then
      If Not IsBinaryFormat(CurrentFiles.ModuleFile, SortingStatus) Then
         ModuleCode = GetModuleCode(CurrentFiles, SortingStatus)
      End If
   Else
      ModuleCode = GetModuleCode(CurrentFiles, SortingStatus)
   End If
   
   CountProceduresToSort ModuleCode, SortingStatus
   SortProcedures ModuleCode
   ReplaceModuleCode ModuleCode, CurrentFiles, SortingStatus
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure reads the module file names from a project file and gives the command to process each module.
Private Sub ProcessProject(CurrentFiles As CurrentFilesStr, SortingStatus As SortingStatusStr)
On Error GoTo ErrorTrap
Dim FileHandle As Long
Dim ProjectData As String

   ChDrive GetDrive(CurrentFiles.ProjectFile)
   ChDir GetDirectory(CurrentFiles.ProjectFile)
   
   FileHandle = FreeFile()
   Open GetFileName(CurrentFiles.ProjectFile) For Input Lock Read Write As FileHandle
      Do Until EOF(FileHandle)
         Line Input #FileHandle, ProjectData
         ProjectData = Trim$(ProjectData)
         If InStr(ProjectData, MODULE_PROPERTIES_DELIMITER) > 0 Then
            If LCase$(Left$(ProjectData, InStr(ProjectData, MODULE_PROPERTIES_DELIMITER) - 1)) = "class" Or LCase$(Left$(ProjectData, InStr(ProjectData, MODULE_PROPERTIES_DELIMITER) - 1)) = "module" Then
               ProjectData = Trim$(Mid$(ProjectData, InStr(ProjectData, MODULE_PROPERTIES_DELIMITER) + 1))
               ProjectData = Trim$(Mid$(ProjectData, InStr(ProjectData, MODULE_PATH_DELIMITER) + 1))
            Else
               ProjectData = Trim$(Mid$(ProjectData, InStr(ProjectData, MODULE_PROPERTIES_DELIMITER) + 1))
            End If
         End If
         If IsModuleFile(ProjectData) Then
            CurrentFiles.ModuleFile = ProjectData
            ProcessModule CurrentFiles, SortingStatus
         End If
      Loop
   Close FileHandle
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError(CurrentPath:=CurrentFiles.ProjectFile)
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure reads the project file names from a project group file and gives the command to process each project.
Private Sub ProcessProjectGroup(CurrentFiles As CurrentFilesStr, SortingStatus As SortingStatusStr)
On Error GoTo ErrorTrap
Dim FileHandle As Long
Dim ProjectGroupData As String
Dim ProjectGroupFullPath As String

   ChDrive GetDrive(CurrentFiles.ProjectGroupFile)
   ChDir GetDirectory(CurrentFiles.ProjectGroupFile)
   
   ProjectGroupFullPath = CurDir$()

   FileHandle = FreeFile()
   Open GetFileName(CurrentFiles.ProjectGroupFile) For Input Lock Read Write As FileHandle
      Do Until EOF(FileHandle)
         Line Input #FileHandle, ProjectGroupData
         ProjectGroupData = Trim$(ProjectGroupData)
         If InStr(ProjectGroupData, PROJECT_PROPERTIES_DELIMITER) > 0 Then
            ProjectGroupData = Trim$(Mid$(ProjectGroupData, InStr(ProjectGroupData, PROJECT_PROPERTIES_DELIMITER) + 1))
         End If
         If IsProjectFile(ProjectGroupData) Then
            CurrentFiles.ProjectFile = ProjectGroupData
            ProcessProject CurrentFiles, SortingStatus
           
            ChDrive GetDrive(ProjectGroupFullPath)
            ChDir ProjectGroupFullPath
         End If
      Loop
   Close FileHandle
EndRoutine:
   Exit Sub
 
ErrorTrap:
   Select Case HandleError(CurrentPath:=CurrentFiles.ProjectGroupFile)
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This function filters the line breaks from the specified text and returns the result.
Private Function RemoveLineBreaks(Text As String) As String
On Error GoTo ErrorTrap
Dim Character As String
Dim Index As Long
Dim NewText As String

   NewText = vbNullString
   For Index = 1 To Len(Text)
      Character = Mid$(Text, Index, 1)
      If Not (Character = vbCr Or Character = vbLf) Then NewText = NewText & Character
   Next Index
   
EndRoutine:
   RemoveLineBreaks = NewText
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure replaces the original module code with the sorted code.
Private Sub ReplaceModuleCode(ModuleCode As ModuleStr, CurrentFiles As CurrentFilesStr, SortingStatus As SortingStatusStr)
On Error GoTo ErrorTrap
Dim FileHandle As Long
Dim Index As Long

   If Not (IsBinaryFormat(CurrentFiles.ModuleFile, SortingStatus) And CommandLineArguments().CheckForBinaryFiles) Then
      FileHandle = FreeFile()
      Open CurrentFiles.ModuleFile For Output Lock Read Write As FileHandle
         With ModuleCode
            Print #FileHandle, .HeaderCode;
            For Index = LBound(.ProcedureNames()) To UBound(.ProcedureNames())
               If Not (.ProcedureEmpty(Index) And CommandLineArguments().DeleteEmptyProcedures) Then
                  Print #FileHandle, .ProcedureCode(Index);
                  Print #FileHandle,
               End If
            Next Index
         End With
      Close FileHandle
   End If
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError(CurrentPath:=CurrentFiles.ModuleFile)
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure requests the user to select a file and returns the result.
Private Function RequestFilePath() As String
On Error GoTo ErrorTrap
Dim Dialog As OPENFILENAME
Dim ErrorCode As Long
Dim Index As Long
Dim Message As String
Dim PathV As String
Dim ReturnValue As Long

   With Dialog
      .flags = OFN_EXPLORER
      .flags = .flags Or OFN_HIDEREADONLY
      .flags = .flags Or OFN_FILEMUSTEXIST
      .flags = .flags Or OFN_LONGNAMES
      .flags = .flags Or OFN_PATHMUSTEXIST
       
      .hInstance = 0
      .hwndOwner = 0
      .lCustData = 0
      .lpfnHook = 0
      .lpstrCustomFilter = vbNullString
      .lpstrDefExt = vbNullString
      .lpstrFile = String$(MAX_STRING, vbNullChar) & vbNullChar
      .lpstrFileTitle = String$(MAX_STRING, vbNullChar) & vbNullChar
       
      .lpstrFilter = "All Supported File Types" & vbNullChar
      For Index = LBound(ModuleTypes()) To UBound(ModuleTypes())
         .lpstrFilter = .lpstrFilter & "*" & ModuleExtensions(Index) & ";"
      Next Index
      For Index = LBound(ProjectTypes()) To UBound(ProjectTypes())
         .lpstrFilter = .lpstrFilter & "*" & ProjectExtensions(Index) & ";"
      Next Index
      For Index = LBound(ProjectGroupTypes()) To UBound(ProjectGroupTypes())
         .lpstrFilter = .lpstrFilter & "*" & ProjectGroupExtensions(Index) & ";"
      Next Index
       
      .lpstrFilter = .lpstrFilter & vbNullChar
      For Index = LBound(ModuleTypes()) To UBound(ModuleTypes())
         .lpstrFilter = .lpstrFilter & ModuleTypes(Index) & " (*" & ModuleExtensions(Index) & ")" & vbNullChar
         .lpstrFilter = .lpstrFilter & "*" & ModuleExtensions(Index) & vbNullChar
      Next Index
      For Index = LBound(ProjectTypes()) To UBound(ProjectTypes())
         .lpstrFilter = .lpstrFilter & ProjectTypes(Index) & " (*" & ProjectExtensions(Index) & ")" & vbNullChar
         .lpstrFilter = .lpstrFilter & "*" & ProjectExtensions(Index) & vbNullChar
      Next Index
      For Index = LBound(ProjectGroupTypes()) To UBound(ProjectGroupTypes())
         .lpstrFilter = .lpstrFilter & ProjectGroupTypes(Index) & " (*" & ProjectGroupExtensions(Index) & ")" & vbNullChar
         .lpstrFilter = .lpstrFilter & "*" & ProjectGroupExtensions(Index) & vbNullChar
      Next Index
      .lpstrFilter = .lpstrFilter & vbNullChar
      
      .lpstrInitialDir = App.Path & vbNullChar
      .lpstrTitle = App.Title & " - Select a module or project to sort." & vbNullChar
      .lpTemplateName = vbNullString
      .lStructSize = Len(Dialog)
      .nFileExtension = 0
      .nFileOffset = 0
      .nFilterIndex = 1
      .nMaxCustomFilter = 0
      .nMaxFile = Len(.lpstrFile)
      .nMaxFileTitle = Len(.lpstrFileTitle)
   End With
   
   PathV = vbNullString
   
   ReturnValue = GetOpenFileNameA(Dialog)
   If ReturnValue = 0 Then
      ErrorCode = CommDlgExtendedError()
      If Not ErrorCode = ERROR_SUCCESS Then
         Message = "Common dialog error:" & vbCrLf
         Message = Message & "Error code: " & ErrorCode & vbCrLf
         Message = Message & "Return value: " & ReturnValue
         MsgBox Message, vbExclamation
      End If
   Else
      PathV = Trim$(Left$(Dialog.lpstrFile, InStr(Dialog.lpstrFile, vbNullChar) - 1))
   End If
   
EndRoutine:
   RequestFilePath = PathV
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

'This procedure sorts the procedures.
Private Sub SortProcedures(ModuleCode As ModuleStr)
On Error GoTo ErrorTrap
Dim Index As Long
Dim OtherIndex As Long
   
   With ModuleCode
      If Not SafeArrayGetDim(.ProcedureNames) = 0 Then
         For Index = LBound(.ProcedureNames()) To UBound(.ProcedureNames())
            For OtherIndex = LBound(.ProcedureNames()) To UBound(.ProcedureNames())
               If LCase$(.ProcedureNames(Index)) < LCase$(.ProcedureNames(OtherIndex)) Then
                  Swap .ProcedureCode(Index), .ProcedureCode(OtherIndex)
                  Swap .ProcedureEmpty(Index), .ProcedureEmpty(OtherIndex)
                  Swap .ProcedureNames(Index), .ProcedureNames(OtherIndex)
               End If
            Next OtherIndex
         Next Index
      End If
   End With
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure swaps the two specified variables with each other.
Private Sub Swap(Variable1 As Variant, Variable2 As Variant)
On Error GoTo ErrorTrap
Dim Variable3 As Variant

   Variable3 = Variable1
   Variable1 = Variable2
   Variable2 = Variable3
EndRoutine:
   Exit Sub
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Sub

'This procedure swaps a property procedure's type and name to ensure proper sorting and returns the result.
Private Function SwapPropertyTypeName(PropertyTypeName As String) As String
On Error GoTo ErrorTrap
Dim Position As Long
Dim PropertyName As String
Dim PropertyType As String
Dim Swapped As String

   Position = InStr(PropertyTypeName, " ")
   PropertyType = Left$(PropertyTypeName, Position - 1)
   PropertyName = Mid$(PropertyTypeName, Position + 1)
   Swapped = PropertyName & " " & PropertyType
   
EndRoutine:
   SwapPropertyTypeName = Swapped
   Exit Function
   
ErrorTrap:
   Select Case HandleError()
      Case vbIgnore
         Resume EndRoutine
      Case vbRetry
         Resume
   End Select
End Function

