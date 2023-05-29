VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''
' VBA-ErrorHandling: Logger
' (c) RadiusCore Ltd - https://radiuscore.co.nz/
'
' Logging for VBA.
'
' @module Logger
' @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' @depend ErrorHandler, FileSystemObject, TextStream
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder ErrorHandling
'@ignoremodule ProcedureNotUsed
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' --------------------------------------------- '
' Windows API Headers
' --------------------------------------------- '

#If Not Mac Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
#End If

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TLogger
    LogFile As TextStream
    LogFilePath As String
    LogTitle As String
    Threshold As VbLogLevel
    Immediate As Boolean
    Console As Boolean
    ConsoleHandle As Long
    IsInitialised As Boolean
End Type

Private This As TLogger

' --------------------------------------------- '
' Public Properties
' --------------------------------------------- '

''
' Turn logging off (0) or only log messages that are >= threshold
'
' LogThreshold = 4 -> Level >= 4 -> Warn and Error
' LogThreshold = 0 -> Off
'
' @property LogThreshold
' @type {vbLogLevel}
' @default 0
''
Public Property Get LogThreshold() As VbLogLevel
    LogThreshold = This.Threshold
End Property
Public Property Let LogThreshold(ByVal Value As VbLogLevel)
    This.Threshold = Value
End Property

''
' Whether logging is enabled.
' Read-only. Use `LogThreshold` to enable/disable logging.
'
' @property LogEnabled
' @type {Boolean}
' @default False
''
Public Property Get LogEnabled() As Boolean
    LogEnabled = Not (LogThreshold = VbLogLevel.vbLogOff)
End Property

''
' Whether to also log information to VBE's immediate window.
'
' @property LogToImmediate
' @type {Boolean}
' @default False
''
Public Property Get LogToImmediate() As Boolean
    LogToImmediate = This.Immediate
End Property
Public Property Let LogToImmediate(ByVal Value As Boolean)
    This.Immediate = Value
End Property

''
' Whether to also log information to a console window (Windows Only).
'
' @property LogToConsole
' @type {Boolean}
' @default False
''
#If Not Mac Then
Public Property Get LogToConsole() As Boolean
    LogToConsole = This.Console
End Property
Public Property Let LogToConsole(ByVal Value As Boolean)
    This.Console = Value
End Property
#End If

' ============================================= '
' Public Methods
' ============================================= '

''
' Log message.
'
' @method Log
' @param {VbLogLevel} Level
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub Log(ByVal Level As VbLogLevel, ByVal Message As String, Optional ByVal From As String = vbNullString)
    If Not This.IsInitialised Then Err.Raise 1004, "Logger", "Application-defined or object-defined error (class not intialised)."  ' Exit if class is not initialised.
    If Me.LogThreshold = VbLogLevel.vbLogOff Or Level < Me.LogThreshold Then Exit Sub   ' Exit if log threshold is not high enough.
    If VBA.FileLen(This.LogFilePath) >= 1024000 Then log_UpdateLogFiles                 ' Create new log file if size limit is met.
#If Not Mac Then
    If This.Console And This.ConsoleHandle = -1 Then log_OpenConsole                    ' Show log in console, if enabled.
#End If
    
    Dim lg_Text As String
    lg_Text = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2) & "|" & _
              log_LogLevelToText(Level) & "|" & _
              VBA.IIf(From = vbNullString, vbNullString, From & "|") & _
              Message                                                           ' Compile string to log.
    
    If This.Immediate Then Debug.Print lg_Text                                  ' Log to immediate window.
    This.LogFile.WriteLine lg_Text                                              ' Log to file.
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' Update log files when:
'  1) Class is initialised, performing the following actions:
'     - add `new instance` line breaks.
'  2) When the existing log file reaches 1KB in size, performing the following actions:
'     - increment existing logs with numeric identifiers;
'     - delete the oldest log; and
'     - create a new (blank) log file.
'
' @method log_UpdateLogFiles
''
Private Sub log_UpdateLogFiles()
    Dim log_FSO As FileSystemObject
    Dim log_Long As Long
    Set log_FSO = New FileSystemObject
    
    ' Create folder if necessary.
    If Not log_FSO.FolderExists(log_FSO.GetParentFolderName(This.LogFilePath)) Then log_FSO.CreateFolder log_FSO.GetParentFolderName(This.LogFilePath)
    
    If This.LogFile Is Nothing Then
        ' Add `new instance` line breaks to existing log file.
        Set This.LogFile = log_FSO.OpenTextFile(This.LogFilePath, ForAppending, True, TristateFalse)
        This.LogFile.WriteBlankLines 1
        This.LogFile.WriteLine "----------" & This.LogTitle & "----------"
    Else
        ' Update log file numbering & delete oldest log.
        With log_FSO
            If .FileExists(This.LogFilePath) Then
                For log_Long = 4 To 0 Step -1
                    If .FileExists(.BuildPath(.GetParentFolderName(This.LogFilePath), .GetBaseName(This.LogFilePath)) & "." & log_Long & "." & .GetExtensionName(This.LogFilePath)) Then
                        If log_Long = 4 Then
                            .DeleteFile .BuildPath(.GetParentFolderName(This.LogFilePath), .GetBaseName(This.LogFilePath)) & "." & log_Long & "." & .GetExtensionName(This.LogFilePath)
                        Else
                            .MoveFile .BuildPath(.GetParentFolderName(This.LogFilePath), .GetBaseName(This.LogFilePath)) & "." & log_Long & "." & .GetExtensionName(This.LogFilePath), _
                                      .BuildPath(.GetParentFolderName(This.LogFilePath), .GetBaseName(This.LogFilePath)) & "." & log_Long + 1 & "." & .GetExtensionName(This.LogFilePath)
                        End If
                    End If
                Next log_Long
                This.LogFile.CloseFile
                Name This.LogFilePath As .BuildPath(.GetParentFolderName(This.LogFilePath), .GetBaseName(This.LogFilePath)) & ".0." & .GetExtensionName(This.LogFilePath)
            End If
            ' Create new, empty text file.
            Set This.LogFile = log_FSO.CreateTextFile(This.LogFilePath, True, False)
        End With
    End If
End Sub

''
' Open new console. Only run if pointer is -1, indicating we haven't launched a console yet.
'
' @method log_OpenConsole
''
#If Not Mac Then
Private Sub log_OpenConsole()
    On Error GoTo log_OpenConsole
    
    This.ConsoleHandle = FindWindow("ConsoleWindowClass", This.LogTitle)        ' Check if logging window already exists.
    If This.ConsoleHandle = 0 Then
        Shell "PowerShell.exe -WindowStyle Normal -NoExit -Command " & _
            "$Host.UI.RawUI.WindowTitle = '" & This.LogTitle & "'; " & _
            "get-content '" & This.LogFilePath & "' -Wait"                      ' Open PowerShell window for console logging.
        Application.Wait VBA.Now + VBA.TimeValue("0:00:02")                     ' Wait to open.
        This.ConsoleHandle = FindWindow("ConsoleWindowClass", This.LogTitle)    ' Get handle for window just opened using title to find it.
    End If
    SetForegroundWindow This.ConsoleHandle                                      ' Bring console to foreground.
    Exit Sub
    
log_OpenConsole:
    This.ConsoleHandle = 0 ' An error occured while opening console, set log handle to 0 so further logging will not be completed.
End Sub
#End If

''
' Get Log name for given `VbLogLevel`.
'
' @method log_LogLevelToText
' @param {VbLogLevel} Level
' @return {String}
''
Private Function log_LogLevelToText(ByVal Level As VbLogLevel) As String
    Select Case Level
    Case VbLogLevel.vbLogTrace
        log_LogLevelToText = "Trace"
    Case VbLogLevel.vbLogDebug
        log_LogLevelToText = "Debug"
    Case VbLogLevel.vbLogInfo
        log_LogLevelToText = "Info "
    Case VbLogLevel.vbLogWarn
        log_LogLevelToText = "WARN "
    Case VbLogLevel.vbLogError
        log_LogLevelToText = "ERROR"
    End Select
End Function

' ============================================= '
' Initialize & Terminate Methods
' ============================================= '

Public Sub Initialise(ByVal LogFilePath As String, ByVal LogTitle As String)
    ' Error handling.
    On Error GoTo lg_ErrorHandling
    If This.IsInitialised Then Err.Raise 1004, "Logger.Initialise", "Application-defined or object-defined error (class already intialised)."
    
    This.LogFilePath = LogFilePath
    This.LogTitle = LogTitle
    log_UpdateLogFiles
    This.ConsoleHandle = -1
    
    This.IsInitialised = True
    Exit Sub
    
lg_ErrorHandling:
    This.IsInitialised = False
    Select Case Err.Number
    Case 1004 ' Re-throw error raised earlier.
        Err.Raise Err.Number, Err.Source, Err.Description
    Case Else
        Err.Raise 1004, "Logger.Initialise", "Application-defined or object-defined error (error initialising class: " & Err.Description & ")."
    End Select
End Sub

Private Sub Class_Terminate()
    If Not This.LogFile Is Nothing Then This.LogFile.CloseFile
End Sub
