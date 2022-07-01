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
' Logging & Error Handler for VBA.
'
' @module Logger
' @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
' @depend FileSystemObject, TextStream
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder VBA-ErrorHandling
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
    Err As Collection
End Type

Private This As TLogger

' --------------------------------------------- '
' Types
' --------------------------------------------- '

''
' Log Levels.
'
' @property vbLogLevel
' @param Off        @param Info
' @param Trace/All  @param Warn
' @param Debug      @param Error
''
Public Enum VbLogLevel
    Off = 0
    Trace = 1
    Debugg = 2
    Info = 3
    Warn = 4
    Error = 5
End Enum

''
' Whether to load or save to Cache.
'
' @proprty VbCache
' @param cacheSave
' @param cacheLoad
''
Public Enum VbCache
    cacheSave = 1
    cacheLoad = 2
End Enum

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
    LogEnabled = Not (LogThreshold = VbLogLevel.Off)
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
' @method LogTrace
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogTrace(ByVal Message As String, Optional ByVal From As String = vbNullString)
    log_Log VbLogLevel.Trace, Message, From
End Sub

''
' @method LogDebug
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogDebug(ByVal Message As String, Optional ByVal From As String = vbNullString)
    log_Log VbLogLevel.Debugg, Message, From
End Sub

''
' @method LogInfo
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogInfo(ByVal Message As String, Optional ByVal From As String = vbNullString)
    log_Log VbLogLevel.Info, Message, From
End Sub

''
' @method LogWarning
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogWarn(ByVal Message As String, Optional ByVal From As String = vbNullString)
    log_Log VbLogLevel.Warn, Message, From
End Sub

''
' @method LogError
' @param {String} Message
' @param {String} [From = ""]
' @param {Long} [ErrNumber = 0]
''
Public Sub LogError(ByVal Message As String, Optional ByVal From As String = vbNullString, Optional ByVal ErrNumber As Long = 0)
    Dim log_ErrorValue As String
    If Not ErrNumber = 0 Then
        log_ErrorValue = ErrNumber
    
        ' For object errors, extract from vbObjectError and get Hex value
        If ErrNumber < 0 Then
            log_ErrorValue = log_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If
        
        log_ErrorValue = log_ErrorValue & ", "
    End If

    log_Log VbLogLevel.Error, log_ErrorValue & Message, From
End Sub

''
' Display error message as warning in dialogue box.
'
' @method ShowWarn
' @param {String} Message
' @param {String} [ErrDescription = ""]
' @param {String} [ErrSource = ""]
' @param {Long} [ErrNumber = 0]
' @param {Boolean} [Log = True]
''
Public Sub ShowWarn(ByVal Message As String, Optional ByVal ErrDescription As String = vbNullString, Optional ByVal ErrSource As String = vbNullString, Optional ByVal ErrNumber As Long = 0, Optional ByVal Log As Boolean = True)
    If Not This.IsInitialised Then Err.Raise 1004, "Logger", "Application-defined or object-defined error (class not intialised)."
    If Log Then
        If Not (ErrDescription = vbNullString And ErrSource = vbNullString And ErrNumber = 0) Then LogError VBA.Replace(ErrDescription, vbNewLine, VBA.Chr$(32)), ErrSource, ErrNumber
        LogWarn Message, ErrSource
    End If
    VBA.MsgBox Message & _
               VBA.IIf(ErrDescription = vbNullString And ErrSource = vbNullString And ErrNumber = 0, vbNullString, vbNewLine & vbNewLine & "---Error Information---" & vbNewLine) & _
               VBA.IIf(ErrDescription = vbNullString, vbNullString, VBA.Replace(ErrDescription, vbNewLine, " ") & vbNewLine) & _
               VBA.IIf(ErrSource = vbNullString, vbNullString, "[" & ErrSource & "]" & vbNewLine) & _
               VBA.IIf(ErrNumber = 0, vbNullString, "(" & ErrNumber - vbObjectError & " / " & ErrNumber & VBA.IIf(ErrNumber < 0, " / " & VBA.LCase$(VBA.Hex$(ErrNumber)), vbNullString) & ")"), _
               vbExclamation + vbOKOnly, "Warning"
End Sub

''
' Display error message in dialogue box.
'
' @method ShowError
' @param {String} Message
' @param {String} [From = ""]
' @param {Long} [ErrNumber = 0]
' @param {Boolean} [Log = True]
''
Public Sub ShowError(ByVal Message As String, Optional ByVal ErrDescription As String = vbNullString, Optional ByVal ErrSource As String = vbNullString, Optional ByVal ErrNumber As Long = 0, Optional ByVal Log As Boolean = True)
    If Not This.IsInitialised Then Err.Raise 1004, "Logger", "Application-defined or object-defined error (class not intialised)."
    If Log Then LogError VBA.Replace(Message, vbNewLine, VBA.Chr$(32)) & VBA.Chr$(32) & VBA.Replace(Err.Description, vbNewLine, VBA.Chr$(32)), ErrSource, ErrNumber
    VBA.MsgBox Message & vbNewLine & _
               VBA.IIf(ErrDescription = vbNullString, vbNullString, vbNewLine & ErrDescription & vbNewLine) & _
               VBA.IIf(ErrSource = vbNullString, vbNullString, "[" & ErrSource & "]" & vbNewLine) & _
               VBA.IIf(ErrNumber = 0, vbNullString, "(" & ErrNumber - vbObjectError & " / " & ErrNumber & VBA.IIf(ErrNumber < 0, " / " & VBA.LCase$(VBA.Hex$(ErrNumber)), vbNullString) & ")" & vbNewLine) & _
               VBA.IIf(ErrDescription = vbNullString And ErrSource = vbNullString And ErrNumber = 0, vbNullString, vbNewLine & "If this error is persistent please contact the developer."), _
               vbCritical + vbOKOnly, "Error"
End Sub

''
' Cache Excel Error Object (`Err`) and reload it, so the error can be persistent through methods
' that require error handling to be used.
'
' @method ErrCache
' @param {VbCache} Operation | Whether to save to or load from the cache.
''
Public Sub ErrCache(ByVal Operation As VbCache)
    Select Case Operation
    Case VbCache.cacheSave
        Set This.Err = New Collection ' Reset cache.
        With This.Err
            .Add Err.Description, "Description"
            .Add Err.HelpContext, "HelpContext"
            .Add Err.HelpFile, "HelpFile"
            .Add Err.Number, "Number"
            .Add Err.Source, "Source"
        End With
    Case VbCache.cacheLoad
        If Not This.Err Is Nothing Then
            With This.Err
                Err.Description = .Item("Description")
                Err.HelpContext = .Item("HelpContext")
                Err.HelpFile = .Item("HelpFile")
                Err.Number = .Item("Number")
                Err.Source = .Item("Source")
            End With
        End If
    End Select
End Sub

' ============================================= '
' Private Methods
' ============================================= '

''
' @method log_Log
' @param {VbLogLevel} Level
' @param {String} Message
' @param {String} [From = ""]
''
Private Sub log_Log(ByVal Level As VbLogLevel, ByVal Message As String, Optional ByVal From As String = vbNullString)
    On Error GoTo log_ErrorHandling
    If Not This.IsInitialised Then Err.Raise 1004, "Logger", "Application-defined or object-defined error (class not intialised)."
    If LogThreshold = VbLogLevel.Off Or Level < LogThreshold Then Exit Sub      ' Exit if log threshold is not high enough.
    If VBA.FileLen(This.LogFilePath) >= 1024000 Then log_UpdateLogFiles         ' Create new log file if size limit is met.
#If Not Mac Then
    If This.Console And This.ConsoleHandle = -1 Then log_OpenConsole            ' Show log in console, if enabled.
#End If
    
    Dim lg_Text As String
    lg_Text = VBA.Format$(VBA.Now, "yyyy-mm-dd hh:mm:ss") & "." & VBA.Right$(VBA.Format$(VBA.Timer, "#0.00"), 2) & "|" & _
              log_LogLevelToText(Level) & "|" & _
              VBA.IIf(From = vbNullString, vbNullString, From & "|") & _
              Message                                                           ' Compile string to log.
    
    If This.Immediate Then Debug.Print lg_Text                                  ' Log to immediate window.
    This.LogFile.WriteLine lg_Text                                              ' Log to file.
    Exit Sub
    
log_ErrorHandling:
    Debug.Print "Error Logging (" & Err.Number & ": " & Err.Description & ")"   ' Suppress any errors while logging. Don't want to break application due to this.
End Sub

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
        Shell "PowerShell.exe -NoExit -Command " & _
            "$Host.UI.RawUI.WindowTitle = '" & This.LogTitle & "'; " & _
            "get-content '" & This.LogFilePath & "' -Wait"                      ' Open PowerShell window for console logging.
        Application.Wait VBA.Now + VBA.TimeValue("0:00:02")                     ' Wait to open.
        This.ConsoleHandle = FindWindow("ConsoleWindowClass", This.LogTitle)    ' Get handle for window just opened. , using title to find it.
    Else
        SetForegroundWindow This.ConsoleHandle                                  ' Window already open, bring to foreground.
    End If
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
    Case VbLogLevel.Trace
        log_LogLevelToText = "Trace"
    Case VbLogLevel.Debugg
        log_LogLevelToText = "Debug"
    Case VbLogLevel.Info
        log_LogLevelToText = "Info "
    Case VbLogLevel.Warn
        log_LogLevelToText = "WARN "
    Case VbLogLevel.Error
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
