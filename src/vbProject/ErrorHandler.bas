Attribute VB_Name = "ErrorHandler"
''
' VBA-ErrorHandling: Error Handler
' (c) RadiusCore Ltd - https://www.radiuscore.co.nz/
'
' Show Errors to the user in clean messages. Optionally attach a Logger
' to integrate automatic logging of error messages.
'
' @module ErrorHandler
' @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder ErrorHandling
'@ignoremodule ProcedureNotUsed
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Private Module

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type TErrorHandler
    ApplicationName As String
    Log As Object ' Logger
    Err As Collection
End Type

Private This As TErrorHandler

' --------------------------------------------- '
' Types
' --------------------------------------------- '

''
' Whether to load or save to Cache.
'
' @proprty VbCache
' @param cacheSave
' @param cacheLoad
''
Public Enum VbCache
    vbCacheSave = 1
    vbCacheLoad = 2
End Enum

''
' Log Levels.
'
' @property vbLogLevel
' @param Off        @param Info
' @param Trace/All  @param Warn
' @param Debug      @param Error
''
Public Enum VbLogLevel
    vbLogOff = 0
    vbLogTrace = 1
    vbLogDebug = 2
    vbLogInfo = 3
    vbLogWarn = 4
    vbLogError = 5
End Enum

' --------------------------------------------- '
' Public Properties
' --------------------------------------------- '

''
' Application name to display in error message header.
'
' @property ApplicationName
' @type {String}
' @default vbNullString
''
Public Property Get ApplicationName() As String
    ApplicationName = This.ApplicationName
End Property
Public Property Let ApplicationName(ByVal Value As String)
    This.ApplicationName = Value
End Property

''
' Whether logging is enabled.
'
' @property IsLoggingEnabled
' @type {Boolean}
' @default False
''
Public Property Get IsLoggingEnabled() As Boolean
    IsLoggingEnabled = Not (This.Log Is Nothing)
End Property

''
' Attach a Logger to Error Handler, allowing error messages to be automatically logged.
'
' @property Log
' @type {Logger}
''
Public Property Get Log() As Object
    Set Log = This.Log
End Property
Public Property Set Log(ByVal Value As Object)
    If Not VBA.TypeName(Value) = "Logger" Then Err.Raise 13, "ErrorHandler", "Type mismatch"
    Set This.Log = Value
End Property

' ============================================= '
' Public Methods
' ============================================= '

''
' Cache Excel Error Object (`Err`) and reload it, so the error can be persistent through methods
' that require error handling to be used.
'
' @method ErrCache
' @param {rcCache} Operation | Whether to save to or load from the cache.
''
Public Sub ErrCache(ByVal Operation As VbCache)
    Select Case Operation
    Case VbCache.vbCacheSave
        Set This.Err = New Collection ' Reset cache.
        With This.Err
            .Add Err.Description, "Description"
            .Add Err.HelpContext, "HelpContext"
            .Add Err.HelpFile, "HelpFile"
            .Add Err.Number, "Number"
            .Add Err.Source, "Source"
        End With
    Case VbCache.vbCacheLoad
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

''
' @method LogTrace
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogTrace(ByVal Message As String, Optional ByVal From As String = vbNullString)
    If ErrorHandler.IsLoggingEnabled Then This.Log.Log VbLogLevel.vbLogTrace, Message, From
End Sub

''
' @method LogDebug
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogDebug(ByVal Message As String, Optional ByVal From As String = vbNullString)
    If ErrorHandler.IsLoggingEnabled Then This.Log.Log VbLogLevel.vbLogDebug, Message, From
End Sub

''
' @method LogInfo
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogInfo(ByVal Message As String, Optional ByVal From As String = vbNullString)
    If ErrorHandler.IsLoggingEnabled Then This.Log.Log VbLogLevel.vbLogInfo, Message, From
End Sub

''
' @method LogWarning
' @param {String} Message
' @param {String} [From = ""]
''
Public Sub LogWarn(ByVal Message As String, Optional ByVal From As String = vbNullString)
    If ErrorHandler.IsLoggingEnabled Then This.Log.Log VbLogLevel.vbLogWarn, Message, From
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
        If ErrNumber < 0 Then log_ErrorValue = log_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        log_ErrorValue = log_ErrorValue & ", "
    End If
    If ErrorHandler.IsLoggingEnabled Then This.Log.Log VbLogLevel.vbLogError, log_ErrorValue & Message, From
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
    ' Log if possible.
    If Log And Not This.Log Is Nothing Then
        If Not (ErrDescription = vbNullString And ErrSource = vbNullString And ErrNumber = 0) Then This.Log.LogError VBA.Replace(ErrDescription, vbNewLine, VBA.Chr$(32)), ErrSource, ErrNumber
        This.Log.LogWarn Message, ErrSource
    End If
    ' Show error message.
    VBA.MsgBox Message & _
               VBA.IIf(ErrDescription = vbNullString And ErrSource = vbNullString And ErrNumber = 0, vbNullString, vbNewLine & vbNewLine & "---Error Information---" & vbNewLine) & _
               VBA.IIf(ErrDescription = vbNullString, vbNullString, VBA.Replace(ErrDescription, vbNewLine, " ") & vbNewLine) & _
               VBA.IIf(ErrSource = vbNullString, vbNullString, "[" & ErrSource & "]" & vbNewLine) & _
               VBA.IIf(ErrNumber = 0, vbNullString, "(" & ErrNumber - vbObjectError & " / " & ErrNumber & VBA.IIf(ErrNumber < 0, " / " & VBA.LCase$(VBA.Hex$(ErrNumber)), vbNullString) & ")"), _
               vbExclamation + vbOKOnly, VBA.IIf(This.ApplicationName = vbNullString, vbNullString, This.ApplicationName & " | ") & "Warning"
End Sub

''
' Display error message in dialogue box.
'
' @method ShowError
' @param {String} Message
' @param {String} [ErrDescription = ""]
' @param {String} [ErrSource = ""]
' @param {Long} [ErrNumber = 0]
' @param {Boolean} [Log = True]
''
Public Sub ShowError(ByVal Message As String, Optional ByVal ErrDescription As String = vbNullString, Optional ByVal ErrSource As String = vbNullString, Optional ByVal ErrNumber As Long = 0, Optional ByVal Log As Boolean = True)
    ' Log if possible.
    If Log And Not This.Log Is Nothing Then
        If Log Then This.Log.LogError VBA.Replace(Message, vbNewLine, VBA.Chr$(32)) & VBA.Chr$(32) & VBA.Replace(Err.Description, vbNewLine, VBA.Chr$(32)), ErrSource, ErrNumber
    End If
    ' Show error message.
    VBA.MsgBox Message & vbNewLine & _
               VBA.IIf(ErrDescription = vbNullString, vbNullString, vbNewLine & ErrDescription & vbNewLine) & _
               VBA.IIf(ErrSource = vbNullString, vbNullString, "[" & ErrSource & "]" & vbNewLine) & _
               VBA.IIf(ErrNumber = 0, vbNullString, "(" & ErrNumber - vbObjectError & " / " & ErrNumber & VBA.IIf(ErrNumber < 0, " / " & VBA.LCase$(VBA.Hex$(ErrNumber)), vbNullString) & ")" & vbNewLine) & _
               VBA.IIf(ErrDescription = vbNullString And ErrSource = vbNullString And ErrNumber = 0, vbNullString, vbNewLine & "If this error is persistent please contact the developer."), _
               vbCritical + vbOKOnly, VBA.IIf(This.ApplicationName = vbNullString, vbNullString, This.ApplicationName & " | ") & "Error"
End Sub
