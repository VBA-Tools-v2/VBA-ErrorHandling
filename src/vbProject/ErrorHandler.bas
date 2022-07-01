Attribute VB_Name = "ErrorHandler"
''
' VBA-ErrorHandling: Error Handler
' (c) RadiusCore Ltd - https://www.radiuscore.co.nz/
'
' Show Errors to the user in clean messages.
'
' @module ErrorHandler
' @author Andrew Pullon | andrew.pullon@radiuscore.co.nz | andrewcpullon@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
' RubberDuck Annotations
' https://rubberduckvba.com/ | https://github.com/rubberduck-vba/Rubberduck/
'
'@folder VBA-ErrorHandling
'@ignoremodule ProcedureNotUsed
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit
Option Private Module

' --------------------------------------------- '
' Constants and Private Variables
' --------------------------------------------- '

Private Type THelper
    Err As Collection
End Type

'@Ignore MoveFieldCloserToUsage
Private This As THelper
Private Const APPLICATIONNAME As String = vbNullString

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
    cacheSave = 1
    cacheLoad = 2
End Enum

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

''
' Display error message as warning in dialogue box.
'
' @method ShowWarn
' @param {String} Message
' @param {String} [ErrDescription = ""]
' @param {String} [From = ""]
' @param {Long} [ErrNumber = 0]
''
Public Sub ShowWarn(ByVal Message As String, Optional ByVal ErrDescription As String = vbNullString, Optional ByVal From As String = vbNullString, Optional ByVal ErrNumber As Long = 0)
    VBA.MsgBox Message & _
               VBA.IIf(ErrDescription = vbNullString And From = vbNullString And ErrNumber = 0, vbNullString, vbNewLine & vbNewLine & "---Error Information---" & vbNewLine) & _
               VBA.IIf(ErrDescription = vbNullString, vbNullString, VBA.Replace(ErrDescription, vbNewLine, " ") & vbNewLine) & _
               VBA.IIf(From = vbNullString, vbNullString, "[" & From & "]" & vbNewLine) & _
               VBA.IIf(ErrNumber = 0, vbNullString, "(" & ErrNumber - vbObjectError & " / " & ErrNumber & VBA.IIf(ErrNumber < 0, " / " & VBA.LCase$(VBA.Hex$(ErrNumber)), vbNullString) & ")"), _
               vbExclamation + vbOKOnly, VBA.IIf(APPLICATIONNAME = vbNullString, vbNullString, APPLICATIONNAME & " | ") & "Warning"
End Sub

''
' Display error message in dialogue box.
'
' @method ShowError
' @param {String} Message
' @param {String} [ErrDescription = ""]
' @param {String} [From = ""]
' @param {Long} [ErrNumber = 0]
''
Public Sub ShowError(ByVal Message As String, Optional ByVal ErrDescription As String = vbNullString, Optional ByVal From As String = vbNullString, Optional ByVal ErrNumber As Long = 0)
    VBA.MsgBox Message & vbNewLine & _
               VBA.IIf(ErrDescription = vbNullString, vbNullString, vbNewLine & ErrDescription & vbNewLine) & _
               VBA.IIf(From = vbNullString, vbNullString, "[" & From & "]" & vbNewLine) & _
               VBA.IIf(ErrNumber = 0, vbNullString, "(" & ErrNumber - vbObjectError & " / " & ErrNumber & VBA.IIf(ErrNumber < 0, " / " & VBA.LCase$(VBA.Hex$(ErrNumber)), vbNullString) & ")" & vbNewLine) & _
               VBA.IIf(ErrDescription = vbNullString And From = vbNullString And ErrNumber = 0, vbNullString, vbNewLine & "If this error is persistent please contact the developer."), _
               vbCritical + vbOKOnly, VBA.IIf(APPLICATIONNAME = vbNullString, vbNullString, APPLICATIONNAME & " | ") & "Error"
End Sub
