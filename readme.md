# VBA-ErrorHandling | [RadiusCore](https://radiuscore.co.nz) VBA Tools

__Status__: _Complete_

Error Handling and Logging for VBA.

# Development Environment

__Environment__

Recommended development environment is 64-bit Office-365 Excel running on Windows 10. Highly recommended to have [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck/) installed to improve the VBE, this is required for Unit Testing.

__Build__

[VBA-Git](https://github.com/VBA-Tools-v2/VBA-Git) is a tool that provides git repository support for Excel/VBA projects. It should be used to build an _.xlam_ file from repository source. VBA-Git will automatically download and add the following project dependencies to the compiled file:
- [VBA-Scripting](https://github.com/VBA-Tools-v2/VBA-Scripting)


# Usage
__Manual__

To manually add  VBA-ErrorHandling to a VBA project, import [ErrorHandler.bas](/src/vbProject/ErrorHandler.bas) and optionally, if logging is desired, [Logger.cls](/src/vbProject/Logger.cls) to the project. If `Logger.cls` is being used, include the requred files from [VBA-Scripting](https://github.com/VBA-Tools-v2/VBA-Scripting/) ([FileSystemObject.cls](https://github.com/VBA-Tools-v2/VBA-Scripting/blob/master/src/vbProject/Scripting/FileSystemObject.cls) & [TextStream.cls](https://github.com/VBA-Tools-v2/VBA-Scripting/blob/master/src/vbProject/Scripting/TextStream.cls)) or include a reference to `Microsoft Scripting Runtime`.

__VBA-Git__

To easily include VBA-ErrorHandling in any VBA project, use VBA-Git to build the target project, ensuring this repository is listed as a dependency in the project's configuration file. An example is included below, however additional information on how to do this can be found in the [VBA-Git ReadMe](https://github.com/VBA-Tools-v2/VBA-Git/blob/master/readme.md), 

```
"VBA-ErrorHandler": {
    "git": "https://github.com/VBA-Tools-v2/vba-errorhandler/",
    "tag": "v1.1.0",
    "key": "{readonly personal access token}"
    "src": ["ErrorHandler.bas"]
}
```

# Example

```VB.net
ErrorHandler.ApplicationName = "VBA-ErrorHandling"
' Enable logging by attaching a Logger to the Error Handler.
Set ErrorHandler.Log = New Logger
ErrorHandler.Log.Initialise LogFilePath:="C:/VBA-Log.log", LogTitle:="Test Log"
ErrorHandler.Log.LogThreshold = Info

' Show a warning error, which will log if a Logger is attached.
ErrorHandler.ShowWarn "An error has occured.", Err.Description, Err.Source, Err.Number, True
' -> 2022-07-01 21:37:50.00|ERROR|{Err.Source}|{Err.Number}, {Err.Description}.
' -> 2022-07-01 21:37:50.00|WARN |{Err.Source}|An error has occured.

' Directly log a warning using attached Logger.
ErrorHandler.LogWarn "Logging has started to the target file.", "ModuleName.MethodName"
' -> 2022-07-01 21:37:50.00|WARN |ModuleName.MethodName|Logging has started to the target file.
```