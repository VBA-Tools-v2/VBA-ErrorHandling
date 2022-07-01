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

To manually add  VBA-ErrorHandling to a VBA project, import __either__ [ErrorHandler.bas](/src/vbProject/ErrorHandler.bas) _or_ [Logger.cls](/src/vbProject/Logger.cls) to the project. If `Logger.cls` is being used, include the requred files from [VBA-Scripting](https://github.com/VBA-Tools-v2/VBA-Scripting/) ([FileSystemObject.cls](https://github.com/VBA-Tools-v2/VBA-Scripting/blob/master/src/vbProject/Scripting/FileSystemObject.cls) & [TextStream.cls](https://github.com/VBA-Tools-v2/VBA-Scripting/blob/master/src/vbProject/Scripting/TextStream.cls)) or include a reference to `Microsoft Scripting Runtime`.

__VBA-Git__

To easily include VBA-ErrorHandling in any VBA project, use VBA-Git to build the target project, ensuring this repository is listed as a dependency in the project's configuration file. An example is included below, however additional information on how to do this can be found in the [VBA-Git ReadMe](https://github.com/VBA-Tools-v2/VBA-Git/blob/master/readme.md), 

```
"VBA-ErrorHandler": {
    "git": "https://github.com/VBA-Tools-v2/vba-errorhandler/",
    "tag": "v1.0.0",
    "key": "{readonly personal access token}"
    "src": ["Logger.cls"]
}
```


# Example

```VB.net
Dim xl_Log As Logger
Set xl_Log = New Logger
xl_Log.Initialise LogFilePath:="C:/VBA-Log.log", LogTitle:="Test Log"
xl_Log.LogThreshold = Info

xl_Log.LogWarn "Logging has started to the target file.", "ModuleName.MethodName"
' -> 2022-07-01 21:37:50.00|WARN |ModuleName.MethodName|Logging has started to the target file.
```