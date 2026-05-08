Attribute VB_Name = "xlwings"
#Const App = "Microsoft Excel" 'Adjust when using outside of Excel
'Version: 0.33.16

'xlwings is distributed under a BSD 3-clause license.
'
'Copyright (C) 2014-present, Zoomer Analytics LLC.
'All rights reserved.
'
'Redistribution and use in source and binary forms, with or without modification,
'are permitted provided that the following conditions are met:
'
'* Redistributions of source code must retain the above copyright notice, this
'  list of conditions and the following disclaimer.
'
'* Redistributions in binary form must reproduce the above copyright notice, this
'  list of conditions and the following disclaimer in the documentation and/or
'  other materials provided with the distribution.
'
'* Neither the name of the copyright holder nor the names of its
'  contributors may be used to endorse or promote products derived from
'  this software without specific prior written permission.
'
'THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
'ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
'WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
'DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR
'ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
'(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
'LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON
'ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

'Attribute VB_Name = "Main"



#If VBA7 Then
    #If Mac Then
        Private Declare PtrSafe Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    #If Win64 Then
        Const XLPyDLLName As String = "xlwings64-0.33.16.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings64-0.33.16.dll" (ByRef Result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
        Declare PtrSafe Function XLPyDLLNDims Lib "xlwings64-0.33.16.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Declare PtrSafe Function XLPyDLLVersion Lib "xlwings64-0.33.16.dll" (tag As String, VERSION As Double, arch As String) As Long
    #Else
        Private Const XLPyDLLName As String = "xlwings32-0.33.16.dll"
        Declare PtrSafe Function XLPyDLLActivateAuto Lib "xlwings32-0.33.16.dll" (ByRef Result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
        Private Declare PtrSafe Function XLPyDLLNDims Lib "xlwings32-0.33.16.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
        Private Declare PtrSafe Function XLPyDLLVersion Lib "xlwings32-0.33.16.dll" (tag As String, VERSION As Double, arch As String) As Long
    #End If
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#Else
    #If Mac Then
        Private Declare Function system Lib "libc.dylib" (ByVal Command As String) As Long
    #End If
    Private Const XLPyDLLName As String = "xlwings32-0.33.16.dll"
    Private Declare Function XLPyDLLActivateAuto Lib "xlwings32-0.33.16.dll" (ByRef Result As Variant, Optional ByVal Config As String = "", Optional ByVal mode As Long = 1) As Long
    Private Declare Function XLPyDLLNDims Lib "xlwings32-0.33.16.dll" (ByRef src As Variant, ByRef dims As Long, ByRef transpose As Boolean, ByRef dest As Variant) As Long
    Private Declare Function LoadLibrary Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
    Declare Function XLPyDLLVersion Lib "xlwings32-0.33.16.dll" (tag As String, VERSION As Double, arch As String) As Long
#End If

Public Const XLWINGS_VERSION As String = "0.33.16"
Public Const PROJECT_NAME As String = "xlwings"

Public Function RunPython(pythonCommand As String)
    ' Public API: Runs the Python command, e.g.: to run the function foo() in module bar, call the function like this:
    ' RunPython "import bar; bar.foo()"
    
    Dim i As Integer
    Dim SourcePythonCommand As String, interpreter As String, PYTHONPATH As String, licenseKey, ActiveFullName As String, ThisFullName As String, AddExcelDir As String
    Dim OPTIMIZED_CONNECTION As Boolean, uses_embedded_code As Boolean
    Dim Wb As Workbook
    Dim sht As Worksheet
    
    ' svpython ?omFpbZ[W{bNX
    ' Call MsgBox(pythonCommand, vbInformation, "xlwings.RunPython")

'    'MsgBox "merge clicked"
    SourcePythonCommand = pythonCommand
    
    #If Mac Then
        interpreter = GetConfig("INTERPRETER_MAC", "")
    #Else
        interpreter = GetConfig("INTERPRETER_WIN", "")
    #End If
    If interpreter = "" Then
        ' Legacy
        ' interpreter = GetConfig("INTERPRETER", "python")
'        interpreter = GetConfig("INTERPRETER", "C:\work\Python\Project\Excel_AddIn\.venv\Scripts\python.exe")
        interpreter = GetConfig("INTERPRETER", "C:\Project\Python\Excel_AddIn\.venv\Scripts\python.exe")
    End If

    ' --- Guard: never allow bare python/pythonw on Windows (prevents PATH -> system python) ---
    #If Not Mac Then
        Call HC_Log.Diag("xlwings", "RunPython resolved interpreter(raw)=" & interpreter)

        If LCase$(interpreter) = "python" Or LCase$(interpreter) = "pythonw" Then
            Dim fallbackInterpreter As String
            fallbackInterpreter = GetConfig("INTERPRETER", "C:\Project\Python\Excel_AddIn\.venv\Scripts\python.exe")
            Call HC_Log.Diag("xlwings", "Guard hit: interpreter was bare. Fallback INTERPRETER=" & fallbackInterpreter)

            ' If fallback is still bare, stop with an explicit error (SSOT)
            If LCase$(fallbackInterpreter) = "python" Or LCase$(fallbackInterpreter) = "pythonw" Or fallbackInterpreter = "" Then
                MsgBox "xlwings interpreter is set to bare python/pythonw. " & vbCrLf & _
                       "This may start system Python via PATH and cause double-launch." & vbCrLf & _
                       "Please set INTERPRETER_WIN to your venv python.exe in xlwings.conf.", vbCritical, "xlwings interpreter error"
                RunPython = -1
                Exit Function
            End If

            interpreter = fallbackInterpreter
            Call HC_Log.Diag("xlwings", "Guard applied: interpreter(final)=" & interpreter)
        End If
    #End If

    ' Check for embedded Python code
    uses_embedded_code = False
    For i = 1 To 2
        If i = 1 Then
            Set Wb = ThisWorkbook
        Else
            Set Wb = ThisWorkbook
        End If
        For Each sht In Wb.Worksheets
            If Right$(sht.Name, 3) = ".py" Then
                uses_embedded_code = True
                Exit For
            End If
        Next
    Next i

    If uses_embedded_code = True Then
        AddExcelDir = "false"
    Else
        AddExcelDir = GetConfig("ADD_WORKBOOK_TO_PYTHONPATH", "true")
    End If

    ' The first 5 args are not technically part of the PYTHONPATH, but it's just easier to add it here (used by xlwings.utils.prepare_sys_path)
    #If Mac Then
        If InStr(ThisWorkbook.FullName, "://") = 0 Then
            ActiveFullName = ToPosixPath(ThisWorkbook.FullName)
            ThisFullName = ToPosixPath(ThisWorkbook.FullName)
        Else
            ActiveFullName = ThisWorkbook.FullName
            ThisFullName = ThisWorkbook.FullName
        End If
    #Else
        ActiveFullName = ThisWorkbook.FullName
        ThisFullName = ThisWorkbook.FullName
    #End If
    
    #If Mac Then
        PYTHONPATH = AddExcelDir & ";" & ActiveFullName & ";" & ThisFullName & ";" & GetConfig("ONEDRIVE_CONSUMER_MAC") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_MAC") & ";" & GetConfig("SHAREPOINT_MAC") & ";" & GetConfig("PYTHONPATH")
    #Else
        PYTHONPATH = AddExcelDir & ";" & ActiveFullName & ";" & ThisFullName & ";" & GetConfig("ONEDRIVE_CONSUMER_WIN") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_WIN") & ";" & GetConfig("SHAREPOINT_WIN") & ";" & GetConfig("PYTHONPATH")
    #End If

    OPTIMIZED_CONNECTION = GetConfig("USE UDF SERVER", False)

    ' PythonCommand with embedded code
    If uses_embedded_code = True Then
        licenseKey = GetConfig("LICENSE_KEY")
        If licenseKey = "" Then
            MsgBox "Embedded code requires a valid LICENSE_KEY."
            Exit Function
        Else
            pythonCommand = "import xlwings.pro;xlwings.pro.runpython_embedded_code('" & SourcePythonCommand & "')"
        End If
    End If

    ' Call Python platform-dependent
    #If Mac Then
        Application.StatusBar = "Running..."  ' Non-blocking way of giving feedback that something is happening
        ExecuteMac pythonCommand, interpreter, PYTHONPATH
    #Else
    
    Dim v As Variant
    v = GetConfig("USE UDF SERVER", False)
    Call HC_Log.Diag("xlwings", "USE UDF SERVER raw=" & CStr(v) & " vartype=" & CStr(VarType(v)))
    Call HC_Log.Diag("xlwings", "OPTIMIZED_CONNECTION = " & CStr(OPTIMIZED_CONNECTION))

    
        If OPTIMIZED_CONNECTION = True Then
            
            XLPy.SetAttr XLPy.Module("xlwings._xlwindows"), "BOOK_CALLER", ThisWorkbook
            
            On Error GoTo err_handling
            
            XLPy.Exec "" & pythonCommand & ""
            GoTo end_err_handling
err_handling:
            ShowError "", Err.Description
            RunPython = -1
            On Error GoTo 0
end_err_handling:
        Else
            If HcTruthyPackagedRunPython() Then
                If Len(Trim$(Environ$("HC_INSTALL_ROOT"))) = 0 Then
                    MsgBox "HC_INSTALL_ROOT is not set. Packaged RunPython requires install root.", vbCritical, "xlwings"
                    RunPython = -1
                ElseIf Not FileExists(interpreter) Then
                    MsgBox "Interpreter not found: " & interpreter, vbCritical, "xlwings"
                    RunPython = -1
                Else
                    Dim scriptPath As String
                    Dim fullPy As String
                    Dim frozenArgs As String
                    fullPy = "import os, sys" & vbLf
                    fullPy = fullPy & "install = (os.environ.get('HC_INSTALL_ROOT') or '').strip()" & vbLf
                    fullPy = fullPy & "if install:" & vbLf
                    fullPy = fullPy & "    if install not in sys.path:" & vbLf
                    fullPy = fullPy & "        sys.path.insert(0, install)" & vbLf
                    fullPy = fullPy & "    try:" & vbLf
                    fullPy = fullPy & "        os.chdir(install)" & vbLf
                    fullPy = fullPy & "    except OSError:" & vbLf
                    fullPy = fullPy & "        pass" & vbLf
                    fullPy = fullPy & "else:" & vbLf
                    fullPy = fullPy & "    root = r'" & Replace(ThisWorkbook.Path, "'", "''") & "'" & vbLf
                    fullPy = fullPy & "    cand = os.path.dirname(root)" & vbLf
                    fullPy = fullPy & "    hc = os.path.join(root, 'hc_main.py')" & vbLf
                    fullPy = fullPy & "    hc2 = os.path.join(cand, 'hc_main.py')" & vbLf
                    fullPy = fullPy & "    if (not os.path.exists(hc)) and os.path.exists(hc2):" & vbLf
                    fullPy = fullPy & "        root = cand" & vbLf
                    fullPy = fullPy & "    if root and root not in sys.path:" & vbLf
                    fullPy = fullPy & "        sys.path.insert(0, root)" & vbLf
                    fullPy = fullPy & pythonCommand
                    scriptPath = GetConfig("TEMP DIR", Environ("Temp")) & "\xlwings-snippet-" & CreateGUID() & ".py"
                    Call HcWriteUtf8File(scriptPath, fullPy)
                    frozenArgs = "--script-file=" & Chr(34) & scriptPath & Chr(34)
                    RunPython = ExecuteWindows(True, interpreter, ParentFolder(interpreter), PYTHONPATH, frozenArgs)
                    On Error Resume Next
                    Kill scriptPath
                    On Error GoTo 0
                End If
            Else
                RunPython = ExecuteWindows(False, pythonCommand, interpreter, PYTHONPATH)
            End If
        End If
    #End If
    
End Function


Sub ExecuteMac(pythonCommand As String, PYTHON_MAC As String, Optional PYTHONPATH As String)
    #If Mac Then
    Dim PythonInterpreter As String, RunCommand As String, Log As String
    Dim ParameterString As String, ExitCode As String, CondaCmd As String, CondaPath As String, CondaEnv As String, LOG_FILE As String

    ' Transform paths
    PYTHONPATH = Replace(PYTHONPATH, "'", "\'") ' Escaping quotes

    If PYTHON_MAC <> "" Then
        If PYTHON_MAC <> "python" And PYTHON_MAC <> "pythonw" Then
            PythonInterpreter = ToPosixPath(PYTHON_MAC)
        Else
            PythonInterpreter = PYTHON_MAC
        End If
    Else
        PythonInterpreter = "python"
    End If

    ' Sandbox location that requires no file access confirmation
    ' TODO: Use same logic with GUID like for Windows. Only here the GUID will need to be passed back to CleanUp()
    LOG_FILE = Environ("HOME") + "/xlwings.log" '/Users/<User>/Library/Containers/com.microsoft.Excel/Data/xlwings.log

    ' Delete Log file just to make sure we don't show an old error
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0

    ' ParameterSting with all paramters (AppleScriptTask only accepts a single parameter)
    ParameterString = PYTHONPATH + ";"
    ParameterString = ParameterString + "|" + PythonInterpreter
    ParameterString = ParameterString + "|" + pythonCommand
    ParameterString = ParameterString + "|" + ThisWorkbook.Name
    ParameterString = ParameterString + "|" + Left(Application.Path, Len(Application.Path) - 4)
    ParameterString = ParameterString + "|" + LOG_FILE

    On Error GoTo AppleScriptErrorHandler
        ExitCode = AppleScriptTask("xlwings-" & XLWINGS_VERSION & ".applescript", "VbaHandler", ParameterString)
    On Error GoTo 0

    ' If there's a log at this point (normally that will be from the shell only, not Python) show it and reset the StatusBar
    On Error Resume Next
        Log = ReadFile(LOG_FILE)
        If Log = "" Then
            Exit Sub
        Else
            ShowError (LOG_FILE)
            Application.StatusBar = False
        End If
        Exit Sub
    On Error GoTo 0

AppleScriptErrorHandler:
    MsgBox "To enable RunPython, please run 'xlwings runpython install' in a terminal once and try again.", vbCritical
    #End If
End Sub



Private Function HcTruthyPackagedRunPython() As Boolean
    Dim s As String
    Dim e As String
    s = LCase$(Trim$(CStr(GetConfig("USE_PACKAGED_RUNPYTHON", "False"))))
    HcTruthyPackagedRunPython = (s = "true" Or s = "1" Or s = "yes")
    e = LCase$(Trim$(Environ$("HC_PACKAGED_DEPLOYMENT")))
    If (e = "1" Or e = "true" Or e = "yes") Then
        HcTruthyPackagedRunPython = True
    End If
End Function

Private Sub HcWriteUtf8File(ByVal filePath As String, ByVal content As String)
    Const AD_TYPE_TEXT As Long = 2
    Const AD_SAVE_OVERWRITE As Long = 2
    Const AD_WRITE_CHAR As Long = 0
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = AD_TYPE_TEXT
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content, AD_WRITE_CHAR
    stm.SaveToFile filePath, AD_SAVE_OVERWRITE
    stm.Close
    Set stm = Nothing
End Sub
Function ExecuteWindows(IsFrozen As Boolean, pythonCommand As String, PYTHON_WIN As String, _
                        Optional PYTHONPATH As String, Optional FrozenArgs As String) As Integer
    ' Call a command window and change to the directory of the Python installation or frozen executable
    ' Note: If Python is called from a different directory with the fully qualified path, pywintypesXX.dll won't be found.
    ' This seems to be a general issue with pywin32, see http://stackoverflow.com/q/7238403/918626
    Dim ShowConsole As Integer
    Dim TempDir As String
    
    ' fBobNp 2026/0215
    Call HC_Log.Diag("xlwings", "Root ExecuteWindows")
    
    If GetConfig("SHOW CONSOLE", False) = True Then
        ShowConsole = 1
    Else
        ShowConsole = 0
    End If

    Dim WaitOnReturn As Boolean: WaitOnReturn = True
    Dim WindowStyle As Integer: WindowStyle = ShowConsole
    Dim DriveCommand As String, RunCommand, condaExcecutable As String
    Dim PythonInterpreter As String, PythonDir As String, CondaCmd As String, CondaPath As String, CondaEnv As String
    Dim ExitCode As Long
    Dim LOG_FILE As String
    
    TempDir = GetConfig("TEMP DIR", Environ("Temp")) 'undocumented setting
    
    LOG_FILE = TempDir & "\xlwings-" & CreateGUID() & ".log"

    If Not IsFrozen And (PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw") Then
        If FileExists(PYTHON_WIN) Then
            PythonDir = ParentFolder(PYTHON_WIN)
        Else
            MsgBox "Could not find Interpreter!", vbCritical
            Exit Function
        End If
    Else
        PythonDir = ""  ' TODO: hack
    End If

    If Left$(PYTHON_WIN, 2) Like "[A-Za-z]:" Then
        ' If Python is installed on a mapped or local drive, change to drive, then cd to path
        DriveCommand = Left$(PYTHON_WIN, 2) & " & cd """ & PythonDir & """ & "
    ElseIf Left$(PYTHON_WIN, 2) = "\\" Then
        ' If Python is installed on a UNC path, temporarily mount and activate a drive letter with pushd
        DriveCommand = "pushd """ & PythonDir & """ & "
    End If

    ' Run Python with the "-c" command line switch: add the path of the python file and run the
    ' Command as first argument, then provide the Name and "from_xl" as 2nd and 3rd arguments.
    ' Then redirect stderr to the LOG_FILE and wait for the call to return.

    If PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw" Then
        PythonInterpreter = Chr(34) & PYTHON_WIN & Chr(34)
    Else
        ' --- ?X ---
        MsgBox "ExecuteWindows received bare python/pythonw." & vbCrLf & _
               "This may start system Python via PATH." & vbCrLf & _
               "Please set INTERPRETER_WIN to your venv python.exe in xlwings.conf.", _
               vbCritical, "xlwings interpreter error"
    
        ExecuteWindows = -1
        Exit Function
    End If
    
    Call HC_Log.Diag("xlwings", "ExecuteWindows PYTHON_WIN=" & PYTHON_WIN)
    Call HC_Log.Diag("xlwings", "ExecuteWindows PythonInterpreter=" & PythonInterpreter)
    Call HC_Log.Diag("xlwings", "ExecuteWindows pythonCommand=" & pythonCommand)
    Call HC_Log.Diag("xlwings", "ExecuteWindows PYTHONPATH=" & PYTHONPATH)

    CondaPath = GetConfig("CONDA PATH")
    CondaEnv = GetConfig("CONDA ENV")
    
    ' Handle spaces in path (for UDFs, this is handled via nested quotes instead, see XLPyCommand)
    CondaPath = Replace(CondaPath, " ", "^ ")
    
    ' Handle ampersands and backslashes in file paths
    PYTHONPATH = Replace(PYTHONPATH, "&", "^&")
    PYTHONPATH = Replace(PYTHONPATH, "\", "\\")
    
    If CondaPath <> "" And CondaEnv <> "" Then
        If CheckConda(CondaPath) = False Then
            Exit Function
        End If
        CondaCmd = CondaPath & "\condabin\conda activate " & CondaEnv & " && "
    Else
        CondaCmd = ""
    End If

    If IsFrozen = False Then
        Dim projRoot As String
        Dim pyBootstrap As String
        ' Bootstrap: project root = folder with hc_main.py (bridge); short-lived invoke via core.excel_session (core.ribbon_invoke).
        
        projRoot = ThisWorkbook.Path
        pyBootstrap = "import os,sys;"
        pyBootstrap = pyBootstrap & "root=r'" & Replace(projRoot, "'", "''") & "';"
        pyBootstrap = pyBootstrap & "cand=os.path.dirname(root);hc=os.path.join(root,'hc_main.py');hc2=os.path.join(cand,'hc_main.py');"
        pyBootstrap = pyBootstrap & "(not os.path.exists(hc) and os.path.exists(hc2)) and (root:=cand);"
        pyBootstrap = pyBootstrap & "(root and root not in sys.path) and sys.path.insert(0, root);"
        
        RunCommand = CondaCmd & PythonInterpreter & " -B -c """ & pyBootstrap & " " & pythonCommand & """ "
    ElseIf IsFrozen = True Then
        RunCommand = Chr(34) & pythonCommand & Chr(34) & " " & FrozenArgs & " "
    End If
    
    ExitCode = WScript.Run("cmd.exe /C " & DriveCommand & _
                       RunCommand & _
                       " --wb=" & """" & ThisWorkbook.Name & """ --from_xl=1" & " --app=" & Chr(34) & _
                       Application.Path & "\" & Application.Name & Chr(34) & " --hwnd=" & Chr(34) & Application.hwnd & Chr(34) & _
                       " 2> """ & LOG_FILE & """ ", _
                       WindowStyle, WaitOnReturn)

    'If ExitCode <> 0 then there's something wrong
    If ExitCode <> 0 Then
        Call ShowError(LOG_FILE)
        ExecuteWindows = -1
    End If

    ' Delete file after the error message has been shown
    On Error Resume Next
        Kill LOG_FILE
    On Error GoTo 0
End Function

Public Function RunFrozenPython(Executable As String, Optional Args As String)
    ' Runs a Python executable that has been frozen by PyInstaller and the like. Call the function like this:
    ' RunFrozenPython "C:\path\to\frozen_executable.exe", "arg1 arg2". Currently not implemented for Mac.

    ' Call Python
    #If Mac Then
        MsgBox "This functionality is not yet supported on Mac." & vbNewLine & _
               "Please run your scripts directly in Python!", vbCritical + vbOKOnly, "Unsupported Feature"
    #Else
        ExecuteWindows True, Executable, ParentFolder(Executable), , Args
    #End If
End Function

#If App = "Microsoft Excel" Then
Function GetUdfModules(Optional Wb As Workbook) As String
#Else
Function GetUdfModules(Optional Wb As Variant) As String
#End If
    Dim i As Integer
    Dim UDF_MODULES As String
    Dim sht As Worksheet

    GetUdfModules = GetConfig("UDF MODULES")
    ' Remove trailing ";"
    If Right$(GetUdfModules, 1) = ";" Then
        GetUdfModules = Left$(GetUdfModules, Len(GetUdfModules) - 1)
    End If
    
    ' Automatically add embedded code sheets
    For Each sht In Wb.Worksheets
        If Right$(sht.Name, 3) = ".py" Then
            If GetUdfModules = "" Then
                GetUdfModules = Left$(sht.Name, Len(sht.Name) - 3)
            Else
                GetUdfModules = GetUdfModules & ";" & Left$(sht.Name, Len(sht.Name) - 3)
            End If
        End If
    Next

    ' Default
    If GetUdfModules = "" Then
        GetUdfModules = Left$(Wb.Name, Len(Wb.Name) - 5) ' assume that it ends in .xls*
    End If
    
End Function

Private Sub CleanUp()
    'On Mac only, this function is being called after Python is done (using Python's atexit handler)
    Dim LOG_FILE As String

    #If MAC_OFFICE_VERSION >= 15 Then
        LOG_FILE = Environ("HOME") + "/xlwings.log" '~/Library/Containers/com.microsoft.Excel/Data/xlwings.log
    #Else
        LOG_FILE = "/tmp/xlwings.log"
    #End If

    'Show the LOG_FILE as MsgBox if not empty
    On Error Resume Next
    If ReadFile(LOG_FILE) <> "" Then
        Call ShowError(LOG_FILE)
    End If
    On Error GoTo 0

    'Clean up
    Application.StatusBar = False
    Application.ScreenUpdating = True
    On Error Resume Next
        #If MAC_OFFICE_VERSION >= 15 Then
            Kill LOG_FILE
        #Else
            KillFileOnMac ToMacPath(ToPosixPath(LOG_FILE))
        #End If
    On Error GoTo 0
End Sub

Function XLPyCommand()
    'TODO: the whole python vs. pythonw should be obsolete now that the console is shown/hidden by the dll
    Dim PYTHON_WIN As String, PYTHONPATH As String, LOG_FILE As String, tail As String, licenseKey As String, LicenseKeyEnvString As String, AddExcelDir As String
    Dim CondaCmd As String, CondaPath As String, CondaEnv As String, ConsoleSwitch As String, FName As String

    Dim DEBUG_UDFS As Boolean
    #If App = "Microsoft Excel" Then
    Dim Wb As Workbook
    #End If

    ' TODO: Doesn't automatically check if code is embedded
    AddExcelDir = GetConfig("ADD_WORKBOOK_TO_PYTHONPATH", "true")

    ' The first 6 args are not technically part of the PYTHONPATH, but it's just easier to add it here (used by xlwings.utils.prepare_sys_path)
    #If App = "Microsoft Excel" Then
        PYTHONPATH = AddExcelDir & ";" & ThisWorkbook.FullName & ";" & ThisWorkbook.FullName & ";" & GetConfig("ONEDRIVE_CONSUMER_WIN") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_WIN") & ";" & GetConfig("SHAREPOINT_WIN") & ";" & GetConfig("PYTHONPATH")
    #Else
        ' Other office apps
        #If App = "Microsoft Word" Then
            FName = ThisDocument.FullName
        #ElseIf App = "Microsoft Access" Then
            FName = CurrentProject.FullName
        #ElseIf App = "Microsoft PowerPoint" Then
            FName = ActivePresentation.FullName
        #End If
        PYTHONPATH = FName & ";" & ";" & GetConfig("ONEDRIVE_CONSUMER_WIN") & ";" & GetConfig("ONEDRIVE_COMMERCIAL_WIN") & ";" & GetConfig("SHAREPOINT_WIN") & ";" & GetConfig("PYTHONPATH")
    #End If

    ' Escaping backslashes and quotes
    PYTHONPATH = Replace(PYTHONPATH, "\", "\\")
    PYTHONPATH = Replace(PYTHONPATH, "'", "\'")
    PYTHONPATH = Replace(PYTHONPATH, "&", "^&")
    
    PYTHON_WIN = GetConfig("INTERPRETER_WIN", "")
    If PYTHON_WIN = "" Then
        ' Legacy
        PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")
    End If
    DEBUG_UDFS = GetConfig("DEBUG UDFS", False)

    ' /showconsole is a fictitious command line switch that's ignored by cmd.exe but used by CreateProcessA in the dll
    ' It's the only setting that's sent over like this at the moment
    If GetConfig("SHOW CONSOLE", False) = True Then
        ConsoleSwitch = "/showconsole"
    Else
        ConsoleSwitch = ""
    End If

    CondaPath = GetConfig("CONDA PATH")
    CondaEnv = GetConfig("CONDA ENV")

    If (PYTHON_WIN = "python" Or PYTHON_WIN = "pythonw") And (CondaPath <> "" And CondaEnv <> "") Then
        CondaCmd = Chr(34) & Chr(34) & CondaPath & "\condabin\conda" & Chr(34) & " activate " & CondaEnv & " && "
        PYTHON_WIN = "cmd.exe " & ConsoleSwitch & " /K " & CondaCmd & "python"
    Else
        PYTHON_WIN = "cmd.exe " & ConsoleSwitch & " /K " & Chr(34) & Chr(34) & PYTHON_WIN & Chr(34)
    End If

    licenseKey = GetConfig("LICENSE_KEY", "")
    If licenseKey <> "" Then
        LicenseKeyEnvString = "os.environ['XLWINGS_LICENSE_KEY']='" & licenseKey & "';"
    Else
        LicenseKeyEnvString = ""
    End If

    If DEBUG_UDFS = True Then
        XLPyCommand = "{506e67c3-55b5-48c3-a035-eed5deea7d6d}"
    Else
        ' Spaces in path of python.exe require quote around path AND quotes around whole command, see:
        ' https://stackoverflow.com/questions/6376113/how-do-i-use-spaces-in-the-command-prompt
        tail = " -B -c ""import sys, os;" & LicenseKeyEnvString & "import xlwings.utils;xlwings.utils.prepare_sys_path(\""" & PYTHONPATH & "\"");import xlwings; xlwings.serve('$(CLSID)')"""
        XLPyCommand = PYTHON_WIN & tail & Chr(34)
    End If
End Function

Private Sub XLPyLoadDLL()
    Dim PYTHON_WIN As String, CondaCmd As String, CondaPath As String, CondaEnv As String

    PYTHON_WIN = GetConfig("INTERPRETER_WIN", "")
    If PYTHON_WIN = "" Then
        ' Legacy
        PYTHON_WIN = GetConfig("INTERPRETER", "pythonw")
    End If
    CondaPath = GetConfig("CONDA PATH")
    CondaEnv = GetConfig("CONDA ENV")

    If (PYTHON_WIN = "python" Or PYTHON_WIN = "pythonw") And (CondaPath <> "" And CondaEnv <> "") Then
        ' This only works if the envs are in their default location
        ' Otherwise you'll have to add the full path for the interpreter in addition to the conda infos
        If CondaEnv = "base" Then
            PYTHON_WIN = CondaPath & "\" & PYTHON_WIN
        Else
            PYTHON_WIN = CondaPath & "\envs\" & CondaEnv & "\" & PYTHON_WIN
        End If
    End If

    If (PYTHON_WIN <> "python" And PYTHON_WIN <> "pythonw") Or (CondaPath <> "" And CondaEnv <> "") Then
        If LoadLibrary(ParentFolder(PYTHON_WIN) + "\" + XLPyDLLName) = 0 Then  ' Standard installation
            If LoadLibrary(ParentFolder(ParentFolder(PYTHON_WIN)) + "\" + XLPyDLLName) = 0 Then  ' Virtualenv
                Err.Raise 1, Description:= _
                    "Could not load " + XLPyDLLName + " from either of the following folders: " _
                    + vbCrLf + ParentFolder(PYTHON_WIN) _
                    + vbCrLf + ", " + ParentFolder(ParentFolder(PYTHON_WIN))
            End If
        End If
    End If
End Sub

Function NDims(ByRef src As Variant, dims As Long, Optional transpose As Boolean = False)
    XLPyLoadDLL
    If 0 <> XLPyDLLNDims(src, dims, transpose, NDims) Then Err.Raise 1001, Description:=NDims
End Function

Function XLPy()
    XLPyLoadDLL
    If 0 <> XLPyDLLActivateAuto(XLPy, XLPyCommand, 1) Then Err.Raise 1000, Description:=XLPy
End Function

Sub KillPy()
    XLPyLoadDLL
    Dim unused
    If 0 <> XLPyDLLActivateAuto(unused, XLPyCommand, -1) Then Err.Raise 1000, Description:=unused
End Sub

Sub ImportPythonUDFsBase(Optional addin As Boolean = False)
    ' This is called from the Ribbon button
    Dim tempPath As String, errorMsg As String
    Dim Wb As Workbook

    If GetConfig("CONDA PATH") <> "" And CheckConda(GetConfig("CONDA PATH")) = False Then
        Exit Sub
    End If

    If addin = True Then
        Set Wb = ThisWorkbook
    Else
        Set Wb = ThisWorkbook
    End If

    On Error GoTo ImportError
        tempPath = XLPy.Str(XLPy.Call(XLPy.Module("xlwings"), "import_udfs", XLPy.Tuple(GetUdfModules(Wb), Wb)))
    Exit Sub
ImportError:
    errorMsg = Err.Description & " " & Hex$(Err.Number)
    ShowError "", errorMsg
End Sub

Sub ImportPythonUDFs()
    ImportPythonUDFsBase
End Sub

Sub ImportPythonUDFsToAddin()
    ImportPythonUDFsBase addin:=True
End Sub

Sub ImportXlwingsUdfsModule(tf As String)
    ' Fallback: This is called from Python as direct pywin32 calls were sometimes failing, see comments in the Python code
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents("xlwings_udfs")
    On Error GoTo 0
    ThisWorkbook.VBProject.VBComponents.Import tf
End Sub

Private Sub GetDLLVersion()
    ' Currently only for testing
    Dim tag As String, arch As String
    Dim ver As Double
    XLPyDLLVersion tag, ver, arch
    Debug.Print tag
    Debug.Print ver
    Debug.Print arch
End Sub



'Attribute VB_Name = "Config"



#If App = "Microsoft Excel" Then
Function GetDirectoryPath(Optional Wb As Workbook) As String
#Else
Function GetDirectoryPath(Optional Wb As Variant) As String
#End If
    ' Leaving this here for now because we currently don't have #Const App in Utils
    Dim Path As String
    #If App = "Microsoft Excel" Then
        On Error Resume Next 'On Mac, this is called when exiting the Python interpreter
            Path = GetDirectory(GetFullName(Wb))
        On Error GoTo 0
    #ElseIf App = "Microsoft Word" Then
        Path = ThisDocument.Path
    #ElseIf App = "Microsoft Access" Then
        Path = CurrentProject.Path ' Won't be transformed for standalone module as ThisProject doesn't exit
    #ElseIf App = "Microsoft PowerPoint" Then
        Path = ActivePresentation.Path ' Won't be transformed for standalone module ThisPresentation doesn't exist
    #Else
        Exit Function
    #End If
    GetDirectoryPath = Path
End Function

Function GetConfigFilePath() As String
    #If Mac Then
        ' ~/Library/Containers/com.microsoft.Excel/Data/xlwings.conf
        GetConfigFilePath = GetMacDir("$HOME", False) & "/" & PROJECT_NAME & ".conf"
    #Else
        GetConfigFilePath = Environ("USERPROFILE") & "\." & PROJECT_NAME & "\" & PROJECT_NAME & ".conf"
    #End If
End Function

Function GetDirectoryConfigFilePath() As String
    Dim pathSeparator As String
    
    #If Mac Then ' Application.PathSeparator doesn't seem to exist in Access...
        pathSeparator = "/"
    #Else
        pathSeparator = "\"
    #End If
    
    GetDirectoryConfigFilePath = GetDirectoryPath(ThisWorkbook) & pathSeparator & PROJECT_NAME & ".conf"
End Function

#If App = "Microsoft Excel" Then
Function GetConfigFromSheet(Wb As Workbook)
    Dim lastCell As Range, cell As Range
    #If Mac Then
    Dim d As Dictionary
    Set d = New Dictionary
    #Else
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    #End If
    Dim sht As Worksheet

    Set sht = Wb.Sheets(PROJECT_NAME & ".conf")

    If sht.Range("A2") = "" Then
        Set lastCell = sht.Range("A1")
    Else
        Set lastCell = sht.Range("A1").End(xlDown)
    End If

    For Each cell In Range(sht.Range("A1"), lastCell)
        d.Add UCase(cell.Value), cell.Offset(0, 1).Value
    Next cell
    Set GetConfigFromSheet = d
End Function
#End If

Function GetConfig(configKey As String, Optional default As String = "", Optional source As String = "") As Variant
    ' ??: 2.2.0 (exezz??: UDFT[o[??ftHgL)
    
    Dim configValue As String
    Dim myConfPath As String
    Dim myExePath As String
    
    ' --- 1. xlwings.conf (??t@C) ??D???? ---
    myConfPath = ThisWorkbook.Path & "\xlwings.conf"
    
    If FileExists(myConfPath) = True Then
        If GetConfigFromFile(myConfPath, configKey, configValue) Then
            GetConfig = configValue
            
            ' ??pX (.\) ??WbN
            If Left(GetConfig, 2) = ".\" Then
                GetConfig = ThisWorkbook.Path & Mid(GetConfig, 2)
            End If
            
            ' st@C??`FbN (fobOp)
            If configKey = "INTERPRETER" Then
                If FileExists(GetConfig) = False Then
                    MsgBox "??t@C??w??????st@C????F" & vbCrLf & GetConfig, vbCritical, "HeaderTool Error"
                End If
            End If

            GetConfig = ExpandEnvironmentStrings(GetConfig)
            Exit Function
        End If
    End If

    '  ?? 1: UDFT[o[[h?? (exes??K{)
    ' ??t@C??w???????AftHg True (L) ??
    If configKey = "USE UDF SERVER" And default = "" Then
        'GetConfig = "True"
        ' fBtHg?? False
        GetConfig = "False"
        Exit Function
    End If
    
    '  ?? 2: W[??
    ' exe????APythont@C(gq??)w??????????
    If configKey = "UDF MODULES" And default = "" Then
        GetConfig = "header_converter"
        Exit Function
    End If


    ' --- 2. TWbN (INTERPRETER) ---
    ' ??t@C????AtH_ header_converter.exe ??g
    If configKey = "INTERPRETER" And default = "" Then
        myExePath = ThisWorkbook.Path & "\header_converter.exe"
        If FileExists(myExePath) Then
            GetConfig = myExePath
            Exit Function
        Else
            ' T????????
            ' (exezzO?????A??????G[o)
            ' MsgBox "st@C(header_converter.exe)????B", vbCritical, "HeaderTool Error"
        End If
    End If

    ' --- 3. W??tH[obNWbN ---
    
    ' Sheet
    #If App = "Microsoft Excel" Then
    If source = "" Or source = "sheet" Then
        If Application.Name = "Microsoft Excel" Then
            If SheetExists(ThisWorkbook, PROJECT_NAME & ".conf") = True Then
                If GetConfigFromSheet(ThisWorkbook).Exists(configKey) = True Then
                    GetConfig = GetConfigFromSheet(ThisWorkbook).Item(configKey)
                    GetConfig = ExpandEnvironmentStrings(GetConfig)
                    Exit Function
                End If
            End If
        End If
    End If
    #End If

    ' Directory Config
    If source = "" Or source = "directory" Then
        #If App = "Microsoft Excel" Then
            If GetFullName(ThisWorkbook) <> "" Then
        #Else
            If InStr(GetDirectoryPath(), "://") = 0 Then
        #End If
            If FileExists(GetDirectoryConfigFilePath()) = True Then
                If GetConfigFromFile(GetDirectoryConfigFilePath(), configKey, configValue) Then
                    GetConfig = configValue
                    GetConfig = ExpandEnvironmentStrings(GetConfig)
                    Exit Function
                End If
            End If
        End If
    End If

    ' User Config
    If source = "" Or source = "user" Then
        If FileExists(GetConfigFilePath()) = True Then
            If GetConfigFromFile(GetConfigFilePath(), configKey, configValue) Then
                GetConfig = configValue
                GetConfig = ExpandEnvironmentStrings(GetConfig)
                Exit Function
            End If
        End If
    End If

    ' Defaults
    GetConfig = default
    GetConfig = ExpandEnvironmentStrings(GetConfig)

End Function

Function SaveConfigToFile(sFileName As String, sName As String, Optional sValue As String) As Boolean
'Adopted from http://peltiertech.com/save-retrieve-information-text-files/

  Dim iFileNumA As Long, iFileNumB As Long, lErrLast As Long
  Dim sFile As String, sXFile As String, sVarName As String, sVarValue As String
      
    
  #If Mac Then
    If Not FileOrFolderExistsOnMac(ParentFolder(sFileName)) Then
  #Else
    If Len(Dir(ParentFolder(sFileName), vbDirectory)) = 0 Then
  #End If
     MkDir ParentFolder(sFileName)
  End If

  ' assume false unless variable is successfully saved
  SaveConfigToFile = False

  ' temporary file
  sFile = sFileName
  sXFile = sFileName & "_temp"

  ' open text file to read settings
  If FileExists(sFile) Then
    'replace existing settings file
    iFileNumA = FreeFile
    Open sFile For Input As iFileNumA
    iFileNumB = FreeFile
    Open sXFile For Output As iFileNumB
      Do While Not EOF(iFileNumA)
        Input #iFileNumA, sVarName, sVarValue
        If sVarName <> sName Then
          Write #iFileNumB, sVarName, sVarValue
        End If
      Loop
      Write #iFileNumB, sName, sValue
      SaveConfigToFile = True
    Close #iFileNumA
    Close #iFileNumB
    FileCopy sXFile, sFile
    Kill sXFile
  Else
    ' make new file
    iFileNumB = FreeFile
    Open sFile For Output As iFileNumB
      Write #iFileNumB, sName, sValue
      SaveConfigToFile = True
    Close #iFileNumB
  End If

End Function

Function GetConfigFromFile(sFile As String, sName As String, Optional sValue As String) As Boolean
'Based on http://peltiertech.com/save-retrieve-information-text-files/

  Dim iFileNum As Long, lErrLast As Long
  Dim sVarName As String, sVarValue As String
  Dim sLine As String, sRawValue As String
  Dim p As Long

  ' assume false unless variable is found
  GetConfigFromFile = False

  ' open text file to read settings
  If FileExists(sFile) Then
    iFileNum = FreeFile
    Open sFile For Input As iFileNum
      Do While Not EOF(iFileNum)
        Line Input #iFileNum, sLine
        sLine = Trim$(sLine)
        sVarName = ""
        sVarValue = ""

        If sLine <> "" Then
          If Left$(sLine, 1) <> "#" And Left$(sLine, 1) <> ";" Then
            If Left$(sLine, 1) = Chr$(34) Then
              p = InStr(2, sLine, Chr$(34) & "," & Chr$(34), vbBinaryCompare)
              If p > 0 And Right$(sLine, 1) = Chr$(34) Then
                sVarName = Mid$(sLine, 2, p - 2)
                sVarValue = Mid$(sLine, p + 3, Len(sLine) - p - 3)
              End If
            Else
              p = InStr(1, sLine, "=", vbBinaryCompare)
              If p > 0 Then
                sVarName = Trim$(Left$(sLine, p - 1))
                sRawValue = Trim$(Mid$(sLine, p + 1))
                If Len(sRawValue) >= 2 Then
                  If Left$(sRawValue, 1) = Chr$(34) And Right$(sRawValue, 1) = Chr$(34) Then
                    sRawValue = Mid$(sRawValue, 2, Len(sRawValue) - 2)
                  End If
                End If
                sVarValue = sRawValue
              End If
            End If
          End If
        End If

        If sVarName <> "" Then
          If LCase$(Trim$(sVarName)) = LCase$(Trim$(sName)) Then
            sValue = sVarValue
            GetConfigFromFile = True
            Exit Do
          End If
        End If
      Loop
    Close #iFileNum
  End If

End Function

'Attribute VB_Name = "Extensions"
Function sql(query, ParamArray tables())
        If TypeOf Application.Caller Is Range Then On Error GoTo failed
        ReDim argsArray(1 To UBound(tables) - LBound(tables) + 2)
        argsArray(1) = query
        For k = LBound(tables) To UBound(tables)
        argsArray(2 + k - LBound(tables)) = tables(k)
        Next k
        If has_dynamic_array() Then
            sql = XLPy.CallUDF("xlwings.ext", "sql_dynamic", argsArray, ThisWorkbook, Application.Caller)
        Else
            sql = XLPy.CallUDF("xlwings.ext", "sql", argsArray, ThisWorkbook, Application.Caller)
        End If
        Exit Function
failed:
        sql = Err.Description
End Function


'Attribute VB_Name = "Utils"



Function WScript(Optional CreateNew As Boolean) As Object
  Static Value As Object
  If CreateNew Or Value Is Nothing Then Set Value = CreateObject("WScript.Shell")
  Set WScript = Value
End Function

Function IsFullName(sFile As String) As Boolean
  ' if sFile includes path, it contains path separator "\" or "/"
  IsFullName = InStr(sFile, "\") + InStr(sFile, "/") > 0
End Function

Function FileExists(ByVal FileSpec As String) As Boolean
    #If Mac Then
        FileExists = FileOrFolderExistsOnMac(FileSpec)
    #Else
        FileExists = FileExistsOnWindows(FileSpec)
    #End If
End Function

Function FileExistsOnWindows(ByVal FileSpec As String) As Boolean
   ' by Karl Peterson MS MVP VB
   Dim Attr As Long
   ' Guard against bad FileSpec by ignoring errors
   ' retrieving its attributes.
   On Error Resume Next
   Attr = GetAttr(FileSpec)
   If Err.Number = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      FileExistsOnWindows = Not ((Attr And vbDirectory) = vbDirectory)
   End If
End Function


Function FileOrFolderExistsOnMac(FileOrFolderstr As String) As Boolean
'Ron de Bruin : 26-June-2015
'Function to test whether a file or folder exist on a Mac in office 2011 and up
'Uses AppleScript to avoid the problem with long names in Office 2011,
'limit is max 32 characters including the extension in 2011.
    Dim ScriptToCheckFileFolder As String
    Dim TestStr As String
    
    #If Mac Then
    If val(Application.VERSION) < 15 Then
        ScriptToCheckFileFolder = "tell application " & Chr(34) & "System Events" & Chr(34) & _
         "to return exists disk item (" & Chr(34) & FileOrFolderstr & Chr(34) & " as string)"
        FileOrFolderExistsOnMac = MacScript(ScriptToCheckFileFolder)
    Else
        On Error Resume Next
        TestStr = Dir(FileOrFolderstr, vbDirectory)
        On Error GoTo 0
        If Not TestStr = vbNullString Then FileOrFolderExistsOnMac = True
    End If
    #End If
End Function

Function ParentFolder(ByVal Folder)
  #If Mac Then
      ParentFolder = Left$(Folder, InStrRev(Folder, "/") - 1)
  #Else
      ParentFolder = Left$(Folder, InStrRev(Folder, "\") - 1)
  #End If
End Function

Function GetDirectory(Path)
    #If Mac Then
    GetDirectory = Left(Path, InStrRev(Path, "/"))
    #Else
    GetDirectory = Left(Path, InStrRev(Path, "\"))
    #End If
End Function

Function KillFileOnMac(Filestr As String)
    'Ron de Bruin
    '30-July-2012
    'Delete files from a Mac.
    'Uses AppleScript to avoid the problem with long file names (on 2011 only)

    Dim ScriptToKillFile As String
    
    #If Mac Then
    ScriptToKillFile = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
    ScriptToKillFile = ScriptToKillFile & "do shell script ""rm "" & quoted form of posix path of " & Chr(34) & Filestr & Chr(34) & Chr(13)
    ScriptToKillFile = ScriptToKillFile & "end tell"

    On Error Resume Next
        MacScript (ScriptToKillFile)
    On Error GoTo 0
    #End If
End Function

Function ToMacPath(PosixPath As String) As String
    ' This function transforms a Posix Path into a MacOS Path
    ' E.g. "/Users/<User>" --> "MacintoshHD:Users:<User>"
    #If Mac Then
    ToMacPath = MacScript("set mac_path to POSIX file " & Chr(34) & PosixPath & Chr(34) & " as string")
    #End If
End Function

Function GetMacDir(Name As String, Normalize As Boolean) As String
    #If Mac Then
        Select Case Name
            Case "$HOME"
                Name = "home folder"
            Case "$APPLICATIONS"
                Name = "applications folder"
            Case "$DOCUMENTS"
                Name = "documents folder"
            Case "$DOWNLOADS"
                Name = "downloads folder"
            Case "$DESKTOP"
                Name = "desktop folder"
            Case "$TMPDIR"
                Name = "temporary items"
        End Select
        GetMacDir = MacScript("return POSIX path of (path to " & Name & ") as string")
        If Normalize = True Then
            'Normalize Excel sandbox location
            GetMacDir = Replace(GetMacDir, "/Library/Containers/com.microsoft.Excel/Data", "")
        End If
    #Else
    #End If
End Function


Function ToPosixPath(ByVal MacPath As String) As String
    'This function accepts relative paths with backward and forward slashes: ThisWorkbook & "\test"
    ' E.g. "MacintoshHD:Users:<User>" --> "/Users/<User>"

    Dim s As String
    Dim LeadingSlash As Boolean
    
    #If Mac Then
    If MacPath = "" Then
        ToPosixPath = ""
    Else
        ToPosixPath = Replace(MacPath, "\", "/")
        ToPosixPath = MacScript("return POSIX path of (" & Chr(34) & MacPath & Chr(34) & ") as string")
    End If
    #End If
End Function

Sub ShowError(FileName As String, Optional Message As String = "")
    ' Shows a MsgBox with the content of a text file

    Dim Content As String
    Dim ErrorSheet As Worksheet

    Const OK_BUTTON_ERROR = 16
    Const AUTO_DISMISS = 0
    
    If Message = "" Then
        Content = ReadFile(FileName)
    Else
        Content = Message
    End If
    

    If GetConfig("SHOW_ERROR_POPUPS", "True") = "False" Then
        If SheetExists(ThisWorkbook, "Error") = False Then
            Set ErrorSheet = ThisWorkbook.Sheets.Add()
            ErrorSheet.Name = "Error"
        Else
            Set ErrorSheet = ThisWorkbook.Sheets("Error")
        End If
        ErrorSheet.Range("A1").Value = Content
    Else
        #If Mac Then
            MsgBox Content, vbCritical, "Error"
        #Else
            Content = Content & vbCrLf
            Content = Content & "Press Ctrl+C to copy this message to the clipboard."
    
            WScript.Popup Content, AUTO_DISMISS, "Error", OK_BUTTON_ERROR
        #End If
    End If
End Sub

Function ExpandEnvironmentStrings(ByVal s As String)
    ' Expand environment variables
    Dim EnvString As String
    Dim PathParts As Variant
    Dim i As Integer
    #If Mac Then
        If Left(s, 1) = "$" Then
            PathParts = Split(s, "/")
            EnvString = PathParts(0)
            ExpandEnvironmentStrings = GetMacDir(EnvString, True)
            For i = 1 To UBound(PathParts)
                If Right$(ExpandEnvironmentStrings, 1) = "/" Then
                    ExpandEnvironmentStrings = ExpandEnvironmentStrings & PathParts(i)
                Else
                    ExpandEnvironmentStrings = ExpandEnvironmentStrings & "/" & PathParts(i)
                End If
            Next i
        Else
            ExpandEnvironmentStrings = s
        End If
    #Else
        ExpandEnvironmentStrings = WScript.ExpandEnvironmentStrings(s)
    #End If
End Function

Function ReadFile(ByVal FileName As String)
    ' Read a text file

    Dim Content As String
    Dim Token As String
    Dim FileNum As Integer
    Dim objShell As Object
    Dim LineBreak As Variant

    #If Mac Then
        FileName = ToMacPath(FileName)
        LineBreak = vbLf
    #Else
        FileName = ExpandEnvironmentStrings(FileName)
        LineBreak = vbCrLf
    #End If

    FileNum = FreeFile
    Content = ""

    ' Read Text File
    Open FileName For Input As #FileNum
        Do While Not EOF(FileNum)
            Line Input #FileNum, Token
            Content = Content & Token & LineBreak
        Loop
    Close #FileNum

    ReadFile = Content
End Function

#If App = "Microsoft Excel" Then
Function SheetExists(Wb As Workbook, sheetName As String) As Boolean
    Dim sht As Worksheet
    On Error Resume Next
        Set sht = Wb.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not sht Is Nothing
End Function
#End If

Function GetBaseName(Wb As String) As String
    Dim extension As String
    extension = LCase$(Right$(Wb, 4))
    If extension = ".xls" Or extension = ".xla" Or extension = ".xlt" Then
        GetBaseName = Left$(Wb, Len(Wb) - 4)
    Else
        GetBaseName = Left$(Wb, Len(Wb) - 5)
    End If
End Function

Function has_dynamic_array() As Boolean
    has_dynamic_array = False
    On Error GoTo ErrHandler
        Application.WorksheetFunction.Unique ("dummy")
        has_dynamic_array = True
    Exit Function
ErrHandler:
    has_dynamic_array = False
End Function

Public Function CreateGUID() As String
    Randomize Timer() + Application.hwnd
    ' https://stackoverflow.com/a/46474125/918626
    Do While Len(CreateGUID) < 32
        If Len(CreateGUID) = 16 Then
            '17th character holds version information
            CreateGUID = CreateGUID & Hex$(8 + CInt(Rnd * 3))
        End If
        CreateGUID = CreateGUID & Hex$(CInt(Rnd * 15))
    Loop
    CreateGUID = Mid(CreateGUID, 1, 8) & "-" & Mid(CreateGUID, 9, 4) & "-" & Mid(CreateGUID, 13, 4) & "-" & Mid(CreateGUID, 17, 4) & "-" & Mid(CreateGUID, 21, 12)
End Function

Function CheckConda(CondaPath As String) As Boolean
    ' Check if the conda executable exists.
    ' If it doesn't, conda is too old and the Interpreter setting has to be used instead of Conda settings
    Dim condaExecutable As String
    Dim condaExists As Boolean
    #If Mac Then
        condaExecutable = CondaPath & "\condabin\conda"
    #Else
        condaExecutable = CondaPath & "\condabin\conda.bat"
    #End If
    ' Replace space escape character ^ to check if path exists
    condaExists = FileExists(Replace(condaExecutable, "^", ""))
    If condaExists = False And CondaPath <> "" Then
        MsgBox "Your Conda version seems to be too old for the Conda settings. Use the Interpreter setting instead."
    End If
    CheckConda = condaExists
End Function

#If App = "Microsoft Excel" Then
Function GetFullName(Wb As Workbook) As String
    ' The only case where this is still used is for directory-based config files, otherwise this is now handled in Python
    ' Unlike the Python version, this doesn't work for SharePoint and will just ignore a directory-based config file silently

    Dim total_found, i_parsing, i_env_var, slash_number As Integer
    Dim found_path, one_drive_path, full_path_name, this_found_path As String

    ' In the majority of cases, ThisWorkbook.FullName will provide the path of the
    ' Excel workbook correctly. Unfortunately, when the user is using OneDrive
    ' this doesn't work. This function will attempt to find the LOCAL path.
    ' This uses code from Daniel Guetta and
    ' https://stackoverflow.com/questions/33734706/excels-fullname-property-with-onedrive
    
    If InStr(Wb.FullName, "://") = 0 Or Wb.Path = "" Then
        GetFullName = Wb.FullName
        Exit Function
    End If
        
    ' According to the link above, there are three possible environment variables
    ' the user's OneDrive folder could be located in
    '      "OneDriveCommercial", "OneDriveConsumer", "OneDrive"
    '
    ' Furthermore, there are two possible formats for OneDrive URLs
    '    1. "https://companyName-my.sharepoint.com/personal/userName_domain_com/Documents" & file.FullName
    '    2. "https://d.docs.live.net/d7bbaa#######1/" & file.FullName
    ' In the first case, we can find the true path by just looking for everything after /Documents. In the
    ' second, we need to look for the fourth slash in the URL
    '
    ' The code below will try every combination of the three environment variables above, and
    ' each of the two methods of parsing the URL. The file is found in *exactly* one of those
    ' locations, then we're good to go.
    '
    ' Note that this still leaves a gap - if this file (file A) is in a location that is NOT covered by the
    ' eventualities above AND a file of the exact same name (file B) exists in one of the locations that is
    ' covered above, then this function will identify File B's location as the location of this workbook,
    ' which would be wrong
    total_found = 0
    
    For i_parsing = 1 To 2
        If i_parsing = 1 Then
            ' Parse using method 1 above; find /Documents and take everything after, INCLUDING the
            ' leading slash
            If InStr(1, Wb.FullName, "/Documents") Then
                full_path_name = Mid(Wb.FullName, InStr(1, Wb.FullName, "/Documents") + Len("/Documents"))
            Else
                full_path_name = ""
            End If
        Else
            ' Parse using method 2; find everything after the fourth slash, including that fourth
            ' slash
            Dim i_pos As Integer
            
            ' Start at the last slash in https://
            i_pos = 8

            For slash_number = 1 To 2
                i_pos = InStr(i_pos + 1, Wb.FullName, "/")
            Next slash_number
            
            full_path_name = Mid(Wb.FullName, i_pos)
        End If
        
        ' Replace forward slahes with backslashes on Windows
        full_path_name = Replace(full_path_name, "/", Application.pathSeparator)
        
        
        If full_path_name <> "" Then
            #If Not Mac Then
            For i_env_var = 1 To 3
                    one_drive_path = Environ(Choose(i_env_var, "OneDriveCommercial", "OneDriveConsumer", "OneDrive"))
                
                    If (one_drive_path <> "") And FileExists(one_drive_path & full_path_name) Then
                        this_found_path = one_drive_path & full_path_name
                        
                        If this_found_path <> found_path Then
                            total_found = total_found + 1
                            found_path = this_found_path
                        End If
                    End If
            Next i_env_var
            #End If
        End If
    Next i_parsing
        
    If total_found = 1 Then
        GetFullName = found_path
        Exit Function
    End If

End Function
#End If
