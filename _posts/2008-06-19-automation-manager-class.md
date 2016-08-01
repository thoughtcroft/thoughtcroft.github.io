---
layout: post
title:  "Automation Manager Class"
date:   2008-06-19
tags:   [access, automation, excel, vba]
---

I needed a way to manage calls to different office automation servers in
a consistent fashion. These were mostly from Access to extract data from
a large number of Excel workbooks. Specifically what I wanted was a way
of managing:

* reuse of any existing instance of Excel or start an instance if there
wasn't one
* save the state of the application - things like the calculation mode
etc and restore them when finished
* work out whether to close the instance when finished (if we started
it) or leave it (if we didn't)
* handle the strange automation errors that can occur and ensure that
the instance is properly terminated in the case where an unrecoverable
error has occurred

Further, I wanted to be able to use this for multiple automation
clients. The following class modules have served me well for this
purpose and I offer them here for those that may have a similar
requirement. Some of this is not for the faint hearted, so send me an
email if you need further explanation.

Typical calling method is as follows:

```vb
' Give this global scope
Public asm As AppStateMgr
Public app as Object

Set asm = New AppStateMgr
'
' Do something and decide we need to open Excel, say
'
Call asm.OpenApplication("Excel",app)
'
' Now open a workbook say and check for errors
'
If asm.CheckApplicationError(app,Err.Number) Then
   ' Something bad happened so deal with it
   ' If the error was catastrophic to Excel
   ' a new instance will be started anyway
Else
   ' Everything is fine, so something else
End If
'
' Finishing up now
'
Call asm.CloseApplication(app)
Set asm = Nothing
```

Create a new Class Module called AppState and copy the following code into it. This describes all the properties associated with an instance of an application that has been started to provide automation services.

```vb
' Class for defining an application state which can be
' saved and restored using the AppStateMgr class

' Developed by Warren Bain on 16/11/2006
' Copyright (c) Thought Croft Pty Ltd
' http:\\www.thoughtcroft.com
' All rights reserved.

Option Explicit

' Pointer calculated from the object used to index the state collection
' and the name of the application object this relates to
Private mstrIndex As String
Private mstrAppName As String

' Was the application instance created by us and how many are using it
Private mblnSelfStarted As Boolean
Private mlngObjectCount As Long

' Common application properties
Private mblnDisplayAlerts As Boolean
Private mblnScreenUpdating As Boolean
Private mblnVisible As Boolean
Private mlnghWnd As Long

' This one is only with Excel
Private meCalculation As Variant

Friend Property Get ObjectCount() As Long
    ObjectCount = mlngObjectCount
End Property

Friend Property Get DisplayAlerts() As Boolean
    DisplayAlerts = mblnDisplayAlerts
End Property

Friend Property Let DisplayAlerts(ByVal blnDisplayAlerts As Boolean)
    mblnDisplayAlerts = blnDisplayAlerts
End Property

Friend Property Get SelfStarted() As Boolean
    SelfStarted = mblnSelfStarted
End Property

Friend Property Let SelfStarted(ByVal blnSelfStarted As Boolean)
    mblnSelfStarted = blnSelfStarted
End Property

Friend Property Get ScreenUpdating() As Boolean
    ScreenUpdating = mblnScreenUpdating
End Property

Friend Property Let ScreenUpdating(ByVal blnScreenUpdating As Boolean)
    mblnScreenUpdating = blnScreenUpdating
End Property

Friend Property Get Visible() As Boolean
    Visible = mblnVisible
End Property

Friend Property Let Visible(ByVal blnVisible As Boolean)
    mblnVisible = blnVisible
End Property

Friend Property Get Index() As String
    Index = mstrIndex
End Property

Friend Property Let Index(ByVal strIndex As String)
    ' Can only assign this if the value is empty
    ' i.e. after it has been set, it is read only!
    If Len(mstrIndex) = 0 Then
        mstrIndex = strIndex
    Else
        Err.Raise vbObjectError + 5, "AppState", _
                  "Can't alter the Index after created!"
    End If
End Property

Friend Function IncrementCount() As Long
    ' Increase the count of objects using this application
    mlngObjectCount = mlngObjectCount + 1
    IncrementCount = mlngObjectCount
End Function

Friend Function DecrementCount() As Long
    ' Decrease the count of objects using this application
    mlngObjectCount = mlngObjectCount - 1
    DecrementCount = mlngObjectCount
End Function

Friend Property Get AppName() As String
    AppName = mstrAppName
End Property

Friend Property Let AppName(ByVal strAppName As String)
    ' Can only assign this if the value is empty
    ' i.e. after it has been set, it is read only!
    If Len(mstrAppName) = 0 Then
        mstrAppName = strAppName
    Else
        Err.Raise vbObjectError + 5, "AppState", _
                  "Can't alter the AppName after created!"
    End If
End Property

Friend Property Get Calculation() As Variant
    Calculation = meCalculation
End Property

Friend Property Let Calculation(ByVal eCalculation As Variant)
    meCalculation = eCalculation
End Property

Friend Property Get WindowsHandle() As Long
    WindowsHandle = mlnghWnd
End Property

Friend Property Let WindowsHandle(ByVal lngWindowsHandle As Long)
    mlnghWnd = lngWindowsHandle
End Property
```

Create a Class Module called AppStateMgr and copy the following code into it. This provides the functions for managing instances of automation clients.

```vb
' Manage the state of automation application objects and associated
' functions for saving, restoring the application state as well as
' handling typical automation errors, etc

' Developed by Warren Bain on 16/11/2006
' Copyright (c) Thought Croft Pty Ltd
' http:\\www.thoughtcroft.com
' All rights reserved.

Option Explicit

Private Const PROCESS_TERMINATE As Long = (&H1)
Private Const SW_SHOWNORMAL = 1

Private Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" ( _
                                       ByVal strClass As String, _
                                       ByVal lpWindow As String) As Long

Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" ( _
                                        ByVal hwnd As Long, _
                                        ByVal Msg As Long, _
                                        ByVal wParam As Long, _
                                        ByVal lParam As Long) As Long

Private Declare Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" ( _
                                                ByVal hwnd As Long) As Long

Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" ( _
                                       ByVal hwnd As Long, _
                                       ByVal nCmdShow As Long) As Long

Private Declare Function apiIsIconic Lib "user32" Alias "IsIconic" ( _
                                     ByVal hwnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
                               ByRef Destination As Any, _
                               ByRef Source As Any, _
                               ByVal Length As Long)

Private Declare Function CloseHandle Lib "kernel32.dll" ( _
                                     ByVal hObject As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32.dll" ( _
                                                  ByVal hwnd As Long, _
                                                  ByRef lpdwProcessId As Long) As Long

Private Declare Function OpenProcess Lib "kernel32.dll" ( _
                                     ByVal dwDesiredAccess As Long, _
                                     ByVal bInheritHandle As Long, _
                                     ByVal dwProcessId As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32.dll" ( _
                                          ByVal hProcess As Long, _
                                          ByVal uExitCode As Long) As Long


' Collection for holding the application states
Private mcolAppStateCol       As Collection

Friend Property Get AppStateCol() As Collection
    ' Return collection object, create if necessary
    If mcolAppStateCol Is Nothing Then
        Set mcolAppStateCol = New Collection
    End If
    Set AppStateCol = mcolAppStateCol
End Property

Public Function CheckApplicationError( _
       ByRef objApp As Object, _
       ByVal lngErrNumber As Long) As Boolean

    ' Check for a range of automation errors that can occur
    ' and try and recover from them, typically by restarting
    ' the application server. Returns True if the App
    ' had to be recovered due to Automation errors

    ' Normally, lngErrNumber will contain the value from
    ' Err.Number after a call to an automation client related
    ' object but can be passed a negative number to force a restart

    Const conSpecificAutomationErrorEndRange = 0
    Const conObjectRequiredError = 424
    Const conGeneralAutomationError = 440
    Const conAutomationNotSupportedError = 458
    Const conRemoteServerLostError = 462
    Const conApplicationDefinedError = 1004

    Dim objState              As AppState

    If objApp Is Nothing Then
        ' Must have application to work with here
        Err.Raise vbObjectError + 5, "AppStateMgr::CheckApplication", _
                  "Must supply a valid App object"
    Else
        Select Case lngErrNumber
        Case Is < conSpecificAutomationErrorEndRange, _
             conObjectRequiredError, _
             conGeneralAutomationError, _
             conAutomationNotSupportedError, _
             conRemoteServerLostError, _
             conApplicationDefinedError
            ' Definitely an unrecoverable automation error
            ' so force a close, getting the state so we can
            ' determine the app type and force a new instance
            Set objState = CloseApplication(objApp, True)
            Call OpenApplication(objState.AppName, objApp, True)
            Set objState = Nothing
            CheckApplicationError = True
        End Select
    End If
End Function

Private Function CheckAppRunning( _
        ByVal strAppName As String, _
        Optional ByVal blnActivate As Boolean) As Boolean

    ' This code was originally written by Dev Ashish
    ' but has been enhanced to cope with other classes
    ' of application by Warren Bain

    Const WM_USER = 1024

    Dim lngH                  As Long
    Dim strClassName          As String
    Dim lngX                  As Long
    Dim lngTmp                As Long

    On Local Error GoTo HandleErrors

    CheckAppRunning = False
    strClassName = GetClassName(strAppName)
    If Len(strClassName) = 0 Then
        lngH = apiFindWindow(vbNullString, strAppName)
    Else
        lngH = apiFindWindow(strClassName, vbNullString)
    End If
    If lngH <> 0 Then
        apiSendMessage lngH, WM_USER + 18, 0, 0
        lngX = apiIsIconic(lngH)
        If lngX <> 0 Then
            lngTmp = apiShowWindow(lngH, SW_SHOWNORMAL)
        End If
        If blnActivate Then
            lngTmp = apiSetForegroundWindow(lngH)
        End If
        CheckAppRunning = True
    End If

ExitHere:
    Exit Function
HandleErrors:
    CheckAppRunning = False
    Resume ExitHere
End Function

Private Sub Class_Terminate()

    ' Destroy the collection object if there are
    ' still members that we haven't terminated before hand

    If Not mcolAppStateCol Is Nothing Then
        Set mcolAppStateCol = Nothing
    End If

End Sub

Public Sub CloseAllApplications( _
       Optional ByVal blnForceKill As Boolean = False)

    ' Walk the collection and call CloseApplication for
    ' each instance we have available to us.

    ' ***************** WARNING ************************
    ' Use extreme care calling as this will terminate the
    ' application without cleaning up any objects still
    ' pointing at it.  May cause host app to crash on exit
    ' **************************************************

    Dim objState              As AppState
    Dim i                     As Integer

    If Not Me.AppStateCol Is Nothing Then
        For i = Me.AppStateCol.Count To 1 Step -1
            ' Create an object for each
            ' member that was created
            ' using GetObjectFromPtr
            Set objState = Me.AppStateCol(i)
            Call CloseApplication(GetObjectFromPtr(objState.Index), _
                                  blnForceKill, _
                                  objState)
        Next i
    End If

End Sub

Public Function CloseApplication( _
       ByRef objApp As Object, _
       Optional ByVal blnForceKill As Boolean = False, _
       Optional ByRef objState As AppState) As AppState

    ' Restore application state and leave it running if it
    ' was not started by us (unless they want us to kill it)
    ' or if there are others using it

    Dim blnShutDown           As Boolean
    Dim blnLastState          As Boolean

    If Not objApp Is Nothing Then
        ' Find the saved status of the application
        ' and work out if we started it - if no
        ' saved state, then assume we didn't start it
        If objState Is Nothing Then
            ' They didn't supply it so, find it
            ' Don't worry about trapping errors
            On Error Resume Next
            Set objState = FindAppState(objApp)
            On Error GoTo 0
        End If
        If objState Is Nothing Then
            ' No state saved so action depends on ForceKill value
            blnShutDown = blnForceKill
        Else
            ' Decrement count of objects to decide if we should
            ' really shut this one down as well
            With objState
                blnLastState = (.DecrementCount <= 0)
                blnShutDown = blnForceKill Or _
                              (.SelfStarted And blnLastState)
            End With
        End If

        ' Did we start it and no-one else using it
        ' or do they want to kill it anyway?
        If blnShutDown Then
            ' Don't bother restoring, just terminate the application
            Call TerminateApplication(objApp, objState)
        ElseIf blnLastState Then
            ' Attempt to restore the application state
            ' as we have finished with it in this process
            ' and remove it from the list
            Call RestoreAppState(objApp, objState)
        End If
        Call RemoveAppState(objState)
        Set CloseApplication = objState
        ' Force release of application object to ensure
        ' application will shutdown normally
        Set objApp = Nothing
    End If
End Function

Private Function FindAppState( _
        ByRef objApp As Object, _
        Optional ByVal strName As String = vbNullString, _
        Optional ByVal blnCreateNew As Boolean = False) As AppState

    ' Retrieve an existing AppSave object or
    ' if not available then add a new one if
    ' caller requests us to CreateNew

    Dim objState              As AppState

    If objApp Is Nothing Then
        ' Can't do this if there isn't any object
             Err.Raise vbObjectError + 5, "AppStateMgr::FindAppState", _
                      "Must supply an instantiated object (not Nothing)!"
    Else
        ' Check compatible parameters
        If Len(strName) = 0 And blnCreateNew Then
            ' Can't create if we don't know the AppName
            Err.Raise vbObjectError + 13, "AppStateMgr::FindAppState", _
                      "Can't specify 'CreateNew' without supplying 'Name'!"
        Else
            On Error Resume Next
            Set objState = Me.AppStateCol(GetObjPtr(objApp))
            On Error GoTo 0
            If objState Is Nothing Then
                If blnCreateNew Then
                    Set objState = SaveAppState(objApp, strName)
                Else
                    Err.Raise vbObjectError + 63, "AppStateMgr::FindAppState", _
                              "Can't find required 'AppState' for this object!"
                End If
            ElseIf objState.AppName <> strName And Len(strName) > 0 Then
                ' Found existing one but it doesn't match
                ' the AppName we are expecting - whoops!
                Err.Raise vbObjectError + 13, "AppStateMgr::FindAppState", _
                          "Conflict with 'AppName' supplied [" & strName & _
                          "] and retrieved [" & objState.AppName & "]!"
            End If
            Set FindAppState = objState
        End If
    End If
End Function

Private Function GeneratePassword( _
       ByVal intLength As Integer) As String

    ' Generates a random string of digits of the requested length

    Dim lngHighNumber         As Long
    Dim lngLowNumber          As Long
    Dim lngRndNumber          As Long

    ' Check we don't exceed our maximum range
    If intLength > 9 Or intLength < 1 Then
        Err.Raise 5, "GetPassword", _
                  "Invalid string length - must be between 1 and 9"
    Else
        ' Work out the numbers
        lngLowNumber = 10 ^ (intLength - 1)
        lngHighNumber = (10 ^ intLength) - 1
        ' Generate a new seed and a new random number
        Randomize
        lngRndNumber = Int((lngHighNumber - lngLowNumber + 1) * Rnd) + lngLowNumber
        ' Format the result as string
        GeneratePassword = Format$(lngRndNumber, String$(intLength, "0"))
    End If
End Function

Private Function GetApplicationHandle(ByRef objApp As Object) As Long

    ' Locate the windows handle for the application
    ' represented by this object

    Dim hwnd                  As Long
    Dim varCaption            As Variant

    On Error Resume Next

    ' Determine the type of object - can make it easy
    ' as the object may store it itself
    If TypeOf objApp Is Access.Application Then
        hwnd = objApp.hWndAccessApp
    ElseIf TypeOf objApp Is Excel.Application Then
        ' This only works for Excel 2002 onwards
        hwnd = objApp.hwnd
    End If
    On Error GoTo 0

    If hwnd = 0 Then
        ' Need to discover it from the Window so we make sure
        ' that the caption is unique for this application
        varCaption = objApp.Caption
        objApp.Caption = GeneratePassword(8)
        hwnd = apiFindWindow(GetClassName(objApp.Name), objApp.Caption)
        objApp.Caption = varCaption
    End If
    GetApplicationHandle = hwnd
End Function

Private Function GetClassName(ByVal strAppName As String) As String

    ' Returns the Class Name for the main window of various
    ' Microsoft software applications

    Select Case LCase$(strAppName)
    Case "excel", "microsoft excel": GetClassName = "XLMain"
    Case "word", "microsoft word": GetClassName = "OpusApp"
    Case "access", "microsoft access": GetClassName = "OMain"
    Case "powerpoint95": GetClassName = "PP7FrameClass"
    Case "powerpoint97": GetClassName = "PP97FrameClass"
    Case "powerpoint2000": GetClassName = "PP9FrameClass"
    Case "powerpoint2002": GetClassName = "PP10FrameClass"
    Case "powerpoint2003": GetClassName = "PP11FrameClass"
    Case "powerpoint2007": GetClassName = "JWinproj-WhimperMainClass"
    Case "project", "microsoft project": GetClassName = "PP9FrameClass"
    Case "notepad": GetClassName = "NOTEPAD"
    Case "paintbrush": GetClassName = "pbParent"
    Case "wordpad": GetClassName = "WordPadClass"
    Case Else: GetClassName = vbNullString
    End Select
End Function

Private Function GetObjectFromPtr(ByVal lPtr As Long) As Object

    ' Based on Bruce McKinney's code for getting an Object from the
    ' object pointer - the reverse of ObjPtr(object).

    Dim objT                  As Object

    On Error GoTo HandleError

    CopyMemory objT, lPtr, 4
    Set GetObjectFromPtr = objT
    Exit Function

HandleError:
    With Err
        .Raise .Number, "AppStateMgr::GetObjectFromPtr" & .Source, _
               .Description, .HelpFile, .HelpContext
    End With
End Function

Private Function GetObjPtr(ByRef obj As Object) As String

    ' This relies on undocumented function to return
    ' the address of the object pointer in memory
    ' which is useful for fast indexing into a collection
    ' of objects or object related data. Returns null string
    ' if object hasn't been assigned yet

    If obj Is Nothing Then
        GetObjPtr = vbNullString
    Else
        GetObjPtr = CStr(ObjPtr(obj))
    End If
End Function

Public Sub OpenApplication( _
       ByVal strAppName As String, _
       ByRef objApp As Object, _
       Optional ByVal blnForceNewInstance As Boolean = False, _
       Optional ByVal blnDisplayAlerts As Boolean = False)

    ' Check if object is already referencing an application and
    ' if not then first try and use existing automation client if
    ' running else start a new one and save its state. The caller
    ' can force us to create a new instance if they wish although
    ' this also depends on the application which may only single instance

    Dim objState              As AppState
    Dim blnSelfStarted        As Boolean

    On Error GoTo HandleErrors

    If objApp Is Nothing Then
        ' Try and locate an existing instance first before starting new one
        If CheckAppRunning(strAppName) And Not blnForceNewInstance Then
            ' Server already running so return reference to it
            Set objApp = GetObject(, strAppName & ".Application")
            blnSelfStarted = False
        Else
            ' Need to start a new instance of required server application
            Set objApp = CreateObject(strAppName & ".Application")
            blnSelfStarted = True
        End If

        ' Now find the state - if it doesn't exist then it will create a
        ' new one and save app state.
        Set objState = FindAppState(objApp, strAppName, True)

        ' Increment the counter and set the DisplayAlert property
        objState.IncrementCount
        objApp.DisplayAlerts = blnDisplayAlerts
        
        ' Save whether we started it but don't update it
        ' if we didn't because may have been done by previous
        ' call using a different object variable
        If blnSelfStarted Then
            objState.SelfStarted = blnSelfStarted
        End If
    End If

ExitHere:
    Exit Sub

HandleErrors:
    With Err
        Select Case .Number
        Case Else
            .Raise .Number, "AppStateMgr::OpenApplication", .Description, .HelpFile, .HelpContext
        End Select
    End With
    Resume ExitHere
End Sub

Private Function RemoveAppState(ByRef objState As AppState)

    ' To remove the supplied AppState object from the
    ' collection - no longer required

    If Not objState Is Nothing Then
        Me.AppStateCol.Remove objState.Index
    End If
End Function

Private Function RestoreAppState( _
        ByRef objApp As Object, _
        Optional ByRef objState As AppState = Nothing) As AppState

    ' To find existing AppState and restore the state
    ' of the supplied application object

    ' If application is already nothing then exit
    If Not objApp Is Nothing Then
        If objState Is Nothing Then
            ' No AppState supplied so go find it -
            ' note that this call will raise an error
            ' if the AppState can't be found
            Set objState = FindAppState(objApp)
        End If

        ' We will have a valid AppState now
        With objApp

            ' *************************************
            ' These properties apply to all objects
            .DisplayAlerts = objState.DisplayAlerts
            .ScreenUpdating = objState.ScreenUpdating

            ' Can only reset this if we started it
            If Not .UserControl Then
                .Visible = objState.Visible
            End If

            ' -------------------------------------
            ' Properties specific to Excel
            If TypeOf objApp Is Excel.Application Then
                ' Can only reset this if we started it
                If Not .UserControl Then
                    .Calculation = objState.Calculation
                End If
            End If

        End With
        Set RestoreAppState = objState
    End If
End Function

Private Function SaveAppState( _
        ByRef objApp As Object, _
        ByVal strName As String) As AppState

    ' To create a new AppState and save the state
    ' of the supplied application object

    Dim objState              As AppState

    If Not objApp Is Nothing Then
        ' Create a new instance and save key state info
        Set objState = New AppState
        With objState

            ' *************************************
            ' These properties apply to all objects
            .Index = GetObjPtr(objApp)
            .AppName = strName
            .WindowsHandle = GetApplicationHandle(objApp)
            .DisplayAlerts = objApp.DisplayAlerts
            .Visible = objApp.Visible
            .ScreenUpdating = objApp.ScreenUpdating

            ' -------------------------------------
            ' Properties specific to Excel
            If TypeOf objApp Is Excel.Application Then
                .Calculation = objApp.Calculation
            End If

            ' Now add to the collection
            Me.AppStateCol.Add objState, .Index
        End With
        Set SaveAppState = objState
    End If
End Function

Private Sub TerminateApplication( _
        ByRef objApp As Object, _
        ByRef objState As AppState)

    ' This will try and exit the application and
    ' also terminate the process where the
    ' automation server is not responding

    Dim hWndApp               As Long
    Dim hProcessID            As Long
    Dim hThreadID             As Long
    Dim hTerminateID          As Long


    On Error Resume Next
    If Not objApp Is Nothing Then
        If objState Is Nothing Then
            hWndApp = GetApplicationHandle(objApp)
        Else
            hWndApp = objState.WindowsHandle
        End If

        ' Now close the application normally - the Quit just
        ' allows the application to close its objects but it
        ' doesn't actually terminate until we close the object
        objApp.Quit
        Set objApp = Nothing
        DoEvents

        If hWndApp <> 0 Then
            ' Find the processid of the selected window in case it didn't
            ' close normally in which case we will get an id back
            hThreadID = GetWindowThreadProcessId(hWndApp, hProcessID)
            If hProcessID <> 0 Then
                ' Acquire a handle with terminate ability and try and kill it
                ' don't worry about failing as there is nothing we can do anyway
                hTerminateID = OpenProcess(PROCESS_TERMINATE, 0, hProcessID)
                Call TerminateProcess(hTerminateID, 0)
                Call CloseHandle(hTerminateID)
            End If
        End If
    End If
End Sub
```
