Attribute VB_Name = "SlideTimer"
Option Explicit

' Slide Timer 1.2.4 | PowerPoint for Mac, VBA7 / 64-bit
' Import SlideTimer.bas AND SlideTimerEvents.cls, then save as .pptm.
' Run Timer_Insert in Normal view. Start the show normally and click play/pause.
' No network, AppleScript, ActiveX, or third-party dependencies at runtime.
' End sound: macOS afplay runs in a separate process; no Objective-C bridge.

#If Mac Then
    Private Declare PtrSafe Function PT_MonotonicTime _
        Lib "/System/Library/Frameworks/QuartzCore.framework/QuartzCore" _
        Alias "CACurrentMediaTime" () As Double
    Private Declare PtrSafe Function PT_USleep Lib "/usr/lib/libSystem.B.dylib" _
        Alias "usleep" (ByVal microseconds As Long) As Long
    Private Declare PtrSafe Sub PT_CFRelease _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFRelease" (ByVal reference As LongPtr)
    Private Declare PtrSafe Function PT_CFAbsoluteTime _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFAbsoluteTimeGetCurrent" () As Double
    Private Declare PtrSafe Function PT_MainRunLoop _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFRunLoopGetMain" () As LongPtr
    Private Declare PtrSafe Function PT_CreateRunLoopTimer _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFRunLoopTimerCreate" (ByVal allocator As LongPtr, ByVal fireDate As Double, _
        ByVal interval As Double, ByVal flags As LongPtr, ByVal order As LongPtr, _
        ByVal callback As LongPtr, ByVal context As LongPtr) As LongPtr
    Private Declare PtrSafe Sub PT_AddRunLoopTimer _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFRunLoopAddTimer" (ByVal runLoop As LongPtr, ByVal timer As LongPtr, ByVal mode As LongPtr)
    Private Declare PtrSafe Sub PT_InvalidateRunLoopTimer _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFRunLoopTimerInvalidate" (ByVal timer As LongPtr)
    Private Declare PtrSafe Function PT_FindSymbol Lib "/usr/lib/libSystem.B.dylib" _
        Alias "dlsym" (ByVal handle As LongPtr, ByVal symbol As String) As LongPtr
    Private Declare PtrSafe Sub PT_CopyPointer Lib "/usr/lib/libSystem.B.dylib" _
        Alias "memmove" (ByRef destination As LongPtr, ByVal source As LongPtr, ByVal size As LongPtr)
    ' Run Apple's audio player asynchronously with a literal argument vector.
    ' There is no shell and no Objective-C call from VBA.
    Private Declare PtrSafe Function PT_Spawn Lib "/usr/lib/libSystem.B.dylib" _
        Alias "posix_spawn" (ByRef processID As Long, ByVal executable As LongPtr, _
        ByVal fileActions As LongPtr, ByVal attributes As LongPtr, _
        ByVal arguments As LongPtr, ByVal environment As LongPtr) As Long
    Private Declare PtrSafe Function PT_WaitPID Lib "/usr/lib/libSystem.B.dylib" _
        Alias "waitpid" (ByVal processID As Long, ByRef status As Long, ByVal options As Long) As Long
    Private Declare PtrSafe Function PT_Kill Lib "/usr/lib/libSystem.B.dylib" _
        Alias "kill" (ByVal processID As Long, ByVal signal As Long) As Long
#End If

Private Const PT_MARKER As String = "PT_TIMER"
Private Const PT_SECONDS As String = "PT_SECONDS"
Private Const PT_SOUND As String = "PT_SOUND"
Private Const PT_ALERT_NAME As String = "PT_ALERT_NAME"
' Used only to remove obsolete start cues from versions 1.2.0-1.2.3.
Private Const PT_LEGACY_CLICK As String = "PT_CLICK_START"
Private Const PT_VERSION As String = "1.2.4"
Private Const PT_MAX_SECONDS As Long = 86399
Private Const PT_ALERT_FILE As String = "/System/Library/Sounds/Ping.aiff"
Private Const PT_AUDIO_PLAYER As String = "/usr/bin/afplay"
Private Const PT_ALERT_REPEATS As Long = 3
Private Const PT_ALERT_GAP As Double = 1.8

Private Type CountdownState
    Remaining As Double
    Deadline As Double
    Status As String
End Type

Private Type SlideTimerState
    Timer As Shape
    TimerSlide As Slide
    Clock As CountdownState
    AlertsLeft As Long
    NextAlert As Double
    LastPaint As Long
    LastStatus As String
    WasVisible As Boolean
End Type

Private mStates() As SlideTimerState
Private mStateCount As Long
Private mOwner As Presentation
Private mEvents As SlideTimerEvents
Private mConfiguring As Boolean
Private mTicking As Boolean
Private mRefreshing As Boolean
Private mLastProblem As String
Private mTestAlertsLeft As Long
Private mTestNextAlert As Double
Private mTestSoundPath As String
#If Mac Then
Private mNativeTimer As LongPtr
Private mAudioPID As Long
#End If

' Renderer cache. Resolve PowerPoint objects once; keep tick updates small.
Private mPartParent As Shape
Private mPartCache As Collection
Private mSegmentShapes(1 To 6, 1 To 7) As Shape
Private mDotShapes(1 To 2, 1 To 2) As Shape
Private mLayoutCharacters As Long
Private mPreviousDisplay As String
Private mPreviousInk As Long
Private mPreviousStatus As String
Private mPreviousUnits As String
Private mPreviousSound As String
Private mPreviousIconInk As Long

' ------------ Commands available in Tools > Macro > Macros ------------

Public Sub Timer_Insert()
    Dim sld As Slide, timer As Shape, duration As Long, soundChoice As String
    If mConfiguring Then Exit Sub
    On Error GoTo Failed
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before inserting a timer.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set sld = EditingSlide()
    If sld Is Nothing Then Exit Sub
    Set timer = FindTimer(sld)
    If Not timer Is Nothing Then
        If ReadTag(timer, "PT_LAYOUT") <> "FLAT" Then Timer_RefreshStyle
        AttachPresentation sld.Parent
        Timer_Configure
        Exit Sub
    End If
    mConfiguring = True
    If Not AskDuration(300, duration) Then GoTo Done
    If Not AskSound("Ping", soundChoice) Then GoTo Done
    Set timer = CreateTimer(sld, duration)
    ApplySoundChoice timer, soundChoice
    AttachPresentation sld.Parent
    timer.Select
Done:
    mConfiguring = False
    Exit Sub
Failed:
    mConfiguring = False
    ReportError "insert the timer", Err.Description
End Sub

Public Sub Timer_Configure()
    Dim timer As Shape, duration As Long, accepted As Boolean, index As Long
    Dim wasRunning As Boolean, oldStatus As String, problem As String
    If mConfiguring Then Exit Sub
    On Error GoTo Failed
    Set timer = CurrentTimer()
    If timer Is Nothing Then
        MsgBox "This slide has no timer. Run Timer_Insert first.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    index = StateFor(timer)
    mConfiguring = True
    UpdateState index
    oldStatus = mStates(index).Clock.Status
    wasRunning = (oldStatus = "Running")
    If wasRunning Then mStates(index).Clock.Status = "Paused"
    If SlideShowWindows.Count > 0 Then
        duration = NextPresetDuration(DurationOf(timer))
        accepted = True
    Else
        accepted = AskDuration(DurationOf(timer), duration)
    End If
    If accepted Then
        timer.Tags.Add PT_SECONDS, CStr(duration)
        ResetCountdown mStates(index).Clock, duration
        mStates(index).AlertsLeft = 0
    ElseIf wasRunning Then
        StartCountdown mStates(index).Clock, DurationOf(timer), ClockSeconds()
    Else
        mStates(index).Clock.Status = oldStatus
    End If
    mConfiguring = False
    PaintState index
    Exit Sub
Failed:
    problem = Err.Description
    mConfiguring = False
    ReportError "configure the timer", problem
End Sub

Public Sub Timer_Initialize()
    ' Optional upgrade cleanup. Timer buttons attach close events themselves;
    ' no initialization macro is required before an ordinary slide show.
    On Error GoTo Failed
    If Presentations.Count = 0 Then Exit Sub
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before cleaning up old start animations.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    AttachPresentation ActivePresentation
    Exit Sub
Failed:
    ReportError "initialize the timer", Err.Description
End Sub

Public Sub Timer_Present()
    ' Optional convenience command retained for existing users.
    Dim win As SlideShowWindow
    On Error GoTo Failed
    If Presentations.Count = 0 Then Exit Sub
    AttachPresentation ActivePresentation
    If SlideShowWindows.Count = 0 Then Set win = mOwner.SlideShowSettings.Run
    Exit Sub
Failed:
    ReportError "start the presentation", Err.Description
End Sub

Public Sub Timer_Toggle()
    Dim timer As Shape, index As Long
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    If SlideShowWindows.Count = 0 Then
        MsgBox "Start Slide Show view, then click the timer's play/pause button.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    If ReadTag(timer, "PT_LAYOUT") <> "FLAT" Then
        MsgBox "Exit the show and run Timer_RefreshStyle to update this timer.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    index = StateFor(timer)
    UpdateState index
    If mStates(index).Clock.Status = "Running" Then
        mStates(index).Clock.Status = "Paused"
    Else
        StartCountdown mStates(index).Clock, DurationOf(timer), ClockSeconds()
        mStates(index).AlertsLeft = 0
    End If
    StartScheduler
    PaintState index
    Exit Sub
Failed:
    ReportError "run the timer", Err.Description
End Sub

Public Sub Timer_Start()
    ' Idempotent action macro: duplicate start requests never pause/restart.
    Dim timer As Shape, index As Long
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    index = StateFor(timer)
    UpdateState index
    If mStates(index).Clock.Status = "Running" Then Exit Sub
    StartCountdown mStates(index).Clock, DurationOf(timer), ClockSeconds()
    mStates(index).AlertsLeft = 0
    StartScheduler
    PaintState index
    Exit Sub
Failed:
    ReportError "start the timer", Err.Description
End Sub

Public Sub Timer_Reset()
    Dim timer As Shape, index As Long
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    index = StateFor(timer)
    ResetCountdown mStates(index).Clock, DurationOf(timer)
    mStates(index).AlertsLeft = 0
    PaintState index
    Exit Sub
Failed:
    ReportError "reset the timer", Err.Description
End Sub

Public Sub Timer_ToggleSound()
    Dim timer As Shape, index As Long
    On Error GoTo Failed
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    index = StateFor(timer)
    If SoundEnabled(timer) Then
        timer.Tags.Add PT_SOUND, "0"
        mStates(index).AlertsLeft = 0
    Else
        timer.Tags.Add PT_SOUND, "1"
    End If
    PaintSoundButton timer
    RefreshShowSurface timer
    Exit Sub
Failed:
    ReportError "change the sound setting", Err.Description
End Sub

Public Sub Timer_AddMinute()
    AdjustDuration 60
End Sub

Public Sub Timer_SubtractMinute()
    AdjustDuration -60
End Sub

Private Sub AdjustDuration(ByVal change As Long)
    Dim timer As Shape, oldDuration As Long, newDuration As Long, index As Long
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    index = StateFor(timer)
    oldDuration = DurationOf(timer)
    newDuration = oldDuration + change
    If newDuration < 1 Then newDuration = 1
    If newDuration > PT_MAX_SECONDS Then newDuration = PT_MAX_SECONDS
    If newDuration = oldDuration Then Exit Sub
    UpdateState index
    AdjustCountdown mStates(index).Clock, newDuration - oldDuration, newDuration, ClockSeconds()
    mStates(index).AlertsLeft = 0
    timer.Tags.Add PT_SECONDS, CStr(newDuration)
    PaintState index
    Exit Sub
Failed:
    ReportError "adjust the duration", Err.Description
End Sub

Public Sub Timer_RefreshStyle()
    Dim sld As Slide, oldTimer As Shape, fresh As Shape, duration As Long
    Dim oldLeft As Single, oldTop As Single, oldWidth As Single, soundWasOn As Boolean, soundNameTag As String
    On Error GoTo Failed
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before updating the timer's appearance.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set sld = EditingSlide()
    If sld Is Nothing Then Exit Sub
    Set oldTimer = FindTimer(sld)
    If oldTimer Is Nothing Then
        Timer_Insert
        Exit Sub
    End If
    duration = DurationOf(oldTimer)
    soundWasOn = SoundEnabled(oldTimer)
    soundNameTag = ReadTag(oldTimer, PT_ALERT_NAME)
    oldLeft = oldTimer.Left: oldTop = oldTimer.Top: oldWidth = oldTimer.Width
    Set fresh = CreateTimer(sld, duration)
    fresh.LockAspectRatio = msoTrue
    fresh.Width = oldWidth
    fresh.Left = oldLeft
    fresh.Top = oldTop
    If Not soundWasOn Then fresh.Tags.Add PT_SOUND, "0"
    If Len(soundNameTag) > 0 Then fresh.Tags.Add PT_ALERT_NAME, soundNameTag
    PaintTimer fresh, duration, "Ready"
    ForgetTimer oldTimer
    oldTimer.Delete
    AttachPresentation sld.Parent
    fresh.Name = "SlideTimer_" & CStr(sld.SlideID)
    fresh.Select
    Exit Sub
Failed:
    ReportError "update the timer's appearance", Err.Description
End Sub

Public Sub Timer_ChooseSound()
    Dim timer As Shape, initialChoice As String, choice As String, index As Long
    Dim problem As String
    If mConfiguring Then Exit Sub
    On Error GoTo Failed
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before choosing a sound.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set timer = CurrentTimer()
    If timer Is Nothing Then
        MsgBox "This slide has no timer. Run Timer_Insert first.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    index = StateFor(timer)
    initialChoice = SoundNameOf(timer)
    If Not SoundEnabled(timer) Then initialChoice = "None"
    mConfiguring = True
    If AskSound(initialChoice, choice) Then
        ApplySoundChoice timer, choice
        mStates(index).AlertsLeft = 0
        PaintSoundButton timer
        RefreshShowSurface timer
    End If
    mConfiguring = False
    Exit Sub
Failed:
    problem = Err.Description
    mConfiguring = False
    ReportError "choose the end sound", problem
End Sub

Public Sub Timer_TestSound()
    Dim timer As Shape
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    If Presentations.Count = 0 Then Exit Sub
    Set timer = CurrentTimer()
    If timer Is Nothing Then
        mTestSoundPath = PT_ALERT_FILE
    Else
        If Not SoundEnabled(timer) Then
            MsgBox "Sound is off for this timer. Turn SOUND ON or run Timer_ChooseSound first.", _
                vbInformation, "Slide Timer"
            Exit Sub
        End If
        mTestSoundPath = SoundPathOf(timer)
    End If
    AttachPresentation ActivePresentation
    mTestAlertsLeft = 0
    PlayAlert mTestSoundPath
    mTestAlertsLeft = PT_ALERT_REPEATS - 1
    mTestNextAlert = ClockSeconds() + PT_ALERT_GAP
    StartScheduler
    Exit Sub
Failed:
    ReportError "play the end sound", Err.Description
End Sub

Public Sub Timer_Remove()
    Dim timer As Shape, sld As Slide
    On Error GoTo Failed
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before removing a timer.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set sld = EditingSlide()
    If sld Is Nothing Then Exit Sub
    Set timer = FindTimer(sld)
    If timer Is Nothing Then Exit Sub
    ForgetTimer timer
    RemoveLegacyStartCue sld
    timer.Delete
    Exit Sub
Failed:
    ReportError "remove the timer", Err.Description
End Sub

Public Sub Timer_Shutdown()
    ' Always call before editing/removing VBA code or pressing VBE Reset.
    ' PresentationBeforeClose invokes this before the project can unload.
    Dim i As Long
    StopScheduler
    mTestAlertsLeft = 0
    On Error Resume Next
    For i = 1 To mStateCount
        PauseCountdown mStates(i).Clock, ClockSeconds()
        mStates(i).AlertsLeft = 0
        If Not mStates(i).Timer Is Nothing Then PaintState i
    Next i
    StopAudioPlayer
    On Error GoTo 0
End Sub

Public Sub Timer_Help()
    MsgBox "SLIDE TIMER " & PT_VERSION & vbCrLf & vbCrLf & _
        "Normal view: Timer_Insert asks for duration and end sound." & vbCrLf & _
        "Start the slide show normally, then click play/pause or the digits." & vbCrLf & _
        "Play/pause or digits: start, pause or resume. RESET: ready again." & vbCrLf & _
        "SET: presets. Plus/minus: one minute. SOUND: end alert on/off." & vbCrLf & vbCrLf & _
        "Timers keep counting across slides, app switches and Esc." & vbCrLf & _
        "Timer_ChooseSound changes the sound; type p 2 to preview choice 2." & vbCrLf & _
        "The alert follows the selected macOS media output and volume." & vbCrLf & _
        "Run Timer_Shutdown before changing the VBA code." & vbCrLf & _
        "Save as .pptm; import both SlideTimer and SlideTimerEvents.", vbInformation, "Slide Timer"
End Sub

' ------------ Timer engine: short callbacks, never a blocking macro loop ------------

Private Sub AttachPresentation(ByVal owner As Presentation)
    If Not mOwner Is Nothing Then
        If Not (owner Is mOwner) Then
            Timer_Shutdown
            Erase mStates
            mStateCount = 0
        End If
    End If
    Set mOwner = owner
    If mEvents Is Nothing Then Set mEvents = New SlideTimerEvents
    Set mEvents.App = Application
    ' Only change slide objects in editing view. This also removes the extra
    ' animation step left by earlier versions without touching user effects.
    If SlideShowWindows.Count = 0 Then RemoveLegacyStartCues owner
End Sub

Public Sub Timer_OnBeforeClose(ByVal owner As Presentation)
    If mOwner Is Nothing Then Exit Sub
    If Not (owner Is mOwner) Then Exit Sub
    ' Keep paused state if a save prompt or another handler cancels closing.
    Timer_Shutdown
End Sub

Public Sub Timer_OnClose(ByVal owner As Presentation)
    If mOwner Is Nothing Then Exit Sub
    If Not (owner Is mOwner) Then Exit Sub
    Timer_Shutdown
    Erase mStates
    mStateCount = 0
    Set mPartParent = Nothing
    Set mPartCache = Nothing
    Set mEvents.App = Nothing
    Set mEvents = Nothing
    Set mOwner = Nothing
End Sub

Private Function StateFor(ByVal timer As Shape) As Long
    Dim i As Long, sld As Slide
    For i = 1 To mStateCount
        If Not mStates(i).Timer Is Nothing Then
            If timer Is mStates(i).Timer Then
                StateFor = i
                Exit Function
            End If
        End If
    Next i
    Set sld = timer.Parent
    AttachPresentation sld.Parent
    mStateCount = mStateCount + 1
    ReDim Preserve mStates(1 To mStateCount)
    With mStates(mStateCount)
        Set .Timer = timer
        Set .TimerSlide = sld
        ResetCountdown .Clock, DurationOf(timer)
        .LastPaint = -1
    End With
    StateFor = mStateCount
End Function

Private Sub ForgetTimer(ByVal timer As Shape)
    Dim i As Long
    For i = 1 To mStateCount
        If Not mStates(i).Timer Is Nothing Then
            If timer Is mStates(i).Timer Then
                mStates(i).Clock.Status = "Ready"
                mStates(i).AlertsLeft = 0
                Set mStates(i).Timer = Nothing
                Set mStates(i).TimerSlide = Nothing
            End If
        End If
    Next i
End Sub

Private Sub StartScheduler()
#If Mac Then
    Dim mode As LongPtr, symbol As LongPtr
    If mNativeTimer <> 0 Then Exit Sub
    On Error GoTo Failed
    ' Common modes is a sentinel checked by pointer identity. An equal
    ' CFString created from text does NOT register a timer in common modes.
    ' RTLD_DEFAULT = -2; dlsym returns the address of the exported pointer.
    symbol = PT_FindSymbol(-2, "kCFRunLoopCommonModes")
    If symbol = 0 Then Err.Raise vbObjectError + 2050, , "Cannot find the run-loop mode."
    PT_CopyPointer mode, symbol, LenB(mode)
    If mode = 0 Then Err.Raise vbObjectError + 2050, , "Cannot read the run-loop mode."
    mNativeTimer = PT_CreateRunLoopTimer(0, PT_CFAbsoluteTime() + 0.1, 0.1, 0, 0, AddressOf Timer_NativeTick, 0)
    If mNativeTimer = 0 Then Err.Raise vbObjectError + 2050, , "Cannot schedule the timer."
    PT_AddRunLoopTimer PT_MainRunLoop(), mNativeTimer, mode
    Exit Sub
Failed:
    StopScheduler
    Err.Raise vbObjectError + 2050, "SlideTimer", "Cannot schedule background updates."
#Else
    Err.Raise vbObjectError + 2050, "SlideTimer", "Slide Timer requires PowerPoint for Mac."
#End If
End Sub

Private Sub StopScheduler()
#If Mac Then
    If mNativeTimer = 0 Then Exit Sub
    PT_InvalidateRunLoopTimer mNativeTimer
    PT_CFRelease mNativeTimer
    mNativeTimer = 0
#End If
End Sub

Public Sub Timer_NativeTick(ByVal nativeTimer As LongPtr, ByVal context As LongPtr)
    ' Core Foundation calls on PowerPoint's main UI thread. No DoEvents,
    ' dialogs or unhandled errors may escape this native callback.
    Dim i As Long, win As SlideShowWindow, busy As Boolean
    On Error GoTo Failed
    If mTicking Or mConfiguring Or mRefreshing Then Exit Sub
    mTicking = True
    Set win = OwnerShowWindow()
    For i = 1 To mStateCount
        If StateIsLive(i) Then
            UpdateState i
            ContinueAlert i
            PaintStateIfChanged i, win
            If mStates(i).Clock.Status = "Running" Or mStates(i).AlertsLeft > 0 Then busy = True
        End If
    Next i
    If mTestAlertsLeft > 0 Then
        busy = True
        If ClockSeconds() >= mTestNextAlert And Not AudioPlayerRunning() Then
            If TryPlayAlert(mTestSoundPath) Then
                mTestAlertsLeft = mTestAlertsLeft - 1
            Else
                mTestAlertsLeft = 0
            End If
            mTestNextAlert = ClockSeconds() + PT_ALERT_GAP
        End If
    End If
    ' Keep polling until the last newly launched child has exited and been reaped.
    If AudioPlayerRunning() Then busy = True
    If Not busy Then StopScheduler
    mTicking = False
    Exit Sub
Failed:
    mLastProblem = Err.Description
    ' Stop native callbacks before showing a problem through Timer_CheckCurrent.
    On Error Resume Next
    StopScheduler
    StopAudioPlayer
    mTicking = False
End Sub

Private Function StateIsLive(ByVal index As Long) As Boolean
    Dim sld As Slide, shp As Shape, shapeId As Long
    On Error GoTo Gone
    If mStates(index).Timer Is Nothing Then Exit Function
    If ReadTag(mStates(index).Timer, PT_MARKER) <> "1" Then GoTo Gone
    ' Resolve through the owning deck too: a deleted slide may leave a COM wrapper.
    Set sld = mOwner.Slides.FindBySlideID(mStates(index).TimerSlide.SlideID)
    shapeId = mStates(index).Timer.Id
    For Each shp In sld.Shapes
        If shp.Id = shapeId Then
            StateIsLive = True
            Exit Function
        End If
    Next shp
Gone:
    mStates(index).Clock.Status = "Ready"
    mStates(index).AlertsLeft = 0
    Set mStates(index).Timer = Nothing
    Set mStates(index).TimerSlide = Nothing
End Function

Private Function OwnerShowWindow() As SlideShowWindow
    Dim win As SlideShowWindow
    On Error GoTo Gone
    If mOwner Is Nothing Then Exit Function
    For Each win In SlideShowWindows
        If win.Presentation Is mOwner Then
            If win.View.State <> ppSlideShowDone Then Set OwnerShowWindow = win
            Exit Function
        End If
    Next win
Gone:
End Function

Private Sub UpdateState(ByVal index As Long)
    If AdvanceCountdown(mStates(index).Clock, ClockSeconds()) Then
        If SoundEnabled(mStates(index).Timer) Then
            mStates(index).AlertsLeft = PT_ALERT_REPEATS
            mStates(index).NextAlert = ClockSeconds()
        End If
    End If
End Sub

Private Sub ContinueAlert(ByVal index As Long)
    If mStates(index).AlertsLeft <= 0 Then Exit Sub
    If mStates(index).Clock.Status <> "Time's up" Or Not SoundEnabled(mStates(index).Timer) Then
        mStates(index).AlertsLeft = 0
    ElseIf ClockSeconds() >= mStates(index).NextAlert Then
        ' Let long selected sounds finish before another timer/repeat plays.
        If AudioPlayerRunning() Then Exit Sub
        If TryPlayAlert(SoundPathOf(mStates(index).Timer)) Then
            mStates(index).AlertsLeft = mStates(index).AlertsLeft - 1
        Else
            mStates(index).AlertsLeft = 0
        End If
        mStates(index).NextAlert = ClockSeconds() + PT_ALERT_GAP
    End If
End Sub

' Pure state transitions. Presentation visibility never changes the clock.
Private Function AdvanceCountdown(ByRef state As CountdownState, ByVal nowSeconds As Double) As Boolean
    If state.Status <> "Running" Then Exit Function
    state.Remaining = state.Deadline - nowSeconds
    If state.Remaining <= 0 Then
        state.Remaining = 0
        state.Status = "Time's up"
        AdvanceCountdown = True
    End If
End Function

Private Sub StartCountdown(ByRef state As CountdownState, ByVal duration As Long, ByVal nowSeconds As Double)
    If state.Status = "Running" Then Exit Sub
    If state.Status = "Time's up" Or state.Remaining <= 0 Then state.Remaining = duration
    state.Deadline = nowSeconds + state.Remaining
    state.Status = "Running"
End Sub

Private Sub PauseCountdown(ByRef state As CountdownState, ByVal nowSeconds As Double)
    Dim finishedNow As Boolean
    finishedNow = AdvanceCountdown(state, nowSeconds)
    If state.Status = "Running" Then state.Status = "Paused"
End Sub

Private Sub ResetCountdown(ByRef state As CountdownState, ByVal duration As Long)
    state.Remaining = duration
    state.Deadline = 0
    state.Status = "Ready"
End Sub

Private Sub AdjustCountdown(ByRef state As CountdownState, ByVal change As Long, _
                            ByVal duration As Long, ByVal nowSeconds As Double)
    If state.Status = "Time's up" Or state.Status = "Ready" Then
        ResetCountdown state, duration
    Else
        state.Remaining = state.Remaining + change
        If state.Remaining < 1 Then state.Remaining = 1
        If state.Remaining > PT_MAX_SECONDS Then state.Remaining = PT_MAX_SECONDS
        If state.Status = "Running" Then state.Deadline = nowSeconds + state.Remaining
    End If
End Sub

Private Function ClockSeconds() As Double
#If Mac Then
    ClockSeconds = PT_MonotonicTime()
#Else
    ClockSeconds = CDbl(Now) * 86400#
#End If
End Function

Private Sub YieldBriefly()
#If Mac Then
    Dim ignored As Long
    ' Used only by the optional native-clock diagnostic, never the engine.
    ignored = PT_USleep(30000)
#End If
End Sub

Private Function TryPlayAlert(ByVal soundPath As String) As Boolean
    On Error GoTo Failed
    PlayAlert soundPath
    TryPlayAlert = True
    Exit Function
Failed:
    ' An unavailable output must not stop this or any other countdown.
    mLastProblem = "End sound: " & Err.Description
End Function

Private Sub PlayAlert(ByVal soundPath As String)
#If Mac Then
    Dim executable() As Byte, soundFile() As Byte, arguments(0 To 2) As LongPtr
    Dim environment(0 To 0) As LongPtr, processID As Long, result As Long
    ' Keep every byte buffer and argv alive until posix_spawn has copied them.
    executable = UTF8CString(PT_AUDIO_PLAYER)
    soundFile = UTF8CString(soundPath)
    arguments(0) = VarPtr(executable(0))
    arguments(1) = VarPtr(soundFile(0))
    arguments(2) = 0
    ' Start a fresh short alert if a previous one is still playing.
    StopAudioPlayer
    result = PT_Spawn(processID, VarPtr(executable(0)), 0, 0, VarPtr(arguments(0)), VarPtr(environment(0)))
    If result <> 0 Then
        Err.Raise vbObjectError + 2051, "SlideTimer", _
            "macOS could not start afplay (error " & CStr(result) & ")."
    End If
    mAudioPID = processID
    ' The child inherits PowerPoint's sandbox. A denied launch/read/playback
    ' becomes an error, never an attempt to bypass the sandbox or use Beep.
#End If
End Sub

Private Function AudioPlayerRunning() As Boolean
#If Mac Then
    Dim result As Long, status As Long
    If mAudioPID <= 0 Then Exit Function
    result = PT_WaitPID(mAudioPID, status, 1) ' WNOHANG: do not wait for audio.
    If result = 0 Then
        AudioPlayerRunning = True
        Exit Function
    End If
    mAudioPID = 0
    If result > 0 And status <> 0 Then
        mLastProblem = "afplay could not finish playback (status " & CStr(status) & _
            "). Check the sound file and selected media output."
    End If
#End If
End Function

Private Sub StopAudioPlayer()
#If Mac Then
    Dim ignored As Long, status As Long
    If Not AudioPlayerRunning() Then Exit Sub
    ' Only signal our own unreaped child; its PID cannot have been reused.
    ignored = PT_Kill(mAudioPID, 15) ' SIGTERM; afplay uses its default handler.
    ignored = PT_WaitPID(mAudioPID, status, 0) ' Reap the terminated child.
    mAudioPID = 0
#End If
End Sub

Private Function UTF8CString(ByVal value As String) As Byte()
    ' Explicit UTF-8, including surrogate pairs. No implicit VBA/native string
    ' conversion, shell parsing or quoting is involved in launching afplay.
    Dim bytes() As Byte, i As Long, used As Long, code As Long, low As Long
    ReDim bytes(0 To Len(value) * 3) ' Includes at least one zero terminator.
    i = 1
    Do While i <= Len(value)
        code = AscW(Mid$(value, i, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& Then
            If i = Len(value) Then GoTo InvalidText
            low = AscW(Mid$(value, i + 1, 1)) And &HFFFF&
            If low < &HDC00& Or low > &HDFFF& Then GoTo InvalidText
            code = &H10000 + (code - &HD800&) * &H400& + low - &HDC00&
            i = i + 1
        ElseIf code >= &HDC00& And code <= &HDFFF& Then
            GoTo InvalidText
        End If
        If code = 0 Then GoTo InvalidText
        If code < &H80 Then
            bytes(used) = code: used = used + 1
        ElseIf code < &H800 Then
            bytes(used) = &HC0 Or (code \ &H40): used = used + 1
            bytes(used) = &H80 Or (code And &H3F): used = used + 1
        ElseIf code < &H10000 Then
            bytes(used) = &HE0 Or (code \ &H1000): used = used + 1
            bytes(used) = &H80 Or ((code \ &H40) And &H3F): used = used + 1
            bytes(used) = &H80 Or (code And &H3F): used = used + 1
        Else
            bytes(used) = &HF0 Or (code \ &H40000): used = used + 1
            bytes(used) = &H80 Or ((code \ &H1000) And &H3F): used = used + 1
            bytes(used) = &H80 Or ((code \ &H40) And &H3F): used = used + 1
            bytes(used) = &H80 Or (code And &H3F): used = used + 1
        End If
        i = i + 1
    Loop
    ReDim Preserve bytes(0 To used)
    UTF8CString = bytes
    Exit Function
InvalidText:
    Err.Raise vbObjectError + 2051, "SlideTimer", "Invalid text in the audio player or sound path."
End Function

' ------------ Slide objects and controls ------------

Private Function CreateTimer(ByVal sld As Slide, ByVal duration As Long) As Shape
    Dim names() As Variant, count As Long, piece As Shape, group As Shape
    Dim x As Single, y As Single, i As Long, offset As Single
    Dim role As String, caption As String, action As String, stage As String
    Dim originalIds As Collection, oldShape As Shape
    Dim problem As String, errorNumber As Long
    On Error GoTo Failed
    Set originalIds = New Collection
    For Each oldShape In sld.Shapes
        originalIds.Add CLng(oldShape.Id)
    Next oldShape
    stage = "drawing the timer case"
    ReDim names(0 To 99)
    x = (ActivePresentation.PageSetup.SlideWidth - 394) / 2
    y = (ActivePresentation.PageSetup.SlideHeight - 288) / 2
    If x < 0 Then x = 0
    If y < 0 Then y = 0

    Set piece = AddSurface(sld, "PT_CaseEdge", x + 5, y + 6, 388, 282, RGB(23, 25, 30), 0.08)
    RememberShape names, count, piece
    Set piece = AddSurface(sld, "PT_Background", x, y, 388, 282, RGB(38, 40, 46), 0.08)
    piece.Line.Visible = msoTrue
    piece.Line.ForeColor.RGB = RGB(65, 68, 75)
    piece.Line.Weight = 0.8
    RememberShape names, count, piece

    Set piece = AddSurface(sld, "PT_Bezel", x + 27, y + 23, 334, 154, RGB(15, 20, 24), 0.06)
    RememberShape names, count, piece
    Set piece = AddSurface(sld, "PT_LCD", x + 33, y + 29, 322, 142, RGB(174, 198, 204), 0.04)
    piece.Line.Visible = msoTrue
    piece.Line.ForeColor.RGB = RGB(112, 141, 152)
    piece.Line.Weight = 1.5
    RememberShape names, count, piece

    Set piece = AddText(sld, "PT_Title", "TIMER", x + 46, y + 35, 55, 13, 7, RGB(54, 76, 87))
    RememberShape names, count, piece
    Set piece = AddText(sld, "PT_Sound", "SOUND ON", x + 267, y + 35, 74, 13, 7, RGB(54, 76, 87))
    SetClick piece, "Timer_ToggleSound"
    RememberShape names, count, piece

    ' All 42 segments are direct children of the final timer group.
    ' There are no intermediate digit or colon groups to survive regrouping.
    For i = 1 To 6
        stage = "creating flat LCD segments for digit " & CStr(i)
        AddDigitSegments sld, i, x + 55, y + 53, names, count
    Next i
    For i = 1 To 2
        stage = "creating LCD separator dots " & CStr(i)
        AddColonDots sld, i, x + 174, y + 53, names, count
    Next i

    Set piece = AddText(sld, "PT_Status", "READY", x + 46, y + 153, 100, 13, 7, RGB(54, 76, 87))
    RememberShape names, count, piece
    Set piece = AddText(sld, "PT_Units", "MIN    SEC", x + 216, y + 153, 124, 13, 7, RGB(54, 76, 87))
    RememberShape names, count, piece

    ' A clear overlay catches digit clicks without covering sound or buttons.
    Set piece = sld.Shapes.AddShape(msoShapeRectangle, x + 46, y + 51, 296, 100)
    piece.Name = "PT_Display"
    piece.Tags.Add "PT_ROLE", "PT_Display"
    piece.Fill.Transparency = 1
    piece.Line.Visible = msoFalse
    SetClick piece, "Timer_Toggle"
    RememberShape names, count, piece

    stage = "creating the round buttons"
    For i = 0 To 3
        offset = x + 34 + i * 88
        Select Case i
            Case 0: role = "PT_Time": caption = "SET": action = "Timer_Configure"
            Case 1: role = "PT_Plus": caption = "+": action = "Timer_AddMinute"
            Case 2: role = "PT_Minus": caption = ChrW(&H2212): action = "Timer_SubtractMinute"
            Case 3: role = "PT_StartBase": caption = "": action = "Timer_Toggle"
        End Select
        Set piece = AddRoundButton(sld, role, offset, y + 201)
        If i = 0 Then
            SetButtonLabel piece, caption, 12
        ElseIf i < 3 Then
            SetButtonLabel piece, caption, 29
        End If
        SetClick piece, action
        RememberShape names, count, piece
    Next i

    Set piece = AddPolygon(sld, "PT_PlayIcon", "0,0 0,20 13,10", x + 311, y + 218)
    piece.Fill.ForeColor.RGB = RGB(227, 238, 240)
    RememberShape names, count, piece
    Set piece = AddSurface(sld, "PT_Pause1", x + 328, y + 218, 3, 20, RGB(227, 238, 240), 0)
    RememberShape names, count, piece
    Set piece = AddSurface(sld, "PT_Pause2", x + 334, y + 218, 3, 20, RGB(227, 238, 240), 0)
    RememberShape names, count, piece
    Set piece = sld.Shapes.AddShape(msoShapeOval, x + 298, y + 201, 54, 54)
    piece.Name = "PT_Start"
    piece.Tags.Add "PT_ROLE", "PT_Start"
    piece.Fill.Transparency = 1
    piece.Line.Visible = msoFalse
    SetClick piece, "Timer_Toggle"
    RememberShape names, count, piece

    Set piece = AddText(sld, "PT_Footer", "COUNTDOWN", x + 35, y + 264, 93, 12, 6, RGB(137, 144, 154))
    RememberShape names, count, piece
    Set piece = AddText(sld, "PT_Reset", "RESET", x + 291, y + 264, 65, 12, 7, RGB(181, 188, 198))
    SetClick piece, "Timer_Reset"
    RememberShape names, count, piece

    ReDim Preserve names(0 To count - 1)
    stage = "grouping the timer components"
    Set group = GroupByIds(sld, names)
    group.Name = "SlideTimer_" & CStr(sld.SlideID)
    group.Tags.Add PT_MARKER, "1"
    group.Tags.Add PT_SECONDS, CStr(duration)
    group.Tags.Add PT_SOUND, "1"
    group.Tags.Add "PT_STYLE", "LCD"
    group.Tags.Add "PT_LAYOUT", "FLAT"
    stage = "binding the timer controls"
    SetClick Part(group, "PT_Display"), "Timer_Toggle"
    SetClick Part(group, "PT_Start"), "Timer_Toggle"
    SetClick Part(group, "PT_StartBase"), "Timer_Toggle"
    SetClick Part(group, "PT_Reset"), "Timer_Reset"
    SetClick Part(group, "PT_Time"), "Timer_Configure"
    SetClick Part(group, "PT_Plus"), "Timer_AddMinute"
    SetClick Part(group, "PT_Minus"), "Timer_SubtractMinute"
    SetClick Part(group, "PT_Sound"), "Timer_ToggleSound"
    SetButtonLabel Part(group, "PT_Time"), "SET", 12
    SetButtonLabel Part(group, "PT_Plus"), "+", 29
    SetButtonLabel Part(group, "PT_Minus"), ChrW(&H2212), 29
    group.AlternativeText = "LCD countdown timer. SET changes the duration. Plus and minus adjust one minute. Play/pause starts or pauses. RESET and SOUND controls are also available."
    stage = "checking the flat LCD components"
    ValidateTimer group
    stage = "painting the LCD digits"
    PaintTimer group, duration, "Ready"
    Set CreateTimer = group
    Exit Function
Failed:
    errorNumber = Err.Number
    problem = Err.Description
    On Error Resume Next
    If Not originalIds Is Nothing Then DeleteNewShapes sld, originalIds
    On Error GoTo 0
    Err.Raise errorNumber, "SlideTimer.CreateTimer", "While " & stage & ": " & problem
End Function

Private Sub RemoveLegacyStartCues(ByVal owner As Presentation)
    Dim sld As Slide
    For Each sld In owner.Slides
        RemoveLegacyStartCue sld
    Next sld
End Sub

Private Sub RemoveLegacyStartCue(ByVal sld As Slide)
    ' Migration only: remove our old tagged start effect/shape, leaving
    ' ordinary slide content and its animation sequence in place.
    Dim i As Long, effect As Effect
    For i = sld.TimeLine.MainSequence.Count To 1 Step -1
        Set effect = sld.TimeLine.MainSequence(i)
        If ReadTag(effect.Shape, PT_LEGACY_CLICK) = "1" Then effect.Delete
    Next i
    For i = sld.Shapes.Count To 1 Step -1
        If ReadTag(sld.Shapes(i), PT_LEGACY_CLICK) = "1" Then sld.Shapes(i).Delete
    Next i
End Sub

Private Function ReadTag(ByVal owner As Object, ByVal tagName As String) As String
    ' Missing tags are normal on slide placeholders and user content.
    ' Enumerate tags by index instead of asking PowerPoint for a missing name.
    Dim i As Long
    For i = 1 To owner.Tags.Count
        If StrComp(owner.Tags.Name(i), tagName, vbTextCompare) = 0 Then
            ReadTag = owner.Tags.Value(i)
            Exit Function
        End If
    Next i
End Function

Private Sub BuildPartCache(ByVal parent As Shape)
    Dim i As Long, child As Shape, role As String, cache As Collection
    Set cache = New Collection
    For i = 1 To parent.GroupItems.Count
        Set child = parent.GroupItems.Item(i)
        role = ReadTag(child, "PT_ROLE")
        If Len(role) = 0 Then role = child.Name
        If Len(role) > 0 Then cache.Add child, UCase$(role)
    Next i
    Set mPartParent = parent
    Set mPartCache = cache
    mLayoutCharacters = 0
    mPreviousDisplay = ""
    mPreviousInk = -1
    mPreviousStatus = ""
    mPreviousUnits = ""
    mPreviousSound = ""
    mPreviousIconInk = -1
End Sub

Private Function Part(ByVal parent As Shape, ByVal role As String) As Shape
    On Error GoTo Missing
    If mPartCache Is Nothing Then
        BuildPartCache parent
    ElseIf Not (parent Is mPartParent) Then
        BuildPartCache parent
    End If
    ' This is a VBA Collection lookup, not a PowerPoint name lookup.
    Set Part = mPartCache.Item(UCase$(role))
    Exit Function
Missing:
    Err.Raise vbObjectError + 2049, "SlideTimer.Part", _
        "Missing timer component: " & role & ". Run Timer_RefreshStyle in Normal view."
End Function

Private Function GroupByIds(ByVal sld As Slide, ByRef ids() As Variant) As Shape
    Dim indexes() As Variant, wanted As Long, current As Long, found As Boolean
    Dim roles() As String, geometry() As String, fields As Variant, field As Long
    Dim child As Shape, group As Shape, position As Long
    ReDim indexes(LBound(ids) To UBound(ids))
    ReDim roles(LBound(ids) To UBound(ids))
    ReDim geometry(LBound(ids) To UBound(ids), 0 To 3)
    fields = Array("PT_LX", "PT_LY", "PT_LW", "PT_LH")
    For wanted = LBound(ids) To UBound(ids)
        found = False
        For current = 1 To sld.Shapes.Count
            Set child = sld.Shapes.Item(current)
            If child.Id = CLng(ids(wanted)) Then
                indexes(wanted) = CLng(current)
                roles(wanted) = ReadTag(child, "PT_ROLE")
                For field = 0 To 3
                    geometry(wanted, field) = ReadTag(child, CStr(fields(field)))
                Next field
                found = True
                Exit For
            End If
        Next current
        If Not found Then
            Err.Raise vbObjectError + 2050, "SlideTimer.GroupByIds", _
                "A newly created shape is missing (ID " & CStr(ids(wanted)) & ")."
        End If
        If wanted > LBound(ids) Then
            If CLng(indexes(wanted)) <= CLng(indexes(wanted - 1)) Then
                Err.Raise vbObjectError + 2054, "SlideTimer", "Timer component order changed before grouping."
            End If
        End If
    Next wanted
    ' Exactly one group, containing only individual shapes in drawing order.
    Set group = sld.Shapes.Range(indexes).Group
    If group.GroupItems.Count <> UBound(ids) - LBound(ids) + 1 Then
        Err.Raise vbObjectError + 2055, "SlideTimer", "PowerPoint returned an incomplete flat timer group."
    End If
    ' Stamp metadata onto the actual children after the grouping operation.
    ' No digit-group identity or tag retention across grouping is required.
    For wanted = LBound(ids) To UBound(ids)
        position = wanted - LBound(ids) + 1
        Set child = group.GroupItems.Item(position)
        child.Tags.Add "PT_ROLE", roles(wanted)
        For field = 0 To 3
            If Len(geometry(wanted, field)) > 0 Then
                child.Tags.Add CStr(fields(field)), geometry(wanted, field)
            End If
        Next field
    Next wanted
    Set GroupByIds = group
End Function

Private Sub DeleteNewShapes(ByVal sld As Slide, ByVal originalIds As Collection)
    Dim i As Long, original As Variant, keep As Boolean
    For i = sld.Shapes.Count To 1 Step -1
        keep = False
        For Each original In originalIds
            If sld.Shapes.Item(i).Id = CLng(original) Then
                keep = True
                Exit For
            End If
        Next original
        If Not keep Then sld.Shapes.Item(i).Delete
    Next i
End Sub

Private Sub SetButtonLabel(ByVal button As Shape, ByVal caption As String, ByVal size As Single)
    ' The caption belongs to the circle itself, so its center follows the
    ' button when PowerPoint groups, moves or resizes the timer.
    With button.TextFrame2
        .HorizontalAnchor = msoAnchorCenter
        .AutoSize = msoAutoSizeNone
        .MarginLeft = 0: .MarginRight = 0
        .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .text = caption
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = "Arial"
            .Font.Size = size
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(227, 238, 240)
        End With
    End With
End Sub

Private Sub RememberShape(ByRef names() As Variant, ByRef count As Long, ByVal shp As Shape)
    names(count) = CLng(shp.Id)
    count = count + 1
End Sub

Private Function AddSurface(ByVal sld As Slide, ByVal objectName As String, _
    ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single, _
    ByVal color As Long, ByVal rounding As Single) As Shape
    Dim shp As Shape
    If rounding > 0 Then
        Set shp = sld.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
        shp.Adjustments(1) = rounding
    Else
        Set shp = sld.Shapes.AddShape(msoShapeRectangle, x, y, w, h)
    End If
    shp.Name = objectName
    shp.Tags.Add "PT_ROLE", objectName
    shp.Fill.ForeColor.RGB = color
    shp.Line.Visible = msoFalse
    Set AddSurface = shp
End Function

Private Function AddRoundButton(ByVal sld As Slide, ByVal objectName As String, _
    ByVal x As Single, ByVal y As Single) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(msoShapeOval, x, y, 54, 54)
    shp.Name = objectName
    shp.Tags.Add "PT_ROLE", objectName
    shp.Fill.ForeColor.RGB = RGB(43, 46, 53)
    shp.Line.ForeColor.RGB = RGB(15, 18, 22)
    shp.Line.Weight = 1.8
    Set AddRoundButton = shp
End Function

Private Function AddPolygon(ByVal sld As Slide, ByVal objectName As String, _
    ByVal path As String, ByVal x As Single, ByVal y As Single) As Shape
    Dim points() As String, pair() As String, builder As FreeformBuilder, i As Long
    points = Split(path, " ")
    pair = Split(points(0), ",")
    Set builder = sld.Shapes.BuildFreeform(msoEditingCorner, x + CSng(pair(0)), y + CSng(pair(1)))
    For i = 1 To UBound(points)
        pair = Split(points(i), ",")
        builder.AddNodes msoSegmentLine, msoEditingAuto, x + CSng(pair(0)), y + CSng(pair(1))
    Next i
    pair = Split(points(0), ",")
    builder.AddNodes msoSegmentLine, msoEditingAuto, x + CSng(pair(0)), y + CSng(pair(1))
    Set AddPolygon = builder.ConvertToShape
    AddPolygon.Name = objectName
    AddPolygon.Tags.Add "PT_ROLE", objectName
    AddPolygon.Line.Visible = msoFalse
End Function

Private Function SegmentPath(ByVal segment As Long) As String
    Select Case segment
        Case 1: SegmentPath = "9,0 45,0 51,6 45,12 9,12 3,6"
        Case 2: SegmentPath = "46,14 52,8 52,42 46,48 40,42 40,20"
        Case 3: SegmentPath = "46,52 52,58 52,92 46,98 40,92 40,58"
        Case 4: SegmentPath = "9,88 45,88 51,94 45,100 9,100 3,94"
        Case 5: SegmentPath = "6,52 12,58 12,86 6,92 0,86 0,58"
        Case 6: SegmentPath = "6,8 12,14 12,42 6,48 0,42 0,14"
        Case 7: SegmentPath = "9,44 45,44 51,50 45,56 9,56 3,50"
    End Select
End Function

Private Sub AddDigitSegments(ByVal sld As Slide, ByVal digitIndex As Long, _
    ByVal x As Single, ByVal y As Single, ByRef ids() As Variant, ByRef count As Long)
    Dim shp As Shape, builder As FreeformBuilder, role As String
    Dim i As Long, j As Long, points() As String, pair() As String
    Dim px As Single, py As Single, firstX As Single, firstY As Single
    Dim minX As Single, minY As Single, maxX As Single, maxY As Single
    For i = 1 To 7
        role = "PT_Digit" & CStr(digitIndex) & "_" & Chr$(64 + i)
        points = Split(SegmentPath(i), " ")
        For j = 0 To UBound(points)
            pair = Split(points(j), ",")
            py = CSng(pair(1))
            px = 10 + CSng(pair(0)) - py * 0.105
            If j = 0 Then
                firstX = px: firstY = py
                minX = px: maxX = px: minY = py: maxY = py
                Set builder = sld.Shapes.BuildFreeform(msoEditingCorner, x + px, y + py)
            Else
                builder.AddNodes msoSegmentLine, msoEditingAuto, x + px, y + py
                If px < minX Then minX = px
                If px > maxX Then maxX = px
                If py < minY Then minY = py
                If py > maxY Then maxY = py
            End If
        Next j
        builder.AddNodes msoSegmentLine, msoEditingAuto, x + firstX, y + firstY
        Set shp = builder.ConvertToShape
        shp.Name = role
        shp.Tags.Add "PT_ROLE", role
        shp.Fill.ForeColor.RGB = RGB(39, 56, 68)
        shp.Line.Visible = msoFalse
        StoreLCDGeometry shp, minX, minY, maxX - minX, maxY - minY
        RememberShape ids, count, shp
    Next i
End Sub

Private Sub AddColonDots(ByVal sld As Slide, ByVal colonIndex As Long, _
    ByVal x As Single, ByVal y As Single, ByRef ids() As Variant, ByRef count As Long)
    Dim shp As Shape, role As String
    role = "PT_Colon" & CStr(colonIndex)
    Set shp = AddSurface(sld, role & "_Top", x + 11, y + 29, 7, 8, RGB(39, 56, 68), 0)
    StoreLCDGeometry shp, 11, 29, 7, 8
    RememberShape ids, count, shp
    Set shp = AddSurface(sld, role & "_Bottom", x + 11, y + 65, 7, 8, RGB(39, 56, 68), 0)
    StoreLCDGeometry shp, 11, 65, 7, 8
    RememberShape ids, count, shp
End Sub

Private Sub StoreLCDGeometry(ByVal shp As Shape, ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single)
    ' Store thousandths as integers so the metadata is locale-independent.
    shp.Tags.Add "PT_LX", CStr(CLng(x * 1000))
    shp.Tags.Add "PT_LY", CStr(CLng(y * 1000))
    shp.Tags.Add "PT_LW", CStr(CLng(w * 1000))
    shp.Tags.Add "PT_LH", CStr(CLng(h * 1000))
End Sub

Private Sub PlaceLCDPart(ByVal shp As Shape, ByVal x As Single, ByVal y As Single, ByVal digitScale As Single)
    shp.LockAspectRatio = msoFalse
    shp.Width = CLng(ReadTag(shp, "PT_LW")) / 1000! * digitScale
    shp.Height = CLng(ReadTag(shp, "PT_LH")) / 1000! * digitScale
    shp.Left = x + CLng(ReadTag(shp, "PT_LX")) / 1000! * digitScale
    shp.Top = y + CLng(ReadTag(shp, "PT_LY")) / 1000! * digitScale
    shp.Visible = msoTrue
End Sub

Private Sub ValidateTimer(ByVal timer As Shape)
    Dim i As Long, j As Long, component As Shape
    For i = 1 To timer.GroupItems.Count
        If timer.GroupItems.Item(i).Type = msoGroup Then
            Err.Raise vbObjectError + 2052, "SlideTimer", "This timer uses the old nested layout. Run Timer_RefreshStyle."
        End If
    Next i
    For i = 1 To 6
        For j = 1 To 7
            Set component = Part(timer, "PT_Digit" & CStr(i) & "_" & Chr$(64 + j))
            If Len(ReadTag(component, "PT_LW")) = 0 Then
                Err.Raise vbObjectError + 2053, "SlideTimer", "LCD segment geometry is missing. Run Timer_RefreshStyle."
            End If
        Next j
    Next i
    For i = 1 To 2
        Set component = Part(timer, "PT_Colon" & CStr(i) & "_Top")
        Set component = Part(timer, "PT_Colon" & CStr(i) & "_Bottom")
    Next i
    Set component = Part(timer, "PT_Time")
    If component.AutoShapeType <> msoShapeOval Then
        Err.Raise vbObjectError + 2051, "SlideTimer", "Refresh the timer to update its round controls."
    End If
    Set component = Part(timer, "PT_Plus")
    Set component = Part(timer, "PT_Minus")
End Sub

Private Function DigitMask(ByVal digit As Long) As String
    Dim masks As Variant
    masks = Array("ABCDEF", "BC", "ABGED", "ABGCD", "FGBC", "AFGCD", "AFGECD", "ABC", "ABCDEFG", "ABCFGD")
    DigitMask = masks(digit)
End Function

Private Sub LayoutLCD(ByVal timer As Shape, ByVal characterCount As Long)
    Dim target As Shape, layout As String
    Dim digitScale As Single, startX As Single, startY As Single, cursor As Single, logicalWidth As Single
    Dim i As Long, j As Long, digitIndex As Long, colonIndex As Long
    Set target = Part(timer, "PT_Display")
    If characterCount > 5 Then
        layout = "00:00:00": logicalWidth = 426
    Else
        layout = "00:00": logicalWidth = 276
    End If
    digitScale = target.Width / logicalWidth
    If target.Height / 100 < digitScale Then digitScale = target.Height / 100
    startX = target.Left + (target.Width - logicalWidth * digitScale) / 2
    startY = target.Top + (target.Height - 100 * digitScale) / 2
    For i = 1 To 6
        For j = 1 To 7
            Set mSegmentShapes(i, j) = Part(timer, "PT_Digit" & CStr(i) & "_" & Chr$(64 + j))
            mSegmentShapes(i, j).Visible = msoFalse
        Next j
    Next i
    For i = 1 To 2
        Set mDotShapes(i, 1) = Part(timer, "PT_Colon" & CStr(i) & "_Top")
        Set mDotShapes(i, 2) = Part(timer, "PT_Colon" & CStr(i) & "_Bottom")
        mDotShapes(i, 1).Visible = msoFalse
        mDotShapes(i, 2).Visible = msoFalse
    Next i
    For i = 1 To Len(layout)
        If Mid$(layout, i, 1) = ":" Then
            colonIndex = colonIndex + 1
            For j = 1 To 2
                PlaceLCDPart mDotShapes(colonIndex, j), startX + cursor * digitScale, startY, digitScale
            Next j
            cursor = cursor + 22
        Else
            digitIndex = digitIndex + 1
            For j = 1 To 7
                PlaceLCDPart mSegmentShapes(digitIndex, j), startX + cursor * digitScale, startY, digitScale
            Next j
            cursor = cursor + 64
        End If
    Next i
    mLayoutCharacters = characterCount
    mPreviousDisplay = ""
    mPreviousInk = -1
End Sub

Private Sub PaintLCD(ByVal timer As Shape, ByVal text As String, ByVal ink As Long)
    Dim target As Shape, segment As Shape
    Dim i As Long, j As Long, digitIndex As Long, colonIndex As Long
    Dim character As String, previousCharacter As String, mask As String, previousMask As String
    Dim inkChanged As Boolean, firstPaint As Boolean, isOn As Boolean, wasOn As Boolean
    ' Ensure the correct object's cache is active before inspecting cache state.
    Set target = Part(timer, "PT_Display")
    If mLayoutCharacters <> Len(text) Then LayoutLCD timer, Len(text)
    If mPreviousDisplay = text And mPreviousInk = ink Then Exit Sub
    inkChanged = (mPreviousInk <> ink)
    firstPaint = (Len(mPreviousDisplay) = 0)
    For i = 1 To Len(text)
        character = Mid$(text, i, 1)
        If character = ":" Then
            colonIndex = colonIndex + 1
            If inkChanged Then
                mDotShapes(colonIndex, 1).Fill.ForeColor.RGB = ink
                mDotShapes(colonIndex, 2).Fill.ForeColor.RGB = ink
            End If
        Else
            digitIndex = digitIndex + 1
            previousCharacter = ""
            If Not firstPaint Then previousCharacter = Mid$(mPreviousDisplay, i, 1)
            If inkChanged Or previousCharacter <> character Then
                mask = DigitMask(CLng(character))
                previousMask = ""
                If Len(previousCharacter) > 0 Then previousMask = DigitMask(CLng(previousCharacter))
                For j = 1 To 7
                    Set segment = mSegmentShapes(digitIndex, j)
                    If inkChanged Then segment.Fill.ForeColor.RGB = ink
                    isOn = (InStr(1, mask, Chr$(64 + j), vbBinaryCompare) > 0)
                    wasOn = (InStr(1, previousMask, Chr$(64 + j), vbBinaryCompare) > 0)
                    If firstPaint Or (isOn Xor wasOn) Then
                        If isOn Then
                            segment.Fill.Transparency = 0
                        Else
                            segment.Fill.Transparency = 0.95
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    mPreviousDisplay = text
    mPreviousInk = ink
End Sub

Private Function AddText(ByVal sld As Slide, ByVal objectName As String, _
    ByVal text As String, ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single, ByVal size As Single, ByVal color As Long) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    shp.Name = objectName
    shp.Tags.Add "PT_ROLE", objectName
    shp.Line.Visible = msoFalse
    With shp.TextFrame
        .MarginLeft = 0: .MarginRight = 0
        .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeNone
        With .TextRange
            .text = text
            .ParagraphFormat.Alignment = ppAlignCenter
            .Font.Name = "Aptos"
            .Font.Size = size
            .Font.Color.RGB = color
        End With
    End With
    With shp.TextFrame2
        .AutoSize = msoAutoSizeNone
        .MarginLeft = 0: .MarginRight = 0
        .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    shp.Width = w
    shp.Height = h
    Set AddText = shp
End Function

Private Sub SetClick(ByVal shp As Shape, ByVal macroName As String)
    With shp.ActionSettings(ppMouseClick)
        .Action = ppActionRunMacro
        .Run = macroName
        .AnimateAction = msoFalse
    End With
End Sub

Private Sub PaintState(ByVal index As Long)
    With mStates(index)
        PaintTimer .Timer, WholeSeconds(.Clock.Remaining), .Clock.Status
        .LastPaint = WholeSeconds(.Clock.Remaining)
        .LastStatus = .Clock.Status
    End With
End Sub

Private Sub PaintStateIfChanged(ByVal index As Long, ByVal win As SlideShowWindow)
    Dim visible As Boolean
    visible = TimerIsVisible(index, win)
    With mStates(index)
        If WholeSeconds(.Clock.Remaining) <> .LastPaint Or .Clock.Status <> .LastStatus Or _
            (visible And Not .WasVisible) Then PaintState index
        .WasVisible = visible
    End With
End Sub

Private Function TimerIsVisible(ByVal index As Long, ByVal win As SlideShowWindow) As Boolean
    On Error GoTo Unavailable
    If win Is Nothing Then Exit Function
    TimerIsVisible = (win.View.Slide.SlideID = mStates(index).TimerSlide.SlideID)
Unavailable:
End Function

Private Sub PaintTimer(ByVal timer As Shape, ByVal seconds As Long, ByVal status As String)
    Dim ink As Long, iconColor As Long, units As String
    If mPartParent Is Nothing Then
        BuildPartCache timer
    ElseIf Not (timer Is mPartParent) Then
        BuildPartCache timer
    End If
    ink = RGB(39, 56, 68)
    If status = "Time's up" Then ink = RGB(115, 60, 59)
    PaintLCD timer, FormatDuration(seconds, DurationOf(timer) >= 3600), ink
    If mPreviousStatus <> status Then
        Part(timer, "PT_Status").TextFrame.TextRange.text = UCase$(status)
        Part(timer, "PT_Start").AlternativeText = "Start, pause or resume the countdown"
        mPreviousStatus = status
        PaintSoundButton timer
    End If
    units = "MIN    SEC"
    If DurationOf(timer) >= 3600 Then units = "HR   MIN   SEC"
    If mPreviousUnits <> units Then
        Part(timer, "PT_Units").TextFrame.TextRange.text = units
        mPreviousUnits = units
    End If
    iconColor = RGB(227, 238, 240)
    If status = "Running" Then iconColor = RGB(185, 228, 210)
    If mPreviousIconInk <> iconColor Then
        Part(timer, "PT_PlayIcon").Fill.ForeColor.RGB = iconColor
        Part(timer, "PT_Pause1").Fill.ForeColor.RGB = iconColor
        Part(timer, "PT_Pause2").Fill.ForeColor.RGB = iconColor
        mPreviousIconInk = iconColor
    End If
    Part(timer, "PT_Display").AlternativeText = FormatDuration(seconds) & ". " & status
    RefreshShowSurface timer
End Sub

Private Sub RefreshShowSurface(ByVal timer As Shape)
    ' Redraw only this timer's visible slide. SlideIndex is the deck index;
    ' CurrentShowPosition would jump to the wrong slide in a custom show.
    Dim win As SlideShowWindow, sld As Slide
    On Error GoTo Done
    If mRefreshing Then Exit Sub
    Set win = OwnerShowWindow()
    If win Is Nothing Then Exit Sub
    Set sld = timer.Parent
    If win.View.Slide.SlideID <> sld.SlideID Then Exit Sub
    If win.View.State <> ppSlideShowRunning Then Exit Sub
    mRefreshing = True
    win.View.GotoSlide sld.SlideIndex, msoFalse
Done:
    mRefreshing = False
End Sub

Private Sub PaintSoundButton(ByVal timer As Shape)
    Dim label As String, control As Shape
    Set control = Part(timer, "PT_Sound")
    control.AlternativeText = "End sound: " & SoundNameOf(timer) & ". Click to mute or enable. Use Timer_ChooseSound to change it."
    label = "SOUND OFF"
    If SoundEnabled(timer) Then label = "SOUND ON"
    If mPreviousSound <> label Then
        control.TextFrame.TextRange.text = label
        mPreviousSound = label
    End If
End Sub

Private Function FindTimer(ByVal sld As Slide) As Shape
    Dim shp As Shape
    For Each shp In sld.Shapes
        If ReadTag(shp, PT_MARKER) = "1" Then
            Set FindTimer = shp
            Exit Function
        End If
    Next shp
End Function

Private Function EditingSlide() As Slide
    On Error GoTo Unavailable
    If Presentations.Count = 0 Then GoTo Unavailable
    If ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide Then
        MsgBox "Switch to Normal view and select a slide first.", vbInformation, "Slide Timer"
        Exit Function
    End If
    Set EditingSlide = ActiveWindow.View.Slide
    Exit Function
Unavailable:
    MsgBox "Open a presentation and select a slide in Normal view first.", vbInformation, "Slide Timer"
End Function

Private Function CurrentTimer() As Shape
    Dim sld As Slide
    If SlideShowWindows.Count > 0 Then
        Set sld = SlideShowWindows(1).View.Slide
    Else
        Set sld = EditingSlide()
    End If
    If sld Is Nothing Then Exit Function
    Set CurrentTimer = FindTimer(sld)
End Function

Private Function DurationOf(ByVal timer As Shape) As Long
    Dim value As Long
    On Error GoTo UseDefault
    value = CLng(ReadTag(timer, PT_SECONDS))
    If value < 1 Or value > PT_MAX_SECONDS Then GoTo UseDefault
    DurationOf = value
    Exit Function
UseDefault:
    DurationOf = 300
End Function

Private Function SoundEnabled(ByVal timer As Shape) As Boolean
    SoundEnabled = (ReadTag(timer, PT_SOUND) <> "0")
End Function


' ------------ Per-timer end sound selection ------------

Private Function SoundNames() As Variant
    ' Canonical filenames matter: PowerPoint uppercases tag values.
    SoundNames = Array("Ping", "Glass", "Hero", "Submarine", "Funk", "Basso", _
        "Blow", "Bottle", "Frog", "Morse", "Pop", "Purr", "Sosumi", "Tink")
End Function

Private Function TryParseSoundChoice(ByVal text As String, ByRef choice As String) As Boolean
    Dim names As Variant, i As Long, number As Long
    text = Trim$(text)
    choice = ""
    If text = "0" Or StrComp(text, "None", vbTextCompare) = 0 Or _
        StrComp(text, "Off", vbTextCompare) = 0 Then
        choice = "None"
        TryParseSoundChoice = True
        Exit Function
    End If
    names = SoundNames()
    If Len(text) > 0 And Len(text) <= 2 Then
        If AllDigits(text) Then
            number = CLng(text)
            If number >= 1 And number <= UBound(names) + 1 Then
                choice = CStr(names(number - 1))
                TryParseSoundChoice = True
                Exit Function
            End If
        End If
    End If
    For i = 0 To UBound(names)
        If StrComp(text, CStr(names(i)), vbTextCompare) = 0 Then
            choice = CStr(names(i))
            TryParseSoundChoice = True
            Exit Function
        End If
    Next i
End Function

Private Function SoundFileForChoice(ByVal choice As String) As String
    Dim canonical As String
    If Not TryParseSoundChoice(choice, canonical) Then
        SoundFileForChoice = PT_ALERT_FILE
    ElseIf canonical <> "None" Then
        SoundFileForChoice = "/System/Library/Sounds/" & canonical & ".aiff"
    End If
End Function

Private Function SoundNameOf(ByVal timer As Shape) As String
    Dim canonical As String
    SoundNameOf = "Ping"
    If TryParseSoundChoice(ReadTag(timer, PT_ALERT_NAME), canonical) Then
        If canonical <> "None" Then SoundNameOf = canonical
    End If
End Function

Private Function SoundPathOf(ByVal timer As Shape) As String
    SoundPathOf = SoundFileForChoice(ReadTag(timer, PT_ALERT_NAME))
End Function

Private Sub ApplySoundChoice(ByVal timer As Shape, ByVal choice As String)
    Dim canonical As String
    If Not TryParseSoundChoice(choice, canonical) Then
        Err.Raise vbObjectError + 2054, "SlideTimer", "Choose a sound from the list."
    End If
    If canonical = "None" Then
        ' Retain the previous selection so the existing SOUND ON button can
        ' enable it again. A newly muted timer falls back to Ping.
        timer.Tags.Add PT_SOUND, "0"
    Else
        timer.Tags.Add PT_ALERT_NAME, canonical
        timer.Tags.Add PT_SOUND, "1"
    End If
    PaintSoundButton timer
End Sub

Private Function AskSound(ByVal initialChoice As String, ByRef choice As String) As Boolean
    Dim names As Variant, prompt As String, answer As String, candidate As String
    Dim i As Long, preview As Boolean, problem As String
    On Error GoTo Failed
    names = SoundNames()
    prompt = "Choose the end sound by number or name:" & vbCrLf & "0: No sound"
    For i = 0 To UBound(names)
        prompt = prompt & vbCrLf & CStr(i + 1) & ": " & CStr(names(i))
    Next i
    prompt = prompt & vbCrLf & vbCrLf & "Preview: type p 2, then press OK. Choose 2 to keep it." & _
        vbCrLf & "Cancel leaves this choice unchanged (or cancels a new timer)."
    mTestAlertsLeft = 0
    Do
        answer = Trim$(InputBox(prompt, "Slide Timer - end sound", initialChoice))
        If Len(answer) = 0 Then
            StopAudioPlayer
            Exit Function
        End If
        preview = (StrComp(Left$(answer, 2), "p ", vbTextCompare) = 0)
        candidate = answer
        If preview Then candidate = Trim$(Mid$(answer, 3))
        If TryParseSoundChoice(candidate, choice) Then
            If preview Then
                StopAudioPlayer
                If choice <> "None" Then PlayAlert SoundFileForChoice(choice)
                initialChoice = choice
            Else
                StopAudioPlayer
                AskSound = True
                Exit Function
            End If
        Else
            MsgBox "Choose 0-14, a listed sound name, or p followed by a choice to preview.", _
                vbExclamation, "Slide Timer"
        End If
    Loop
Failed:
    problem = Err.Description
    On Error Resume Next
    StopAudioPlayer
    On Error GoTo 0
    Err.Raise vbObjectError + 2054, "SlideTimer", problem
End Function

' ------------ Duration input and formatting ------------

Private Function NextPresetDuration(ByVal current As Long) As Long
    Dim presets As Variant, i As Long
    presets = Array(60, 120, 180, 300, 600, 900, 1200, 1800, 2700, 3600)
    For i = 0 To UBound(presets)
        If presets(i) > current Then
            NextPresetDuration = presets(i)
            Exit Function
        End If
    Next i
    NextPresetDuration = presets(0)
End Function

Private Function AskDuration(ByVal initialSeconds As Long, ByRef seconds As Long) As Boolean
    Dim answer As String
    Do
        answer = InputBox("Enter minutes, mm:ss, or hh:mm:ss." & vbCrLf & _
            "Examples: 5 = five minutes; 0:10 = ten seconds; 1:30 = ninety seconds." & vbCrLf & _
            "Range: 0:01 to 23:59:59. Cancel keeps the current setting.", _
            "Slide Timer - duration", FormatDuration(initialSeconds))
        If Len(answer) = 0 Then Exit Function
        If TryParseDuration(answer, seconds) Then
            AskDuration = True
            Exit Function
        End If
        MsgBox "Enter a positive whole number of minutes, mm:ss, or hh:mm:ss." & _
            vbCrLf & "Seconds, and minutes after an hour, must be 00-59." & _
            vbCrLf & "The maximum duration is 23:59:59.", vbExclamation, "Slide Timer"
    Loop
End Function

Private Function TryParseDuration(ByVal text As String, ByRef seconds As Long) As Boolean
    Dim parts() As String, i As Long, n As Long, total As Long
    On Error GoTo Invalid
    text = Trim$(text)
    If Len(text) = 0 Or Len(text) > 20 Then Exit Function
    parts = Split(text, ":")
    If UBound(parts) > 2 Then Exit Function
    For i = 0 To UBound(parts)
        If Not AllDigits(parts(i)) Then Exit Function
        If Len(parts(i)) > 5 Then Exit Function
        n = CLng(parts(i))
        If i > 0 And n > 59 Then Exit Function
        total = total * 60 + n
        If total > PT_MAX_SECONDS Then Exit Function
    Next i
    If UBound(parts) = 0 Then total = total * 60
    If total < 1 Or total > PT_MAX_SECONDS Then Exit Function
    seconds = total
    TryParseDuration = True
Invalid:
End Function

Private Function AllDigits(ByVal text As String) As Boolean
    Dim i As Long, character As String
    If Len(text) = 0 Then Exit Function
    For i = 1 To Len(text)
        character = Mid$(text, i, 1)
        If character < "0" Or character > "9" Then Exit Function
    Next i
    AllDigits = True
End Function

Private Function FormatDuration(ByVal seconds As Long, Optional ByVal forceHours As Boolean = False) As String
    If seconds < 0 Then seconds = 0
    If seconds >= 3600 Or forceHours Then
        FormatDuration = Format$(seconds \ 3600, "00") & ":" & _
            Format$((seconds Mod 3600) \ 60, "00") & ":" & Format$(seconds Mod 60, "00")
    Else
        FormatDuration = Format$(seconds \ 60, "00") & ":" & Format$(seconds Mod 60, "00")
    End If
End Function

Private Function WholeSeconds(ByVal remaining As Double) As Long
    ' Ceiling avoids showing 00:00 before the deadline. Never round to nearest.
    If remaining <= 0 Then
        WholeSeconds = 0
    Else
        WholeSeconds = CLng(-Int(-remaining))
    End If
End Function

Private Sub ReportError(ByVal action As String, ByVal detail As String)
    MsgBox "Slide Timer " & PT_VERSION & " could not " & action & "." & vbCrLf & vbCrLf & detail, _
        vbExclamation, "Slide Timer"
End Sub

' ------------ Optional native diagnostics ------------

Public Sub Timer_CheckCurrent()
    Dim timer As Shape
    On Error GoTo Failed
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before checking the timer.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set timer = CurrentTimer()
    If timer Is Nothing Then
        MsgBox "No timer on this slide. Run Timer_Insert first.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    ValidateTimer timer
    If Len(mLastProblem) > 0 Then
        MsgBox "Last background problem: " & mLastProblem & vbCrLf & _
            "Fix the problem, then use the timer controls or run Timer_TestSound.", vbExclamation, "Slide Timer"
        mLastProblem = ""
        Exit Sub
    End If
    PaintState StateFor(timer)
    MsgBox "Slide Timer " & PT_VERSION & ": the flat LCD layout, all 42 segments and round controls passed." & _
        vbCrLf & "Test a 0:05 countdown in Slide Show view next.", vbInformation, "Slide Timer"
    Exit Sub
Failed:
    ReportError "check the current timer", Err.Description
End Sub

Public Sub Timer_RunSelfTests()
    Dim total As Long, count As Long, first As Double, state As CountdownState, finishedNow As Boolean, encoded() As Byte, soundChoice As String
    On Error GoTo Failed
    AssertValue TryParseDuration("5", total) And total = 300, "minutes", count
    AssertValue TryParseDuration("1:30", total) And total = 90, "mm:ss", count
    AssertValue TryParseDuration("0:01", total) And total = 1, "one second", count
    AssertValue TryParseDuration("1:02:03", total) And total = 3723, "hh:mm:ss", count
    AssertValue TryParseDuration("23:59:59", total) And total = 86399, "maximum", count
    AssertValue TryParseDuration(" 10 ", total) And total = 600, "outer spaces", count
    AssertValue Not TryParseDuration("0", total), "reject zero", count
    AssertValue Not TryParseDuration("-1", total), "reject negative", count
    AssertValue Not TryParseDuration("1.5", total), "reject decimal", count
    AssertValue Not TryParseDuration("1e2", total), "reject exponent", count
    AssertValue Not TryParseDuration("1:60", total), "reject seconds 60", count
    AssertValue Not TryParseDuration("1:60:00", total), "reject minutes 60", count
    AssertValue Not TryParseDuration("24:00:00", total), "reject over maximum", count
    AssertValue Not TryParseDuration("1::2", total), "reject empty segment", count
    AssertValue Not TryParseDuration("1:2:3:4", total), "reject extra segment", count
    AssertValue Not TryParseDuration("999999999999999999999", total), "reject overflow", count
    AssertValue Not TryParseDuration("", total), "reject empty", count
    AssertValue FormatDuration(0) = "00:00", "format zero", count
    AssertValue FormatDuration(90) = "01:30", "format minutes", count
    AssertValue FormatDuration(3600) = "01:00:00", "format hours", count
    AssertValue FormatDuration(3599, True) = "00:59:59", "keep hours visible", count
    AssertValue DigitMask(0) = "ABCDEF", "LCD zero", count
    AssertValue DigitMask(1) = "BC", "LCD one", count
    AssertValue DigitMask(8) = "ABCDEFG", "LCD eight", count
    AssertValue Len(SegmentPath(7)) > 0, "LCD segment geometry", count
    AssertValue WholeSeconds(0.01) = 1, "ceiling near zero", count
    AssertValue WholeSeconds(1.01) = 2, "ceiling above one", count
    AssertValue WholeSeconds(1#) = 1, "ceiling exact", count
    AssertValue WholeSeconds(-1#) = 0, "clamp negative", count
    ResetCountdown state, 300
    StartCountdown state, 300, 1000
    finishedNow = AdvanceCountdown(state, 1012.5)
    AssertValue Not finishedNow And state.Remaining = 287.5, "elapsed off-slide", count
    StartCountdown state, 300, 1015
    AssertValue state.Deadline = 1300, "duplicate start keeps deadline", count
    PauseCountdown state, 1020.25
    AssertValue state.Status = "Paused" And state.Remaining = 279.75, "fractional pause", count
    finishedNow = AdvanceCountdown(state, 1100)
    AssertValue state.Remaining = 279.75, "paused time stays frozen", count
    StartCountdown state, 300, 1100
    AssertValue state.Deadline = 1379.75, "resume from remaining", count
    finishedNow = AdvanceCountdown(state, 1400)
    AssertValue finishedNow And state.Status = "Time's up" And state.Remaining = 0, "finish while away", count
    AssertValue Not AdvanceCountdown(state, 1500), "completion happens once", count
    StartCountdown state, 300, 1500
    AssertValue state.Deadline = 1800, "restart finished timer", count
    ResetCountdown state, 60
    StartCountdown state, 60, 2000
    PauseCountdown state, 2000.01
    AssertValue state.Status = "Paused" And state.Remaining > 59, "pause before first paint", count
    AdjustCountdown state, 60, 120, 2001
    AssertValue state.Status = "Paused" And state.Remaining > 119, "adjust paused timer", count
    ResetCountdown state, 60
    StartCountdown state, 60, 2100
    finishedNow = AdvanceCountdown(state, 2110)
    AdjustCountdown state, -60, 1, 2110
    AssertValue state.Remaining = 1 And state.Deadline = 2111, "running adjustment clamps", count
    ResetCountdown state, 300
    AssertValue state.Status = "Ready" And state.Remaining = 300, "explicit reset", count
    encoded = UTF8CString("Ping")
    AssertValue UBound(encoded) = 4 And encoded(0) = 80 And encoded(4) = 0, "ASCII argument terminator", count
    encoded = UTF8CString(ChrW(&HE9))
    AssertValue UBound(encoded) = 2 And encoded(0) = &HC3 And encoded(1) = &HA9 And encoded(2) = 0, "UTF-8 filename", count
    encoded = UTF8CString(ChrW(&HD83D) & ChrW(&HDD14))
    AssertValue UBound(encoded) = 4 And encoded(0) = &HF0 And encoded(1) = &H9F And _
        encoded(2) = &H94 And encoded(3) = &H94 And encoded(4) = 0, "UTF-8 surrogate pair", count
    AssertValue TryParseSoundChoice("1", soundChoice) And soundChoice = "Ping", "sound number", count
    AssertValue TryParseSoundChoice("3", soundChoice) And soundChoice = "Hero", "Hero selection", count
    AssertValue TryParseSoundChoice("14", soundChoice) And soundChoice = "Tink", "last sound", count
    AssertValue TryParseSoundChoice(" gLaSs ", soundChoice) And soundChoice = "Glass", "canonical sound name", count
    AssertValue TryParseSoundChoice("0", soundChoice) And soundChoice = "None", "no sound", count
    AssertValue TryParseSoundChoice("OFF", soundChoice) And soundChoice = "None", "mute alias", count
    AssertValue Not TryParseSoundChoice("-1", soundChoice), "reject negative sound", count
    AssertValue Not TryParseSoundChoice("15", soundChoice), "reject unknown sound number", count
    AssertValue Not TryParseSoundChoice("1.5", soundChoice), "reject fractional sound", count
    AssertValue Not TryParseSoundChoice("999999999999", soundChoice), "reject sound overflow", count
    AssertValue Not TryParseSoundChoice("../Hero", soundChoice), "reject unlisted sound path", count
    AssertValue SoundFileForChoice("HERO") = "/System/Library/Sounds/Hero.aiff", "uppercase tag filename", count
    AssertValue SoundFileForChoice("None") = "", "muted sound path", count
    AssertValue SoundFileForChoice("") = PT_ALERT_FILE, "legacy sound fallback", count
    first = ClockSeconds()
    YieldBriefly
    AssertValue ClockSeconds() > first, "native clock and sleep", count
    MsgBox CStr(count) & " checks passed." & vbCrLf & _
        "Next: run Timer_TestSound, then test a 0:05 timer in Slide Show view.", _
        vbInformation, "Slide Timer diagnostics"
    Exit Sub
Failed:
    MsgBox "A diagnostic check failed: " & Err.Description, vbExclamation, "Slide Timer diagnostics"
End Sub

Private Sub AssertValue(ByVal condition As Boolean, ByVal label As String, ByRef count As Long)
    If Not condition Then Err.Raise vbObjectError + 2048, "SlideTimer", label
    count = count + 1
End Sub
