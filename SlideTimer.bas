Attribute VB_Name = "SlideTimer"
Option Explicit

' Slide Timer 1.1.9 | PowerPoint for Mac, VBA7 / 64-bit
' Import this module into a presentation and save the presentation as .pptm.
' Run Timer_Insert in Normal view. Click play/pause in Slide Show view.
' No network, AppleScript, ActiveX, or third-party dependencies at runtime.
' The only file read is the macOS system sound used for the end alert.

#If Mac Then
    Private Declare PtrSafe Function PT_MonotonicTime _
        Lib "/System/Library/Frameworks/QuartzCore.framework/QuartzCore" _
        Alias "CACurrentMediaTime" () As Double
    Private Declare PtrSafe Function PT_USleep Lib "/usr/lib/libSystem.B.dylib" _
        Alias "usleep" (ByVal microseconds As Long) As Long
    ' VBA Beep is silent in PowerPoint for Mac, so the end alert loads a
    ' system sound file and plays it through the macOS sound service.
    Private Declare PtrSafe Sub PT_PlayAlertSound _
        Lib "/System/Library/Frameworks/AudioToolbox.framework/AudioToolbox" _
        Alias "AudioServicesPlayAlertSound" (ByVal soundID As Long)
    Private Declare PtrSafe Function PT_CreateSystemSoundID _
        Lib "/System/Library/Frameworks/AudioToolbox.framework/AudioToolbox" _
        Alias "AudioServicesCreateSystemSoundID" (ByVal fileURL As LongPtr, ByRef soundID As Long) As Long
    Private Declare PtrSafe Function PT_CFStringCreate _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFStringCreateWithCString" (ByVal allocator As LongPtr, ByVal text As String, ByVal encoding As Long) As LongPtr
    Private Declare PtrSafe Function PT_CFURLCreate _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFURLCreateWithFileSystemPath" (ByVal allocator As LongPtr, ByVal filePath As LongPtr, _
        ByVal pathStyle As LongPtr, ByVal isDirectory As Byte) As LongPtr
    Private Declare PtrSafe Sub PT_CFRelease _
        Lib "/System/Library/Frameworks/CoreFoundation.framework/CoreFoundation" _
        Alias "CFRelease" (ByVal reference As LongPtr)
#End If

Private Const PT_MARKER As String = "PT_TIMER"
Private Const PT_SECONDS As String = "PT_SECONDS"
Private Const PT_SOUND As String = "PT_SOUND"
Private Const PT_VERSION As String = "1.1.9"
Private Const PT_MAX_SECONDS As Long = 86399
Private Const PT_ALERT_FILE As String = "/System/Library/Sounds/Hero.aiff"
Private Const PT_ALERT_REPEATS As Long = 3
Private Const PT_ALERT_GAP As Double = 1.2

Private mLoopActive As Boolean
Private mRunning As Boolean
Private mFinished As Boolean
Private mConfiguring As Boolean
Private mRemaining As Double
Private mDeadline As Double
Private mLastPaint As Long
Private mTimer As Shape
Private mSlide As Slide
Private mWindow As SlideShowWindow
Private mAlertSound As Long
Private mPausedTimer As Shape
Private mPausedRemaining As Double
Private mAlertsLeft As Long
Private mNextAlert As Double

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
    Dim sld As Slide, timer As Shape, duration As Long
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
        Timer_Configure
        Exit Sub
    End If
    If Not AskDuration(300, duration) Then Exit Sub
    Set timer = CreateTimer(sld, duration)
    timer.Select
    Exit Sub
Failed:
    ReportError "insert the timer", Err.Description
End Sub

Public Sub Timer_Configure()
    Dim timer As Shape, duration As Long, wasRunning As Boolean
    Dim active As Boolean, accepted As Boolean, problem As String
    If mConfiguring Then Exit Sub
    On Error GoTo Failed
    Set timer = CurrentTimer()
    If timer Is Nothing Then
        MsgBox "This slide has no timer. Run Timer_Insert first.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    mConfiguring = True
    active = IsActiveTimer(timer)
    If active Then
        CaptureRemaining
        wasRunning = mRunning
        mRunning = False
    End If
    If SlideShowWindows.Count > 0 Then
        ' A dialog opened during a slide show sits behind the full-screen
        ' show and cannot receive input, so SET steps through presets here.
        duration = NextPresetDuration(DurationOf(timer))
        accepted = True
    Else
        accepted = AskDuration(DurationOf(timer), duration)
    End If
    If accepted Then
        timer.Tags.Add PT_SECONDS, CStr(duration)
        If IsPausedTimer(timer) Then ClearPaused
        If active Then
            mRemaining = duration
            mRunning = False
            mFinished = False
        End If
        PaintTimer timer, duration, "Ready"
    ElseIf active Then
        ' Cancel freezes the elapsed time spent in the dialog.
        mRunning = wasRunning
        If mRunning Then mDeadline = ClockSeconds() + mRemaining
        PaintActiveTimer
    End If
    mConfiguring = False
    Exit Sub
Failed:
    problem = Err.Description
    mConfiguring = False
    If active Then
        mRunning = False
        On Error Resume Next
        PaintActiveTimer
        On Error GoTo 0
    End If
    ReportError "configure the timer", problem
End Sub

Public Sub Timer_Toggle()
    Dim timer As Shape, win As SlideShowWindow, problem As String
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    If SlideShowWindows.Count = 0 Then
        MsgBox "Start Slide Show view, then click the timer's play/pause button.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set win = SlideShowWindows(1)
    Set timer = FindTimer(win.View.Slide)
    If timer Is Nothing Then Exit Sub
    If ReadTag(timer, "PT_LAYOUT") <> "FLAT" Then
        MsgBox "Exit the show and run Timer_RefreshStyle to update this timer.", vbInformation, "Slide Timer"
        Exit Sub
    End If

    ' DoEvents can dispatch another click while RunTimerLoop is running.
    ' That click only changes state. It never creates a second loop.
    If mLoopActive Then
        If Not IsActiveTimer(timer) Then Exit Sub
        If mRunning Then
            CaptureRemaining
            mRunning = False
        Else
            If mFinished Or mRemaining <= 0 Then mRemaining = DurationOf(timer)
            mFinished = False
            mDeadline = ClockSeconds() + mRemaining
            mRunning = True
        End If
        PaintActiveTimer
        Exit Sub
    End If

    Set mTimer = timer
    Set mWindow = win
    Set mSlide = win.View.Slide
    mRemaining = DurationOf(timer)
    If IsPausedTimer(timer) Then mRemaining = mPausedRemaining
    ClearPaused
    PaintTimer timer, WholeSeconds(mRemaining), "Ready"
    mDeadline = ClockSeconds() + mRemaining
    mFinished = False
    mRunning = True
    mLoopActive = True
    mLastPaint = -1
    PaintActiveTimer
    RunTimerLoop
    Exit Sub
Failed:
    problem = Err.Description
    StopTimerSession
    ReportError "run the timer", problem
End Sub

Public Sub Timer_Reset()
    Dim timer As Shape
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    If IsActiveTimer(timer) Then
        mRunning = False
        mFinished = False
        mRemaining = DurationOf(timer)
    End If
    If IsPausedTimer(timer) Then ClearPaused
    PaintTimer timer, DurationOf(timer), "Ready"
    Exit Sub
Failed:
    ReportError "reset the timer", Err.Description
End Sub

Public Sub Timer_ToggleSound()
    Dim timer As Shape
    On Error GoTo Failed
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    If SoundEnabled(timer) Then
        timer.Tags.Add PT_SOUND, "0"
    Else
        timer.Tags.Add PT_SOUND, "1"
    End If
    PaintSoundButton timer
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
    Dim timer As Shape, oldDuration As Long, newDuration As Long, actualChange As Long
    On Error GoTo Failed
    If mConfiguring Then Exit Sub
    Set timer = CurrentTimer()
    If timer Is Nothing Then Exit Sub
    oldDuration = DurationOf(timer)
    newDuration = oldDuration + change
    If newDuration < 1 Then newDuration = 1
    If newDuration > PT_MAX_SECONDS Then newDuration = PT_MAX_SECONDS
    actualChange = newDuration - oldDuration
    If actualChange = 0 Then Exit Sub
    If IsActiveTimer(timer) Then
        CaptureRemaining
        If mFinished Then
            mRemaining = newDuration
        Else
            mRemaining = mRemaining + actualChange
            If mRemaining < 1 Then mRemaining = 1
            If mRemaining > PT_MAX_SECONDS Then mRemaining = PT_MAX_SECONDS
        End If
        mFinished = False
        If mRunning Then mDeadline = ClockSeconds() + mRemaining
        timer.Tags.Add PT_SECONDS, CStr(newDuration)
        PaintActiveTimer
    ElseIf IsPausedTimer(timer) Then
        mPausedRemaining = mPausedRemaining + actualChange
        If mPausedRemaining < 1 Then mPausedRemaining = 1
        If mPausedRemaining > PT_MAX_SECONDS Then mPausedRemaining = PT_MAX_SECONDS
        timer.Tags.Add PT_SECONDS, CStr(newDuration)
        PaintTimer timer, WholeSeconds(mPausedRemaining), "Paused"
    Else
        timer.Tags.Add PT_SECONDS, CStr(newDuration)
        PaintTimer timer, newDuration, "Ready"
    End If
    Exit Sub
Failed:
    ReportError "adjust the duration", Err.Description
End Sub

Public Sub Timer_RefreshStyle()
    Dim sld As Slide, oldTimer As Shape, fresh As Shape, duration As Long
    Dim oldLeft As Single, oldTop As Single, oldWidth As Single, soundWasOn As Boolean
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
    oldLeft = oldTimer.Left: oldTop = oldTimer.Top: oldWidth = oldTimer.Width
    Set fresh = CreateTimer(sld, duration)
    fresh.LockAspectRatio = msoTrue
    fresh.Width = oldWidth
    fresh.Left = oldLeft
    fresh.Top = oldTop
    If Not soundWasOn Then fresh.Tags.Add PT_SOUND, "0"
    PaintTimer fresh, duration, "Ready"
    oldTimer.Delete
    fresh.Name = "SlideTimer_" & CStr(sld.SlideID)
    fresh.Select
    Exit Sub
Failed:
    ReportError "update the timer's appearance", Err.Description
End Sub

Public Sub Timer_TestSound()
    ' The same alert played at zero: Hero, three times, at the alert volume.
    Dim i As Long, ignored As Long
    For i = 1 To PT_ALERT_REPEATS
#If Mac Then
        If i > 1 Then ignored = PT_USleep(CLng(PT_ALERT_GAP * 1000000#))
#End If
        PlayAlert
    Next i
End Sub

Public Sub Timer_Remove()
    Dim timer As Shape
    On Error GoTo Failed
    If SlideShowWindows.Count > 0 Then
        MsgBox "Exit the slide show before removing a timer.", vbInformation, "Slide Timer"
        Exit Sub
    End If
    Set timer = CurrentTimer()
    If Not timer Is Nothing Then timer.Delete
    Exit Sub
Failed:
    ReportError "remove the timer", Err.Description
End Sub

Public Sub Timer_Help()
    MsgBox "SLIDE TIMER " & PT_VERSION & vbCrLf & vbCrLf & _
        "Normal view: run Timer_Insert to add a timer to the current slide." & vbCrLf & _
        "Set a duration with Timer_Configure." & vbCrLf & vbCrLf & _
        "Slide Show view:" & vbCrLf & _
        "Play/pause or countdown: start, pause, or resume." & vbCrLf & _
        "Reset: restore the duration and wait for play/pause." & vbCrLf & _
        "SET: next preset (1-60 min). Plus/minus: adjust one minute." & vbCrLf & _
        "Sound: enable or mute the end alert." & vbCrLf & vbCrLf & _
        "Pause releases the show: Esc and slide navigation work normally." & vbCrLf & _
        "Leaving the slide while it runs resets the timer." & vbCrLf & _
        "Save as .pptm to keep the code. Import this module into each deck.", _
        vbInformation, "Slide Timer"
End Sub

' ------------ Timer engine ------------

Private Sub RunTimerLoop()
    Dim problem As String, nextVisibilityCheck As Double, idle As Boolean
    On Error GoTo Failed
    Do While mLoopActive
        DoEvents
        If ClockSeconds() >= nextVisibilityCheck Then
            If Not TimerStillVisible() Then Exit Do
            nextVisibilityCheck = ClockSeconds() + 0.15
        End If
        If Not mConfiguring Then
            If mRunning Then
                CaptureRemaining
                If WholeSeconds(mRemaining) <> mLastPaint Then
                    PaintActiveTimer
                End If
            End If
        End If
        If mAlertsLeft > 0 Then ContinueAlert
        ' A running macro holds up slide navigation and Esc. Whenever the
        ' countdown is not running (paused, reset or finished) and no alert is
        ' pending, end the loop and leave the slide to PowerPoint.
        If Not mRunning And mAlertsLeft = 0 And Not mConfiguring Then
            idle = True
            Exit Do
        End If
        YieldBriefly
    Loop
    If idle Then
        EndIdleSession
    Else
        StopTimerSession
    End If
    Exit Sub
Failed:
    problem = Err.Description
    StopTimerSession
    ReportError "update the timer", problem
End Sub

Private Sub CaptureRemaining()
    If mRunning Then
        mRemaining = mDeadline - ClockSeconds()
        If mRemaining <= 0 Then
            mRemaining = 0
            mRunning = False
            mFinished = True
            PaintActiveTimer
            ' Commit completion before the alert, even if a pause/settings
            ' click discovers that the deadline has just passed.
            If SoundEnabled(mTimer) Then StartAlert
        End If
    End If
End Sub

Private Sub StopTimerSession()
    ' Cleanup also runs on Esc, navigation, or a deleted/closed slide.
    ' Only visual text changes here. The configured duration stays in tags.
    mRunning = False
    mLoopActive = False
    mFinished = False
    mAlertsLeft = 0
    mConfiguring = False
    ClearPaused
    On Error Resume Next
    If Not mTimer Is Nothing Then PaintTimer mTimer, DurationOf(mTimer), "Ready"
    Set mTimer = Nothing
    Set mSlide = Nothing
    Set mWindow = Nothing
    On Error GoTo 0
End Sub

Private Sub EndIdleSession()
    ' Like StopTimerSession, but keep the display as it is. A paused time is
    ' remembered so the next play/pause click resumes it; after "Time's up"
    ' or a reset, the next click starts from the configured duration.
    If Not mFinished And mRemaining > 0 And WholeSeconds(mRemaining) < DurationOf(mTimer) Then
        Set mPausedTimer = mTimer
        mPausedRemaining = mRemaining
    Else
        ClearPaused
    End If
    mRunning = False
    mLoopActive = False
    mFinished = False
    mAlertsLeft = 0
    Set mTimer = Nothing
    Set mSlide = Nothing
    Set mWindow = Nothing
End Sub

Private Function TimerStillVisible() As Boolean
    On Error GoTo Gone
    If mWindow Is Nothing Then Exit Function
    If mSlide Is Nothing Then Exit Function
    If SlideShowWindows.Count = 0 Then Exit Function
    If mWindow.View.State = ppSlideShowDone Then Exit Function
    If mWindow.View.Slide.SlideID <> mSlide.SlideID Then Exit Function
    ' Accessing the live shape also detects deletion or a closed presentation.
    If ReadTag(mTimer, PT_MARKER) <> "1" Then Exit Function
    TimerStillVisible = True
Gone:
End Function

Private Function ClockSeconds() As Double
#If Mac Then
    ' Monotonic time avoids the one-second Mac VBA Timer resolution and
    ' remains independent of changes to the system clock or midnight.
    ClockSeconds = PT_MonotonicTime()
#Else
    ' Secondary fallback only. The supported target for this package is Mac.
    ClockSeconds = CDbl(Now) * 86400#
#End If
End Function

Private Sub YieldBriefly()
#If Mac Then
    Dim ignored As Long
    ' Sleep for 30 ms, then yield through DoEvents on the next iteration.
    ' This avoids a busy spin while keeping click/keyboard controls responsive.
    ignored = PT_USleep(30000)
#End If
End Sub

Private Sub StartAlert()
    ' Play the first alert now. RunTimerLoop plays the rest, so the controls
    ' stay responsive between repeats.
    PlayAlert
    mAlertsLeft = PT_ALERT_REPEATS - 1
    mNextAlert = ClockSeconds() + PT_ALERT_GAP
End Sub

Private Sub ContinueAlert()
    ' Reset, restart or mute during the alert cancels the remaining repeats.
    If Not mFinished Or Not SoundEnabled(mTimer) Then
        mAlertsLeft = 0
    ElseIf ClockSeconds() >= mNextAlert Then
        PlayAlert
        mAlertsLeft = mAlertsLeft - 1
        mNextAlert = ClockSeconds() + PT_ALERT_GAP
    End If
End Sub

Private Sub PlayAlert()
#If Mac Then
    ' Played at the alert volume from System Settings > Sound.
    On Error GoTo Fallback
    PT_PlayAlertSound AlertSoundID()
    Exit Sub
Fallback:
#End If
    Beep
End Sub

#If Mac Then
Private Function AlertSoundID() As Long
    Dim path As LongPtr, url As LongPtr, created As Long
    If mAlertSound = 0 Then
        On Error GoTo Done
        ' &H8000100 = kCFStringEncodingUTF8; 0 = POSIX path style.
        path = PT_CFStringCreate(0, PT_ALERT_FILE, &H8000100)
        If path <> 0 Then url = PT_CFURLCreate(0, path, 0, 0)
        If url <> 0 Then
            If PT_CreateSystemSoundID(url, created) = 0 Then mAlertSound = created
        End If
Done:
        If url <> 0 Then PT_CFRelease url
        If path <> 0 Then PT_CFRelease path
    End If
    ' Without the sound file, fall back to the user's chosen alert sound
    ' (kSystemSoundID_UserPreferredAlert).
    AlertSoundID = mAlertSound
    If AlertSoundID = 0 Then AlertSoundID = &H1000
End Function
#End If

Private Function IsPausedTimer(ByVal timer As Shape) As Boolean
    On Error GoTo Different
    If mPausedTimer Is Nothing Or timer Is Nothing Then Exit Function
    If timer Is mPausedTimer Then IsPausedTimer = True
Different:
End Function

Private Sub ClearPaused()
    Set mPausedTimer = Nothing
    mPausedRemaining = 0
End Sub

Private Function IsActiveTimer(ByVal timer As Shape) As Boolean
    On Error GoTo Different
    If Not mLoopActive Then Exit Function
    If mTimer Is Nothing Or timer Is Nothing Then Exit Function
    If timer Is mTimer Then IsActiveTimer = True
Different:
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

Private Sub PaintActiveTimer()
    Dim status As String
    If mFinished Then
        status = "Time's up"
    ElseIf mRunning Then
        status = "Running"
    ElseIf mRemaining = DurationOf(mTimer) Then
        status = "Ready"
    Else
        status = "Paused"
    End If
    PaintTimer mTimer, WholeSeconds(mRemaining), status
End Sub

Private Sub PaintTimer(ByVal timer As Shape, ByVal seconds As Long, ByVal status As String)
    Dim ink As Long, iconColor As Long, units As String
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
    If IsActiveTimer(timer) Then
        mLastPaint = seconds
        RefreshShowSurface timer
    End If
End Sub

Private Sub RefreshShowSurface(ByVal timer As Shape)
    ' PowerPoint for Mac does not repaint a running slide show after shape
    ' changes made from VBA; the LCD would only update on the next click.
    ' Re-entering the current slide without resetting it forces a redraw.
    Dim position As Long
    On Error Resume Next
    If mWindow Is Nothing Or mSlide Is Nothing Then Exit Sub
    ' Never pull the show back if the presenter has just moved on.
    If mWindow.View.Slide.SlideID <> mSlide.SlideID Then Exit Sub
    position = mWindow.View.CurrentShowPosition
    If position > 0 Then mWindow.View.GotoSlide position, msoFalse
End Sub

Private Sub PaintSoundButton(ByVal timer As Shape)
    Dim label As String, control As Shape
    Set control = Part(timer, "PT_Sound")
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
    PaintTimer timer, DurationOf(timer), "Ready"
    MsgBox "Slide Timer " & PT_VERSION & ": the flat LCD layout, all 42 segments and round controls passed." & _
        vbCrLf & "Test a 0:05 countdown in Slide Show view next.", vbInformation, "Slide Timer"
    Exit Sub
Failed:
    ReportError "check the current timer", Err.Description
End Sub

Public Sub Timer_RunSelfTests()
    Dim total As Long, count As Long, first As Double
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
