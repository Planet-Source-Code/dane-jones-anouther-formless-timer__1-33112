Attribute VB_Name = "basTimer"
Option Explicit

'** This Module is required to use the clsTimer

'** All Public Procedures are Hidden for limited protection **

'** Define Constants
    Private Const lngKeyCode = &H8563256  '<--- Used to help protect against user from calling automated procedures
    '** Why the KeyCode? _
        Well actually I had to put this into place to protect you. _
        You see since certain callback procedures require Public _
        declaration they would also be available to any application _
        that creates your ActiveX object and that is asking for _
        trouble.  Granted you can hide the procedures from the Calling _
        apps interface but the procedures can still be called.  Not to _
        mention all the programer has to do is select the "Show hidden _
        members" option to view these hidden procedures.  So if you _
        want to protect your code and yourself I recomend changing the _
        KeyCode here and in basTimer.

'** Define Objects
    Private ClassCollection As New Collection

Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTimer As Long)
Attribute TimerProc.VB_Description = "Special CallBack Procedure - DO NOT USE!"
Attribute TimerProc.VB_MemberFlags = "40"
    Dim CallingClass As clsTimer
    On Error Resume Next
    Debug.Print "Timer was called: " & idEvent
    
    '** Test Perameter Validity
        If hwnd = 0 Then
        '** Get Timer Class Reference that created this timer event
            Set CallingClass = ClassCollection.Item(CStr(idEvent))
            
            If Err.Number = 0 Then
                '** Call Timer_Event in Class object
                    CallByName CallingClass, "Timer_Event", VbMethod, lngKeyCode
            End If
        End If
End Sub

Public Sub AddClassRef(ByRef CallingClass As clsTimer, TimerID As Long, Optional KeyCode As Long)
Attribute AddClassRef.VB_Description = "Special CallBack Procedure - DO NOT USE!"
Attribute AddClassRef.VB_MemberFlags = "40"
    '** Validate KeyCode
        If lngKeyCode = KeyCode Then
            '** Add Reference of Calling Class Collection
                ClassCollection.Add CallingClass, CStr(TimerID)
        End If
End Sub

Public Sub RemoveClassRef(TimerID As Long, Optional KeyCode As Long)
Attribute RemoveClassRef.VB_Description = "Special CallBack Procedure - DO NOT USE!"
Attribute RemoveClassRef.VB_MemberFlags = "40"
    '** Validate KeyCode
        If lngKeyCode = KeyCode Then
            '** Remove Reference of Calling Class Collection
                ClassCollection.Remove CStr(TimerID)
        End If
End Sub
