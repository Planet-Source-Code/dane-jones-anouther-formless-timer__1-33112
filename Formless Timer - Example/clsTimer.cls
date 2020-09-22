VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'** This Class requires the presence of the basTimer Module _
    that was provided with the example from PSC.

'** Define Constants
    Private Const lngKeyCode = &H8563256  '<--- Used to help protect against user from calling automated procedures
    '** Why The KeyCode? _
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

'** Define Variables
    Private lngInterval As Long
    Private bolEnable As Boolean

'** Define Events
    Public Event Timer()

'** Declare DLL Function Calls
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Sub Class_Initialize()
    '** Initilize default values
        bolEnable = False
        lngInterval = 1000
End Sub

Private Sub Class_Terminate()
    '** Clean Up
        Enable = False
End Sub

'** Interval Property
    Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer's Timer event."
        Interval = lngInterval
    End Property

    Public Property Let Interval(ByVal vNewValue As Long)
        lngInterval = vNewValue
    End Property

'** Enable Property
    Public Property Get Enable() As Boolean
Attribute Enable.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
        Enable = bolEnable
    End Property
    
    Public Property Let Enable(ByVal vNewValue As Boolean)
        Static tmrID As Long
        
        '** If supplied Value is the same then cancel assignment
            If vNewValue = bolEnable Then Exit Property
            
        '** Assign to Variable
            bolEnable = vNewValue
        
        '** Enable or Disable timer
            If vNewValue Then
                '** Create Timer
                    tmrID = SetTimer(0, 0, lngInterval, AddressOf TimerProc)
                    
                '** Create Reference for this timer class
                    AddClassRef Me, tmrID, lngKeyCode
            Else
                '** Destroy Active Timer
                    KillTimer 0, tmrID
                    
                '** Remove Reference forthis timer class
                    RemoveClassRef tmrID, lngKeyCode
            End If
        
    End Property

Public Sub Timer_Event(Optional KeyCode As Long)
Attribute Timer_Event.VB_Description = "Special CallBack Procedure - DO NOT USE!"
Attribute Timer_Event.VB_MemberFlags = "40"
    '** HIDDEN PROCEDURE: This Procedure is called when the timers event fires
    '** Validate KeyCode
        If lngKeyCode = KeyCode Then
            Debug.Print "Success!"
            RaiseEvent Timer
        End If
End Sub
