VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   3705
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Individual Timer Object"
      Height          =   3615
      Left            =   4920
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdTimerObj 
         Caption         =   "Turn Timer On"
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   3000
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Height          =   2595
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Timer Array"
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.ListBox List1 
         Height          =   2595
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Turn Timer 1 Off"
         Height          =   495
         Index           =   1
         Left            =   0
         TabIndex        =   3
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Turn Timer 2 Off"
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   2
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton cmd 
         Caption         =   "Turn Timer 3 Off"
         Height          =   495
         Index           =   3
         Left            =   3120
         TabIndex        =   1
         Top             =   2880
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'CTimer is a base Api timer class that also supports being used as the base
'class as an array of Ctimer classes.
'
'Ctimers wraps the functionality of CTimer into an array type object so that
'you can manage and receive events from multiple CTimer objects through the
'events of one object, accessing each class by array index.
'
'VB does not support
' dim WithEvents myAry() as CTimer
'
'So this is one way to add in that kind of support
'
'How does it do it?
'Well I devised this technique specifically for these timer classes (first attempt
'at adding events to collection objects). Because the timer event has to be contained
'in an external module for the API callback function, and the event manually raised
'back in the class instance, I just extended that a little more to make it smart
'enough to be able to tell if a specific timer was a standalone object or part of
'an array of timers created through the CTimers class.
'
'If it was part of an array, then the event is instead raised in the parent class
'and passed the index of that specific timer. (While the base CTimer class still
'handles all of the actual Timer Enable/Disable management)
'
'This technique to forward events to parent class might not transfer to well to
'other scenarios where there was not a global callback function (all timers are
'routed through TimerProc) Although something can probably be devised with
'CallWindowProc Api, Anyway this is a simple clean layout for timers
'
'CTimer started life from Bruce McKinney's HardcoreVB book, however its functionality
'has been completly rewritten. CTimer can be used directly as a standalone Timer
'object, or through CTimers as part of an array of timers
'
'You are free to use this code in any Commercial/Non-Commercial applications as you
'wish.
'
'Author:  David Zimmer
'  Date:  Oct 16th 2003
'  Site:  http://sandsprite.com


Private WithEvents timerArray As CTimers     'Array of CTimer Objects
Attribute timerArray.VB_VarHelpID = -1
Private WithEvents standAloneTimer As CTimer 'Single Ctimer object
Attribute standAloneTimer.VB_VarHelpID = -1

 Private Sub cmd_Click(Index As Integer)
    timerArray(Index).Enabled = Not timerArray(Index).Enabled
    cmd(Index).Caption = "Turn Timer " & Index & IIf(timerArray(Index).Enabled, " Off", " On")
End Sub

Private Sub cmdTimerObj_Click()
    standAloneTimer.Enabled = Not standAloneTimer.Enabled
    cmdTimerObj.Caption = "Turn Timer " & IIf(standAloneTimer.Enabled, " Off", " On")
End Sub

Private Sub Form_Load()
    Dim i
    
    Set timerArray = New CTimers
       
    For i = 1 To 3
        timerArray.Add
        timerArray(i).Interval = i * 1000
        timerArray(i).Enable
    Next
    
    Set standAloneTimer = New CTimer
    standAloneTimer.Interval = 2000
     
End Sub

Private Sub standAloneTimer_Timer()
    List2.AddItem "Standalone Timer"
End Sub

Private Sub timerArray_Timer(ByVal Index As Integer)
    List1.AddItem "timer: " & Index
End Sub



Private Sub List1_Click()
    List1.Clear
End Sub
Private Sub List2_Click()
    List2.Clear
End Sub

