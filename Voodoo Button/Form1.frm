VERSION 5.00
Begin VB.Form frmButton 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voodoo Button"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tquit2 
      Left            =   1800
      Top             =   1320
   End
   Begin VB.Timer Tquit1 
      Left            =   1320
      Top             =   1320
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Quit"
      Height          =   315
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Timer Tbutton2 
      Left            =   600
      Top             =   1320
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Start"
      Height          =   315
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton cmdVoodoo 
      BackColor       =   &H000000C0&
      Caption         =   "Voodoo"
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.Timer Tbutton1 
      Left            =   120
      Top             =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   5040
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frmButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Very simple examble to make special
'effects for your project

Private Sub cmdExit_Click()

'Now Tquit1 and Tquit2 can
'execute their commands
Tquit1.Enabled = True
Tquit2.Enabled = True

End Sub

Private Sub cmdStart_Click()

'Tbutton1 can execute it's commands
Tbutton1.Enabled = True

'Tbutton2 can't execute it's commands
Tbutton2.Enabled = False

End Sub

Private Sub Form_Load()

'As you can see I haven't changed the following values
'from the properties so that every one can understand
'what this project needs to work.
Tbutton1.Enabled = False
Tbutton2.Enabled = False
Tquit1.Enabled = False
Tquit2.Enabled = False
Tbutton1.Interval = 5
Tbutton2.Interval = 5
Tquit1.Interval = 5
Tquit2.Interval = 100
'If you want you can change them directly from the properties.

End Sub

Private Sub Tbutton1_Timer()

'Now the button will start to move
'with this direction  <-----
cmdVoodoo.Left = cmdVoodoo.Left - 20

'When cmdVoodoo.Left goes to 120
'then Tbutton1 disables and
'Tbutton2 enables
If cmdVoodoo.Left = 120 Then
    Tbutton1.Enabled = False
    Tbutton2.Enabled = True
End If

End Sub

Private Sub Tbutton2_Timer()

'Now the button will start to move
'with this direction  ----->
cmdVoodoo.Left = cmdVoodoo.Left + 20

'When cmdVoodoo.Left goes to 120
'then Tbutton2 disables and
'Tbutton1 enables
If cmdVoodoo.Left = 4200 Then
    Tbutton2.Enabled = False
    Tbutton1.Enabled = True
End If

End Sub

Private Sub Tquit1_Timer()

'The form leaves with speed
'to the left of the screen
Me.Left = Me.Left - 2000

End Sub

Private Sub Tquit2_Timer()

'Close project
Unload Me

End Sub
