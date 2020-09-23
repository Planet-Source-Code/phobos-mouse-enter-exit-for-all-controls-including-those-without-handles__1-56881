VERSION 5.00
Begin VB.Form frmMouseOver 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   675
      ScaleHeight     =   2355
      ScaleWidth      =   6150
      TabIndex        =   1
      Top             =   435
      Width           =   6210
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   540
         Left            =   4800
         TabIndex        =   2
         Tag             =   "Skinned"
         Top             =   315
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         FillStyle       =   4  'Upward Diagonal
         Height          =   540
         Left            =   4785
         Shape           =   4  'Rounded Rectangle
         Tag             =   "Skinned"
         Top             =   1590
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   315
         Left            =   4845
         TabIndex        =   3
         Tag             =   "Skinned"
         Top             =   1065
         Width           =   990
      End
      Begin VB.Image Image1 
         Height          =   1815
         Index           =   2
         Left            =   2520
         Picture         =   "frmMouseOver.frx":0000
         Tag             =   "Skinned"
         Top             =   270
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   1815
         Index           =   1
         Left            =   315
         Picture         =   "frmMouseOver.frx":AA6A
         Tag             =   "Skinned"
         Top             =   270
         Width           =   1800
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1710
      TabIndex        =   0
      Tag             =   "Skinned"
      Top             =   3240
      Width           =   4170
   End
End
Attribute VB_Name = "frmMouseOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Test program to demonstrate trapping of MouseEnter, MouseExit and MouseWheel events.
'
' This method uses subclassing to detect the events, and will only report MouseEnter
' and MouseExit events for controls with the text "Skinnned" in the tag field.
'
' This program is a progression on the MouseEnter/MouseExit work submitted by Evan Toder.
'
' This new method allows you to caputer Events for controls that do not have a hwnd
' property (such as images and labels).

Private WithEvents cMouseEvents As clsMouseEvents
Attribute cMouseEvents.VB_VarHelpID = -1

Private Sub Form_Load()
    ' Start listening for MouseEnter and MouseExit events.
    Set cMouseEvents = New clsMouseEvents
    Call cMouseEvents.HookMouseEvents(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' The form is being unloaded so we clean up.
    Set cMouseEvents = Nothing
End Sub

Private Sub cMouseEvents_MouseEnter(ctlEntered As Control)
    ' MouseEnter Event detected.
    If ctlEntered.Name = "Image1" Then
        Text1.Text = "MouseEnter " & ctlEntered.Name & " (index item " & ctlEntered.Index & ")"
    Else
        Text1.Text = "MouseEnter " & ctlEntered.Name
    End If
End Sub

Private Sub cMouseEvents_MouseExit(ctlExited As Control)
    ' MouseExit Event detected.
    If ctlExited.Name = "Image1" Then
        Text1.Text = "MouseExit " & ctlExited.Name & " (index item " & ctlExited.Index & ")"
    Else
        Text1.Text = "MouseExit " & ctlExited.Name
    End If
End Sub

Private Sub cMouseEvents_MouseWheelUp(lngMoueX As Long, lngMouseY As Long)
    ' MouseWheelUp Event detected.
    Text1.Text = "MouseWheel Up"
End Sub

Private Sub cMouseEvents_MouseWheelDown(lngMoueX As Long, lngMouseY As Long)
    ' MouseWheelDown Event detected.
    Text1.Text = "MouseWheel Down"
End Sub
