VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test open cash drawer"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4545
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Open Drawer 6"
      Height          =   615
      Left            =   2520
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Open Drawer 5"
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Port"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2895
      Begin VB.OptionButton Option2 
         Caption         =   "COM 2"
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "COM 1"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open Drawer 4"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Open Drawer 3"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open Drawer 2"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Drawer 1"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo errorepson
    
    'Sets the Com Port number (this can be changed here)
    If Option1.Value = True Then
        MSComm1.CommPort = 1
        ElseIf Option2.Value = True Then
        MSComm1.CommPort = 2
    End If
    MSComm1.Settings = "9600,n,8,1"
    'Sets the Baud rate (9600 in this case)
    MSComm1.PortOpen = True
    'opens the port
    MSComm1.Output = Chr$(27) + "p" + "0" + "zz"
    'Sends a pulse to the Com Port
    MSComm1.PortOpen = False
    'Closes the port when finished
    Exit Sub

errorepson:
    MsgBox "There was a problema, cash drawer could not be opened"
    Exit Sub
End Sub

Private Sub Command2_Click()
On Error GoTo errorepson

    If Option1.Value = True Then
        MSComm1.CommPort = 1
        ElseIf Option2.Value = True Then
        MSComm1.CommPort = 2
    End If
    
    MSComm1.PortOpen = True
    MSComm1.Output = Chr$(27) & "p" & 0 & vbCrLf
    MSComm1.Output = Chr$(27) & "@"
    MSComm1.PortOpen = False
    
    Exit Sub
errorepson:
    MsgBox "There was a problema, cash drawer could not be opened"
    Exit Sub
End Sub

Private Sub Command3_Click()
On Error GoTo errorepson

Dim Puerto As String

    If Option1.Value = True Then
        Puerto = "COM1"
    ElseIf Option2.Value = True Then
        Puerto = "COM2"
    End If

    Open Puerto For Output Access Write As #1
    Print #1, Chr(27) & Chr(112) & Chr(0)
    Close #1

    Exit Sub
errorepson:
    MsgBox "There was a problema, cash drawer could not be opened"
    Exit Sub
End Sub

Private Sub Command4_Click()
On Error GoTo errorepson
Dim Puerto As String

    If Option1.Value = True Then
        Puerto = "COM1:"
    ElseIf Option2.Value = True Then
        Puerto = "COM2:"
    End If

    Open Puerto For Output As #1
    Print #1, Chr$(&H1B); "p"; Chr$(0); Chr$(100); Chr$(250)
    Close #1
    
    Exit Sub

errorepson:
    MsgBox "There was a problema, cash drawer could not be opened"
    Exit Sub
End Sub

Private Sub Command5_Click()
On Error GoTo errorepson

    'Set The Port Number
    If Option1.Value = True Then
        MSComm1.CommPort = 1
        ElseIf Option2.Value = True Then
        MSComm1.CommPort = 2
    End If
    
    'Open The COM Port
    MSComm1.PortOpen = True
    
    'Set The COM Port Settings
    MSComm1.Settings = "19200,N,8,1"
    
    'Open The Draw
    MSComm1.Output = Chr(&H1B&) & Chr(&H70&) & Chr(&H0&) & "1" & "4" & Chr(13)
    
    'Close The COM Port
    MSComm1.PortOpen = False
    
    Exit Sub

errorepson:
    MsgBox "There was a problema, cash drawer could not be opened"
    Exit Sub

End Sub

Private Sub Command6_Click()
On Error GoTo errorepson

    If Option1.Value = True Then
        MSComm1.CommPort = 1
        ElseIf Option2.Value = True Then
        MSComm1.CommPort = 2
    End If
    
    MSComm1.PortOpen = True
    MSComm1.Output = Chr$(&H1B) + Chr$(&H70) + Chr$(0) + Chr$(25) + Chr$(250)
    MSComm1.PortOpen = False
    
    Exit Sub
    
errorepson:
    MsgBox "There was a problema, cash drawer could not be opened"
    Exit Sub
End Sub
