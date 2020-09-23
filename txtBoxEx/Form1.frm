VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Text box extension class !"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6390
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load text file"
      Height          =   375
      Left            =   240
      TabIndex        =   20
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll to caret"
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3600
      Top             =   3960
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options "
      Height          =   2055
      Left            =   3600
      TabIndex        =   13
      Top             =   1680
      Width           =   2655
      Begin VB.CheckBox Check5 
         Caption         =   "Disable paste"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Lose focus with ""enter"" (doesn't work if multiline)"
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Read only (not disabled!)"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Select text on entry"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Context menu "
      Height          =   1335
      Left            =   3600
      TabIndex        =   9
      Top             =   240
      Width           =   2655
      Begin VB.OptionButton Option2 
         Caption         =   "Custom (!)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "None"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Default"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Allowed entry "
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
      Begin VB.CheckBox Check1 
         Caption         =   "Upper case"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Low case"
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Currency"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Numeric"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Alfa"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Everything"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Extension of the textbox! "
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Label Label3 
      Caption         =   "File path:"
      Height          =   255
      Left            =   1920
      TabIndex        =   22
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Current line number"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Number of lines:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu Prova 
         Caption         =   "This is..."
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu Pippo 
         Caption         =   "... my custom menu!"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim myTxtbox As cTxtBoxEx

Dim myCheck As TxtExCheck
Dim myCase As AlfaCase

Private Sub Check1_Click(Index As Integer)
    Dim valueL As Boolean
    Dim valueU As Boolean
    
    valueL = CBool(Check1(0).Value)
    valueU = CBool(Check1(1).Value)

    myCase = 0

    If valueL Then
        myCase = myCase Or TxtLowerCase
    End If

    If valueU Then
        myCase = myCase Or TxtUpperCase
    End If
    
    setCheck
End Sub

Private Sub Check2_Click()
    myTxtbox.SelectOnEntry = CBool(Check2.Value)
End Sub

Private Sub Check3_Click()
    myTxtbox.ReadOnly = CBool(Check3.Value)
End Sub

Private Sub Check4_Click()
    myTxtbox.EnterLoseFocus = CBool(Check4.Value)
End Sub

Private Sub Check5_Click()
    If Check5.Value = 1 Then
        myTxtbox.PasteEnabled = False
    Else
        myTxtbox.PasteEnabled = True
    End If
End Sub

Private Sub Command1_Click()
    myTxtbox.ScrollCaret
End Sub

Private Sub Command2_Click()
    
    On Error GoTo myErrHandle
    myTxtbox.LoadFile Text3.Text
    Exit Sub
    
myErrHandle:
    MsgBox "Error loading file: " & Err.Number & " " & Err.Description, vbExclamation
End Sub

Private Sub Form_Load()
    Set myTxtbox = New cTxtBoxEx
    Set myTxtbox.TextBoxRef = Text1
    
    Option1(0).Value = True
    Check1(0).Value = 1
    Check1(1).Value = 1
    Option2(0).Value = True
    
    Text3.Text = App.Path & "\cTxtBoxEx.cls"
End Sub


Private Sub Option1_Click(Index As Integer)

    Select Case Index
    
        Case 0
            myCheck = TxtNone
        
        Case 1
            myCheck = TxtAlfa
        
        Case 2
            myCheck = TxtNumeric
            
        Case 3
            myCheck = TxtCurrency

    End Select
    
    setCheck
    
    Check1(0).Enabled = Index = 1
    Check1(1).Enabled = Index = 1
    
End Sub

Private Sub setCheck()
    myTxtbox.PerformCheck myCheck, myCase
End Sub

Private Sub Option2_Click(Index As Integer)
    Select Case Index
        Case 0
            myTxtbox.SetContextMenu TxtMenuDefault
            
        Case 1
            myTxtbox.SetContextMenu TxtMenuNone
            
        Case 2
            myTxtbox.SetContextMenu TxtMenuCustom, mnuPopup
            
    End Select
End Sub

Private Sub Timer1_Timer()
    '
    Text2(0).Text = myTxtbox.LineCount
    Text2(1).Text = myTxtbox.CurrentLineNum
End Sub
