VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simple Way to Check for Changes on a Form"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Tag             =   "Validate"
      Top             =   825
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   915
      Left            =   240
      TabIndex        =   4
      Top             =   1350
      Width           =   3975
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Tag             =   "Validate"
         Top             =   525
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Tag             =   "Validate"
         Top             =   225
         Width           =   2415
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      Left            =   2400
      TabIndex        =   3
      Tag             =   "Validate"
      Text            =   "Combo1"
      Top             =   300
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Tag             =   "Validate"
      Text            =   "Text2"
      Top             =   825
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Tag             =   "Validate"
      Text            =   "Text1"
      Top             =   300
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim varSnapShot() As Variant

Private Sub Form_Load()
Combo1.AddItem "a"
Combo1.AddItem "b"
Combo1.AddItem "c"
Combo1.AddItem "d"

Call subRead
Call subSnapShot

End Sub

'=============================================================================
'Simple Example - Normaly load recordset here
'=============================================================================
Private Sub subRead()
    Me.Text1.Text = "Text1"
    Me.Text2.Text = "Text2"
    Me.Option1.Value = True
End Sub

'=============================================================================
'Validate Controls that have the word "Validate" in the Tag property
'We will take a snapshot of values on the form. When we exit, we will compare the values
'=============================================================================
Private Sub subSnapShot()
    Dim i As Integer
    Dim ctlLoop As Control
    For Each ctlLoop In Me.Controls
        If ctlLoop.Tag = "Validate" Then
            i = i + 1
            ReDim Preserve varSnapShot(i)
            varSnapShot(i) = ctlLoop
        End If
    Next
    Set ctlLoop = Nothing
End Sub

'=============================================================================
'Validate controls that have the word "Validate" in the Tag property
'When we exit the form we will check for changes to the entry fields
'=============================================================================
Private Function IsChanged() As Boolean
    Dim i As Integer
    Dim ctlLoop As Control
    For Each ctlLoop In Me.Controls
            If ctlLoop.Tag = "Validate" Then
                i = i + 1
                If varSnapShot(i) <> ctlLoop Then
                    IsChanged = True
                    Exit Function
                End If
            End If
    Next
    IsChanged = False
    Set ctlLoop = Nothing
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub


'=============================================================================
'Unload Event
'=============================================================================
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call subExitEdit
    Unload Me
End Sub

'=============================================================================
'Check for Changes
'=============================================================================
Private Sub subExitEdit()
    If IsChanged = True Then
        Dim lngResult As Long
        lngResult = MsgBox("There are changes to X. Would you like to save the changes?", vbYesNo + vbQuestion, "Save X")
        If lngResult = vbYes Then
            Call subSave
        End If
    End If
End Sub

Private Sub subSave()
    'Save Data
End Sub

