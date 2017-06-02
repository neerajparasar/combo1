VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form1"
   ClientHeight    =   5205
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10575
   BeginProperty Font 
      Name            =   "PanRoman"
      Size            =   9.75
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Wide Latin"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   6120
      List            =   "Form1.frx":0019
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   840
      Width           =   3855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0052
      Left            =   360
      List            =   "Form1.frx":006B
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "PanRoman"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   11
      Top             =   3480
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "PanRoman"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   10
      Top             =   3480
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_Click()
Label1.Caption = Combo1.Text
End Sub



Private Sub Combo2_Click()
Label2.Caption = Combo2.Text
End Sub

Private Sub Command1_Click()
If Len(Text1.Text) > 0 Then
Combo1.AddItem (Text1.Text)
Text1.Text = ""
End If
End Sub

Private Sub Command2_Click()
If Len(Text2.Text) > 0 Then
Combo2.AddItem (Text2.Text)
Text2.Text = ""
End If
End Sub

Private Sub Command3_Click()
If Combo1.Text = "" Then
MsgBox "no items to add"
Else

Label1.Caption = ""
Combo2.AddItem (Combo1.Text)
Combo1.RemoveItem Combo1.ListIndex
End If
End Sub

Private Sub Command4_Click()
If Combo2.Text = "" Then
MsgBox "no items to add"
Else
Label2.Caption = ""
Combo1.AddItem (Combo2.Text)
Combo2.RemoveItem Combo2.ListIndex
End If
End Sub

Private Sub Command5_Click()
Label1.Caption = ""
Dim i, j As Integer
j = Combo1.ListCount
If j = 0 Then
MsgBox "No more items"
Else
For i = 0 To j - 1
Combo2.AddItem Combo1.List(i)
Next i
End If
Combo1.Clear
End Sub

Private Sub Command6_Click()
Label2.Caption = ""
Dim i, j As Integer
j = Combo2.ListCount
If j = 0 Then
MsgBox "no more items"
Else
For i = 0 To j - 1
Combo1.AddItem Combo2.List(i)
Next i
End If
Combo2.Clear
End Sub

