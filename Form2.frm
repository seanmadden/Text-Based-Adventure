VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Ye Olde Map"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9825
   LinkTopic       =   "Form2"
   ScaleHeight     =   7320
   ScaleWidth      =   9825
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblInfo 
      Caption         =   "* is your current location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   1
      Top             =   6960
      Width           =   2895
   End
   Begin VB.Label lblmap 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()
    On Error Resume Next
    lblmap(7).Visible = False
    Dim t, l, curx, cury As Integer
    'curx and cury is the location in the array that we are working on
    'if curx = 6 and cury = 6 then count over 6 and down 6 on the label
    curx = 0
    cury = 0
    t = 700
    l = 2000

    For i = 1 To 64
        
        
        Load lblmap(i)
        lblmap(i).Visible = False
        lblmap(i).Top = t
        lblmap(i).Left = l
        
        
        l = l + lblmap(i).Width
        If current.Descr(cury, curx) <> "" Then
            'if there is a description for that room, make the label visible
            lblmap(i).Visible = True
        End If
        
        
        lblmap(i).Caption = ""
        If playery = curx And playerx = cury Then
            'if the player is currently in the room add a star to indicate
            'this is mostly used when form_load is called when a player moves
            lblmap(i).Caption = "*"
        End If
        
        curx = curx + 1
        'add one to curx because we are working across the top
        If i Mod 8 = 0 Then 'a row is finished, reset the left, add a row to the top
                            'and add one to cury
            t = t + lblmap(i).Height
            cury = cury + 1
            l = 2000
            curx = 0
        End If
    Next i
End Sub

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub lblmap_Click(Index As Integer)
    MsgBox Index
End Sub
