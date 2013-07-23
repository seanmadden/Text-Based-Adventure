VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Avast"
   ClientHeight    =   6855
   ClientLeft      =   810
   ClientTop       =   1455
   ClientWidth     =   11025
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   11025
   Begin VB.CommandButton cmdMap 
      Caption         =   "Map"
      Height          =   375
      Left            =   8760
      TabIndex        =   6
      Top             =   3600
      Width           =   735
   End
   Begin VB.Timer tmrtick 
      Interval        =   10000
      Left            =   9360
      Top             =   4680
   End
   Begin VB.Timer tmrcombat 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   9360
      Top             =   5280
   End
   Begin RichTextLib.RichTextBox txtOutput 
      Height          =   5895
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10398
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      MousePointer    =   1
      TextRTF         =   $"main.frx":0000
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   6480
      Width           =   8295
   End
   Begin VB.Label lblroom 
      Height          =   735
      Left            =   8640
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblhp 
      Height          =   255
      Left            =   8640
      TabIndex        =   4
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblClass 
      Height          =   615
      Left            =   8640
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblName 
      Height          =   375
      Left            =   8640
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim town As zone
Dim wild1 As zone


Dim charname As String
Dim action As Integer
Dim lastaction As Integer
Dim playerclass As String
Dim humanplayer As player

Dim orc As monster

Dim curx, cury As Integer
Dim exits As String

Dim start As Integer

Dim CombatState As String

Dim monster As Boolean
Dim attackable As String
Dim target As String

Dim sword As weapon

Dim exitto As String

Dim currentarea As String

Dim arrcount, arc1 As Integer
Dim comm(50) As String

Private Sub cmdMap_Click()
    Load Form2
    Form2.Show
End Sub


Private Sub Form_Load()
    currentarea = "town"
    Dim contain As String
    txtOutput.text = "Ahoy there sir/madam, what is thoust name?"
    action = 1
    CombatState = "ready"
    
    playerx = 7
    playery = 3
    
    humanplayer.hp = 20
    humanplayer.dex = 15
    humanplayer.mana = 20
    'humanplayer.str = 15
    humanplayer.str = 9999
    
    humanplayer.maxhp = 20
    humanplayer.maxmana = 20
    
    humanplayer.xp = 1
    humanplayer.exptnl = 1000
    humanplayer.lvl = 1
    
    humanplayer.weapon = "sword"
    
    sword.mindamage = 2
    sword.maxdamage = 4
    
    
    Open "U:\GameVB\1.0\maps/town.txt" For Input As #1
    
    For i = 0 To 23
        Input #1, x, y, exits, Descr, contain, exitto
        town.room(x, y) = exits
        town.Descr(x, y) = Descr
        town.contains(x, y) = contain
        town.exitto(x, y) = exitto
    Next i
    current = town
    Close #1
    
    curx = 7
    cury = 3
    exits = current.room(curx, cury)
    
End Sub

Private Sub tmrcombat_Timer()
    CombatState = "combat"
    Randomize
    Dim damage As Integer
    Dim damage2 As Integer
    Dim crit As Boolean
    
    lblhp.Caption = humanplayer.hp
    
    Dim chance As Integer
    chance = Int(Rnd * 100 + humanplayer.dex)
    If chance > 90 And chance < 100 Then
        at ("You missed the " & target & vbCrLf & target & " hits you for " & damage2)
    ElseIf chance > 100 Then
        crit = True
        damage = damage + Int(Rnd * humanplayer.str)
    End If
    
    damage2 = Int(Rnd * orc.str)
    humanplayer.hp = humanplayer.hp - damage2
    damage = Int(Rnd * humanplayer.str)
    orc.hp = orc.hp - damage
    
    If crit = True Then
        at (target & " hits you for " & damage2 & " points of damage" & vbCrLf & "You CRITICALLY hit " & target & " for " & damage & " damage")
        crit = False
    End If
        
    If orc.hp <= 0 Then
        Dim expgiven As Integer
        expgiven = orc.expgiven
        humanplayer.xp = humanplayer.xp + expgiven
        at (target & " hits you for " & damage2 & " points of damage" & vbCrLf & "You hit " & target & " for " & damage & " damage" & vbCrLf & "You have slain the " & target & vbCrLf & "You gained " & expgiven & " eperience points")
        CombatState = "ready"
        lblhp.Caption = humanplayer.hp
        
        Do While humanplayer.xp > humanplayer.exptnl
            If humanplayer.xp > humanplayer.exptnl Then
                humanplayer.exptnl = humanplayer.exptnl * 2
                humanplayer.lvl = humanplayer.lvl + 1
                at ("You have gained a level, You will need " & humanplayer.exptnl & " Experience points for your next level")
            End If
        Loop
        
        current.contains(curx, cury) = ""
        Call mob
        tmrcombat.Enabled = False
    Else
        at (target & " hits you for " & damage2 & " points of damage" & vbCrLf & "You hit " & target & " for " & damage & " damage")
    End If
    
    

End Sub

Private Sub tmrtick_Timer()
    If humanplayer.hp <> humanplayer.maxhp Then
        If CombatState = "rest" Then
            humanplayer.hp = humanplayer.hp + (humanplayer.hp / 2) + 5
        Else
            humanplayer.hp = humanplayer.hp + (humanplayer.hp / 4) + 1
        End If
    End If
    If humanplayer.hp > humanplayer.maxhp Then
        humanplayer.hp = humanplayer.maxhp
    End If
End Sub



Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
'I need 38(up) and 40(down)
On Error GoTo hell
    If KeyCode = 13 Then
        KeyCode = 0
        arrcount = 0

        
        If txtInput.text = "" And action = 0 Then
            at ("")
            Exit Sub
        Else
        
        comm(arc1) = txtInput.text
        arc1 = arc1 + 1
        If arc1 > 50 Then
            arc1 = 0
        End If
        
            Select Case action
                Case 1
                    charname = txtInput.text
                    lastaction = action
                    action = 2
                    Call at("So, your name is " & charname & "? [Yes/No]", False)
                    
                Case 2
                    If UCase(txtInput.text) = "YES" Then
                        Call at("Alright then.", False)
                        If lastaction = 1 Then
                            lblName.Caption = charname
                            lastaction = action
                            action = 3
                            Call at("What class would you like to be?" & vbCrLf & "P - Pirate" & vbCrLf & "N - Ninja", False)
                        ElseIf lastaction = 3 Then
                            Call at("You are now a " & playerclass & "")
                            at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury))
                            lblClass = playerclass
                            action = 0
                        End If
                    ElseIf UCase(txtInput.text) = "NO" Then
                        If lastaction = 1 Then
                            Call at("What is your name?", False)
                        ElseIf lastaction = 3 Then
                            Call at("What class would you like to be?" & vbCrLf & "P - Pirate" & vbCrLf & "N - Ninja", False)
                        End If
                        action = lastaction
                    Else
                        Call at("I said YES or NO.", False)
                    End If
                    
                Case 3
                    If UCase(txtInput.text) = "P" Then
                        Call at("Looks like you want to be a Pirate. [Yes/No]", False)
                        playerclass = "pirate"
                        lastaction = action
                        action = 2
                    ElseIf UCase(txtInput.text) = "N" Then
                        Call at("Looks like you want to be a Ninja. [Yes/No]", False)
                        playerclass = "ninja"
                        lastaction = action
                        action = 2
                    Else
                        Call at("That class doesnt exist.", False)
                        Call at("Pick again.", False)
                        Call at("What class would you like to be?" & vbCrLf & "P - Pirate" & vbCrLf & "N - Ninja", False)
                    End If
                Case Else
                

                
                If InStr(UCase(txtInput.text), "K ") > 0 And monster = True And CombatState = "ready" Then
                    If InStr(UCase(txtInput.text), attackable) > 0 Then
                        Dim temp As String
                        For i = 1 To Len(txtInput.text)
                            If i = 1 Then
                                temp = UCase(Mid(txtInput.text, i, 1))
                            Else
                                temp = Mid(txtInput.text, i, 1)
                            End If
                            
                            If temp <> " " Then
                                target = target + temp
                            Else
                                target = ""
                            End If
                        Next i
                        
                        at ("You attack the " & target)
                        CombatState = "combat"
                        tmrcombat.Enabled = True
                        txtInput.text = ""
                        Exit Sub
                    End If

                ElseIf InStr(UCase(txtInput.text), "K ") > 0 And CombatState = "rest" Then
                    at ("You're resting!")
                End If
                
                    If LCase(txtInput.text) = "flee" Then
                        If CombatState = "combat" Then
                            Randomize
                            Dim flee As Integer
                            Dim dir As Integer
                            flee = Int(Rnd() * 100 + 1)
                            If flee < 90 Then
                                CombatState = "ready"
                                tmrcombat.Enabled = False
                                at ("You flee in panic")
                                If InStr(exits, "north") > 0 And flee < 20 Then
                                    Call gonorth
                                ElseIf InStr(exits, "south") > 0 And flee > 21 And flee < 50 Then
                                    Call gosouth
                                ElseIf InStr(exits, "east") > 0 And flee > 51 And flee < 70 Then
                                    Call goeast
                                ElseIf InStr(exits, "west") > 0 And flee > 71 And flee < 90 Then
                                    Call gowest
                                End If
                                Exit Sub
                            Else
                                at ("You fail to flee!")
                                Exit Sub
                            End If
                        Else
                            at ("You can only do that while in combat!")
                            txtInput.text = ""
                            Exit Sub
                        End If
                    End If
                    
                    If UCase(txtInput.text) = "HELP" Then
                        Call at("Made by me!")
                    ElseIf UCase(txtInput.text) = "LOOK" Or UCase(txtInput.text) = "L" Then
                        monster = False
                        Call mob
                        If monster = True Then
                            at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury) & vbCrLf & "There is an orc here")
                        Else
                            at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury))
                        End If
                        'at (current.Descr(curX, curY) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curX, curY))
                        exits = current.room(curx, cury)
                        lblroom.Caption = "(" & curx & "," & cury & ")"
                    ElseIf LCase(txtInput.text) = "north" Then
                        Call gonorth
                    ElseIf LCase(txtInput.text) = "south" Then
                        Call gosouth
                    ElseIf LCase(txtInput.text) = "east" Then
                        Call goeast
                    ElseIf LCase(txtInput.text) = "west" Then
                        Call gowest
                        
                    ElseIf LCase(txtInput.text) = "w" Then
                        Call gowest
                    ElseIf LCase(txtInput.text) = "n" Then
                        Call gonorth
                    ElseIf LCase(txtInput.text) = "s" Then
                        Call gosouth
                    ElseIf LCase(txtInput.text) = "e" Then
                        Call goeast
                    ElseIf LCase(txtInput.text) = "map" Then
                        Form2.Show
                    ElseIf UCase(txtInput.text) = "STAND" Or UCase(txtInput.text) = "ST" And CombatState = "rest" Then
                        CombatState = "ready"
                        at ("You stand up and ready yourself")
                    ElseIf UCase(txtInput.text) = "INV" Then
                        Dim inv As String
                        For i = 1 To 50
                            If humanplayer.inventory(i) <> "" Then
                                inv = inv + vbCrLf + humanplayer.inventory(i)
                                at (inv)
                            End If
                        Next i
                    ElseIf UCase(txtInput.text) = "SCORE" Or UCase(txtInput.text) = "SC" Then
                        at ("You are " & charname & " a level " & humanplayer.lvl & " " & playerclass & vbCrLf & vbCrLf & _
                        "XP: " & humanplayer.xp & "/" & humanplayer.exptnl)
                    ElseIf UCase(txtInput.text) = "R" Or UCase(txtInput.text) = "REST" Then
                        If CombatState <> "rest" Then
                            CombatState = "rest"
                            at ("You sit down awhile and rest")
                        Else
                            at ("You're already resting")
                        End If
                        
                    Else
                        txtInput.text = ""
                        Call at("Arr, invalid command")
                    End If
                    
            End Select
            
            'txtOutput.text = txtOutput.text + vbCrLf + txtInput.text
            txtOutput.SelStart = Len(txtOutput) - 1
            txtInput.text = ""
            lblhp.Caption = "HP: " & humanplayer.hp
        End If


    ElseIf KeyCode = 38 Then
        arrcount = arrcount - 1
        txtInput.text = comm(arrcount)
    ElseIf KeyCode = 40 Then
        

        arrcount = arrcount + 1
        If arrcount > start Then
            arrcount = arrcount - 1
        End If
        txtInput.text = comm(arrcount)
        
    End If

    Exit Sub
hell:
arrcount = 0
    For i = 50 To 0 Step -1
        If comm(i) <> "" Then
            start = i
            Exit For
            End If
        Next i
    arrcount = start
End Sub

Public Sub at(ByVal text As String, Optional prompt As Boolean = True)
        If humanplayer.hp <= 0 Then
            txtOutput.text = txtOutput.text & vbCrLf + target & " has slain you!"
            lblhp.Caption = humanplayer.hp
            tmrcombat.Enabled = False
            txtOutput.SelStart = Len(txtOutput) - 1
            txtInput.text = ""
            Exit Sub
        End If
                
        txtOutput.text = txtOutput.text + vbCrLf + text + vbCrLf
        If prompt = True Then
            txtOutput.text = txtOutput.text + vbCrLf + "[" & humanplayer.hp & "/" & humanplayer.maxhp & "hp " & humanplayer.mana & "/" & humanplayer.maxmana & "m]" + vbCrLf
        End If
        txtOutput.SelStart = Len(txtOutput) - 1
End Sub

Public Sub gowest()
    If CombatState = "ready" Then
        If InStr(exits, "west") > 0 Then
            
            cury = cury - 1
            
            If cury < 0 Then
                cury = cury + 1
                Call area
                cury = 7
            End If
            monster = False
            Call mob
            If monster = True Then
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury) & vbCrLf & "There is an orc here")
            Else
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury))
            End If
            'at (currentDescr(curX, curY) & vbCrLf & "Obvious Exits: " & vbCrLf & currentroom(curX, curY))
            exits = current.room(curx, cury)
            lblroom.Caption = "(" & curx & "," & cury & ")"
                                
        Else
            at ("You cant go there!")
        End If
    Else
        at ("You're still resting!")
    End If
    playery = cury
    Call Form2.Form_Load
End Sub

Public Sub goeast()
    If CombatState = "ready" Then
        If InStr(exits, "east") > 0 Then
            
            cury = cury + 1
            If cury > 7 Then
            cury = cury - 1
                Call area
                cury = 0
            End If
            
            monster = False
            Call mob
            If monster = True Then
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury) & vbCrLf & "There is an orc here")
            Else
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury))
            End If
            'at (current.Descr(curX, curY) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curX, curY))
            exits = current.room(curx, cury)
            lblroom.Caption = "(" & curx & "," & cury & ")"
        Else
            at ("You cant go there!")
        End If
    Else
        at ("You're still resting!")
    End If
    playery = cury
    Call Form2.Form_Load
End Sub

Public Sub gonorth()
    If CombatState = "ready" Then
        If InStr(exits, "north") > 0 Then
            curx = curx - 1
            If curx < 0 Then
            curx = curx + 1
                Call area
                curx = 7
            End If
            monster = False
            Call mob
            If monster = True Then
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury) & vbCrLf & "There is an orc here")
            Else
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury))
            End If
            'at (current.Descr(curX, curY) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curX, curY))
            exits = current.room(curx, cury)
            lblroom.Caption = "(" & curx & "," & cury & ")"
        Else
            at ("You cant go there!")
        End If
    Else
        at ("You're still resting!")
    End If
    playerx = curx
    Call Form2.Form_Load
End Sub

Public Sub gosouth()
    If CombatState = "ready" Then
        If InStr(exits, "south") > 0 Then
            curx = curx + 1
            If curx > 8 Then
                curx = curx - 1
                Call area
                curx = 0
            End If
            monster = False
            Call mob
            If monster = True Then
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury) & vbCrLf & "There is an orc here")
            Else
                at (current.Descr(curx, cury) & vbCrLf & "Obvious Exits: " & vbCrLf & current.room(curx, cury))
            End If
            exits = current.room(curx, cury)
            lblroom.Caption = "(" & curx & "," & cury & ")"
            
        Else
            at ("You cant go there!")
        End If
    Else
        at ("You're still resting!")
    End If
        playerx = curx
        Call Form2.Form_Load
End Sub

Public Sub mob()
      If InStr(current.contains(curx, cury), "orc") > 0 Then
        orc.hp = 20
        orc.maxhp = 20
        orc.mana = 20
        orc.dex = 10
        orc.str = 10
        orc.maxmana = 20
        orc.expgiven = 9999
        orc.lvl = 1
        monster = True
        'target = "orc"
    End If

End Sub

Public Sub area()
    On Error Resume Next
    
    If current.exitto(curx, cury) <> "" Then
        Select Case currentarea
            Case "town"
                town = current
                Open "U:\GameVB\1.0\maps/" & town.exitto(curx, cury) & ".txt" For Input As #1
                For i = 0 To 63
                    Input #1, x, y, exits, Descr, contain, exitto
                    wild1.room(x, y) = exits
                    wild1.Descr(x, y) = Descr
                    wild1.contains(x, y) = contain
                    wild1.exitto(x, y) = exitto
                Next i
                Close #1
                current = wild1
                currentarea = "wild1"
            Case "wild1"
                current = town
                currentarea = "town"
        End Select
        
    End If
    
End Sub

Private Sub txtOutput_Click()
    txtInput.SetFocus
End Sub
