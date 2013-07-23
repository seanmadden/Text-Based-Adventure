Attribute VB_Name = "Module1"

Public current As zone
Public playerx, playery As Integer

Public Type player
    hp As Integer
    mana As Integer
    str As Integer
    dex As Integer
    weapon As String
    inventory(50) As String
    
    maxhp As Integer
    maxmana As Integer
    
    xp As Double
    exptnl As Double
    lvl As Integer
End Type

Public Type monster
    hp As Integer
    mana As Integer
    str As Integer
    dex As Integer
    
    maxhp As Integer
    maxmana As Integer
    expgiven As Double
    
    lvl As Integer
End Type

Public Type zone
    room(7, 7) As String
    Descr(7, 7) As String
    contains(7, 7) As String
    exitto(7, 7) As String
End Type

Public Type weapon
    mindamage As Integer
    maxdamage As Integer
End Type

