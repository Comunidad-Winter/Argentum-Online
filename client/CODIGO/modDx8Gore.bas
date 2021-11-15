Attribute VB_Name = "modDx8_Gore"
Option Explicit

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' This program is free software; you can redistribute it and/or modify
' it under the terms of the Affero General Public License;
' either version 1 of the License, or any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' Affero General Public License for more details.
'
' You should have received a copy of the Affero General Public License
' along with this program; if not, you can find it at http://www.affero.org/oagpl.html
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' / Extractos de vbGore utilizados en Boskorcha AO
' / Flechas, daño y partículas.

' - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

' CONST
Public Const DamageDisplayTime As Integer = 2000

' TYPES -
Public Type Projectile

    X           As Single
    Y           As Single
    tX          As Single
    tY          As Single
    RotateSpeed As Byte
    Rotate      As Single
    Grh         As Grh
    
End Type

Public Type structDamage

    Pos     As Position
    value   As String
    Counter As Single
    Width   As Integer
    R       As Byte
    G       As Byte
    B       As Byte
    
End Type

' DECLARES -
Public ProjectileList() As Projectile
Public DamageList()     As structDamage

Public LastProjectile   As Integer    'Last projectile index used
Public LastDamage       As Integer    'Last damage counter text index used

Public Sub Engine_Projectile_Create(ByVal AttackerIndex As Integer, ByVal TargetIndex As Integer, ByVal GrhIndex As Long, ByVal Rotation As Byte)
'*****************************************************************
'Creates a projectile for a ranged weapon
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Create
'*****************************************************************
Dim ProjectileIndex As Integer

    If AttackerIndex = 0 Then Exit Sub
    If TargetIndex = 0 Then Exit Sub
    If AttackerIndex > UBound(charlist) Then Exit Sub
    If TargetIndex > UBound(charlist) Then Exit Sub

    'Get the next open projectile slot
    Do
        ProjectileIndex = ProjectileIndex + 1
        
        'Update LastProjectile if we go over the size of the current array
        If ProjectileIndex > LastProjectile Then
            LastProjectile = ProjectileIndex
            ReDim Preserve ProjectileList(1 To LastProjectile)
            Exit Do
        End If
        
    Loop While ProjectileList(ProjectileIndex).Grh.GrhIndex > 0
    
    'Figure out the initial rotation value
    ProjectileList(ProjectileIndex).Rotate = Engine_GetAngle(charlist(AttackerIndex).Pos.X, charlist(AttackerIndex).Pos.Y, charlist(TargetIndex).Pos.X, charlist(TargetIndex).Pos.Y)
    
    'Fill in the values
    ProjectileList(ProjectileIndex).tX = charlist(TargetIndex).Pos.X * 32 '+ charlist(TargetIndex).MoveOffsetX
    ProjectileList(ProjectileIndex).tY = charlist(TargetIndex).Pos.Y * 32  '+ charlist(TargetIndex).MoveOffsetY
    ProjectileList(ProjectileIndex).RotateSpeed = Rotation
    ProjectileList(ProjectileIndex).X = charlist(AttackerIndex).Pos.X * 32 ' * 32 '+ charlist(AttackerIndex).MoveOffsetX
    ProjectileList(ProjectileIndex).Y = charlist(AttackerIndex).Pos.Y * 32 - 10 ' * 32 '+ charlist(AttackerIndex).MoveOffset

    InitGrh ProjectileList(ProjectileIndex).Grh, GrhIndex
    
End Sub

Public Sub Engine_Projectile_Erase(ByVal ProjectileIndex As Integer)
'*****************************************************************
'Erase a projectile by the projectile index
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Projectile_Erase
'*****************************************************************
 
    'Clear the selected index
    ProjectileList(ProjectileIndex).Grh.FrameCounter = 0
    ProjectileList(ProjectileIndex).Grh.GrhIndex = 0
    ProjectileList(ProjectileIndex).X = 0
    ProjectileList(ProjectileIndex).Y = 0
    ProjectileList(ProjectileIndex).tX = 0
    ProjectileList(ProjectileIndex).tY = 0
    ProjectileList(ProjectileIndex).Rotate = 0
    ProjectileList(ProjectileIndex).RotateSpeed = 0
 
    'Update LastProjectile
    If ProjectileIndex = LastProjectile Then
        Do Until ProjectileList(ProjectileIndex).Grh.GrhIndex > 1
            'Move down one projectile
            LastProjectile = LastProjectile - 1
            If LastProjectile = 0 Then Exit Do
        Loop
        If ProjectileIndex <> LastProjectile Then
            'We still have projectiles, resize the array to end at the last used slot
            If LastProjectile > 0 Then
                ReDim Preserve ProjectileList(1 To LastProjectile)
            Else
                Erase ProjectileList
            End If
        End If
    End If
 
End Sub

Public Sub Engine_Damage_Create(ByVal X As Integer, ByVal Y As Integer, ByVal value As Integer, ByVal R As Byte, ByVal G As Byte, ByVal B As Byte)

' Note: Crea el daño en el mapa

Dim DamageIndex As Integer

'Get the next open damage slot

    Do
        DamageIndex = DamageIndex + 1

        'Update LastDamage if we go over the size of the current array
        If DamageIndex > LastDamage Then
            LastDamage = DamageIndex
            ReDim Preserve DamageList(1 To LastDamage)
            Exit Do
        End If

    Loop While DamageList(DamageIndex).Counter > 0
    
    Debug.Print value
    
    'Set the values
    If Not value Then DamageList(DamageIndex).value = "Miss" Else
    DamageList(DamageIndex).value = value
    DamageList(DamageIndex).Counter = DamageDisplayTime
    DamageList(DamageIndex).Width = Engine_GetTextWidth(cfonts(1), DamageList(DamageIndex).value)
    DamageList(DamageIndex).Pos.X = X
    DamageList(DamageIndex).Pos.Y = Y
    DamageList(DamageIndex).R = R
    DamageList(DamageIndex).G = G
    DamageList(DamageIndex).B = B

End Sub

Public Sub Engine_Damage_Erase(ByVal DamageIndex As Integer)

' Note: Quita el daño del mapa

    'Clear the selected index
    DamageList(DamageIndex).Counter = 0
    DamageList(DamageIndex).value = vbNullString
    DamageList(DamageIndex).Width = 0

    'Update LastDamage
    If DamageIndex = LastDamage Then
        Do Until DamageList(LastDamage).Counter > 0

            'Move down one splatter
            LastDamage = LastDamage - 1

            If LastDamage = 0 Then
                Erase DamageList
                Exit Sub
            Else
                'We still have damage text, resize the array to end at the last used slot
                ReDim Preserve DamageList(1 To LastDamage)
            End If

        Loop
    End If

End Sub

Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEn ... ne_TPtoSPX" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
 
    Engine_TPtoSPX = Engine_PixelPosX(X - minX) + OffsetCounterX - 24

End Function
 
Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEn ... ne_TPtoSPY" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
 
    Engine_TPtoSPY = Engine_PixelPosY(Y - minY) + OffsetCounterY - 23
 
End Function

Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEn ... _PixelPosX" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    Engine_PixelPosX = (X - 1) * TilePixelWidth - 32 * 4
 
End Function
 
Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'More info: http://www.vbgore.com/GameClient.TileEn ... _PixelPosY" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    Engine_PixelPosY = (Y - 1) * TilePixelWidth - 32 * 4
 
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEn ... e_GetAngle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'************************************************************
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function
 
End Function



