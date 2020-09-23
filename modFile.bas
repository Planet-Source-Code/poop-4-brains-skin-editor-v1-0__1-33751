Attribute VB_Name = "modFile"
Type Colors
LightColor As Long
DarkColor As Long
Back As Long
Caption As Long
End Type

Public Type SkinInfo
Title As String
bntMax As Boolean
bntClose As Boolean
bntMin As Boolean

TitleHeight As Long

TColors As Colors
WColors As Colors

Width As Long
Height As Long
End Type

Public Skin As SkinInfo

Function SaveSkin(Path As String)
Open Path For Output As #1
Write #1, Skin.Title, Skin.Width, Skin.Height, Skin.TitleHeight
Write #1, Skin.bntClose, Skin.bntMin, Skin.bntMax
Write #1, Skin.TColors.Back, Skin.TColors.Caption, Skin.TColors.DarkColor, Skin.TColors.LightColor
Write #1, Skin.WColors.Back, Skin.WColors.Caption, Skin.WColors.DarkColor, Skin.WColors.LightColor
Close #1
End Function

Function LoadSkin(Path As String)
Open Path For Input As #1
Input #1, Skin.Title, Skin.Width, Skin.Height, Skin.TitleHeight
Input #1, Skin.bntClose, Skin.bntMin, Skin.bntMax
Input #1, Skin.TColors.Back, Skin.TColors.Caption, Skin.TColors.DarkColor, Skin.TColors.LightColor
Input #1, Skin.WColors.Back, Skin.WColors.Caption, Skin.WColors.DarkColor, Skin.WColors.LightColor
Close #1
End Function

Function ConvertToBMP(bmpPath As String)

End Function
