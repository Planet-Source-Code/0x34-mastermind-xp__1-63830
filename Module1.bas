Attribute VB_Name = "Module1"
Option Explicit

Global Wins As Integer
Global Losses As Integer
Global SPly As Integer


Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" ( _
    ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const csndsync = &H0, csndasync = &H1
Const csndnodefault = &H2
Const csndloop = &H8, csndnostop = &H10



' That's all folks!
