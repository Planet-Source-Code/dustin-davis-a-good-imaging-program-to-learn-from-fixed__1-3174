Attribute VB_Name = "Module1"
'This gets the mouse position!
Public Type POINTAPI
       X As Long
       
Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function Get_XY()

    Dim MP As POINTAPI, posx As Long, posy As Long
    Call GetCursorPos(MP)
    posx = MP.X
    posy = MP.Y
    Form1.MX = posx
    Form1.MY = posy
End Function
