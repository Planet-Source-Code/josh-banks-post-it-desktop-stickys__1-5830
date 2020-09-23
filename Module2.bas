Attribute VB_Name = "Module2"
'Keep Form on Top,
'Delete if you dont need it
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Global Const conHwndTopmost = -1
    Global Const conSwpNoActivate = &H10
    Global Const conSwpShowWindow = &H40

