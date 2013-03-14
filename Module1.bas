Attribute VB_Name = "Module1"
Public Declare Sub mciSendStringA Lib "winmm.dll" (ByVal lpstrCommand As String, _
                                                    ByVal lpstrReturnString As Any, _
                                                    ByVal uReturnLength As Long, _
                                                    ByVal hwndCallback As Long)


