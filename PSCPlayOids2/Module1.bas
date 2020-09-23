Attribute VB_Name = "Module1"
' Module1.bas

' Playoids2

Option Explicit
Option Base 1              'Arrays starting at subscript 1

Public Type POINTAPI
        xk As Long
        yk As Long
End Type

'Use:
'Dim pp As POINTAPI
'res& = LineTo(Object.hdc, x, y)
'res& = MoveToEx(Object.hdc, x, y, pp)

'LineTo and MoveToEx are faster the VB equivalents
Declare Function LineTo Lib "gdi32" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" _
(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
lpPoint As POINTAPI) As Long

Public Sub FixExtension(FSpec$, Ext$)
' In: FixExtension FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub


