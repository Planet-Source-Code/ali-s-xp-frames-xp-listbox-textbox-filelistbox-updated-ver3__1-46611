VERSION 5.00
Begin VB.UserControl XPListBox 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   ControlContainer=   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   4410
   ToolboxBitmap   =   "XPListBox.ctx":0000
End
Attribute VB_Name = "XPListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const BorderPixelX As Long = 3
Private Const BorderPixelY As Long = 3
'       --------------------------------
'       ------------- By ---------------
'       ----- Mohammad Ali Sohrabi -----
'       ------ ali6236@yahoo.com -------
'       ------- !!!Freeware!!! ---------
'       --------------------------------

Private Function TwipX(lngPixel As Long) As Long
    TwipX = ScaleX(lngPixel, vbPixels, vbTwips)
End Function
Private Function TwipY(lngPixel As Long) As Long
    TwipY = ScaleY(lngPixel, vbPixels, vbTwips)
End Function

Private Sub UserControl_Resize()
On Error Resume Next
    Dim i As Object
    For Each i In UserControl.ContainedControls
        i.BorderStyle = 0
        i.Appearance = 0
        i.Move -TwipX(1), -TwipY(1), ScaleWidth + TwipX(2), ScaleHeight + TwipY(2)
        If i.IntegralHeight Then
            UserControl.Height = i.Height - TwipX(2)
        End If
    Next
End Sub
