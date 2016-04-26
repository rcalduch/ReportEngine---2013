Public Class csScreenImage

  Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
  Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Integer) As Integer
  Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
  Private Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
  Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
  Private Declare Function BitBlt Lib "GDI32" (ByVal srchDC As Integer, ByVal srcX As Integer, ByVal srcY As Integer, ByVal srcW As Integer, ByVal srcH As Integer, ByVal desthDC As Integer, ByVal destX As Integer, ByVal destY As Integer, ByVal op As Integer) As Integer
  Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Integer) As Integer
  Private Declare Function DeleteObject Lib "GDI32" (ByVal hObj As Integer) As Integer
  Const SRCCOPY As Integer = &HCC0020

  Public Shared Function GetScreenImage() As System.Drawing.Bitmap
    Dim ImgScr As System.Drawing.Bitmap
    Dim hSDC, hMDC As Integer
    Dim hBMP, hBMPOld As Integer
    Dim FW, FH As Integer
    Dim r As Integer

    hSDC = CreateDC("DISPLAY", "", "", "")
    hMDC = CreateCompatibleDC(hSDC)

    FW = GetDeviceCaps(hSDC, 8)
    FH = GetDeviceCaps(hSDC, 10)
    hBMP = CreateCompatibleBitmap(hSDC, FW, FH)

    hBMPOld = SelectObject(hMDC, hBMP)
    r = BitBlt(hMDC, 0, 0, FW, FH, hSDC, 0, 0, 13369376)
    hBMP = SelectObject(hMDC, hBMPOld)

    r = DeleteDC(hSDC)
    r = DeleteDC(hMDC)

    ImgScr = System.Drawing.Image.FromHbitmap(New IntPtr(hBMP))
    DeleteObject(hBMP)

    Return ImgScr

  End Function

  Public Shared Function GetScreenSnapshot() As System.Drawing.Image
    Return GetScreenSnapshot(False)
  End Function

  Public Shared Function GetScreenSnapshot(ByVal activeWindowOnly As Boolean) As System.Drawing.Image
    ' Alt-Print Screen captures the active window only
    If activeWindowOnly Then
      System.Windows.Forms.SendKeys.SendWait("%{PRTSC}")
    Else
      System.Windows.Forms.SendKeys.SendWait("{PRTSC 2}")
    End If

    ' return the bitmap now in the clipboard
    Return DirectCast(System.Windows.Forms.Clipboard.GetDataObject().GetData(System.Windows.Forms.DataFormats.Bitmap), System.Drawing.Image)
  End Function

End Class
