VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.UserControl adRotator 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7290
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   76
   ScaleMode       =   0  'User
   ScaleWidth      =   468
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   840
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   120
      Top             =   0
   End
   Begin VB.Label lblFailed 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Failed to Load"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Stretch         =   -1  'True
      Top             =   240
      Width           =   7290
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Support This Software by Visiting Our Sponsors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
   End
   Begin VB.Shape Shape1 
      Height          =   900
      Left            =   0
      Top             =   240
      Width           =   7290
   End
End
Attribute VB_Name = "adRotator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Feedback and voting is helpful and encouraged.
'
' Add banner advertising to your apps EASILY
' with this control.
' This is a self contained control with an image and label
' that uses the inet control and a timer to retrieve
' images.
' The banner information is handled by three arrays and has
' routines for adding banners and getting the images from a
' server. The banner interval can be changed on the fly.
' The images are stored in the same folder as the app but
' are deleted before the next image loads and after the
' control terminates.
'
' The commenting is poor but the control works well.
'
' GetInternetFile routine by Blake Bell
' The code is copyrighted by its respective owners and
' comes with no warranty, but you can use it in any
' project you want.
'
' Steve D.
'

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private adName() As String
Private adURL() As String
Private bannerURL() As String

Private currentImageFile As String

Public adNum As Integer

Public adInterval As Integer


Public Sub clearAds()

ReDim adName(0)
ReDim adURL(0)
ReDim bannerURL(0)

End Sub

Public Sub addAd(name As String, URL As String, imageURL As String)
Dim i As Integer

i = UBound(adName)
ReDim Preserve adName(i + 1)
ReDim Preserve adURL(i + 1)
ReDim Preserve bannerURL(i + 1)

adName(i + 1) = name
adURL(i + 1) = URL
bannerURL(i + 1) = imageURL

End Sub



Private Sub Image1_Click()
LinkLaunch 0, adURL(adNum)
End Sub

Private Sub Timer1_Timer()
If adNum > 0 Then
If adNum < UBound(adName) Then
adNum = adNum + 1
Else
adNum = 1
End If

getBanner (adNum)
End If

'this will change the interval if it has been altered externally
Timer1.Interval = adInterval

End Sub

Private Sub UserControl_Initialize()
clearAds
adNum = 0
adInterval = 5000

End Sub

Public Sub getBanner(bNum As Integer)
On Error GoTo errored

lblFailed.Visible = False

Dim f As String

If bNum > 0 Then
Image1.Picture = LoadPicture()
Call GetInternetFile(Inet1, bannerURL(bNum), App.Path)
Image1.Picture = LoadPicture(App.Path & "\" & currentImageFile)
End If
Exit Sub

errored:
lblFailed.Visible = True

End Sub


Private Sub GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String)
On Error Resume Next

If myURL = "" Then Exit Sub

If currentImageFile <> "" Then
Kill App.Path & "\" & currentImageFile
End If

    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Sub
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    myFile$ = DestDIR + "\" + RealFile$
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    currentImageFile = RealFile$

End Sub

Private Sub UserControl_Resize()
Shape1.Width = ScaleWidth
Label1.Width = ScaleWidth
Image1.Width = ScaleWidth
Label1.Top = 0
Shape1.Height = ScaleHeight - Label1.Height
Image1.Height = ScaleHeight - Label1.Height
lblFailed.Left = ScaleWidth - lblFailed.Width - 5
lblFailed.Top = ScaleHeight - lblFailed.Height

End Sub


Public Function LinkLaunch(myHwnd As Long, sURL As String) As Integer
    On Error GoTo ErrHandler
    Dim RetVal As Integer
    Dim SW_SHOWNORMAL
    
    RetVal = ShellExecute(myHwnd, "open", sURL, "", "", SW_SHOWNORMAL)
    LinkLaunch = RetVal
    
Exit Function
ErrHandler:
    MsgBox "Web Page not available at this time or the URL is improperly formatted.", vbExclamation, "URL Error"
End Function

Private Sub UserControl_Terminate()
On Error Resume Next

Kill App.Path & "\" & currentImageFile

End Sub
