VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Ad Rotator Test"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin Project1.adRotator adRotator1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2143
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
adRotator1.adNum = 1
adRotator1.addAd "ad1", "http://www.planetsourcecode.com/", "http://www.planetsourcecode.com/banners/banner1.gif"
adRotator1.getBanner (1)
End Sub
