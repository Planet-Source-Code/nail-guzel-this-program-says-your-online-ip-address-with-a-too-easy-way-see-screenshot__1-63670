VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   ScaleHeight     =   885
   ScaleWidth      =   2100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "tell my ip.."
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "your ip address"
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub Command1_Click()
Dim ip As String
download = URLDownloadToFile(0, "http://k.domaindlx.com/nailgg/tr/ip.asp", "c:\windows\temp\a1.tmp", 0, 0)

Open "c:\windows\temp\a1.tmp" For Input As #1
 Input #1, ip
Close

Text1 = ip
End Sub
