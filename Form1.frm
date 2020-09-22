VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fair Ftp"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4740
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3840
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Upload"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3720
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   ".."
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "21"
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Fair Ftp By Calle Callecalle_10@hotmail.com"
      Height          =   255
      Left            =   1560
      TabIndex        =   16
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Logo By Gouranga"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "Form1.frx":08CA
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label Label7 
      Caption         =   "0 Bytes."
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Remote File"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Local File"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "FTP UserName"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "FTP Password"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "FTP RemotePort"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "FTP RemoteHost"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
CD1.ShowOpen
Text5.Text = CD1.FileName
Text6.Text = CD1.FileTitle
Label7.Caption = ShowFileSize(Text5) & " Bytes."
End Sub

Function ShowFileSize(file)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(file)
    ShowFileSize = f.Size
End Function

Private Sub Command2_Click()
Inet1.AccessType = icUseDefault
Inet1.Protocol = icFTP
Inet1.RemoteHost = Text1.Text
Inet1.RemotePort = Text2.Text
Inet1.Password = Text3.Text
Inet1.UserName = Text4.Text
Inet1.RequestTimeout = "60"
strRemote = Text6.Text
strLocal = Text5.Text
Form1.Inet1.Execute , "PUT """ & strLocal & """ " & strRemote
End Sub

Private Sub Label10_Click()

End Sub

