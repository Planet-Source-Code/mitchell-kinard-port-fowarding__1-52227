VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Port Manager v1.0"
   ClientHeight    =   5070
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   7770
   FillColor       =   &H8000000D&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Caption         =   "Stop"
      Height          =   255
      Left            =   5400
      MaskColor       =   &H8000000B&
      TabIndex        =   9
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Start"
      Height          =   255
      Left            =   120
      MaskColor       =   &H8000000A&
      TabIndex        =   8
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "Port Input/Output Data:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   7575
      Begin MSWinsockLib.Winsock Winsock2 
         Left            =   6720
         Top             =   360
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   6720
         Top             =   3240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtPLOG 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   3495
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   7335
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      Caption         =   "Send To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtrport 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   3960
         TabIndex        =   6
         Text            =   "3000"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtrIP 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   480
         TabIndex        =   4
         Text            =   "192.168.0.1"
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000007&
         Caption         =   "Port:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   260
         Width           =   495
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000007&
         Caption         =   "IP:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   250
         Width           =   255
      End
   End
   Begin VB.TextBox txtlport 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "80"
      Top             =   225
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Local Port:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu cdata 
      Caption         =   "Clear Data"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cdata_Click()
txtPLOG.Text = ""
End Sub

Private Sub Command1_Click()
txtPLOG.Text = txtPLOG.Text & "Fowarding Port [" & txtlport.Text & "] To:" & txtrIP.Text & " [" & txtrport.Text & "]" & vbCrLf
Winsock1.Close
Winsock1.LocalPort = txtlport.Text
Winsock1.Listen
End Sub

Private Sub Command2_Click()
txtPLOG.Text = txtPLOG.Text & "Fowarding Stopped!" & vbCrLf
Winsock1.Close
Winsock2.Close
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
txtPLOG.Text = txtPLOG.Text & "Connection From:" & Winsock1.RemoteHostIP & vbCrLf
Winsock1.Close
Winsock1.Accept requestID
Winsock2.Close
Winsock2.RemoteHost = txtrIP.Text
Winsock2.RemotePort = txtrport.Text
Winsock2.Connect
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim rDATA As String
Winsock1.GetData rDATA
txtPLOG.Text = txtPLOG.Text & rDATA
Winsock2.SendData rDATA
End Sub

Private Sub Winsock2_Close()
Winsock1.Close
Winsock2.Close
Winsock1.Listen
End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim sDATA As String
Winsock2.GetData sDATA
txtPLOG.Text = txtPLOG.Text & sDATA
Winsock1.SendData sDATA
End Sub

