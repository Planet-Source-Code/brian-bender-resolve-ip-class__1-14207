VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   300
      TabIndex        =   5
      Top             =   2280
      Width           =   5295
      Begin VB.CommandButton Command1 
         Caption         =   "Convert IP to Host Name"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1020
         Width           =   2595
      End
      Begin VB.TextBox txtIPaddress 
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Text            =   "64.58.76.177"
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label lblHostName 
         Height          =   255
         Left            =   3060
         TabIndex        =   9
         Top             =   660
         Width           =   1875
      End
      Begin VB.Label Label1 
         Caption         =   "IP Address:"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   5295
      Begin VB.CommandButton Command2 
         Caption         =   "Convert Host Name to IP Address"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtHostName 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "Microsoft.com"
         Top             =   540
         Width           =   2595
      End
      Begin VB.Label lblIPAddress 
         Height          =   255
         Left            =   3300
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Host Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oIP As New cResolveIP

Private Sub Command1_Click()
    lblHostName = oIP.GetHostNameByIPAddress(Trim(txtIPaddress.Text))
    
End Sub

Private Sub Command2_Click()
   lblIPAddress = oIP.GetIPAddressbyHostName(Trim(txtHostName.Text))
   
End Sub


