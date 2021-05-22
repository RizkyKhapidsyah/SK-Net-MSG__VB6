VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetMSG Plus v1.0"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMacAdr 
      Caption         =   "Mac Address"
      Height          =   3840
      Left            =   0
      TabIndex        =   21
      Top             =   8475
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   6675
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1125
      End
      Begin VB.TextBox txtMac 
         Height          =   3450
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Top             =   315
         Width           =   6405
      End
      Begin VB.CommandButton cmdGetMac 
         Caption         =   "Inspect"
         Height          =   300
         Left            =   6675
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   300
         Width           =   1125
      End
      Begin VB.CommandButton cmdCopy 
         Caption         =   "Copy"
         Height          =   300
         Left            =   6675
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   675
         Width           =   1125
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "NetMSG Plus"
      Height          =   3840
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7965
      Begin VB.CommandButton cmdLock 
         Caption         =   "Lock Station"
         Height          =   375
         Left            =   6705
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "This Locks the Workstation you are on."
         Top             =   1440
         Width           =   1155
      End
      Begin VB.OptionButton optUsers 
         Caption         =   "Users"
         Height          =   255
         Left            =   5805
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "This option requires a user name to be input."
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton optDomain 
         Caption         =   "Domain"
         Height          =   255
         Left            =   4725
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "This option requires a valid domain. If left blank it will assume the current domain."
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   375
         Left            =   6705
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "About the Program."
         Top             =   300
         Width           =   1155
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         Height          =   375
         Left            =   6705
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Uh!"
         Top             =   735
         Width           =   1155
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   375
         Left            =   6705
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Resets the Interface"
         Top             =   2895
         Width           =   1155
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6705
         TabIndex        =   12
         ToolTipText     =   "Sends the current message."
         Top             =   3375
         Width           =   1155
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   1050
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   705
         Width           =   5475
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1065
         TabIndex        =   10
         Top             =   360
         Width           =   3480
      End
      Begin VB.CommandButton cmdMacAdr 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mac Address"
         Height          =   375
         Left            =   6705
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "This option uses Arp to extract Mac Addresses"
         Top             =   1860
         Width           =   1155
      End
      Begin VB.CommandButton cmdPing 
         Caption         =   "Ping Host"
         Height          =   375
         Left            =   6705
         TabIndex        =   8
         ToolTipText     =   "This option allows you to ping any valid Host / IP Address"
         Top             =   2280
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Message:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   225
         TabIndex        =   20
         Top             =   705
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Domain:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   315
         TabIndex        =   19
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame fraPing 
      Caption         =   "Ping"
      Height          =   3840
      Left            =   0
      TabIndex        =   1
      Top             =   4575
      Visible         =   0   'False
      Width           =   7965
      Begin VB.CommandButton cmdPReturn 
         Caption         =   "OK"
         Height          =   390
         Left            =   6675
         TabIndex        =   6
         Top             =   675
         Width           =   1140
      End
      Begin VB.TextBox txtPingResults 
         Height          =   3090
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   675
         Width           =   6390
      End
      Begin VB.CommandButton cmdPingIt 
         Caption         =   "Ping!"
         Height          =   315
         Left            =   6675
         TabIndex        =   4
         Top             =   300
         Width           =   1140
      End
      Begin VB.TextBox txtPingAdr 
         Height          =   315
         Left            =   1575
         TabIndex        =   2
         Text            =   "Enter Host / IP Address to Ping..."
         Top             =   300
         Width           =   4965
      End
      Begin VB.Label lblPing 
         AutoSize        =   -1  'True
         Caption         =   "Ping Destination:"
         Height          =   195
         Left            =   225
         TabIndex        =   3
         Top             =   375
         Width           =   1200
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   225
      Top             =   1650
   End
   Begin ComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   3855
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10855
            MinWidth        =   457
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "8:17"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   1806
            MinWidth        =   1806
            TextSave        =   "22/05/2021"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NERR_Success As Long = 0&
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
Private Declare Function NetMessageBufferSend Lib "NETAPI32.DLL" _
(yServer As Any, yToName As Byte, yFromName As Any, yMsg As Byte, ByVal lSize As Long) As Long

Dim Str As String

'###########################################################################################################
' Lock Workstation
'###########################################################################################################

Private Sub cmdLock_Click()
    Dim bUserCancel As Boolean
    Call LockWorkStation
End Sub

'###########################################################################################################
'Deals with the Ping Command
'###########################################################################################################

Private Sub cmdPing_Click()
    fraPing.Top = 0
    fraPing.Left = 0
    fraMain.Visible = False
    fraPing.Visible = True
    txtPingResults.Text = ""
    txtPingAdr.Text = "Enter Host / IP Address to Ping..."
    sb.Panels(1).Text = "Ping a Host / IP Address"
End Sub

Private Sub cmdPReturn_Click()
    Call txtPingAdr_Click
    fraPing.Visible = False
    fraMain.Visible = True
    sb.Panels(1).Text = App.Title
End Sub

Private Sub cmdPingIt_Click()
    txtPingResults.Text = ""
    txtPingResults.Text = MGetCmdOutput.GetCommandOutput("ping " & txtPingAdr.Text, True, False, True)
End Sub

Private Sub txtPingAdr_Click()
    txtPingAdr.Text = "127.0.0.1"
    txtPingResults.Text = ""
    sb.Panels(1).Text = "Pinging: " & txtPingAdr.Text
End Sub

Private Sub txtPingAdr_Change()
    If txtPingAdr.Text <> "" Then
        sb.Panels(1).Text = "Pinging: " & txtPingAdr.Text
    Else
        MsgBox "Entry cannot be empty. Please enter a valid Host / IP address", vbCritical + vbOKOnly
    End If
End Sub

'###########################################################################################################
' Mac Address Function
'###########################################################################################################

Private Sub cmdMacAdr_Click()
    fraMacAdr.Top = 0
    fraMacAdr.Left = 0
    fraMain.Visible = False
    fraMacAdr.Visible = True
    txtMac.Text = ""
    sb.Panels(1).Text = "Mac Address Extractor"
End Sub

Private Sub cmdCopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtMac.Text
End Sub

Private Sub cmdGetMac_Click()
    txtMac.Text = MGetCmdOutput.GetCommandOutput("arp -a", True, False, True)
End Sub

Private Sub OKButton_Click()
    fraMacAdr.Visible = False
    fraMain.Visible = True
    sb.Panels(1).Text = App.Title
End Sub

'###########################################################################################################
' Button Commands
'###########################################################################################################

Private Sub cmdAbout_Click()
    sb.Panels(1).Text = "About NetMsg v1.0"
    Load frmAbout
    frmAbout.Show , Me
    sb.Panels(1).Text = App.Title
End Sub

Private Sub cmdExit_Click()
    Unload Me: End
End Sub

Private Sub cmdReset_Click()
    Text1.Text = "/DOMAIN:"
    Text2.Text = ""
    Text1.SetFocus
    optUsers.Value = False
    optDomain.Value = True
End Sub

Private Sub cmdSend_Click()
    
    Dim c, m As String
    
    c = Text1.Text
    m = Text2.Text
    
    If (Text1.Text = "") Then
        MsgBox "Enter Computer/User Name"
        Text1.SetFocus
    ElseIf (Text2.Text = "") Then
        MsgBox "Enter Your Message"
        Text2.SetFocus
    End If
    
    Call GetCommandOutput("net send " & c & " " & m, True, False, True)
    
End Sub

Private Sub Form_Activate()
    'reset to base values
    cmdReset_Click
End Sub

Private Sub optDomain_Click()
    If optUsers.Value = True Then
        optDomain.Value = False
        Label1.Caption = "User:"
        Text1.Text = "Enter User Name"
    ElseIf optUsers.Value = False Then
        optDomain.Value = True
        Label1.Caption = "Domain:"
        Text1.Text = "/DOMAIN:"
    End If
End Sub

Private Sub optDomain_GotFocus()
    sb.Panels(1).Text = "This option uses the current domain."
End Sub

Private Sub optUsers_Click()
    If optDomain.Value = True Then
        optUsers.Value = False
        Label1.Caption = "Domain:"
        Text1.Text = "/DOMAIN:"
    ElseIf optDomain.Value = False Then
        optUsers.Value = True
        Label1.Caption = "User:"
        Text1.Text = "Enter User Name"
    End If
End Sub

Private Sub optUsers_GotFocus()
    sb.Panels(1).Text = "This option requires a user name."
End Sub

Private Sub Text1_GotFocus()
    sb.Panels(1).Text = "Enter Details"
    Timer1.Enabled = False
End Sub

Private Sub Text2_GotFocus()
    sb.Panels(1).Text = "Enter Your Message"
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    
    If Not Text2.Text = "" Then
        cmdSend.Enabled = True
    Else
        cmdSend.Enabled = False
    End If
End Sub

'###########################################################################################################
' Send Message Function
'###########################################################################################################

Public Function Sendmsg(strTo As String, strFrom As String, strMessage As String) As Boolean
   
    Dim bytTo() As Byte
    Dim bytFrom() As Byte
    Dim bytMsg() As Byte
    Dim bytName() As Byte
    
    bytTo = strTo & vbNullChar
    bytName = strFrom & vbNullChar
    bytMsg = strMessage & vbNullChar

    Sendmsg = (NetMessageBufferSend(ByVal 0&, bytName(0), _
              ByVal 0&, bytMsg(0), UBound(bytMsg)) = NERR_Success)
End Function

