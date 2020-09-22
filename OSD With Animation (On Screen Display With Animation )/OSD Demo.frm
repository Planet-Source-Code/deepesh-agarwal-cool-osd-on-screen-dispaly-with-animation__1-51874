VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OSD TestPad"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdShowMe 
      Caption         =   "Show OSD"
      Default         =   -1  'True
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "OSD Demo.frx":0000
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
' GUYS get my free NSIS based installer for your VB apps
' (1). Its Easy (Wizard Based)
' (2). Based on rock Solid NSIS Super-PIMP Tecnology
' (3). Will compress your project size to almost half, YES your installer size
'      will be less then your total file size
' (4). Support for adding Splash Skin with fading effect and Sound.
' Many many more, visit product page at www.deepeshagarwal.tk

'=========================================================================================
'  OSD With Animation (On Screen Display With Animation ) Demonstrator
'  This Form Demostrate how to make a OSD with animation
'  To Make USE in your project:-
'  (1). Add the form OSDWIN.frm to ur project.
'  (2). Add Sub Below into the calling Form
  ' Private Sub OSD(ByVal xOSDtext As String)
  '       On Error Resume Next
  '       Unload OSDwin
  '       With OSDwin
  '          .text.Caption = xOSDtext
  '          .Show
  '       End With
  ' End Sub
'  (3). Call Sub OSD as and when Needed as Shown below
'=========================================================================================
' This Code was Adapted from someone's else code but as i can't find authors name,i am unable to mention
' his\her name, Just from project properties his company name is "Compaq". if ur the one THANKS.
'
'  Coded By: Deepesh Agarwal
'  WebSite: http://www.deepeshagarwal.tk
'  E-mail: agarwal_deepesh@indiatimes.com
'  Visit my site for Free-Software's like:
'  1). The-AdPolice - Blocks 21000+ adservers to save bandwidth
'  2). Dr. System 2.0 -  Schedule Computer Maintainence - A must for every computer user
'  3). Service Controller XP (A Must For XP User) - Start,Stop,Pause and change startup type of 2000/XP services with recommended settings for different system config.
'   And Many More........
'=========================================================================================



Option Explicit


Private Sub OSD(ByVal xOSDtext As String)

'Now, we have to load our OSD form at desired location and show the OSD.
    On Error Resume Next
    'unload If already loaded
    Unload OSDwin
    'Show
    With OSDwin
        .Text.Caption = xOSDtext
        .Show
    End With

End Sub


Private Sub CmdShowMe_Click()
'Calling the SUB to show the OSD
    OSD Text1.Text

End Sub
