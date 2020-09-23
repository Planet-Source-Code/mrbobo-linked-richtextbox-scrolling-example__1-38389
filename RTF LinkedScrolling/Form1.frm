VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Linking RichTextBox scrolling example"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RichTextBox2 
      Height          =   4455
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":008B
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********Copyright PSST Software 2002**********************
'Submitted to Planet Source Code - August 2002
'If you got it elsewhere - they stole it from PSC.

'Written by MrBobo - enjoy
'Please visit our website - www.psst.com.au

'Linking Richtextbox scrolling example
'This could be tweaked for use on many controls
'other than RichTextBox's
Option Explicit
Private Sub Form_Load()
    Dim z As Long, temp As String
    'Put some text in the RichTextBox's
    For z = 1 To 300
        temp = temp & "Line number " & z & vbCrLf
    Next
    RichTextBox1.Text = temp
    RichTextBox2.Text = temp
    'Link the scrollers
    fSubClass RichTextBox1.hwnd, RichTextBox2.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Unlink the scrollers
    pUnSubClass RichTextBox1.hwnd, RichTextBox2.hwnd
End Sub

