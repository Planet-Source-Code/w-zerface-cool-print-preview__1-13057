VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrintPreview 
      Caption         =   "Print Preview"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   4560
      Width           =   1395
   End
   Begin RichTextLib.RichTextBox rtbPrint 
      Height          =   4395
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7752
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrintPreview_Click()

    PrintPreview rtbPrint, 1, 1, 1, 1, 1
End Sub

Private Sub Form_Load()

    rtbPrint.RightMargin = rtbPrint.Width
    rtbPrint.FileName = App.Path & "\test.rtf"
    
End Sub

