VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rich Scroll Example"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   1035
      Left            =   5580
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frmAutoScroll.frx":0000
      Top             =   120
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      Left            =   4620
      TabIndex        =   1
      Top             =   60
      Width           =   255
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   0   'False
      TextRTF         =   $"frmAutoScroll.frx":0ED6
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_SCROLL As Long = &HB5
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINESCROLL = &HB6

Dim PSP As Integer

Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
    RichTextBox1.Text = Text1.Text
    Dim lCount As Long
    lCount = SendMessage(RichTextBox1.hwnd, EM_GETLINECOUNT, 0, ByVal 0&)
    VScroll1.Min = 0
    VScroll1.Max = lCount - ((RichTextBox1.Height - 60) / Me.TextHeight("A"))
    VScroll1.SmallChange = 1
    VScroll1.LargeChange = 10
    PSP = 0
End Sub

Private Sub VScroll1_Change()
VScroll1_Scroll
End Sub

Private Sub VScroll1_Scroll()
    Dim l As Long
    Dim i
    With RichTextBox1
            
        If VScroll1.Value > PSP Then
            For i = PSP + 1 To VScroll1.Value
                    l = SendMessage(.hwnd, EM_SCROLL, 1, 0)
            Next i
        ElseIf VScroll1.Value < PSP Then
            For i = VScroll1.Value + 1 To PSP
                l = SendMessage(.hwnd, EM_SCROLL, 0, 1)
            Next i
        End If

        PSP = VScroll1.Value
        
    End With
End Sub
