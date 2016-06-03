VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "VB Notepad"
   ClientHeight    =   6930
   ClientLeft      =   1050
   ClientTop       =   1575
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   12720
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox textArea 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu new 
         Caption         =   "New"
      End
      Begin VB.Menu open 
         Caption         =   "Open"
      End
      Begin VB.Menu save 
         Caption         =   "Save"
      End
      Begin VB.Menu save_as 
         Caption         =   "Save As .."
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu font 
      Caption         =   "Font"
   End
   Begin VB.Menu help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isSaveAsExecuted As Integer

Dim fileName As String
Private Sub exit_Click()
    Unload Me
End Sub

Private Sub font_Click()
    CommonDialog1.ShowFont
    textArea.font = CommonDialog1.FontName
    textArea.FontSize = CommonDialog1.FontSize
End Sub

Private Sub Form_Load()
    textArea.Left = 0
    textArea.Top = 0
    textArea.Width = Me.Width * 0.98
    textArea.Height = Me.Height * 0.98
    isSaveAsExecuted = 0
End Sub

Private Sub Form_Resize()
    textArea.Width = Me.Width * 0.98
    textArea.Height = Me.Height * 0.98
End Sub

Private Sub help_Click()
    msg = MsgBox("VB Notepad application created by c0de-w0rks.", vbInformation, "About VB Notepad")
End Sub

Private Sub new_Click()
    Call save_Click
    textArea.Text = ""
End Sub

Private Sub open_Click()
    Call save_Click
    CommonDialog1.ShowOpen
    Dim openFile, content As String
    openFile = CommonDialog1.fileName
    textArea.Text = ""
    Form1.Caption = openFile
    Open openFile For Input As #1
        Do
            Input #1, content
            textArea.Text = textArea.Text & content & vbCrLf
        Loop Until EOF(1)
    Close #1
End Sub

Private Sub save_as_Click()
    Dim name As String
    CommonDialog1.ShowSave
    fileName = CommonDialog1.fileName & ".txt"
    Open fileName For Output As #1
    Print #1, textArea.Text
    Close #1
    isSaveAsExecuted = 1
End Sub

Private Sub save_Click()
    If isSaveAsExecuted = 0 Then
        Call save_as_Click
        Else
            Open fileName For Output As #1
            Print #1, textArea.Text
            Close #1
    End If
End Sub
