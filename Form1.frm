VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Test"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract Calc.exe"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.text <> "" Then
Dim datei As Byte
fsiz = ShowFileSize(App.Path & "\" & App.EXEName & ".exe")
Open (App.Path & "\" & App.EXEName & ".exe") For Binary As #1
Open Text1.text For Binary As #2
prog1 = fsiz - 94208
prog2 = 94208
Get #1, prog1 + 1, datei
Put #2, 1, datei
For a = 1 To 94208
Get #1, , datei
Put #2, , datei
Next
Close
End If
End Sub
Function ShowFileSize(file)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(file)
    ShowFileSize = f.Size
    's = UCase(f.Name) & " uses " & f.Size & " bytes."
    'MsgBox s, 0, "Folder Size Info"
End Function
'94208

Private Sub Form_Load()
Text1.text = App.Path & "\" & "Test2.exe"
End Sub
