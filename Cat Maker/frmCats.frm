VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   5640
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   5055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Text1.Text = "CREATE TABLE `eqitems` " & vbCrLf & "("

Open App.Path & "\cats.txt" For Input As #1
Do Until (EOF(1))

    Line Input #1, Phrase
    If Phrase = "" Then
        GoTo skip:
    End If

If Left(Phrase, 18) = "    <td width=" & Chr(34) & "20%" Then

'    <td width="20%" class="spelllabel">aagi</td>

printout = Left(Phrase, Len(Phrase) - 5)
printoutb = Right(printout, Len(printout) - 39)

Text1.Text = Text1.Text & "`" & printoutb & "` TEXT NOT NULL , " & vbCrLf

End If

skip:
Loop
Close #1

Text1.Text = Text1.Text & ");"

End Sub

