VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11550
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   17775
   LinkTopic       =   "Form1"
   ScaleHeight     =   11550
   ScaleWidth      =   17775
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   11160
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   10800
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   10455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   17535
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      URL             =   "http://"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_DownloadingFile     As Boolean
Private m_DownloadingFileSize As Long
Private m_LocalSaveFile       As String
Dim RemoteFileToGet As String
Dim LocalFileToSave As String

 Private Declare Function ShellExecute Lib _
                "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, ByVal lpOperation As String, _
                ByVal lpFile As String, ByVal lpParameters As String, _
                ByVal lpDirectory As String, ByVal nShowCmd As Long _
                ) As Long


Private Sub Command1_Click()
    Dim s As String
Command1.Enabled = False

Open App.Path & "\insert.txt" For Input As #3
Line Input #3, insertline
Close #3

Open App.Path & "\itemlist.txt" For Input As #1
Do Until (EOF(1))

    Input #1, idNum, itemName, Lucylink

RemoteFileToGet = "http://lucy.allakhazam.com/itemraw.html?id=" & idNum
LocalFileToSave = "temp.txt"
'DownloadFile

Open App.Path & "\temp.txt" For Input As #2

Line Input #2, Phrase
    
s = Phrase

X = Split(s, vbLf)
For r = 0 To UBound(X) - 1
'RealText X(r)

If Left(X(r), 18) = "    <td width=" & Chr(34) & "20%" Then
'    <td width="20%" class="spelllabel">aagi</td>
printout = Left(X(r), Len(X(r)) - 5)
printoutb = Right(printout, Len(printout) - 39)

Text1.Text = Text1.Text & printoutb

ElseIf Left(X(r), 18) = "    <td width=" & Chr(34) & "30%" Then
'    <td width="30%">0</td>
printout = Left(X(r), Len(X(r)) - 5)
printoutb = Right(printout, Len(printout) - 20)

Text1.Text = Text1.Text & " - " & printoutb & vbCrLf

End If

Next


Close #2



'Text1.Text = Text1.Text & idNum & " / " & itemName & vbCrLf


pb1.Value = pb1.Value + 1

Loop
Close #1


End Sub


Sub DownloadFile()

Dim vtData()  As Byte
Dim FreeNr    As Integer
Dim SizeDone  As Long
Dim bDone     As Boolean
Dim GetPerc   As Integer

m_LocalSaveFile = App.Path & "\" & LocalFileToSave
Inet1.Execute RemoteFileToGet, "GET " & Chr(34) & App.Path & "\" & LocalFileToSave & Chr(34)


FreeNr = FreeFile
    
    Open App.Path & "\" & LocalFileToSave For Binary Access Write As FreeNr

                Do While Not bDone
                    If Inet1.StillExecuting = False Then
                    
                    vtData = Inet1.GetChunk(1024, icByteArray)
                    
                    SizeDone = SizeDone + UBound(vtData)
                                                 
                              
                    Put #FreeNr, , vtData()
                    If UBound(vtData) = -1 Then
                        bDone = True
                    Else
                        DoEvents
                    End If
                    
                    Else
                    DoEvents
                    End If
                Loop
                
                Close FreeNr
    
    TransferSuccess = True


End Sub

