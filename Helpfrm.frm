VERSION 5.00
Begin VB.Form Helpfrm 
   Caption         =   "Help"
   ClientHeight    =   6285
   ClientLeft      =   2550
   ClientTop       =   1695
   ClientWidth     =   7140
   Icon            =   "Helpfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7140
   Begin VB.CommandButton Command1 
      Caption         =   "Return"
      Height          =   255
      Left            =   5880
      TabIndex        =   1
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   5535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Visit Me On the Web: Http://come.to/phishbowler"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   4815
   End
End
Attribute VB_Name = "Helpfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
AnswerCrack.Show
End Sub

Private Sub Form_Load()
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & "--------------------- ANSWER CRACK MACHINE --------------------- "
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine

Helpfrm.Text1 = Helpfrm.Text1 & "Many answering machines have a code that allows you to hear your messages when you are not around. Answer Machine Crack is for determining it." & vbNewLine

Helpfrm.Text1 = Helpfrm.Text1 & "This program entitled Answer Crack is meant strictly for educational purposes. You may look at the code, run it, under the condition that you do not use it for illegal purposes. " & vbNewLine & "First you must decide whether you are trying to crack a 2 digit code, or a 3 digit code." & vbNewLine & "The two digit code processes all possible 2 digit codes 00-99 as quickly as possible," & vbNewLine & "repeat digits have been eliminated as best as possible. Such as  00102 (00, 01, 10, 02)." & vbNewLine _
& "3 digit codes process all possible 3 digit codes from 000-999. Since 3 digit codes are" & vbNewLine _
& "so long, and many times it is difficult to see exactly what point the code has been" & vbNewLine _
& "cracked, you may choose to break the code off into individual sections. This" & vbNewLine _
& "would mean you would call your answering machine an equal amount of times as" & vbNewLine _
& "you have chosen sections. So 20 sections would mean 20 calls, but the length of" & vbNewLine _
& "each call may only be 6 seconds.  In 6 seconds you may break it down to say" & vbNewLine _
& "under 10 three digit codes for one section." & vbNewLine _
& "" & vbNewLine _
& "When you choose one of the sections it will begin to play." & vbNewLine

Helpfrm.Text1 = Helpfrm.Text1 & "This will invoke the " & Chr$(34) & "Tone Playing" & Chr$(34) & " (You can also repeat play by clicking the button)" & vbNewLine _
& "You should have your phone ready and held up to the speaker so that," & vbNewLine _
& "the answering machine catches all the tones being played. Make sure your" & vbNewLine _
& "sound card is working properly, you will actually *hear* the tones being played." & vbNewLine _
& "(I think that's obvious though huh?) Once the tones are done being played," & vbNewLine _
& "Listen to hear a menu for the machine. Or even try hitting a digit on the phone" & vbNewLine _
& "to try and invoke a menu. You must go through all the sections that you choose," & vbNewLine _
& "this takes time to do, and repetition.. hence the " & Chr$(34) & "cracking" & Chr$(34) & " term coined in the title." & vbNewLine

Helpfrm.Text1 = Helpfrm.Text1 & "Double Click the individual lists to remove duplicates from them." & vbNewLine _
& "They are not sorted so that you might have an idea of when a particular code had worked."" & vbNewLine" _
& "" & vbNewLine _
& "There is still much improvement that can made on this. The tones have been" & vbNewLine _
& "tested with just dialing ordinary phone numbers. Not to crack answering machines." & vbNewLine _
& "Point is, the phone recognizes each tone clearly.  I have had sucess with the speed that they are" & vbNewLine _
& "played as well. This might not be true for everyone, so there are no guarantees." & vbNewLine _
& "" & vbNewLine _
& "If anyone can organize the numbers for the 3 digit code in a more efficient way," & vbNewLine _
& "or would like to fix this program up, you are more than welcome to." & vbNewLine _
& "There is definitely room for improvement, I just don't have the time." & vbNewLine _
& "You may distribute this, so long as you leave credit to me and keep the disclaimer" & vbNewLine _
& "that this is strictly meant for educational purposes." & vbNewLine

Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & vbNewLine
Helpfrm.Text1 = Helpfrm.Text1 & "Copyright " & Chr$(169) & "1998 by Phishbowler. All Rights Reserved."



End Sub

Private Sub Label1_Click()
Shell ("Start http://come.to/phishbowler"), vbMinimizedFocus

End Sub
