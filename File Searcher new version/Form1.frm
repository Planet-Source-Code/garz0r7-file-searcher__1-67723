VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "File Searcher :: By Garz0r7"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   10845
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView2 
      Height          =   2655
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Folder Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Folder"
         Object.Width           =   10585
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8281
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Folder"
         Object.Width           =   8820
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size (Kb)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   7320
      Top             =   4680
   End
   Begin VB.TextBox WhatSearch 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2160
      TabIndex        =   7
      Text            =   "read"
      Top             =   720
      Width           =   3855
   End
   Begin VB.TextBox SearchPath 
      BackColor       =   &H80000009&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2160
      TabIndex        =   4
      Text            =   "c:\"
      Top             =   240
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2040
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Left            =   5760
      Top             =   4440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Search.."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000007&
      Height          =   420
      Left            =   360
      TabIndex        =   1
      Top             =   6840
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   360
      TabIndex        =   0
      Top             =   7320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label y 
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label finded2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   11
      Top             =   9480
      Width           =   3615
   End
   Begin VB.Label finded1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   9480
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Click  A File Name (Or Folder Name ) To Open It's Directory !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   1440
      Width           =   10695
   End
   Begin VB.Label Label3 
      Caption         =   "Text to Find:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Path To Investigate:"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'BY Garz0r7
'ceid student,patra,greece
'24012007

Option Explicit
Const FOL = "******************"
Dim i, j, temp, pointer, flag As Integer  'timer2 variables
Dim s As String 'timer2 variable

Dim nul, nul2, effe As Integer

Dim pointer_from, pointer_to, files, counter1, counter2 As Long
Private Sub Command1_Click()
Dim folder_depth As Integer

counter1 = 0
counter2 = 0
List1.Clear
ListView1.ListItems.Clear
ListView2.ListItems.Clear
Command1.Enabled = False
pointer = 0




'Setup our Dir1 List
Dir1 = SearchPath

'Add the path of the folder in
'which the folder will take place!
'(you could here add more paths...)
List1.AddItem (Dir1)

'set the pointers
pointer_from = 0
pointer_to = List1.ListCount - 1

folder_depth = -1

Do
DoEvents
    'analyze folders with depth=folder_depth
    folder_depth = folder_depth + 1

    'Start analysis!
    nul = analyze(pointer_from, pointer_to)

Loop Until nul = 1 'no more folders to investigate!


Command1.Enabled = True 'ready for next search

End Sub
Function analyze(ByVal pfrom, ByVal pto)
Dim k1, k2 As Integer

'Analyze the area which pointers show
For k1 = pfrom To pto
DoEvents

    Dir1 = List1.List(k1) 'Search for folders this Path
    
    For k2 = 0 To Dir1.ListCount - 1
    DoEvents
    
        'Add the folders of this path into list1!
        List1.AddItem (Dir1.List(k2))
        
    Next

Next

'refresh pointers
pointer_from = pto + 1
pointer_to = List1.ListCount

If pointer_from <= pointer_to Then
    analyze = 0 'continue,there are folders to analyze
Else
    analyze = 1 'no folders left,end analysis!
End If

End Function

Private Sub Form_Load()
Form1.Show
y = last("dfdfdd")
End Sub
Private Sub Form_Unload(Cancel As Integer)

FORM2.Show


End Sub

Private Sub List3_Click()
Shell "explorer.exe " & List3, vbNormalFocus
End Sub

Private Sub ListView1_DblClick()


    Shell "explorer.exe " & ListView1.SelectedItem.SubItems(1), vbNormalFocus


End Sub

Private Sub ListView2_DblClick()
Shell "explorer.exe " & ListView2.SelectedItem.SubItems(1), vbNormalFocus
End Sub

Private Sub Timer1_Timer()
List2.List(2) = "Folders so far : " + Str(List1.ListCount)
List2.List(3) = "Time : " + Str(Int((Timer - t2) * 100) / 100) + " secs"

'this will make some effects!
effe = effe + 1
If effe < 3 Then
    label_wait.ForeColor = vbBlack
End If

If effe > 3 Then
    label_wait.ForeColor = vbRed
End If

If effe > 6 Then
    effe = 0
End If

End Sub


Private Sub Timer2_Timer()
'ok here is the good stuff..
'the file search take place simultaneously
'with the folder invastigation!
'that's why we are using timer and not function.

If List1.ListCount > pointer Then


'means that there are folders
'which have been invastigated but we
'havent yet search for files inside them.

temp = List1.ListCount

'for each of these folders
'make a search for the file
'we looking for

For i = pointer To temp - 1
DoEvents

File1 = List1.List(i) 'load folder's files into file1

If InStr(1, LCase(last(List1.List(i))), LCase(WhatSearch)) Then

    'founded a folder with the text we are searching!
    ListView2.ListItems.Add , , last(List1.List(i))
    ListView2.ListItems(ListView2.ListItems.Count).SubItems(1) = List1.List(i)
    'ListView2.ListItems(ListView1.ListItems.Count).SubItems(2) = FOL
    counter1 = counter1 + 1
    
End If

'refresh our file counter!!
'it holds all the files we have found
'in our search so for(look *** (above) for the const 4)

finded1 = Str(counter1) + " Folders "
finded2 = Str(counter2) + " Files "

    For j = 0 To File1.ListCount - 1
    DoEvents
        
    
        If InStr(1, LCase(File1.List(j)), LCase(WhatSearch)) Then
        flag = 2
            'FIND WHAT YOU LOOKING FOR!
            'ADD IT IN THE LIST.
            s = File1.List(j)
 
            ListView1.ListItems.Add , , s
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(1) = List1.List(i)
            On Error GoTo er:
            'bad filename!
            ListView1.ListItems(ListView1.ListItems.Count).SubItems(2) = Round(FileLen(List1.List(i) + "\" + s) / 1024, 2)
er:
            counter2 = counter2 + 1
        End If

    Next

Next


pointer = temp


End If


End Sub
Function last(ByVal s As String) As String
Dim st As String
Dim k As Integer
k = 0

'INPOUT:A STRING LIKE "C:/MICRO/TAKER/CD"
'OUTPUT:THE LAST WORD OF THE STRING.HERE:"CD"

y = s
While (InStr(1, Right(s, k), "\") = 0) And k <= Len(s)
st = Right(s, k)
k = k + 1
Wend

last = st
End Function
