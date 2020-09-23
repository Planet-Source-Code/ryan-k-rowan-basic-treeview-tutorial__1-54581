VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Basic Treeview Example"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Explanation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   5265
      TabIndex        =   2
      Top             =   735
      Width           =   7650
      Begin VB.Frame fraNode 
         Caption         =   "Your Node"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   150
         TabIndex        =   6
         Top             =   5670
         Visible         =   0   'False
         Width           =   7350
         Begin VB.CommandButton cmdGo 
            Caption         =   "GO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5940
            TabIndex        =   13
            Top             =   285
            Width           =   1245
         End
         Begin VB.TextBox txtText 
            Height          =   285
            Left            =   4245
            TabIndex        =   10
            Top             =   330
            Width           =   1530
         End
         Begin VB.TextBox txtKey 
            Height          =   285
            Left            =   1110
            TabIndex        =   8
            Top             =   315
            Width           =   1530
         End
         Begin VB.Label lblOutput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   870
            TabIndex        =   12
            Top             =   705
            Width           =   6375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Output"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   105
            TabIndex        =   11
            Top             =   690
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "[Text]"
            Height          =   195
            Left            =   3420
            TabIndex        =   9
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "[Key]"
            Height          =   195
            Left            =   285
            TabIndex        =   7
            Top             =   345
            Width           =   360
         End
      End
      Begin VB.TextBox txtExplain 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4755
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   405
         Width           =   7410
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<< Previous"
         Height          =   315
         Left            =   2130
         TabIndex        =   4
         Top             =   5310
         Width           =   1350
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Next >>"
         Height          =   315
         Left            =   4770
         TabIndex        =   3
         Top             =   5310
         Width           =   1350
      End
   End
   Begin MSComctlLib.TreeView treeview1 
      Height          =   6720
      Left            =   285
      TabIndex        =   0
      Top             =   885
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   11853
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblOne 
      AutoSize        =   -1  'True
      Caption         =   "Treeview1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1815
      TabIndex        =   1
      Top             =   525
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************
' Basic Treeview Example By Frawg
'
' Feel free to rip the code apart and play
' with it til you feel comfortable and
' knowledgable with a treeview control :)
'
'   I also took the liberty of setting apart the
'actual treeview code by comments out in the
'margin, might make it easier for everyone to find
'*********************************************

'**************************************
' Naming convention used:
'****************************
' s<name> = String
' n<name> = Integer
' b<name> = Boolean
' d<name> = double
' cmd<name> = Button Object
' lst<name> = Listview Object
' txt<name> = Textbox Object
' lbl<name> = Label Object
' tv<name> = Treeview Object
' ws<name> = Winsock Object
' fra<name> = Frame Object
'**************************************


'Globals
Public nEXPLAIN_PAGE As Integer 'used in the explain box for page choice via next/prev

Private Sub cmdGo_Click()
   On Error GoTo erh

    treeview1.Nodes.Add , , txtKey.Text, txtText.Text
    Exit Sub
    
erh:
   'instead of letting the program crash as it normally would if a bad key were entered,
   'i simply let it drop to this and throw a msgbox with the error instead
   MsgBox Error, vbExclamation, "Error!"
End Sub

Private Sub cmdNext_Click()
    nEXPLAIN_PAGE = nEXPLAIN_PAGE + 1
    ShowPage
End Sub

Private Sub cmdPrevious_Click()
    nEXPLAIN_PAGE = nEXPLAIN_PAGE - 1
    ShowPage
End Sub

Private Sub Form_Load()
    nEXPLAIN_PAGE = 8
    ShowPage
End Sub



Sub ShowPage()
   With txtExplain
    Select Case nEXPLAIN_PAGE
        Case 1:
            cmdPrevious.Enabled = False
            .Text = "Welcome to my attempt at teaching" & vbCrLf
            Buffer "you how to use a treeview control!" & vbCrLf
            ShowNext
        Case 2:
            cmdPrevious.Enabled = True
            .Text = "   Ok, the thing i notice most about people putting" & vbCrLf
            Buffer "tutorials out about Treeview controls (or most"
            Buffer "controls for that matter), is the lack of information"
            Buffer "on how to initially setup the control for use as YOU"
            Buffer "want it to be. So taking that into consideration, that"
            Buffer "is what we are going to do right now."
            Buffer vbCrLf & vbCrLf
            Buffer "Treeview1 Setup"
            Buffer "------------------------------------" & vbCrLf
            Buffer "1.) right-click on the control and choose: Properties"
            Buffer "2.) find 'Style' and choose: 7 - tvwTreelinesPlusMinusPictureText"
            Buffer "3.) find 'LineStyle' and choose: 1 - tvwRootLines"
            Buffer "4.) click 'Apply' and 'Ok'" & vbCrLf
            Buffer "Note ** You can't click on anything on the control at the moment,"
            Buffer "        I have it all set up for you."
            Buffer vbCrLf & vbCrLf & "Explanations of choices:"
            Buffer "-----"
            Buffer "tvwTreelinsPlusMinusPictureText - the most robust choice with the most options available.  Let's break it apart piece by piece."
            Buffer "   * tvw - simply means 'Treeview'"
            Buffer "   * Treelines - shows the treer lines from parent to child"
            Buffer "   * PlusMinus - shows a + or - box to expand or close a branch"
            Buffer "   * Picture - let's you use pictures (icons) at each branch point"
            Buffer "   * Text - let's you see any text your heart desires to put there"
            Buffer vbCrLf & "- tvwRootLines"
            Buffer "   * tvw - simply means 'Treeview'"
            Buffer "   * RootLines - shows lines from 'Root' or 'Main' branch"
            ShowNext
        Case 3:
            .Text = "   OK! now that the control is set up, let's explain a thing or two about the treeview command!"
            Buffer vbCrLf & vbCrLf & "To add a Node (branch):" & vbCrLf
            Buffer "treeview1.Nodes.Add [Relative], [Relationship], [Key], [Text], [Image], [SelectedImage]"
            Buffer vbCrLf & "Now most of you look at this and say, 'What the...'."
            Buffer "The first part 'treeview1.nodes.add' is pretty straight forward,"
            Buffer "you want to add a node to treeview1."
            Buffer "[Relative] - who's it's daddy? (exactly what it means too)"
            Buffer "[Relationship] - whether it is a child of a parent node"
            Buffer "[Key] - a UNIQUE text string that identifies this 'parent'"
            Buffer "[Text] - the text that you want to say in the treeview"
            Buffer "[Image] - not covered (not essential, just eye candy)  :)"
            Buffer "[SelectedImage] - same as [Image]"
            Buffer vbCrLf & "One thing to remember is that even child nodes can have children"
            Buffer "No, that does not make the parent node of the first child a grandparent node,"
            Buffer "but it does make the first child node a child node and a parent at the same time."
            ShowNext
        Case 4:
            .Text = " Your First Treeview Command!" & vbCrLf
            Buffer "-------------------------------" & vbCrLf
            Buffer "Step 1:"
            Buffer "    To set an initial node (root node) we use the following command:"
            Buffer "        treeview1.Nodes.Add , , ," & Chr(34) & "Root Node" & Chr(34)
            Buffer vbCrLf & "  As you can see in treeview1, that is what is displayed."
            Buffer "A simple node, no parents, no children, no key either!"
'*** CODE
            treeview1.Nodes.Clear
            treeview1.Nodes.Add , , , "Root Node"
'*** END CODE
            ShowNext
        Case 5:
            .Text = " Child Node!" & vbCrLf
            Buffer "-------------------------------" & vbCrLf
            Buffer " To set a child of that root node we made, we had to have used a [Key]."
            Buffer "So, what we need to do is refine our last command to accomodate:"
            Buffer vbCrLf & " treeview1.Nodes.Add , , " & Chr(34) & "MyKey" & Chr(34) & ", " & Chr(34) & "Root Node" & Chr(34)
            Buffer vbCrLf & "This will assign the text 'MyKey' as the [Key] of the parent node"
            Buffer "To set a child node that will expand/contract using our 'Root Node' as the parent node, we simply:"
            Buffer vbCrLf & "treeview1.Nodes.Add " & Chr(34) & "MyKey" & Chr(34) & ", tvwChild, , " & Chr(34) & "Child Node #1" & Chr(34)
            Buffer vbCrLf & "You will see that a + has shown up next to the 'Root Node'. If you click on the +, the branch will expand and show you the new child node!  In contrast, clicking on the - will contract the child node so it is hidden.  It is still there, just out of sight."
'*** CODE
            treeview1.Nodes.Clear
            treeview1.Nodes.Add , , "MyKey", "Root Node"
            treeview1.Nodes.Add "MyKey", tvwChild, , "Child Node #1"
'*** END CODE
            ShowNext
        Case 6:
            .Text = ""
            Buffer vbCrLf & "If you are paying attention, you will notice that we did not set a [Key] for the child node, this means that we will not be able to set a child node under it. We can modify that by using this instead of the previous child node command:"
            Buffer vbCrLf & "treeview1.Nodes.Add " & Chr(34) & "MyKey" & Chr(34) & ", tvwChild, " & Chr(34) & "ChildKey" & Chr(34) & ", " & Chr(34) & "Child Node #1 (with [Key])" & Chr(34)
            Buffer vbCrLf & "Now that the child node has a key, we can set a child node under it was well. Notice now that 'child node #1' now has a plus sign beside it as well."
            Buffer vbCrLf & "You can have as many child nodes cascading like this as you like, however, for every branch on the 'tree', the [Key] must be unique to each branch."
'*** CODE
            treeview1.Nodes.Clear
            treeview1.Nodes.Add , , "MyKey", "Root Node"
            treeview1.Nodes.Add "MyKey", tvwChild, "ChildKey", "Child Node #1 (with [Key])"
            treeview1.Nodes.Add "ChildKey", tvwChild, , "Child of a child node"
'*** END CODE
            ShowNext
        Case 7:
            .Text = ""
            Buffer "Now let's set mutliple root nodes and each with their own child. Root nodes have no [Relative], so no need to try and put one in there."
            Buffer vbCrLf
            Buffer "treeview1.Nodes.Add , , " & Chr(34) & "Root1" & Chr(34) & ", " & Chr(34) & "Root #1" & Chr(34)
            Buffer "treeview1.Nodes.Add " & Chr(34) & "Root1" & Chr(34) & ", tvwChild, " & Chr(34) & "Root1_Child1" & Chr(34) & ", " & Chr(34) & "Root #1's Child #1" & Chr(34)
            Buffer "treeview1.Nodes.Add " & Chr(34) & "Root1_Child1" & Chr(34) & ", tvwChild, , " & Chr(34) & "Root#1, Child#1, Child" & Chr(34)
            Buffer "treeview1.Nodes.Add " & Chr(34) & "Root1" & Chr(34) & ", tvwChild, " & Chr(34) & "Root1_Child2" & Chr(34) & ", " & Chr(34) & "Root #1's Child #2" & Chr(34)
            Buffer "treeview1.Nodes.Add, , " & Chr(34) & "Root2" & Chr(34) & ", " & Chr(34) & "Root #2" & Chr(34)
            Buffer "treeview1.Nodes.Add " & Chr(34) & "Root2" & Chr(34) & ", tvwChild, " & Chr(34) & "Root2_Child1" & Chr(34) & ", " & Chr(34) & "Root #2's Child #1" & Chr(34)
            Buffer "treeview1.Nodes.Add , , ,  " & Chr(34) & "Root #3" & Chr(34)
            Buffer vbCrLf & "Now anyone will tell ya that that looks a little cryptis in that format, let's set it up so that it makes a little more sense."
            Buffer vbCrLf
            Buffer "treeview1.Nodes.Add , , " & Chr(34) & "Root1" & Chr(34) & ", " & Chr(34) & "Root #1" & Chr(34)
            Buffer "    treeview1.Nodes.Add " & Chr(34) & "Root1" & Chr(34) & ", tvwChild, " & Chr(34) & "Root1_Child1" & Chr(34) & ", " & Chr(34) & "Root #1's Child #1" & Chr(34)
            Buffer "        treeview1.Nodes.Add " & Chr(34) & "Root1_Child1" & Chr(34) & ", tvwChild, , " & Chr(34) & "Root#1, Child#1, Child" & Chr(34)
            Buffer "    treeview1.Nodes.Add " & Chr(34) & "Root1" & Chr(34) & ", tvwChild, " & Chr(34) & "Root1_Child2" & Chr(34) & ", " & Chr(34) & "Root #1's Child #2" & Chr(34)
            Buffer "treeview1.Nodes.Add, , " & Chr(34) & "Root2" & Chr(34) & ", " & Chr(34) & "Root #2" & Chr(34)
            Buffer "    treeview1.Nodes.Add " & Chr(34) & "Root2" & Chr(34) & ", tvwChild, " & Chr(34) & "Root2_Child1" & Chr(34) & ", " & Chr(34) & "Root #2's Child #1" & Chr(34)
            Buffer "treeview1.Nodes.Add , , ,  " & Chr(34) & "Root #3" & Chr(34)
            Buffer vbCrLf & "   Indenting the nodes as they would appear in the treeview adds a small level of readability to the code. You can quickly find any node that may be giving you a problem and isolate it without having to search through the entire list.  Treeviews can have an extremely high number of nodes.  Imagine putting an entire text book in a treeview! How many nodes would you have for that??"
'*** CODE
            treeview1.Nodes.Clear
            treeview1.Nodes.Add , , "Root1", "Root #1"
                treeview1.Nodes.Add "Root1", tvwChild, "Root1_Child1", "Root #1's Child #1"
                    treeview1.Nodes.Add "Root1_Child1", tvwChild, , "Root#1, Child#1, Child"
                treeview1.Nodes.Add "Root1", tvwChild, "Root1_Child2", "Root #1's Child #2"
            treeview1.Nodes.Add , , "Root2", "Root #2"
                treeview1.Nodes.Add "Root2", tvwChild, "Root2_Child1", "Root #2's Child #1"
            treeview1.Nodes.Add , , , "Root #3"
'*** END CODE
            fraNode.Visible = False
            ShowNext
        Case 8:
            .Text = ""
            Buffer "    Ok, now i think it may be time for you to try to add a root node on your own! I have cleared the treeview (using treeview1.Nodes.Clear) for you. Notice at the bottom that the frame 'Your Node' is now visible.  Fill in the fields and click 'GO'."
            Buffer vbCrLf & "Remember, the format for a treeview command is:"
            Buffer "treeview1.Nodes.Add [Relative], [Relationship], [Key], [Text], [Image], [SelectedImage]"
            Buffer vbCrLf & "*** Note - you can double click on any node to view it's key"
'*** CODE (sort of)
'   the code from this is in: Private Sub cmdGo_Click()
'   if listed correctly, all the way at the top
'*** END CODE (sort of)
            fraNode.Visible = True
            treeview1.Nodes.Clear
            cmdNext.Enabled = True
        Case 9:
            .Text = ""
            Buffer "    Now you are ready to start using your treeview talent on your own. Feel free to rip this program apart and use what you can or need to from it. I hope this little tutorial was helpful to at least 1 person, then it was all worth it :)"
            Buffer vbCrLf & vbCrLf & "  If you have any questions, suggestions, modifications you wish to let me know about, simply email me at the address provided below and I will answer back as soon as possible. And before you ask:  Yes, I actually read my email. I have no idea where i find the time, but I manage ;)"
            Buffer vbCrLf & "Email:  frawg@dragons-ire.com"
            Buffer vbCrLf & vbCrLf & "Happy Coding!"
            fraNode.Visible = False
            cmdNext.Enabled = False
    End Select
   End With
End Sub

Sub Buffer(Data As String)
    With txtExplain
        .Text = .Text & Data & vbCrLf
    End With
End Sub

Sub ShowNext()
    Buffer vbCrLf & vbCrLf & "***** CLICK NEXT TO CONTINUE *****"
End Sub

Private Sub treeview1_DblClick()
    MsgBox "[Key]:  " & Chr(34) & treeview1.SelectedItem.Key & Chr(34), vbOKOnly, "Root node: " & treeview1.SelectedItem.Text
End Sub

Private Sub txtKey_Change()
    UpdateOutput
End Sub

Private Sub txtRelationship_Change()
    UpdateOutput
End Sub

Private Sub txtRelative_Change()
    UpdateOutput
End Sub

Private Sub txtText_Change()
    UpdateOutput
End Sub

Sub UpdateOutput()
    lblOutput.Caption = "treeview1.Nodes.Add , , " & Chr(34) & txtKey.Text & Chr(34) & ", " & Chr(34) & txtText.Text & Chr(34)
End Sub
