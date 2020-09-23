VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Free VB stuff!!"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1620
      Left            =   7875
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.PictureBox picProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2025
      ScaleHeight     =   255
      ScaleWidth      =   4530
      TabIndex        =   2
      Top             =   90
      Width           =   4560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Get me free VB stuff!!"
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1770
   End
   Begin VB.ListBox List1 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   945
      Width           =   6675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Balloon"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   1530
      TabIndex        =   4
      Top             =   495
      Width           =   5100
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim m_mainpages() As String

Private WithEvents cHtml As cHtmlDoc
Attribute cHtml.VB_VarHelpID = -1
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
 
Private Sub cHtml_docCreateProgress(progress As String)

  Caption = progress
  
End Sub

Private Sub cHtml_docCreationFailed()

   MsgBox "Error!  check your internet connection"

End Sub

Private Sub cHtml_linksDone()

  If IE_state = START_PAGE Then
     Dim lcnt   As Long, lcnt2    As Long
     
     'set the scale vals for the picProgress control
     picProgress.ScaleWidth = UBound(m_mainpages) - 1
     picProgress.DrawWidth = 2
     picProgress.ScaleHeight = 10
     
     IE_state = SOFTWARE_PAGE
     
     'go through each of the page links extracted from the
     'first page
     For lcnt = 0 To UBound(m_mainpages)
         cHtml.document_from_url m_mainpages(lcnt), , True
          'increment progress var
         For lcnt2 = 0 To 10
            picProgress.Line (0, lcnt2)-(lcnt, lcnt2), RGB(250 - (lcnt2 * 10), 150, 150)
         Next lcnt2
     Next lcnt
  
     picProgress.Cls
     Command1.Enabled = True
  End If
 
  
End Sub

Private Sub cHtml_linksReady(aLink As MSHTML.DispHTMLAnchorElement)
   
If IE_state = START_PAGE Then
   'these is the list of pages that have all the free vb stuff
   If InStr(1, aLink, "http://www.winsite.com/tech/vb/page") Then
      If Not isDuplicate(List1, aLink.href) Then
         Dim upper  As Long
         
         If code.IsArray(m_mainpages) Then
           upper = UBound(m_mainpages) + 1
         Else
           upper = 0
         End If
         
         ReDim Preserve m_mainpages(upper)
         m_mainpages(upper) = aLink.href
      End If
   End If
 
 ElseIf IE_state = SOFTWARE_PAGE Then
     
     'here comes the free stuff!!
     If InStr(1, aLink, "http://www.winsite.com/bin/Info?") Then
      If Not isDuplicate(List1, aLink.href) Then
         List1.AddItem aLink.outerText
         List2.AddItem aLink.href
         Label1 = List1.ListCount & "(mostly) Free items listed !!"
      End If
   End If
 End If
 
End Sub

 
Private Sub Command1_Click()
 
 'go to the start page which holds the long list
 'of all the other pages that have all the free stuff
 MsgBox "Please be patient." & vbCrLf & _
        "It takes several minutes to download and parse" & vbCrLf & _
        "the links to hundreds of (mostly) FREE vb related items."
 IE_state = START_PAGE
 Debug.Print "command1 clicked"
 Command1.Enabled = False
 cHtml.document_from_url "http://www.winsite.com/tech/vb/index.html", , True
 
End Sub

Private Sub Form_Load()

  Set cHtml = New cHtmlDoc
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set cHtml = Nothing
End Sub

 
Private Sub List1_DblClick()
   'if the user has hilighted a free vb item
   'then create a new IE object if its not already there
   'and "insert it into" frmStuff
  If List1.SelCount > 0 Then
  
    If frmStuff.Visible = False Then
       frmStuff.Show vbModeless, Me
    End If
    
    frmStuff.IE.navigate List2.List(List1.ListIndex)
  End If

End Sub
