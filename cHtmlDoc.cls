VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHtmlDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Enum enElemType
   links = 0
End Enum

'events
Event docCreateProgress(progress As String)
Event docCreationFailed()
Event docReady(htmlDoc As HTMLDocument)
Event bodyReady(htmlBody As htmlBody)
Event linksReady(aLink As HTMLAnchorElement)
Event linksDone()
Event inputElementsReady(aInputElement As HTMLInputElement)
Event InputElementsDone()
Event tableCellsReady(aTableCell As HTMLTableCell)
Event TableCellsDone()
Event imagesReady(aImage As HTMLImg)
Event ImagesDone()

Dim m_start_time     As Long
 


Sub document_from_url(url As String, Optional timeoutVal As Long = 5000, _
                      Optional ReturnLinks As Boolean, _
                      Optional ReturnInputElements As Boolean, _
                      Optional returnTableCels As Boolean, _
                      Optional returnImages As Boolean)
 
 Dim HTML As New HTMLDocument
 Dim odoc As HTMLDocument
 
 Debug.Print "creating document from url  " & url
 'this is how we get a webpage without a browser or visible object
 Set odoc = HTML.createDocumentFromUrl(url, vbNullString)
  
 Do  'gives you feedback as to the page progress
    RaiseEvent docCreateProgress(odoc.readyState)
    
    'we need a timeout so we avoid endless loops in
    'case your internet connection is not cooperating
    If Not (timeout(timeoutVal)) Then
       If odoc.readyState = "complete" Then
          'you are ready to access the dozens of events/methods
          'of both the document and the documents body
          Debug.Print "document object set and ready"
          RaiseEvent docReady(odoc)
          RaiseEvent bodyReady(odoc.body)
          'return only the info we want to give our friend
          'intel a break
          If ReturnLinks = True Then Call GetAllLinks(odoc)
          If ReturnInputElements = True Then Call GetAllInputElements(odoc)
          If returnTableCels = True Then Call GetAllTableCells(odoc)
          If returnImages = True Then Call GetAllImages(odoc)
          'give our friend PONY a break too
          Set odoc = Nothing
          Set HTML = Nothing
          Exit Sub
       End If
    End If
 
    DoEvents
  Loop
 
  
End Sub

'returns all the links objects
Private Sub GetAllLinks(odoc As HTMLDocument)
 
 Dim i   As Integer
 
 For i = 0 To odoc.links.length - 1
   RaiseEvent linksReady(odoc.links(i))
 Next i
 
 RaiseEvent linksDone
 
End Sub

'returns all the input elements objects
Private Sub GetAllInputElements(odoc As HTMLDocument)
 
 Dim i   As Integer

 For i = 0 To odoc.getElementsByTagName("INPUT").length - 1
    RaiseEvent inputElementsReady(odoc.getElementsByTagName("INPUT")(i))
 Next i
 
 RaiseEvent InputElementsDone
 
End Sub

'returns all the table cell objects
Private Sub GetAllTableCells(odoc As HTMLDocument)
  
  Dim oCell As HTMLTableCell
  Dim oRow  As HTMLTableRow
  Dim i As Integer, ii As Integer
  
  With odoc
  
   For i = 0 To odoc.getElementsByTagName("TR").length - 1
   
       Set oRow = odoc.getElementsByTagName("TR")(i)
       
       For ii = 0 To oRow.cells.length - 1
          Set oCell = oRow.cells(ii)
          RaiseEvent tableCellsReady(oCell)
       Next ii
       
   Next i
   
   Set oCell = Nothing
   Set oRow = Nothing
   
   RaiseEvent TableCellsDone
   
  End With
  
End Sub

'returns all the image objects
Private Sub GetAllImages(odoc As HTMLDocument)

 Dim i As Integer
 
 For i = 0 To odoc.images.length - 1
    RaiseEvent imagesReady(odoc.images(i))
 Next i
 
 RaiseEvent ImagesDone
 
End Sub

'this is out function that acts as the loop opt out
'in case of failing website or internet connection
Private Function timeout(timeoutVal As Long) As Boolean
    
    Static bStarted    As Boolean
    Static start_time  As Long
     
    If Not (bStarted) Then
       bStarted = True
       start_time = GetTickCount
    End If
    
    If (GetTickCount - start_time) > timeoutVal Then
        timeout = True
        GoTo clean_up:
    End If
    
clean_up:
        bStarted = False
        start_time = 0
End Function

 
