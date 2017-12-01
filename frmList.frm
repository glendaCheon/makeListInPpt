VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmList 
   Caption         =   "목차 만들기"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7635
   OleObjectBlob   =   "frmList.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMakeList_Click()
    
    Dim title As String
    Dim pageNum As Integer
    Dim forListDot As String
    Dim tbxName As String
    
    Dim allPresentationCnt As Integer
    Dim currSlide As Slide
    Dim currTitle As String, prevTitle As String
    Dim currTbx As TextBox
    
    tbxName = "텍스트 개체 틀 1"
    forListDot = vbTab & "...................."
    
    allPresentationCnt = ActivePresentation.Slides.Count
    
    For i = 1 To allPresentationCnt
        Set currSlide = ActivePresentation.Slides(i)
        
        If isShapeExists(tbxName, currSlide.shapes) Then
            pageNum = ActivePresentation.Slides(i).SlideNumber
            currTitle = ActivePresentation.Slides(i).shapes(tbxName).TextFrame.TextRange.Text
            
            If currTitle <> prevTitle Then
                title = title & currTitle & forListDot & pageNum & vbLf
                prevTitle = currTitle
            End If
        End If
    Next
        
    
    
    Me.tbxListResult.Text = title
    
    
End Sub
Function isShapeExists(shpName As String, shapes As shapes)
    Dim currShape As shape
    Dim result As Boolean
    
    result = False

    For Each currShape In shapes
        If currShape.Name = shpName Then
            result = True
            Exit For
        End If
    Next
    
    isShapeExists = result
        
End Function
    

