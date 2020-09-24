VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConventional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub QuicksortConven(SortElems() As String, ByVal xFrom As Long, ByVal xThru As Long, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare)

  'Sorts (part of) a table of strings by moving the strings

  Dim xLeft As Long
  Dim xRite As Long
  Dim Pivot As String 'this receives table elements

    Do While xFrom < xThru  'we have something to sort (@ least two elements)
        xLeft = xFrom
        xRite = xThru
        Pivot = SortElems(xLeft) 'get pivot elem and make room
        Do
            Do Until xRite = xLeft
                If StrComp(SortElems(xRite), Pivot, Compare) < 0 Then 'is less than pivot
                    SortElems(xLeft) = SortElems(xRite) 'so move it to the left
                    xLeft = xLeft + 1 'leave the item just moved alone for now
                    Exit Do 'loop 
                End If
                xRite = xRite - 1
            Loop
            Do Until xLeft = xRite
                If StrComp(SortElems(xLeft), Pivot, Compare) > 0 Then 'is greater than pivot
                    SortElems(xRite) = SortElems(xLeft) 'so move it to the right
                    xRite = xRite - 1 'leave the item just moved alone for now
                    Exit Do 'loop 
                End If
                xLeft = xLeft + 1
            Loop
        Loop Until xLeft = xRite
        'now the indexes have met and all bigger items are to the right and all smaller items are left
        SortElems(xRite) = Pivot 'insert Pivot and sort the two areas left and right of it
        If xLeft - xFrom < xThru - xRite Then 'smaller part 1st to reduce recursion depth
            xLeft = xFrom
            xFrom = xRite + 1
            xRite = xRite - 1
          Else 'NOT XLEFT...
            xRite = xThru
            xThru = xLeft - 1
            xLeft = xLeft + 1
        End If
        If xLeft < xRite Then 'smaller part is not empty...
            QuicksortConven SortElems, xLeft, xRite, Compare '...so sort it
        End If
    Loop

End Sub

':) Ulli's VB Code Formatter V2.23.12 (2007-Mrz-25 13:48)  Decl: 1  Code: 51  Total: 52 Lines
':) CommentOnly: 4 (7,7%)  Commented: 14 (26,9%)  Empty: 6 (11,5%)  Max Logic Depth: 5
