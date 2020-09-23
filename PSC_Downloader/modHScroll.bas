Attribute VB_Name = "modHScroll"
Option Explicit

'If the text in a listbox is likely to exceed the width of the listbox, then use the following code to add items to the listbox. By using the routine below a horizontal scroll bar will be displayed if the width of the added item exceeds the width of the listbox.

Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'Purpose     :  Adds items to a listbox and if neccessary sets the
'               width of the horizontal scroll bar to the maximum width of the
'               items in the listbox.
'Inputs      :  lbListbox                   The listbox to add the item to.
'               sItemText                   The text to add to the listbox.
'               [iIndex]                    The position within the object where the new item or row is placed.
'Outputs     :  Returns True on success
'Author      :  Andrew Baker
'Date        :  10/02/2001 15:50
'Notes       :
'Revisions   :
'Assumptions :

Function ListboxAddItem(lbListbox As ListBox, sItemText As String, Optional iIndex As Integer = -1) As Boolean
    Dim fTextWidth As Single, fExistScrollWidth As Single
    Dim oParentFont As StdFont
    Const LB_SETHORIZONTALEXTENT = &H194, LB_GETHORIZONTALEXTENT = &H193

    On Error Resume Next

    'Add item to listbox
    If iIndex > -1 Then
        lbListbox.AddItem sItemText, iIndex
    Else
        lbListbox.AddItem sItemText
    End If

    'Store the form's original font
    Set oParentFont = lbListbox.Parent.Font
    'Set the form's font to the listbox's font
    Set lbListbox.Parent.Font = lbListbox.Font
    'Get width of text on the form
    fTextWidth = lbListbox.Parent.TextWidth(sItemText & " ")        'Extra space allows for vertical scroll bar
    'Restore the form's font
    Set lbListbox.Parent.Font = oParentFont
    
    'Get the width of the existing scroll bar
    fExistScrollWidth = SendMessageA(lbListbox.hwnd, LB_GETHORIZONTALEXTENT, 0, 0)
    
    If lbListbox.Parent.ScaleMode = vbTwips Then
        'Change twips to pixels
        fTextWidth = fTextWidth / Screen.TwipsPerPixelX
    End If
    
    If fTextWidth > fExistScrollWidth Then
        'Increase width of scroll bar
        Call SendMessageA(lbListbox.hwnd, LB_SETHORIZONTALEXTENT, fTextWidth, 0)
    End If
    ListboxAddItem = (Err.Number = 0)
End Function


'Purpose     :  Modifies the text of an item in a listbox and if neccessary sets the
'               width of the horizontal scroll bar to the maximum width of the
'               items in the listbox.
'Inputs      :  lbListbox                   The listbox to update the item in.
'               sNewItemText                The new text for the item in the listbox.
'               [iIndex]                    The index of the item to update within the listbox.
'Outputs     :  Returns True on Success
'Author      :  Andrew Baker
'Date        :  10/02/2001 15:50
'Notes       :
'Revisions   :
'Assumptions :

Function ListboxUpdateItem(lbListbox As ListBox, sNewItemText As String, iIndex As Integer) As Boolean
    Dim fTextWidth As Single, fExistScrollWidth As Single
    Dim oParentFont As StdFont
    Const LB_SETHORIZONTALEXTENT = &H194, LB_GETHORIZONTALEXTENT = &H193
    
    'Add item to listbox
    On Error GoTo ErrFailed
    If lbListbox.List(iIndex) <> sNewItemText Then
        lbListbox.List(iIndex) = sNewItemText
        'Get width of text
        Set oParentFont = lbListbox.Parent.Font
        Set lbListbox.Parent.Font = lbListbox.Font
        fTextWidth = lbListbox.Parent.TextWidth(sNewItemText & " ")        'Extra space allows for vertical scroll bar
        Set lbListbox.Parent.Font = oParentFont
        fExistScrollWidth = SendMessageA(lbListbox.hwnd, LB_GETHORIZONTALEXTENT, 0, 0)
        
        If lbListbox.Parent.ScaleMode = vbTwips Then
            'Change twips to pixels
            fTextWidth = fTextWidth / Screen.TwipsPerPixelX
        End If
        
        If fTextWidth > fExistScrollWidth Then
            'Increase width of scroll bar
            Call SendMessageA(lbListbox.hwnd, LB_SETHORIZONTALEXTENT, fTextWidth, 0)
        End If
    End If
    ListboxUpdateItem = True

    Exit Function
    
ErrFailed:
    Debug.Print "Error in ListboxAddItem: " & lbListbox.Name & " Description: " & Err.Description
    ListboxUpdateItem = False
End Function

