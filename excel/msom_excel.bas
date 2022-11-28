Attribute VB_Name = "msom_excel"
'    ms_office_macros
'    Copyright (C) 2022  Andy Frank Schoknecht
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.


Function str_arr_remove_at(arr() As String, ByVal index As Integer)
    Dim i As Integer
    Dim arr_cpy() As String
    
    arr_cpy = arr
    
    For i = index To (UBound(arr_cpy) - 1)
        arr_cpy(i) = arr_cpy(i + 1)
    Next i
    
    ReDim Preserve arr_cpy(UBound(arr_cpy) - 1)
    
    str_arr_remove_at = arr_cpy
End Function

Function str_arr_unique(arr() As String)
    Dim seen_vals() As String
    Dim i, cmp As Integer
    
    ReDim Preserve seen_vals(0)

    For i = LBound(arr) To UBound(arr)
        ' if value already seen, skip
        For cmp = LBound(seen_vals) To UBound(seen_vals)
            If seen_vals(cmp) = arr(i) Then
                GoTo MSsux_no_continue
            End If
        Next
        
        ' add value to seen vals
        ReDim Preserve seen_vals(UBound(seen_vals) + 1)
        seen_vals(UBound(seen_vals) - 1) = arr(i)
MSsux_no_continue:
    Next
    
    str_arr_unique = seen_vals
End Function

Function str_arr_noempty(arr() As String)
    Dim i As Integer
    Dim arr_cpy() As String
    
    arr_cpy = arr

    For i = LBound(arr_cpy) To UBound(arr_cpy)
        ' if value empty, remove from arr
        If Len(arr_cpy(i)) = 0 Then
            arr_cpy = str_arr_remove_at(arr_cpy, i)
        End If
    Next
    
    str_arr_noempty = arr_cpy
End Function

Sub import_dir_msgs_field()
    Dim i As Integer
    Dim header As Excel.Range
    Dim header_item As Excel.Range
    Dim target_day
    Dim sheet As Excel.Worksheet
    
    target_day = Date + msom_excel_cfg.IMPORT_CUR_DAY_OFFSET
    
    ' in header, find cur date with configured offset
    Set sheet = Application.ActiveSheet
    Set header = sheet.Range(msom_excel_cfg.IMPORT_HEADER_RANGE)
    
    For i = 1 To header.Count
        Set header_item = header.item(1, i)
        If header_item = target_day Then
            Exit For
        End If
    Next
    
    Dim active_col, orient_col As Excel.Range
    Dim rng_a, rng_b As String
    
    ' get active column
    rng_a = Split(header_item.Address, "$")(1) & "1"
    rng_b = Mid(rng_a, 1, 1) & msom_excel_cfg.IMPORT_LAST_ROW
    
    Set active_col = sheet.Range(rng_a & ":" & rng_b)
    
    ' get orientation column
    rng_a = msom_excel_cfg.IMPORT_ORIENTATION_COL & Mid(rng_a, 2)
    rng_b = msom_excel_cfg.IMPORT_ORIENTATION_COL & Mid(rng_b, 2)
    
    Set orient_col = sheet.Range(rng_a & ":" & rng_b)
    
    Dim line As String
    Dim content() As String
    
    ' MS allows you to basically create an empty vector, without them knowing that it is empty...
    ReDim Preserve content(0)
    ' let that sink in
    
    ' open import file
    Open msom_excel_cfg.IMPORT_PATH For Input As #1
    
    ' read lines
    Do While Not EOF(1)
        Input #1, line
        ReDim Preserve content(UBound(content) + 1)
        content(UBound(content) - 1) = line
    Loop
    
    ' close file
    Close #1
    
    ' eliminate redundancies, remove empty lines
    content = str_arr_unique(content)
    content = str_arr_noempty(content)
    
    ' import
    Dim found_cell As Excel.Range
    Dim answer As Integer
        
    For i = LBound(content) To UBound(content)
        ' orient_col: find cell that matches field
        Set found_cell = orient_col.Find(content(i))
        
        ' if not found, skip
        If found_cell Is Nothing Then
            GoTo MSsux_no_continue
        End If
        
        ' if cell in active column is already filled
        If active_col.item(found_cell.Row).Value <> "" Then
            ' ask user if overwrite
            active_col.item(found_cell.Row).Select
            answer = MsgBox("The cell " & active_col.item(found_cell.Row).Address & " contains '" & active_col.item(found_cell.Row).Value & "'" & Chr(10) & "Do you want to overwrite?", vbQuestion + vbYesNoCancel)
            
            If answer = vbYes Then
                ' set cell in active column
                active_col.item(found_cell.Row).Value = msom_excel_cfg.IMPORT_VALUE
            ElseIf answer = vbCancel Then
                Exit For
            End If
        Else
            ' set cell in active column
            active_col.item(found_cell.Row).Value = msom_excel_cfg.IMPORT_VALUE
        End If
MSsux_no_continue:
    Next
End Sub

