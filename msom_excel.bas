Attribute VB_Name = "msom_excel"
' ms_office_macros
' Copyright (C) 2022  Andy Frank Schoknecht
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along
' with this program; if not see
' <https://www.gnu.org/licenses/old-licenses/gpl-2.0.html>.

Sub flag_col_cells_by_import_list()
    Dim i As Integer
    Dim header As Excel.Range
    Dim header_item As Excel.Range
    Dim target_day
    Dim found_target As Boolean
    Dim sheet As Excel.Worksheet
    
    found_target = False
    target_day = Date + msom_excel_cfg.FLAG_BY_IMPORT_CUR_DAY_OFFSET
    
    ' in header, find cur date with configured offset
    Set sheet = Application.ActiveSheet
    Set header = sheet.Range(msom_excel_cfg.FLAG_BY_IMPORT_HEADER_RANGE)
    
    For i = 1 To header.Count
        Set header_item = header.Item(1, i)
        If header_item = target_day Then
            found_target = True
            Exit For
        End If
    Next
    
    ' It just occured to me,
    ' that MS created a language with variable declaration and types,
    ' that does not raise any red flag upon seeing an unknown variable name in an if expression.
    ' ( If target_found = False ... )
    ' Even Python would have caught on to this.
    
    ' if target not found, msg and exit
    If found_target = False Then
        MsgBox "Target '" & target_day & "' could not be found"
        Exit Sub
    End If
    
    Dim target_col, orient_col As Excel.Range
    Dim rng_a, rng_b As String
    
    ' get active column
    rng_a = Split(header_item.Address, "$")(1) & "1"
    rng_b = Mid(rng_a, 1, 1) & msom_excel_cfg.FLAG_BY_IMPORT_LAST_ROW
    
    Set target_col = sheet.Range(rng_a & ":" & rng_b)
    
    ' get orientation column
    rng_a = msom_excel_cfg.FLAG_BY_IMPORT_ORIENTATION_COL & Mid(rng_a, 2)
    rng_b = msom_excel_cfg.FLAG_BY_IMPORT_ORIENTATION_COL & Mid(rng_b, 2)
    
    Set orient_col = sheet.Range(rng_a & ":" & rng_b)
    
    Dim line As String
    Dim content() As String
    
    ' MS allows you to basically create an empty vector, without them knowing that it is empty...
    ReDim Preserve content(0)
    ' let that sink in
    
    ' open import file
    Open msom_excel_cfg.FLAG_BY_IMPORT_PATH For Input As #1
    
    ' read lines
    Do While Not EOF(1)
        Input #1, line
        ReDim Preserve content(UBound(content) + 1)
        content(UBound(content) - 1) = line
    Loop
    
    ' close file
    Close #1
    
    ' eliminate redundancies, remove empty lines
    content = msom_lib.str_arr_unique(content)
    content = msom_lib.str_arr_noempty(content)
    
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
        
        ' if cell is striked-through by border-formatting
        If target_col.Item(found_cell.Row).Borders(xlDiagonalUp).LineStyle <> XlLineStyle.xlLineStyleNone Or target_col.Item(found_cell.Row).Borders(xlDiagonalDown).LineStyle <> XlLineStyle.xlLineStyleNone Or target_col.Item(found_cell.Row).Borders(xlInsideVertical).LineStyle <> XlLineStyle.xlLineStyleNone Or target_col.Item(found_cell.Row).Borders(xlInsideHorizontal).LineStyle <> XlLineStyle.xlLineStyleNone Then
            ' ask if remove strike and flag
            target_col.Item(found_cell.Row).Select
            answer = MsgBox("The cell " & target_col.Item(found_cell.Row).Address & " is striked through." & Chr(10) & "Remove strike and set flag?", vbYesNoCancel)
            
            ' remove strike, flag cell in active column
            If answer = vbYes Then
                target_col.Item(found_cell.Row).Borders(xlDiagonalUp).LineStyle = XlLineStyle.xlLineStyleNone
                target_col.Item(found_cell.Row).Borders(xlDiagonalDown).LineStyle = XlLineStyle.xlLineStyleNone
                target_col.Item(found_cell.Row).Borders(xlInsideVertical).LineStyle = XlLineStyle.xlLineStyleNone
                target_col.Item(found_cell.Row).Borders(xlInsideHorizontal).LineStyle = XlLineStyle.xlLineStyleNone
                target_col.Item(found_cell.Row).Value = msom_excel_cfg.FLAG_BY_IMPORT_FLAG
                
            ElseIf answer = vbCancel Then
                Exit For
            End If
        
        ' if cell in active column is already filled
        ElseIf target_col.Item(found_cell.Row).Value <> "" Then
            ' ask user if overwrite
            target_col.Item(found_cell.Row).Select
            answer = MsgBox("The cell " & target_col.Item(found_cell.Row).Address & " contains '" & target_col.Item(found_cell.Row).Value & "'" & Chr(10) & "Do you want to overwrite?", vbQuestion + vbYesNoCancel)
            
            ' flag cell in active column
            If answer = vbYes Then
                target_col.Item(found_cell.Row).Value = msom_excel_cfg.FLAG_BY_IMPORT_FLAG
                
            ElseIf answer = vbCancel Then
                Exit For
            End If
        Else
            ' flag cell in active column
            target_col.Item(found_cell.Row).Value = msom_excel_cfg.FLAG_BY_IMPORT_FLAG
        End If
MSsux_no_continue:
    Next
End Sub


