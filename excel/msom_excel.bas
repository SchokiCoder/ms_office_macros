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


Function str_arr_unique(str_arr() As String)
    Dim seen_vals() As String
    Dim i As Integer
    
    For i = LBound(str_arr) To UBound(str_arr)
        ' if value already seen, remove from arr
    Next
    
    str_arr_unique = str_arr
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
        Set header_item = header.Item(1, i)
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
    Dim found_cell As Excel.Range
    Dim answer As Integer
    
    ' open import file
    Open msom_excel_cfg.IMPORT_PATH For Input As #1
    
    ' I am about to use the goto statement because MS can't afford a continue-like keyword.
    ' I know, that i could rewrite it in a way that it doesn't need that.
    ' However i think it is amusing to see how MS just fails all over the place.
    
    ' read lines
	read into str arr and then use my unique function
	then go over each line
	
    Do While Not EOF(1)
MSCouldntAffordContinueKeyword:
        Input #1, line
        
        ' orient_col: find cell that matches field
        Set found_cell = orient_col.Find(line)
        
        ' if not found, skip
        If found_cell Is Nothing Then
            GoTo MSCouldntAffordContinueKeyword
        End If
        
        ' if cell in active column is already filled
        If active_col.Item(found_cell.Row).Value <> "" Then
            ' ask user if overwrite
            active_col.Item(found_cell.Row).Select
            answer = MsgBox("The cell " & active_col.Item(found_cell.Row).Address & " contains '" & active_col.Item(found_cell.Row).Value & "'" & Chr(10) & "Do you want to overwrite?", vbQuestion + vbYesNo)
            
            If answer = vbYes Then
                ' set cell in active column
                active_col.Item(found_cell.Row).Value = msom_excel_cfg.IMPORT_VALUE
            End If
        Else
            ' set cell in active column
            active_col.Item(found_cell.Row).Value = msom_excel_cfg.IMPORT_VALUE
        End If
    Loop
    
    ' close
    Close #1
End Sub

