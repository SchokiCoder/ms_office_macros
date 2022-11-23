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


' hands on your
'       ____    ____    _   _   ____   _    ____
'      /  __|  /    \  | \ | | |  __| | |  /  __|
'     |  /    |  /\  | |  \| | | |_   | | | |
'  -- | |     | |  | | |   | | |  _|  | | | |  _  --
'     |  \__  |  \/  | | |\  | | |    | | | |_| |
'      \____|  \____/  |_| \_| |_|    |_|  \____/
'


Const IMPORT_PATH = "export.csv"

' horizontal and one row only, format = begin:end
Const IMPORT_HEADER_RANGE = "B4:H4"

Const IMPORT_CUR_DAY_OFFSET = -1

Const IMPORT_LAST_ROW = "230"

Const IMPORT_ORIENTATION_COL = "A"

Const IMPORT_VALUE = "io"


' don't touch that darn
'       __    ____    _   _   ____     ____   ____
'      / _|  /    \  | | | | |  _ \   /  __| |  __|
'     / /   |  /\  | | | | | | |_) | |  /    | |_
'  -- \ \   | |  | | | | | | |    /  | |     |  _| --
'    _/ /   |  \/  | | \_/ | | |\ \  |  \__  | |__
'   |__/     \____/   \___/  |_| \_\  \____| |____|
'


Sub import_dir_msgs_field()
    Dim i As Integer
    Dim header As Excel.Range
    Dim header_item As Excel.Range
    Dim target_day
    Dim sheet As Excel.Worksheet
    
    target_day = Date + IMPORT_CUR_DAY_OFFSET
    
    ' in header, find cur date with configured offset
    Set sheet = Application.ActiveSheet
    Set header = sheet.Range(IMPORT_HEADER_RANGE)
    
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
    rng_b = Mid(rng_a, 1, 1) & IMPORT_LAST_ROW
    
    Set active_col = sheet.Range(rng_a & ":" & rng_b)
    
    ' get orientation column
    rng_a = IMPORT_ORIENTATION_COL & Mid(rng_a, 2)
    rng_b = IMPORT_ORIENTATION_COL & Mid(rng_b, 2)
    
    Set orient_col = sheet.Range(rng_a & ":" & rng_b)
    
    Dim line As String
    Dim found_cell As Excel.Range
    
    ' open import file
    Open IMPORT_PATH For Input As #1
    
    ' I am about to use the goto statement because MS can't afford a continue-like keyword.
    ' I know, that i could rewrite it in a way that it doesn't need that.
    ' However i think it is amusing to see how MS just fails all over the place.
    
    ' read lines
    Do While Not EOF(1)
MSCouldntAffordContinueKeyword:
        Input #1, line
        
        ' orient_col: find cell that matches field
        Set found_cell = orient_col.Find(line)
        
        ' if not found, skip
        If found_cell Is Nothing Then
            GoTo MSCouldntAffordContinueKeyword
        End If
        
        ' set cell in active column
        active_col.Item(found_cell.Row).Value = IMPORT_VALUE
    Loop
    
    ' close
    Close #1
End Sub

