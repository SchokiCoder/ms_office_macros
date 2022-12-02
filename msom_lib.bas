Attribute VB_Name = "msom_lib"
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


Function write_list_csv(path As String, list() As String)
    Dim i As Integer

    ' if file already exists, delete
    If Len(dir(path)) <> 0 Then
        Kill path
    End If
    
    ' write
    Open path For Append As #1
    
    For i = LBound(list) To (UBound(list) - 1)
        Write #1, list(i)
    Next
    
    Close #1
End Function

Function str_arr_insert_at(arr() As String, ByVal index As Integer, str As String)
    Dim i As Integer
    Dim arr_cpy() As String
    
    arr_cpy = arr
    
    ReDim Preserve arr_cpy(UBound(arr_cpy) + 1)
    
    For i = UBound(arr_cpy) To index
        arr_cpy(i) = arr_cpy(i - 1)
    Next
    
    arr_cpy(index) = str
    
    str_arr_insert_at = arr_cpy
End Function

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

