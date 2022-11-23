Attribute VB_Name = "msom_outlook_lib"
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


Sub write_list_csv(path As String, list() As String)
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
End Sub

