Attribute VB_Name = "outlook_msom"
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


Sub forward_new_task()
    Dim fwd As Outlook.MailItem
    
    ' get cur selected item, forward
    If Application.ActiveExplorer.Selection.Count <= 0 Then
        MsgBox "Please, first select an email to forward"
        Exit Sub
    End If
    
    Set fwd = Application.ActiveExplorer.Selection.Item(1).Forward
    
    ' add recipent
    fwd.Recipients.Add (outlook_msom_cfg.FORWARD_TASK_RECIPENT)
    
    ' find last filled line
    Dim i As Integer
    Dim line_empty As Boolean
    
    line_empty = True
    
    i = 1
    
    Do While i < Len(fwd.Body)
        ' if new-line found
        If Mid(fwd.Body, (Len(fwd.Body) - i), 1) = Chr(10) Then
            
            'if line had been filled
            If line_empty = False Then
                Exit Do
            End If
            
        ' else if not tab or space found
        ElseIf Mid(fwd.Body, (Len(fwd.Body) - i), 1) <> Chr(9) Or Mid(fwd.Body, (Len(fwd.Body) - i), 1) <> Chr(20) Then
            line_empty = False
        End If
        
        i = i + 1
    Loop
    
    ' go back to end of line
    Do While i > 0
        i = i - 1
        
        If Mid(fwd.Body, (Len(fwd.Body) - i), 1) = Chr(10) Then
            Exit Do
        End If
    Loop
    
    ' shorten body
    fwd.Body = Mid(fwd.Body, 1, (Len(fwd.Body) - i))
    
    ' find tail
    Dim line_count As Integer
    
    i = Len(fwd.Body) - 1
    
    Do While i > 0
        If Mid(fwd.Body, i, 1) = Chr(10) Then
            line_count = line_count + 1
            
            If line_count >= outlook_msom_cfg.FORWARD_TASK_TAIL Then
                Exit Do
            End If
        End If
        
        i = i - 1
    Loop
    
    ' remove tail
    fwd.Body = Mid(fwd.Body, 1, i)
    
    ' prepare signature
    ' MS won't allow me to just access the users saved signatures, so now it is a const in the config...
    ' MS can't express a linebreak as a const char, so now it gets replaced at runtime...
    ' @MS: Stop using backslashes for paths and start using them for goddamn escape sequences!
    Dim sign As String
    
    sign = outlook_msom_cfg.FORWARD_TASK_SIGNATURE
    sign = Replace(sign, outlook_msom_cfg.LINEBREAK, Chr(10))
    
    ' add signature
    fwd.Body = fwd.Body & Chr(10) & sign
    
    ' show message
    fwd.GetInspector.Display
    
End Sub


Sub export_dir_msgs_field()
    Dim dir As Outlook.Folder
    
    ' IN DEVELOPMENT, exit
    MsgBox "This macro is not done yet"
    Exit Sub
    
    ' goto target dir
    Set dir = Application.Session.Folders.Item(outlook_msom_cfg.EXPORT_USER).Folders.Item(outlook_msom_cfg.EXPORT_DIR)
    
    ' if dir is empty, msgbox and exit sub
    If dir.Items.Count = 0 Then
        MsgBox "No messages found"
        Exit Sub
    End If
    
    ' iterate mails
    Dim i As Integer
    Dim msg As Outlook.MailItem
    Dim field As String
    Dim field_list() As String
    
    Set msg = dir.Items.GetFirst
    
    For i = 0 To 1 'dir.Items.Count <--- UNCOMMENT ---
        ' save field to array
        field = Split(msg.Body, outlook_msom_cfg.EXPORT_DELIM)(outlook_msom_cfg.EXPORT_FIELD)
        
        ReDim Preserve field_list(i + 1)
        field_list(i) = field
        
        ' next
        Set msg = dir.Items.GetNext
    Next
    
    ' write csv file
    'Dim cells() As String
    'ReDim cells(Len(field_list), 0)
    
    'For i = 0 To LBound(field_list)
    '    cells(i, 0) = field_list(i)
    'Next
    
    'write_csv "export.csv", cells
    'Function write_csv(filepath As String, cells() As String)
    'End
    
    
    
    ' NOW, write an Excel macro that imports it all and checks for validity etc...
End Sub

