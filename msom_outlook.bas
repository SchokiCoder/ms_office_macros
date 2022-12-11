Attribute VB_Name = "msom_outlook"
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

Sub forward_task()
    Dim mail As Outlook.MailItem
    
    ' get cur selected item, forward
    If Application.ActiveExplorer.Selection.Count <= 0 Then
        MsgBox "Please, first select an email to forward"
        Exit Sub
    End If
    
    Set mail = Application.ActiveExplorer.Selection.Item(1).Forward
    
    ' add recipent
    mail.Recipients.Add (msom_outlook_cfg.FORWARD_TASK_RECIPIENT)
    
    ' find tail marker
    Dim i As Integer
    Dim lines() As String
    Dim found_marker As Boolean
    Dim tail_line_at As Integer
    
    found_marker = False
    
    ' from first line to last, search tail marker
    lines = Split(mail.HTMLBody, Chr(10))
    
    For i = LBound(lines) To UBound(lines)
        If InStr(lines(i), msom_outlook_cfg.FORWARD_TASK_TAIL_MARKER) Then
            found_marker = True
            tail_line_at = i
            Exit For
        End If
    Next
    
    ' if no marker, msg and exit
    If found_marker = False Then
        MsgBox "No tail marker found, task can not be forwarded."
        Exit Sub
    End If
    
    ' delete lines, from marker line until tags: html, font, p
    For i = tail_line_at To UBound(lines)
        If InStr(lines(i), "</HTML>") Or InStr(lines(i), "</FONT>") Or InStr(lines(i), "</P>") Then
            Exit For
        End If
        
        lines = msom_lib.str_arr_remove_at(lines, i)
    Next
    
    ' insert signature
    lines = msom_lib.str_arr_insert_at(lines, tail_line_at, msom_outlook_cfg.FORWARD_TASK_SIGNATURE)
    
    ' build string from arr
    Dim newbody As String
    
    For i = LBound(lines) To UBound(lines)
        newbody = newbody & lines(i) & Chr(10)
    Next
    
    ' set mail htmlbody and set mail body to itself
    ' this is done because otherwise the formatting is changed in weird ways...
    ' MS knows why, i don't
    mail.HTMLBody = newbody
    mail.Body = mail.Body
    
    ' display
    mail.GetInspector.Display
End Sub

Sub export_dir_msgs_field()
    Dim dir As Outlook.Folder
    
    ' goto target dir
    Set dir = Application.Session.Folders.Item(msom_outlook_cfg.EXPORT_USER).Folders.Item(msom_outlook_cfg.EXPORT_DIR)
    
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
    
    
    ' Why i am about to use (i + 1) as index:
    ' Outlook.Folder.Items.Item(index) <- index begins at 1
    ' MS allows VBA writers to choose at which num the array begins.
    ' MS: JuSt UsE lBoUnD
    ' Just pick a damn starting number, like a sane lang-designer would
    
    
    ' for all msgs: get msg and field, save field to list
    For i = 0 To (dir.Items.Count - 1)
        Set msg = dir.Items.Item(i + 1)
        field = Split(msg.Body, msom_outlook_cfg.EXPORT_DELIM)(msom_outlook_cfg.EXPORT_FIELD)
        
        ReDim Preserve field_list(i + 1)
        field_list(i) = field
    Next
    
    ' write csv file
    write_list_csv msom_outlook_cfg.EXPORT_PATH, field_list
End Sub

