Attribute VB_Name = "modMain"
                                        
                                        Option Explicit
'______________________________________________________________________________________________


   
    

Public Sub AddtoList(lvStore As ListView, Optional ByVal Col2 As String = "Server", _
                                          Optional ByVal Col3 As String = "App Started", _
                                          Optional ByVal Col4 As String = "")
                                          
' If there is an error handle it down at Err_Handler
    On Error GoTo Err_Handler
' oItem refers to our listbox that we are adding info to
    Dim oItem As ListItem
' So we don't have an over-load of msgs this will check and clear
    If lvStore.ListItems.Count > 5000 Then
    
        lvStore.ListItems.Clear ' Clearing  Messages now
        
        DoEvents ' Now do this
        
        Set oItem = lvStore.ListItems.Add(, , Now()) ' Store the time-stamp in 1st Column
        
        oItem.SubItems(1) = "Server" ' Column 2
        
        oItem.SubItems(2) = "List reached 5000 entries, cleared." ' and our Last Column 3
        
        oItem.EnsureVisible ' SelfExplanitory
        Dim sb As StatusBar
        
    End If
    ' ^ Same as above ^
    Set oItem = lvStore.ListItems.Add(, , Now()) ' stamp the time
    oItem.SubItems(1) = Col2 ' where did this come from
    oItem.SubItems(2) = Col3 ' Whats the background like
    oItem.SubItems(3) = Col4
    
    oItem.EnsureVisible
    
    
Err_Handler: ' Where we go on_Error
    Set oItem = Nothing
    Exit Sub
    
End Sub

Public Sub sBar(sb As StatusBar, Optional ByVal Pan1 As String = "Local IP", _
                                 Optional ByVal Pan2 As String = "9456", _
                                 Optional ByVal Pan3 As String = "No Current Messages")
                                 
    On Error GoTo errHandler:
                                 
    ' Handle the status bar
    sb.Panels(1).Text = Pan1 'Panel one
    
    sb.Panels(2).Text = Pan2 'Panel two
    
    sb.Panels(3).Text = Pan3 'Panel three
    
errHandler:
    Exit Sub
End Sub
