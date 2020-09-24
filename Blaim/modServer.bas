Attribute VB_Name = "modServer"
                                        Option Explicit
                                        
                                        
' Commands
    Private Enum Client_Commands
        UserID = 1                      ' So far this is for testing purposes
        Password = 2
        Message = 3
    End Enum
    
' Client Info
    Public Type Clients
        UserID As String
        Current_Socket As Integer
        Password As String
        Frozen As Boolean
    End Type
        
                                        
                                        


Public Function Check_Message(ByVal Msg As Variant)
                
    Dim Temp As String
    
    Temp = Trim(Left(Msg, 1))
    
                              
    Select Case Temp ' Handle the Received Message
        
        Case 0 ' !ID
            Change_ID Trim(Msg)
            
        Case 1 ' !IM
            'Req_Chat Trim(Msg)
        
        Case 2 ' !PW
             Change_PW Trim(Msg)
        
        Case 3 ' !In
            'Client_IN Trim(Msg)
        
        Case 4 ' !Out
            'Client_Out Trim(Msg)
        
        Case 5 ' !NA
             New_Account Trim(Msg)
        
        Case 6 ' !DA
            'Deactivate_Account
            
    End Select
        
End Function

Private Function Change_ID(ByVal Msg As Variant)

    Dim Temp As String
    
    Dim nTemp As String
    
    Dim oID As String
    
    Dim nID As String
    
    Dim oNum As Integer
    
    Dim nNum As Integer
    'Format
    '00000NewNameOldName
    
    'Store the Part of the string we need
    Temp = Left(Msg, 5) 'Left() & Right() functions read as fallows:
                                                            ' MyString = "Hello World"
                                                            ' MyStr = Right(MyString,5)
                                                            ' MyStr Now = "World"
    
    'Again we are reading out the wanted segments of the string
    nTemp = Right(Temp, 2) 'Start finding the Length of the old ID
    
    'Check to see if the first char is a Zero(0) if so then it is a # below 10
    If Left(nTemp, 1) = "0" Then
        
        'Now, that we figured that out, lets store the remainging # which we need here
        oNum = Int(Right(nTemp, 1))
    'If it is not a Zero(0) then it is a 10 or higher
    Else
        
        'So we will store both chars here
        oNum = Int(nTemp) 'The Int() Function takes a string and converts it to an Interger
    
    End If
    
    nTemp = Right(Temp, 4) 'Start finding the Length of the new ID
    
    If Left(nTemp, 1) = "0" Then
        
        Temp = Left(nTemp, 2)
        
        nNum = Right(Temp, 1)
        
    Else
        
        nNum = Int(Left(nTemp, 2))
    
    End If
    
    oID = Trim(Right(Msg, oNum)) 'Store the Clients old ID
    
    Temp = Trim(Right(Msg, oNum + nNum))
    
    nID = Trim(Left(Temp, nNum)) 'Store the Clients new ID
    
    MsgBox nID
    
            
End Function

Private Function Change_PW(ByVal Msg As Variant)

    Dim Temp As String
    
    Dim nTemp As String
    
    Dim nLen As Integer ' The length of the New Password will be stored here
    
    Dim oLen As Integer ' The length of the Old Password will be stored here
    
    Dim nPW As String * 9 ' The New Password will be stored here
    
    Dim oPW As String * 9 ' The Old Password for crossReferancing will be stored here
    
    Temp = Trim(Left(Msg, 5)) ' The Trim() Function is used to trim spaces
    
    nTemp = Int(Trim(Right(Temp, 2))) 'Start find old pw length
    
    'Check to see if the first char is a 0 if it is then its a # below 10
    If Left(nTemp, 1) = "0" Then
    
        'Now that we know it is store it
        oLen = Int(Right(nTemp, 1))
        
    'If it is not a 0 then its 10 or higher
    Else
    
        'So lets store it
        oLen = Int(nTemp)
        
    End If
    
    'Lets only store the small portion that we need
    nTemp = Trim(Left(Msg, 3)) 'again just for reasurance the Trim() function
    
    'Only need two chars
    Temp = Right(nTemp, 2) 'Start find new pw length
    
    'Again checking for a Zero
    If Left(Temp, 1) = "0" Then
    
        'k we found it now store it
        nLen = Int(Right(Temp, 1))
        
    'If it's not a 0 then its 10 or higher :)
    Else
    
        'Ok lets store it
        nLen = Int(Temp)
        
    End If
    
    'Now this is your old password being stored here " eHem old PassWord "
    oPW = Trim(Right(Msg, oLen)) 'Store old pw
    
    'Getting the portion we need
    Temp = Trim(Right(Msg, oLen + nLen))
    
    'Storing the new PassWord
    nPW = Trim(Left(Temp, nLen)) 'Store new pw
    
    'Try it out
    MsgBox nPW 'Test
    
    
End Function

Private Function New_Account(ByVal Msg As Variant)

    'Format
    ' 5000UserNameUserPassword
    
    Dim Temp As String
    
    Dim nTemp As String
    
    Dim uName As String
    
    Dim uPass As String
    
    Dim nLen As Integer
    
    Dim pLen As Integer
    
    'The password Can only be a maximum of 9 so no need to read out more than this
    Temp = Left(Msg, 4)
    
    pLen = Int(Right(Temp, 1))
    
    uPass = Right(Msg, pLen)
'_________________________________________
    
    'Again reading out need portion
    Temp = Left(Msg, 3)
    
    nTemp = Right(Temp, 2) 'Store #'s to check for 0
    
    If Left(nTemp, 1) = "0" Then ' Check if there is a 0
        
        nLen = Int(Right(nTemp, 1)) 'If so store the other #
    
    Else ' if no 0 then its 10 or higher
        
        nLen = Int(nTemp) ' So store it here
    
    End If
    
    Temp = Right(Msg, nLen + pLen) 'Read out need portion
    
    uName = Left(Temp, nLen) 'Store the Name
    
    Add_New frmMain.db 'Add a new section to the Database
    
    frmMain.tID.Text = uName 'Store the NewUsers ID
    
    frmMain.tPW.Text = uPass 'Store the new Password
    
    frmMain.tWarn.Text = "0" 'Set the warnings to 0 as they are a new user
    
    frmMain.tDate.Text = CStr(Date) 'Store the date for which they signed on with us
    
    Update_Db frmMain.db 'Now up date the information so that it is stored and viewable to us
    
   ' Refresh_Db frmMain.db
    
    'Show the info on the server side screen
    AddtoList frmMain.lv1, , "New User: " & uName
    
    
End Function
