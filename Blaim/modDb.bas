Attribute VB_Name = "modDb"
                                        Option Explicit

    'Functions for controlling the DataBase storing and movement Routines
Public Function Add_New(Data1 As Data)

    'This will open a new section of the database to store more information
    Data1.Recordset.AddNew

End Function

Public Function Pro_Delete(Data1 As Data)

    'This will delete the entire Clints profile stored on the database
    Data1.Recordset.Delete

    Data1.Recordset.MoveNext
End Function

Public Function Refresh_Db(Data1 As Data)
    
    'This will refresh the database in case there are multiple connections we want only the
    ' most recent data
    Data1.Refresh

End Function

Public Function Update_Db(Data1 As Data)
    
    'Update a record that has been changed/new or modified
    Data1.UpdateRecord
    
    'Not needed nor used here just for you to play with
    Data1.Recordset.Bookmark = Data1.Recordset.LastModified 'will show last entry you were in

End Function
