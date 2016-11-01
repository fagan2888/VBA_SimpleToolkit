Attribute VB_Name = "CreateObject"
Public Function edhCreateDbHandler(dbConfig As dbConfig) As DbHandler

    Set edhCreateDbHandler = New DbHandler
    edhCreateDbHandler.initialize dbConfig

End Function

Public Function edhCreateExcelAsDataBaseHandler(path As String) As DbHandler
    
    Dim config As dbConfig
    Set config = New dbConfig
    
    config.initialize "Microsoft.ACE.OLEDB.12.0", "Data Source=" & path & ";" & _
            "Extended Properties=""Excel 8.0;HDR=Yes;"";"
 
    Set edhCreateExcelAsDataBaseHandler = edhCreateDbHandler(config)
  
End Function
