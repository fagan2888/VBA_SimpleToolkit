Attribute VB_Name = "CreateObject"
Public Function edhCreateDbHandler(dbConfig As dbConfig) As DbHandler

    Set edhCreateDbHandler = New DbHandler
    edhCreateDbHandler.initialize dbConfig

End Function

Public Function edhCreateNiceDbHandler() As DbHandler
    Dim config As dbConfig
    Set config = New dbConfig
    config.initialize "OraOLEDB.Oracle.1", "User ID=nice;Password=nps1260;Data Source=PNI_TITAN"
    Set edhCreateNiceDbHandler = edhCreateDbHandler(config)
End Function

Public Function edhCreateExcelAsDataBaseHandler(path As String) As DbHandler
    
    Dim config As dbConfig
    Set config = New dbConfig
    
    config.initialize "Microsoft.ACE.OLEDB.12.0", "Data Source=" & path & ";" & _
            "Extended Properties=""Excel 8.0;HDR=Yes;"";"
 
    Set edhCreateExcelAsDataBaseHandler = edhCreateDbHandler(config)
  
End Function
