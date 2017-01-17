Var Linea                   : A80   = ""
Var Linea1                   : A20   = "Linea1"
Var Linea2                  : A20   = "Linea2"
Var Linea3                   : A20   = "Linea3"
Var Linea4                   : A20   = "Linea4"
Var Linea5                   : A20   = "Linea5"
Var Linea6                   : A20   = "Linea6"
Var Linea7                   : A20   = "Linea7"
Var SqlStr1                     : A3000
Var SqlStr                      : A3000 = ""
Var hODBCDLL            : N12   = 0

Event print_header : Header_1
	call Consulta_DomicilioFiscal(Linea1)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[1] = Linea
	call Consulta_DomicilioFiscal(Linea2)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[2] = Linea
	call Consulta_DomicilioFiscal(Linea3)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[3] = Linea
	call Consulta_DomicilioFiscal(Linea4)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[4] = Linea
	call Consulta_DomicilioFiscal(Linea5)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[5] = Linea
	call Consulta_DomicilioFiscal(Linea6)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[6] = Linea
	call Consulta_DomicilioFiscal(Linea7)
	Split SqlStr1, ";", SqlStr1
      Linea = SqlStr1
	@header[7] = Linea
	
ENDEVENT 
Sub Conecta_Base_Datos_Sybase
    Call Load_ODBC_DLL_SyBase
    Call ConnectDB_SyBase
EndSub
Sub ConnectDB_SyBase
    DLLCALL_CDECL hODBCDLL, sqlInitConnection("micros","ODBC;UID=dba;PWD=Password1", "")//BurguerKing
EndSub

Sub Load_ODBC_DLL_SyBase
  	hODBCDLL  = 0

	  If @WSTYPE = 3
	      DLLFree hODBCDLL
	      DLLLoad hODBCDLL, "\cf\micros\bin\MDSSysUtilsProxy.dll"
	   Else
	      DLLFree hODBCDLL
	      DLLLoad hODBCDLL, "MDSSysUtilsProxy.dll"
	   EndIf
    
	  IF hODBCDLL = 0 
	    ExitWithError "No se Puede Cargar DLL (MDSSysUtilsProxy.DLL)"
	  ENDIF
EndSub
Sub UnloadDLL_SyBase
  DLLCALL_CDECL hODBCDLL, sqlCloseConnection()
    
  DLLFree  hODBCDLL
  hODBCDLL = 0
EndSub
Sub Consulta_Sql

  Call Conecta_Base_Datos_Sybase
  
  DLLCALL_CDECL hODBCDLL, sqlGetRecordSet(SqlStr)
    SqlStr1 = ""
    DLLCALL_CDECL hODBCDLL, sqlGetLastErrorString(ref SqlStr1)
     
    IF (SqlStr1 <> "")
      
      Call UnloadDLL_SyBase
      ExitCancel
    ENDIF
     
    SqlStr1 = ""
    // si no obtiene error hace la consulta
    // regresa el valor en SqlStr1 
    DLLCALL_CDECL hODBCDLL, sqlGetFirst(ref SqlStr1)
EndSub

Sub Consulta_DomicilioFiscal( ref Etiqueta_Crm )
  Format SqlStr as "SELECT ",Etiqueta_Crm, " FROM custom.domicilios_fiscales "
  Call Consulta_Sql
  Call UnloadDLL_SyBase
EndSub