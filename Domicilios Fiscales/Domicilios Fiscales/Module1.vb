Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text
Imports System.IO

Module Module1
    Dim Connection

    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook
    Dim xlWorkSheet As Excel.Worksheet
    Dim Existe
    Dim datoColum As String
    Dim numeroColum As Double
    Dim numeroTienda As String
    Dim numeroTiendaEn As Integer
    Dim datoColumEn As Integer
    Dim DomicilioFiscalNuevo As String
    Dim Linea1 As String
    Dim Linea2 As String
    Dim Linea3 As String
    Dim Linea4 As String
    Dim Linea5 As String
    Dim Linea6 As String
    Dim Linea7 As String
    Dim Tienda As String

    Dim myStreamWriter As StreamWriter
    Dim myStreamReader As StreamReader
    Public Const FileLog As String = "Log_Domicilio_fiscal.TxT"
    Public Registro_Log As String



    Sub Main()
        Try

            'Agregar datos a las celdas de la primera hoja en el libro nuevo
            Console.WriteLine("Selecione la tienda en el cual se ejecutara el programa")
            Console.WriteLine("Escribe 1 si es StarBucks o 2 si es BurgerKing")
            Tienda = Console.ReadLine()
            If Tienda = 1 Or Tienda = 2 Then
                Call StartConnection()
                Console.WriteLine("Creando la tabla Domicilio Fiscal")
                graba_log("Creando la tabla Domicilio Fiscal")
                Dim iReturn
                TablesExist()

                If Existe = False Then
                    Call DropTablesCatErro()
                    Call CreateTablesCustomDomicilios()
                Else

                    Call CreateTablesCustomDomicilios()

                End If
                graba_log("Se creo la tabla Domicilio Fiscal")
                Console.WriteLine("Se creo la tabla Domicilio Fiscal")
                Console.WriteLine("Buscando datos de la Tienda")
                graba_log("Buscando datos de la Tienda")
                Call BuscaTablaMicros()
                Console.WriteLine("Se obtuvo los datos de la tienda")
                Console.WriteLine("Buscando Datos en el archivo excel")
                Call LeerExcel()
                Console.WriteLine("Se encontro los datos en el archivo")
                Call BorrarDatosFiscales()
                Console.WriteLine("Borrando vieja dirección fiscal")
                InsertDomicilioFiscal()
                Console.WriteLine("Se inserto la nueva dirección fiscal")
                graba_log("Se Completo la actualización del Domicilio Fiscal")

            Else
                Console.WriteLine("Opción invalida")
            End If
        Catch ex As Exception
            graba_log(ex.ToString())

        End Try

       
    End Sub
    Sub DropTablesCatErro()
        Dim RS
        RS = Connection.Execute("drop table custom.domicilios_fiscales")
    End Sub
    Function TablesExist()
        Dim RS
        Dim strResponse

        RS = Connection.Execute("select * from sysobjects where name = 'custom.domicilios_fiscales'")


        graba_log("Existe>" + RS.Eof.ToString())
        Existe = RS.Eof
    End Function
    Sub StartConnection()
        Try
            Connection = CreateObject("ADODB.Connection")
            If Tienda = 1 Then
                Connection.Open("DSN=Micros;UID=dba;PWD=Db@M@5t3r$;Mode=Read")
                'Connection.Open("DSN=Micros;UID=dba;PWD=micros220965;Mode=Read")

            Else
                Connection.Open("DSN=Micros;UID=dba;PWD=Password1;Mode=Read")
            End If
        Catch ex As Exception
            Console.WriteLine("Tienda no seleccionada correctamente")
            graba_log("Tienda no seleccionada correctamente")
        End Try
       
    End Sub
    Sub CreateTablesCustomDomicilios()
        Try
            Dim RS
            Dim sqlCmd
            sqlCmd = "CREATE TABLE custom.domicilios_fiscales ("
            sqlCmd = sqlCmd & " Id	                integer		NOT NULL DEFAULT autoincrement,"
            sqlCmd = sqlCmd & " NumeroTienda	  	varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea1			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea2			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea3			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea4			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea5			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea6			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " Linea7			varchar(25)	NOT NULL,"
            sqlCmd = sqlCmd & " PRIMARY KEY (Id) )"
            RS = Connection.Execute(sqlCmd)
        Catch ex As Exception
            'Console.WriteLine("Tienda no seleccionada correctamente")
            'graba_log("Tienda no seleccionada correctamente")
        End Try
        
    End Sub
    Sub BuscaTablaMicros()
        Dim RS
        Dim sqlCmd
        Try
            sqlCmd = "SELECT obj_num FROM micros.rest_def"
            RS = Connection.Execute(sqlCmd)
            numeroTienda = RS.GetString
            graba_log("Numero de tienda>> " + numeroTienda)

        Catch ex As Exception
            graba_log("BuscaTabla error>> " + ex.ToString())
        End Try
       
      

    End Sub
    Sub InsertDomicilioFiscal()
        Try
            Dim RS
            Dim sqlCmd
            sqlCmd = "INSERT INTO custom.domicilios_fiscales (NumeroTienda, Linea1, Linea2, Linea3, Linea4, Linea5, Linea6, Linea7) VALUES (" + numeroTienda + ", '" + Linea1 + "', '" + Linea2 + "', '" + Linea3 + "', '" + Linea4 + "', '" + Linea5 + "', '" + Linea6 + "', '" + Linea7 + "')"
            RS = Connection.Execute(sqlCmd)
            graba_log("Se insertaron los datos en la tabla domicilios_fiscales ")

        Catch ex As Exception
            graba_log("Error Insertar Domicilio>>> " + ex.ToString())
        End Try
      
    End Sub
    Sub BorrarDatosFiscales()
        Try
            Dim RS
            Dim sqlCmd
            If Tienda = 1 Then
                sqlCmd = "Update micros.hdr_def set line_01 = ' ', line_02 = '@@Header_1', line_03 = ' ', line_04 = ' ', line_05 = ' ', line_06 = ' ' Where obj_num = 101"
                RS = Connection.Execute(sqlCmd)

            Else
                sqlCmd = "Update micros.hdr_def set line_01 = ' ', line_02 = '@@Header_1', line_03 = ' ', line_04 = ' ', line_05 = ' ', line_06 = ' ' Where obj_num = 100"
                RS = Connection.Execute(sqlCmd)
            End If

           
            graba_log("Se Borraron los datos de los headers de la tabla hdr_def")
        Catch ex As Exception
            graba_log("Error Borrar datos Fiscales>>> " + ex.ToString())
        End Try
       
    End Sub
    Sub LeerExcel()

        Try
            xlApp = New Excel.ApplicationClass
            ' xlWorkBook = xlApp.Workbooks.Open("C:\Documents and Settings\Administrador\Escritorio\Debug\DomiFiscal.xlsx")
            'xlWorkBook = xlApp.Workbooks.Open("C:\Users\Administrator\Desktop\Debug\DomiFiscal.xlsx")
            xlWorkBook = xlApp.Workbooks.Open("C:\Program Files\DomiciliosFiscales\DomiFiscal.xlsx")

            xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)
            numeroTiendaEn = CInt(numeroTienda)
            graba_log("Numero de tienda: " + numeroTiendaEn.ToString())
            DomicilioFiscalNuevo = "N"
            For index As Integer = 2 To 283
                datoColum = xlWorkSheet.Cells(index, 1).value
                datoColumEn = CInt(datoColum)
                'graba_log("Numero de tiendas en el archivo: " + datoColumEn.ToString())
                If numeroTiendaEn = datoColumEn Then
                    'datoColum = xlWorkSheet.Cells(1, index).value
                    DomicilioFiscalNuevo = "S"
                    Linea1 = xlWorkSheet.Cells(index, 2).value
                    graba_log("Domicilio encontrado 1: " + Linea1)
                    Linea2 = xlWorkSheet.Cells(index, 3).value
                    graba_log("Domicilio encontrado 2: " + Linea2)
                    Linea3 = xlWorkSheet.Cells(index, 4).value
                    graba_log("Domicilio encontrado 3:  " + Linea3)
                    Linea4 = xlWorkSheet.Cells(index, 5).value
                    graba_log("Domicilio encontrado 4: " + Linea4)
                    Linea5 = xlWorkSheet.Cells(index, 6).value
                    graba_log("Domicilio encontrado 5: " + Linea5)
                    Linea6 = xlWorkSheet.Cells(index, 7).value
                    graba_log("Domicilio encontrado 6: " + Linea6)
                    Linea7 = xlWorkSheet.Cells(index, 8).value
                    graba_log("Domicilio encontrado 7: " + Linea7)

                    Console.WriteLine("Domicilio encontrado " + Linea1 + " " + Linea2 + " " + Linea3 + " " + Linea4 + " " + Linea5 + " " + Linea6 + " " + Linea7)
                    graba_log("Domicilio encontrado: " + Linea1 + " " + Linea2 + " " + Linea3 + " " + Linea4 + " " + Linea5 + " " + Linea6 + " " + Linea7)
                End If


            Next

            If DomicilioFiscalNuevo.Equals("N") Then
                graba_log("No se encontro la tienda dentro del archivo")
                xlWorkBook.Close()
                xlApp.Quit()
                Environment.Exit(0)
            End If

            xlWorkBook.Close()
            xlApp.Quit()
        Catch ex As Exception
            graba_log("BuscaEnArchivo>> " + ex.ToString())
        End Try

    End Sub

    Private Sub graba_log(ByVal Reg_Log As String)
        Try
            myStreamWriter = File.AppendText(FileLog)
            myStreamWriter.WriteLine(Now.ToLongDateString & " " & Now.ToLongTimeString & ">>" & Reg_Log)
            myStreamWriter.Flush()
        Catch Errores_graba_log As Exception
            Registro_Log = "Errores_graba_log               :" + Errores_graba_log.ToString
        Finally
            If Not myStreamWriter Is Nothing Then
                myStreamWriter.Close()
            End If
        End Try
    End Sub
End Module
