Imports Sap.Data.Hana
Imports SAPbobsCOM

Module SubMain

    Public SBOCompany As SAPbobsCOM.Company

    Sub Main()

        Conectar()
        cerrarOrdenesVentas()
        Desconectar()

    End Sub

    Public Function Conectar()

        Try

            SBOCompany = New SAPbobsCOM.Company

            SBOCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            SBOCompany.Server = My.Settings.Server
            SBOCompany.LicenseServer = My.Settings.LicenseServer
            SBOCompany.DbUserName = My.Settings.DbUserName
            SBOCompany.DbPassword = My.Settings.DbPassword

            SBOCompany.CompanyDB = My.Settings.CompanyDB

            SBOCompany.UserName = My.Settings.UserName
            SBOCompany.Password = My.Settings.Password

            SBOCompany.Connect()

        Catch ex As Exception

            MsgBox("Error al Conectar: " & ex.Message)

        End Try

    End Function


    Public Function cerrarOrdenesVentas()
        Dim stQueryH As String
        Dim oRecSetH As SAPbobsCOM.Recordset
        Dim docNum As String
        Dim docEntry As String
        Dim oOrder As SAPbobsCOM.Documents
        Dim RetVal As Long

        Try

            oOrder = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

            oRecSetH = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQueryH = "Call ""ACerrarOrdenesAntiguas"""
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oRecSetH.MoveFirst()

                For i = 0 To oRecSetH.RecordCount - 1

                    docEntry = oRecSetH.Fields.Item("DocEntry").Value

                    'Retrieve the document record to close from the database
                    RetVal = oOrder.GetByKey(docEntry)

                    'Close the record
                    RetVal = oOrder.Close()

                    'Limpia valor de RetVal
                    RetVal = Nothing

                    'Se mueve a la siguiente linea
                    oRecSetH.MoveNext()

                Next

            End If

        Catch ex As Exception
            MsgBox("cerrarOrdenesVentas: " & ex.Message)
            Return -1
        End Try
    End Function


    Public Function Desconectar()

        Try

            SBOCompany.Disconnect()

        Catch ex As Exception

            MsgBox("Error al tratar de cerrar conexión con SAP B1: " & ex.Message)

        End Try

    End Function


End Module
