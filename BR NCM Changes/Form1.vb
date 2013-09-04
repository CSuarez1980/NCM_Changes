Public Class Form1
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BGWL7P.RunWorkerAsync()
        'BGWG4P.RunWorkerAsync()
        'BGWGBP.RunWorkerAsync()
        'BGWL6P.RunWorkerAsync()
    End Sub

    Private Sub BGWL7P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGWL7P.DoWork
        Const SAPBox As String = "L7P"
        Dim Rep As New SAPCOM.POs_Report(SAPBox, Environ("USERID"), "LAT")
        Dim cn As New OAConnection.Connection
        Dim LAPlants As New DataTable
        Dim MaxPO As New DataTable
        Dim PO As Double = 0
        LAPlants = cn.RunSentence("Select * From Plant Where Country = 'BR'").Tables(0)
        MaxPO = cn.RunSentence("Select Top 1 [Doc Number] From BR_NCM_Changes Where SAPBox = '" & SAPBox & "' Order By [Doc Number] Desc").Tables(0)

        PO = MaxPO.Rows(0)("Doc Number")
        PO += 1

        If PO > 1 Then
            If LAPlants.Rows.Count > 0 Then
                For Each R As DataRow In LAPlants.Rows
                    Rep.IncludePlant(R("Code"))
                Next

                Rep.IncludeDocumentFromTo(PO.ToString, "3999999999")
                'Rep.IncludeDocsDatedFromTo(Today, Today)

                Rep.Execute()
                If Rep.Success Then
                    Dim EKPO As New SAPCOM.EKPO_Report(SAPBox, Environ("USERID"), "LAT")
                   
                    For Each R As DataRow In Rep.Data.Rows
                        If (R("UOM") = "LE") Or (R("UOM") = "ACT") Or (R("Del Indicator") = "L") Or (R("Del Indicator") = "S") Then
                            R.Delete()
                        Else
                            EKPO.IncludeDocument(R("Doc Number"))
                        End If
                    Next

                    EKPO.AddCustomField("J_1BNBM", "NCM Code")
                    Rep.Data.AcceptChanges()

                    If Rep.Data.Rows.Count > 0 Then
                        EKPO.Execute()
                        If EKPO.Success Then
                            Dim SAP As New DataColumn
                            SAP.ColumnName = "SAPBox"
                            SAP.Caption = "SAPBox"
                            SAP.DefaultValue = SAPBox

                            Rep.Data.Columns.Add("SAP NCM")
                            Rep.Data.Columns.Add("New NCM")
                            Rep.Data.Columns.Add("Material Usage")
                            Rep.Data.Columns.Add("Material Origen")
                            Rep.Data.Columns.Add("Change")
                            Rep.Data.Columns.Add("Changed by DB")

                            Rep.Data.Columns.Add(SAP)

                            For Each R As DataRow In Rep.Data.Rows
                                Dim NCM = (From C In EKPO.Data.AsEnumerable() _
                                           Where ((C.Item("Doc Number") = R("Doc Number")) And (C.Item("Item Number") = R("Item Number"))) _
                                           Select C.Item("NCM Code"))

                                If Not NCM Is Nothing Then
                                    R("SAP NCM") = NCM(0)
                                End If

                                Dim T As New DataTable
                                T = cn.GetBRTable(R("Vendor"), R("Short Text"))
                                If Not T Is Nothing AndAlso T.Rows.Count > 0 Then
                                    R("Change") = "X"

                                    'Modificado porque el NCM
                                    Dim pNCM As String = T.Rows(0)("NCM Code").ToString.PadLeft(10, "0")
                                    IIf(pNCM = "0ISS_00000", pNCM = "ISS_00000", False) '

                                    R("New NCM") = pNCM
                                    R("Material Usage") = T.Rows(0)("Material Usage")
                                    R("Material Origen") = T.Rows(0)("Material Origen")
                                End If
                            Next
                        Else
                            Dim Attach() As String
                            ReDim Attach(1)

                            Attach(0) = ""
                            cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & EKPO.ErrMessage, "", False, "HTML", , True)
                        End If
                        Rep.Data.AcceptChanges()
                        cn.AppendTableToSqlServer("BR_NCM_Changes", Rep.Data)
                    End If
                Else
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = ""
                    cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & Rep.ErrMessage, "", False, "HTML", , True)

                End If
            Else
                MsgBox("No plants found.", MsgBoxStyle.Critical)
            End If
        End If

        'Esto es para intentar modificar alguna que no se pudo modificar anteriormente por caida de script:
        Dim DT As New DataTable
        Dim P As New DataTable

        DT = cn.RunSentence("Select * From BR_NCM_Changes Where (([Change] = 'X') And (SAPBox = '" & SAPBox & "'))").Tables(0)
        P = cn.RunSentence("Select * From BR_NCM_Changes_Pass Where SAPBox = '" & SAPBox & "'").Tables(0)


        If DT.Rows.Count > 0 Then
            Dim SAPCn As New SAPCOM.SAPConnector
            Dim CD As SAPCOM.ConnectionData

            CD.Box = SAPBox
            CD.Login = P.Rows(0)("TNumber")
            CD.Password = cn.Encrypt(P.Rows(0)("Pass"))

            'CD.Box = "L7A"
            'CD.Login = "BM4691"
            'CD.Password = "iker2012"

            Dim Conn As Object = SAPCn.GetSAPConnection(CD)
            Dim iSAP As New SAPConection.c_SAP(CD.Box)

            iSAP.UserName = CD.Login
            iSAP.Password = CD.Password
            iSAP.OpenConnection(False)

            Dim BRF As New SAPConection.BRF_Fixing(iSAP.GUI)
            For Each Row In DT.Rows

                Dim POChange As New SAPCOM.POChanges(Conn, Row("Doc Number"))
                'Dim POChange As New SAPCOM.POChanges(Conn, "3062266423")

                If POChange.IsReady Then
                    POChange.MaterialOrigin(Row("Item Number")) = Row("Material Origen").ToString.ToUpper.Trim
                    POChange.MaterialUsage(Row("Item Number")) = Row("Material Usage").ToString.ToUpper.Trim

                    'POChange.BrasNCMCode(Row("Item Number")) = Row("New NCM").ToString.ToUpper.Trim
                    'POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    If Not DBNull.Value.Equals(Row("New NCM")) Then
                        POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    Else
                        POChange.BrasNCMCode(Row("Item Number")) = ""
                    End If


                    POChange.CommitChanges()
                    If Not POChange.Success Then
                        Dim er As String
                        Dim EM As String = ""

                        For Each er In POChange.Results
                            EM = EM & Chr(13) & er
                        Next

                        Dim Attach() As String
                        ReDim Attach(1)

                        Attach(0) = ""
                        cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: PO" & Row("Doc Number") & "-" & Row("Item Number") & " - " & EM, "", False, "HTML", , True)
                    Else
                        If iSAP.Conected Then
                            BRF.Documents.Clear()
                            BRF.IncludePO(New SAPConection.BRF_PO(Row("Doc Number"), Row("Item Number")))
                            'BRF.IncludePO(New SAPConection.BRF_PO("3062266423", "10"))
                            BRF.Execute()
                            cn.ExecuteInServer("Update BR_NCM_Changes Set [Change] = '', [Change By DB] = 'X' Where (([Doc Number] = '" & Row("Doc Number") & "') And ([Item Number] = '" & Row("Item Number") & "') And (SAPBox = '" & SAPBox & "'))")
                        End If
                    End If
                Else
                    MsgBox("Error getting SAP Connection.", MsgBoxStyle.Exclamation)
                End If
                ' End If
            Next
            iSAP.CloseConnection()

            Dim xlPath As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\" & Replace(Replace(Now.ToString, "/", "-"), ":", "-") & "- SAPBox" & SAPBox & ".xlsx"
            If Not Rep.Data Is Nothing Then
                For Each r As DataRow In Rep.Data.Rows
                    If Not DBNull.Value.Equals(r("Change")) Then
                        If (r("Change") = "X") Then
                            r.Delete()
                        End If
                    End If
                Next

                Rep.Data.AcceptChanges()

                If cn.ExportDataTableToXL(Rep.Data, xlPath) Then
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = xlPath
                    cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "Auto NCM code and BRF+ report for " & SAPBox & " without changes done.", "", False, "HTML", , True)
                End If

            End If
        Else
            Dim Attach() As String
            ReDim Attach(1)

            Attach(0) = ""
            cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "No new catalogs found for changes.", "", False, "HTML", , True)

        End If
    End Sub
    Private Sub BGWG4P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGWG4P.DoWork
        Const SAPBox As String = "G4P"
        Dim Rep As New SAPCOM.POs_Report(SAPBox, Environ("USERID"), "LAT")
        Dim cn As New OAConnection.Connection
        Dim LAPlants As New DataTable
        Dim MaxPO As New DataTable
        Dim PO As Double = 0
        LAPlants = cn.RunSentence("Select * From Plant Where Country = 'BR'").Tables(0)
        MaxPO = cn.RunSentence("Select Top 1 [Doc Number] From BR_NCM_Changes Where SAPBox = '" & SAPBox & "' Order By [Doc Number] Desc").Tables(0)

        PO = MaxPO.Rows(0)("Doc Number")
        PO += 1

        If PO > 1 Then
            If LAPlants.Rows.Count > 0 Then
                For Each R As DataRow In LAPlants.Rows
                    Rep.IncludePlant(R("Code"))
                Next

                Rep.IncludeDocumentFromTo(PO.ToString, "3999999999")
                'Rep.IncludeDocsDatedFromTo(Today, Today)

                Rep.Execute()
                If Rep.Success Then
                    Dim EKPO As New SAPCOM.EKPO_Report(SAPBox, Environ("USERID"), "LAT")

                    For Each R As DataRow In Rep.Data.Rows
                        If (R("UOM") = "LE") Or (R("UOM") = "ACT") Or (R("Del Indicator") = "L") Or (R("Del Indicator") = "S") Then
                            R.Delete()
                        Else
                            EKPO.IncludeDocument(R("Doc Number"))
                        End If
                    Next

                    EKPO.AddCustomField("J_1BNBM", "NCM Code")
                    Rep.Data.AcceptChanges()

                    If Rep.Data.Rows.Count > 0 Then
                        EKPO.Execute()
                        If EKPO.Success Then
                            Dim SAP As New DataColumn
                            SAP.ColumnName = "SAPBox"
                            SAP.Caption = "SAPBox"
                            SAP.DefaultValue = SAPBox

                            Rep.Data.Columns.Add("SAP NCM")
                            Rep.Data.Columns.Add("New NCM")
                            Rep.Data.Columns.Add("Material Usage")
                            Rep.Data.Columns.Add("Material Origen")
                            Rep.Data.Columns.Add("Change")
                            Rep.Data.Columns.Add("Changed by DB")

                            Rep.Data.Columns.Add(SAP)

                            For Each R As DataRow In Rep.Data.Rows
                                Dim NCM = (From C In EKPO.Data.AsEnumerable() _
                                           Where ((C.Item("Doc Number") = R("Doc Number")) And (C.Item("Item Number") = R("Item Number"))) _
                                           Select C.Item("NCM Code"))

                                If Not NCM Is Nothing Then
                                    R("SAP NCM") = NCM(0)
                                End If

                                Dim T As New DataTable
                                T = cn.GetBRTable(R("Vendor"), R("Short Text"))
                                If Not T Is Nothing AndAlso T.Rows.Count > 0 Then
                                    R("Change") = "X"

                                    'Modificado porque el NCM
                                    Dim pNCM As String = T.Rows(0)("NCM Code").ToString.PadLeft(10, "0")
                                    IIf(pNCM = "0ISS_00000", pNCM = "ISS_00000", False)

                                    R("New NCM") = pNCM
                                    R("Material Usage") = T.Rows(0)("Material Usage")
                                    R("Material Origen") = T.Rows(0)("Material Origen")
                                End If
                            Next
                        Else
                            Dim Attach() As String
                            ReDim Attach(1)

                            Attach(0) = ""
                            cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & EKPO.ErrMessage, "", False, "HTML", , True)
                        End If

                        Rep.Data.AcceptChanges()
                        cn.AppendTableToSqlServer("BR_NCM_Changes", Rep.Data)
                    End If
                Else
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = ""
                    cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & Rep.ErrMessage, "", False, "HTML", , True)

                End If
            Else
                MsgBox("No plants found.", MsgBoxStyle.Critical)
            End If
        End If

        'Esto es para intentar modificar alguna que no se pudo modificar anteriormente por caida de script:
        Dim DT As New DataTable
        Dim P As New DataTable

        DT = cn.RunSentence("Select * From BR_NCM_Changes Where (([Change] = 'X') And (SAPBox = '" & SAPBox & "'))").Tables(0)
        P = cn.RunSentence("Select * From BR_NCM_Changes_Pass Where SAPBox = '" & SAPBox & "'").Tables(0)


        If DT.Rows.Count > 0 Then
            Dim SAPCn As New SAPCOM.SAPConnector
            Dim CD As SAPCOM.ConnectionData

            CD.Box = SAPBox
            CD.Login = P.Rows(0)("TNumber")
            CD.Password = cn.Encrypt(P.Rows(0)("Pass"))

            'CD.Box = "L7A"
            'CD.Login = "BM4691"
            'CD.Password = "iker2012"

            Dim Conn As Object = SAPCn.GetSAPConnection(CD)
            Dim iSAP As New SAPConection.c_SAP(CD.Box)

            iSAP.UserName = CD.Login
            iSAP.Password = CD.Password
            iSAP.OpenConnection(False)

            Dim BRF As New SAPConection.BRF_Fixing(iSAP.GUI)

            For Each Row In DT.Rows

                Dim POChange As New SAPCOM.POChanges(Conn, Row("Doc Number"))
                'Dim POChange As New SAPCOM.POChanges(Conn, "3062266423")

                If POChange.IsReady Then
                    POChange.MaterialOrigin(Row("Item Number")) = Row("Material Origen").ToString.ToUpper.Trim
                    POChange.MaterialUsage(Row("Item Number")) = Row("Material Usage").ToString.ToUpper.Trim

                    'POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    If Not DBNull.Value.Equals(Row("New NCM")) Then
                        POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    Else
                        POChange.BrasNCMCode(Row("Item Number")) = ""
                    End If
                    'POChange.BrasNCMCode(Row("Item Number")) = Row("New NCM").ToString.ToUpper.Trim

                    POChange.CommitChanges()
                    If Not POChange.Success Then
                        Dim er As String
                        Dim EM As String = ""

                        For Each er In POChange.Results
                            EM = EM & Chr(13) & er
                        Next

                        Dim Attach() As String
                        ReDim Attach(1)

                        Attach(0) = ""
                        cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: PO" & Row("Doc Number") & "-" & Row("Item Number") & " - " & EM, "", False, "HTML", , True)
                    Else
                        If iSAP.Conected Then
                            BRF.Documents.Clear()
                            BRF.IncludePO(New SAPConection.BRF_PO(Row("Doc Number"), Row("Item Number")))
                            'BRF.IncludePO(New SAPConection.BRF_PO("3062266423", "10"))
                            BRF.Execute()
                            cn.ExecuteInServer("Update BR_NCM_Changes Set [Change] = '', [Change By DB] = 'X' Where (([Doc Number] = '" & Row("Doc Number") & "') And ([Item Number] = '" & Row("Item Number") & "') And (SAPBox = '" & SAPBox & "'))")
                        End If
                    End If
                Else
                    MsgBox("Error getting SAP Connection.", MsgBoxStyle.Exclamation)
                End If
                ' End If
            Next
            iSAP.CloseConnection()

            Dim xlPath As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\" & Replace(Replace(Now.ToString, "/", "-"), ":", "-") & "- SAPBox" & SAPBox & ".xlsx"
            If Not Rep.Data Is Nothing Then
                For Each r As DataRow In Rep.Data.Rows
                    If Not DBNull.Value.Equals(r("Change")) Then
                        If (r("Change") = "X") Then
                            r.Delete()
                        End If
                    End If
                Next

                Rep.Data.AcceptChanges()

                If cn.ExportDataTableToXL(Rep.Data, xlPath) Then
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = xlPath
                    cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "Auto NCM code and BRF+ report for " & SAPBox & " without changes done.", "", False, "HTML", , True)
                End If
            End If
        Else
            Dim Attach() As String
            ReDim Attach(1)

            Attach(0) = ""
            cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "No new catalogs found for changes.", "", False, "HTML", , True)

        End If
    End Sub
    Private Sub BGWL6P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGWL6P.DoWork
        Const SAPBox As String = "L6P"
        Dim Rep As New SAPCOM.POs_Report(SAPBox, Environ("USERID"), "LAT")
        Dim cn As New OAConnection.Connection
        Dim LAPlants As New DataTable
        Dim MaxPO As New DataTable
        Dim PO As Double = 0
        LAPlants = cn.RunSentence("Select * From Plant Where Country = 'BR'").Tables(0)

        MaxPO = cn.RunSentence("Select Top 1 [Doc Number] From BR_NCM_Changes Where SAPBox = '" & SAPBox & "' Order By [Doc Number] Desc").Tables(0)

        PO = MaxPO.Rows(0)("Doc Number")
        PO += 1

        If PO > 1 Then
            If LAPlants.Rows.Count > 0 Then
                For Each R As DataRow In LAPlants.Rows
                    Rep.IncludePlant(R("Code"))
                Next

                Rep.IncludeDocumentFromTo(PO.ToString, "3999999999")
                'Rep.IncludeDocsDatedFromTo(Today, Today)

                Rep.Execute()
                If Rep.Success Then
                    Dim EKPO As New SAPCOM.EKPO_Report(SAPBox, Environ("USERID"), "LAT")

                    For Each R As DataRow In Rep.Data.Rows
                        If (R("UOM") = "LE") Or (R("UOM") = "ACT") Or (R("Del Indicator") = "L") Or (R("Del Indicator") = "S") Then
                            R.Delete()
                        Else
                            EKPO.IncludeDocument(R("Doc Number"))
                        End If
                    Next

                    EKPO.AddCustomField("J_1BNBM", "NCM Code")
                    Rep.Data.AcceptChanges()

                    If Rep.Data.Rows.Count > 0 Then
                        EKPO.Execute()
                        If EKPO.Success Then
                            Dim SAP As New DataColumn
                            SAP.ColumnName = "SAPBox"
                            SAP.Caption = "SAPBox"
                            SAP.DefaultValue = SAPBox

                            Rep.Data.Columns.Add("SAP NCM")
                            Rep.Data.Columns.Add("New NCM")
                            Rep.Data.Columns.Add("Material Usage")
                            Rep.Data.Columns.Add("Material Origen")
                            Rep.Data.Columns.Add("Change")
                            Rep.Data.Columns.Add("Changed by DB")

                            Rep.Data.Columns.Add(SAP)

                            For Each R As DataRow In Rep.Data.Rows
                                Dim NCM = (From C In EKPO.Data.AsEnumerable() _
                                           Where ((C.Item("Doc Number") = R("Doc Number")) And (C.Item("Item Number") = R("Item Number"))) _
                                           Select C.Item("NCM Code"))

                                If Not NCM Is Nothing Then
                                    R("SAP NCM") = NCM(0)
                                End If

                                Dim T As New DataTable
                                T = cn.GetBRTable(R("Vendor"), R("Short Text"))
                                If Not T Is Nothing AndAlso T.Rows.Count > 0 Then
                                    R("Change") = "X"

                                    'Modificado porque el NCM
                                    Dim pNCM As String = T.Rows(0)("NCM Code").ToString.PadLeft(10, "0")
                                    IIf(pNCM = "0ISS_00000", pNCM = "ISS_00000", False) '

                                    R("Material Usage") = T.Rows(0)("Material Usage")
                                    R("Material Origen") = T.Rows(0)("Material Origen")
                                End If
                            Next
                        Else
                            Dim Attach() As String
                            ReDim Attach(1)

                            Attach(0) = ""
                            cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & EKPO.ErrMessage, "", False, "HTML", , True)
                        End If

                        Rep.Data.AcceptChanges()
                        cn.AppendTableToSqlServer("BR_NCM_Changes", Rep.Data)
                    End If
                Else
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = ""
                    cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & Rep.ErrMessage, "", False, "HTML", , True)
                End If
            Else
                MsgBox("No plants found.", MsgBoxStyle.Critical)
            End If
        End If

        'Esto es para intentar modificar alguna que no se pudo modificar anteriormente por caida de script:
        Dim DT As New DataTable
        Dim P As New DataTable

        DT = cn.RunSentence("Select * From BR_NCM_Changes Where (([Change] = 'X') And (SAPBox = '" & SAPBox & "'))").Tables(0)
        P = cn.RunSentence("Select * From BR_NCM_Changes_Pass Where SAPBox = '" & SAPBox & "'").Tables(0)


        If DT.Rows.Count > 0 Then
            Dim SAPCn As New SAPCOM.SAPConnector
            Dim CD As SAPCOM.ConnectionData

            CD.Box = SAPBox
            CD.Login = P.Rows(0)("TNumber")
            CD.Password = cn.Encrypt(P.Rows(0)("Pass"))

            'CD.Box = "L7A"
            'CD.Login = "BM4691"
            'CD.Password = "iker2012"

            Dim Conn As Object = SAPCn.GetSAPConnection(CD)
            Dim iSAP As New SAPConection.c_SAP(CD.Box)

            iSAP.UserName = CD.Login
            iSAP.Password = CD.Password
            iSAP.OpenConnection(False)

            Dim BRF As New SAPConection.BRF_Fixing(iSAP.GUI)

            For Each Row In DT.Rows

                Dim POChange As New SAPCOM.POChanges(Conn, Row("Doc Number"))
                'Dim POChange As New SAPCOM.POChanges(Conn, "3062266423")

                If POChange.IsReady Then
                    POChange.MaterialOrigin(Row("Item Number")) = Row("Material Origen").ToString.ToUpper.Trim
                    POChange.MaterialUsage(Row("Item Number")) = Row("Material Usage").ToString.ToUpper.Trim

                    'POChange.BrasNCMCode(Row("Item Number")) = Row("New NCM").ToString.ToUpper.Trim

                    If Not DBNull.Value.Equals(Row("New NCM")) Then
                        POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    Else
                        POChange.BrasNCMCode(Row("Item Number")) = ""
                    End If



                    POChange.CommitChanges()
                    If Not POChange.Success Then
                        Dim er As String
                        Dim EM As String = ""

                        For Each er In POChange.Results
                            EM = EM & Chr(13) & er
                        Next

                        Dim Attach() As String
                        ReDim Attach(1)

                        Attach(0) = ""
                        cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: PO" & Row("Doc Number") & "-" & Row("Item Number") & " - " & EM, "", False, "HTML", , True)
                    Else
                        If iSAP.Conected Then
                            BRF.Documents.Clear()
                            BRF.IncludePO(New SAPConection.BRF_PO(Row("Doc Number"), Row("Item Number")))
                            'BRF.IncludePO(New SAPConection.BRF_PO("3062266423", "10"))
                            BRF.Execute()
                            cn.ExecuteInServer("Update BR_NCM_Changes Set [Change] = '', [Change By DB] = 'X' Where (([Doc Number] = '" & Row("Doc Number") & "') And ([Item Number] = '" & Row("Item Number") & "') And (SAPBox = '" & SAPBox & "'))")
                        End If
                    End If
                Else
                    MsgBox("Error getting SAP Connection.", MsgBoxStyle.Exclamation)
                End If
                ' End If
            Next
            iSAP.CloseConnection()

            Dim xlPath As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\" & Replace(Replace(Now.ToString, "/", "-"), ":", "-") & "- SAPBox" & SAPBox & ".xlsx"
            If Not Rep.Data Is Nothing Then
                For Each r As DataRow In Rep.Data.Rows
                    If Not DBNull.Value.Equals(r("Change")) Then
                        If (r("Change") = "X") Then
                            r.Delete()
                        End If
                    End If
                Next

                Rep.Data.AcceptChanges()

                If cn.ExportDataTableToXL(Rep.Data, xlPath) Then
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = xlPath
                    cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "Auto NCM code and BRF+ report for " & SAPBox & " without changes done.", "", False, "HTML", , True)
                End If
            End If
        Else
            Dim Attach() As String
            ReDim Attach(1)

            Attach(0) = ""
            cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "No new catalogs found for changes.", "", False, "HTML", , True)

        End If
    End Sub
    Private Sub BGWGBP_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGWGBP.DoWork
        Const SAPBox As String = "GBP"
        Dim Rep As New SAPCOM.POs_Report(SAPBox, Environ("USERID"), "LAT")
        Dim cn As New OAConnection.Connection
        Dim LAPlants As New DataTable
        Dim MaxPO As New DataTable
        Dim PO As Double = 0
        LAPlants = cn.RunSentence("Select * From Plant Where Country = 'BR'").Tables(0)
        MaxPO = cn.RunSentence("Select Top 1 [Doc Number] From BR_NCM_Changes Where SAPBox = '" & SAPBox & "' Order By [Doc Number] Desc").Tables(0)

        PO = MaxPO.Rows(0)("Doc Number")
        PO += 1

        If PO > 1 Then
            If LAPlants.Rows.Count > 0 Then
                For Each R As DataRow In LAPlants.Rows
                    Rep.IncludePlant(R("Code"))
                Next

                Rep.IncludeDocumentFromTo(PO.ToString, "3999999999")
                'Rep.IncludeDocsDatedFromTo(Today, Today)

                Rep.Execute()
                If Rep.Success Then
                    Dim EKPO As New SAPCOM.EKPO_Report(SAPBox, Environ("USERID"), "LAT")

                    For Each R As DataRow In Rep.Data.Rows
                        If (R("UOM") = "LE") Or (R("UOM") = "ACT") Or (R("Del Indicator") = "L") Or (R("Del Indicator") = "S") Then
                            R.Delete()
                        Else
                            EKPO.IncludeDocument(R("Doc Number"))
                        End If
                    Next

                    EKPO.AddCustomField("J_1BNBM", "NCM Code")
                    Rep.Data.AcceptChanges()

                    If Rep.Data.Rows.Count > 0 Then
                        EKPO.Execute()
                        If EKPO.Success Then
                            Dim SAP As New DataColumn
                            SAP.ColumnName = "SAPBox"
                            SAP.Caption = "SAPBox"
                            SAP.DefaultValue = SAPBox

                            Rep.Data.Columns.Add("SAP NCM")
                            Rep.Data.Columns.Add("New NCM")
                            Rep.Data.Columns.Add("Material Usage")
                            Rep.Data.Columns.Add("Material Origen")
                            Rep.Data.Columns.Add("Change")
                            Rep.Data.Columns.Add("Changed by DB")

                            Rep.Data.Columns.Add(SAP)

                            For Each R As DataRow In Rep.Data.Rows
                                Dim NCM = (From C In EKPO.Data.AsEnumerable() _
                                           Where ((C.Item("Doc Number") = R("Doc Number")) And (C.Item("Item Number") = R("Item Number"))) _
                                           Select C.Item("NCM Code"))

                                If Not NCM Is Nothing Then
                                    R("SAP NCM") = NCM(0)
                                End If

                                Dim T As New DataTable
                                T = cn.GetBRTable(R("Vendor"), R("Short Text"))
                                If Not T Is Nothing AndAlso T.Rows.Count > 0 Then
                                    R("Change") = "X"

                                    'Modificado porque el NCM
                                    Dim pNCM As String = T.Rows(0)("NCM Code").ToString.PadLeft(10, "0")
                                    IIf(pNCM = "0ISS_00000", pNCM = "ISS_00000", False) '

                                    R("New NCM") = pNCM
                                    R("Material Usage") = T.Rows(0)("Material Usage")
                                    R("Material Origen") = T.Rows(0)("Material Origen")
                                End If
                            Next
                        Else
                            Dim Attach() As String
                            ReDim Attach(1)

                            Attach(0) = ""
                            cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & EKPO.ErrMessage, "", False, "HTML", , True)
                        End If

                        Rep.Data.AcceptChanges()
                        cn.AppendTableToSqlServer("BR_NCM_Changes", Rep.Data)
                    End If
                Else
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = ""
                    cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: " & Rep.ErrMessage, "", False, "HTML", , True)

                End If
            Else
                MsgBox("No plants found.", MsgBoxStyle.Critical)
            End If
        End If

        'Esto es para intentar modificar alguna que no se pudo modificar anteriormente por caida de script:
        Dim DT As New DataTable
        Dim P As New DataTable

        DT = cn.RunSentence("Select * From BR_NCM_Changes Where (([Change] = 'X') And (SAPBox = '" & SAPBox & "'))").Tables(0)
        P = cn.RunSentence("Select * From BR_NCM_Changes_Pass Where SAPBox = '" & SAPBox & "'").Tables(0)


        If DT.Rows.Count > 0 Then
            Dim SAPCn As New SAPCOM.SAPConnector
            Dim CD As SAPCOM.ConnectionData

            CD.Box = SAPBox
            CD.Login = P.Rows(0)("TNumber")
            CD.Password = cn.Encrypt(P.Rows(0)("Pass"))

            'CD.Box = "L7A"
            'CD.Login = "BM4691"
            'CD.Password = "iker2012"

            Dim Conn As Object = SAPCn.GetSAPConnection(CD)
            Dim iSAP As New SAPConection.c_SAP(CD.Box)

            iSAP.UserName = CD.Login
            iSAP.Password = CD.Password
            iSAP.OpenConnection(False)

            Dim BRF As New SAPConection.BRF_Fixing(iSAP.GUI)

            For Each Row In DT.Rows

                Dim POChange As New SAPCOM.POChanges(Conn, Row("Doc Number"))
                'Dim POChange As New SAPCOM.POChanges(Conn, "3062266423")

                If POChange.IsReady Then
                    POChange.MaterialOrigin(Row("Item Number")) = Row("Material Origen").ToString.ToUpper.Trim
                    POChange.MaterialUsage(Row("Item Number")) = Row("Material Usage").ToString.ToUpper.Trim

                    'POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    If Not DBNull.Value.Equals(Row("New NCM")) Then
                        POChange.BrasNCMCode(Row("Item Number")) = IIf(Row("New NCM") = "0ISS_00000", "ISS_00000", Row("New NCM"))
                    Else
                        POChange.BrasNCMCode(Row("Item Number")) = ""
                    End If
                    'POChange.BrasNCMCode(Row("Item Number")) = Row("New NCM").ToString.ToUpper.Trim

                    POChange.CommitChanges()
                    If Not POChange.Success Then
                        Dim er As String
                        Dim EM As String = ""

                        For Each er In POChange.Results
                            EM = EM & Chr(13) & er
                        Next

                        Dim Attach() As String
                        ReDim Attach(1)

                        Attach(0) = ""
                        cn.SendOutlookMail("Error AUTO NCM Changes: " & SAPBox, Attach, Environ("USERID") & "@PG.com", "", "Error: PO" & Row("Doc Number") & "-" & Row("Item Number") & " - " & EM, "", False, "HTML", , True)
                    Else
                        If iSAP.Conected Then
                            BRF.Documents.Clear()
                            BRF.IncludePO(New SAPConection.BRF_PO(Row("Doc Number"), Row("Item Number")))
                            'BRF.IncludePO(New SAPConection.BRF_PO("3062266423", "10"))
                            BRF.Execute()
                            cn.ExecuteInServer("Update BR_NCM_Changes Set [Change] = '', [Change By DB] = 'X' Where (([Doc Number] = '" & Row("Doc Number") & "') And ([Item Number] = '" & Row("Item Number") & "') And (SAPBox = '" & SAPBox & "'))")
                        End If
                    End If
                Else
                    MsgBox("Error getting SAP Connection.", MsgBoxStyle.Exclamation)
                End If
                ' End If
            Next
            iSAP.CloseConnection()

            Dim xlPath As String = My.Computer.FileSystem.SpecialDirectories.CurrentUserApplicationData & "\" & Replace(Replace(Now.ToString, "/", "-"), ":", "-") & "- SAPBox" & SAPBox & ".xlsx"
            If Not Rep.Data Is Nothing Then


                For Each r As DataRow In Rep.Data.Rows
                    If Not DBNull.Value.Equals(r("Change")) Then
                        If (r("Change") = "X") Then
                            r.Delete()
                        End If
                    End If
                Next

                Rep.Data.AcceptChanges()

                If cn.ExportDataTableToXL(Rep.Data, xlPath) Then
                    Dim Attach() As String
                    ReDim Attach(1)

                    Attach(0) = xlPath
                    cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "Auto NCM code and BRF+ report for " & SAPBox & " without changes done.", "", False, "HTML", , True)
                End If
            End If
        Else
            Dim Attach() As String
            ReDim Attach(1)

            Attach(0) = ""
            cn.SendOutlookMail("AUTO NCM Changes: " & SAPBox, Attach, P.Rows(0)("TNumber") & "@PG.com", "", "No new catalogs found for changes.", "", False, "HTML", , True)

        End If
    End Sub

    Private Sub The_End()
        If Not BGWL7P.IsBusy AndAlso Not BGWG4P.IsBusy AndAlso Not BGWGBP.IsBusy AndAlso Not BGWL6P.IsBusy Then
            End
        End If
    End Sub

    Private Sub BGWG4P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWG4P.RunWorkerCompleted
        BGWGBP.RunWorkerAsync()
    End Sub
    Private Sub BGWGBP_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWGBP.RunWorkerCompleted
        The_End()
    End Sub
    Private Sub BGWL6P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWL6P.RunWorkerCompleted
        BGWG4P.RunWorkerAsync()
    End Sub
    Private Sub BGWL7P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWL7P.RunWorkerCompleted
        BGWL6P.RunWorkerAsync()
    End Sub
End Class
