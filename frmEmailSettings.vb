Imports SpilCommon
Imports System.Text.RegularExpressions

Public Class frmEmailSettings
    Dim oSQL As New clsSqlConn
    Dim dsValues As DataSet = Nothing
    Dim oSQLQuery As String = ""
    Dim messageValue As Integer = 0
    Dim isSetDefAccount As Integer = 0
    Dim isManully As Boolean = False

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        Try
            _cusDef = New SpilCommon.CustomerDef(1)
            txtOrderConfirmation.Text = _cusDef.EmailMsgCrdNote
            txtPerformaInvoice.Text = _cusDef.EmailMsgNcr
            txtQuote.Text = _cusDef.EmailMsgQuote
            txtSalesOrder.Text = _cusDef.EmailMsgSalesOrder
            txtStatement.Text = _cusDef.EmailMsgStatement
            txtTaxInvoice.Text = _cusDef.EmailMsgTaxInv
            utxtEmailNtfyBody.Text = _cusDef.EmailNotificationBody
            utxtEmailNtfySubject.Text = _cusDef.EmailNotificationHeader
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error in CliDef table")
        End Try
        dataset()
        'SetAccountsDetails()
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        If tcMain.SelectedTab.Key = 0 Then
            Dim con As New SpilDataBaseConnection.SecondaryDBConn
            Try
                con.SecondaryDbBeginTrans()
                Dim _cusDef As SpilCommon.CustomerDef = New SpilCommon.CustomerDef()
                _cusDef.EmailMsgCrdNote = txtOrderConfirmation.Text
                _cusDef.EmailMsgNcr = txtPerformaInvoice.Text
                _cusDef.EmailMsgQuote = txtQuote.Text
                _cusDef.EmailMsgSalesOrder = txtSalesOrder.Text
                _cusDef.EmailMsgStatement = txtStatement.Text
                _cusDef.EmailMsgTaxInv = txtTaxInvoice.Text
                _cusDef.EmailNotificationBody = utxtEmailNtfySubject.Text
                _cusDef.EmailNotificationHeader = utxtEmailNtfyBody.Text
                _cusDef.Save(con)
                con.SecondaryDbCommitTrans()
                If MsgBox("Default Email Contents saved successfully! Do you want to exit?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Error in CliDef table")
                con.SecondaryDbRollback()
            End Try
        ElseIf tcMain.SelectedTab.Key = 1 Then
            CCEmailDataUpdate()
        ElseIf tcMain.SelectedTab.Key = 2 Then
            If uChkSetDefault.Checked = True And uTxtDeEmail.Text <> "" Then
                DefaultEmailDataUpdate()
            Else
                If MsgBox("No default email for Invoice and Statement! Do you want to exit?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                    Me.Close()
                End If
            End If
        End If
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Dispose()
        Me.Close()
    End Sub
   
    Private Function EmailCCDataGet(oSQLQuery) As DataSet
        Try
            If IsNothing(oSQL) Then
                oSQL = New clsSqlConn
            End If
            If oSQLQuery = String.Empty Then
                oSQLQuery = "SELECT ModuleName, DefaultCCEmail, DefaultBccEmail, DefaultSender,DefaultSenderEmail FROM spilEmailCCDefault"
            End If
            dsValues = oSQL.GET_INSERT_UPDATE(oSQLQuery)
            Return dsValues
        Catch ex As Exception
            Return dsEmailCCData
        Finally
            oSQL.Dispose()
            oSQLQuery = ""
            dsValues = Nothing
        End Try
    End Function

    Private Sub dataset()
        Try
            Dim dsCCEmail As DataSet = EmailCCDataGet(oSQLQuery)
            For Each dr As DataRow In dsCCEmail.Tables(0).Rows
                'Cc
                If dr("moduleName") = "SalesOrder" Then
                    utxtSalesOrder.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtSalesOrderBcc.Value = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "ProformaInvoice" Then
                    utxtProformaInvoice.Text = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtProformaInvoiceBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "OrderConfirmation" Then
                    utxtOrderConfirmation.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtOrderConfirmationBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "TaxInvoice" Then
                    utxtTaxInvoice.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtTaxInvoiceBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                    If dr("DefaultSenderEmail") <> "" Then
                        uChkSetDefault.Checked = True
                    Else
                        uChkSetDefault.Checked = False
                    End If

                    isSetDefAccount = dr("DefaultSender")
                    utxtTaxInvoice.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtTaxInvoiceBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "Quotation" Then
                    utxtQuotation.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtQuotationBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "NCR" Then
                    utxtNCR.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtNCRBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "CreditNote" Then
                    utxtCreditNote.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtCreditNoteBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "DeliveryDocket" Then
                    utxtDeliveryDocket.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtDeliveryDocketBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                ElseIf dr("moduleName") = "Statment" Then
                    utxtStatment.Value = IIf(IsDBNull(dr("defaultCCEmail")), "", dr("defaultCCEmail"))
                    utxtStatmentBcc.Text = IIf(IsDBNull(dr("defaultBccEmail")), "", dr("defaultBccEmail"))

                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "SPIL-Glass Error")
        Finally
            oSQL = Nothing
        End Try
        'Me.Dispose()
    End Sub
    Private Sub CCEmailDataUpdate()
        Try
          
            Dim SqlQuery As String
            Dim ds As DataSet = EmailCCDataGet(oSQLQuery)
            For Each dr As DataRow In ds.Tables(0).Rows
                If dr("moduleName") = "SalesOrder" Then
                    dr("defaultCCEmail") = IIf(utxtSalesOrder.Value = Nothing, "", utxtSalesOrder.Value)
                    dr("defaultBccEmail") = IIf(utxtSalesOrderBcc.Value = Nothing, "", utxtSalesOrderBcc.Value)

                ElseIf dr("moduleName") = "ProformaInvoice" Then
                    dr("defaultCCEmail") = IIf(utxtProformaInvoice.Value = Nothing, "", utxtProformaInvoice.Value)
                    dr("defaultBccEmail") = IIf(utxtProformaInvoiceBcc.Value = Nothing, "", utxtProformaInvoiceBcc.Value)

                ElseIf dr("moduleName") = "OrderConfirmation" Then
                    dr("defaultCCEmail") = IIf(utxtOrderConfirmation.Value = Nothing, "", utxtOrderConfirmation.Value)
                    dr("defaultBccEmail") = IIf(utxtOrderConfirmationBcc.Value = Nothing, "", utxtOrderConfirmationBcc.Value)

                ElseIf dr("moduleName") = "TaxInvoice" Then
                    dr("defaultCCEmail") = IIf(utxtTaxInvoice.Value = Nothing, "", utxtTaxInvoice.Value)
                    dr("defaultBccEmail") = IIf(utxtTaxInvoiceBcc.Value = Nothing, "", utxtTaxInvoiceBcc.Value)

                ElseIf dr("moduleName") = "Quotation" Then
                    dr("defaultCCEmail") = IIf(utxtQuotation.Value = Nothing, "", utxtQuotation.Value)
                    dr("defaultBccEmail") = IIf(utxtQuotationBcc.Value = Nothing, "", utxtQuotationBcc.Value)

                ElseIf dr("moduleName") = "NCR" Then
                    dr("defaultCCEmail") = IIf(utxtNCR.Value = Nothing, "", utxtNCR.Value)
                    dr("defaultBccEmail") = IIf(utxtNCRBcc.Value = Nothing, "", utxtNCRBcc.Value)

                ElseIf dr("moduleName") = "CreditNote" Then
                    dr("defaultCCEmail") = IIf(utxtCreditNote.Value = Nothing, "", utxtCreditNote.Value)
                    dr("defaultBccEmail") = IIf(utxtCreditNoteBcc.Value = Nothing, "", utxtCreditNoteBcc.Value)

                ElseIf dr("moduleName") = "DeliveryDocket" Then
                    dr("defaultCCEmail") = IIf(utxtDeliveryDocket.Value = Nothing, "", utxtDeliveryDocket.Value)
                    dr("defaultBccEmail") = IIf(utxtDeliveryDocketBcc.Value = Nothing, "", utxtDeliveryDocketBcc.Value)

                ElseIf dr("moduleName") = "Statment" Then
                    dr("defaultCCEmail") = IIf(utxtStatment.Value = Nothing, "", utxtStatment.Value)
                    dr("defaultBccEmail") = IIf(utxtStatmentBcc.Value = Nothing, "", utxtStatmentBcc.Value)

                End If
            Next
            'Email fields validations
            For Each dr As DataRow In ds.Tables(0).Rows
                If dr("defaultCCEmail") <> String.Empty Then
                    If dr("defaultCCEmail").Contains(";") Then
                        Dim emailddress As String() = Split(dr("defaultCCEmail"), ";")
                        Dim i As Integer
                        For i = 0 To UBound(emailddress)
                            Dim emailWithoutSpace As String = emailddress(i)
                            If IsValidEmailFormat(emailWithoutSpace.Replace(" ", "")) = False Then
                                MsgBox("Please enter a valid email address")
                                Exit Sub
                            End If
                        Next
                    Else
                        If IsValidEmailFormat(dr("defaultCCEmail")) = False Then
                            MsgBox("Please enter a valid email address")
                            Exit Sub
                        End If
                    End If
                End If
                If dr("defaultBccEmail") <> String.Empty Then
                    If dr("defaultCCEmail").Contains(";") Then
                        Dim emailddress As String() = Split(dr("defaultBccEmail"), ";")
                        Dim i As Integer
                        For i = 0 To UBound(emailddress)
                            Dim emailWithoutSpace As String = emailddress(i)
                            If IsValidEmailFormat(emailWithoutSpace.Replace(" ", "")) = False Then
                                MsgBox("Please enter a valid email address")
                                Exit Sub
                            End If
                        Next
                    Else
                        If IsValidEmailFormat(dr("defaultBccEmail")) = False Then
                            MsgBox("Please enter a valid email address")
                            Exit Sub
                        End If
                    End If
                End If
            Next

            If IsNothing(oSQL) Then
                oSQL = New clsSqlConn
            End If
            oSQL.Begin_Trans()
            If UpdateDefaultEmail(ds, oSQL) = 0 Then
                ShowMessage("Data not saved.")
                oSQL.Rollback_Trans()
                Exit Sub
            End If
            oSQL.Commit_Trans()
            If MsgBox("Default Cc And Bcc emails saved successfully! Do you want to exit?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                Me.Close()
            Else
                'dataset()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            objExemption = Nothing
        End Try
    End Sub
    Sub dataInvStaEmailUpdate()
        Dim deEmail As String = uTxtDeEmail.Value
    End Sub
    Public Function UpdateDefaultEmail(ds As DataSet, oSQLNew As clsSqlConn) As Integer
        Try
            oSQLQuery = ""
            Dim bError As Boolean = False
            For Each dr As DataRow In ds.Tables(0).Rows
                oSQLQuery = "UPDATE spilEmailCCDefault SET DefaultCCEmail = '" & dr("DefaultCCEmail") & "', DefaultBccEmail = '" & dr("DefaultBccEmail") & "' WHERE ModuleName = '" & dr("ModuleName") & "'  "
                If oSQLNew.Exe_Query_Trans(oSQLQuery) = 0 Then
                    bError = True
                End If
            Next

            If bError Then
                Return 0
            Else
                Return 1
            End If
        Catch ex As Exception
            Return 0
        Finally
            oSQLQuery = ""
        End Try
    End Function
    Private Sub ShowMessage(strError As String)
        MsgBox(strError, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Message")
    End Sub
    Function IsValidEmailFormat(ByVal s As String) As Boolean
        Return Regex.IsMatch(s, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$")
    End Function

    Private Sub frmEmailSettings_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Me.Dispose()
    End Sub

    Private Sub uChkSetDefault_CheckedChanged(sender As Object, e As EventArgs) Handles uChkSetDefault.CheckedChanged
        Try
            Dim isFieldsEmplty As Boolean = True
            If uChkSetDefault.Checked = True Then
                Me.uCmbDeAcc.Visible = True
                Me.uTxtDeEmail.Visible = True
                lblAccounts.Visible = True
                lblEmail.Visible = True

            ElseIf uChkSetDefault.Checked = False Then
                If uTxtDeEmail.Text <> "" Or uCmbDeAcc.Value <> Nothing Then
                    If MsgBox("Do you want to unset default email for Invoice and Statement?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                        oSQLQuery = "UPDATE spilEmailCCDefault SET DefaultSender = '0', DefaultSenderEmail='' WHERE ModuleName IN ('TaxInvoice','Statment') "
                        oSQL.GET_INSERT_UPDATE(oSQLQuery)
                        isSetDefAccount = 0
                        Me.uTxtDeEmail.Visible = False
                        Me.uCmbDeAcc.Visible = False
                        lblAccounts.Visible = False
                        lblEmail.Visible = False
                        isFieldsEmplty = True
                        uCmbDeAcc.Value = Nothing
                        uTxtDeEmail.Text = ""
                    Else
                        isFieldsEmplty = False
                        uChkSetDefault.Checked = True
                    End If
                End If
                If isFieldsEmplty = True Then
                    Me.uTxtDeEmail.Visible = False
                    Me.uCmbDeAcc.Visible = False
                    lblAccounts.Visible = False
                    lblEmail.Visible = False
                End If
            Else
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            oSQLQuery = ""
        End Try
    End Sub

    Private Sub uCmbDeAcc_VisibleChanged(sender As Object, e As EventArgs) Handles uCmbDeAcc.VisibleChanged
        If uCmbDeAcc.Visible = True Then
            SetAccountsDetails()
        End If
    End Sub
    Sub SetAccountsDetails()
        Try
            oSQLQuery = "SELECT AgentID, AgentName, emailFrom FROM spilAgentEmailSettings"
            dsValues = EmailCCDataGet(oSQLQuery)
            uCmbDeAcc.DataSource = dsValues
            uCmbDeAcc.DisplayMember = "AgentName"
            uCmbDeAcc.ValueMember = "AgentID"

            uCmbDeAcc.DisplayLayout.Bands(0).ColHeadersVisible = False
            uCmbDeAcc.DisplayLayout.Bands(0).Columns("AgentID").Hidden = True
            uCmbDeAcc.DisplayLayout.Bands(0).Columns("emailFrom").Hidden = True

            If isSetDefAccount <> 0 Then
                uCmbDeAcc.Value = isSetDefAccount
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            oSQLQuery = ""
        End Try
    End Sub

    Private Sub uCmbDeAcc_ValueChanged(sender As Object, e As EventArgs) Handles uCmbDeAcc.ValueChanged
        If isManully = False Then
            Try
                If IsNothing(uCmbDeAcc.SelectedRow) Then
                    uTxtDeEmail.Value = ""
                Else
                    uTxtDeEmail.Value = IIf(uCmbDeAcc.SelectedRow.Cells("emailFrom").Value = Nothing, "", uCmbDeAcc.SelectedRow.Cells("emailFrom").Value)
                End If
            Catch
                uTxtDeEmail.Value = ""
            End Try
        End If
    End Sub
    Sub DefaultEmailDataUpdate()
        Try
            If IsNothing(oSQL) Then
                oSQL = New clsSqlConn
            End If
            Dim sqlQuery As String = ""
            sqlQuery = "SELECT emailSMTPServer, emailOutgoigPort, emailUserName, emailPassword, AgentID FROM spilAgentEmailSettings WHERE emailUserName='" & uTxtDeEmail.Text & "'"
            Dim avalbleEmailAccount = oSQL.GET_INSERT_UPDATE(sqlQuery)
            Dim defaultAgentID As Integer = 0
            Dim defaultemailUserName As String = ""
            For Each requiredAccountDetails As DataRow In avalbleEmailAccount.Tables(0).Rows
                defaultAgentID = requiredAccountDetails("AgentID")
                defaultemailUserName = requiredAccountDetails("emailUserName")
            Next

            If avalbleEmailAccount.Tables(0).Rows.Count > 0 Then
                'There is an email account in spilAgentEmailSettings
                sqlQuery = "UPDATE spilEmailCCDefault SET DefaultSender = '" & defaultAgentID & "', DefaultSenderEmail='" & defaultemailUserName & "' WHERE ModuleName IN ('TaxInvoice','Statment') "

                If oSQL.Exe_Query(sqlQuery) = 1 Then
                    If MsgBox("Default email for Invoice and Statement saved successfully! Do you want to exit?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Confirmation") = MsgBoxResult.Yes Then
                        Me.Close()
                    End If
                Else
                    ShowMessage("Default email not saved. Please check the email again")
                End If
            Else
                'There is no email account in spilAgentEmailSettings
                MsgBox("There is no matching email-username for entered email!" & vbCrLf & "Please check the email-username in selected user's email settings", MsgBoxStyle.Exclamation, "SPIL Glass")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub uTxtDeEmail_AfterEnterEditMode(sender As Object, e As EventArgs) Handles uTxtDeEmail.AfterEnterEditMode
        isManully = True
        uCmbDeAcc.ValueMember = "Manully"
        uCmbDeAcc.ValueMember = 0
        uCmbDeAcc.Value = Nothing
    End Sub

    Private Sub uCmbDeAcc_Click(sender As Object, e As EventArgs) Handles uCmbDeAcc.Click
        isManully = False
    End Sub

    Private Sub uCmbDeAcc_BeforeDropDown(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles uCmbDeAcc.BeforeDropDown
        SetAccountsDetails()
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class