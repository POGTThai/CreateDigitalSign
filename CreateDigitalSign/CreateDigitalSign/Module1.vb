Imports System
Imports System.IO
Imports System.Security
Imports System.Text
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.Security.Cryptography.X509Certificates

Imports iTextSharp.text.pdf
Imports iTextSharp.text.pdf.security
Imports iTextSharp.text.pdf.BarcodeQRCode
Imports iTextSharp.text

Module Module1
    ReadOnly SettingFile As String = AppDomain.CurrentDomain.BaseDirectory.Trim
    ReadOnly SettingFileName As String = SettingFile & "Setting.txt"
    Dim CertPath, MasterFilePDF, DigitalSignPDF As String
    Dim objReader As StreamReader
    Dim strSettingDetail As String
    Dim arrSettingDetail() As String
    Dim strMasterPDF, strDestinationPDF, strCertificatePath, strCertificatePassword, strImagePath As String
    Dim strDigitalSignReason, strDigitalSignLocation, strRealtime, strLogPath, strMoveFileTo, strErrorPath, strErrorFolder As String
    Dim strImageVisible, strDigitalVisible, strDigitalPosition As String
    Sub Main()
        ReadSettingFile()
        StampCertificateDetail()
    End Sub
    Public Sub ReadSettingFile()
        If File.Exists(SettingFileName) Then
            objReader = New StreamReader(SettingFileName)
            Do While objReader.Peek() <> -1
                strSettingDetail = objReader.ReadLine.Trim.Replace(" ", "")

                'Split data 
                arrSettingDetail = Split(strSettingDetail, "=")

                Select Case arrSettingDetail(0).ToString.ToUpper
                    Case "MASTERPDF"
                        strMasterPDF = Trim(arrSettingDetail(1))
                        If Not Directory.Exists(strMasterPDF) Then
                            MessageBox.Show("Directory does not exist " & strMasterPDF, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End
                        End If
                    Case "DESTINATIONPDF"
                        strDestinationPDF = Trim(arrSettingDetail(1))
                        If Not Directory.Exists(strDestinationPDF) Then
                            MessageBox.Show("Directory does not exist " & strDestinationPDF, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End
                        End If
                    Case "LOGFILEPATH"
                        strLogPath = Trim(arrSettingDetail(1))
                        If Not Directory.Exists(strLogPath) Then
                            MessageBox.Show("Directory does not exist " & strLogPath, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End
                        End If
                    Case "CERTIFICATEPATH"
                        strCertificatePath = Trim(arrSettingDetail(1))
                        If Not File.Exists(strCertificatePath) Then
                            MessageBox.Show("Directory does not exist " & strCertificatePath, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End
                        End If
                    Case "MOVEFILETO"
                        strMoveFileTo = Trim(arrSettingDetail(1))
                        If Not Directory.Exists(strMoveFileTo) Then
                            MessageBox.Show("Directory does not exist " & strMoveFileTo, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End
                        End If
                    Case "CERTIFICATEPASSWORD"
                        strCertificatePassword = Trim(arrSettingDetail(1))

                    Case "IMAGEPATH"
                        strImagePath = Trim(arrSettingDetail(1))
                        If strImagePath <> "" Then
                            If Not File.Exists(strImagePath) Then
                                MessageBox.Show("Directory does not exist " & strImagePath, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End
                            End If
                        End If

                    'Add error folder for moving file from master folder.
                    Case "ERRORPATH"
                        strErrorPath = Trim(arrSettingDetail(1)).ToUpper
                        If Not Directory.Exists(strErrorPath) Then
                            MessageBox.Show("Directory does not exist " & strErrorPath, "Please check!!!!!", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            End
                        End If

                    Case "IMAGEVISIBLE"
                        strImageVisible = Trim(arrSettingDetail(1)).ToUpper
                    Case "DIGITALVISIBLE"
                        strDigitalVisible = Trim(arrSettingDetail(1)).ToUpper
                    Case "DIGITALSIGNREASON"
                        strDigitalSignReason = Trim(arrSettingDetail(1))

                    Case "DIGITALSIGNLOCATION"
                        strDigitalSignLocation = Trim(arrSettingDetail(1))
                    Case "DIGITALPOSITION"
                        strDigitalPosition = Trim(arrSettingDetail(1))
                    Case "REALTIME"
                        strRealtime = Trim(arrSettingDetail(1)).ToUpper
                End Select

            Loop
        End If

        objReader.Close()
        objReader.Dispose()
        objReader = Nothing

    End Sub
    Public Sub StampCertificateDetail()

        'Get file from folder
        Dim listfile As FileInfo
        Dim filename, MasterPDFPath, DestinationPDFPath As String
        Dim dir As DirectoryInfo
        Dim strLogDetail As String
        Dim strMoveFileToPath As String

StampCert:

        Try
            dir = New DirectoryInfo(strMasterPDF)
            If dir.GetFiles.Count <> 0 Then

                For Each listfile In dir.GetFiles
                    filename = listfile.Name
                    MasterPDFPath = Path.Combine(strMasterPDF, filename)

                    If File.Exists(MasterPDFPath) Then


                        'Get detail on certificate (Cert path and Password)
                        Dim cert As X509Certificate2 = New X509Certificate2(strCertificatePath, strCertificatePassword)

                        'Source File
                        Dim Reader As PdfReader
                        PdfReader.unethicalreading = True
                        Reader = New PdfReader(MasterPDFPath)

                        'Target File
                        DestinationPDFPath = Path.Combine(strDestinationPDF, filename)
                        Dim File_stream As FileStream = New FileStream(DestinationPDFPath, FileMode.Create, FileAccess.ReadWrite)


                        'Create digital signature
                        Dim stamper As PdfStamper = PdfStamper.CreateSignature(Reader, File_stream, "\0", Nothing, True)
                        Dim sap As PdfSignatureAppearance = stamper.SignatureAppearance
                        Dim cp As Org.BouncyCastle.X509.X509CertificateParser = New Org.BouncyCastle.X509.X509CertificateParser()
                        Dim chain As Org.BouncyCastle.X509.X509Certificate() = New Org.BouncyCastle.X509.X509Certificate() {cp.ReadCertificate(cert.RawData)}
                        Dim externalSign As IExternalSignature = New X509Certificate2Signature(cert, "SHA-1")
                        Dim imagepath As String
                        Dim SignFileImage As iTextSharp.text.Image
                        Dim arrPosition() As String

                        If strDigitalSignReason IsNot Nothing Then
                            sap.Reason = strDigitalSignReason
                        End If

                        If strDigitalSignLocation IsNot Nothing Then
                            sap.Location = strDigitalSignLocation
                        End If

                        '========================================================================================================
                        'Syntax below is image with digital sign detail
                        If strImageVisible = "Y" Then
                            'Image path for digital sign
                            SignFileImage = iTextSharp.text.Image.GetInstance(strImagePath)
                            sap.SignatureGraphic = SignFileImage
                            sap.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.GRAPHIC_AND_DESCRIPTION
                        Else
                            'Syntax below is digital sign detail only.
                            sap.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.DESCRIPTION
                        End If

                        'Syntax below is name of digital sign only.
                        'sap.SignatureRenderingMode = PdfSignatureAppearance.RenderingMode.NAME_AND_DESCRIPTION
                        '========================================================================================================

                        'Set visible digital signature.
                        If strDigitalVisible = "Y" Then
                            arrPosition = Split(strDigitalPosition, ",")
                            sap.SetVisibleSignature(New iTextSharp.text.Rectangle(arrPosition(0), arrPosition(1), arrPosition(2), arrPosition(3)), 1, Nothing)
                        End If

                        'Create digital sign on pdf file
                        MakeSignature.SignDetached(sap, externalSign, chain, Nothing, Nothing, Nothing, 0, CryptoStandard.CMS)

                        strLogDetail = filename & " is successful."


                        strMoveFileToPath = Path.Combine(strMoveFileTo, filename)

                        If File.Exists(strMoveFileToPath) Then
                            File.Delete(strMoveFileToPath)
                        End If
                        System.IO.File.Move(MasterPDFPath, strMoveFileToPath)
                        writelog(strLogDetail)


                    End If
                Next
            End If



        Catch ex As Exception
            strErrorFolder = checkFolderError()
            writelog(ex.Message & " (" & MasterPDFPath & " move to Error folder(" & strErrorFolder & "))")
            listfile.MoveTo(Path.Combine(strErrorFolder, listfile.Name))

        End Try

        If strRealtime = "Y" Then
            Threading.Thread.Sleep(30)
            dir = Nothing
            GoTo StampCert
            'StampCertificateDetail()
        End If

    End Sub

    Function checkFolderError()
        Dim currErrorFolder As String
        currErrorFolder = Path.Combine(strErrorPath, DateTime.Now.ToString("yyyyMM", CultureInfo.InvariantCulture))

        If Not Directory.Exists(currErrorFolder) Then
            Directory.CreateDirectory(currErrorFolder)
        End If

        Return currErrorFolder

    End Function

    Sub writelog(ByVal strMsg)
        Dim logWriter As StreamWriter
        Dim strLogName As String
        Try
            strLogName = Path.Combine(strLogPath, DateTime.Now.ToString("yyyyMMdd", CultureInfo.InvariantCulture) & ".txt")

            strMsg = DateTime.Now.ToString & ControlChars.Tab & ControlChars.Tab & ControlChars.Tab & strMsg

            logWriter = New StreamWriter(strLogName, True)
            With logWriter
                .WriteLine(strMsg)
            End With

            logWriter.Flush()
            logWriter.Close()
            logWriter.Dispose()


        Catch ex As Exception
            MsgBox("Error on writlog sub!!!!  " & ex.Message.ToString)
        End Try

        strLogName = Nothing

    End Sub

End Module
