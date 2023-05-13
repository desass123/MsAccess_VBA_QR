Imports QRCoder

Public Sub InsertInto(id data As INT,Barcode_data data As String,bmp_data data As Bitmap) 
     
    Dim StrSQL As String
     
    StrSQL= " INSERT INTO Report_Barcodes" _ 
        & "(ID,barcode_Data, barcode) VALUES " _ 
        & "('"& id &"', '"& Barcode_data &"', '"&bmp_data &"');" 
         
    DoCmd.SetWarnings False

    DoCmd.RunSQL StrSQL

    DoCmd.SetWarnings True 
     
Public End Sub 

Public Function GenerateQRCode(ByVal data As String) As Bitmap
    Dim qrGenerator As New QRCodeGenerator()
    Dim qrCodeData As QRCodeData = qrGenerator.CreateQrCode(data, QRCodeGenerator.ECCLevel.Q)
    Dim qrCode As New QRCode(qrCodeData)
    Dim qrCodeImage As Bitmap = qrCode.GetGraphic(20)
    Return qrCodeImage
End Function

Private Sub Form_AfterUpdate()
    Dim Barcode_data  As String 'changed Coloumn name'

    id = ME!ID
    Barcode_data = ME!value1 & ME!value2 & ME!value3 'ME = form name'
    
    CreateReport_BarcodesTable() 'Report_Barcodes'

    InsertInto(ME!ID,Barcode_data,GenerateQRCode(Barcode_data))
End Sub

Sub CreateReport_BarcodesTable()
    On Error GoTo ErrHandler
    
    CurrentDb.Execute "CREATE TABLE Report_Barcodes (ID INT, barcode_Data TEXT(255), barcode OLEOBJECT, PRIMARY KEY(ID));"
    
ExitHere:
    Exit Sub
    
ErrHandler:
    
    Resume ExitHere
End Sub