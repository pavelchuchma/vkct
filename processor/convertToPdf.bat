@setlocal
:: set inst="-env:UserInstallation=file:///Temp/LibreOffice_Conversion_${USER}"
set inst="-env:UserInstallation=$SYSUSERCONFIG/LibreOffice/pdfExport.tmp"
"C:\Program Files\LibreOffice\program\scalc.exe" --convert-to pdf:writer_pdf_Export %INST% --outdir . vysledky2022.xlsx
echo %ERRORLEVEL%