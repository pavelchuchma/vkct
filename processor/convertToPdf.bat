@setlocal
:: set inst="-env:UserInstallation=file:///Temp/LibreOffice_Conversion_${USER}"
call :convertToPdf "vysledky2024.xlsx" "VKCT 2024.pdf"
goto :eof

:convertToPdf
    set src=%~1
    set out=%~dpn1.pdf
    set target=%~2
    if exist "%out%" del "%out%" || pause
    if exist "%target%" del "%target%" || pause

    set inst="-env:UserInstallation=$SYSUSERCONFIG/LibreOffice/pdfExport.tmp"
    "C:\Program Files\LibreOffice\program\scalc.exe" --convert-to pdf:writer_pdf_Export %INST% --outdir . vysledky2024.xlsx || pause
    move "%out%" "%target%" || pause
