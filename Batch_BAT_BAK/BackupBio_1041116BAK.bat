@echo on
REM ------1.備份資料庫------
@echo off
osql -U biorest -P biorest7632 -S localhost -Q "BACKUP DATABASE DB_BIO TO DISK='c:\Batch\DB_BIO.bak'"
pause
osql -U biorest -P biorest7632 -S localhost -Q "BACKUP DATABASE DB_SEC TO DISK='c:\Batch\DB_SEC.bak'"
pause
@echo on
REM ------2.加密壓縮成ZIP FILE------
@echo off
cd C:\7-Zip
7z -pbiobank-707 a -tzip C:\Batch\DB_BIO.zip C:\Batch\DB_BIO.bak
7z -pbiobank-707 a -tzip C:\Batch\DB_SEC.zip C:\Batch\DB_SEC.bak  

@echo on
REM ------3.搬移至副機------
@echo off
REM move C:\Batch\DB_BIO.zip \\192.168.1.3\Batch\
REM move C:\Batch\DB_SEC.zip \\192.168.1.3\Batch\

@echo on
REM ------4.Delete Temp File------
@echo off
del C:\Batch\DB_BIO.bak  
del C:\Batch\DB_SEC.bak  

REM @ echo on
