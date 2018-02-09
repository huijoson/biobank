@Echo off
copy  C:\Batch\DB_BIO.zip C:\Batch2\
copy  C:\Batch\DB_SEC.zip C:\Batch2\
cd C:\7-Zip

@Echo on
REM ------1.解壓縮 DB_EXE.zip  -o/目錄名稱------
@Echo off
7z -pPWD x -y -o/Batch2  C:\Batch2\DB_BIO.zip  
7z -pPWD x -y -o/Batch2  C:\Batch2\DB_SEC.zip  

cd \
@Echo on
REM ------2.條件還原資料庫 DB_BIO & SB_SEC------
@Echo off

osql -U ID -P PWD -S localhost  -d master -Q "ALTER DATABASE [DB_BIO] SET  SINGLE_USER WITH ROLLBACK IMMEDIATE"

osql -U ID -P PWD-S localhost  -d master -Q "RESTORE DATABASE DB_BIO FROM DISK='c:\Batch2\DB_BIO.bak' with Replace"
pause




@Echo on
REM ------3.刪除解壓縮的暫存檔案------
@Echo off
del C:\Batch2\DB_BIO.bak
del C:\Batch2\DB_SEC.bak

cd C:\Batch2

@Echo on 
REM ------4.完成------