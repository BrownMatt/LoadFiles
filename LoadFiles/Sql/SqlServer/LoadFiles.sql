-------------------------------------------------------------
-- Create the FileSheet header table
IF OBJECT_ID('FileSheet', 'U') IS NOT NULL
    BEGIN
        PRINT 'The FileSheet table exists';
    END
ELSE
    BEGIN
        PRINT 'The FileSheet table does not exist';

        CREATE TABLE FileSheet
        (
            FLE_Id INT IDENTITY(1,1) NOT NULL
            ,FLE_File_Name VARCHAR(300) NOT NULL
            ,FLE_File_Sheet VARCHAR(300) NOT NULL
            ,FLE_File_Hash VARCHAR(100) NOT NULL
            CONSTRAINT [PK_FileSheet] PRIMARY KEY CLUSTERED
            (
                [FLE_Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
        )
    END
        GO

-------------------------------------------------------------
-- Create the FileSheet data table
IF OBJECT_ID('FileSheetData', 'U') IS NOT NULL
    BEGIN
        PRINT 'The FileSheetData table exists';
    END
ELSE
    BEGIN
        PRINT 'The FileSheetData table does not exist';

        CREATE TABLE FileSheetData
        (
            FSD_Id INT IDENTITY(1,1) NOT NULL
            ,FSD_FLE_Id INT NOT NULL
            ,FSD_Row INT NOT NULL
            ,FSD_Col INT NOT NULL
            ,FSD_Data VARCHAR(2000) NULL
            CONSTRAINT [PK_FileSheetData] PRIMARY KEY CLUSTERED
            (
                [FSD_Id] ASC
            )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
        )
    END
        GO

-------------------------------------------------------------
-- Gets existing or new ID for FileSheet
--  With logic to reload data when file hash changes
CREATE OR ALTER   PROCEDURE dbo.GetSheetId(
    @FileName VARCHAR (300),
    @Sheet varchar(300),
    @Hash varchar(100)
)
AS
BEGIN

	SET NOCOUNT ON;
	IF(EXISTS(SELECT TOP 1 1 FROM FileSheet WHERE FLE_File_Name = @FileName AND FLE_File_Sheet = @Sheet))
    BEGIN
        IF(EXISTS(SELECT TOP 1 1 FROM FileSheet
                    WHERE FLE_File_Name = @FileName
                      AND FLE_File_Sheet = @Sheet
                      AND FLE_File_Hash = @Hash))
        BEGIN
            SELECT -3
        END
        ELSE
        BEGIN
            -- Get the current sheet id
            DECLARE @sheetId INT;
            SELECT @sheetId=FLE_Id FROM FileSheet
            WHERE FLE_File_Name = @FileName
              AND FLE_File_Sheet = @Sheet;
            DELETE FROM FileSheetData WHERE FSD_FLE_Id = @sheetId;
            SELECT @sheetId;
        END
    END
    ELSE
        BEGIN -- File does not exist
            INSERT INTO FileSheet(FLE_File_Name, FLE_File_Sheet, FLE_File_Hash)
                VALUES(@FileName, @Sheet, @Hash);
            SELECT SCOPE_IDENTITY();
        END

END


/*--drop table regions;
----------------------------------------
-- Sample mapping work
CREATE TABLE regions
( theRow int, regionId VARCHAR(10), regionDescription VARCHAR(100) );

-- first column
INSERT INTO regions(theRow, regionId)
SELECT FSD_Row, fsd.FSD_Data --, *
FROM dbo.FileSheetData fsd
INNER JOIN dbo.FileSheet fs ON fs.FLE_Id = fsd.FSD_FLE_Id
WHERE fs.FLE_File_Name = 'regions.csv'
AND fsd.FSD_Col = 0   -- We are working on the first column
AND fsd.FSD_Row > 1   -- This data has a header row
AND fsd.FSD_Data <> '' -- We are going to ignore blanks

-- second column
UPDATE regions
SET regionDescription = RgnData.FSD_Data
FROM dbo.FileSheetData AS RgnData
INNER JOIN dbo.FileSheet AS RgnHdr  ON RgnData.FSD_FLE_Id = RgnHdr.FLE_Id AND RgnHdr.FLE_File_Name = 'regions.csv'
INNER JOIN dbo.regions AS [R] ON R.theRow = RgnData.FSD_Row AND RgnData.FSD_Col = 1
WHERE [R].theRow = RgnData.FSD_Row
AND RgnData.FSD_Row > 1  -- Ignore the header row
*/
