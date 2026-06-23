-- Nöbet listesi dosya takip tablosu (isteğe bağlı)
-- SQL Server

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'NobetListeDosyalar')
BEGIN
    CREATE TABLE NobetListeDosyalar (
        id INT IDENTITY(1,1) PRIMARY KEY,
        ay_klasor NVARCHAR(50) NOT NULL,
        dosya_adi NVARCHAR(255) NOT NULL,
        aktif BIT NOT NULL DEFAULT 0,
        olusturma_tarihi DATETIME NOT NULL DEFAULT GETDATE(),
        guncelleme_tarihi DATETIME NULL,
        CONSTRAINT UQ_NobetListeDosyalar_AyDosya UNIQUE (ay_klasor, dosya_adi)
    );

    CREATE INDEX IX_NobetListeDosyalar_AyKlasor
        ON NobetListeDosyalar (ay_klasor);
END
