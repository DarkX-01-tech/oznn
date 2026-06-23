-- Nöbet Liste Sistemi - Veritabanı Şeması (SQL Server)

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'NobetListeDosyalar')
BEGIN
    CREATE TABLE NobetListeDosyalar (
        id INT IDENTITY(1,1) PRIMARY KEY,
        yil INT NOT NULL,
        bina NVARCHAR(50) NOT NULL,
        ay_klasor NVARCHAR(20) NOT NULL,
        dosya_adi NVARCHAR(255) NOT NULL,
        baslik NVARCHAR(255) NOT NULL,
        aktif BIT NOT NULL DEFAULT 1,
        yukleyen NVARCHAR(100) NULL,
        olusturma_tarihi DATETIME NOT NULL DEFAULT GETDATE(),
        guncelleme_tarihi DATETIME NULL,
        CONSTRAINT UQ_NobetListeDosyalar UNIQUE (yil, bina, ay_klasor, dosya_adi)
    );
END

IF NOT EXISTS (SELECT 1 FROM sys.tables WHERE name = 'NobetAdminKullanicilar')
BEGIN
    CREATE TABLE NobetAdminKullanicilar (
        id INT IDENTITY(1,1) PRIMARY KEY,
        kullanici_adi NVARCHAR(50) NOT NULL UNIQUE,
        sifre NVARCHAR(255) NOT NULL,
        ad_soyad NVARCHAR(100) NULL,
        aktif BIT NOT NULL DEFAULT 1,
        olusturma_tarihi DATETIME NOT NULL DEFAULT GETDATE()
    );

    INSERT INTO NobetAdminKullanicilar (kullanici_adi, sifre, ad_soyad)
    VALUES ('admin', 'admin123', 'Sistem Yöneticisi');
END
