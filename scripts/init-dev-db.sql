CREATE DATABASE IF NOT EXISTS oznn_yemek;
CREATE USER IF NOT EXISTS 'oznn'@'localhost' IDENTIFIED BY 'oznn_dev';
GRANT ALL ON oznn_yemek.* TO 'oznn'@'localhost';
FLUSH PRIVILEGES;

USE oznn_yemek;

CREATE TABLE IF NOT EXISTS yemek_listesi (
  id INT AUTO_INCREMENT PRIMARY KEY,
  tarih DATE NOT NULL,
  gun_adi VARCHAR(20),
  ogle_corba VARCHAR(255),
  ogle_ana_yemek VARCHAR(255),
  ogle_yan_urun VARCHAR(255),
  ogle_tatli VARCHAR(255),
  aksam_corba VARCHAR(255),
  aksam_ana_yemek VARCHAR(255),
  aksam_yan_urun VARCHAR(255),
  aksam_tatli VARCHAR(255),
  aktif TINYINT(1) DEFAULT 1
);

CREATE TABLE IF NOT EXISTS diyet_yemek_listesi LIKE yemek_listesi;

CREATE TABLE IF NOT EXISTS yemek_kisi_sayisi (
  id INT AUTO_INCREMENT PRIMARY KEY,
  tarih DATE,
  menu_tipi VARCHAR(20),
  kisi_sayisi INT,
  kullanici VARCHAR(100),
  kayit_tarihi DATETIME
);

CREATE TABLE IF NOT EXISTS yemek_degisiklik_log (
  id INT AUTO_INCREMENT PRIMARY KEY,
  menu_tarih DATE,
  degisiklik_tarihi DATETIME,
  ogun_tipi VARCHAR(50),
  eski_deger TEXT,
  yeni_deger TEXT,
  kullanici VARCHAR(100),
  menu_tipi VARCHAR(20),
  yil INT,
  ay INT
);

CREATE TABLE IF NOT EXISTS yemek_guncelleme_log (
  id INT AUTO_INCREMENT PRIMARY KEY,
  yil INT,
  ay INT,
  guncelleme_tarihi DATETIME,
  guncelleme_tipi VARCHAR(50),
  kullanici VARCHAR(100),
  aciklama TEXT
);

INSERT INTO yemek_listesi (tarih, gun_adi, ogle_corba, ogle_ana_yemek, ogle_yan_urun, ogle_tatli, aksam_corba, aksam_ana_yemek, aksam_yan_urun, aksam_tatli, aktif)
SELECT * FROM (
  SELECT '2026-06-10' AS tarih, 'Salı' AS gun_adi, 'Mercimek Çorbası', 'Tavuk Sote', 'Pilav', 'Sütlaç', 'Ezogelin', 'Köfte', 'Makarna', 'Meyve', 1 AS aktif
  UNION ALL
  SELECT '2026-06-11', 'Çarşamba', 'Domates Çorbası', 'Balık', 'Sebze', 'Helva', 'Tavuk Çorbası', 'Et Güveç', 'Bulgur', 'Ayran', 1
  UNION ALL
  SELECT '2026-06-12', 'Perşembe', 'Yayla Çorbası', 'Karnıyarık', 'Cacık', 'Revani', 'Mercimek', 'Tavuk Izgara', 'Püre', 'Komposto', 1
) AS seed
WHERE NOT EXISTS (SELECT 1 FROM yemek_listesi LIMIT 1);
