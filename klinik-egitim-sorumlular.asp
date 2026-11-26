<!-- #include file="admin/database/Connection.asp" -->
<!-- #include file="ayarlar.asp" -->
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>MÜ Pendik E.A.H. Portal</title>
<link rel="icon" href="images/hastane_portal_logo.png"/>
<style type="text/css">
@font-face{font-family:'Open Sans';src:url('/webfonts/OpenSans-Regular.ttf') format('truetype');font-weight:400;font-style:normal}
@font-face{font-family:'Open Sans';src:url('/webfonts/OpenSans-SemiBold.ttf') format('truetype');font-weight:600;font-style:normal}
@font-face{font-family:'Open Sans';src:url('/webfonts/OpenSans-ExtraBold.ttf') format('truetype');font-weight:800;font-style:normal}
body{background:url('images/arka-plan.png') no-repeat center center fixed;background-size:cover;font-family:'Open Sans',sans-serif;margin:0;padding:0;overflow-x:hidden}
.golgeliKutu{box-shadow:0 8px 16px rgba(0,0,0,.2);width:900px;margin:40px auto;background-color:#fff;border-radius:8px;overflow:hidden;transition:box-shadow .3s,background-color .3s}
.golgeliKutu:hover{background-color:#f9f9f9;box-shadow:0 12px 24px rgba(0,0,0,.3)}
.baslik{font-size:24px;color:#850303;text-align:center;margin:25px 0 0;font-weight:600;text-transform:uppercase;letter-spacing:1px;padding-bottom:0;text-shadow:1px 1px 4px rgba(0,0,0,.3);transition:color .3s,border-color .3s;font-family:'Open Sans',sans-serif}
.baslik:hover{color:#343a40;border-color:#c70039}
.baslik::after{content:"";display:block;width:725px;height:2px;background-color:#343a40;margin:15px auto 0}
.tablo-toolbar{display:flex;align-items:center;justify-content:flex-end;padding:8px 18px}
.search-container{margin:0 12px 0 20px;display:flex;justify-content:flex-end;align-items:center}
.search-container form{display:inline-block;position:relative}
.search-container input[type="text"]{width:16ch;min-width:16ch;height:40px;padding:0 34px 0 14px;border:1px solid #cfcfcf;border-radius:22px;font-size:16px;line-height:40px;box-sizing:border-box;transition:width .35s ease,border-color .25s,box-shadow .25s;box-shadow:0 2px 6px rgba(0,0,0,.08);color:#666;background:#fff;transform-origin:right center}
.search-container input[type="text"]::placeholder{color:#8a8a8a}
.search-container input[type="text"].expanded,.search-container input[type="text"]:focus{width:280px}
.search-container input[type="text"]:focus{border-color:#45b8c3;outline:none;box-shadow:0 4px 10px rgba(0,0,0,.12)}
.search-container input[type="text"]:hover{border-color:#45b8c3}
.clear-search-btn{position:absolute;right:10px;top:50%;transform:translateY(-50%);width:18px;height:18px;border:none;border-radius:50%;background:transparent!important;box-shadow:none!important;cursor:pointer;padding:0;display:flex;align-items:center;justify-content:center;transition:all .2s ease}
.clear-search-btn:hover,.clear-search-btn.has-text{background:transparent!important;box-shadow:none!important}
.clear-search-btn .icon-search,.clear-search-btn .icon-clear{position:absolute;left:50%;top:50%;transform:translate(-50%,-50%) scale(1);transition:opacity .22s ease,transform .22s ease,color .22s ease;user-select:none;pointer-events:none;line-height:1}
.clear-search-btn .icon-search{opacity:1;width:12px;height:12px;color:#8a8a8a;font-size:0}
.clear-search-btn .icon-search::before{content:"";position:absolute;left:0;top:0;width:8px;height:8px;border:2px solid currentColor;border-radius:50%;box-sizing:border-box}
.clear-search-btn .icon-search::after{content:"";position:absolute;width:2px;height:5px;background:currentColor;border-radius:2px;transform:rotate(-45deg);right:0;bottom:0;transform-origin:center}
.clear-search-btn .icon-clear{opacity:0;transform:translate(-50%,-50%) scale(.7) rotate(-70deg);color:#8a8a8a;font-size:16px;font-weight:700}
.clear-search-btn.has-text .icon-search{opacity:0;transform:translate(-50%,-50%) scale(.7) rotate(70deg)}
.clear-search-btn.has-text .icon-clear{opacity:1;transform:translate(-50%,-50%) scale(1) rotate(0)}
.search-container input[type="text"]:focus~.clear-search-btn .icon-search,.search-container input[type="text"]:focus~.clear-search-btn .icon-clear{color:#45b8c3}
#noMatchesRow{width:100%;min-height:360px;margin:0;padding:32px 30px 35px;box-sizing:border-box;display:none;flex-direction:column;align-items:center;justify-content:center;text-align:center;background:linear-gradient(135deg,rgba(69,184,195,.045) 0%,#fff 48%,rgba(133,3,3,.018) 100%);border:1px solid #e1e7e9;border-left:4px solid #25abb9;border-bottom:4px solid #25abb9;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,.07),0 1px 3px rgba(0,0,0,.04);position:relative;overflow:hidden;transition:all .25s ease}
#noMatchesRow:hover{border-left-color:#0d5f6b;border-bottom-color:#0d5f6b;box-shadow:0 6px 18px rgba(0,0,0,.14);transform:translateY(-1px)}
#noMatchesRow::after{content:"";position:absolute;top:0;left:14%;width:72%;height:1px;background:linear-gradient(90deg,transparent,rgba(69,184,195,.45),transparent)}
#noMatchesRow .no-result-icon{width:64px;height:64px;flex:0 0 64px;margin:0 0 18px;display:flex;align-items:center;justify-content:center;border-radius:50%;border:1px solid rgba(69,184,195,.38);background:rgba(69,184,195,.08);box-shadow:0 4px 12px rgba(37,171,185,.1);position:relative}
#noMatchesRow .no-result-icon::before{content:"";position:absolute;width:22px;height:22px;left:16px;top:14px;border:3px solid #25abb9;border-radius:50%;box-sizing:border-box}
#noMatchesRow .no-result-icon::after{content:"";position:absolute;width:13px;height:3px;left:35px;top:39px;background:#25abb9;border-radius:3px;transform:rotate(45deg);transform-origin:left center}
#noMatchesRow .no-result-title{margin:0 0 13px;font-family:'Open Sans','Segoe UI',Tahoma,sans-serif;font-size:20px;line-height:1.45;font-weight:800;color:#25abb9;text-align:center;letter-spacing:.1px}
#noMatchesRow .no-result-divider{width:72%;height:1px;margin:0 auto 18px;background:linear-gradient(90deg,transparent 0%,#8bd3da 20%,#45b8c3 50%,#8bd3da 80%,transparent 100%);position:relative}
#noMatchesRow .no-result-divider::after{content:"";position:absolute;width:8px;height:8px;left:50%;top:50%;transform:translate(-50%,-50%);background:#45b8c3;border:2px solid #fff;border-radius:50%;box-shadow:0 0 0 1px #45b8c3}
#noMatchesRow .no-result-text{margin:0;font-family:'Open Sans','Segoe UI',Tahoma,sans-serif;font-size:13px;line-height:1.85;font-weight:500;color:#666;text-align:center}
.tablo-container{box-sizing:border-box;max-height:1290px;overflow-y:auto;overflow-x:hidden;overscroll-behavior:contain;margin:0 18px 14px;border:1px solid #e1edef;border-radius:14px;background:#fff;box-shadow:0 1px 2px rgba(30,140,153,.06),0 8px 24px rgba(30,140,153,.08);min-height:0;color:rgba(37,171,185,.28);transition:color .45s ease,box-shadow .3s ease}
.tablo-container:hover{color:#45b8c3;box-shadow:0 1px 2px rgba(30,140,153,.08),0 12px 28px rgba(30,140,153,.12)}
.tablo-container::-webkit-scrollbar{width:8px}
.tablo-container::-webkit-scrollbar-track{background:transparent;margin:46px 0 10px}
.tablo-container::-webkit-scrollbar-thumb{background-color:transparent;border-radius:10px;border:1px solid transparent;background-clip:padding-box;box-shadow:inset 0 0 0 10px}
.tablo-container::-webkit-scrollbar-thumb:hover{box-shadow:inset 0 0 0 10px #1e8c99}
@supports(-moz-appearance:none){
.tablo-container{scrollbar-width:thin;scrollbar-color:#b9e2e7 transparent;transition:scrollbar-color .45s ease,box-shadow .3s ease}
.tablo-container:hover{scrollbar-color:#45b8c3 transparent}
}
.bolum-tablo{width:100%;border-collapse:separate;border-spacing:0;font-family:'Open Sans',sans-serif;color:#444}
.bolum-tablo thead th{position:-webkit-sticky;position:sticky;top:0;z-index:5;background:linear-gradient(180deg,#fff 0%,#f4fafb 100%);color:#1e8c99;padding:14px 16px 12px;font-size:11px;font-weight:800;letter-spacing:1.2px;text-transform:uppercase;text-align:left;border-bottom:2px solid #25abb9;box-shadow:0 6px 10px -8px rgba(30,140,153,.45);white-space:nowrap}
.bolum-tablo thead th.sira-no-th{text-align:center;width:64px}
.bolum-tablo tbody td{padding:11px 16px;vertical-align:middle;border-bottom:1px solid #eef3f4;background-color:#fff;transition:background-color .25s ease,color .25s ease,box-shadow .25s ease}
.bolum-tablo tbody tr:nth-child(even) td{background-color:#fafdfd}
.bolum-tablo tbody tr:last-child td{border-bottom:none}
.bolum-tablo tbody tr:hover td{background-color:#eef8f9}
.bolum-tablo tbody td:first-child{box-shadow:inset 0 0 0 #25abb9}
.bolum-tablo tbody tr:hover td:first-child{box-shadow:inset 4px 0 0 #25abb9}
.sira-no-cell{text-align:center;width:64px}
.sira-rozet{position:relative;display:inline-flex;align-items:center;justify-content:center;width:30px;height:30px;border-radius:50%;background:linear-gradient(145deg,#fff 0%,#e9f6f7 100%);color:#1e8c99;font-size:12px;font-weight:700;font-variant-numeric:tabular-nums;box-shadow:inset 0 0 0 1.5px #a9dce1,0 2px 5px rgba(30,140,153,.12);transition:transform .3s cubic-bezier(.34,1.56,.64,1),box-shadow .3s ease,color .3s ease}
.sira-rozet::before{content:"";position:absolute;top:0;left:0;right:0;bottom:0;border-radius:50%;background:linear-gradient(145deg,#45b8c3 0%,#1e8c99 100%);opacity:0;transform:scale(.4);transition:opacity .3s ease,transform .3s ease}
.sira-rozet span{position:relative;z-index:1}
.bolum-tablo tbody tr:hover .sira-rozet{color:#fff;transform:scale(1.1);box-shadow:inset 0 0 0 1.5px transparent,0 4px 12px rgba(30,140,153,.35),0 0 0 4px rgba(69,184,195,.15)}
.bolum-tablo tbody tr:hover .sira-rozet::before{opacity:1;transform:scale(1)}
.bolum-adi-cell{font-size:13px;font-weight:600;color:#2b3a42;transition:color .25s ease,padding-left .25s ease}
.bolum-tablo tbody tr:hover .bolum-adi-cell{color:#1e8c99;padding-left:20px}
.sorumlu-adi-cell{font-size:13px;font-weight:600;color:#3d4a50;line-height:1.5}
.sorumlu-unvan{display:block;font-size:10.5px;font-weight:600;color:#8a9aa0;letter-spacing:.3px;margin-bottom:1px}
.bolum-tablo tbody tr:hover .sorumlu-unvan{color:#45b8c3}
.bolum-tablo tr.search-hidden{display:none!important}
#Table_01>tbody>tr:nth-child(2)>td:first-child{background:#f9fcfd;border-right:1px solid #dae7ea}
#scrollTopBtn{position:fixed;bottom:30px;right:30px;width:42px;height:42px;background-color:#25abb9;color:#fff;font-size:20px;text-align:center;line-height:42px;border-radius:50%;box-shadow:0 4px 10px rgba(37,171,185,.4);cursor:pointer;transition:all .3s ease,opacity .5s ease;opacity:0;visibility:hidden;z-index:999}
#scrollTopBtn:hover{background-color:#1e8c99;box-shadow:0 6px 14px rgba(37,171,185,.55);transform:scale(1.1)}
</style>
<script>
document.addEventListener("DOMContentLoaded",function(){
  setTimeout(function(){window.scrollTo({top:0,behavior:'smooth'})},15000);
  var btn=document.getElementById("scrollTopBtn");
  window.addEventListener("scroll",function(){
    if(window.scrollY>300){btn.style.visibility='visible';btn.style.opacity='1'}
    else{btn.style.opacity='0';btn.style.visibility='hidden'}
  });
});
function scrollToTop(){window.scrollTo({top:0,behavior:'smooth'})}
</script>
</head>
<body>
<div id="scrollTopBtn" onclick="scrollToTop()">^</div>
<div align="center" class="golgeliKutu">
<table id="Table_01" width="900" border="0" cellpadding="0" cellspacing="0">
<tr>
<td colspan="3"><!--#include virtual="/banner.asp"--></td>
</tr>
<tr>
<td bgcolor="#FFFFFF" width="166" height="604" valign="top"><!--#include file="sol.asp"--></td>
<td bgcolor="#FFFFFF" width="734" height="604" valign="top" align="center">
<div class="baslik">KLİNİK EĞİTİM SORUMLULAR</div>
<div class="tablo-toolbar">
<div class="search-container">
<form onsubmit="return false;">
<input type="text" id="search-bar" placeholder="Bölüm Ara...">
<button type="button" id="clearSearchBtn" class="clear-search-btn" title="Ara" aria-label="Ara">
<span class="icon-search" aria-hidden="true"></span>
<span class="icon-clear" aria-hidden="true">&times;</span>
</button>
</form>
</div>
</div>
<div class="tablo-container">
<table class="bolum-tablo">
<thead>
<tr>
<th class="sira-no-th">Sıra</th>
<th>Bölüm</th>
<th>KLİNİK EĞİTİM SORUMLUSU</th>
</tr>
</thead>
<tbody id="tabloGovde">
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>1</span></span></td><td class="bolum-adi-cell">Acil Tıp</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Arzu Denizbaşı Altınok</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>2</span></span></td><td class="bolum-adi-cell">Adli Tıp</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Mehmet Akif İnanıcı</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>3</span></span></td><td class="bolum-adi-cell">Aile Hekimliği</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Memnune Çiğdem Apaydın Kaya</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>4</span></span></td><td class="bolum-adi-cell">Algoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Savaş Şencan</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>5</span></span></td><td class="bolum-adi-cell">Anatomi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ümit Süleyman Şehirli</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>6</span></span></td><td class="bolum-adi-cell">Anesteziyoloji ve Reanimasyon</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Alper Kararmaz</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>7</span></span></td><td class="bolum-adi-cell">Anesteziyoloji ve Reanimasyon Yoğun Bakım</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Fethi Gül</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>8</span></span></td><td class="bolum-adi-cell">Beyin ve Sinir Cerrahisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Fatih Bayraklı</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>9</span></span></td><td class="bolum-adi-cell">Çocuk Cerrahisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Gürsu Kıyan</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>10</span></span></td><td class="bolum-adi-cell">Çocuk Endokrinolojisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Serap Demircioğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>11</span></span></td><td class="bolum-adi-cell">Çocuk Enfeksiyon Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Sevliya Öcal Demir</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>12</span></span></td><td class="bolum-adi-cell">Çocuk Gastroenterolojisi Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Deniz Ertem Şahinoğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>13</span></span></td><td class="bolum-adi-cell">Çocuk Genetik Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Huriye Nursel Elçioğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>14</span></span></td><td class="bolum-adi-cell">Çocuk Göğüs Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ela Erdem Eralp</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>15</span></span></td><td class="bolum-adi-cell">Çocuk Hematolojisi ve Onkolojisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ahmet Koç</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>16</span></span></td><td class="bolum-adi-cell">Çocuk İmmünoloji ve Alerji Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ahmet Oğuzhan Özen</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>17</span></span></td><td class="bolum-adi-cell">Çocuk Kardiyolojisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Figen Akalın</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>18</span></span></td><td class="bolum-adi-cell">Çocuk Metabolizma Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Burcu Öztürk Hişmi</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>19</span></span></td><td class="bolum-adi-cell">Çocuk Nefrolojisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> İbrahim Gökçe</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>20</span></span></td><td class="bolum-adi-cell">Çocuk Nöroloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Olcay Ünver</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>21</span></span></td><td class="bolum-adi-cell">Çocuk Ürolojisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Çağrı Akın Şekerci</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>22</span></span></td><td class="bolum-adi-cell">Çocuk Sağlığı ve Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Serap Demircioğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>23</span></span></td><td class="bolum-adi-cell">Çocuk ve Ergen Ruh Sağlığı ve Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Neşe Perdahlı Fiş</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>24</span></span></td><td class="bolum-adi-cell">Çocuk Yoğun Bakım</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Makbule Nilüfer Yalındağ</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>25</span></span></td><td class="bolum-adi-cell">Deri ve Zührevi Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ayşe Deniz Yücelten</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>26</span></span></td><td class="bolum-adi-cell">Enfeksiyon Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Lütfiye Mülazimoğlu Durmuşoğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>27</span></span></td><td class="bolum-adi-cell">Endokrinoloji ve Metabolizma Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Dilek Yavuz</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>28</span></span></td><td class="bolum-adi-cell">Fiziksel Tıp ve Rehabilitasyon</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Mehmet Tuncay Duruöz</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>29</span></span></td><td class="bolum-adi-cell">Fizyoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Özgür Kasımay Berköz</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>30</span></span></td><td class="bolum-adi-cell">Gastroenteroloji Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Özlen Atuğ</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>31</span></span></td><td class="bolum-adi-cell">Genel Cerrahi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Asım Cingi</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>32</span></span></td><td class="bolum-adi-cell">Geriatri</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Aslı Tufan Çinçin</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>33</span></span></td><td class="bolum-adi-cell">Göğüs Cerrahisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Korkut Bostancı</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>34</span></span></td><td class="bolum-adi-cell">Göğüs Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Sait Karakurt</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>35</span></span></td><td class="bolum-adi-cell">Göz Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Volkan Dericioğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>36</span></span></td><td class="bolum-adi-cell">Halk Sağlığı</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Nadiye Pınar Ay</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>37</span></span></td><td class="bolum-adi-cell">Hematoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Işık Atagündüz</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>38</span></span></td><td class="bolum-adi-cell">Histoloji-Embriyoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Feriha Ercan</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>39</span></span></td><td class="bolum-adi-cell">İç Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Mehmet Koç</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>40</span></span></td><td class="bolum-adi-cell">Jinekolojik Onkoloji Cerrahisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Tanju Pekin</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>41</span></span></td><td class="bolum-adi-cell">Kadın Hastalıkları ve Doğum</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Suna Kabil Kucur</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>42</span></span></td><td class="bolum-adi-cell">Kalp ve Damar Cerrahisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Sinan Arsan</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>43</span></span></td><td class="bolum-adi-cell">Kardiyoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ahmet Altuğ Çinçin</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>44</span></span></td><td class="bolum-adi-cell">Kulak, Burun ve Boğaz Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Çağatay Oysu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>45</span></span></td><td class="bolum-adi-cell">Merkez Laboratuvar - Biyokimya</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Önder Şirikçi</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>46</span></span></td><td class="bolum-adi-cell">Nefroloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Zübeyde Serhan Tuğlular</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>47</span></span></td><td class="bolum-adi-cell">Neonatoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Hülya Selva Bilgen</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>48</span></span></td><td class="bolum-adi-cell">Nörofizyoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> İpek Midi</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>49</span></span></td><td class="bolum-adi-cell">Nöroloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> İpek Midi</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>50</span></span></td><td class="bolum-adi-cell">Nükleer Tıp</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Tanju Yusuf Erdil</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>51</span></span></td><td class="bolum-adi-cell">Ortopedi ve Travmatoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Özgür Baysal</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>52</span></span></td><td class="bolum-adi-cell">Perinatoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Münip Akalın</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>53</span></span></td><td class="bolum-adi-cell">Plastik ve Rekonstrüktif Cerrahi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Doç. Dr.</span> Zeynep Deniz Akdeniz Doğan</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>54</span></span></td><td class="bolum-adi-cell">Radyasyon Onkolojisi</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Zerrin Özgen</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>55</span></span></td><td class="bolum-adi-cell">Radyoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Nuri Çagatay Çimşit</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>56</span></span></td><td class="bolum-adi-cell">Romatoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Güzide Nevsun İnanç</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>57</span></span></td><td class="bolum-adi-cell">Ruh Sağlığı ve Hastalıkları</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Mesut Yıldız</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>58</span></span></td><td class="bolum-adi-cell">Tıbbi Farmakoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Hasan Raci Yananlı</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>59</span></span></td><td class="bolum-adi-cell">Tıbbi Genetik</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Ahmet Arman</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>60</span></span></td><td class="bolum-adi-cell">Tıbbi Mikrobiyoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Zeynep Arzu İlki</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>61</span></span></td><td class="bolum-adi-cell">Tıbbi Onkoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> İbrahim Vedat Bayoğlu</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>62</span></span></td><td class="bolum-adi-cell">Tıbbi Patoloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Hüseyin Kemal Türköz</td></tr>
<tr><td class="sira-no-cell"><span class="sira-rozet"><span>63</span></span></td><td class="bolum-adi-cell">Üroloji</td><td class="sorumlu-adi-cell"><span class="sorumlu-unvan">Prof. Dr.</span> Haydar Kamil Çam</td></tr>
</tbody>
</table>
</div>
<div id="noMatchesWrap" style="padding:0 18px 14px;">
<div id="noMatchesRow" style="display:none;">
<div class="no-result-icon" aria-hidden="true"></div>
<div class="no-result-title">"Aradığınız Bölüm Bulunamadı"</div>
<div class="no-result-divider" aria-hidden="true"></div>
<div class="no-result-text">
Arama Kriterlerinize Uygun Bir Bölüm Veya Sorumlu Bulunamadı.<br>
Farklı Bir Kelime Veya İfade İle Tekrar Deneyebilirsiniz.
</div>
</div>
</div>
</td>
</tr>
</table>
</div>
<script>
function normalize(txt){
  return txt.toLocaleLowerCase('tr-TR').normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^\w\s]/g,'');
}
function debounce(fn,delay){
  var t;
  return function(){
    clearTimeout(t);
    var args=arguments,ctx=this;
    t=setTimeout(function(){fn.apply(ctx,args)},delay);
  };
}
function getGroupSearchText(group){
  return normalize(group.innerText?' '+group.innerText:'');
}
function searchAnnouncements(){
  var raw=document.getElementById('search-bar').value.trim();
  var tokens=normalize(raw).split(/\s+/).filter(Boolean);
  var groups=document.querySelectorAll('#tabloGovde tr');
  var noRow=document.getElementById('noMatchesRow');
  var kutu=document.querySelector('#Table_01 .tablo-container');
  if(tokens.length===0){
    groups.forEach(function(g){g.classList.remove('search-hidden')});
    if(noRow)noRow.style.display='none';
    if(kutu)kutu.style.display='';
    return;
  }
  var matchFound=false;
  groups.forEach(function(group){
    var combo=getGroupSearchText(group);
    var hit=tokens.every(function(tok){return combo.indexOf(tok)!==-1});
    if(hit){group.classList.remove('search-hidden');matchFound=true}
    else group.classList.add('search-hidden');
  });
  if(noRow)noRow.style.display=matchFound?'none':'flex';
  if(kutu)kutu.style.display=matchFound?'':'none';
}
function updateClearButtonVisibility(){
  var searchBar=document.getElementById('search-bar');
  var clearBtn=document.getElementById('clearSearchBtn');
  if(!searchBar||!clearBtn)return;
  if(searchBar.value.trim().length>0){
    clearBtn.classList.add('has-text');
    clearBtn.title="Temizle";
    clearBtn.setAttribute("aria-label","Aramayı temizle");
  }else{
    clearBtn.classList.remove('has-text');
    clearBtn.title="Ara";
    clearBtn.setAttribute("aria-label","Ara");
  }
}
function updateSearchExpandState(){
  var searchBar=document.getElementById('search-bar');
  if(!searchBar)return;
  if(document.activeElement===searchBar)searchBar.classList.add('expanded');
  else searchBar.classList.remove('expanded');
}
document.addEventListener('DOMContentLoaded',function(){
  var searchBar=document.getElementById('search-bar');
  var clearBtn=document.getElementById('clearSearchBtn');
  if(searchBar){
    var runSearch=debounce(function(){
      searchAnnouncements();
      updateClearButtonVisibility();
      updateSearchExpandState();
    },200);
    searchBar.addEventListener('keyup',runSearch);
    searchBar.addEventListener('input',runSearch);
    searchBar.addEventListener('focus',updateSearchExpandState);
    searchBar.addEventListener('blur',function(){setTimeout(updateSearchExpandState,10)});
  }
  if(clearBtn&&searchBar){
    clearBtn.addEventListener('click',function(){
      if(searchBar.value.trim().length>0){
        searchBar.value='';
        searchAnnouncements();
        updateClearButtonVisibility();
        searchBar.blur();
      }else{
        searchBar.focus();
      }
      updateSearchExpandState();
    });
  }
  updateClearButtonVisibility();
  updateSearchExpandState();
});
</script>
<script>
(function(){
  var raf=0;
  function updateTableHeight(){
    raf=0;
    var menu=document.getElementById('cssmenu');
    var table=document.querySelector('#Table_01 .tablo-container');
    if(!menu||!table)return;
    var menuCell=menu.closest('td');
    var contentCell=document.querySelector('#Table_01 .baslik').closest('td');
    if(!menuCell||!contentCell)return;
    var menuHeight=Math.ceil(menu.getBoundingClientRect().height);
    if(menuHeight<1)return;
    var wrap=document.getElementById('noMatchesWrap');
    var wrapHeight=wrap?Math.ceil(wrap.getBoundingClientRect().height):0;
    var contentTop=contentCell.getBoundingClientRect().top;
    var tableTop=table.getBoundingClientRect().top;
    var bottomMargin=parseFloat(getComputedStyle(table).marginBottom)||0;
    var available=Math.max(200,Math.floor(menuHeight-(tableTop-contentTop)-bottomMargin-wrapHeight));
    if(table.style.height!==available+'px')table.style.height=available+'px';
    if(table.style.maxHeight!=='none')table.style.maxHeight='none';
  }
  function requestUpdate(){
    if(raf)return;
    raf=requestAnimationFrame(updateTableHeight);
  }
  function init(){
    var menu=document.getElementById('cssmenu');
    if(!menu)return;
    requestUpdate();
    if(window.ResizeObserver)new ResizeObserver(requestUpdate).observe(menu);
    menu.addEventListener('click',requestUpdate);
    menu.addEventListener('input',requestUpdate);
    menu.addEventListener('transitionend',requestUpdate,true);
    window.addEventListener('resize',requestUpdate,{passive:true});
    window.addEventListener('load',requestUpdate);
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);
  else init();
})();
</script>
</body>
</html>
