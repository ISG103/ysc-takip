<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ECE TRAFO - İSG Yönetim Bilgi Sistemi</title>
  <!-- SheetJS & FontAwesome -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  
  <style>
    :root {
      --sidebar: #0f172a;
      --brand-blue: #0284c7;
      --bg: #f3f4f6;
      --border: #cbd5e1;
      --text: #0f172a;
      --text-muted: #64748b;
      --danger: #dc2626;
      --warning: #d97706;
      --success: #16a34a;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', system-ui, sans-serif; }
    body { background-color: var(--bg); color: var(--text); display: flex; height: 100vh; overflow: hidden; }
    aside { width: 290px; background: var(--sidebar); color: white; display: flex; flex-direction: column; flex-shrink: 0; }
    .brand-box { padding: 16px; background: #ffffff; margin: 12px 14px; border-radius: 8px; text-align: center; }
    .brand-logo-img { max-height: 50px; width: auto; object-fit: contain; }
    .brand-portal-title { font-size: 0.72rem; font-weight: 800; color: #0f172a; margin-top: 4px; border-top: 1px solid #e2e8f0; padding-top: 4px; }
    .nav-list { list-style: none; padding: 6px 0; flex: 1; overflow-y: auto; }
    .nav-category { font-size: 0.65rem; color: #64748b; font-weight: 800; padding: 10px 20px 4px; text-transform: uppercase; }
    .nav-item { padding: 10px 20px; display: flex; align-items: center; gap: 12px; color: #94a3b8; cursor: pointer; font-size: 0.86rem; transition: 0.2s; }
    .nav-item:hover, .nav-item.active { background-color: rgba(255,255,255,0.06); color: #ffffff; border-left: 4px solid var(--brand-blue); }
    .nav-item i { width: 18px; text-align: center; }
    main { flex: 1; display: flex; flex-direction: column; overflow-y: auto; }
    header { background: white; padding: 14px 28px; border-bottom: 1px solid var(--border); display: flex; justify-content: space-between; align-items: center; }
    .plant-badge { background: #fef3c7; border: 1px solid #fde68a; color: #b45309; padding: 4px 12px; border-radius: 20px; font-size: 0.78rem; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; }
    .cloud-badge { padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; display: inline-flex; align-items: center; gap: 6px; background: #e0f2fe; color: #0369a1; }
    .content { padding: 24px 28px; flex: 1; }
    .metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; margin-bottom: 22px; }
    .metric-card { background: white; padding: 16px; border-radius: 8px; border: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center; }
    .metric-num { font-size: 1.55rem; font-weight: 800; margin-top: 2px; }
    .metric-title { font-size: 0.7rem; font-weight: 700; color: var(--text-muted); text-transform: uppercase; }
    .table-actions { display: flex; justify-content: space-between; align-items: center; margin-bottom: 14px; gap: 12px; }
    .search-box { position: relative; width: 320px; }
    .search-box input { width: 100%; padding: 8px 12px 8px 34px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; outline: none; }
    .search-box i { position: absolute; left: 10px; top: 50%; transform: translateY(-50%); color: var(--text-muted); }
    .btn { padding: 8px 14px; border-radius: 6px; font-size: 0.88rem; font-weight: 600; cursor: pointer; border: none; display: inline-flex; align-items: center; gap: 8px; }
    .btn-primary { background: var(--brand-blue); color: white; }
    .btn-excel { background: #107c41; color: white; }
    .btn-edit { background: #0284c7; color: white; padding: 5px 8px; font-size: 0.78rem; border-radius: 4px; }
    .btn-danger { background: var(--danger); color: white; padding: 5px 8px; font-size: 0.78rem; border-radius: 4px; }
    .table-container { background: white; border-radius: 8px; border: 1px solid var(--border); overflow-x: auto; }
    table { width: 100%; border-collapse: collapse; text-align: left; font-size: 0.86rem; }
    th { background: #f8fafc; padding: 12px 14px; color: #334155; font-weight: 700; border-bottom: 1px solid var(--border); }
    td { padding: 12px 14px; border-bottom: 1px solid #f1f5f9; }
    .badge { display: inline-block; padding: 3px 8px; border-radius: 999px; font-size: 0.72rem; font-weight: 700; }
    .badge-success { background: #dcfce7; color: var(--success); }
    .badge-warning { background: #fef3c7; color: var(--warning); }
    .badge-danger { background: #fee2e2; color: var(--danger); }
    .modal { position: fixed; inset: 0; background: rgba(15, 23, 42, 0.6); display: flex; align-items: center; justify-content: center; z-index: 1000; opacity: 0; visibility: hidden; transition: all 0.2s; }
    .modal.active { opacity: 1; visibility: visible; }
    .modal-content { background: white; width: 540px; border-radius: 8px; padding: 24px; max-height: 90vh; overflow-y: auto; }
    .modal-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px; border-bottom: 1px solid #f1f5f9; padding-bottom: 10px; }
    .form-group { margin-bottom: 12px; }
    .form-group label { display: block; font-size: 0.82rem; font-weight: 600; margin-bottom: 4px; color: #475569; }
    .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; font-size: 0.88rem; }
    .modal-footer { display: flex; justify-content: flex-end; gap: 8px; margin-top: 20px; }
    .hidden { display: none !important; }

    .class-card { border: 2px solid var(--border); border-radius: 8px; padding: 16px; margin-bottom: 14px; cursor: pointer; transition: all 0.2s; background: white; display: flex; justify-content: space-between; align-items: center; }
    .class-card:hover { border-color: var(--brand-blue); background: #f8fafc; }
    .class-card.selected { border-color: var(--brand-blue); background: #eff6ff; }
  </style>
</head>
<body>
  <aside>
    <div class="brand-box">
      <img src="ece_logo.png" alt="ECE TRAFO" class="brand-logo-img" onerror="this.style.display='none'; document.getElementById('logo-fallback').style.display='block';">
      <div id="logo-fallback" style="display:none; font-size: 1.3rem; font-weight:900; color:#0f172a;">ECE TRAFO</div>
      <div class="brand-portal-title">İSG YÖNETİM PORTALI</div>
    </div>
    <ul class="nav-list">
      <div class="nav-category">PERSONEL & SAĞLIK</div>
      <li class="nav-item active" onclick="switchNav('personel', this)"><i class="fa-solid fa-users"></i> Personel & Muayene Takibi</li>
      <li class="nav-item" onclick="switchNav('sertifika', this)"><i class="fa-solid fa-award"></i> İlk Yardım & MYK Belgeleri</li>
      <li class="nav-item" onclick="switchNav('kkd', this)"><i class="fa-solid fa-vest"></i> KKD Zimmet Takibi</li>
      
      <div class="nav-category">SAHA GÜVENLİĞİ</div>
      <li class="nav-item" onclick="switchNav('dof', this)"><i class="fa-solid fa-clipboard-check"></i> Saha Denetimi & DÖF</li>
      <li class="nav-item" onclick="switchNav('izin', this)"><i class="fa-solid fa-file-signature"></i> Özel İş İzinleri (PTW)</li>
      <li class="nav-item" onclick="switchNav('kazalar', this)"><i class="fa-solid fa-triangle-exclamation"></i> İş Kazası Defteri</li>

      <div class="nav-category">YASAL AYARLAR & MEVZUAT</div>
      <li class="nav-item" onclick="switchNav('siniflar', this)" style="color:#38bdf8; font-weight:700;"><i class="fa-solid fa-sliders"></i> İSG Sınıf Ayarları</li>
      <li class="nav-item" onclick="switchNav('istatistik', this)"><i class="fa-solid fa-calculator"></i> SGK Kaza Oranları Motoru</li>
    </ul>
  </aside>

  <main>
    <header>
      <div>
        <h2 style="font-size:1.15rem; font-weight:800;" id="pageTitle">Personel Eğitim & Sağlık İzleme</h2>
        <div style="display:flex; gap:8px; margin-top:4px;">
          <div class="plant-badge" id="headerBadge"><i class="fa-solid fa-shield-halved"></i> Aktif Sınıf: TEHLİKELİ (Eğitim: 2 Yıl / Muayene: 3 Yıl)</div>
          <div class="cloud-badge" id="cloudStatus"><i class="fa-solid fa-cloud-arrow-up"></i> Bulut Hazır</div>
        </div>
      </div>
      <button class="btn btn-excel" onclick="excelIndir()"><i class="fa-solid fa-file-excel"></i> Excel Çıktısı (.xlsx)</button>
    </header>

    <div class="content">
      <div class="metric-grid">
        <div class="metric-card"><div><div class="metric-title">Toplam Personel</div><div class="metric-num" id="stat_toplam_p">0</div></div><i class="fa-solid fa-users" style="font-size:1.5rem; color:var(--brand-blue);"></i></div>
        <div class="metric-card"><div><div class="metric-title">Yenileme Alarmı</div><div class="metric-num" id="stat_alarm" style="color:var(--danger)">0</div></div><i class="fa-solid fa-bell" style="font-size:1.5rem; color:var(--danger);"></i></div>
        <div class="metric-card"><div><div class="metric-title">Açık DÖF</div><div class="metric-num" id="stat_acik_dof" style="color:var(--warning)">0</div></div><i class="fa-solid fa-clipboard-check" style="font-size:1.5rem; color:var(--warning);"></i></div>
        <div class="metric-card"><div><div class="metric-title">Kaza Sayısı</div><div class="metric-num" id="stat_kaza">0</div></div><i class="fa-solid fa-bandage" style="font-size:1.5rem; color:#475569;"></i></div>
      </div>

      <!-- 1. PERSONEL -->
      <section id="sec-personel">
        <div class="table-actions">
          <div class="search-box"><i class="fa-solid fa-magnifying-glass"></i><input type="text" id="p_ara" placeholder="Personel, Görev veya Bölüm Ara..." oninput="personelCiz()"></div>
          <button class="btn btn-primary" onclick="openPersonelModal()"><i class="fa-solid fa-plus"></i> Yeni Personel Ekle</button>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Ad Soyad</th>
                <th>Bölüm / İstasyon</th>
                <th>Görevi / Pozisyon</th>
                <th>İşe Başlama</th>
                <th>İSG Eğitimi Bitiş</th>
                <th>Periyodik Muayene Bitiş</th>
                <th style="min-width:90px; text-align:center;">İşlemler</th>
              </tr>
            </thead>
            <tbody id="personelTablo"></tbody>
          </table>
        </div>
      </section>

      <!-- 2. İSG SINIFLARI SEKMESİ -->
      <section id="sec-siniflar" class="hidden">
        <div style="background:white; border-radius:8px; border:1px solid var(--border); padding:24px; max-width:700px;">
          <h3 style="margin-bottom:6px; font-size:1.15rem; color:#0f172a;"><i class="fa-solid fa-sliders" style="color:var(--brand-blue);"></i> Fabrika İSG Tehlike Sınıfı Seçimi</h3>
          <p style="font-size:0.85rem; color:#64748b; margin-bottom:18px;">Buradan fabrikanız için geçerli olan sınıfı seçin. Yeni personel kaydında eğitim ve sağlık süreleri otomatik olarak buna göre hesaplanır.</p>

          <div class="class-card selected" id="card_tehlikeli" onclick="sinifSec('tehlikeli')">
            <div>
              <div style="font-weight:700; color:#b45309; font-size:1rem;"><i class="fa-solid fa-triangle-exclamation"></i> TEHLİKELİ SINIF (Varsayılan)</div>
              <div style="font-size:0.84rem; color:#475569; margin-top:4px;">İSG Eğitimi: <strong>2 Yılda 1 (12 Saat)</strong> | Periyodik Sağlık: <strong>3 Yılda 1</strong></div>
            </div>
            <i class="fa-solid fa-circle-check" id="check_tehlikeli" style="color:var(--brand-blue); font-size:1.3rem;"></i>
          </div>

          <div class="class-card" id="card_cok_tehlikeli" onclick="sinifSec('cok_tehlikeli')">
            <div>
              <div style="font-weight:700; color:#dc2626; font-size:1rem;"><i class="fa-solid fa-radiation"></i> ÇOK TEHLİKELİ SINIF</div>
              <div style="font-size:0.84rem; color:#475569; margin-top:4px;">İSG Eğitimi: <strong>1 Yılda 1 (16 Saat)</strong> | Periyodik Sağlık: <strong>1 Yılda 1</strong></div>
            </div>
            <i class="fa-regular fa-circle" id="check_cok_tehlikeli" style="color:#cbd5e1; font-size:1.3rem;"></i>
          </div>

          <div class="class-card" id="card_az_tehlikeli" onclick="sinifSec('az_tehlikeli')">
            <div>
              <div style="font-weight:700; color:#16a34a; font-size:1rem;"><i class="fa-solid fa-circle-check"></i> AZ TEHLİKELİ SINIF</div>
              <div style="font-size:0.84rem; color:#475569; margin-top:4px;">İSG Eğitimi: <strong>3 Yılda 1 (8 Saat)</strong> | Periyodik Sağlık: <strong>5 Yılda 1</strong></div>
            </div>
            <i class="fa-regular fa-circle" id="check_az_tehlikeli" style="color:#cbd5e1; font-size:1.3rem;"></i>
          </div>

          <div class="class-card" id="card_ozel" onclick="sinifSec('ozel')">
            <div style="width:100%;">
              <div style="font-weight:700; color:#0284c7; font-size:1rem;"><i class="fa-solid fa-pen-ruler"></i> ÖZEL / FABRİKAYA ÖZEL PERİYOT</div>
              <div style="font-size:0.84rem; color:#475569; margin-top:4px;">Süreleri kendiniz belirleyin:</div>
              <div style="display:flex; gap:12px; margin-top:8px;">
                <div style="flex:1;">
                  <label style="font-size:0.75rem; font-weight:600;">Eğitim Kaç Yılda Bir?</label>
                  <input type="number" id="ozel_egitim_yil" value="2" min="1" max="10" style="padding:6px; border:1px solid #cbd5e1; border-radius:4px; width:100%;">
                </div>
                <div style="flex:1;">
                  <label style="font-size:0.75rem; font-weight:600;">Muayene Kaç Yılda Bir?</label>
                  <input type="number" id="ozel_saglik_yil" value="3" min="1" max="10" style="padding:6px; border:1px solid #cbd5e1; border-radius:4px; width:100%;">
                </div>
              </div>
            </div>
            <i class="fa-regular fa-circle" id="check_ozel" style="color:#cbd5e1; font-size:1.3rem; margin-left:12px;"></i>
          </div>

          <button class="btn btn-primary" style="margin-top:14px; width:100%; justify-content:center;" onclick="sinifAyariKaydet()">
            <i class="fa-solid fa-floppy-disk"></i> Seçilen Sınıfı Aktif Yap ve Kaydet
          </button>
        </div>
      </section>

      <!-- 3. SERTİFİKA -->
      <section id="sec-sertifika" class="hidden">
        <div class="table-actions">
          <button class="btn btn-primary" onclick="openModal('modal-sertifika')"><i class="fa-solid fa-plus"></i> Belge Ekle</button>
        </div>
        <div class="table-container">
          <table>
            <thead><tr><th>Personel</th><th>Belge Türü</th><th>Kurum / No</th><th>Vize Tarihi</th><th>Durum</th><th>İşlem</th></tr></thead>
            <tbody id="sertifikaTablo"></tbody>
          </table>
        </div>
      </section>

      <!-- 4. KKD -->
      <section id="sec-kkd" class="hidden">
        <div class="table-actions">
          <button class="btn btn-primary" onclick="openModal('modal-kkd')"><i class="fa-solid fa-plus"></i> KKD Zimmetle</button>
        </div>
        <div class="table-container">
          <table>
            <thead><tr><th>Personel</th><th>Donanım</th><th>Veriliş</th><th>Değişim</th><th>Durum</th><th>İşlem</th></tr></thead>
            <tbody id="kkdTablo"></tbody>
          </table>
        </div>
      </section>

      <!-- 5. DÖF -->
      <section id="sec-dof" class="hidden">
        <div class="table-actions">
          <button class="btn btn-primary" onclick="openModal('modal-dof')"><i class="fa-solid fa-plus"></i> DÖF Aç</button>
        </div>
        <div class="table-container">
          <table>
            <thead><tr><th>Tarih</th><th>Bölüm</th><th>Uygunsuzluk</th><th>Aksiyon</th><th>Sorumlu</th><th>Statü</th><th>İşlem</th></tr></thead>
            <tbody id="dofTablo"></tbody>
          </table>
        </div>
      </section>

      <!-- 6. İŞ İZNİ -->
      <section id="sec-izin" class="hidden">
        <div class="table-actions">
          <button class="btn btn-primary" onclick="openModal('modal-izin')"><i class="fa-solid fa-plus"></i> İş İzni Aç</button>
        </div>
        <div class="table-container">
          <table>
            <thead><tr><th>İzin No</th><th>Tür</th><th>Alan</th><th>Personel</th><th>Statü</th><th>İşlem</th></tr></thead>
            <tbody id="izinTablo"></tbody>
          </table>
        </div>
      </section>

      <!-- 7. KAZALAR -->
      <section id="sec-kazalar" class="hidden">
        <div class="table-actions">
          <button class="btn btn-primary" onclick="openModal('modal-kaza')"><i class="fa-solid fa-plus"></i> Kaza Bildir</button>
        </div>
        <div class="table-container">
          <table>
            <thead><tr><th>Tarih</th><th>Kazazede</th><th>Tür</th><th>Kayıp Gün</th><th>İşlem</th></tr></thead>
            <tbody id="kazaTablo"></tbody>
          </table>
        </div>
      </section>

      <!-- 8. YASAL İSTATİSTİK -->
      <section id="sec-istatistik" class="hidden">
        <div style="background:white; border-radius:8px; border:1px solid var(--border); padding:24px; max-width:650px;">
          <h3 style="margin-bottom:12px; font-size:1.1rem; color:#0f172a;"><i class="fa-solid fa-calculator" style="color:var(--brand-blue);"></i> SGK / Bakanlık Resmi Kaza Oranları</h3>
          <div class="form-group">
            <label>Dönemdeki Fiili Çalışma Saati (Adam x Saat)</label>
            <input type="number" id="stat_fiili_saat" value="100000" oninput="hesaplaIstatistik()">
          </div>
          <div style="background:#f8fafc; padding:18px; border-radius:8px; margin-top:15px; border:1px solid #e2e8f0;">
            <div style="margin-bottom:14px;">
              <strong>Kaza Sıklık Oranı (KSO): </strong>
              <span id="res_kso" style="font-size:1.2rem; font-weight:800; color:var(--brand-blue);">0.00</span>
              <div style="font-size:0.75rem; color:#64748b;">(SGK Formülü: Toplam Kaza Sayısı × 1.000.000 / Fiili Çalışma Saati)</div>
            </div>
            <div>
              <strong>Kaza Ağırlık Oranı (KAO): </strong>
              <span id="res_kao" style="font-size:1.2rem; font-weight:800; color:var(--danger);">0.00</span>
              <div style="font-size:0.75rem; color:#64748b;">(SGK Formülü: Toplam Kayıp Gün × 1.000 / Fiili Çalışma Saati)</div>
            </div>
          </div>
        </div>
      </section>
    </div>
  </main>

  <!-- MODAL: PERSONEL (EKLEME & DÜZENLEME) -->
  <div class="modal" id="modal-personel">
    <div class="modal-content">
      <div class="modal-header">
        <h3 id="modalPersonelBaslik">Yeni Personel Kaydı</h3>
        <i class="fa-solid fa-xmark" style="cursor:pointer;" onclick="closeModal('modal-personel')"></i>
      </div>
      <!-- Düzenlenen Personelin ID'sini gizli tutar -->
      <input type="hidden" id="p_edit_id">

      <div class="form-group"><label>Adı Soyadı</label><input type="text" id="p_ad" placeholder="Ad Soyad giriniz"></div>
      <div class="form-group">
        <label>Fabrika Bölümü / İstasyon</label>
        <select id="p_bolum">
          <option>Kazan Kaynak & Montaj Holü</option>
          <option>Kumlama & Boyahane İstasyonu</option>
          <option>Mekanik İşleme & CNC Atölyesi</option>
          <option>Lazer & Giyotin Saç Kesim</option>
          <option>Elektrik & Mekanik Bakım</option>
          <option>Açık Saha & Sevkiyat</option>
          <option>İdari İşler & Kalite</option>
        </select>
      </div>
      <div class="form-group"><label>Görevi / Pozisyonu</label><input type="text" id="p_gorev" placeholder="Örn: Kaynakçı, Vinç Operatörü, Formen"></div>
      <div class="form-group"><label>İşe Başlama Tarihi</label><input type="date" id="p_ise_baslama"></div>
      
      <div style="background:#f8fafc; border:1px solid #e2e8f0; padding:12px; border-radius:6px; margin-bottom:12px;">
        <div class="form-group">
          <label>Son Yapılan İSG Eğitimi Tarihi</label>
          <input type="date" id="p_egitim_yapilis" onchange="otomatikTarihHesapla()">
        </div>
        <div class="form-group" style="margin-bottom:0;">
          <label>İSG Eğitimi Bitiş (Yenileme) Tarihi</label>
          <input type="date" id="p_egitim">
        </div>
      </div>

      <div style="background:#f8fafc; border:1px solid #e2e8f0; padding:12px; border-radius:6px; margin-bottom:12px;">
        <div class="form-group">
          <label>Son Sağlık Muayenesi Tarihi</label>
          <input type="date" id="p_saglik_yapilis" onchange="otomatikTarihHesapla()">
        </div>
        <div class="form-group" style="margin-bottom:0;">
          <label>Sağlık Raporu Bitiş (Yenileme) Tarihi</label>
          <input type="date" id="p_saglik">
        </div>
      </div>

      <div class="modal-footer">
        <button class="btn btn-primary" id="btnPersonelKaydet" onclick="personelKaydet()">Kaydet</button>
      </div>
    </div>
  </div>

  <!-- Diğer Modallar -->
  <div class="modal" id="modal-sertifika">
    <div class="modal-content">
      <div class="modal-header"><h3>Belge Ekle</h3><i class="fa-solid fa-xmark" style="cursor:pointer;" onclick="closeModal('modal-sertifika')"></i></div>
      <div class="form-group"><label>Personel</label><input type="text" id="s_ad"></div>
      <div class="form-group"><label>Belge Türü</label><select id="s_tur"><option>İlk Yardımcı Sertifikası (3 Yıl Vize)</option><option>MYK Çelik Kaynakçısı</option><option>Tavan Vinci Operatörü</option><option>Forklift Operatörü</option><option>Kazan Operatörü</option></select></div>
      <div class="form-group"><label>Kurum / No</label><input type="text" id="s_kurum"></div>
      <div class="form-group"><label>Vize Bitiş</label><input type="date" id="s_bitis"></div>
      <div class="modal-footer"><button class="btn btn-primary" onclick="sertifikaEkle()">Kaydet</button></div>
    </div>
  </div>

  <div class="modal" id="modal-kkd">
    <div class="modal-content">
      <div class="modal-header"><h3>KKD Zimmet</h3><i class="fa-solid fa-xmark" style="cursor:pointer;" onclick="closeModal('modal-kkd')"></i></div>
      <div class="form-group"><label>Personel</label><input type="text" id="kkd_p"></div>
      <div class="form-group"><label>Donanım</label><input type="text" id="kkd_tur" placeholder="Örn: S3 İş Ayakkabısı"></div>
      <div class="form-group"><label>Veriliş Tarihi</label><input type="date" id="kkd_verilis"></div>
      <div class="form-group"><label>Değişim Periyodu (Ay)</label><select id="kkd_omur"><option value="6">6 Ay</option><option value="12" selected>12 Ay</option></select></div>
      <div class="modal-footer"><button class="btn btn-primary" onclick="kkdEkle()">Kaydet</button></div>
    </div>
  </div>

  <div class="modal" id="modal-dof">
    <div class="modal-content">
      <div class="modal-header"><h3>Saha DÖF Girişi</h3><i class="fa-solid fa-xmark" style="cursor:pointer;" onclick="closeModal('modal-dof')"></i></div>
      <div class="form-group"><label>Tarih</label><input type="date" id="dof_tarih"></div>
      <div class="form-group"><label>Bölüm</label><input type="text" id="dof_bolum"></div>
      <div class="form-group"><label>Tespit</label><textarea id="dof_tespit" rows="2"></textarea></div>
      <div class="form-group"><label>Aksiyon</label><textarea id="dof_aksiyon" rows="2"></textarea></div>
      <div class="form-group"><label>Sorumlu</label><input type="text" id="dof_sorumlu"></div>
      <div class="form-group"><label>Statü</label><select id="dof_statu"><option>Açık</option><option>Kapatıldı</option></select></div>
      <div class="modal-footer"><button class="btn btn-primary" onclick="dofEkle()">Kaydet</button></div>
    </div>
  </div>

  <div class="modal" id="modal-izin">
    <div class="modal-content">
      <div class="modal-header"><h3>İş İzni Aç</h3><i class="fa-solid fa-xmark" style="cursor:pointer;" onclick="closeModal('modal-izin')"></i></div>
      <div class="form-group"><label>İzin No</label><input type="text" id="ptw_no" placeholder="Örn: PTW-01"></div>
      <div class="form-group"><label>Tür</label><select id="ptw_tur"><option>Kazan İçi Kapalı Alan Girişi</option><option>Sıcak Çalışma (Kaynak/Alev)</option><option>Yüksekte Çalışma</option></select></div>
      <div class="form-group"><label>Alan</label><input type="text" id="ptw_alan"></div>
      <div class="form-group"><label>Personel</label><input type="text" id="ptw_personel"></div>
      <div class="modal-footer"><button class="btn btn-primary" onclick="izinEkle()">Kaydet</button></div>
    </div>
  </div>

  <div class="modal" id="modal-kaza">
    <div class="modal-content">
      <div class="modal-header"><h3>Kaza Kaydı</h3><i class="fa-solid fa-xmark" style="cursor:pointer;" onclick="closeModal('modal-kaza')"></i></div>
      <div class="form-group"><label>Tarih</label><input type="date" id="k_tarih"></div>
      <div class="form-group"><label>Kazazede</label><input type="text" id="k_ad"></div>
      <div class="form-group"><label>Tür</label><select id="k_tur"><option>Hafif Yaralanma</option><option>İş Göremezlik Raporlu Kaza</option><option>Ramak Kala</option></select></div>
      <div class="form-group"><label>Kayıp Gün</label><input type="number" id="k_gun" value="0"></div>
      <div class="modal-footer"><button class="btn btn-primary" onclick="kazaEkle()">Kaydet</button></div>
    </div>
  </div>

  <script>
    // BULUT VERİTABANI BAĞLANTISI (JSONBin.io)
    const BIN_ID = '6aba83d3ffd5d1605337f957';
    const MASTER_KEY = '$2a$10$AgSMqKfELFxqPVn8PEo3ie2lsSPAbvq/nawt/ATJSKDHJgJUjPs6u';
    const BIN_URL = `https://api.jsonbin.io/v3/b/${BIN_ID}`;

    let db = {
      p: JSON.parse(localStorage.getItem('isg_p')) || [],
      s: JSON.parse(localStorage.getItem('isg_s')) || [],
      dof: JSON.parse(localStorage.getItem('isg_dof')) || [],
      izin: JSON.parse(localStorage.getItem('isg_izin')) || [],
      k: JSON.parse(localStorage.getItem('isg_k')) || [],
      kkd: JSON.parse(localStorage.getItem('isg_kkd')) || [],
      ayar: JSON.parse(localStorage.getItem('isg_ayar')) || {
        sinif: 'tehlikeli',
        egitimYil: 2,
        saglikYil: 3,
        baslik: 'TEHLİKELİ (Eğitim: 2 Yıl / Muayene: 3 Yıl)'
      }
    };

    function updateCloudStatus(status, text) {
      const el = document.getElementById('cloudStatus');
      if (status === 'loading') {
        el.style.background = '#fef3c7'; el.style.color = '#b45309';
        el.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> ${text}`;
      } else if (status === 'success') {
        el.style.background = '#dcfce7'; el.style.color = '#15803d';
        el.innerHTML = `<i class="fa-solid fa-cloud-check"></i> ${text}`;
      } else {
        el.style.background = '#fee2e2'; el.style.color = '#b91c1c';
        el.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> ${text}`;
      }
    }

    async function buluttanYukle() {
      updateCloudStatus('loading', 'Buluttan Alınıyor...');
      try {
        const res = await fetch(BIN_URL + '/latest', {
          headers: { 'X-Master-Key': MASTER_KEY }
        });
        if (res.ok) {
          const data = await res.json();
          if (data.record && typeof data.record === 'object') {
            db.p = data.record.p || db.p;
            db.s = data.record.s || db.s;
            db.dof = data.record.dof || db.dof;
            db.izin = data.record.izin || db.izin;
            db.k = data.record.k || db.k;
            db.kkd = data.record.kkd || db.kkd;
            if (data.record.ayar) db.ayar = data.record.ayar;
            
            for(let key in db) { localStorage.setItem('isg_' + key, JSON.stringify(db[key])); }
            updateCloudStatus('success', 'Bulut Canlı');
            cizimlerinHepsiniYenile();
          }
        } else {
          updateCloudStatus('error', 'Bulut Okunamadı');
        }
      } catch (err) {
        console.error(err);
        updateCloudStatus('success', 'Yerel Aktif (Bulut Beklemede)');
      }
    }

    async function sync() {
      for(let key in db) { localStorage.setItem('isg_' + key, JSON.stringify(db[key])); }
      updateMetrics();
      hesaplaIstatistik();

      updateCloudStatus('loading', 'Buluta Yazılıyor...');
      try {
        const res = await fetch(BIN_URL, {
          method: 'PUT',
          headers: {
            'Content-Type': 'application/json',
            'X-Master-Key': MASTER_KEY
          },
          body: JSON.stringify(db)
        });
        if (res.ok) {
          updateCloudStatus('success', 'Bulut Senkronize');
        } else {
          updateCloudStatus('error', 'Buluta Yazılamadı');
        }
      } catch (e) {
        console.error(e);
        updateCloudStatus('error', 'İnternet Bağlantısı Yok');
      }
    }

    function openModal(id) { document.getElementById(id).classList.add('active'); }
    function closeModal(id) { document.getElementById(id).classList.remove('active'); }

    function switchNav(target, elem) {
      document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
      elem.classList.add('active');
      ['personel', 'siniflar', 'sertifika', 'kkd', 'dof', 'izin', 'kazalar', 'istatistik'].forEach(s => {
        const el = document.getElementById('sec-' + s);
        if(el) el.classList.add('hidden');
      });
      document.getElementById('sec-' + target).classList.remove('hidden');

      const titles = {
        personel: 'Personel Eğitim & Sağlık İzleme',
        siniflar: 'İSG Tehlike Sınıfı ve Periyot Ayarları',
        sertifika: 'İlk Yardımcı, MYK ve Sertifika Yönetimi',
        kkd: 'KKD Zimmet Takibi',
        dof: 'DÖF & Saha Denetimi',
        izin: 'Özel İş İzinleri (PTW)',
        kazalar: 'İş Kazaları Defteri',
        istatistik: 'SGK Kaza Oranları Motoru'
      };
      document.getElementById('pageTitle').innerText = titles[target] || 'İSG Sistemi';
    }

    let geciciSecim = db.ayar.sinif;
    function sinifSec(sinifKey) {
      geciciSecim = sinifKey;
      ['tehlikeli', 'cok_tehlikeli', 'az_tehlikeli', 'ozel'].forEach(k => {
        document.getElementById('card_' + k).classList.remove('selected');
        document.getElementById('check_' + k).className = 'fa-regular fa-circle';
        document.getElementById('check_' + k).style.color = '#cbd5e1';
      });
      document.getElementById('card_' + sinifKey).classList.add('selected');
      document.getElementById('check_' + sinifKey).className = 'fa-solid fa-circle-check';
      document.getElementById('check_' + sinifKey).style.color = '#0284c7';
    }

    function sinifAyariKaydet() {
      if(geciciSecim === 'tehlikeli') {
        db.ayar = { sinif: 'tehlikeli', egitimYil: 2, saglikYil: 3, baslik: 'TEHLİKELİ (Eğitim: 2 Yıl / Muayene: 3 Yıl)' };
      } else if(geciciSecim === 'cok_tehlikeli') {
        db.ayar = { sinif: 'cok_tehlikeli', egitimYil: 1, saglikYil: 1, baslik: 'ÇOK TEHLİKELİ (Eğitim: 1 Yıl / Muayene: 1 Yıl)' };
      } else if(geciciSecim === 'az_tehlikeli') {
        db.ayar = { sinif: 'az_tehlikeli', egitimYil: 3, saglikYil: 5, baslik: 'AZ TEHLİKELİ (Eğitim: 3 Yıl / Muayene: 5 Yıl)' };
      } else if(geciciSecim === 'ozel') {
        const ey = parseInt(document.getElementById('ozel_egitim_yil').value) || 1;
        const sy = parseInt(document.getElementById('ozel_saglik_yil').value) || 1;
        db.ayar = { sinif: 'ozel', egitimYil: ey, saglikYil: sy, baslik: `ÖZEL AYAR (Eğitim: ${ey} Yıl / Muayene: ${sy} Yıl)` };
      }
      sync();
      updateHeaderBadge();
      alert(`Fabrika aktif sınıfı buluta kaydedildi:\n${db.ayar.baslik}`);
    }

    function updateHeaderBadge() {
      document.getElementById('headerBadge').innerHTML = `<i class="fa-solid fa-shield-halved"></i> Aktif Sınıf: ${db.ayar.baslik}`;
    }

    function otomatikTarihHesapla() {
      const egitimYapilis = document.getElementById('p_egitim_yapilis').value;
      const saglikYapilis = document.getElementById('p_saglik_yapilis').value;
      if(egitimYapilis) {
        let d = new Date(egitimYapilis);
        d.setFullYear(d.getFullYear() + db.ayar.egitimYil);
        document.getElementById('p_egitim').value = d.toISOString().slice(0, 10);
      }
      if(saglikYapilis) {
        let d = new Date(saglikYapilis);
        d.setFullYear(d.getFullYear() + db.ayar.saglikYil);
        document.getElementById('p_saglik').value = d.toISOString().slice(0, 10);
      }
    }

    function getStatus(dateStr) {
      if(!dateStr) return { text: "Tarih Yok", cls: "badge-danger", diff: -999 };
      const target = new Date(dateStr);
      const today = new Date();
      today.setHours(0,0,0,0);
      const diff = Math.ceil((target - today) / (1000 * 60 * 60 * 24));
      if (diff < 0) return { text: `Süresi Doldu (${Math.abs(diff)} gün)`, cls: "badge-danger", diff };
      if (diff <= 60) return { text: `Yaklaştı (${diff} gün)`, cls: "badge-warning", diff };
      return { text: `Geçerli (${diff} gün)`, cls: "badge-success", diff };
    }

    function updateMetrics() {
      document.getElementById('stat_toplam_p').innerText = db.p.length;
      document.getElementById('stat_acik_dof').innerText = db.dof.filter(x => x.statu !== 'Kapatıldı').length;
      document.getElementById('stat_kaza').innerText = db.k.length;

      let alarm = 0;
      db.p.forEach(x => {
        if(getStatus(x.egitim).diff <= 60 || getStatus(x.saglik).diff <= 60) alarm++;
      });
      document.getElementById('stat_alarm').innerText = alarm;
    }

    function personelCiz() {
      const q = document.getElementById('p_ara').value.toLowerCase();
      const tbody = document.getElementById('personelTablo');
      tbody.innerHTML = '';
      db.p.filter(x => x.ad.toLowerCase().includes(q) || (x.bolum||'').toLowerCase().includes(q) || (x.gorev||'').toLowerCase().includes(q)).forEach(p => {
        const e = getStatus(p.egitim);
        const s = getStatus(p.saglik);
        tbody.innerHTML += `<tr>
          <td><strong>${p.ad}</strong></td>
          <td>${p.bolum||'-'}</td>
          <td>${p.gorev||'-'}</td>
          <td>${p.ise_baslama||'-'}</td>
          <td><span class="badge ${e.cls}">${e.text}</span> <div style="font-size:10px; color:#64748b;">${p.egitim||''}</div></td>
          <td><span class="badge ${s.cls}">${s.text}</span> <div style="font-size:10px; color:#64748b;">${p.saglik||''}</div></td>
          <td style="text-align:center; white-space:nowrap;">
            <button class="btn btn-edit" title="Düzenle" onclick="personelDuzenle(${p.id})"><i class="fa-solid fa-pen-to-square"></i></button>
            <button class="btn btn-danger" title="Sil" onclick="sil('p', ${p.id})"><i class="fa-solid fa-trash"></i></button>
          </td>
        </tr>`;
      });
    }

    // PERSONEL EKLEME VE DÜZENLEME FONKSİYONLARI
    function openPersonelModal() {
      document.getElementById('modalPersonelBaslik').innerText = "Yeni Personel Kaydı";
      document.getElementById('btnPersonelKaydet').innerText = "Kaydet";
      document.getElementById('p_edit_id').value = "";
      document.getElementById('p_ad').value = '';
      document.getElementById('p_gorev').value = '';
      document.getElementById('p_ise_baslama').value = '';
      document.getElementById('p_egitim_yapilis').value = '';
      document.getElementById('p_egitim').value = '';
      document.getElementById('p_saglik_yapilis').value = '';
      document.getElementById('p_saglik').value = '';
      openModal('modal-personel');
    }

    function personelDuzenle(id) {
      const p = db.p.find(x => x.id === id);
      if(!p) return;

      document.getElementById('modalPersonelBaslik').innerText = "Personel Bilgilerini Düzenle";
      document.getElementById('btnPersonelKaydet').innerText = "Değişiklikleri Güncelle";
      document.getElementById('p_edit_id').value = p.id;
      document.getElementById('p_ad').value = p.ad || '';
      document.getElementById('p_bolum').value = p.bolum || 'Kazan Kaynak & Montaj Holü';
      document.getElementById('p_gorev').value = p.gorev || '';
      document.getElementById('p_ise_baslama').value = p.ise_baslama || '';
      document.getElementById('p_egitim_yapilis').value = p.egitim_yapilis || '';
      document.getElementById('p_egitim').value = p.egitim || '';
      document.getElementById('p_saglik_yapilis').value = p.saglik_yapilis || '';
      document.getElementById('p_saglik').value = p.saglik || '';

      openModal('modal-personel');
    }

    function personelKaydet() {
      const editId = document.getElementById('p_edit_id').value;
      const ad = document.getElementById('p_ad').value.trim();
      const bolum = document.getElementById('p_bolum').value;
      const gorev = document.getElementById('p_gorev').value.trim();
      const ise_baslama = document.getElementById('p_ise_baslama').value;
      const egitim_yapilis = document.getElementById('p_egitim_yapilis').value;
      const egitim = document.getElementById('p_egitim').value;
      const saglik_yapilis = document.getElementById('p_saglik_yapilis').value;
      const saglik = document.getElementById('p_saglik').value;
      
      if(!ad) return alert("Personel adı boş bırakılamaz.");

      if(editId) {
        // MEVCUT PERSONELİ GÜNCELLE
        const index = db.p.findIndex(x => x.id == editId);
        if(index !== -1) {
          db.p[index] = {
            ...db.p[index],
            ad, bolum, gorev, ise_baslama, egitim_yapilis, egitim, saglik_yapilis, saglik
          };
        }
      } else {
        // YENİ PERSONEL EKLE
        db.p.push({
          id: Date.now(),
          ad, bolum, gorev, ise_baslama, egitim_yapilis, egitim, saglik_yapilis, saglik
        });
      }

      sync();
      personelCiz();
      closeModal('modal-personel');
    }

    function sertifikaCiz() {
      const tbody = document.getElementById('sertifikaTablo');
      tbody.innerHTML = '';
      db.s.forEach(s => {
        const d = getStatus(s.bitis);
        tbody.innerHTML += `<tr><td>${s.ad}</td><td>${s.tur}</td><td>${s.kurum||'-'}</td><td>${s.bitis||'-'}</td><td><span class="badge ${d.cls}">${d.text}</span></td><td><button class="btn btn-danger" onclick="sil('s', ${s.id})"><i class="fa-solid fa-trash"></i></button></td></tr>`;
      });
    }
    function sertifikaEkle() {
      const ad = document.getElementById('s_ad').value.trim();
      const tur = document.getElementById('s_tur').value;
      const kurum = document.getElementById('s_kurum').value.trim();
      const bitis = document.getElementById('s_bitis').value;
      if(!ad) return alert("Ad girin.");
      db.s.push({ id: Date.now(), ad, tur, kurum, bitis });
      sync(); sertifikaCiz(); closeModal('modal-sertifika');
    }

    function kkdCiz() {
      const tbody = document.getElementById('kkdTablo');
      tbody.innerHTML = '';
      db.kkd.forEach(k => {
        const d = getStatus(k.degisim);
        tbody.innerHTML += `<tr><td>${k.personel}</td><td>${k.tur}</td><td>${k.verilis}</td><td>${k.degisim}</td><td><span class="badge ${d.cls}">${d.text}</span></td><td><button class="btn btn-danger" onclick="sil('kkd', ${k.id})"><i class="fa-solid fa-trash"></i></button></td></tr>`;
      });
    }
    function kkdEkle() {
      const personel = document.getElementById('kkd_p').value.trim();
      const tur = document.getElementById('kkd_tur').value.trim();
      const verilis = document.getElementById('kkd_verilis').value;
      const omur = parseInt(document.getElementById('kkd_omur').value);
      if(!personel || !verilis) return alert("Ad ve tarih zorunlu.");
      let d = new Date(verilis); d.setMonth(d.getMonth() + omur);
      db.kkd.push({ id: Date.now(), personel, tur, verilis, degisim: d.toISOString().slice(0,10) });
      sync(); kkdCiz(); closeModal('modal-kkd');
    }

    function dofCiz() {
      const tbody = document.getElementById('dofTablo');
      tbody.innerHTML = '';
      db.dof.forEach(d => {
        tbody.innerHTML += `<tr><td>${d.tarih}</td><td>${d.bolum}</td><td>${d.tespit}</td><td>${d.aksiyon}</td><td>${d.sorumlu}</td><td><span class="badge ${d.statu==='Kapatıldı'?'badge-success':'badge-danger'}">${d.statu}</span></td><td><button class="btn btn-danger" onclick="sil('dof', ${d.id})"><i class="fa-solid fa-trash"></i></button></td></tr>`;
      });
    }
    function dofEkle() {
      const tarih = document.getElementById('dof_tarih').value;
      const bolum = document.getElementById('dof_bolum').value;
      const tespit = document.getElementById('dof_tespit').value;
      const aksiyon = document.getElementById('dof_aksiyon').value;
      const sorumlu = document.getElementById('dof_sorumlu').value;
      const statu = document.getElementById('dof_statu').value;
      db.dof.push({ id: Date.now(), tarih, bolum, tespit, aksiyon, sorumlu, statu });
      sync(); dofCiz(); closeModal('modal-dof');
    }

    function izinCiz() {
      const tbody = document.getElementById('izinTablo');
      tbody.innerHTML = '';
      db.izin.forEach(i => {
        tbody.innerHTML += `<tr><td>${i.no}</td><td>${i.tur}</td><td>${i.alan}</td><td>${i.personel}</td><td><span class="badge badge-success">Aktif</span></td><td><button class="btn btn-danger" onclick="sil('izin', ${i.id})"><i class="fa-solid fa-trash"></i></button></td></tr>`;
      });
    }
    function izinEkle() {
      const no = document.getElementById('ptw_no').value;
      const tur = document.getElementById('ptw_tur').value;
      const alan = document.getElementById('ptw_alan').value;
      const personel = document.getElementById('ptw_personel').value;
      db.izin.push({ id: Date.now(), no, tur, alan, personel });
      sync(); izinCiz(); closeModal('modal-izin');
    }

    function kazaCiz() {
      const tbody = document.getElementById('kazaTablo');
      tbody.innerHTML = '';
      db.k.forEach(k => {
        tbody.innerHTML += `<tr><td>${k.tarih}</td><td>${k.ad}</td><td>${k.tur}</td><td>${k.gun}</td><td><button class="btn btn-danger" onclick="sil('k', ${k.id})"><i class="fa-solid fa-trash"></i></button></td></tr>`;
      });
    }
    function kazaEkle() {
      const tarih = document.getElementById('k_tarih').value;
      const ad = document.getElementById('k_ad').value;
      const tur = document.getElementById('k_tur').value;
      const gun = document.getElementById('k_gun').value;
      db.k.push({ id: Date.now(), tarih, ad, tur, gun });
      sync(); kazaCiz(); closeModal('modal-kaza');
    }

    function hesaplaIstatistik() {
      const fiiliSaat = parseFloat(document.getElementById('stat_fiili_saat').value) || 1;
      const kazaSayisi = db.k.filter(x => x.tur !== 'Ramak Kala').length;
      let toplamKayip = 0;
      db.k.forEach(x => { if(x.tur !== 'Ramak Kala') toplamKayip += Number(x.gun)||0; });

      const kso = ((kazaSayisi * 1000000) / fiiliSaat).toFixed(2);
      const kao = ((toplamKayip * 1000) / fiiliSaat).toFixed(2);

      document.getElementById('res_kso').innerText = kso;
      document.getElementById('res_kao').innerText = kao;
    }

    function sil(kat, id) {
      if(confirm("Silinsin mi? Bu işlem buluttan da silecektir.")) {
        db[kat] = db[kat].filter(x => x.id !== id);
        sync();
        if(kat==='p') personelCiz(); if(kat==='s') sertifikaCiz();
        if(kat==='kkd') kkdCiz(); if(kat==='dof') dofCiz();
        if(kat==='izin') izinCiz(); if(kat==='k') kazaCiz();
      }
    }

    function excelIndir() {
      const wb = XLSX.utils.book_new();
      if(db.p.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(db.p), "Personel");
      if(db.s.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(db.s), "Sertifikalar");
      if(db.dof.length > 0) XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(db.dof), "DOF");
      XLSX.writeFile(wb, `ECE_TRAFO_ISG.xlsx`);
    }

    function cizimlerinHepsiniYenile() {
      sinifSec(db.ayar.sinif);
      updateHeaderBadge();
      personelCiz(); sertifikaCiz(); kkdCiz(); dofCiz(); izinCiz(); kazaCiz(); updateMetrics(); hesaplaIstatistik();
    }

    cizimlerinHepsiniYenile();
    buluttanYukle();
  </script>
</body>
</html>
