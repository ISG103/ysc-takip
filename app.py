from flask import Flask, render_template_string, request, redirect, session, url_for, send_file
import sqlite3
import os
import io
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = "super_gizli_anahtar_ysc_guvenlik_2026"
DB_NAME = "envanter.db"
UPLOAD_FOLDER = 'static/uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# GÜVENLİK ŞİFRELERİ
YONETICI_SIFRESI = "1234"   # Panel giriş ve lokasyon düzenleme şifresi
KONTROL_SIFRESI  = "1234"   # Sahada telefondan tüp onaylama PIN kodu

# LOGO DOĞRUDAN LİNKİ (HIZLI CDN - ASLA KAYBOLMAZ)
LOGO_SRC = "https://i.ibb.co/LdQyM8r/ece-logo.png"

def veritabanini_hazirla():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS tupler (
            kod TEXT PRIMARY KEY,
            tip TEXT,
            lokasyon TEXT,
            son_kontrol TEXT,
            sonraki_kontrol TEXT,
            kontrol_eden TEXT,
            durum TEXT,
            foto_yol TEXT
        )
    ''')
    try:
        cursor.execute("ALTER TABLE tupler ADD COLUMN foto_yol TEXT")
    except:
        pass

    cursor.execute("SELECT COUNT(*) FROM tupler")
    if cursor.fetchone()[0] == 0:
        bugun = datetime.now().strftime("%Y-%m-%d")
        sonraki = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")
        for i in range(1, 101):
            kod = f"YSC-{i:03d}"
            cursor.execute('''
                INSERT INTO tupler (kod, tip, lokasyon, son_kontrol, sonraki_kontrol, kontrol_eden, durum, foto_yol)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (kod, "6 KG KKT", f"Fabrika Alanı / Kolon-{i}", bugun, sonraki, "Sistem", "Gecerli", ""))
        conn.commit()
    conn.close()

veritabanini_hazirla()

# -------------------------------------------------------------
# 1. MOBİL EKRAN
# -------------------------------------------------------------
MOBIL_HTML = '''
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>YSC Kontrol Kartı</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #f1f5f9; padding: 15px; margin: 0; }
        .kart { background: white; border-radius: 14px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); max-width: 450px; margin: auto; }
        .logo-kutu { text-align: center; margin-bottom: 15px; }
        .logo-kutu img { max-height: 60px; max-width: 220px; object-fit: contain; }
        .baslik { font-size: 24px; font-weight: bold; color: #0f172a; margin-bottom: 5px; }
        .rozet { display: inline-block; padding: 5px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; margin-bottom: 15px; }
        .Gecerli { background: #dcfce7; color: #166534; }
        .Gecikmis { background: #fee2e2; color: #991b1b; }
        .satir { border-bottom: 1px solid #e2e8f0; padding: 10px 0; font-size: 14px; }
        .satir strong { color: #475569; display: inline-block; width: 130px; }
        .kontrol-kutusu { margin-top: 20px; background: #f8fafc; border-radius: 8px; padding: 15px; border: 1px solid #cbd5e1; }
        .kontrol-kutusu label { display: block; margin: 10px 0; font-size: 15px; cursor: pointer; color: #1e293b; }
        .girdi { width: 100%; box-sizing: border-box; padding: 10px; border-radius: 6px; border: 1px solid #cbd5e1; margin-top: 5px; margin-bottom: 12px; font-size: 15px; }
        .buton { display: block; width: 100%; background: #2563eb; color: white; border: none; padding: 14px; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 10px; }
        .mesaj { background: #e0f2fe; color: #0369a1; padding: 10px; border-radius: 6px; margin-bottom: 15px; font-size: 14px; text-align: center; }
        .yetki-alani { text-align: center; margin-top: 25px; padding-top: 15px; border-top: 1px dashed #cbd5e1; }
        .yetki-link { color: #64748b; font-size: 13px; text-decoration: none; }
        .tup-foto { width: 100%; max-height: 250px; object-fit: cover; border-radius: 8px; margin-top: 12px; border: 1px solid #cbd5e1; }
    </style>
</head>
<body>
    <div class="kart">
        <!-- LOGO -->
        <div class="logo-kutu">
            <img src="{{ logo_src }}" alt="Ece Trafo Logo">
        </div>

        {% if basarili %}<div class="mesaj">✓ Muayene kaydı başarıyla güncellendi!</div>{% endif %}
        
        <span class="rozet {{ tup[6] }}">{{ '✓ GEÇERLİ' if tup[6] == 'Gecerli' else '⚠ SÜRESİ GEÇMİŞ' }}</span>
        <div class="baslik">{{ tup[0] }}</div>
        <div style="color: #64748b; margin-bottom: 15px; font-size: 15px;">{{ tup[1] }}</div>
        
        <div class="satir"><strong>Lokasyon:</strong> {{ tup[2] }}</div>
        <div class="satir"><strong>Son Kontrol:</strong> {{ tup[3] }}</div>
        <div class="satir"><strong>Sonraki Kontrol:</strong> {{ tup[4] }}</div>
        <div class="satir"><strong>Son Denetleyen:</strong> {{ tup[5] }}</div>

        {% if tup[7] %}
            <div style="margin-top: 12px;">
                <span style="font-size: 12px; color: #64748b; font-weight: bold;">Son Denetim Fotoğrafı:</span>
                <img src="/static/uploads/{{ tup[7] }}" class="tup-foto" alt="Tüp Fotoğrafı">
            </div>
        {% endif %}

        {% if yetkili %}
            <form method="POST" action="/kontrol-kaydet/{{ tup[0] }}" enctype="multipart/form-data" class="kontrol-kutusu">
                <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
                    <h4 style="margin:0; color:#0f172a;">Yetkili Saha Denetimi:</h4>
                    <a href="/denetci-cikis/{{ tup[0] }}" style="font-size:11px; color:#ef4444; text-decoration:none;">(Yetkiyi Kapat)</a>
                </div>
                <label><input type="checkbox" required checked> Basınç İbresi Normal (Yeşilde)</label>
                <label><input type="checkbox" required checked> Emniyet Pimi ve Mühür Tam</label>
                <label><input type="checkbox" required checked> Hortum ve Tetik Mekanizması Sağlam</label>
                <label><input type="checkbox" required checked> Cihazın Önü Açık ve Erişilebilir</label>
                
                <label style="margin-top: 12px; font-weight: 600;">Kontrol Eden Personel:
                    <input type="text" name="personel" class="girdi" placeholder="Ad Soyad" required>
                </label>

                <label style="margin-top: 5px; font-weight: 600;">Fotoğraf Çek / Yükle (Opsiyonel):
                    <input type="file" name="foto" accept="image/*" capture="environment" class="girdi" style="padding: 6px;">
                </label>

                <button type="submit" class="buton">✓ Kontrolü Onayla ve Kaydet</button>
            </form>
        {% else %}
            <div class="yetki-alani">
                <a href="/denetci-giris/{{ tup[0] }}" class="yetki-link">🔒 Denetim Personeli Girişi</a>
            </div>
        {% endif %}
    </div>
</body>
</html>
'''

# -------------------------------------------------------------
# 2. DENETÇİ DOĞRULAMA
# -------------------------------------------------------------
DENETCI_LOGIN_HTML = '''
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Denetçi Doğrulama</title>
    <style>
        body { font-family: sans-serif; background: #f8fafc; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
        .kutu { background: white; padding: 25px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); width: 290px; text-align: center; }
        .logo-kutu { margin-bottom: 15px; }
        .logo-kutu img { max-height: 45px; }
        input { width: 100%; box-sizing: border-box; padding: 12px; border: 1px solid #cbd5e1; border-radius: 6px; margin: 15px 0; font-size: 18px; text-align: center; }
        button { width: 100%; padding: 12px; background: #2563eb; color: white; border: none; border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; }
        .hata { color: #dc2626; font-size: 13px; margin-bottom: 10px; }
    </style>
</head>
<body>
    <div class="kutu">
        <div class="logo-kutu">
            <img src="{{ logo_src }}" alt="Logo">
        </div>
        <h3 style="margin-top:0;">Yetkili Denetçi PIN</h3>
        <p style="font-size:13px; color:#64748b;">Kontrol onayı verebilmek için lütfen PIN kodunu girin.</p>
        {% if hata %}<div class="hata">{{ hata }}</div>{% endif %}
        <form method="POST">
            <input type="password" name="pin" placeholder="PIN Kodu" pattern="[0-9]*" inputmode="numeric" required autofocus>
            <button type="submit">Doğrula ve Aç</button>
            <a href="/tup/{{ kod }}" style="display:block; margin-top:12px; color:#64748b; font-size:12px; text-decoration:none;">Vazgeç</a>
        </form>
    </div>
</body>
</html>
'''

# -------------------------------------------------------------
# 3. YÖNETİCİ GİRİŞİ & PANELİ
# -------------------------------------------------------------
LOGIN_HTML = '''
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Yönetici Girişi</title>
    <style>
        body { font-family: sans-serif; background: #f8fafc; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
        .kutu { background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); width: 320px; text-align: center; }
        input { width: 100%; box-sizing: border-box; padding: 12px; border: 1px solid #cbd5e1; border-radius: 6px; margin: 15px 0; font-size: 16px; }
        button { width: 100%; padding: 12px; background: #2563eb; color: white; border: none; border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; }
        .hata { color: #dc2626; font-size: 14px; margin-bottom: 10px; }
    </style>
</head>
<body>
    <div class="kutu">
        <img src="{{ logo_src }}" alt="Logo" style="max-height: 50px; margin-bottom: 15px;">
        <h3 style="margin-top:0;">Yönetici Girişi</h3>
        {% if hata %}<div class="hata">{{ hata }}</div>{% endif %}
        <form method="POST">
            <input type="password" name="sifre" placeholder="Yönetici Şifresi" required autofocus>
            <button type="submit">Giriş Yap</button>
        </form>
    </div>
</body>
</html>
'''

PANEL_HTML = '''
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <title>YSC Envanter & Denetim Paneli</title>
    <style>
        body { font-family: sans-serif; background: #f8fafc; padding: 25px; margin: 0; }
        .container { max-width: 1200px; margin: auto; }
        .ust-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }
        .logo-ve-baslik { display: flex; align-items: center; gap: 18px; }
        .logo-ve-baslik img { max-height: 50px; max-width: 180px; object-fit: contain; }
        .aksiyonlar { display: flex; gap: 10px; align-items: center; }
        .ozet-kutulari { display: flex; gap: 20px; margin-bottom: 25px; }
        .kutu { flex: 1; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .sayi { font-size: 28px; font-weight: bold; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        th, td { padding: 12px 14px; text-align: left; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
        th { background: #f1f5f9; color: #475569; }
        .badge { padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: bold; }
        .Gecerli { background: #dcfce7; color: #166534; }
        .Gecikmis { background: #fee2e2; color: #991b1b; }
        .btn-excel { background: #15803d; color: white; text-decoration: none; padding: 8px 16px; border-radius: 6px; font-size: 13px; font-weight: bold; display: inline-flex; align-items: center; gap: 6px; }
        .btn-duzenle { background: #0284c7; color: white; text-decoration: none; padding: 5px 10px; border-radius: 5px; font-size: 12px; font-weight: bold; }
        .btn-cikis { background: #ef4444; color: white; text-decoration: none; padding: 8px 14px; border-radius: 6px; font-size: 13px; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <div class="ust-bar">
            <div class="logo-ve-baslik">
                <img src="{{ logo_src }}" alt="Ece Trafo Logo">
                <h2 style="margin:0;">Yangın Söndürme Cihazı (YSC) Takip Paneli</h2>
            </div>
            <div class="aksiyonlar">
                <a href="/excel-indir" class="btn-excel">📥 Denetim Excel Raporu İndir</a>
                <a href="/cikis" class="btn-cikis">Çıkış</a>
            </div>
        </div>

        <div class="ozet-kutulari">
            <div class="kutu"><div>Toplam Ekipman</div><div class="sayi" style="color: #2563eb;">{{ toplam }}</div></div>
            <div class="kutu"><div>Geçerli / Kontrol Edilmiş</div><div class="sayi" style="color: #16a34a;">{{ gecerli }}</div></div>
            <div class="kutu"><div>Gecikmiş / Kontrol Bekleyen</div><div class="sayi" style="color: #dc2626;">{{ gecikmis }}</div></div>
        </div>

        <table>
            <thead>
                <tr>
                    <th>Kod</th><th>Tip</th><th>Lokasyon</th><th>Son Kontrol</th><th>Sonraki Kontrol</th><th>Kontrol Eden</th><th>Fotoğraf</th><th>Durum</th><th>İşlem</th>
                </tr>
            </thead>
            <tbody>
                {% for t in tupler %}
                <tr>
                    <td><strong><a href="/tup/{{ t[0] }}" target="_blank">{{ t[0] }}</a></strong></td>
                    <td>{{ t[1] }}</td>
                    <td><span style="color: #0369a1; font-weight: 500;">{{ t[2] }}</span></td>
                    <td>{{ t[3] }}</td>
                    <td>{{ t[4] }}</td>
                    <td>{{ t[5] }}</td>
                    <td>
                        {% if t[7] %}
                            <a href="/static/uploads/{{ t[7] }}" target="_blank" style="color:#2563eb; font-weight:bold; font-size:12px;">📷 Gör</a>
                        {% else %}
                            <span style="color:#94a3b8; font-size:12px;">-</span>
                        {% endif %}
                    </td>
                    <td><span class="badge {{ t[6] }}">{{ t[6] }}</span></td>
                    <td><a href="/duzenle/{{ t[0] }}" class="btn-duzenle">✎ Düzenle</a></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
'''

DUZENLE_HTML = '''
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tüp Bilgilerini Düzenle</title>
    <style>
        body { font-family: sans-serif; background: #f1f5f9; padding: 20px; margin: 0; }
        .kart { background: white; border-radius: 12px; padding: 25px; max-width: 500px; margin: auto; box-shadow: 0 4px 10px rgba(0,0,0,0.08); }
        .girdi { width: 100%; box-sizing: border-box; padding: 10px; border-radius: 6px; border: 1px solid #cbd5e1; margin-top: 6px; margin-bottom: 15px; font-size: 15px; }
        .kaydet-btn { background: #16a34a; color: white; border: none; padding: 12px 20px; border-radius: 6px; font-weight: bold; cursor: pointer; width: 100%; font-size: 16px; }
        .iptal-btn { display: block; text-align: center; margin-top: 10px; color: #64748b; text-decoration: none; font-size: 14px; }
    </style>
</head>
<body>
    <div class="kart">
        <h2>{{ tup[0] }} Düzenle</h2>
        <form method="POST">
            <label><strong>Ekipman Tipi / Kapasitesi:</strong></label>
            <input type="text" name="tip" class="girdi" value="{{ tup[1] }}" required>
            <label><strong>Yeni Lokasyon (Bina / Kat / Bölüm):</strong></label>
            <input type="text" name="lokasyon" class="girdi" value="{{ tup[2] }}" required>
            <button type="submit" class="kaydet-btn">Bilgileri Güncelle</button>
            <a href="/panel" class="iptal-btn">← Panele Geri Dön</a>
        </form>
    </div>
</body>
</html>
'''

# -------------------------------------------------------------
# ROTALAR
# -------------------------------------------------------------

@app.route('/tup/<kod>')
def tup_detay(kod):
    basarili = request.args.get('kaydedildi', False)
    yetkili = session.get('denetci_yetkisi', False)
    
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT * FROM tupler WHERE kod = ?", (kod,))
    tup = c.fetchone()
    conn.close()
    
    if not tup:
        return "Ekipman bulunamadı!", 404
        
    return render_template_string(MOBIL_HTML, tup=tup, basarili=basarili, yetkili=yetkili, logo_src=LOGO_SRC)

@app.route('/denetci-giris/<kod>', methods=['GET', 'POST'])
def denetci_giris(kod):
    hata = None
    if request.method == 'POST':
        if request.form.get('pin') == KONTROL_SIFRESI:
            session['denetci_yetkisi'] = True
            return redirect(f"/tup/{kod}")
        else:
            hata = "Hatalı PIN kodu!"
    return render_template_string(DENETCI_LOGIN_HTML, kod=kod, hata=hata, logo_src=LOGO_SRC)

@app.route('/denetci-cikis/<kod>')
def denetci_cikis(kod):
    session.pop('denetci_yetkisi', None)
    return redirect(f"/tup/{kod}")

@app.route('/kontrol-kaydet/<kod>', methods=['POST'])
def kontrol_kaydet(kod):
    if not session.get('denetci_yetkisi'):
        return "Yetkisiz işlem!", 403

    personel = request.form.get('personel', 'Yetkili Personel')
    su_an = datetime.now().strftime("%Y-%m-%d %H:%M")
    sonraki = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")
    
    foto_adi = ""
    if 'foto' in request.files:
        foto = request.files['foto']
        if foto and foto.filename != '':
            uzanti = os.path.splitext(foto.filename)[1]
            foto_adi = f"{kod}_{int(datetime.now().timestamp())}{uzanti}"
            foto.save(os.path.join(UPLOAD_FOLDER, foto_adi))

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    if foto_adi != "":
        c.execute('''
            UPDATE tupler 
            SET son_kontrol = ?, sonraki_kontrol = ?, kontrol_eden = ?, durum = 'Gecerli', foto_yol = ?
            WHERE kod = ?
        ''', (su_an, sonraki, personel, foto_adi, kod))
    else:
        c.execute('''
            UPDATE tupler 
            SET son_kontrol = ?, sonraki_kontrol = ?, kontrol_eden = ?, durum = 'Gecerli'
            WHERE kod = ?
        ''', (su_an, sonraki, personel, kod))
        
    conn.commit()
    conn.close()
    
    return redirect(f"/tup/{kod}?kaydedildi=1")

@app.route('/login', methods=['GET', 'POST'])
def login():
    hata = None
    if request.method == 'POST':
        if request.form.get('sifre') == YONETICI_SIFRESI:
            session['giris_yapti'] = True
            return redirect('/panel')
        else:
            hata = "Hatalı şifre!"
    return render_template_string(LOGIN_HTML, hata=hata, logo_src=LOGO_SRC)

@app.route('/cikis')
def cikis():
    session.pop('giris_yapti', None)
    return redirect('/login')

@app.route('/panel')
def yonetici_paneli():
    if not session.get('giris_yapti'):
        return redirect('/login')
        
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    bugun = datetime.now().strftime("%Y-%m-%d")
    c.execute("UPDATE tupler SET durum = 'Gecikmis' WHERE sonraki_kontrol < ?", (bugun,))
    conn.commit()
    c.execute("SELECT * FROM tupler ORDER BY kod ASC")
    tupler = c.fetchall()
    toplam = len(tupler)
    gecerli = sum(1 for t in tupler if t[6] == 'Gecerli')
    gecikmis = toplam - gecerli
    conn.close()
    return render_template_string(PANEL_HTML, tupler=tupler, toplam=toplam, gecerli=gecerli, gecikmis=gecikmis, logo_src=LOGO_SRC)

@app.route('/excel-indir')
def excel_indir():
    if not session.get('giris_yapti'):
        return redirect('/login')

    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT kod, tip, lokasyon, son_kontrol, sonraki_kontrol, kontrol_eden, durum FROM tupler ORDER BY kod ASC")
    veriler = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "YSC Denetim Listesi"

    ws.merge_cells('A1:G1')
    ws['A1'] = "ECE TRAFO - YANGIN SÖNDÜRME CİHAZLARI PERİYODİK KONTROL RAPORU"
    ws['A1'].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
    ws['A1'].fill = PatternFill(start_color="1E3A8A", fill_type="solid")
    ws['A1'].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    sutunlar = ["Ekipman Kodu", "Tip / Kapasite", "Lokasyon", "Son Kontrol Tarihi", "Sonraki Kontrol Tarihi", "Denetleyen", "Durum"]
    ws.append([])
    ws.append(sutunlar)

    for col in range(1, 8):
        hucre = ws.cell(row=3, column=col)
        hucre.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        hucre.fill = PatternFill(start_color="3B82F6", fill_type="solid")
        hucre.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 25

    for sira, row in enumerate(veriler, start=4):
        ws.append(list(row))
        durum_hucre = ws.cell(row=sira, column=7)
        if row[6] == 'Gecerli':
            durum_hucre.fill = PatternFill(start_color="DCFCE7", fill_type="solid")
            durum_hucre.font = Font(color="15803D", bold=True)
        else:
            durum_hucre.fill = PatternFill(start_color="FEE2E2", fill_type="solid")
            durum_hucre.font = Font(color="B91C1C", bold=True)

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    dosya_adi = f"YSC_Denetim_Raporu_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(buffer, as_attachment=True, download_name=dosya_adi, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/duzenle/<kod>', methods=['GET', 'POST'])
def duzenle(kod):
    if not session.get('giris_yapti'):
        return redirect('/login')
        
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    if request.method == 'POST':
        yeni_tip = request.form.get('tip')
        yeni_lokasyon = request.form.get('lokasyon')
        c.execute("UPDATE tupler SET tip = ?, lokasyon = ? WHERE kod = ?", (yeni_tip, yeni_lokasyon, kod))
        conn.commit()
        conn.close()
        return redirect('/panel')
        
    c.execute("SELECT * FROM tupler WHERE kod = ?", (kod,))
    tup = c.fetchone()
    conn.close()
    return render_template_string(DUZENLE_HTML, tup=tup)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
