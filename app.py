from flask import Flask, render_template_string, request, redirect
import sqlite3
from datetime import datetime, timedelta

app = Flask(__name__)
DB_NAME = "envanter.db"

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
            durum TEXT
        )
    ''')
    cursor.execute("SELECT COUNT(*) FROM tupler")
    if cursor.fetchone()[0] == 0:
        bugun = datetime.now().strftime("%Y-%m-%d")
        sonraki = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")
        for i in range(1, 101):
            kod = f"YSC-{i:03d}"
            cursor.execute('''
                INSERT INTO tupler VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (kod, "6 KG KKT", f"Fabrika Alanı / Kolon-{i}", bugun, sonraki, "Sistem", "Gecerli"))
        conn.commit()
    conn.close()

veritabanini_hazirla()

MOBIL_HTML = '''
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>YSC Kontrol Kartı</title>
    <style>
        body { font-family: -apple-system, sans-serif; background: #f1f5f9; padding: 15px; margin: 0; }
        .kart { background: white; border-radius: 14px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.08); max-width: 450px; margin: auto; }
        .baslik { font-size: 22px; font-weight: bold; color: #0f172a; margin-bottom: 5px; }
        .rozet { display: inline-block; padding: 5px 12px; border-radius: 20px; font-size: 13px; font-weight: bold; margin-bottom: 15px; }
        .Gecerli { background: #dcfce7; color: #166534; }
        .Gecikmis { background: #fee2e2; color: #991b1b; }
        .satir { border-bottom: 1px solid #e2e8f0; padding: 10px 0; font-size: 14px; }
        .satir strong { color: #475569; display: inline-block; width: 130px; }
        .kontrol-kutusu { margin-top: 20px; background: #f8fafc; border-radius: 8px; padding: 15px; border: 1px solid #e2e8f0; }
        .kontrol-kutusu label { display: block; margin: 10px 0; font-size: 15px; cursor: pointer; color: #1e293b; }
        .girdi { width: 100%; box-sizing: border-box; padding: 10px; border-radius: 6px; border: 1px solid #cbd5e1; margin-top: 5px; margin-bottom: 12px; }
        .buton { display: block; width: 100%; background: #2563eb; color: white; border: none; padding: 14px; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; margin-top: 10px; }
        .mesaj { background: #e0f2fe; color: #0369a1; padding: 10px; border-radius: 6px; margin-bottom: 15px; font-size: 14px; text-align: center; }
    </style>
</head>
<body>
    <div class="kart">
        {% if basarili %}<div class="mesaj">✓ Muayene kaydı başarıyla güncellendi!</div>{% endif %}
        <span class="rozet {{ tup[6] }}">{{ '✓ GEÇERLİ' if tup[6] == 'Gecerli' else '⚠ SÜRESİ GEÇMİŞ' }}</span>
        <div class="baslik">{{ tup[0] }}</div>
        <div style="color: #64748b; margin-bottom: 15px; font-size: 15px;">{{ tup[1] }}</div>
        <div class="satir"><strong>Lokasyon:</strong> {{ tup[2] }}</div>
        <div class="satir"><strong>Son Kontrol:</strong> {{ tup[3] }}</div>
        <div class="satir"><strong>Sonraki Kontrol:</strong> {{ tup[4] }}</div>
        <div class="satir"><strong>Son Denetleyen:</strong> {{ tup[5] }}</div>
        <form method="POST" action="/kontrol-kaydet/{{ tup[0] }}" class="kontrol-kutusu">
            <h4 style="margin: 0 0 10px 0; color: #0f172a;">Saha Denetim Kontrolü:</h4>
            <label><input type="checkbox" required checked> Basınç İbresi Normal (Yeşilde)</label>
            <label><input type="checkbox" required checked> Emniyet Pimi ve Mühür Tam</label>
            <label><input type="checkbox" required checked> Hortum ve Tetik Mekanizması Sağlam</label>
            <label><input type="checkbox" required checked> Cihazın Önü Açık ve Erişilebilir</label>
            <label style="margin-top: 15px; font-weight: 600;">Kontrol Eden Personel:
                <input type="text" name="personel" class="girdi" placeholder="Ad Soyad" required>
            </label>
            <button type="submit" class="buton">✓ Kontrolü Onayla ve Kaydet</button>
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
        .container { max-width: 1100px; margin: auto; }
        .ozet-kutulari { display: flex; gap: 20px; margin-bottom: 25px; }
        .kutu { flex: 1; background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .sayi { font-size: 28px; font-weight: bold; margin-top: 5px; }
        table { width: 100%; border-collapse: collapse; background: white; border-radius: 10px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        th, td { padding: 12px 16px; text-align: left; border-bottom: 1px solid #f1f5f9; font-size: 14px; }
        th { background: #f1f5f9; color: #475569; }
        .badge { padding: 4px 8px; border-radius: 6px; font-size: 12px; font-weight: bold; }
        .Gecerli { background: #dcfce7; color: #166534; }
        .Gecikmis { background: #fee2e2; color: #991b1b; }
    </style>
</head>
<body>
    <div class="container">
        <h2>Yangın Söndürme Cihazı (YSC) Takip Paneli</h2>
        <div class="ozet-kutulari">
            <div class="kutu"><div>Toplam Ekipman</div><div class="sayi" style="color: #2563eb;">{{ toplam }}</div></div>
            <div class="kutu"><div>Geçerli / Kontrol Edilmiş</div><div class="sayi" style="color: #16a34a;">{{ gecerli }}</div></div>
            <div class="kutu"><div>Gecikmiş / Kontrol Bekleyen</div><div class="sayi" style="color: #dc2626;">{{ gecikmis }}</div></div>
        </div>
        <table>
            <thead>
                <tr>
                    <th>Kod</th><th>Tip</th><th>Lokasyon</th><th>Son Kontrol</th><th>Sonraki Kontrol</th><th>Kontrol Eden</th><th>Durum</th>
                </tr>
            </thead>
            <tbody>
                {% for t in tupler %}
                <tr>
                    <td><strong><a href="/tup/{{ t[0] }}" target="_blank">{{ t[0] }}</a></strong></td>
                    <td>{{ t[1] }}</td><td>{{ t[2] }}</td><td>{{ t[3] }}</td><td>{{ t[4] }}</td><td>{{ t[5] }}</td>
                    <td><span class="badge {{ t[6] }}">{{ t[6] }}</span></td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>
</body>
</html>
'''

@app.route('/tup/<kod>')
def tup_detay(kod):
    basarili = request.args.get('kaydedildi', False)
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute("SELECT * FROM tupler WHERE kod = ?", (kod,))
    tup = c.fetchone()
    conn.close()
    if not tup:
        return "Ekipman bulunamadı!", 404
    return render_template_string(MOBIL_HTML, tup=tup, basarili=basarili)

@app.route('/kontrol-kaydet/<kod>', methods=['POST'])
def kontrol_kaydet(kod):
    personel = request.form.get('personel', 'Bilinmeyen Personel')
    su_an = datetime.now().strftime("%Y-%m-%d %H:%M")
    sonraki = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    c.execute('''
        UPDATE tupler 
        SET son_kontrol = ?, sonraki_kontrol = ?, kontrol_eden = ?, durum = 'Gecerli'
        WHERE kod = ?
    ''', (su_an, sonraki, personel, kod))
    conn.commit()
    conn.close()
    return redirect(f"/tup/{kod}?kaydedildi=1")

@app.route('/panel')
def yonetici_paneli():
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
    return render_template_string(PANEL_HTML, tupler=tupler, toplam=toplam, gecerli=gecerli, gecikmis=gecikmis)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)