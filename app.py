import os
from flask import Flask, request, render_template, send_file, jsonify
import pandas as pd

app = Flask(__name__)

UPLOAD_FOLDER = r'C:\HIT\WebRapor\data\uploads'
REPORT_FOLDER = r'C:\HIT\WebRapor\reports'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

def mesai_kontrol(etiketler):
    """
    Mesai durumu tespiti:
      - 'mesaidişi', 'mesaidisi', 'mesaidışı' kelimeleri varsa 'Mesai Dışı'
      - Aksi takdirde 'Mesai İçi'
    """
    if isinstance(etiketler, str):
        etiketler = etiketler.strip().lower().replace(' ', '')
        if 'mesaidişi' in etiketler or 'mesaidisi' in etiketler or 'mesaidışı' in etiketler:
            return 'Mesai Dışı'
    return 'Mesai İçi'

def talebin_gelis_yolu_kontrol(etiket):
    """
    Etiketler sütunundaki veriyi inceleyerek, hangi anahtar kelime geçiyorsa 
    onu döndürüyoruz. Bu versiyonda Email/Mail/E-posta/Eposta/E-mail hepsi 
    'Mail' kategorisi altında toplanır. Ayrıca Bitrix yerine CRM kullanılmıştır.
    """
    if not isinstance(etiket, str):
        return "Eksik girilmiş"
    
    normalized = etiket.casefold()
    
    # Anahtar kelimeleri tek bir sözlükte topluyoruz.
    synonyms = {
        "Mail": ["mail", "email", "eposta", "e-posta", "e-mail"],
        "CRM": ["crm", "bitrix"],
        "Whmcs": ["whmcs"],
        "Gsm": ["gsm"],
        "Zoom": ["zoom"],
        "Whatsapp": ["whatsapp"]
    }
    
    for main_key, variants in synonyms.items():
        for v in variants:
            if v in normalized:
                return main_key
    return "Eksik girilmiş"

# Kişi -> Departman eşleştirmesi (kişisel isimler yerine genel tanımlayıcılar ve 
# departman isimleri daha genel ifadelerle değiştirilmiştir)
assignee_to_department = {
    "Kişi 1": "Departman 1",
    "Kişi 2": "Departman 2",
    "Kişi 3": "Departman 3",
    "Kişi 4": "Departman 4",
    "Kişi 5": "Departman 5",
    "Kişi 6": "Departman 6",
    "Kişi 7": "Departman 7",
    "Kişi 8": "Departman 8",
    "Kişi 9": "Departman 9",
    "Kişi 10": "Departman 10",
    "Kişi 11": "Departman 11",
    "Kişi 12": "Departman 12",
    "Kişi 13": "Departman 13",
    "Kişi 14": "Departman 14",
    "Kişi 15": "Departman 15",
    "Kişi 16": "Departman 16",
    "Kişi 17": "Departman 17",
    "Kişi 18": "Departman 18",
    "Kişi 19": "Departman 19",
    "Kişi 20": "Departman 20",
    "Kişi 21": "Departman 21",
    "Kişi 22": "Departman 22",
    "Kişi 23": "Departman 23",
    "Kişi 24": "Departman 24",
    "Kişi 25": "Departman 25",
    "Kişi 26": "Departman 26",
    "Kişi 27": "Departman 27",
    "Kişi 28": "Departman 28",
    "Kişi 29": "Departman 29",
    "Kişi 30": "Departman 30",
    "Kişi 31": "Departman 31",
    "Kişi 32": "Departman 32",
    "Kişi 33": "Departman 33",
    "Kişi 34": "Departman 34",
    "Kişi 35": "Departman 35",
    "Kişi 36": "Departman 36",
    "Kişi 37": "Departman 37",
    "Kişi 38": "Departman 38",
    "Kişi 39": "Departman 39",
    "Kişi 40": "Departman 40",
    "Kişi 41": "Departman 41",
    "Kişi 42": "Departman 42",
    "Kişi 43": "Departman 43",
    "Kişi 44": "Departman 44",
    "Kişi 45": "Departman 45",
    "Kişi 46": "Departman 46",
    "Kişi 47": "Departman 47",
    "Kişi 48": "Departman 48",
    "Kişi 49": "Departman 49",
    "Kişi 50": "Departman 50",
    "Kişi 51": "Departman 51"
}

def rapor_olustur(dosya_yolu, rapor_kayit_yolu):
    print(f"Yüklenen dosya yolu: {dosya_yolu}")
    df = pd.read_excel(dosya_yolu)
    print(f"Yüklenen veri çerçevesi boyutu: {df.shape}")
    df.columns = df.columns.str.strip().str.lower()
    
    # Etiketler sütununu kaydedelim ki talebin geliş yolu oluşturabilelim.
    if 'etiketler' in df.columns:
        df['orijinal_etiketler'] = df['etiketler']
    
    # Sütun yeniden adlandırma (assignee -> sorumlu kişi, etiketler -> mesai durumu)
    column_mapping = {
        'assignee': 'sorumlu kişi',
        'etiketler': 'mesai durumu'
    }
    df.rename(columns=column_mapping, inplace=True)
    
    # Gerekli sütunlar ve varsayılan değerler
    required_columns = {
        'görev': 'Bilinmeyen Görev',
        'oluşturan': 'Bilinmeyen Kişi',
        'sorumlu kişi': 'Bilinmeyen Kişi',
        'durum': 'Bilinmeyen Durum',
        'talebin iletildiği departman': 'Bilinmeyen Departman',
        'müşteri': 'Bilinmeyen Müşteri',
        'talebin geliş yolu': 'Bilinmeyen Yol',
        'oluşturulma tarihi': pd.NaT,
        'şu tarihte kapatılmış': pd.NaT,
        'ay': 'Bilinmeyen Ay',
        'mesai durumu': 'Mesai İçi'
    }
    
    for col, default_value in required_columns.items():
        if col not in df.columns:
            df[col] = default_value

    # "talebi ileten departman" varsa, kaldırıyoruz.
    if 'talebi ileten departman' in df.columns:
        df.drop(columns=['talebi ileten departman'], inplace=True)
        
    # Tarih sütunlarını dönüştürme
    df['oluşturulma tarihi'] = pd.to_datetime(df['oluşturulma tarihi'], errors='coerce')
    df['oluşturulma tarihi'] = df['oluşturulma tarihi'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
    df['şu tarihte kapatılmış'] = pd.to_datetime(df['şu tarihte kapatılmış'], errors='coerce')
    df['şu tarihte kapatılmış'] = df['şu tarihte kapatılmış'].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
    
    # Ay isimlerini Türkçe'ye çevirme
    ay_mapping = {
        'January': 'Ocak', 'February': 'Şubat', 'March': 'Mart', 'April': 'Nisan',
        'May': 'Mayıs', 'June': 'Haziran', 'July': 'Temmuz', 'August': 'Ağustos',
        'September': 'Eylül', 'October': 'Ekim', 'November': 'Kasım', 'December': 'Aralık'
    }
    df['ay'] = pd.to_datetime(df['oluşturulma tarihi'], errors='coerce').dt.strftime('%B')
    df['ay'] = df['ay'].map(ay_mapping).fillna('Bilinmeyen Ay')
    
    # Mesai durumu kontrolü
    df['mesai durumu'] = df['mesai durumu'].apply(mesai_kontrol)
    
    # Departman eşleştirme
    df['talebin iletildiği departman'] = df['sorumlu kişi'].map(assignee_to_department).fillna('Bilinmeyen Departman')
    df['oluşturan departman'] = df['oluşturan'].map(assignee_to_department).fillna('Bilinmeyen Departman')
    
    # Talebin Geliş Yolunu, orijinal etiketler üzerinden oluşturuyoruz.
    if 'orijinal_etiketler' in df.columns:
        df['talebin geliş yolu'] = df['orijinal_etiketler'].apply(talebin_gelis_yolu_kontrol)
    
    # ---------------------------------------------------
    # TABLO 1: Mesai Durumu + Departman satırlarda, Ay sütunlarda
    # ---------------------------------------------------
    table1 = pd.pivot_table(
        df,
        index=['mesai durumu', 'talebin iletildiği departman'],
        columns='ay',
        aggfunc='size',
        fill_value=0
    )
    table1['Genel Toplam'] = table1.sum(axis=1)
    table1 = table1.reset_index()

    # ---------------------------------------------------
    # TABLO 2: Mesai Dışında En Çok Talep İleten 10 Müşteri
    # ---------------------------------------------------
    df_mesai_dis = df[df['mesai durumu'] == 'Mesai Dışı']
    table2 = pd.pivot_table(
        df_mesai_dis,
        index='müşteri',
        columns='ay',
        aggfunc='size',
        fill_value=0
    )
    table2['Genel Toplam'] = table2.sum(axis=1)
    table2 = table2.sort_values('Genel Toplam', ascending=False).head(10)
    
    # ---------------------------------------------------
    # TABLO 3: En Çok Talep İleten 10 Müşteri (Tüm mesai durumları)
    # ---------------------------------------------------
    table3 = pd.pivot_table(
        df,
        index='müşteri',
        columns='ay',
        aggfunc='size',
        fill_value=0
    )
    table3['Genel Toplam'] = table3.sum(axis=1)
    table3 = table3.sort_values('Genel Toplam', ascending=False).head(10)
    
    # ---------------------------------------------------
    # TABLO 4: Talebin Geliş Yolu Bazında Talepler
    # ---------------------------------------------------
    table4 = pd.pivot_table(
        df,
        index='talebin geliş yolu',
        columns='ay',
        aggfunc='size',
        fill_value=0
    )
    table4['Genel Toplam'] = table4.sum(axis=1)
    table4 = table4.reset_index()
    
    # Excel'e yazdırma
    rapor_yolu = os.path.join(rapor_kayit_yolu, 'aylik_ticket_raporu.xlsx')
    
    with pd.ExcelWriter(rapor_yolu, engine='xlsxwriter') as writer:
        # 1) Ana Data sayfası: ham veri
        df.to_excel(writer, sheet_name='Ana Data', index=False)

        # 2) Özet Bilgi sayfası: özet tablolar
        sheet_name = 'Özet Bilgi'
        workbook = writer.book
        
        # -- Tablo 1 Yazdır --
        table1_startrow = 2
        table1.to_excel(
            writer, sheet_name=sheet_name,
            startrow=table1_startrow, startcol=0, index=False
        )
        
        ws = writer.sheets[sheet_name]
        ws.write(0, 0, "Departmanlara iletilen Ticket Adetleri")
        ws.write(1, 0, "Mesai Durumu & Departman & Aylık Ticket Sayıları")

        # --- Renklendirme (Tablo 1): Mesai Dışı kırmızı, Mesai İçi yeşil ---
        red_bg = workbook.add_format({'bg_color': '#FFC7CE'})
        green_bg = workbook.add_format({'bg_color': '#C6EFCE'})

        t1_rows, t1_cols = table1.shape
        first_data_row = table1_startrow + 1
        for i in range(t1_rows):
            mesai_val = table1.iloc[i, 0]
            excel_row = first_data_row + i
            if mesai_val == "Mesai Dışı":
                ws.set_row(excel_row, None, red_bg)
            else:
                ws.set_row(excel_row, None, green_bg)

        # -- Tablo 2 Yazdır --
        table2_startrow = table1_startrow + t1_rows + 4
        ws.write(table2_startrow, 0, "Mesai Dışında En Çok Talep İleten 10 Müşteri")
        table2.to_excel(
            writer, sheet_name=sheet_name,
            startrow=table2_startrow + 1, startcol=0
        )

        # -- Tablo 3 Yazdır --
        table3_startrow = table2_startrow + table2.shape[0] + 4
        ws.write(table3_startrow, 0, "En Çok Talep İleten 10 Müşteri (Tüm Durumlar)")
        table3.to_excel(
            writer, sheet_name=sheet_name,
            startrow=table3_startrow + 1, startcol=0
        )
        
        # -- Tablo 4 Yazdır --
        table4_startrow = table3_startrow + table3.shape[0] + 4
        ws.write(table4_startrow, 0, "Talebin Geliş Yolu Bazında Talepler")
        table4.to_excel(
            writer, sheet_name=sheet_name,
            startrow=table4_startrow + 1, startcol=0, index=False
        )

    return df, rapor_yolu

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/preview', methods=['POST'])
def preview():
    try:
        dosya = request.files.get('file')
        if not dosya:
            return jsonify({"error": "Dosya yüklenmedi."}), 400
        
        dosya_yolu = os.path.join(UPLOAD_FOLDER, dosya.filename)
        dosya.save(dosya_yolu)
        
        df_processed, _ = rapor_olustur(dosya_yolu, REPORT_FOLDER)
        
        if df_processed.empty:
            return jsonify({"error": "İşlenen veri boş! Lütfen Excel dosyanızı kontrol edin."}), 400
        
        df_processed = df_processed.astype(str)
        print("İşlenen verinin ilk 10 satırı:")
        print(df_processed.head(10))
        preview_rows = df_processed.head(10).to_dict(orient='records')
        
        return jsonify(preview_rows)
    except Exception as e:
        print("Hata oluştu:", str(e))
        return jsonify({"error": str(e)}), 500

@app.route('/download', methods=['POST'])
def download():
    try:
        dosya = request.files.get('file')
        if not dosya:
            return jsonify({"error": "Dosya yüklenmedi."}), 400
        
        dosya_yolu = os.path.join(UPLOAD_FOLDER, dosya.filename)
        dosya.save(dosya_yolu)
        
        _, rapor_yolu = rapor_olustur(dosya_yolu, REPORT_FOLDER)
        
        if not os.path.exists(rapor_yolu):
            return jsonify({"error": "Rapor oluşturulamadı, dosya bulunamadı."}), 500
        
        return send_file(rapor_yolu, as_attachment=True)
    except Exception as e:
        print("Hata oluştu:", str(e))
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)

