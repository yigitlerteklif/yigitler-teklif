import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import urllib.parse

# Kurumsal Sayfa Yapılandırması
st.set_page_config(layout="wide", page_title="Yiğitler Teklif Programı", page_icon="🏢")

# Özel CSS Tasarımları
st.markdown("""
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 1.5rem;}
    .stRadio > div {flex-direction: row; gap: 15px;}
    div[data-testid="stExpander"] {border: 2px solid #ffcc00; border-radius: 8px; background-color: #fffdf2;}
    .metric-box {background-color: #f1f5f9; padding: 15px; border-radius: 8px; text-align: center; border: 1px solid #cbd5e1;}
    </style>
""", unsafe_allow_html=True)

# 1. VERİTABANI BAĞLANTISI (V6.3 - Hizalama Hatası Giderildi)
conn = sqlite3.connect("yigitler_bayi_v6.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute('''
    CREATE TABLE IF NOT EXISTS musteriler (
        id INTEGER PRIMARY KEY AUTOINCREMENT, isim TEXT, telefon TEXT, adres TEXT
    )''')
cursor.execute('''
    CREATE TABLE IF NOT EXISTS teklifler (
        id INTEGER PRIMARY KEY AUTOINCREMENT, musteri_id INTEGER, urunler TEXT,
        brut_toplam REAL, bundle_indirimi REAL, musteri_fiyati REAL,
        lojistik_maliyet REAL, pos_komisyon REAL, net_kar REAL, durum TEXT, tarih TEXT,
        sube TEXT, personel TEXT
    )''')
cursor.execute('''
    CREATE TABLE IF NOT EXISTS urunler (
        urun_kodu TEXT PRIMARY KEY, urun_adi TEXT, grup1_marka TEXT, grup2_kategori TEXT,
        brut_maliyet REAL, fiyat_farki REAL, satis_fiyati REAL, lojistik_maliyet REAL
    )''')
cursor.execute('''
    CREATE TABLE IF NOT EXISTS banka_komisyonlari (
        banka_adi TEXT PRIMARY KEY,
        taksit_1 REAL, taksit_2 REAL, taksit_3 REAL, taksit_4 REAL, taksit_5 REAL,
        taksit_6 REAL, taksit_7 REAL, taksit_8 REAL, taksit_9 REAL, taksit_10 REAL,
        taksit_11 REAL, taksit_12 REAL
    )''')
cursor.execute('CREATE TABLE IF NOT EXISTS personeller (personel_adi TEXT PRIMARY KEY)')
cursor.execute('CREATE TABLE IF NOT EXISTS subeler (sube_adi TEXT PRIMARY KEY)')
conn.commit()

# 2. ÜST KURUMSAL LOGO VE BAŞLIK BARI
st.markdown("""
    <div style='display: flex; justify-content: space-between; align-items: center; background-color: #0f2c59; padding: 15px 30px; border-radius: 10px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
        <img src='https://wikimedia.org' width='110' style='filter: brightness(0) invert(1);'>
        <h1 style='text-align: center; color: white; margin: 0; font-family: sans-serif; font-size: 26px; font-weight: bold; letter-spacing: 1px;'>🏢 YİĞİTLER TEKLİF & CRM YÖNETİMİ</h1>
        <img src='https://istikbal.com.tr' width='130' style='background-color:white; padding:5px; border-radius:5px;'>
    </div>
""", unsafe_allow_html=True)

# YAN MENÜ TASARIMI
st.sidebar.markdown("<h2 style='color:#0f2c59; text-align:center;'>📱 Menü</h2>", unsafe_allow_html=True)
sayfa = st.sidebar.radio("İşlem Yapılacak Ekran:", ["📝 Teklif Oluştur (Satış)", "📊 Merkezi CRM & Takip Panel", "🔄 Merkez - Excel Yükleme Odası"])

# 3. SAYFA: ÇOKLU EXCEL YÜKLEME ODASI
if sayfa == "🔄 Merkez - Excel Yükleme Odası":
    st.header("🔄 Gelişmiş Excel Veri Giriş ve Finans Odası")
    st.markdown("---")

    c_ex1, c_ex2 = st.columns(2)
    with c_ex1:
        st.subheader("📦 1. Ürün Portföyü Yükle")
        urun_dosya = st.file_uploader("Ürün Excel Listesini Seçin", type=["xlsx", "xls"], key="u_key")
        if urun_dosya:
            df_u = pd.read_excel(urun_dosya)
            st.dataframe(df_u.head(2), use_container_width=True)
            if st.button("🚀 Ürünleri Güncelle"):
                cursor.execute("DELETE FROM urunler")
                for _, row in df_u.iterrows():
                    cursor.execute("INSERT OR REPLACE INTO urunler VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                   (str(row['urun_kodu']), str(row['urun_adi']), str(row['grup1_marka']), str(row['grup2_kategori']),
                                    float(row['brut_maliyet']), float(row['fiyat_farki']), float(row['satis_fiyati']), float(row['lojistik_maliyet'])))
                conn.commit(); st.success("Ürünler Güncellendi!")

        st.write("---")
        st.subheader("🧑‍💼 3. Personel (Satış Temsilcisi) Listesi Yükle")
        st.info("Sütun Başlığı: 'personel_adi'")
        per_dosya = st.file_uploader("Personel Excel Listesini Seçin", type=["xlsx", "xls"], key="p_key")
        if per_dosya:
            df_p = pd.read_excel(per_dosya)
            st.dataframe(df_p.head(2), use_container_width=True)
            if st.button("🚀 Personel Listesini Güncelle"):
                cursor.execute("DELETE FROM personeller")
                for _, row in df_p.iterrows():
                    cursor.execute("INSERT OR REPLACE INTO personeller VALUES (?)", (str(row['personel_adi']),))
                conn.commit(); st.success("Personel Listesi Güncellendi!")

    with c_ex2:
        st.subheader("💳 2. Banka Taksit Matrisi Yükle")
        komisyon_dosya = st.file_uploader("Banka Komisyon Excel Listesini Seçin", type=["xlsx", "xls"], key="k_key")
        if komisyon_dosya:
            df_k = pd.read_excel(komisyon_dosya)
            st.dataframe(df_k.head(2), use_container_width=True)
            if st.button("🚀 Banka Komisyonlarını Güncelle"):
                cursor.execute("DELETE FROM banka_komisyonlari")
                for _, row in df_k.iterrows():
                    cursor.execute("INSERT OR REPLACE INTO banka_komisyonlari VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                                   (str(row['banka_adi']), float(row['taksit_1']), float(row['taksit_2']), float(row['taksit_3']), float(row['taksit_4']), float(row['taksit_5']), float(row['taksit_6']), float(row['taksit_7']), float(row['taksit_8']), float(row['taksit_9']), float(row['taksit_10']), float(row['taksit_11']), float(row['taksit_12'])))
                conn.commit(); st.success("Banka Komisyonları Güncellendi!")

        st.write("---")
        st.subheader("🏪 4. Şube Listesi Yükle")
        st.info("Sütun Başlığı: 'sube_adi'")
        sube_dosya = st.file_uploader("Şube Excel Listesini Seçin", type=["xlsx", "xls"], key="s_key")
        if sube_dosya:
            df_s = pd.read_excel(sube_dosya)
            st.dataframe(df_s.head(2), use_container_width=True)
            if st.button("🚀 Şube Listesini Güncelle"):
                cursor.execute("DELETE FROM subeler")
                for _, row in df_s.iterrows():
                    cursor.execute("INSERT OR REPLACE INTO subeler VALUES (?)", (str(row['sube_adi']),))
                conn.commit(); st.success("Şube Listesi Güncellendi!")

# 4. SAYFA: TEKLİF OLUŞTURMA
elif sayfa == "📝 Teklif Oluştur (Satış)":
    st.subheader("📝 Adım Adım Güvenli Teklif Hazırlama")

    # ADIM 1: OPERASYONEL BİLGİLER (ŞUBE VE PERSONEL SEÇİMİ)
    with st.container(border=True):
        st.markdown("#### 🏪 Adım 1: Şube ve Satış Temsilcisi Bilgisi")
        sube_list_df = pd.read_sql_query("SELECT * FROM subeler", conn)
        per_list_df = pd.read_sql_query("SELECT * FROM personeller", conn)

        c_op1, c_op2 = st.columns(2)
        with c_op1:
            if sube_list_df.empty:
                st.warning("Lütfen şube listesi excelini yükleyin.")
                secilen_sube = "Bilinmeyen Şube"
            else:
                secilen_sube = st.selectbox("İşlem Yapılan Şubeyi Seçin:", sube_list_df['sube_adi'].tolist())
        with c_op2:
            if per_list_df.empty:
                st.warning("Lütfen personel listesi excelini yükleyin.")
                secilen_personel = "Bilinmeyen Personel"
            else:
                secilen_personel = st.selectbox("Teklifi Hazırlayan Satış Temsilcisi:", per_list_df['personel_adi'].tolist())

    # ADIM 2: MÜŞTERI BİLGİ KARTI VE ADRES
    with st.container(border=True):
        st.markdown("#### 👤 Adım 2: Müşteri Seçimi ve Detaylı Adres")
        musteri_df = pd.read_sql_query("SELECT id, isim FROM musteriler", conn)
        musteri_tipi = st.radio("Müşteri İşlem Türü:", ["Kayıtlı Müşterilerden Seç", "Yeni Müşteri Tanımla"], label_visibility="collapsed")

        m_id = None  # Yeni müşteri henüz kaydedilmediyse None kalır
        c1, c2 = st.columns(2)
        if musteri_tipi == "Yeni Müşteri Tanımla" or musteri_df.empty:
            with c1: m_isim = st.text_input("Müşteri Adı Soyadı / Firma Adı:")
            with c2: m_tel = st.text_input("Telefon Numarası (Örn: 5321234567):")
            m_adres = st.text_area("Teslimat / Fatura Adresi:")
            if st.button("➕ Müşteriyi ve Adresini Kaydet", type="primary"):
                if m_isim and m_tel:
                    cursor.execute("INSERT INTO musteriler (isim, telefon, adres) VALUES (?, ?, ?)", (m_isim, m_tel, m_adres))
                    conn.commit()
                    st.success(f"✔️ {m_isim} adresiyle birlikte kaydedildi. Yukarıdan 'Kayıtlı Müşterilerden Seç' diyebilirsiniz.")
                    st.rerun()
                else: st.error("Ad ve telefon alanları zorunludur!")
        else:
            with c1:
                secilen_musteri = st.selectbox("Sistemde Kayıtlı Müşteriler:", musteri_df['isim'].tolist())
                m_id = int(musteri_df.loc[musteri_df['isim'] == secilen_musteri, 'id'].iloc[0])
                m_data = pd.read_sql_query(f"SELECT telefon, adres FROM musteriler WHERE id={m_id}", conn).iloc[0]
                m_tel = m_data['telefon']
                m_adres = m_data['adres']
            with c2:
                st.write("")
                st.markdown(f"**📞 İletişim:** {m_tel}")
                st.markdown(f"**📍 Güncel Adres:** {m_adres}")

    # ADIM 3: ÖDEME VE KAMPANYA (Hizalama Hatası Çözüldü)
    with st.container(border=True):
        st.markdown("#### 💳 Adım 3: Ödeme Yöntemi ve Kampanya Tanımı")
        banka_df = pd.read_sql_query("SELECT * FROM banka_komisyonlari", conn)
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            if banka_df.empty:
                st.warning("Lütfen banka komisyon listesi excelini yükleyin.")
                secilen_banka = None
            else:
                secilen_banka = st.selectbox("Ödeme Bankası:", banka_df['banka_adi'].tolist())
        with c_f2:
            taksit_sayisi = st.selectbox("Taksit Sayısı:", list(range(1, 13)))
        with c_f3:
            bundle_indirim_tutari = st.number_input("Kampanya / Ek İndirim (TL):", min_value=0.0, value=0.0, step=50.0)

    # ADIM 4: ÜRÜN SEÇİMİ VE SEPET
    with st.container(border=True):
        st.markdown("#### 📦 Adım 4: Ürün Seçimi")
        urun_df = pd.read_sql_query("SELECT * FROM urunler", conn)

        if "sepet" not in st.session_state:
            st.session_state.sepet = []

        if urun_df.empty:
            st.warning("Lütfen ürün listesi excelini yükleyin.")
        else:
            c_u1, c_u2, c_u3 = st.columns([3, 1, 1])
            with c_u1:
                secilen_urun_adi = st.selectbox("Ürün Seçin:", urun_df['urun_adi'].tolist())
            with c_u2:
                adet = st.number_input("Adet:", min_value=1, value=1, step=1)
            with c_u3:
                st.write("")
                st.write("")
                if st.button("➕ Sepete Ekle"):
                    urun_row = urun_df[urun_df['urun_adi'] == secilen_urun_adi].iloc[0]
                    st.session_state.sepet.append({
                        "urun_kodu": urun_row['urun_kodu'],
                        "urun_adi": urun_row['urun_adi'],
                        "adet": int(adet),
                        "satis_fiyati": float(urun_row['satis_fiyati']),
                        "brut_maliyet": float(urun_row['brut_maliyet']),
                        "lojistik_maliyet": float(urun_row['lojistik_maliyet']),
                    })
                    st.rerun()

            if st.session_state.sepet:
                sepet_df = pd.DataFrame(st.session_state.sepet)
                sepet_df["Ara Toplam"] = sepet_df["adet"] * sepet_df["satis_fiyati"]
                st.dataframe(
                    sepet_df[["urun_adi", "adet", "satis_fiyati", "Ara Toplam"]].rename(columns={
                        "urun_adi": "Ürün", "adet": "Adet", "satis_fiyati": "Birim Fiyat"
                    }),
                    use_container_width=True, hide_index=True
                )

                sil_secimi = st.selectbox(
                    "Sepetten çıkarılacak ürün:",
                    ["--"] + [f"{i}: {s['urun_adi']} (x{s['adet']})" for i, s in enumerate(st.session_state.sepet)]
                )
                if sil_secimi != "--" and st.button("🗑️ Ürünü Sepetten Çıkar"):
                    idx = int(sil_secimi.split(":")[0])
                    st.session_state.sepet.pop(idx)
                    st.rerun()
            else:
                st.info("Sepetiniz boş. Yukarıdan ürün ekleyip 'Sepete Ekle' butonuna basın.")

    # ADIM 5: TEKLİF ÖZETİ, HESAPLAMA VE KAYIT
    with st.container(border=True):
        st.markdown("#### 🧮 Adım 5: Teklif Özeti ve Onay")

        if not st.session_state.get("sepet"):
            st.info("Özet görmek için önce Adım 4'ten sepete ürün ekleyin.")
        else:
            sepet_df = pd.DataFrame(st.session_state.sepet)
            brut_toplam = float((sepet_df["adet"] * sepet_df["satis_fiyati"]).sum())
            toplam_brut_maliyet = float((sepet_df["adet"] * sepet_df["brut_maliyet"]).sum())
            toplam_lojistik = float((sepet_df["adet"] * sepet_df["lojistik_maliyet"]).sum())

            musteri_fiyati = max(brut_toplam - bundle_indirim_tutari, 0.0)

            pos_komisyon_oran = 0.0
            if secilen_banka:
                banka_row = banka_df[banka_df['banka_adi'] == secilen_banka].iloc[0]
                pos_komisyon_oran = float(banka_row[f"taksit_{taksit_sayisi}"])
            pos_komisyon = musteri_fiyati * (pos_komisyon_oran / 100)

            net_kar = musteri_fiyati - toplam_brut_maliyet - toplam_lojistik - pos_komisyon

            m1, m2, m3, m4 = st.columns(4)
            m1.markdown(f"<div class='metric-box'><b>Brüt Toplam</b><br>{brut_toplam:,.2f} ₺</div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-box'><b>Müşteri Fiyatı</b><br>{musteri_fiyati:,.2f} ₺</div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-box'><b>POS Komisyonu</b><br>{pos_komisyon:,.2f} ₺</div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='metric-box'><b>Net Kar</b><br>{net_kar:,.2f} ₺</div>", unsafe_allow_html=True)

            durum = st.selectbox("Teklif Durumu:", ["Beklemede", "Onaylandı", "İptal"])
            urunler_str = ", ".join([f"{s['urun_adi']} x{s['adet']}" for s in st.session_state.sepet])

            c_kaydet, c_wp = st.columns(2)
            with c_kaydet:
                if st.button("💾 Teklifi Kaydet", type="primary"):
                    cursor.execute(
                        """INSERT INTO teklifler
                            (musteri_id, urunler, brut_toplam, bundle_indirimi, musteri_fiyati,
                             lojistik_maliyet, pos_komisyon, net_kar, durum, tarih, sube, personel)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (m_id, urunler_str, brut_toplam, bundle_indirim_tutari, musteri_fiyati,
                         toplam_lojistik, pos_komisyon, net_kar, durum,
                         datetime.now().strftime("%Y-%m-%d %H:%M"), secilen_sube, secilen_personel)
                    )
                    conn.commit()
                    st.success("✅ Teklif başarıyla kaydedildi!")
                    st.session_state.sepet = []
                    st.rerun()
            with c_wp:
                if m_tel:
                    telefon_temiz = str(m_tel).strip().replace(" ", "").replace("-", "")
                    wp_mesaj = f"Merhaba, teklifiniz hazır:\n{urunler_str}\nToplam: {musteri_fiyati:,.2f} TL"
                    wp_link = "https://wa.me/90" + telefon_temiz + "?text=" + urllib.parse.quote(wp_mesaj)
                    st.link_button("📲 WhatsApp'tan Gönder", wp_link)
                else:
                    st.caption("WhatsApp linki için müşteri telefonu gerekli.")

# 5. SAYFA: MERKEZİ CRM & TAKİP PANELİ
elif sayfa == "📊 Merkezi CRM & Takip Panel":
    st.header("📊 Merkezi CRM & Teklif Takip Paneli")
    st.markdown("---")

    teklif_df = pd.read_sql_query("""
        SELECT t.id, m.isim AS musteri, t.urunler, t.brut_toplam, t.bundle_indirimi,
               t.musteri_fiyati, t.lojistik_maliyet, t.pos_komisyon, t.net_kar,
               t.durum, t.tarih, t.sube, t.personel
        FROM teklifler t
        LEFT JOIN musteriler m ON t.musteri_id = m.id
        ORDER BY t.id DESC
    """, conn)

    if teklif_df.empty:
        st.info("Henüz kayıtlı teklif bulunmuyor. 'Teklif Oluştur' sayfasından yeni teklif kaydedebilirsiniz.")
    else:
        c_fil1, c_fil2, c_fil3 = st.columns(3)
        with c_fil1:
            sube_filtre = st.multiselect("Şube Filtrele:", sorted(teklif_df['sube'].dropna().unique()))
        with c_fil2:
            personel_filtre = st.multiselect("Personel Filtrele:", sorted(teklif_df['personel'].dropna().unique()))
        with c_fil3:
            durum_filtre = st.multiselect("Durum Filtrele:", sorted(teklif_df['durum'].dropna().unique()))

        filtre_df = teklif_df.copy()
        if sube_filtre:
            filtre_df = filtre_df[filtre_df['sube'].isin(sube_filtre)]
        if personel_filtre:
            filtre_df = filtre_df[filtre_df['personel'].isin(personel_filtre)]
        if durum_filtre:
            filtre_df = filtre_df[filtre_df['durum'].isin(durum_filtre)]

        m1, m2, m3, m4 = st.columns(4)
        m1.markdown(f"<div class='metric-box'><b>Toplam Teklif</b><br>{len(filtre_df)}</div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-box'><b>Toplam Ciro</b><br>{filtre_df['musteri_fiyati'].sum():,.2f} ₺</div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-box'><b>Toplam Net Kar</b><br>{filtre_df['net_kar'].sum():,.2f} ₺</div>", unsafe_allow_html=True)
        onay_orani = (filtre_df['durum'] == 'Onaylandı').mean() * 100 if len(filtre_df) else 0
        m4.markdown(f"<div class='metric-box'><b>Onay Oranı</b><br>{onay_orani:.1f}%</div>", unsafe_allow_html=True)

        st.markdown("---")
        st.dataframe(filtre_df, use_container_width=True, hide_index=True)

        st.markdown("---")
        st.subheader("🔄 Teklif Durumu Güncelle")
        c_upd1, c_upd2, c_upd3 = st.columns(3)
        with c_upd1:
            secilen_teklif_id = st.selectbox("Teklif ID Seçin:", filtre_df['id'].tolist())
        with c_upd2:
            yeni_durum = st.selectbox("Yeni Durum:", ["Beklemede", "Onaylandı", "İptal"])
        with c_upd3:
            st.write("")
            st.write("")
            if st.button("✅ Durumu Güncelle"):
                cursor.execute("UPDATE teklifler SET durum=? WHERE id=?", (yeni_durum, secilen_teklif_id))
                conn.commit()
                st.success("Durum güncellendi!")
                st.rerun()
