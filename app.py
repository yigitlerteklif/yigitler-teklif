import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
import urllib.parse
import io

# Kurumsal Sayfa Yapılandırması
st.set_page_config(layout="wide", page_title="Yiğitler Teklif Programı", page_icon="🏢")

# ============================ YARDIMCI FONKSİYONLAR ============================

def para(x):
    """Türkçe para formatı: 18.000,50 ₺"""
    try:
        s = f"{float(x):,.2f}"
    except (TypeError, ValueError):
        return "-"
    return s.replace(",", "X").replace(".", ",").replace("X", ".") + " ₺"

def sablon_excel_olustur(kolonlar, ornek_satirlar):
    """Verilen kolon başlıkları ve örnek satır(lar)ıyla indirilebilir bir xlsx şablonu üretir."""
    buffer = io.BytesIO()
    df = pd.DataFrame(ornek_satirlar, columns=kolonlar)
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Sablon")
    buffer.seek(0)
    return buffer

def wp_telefon_temizle(tel):
    """Telefonu wa.me formatına çevirir: baştaki 0, +90 veya 90 temizlenir → 90XXXXXXXXXX"""
    t = str(tel).strip().replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    if t.startswith("+"):
        t = t[1:]
    if t.startswith("90") and len(t) == 12:
        t = t[2:]
    t = t.lstrip("0")
    return "90" + t if len(t) == 10 else None  # 5XXXXXXXXX bekleniyor

# Streamlit sürümüne göre genişlik parametresi (use_container_width kaldırılıyor)
try:
    _ST_MAJ, _ST_MIN = (int(x) for x in st.__version__.split(".")[:2])
    ST_YENI = (_ST_MAJ, _ST_MIN) >= (1, 46)
except Exception:
    ST_YENI = False

def genis_dataframe(df, **kw):
    if ST_YENI:
        st.dataframe(df, width="stretch", **kw)
    else:
        st.dataframe(df, use_container_width=True, **kw)

# Özel CSS Tasarımları
st.markdown("""
    <style>
    .block-container {padding-top: 1.5rem; padding-bottom: 1.5rem;}
    .stRadio > div {flex-direction: row; gap: 15px;}
    div[data-testid="stExpander"] {border: 2px solid #ffcc00; border-radius: 8px; background-color: #fffdf2;}
    /* DÜZELTME: koyu temada rakamların kaybolmaması için yazı rengi sabitlendi */
    .metric-box {
        background-color: #f1f5f9;
        color: #0f2c59 !important;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        border: 1px solid #cbd5e1;
        font-size: 16px;
    }
    .metric-box b {color: #334155; font-size: 13px; display:block; margin-bottom:4px;}
    .metric-deger {color: #0f2c59; font-weight: 700; font-size: 20px;}
    .kayit-rozet {display:inline-block; background:#e2e8f0; color:#0f2c59; padding:3px 10px;
                  border-radius:12px; font-size:12px; margin-right:6px; font-weight:600;}
    </style>
""", unsafe_allow_html=True)

# 1. VERİTABANI BAĞLANTISI (V7 - Mutlak yol: hangi klasörden çalıştırılırsa çalıştırılsın aynı DB)
DB_YOLU = os.path.join(os.path.dirname(os.path.abspath(__file__)), "yigitler_bayi_v6.db")
conn = sqlite3.connect(DB_YOLU, check_same_thread=False)
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
cursor.execute('''
    CREATE TABLE IF NOT EXISTS kampanyalar (
        kampanya_adi TEXT PRIMARY KEY, kategori TEXT, gerekli_adet INTEGER, indirim_tutari REAL
    )''')
conn.commit()

def tablo_sayisi(tablo):
    try:
        return conn.execute(f"SELECT COUNT(*) FROM {tablo}").fetchone()[0]
    except Exception:
        return 0

def guvenli_toplu_yukle(tablo, df, gerekli_kolonlar, sayisal_kolonlar, insert_sql, satir_hazirla):
    """Önce sütunları doğrular, sonra tek transaction içinde DELETE+INSERT yapar.
    Herhangi bir hata olursa ROLLBACK → eski veriler korunur."""
    eksik = [k for k in gerekli_kolonlar if k not in df.columns]
    if eksik:
        st.error(f"❌ Excel'de eksik sütun(lar) var: {', '.join(eksik)}. Mevcut veriler SİLİNMEDİ. "
                 f"Lütfen örnek şablonu indirip başlıkları birebir kullanın.")
        return False
    df = df.copy()
    for k in sayisal_kolonlar:
        df[k] = pd.to_numeric(df[k], errors="coerce").fillna(0.0)  # boş hücre = 0 (NaN kâr hesabını bozmasın)
    try:
        cursor.execute(f"DELETE FROM {tablo}")
        for _, row in df.iterrows():
            cursor.execute(insert_sql, satir_hazirla(row))
        conn.commit()
        return True
    except Exception as e:
        conn.rollback()
        st.error(f"❌ Yükleme sırasında hata oluştu, eski veriler geri yüklendi (rollback). Detay: {e}")
        return False

# 2. ÜST KURUMSAL BAŞLIK BARI (kırık logo linkleri kaldırıldı — kendi logonuzu eklemek için
#    aşağıdaki yoruma gerçek bir .png/.svg görsel adresi ya da yerel dosya yazabilirsiniz)
st.markdown("""
    <div style='display: flex; justify-content: center; align-items: center; background-color: #0f2c59; padding: 18px 30px; border-radius: 10px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);'>
        <h1 style='text-align: center; color: white; margin: 0; font-family: sans-serif; font-size: 26px; font-weight: bold; letter-spacing: 1px;'>🏢 YİĞİTLER TEKLİF & CRM YÖNETİMİ</h1>
    </div>
""", unsafe_allow_html=True)
# Logo eklemek isterseniz (örnek):
# st.image("logo.png", width=140)  # app.py ile aynı klasöre logo.png koyun

# YAN MENÜ TASARIMI
st.sidebar.markdown("<h2 style='color:#0f2c59; text-align:center;'>📱 Menü</h2>", unsafe_allow_html=True)
sayfa = st.sidebar.radio("İşlem Yapılacak Ekran:", ["📝 Teklif Oluştur (Satış)", "📊 Merkezi CRM & Takip Panel", "🔄 Merkez - Excel Yükleme Odası"])

# Yan menüde sistemdeki kayıt durumu — "yükledim ama görünmüyor" şüphesini anında giderir
st.sidebar.markdown("---")
st.sidebar.markdown(
    f"<div style='font-size:13px'>"
    f"<span class='kayit-rozet'>📦 Ürün: {tablo_sayisi('urunler')}</span>"
    f"<span class='kayit-rozet'>💳 Banka: {tablo_sayisi('banka_komisyonlari')}</span><br><br>"
    f"<span class='kayit-rozet'>🧑‍💼 Personel: {tablo_sayisi('personeller')}</span>"
    f"<span class='kayit-rozet'>🏪 Şube: {tablo_sayisi('subeler')}</span><br><br>"
    f"<span class='kayit-rozet'>🎁 Kampanya: {tablo_sayisi('kampanyalar')}</span>"
    f"<span class='kayit-rozet'>👤 Müşteri: {tablo_sayisi('musteriler')}</span>"
    f"</div>", unsafe_allow_html=True)
st.sidebar.caption(f"🗄️ Veritabanı: {DB_YOLU}")

# 3. SAYFA: ÇOKLU EXCEL YÜKLEME ODASI
if sayfa == "🔄 Merkez - Excel Yükleme Odası":
    st.header("🔄 Gelişmiş Excel Veri Giriş ve Finans Odası")
    st.markdown("---")

    c_ex1, c_ex2 = st.columns(2)
    with c_ex1:
        st.subheader("📦 1. Ürün Portföyü Yükle")
        st.caption(f"Sistemde kayıtlı ürün: **{tablo_sayisi('urunler')}**")
        st.info("Sütun Başlıkları: 'urun_kodu', 'urun_adi', 'grup1_marka', 'grup2_kategori', 'brut_maliyet', 'fiyat_farki', 'satis_fiyati', 'lojistik_maliyet'")
        st.download_button(
            "📥 Örnek Şablonu İndir",
            data=sablon_excel_olustur(
                ["urun_kodu", "urun_adi", "grup1_marka", "grup2_kategori", "brut_maliyet", "fiyat_farki", "satis_fiyati", "lojistik_maliyet"],
                [["URN001", "Örnek Koltuk Takımı", "Örnek Marka", "Oturma Grubu", 5000, 200, 7500, 150],
                 ["URN002", "Örnek Buzdolabı", "Örnek Marka", "Beyaz Eşya", 8000, 300, 11500, 200]]
            ),
            file_name="urun_sablonu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="urun_sablon_indir"
        )
        urun_dosya = st.file_uploader("Ürün Excel Listesini Seçin", type=["xlsx", "xls"], key="u_key")
        if urun_dosya:
            df_u = pd.read_excel(urun_dosya)
            genis_dataframe(df_u.head(2))
            if st.button("🚀 Ürünleri Güncelle"):
                ok = guvenli_toplu_yukle(
                    "urunler", df_u,
                    ["urun_kodu", "urun_adi", "grup1_marka", "grup2_kategori", "brut_maliyet", "fiyat_farki", "satis_fiyati", "lojistik_maliyet"],
                    ["brut_maliyet", "fiyat_farki", "satis_fiyati", "lojistik_maliyet"],
                    "INSERT OR REPLACE INTO urunler VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    lambda r: (str(r['urun_kodu']).strip(), str(r['urun_adi']).strip(), str(r['grup1_marka']), str(r['grup2_kategori']).strip(),
                               float(r['brut_maliyet']), float(r['fiyat_farki']), float(r['satis_fiyati']), float(r['lojistik_maliyet']))
                )
                if ok:
                    st.success(f"✅ {len(df_u)} ürün güncellendi!")
                    st.rerun()

        st.write("---")
        st.subheader("🧑‍💼 3. Personel (Satış Temsilcisi) Listesi Yükle")
        st.caption(f"Sistemde kayıtlı personel: **{tablo_sayisi('personeller')}**")
        st.info("Sütun Başlığı: 'personel_adi'")
        st.download_button(
            "📥 Örnek Şablonu İndir",
            data=sablon_excel_olustur(["personel_adi"], [["Ahmet Yılmaz"], ["Ayşe Demir"]]),
            file_name="personel_sablonu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="personel_sablon_indir"
        )
        per_dosya = st.file_uploader("Personel Excel Listesini Seçin", type=["xlsx", "xls"], key="p_key")
        if per_dosya:
            df_p = pd.read_excel(per_dosya)
            genis_dataframe(df_p.head(2))
            if st.button("🚀 Personel Listesini Güncelle"):
                ok = guvenli_toplu_yukle(
                    "personeller", df_p, ["personel_adi"], [],
                    "INSERT OR REPLACE INTO personeller VALUES (?)",
                    lambda r: (str(r['personel_adi']).strip(),)
                )
                if ok:
                    st.success("✅ Personel listesi güncellendi!")
                    st.rerun()

    with c_ex2:
        st.subheader("💳 2. Banka Taksit Matrisi Yükle")
        st.caption(f"Sistemde kayıtlı banka: **{tablo_sayisi('banka_komisyonlari')}**")
        st.info("Sütun Başlıkları: 'banka_adi', 'taksit_1', 'taksit_2', ..., 'taksit_12' (komisyon oranları % olarak). Boş bırakılan hücreler 0 kabul edilir.")
        st.download_button(
            "📥 Örnek Şablonu İndir",
            data=sablon_excel_olustur(
                ["banka_adi", "taksit_1", "taksit_2", "taksit_3", "taksit_4", "taksit_5", "taksit_6",
                 "taksit_7", "taksit_8", "taksit_9", "taksit_10", "taksit_11", "taksit_12"],
                [["Örnek Banka", 0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5]]
            ),
            file_name="banka_komisyon_sablonu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="banka_sablon_indir"
        )
        komisyon_dosya = st.file_uploader("Banka Komisyon Excel Listesini Seçin", type=["xlsx", "xls"], key="k_key")
        if komisyon_dosya:
            df_k = pd.read_excel(komisyon_dosya)
            genis_dataframe(df_k.head(2))
            if st.button("🚀 Banka Komisyonlarını Güncelle"):
                taksit_kolonlari = [f"taksit_{i}" for i in range(1, 13)]
                ok = guvenli_toplu_yukle(
                    "banka_komisyonlari", df_k,
                    ["banka_adi"] + taksit_kolonlari, taksit_kolonlari,
                    "INSERT OR REPLACE INTO banka_komisyonlari VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    lambda r: tuple([str(r['banka_adi']).strip()] + [float(r[f"taksit_{i}"]) for i in range(1, 13)])
                )
                if ok:
                    st.success(f"✅ {len(df_k)} banka güncellendi!")
                    st.rerun()

        st.write("---")
        st.subheader("🏪 4. Şube Listesi Yükle")
        st.caption(f"Sistemde kayıtlı şube: **{tablo_sayisi('subeler')}**")
        st.info("Sütun Başlığı: 'sube_adi'")
        st.download_button(
            "📥 Örnek Şablonu İndir",
            data=sablon_excel_olustur(["sube_adi"], [["Merkez Şube"], ["Bornova Şube"]]),
            file_name="sube_sablonu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="sube_sablon_indir"
        )
        sube_dosya = st.file_uploader("Şube Excel Listesini Seçin", type=["xlsx", "xls"], key="s_key")
        if sube_dosya:
            df_s = pd.read_excel(sube_dosya)
            genis_dataframe(df_s.head(2))
            if st.button("🚀 Şube Listesini Güncelle"):
                ok = guvenli_toplu_yukle(
                    "subeler", df_s, ["sube_adi"], [],
                    "INSERT OR REPLACE INTO subeler VALUES (?)",
                    lambda r: (str(r['sube_adi']).strip(),)
                )
                if ok:
                    st.success("✅ Şube listesi güncellendi!")
                    st.rerun()

    st.write("---")
    st.subheader("🎁 5. Bundle Kampanya Listesi Yükle")
    st.caption(f"Sistemde kayıtlı kampanya: **{tablo_sayisi('kampanyalar')}**")
    st.info(
        "Sütun Başlıkları: 'kampanya_adi', 'kategori', 'gerekli_adet', 'indirim_tutari'. "
        "Örnek: Kampanya1 / Beyaz Eşya / 2 / 5000 → '2 li Beyaz Eşya alımına 5000 TL indirim'. "
        "'kategori' alanı ürünlerdeki 'grup2_kategori' ile birebir eşleşmelidir."
    )
    st.download_button(
        "📥 Örnek Şablonu İndir",
        data=sablon_excel_olustur(
            ["kampanya_adi", "kategori", "gerekli_adet", "indirim_tutari"],
            [["Kampanya1", "Beyaz Eşya", 2, 5000],
             ["Kampanya2", "Oturma Grubu", 1, 1500]]
        ),
        file_name="kampanya_sablonu.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="kampanya_sablon_indir"
    )
    kampanya_dosya = st.file_uploader("Kampanya Excel Listesini Seçin", type=["xlsx", "xls"], key="kam_key")
    if kampanya_dosya:
        df_kam = pd.read_excel(kampanya_dosya)
        genis_dataframe(df_kam.head(2))
        if st.button("🚀 Kampanyaları Güncelle"):
            ok = guvenli_toplu_yukle(
                "kampanyalar", df_kam,
                ["kampanya_adi", "kategori", "gerekli_adet", "indirim_tutari"],
                ["gerekli_adet", "indirim_tutari"],
                "INSERT OR REPLACE INTO kampanyalar VALUES (?, ?, ?, ?)",
                lambda r: (str(r['kampanya_adi']).strip(), str(r['kategori']).strip(),
                           int(r['gerekli_adet']), float(r['indirim_tutari']))
            )
            if ok:
                st.success("✅ Kampanyalar güncellendi!")
                st.rerun()

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
                st.warning("Lütfen şube listesi excelini yükleyin. (Yan menüdeki rozetlerden kayıt sayısını görebilirsiniz.)")
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
        musteri_df = pd.read_sql_query("SELECT id, isim, telefon, adres FROM musteriler", conn)
        musteri_tipi = st.radio("Müşteri İşlem Türü:", ["Kayıtlı Müşterilerden Seç", "Yeni Müşteri Tanımla"], label_visibility="collapsed")

        m_id = None   # Müşteri seçilmeden teklif kaydına İZİN VERİLMEZ (sahipsiz teklif engeli)
        m_tel = ""
        c1, c2 = st.columns(2)
        if musteri_tipi == "Yeni Müşteri Tanımla" or musteri_df.empty:
            with c1: m_isim = st.text_input("Müşteri Adı Soyadı / Firma Adı:")
            with c2: m_tel = st.text_input("Telefon Numarası (Örn: 5321234567):")
            m_adres = st.text_area("Teslimat / Fatura Adresi:")
            if st.button("➕ Müşteriyi ve Adresini Kaydet", type="primary"):
                if m_isim and m_tel:
                    cursor.execute("INSERT INTO musteriler (isim, telefon, adres) VALUES (?, ?, ?)", (m_isim.strip(), m_tel.strip(), m_adres.strip()))
                    conn.commit()
                    # Yeni müşteri otomatik seçilsin diye id'sini session'a yazıyoruz
                    st.session_state["son_kayitli_musteri_id"] = cursor.lastrowid
                    st.success(f"✔️ {m_isim} kaydedildi ve otomatik seçildi.")
                    st.rerun()
                else: st.error("Ad ve telefon alanları zorunludur!")
            st.info("ℹ️ Teklif kaydedebilmek için önce müşteriyi kaydedin — kayıt sonrası otomatik seçilir.")
        else:
            with c1:
                # Aynı isimli müşteriler karışmasın diye 'ID — İsim' formatı
                musteri_df["etiket"] = "#" + musteri_df["id"].astype(str) + " — " + musteri_df["isim"].astype(str)
                etiketler = musteri_df["etiket"].tolist()
                varsayilan_idx = 0
                son_id = st.session_state.pop("son_kayitli_musteri_id", None)
                if son_id is not None and (musteri_df["id"] == son_id).any():
                    varsayilan_idx = int(musteri_df.index[musteri_df["id"] == son_id][0])
                secilen_etiket = st.selectbox("Sistemde Kayıtlı Müşteriler:", etiketler, index=varsayilan_idx)
                m_row = musteri_df.loc[musteri_df["etiket"] == secilen_etiket].iloc[0]
                m_id = int(m_row["id"])
                m_tel = m_row["telefon"]
                m_adres = m_row["adres"]
            with c2:
                st.write("")
                st.markdown(f"**📞 İletişim:** {m_tel}")
                st.markdown(f"**📍 Güncel Adres:** {m_adres}")

        # Müşteri değişince sepeti sıfırla (yanlış müşteriye eski sepetle teklif kesilmesin)
        aktif = m_id if m_id is not None else "-yeni-"
        if st.session_state.get("sepet_musteri") not in (None, aktif):
            st.session_state.sepet = []
            st.toast("Müşteri değişti — sepet temizlendi.", icon="🧹")
        st.session_state["sepet_musteri"] = aktif

    # ADIM 3: ÖDEME VE KAMPANYA
    with st.container(border=True):
        st.markdown("#### 💳 Adım 3: Ödeme Yöntemi ve Kampanya Tanımı")
        banka_df = pd.read_sql_query("SELECT * FROM banka_komisyonlari", conn)
        kampanya_df = pd.read_sql_query("SELECT * FROM kampanyalar", conn)
        c_f1, c_f2, c_f3 = st.columns(3)
        with c_f1:
            if banka_df.empty:
                st.warning(f"Lütfen banka komisyon listesi excelini yükleyin. (Şu anki veritabanında {tablo_sayisi('banka_komisyonlari')} banka var → dosya: {os.path.basename(DB_YOLU)})")
                secilen_banka = None
            else:
                secilen_banka = st.selectbox("Ödeme Bankası:", banka_df['banka_adi'].tolist())
        with c_f2:
            taksit_sayisi = st.selectbox("Taksit Sayısı:", list(range(1, 13)))
        with c_f3:
            if kampanya_df.empty:
                secilen_kampanya = None
                manuel_indirim_tutari = st.number_input("Kampanya / Ek İndirim (TL):", min_value=0.0, value=0.0, step=50.0)
            else:
                kampanya_secenekleri = ["Yok (Manuel İndirim Gir)"] + kampanya_df['kampanya_adi'].tolist()
                secim = st.selectbox("Bundle Kampanya:", kampanya_secenekleri)
                if secim == "Yok (Manuel İndirim Gir)":
                    secilen_kampanya = None
                    manuel_indirim_tutari = st.number_input("Kampanya / Ek İndirim (TL):", min_value=0.0, value=0.0, step=50.0)
                else:
                    secilen_kampanya = secim
                    manuel_indirim_tutari = 0.0
        if secilen_kampanya:
            kam_row = kampanya_df[kampanya_df['kampanya_adi'] == secilen_kampanya].iloc[0]
            st.caption(
                f"📌 Koşul: '{kam_row['kategori']}' kategorisinden en az {int(kam_row['gerekli_adet'])} adet ürün → "
                f"{para(kam_row['indirim_tutari'])} indirim. Uygunluk, Adım 5'te sepetinize göre otomatik kontrol edilir."
            )

    # ADIM 4: ÜRÜN SEÇİMİ VE SEPET (ürün koduyla arama + kod görünümü)
    with st.container(border=True):
        st.markdown("#### 📦 Adım 4: Ürün Seçimi")
        urun_df = pd.read_sql_query("SELECT * FROM urunler", conn)

        if "sepet" not in st.session_state:
            st.session_state.sepet = []

        if urun_df.empty:
            st.warning("Lütfen ürün listesi excelini yükleyin.")
        else:
            urun_df["urun_kodu"] = urun_df["urun_kodu"].astype(str)
            urun_df["etiket"] = urun_df["urun_kodu"] + " — " + urun_df["urun_adi"].astype(str)

            arama = st.text_input("🔎 Ürün Ara (ürün kodu veya ürün adı yazın):", placeholder="Örn: URN001 veya Koltuk")
            if arama:
                mask = (urun_df["urun_kodu"].str.contains(arama, case=False, na=False) |
                        urun_df["urun_adi"].astype(str).str.contains(arama, case=False, na=False))
                liste_df = urun_df[mask]
                if liste_df.empty:
                    st.warning(f"'{arama}' ile eşleşen ürün bulunamadı — tüm liste gösteriliyor.")
                    liste_df = urun_df
            else:
                liste_df = urun_df

            c_u1, c_u2, c_u3 = st.columns([3, 1, 1])
            with c_u1:
                secilen_etiket = st.selectbox(f"Ürün Seçin ({len(liste_df)} ürün listeleniyor):", liste_df['etiket'].tolist())
            with c_u2:
                adet = st.number_input("Adet:", min_value=1, value=1, step=1)
            with c_u3:
                st.write("")
                st.write("")
                if st.button("➕ Sepete Ekle"):
                    secilen_kod = secilen_etiket.split(" — ")[0]
                    urun_row = urun_df[urun_df['urun_kodu'] == secilen_kod].iloc[0]
                    # Aynı ürün zaten sepetteyse adet artır (mükerrer satır açma)
                    mevcut = next((s for s in st.session_state.sepet if s["urun_kodu"] == secilen_kod), None)
                    if mevcut:
                        mevcut["adet"] += int(adet)
                    else:
                        st.session_state.sepet.append({
                            "urun_kodu": urun_row['urun_kodu'],
                            "urun_adi": urun_row['urun_adi'],
                            "grup2_kategori": urun_row['grup2_kategori'],
                            "adet": int(adet),
                            "satis_fiyati": float(urun_row['satis_fiyati']),
                            "brut_maliyet": float(urun_row['brut_maliyet']),
                            "lojistik_maliyet": float(urun_row['lojistik_maliyet']),
                        })
                    st.rerun()

            if st.session_state.sepet:
                sepet_df = pd.DataFrame(st.session_state.sepet)
                sepet_df["Ara Toplam"] = sepet_df["adet"] * sepet_df["satis_fiyati"]
                goster = sepet_df[["urun_kodu", "urun_adi", "adet", "satis_fiyati", "Ara Toplam"]].rename(columns={
                    "urun_kodu": "Kod", "urun_adi": "Ürün", "adet": "Adet", "satis_fiyati": "Birim Fiyat"
                })
                genis_dataframe(goster, hide_index=True)

                sil_secimi = st.selectbox(
                    "Sepetten çıkarılacak ürün:",
                    ["--"] + [f"{i}: {s['urun_kodu']} — {s['urun_adi']} (x{s['adet']})" for i, s in enumerate(st.session_state.sepet)]
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

            # Bundle kampanya uygunluk kontrolü (sepetin son haline göre)
            bundle_indirim_tutari = manuel_indirim_tutari
            if secilen_kampanya:
                kam_row = kampanya_df[kampanya_df['kampanya_adi'] == secilen_kampanya].iloc[0]
                gerekli_adet = int(kam_row['gerekli_adet'])
                kategori_adet = int(sepet_df.loc[sepet_df['grup2_kategori'] == kam_row['kategori'], 'adet'].sum())
                if kategori_adet >= gerekli_adet:
                    bundle_indirim_tutari = float(kam_row['indirim_tutari'])
                    st.success(
                        f"🎉 '{secilen_kampanya}' kampanyası uygulandı: '{kam_row['kategori']}' kategorisinden "
                        f"{kategori_adet} adet var (gerekli: {gerekli_adet}) → {para(bundle_indirim_tutari)} indirim."
                    )
                else:
                    bundle_indirim_tutari = 0.0
                    st.warning(
                        f"⚠️ '{secilen_kampanya}' kampanyası için '{kam_row['kategori']}' kategorisinden en az "
                        f"{gerekli_adet} adet gerekiyor. Sepetinizde şu an {kategori_adet} adet var — indirim uygulanmadı."
                    )

            musteri_fiyati = max(brut_toplam - bundle_indirim_tutari, 0.0)

            pos_komisyon_oran = 0.0
            if secilen_banka:
                banka_row = banka_df[banka_df['banka_adi'] == secilen_banka].iloc[0]
                oran = banka_row.get(f"taksit_{taksit_sayisi}", 0.0)
                pos_komisyon_oran = 0.0 if pd.isna(oran) else float(oran)
            pos_komisyon = musteri_fiyati * (pos_komisyon_oran / 100)

            net_kar = musteri_fiyati - toplam_brut_maliyet - toplam_lojistik - pos_komisyon

            m1, m2, m3, m4 = st.columns(4)
            m1.markdown(f"<div class='metric-box'><b>Brüt Toplam</b><span class='metric-deger'>{para(brut_toplam)}</span></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-box'><b>Müşteri Fiyatı</b><span class='metric-deger'>{para(musteri_fiyati)}</span></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-box'><b>POS Komisyonu (%{pos_komisyon_oran:g})</b><span class='metric-deger'>{para(pos_komisyon)}</span></div>", unsafe_allow_html=True)
            kar_renk = "#166534" if net_kar >= 0 else "#b91c1c"
            m4.markdown(f"<div class='metric-box'><b>Net Kar</b><span class='metric-deger' style='color:{kar_renk}'>{para(net_kar)}</span></div>", unsafe_allow_html=True)

            durum = st.selectbox("Teklif Durumu:", ["Beklemede", "Onaylandı", "İptal"])
            urunler_str = ", ".join([f"{s['urun_adi']} x{s['adet']}" for s in st.session_state.sepet])

            c_kaydet, c_wp = st.columns(2)
            with c_kaydet:
                if st.button("💾 Teklifi Kaydet", type="primary"):
                    if m_id is None:
                        st.error("❌ Teklif kaydedilemedi: Önce Adım 2'den bir müşteri seçin veya yeni müşteriyi KAYDEDİN. "
                                 "Sahipsiz (müşterisiz) teklif oluşturulamaz.")
                    else:
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
                wp_no = wp_telefon_temizle(m_tel) if m_tel else None
                if wp_no:
                    wp_mesaj = f"Merhaba, teklifiniz hazır:\n{urunler_str}\nToplam: {para(musteri_fiyati)}"
                    wp_link = "https://wa.me/" + wp_no + "?text=" + urllib.parse.quote(wp_mesaj)
                    st.link_button("📲 WhatsApp'tan Gönder", wp_link)
                elif m_tel:
                    st.caption(f"⚠️ '{m_tel}' geçerli bir cep numarasına çevrilemedi (10 haneli 5XX... bekleniyor).")
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
        m1.markdown(f"<div class='metric-box'><b>Toplam Teklif</b><span class='metric-deger'>{len(filtre_df)}</span></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-box'><b>Toplam Ciro</b><span class='metric-deger'>{para(filtre_df['musteri_fiyati'].sum())}</span></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-box'><b>Toplam Net Kar</b><span class='metric-deger'>{para(filtre_df['net_kar'].sum())}</span></div>", unsafe_allow_html=True)
        onay_orani = (filtre_df['durum'] == 'Onaylandı').mean() * 100 if len(filtre_df) else 0
        m4.markdown(f"<div class='metric-box'><b>Onay Oranı</b><span class='metric-deger'>{onay_orani:.1f}%</span></div>", unsafe_allow_html=True)

        st.markdown("---")
        genis_dataframe(filtre_df, hide_index=True)

        st.markdown("---")
        st.subheader("🔄 Teklif Durumu Güncelle")
        if filtre_df.empty:
            st.info("Filtrelere uyan teklif yok — durum güncellemek için filtreleri genişletin.")
        else:
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
