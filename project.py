import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from scipy import stats
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="VaR Advanced Simulator", page_icon="📊", layout="wide")

st.markdown("""
<style>
.main-header {font-size: 2.5rem; font-weight: bold; color: #1f2937; margin-bottom: 0.5rem;}
.sub-header {font-size: 1.2rem; color: #6b7280; margin-bottom: 2rem;}
.success-box {padding: 1rem; background-color: #d1fae5; border-left: 4px solid #10b981; border-radius: 0.5rem; margin: 1rem 0; color: black;}
.warning-box {padding: 1rem; background-color: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 0.5rem; margin: 1rem 0; color: black;}
.error-box {padding: 1rem; background-color: #fee2e2; border-left: 4px solid #ef4444; border-radius: 0.5rem; margin: 1rem 0; color: black;}
.info-box {padding: 1rem; background-color: #dbeafe; border-left: 4px solid #3b82f6; border-radius: 0.5rem; margin: 1rem 0; color: black;}
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-header">📊 Value at Risk Advanced Simulator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Aplikasi Manajemen Risiko dengan Monte Carlo & Cornish-Fisher Expansion</p>', unsafe_allow_html=True)

# Daftar saham Indonesia
STOCK_LIST_IDN = {
    "ADRO": "Adaro Energy", "AKRA": "AKR Corporindo", "AMRT": "Sumber Alfaria Trijaya",
    "ANTM": "Aneka Tambang", "ASII": "Astra International", "BBCA": "Bank Central Asia",
    "BBNI": "Bank Negara Indonesia", "BBRI": "Bank Rakyat Indonesia", "BMRI": "Bank Mandiri",
    "BRPT": "Barito Pacific", "CPIN": "Charoen Pokphand Indonesia", "EXCL": "XL Axiata",
    "GOTO": "GoTo Gojek Tokopedia", "ICBP": "Indofood CBP", "INCO": "Vale Indonesia",
    "INDF": "Indofood Sukses Makmur", "INKP": "Indah Kiat Pulp & Paper", "ISAT": "Indosat Ooredoo",
    "ITMG": "Indo Tambangraya Megah", "JPFA": "Japfa Comfeed Indonesia", "KLBF": "Kalbe Farma",
    "MBMA": "Marga Abhinaya Abadi", "MDKA": "Merdeka Copper Gold", "MEDC": "Medco Energi",
    "PGAS": "Perusahaan Gas Negara", "PTBA": "Bukit Asam", "SMGR": "Semen Indonesia",
    "TLKM": "Telkom Indonesia", "UNTR": "United Tractors", "UNVR": "Unilever Indonesia"
}

# Daftar saham internasional populer
STOCK_LIST_INTL = {
    "AAPL": "Apple Inc.", "MSFT": "Microsoft Corporation", "GOOGL": "Alphabet Inc. (Google)",
    "AMZN": "Amazon.com Inc.", "TSLA": "Tesla Inc.", "META": "Meta Platforms (Facebook)",
    "NVDA": "NVIDIA Corporation", "BRK-B": "Berkshire Hathaway", "JPM": "JPMorgan Chase",
    "JNJ": "Johnson & Johnson", "V": "Visa Inc.", "PG": "Procter & Gamble",
    "UNH": "UnitedHealth Group", "MA": "Mastercard", "HD": "Home Depot",
    "DIS": "Walt Disney Company", "BAC": "Bank of America", "NFLX": "Netflix Inc.",
    "ADBE": "Adobe Inc.", "CRM": "Salesforce Inc.", "CSCO": "Cisco Systems",
    "INTC": "Intel Corporation", "PFE": "Pfizer Inc.", "KO": "Coca-Cola Company",
    "PEP": "PepsiCo Inc.", "WMT": "Walmart Inc.", "NKE": "Nike Inc.",
    "MRK": "Merck & Co.", "T": "AT&T Inc.", "VZ": "Verizon Communications"
}

tab1, tab2, tab3 = st.tabs(["📥 Download Data Saham", "📤 Upload & Uji Normalitas", "📊 Perhitungan VaR"])

with tab1:
    st.header("📥 Download Data Saham dari Yahoo Finance")
    
    st.markdown("""
    <div class="info-box">
        <strong>ℹ️ Informasi:</strong>
        <ul>
            <li><strong>Saham Indonesia:</strong> Saat Input Manual gunakan kode saham tanpa .JK karena sistem otomatis mengisi .JK (contoh: BBCA, TLKM)</li>
            <li><strong>Saham Internasional:</strong> Gunakan ticker langsung (contoh: AAPL, MSFT, GOOGL)</li>
            <li>Atau pilih dari daftar preset saham populer yang tersedia</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    # Pilihan jenis pasar
    market_type = st.radio(
        "Pilih Jenis Pasar",
        ["🇮🇩 Saham Indonesia (IDX)", "🌍 Saham Internasional"],
        horizontal=True
    )
    
    # Pilihan mode input
    input_mode = st.radio(
        "Pilih Mode Input",
        ["📋 Pilih dari Daftar Preset", "✍️ Input Manual Ticker"],
        horizontal=True
    )
    
    col1, col2 = st.columns(2)
    
    with col1:
        if input_mode == "📋 Pilih dari Daftar Preset":
            if market_type == "🇮🇩 Saham Indonesia (IDX)":
                selected_stock = st.selectbox(
                    "Pilih Saham Indonesia",
                    options=list(STOCK_LIST_IDN.keys()),
                    format_func=lambda x: f"{x}.JK - {STOCK_LIST_IDN[x]}",
                    index=list(STOCK_LIST_IDN.keys()).index("JPFA")
                )
                ticker_symbol = f"{selected_stock}.JK"
                stock_name = STOCK_LIST_IDN[selected_stock]
            else:
                selected_stock = st.selectbox(
                    "Pilih Saham Internasional",
                    options=list(STOCK_LIST_INTL.keys()),
                    format_func=lambda x: f"{x} - {STOCK_LIST_INTL[x]}",
                    index=list(STOCK_LIST_INTL.keys()).index("AAPL")
                )
                ticker_symbol = selected_stock
                stock_name = STOCK_LIST_INTL[selected_stock]
        else:
            if market_type == "🇮🇩 Saham Indonesia (IDX)":
                manual_ticker = st.text_input(
                    "Masukkan Kode Saham Indonesia (tanpa .JK)",
                    value="BBCA",
                    help="Contoh: BBCA, TLKM, ASII, UNVR"
                ).upper().strip()
                ticker_symbol = f"{manual_ticker}.JK"
                stock_name = manual_ticker
                
                st.info(f"📊 Ticker yang akan diunduh: **{ticker_symbol}**")
            else:
                manual_ticker = st.text_input(
                    "Masukkan Ticker Saham Internasional",
                    value="AAPL",
                    help="Contoh: AAPL, MSFT, TSLA, GOOGL, AMZN"
                ).upper().strip()
                ticker_symbol = manual_ticker
                stock_name = manual_ticker
                
                st.info(f"📊 Ticker yang akan diunduh: **{ticker_symbol}**")
        
        start_date = st.date_input("Tanggal Mulai", value=datetime.now() - timedelta(days=365))
    
    with col2:
        st.write("")
        st.write("")
        end_date = st.date_input("Tanggal Akhir", value=datetime.now())
    
    if st.button("📥 Lihat dan Download Data Saham", type="primary", use_container_width=True):
        try:
            # Validasi ticker tidak kosong
            if not ticker_symbol or ticker_symbol.strip() == "" or ticker_symbol.strip() == ".JK":
                st.error("❌ Ticker saham tidak boleh kosong!")
            else:
                with st.spinner(f'Mengunduh data {ticker_symbol}...'):
                    data = yf.download(ticker_symbol, start=start_date, end=end_date, progress=False)
                    
                    if data.empty:
                        st.error(f"❌ Data tidak ditemukan untuk ticker **{ticker_symbol}**")
                        st.warning("""
                        **Kemungkinan penyebab:**
                        - Ticker salah atau tidak tersedia di Yahoo Finance
                        - Tidak ada data pada periode yang dipilih
                        - Format ticker salah (untuk saham Indonesia jangan ditulis .Jk lagi karena sistem sudah otomatis mengisi .JK)
                        
                        **Contoh ticker yang benar:**
                        - Indonesia: BBCA, TLKM, ASII
                        - Internasional: AAPL, MSFT, GOOGL, TSLA
                        """)
                    else:
                        data = data.reset_index()
                        close_price = data[['Date', 'Open', 'High', 'Low', 'Close', 'Volume']].copy()
                        close_price.columns = ['Date', 'Open', 'High', 'Low', 'price.close', 'Volume']
                        close_price['Log Return'] = close_price['price.close'].pct_change().apply(
                            lambda x: np.log(1 + x) if pd.notna(x) and x != -1 else np.nan
                        )
                        
                        st.session_state['downloaded_data'] = close_price
                        st.session_state['stock_ticker'] = ticker_symbol
                        
                        st.markdown(f"""
                        <div class="success-box">
                            <h3>✅ Data Berhasil Diunduh!</h3>
                            <p><strong>Saham:</strong> {ticker_symbol} - {stock_name}</p>
                            <p><strong>Periode:</strong> {start_date} s/d {end_date}</p>
                            <p><strong>Total Data:</strong> {len(close_price)} baris</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.subheader("Preview Data (Keseluruhan)")
                        st.dataframe(close_price, use_container_width=True, height=400)
                        
                        st.subheader("📈 Statistik Deskriptif Log Return")
                        log_returns = close_price['Log Return'].dropna()
                        
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("Mean", f"{log_returns.mean():.6f}")
                            st.metric("Std Dev", f"{log_returns.std():.6f}")
                        with col2:
                            st.metric("Min", f"{log_returns.min():.6f}")
                            st.metric("Max", f"{log_returns.max():.6f}")
                        with col3:
                            st.metric("Skewness", f"{stats.skew(log_returns):.6f}")
                            st.metric("Kurtosis", f"{stats.kurtosis(log_returns):.6f}")
                        with col4:
                            st.metric("Count", f"{len(log_returns)}")
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            close_price.to_excel(writer, index=False, sheet_name='Data')
                        
                        # Buat nama file yang aman
                        safe_filename = ticker_symbol.replace(".", "_")
                        st.download_button(
                            label="💾 Download ke Excel",
                            data=output.getvalue(),
                            file_name=f"{safe_filename}_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
            st.info("💡 Pastikan ticker yang Anda masukkan benar dan tersedia di Yahoo Finance")

with tab2:
    st.header("📤 Upload Data & Uji Normalitas")
    
    st.markdown("""
    <div class="info-box">
        <strong>ℹ️ Informasi Uji Normalitas:</strong>
        <ul>
            <li>Data > 50: Menggunakan uji <strong>Kolmogorov-Smirnov</strong></li>
            <li>Data ≤ 50: Menggunakan uji <strong>Shapiro-Wilk</strong></li>
            <li>Data dianggap normal jika p-value > α (taraf signifikansi)</li>
            <li>File harus memiliki kolom harga <strong>'Close'</strong>, <strong>'close'</strong>, <strong>'Close Price'</strong>, atau <strong>'price.close'</strong></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    # Pilihan Taraf Signifikansi
    alpha_test = st.selectbox(
        "Pilih Taraf Signifikansi (α) untuk Uji Normalitas",
        options=[0.01, 0.05, 0.10],
        format_func=lambda x: f"{int(x*100)}% (α = {x})",
        index=1  # Default 5%
    )
    
    uploaded_file = st.file_uploader("Upload File Excel Data Saham", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Cari kolom close price dengan berbagai kemungkinan nama
            close_col = None
            possible_names = ['Close', 'close', 'Close Price', 'close price', 'price.close', 
                             'CLOSE', 'Closing Price', 'closing price']
            
            for col_name in possible_names:
                if col_name in df.columns:
                    close_col = col_name
                    break
            
            if close_col is None:
                st.error("❌ File harus memiliki kolom harga penutupan (Close/close/Close Price/price.close)")
                st.info("📋 Kolom yang tersedia: " + ", ".join(df.columns.tolist()))
            else:
                st.success(f"✅ Kolom harga ditemukan: '{close_col}'")
                
                # Hitung Log Return: Ln(Pt / Pt-1)
                df['Log Return'] = df[close_col].pct_change().apply(
                    lambda x: np.log(1 + x) if pd.notna(x) and x != -1 else np.nan
                )
                
                log_returns = df['Log Return'].dropna()
                n = len(log_returns)
                
                st.success(f"✅ File berhasil diupload! Total data log return: {n} baris")
                
                st.subheader("Preview Data (Keseluruhan)")
                display_cols = [col for col in df.columns if col in ['Date', 'date', close_col, 'Log Return']]
                # Gunakan container dengan height untuk scrollable dataframe
                st.dataframe(df[display_cols], use_container_width=True, height=400)
                
                st.subheader("🔬 Hasil Uji Normalitas")
                
                if n > 50:
                    stat, p_value = stats.kstest(log_returns, 'norm', args=(log_returns.mean(), log_returns.std()))
                    test_name = "Kolmogorov-Smirnov"
                else:
                    stat, p_value = stats.shapiro(log_returns)
                    test_name = "Shapiro-Wilk"
                
                is_normal = p_value > alpha_test
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Jumlah Data", n)
                with col2:
                    st.metric("Metode Uji", test_name)
                with col3:
                    st.metric("P-Value", f"{p_value:.4f}")
                with col4:
                    st.metric("Taraf Signifikansi", f"{int(alpha_test*100)}%")
                
                if is_normal:
                    st.markdown(f"""
                    <div class="success-box">
                        <h3>✅ Data Berdistribusi NORMAL</h3>
                        <p><strong>P-Value:</strong> {p_value:.4f} > {alpha_test}</p>
                        <p><strong>Kesimpulan:</strong> Data dapat digunakan untuk perhitungan VaR dengan metode Monte Carlo menggunakan distribusi Normal.</p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="error-box">
                        <h3>⚠️ Data TIDAK Berdistribusi Normal</h3>
                        <p><strong>P-Value:</strong> {p_value:.4f} ≤ {alpha_test}</p>
                        <p><strong>Kesimpulan:</strong> Data tidak berdistribusi normal.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("""
                    <div class="warning-box">
                        <h4>📌 Solusi untuk Data Tidak Normal:</h4>
                        <p>Aplikasi ini akan menggunakan <strong>metode Cornish-Fisher Expansion</strong> untuk memperhitungkan skewness dan kurtosis dalam estimasi VaR.</p>
                        <p>Metode ini lebih akurat untuk data yang tidak berdistribusi normal karena menyesuaikan quantile berdasarkan karakteristik distribusi empiris.</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.session_state['uploaded_data'] = df
                st.session_state['log_returns'] = log_returns
                st.session_state['is_normal'] = is_normal
                st.session_state['p_value'] = p_value
                st.session_state['test_name'] = test_name
                
                st.subheader("📊 Statistik Deskriptif")
                mean_ret = log_returns.mean()
                std_ret = log_returns.std()
                skew_ret = stats.skew(log_returns)
                kurt_ret = stats.kurtosis(log_returns)
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Mean", f"{mean_ret:.7f}")
                    st.metric("Median", f"{log_returns.median():.7f}")
                with col2:
                    st.metric("Std Dev", f"{std_ret:.7f}")
                    st.metric("Variance", f"{log_returns.var():.7f}")
                with col3:
                    st.metric("Skewness", f"{skew_ret:.7f}")
                    st.metric("Kurtosis", f"{kurt_ret:.7f}")
                with col4:
                    st.metric("Min", f"{log_returns.min():.7f}")
                    st.metric("Max", f"{log_returns.max():.7f}")
        except Exception as e:
            st.error(f"❌ Error membaca file: {str(e)}")

with tab3:
    st.header("📊 Perhitungan Value at Risk (VaR)")
    
    if 'log_returns' not in st.session_state:
        st.warning("⚠️ Silakan upload data dan lakukan uji normalitas terlebih dahulu di Tab 2")
    else:
        log_returns = st.session_state['log_returns']
        is_normal = st.session_state['is_normal']
        
        st.subheader("⚙️ Parameter Simulasi")
        
        if is_normal:
            st.markdown("""
            <div class="info-box">
                <strong>📊 Metode:</strong> Monte Carlo dengan distribusi <strong>Normal</strong><br>
                Data berdistribusi normal, menggunakan mean dan standar deviasi untuk simulasi.
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2 = st.columns(2)
            with col1:
                v0 = st.number_input("Nilai Investasi Awal (V0) dalam Rupiah", 
                                     min_value=100000, max_value=9007199254740991, 
                                     value=100000000, step=1000000, format="%d")
                st.write(f"**Nilai Investasi:** Rp {v0:,.0f}")
            
            with col2:
                iterations = st.number_input("Jumlah Iterasi Monte Carlo", 
                                             min_value=5, max_value=1000000, 
                                             value=1000, step=100)
        else:
            st.markdown("""
            <div class="warning-box">
                <strong>📊 Metode:</strong> VaR dengan <strong>Cornish-Fisher Expansion</strong><br>
                Data tidak berdistribusi normal, menggunakan adjusted quantile berdasarkan skewness dan kurtosis.
            </div>
            """, unsafe_allow_html=True)
            
            v0 = st.number_input("Nilai Investasi Awal (V0) dalam Rupiah", 
                                 min_value=100000, max_value=9007199254740991, 
                                 value=100000000, step=1000000, format="%d")
            st.write(f"**Nilai Investasi:** Rp {v0:,.0f}")
            iterations = None  # Tidak perlu iterasi untuk Cornish-Fisher
        
        st.markdown("---")
        st.subheader("📋 Parameter Holding Period & Confidence Level")
        
        # Opsi preset
        st.write("**Preset Standar:**")
        use_preset = st.checkbox("Gunakan preset standar (99%, 95%, 90% untuk 1, 5, 20 hari)", value=True)
        
        # Input manual
        if not use_preset:
            st.write("**Input Manual:**")
            col1, col2 = st.columns(2)
            with col1:
                custom_conf_levels = st.multiselect(
                    "Pilih Confidence Level (%)",
                    options=[99, 95, 90, 85, 80, 75],
                    default=[99, 95, 90],
                    help="Pilih satu atau lebih confidence level"
                )
            with col2:
                custom_hp = st.text_input(
                    "Holding Period (hari, pisahkan dengan koma)",
                    value="1,5,20",
                    help="Contoh: 1,5,10,20"
                )
                try:
                    holding_periods = [int(x.strip()) for x in custom_hp.split(',')]
                except:
                    st.error("Format holding period tidak valid! Gunakan format: 1,5,20")
                    holding_periods = [1, 5, 20]

        # TOMBOL HITUNG VaR
        if st.button("🚀 Hitung Value at Risk", type="primary", use_container_width=True):

            with st.spinner('Menghitung VaR...'):

                # Parameter
                if use_preset:
                    alphas = [0.01, 0.05, 0.10]
                    holding_periods = [1, 5, 20]
                else:
                    alphas = [(100 - cl) / 100 for cl in custom_conf_levels]
                    # holding_periods sudah didefinisikan di atas
                
                mean_return = log_returns.mean()
                std_return = log_returns.std()
                skewness = stats.skew(log_returns)
                kurtosis = stats.kurtosis(log_returns)

                # Hasil simulasi
                results = []

                if is_normal:
                    # METODE MONTE CARLO (NORMAL)
                    for alpha in alphas:
                        for hp in holding_periods:
                            simulations = []

                            for _ in range(iterations):
                                cum_return = 0
                                for _ in range(hp):
                                    daily_return = np.random.normal(mean_return, std_return)
                                    cum_return += daily_return
                                simulations.append(cum_return)

                            simulations.sort()
                            var_index = int(alpha * iterations)
                            var_value = simulations[var_index]
                            var_amount = abs(var_value * v0)

                            results.append({
                                'Metode': 'Monte Carlo (Normal)',
                                'Confidence Level': f"{int((1-alpha)*100)}%",
                                'Alpha': f"{int(alpha*100)}%",
                                'Holding Period': f"{hp} hari",
                                'VaR Value': var_value,
                                'Kerugian Maksimal (Rp)': var_amount
                            })
                else:
                    # METODE CORNISH-FISHER EXPANSION
                    for alpha in alphas:
                        for hp in holding_periods:
                            # Z-score normal
                            z_alpha = stats.norm.ppf(alpha)
                            
                            # Cornish-Fisher adjusted quantile
                            z_cf = (z_alpha + 
                                   (z_alpha**2 - 1) * skewness / 6 + 
                                   (z_alpha**3 - 3*z_alpha) * kurtosis / 24 - 
                                   (2*z_alpha**3 - 5*z_alpha) * skewness**2 / 36)
                            
                            # VaR calculation
                            var_value = mean_return * hp + z_cf * std_return * np.sqrt(hp)
                            var_amount = abs(var_value * v0)

                            results.append({
                                'Metode': 'Cornish-Fisher',
                                'Confidence Level': f"{int((1-alpha)*100)}%",
                                'Alpha': f"{int(alpha*100)}%",
                                'Holding Period': f"{hp} hari",
                                'Z-score (Normal)': z_alpha,
                                'Z-score (CF Adjusted)': z_cf,
                                'VaR Value': var_value,
                                'Kerugian Maksimal (Rp)': var_amount
                            })

                # Convert to DataFrame
                df_results = pd.DataFrame(results)

                # INTERPRETASI – VARIABEL
                try:
                    example_99_1 = df_results[(df_results['Alpha'] == '1%') & 
                                              (df_results['Holding Period'] == '1 hari')].iloc[0] if not df_results[(df_results['Alpha'] == '1%') & (df_results['Holding Period'] == '1 hari')].empty else None
                except:
                    example_99_1 = None

                try:
                    example_95_1 = df_results[(df_results['Alpha'] == '5%') & 
                                              (df_results['Holding Period'] == '1 hari')].iloc[0] if not df_results[(df_results['Alpha'] == '5%') & (df_results['Holding Period'] == '1 hari')].empty else None
                except:
                    example_95_1 = None

                try:
                    example_90_1 = df_results[(df_results['Alpha'] == '10%') & 
                                              (df_results['Holding Period'] == '1 hari')].iloc[0] if not df_results[(df_results['Alpha'] == '10%') & (df_results['Holding Period'] == '1 hari')].empty else None
                except:
                    example_90_1 = None

                st.success("✅ Perhitungan VaR selesai!")

                # TAMPILKAN HASIL VaR
                st.subheader("📈 Statistik Data")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Mean Log Return", f"{mean_return:.7f}")
                col2.metric("Std Deviation", f"{std_return:.7f}")
                col3.metric("Skewness", f"{skewness:.7f}")
                col4.metric("Kurtosis", f"{kurtosis:.7f}")
                
                if is_normal:
                    st.info(f"🔄 **Jumlah Iterasi Monte Carlo:** {iterations:,}")

                st.subheader("📊 Hasil Perhitungan VaR")

                df_display = df_results.copy()
                df_display['VaR Value'] = df_display['VaR Value'].apply(lambda x: f"{x:.6f}")
                df_display['Kerugian Maksimal (Rp)'] = df_display['Kerugian Maksimal (Rp)'].apply(
                    lambda x: f"Rp {x:,.0f}"
                )

                st.dataframe(df_display, use_container_width=True)

                # INTERPRETASI HASIL
                st.subheader("📝 Interpretasi Hasil")
                
                interpretation_text = "<div class='info-box'><h4>💡 Interpretasi VaR:</h4>"
                
                if example_99_1 is not None:
                    loss_99 = example_99_1['Kerugian Maksimal (Rp)']
                    interpretation_text += f"<p><strong>VaR 99% (1 hari):</strong> Dengan tingkat kepercayaan 99%, kerugian maksimal dalam 1 hari tidak akan melebihi <strong>Rp {loss_99:,.0f}</strong>.</p>"
                
                if example_95_1 is not None:
                    loss_95 = example_95_1['Kerugian Maksimal (Rp)']
                    interpretation_text += f"<p><strong>VaR 95% (1 hari):</strong> Dengan tingkat kepercayaan 95%, kerugian maksimal dalam 1 hari tidak akan melebihi <strong>Rp {loss_95:,.0f}</strong>.</p>"
                
                if example_90_1 is not None:
                    loss_90 = example_90_1['Kerugian Maksimal (Rp)']
                    interpretation_text += f"<p><strong>VaR 90% (1 hari):</strong> Dengan tingkat kepercayaan 90%, kerugian maksimal dalam 1 hari tidak akan melebihi <strong>Rp {loss_90:,.0f}</strong>.</p>"
                
                interpretation_text += """
    <hr style="margin-top:10px;margin-bottom:10px;">
    <p><strong>Catatan:</strong></p>
    <ul>
        <li>VaR mengukur kerugian maksimal pada tingkat kepercayaan tertentu dalam kondisi pasar normal.</li>
        <li>Semakin tinggi confidence level (semakin rendah alpha), semakin besar nilai VaR.</li>
        <li>Semakin panjang holding period, semakin besar risiko kerugian potensial.</li>
"""
                
                if is_normal:
                    interpretation_text += "        <li>Monte Carlo mengasumsikan distribusi normal dari return dan menggunakan simulasi acak untuk mengestimasi VaR.</li>\n    </ul>\n</div>"
                else:
                    interpretation_text += "        <li>Cornish-Fisher menyesuaikan estimasi VaR berdasarkan skewness dan kurtosis distribusi, memberikan hasil yang lebih akurat untuk data non-normal.</li>\n    </ul>\n</div>"
                
                st.markdown(interpretation_text, unsafe_allow_html=True)
                
                # KESIMPULAN DAN MASUKAN
                st.subheader("📌 Kesimpulan dan Rekomendasi")
                
                # Analisis risiko berdasarkan hasil
                max_loss = df_results['Kerugian Maksimal (Rp)'].max()
                max_loss_pct = (max_loss / v0) * 100
                
                if max_loss_pct < 5:
                    risk_level = "RENDAH"
                    risk_color = "success-box"
                    recommendation = "Portofolio menunjukkan risiko yang rendah. Investasi ini relatif aman dengan potensi kerugian maksimal di bawah 5% dari nilai investasi."
                elif max_loss_pct < 15:
                    risk_level = "SEDANG"
                    risk_color = "warning-box"
                    recommendation = "Portofolio menunjukkan risiko sedang. Pertimbangkan untuk melakukan diversifikasi atau hedging untuk mengurangi eksposur risiko."
                else:
                    risk_level = "TINGGI"
                    risk_color = "error-box"
                    recommendation = "Portofolio menunjukkan risiko tinggi dengan potensi kerugian melebihi 15%. Sangat disarankan untuk melakukan review strategi investasi dan diversifikasi portofolio."
                
                st.markdown(f"""
<div class="{risk_color}">
    <h4>🎯 Tingkat Risiko: {risk_level}</h4>
    <p><strong>Potensi Kerugian Maksimal:</strong> Rp {max_loss:,.0f} ({max_loss_pct:.2f}% dari investasi)</p>
    <p><strong>Rekomendasi:</strong> {recommendation}</p>
    <hr style="margin-top:10px;margin-bottom:10px;">
    <p><strong>Saran Tambahan:</strong></p>
    <ul>
        <li>Lakukan monitoring berkala terhadap pergerakan harga saham</li>
        <li>Pertimbangkan untuk menetapkan stop-loss sesuai dengan toleransi risiko Anda</li>
        <li>Diversifikasi portofolio untuk mengurangi risiko konsentrasi</li>
        <li>{'Pertimbangkan metode hedging untuk melindungi nilai investasi dari fluktuasi ekstrem' if risk_level == 'TINGGI' else 'Tetap waspada terhadap kondisi pasar yang dapat mempengaruhi nilai investasi'}</li>
    </ul>
</div>
""", unsafe_allow_html=True)

                # Download Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False, sheet_name='VaR Results')
                    
                    # Tambahkan sheet statistik
                    stats_data = {
                        'Statistik': ['Mean', 'Std Dev', 'Skewness', 'Kurtosis', 'Min', 'Max', 'Count'],
                        'Nilai': [mean_return, std_return, skewness, kurtosis, 
                                 log_returns.min(), log_returns.max(), len(log_returns)]
                    }
                    pd.DataFrame(stats_data).to_excel(writer, index=False, sheet_name='Statistik')

                st.download_button(
                    label="💾 Download Hasil VaR ke Excel",
                    data=output.getvalue(),
                    file_name="var_results_complete.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
