import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from scipy import stats
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="VaR Monte Carlo Simulator", page_icon="📊", layout="wide")

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

st.markdown('<p class="main-header">📊 Value at Risk Monte Carlo Simulator</p>', unsafe_allow_html=True)
st.markdown('<p class="sub-header">Aplikasi Manajemen Risiko dengan Metode Monte Carlo Simulation</p>', unsafe_allow_html=True)

STOCK_LIST = {
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

tab1, tab2, tab3 = st.tabs(["📥 Download Data Saham", "📤 Upload & Uji Normalitas", "📊 Perhitungan VaR"])

with tab1:
    st.header("📥 Download Data Saham dari Yahoo Finance")
    
    col1, col2 = st.columns(2)
    
    with col1:
        selected_stock = st.selectbox(
            "Pilih Saham",
            options=list(STOCK_LIST.keys()),
            format_func=lambda x: f"{x}.JK - {STOCK_LIST[x]}",
            index=list(STOCK_LIST.keys()).index("JPFA")
        )
        start_date = st.date_input("Tanggal Mulai", value=datetime.now() - timedelta(days=365))
    
    with col2:
        st.write("")
        st.write("")
        end_date = st.date_input("Tanggal Akhir", value=datetime.now())
    
    if st.button("📥 Download Data Saham", type="primary", use_container_width=True):
        try:
            with st.spinner(f'Mengunduh data {selected_stock}.JK...'):
                ticker = f"{selected_stock}.JK"
                data = yf.download(ticker, start=start_date, end=end_date, progress=False)
                
                if data.empty:
                    st.error("❌ Data tidak ditemukan. Pastikan ticker dan tanggal sudah benar.")
                else:
                    data = data.reset_index()
                    close_price = data[['Date', 'Open', 'High', 'Low', 'Close', 'Volume']].copy()
                    close_price.columns = ['Date', 'Open', 'High', 'Low', 'price.close', 'Volume']
                    close_price['Log Return'] = close_price['price.close'].pct_change().apply(
                        lambda x: np.log(1 + x) if pd.notna(x) and x != -1 else np.nan
                    )
                    
                    st.session_state['downloaded_data'] = close_price
                    st.session_state['stock_ticker'] = selected_stock
                    
                    st.markdown(f"""
                    <div class="success-box">
                        <h3>✅ Data Berhasil Diunduh!</h3>
                        <p><strong>Saham:</strong> {selected_stock}.JK - {STOCK_LIST[selected_stock]}</p>
                        <p><strong>Periode:</strong> {start_date} s/d {end_date}</p>
                        <p><strong>Total Data:</strong> {len(close_price)} baris</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.subheader("Preview Data (10 Data Terakhir)")
                    st.dataframe(close_price.tail(10), use_container_width=True)
                    
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
                    
                    st.download_button(
                        label="💾 Download ke Excel",
                        data=output.getvalue(),
                        file_name=f"{selected_stock}_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

with tab2:
    st.header("📤 Upload Data & Uji Normalitas")
    
    st.markdown("""
    <div class="info-box">
        <strong>ℹ️ Informasi Uji Normalitas:</strong>
        <ul>
            <li>Data > 50: Menggunakan uji <strong>Kolmogorov-Smirnov</strong></li>
            <li>Data ≤ 50: Menggunakan uji <strong>Shapiro-Wilk</strong></li>
            <li>Data dianggap normal jika p-value > 0.05</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader("Upload File Excel Data Saham", type=['xlsx', 'xls'], help="File harus memiliki kolom 'Log Return'")
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            
            if 'Log Return' not in df.columns:
                st.error("❌ File harus memiliki kolom 'Log Return'")
            else:
                log_returns = df['Log Return'].dropna()
                n = len(log_returns)
                
                st.success(f"✅ File berhasil diupload! Total data: {n} baris")
                
                st.subheader("Preview Data")
                st.dataframe(df.tail(10), use_container_width=True)
                
                st.subheader("🔬 Hasil Uji Normalitas")
                
                if n > 50:
                    stat, p_value = stats.kstest(log_returns, 'norm', args=(log_returns.mean(), log_returns.std()))
                    test_name = "Kolmogorov-Smirnov"
                else:
                    stat, p_value = stats.shapiro(log_returns)
                    test_name = "Shapiro-Wilk"
                
                is_normal = p_value > 0.05
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Jumlah Data", n)
                with col2:
                    st.metric("Metode Uji", test_name)
                with col3:
                    st.metric("P-Value", f"{p_value:.4f}")
                
                if is_normal:
                    st.markdown(f"""
                    <div class="success-box">
                        <h3>✅ Data Berdistribusi NORMAL</h3>
                        <p><strong>P-Value:</strong> {p_value:.4f} > 0.05</p>
                        <p><strong>Kesimpulan:</strong> Data dapat digunakan untuk perhitungan VaR dengan metode Monte Carlo menggunakan distribusi Normal.</p>
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.markdown(f"""
                    <div class="error-box">
                        <h3>⚠️ Data TIDAK Berdistribusi Normal</h3>
                        <p><strong>P-Value:</strong> {p_value:.4f} ≤ 0.05</p>
                        <p><strong>Kesimpulan:</strong> Data tidak berdistribusi normal.</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.markdown("""
                    <div class="warning-box">
                        <h4>📌 Solusi untuk Data Tidak Normal:</h4>
                        <p>Aplikasi ini akan menggunakan <strong>metode Bootstrap (Empirical Distribution)</strong> untuk simulasi Monte Carlo.</p>
                        <p>Metode ini akan melakukan resampling dari data historis yang ada, sehingga tidak mengasumsikan distribusi normal.</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                st.session_state['uploaded_data'] = df
                st.session_state['log_returns'] = log_returns
                st.session_state['is_normal'] = is_normal
                st.session_state['p_value'] = p_value
                st.session_state['test_name'] = test_name
                
                st.subheader("📊 Statistik Deskriptif")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Mean", f"{log_returns.mean():.7f}")
                    st.metric("Median", f"{log_returns.median():.7f}")
                with col2:
                    st.metric("Std Dev", f"{log_returns.std():.7f}")
                    st.metric("Variance", f"{log_returns.var():.7f}")
                with col3:
                    st.metric("Skewness", f"{stats.skew(log_returns):.7f}")
                    st.metric("Kurtosis", f"{stats.kurtosis(log_returns):.7f}")
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
        
        col1, col2 = st.columns(2)
        with col1:
            v0 = st.number_input("Nilai Investasi Awal (V0) dalam Rupiah", 
                                 min_value=1000000, value=100000000, step=1000000, format="%d")
            st.write(f"**Nilai Investasi:** Rp {v0:,.0f}")
        
        with col2:
            iterations = st.number_input("Jumlah Iterasi Monte Carlo", 
                                         min_value=100, max_value=10000, value=100, step=100)
        
        if is_normal:
            st.markdown("""
            <div class="info-box">
                <strong>📊 Metode:</strong> Monte Carlo dengan distribusi <strong>Normal</strong><br>
                Data berdistribusi normal, menggunakan mean dan standar deviasi untuk simulasi.
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="warning-box">
                <strong>📊 Metode:</strong> Monte Carlo dengan <strong>Bootstrap (Empirical Distribution)</strong><br>
                Data tidak berdistribusi normal, menggunakan resampling dari data historis.
            </div>
            """, unsafe_allow_html=True)

        # ===========================
        # TOMBOL HITUNG VaR
        # ===========================
        if st.button("🚀 Hitung Value at Risk", type="primary", use_container_width=True):

            with st.spinner('Menghitung VaR...'):

                # Parameter
                alphas = [0.01, 0.05, 0.10]
                holding_periods = [1, 5, 20]
                mean_return = log_returns.mean()
                std_return = log_returns.std()

                # Hasil simulasi
                results = []

                for alpha in alphas:
                    for hp in holding_periods:
                        simulations = []

                        for _ in range(iterations):
                            cum_return = 0
                            for _ in range(hp):
                                if is_normal:
                                    daily_return = np.random.normal(mean_return, std_return)
                                else:
                                    daily_return = np.random.choice(log_returns)
                                cum_return += daily_return
                            simulations.append(cum_return)

                        simulations.sort()
                        var_index = int(alpha * iterations)
                        var_value = simulations[var_index]
                        var_amount = abs(var_value * v0)

                        results.append({
                            'Confidence Level': f"{int((1-alpha)*100)}%",
                            'Alpha': f"{int(alpha*100)}%",
                            'Holding Period': f"{hp} hari",
                            'VaR Value': var_value,
                            'Kerugian Maksimal (Rp)': var_amount
                        })

                # Convert to DataFrame
                df_results = pd.DataFrame(results)

                # ======================
                # INTERPRETASI – VARIABEL
                # ======================
                example_99_1 = df_results[(df_results['Alpha'] == '1%') & 
                                          (df_results['Holding Period'] == '1 hari')].iloc[0]

                example_95_1 = df_results[(df_results['Alpha'] == '5%') & 
                                          (df_results['Holding Period'] == '1 hari')].iloc[0]

                example_90_1 = df_results[(df_results['Alpha'] == '10%') & 
                                          (df_results['Holding Period'] == '1 hari')].iloc[0]

                st.success("✅ Perhitungan VaR selesai!")

                # ======================
                # TAMPILKAN HASIL VaR
                # ======================
                st.subheader("📈 Statistik Data")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Mean Log Return", f"{mean_return:.7f}")
                col2.metric("Std Deviation", f"{std_return:.7f}")
                col3.metric("Jumlah Iterasi", iterations)
                col4.metric("Nilai Investasi", f"Rp {v0:,.0f}")

                st.subheader("📊 Hasil Perhitungan VaR")

                df_display = df_results.copy()
                df_display['VaR Value'] = df_display['VaR Value'].apply(lambda x: f"{x:.6f}")
                df_display['Kerugian Maksimal (Rp)'] = df_display['Kerugian Maksimal (Rp)'].apply(
                    lambda x: f"Rp {x:,.0f}"
                )

                st.dataframe(df_display, use_container_width=True)

                # ======================
                # INTERPRETASI HASIL
                # ======================
                st.subheader("📝 Interpretasi Hasil")

                st.markdown(f"""
<div class="info-box">
    <h4>💡 Interpretasi:</h4>
    <p><strong>VaR 99% (1 hari):</strong> Dengan tingkat kepercayaan 99%, kerugian maksimal dalam 1 hari tidak akan melebihi <strong>Rp {example_99_1['Kerugian Maksimal (Rp)']:,.0f}</strong>.</p>
    <p><strong>VaR 95% (1 hari):</strong> Dengan tingkat kepercayaan 95%, kerugian maksimal dalam 1 hari tidak akan melebihi <strong>Rp {example_95_1['Kerugian Maksimal (Rp)']:,.0f}</strong>.</p>
    <p><strong>VaR 90% (1 hari):</strong> Dengan tingkat kepercayaan 90%, kerugian maksimal dalam 1 hari tidak akan melebihi <strong>Rp {example_90_1['Kerugian Maksimal (Rp)']:,.0f}</strong>.</p>
    <hr style="margin-top:10px;margin-bottom:10px;">
    <p><strong>Catatan:</strong></p>
    <ul>
        <li>VaR mengukur kerugian maksimal pada tingkat kepercayaan tertentu.</li>
        <li>Semakin tinggi confidence level, semakin besar nilai VaR.</li>
        <li>Semakin panjang holding period, semakin besar risiko kerugian.</li>
    </ul>
</div>
""", unsafe_allow_html=True)

                # Download Excel
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_results.to_excel(writer, index=False, sheet_name='VaR Results')

                st.download_button(
                    label="💾 Download Hasil VaR ke Excel",
                    data=output.getvalue(),
                    file_name="var_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                    )