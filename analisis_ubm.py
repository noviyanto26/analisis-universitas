import streamlit as st
import pandas as pd
import io

def process_data(df, selected_semester, selected_prodi):
    """
    Fungsi untuk memproses DataFrame input berdasarkan semester dan prodi yang dipilih.
    """
    # Daftar universitas target untuk analisis
    target_universities = [
        'Universitas Ciputra Surabaya',
        'Universitas Katolik Widya Mandala Surabaya',
        'Universitas Kristen Petra',
        'Universitas Surabaya',
        'Universitas Bunda Mulia'
    ]

    # 1. Buat daftar master semua program studi unik dari data mentah
    all_prodi_master = df[['kode_prodi', 'nm_prodi', 'nm_jenj_didik']].drop_duplicates()
    all_prodi_master.rename(columns={'kode_prodi': 'Kode', 'nm_prodi': 'Nama Prodi', 'nm_jenj_didik': 'Jenjang'}, inplace=True)
    
    # Filter daftar master ini berdasarkan prodi yang dipilih pengguna
    master_prodi = all_prodi_master[all_prodi_master['Nama Prodi'].isin(selected_prodi)].sort_values('Kode')

    # 2. Filter data utama berdasarkan semester, universitas, dan prodi yang dipilih
    df_filtered = df[
        (df['id_smt'] == selected_semester) &
        (df['nama_pt'].isin(target_universities)) &
        (df['nm_prodi'].isin(selected_prodi))
    ].copy()

    if df_filtered.empty:
        st.warning(f"Tidak ada data ditemukan untuk kombinasi semester dan prodi yang dipilih.")
        return None

    # 3. Agregasi dan Pivot data yang sudah difilter
    rekap = df_filtered.groupby(['nama_pt', 'kode_prodi'])['jumlah_mhs'].sum().reset_index()
    pivot_rekap = rekap.pivot_table(
        index='kode_prodi',
        columns='nama_pt',
        values='jumlah_mhs'
    )
    pivot_rekap.rename_axis(index={'kode_prodi': 'Kode'}, inplace=True)

    # 4. Gabungkan daftar master prodi dengan data pivot
    final_df = pd.merge(master_prodi, pivot_rekap, on='Kode', how='left').fillna(0)

    # 5. Hitung persentase dan siapkan kolom output
    output_df = final_df[['Kode', 'Nama Prodi', 'Jenjang']].copy()
    
    universities_in_data = [univ for univ in target_universities if univ in final_df.columns]
    total_per_univ = final_df[universities_in_data].sum()

    for univ in target_universities:
        jumlah_col_name = f'{univ}_jumlah'
        percent_col_name = f'{univ}_%'
        
        output_df[jumlah_col_name] = final_df.get(univ, 0).astype(int)
        
        if total_per_univ.get(univ, 0) > 0:
            output_df[percent_col_name] = (output_df[jumlah_col_name] / total_per_univ[univ] * 100).round(2)
        else:
            output_df[percent_col_name] = 0.0

    return output_df


# --- Konfigurasi Aplikasi Streamlit ---
st.set_page_config(page_title="Analisis Data Mahasiswa", layout="wide")
st.title("📊 Aplikasi Analisis dan Rekapitulasi Data Mahasiswa")
st.write(
    "Unggah file data mentah untuk menghasilkan rekapitulasi."
)

uploaded_file = st.file_uploader(
    "Pilih file Excel data mentah",
    type=['xlsx']
)

if uploaded_file is not None:
    try:
        st.info("Membaca file... Harap tunggu.")
        input_df = pd.read_excel(uploaded_file)
        st.success("File berhasil dibaca.")
        
        # --- KONTROL INPUT PENGGUNA ---
        st.subheader("Pengaturan Analisis")
        col1, col2 = st.columns(2)

        with col1:
            semester_list = sorted(input_df['id_smt'].unique(), reverse=True)
            selected_semester = st.selectbox(
                "**1. Pilih Semester (id_smt):**",
                semester_list
            )
        
        with col2:
            prodi_list = sorted(input_df['nm_prodi'].unique())
            selected_prodi = st.multiselect(
                "**2. Pilih Nama Prodi (bisa lebih dari satu):**",
                prodi_list
            )

        # Tombol untuk memulai proses analisis
        if st.button("🚀 3. Buat Rekapitulasi"):
            if not selected_prodi:
                st.warning("Mohon pilih minimal satu Nama Prodi untuk dianalisis.")
            else:
                with st.spinner(f"Memproses data untuk semester {selected_semester}..."):
                    output_df = process_data(input_df, selected_semester, selected_prodi)

                    if output_df is not None:
                        st.success("🎉 Analisis Selesai!")
                        st.subheader(f"Hasil Rekapitulasi Data")
                        
                        # --- Tampilan di Streamlit ---
                        display_df = output_df.copy()
                        display_columns = [('Kode', ''), ('Nama Prodi', ''), ('Jenjang', '')]
                        univ_cols = [c.replace('_jumlah', '') for c in display_df.columns if '_jumlah' in c]
                        for univ in univ_cols:
                            display_columns.append((univ, 'Jumlah Mhs'))
                            display_columns.append((univ, '%'))
                        display_df.columns = pd.MultiIndex.from_tuples(display_columns)
                        st.dataframe(display_df.style.format(precision=2))

                        # --- Fungsi Unduh Excel ---
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                            header_row1 = ['Kode', 'Nama Prodi', 'Jenjang']
                            header_row2 = ['', '', '']
                            for univ in univ_cols:
                                header_row1.extend([univ, ''])
                                header_row2.extend(['Jumlah Mhs', '%'])
                            
                            pd.DataFrame([header_row1]).to_excel(writer, index=False, header=False, startrow=0)
                            pd.DataFrame([header_row2]).to_excel(writer, index=False, header=False, startrow=1)
                            
                            output_df.to_excel(writer, index=False, header=False, startrow=2)
                        
                        output_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 **4. Unduh Hasil Rekapitulasi (Excel)**",
                            data=output_buffer,
                            file_name=f"Rekapitulasi_{selected_semester}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")