import wbdata
import pandas as pd
from datetime import datetime

# Ülkeleri ve göstergeleri tanımlıyoruz
countries = ["TUR", "USA"]
indicator = {"NY.GDP.MKTP.CD": "GDP (Current USD)"}

# Tarih aralığı
date_range = (datetime(2020, 1, 1), datetime(2024, 1, 1))

# Dosya yolu
output_path = "wbdatafile.xlsx"

try:
    # Önce tüm verileri toplayalım
    all_data = {}
    for country in countries:
        # Veriyi al
        data = wbdata.get_dataframe(indicator, country, date_range)
        if not data.empty:
            # Sayısal formatı düzelt
            data = data.reset_index()
            # GDP kolonunu tam sayıya çevir ve binlik ayraçlarla formatla
            data['GDP (Current USD)'] = data['GDP (Current USD)'].apply(lambda x: '{:,.0f}'.format(x))
            all_data[country] = data
            print(f"{country} için veriler alındı.")
        else:
            print(f"{country} için veri bulunamadı.")

    # Excel'e kaydet
    if all_data:
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            # Her ülke için ayrı sayfa oluştur
            for country, data in all_data.items():
                # Excel'e yazdır
                data.to_excel(writer, sheet_name=country, index=False)
                
                # Excel çalışma kitabını ve sayfayı al
                workbook = writer.book
                worksheet = writer.sheets[country]
                
                # Sayı formatını ayarla
                number_format = workbook.add_format({'num_format': '#,##0'})
                
                # GDP sütununa format uygula (B sütunu)
                worksheet.set_column('B:B', 20, number_format)
                
                # Tarih sütununu genişlet
                worksheet.set_column('A:A', 15)
                
        print(f"Veriler başarıyla '{output_path}' konumuna kaydedildi.")
    else:
        print("Hiçbir veri bulunamadı!")

except PermissionError:
    print("Dosya erişim hatası: Excel dosyası açık olabilir veya yazma izniniz olmayabilir.")
except Exception as e:
    print(f"Bir hata oluştu: {str(e)}")
    print("\nDaha fazla bilgi için:")
    import traceback
    traceback.print_exc()
