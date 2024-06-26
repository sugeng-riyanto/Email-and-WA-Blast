import pandas as pd

# Announcement Template
announcement_data = {
    "Nama_Siswa": ["Umar bin Khatab"],
    "Email": ["sugeng.riyanto@shb.sch.id"],
    "Subject": ["New Announcement"],
    "Description": ["Description of the announcement"],
    "Link": ["https://www.google.com/?safe=active&ssui=on"],
    "Phone Number": ["081234567890"]
}
announcement_df = pd.DataFrame(announcement_data)
announcement_df.to_excel("announcement_template.xlsx", index=False)

# Invoice Template
invoice_data = {
    "customer_name": ["Umar bin Khatab"],
    "customer_email": ["sugeng.riyanto@shb.sch.id"],
    "Grade": ["10"],
    "virtual_account": ["1234567890"],
    "trx_amount": [500000],
    "expired_date": ["2023-07-01"],
    "expired_time": ["23:59"],
    "description": ["Invoice description"],
    "link": ["https://www.google.com/?safe=active&ssui=on"],
    "Subject": ["Invoice Subject"],
    "Phone Number": ["081234567890"]
}
invoice_df = pd.DataFrame(invoice_data)
invoice_df.to_excel("invoice_template.xlsx", index=False)

# Reminder Template
reminder_data = {
    "Nama_Siswa": ["Umar bin Khatab"],
    "Email": ["sugeng.riyanto@shb.sch.id"],
    "Grade": ["10"],
    "virtual_account": ["1234567890"],
    "bulan_berjalan": [100000],
    "Ket_1": ["Description 1"],
    "SPP_30hari": [200000],
    "Ket_2": ["Description 2"],
    "Denda": [50000],
    "Ket_3": ["Description 3"],
    "Ket_4": ["Description 4"],
    "Total": [350000],
    "Subject": ["Reminder Subject"],
    "Phone Number": ["081234567890"]
}
reminder_df = pd.DataFrame(reminder_data)
reminder_df.to_excel("reminder_template.xlsx", index=False)
