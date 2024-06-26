import pandas as pd

# Announcement Template
announcement_data = {
    "Nama_Siswa": ["John Doe"],
    "Email": ["john.doe@example.com"],
    "Subject": ["New Announcement"],
    "Description": ["Description of the announcement"],
    "Link": ["https://example.com"],
    "Phone Number": ["081234567890"]
}
announcement_df = pd.DataFrame(announcement_data)
announcement_df.to_excel("announcement_template.xlsx", index=False)

# Invoice Template
invoice_data = {
    "customer_name": ["John Doe"],
    "customer_email": ["john.doe@example.com"],
    "Grade": ["10"],  # Ensure Grade is a string
    "virtual_account": ["1234567890"],
    "trx_amount": [500000],  # Ensure trx_amount is an integer or float
    "expired_date": ["2023-07-01"],
    "expired_time": ["23:59"],
    "description": ["Invoice description"],
    "link": ["https://example.com"],
    "Subject": ["Invoice Subject"],
    "Phone Number": ["081234567890"]
}
invoice_df = pd.DataFrame(invoice_data)
invoice_df['trx_amount'] = invoice_df['trx_amount'].astype(float)  # Convert trx_amount to float
invoice_df.to_excel("invoice_template.xlsx", index=False)

# Reminder Template
reminder_data = {
    "Nama_Siswa": ["John Doe"],
    "Email": ["john.doe@example.com"],
    "Grade": ["10"],  # Ensure Grade is a string
    "virtual_account": ["1234567890"],
    "bulan_berjalan": [100000],  # Ensure bulan_berjalan is an integer or float
    "Ket_1": ["Description 1"],
    "SPP_30hari": [200000],  # Ensure SPP_30hari is an integer or float
    "Ket_2": ["Description 2"],
    "Denda": [50000],  # Ensure Denda is an integer or float
    "Ket_3": ["Description 3"],
    "Ket_4": ["Description 4"],
    "Total": [350000],  # Ensure Total is an integer or float
    "Subject": ["Reminder Subject"],
    "Phone Number": ["081234567890"]
}
reminder_df = pd.DataFrame(reminder_data)
reminder_df['bulan_berjalan'] = reminder_df['bulan_berjalan'].astype(float)  # Convert to float
reminder_df['SPP_30hari'] = reminder_df['SPP_30hari'].astype(float)  # Convert to float
reminder_df['Denda'] = reminder_df['Denda'].astype(float)  # Convert to float
reminder_df['Total'] = reminder_df['Total'].astype(float)  # Convert to float
reminder_df.to_excel("reminder_template.xlsx", index=False)
