import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
import time
import warnings
import webbrowser

def apply_dark_mode():
    dark_mode_css = """
    <style>
    /* Set dark background and text color */
    .css-1d391kg, .css-12oz5g7, .css-1y4p8pa {
        background-color: #0e1117;
        color: #ffffff;
    }
    /* Sidebar background color */
    .css-1d3fmxh {
        background-color: #0e1117;
    }
    /* Adjust text color */
    .css-17eq0hr {
        color: #ffffff;
    }
    </style>
    """
    st.markdown(dark_mode_css, unsafe_allow_html=True)

apply_dark_mode()
# Suppress specific warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

# SMTP configuration
your_name = "Sekolah Harapan Bangsa"
your_email = "shsmodernhill@shb.sch.id"
your_password = "jvvmdgxgdyqflcrf"

server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

# Utility function to check allowed file extensions
ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def send_whatsapp_messages(data, announcement=False, invoice=False, proof_payment=False):
    # Open WhatsApp Web once
    webbrowser.open("https://web.whatsapp.com")
    st.info("Please scan the QR code in the opened WhatsApp Web window.")
    
    # Wait for user to scan QR code and login (increase if needed)
    time.sleep(45)

    for index, row in data.iterrows():
        phone_number = str(row['Phone Number'])
        if not phone_number.startswith('+62'):
            phone_number = f'+62{phone_number.lstrip("0")}'
        
        if announcement:
            message = f"""
            Kepada Yth. Orang Tua/Wali Murid *{row['Nama_Siswa']}*,
            Kami hendak menyampaikan info mengenai:
            *Subject:* {row['Subject']}
            *Description:* {row['Description']}
            *Link:* {row['Link']}
            Terima kasih atas kerjasamanya.
            Admin Sekolah
            
            Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:
            • Ibu Penna (Kasir): https://bit.ly/mspennashb
            • Bapak Supatmin (Admin SMP & SMA): https://bit.ly/wamrsupatminshb4
            """
        elif invoice:
            message = f"""
            Kepada Yth. Orang Tua/Wali Murid *{row['customer_name']}* (Kelas *{row['Grade']}*),
            Kami hendak menyampaikan info mengenai:
            • *Subject:* {row['Subject']}
            • *Batas Tanggal Pembayaran:* {row['expired_date']}
            • *Sebesar:* Rp. {row['trx_amount']:,.2f}
            • Pembayaran via nomor *virtual account* (VA) BNI/Bank: *{row['virtual_account']}*
        Terima kasih atas kerjasamanya.
        Admin Sekolah
        Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:
            • Ibu Penna (Kasir): https://bit.ly/mspennashb
            • Bapak Supatmin (Admin SMP & SMA): https://bit.ly/wamrsupatminshb4
            """
        elif proof_payment:
            message = f"""
            Kepada Yth. Orang Tua/Wali Murid *{row['Nama_Siswa']}* (Kelas *{row['Grade']}*),
            Kami hendak menyampaikan info mengenai SPP:
            • *SPP yang sedang berjalan:* {row['bulan_berjalan']:,.2f} ({row['Ket_1']})
            • *Denda:* {row['Denda']:,.2f} ({row['Ket_3']})
            • *SPP bulan-bulan sebelumnya:* {row['SPP_30hari']:,.2f} ({row['Ket_2']})
            • *Keterangan:* {row['Ket_4']}
            • *Total tagihan:* {row['Total']:,.2f}
            Terima kasih atas kerjasamanya.
            Admin Sekolah
            
            Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:
            • Ibu Penna (Kasir): https://bit.ly/mspennashb
            • Bapak Supatmin (Admin SMP & SMA): https://bit.ly/wamrsupatminshb4
            """
        else:
            continue

        url = f"https://wa.me/{phone_number}?text={webbrowser.quote(message)}"
        webbrowser.open(url)
        st.info(f"Sending message to {phone_number}. Please confirm the send action in the opened WhatsApp window.")
        time.sleep(10)  # Wait for user to confirm the send action

def send_emails(email_list, announcement=False, invoice=False, proof_payment=False):
    for idx, entry in enumerate(email_list):
        if announcement:
            subject = entry['Subject']
            name = entry['Nama_Siswa']
            email = entry['Email']
            description = entry['Description']
            link = entry['Link']
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span><br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai:</p>
            <ul>
                <li><strong>Subject:</strong> {subject}</li>
                <li><strong>Description:</strong> {description}</li>
                <li><strong>Link:</strong> {link}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
        elif invoice:
            subject = entry['Subject']
            grade = entry['Grade']
            va = entry['virtual_account']
            name = entry['customer_name']
            email = entry['customer_email']
            nominal = "{:,.2f}".format(entry['trx_amount'])
            expired_date = entry['expired_date']
            expired_time = entry['expired_time']
            description = entry['description']
            link = entry['link']
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)<br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai:</p>
            <ul>
                <li><strong>Subject:</strong> {subject}</li>
                <li><strong>Batas Tanggal Pembayaran:</strong> {expired_date}</li>
                <li><strong>Sebesar:</strong> Rp. {nominal}</li>
                <li><strong>Pembayaran via nomor virtual account (VA) BNI/Bank:</strong> {va}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
        elif proof_payment:
            subject = entry['Subject']
            va = entry['virtual_account']
            name = entry['Nama_Siswa']
            email = entry['Email']
            nominal = "{:,.2f}".format(entry['Nominal'])
            periode = entry['Periode']
            description = entry['Description']
            message = f"""
            Kepada Yth.<br>Orang Tua/Wali Murid <span style="color: #007bff;">{name}</span> (Kelas <span style="color: #007bff;">{grade}</span>)<br>
            <p>Salam Hormat,</p>
            <p>Kami hendak menyampaikan info mengenai:</p>
            <ul>
                <li><strong>Subject:</strong> {subject}</li>
                <li><strong>Nominal:</strong> Rp. {nominal}</li>
                <li><strong>Periode:</strong> {periode}</li>
                <li><strong>Description:</strong> {description}</li>
            </ul>
            <p>Terima kasih atas kerjasamanya.</p>
            <p>Admin Sekolah</p>
            <p>Jika ada pertanyaan atau hendak konfirmasi dapat menghubungi:</p>
            <strong>Ibu Penna (Kasir):</strong> <a href='https://bit.ly/mspennashb' style="color: #007bff;">https://bit.ly/mspennashb</a><br>
            <strong>Bapak Supatmin (Admin SMP & SMA):</strong> <a href='https://bit.ly/wamrsupatminshb4' style="color: #007bff;">https://bit.ly/wamrsupatminshb4</a>
            """
        else:
            continue

        msg = MIMEMultipart()
        msg['From'] = your_email
        msg['To'] = email
        msg['Subject'] = subject
        msg.attach(MIMEText(message, 'html'))

        server.sendmail(your_email, email, msg.as_string())
        st.success(f"Email sent to {email}")

    server.quit()

# Main function
def main():
    st.title('Email and WhatsApp Blast System')
    st.sidebar.title('Actions')

    # Upload Excel file
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        if allowed_file(uploaded_file.name):
            st.success('File uploaded successfully.')
            # Load data
            data = pd.read_excel(uploaded_file)
            st.dataframe(data)

            # Email Blast
            if st.sidebar.button('Send Email Blast'):
                with st.spinner('Sending emails...'):
                    send_emails(data.to_dict('records'), announcement=True)

            # WhatsApp Blast
            if st.sidebar.button('Send WhatsApp Blast'):
                with st.spinner('Sending WhatsApp messages...'):
                    send_whatsapp_messages(data, announcement=True)

        else:
            st.error("Please upload a valid Excel file (.xlsx).")

if __name__ == "__main__":
    main()
