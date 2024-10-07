import csv
import tkinter as tk
from tkinter import filedialog, messagebox
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from PIL import Image, ImageDraw, ImageFont

# Browse file yang akan digunakan untuk mengirim Email
def browse_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("CSV files", "*.csv"), ("all files", "*.*")))
    entry_file.delete(0, tk.END)
    entry_file.insert(0, filename)

# Membuat gambar slip gaji
def generate_salary_slip(employee_data):
    # Create a blank image with white background
    image = Image.new("RGB", (500, 200), "white")
    draw = ImageDraw.Draw(image)

    # Memuat font yang digunakan
    font_path = "Lato-Regular.ttf"  # Path to your font file
    font = ImageFont.truetype(font_path, 18)
    font_body = ImageFont.truetype(font_path, 15)

    # Menulis text pada gambar
    draw.text((200, 10), "SLIP GAJI", fill="black", font=font)
    draw.text((50, 30), f"Nama: {employee_data['Nama']} ({employee_data['NIP']})", fill="black", font=font)
    draw.text((50, 50), f"Bulan: {employee_data['Bulan']}", fill="black", font=font)
    draw.text((50, 70), "---------------------------------------------------------------", fill="black", font=font)
    draw.text((50, 100), f"Gaji Pokok: Rp {employee_data['Gaji']}", fill="black", font=font_body)
    draw.text((50, 120), f"Tunjangan Kinerja: Rp {employee_data['Tunjangan Kinerja']}", fill="black", font=font_body)
    draw.text((50, 140), "----------------------", fill="black", font=font_body) 
    draw.text((50, 160), f"Total Gaji: Rp {employee_data['Total Gaji']}", fill="black", font=font_body)  
    
    # Menyimpan gambar
    image.save(f"slip_gaji_{employee_data['Nama']}.png")
    return f"slip_gaji_{employee_data['Nama']}.png"

# Fungsi untuk mengecek koneksi email
def check_email_connection(sender_email, password):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        server.quit()
        return True
    except smtplib.SMTPAuthenticationError:
        return False

# Mengirim Email
def send_emails():
    sender_email = entry_sender_email.get()
    password = entry_password.get()
    message = text_message.get("1.0", tk.END)
    file_path = entry_file.get()
    
    # Mengecek koneksi email
    if not check_email_connection(sender_email, password):
        messagebox.showerror("Error", "Email atau password salah. Gagal terhubung ke server.")
        return

    # Membaca data karyawan dari file CSV yang diupload
    employee_data = []
    with open(file_path, 'r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            employee_data.append(row)
    
    for employee in employee_data:
        receiver_email = employee["Email"]
        
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = "Slip Gaji Bulan Ini"
        
        body = message.replace("[nama_karyawan]", employee["Nama"])
        msg.attach(MIMEText(body, 'plain'))
        
        # Generate salary slip image
        slip_image = generate_salary_slip(employee)
        
        # Attach slip gaji image
        with open(slip_image, "rb") as file:
            img_data = file.read()
        image = MIMEImage(img_data, name="slip_gaji.png")
        msg.attach(image)
        
        server = smtplib.SMTP('smtp.gmail.com', 587) # Update SMTP server and port
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)
        server.quit()
    
    messagebox.showinfo("Info", "Email slip gaji telah dikirimkan")
    print("Emails sent successfully!")

# Membuat tampilan GUI
root = tk.Tk()
root.title("Sistem Pengiriman Slip Gaji")

label_sender_email = tk.Label(root, text="Email Pengirim:")
label_sender_email.grid(row=0, column=0, sticky="w")

# Input email pengirim
entry_sender_email = tk.Entry(root)
entry_sender_email.grid(row=0, column=1)

label_password = tk.Label(root, text="Password:")
label_password.grid(row=1, column=0, sticky="w")

# Input password email pengirim
entry_password = tk.Entry(root, show="*")
entry_password.grid(row=1, column=1)

label_message = tk.Label(root, text="Pesan Email:")
label_message.grid(row=2, column=0, sticky="w")

# Tulis pesan yang dikirim
text_message = tk.Text(root, height=5, width=50)
text_message.grid(row=2, column=1)

label_file = tk.Label(root, text="Pilih File CSV Data Karyawan:")
label_file.grid(row=3, column=0, sticky="w")

# Input file CSV
entry_file = tk.Entry(root)
entry_file.grid(row=3, column=1)

# Tombol untuk browse file CSV
button_browse = tk.Button(root, text="Browse", command=browse_file)
button_browse.grid(row=3, column=2)

# Tombol untuk mengirim email
button_send = tk.Button(root, text="Kirim Email", command=send_emails)
button_send.grid(row=4, column=1)

if __name__ == '__main__':
    root.mainloop()
