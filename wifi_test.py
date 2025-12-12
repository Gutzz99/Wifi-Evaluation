import speedtest
import time
import pandas as pd
from datetime import datetime
import os
import requests
import socket
import psutil
import subprocess
import re

# --- Speedtest ---
def run_speedtest():
    st = speedtest.Speedtest()
    st.get_best_server()
    download = st.download() / 1_000_000  # Mbps
    upload = st.upload() / 1_000_000
    ping = st.results.ping
    return round(download, 2), round(upload, 2), round(ping, 2)

# --- Ambil SSID aktif dan sinyal ---
def get_ssid_and_signal():
    try:
        output = subprocess.check_output("netsh wlan show interfaces", shell=True).decode()
        ssid_match = re.search(r"^\s*SSID\s*:\s*(.+)$", output, re.MULTILINE)
        signal_match = re.search(r"^\s*Signal\s*:\s*(\d+)%", output, re.MULTILINE)

        ssid = ssid_match.group(1).strip() if ssid_match else "Tidak Diketahui"
        signal_percent = int(signal_match.group(1)) if signal_match else None
        signal_dbm = int((signal_percent / 2) - 100) if signal_percent is not None else None

        return ssid, signal_dbm
    except Exception:
        return "Tidak Diketahui", None

# --- IP Lokal & Publik ---
def get_ip_info():
    try:
        ip_local = socket.gethostbyname(socket.gethostname())
    except:
        ip_local = "Tidak diketahui"

    try:
        ip_public = requests.get("https://api.ipify.org").text
    except:
        ip_public = "Tidak diketahui"

    return ip_local, ip_public

# --- Deteksi VPN/Proxy lokal (interface) ---
def is_using_vpn_or_proxy():
    try:
        interfaces = psutil.net_if_addrs()
        for iface_name in interfaces:
            if "tun" in iface_name.lower() or "tap" in iface_name.lower() or "vpn" in iface_name.lower():
                return f"Ya (interface: {iface_name})"
        return "Tidak"
    except:
        return "Tidak yakin"

# --- Deteksi WARP/Cloudflare dari IP publik ---
def is_cloudflare_warp(ip=None):
    try:
        if not ip:
            ip = requests.get("https://api.ipify.org").text
        resp = requests.get(f"http://ip-api.com/json/{ip}").json()
        org = resp.get("org", "").lower()
        asn = resp.get("as", "").lower()

        if "cloudflare" in org or "cloudflare" in asn:
            return f"Ya (kemungkinan WARP / Cloudflare: {org})"
        return "Tidak"
    except:
        return "Tidak yakin"

# --- Komentar berdasarkan kecepatan unduh ---
def generate_comment(download_speed):
    if download_speed > 80:
        return "Kecepatan super tinggi, ideal untuk streaming 4K/8K, gaming tanpa lag, dan penggunaan banyak perangkat."
    elif 50 <= download_speed <= 80:
        return "Kecepatan tinggi, nyaman untuk streaming HD/4K, gaming online, dan banyak perangkat."
    elif 30 <= download_speed < 50:
        return "Cukup untuk streaming HD dan gaming ringan, performa stabil untuk beberapa perangkat."
    elif 20 <= download_speed < 30:
        return "Masih layak untuk streaming SD/HD, tapi mulai terasa lambat jika digunakan bersamaan."
    elif 10 <= download_speed < 20:
        return "Cukup lambat, streaming HD dan gaming akan terganggu, cocok hanya untuk aktivitas ringan."
    elif 5 <= download_speed < 10:
        return "Kecepatan rendah, hanya cukup untuk browsing dan pesan instan."
    else:
        return "Sangat lambat, tidak direkomendasikan untuk aktivitas online apa pun."

# --- Simpan ke Excel ---
def save_to_excel(data_row, filename="Evaluasi_WiFi.xlsx"):
    headers = [
        "Tanggal", "SSID", "Download (Mbps)", "Upload (Mbps)", "Ping (ms)",
        "Sinyal (dBm)", "IP Lokal", "IP Publik", "VPN/Proxy?", "WARP Aktif?", "Komentar"
    ]

    df = pd.DataFrame(columns=headers)
    new_row = pd.DataFrame([dict(zip(headers, data_row))])

    if os.path.exists(filename):
        df = pd.read_excel(filename)
        df = pd.concat([df, new_row], ignore_index=True)
    else:
        df = new_row

    df.to_excel(filename, index=False)
    print(f"✅ Data berhasil disimpan ke {filename}")

# --- Main ---
def main():
    print("🚀 Mengambil data WiFi...")
    waktu = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ssid, signal = get_ssid_and_signal()
    download, upload, ping = run_speedtest()
    ip_local, ip_public = get_ip_info()
    vpn_status = is_using_vpn_or_proxy()
    warp_status = is_cloudflare_warp(ip_public)
    komentar = generate_comment(download)

    row = [
        waktu, ssid, download, upload, ping,
        signal, ip_local, ip_public, vpn_status, warp_status, komentar
    ]

    save_to_excel(row)

if __name__ == "__main__":
    main()
