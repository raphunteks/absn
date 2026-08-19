import { NextResponse, NextRequest } from 'next/server';

// =====================================================================
// REDIS CLIENT UNTUK CLOUD SYNC TEMPLATE WA & QUEUE BOT
// =====================================================================
class Redis {
  url: string;
  token: string;

  constructor(config: { url: string; token: string }) {
    this.url = config.url || '';
    this.token = config.token || '';
    if (this.url.endsWith('/')) this.url = this.url.slice(0, -1);
  }

  static fromEnv() {
    let url = process.env.NEXT_PUBLIC_UPSTASH_REDIS_REST_URL || process.env.NEXT_PUBLIC_KV_REST_API_URL || '';
    let token = process.env.NEXT_PUBLIC_UPSTASH_REDIS_REST_TOKEN || process.env.NEXT_PUBLIC_KV_REST_API_TOKEN || '';
    return new Redis({ url, token });
  }

  async get(key: string) {
    if (!this.url || !this.token) return null;
    try {
      const res = await fetch(this.url, { 
        method: 'POST',
        headers: { Authorization: `Bearer ${this.token}`, 'Content-Type': 'application/json' }, 
        body: JSON.stringify(["GET", key]), 
        cache: 'no-store' 
      });
      if (!res.ok) throw new Error('Fetch failed');
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      if (data.result === null || data.result === undefined) return null;
      try {
        return typeof data.result === 'string' ? JSON.parse(data.result) : data.result;
      } catch (e) {
        return data.result; 
      }
    } catch (e) { 
      return null; 
    }
  }

  async set(key: string, value: any) {
    if (!this.url || !this.token) return;
    try {
      const strVal = typeof value === 'string' ? value : JSON.stringify(value);
      const res = await fetch(this.url, { 
        method: 'POST', 
        headers: { Authorization: `Bearer ${this.token}`, 'Content-Type': 'application/json' }, 
        body: JSON.stringify(["SET", key, strVal]) 
      });
      const data = await res.json();
      if (data.error) throw new Error(data.error);
      return data;
    } catch (e) {
      throw e;
    }
  }
}

// =====================================================================
// DAFTAR TEMPLATE DEFAULT LENGKAP (18 SKENARIO)
// =====================================================================
const defaultFormats = [
  { id: 1, title: "Pembukaan Sesi", description: "Dikirim tepat saat jam shift dimulai.", template: "🔔 *NOTIFIKASI ABSENSI DIBUKA* 🔔\n\nHalo *[Nama Lengkap]*, sesi absensi untuk *[Shift]* Dept. RKG hari ini telah resmi *DIBUKA*.\n\n📋 *Detail Sesi Absensi:*\n• Kelompok: *[Kelompok]*\n• Jam Tepat Waktu: *[Jam Sesi]* WITA\n• Batas Tutup Sesi: *[Jam Tutup]* WITA\n\nYuk, segera lakukan validasi kehadiran Anda sekarang melalui portal resmi kami:\n[Link]\n\nSelamat bertugas! 🏥" },
  { id: 2, title: "Pengingat Sisa Waktu", description: "Hanya untuk MHS yang belum absen (Sisa toleransi).", template: "⚠️ *PENGINGAT TERAKHIR ABSENSI* ⚠️\n\nPanggilan kepada *[Nama Lengkap]*! Sistem mendeteksi Anda *BELUM* melakukan absensi untuk *[Shift]* hari ini.\n\nWaktu absensi Anda hampir habis. Sesi ini akan ditutup secara permanen pada pukul *[Jam Tutup]* WITA. Jika Anda tidak melakukan absensi setelah jam tersebut, sistem akan otomatis mencatat status Anda sebagai *TIDAK HADIR (ALPHA)*.\n\nMohon segera menuju area batas kampus dan selesaikan absen Anda di sini:\n[Link]" },
  { id: 3, title: "Berhasil Absen", description: "Real-time saat MHS klik Selesai.", template: "✅ *ABSENSI BERHASIL DITERIMA* ✅\n\nTerima kasih *[Nama Lengkap]*! Data kehadiran Anda untuk *[Shift]* telah berhasil diamankan ke dalam Database Dept. RKG.\n\n📌 *Bukti Rekam Kehadiran:*\n• Waktu Absen: *[Jam Absen]* WITA\n• Kelompok: *[Kelompok]*\n• Keamanan: Tervalidasi (Face ID + GPS)\n\nSistem telah menyimpan log Anda. Silakan cek ringkasan kehadiran di dashboard Anda:\n[Link]\n\nSemangat untuk stase hari ini! 🌟" },
  { id: 4, title: "Rekap Akhir", description: "Status Hadir/Telat/Alpha setelah sesi ditutup.", template: "📊 *STATUS AKHIR KEHADIRAN* 📊\n\nHalo *[Nama Lengkap]*, batas waktu absensi untuk *[Shift]* telah resmi *DITUTUP* pada pukul *[Jam Tutup]* WITA.\n\nBerikut adalah rekapan status kehadiran Anda untuk sesi ini:\n\n*Status Anda:*\n[Status Kehadiran]\n\nData ini telah disinkronisasikan ke dalam rekapitulasi penilaian Dept. RKG (*[Kelompok]*). Transparansi data Anda bisa diakses kembali melalui:\n[Link]" },
  { id: 5, title: "Keamanan Logout", description: "Pelepasan otoritas perangkat HP.", template: "🛡️ *PEMBERITAHUAN KEAMANAN SISTEM* 🛡️\n\nHalo *[Nama Lengkap]*,\nAkses absensi untuk akun Anda baru saja dilepaskan (Logout) dari perangkat sebelumnya pada pukul *[Jam Absen]* WITA.\n\nSaat ini akun Anda dalam status *KOSONG/TIDAK TERIKAT*. Untuk sesi absensi berikutnya, perangkat/HP pertama yang Anda gunakan untuk login akan otomatis menjadi perangkat permanen (Terkunci) untuk akun Anda.\n\nJika ini bukan tindakan Anda atau instruksi dari Admin, segera hubungi Admin Dept. RKG." },
  { id: 6, title: "Welcome Stase", description: "H-1 sebelum tanggal stase dimulai.", template: "👋 *SELAMAT DATANG DI DEPT. RKG!* 🏥\n\nHalo *[Nama Lengkap]*,\nBerdasarkan jadwal akademik, masa stase Anda untuk *[Kelompok]* akan resmi dimulai pada:\n🗓️ *[Tanggal Mulai]* hingga *[Tanggal Akhir]*.\n\nPastikan Anda:\n1. Membuka portal absensi di: [Link]\n2. Melakukan Login perdana untuk mengunci perangkat/HP Anda.\n3. Selalu mengaktifkan fitur Lokasi (GPS) dan Kamera saat melakukan absen.\n\nSelamat bertugas dan tetap jaga kedisiplinan! 🌟" },
  { id: 7, title: "Rapor Mingguan", description: "Persentase kehadiran per minggu.", template: "📊 *RAPOR KEHADIRAN DEPT. RKG* 📊\n\nHalo *[Nama Lengkap]*, berikut adalah rekapitulasi kehadiran Anda untuk *[Kelompok]* sejauh ini:\n\n✅ Tepat Waktu: *[Total Hadir]* Sesi\n⚠️ Terlambat: *[Total Terlambat]* Sesi\n❌ Tidak Hadir (Alpha): *[Total Alpha]* Sesi\n\nCek rincian selengkapnya di portal: [Link]" },
  { id: 8, title: "Perubahan Jadwal", description: "Notif saat admin mengedit jam shift.", template: "🔄 *INFO PERUBAHAN JADWAL ABSENSI* 🔄\n\nPerhatian *[Nama Lengkap]*, terdapat penyesuaian jadwal pada sistem absensi Dept. RKG untuk *[Shift]*.\n\nJadwal absensi terbaru Anda:\n⏳ Mulai Buka: *[Jam Sesi]* WITA\n⏳ Tutup Sesi: *[Jam Tutup]* WITA (Sudah termasuk batas toleransi)\n\nMohon sesuaikan jam kedatangan Anda dengan jadwal terbaru ini.\nTerima kasih atas perhatiannya!" },
  { id: 9, title: "SP Otomatis", description: "Peringatan jika akumulasi Alpha terlalu banyak.", template: "🚨 *SURAT PERINGATAN KEHADIRAN (SISTEM)* 🚨\n\nHalo *[Nama Lengkap]* (*[Kelompok]*),\n\nSistem mendeteksi bahwa Anda telah mencatatkan *[Total Alpha]*x Tidak Hadir (Alpha) dan *[Total Terlambat]*x Terlambat selama masa stase ini berjalan.\n\nSesuai regulasi Dept. RKG, tingkat ketidakhadiran ini telah mencapai batas yang memerlukan perhatian khusus. Mohon segera menghadap ke Admin / Koordinator Dept. RKG untuk mengklarifikasi kehadiran Anda.\n\nAbaikan pesan ini jika Anda sudah melakukan konfirmasi.\nCek riwayat Anda di: [Link]" },
  { id: 10, title: "Pop-up Bukaan", description: "Versi singkat pembukaan sesi.", template: "Halo *[Nama Lengkap]*, sesi *[Shift]* Dept. RKG (*[Jam Sesi]*) sudah dibuka! Batas akhir klik absen adalah jam *[Jam Tutup]*. Segera absen di: [Link] (*[Kelompok]*)." },
  { id: 11, title: "Login Ilegal", description: "Peringatan akses dari HP yang tidak dikenal.", template: "🛑 *PERINGATAN KEAMANAN AKUN* 🛑\n\nHalo *[Nama Lengkap]*,\nSistem mendeteksi adanya percobaan akses ke akun absensi Anda dari *Perangkat/HP yang tidak dikenal* pada pukul *[Jam Absen]* WITA.\n\nSistem telah **MEMBLOKIR** akses tersebut karena akun Anda saat ini sudah terkunci secara aman pada perangkat utama Anda (Fitur Device Fingerprinting aktif).\n\nJika Anda baru saja mengganti HP, harap segera hubungi Admin Dept. RKG untuk meminta pelepasan akses (Unlink Device).\nJaga selalu kerahasiaan Kata Sandi Anda! 🔒" },
  { id: 12, title: "Penghapusan Admin", description: "Peringatan saat absen dibatalkan karena kecurangan.", template: "❌ *PEMBATALAN RIWAYAT KEHADIRAN* ❌\n\nPerhatian *[Nama Lengkap]* (*[Kelompok]*),\nAdmin Dept. RKG baru saja meninjau dan **MENGHAPUS** data kehadiran Anda untuk *[Shift]* tertanggal *[Tanggal]*.\n\n*Alasan:* Foto bukti kehadiran tidak valid / Indikasi ketidaksesuaian data.\nStatus kehadiran Anda pada sesi tersebut saat ini dikembalikan menjadi *TIDAK HADIR*.\n\nMohon untuk selalu menggunakan foto *selfie real-time* dan mematuhi tata tertib absensi Dept. RKG." },
  { id: 13, title: "Update GPS", description: "Notif pergeseran titik Geofence kampus.", template: "📍 *PEMBARUAN TITIK LOKASI ABSENSI* 📍\n\nHalo rekan-rekan mahasiswa Dept. RKG!\nTerdapat pembaruan pada titik pusat radar Lokasi Absensi (Geofence) yang berlaku mulai hari ini.\n\n• Titik Lokasi Baru: *[Nama Lokasi Geofence]*\n• Jangkauan Radar: *[Radius]* Meter\n\nPastikan Anda berada di area gedung tersebut dan selalu mengizinkan akses GPS pada browser Anda sebelum melakukan absensi di: [Link]" },
  { id: 14, title: "Ubah Password", description: "Kredensial baru hasil reset admin.", template: "🔑 *PEMBARUAN KATA SANDI AKUN* 🔑\n\nHalo *[Nama Lengkap]*,\nKata sandi (Password) untuk akun absensi Dept. RKG Anda baru saja di-reset oleh Administrator.\n\nBerikut adalah kredensial terbaru Anda:\n👤 NIM: *[NIM]*\n🔐 Sandi Baru: *[Password]*\n\nSegera gunakan sandi ini untuk mengakses portal di [Link] dan jangan bagikan kredensial ini kepada siapa pun!" },
  { id: 17, title: "Hari Terakhir", description: "Pengingat cross-check di hari terakhir stase.", template: "🏁 *HARI TERAKHIR STASE DEPT. RKG* 🏁\n\nHalo *[Nama Lengkap]*, \nHari ini adalah hari terakhir untuk periode stase *[Kelompok]*.\n\nMohon segera login ke [Link] dan periksa *Rincian Kehadiran* Anda. Pastikan seluruh absensi (Hadir/Terlambat/Alpha) telah terekam dengan benar sebelum Admin menutup buku dan mengekspor laporan akhir (Excel) untuk penilaian.\n\nTerima kasih atas kerja kerasnya selama berada di Dept. RKG! Sukses selalu! ✨" },
  { id: 18, title: "Pesan Broadcast", description: "Pesan custom (pengumuman) dari Admin.", template: "📢 *PENGUMUMAN DEPT. RKG* 📢\nKepada Yth. Seluruh Mahasiswa *[Kelompok]*,\n\n*[Pesan Custom]*\n\n---\n_Pesan ini di-generate otomatis oleh Sistem Admin. Harap segera dilaksanakan._" },
  { id: 20, title: "Rekap Admin", description: "Rangkuman total Hadir/Alpha harian untuk Admin.", template: "📈 *REKAPITULASI ABSENSI HARIAN DEPT. RKG* 📈\n\nHalo Admin, berikut adalah laporan singkat kehadiran mahasiswa untuk hari ini (*[Tanggal]*):\n\n👥 *Total Mahasiswa Aktif:* [Total Mhs] Entitas\n✅ *Hadir Tepat Waktu:* [Total Hadir] Sesi\n⚠️ *Terlambat:* [Total Terlambat] Sesi\n❌ *Tidak Hadir (Alpha):* [Total Alpha] Sesi\n\nSeluruh data telah diamankan ke Cloud. Untuk melihat rincian nama mahasiswa atau mengunduh laporan Excel, silakan akses Dashboard Utama: [Link]" },
  { id: 25, title: "Onboarding Baru", description: "Penyebaran password saat input MHS baru (Excel/Manual).", template: "🎉 *SELAMAT DATANG DI DEPT. RKG!* 🎉\n\nHalo *[Nama Lengkap]*, selamat bergabung! \nData Anda telah berhasil didaftarkan oleh Administrator ke dalam Sistem Absensi Digital Terpadu Dept. RKG.\n\nBerikut adalah rincian informasi dan kredensial akses Anda:\n👤 *Nama:* [Nama Lengkap]\n🆔 *NIM:* [NIM]\n👥 *Kelompok:* [Kelompok]\n🗓️ *Masa Stase:* [Tanggal Mulai] - [Tanggal Akhir]\n🔑 *Kata Sandi:* [Password]\n\nAgar Anda dapat mulai melakukan absensi, ikuti langkah wajib berikut:\n1. Buka portal resmi absensi di: [Link]\n2. Klik tombol \"Mulai Absensi\" di layar utama.\n3. Masukkan *NIM* dan *Kata Sandi* Anda dengan benar.\n4. Pastikan Anda *Mengizinkan (Allow)* akses Kamera dan Lokasi (GPS) pada browser HP Anda.\n5. Lakukan absensi perdana Anda sesuai jadwal yang ditentukan.\n\n⚠️ *PENTING:* Perangkat/HP pertama yang Anda gunakan untuk login akan langsung *DIKUNCI (Terikat)* dengan akun Anda. Jangan pernah menitipkan akun ke HP teman! \n\nJaga kerahasiaan kata sandi Anda. Selamat bertugas dan sukses untuk stasenya! 🏥✨" }
];

// FUNGSI AUTO-SEED: Mengecek apakah template format sudah ada di Redis, jika belum maka masukkan
async function initFormatsInRedis() {
  const redis = Redis.fromEnv();
  const existingFormats = await redis.get('axaxyz_formats');
  
  if (!existingFormats || existingFormats.length === 0) {
    await redis.set('axaxyz_formats', defaultFormats);
    return defaultFormats;
  }
  return existingFormats;
}

// =====================================================================
// GET REQUEST: MENGAMBIL ANTRIAN (BOT) & MENAMPILKAN DOCS (BROWSER)
// =====================================================================
export async function GET(request: NextRequest) {
  // Jika Bot mengirimkan parameter ?action=pull
  const { searchParams } = new URL(request.url);
  const action = searchParams.get('action');

  if (action === 'pull') {
    const redis = Redis.fromEnv();
    const queue = await redis.get('axaxyz_wa_queue') || [];
    return NextResponse.json({ success: true, queue }, { status: 200 });
  }

  // Jika diakses biasa oleh Browser: Tampilkan API Docs
  await initFormatsInRedis();

  const apiDocs = {
    app_name: "Sistem Absensi Dept. RKG",
    endpoint: "https://absensi.maksaarsyad.xyz/api/wa",
    bot_pull_endpoint: "GET https://absensi.maksaarsyad.xyz/api/wa?action=pull",
    method: "POST",
    description: "REST API Gateway All-in-One: Merakit pesan dari template dinamis Redis & Memasukkannya ke dalam Queue untuk di-pull oleh Bot WA.",
    headers: {
      "Content-Type": "application/json"
    },
    payload_example: {
      no_hp: "6281234567890",
      scenario: 18,
      data: {
        namaLengkap: "M. Azhar Arsyad",
        nim: "161202300030",
        kelompok: "Cluster II 2025",
        shift: "Shift Pagi",
        jamSesi: "07:30",
        jamTutup: "08:55",
        jamAbsen: "07:45",
        statusAkhir: "Hadir",
        tanggalMulai: "01/08/2026",
        tanggalAkhir: "31/08/2026",
        password: "123",
        radius: "500",
        lokasiGeofence: "Gedung Rektorat",
        pesanCustom: "Praktikum dialihkan ke Ruang B lantai 2 (Skenario 18)",
        link: "https://absensi.maksaarsyad.xyz/"
      }
    },
    available_scenarios: "Silakan login ke Admin Panel -> Manajemen Format untuk melihat dan mengatur seluruh Skenario ID secara dinamis."
  };

  return NextResponse.json(apiDocs, { status: 200 });
}

// =====================================================================
// POST REQUEST: GENERATE WHATSAPP MESSAGE & PUSH KE QUEUE (REDIS)
// =====================================================================
export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { no_hp, scenario, data } = body;

    if (!no_hp || !scenario || !data) {
      return NextResponse.json(
        { success: false, error: "Payload tidak lengkap. Pastikan mengirim no_hp, scenario, dan data." },
        { status: 400 }
      );
    }

    // Pastikan Format Tersedia (Ambil dari Redis, sekalian Auto-seed)
    const formats = await initFormatsInRedis();

    // Default Variables Fallback untuk mengatasi Error Null
    const link = data.link || "https://absensi.maksaarsyad.xyz/";
    const tglMulai = data.tanggalMulai || "Belum Diatur";
    const tglAkhir = data.tanggalAkhir || "Belum Diatur";

    // Cari Skenario yang Cocok dari Database
    const matchedFormat = formats.find((f: any) => f.id === scenario);

    if (!matchedFormat) {
       return NextResponse.json(
         { success: false, error: `Skenario dengan ID ${scenario} tidak ditemukan dalam Manajemen Format di Admin.` },
         { status: 404 }
       );
    }

    // Proses Replacement String Dinamis (Mencegah crash dengan String conversion)
    let messageText = matchedFormat.template
      .replace(/\[Nama Lengkap\]/g, String(data.namaLengkap || ''))
      .replace(/\[NIM\]/g, String(data.nim || ''))
      .replace(/\[Kelompok\]/g, String(data.kelompok || ''))
      .replace(/\[Shift\]/g, String(data.shift || ''))
      .replace(/\[Jam Sesi\]/g, String(data.jamSesi || ''))
      .replace(/\[Jam Tutup\]/g, String(data.jamTutup || ''))
      .replace(/\[Jam Absen\]/g, String(data.jamAbsen || ''))
      .replace(/\[Tanggal Mulai\]/g, String(tglMulai))
      .replace(/\[Tanggal Akhir\]/g, String(tglAkhir))
      .replace(/\[Tanggal\]/g, String(data.tanggal || ''))
      .replace(/\[Password\]/g, String(data.password || ''))
      .replace(/\[Pesan Custom\]/g, String(data.pesanCustom || ''))
      .replace(/\[Radius\]/g, String(data.radius || ''))
      .replace(/\[Nama Lokasi Geofence\]/g, String(data.lokasiGeofence || ''))
      .replace(/\[Total Mhs\]/g, String(data.totalMhs || '0'))
      .replace(/\[Total Hadir\]/g, String(data.totalHadir || '0'))
      .replace(/\[Total Terlambat\]/g, String(data.totalTerlambat || '0'))
      .replace(/\[Total Alpha\]/g, String(data.totalAlpha || '0'))
      .replace(/\[Link\]/g, String(link));

    // Proses Khusus Variabel Gabungan (Seperti [Status Kehadiran])
    if (data.statusAkhir) {
      let strStatus = "";
      if (data.statusAkhir === "Hadir") strStatus = `🟢 *TEPAT WAKTU / HADIR* (Terekam pada: *${data.jamAbsen}* WITA)`;
      else if (data.statusAkhir === "Terlambat") strStatus = `🟡 *TERLAMBAT* (Terekam pada: *${data.jamAbsen}* WITA)`;
      else strStatus = `🔴 *TIDAK HADIR (ALPHA)* (Tidak ada data rekam jejak masuk)`;
      
      messageText = messageText.replace(/\[Status Kehadiran\]/g, strStatus);
    }

    // PUSH KE MESSAGE QUEUE (REDIS)
    const redis = Redis.fromEnv();
    const currentQueue = await redis.get('axaxyz_wa_queue') || [];
    
    // Pesan baru dengan ID unik agar mudah dihapus oleh Bot nanti
    const newMessage = { 
       id: Math.random().toString(36).substr(2, 9), 
       target_number: String(no_hp), 
       formatted_message: messageText 
    };
    
    currentQueue.push(newMessage);
    await redis.set('axaxyz_wa_queue', currentQueue);

    // Mengembalikan JSON Response Sukses ke Frontend Web
    return NextResponse.json({
      success: true,
      message: "Pesan berhasil dirakit dan dimasukkan ke Antrian (Queue).",
      data: newMessage
    }, { status: 200 });

  } catch (error: any) {
    return NextResponse.json(
      { success: false, error: "Gagal memproses request", details: error.message },
      { status: 500 }
    );
  }
}

// =====================================================================
// DELETE REQUEST: MENGHAPUS PESAN DARI ANTRIAN SETELAH BOT BERHASIL
// =====================================================================
export async function DELETE(request: Request) {
  try {
    const { message_id } = await request.json();
    if (!message_id) return NextResponse.json({ success: false, error: "message_id wajib dikirim dalam payload JSON." }, { status: 400 });

    const redis = Redis.fromEnv();
    let currentQueue = await redis.get('axaxyz_wa_queue') || [];
    
    // Hapus pesan yang sudah dikirim bot berdasarkan ID
    currentQueue = currentQueue.filter((msg: any) => msg.id !== message_id);
    
    await redis.set('axaxyz_wa_queue', currentQueue);

    return NextResponse.json({ success: true, message: `Antrian dengan ID ${message_id} berhasil dihapus.` }, { status: 200 });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
