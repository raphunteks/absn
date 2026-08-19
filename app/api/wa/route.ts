import { NextResponse } from 'next/server';

// Interface untuk Redis
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
}

// =====================================================================
// API DOCS: GET REQUEST
// Akses dokumentasi di: https://absensi.maksaarsyad.xyz/api/wa
// =====================================================================
export async function GET() {
  const apiDocs = {
    app_name: "Sistem Absensi Dept. RKG",
    endpoint: "https://absensi.maksaarsyad.xyz/api/wa",
    method: "POST",
    description: "REST API Gateway untuk Bot WhatsApp. Menggunakan Template Dinamis dari Database (CloudStore).",
    headers: {
      "Content-Type": "application/json"
    },
    payload_example: {
      no_hp: "081234567890",
      scenario: 25,
      data: {
        namaLengkap: "M. Azhar Arsyad",
        nim: "161202300030",
        kelompok: "Cluster II 2025",
        shift: "Shift Pagi",
        jamSesi: "07:30 - 08:45",
        jamTutup: "08:55",
        jamAbsen: "07:45",
        statusAkhir: "Hadir",
        tanggalMulai: "01/08/2026",
        tanggalAkhir: "31/08/2026",
        password: "123",
        radius: "500",
        lokasiGeofence: "Gedung Rektorat",
        pesanCustom: "Praktikum dialihkan ke Ruang B lantai 2",
        link: "https://absensi.maksaarsyad.xyz/"
      }
    },
    available_scenarios: "Diambil secara live dari Manajemen Format di Admin Dashboard."
  };

  return NextResponse.json(apiDocs, { status: 200 });
}

// =====================================================================
// POST REQUEST: GENERATE WHATSAPP MESSAGE (DINAMIS DARI REDIS)
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

    // Default Variables
    const link = data.link || "https://absensi.maksaarsyad.xyz/";
    const tglMulai = data.tanggalMulai || "Belum Diatur";
    const tglAkhir = data.tanggalAkhir || "Belum Diatur";

    // 1. Ambil array Format dari CloudStore (Upstash Redis)
    const redis = Redis.fromEnv();
    const cloudFormats = await redis.get('axaxyz_formats');
    
    // 2. Fallback Default Format jika Redis Kosong / Gagal
    let formats = cloudFormats || [
      { id: 1, template: "🔔 *NOTIFIKASI ABSENSI DIBUKA* 🔔\n\nHalo *[Nama Lengkap]*, sesi absensi untuk *[Shift]* Dept. RKG hari ini telah resmi *DIBUKA*.\n\n📋 *Detail Sesi Absensi:*\n• Kelompok: *[Kelompok]*\n• Jam Tepat Waktu: *[Jam Sesi]* WITA\n• Batas Tutup Sesi: *[Jam Tutup]* WITA\n\nYuk, segera lakukan validasi kehadiran Anda sekarang melalui portal resmi kami:\n[Link]\n\nSelamat bertugas! 🏥" },
      { id: 2, template: "⚠️ *PENGINGAT TERAKHIR ABSENSI* ⚠️\n\nPanggilan kepada *[Nama Lengkap]*! Sistem mendeteksi Anda *BELUM* melakukan absensi untuk *[Shift]* hari ini.\n\nWaktu absensi Anda hampir habis. Sesi ini akan ditutup secara permanen pada pukul *[Jam Tutup]* WITA.\n\nMohon segera selesaikan absen Anda di sini:\n[Link]" },
      { id: 3, template: "✅ *ABSENSI BERHASIL DITERIMA* ✅\n\nTerima kasih *[Nama Lengkap]*! Data kehadiran Anda untuk *[Shift]* telah diamankan ke Database Dept. RKG.\n\n📌 *Bukti Rekam Kehadiran:*\n• Waktu Absen: *[Jam Absen]* WITA\n• Kelompok: *[Kelompok]*\n• Keamanan: Tervalidasi (Face ID + GPS)\n\nSilakan cek ringkasan kehadiran di dashboard Anda:\n[Link]" },
      { id: 25, template: "🎉 *SELAMAT DATANG DI DEPT. RKG!* 🎉\n\nHalo *[Nama Lengkap]*, selamat bergabung! \nData Anda telah didaftarkan ke Sistem Absensi Digital Dept. RKG.\n\nKredensial Akses Anda:\n👤 *Nama:* [Nama Lengkap]\n🆔 *NIM:* [NIM]\n👥 *Kelompok:* [Kelompok]\n🗓️ *Masa Stase:* [Tanggal Mulai] - [Tanggal Akhir]\n🔑 *Kata Sandi:* [Password]\n\nAgar Anda dapat mulai absen, ikuti langkah berikut:\n1. Buka portal di: [Link]\n2. Klik tombol \"Mulai Absensi\" di layar utama.\n3. Masukkan NIM dan Kata Sandi Anda.\n\n⚠️ PENTING: Perangkat/HP pertama yang Anda gunakan login akan DIKUNCI (Terikat) dengan akun Anda.\nJaga kerahasiaan sandi Anda!" }
    ];

    // 3. Cari Skenario yang Cocok
    const matchedFormat = formats.find((f: any) => f.id === scenario);

    if (!matchedFormat) {
       return NextResponse.json(
         { success: false, error: `Skenario dengan ID ${scenario} tidak ditemukan dalam Manajemen Format di Admin.` },
         { status: 404 }
       );
    }

    // 4. Proses Replacement String (Mengubah Variabel Tanda Kurung Siku menjadi Teks Asli)
    let messageText = matchedFormat.template
      .replace(/\[Nama Lengkap\]/g, data.namaLengkap || '')
      .replace(/\[NIM\]/g, data.nim || '')
      .replace(/\[Kelompok\]/g, data.kelompok || '')
      .replace(/\[Shift\]/g, data.shift || '')
      .replace(/\[Jam Sesi\]/g, data.jamSesi || '')
      .replace(/\[Jam Tutup\]/g, data.jamTutup || '')
      .replace(/\[Jam Absen\]/g, data.jamAbsen || '')
      .replace(/\[Tanggal Mulai\]/g, tglMulai)
      .replace(/\[Tanggal Akhir\]/g, tglAkhir)
      .replace(/\[Password\]/g, data.password || '')
      .replace(/\[Link\]/g, link);

    // Variabel Khusus Tambahan (Jika diperlukan di skenario custom admin)
    if (data.statusAkhir) {
      let strStatus = "";
      if (data.statusAkhir === "Hadir") strStatus = `🟢 *TEPAT WAKTU / HADIR* (Terekam pada: *${data.jamAbsen}* WITA)`;
      else if (data.statusAkhir === "Terlambat") strStatus = `🟡 *TERLAMBAT* (Terekam pada: *${data.jamAbsen}* WITA)`;
      else strStatus = `🔴 *TIDAK HADIR (ALPHA)* (Tidak ada data rekam jejak masuk)`;
      
      // Jika admin menaruh "[Status Kehadiran]" di format
      messageText = messageText.replace(/\[Status Kehadiran\]/g, strStatus);
    }

    // Mengembalikan JSON berisi Teks yang sudah dirakit (Formatted String)
    return NextResponse.json({
      success: true,
      target_number: no_hp,
      formatted_message: messageText,
      meta_info: {
        scenario_executed: scenario,
        timestamp: new Date().toISOString()
      }
    }, { status: 200 });

  } catch (error: any) {
    return NextResponse.json(
      { success: false, error: "Gagal memproses request", details: error.message },
      { status: 500 }
    );
  }
}
