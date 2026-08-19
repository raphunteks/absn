import { NextResponse } from 'next/server';

// =====================================================================
// API DOCS: GET REQUEST
// Akses dokumentasi di: https://absensi.maksaarsyad.xyz/api/wa
// =====================================================================
export async function GET() {
  const apiDocs = {
    app_name: "Sistem Absensi Dept. RKG",
    endpoint: "https://absensi.maksaarsyad.xyz/api/wa",
    method: "POST",
    description: "REST API Gateway untuk Bot WhatsApp (JSON Format).",
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
    available_scenarios: [
      1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 17, 18, 20, 25
    ]
  };

  return NextResponse.json(apiDocs, { status: 200 });
}

// =====================================================================
// POST REQUEST: GENERATE WHATSAPP MESSAGE
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

    // Default Fallback
    const link = data.link || "https://absensi.maksaarsyad.xyz/";
    const tglMulai = data.tanggalMulai || "Belum Diatur";
    const tglAkhir = data.tanggalAkhir || "Belum Diatur";

    let messageText = "";

    switch (scenario) {
      case 1:
        messageText = `🔔 *NOTIFIKASI ABSENSI DIBUKA* 🔔\n\nHalo *${data.namaLengkap}*, sesi absensi untuk *${data.shift}* Dept. RKG hari ini telah resmi *DIBUKA*.\n\n📋 *Detail Sesi Absensi:*\n• Kelompok: *${data.kelompok}*\n• Jam Tepat Waktu: *${data.jamSesi}* WITA\n• Batas Tutup Sesi: *${data.jamTutup}* WITA\n\nYuk, segera lakukan validasi kehadiran Anda (Face ID & Lokasi) sekarang melalui portal resmi kami:\n${link}\n\nSelamat bertugas! 🏥`;
        break;

      case 2:
        messageText = `⚠️ *PENGINGAT TERAKHIR ABSENSI* ⚠️\n\nPanggilan kepada *${data.namaLengkap}*! Sistem mendeteksi Anda *BELUM* melakukan absensi untuk *${data.shift}* hari ini.\n\nWaktu absensi Anda hampir habis. Sesi ini akan ditutup secara permanen pada pukul *${data.jamTutup}* WITA. Jika Anda tidak melakukan absensi setelah jam tersebut, sistem akan otomatis mencatat status Anda sebagai *TIDAK HADIR (ALPHA)*.\n\nMohon segera menuju area batas kampus dan selesaikan absen Anda di sini:\n${link}`;
        break;

      case 3:
        messageText = `✅ *ABSENSI BERHASIL DITERIMA* ✅\n\nTerima kasih *${data.namaLengkap}*! Data kehadiran Anda untuk *${data.shift}* telah berhasil diamankan ke dalam Database Dept. RKG.\n\n📌 *Bukti Rekam Kehadiran:*\n• Waktu Absen: *${data.jamAbsen}* WITA\n• Kelompok: *${data.kelompok}*\n• Keamanan: Tervalidasi (Face ID + GPS)\n\nSistem telah menyimpan log Anda. Silakan cek ringkasan kehadiran di dashboard Anda:\n${link}\n\nSemangat untuk stase hari ini! 🌟`;
        break;

      case 4:
        let strStatus = "";
        if (data.statusAkhir === "Hadir") strStatus = `🟢 *TEPAT WAKTU / HADIR* (Terekam pada: *${data.jamAbsen}* WITA)`;
        else if (data.statusAkhir === "Terlambat") strStatus = `🟡 *TERLAMBAT* (Terekam pada: *${data.jamAbsen}* WITA)`;
        else strStatus = `🔴 *TIDAK HADIR (ALPHA)* (Tidak ada data rekam jejak masuk)`;

        messageText = `📊 *STATUS AKHIR KEHADIRAN* 📊\n\nHalo *${data.namaLengkap}*, batas waktu absensi untuk *${data.shift}* telah resmi *DITUTUP* pada pukul *${data.jamTutup}* WITA.\n\nBerikut adalah rekapan status kehadiran Anda untuk sesi ini:\n\n*Status Anda:*\n${strStatus}\n\nData ini telah disinkronisasikan ke dalam rekapitulasi penilaian Dept. RKG (*${data.kelompok}*). Transparansi data Anda bisa diakses kembali melalui:\n${link}`;
        break;

      case 5:
        messageText = `🛡️ *PEMBERITAHUAN KEAMANAN SISTEM* 🛡️\n\nHalo *${data.namaLengkap}*, \nAkses absensi untuk akun Anda baru saja dilepaskan (Logout) dari perangkat sebelumnya pada pukul *${data.jamAbsen}* WITA.\n\nSaat ini akun Anda dalam status *KOSONG/TIDAK TERIKAT*. Untuk sesi absensi berikutnya, perangkat/HP pertama yang Anda gunakan untuk login akan otomatis menjadi perangkat permanen (Terkunci) untuk akun Anda.\n\nJika ini bukan tindakan Anda atau instruksi dari Admin, segera hubungi Admin Dept. RKG.`;
        break;

      case 6:
        messageText = `👋 *SELAMAT DATANG DI DEPT. RKG!* 🏥\n\nHalo *${data.namaLengkap}*, \nBerdasarkan jadwal akademik, masa stase Anda untuk *${data.kelompok}* akan resmi dimulai pada:\n🗓️ *${tglMulai}* hingga *${tglAkhir}*.\n\nPastikan Anda:\n1. Membuka portal absensi di: ${link}\n2. Melakukan Login perdana untuk mengunci perangkat/HP Anda.\n3. Selalu mengaktifkan fitur Lokasi (GPS) dan Kamera saat melakukan absen.\n\nSelamat bertugas dan tetap jaga kedisiplinan! 🌟`;
        break;

      case 7:
        messageText = `📊 *RAPOR KEHADIRAN DEPT. RKG* 📊\n\nHalo *${data.namaLengkap}*, berikut adalah rekapitulasi kehadiran Anda untuk *${data.kelompok}* sejauh ini:\n\n✅ Tepat Waktu: *${data.totalHadir}* Sesi\n⚠️ Terlambat: *${data.totalTerlambat}* Sesi\n❌ Tidak Hadir (Alpha): *${data.totalAlpha}* Sesi\n\nCek rincian selengkapnya di portal: ${link}`;
        break;

      case 8:
        messageText = `🔄 *INFO PERUBAHAN JADWAL ABSENSI* 🔄\n\nPerhatian *${data.namaLengkap}*, terdapat penyesuaian jadwal pada sistem absensi Dept. RKG untuk *${data.shift}*.\n\nJadwal absensi terbaru Anda:\n⏳ Mulai Buka: *${data.jamSesi.split(' - ')[0]}* WITA\n⏳ Tutup Sesi: *${data.jamTutup}* WITA (Sudah termasuk batas toleransi)\n\nMohon sesuaikan jam kedatangan Anda dengan jadwal terbaru ini. \nTerima kasih atas perhatiannya!`;
        break;

      case 9:
        messageText = `🚨 *SURAT PERINGATAN KEHADIRAN (SISTEM)* 🚨\n\nHalo *${data.namaLengkap}* (*${data.kelompok}*),\n\nSistem mendeteksi bahwa Anda telah mencatatkan *${data.totalAlpha}x* Tidak Hadir (Alpha) dan *${data.totalTerlambat}x* Terlambat selama masa stase ini berjalan.\n\nSesuai regulasi Dept. RKG, tingkat ketidakhadiran ini telah mencapai batas yang memerlukan perhatian khusus. Mohon segera menghadap ke Admin / Koordinator Dept. RKG untuk mengklarifikasi kehadiran Anda.\n\nAbaikan pesan ini jika Anda sudah melakukan konfirmasi.\nCek riwayat Anda di: ${link}`;
        break;

      case 10:
        messageText = `Halo *${data.namaLengkap}*, sesi *${data.shift}* Dept. RKG (*${data.jamSesi}*) sudah dibuka! Batas akhir klik absen adalah jam *${data.jamTutup}*. Segera absen di: ${link} (*${data.kelompok}*).`;
        break;

      case 11:
        messageText = `🛑 *PERINGATAN KEAMANAN AKUN* 🛑\n\nHalo *${data.namaLengkap}*,\nSistem mendeteksi adanya percobaan akses ke akun absensi Anda dari *Perangkat/HP yang tidak dikenal* pada pukul *${data.jamAbsen}* WITA.\n\nSistem telah **MEMBLOKIR** akses tersebut karena akun Anda saat ini sudah terkunci secara aman pada perangkat utama Anda (Fitur Device Fingerprinting aktif).\n\nJika Anda baru saja mengganti HP, harap segera hubungi Admin Dept. RKG untuk meminta pelepasan akses (Unlink Device).\nJaga selalu kerahasiaan Kata Sandi Anda! 🔒`;
        break;

      case 12:
        messageText = `❌ *PEMBATALAN RIWAYAT KEHADIRAN* ❌\n\nPerhatian *${data.namaLengkap}* (*${data.kelompok}*),\nAdmin Dept. RKG baru saja meninjau dan **MENGHAPUS** data kehadiran Anda untuk *${data.shift}* tertanggal *${data.tanggal}*.\n\n*Alasan:* Foto bukti kehadiran tidak valid / Indikasi ketidaksesuaian data.\nStatus kehadiran Anda pada sesi tersebut saat ini dikembalikan menjadi *TIDAK HADIR*.\n\nMohon untuk selalu menggunakan foto *selfie real-time* dan mematuhi tata tertib absensi Dept. RKG.`;
        break;

      case 13:
        messageText = `📍 *PEMBARUAN TITIK LOKASI ABSENSI* 📍\n\nHalo rekan-rekan mahasiswa Dept. RKG!\nTerdapat pembaruan pada titik pusat radar Lokasi Absensi (Geofence) yang berlaku mulai hari ini.\n\n• Titik Lokasi Baru: *${data.lokasiGeofence}*\n• Jangkauan Radar: *${data.radius}* Meter\n\nPastikan Anda berada di area gedung tersebut dan selalu mengizinkan akses GPS pada browser Anda sebelum melakukan absensi di: ${link}`;
        break;

      case 14:
        messageText = `🔑 *PEMBARUAN KATA SANDI AKUN* 🔑\n\nHalo *${data.namaLengkap}*,\nKata sandi (Password) untuk akun absensi Dept. RKG Anda baru saja di-reset oleh Administrator.\n\nBerikut adalah kredensial terbaru Anda:\n👤 NIM: *${data.nim}*\n🔐 Sandi Baru: *${data.password}*\n\nSegera gunakan sandi ini untuk mengakses portal di ${link} dan jangan bagikan kredensial ini kepada siapa pun!`;
        break;

      case 17:
        messageText = `🏁 *HARI TERAKHIR STASE DEPT. RKG* 🏁\n\nHalo *${data.namaLengkap}*, \nHari ini adalah hari terakhir untuk periode stase *${data.kelompok}*.\n\nMohon segera login ke ${link} dan periksa *Rincian Kehadiran* Anda. Pastikan seluruh absensi (Hadir/Terlambat/Alpha) telah terekam dengan benar sebelum Admin menutup buku dan mengekspor laporan akhir (Excel) untuk penilaian.\n\nTerima kasih atas kerja kerasnya selama berada di Dept. RKG! Sukses selalu! ✨`;
        break;

      case 18:
        messageText = `📢 *PENGUMUMAN DEPT. RKG* 📢\nKepada Yth. Seluruh Mahasiswa *${data.kelompok}*,\n\n*${data.pesanCustom}*\n\n---\n_Pesan ini di-generate otomatis oleh Sistem Admin. Harap segera dilaksanakan._`;
        break;

      case 20:
        messageText = `📈 *REKAPITULASI ABSENSI HARIAN DEPT. RKG* 📈\n\nHalo Admin, berikut adalah laporan singkat kehadiran mahasiswa untuk hari ini (*${data.tanggal}*):\n\n👥 *Total Mahasiswa Aktif:* ${data.totalMhs} Entitas\n✅ *Hadir Tepat Waktu:* ${data.totalHadir} Sesi\n⚠️ *Terlambat:* ${data.totalTerlambat} Sesi\n❌ *Tidak Hadir (Alpha):* ${data.totalAlpha} Sesi\n\nSeluruh data telah diamankan ke Cloud. Untuk melihat rincian nama mahasiswa atau mengunduh laporan Excel, silakan akses Dashboard Utama: ${link}`;
        break;

      case 25:
        messageText = `🎉 *SELAMAT DATANG DI DEPT. RKG!* 🎉\n\nHalo *${data.namaLengkap}*, selamat bergabung! \nData Anda telah berhasil didaftarkan oleh Administrator ke dalam Sistem Absensi Digital Terpadu Dept. RKG.\n\nBerikut adalah rincian informasi dan kredensial akses Anda:\n👤 *Nama:* ${data.namaLengkap}\n🆔 *NIM:* ${data.nim}\n👥 *Kelompok:* ${data.kelompok}\n🗓️ *Masa Stase:* ${tglMulai} - ${tglAkhir}\n🔑 *Kata Sandi:* ${data.password}\n\nAgar Anda dapat mulai melakukan absensi, ikuti langkah wajib berikut:\n1. Buka portal resmi absensi di: ${link}\n2. Klik tombol *"Mulai Absensi"* di layar utama.\n3. Masukkan *NIM* dan *Kata Sandi* Anda dengan benar.\n4. Pastikan Anda *Mengizinkan (Allow)* akses Kamera dan Lokasi (GPS) pada browser HP Anda.\n5. Lakukan absensi perdana Anda sesuai jadwal yang ditentukan.\n\n⚠️ *PENTING:* Perangkat/HP pertama yang Anda gunakan untuk login akan langsung *DIKUNCI (Terikat)* dengan akun Anda. Jangan pernah menitipkan akun ke HP teman! \n\nJaga kerahasiaan kata sandi Anda. Selamat bertugas dan sukses untuk stasenya! 🏥✨`;
        break;

      default:
        return NextResponse.json(
          { success: false, error: "Skenario tidak ditemukan." },
          { status: 404 }
        );
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
