"use client";

import React, { useState } from 'react';
import { Terminal, Copy, CheckCircle2, Zap, Send, FileJson, Server, Activity, ArrowRight } from 'lucide-react';

export default function ApiDocs() {
  const [copiedReq, setCopiedReq] = useState(false);
  const [copiedRes, setCopiedRes] = useState(false);

  const reqJson = `{
  "no_hp": "081234567890",
  "scenario": 25,
  "data": {
    "namaLengkap": "M. Azhar Arsyad",
    "nim": "161202300030",
    "kelompok": "Cluster II 2025",
    "shift": "Shift Pagi",
    "jamSesi": "07:30 - 08:45",
    "jamTutup": "08:55",
    "jamAbsen": "07:45",
    "statusAkhir": "Hadir",
    "tanggalMulai": "01/08/2026",
    "tanggalAkhir": "31/08/2026",
    "password": "123",
    "radius": "500",
    "lokasiGeofence": "Gedung Rektorat",
    "pesanCustom": "Pindah ke Lab 2 ya!",
    "link": "https://absensi.maksaarsyad.xyz/"
  }
}`;

  const resJson = `{
  "success": true,
  "target_number": "081234567890",
  "formatted_message": "🎉 *SELAMAT DATANG DI DEPT. RKG!* 🎉\\n\\nHalo *M. Azhar Arsyad*...",
  "meta_info": {
    "scenario_executed": 25,
    "timestamp": "${new Date().toISOString()}"
  }
}`;

  const copyToClipboard = (text: string, type: 'req' | 'res') => {
    navigator.clipboard.writeText(text);
    if (type === 'req') {
      setCopiedReq(true);
      setTimeout(() => setCopiedReq(false), 2000);
    } else {
      setCopiedRes(true);
      setTimeout(() => setCopiedRes(false), 2000);
    }
  };

  const scenarios = [
    { id: 1, title: "Pembukaan Sesi", desc: "Dikirim tepat saat jam shift dimulai." },
    { id: 2, title: "Pengingat Sisa Waktu", desc: "Hanya untuk MHS yang belum absen (Sisa toleransi)." },
    { id: 3, title: "Berhasil Absen", desc: "Real-time saat MHS klik Selesai." },
    { id: 4, title: "Rekap Akhir", desc: "Status Hadir/Telat/Alpha setelah sesi ditutup." },
    { id: 5, title: "Keamanan Logout", desc: "Pelepasan otoritas perangkat HP." },
    { id: 6, title: "Welcome Stase", desc: "H-1 sebelum tanggal stase dimulai." },
    { id: 7, title: "Rapor Mingguan", desc: "Persentase kehadiran per minggu." },
    { id: 8, title: "Perubahan Jadwal", desc: "Notif saat admin mengedit jam shift." },
    { id: 9, title: "SP Otomatis", desc: "Peringatan jika akumulasi Alpha terlalu banyak." },
    { id: 10, title: "Pop-up Bukaan", desc: "Versi singkat pembukaan sesi." },
    { id: 11, title: "Login Ilegal", desc: "Peringatan akses dari HP yang tidak dikenal." },
    { id: 12, title: "Penghapusan Admin", desc: "Peringatan saat absen dibatalkan karena kecurangan." },
    { id: 13, title: "Update GPS", desc: "Notif pergeseran titik Geofence kampus." },
    { id: 14, title: "Ubah Password", desc: "Kredensial baru hasil reset admin." },
    { id: 17, title: "Hari Terakhir", desc: "Pengingat cross-check di hari terakhir stase." },
    { id: 18, title: "Pesan Broadcast", desc: "Pesan custom (pengumuman) dari Admin." },
    { id: 20, title: "Rekap Admin", desc: "Rangkuman total Hadir/Alpha harian untuk Admin." },
    { id: 25, title: "Onboarding Baru", desc: "Penyebaran password saat input MHS baru (Excel/Manual)." }
  ];

  return (
    <div className="min-h-screen bg-[#FFF06C] text-black font-sans selection:bg-pink-400 selection:text-black pb-20">
      
      {/* NAVBAR */}
      <nav className="border-b-[4px] border-black bg-white p-4 flex justify-between items-center sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="bg-[#FF6B6B] border-[3px] border-black p-2 shadow-[3px_3px_0px_0px_rgba(0,0,0,1)]">
            <Zap className="w-6 h-6 text-black" strokeWidth={3} />
          </div>
          <h1 className="text-xl md:text-2xl font-black uppercase tracking-widest">Bot WA Gateway</h1>
        </div>
        <a 
          href="/"
          className="hidden sm:flex items-center gap-2 bg-[#4ECDC4] hover:bg-[#3BBAA8] border-[3px] border-black font-black uppercase text-sm px-6 py-2 shadow-[4px_4px_0px_0px_rgba(0,0,0,1)] active:shadow-[0px_0px_0px_0px_rgba(0,0,0,1)] active:translate-y-[4px] active:translate-x-[4px] transition-all"
        >
          Kembali ke Web Utama
        </a>
      </nav>

      <main className="max-w-6xl mx-auto p-4 md:p-8 space-y-12 mt-8">
        
        {/* HERO SECTION */}
        <div className="bg-[#FF9FF3] border-[4px] border-black p-6 md:p-10 shadow-[8px_8px_0px_0px_rgba(0,0,0,1)] relative overflow-hidden">
          <div className="absolute -right-10 -top-10 opacity-20 transform rotate-12 pointer-events-none">
            <Server className="w-64 h-64" />
          </div>
          <h2 className="text-4xl md:text-6xl font-black uppercase mb-4 leading-none">API Docs Endpoint</h2>
          <p className="text-lg md:text-xl font-bold bg-white inline-block px-4 py-2 border-[3px] border-black shadow-[4px_4px_0px_0px_rgba(0,0,0,1)] mb-6 flex items-center gap-3 w-fit break-all">
            <span className="bg-[#FF6B6B] text-white px-2 py-1 text-sm border-2 border-black">POST</span>
            https://absensi.maksaarsyad.xyz/api/wa
          </p>
          <p className="font-bold text-lg max-w-2xl">
            Gateway ini bertugas mengonversi <span className="bg-white border-2 border-black px-2 mx-1">JSON Payload</span> menjadi teks pesan WhatsApp yang rapi dan terstruktur sesuai <span className="bg-white border-2 border-black px-2 mx-1">Scenario ID</span>.
          </p>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
          
          {/* JSON REQUEST */}
          <div className="space-y-4">
            <div className="flex items-center gap-3">
              <div className="bg-white border-[3px] border-black p-2 rounded-full shadow-[2px_2px_0px_0px_rgba(0,0,0,1)]">
                <Send className="w-5 h-5" />
              </div>
              <h3 className="text-2xl font-black uppercase tracking-wider">Format Body Request</h3>
            </div>
            <div className="bg-[#1E1E1E] text-[#00FF41] border-[4px] border-black p-0 shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] relative flex flex-col">
              <div className="bg-[#FF6B6B] border-b-[4px] border-black p-3 flex justify-between items-center">
                <span className="font-black text-black uppercase tracking-widest text-sm flex items-center gap-2">
                  <FileJson className="w-4 h-4"/> request.json
                </span>
                <button 
                  onClick={() => copyToClipboard(reqJson, 'req')}
                  className="bg-white text-black border-[2px] border-black px-3 py-1 font-bold text-xs uppercase shadow-[2px_2px_0px_0px_rgba(0,0,0,1)] active:shadow-[0px_0px_0px_0px] active:translate-y-[2px] active:translate-x-[2px] transition-all flex items-center gap-1"
                >
                  {copiedReq ? <><CheckCircle2 className="w-3 h-3"/> Copied</> : <><Copy className="w-3 h-3"/> Copy</>}
                </button>
              </div>
              <pre className="p-4 md:p-6 text-xs md:text-sm overflow-x-auto whitespace-pre-wrap font-mono leading-relaxed">
                <code>{reqJson}</code>
              </pre>
            </div>
          </div>

          {/* JSON RESPONSE */}
          <div className="space-y-4">
            <div className="flex items-center gap-3">
              <div className="bg-white border-[3px] border-black p-2 rounded-full shadow-[2px_2px_0px_0px_rgba(0,0,0,1)]">
                <Terminal className="w-5 h-5" />
              </div>
              <h3 className="text-2xl font-black uppercase tracking-wider">Response Generator</h3>
            </div>
            <div className="bg-[#F8F9FA] text-black border-[4px] border-black p-0 shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] relative flex flex-col">
              <div className="bg-[#4ECDC4] border-b-[4px] border-black p-3 flex justify-between items-center">
                <span className="font-black text-black uppercase tracking-widest text-sm flex items-center gap-2">
                  <Activity className="w-4 h-4"/> Status: 200 OK
                </span>
                <button 
                  onClick={() => copyToClipboard(resJson, 'res')}
                  className="bg-white text-black border-[2px] border-black px-3 py-1 font-bold text-xs uppercase shadow-[2px_2px_0px_0px_rgba(0,0,0,1)] active:shadow-[0px_0px_0px_0px] active:translate-y-[2px] active:translate-x-[2px] transition-all flex items-center gap-1"
                >
                  {copiedRes ? <><CheckCircle2 className="w-3 h-3"/> Copied</> : <><Copy className="w-3 h-3"/> Copy</>}
                </button>
              </div>
              <pre className="p-4 md:p-6 text-xs md:text-sm overflow-x-auto whitespace-pre-wrap font-mono font-bold leading-relaxed text-[#D80032]">
                <code>{resJson}</code>
              </pre>
            </div>
          </div>
        </div>

        {/* DICTIONARY OF SCENARIOS */}
        <div className="mt-12 space-y-6">
          <div className="bg-white border-[4px] border-black shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] p-6 flex flex-col sm:flex-row justify-between items-start sm:items-center">
             <div>
                <h3 className="text-3xl font-black uppercase tracking-wider mb-2">Daftar Skenario (IDs)</h3>
                <p className="font-bold">Gunakan ID ini pada property <code className="bg-[#FFF06C] px-2 py-0.5 border-2 border-black">scenario</code> di JSON Request.</p>
             </div>
             <div className="bg-[#45AAF2] border-[3px] border-black text-black font-black uppercase px-4 py-2 mt-4 sm:mt-0 shadow-[3px_3px_0px_0px_rgba(0,0,0,1)] rotate-2">
               Total: 18 Templates
             </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
            {scenarios.map((sc, index) => {
              // Menentukan warna rotasi berulang untuk Neubrutalism
              const colors = ['bg-[#FF6B6B]', 'bg-[#4ECDC4]', 'bg-[#45AAF2]', 'bg-[#FF9FF3]', 'bg-[#FEEA00]', 'bg-[#A3CB38]'];
              const cardColor = colors[index % colors.length];

              return (
                <div key={sc.id} className={`${cardColor} border-[3px] border-black p-4 shadow-[4px_4px_0px_0px_rgba(0,0,0,1)] hover:-translate-y-1 hover:-translate-x-1 hover:shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] transition-transform flex flex-col group`}>
                  <div className="flex justify-between items-start mb-3">
                    <span className="bg-white text-black font-black text-lg px-2 border-2 border-black shadow-[2px_2px_0px_0px_rgba(0,0,0,1)]">
                      ID: {sc.id}
                    </span>
                    <ArrowRight className="w-5 h-5 text-black opacity-0 group-hover:opacity-100 transition-opacity" strokeWidth={3} />
                  </div>
                  <h4 className="font-black uppercase text-lg leading-tight mb-1">{sc.title}</h4>
                  <p className="font-bold text-black/80 text-xs leading-snug">{sc.desc}</p>
                </div>
              );
            })}
          </div>
        </div>

      </main>
    </div>
  );
}
