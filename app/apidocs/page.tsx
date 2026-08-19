"use client";

import React, { useState, useEffect } from 'react';
import { Terminal, Copy, CheckCircle2, Zap, Send, FileJson, Server, Activity, ArrowRight, Loader2 } from 'lucide-react';

// ==========================================
// UPSTASH REDIS CLOUD CLIENT (For Dynamic Fetch)
// ==========================================
class Redis {
  url: string;
  token: string;
  
  constructor(config: { url: string; token: string }) {
    this.url = config.url || '';
    this.token = config.token || '';
    if (this.url.endsWith('/')) this.url = this.url.slice(0, -1);
  }
  
  static fromEnv() {
    let url = '';
    let token = '';
    if (typeof process !== 'undefined' && process.env) {
      url = process.env.NEXT_PUBLIC_UPSTASH_REDIS_REST_URL || process.env.NEXT_PUBLIC_KV_REST_API_URL || '';
      token = process.env.NEXT_PUBLIC_UPSTASH_REDIS_REST_TOKEN || process.env.NEXT_PUBLIC_KV_REST_API_TOKEN || '';
    }
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

interface FormatWA { id: number; title: string; description: string; template: string; }

export default function ApiDocs() {
  const [copiedReq, setCopiedReq] = useState(false);
  const [copiedRes, setCopiedRes] = useState(false);
  const [selectedScenario, setSelectedScenario] = useState<number>(1);
  const [formats, setFormats] = useState<FormatWA[]>([]);
  const [isLoading, setIsLoading] = useState(true);

  // Fetch Formats from Redis dynamically
  useEffect(() => {
    const fetchFormats = async () => {
      const redis = Redis.fromEnv();
      const dbFormats = await redis.get('axaxyz_formats');
      if (dbFormats && dbFormats.length > 0) {
         setFormats(dbFormats);
         // Set default selected to the first one available
         if (dbFormats.find((f:any) => f.id === 25)) {
             setSelectedScenario(25);
         } else {
             setSelectedScenario(dbFormats[0].id);
         }
      }
      setIsLoading(false);
    };
    fetchFormats();
  }, []);

  // Set Favicon & Title
  useEffect(() => {
    if (typeof window !== 'undefined') {
      let link = document.querySelector("link[rel~='icon']") as HTMLLinkElement;
      if (!link) {
        link = document.createElement('link');
        link.rel = 'icon';
        document.head.appendChild(link);
      }
      link.href = '/axalogo.png';
      link.type = 'image/png';
      document.title = "API Docs - Bot WA Dept. RKG";
    }
  }, []);

  // Dummy Data for Generator Preview
  const dummyData = {
     namaLengkap: "M. Azhar Arsyad",
     nim: "161202300030",
     kelompok: "Cluster II 2025",
     shift: "Shift Pagi",
     jamSesi: "07:30",
     jamTutup: "08:55",
     jamAbsen: "07:45",
     tanggalMulai: "01/08/2026",
     tanggalAkhir: "31/08/2026",
     tanggal: "20/08/2026",
     password: "123",
     radius: "500",
     lokasiGeofence: "Gedung Rektorat",
     pesanCustom: "Praktikum dialihkan ke Ruang B lantai 2 (Skenario 18)",
     link: "https://absensi.maksaarsyad.xyz/",
     totalMhs: "120",
     totalHadir: "110",
     totalTerlambat: "5",
     totalAlpha: "5"
  };

  // Dynamic JSON Generator
  const getScenarioData = (id: number) => {
    const baseReq: any = {
      no_hp: "081234567890",
      scenario: id,
      data: dummyData
    };

    const targetFormat = formats.find(f => f.id === id);
    let expectedResMsg = targetFormat ? targetFormat.template : "Template tidak ditemukan di database...";

    // Parse status kehadiran specifically
    if (expectedResMsg.includes('[Status Kehadiran]')) {
        expectedResMsg = expectedResMsg.replace(/\[Status Kehadiran\]/g, `🟢 *TEPAT WAKTU / HADIR* (Terekam pada: *${dummyData.jamAbsen}* WITA)`);
    }

    // Replace other placeholders dynamically
    expectedResMsg = expectedResMsg
      .replace(/\[Nama Lengkap\]/g, dummyData.namaLengkap)
      .replace(/\[NIM\]/g, dummyData.nim)
      .replace(/\[Kelompok\]/g, dummyData.kelompok)
      .replace(/\[Shift\]/g, dummyData.shift)
      .replace(/\[Jam Sesi\]/g, dummyData.jamSesi)
      .replace(/\[Jam Tutup\]/g, dummyData.jamTutup)
      .replace(/\[Jam Absen\]/g, dummyData.jamAbsen)
      .replace(/\[Tanggal Mulai\]/g, dummyData.tanggalMulai)
      .replace(/\[Tanggal Akhir\]/g, dummyData.tanggalAkhir)
      .replace(/\[Tanggal\]/g, dummyData.tanggal)
      .replace(/\[Password\]/g, dummyData.password)
      .replace(/\[Pesan Custom\]/g, dummyData.pesanCustom)
      .replace(/\[Radius\]/g, dummyData.radius)
      .replace(/\[Nama Lokasi Geofence\]/g, dummyData.lokasiGeofence)
      .replace(/\[Total Mhs\]/g, dummyData.totalMhs)
      .replace(/\[Total Hadir\]/g, dummyData.totalHadir)
      .replace(/\[Total Terlambat\]/g, dummyData.totalTerlambat)
      .replace(/\[Total Alpha\]/g, dummyData.totalAlpha)
      .replace(/\[Link\]/g, dummyData.link);

    // Escape newline for JSON display
    const resString = `{
  "success": true,
  "target_number": "081234567890",
  "formatted_message": "${expectedResMsg.replace(/\n/g, '\\n')}",
  "meta_info": {
    "scenario_executed": ${id},
    "timestamp": "${new Date().toISOString()}"
  }
}`;

    return {
      req: JSON.stringify(baseReq, null, 2),
      res: resString
    };
  };

  const { req, res } = getScenarioData(selectedScenario);

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

  const handleCardClick = (id: number) => {
    setSelectedScenario(id);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  if (isLoading) {
     return (
       <div className="min-h-screen bg-[#FFF06C] flex flex-col items-center justify-center p-4">
         <Loader2 className="w-12 h-12 text-black animate-spin mb-4" />
         <h2 className="text-xl font-black text-black tracking-widest uppercase mb-2">Memuat API Docs...</h2>
         <p className="text-black/80 text-xs font-bold uppercase">Mengambil data Skenario dari Database Cloud</p>
       </div>
     );
  }

  return (
    <div className="min-h-screen bg-[#FFF06C] text-black font-sans selection:bg-pink-400 selection:text-black pb-20">
      
      {/* NAVBAR */}
      <nav className="border-b-[4px] border-black bg-white p-4 flex justify-between items-center sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <div className="bg-[#FF6B6B] border-[3px] border-black p-2 shadow-[3px_3px_0px_0px_rgba(0,0,0,1)] flex items-center justify-center">
            {/* LOGO WEB SAYA */}
            <img src="/axalogo.png" alt="Logo Dept RKG" className="w-6 h-6 object-contain" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
            <Zap className="w-6 h-6 text-black hidden" strokeWidth={3} />
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
            Gateway ini bertugas mengonversi <span className="bg-white border-2 border-black px-2 mx-1">JSON Payload</span> menjadi teks pesan WhatsApp yang rapi dan terstruktur sesuai <span className="bg-white border-2 border-black px-2 mx-1">Scenario ID</span>. Pilih skenario di bawah untuk melihat preview dinamis dari database!
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
            <div className="bg-[#1E1E1E] text-[#00FF41] border-[4px] border-black p-0 shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] relative flex flex-col transition-all">
              <div className="bg-[#FF6B6B] border-b-[4px] border-black p-3 flex justify-between items-center">
                <span className="font-black text-black uppercase tracking-widest text-sm flex items-center gap-2">
                  <FileJson className="w-4 h-4"/> request.json (Scenario: {selectedScenario})
                </span>
                <button 
                  onClick={() => copyToClipboard(req, 'req')}
                  className="bg-white text-black border-[2px] border-black px-3 py-1 font-bold text-xs uppercase shadow-[2px_2px_0px_0px_rgba(0,0,0,1)] active:shadow-[0px_0px_0px_0px] active:translate-y-[2px] active:translate-x-[2px] transition-all flex items-center gap-1 cursor-pointer"
                >
                  {copiedReq ? <><CheckCircle2 className="w-3 h-3"/> Copied</> : <><Copy className="w-3 h-3"/> Copy</>}
                </button>
              </div>
              <pre className="p-4 md:p-6 text-xs md:text-sm overflow-x-auto whitespace-pre-wrap font-mono leading-relaxed">
                <code>{req}</code>
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
            <div className="bg-[#F8F9FA] text-black border-[4px] border-black p-0 shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] relative flex flex-col transition-all">
              <div className="bg-[#4ECDC4] border-b-[4px] border-black p-3 flex justify-between items-center">
                <span className="font-black text-black uppercase tracking-widest text-sm flex items-center gap-2">
                  <Activity className="w-4 h-4"/> Status: 200 OK
                </span>
                <button 
                  onClick={() => copyToClipboard(res, 'res')}
                  className="bg-white text-black border-[2px] border-black px-3 py-1 font-bold text-xs uppercase shadow-[2px_2px_0px_0px_rgba(0,0,0,1)] active:shadow-[0px_0px_0px_0px] active:translate-y-[2px] active:translate-x-[2px] transition-all flex items-center gap-1 cursor-pointer"
                >
                  {copiedRes ? <><CheckCircle2 className="w-3 h-3"/> Copied</> : <><Copy className="w-3 h-3"/> Copy</>}
                </button>
              </div>
              <pre className="p-4 md:p-6 text-xs md:text-sm overflow-x-auto whitespace-pre-wrap font-mono font-bold leading-relaxed text-[#D80032]">
                <code>{res}</code>
              </pre>
            </div>
          </div>
        </div>

        {/* DICTIONARY OF SCENARIOS DARI REDIS */}
        <div className="mt-12 space-y-6">
          <div className="bg-white border-[4px] border-black shadow-[6px_6px_0px_0px_rgba(0,0,0,1)] p-6 flex flex-col sm:flex-row justify-between items-start sm:items-center">
             <div>
                <h3 className="text-3xl font-black uppercase tracking-wider mb-2">Daftar Skenario (Load dari Database)</h3>
                <p className="font-bold">Gunakan ID ini pada property <code className="bg-[#FFF06C] px-2 py-0.5 border-2 border-black">scenario</code> di JSON Request.</p>
             </div>
             <div className="bg-[#45AAF2] border-[3px] border-black text-black font-black uppercase px-4 py-2 mt-4 sm:mt-0 shadow-[3px_3px_0px_0px_rgba(0,0,0,1)] rotate-2">
                Total: {formats.length} Skenario
             </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
            {formats.map((sc, index) => {
              // Warna berulang untuk Neubrutalism
              const colors = ['bg-[#FF6B6B]', 'bg-[#4ECDC4]', 'bg-[#45AAF2]', 'bg-[#FF9FF3]', 'bg-[#FEEA00]', 'bg-[#A3CB38]'];
              const cardColor = colors[index % colors.length];
              const isSelected = selectedScenario === sc.id;

              return (
                <div 
                  key={sc.id} 
                  onClick={() => handleCardClick(sc.id)}
                  className={`
                    ${cardColor} border-[3px] border-black p-4 transition-all flex flex-col group cursor-pointer
                    ${isSelected ? 'shadow-[0px_0px_0px_0px_rgba(0,0,0,1)] translate-y-[4px] translate-x-[4px] ring-4 ring-black/50' : 'shadow-[4px_4px_0px_0px_rgba(0,0,0,1)] hover:-translate-y-1 hover:-translate-x-1 hover:shadow-[6px_6px_0px_0px_rgba(0,0,0,1)]'}
                  `}
                >
                  <div className="flex justify-between items-start mb-3">
                    <span className={`bg-white text-black font-black text-lg px-2 border-2 border-black ${isSelected ? 'shadow-[1px_1px_0px_0px_rgba(0,0,0,1)]' : 'shadow-[2px_2px_0px_0px_rgba(0,0,0,1)]'}`}>
                      ID: {sc.id}
                    </span>
                    <ArrowRight className={`w-5 h-5 text-black transition-all duration-300 ${isSelected ? 'opacity-100 rotate-0' : 'opacity-0 -translate-x-2 group-hover:opacity-100 group-hover:translate-x-0'}`} strokeWidth={3} />
                  </div>
                  <h4 className="font-black uppercase text-lg leading-tight mb-1">{sc.title}</h4>
                  <p className="font-bold text-black/80 text-xs leading-snug">{sc.description}</p>
                </div>
              );
            })}
          </div>
        </div>

      </main>
    </div>
  );
}
