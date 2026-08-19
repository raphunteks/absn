"use client";

import React, { useState, useEffect, createContext, useContext, useRef, useCallback, useMemo } from 'react';
import { 
  Camera, MapPin, Clock, QrCode, CheckCircle2, AlertCircle, 
  BarChart3, Settings, FileText, LogOut, Users, Download, Plus, Trash2,
  RefreshCcw, ChevronRight, Fingerprint, Map, Activity, Key, Upload, Database, Navigation,
  Printer, X, CreditCard, Eye, EyeOff, Lock, ShieldCheck, Loader2, User, Cloud, CloudOff,
  ServerCrash, Maximize, Menu, Network, Edit, Calendar, UserX, ScanFace, ActivitySquare, MessageSquare, MessageCircle
} from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, AreaChart, Area } from 'recharts';

import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { format, startOfDay, endOfDay } from 'date-fns';

// ==========================================
// DATE HELPER UNTUK MENGHINDARI BUG TIMEZONE UTC
// ==========================================
const getLocalYYYYMMDD = (dateInput: Date | string) => {
  const d = new Date(dateInput);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

// ==========================================
// DYNAMIC SCRIPT LOADER UNTUK EXCEL (XLSX)
// ==========================================
const loadXlsx = async () => {
  if ((window as any).XLSX) return (window as any).XLSX;
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => resolve((window as any).XLSX);
    script.onerror = reject;
    document.head.appendChild(script);
  });
};

// ==========================================
// UPSTASH REDIS CLOUD CLIENT
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

const cn = (...inputs: ClassValue[]) => {
  return twMerge(clsx(inputs));
};

const calculateDistance = (lat1: number, lon1: number, lat2: number, lon2: number) => {
  const R = 6371e3; 
  const φ1 = lat1 * Math.PI / 180;
  const φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180;
  const Δλ = (lon2 - lon1) * Math.PI / 180;

  const a = Math.sin(Δφ / 2) * Math.sin(Δφ / 2) +
            Math.cos(φ1) * Math.cos(φ2) *
            Math.sin(Δλ / 2) * Math.sin(Δλ / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c; 
};

// EXPORT TO EXCEL DYNAMICALLY
const exportToExcel = async (logs: Log[]) => {
  try {
    const XLSX = await loadXlsx();
    const dataToExport = logs.map(log => ({
      'ID Log': log.id,
      'NIM': log.nim,
      'Nama Mahasiswa': log.name,
      'Kelompok': log.clusterName || 'Tanpa Kelompok',
      'Tanggal': new Date(log.timestamp).toLocaleDateString('id-ID'),
      'Waktu': new Date(log.timestamp).toLocaleTimeString('id-ID'),
      'Jadwal': log.sessionName,
      'Status': log.status,
      'Latitude': log.location.lat,
      'Longitude': log.location.lng,
      'Link G-Maps': `https://www.google.com/maps?q=${log.location.lat},${log.location.lng}`
    }));

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Riwayat Absensi");
    XLSX.writeFile(workbook, `Laporan_Absensi_${getLocalYYYYMMDD(new Date())}.xlsx`);
  } catch (error) {
    alert("Gagal memuat file Excel. Pastikan Anda memiliki koneksi internet.");
  }
};

const redis = Redis.fromEnv();

const CloudStore = {
  isAvailable: () => { 
    return redis !== null && redis.url !== '' && redis.token !== ''; 
  },
  async get(key: string) {
    if (!this.isAvailable()) return null;
    try {
      const data = await redis.get(key);
      return typeof data === 'string' ? JSON.parse(data) : data;
    } catch (e) {
      return null;
    }
  },
  async set(key: string, value: any) {
    if (!this.isAvailable()) throw new Error("Cloud Store Unavailable");
    await redis.set(key, value);
  }
};

// ==========================================
// INTERFACES
// ==========================================
interface Cluster { id: string; name: string; startDate?: string; endDate?: string; }
interface Session { id: string; name: string; startTime: string; endTime: string; toleranceMinutes: number; isActive: boolean; }
interface Log { id: string; nim: string; name: string; clusterName?: string; timestamp: string; sessionName: string; status: 'Hadir' | 'Terlambat'; location: { lat: number; lng: number }; photoBase64: string; deviceId: string; }
interface Student { id: string; nim: string; name: string; password?: string; deviceId?: string | null; clusterId?: string; }
interface Geofence { lat: number; lng: number; radius: number; name?: string; }
interface AdminUser { id: string; username: string; password?: string; }
interface ChatFormat { id: number; title: string; description: string; template: string; }

type SyncStatus = 'offline' | 'synced' | 'syncing' | 'error';

interface AppContextType {
  clusters: Cluster[];
  sessions: Session[];
  logs: Log[];
  students: Student[];
  geofence: Geofence;
  admins: AdminUser[];
  formats: ChatFormat[];
  isCloudSync: boolean;
  syncStatus: SyncStatus;
  addCluster: (cluster: Omit<Cluster, 'id'>) => void;
  updateCluster: (id: string, cluster: Omit<Cluster, 'id'>) => void;
  deleteCluster: (id: string) => void;
  addLog: (log: Omit<Log, 'id' | 'timestamp'>) => void;
  deleteLog: (id: string) => void;
  updateSession: (id: string, updates: Partial<Session>) => void;
  addSession: (session: Omit<Session, 'id'>) => void;
  deleteSession: (id: string) => void;
  addStudent: (student: Omit<Student, 'id'>) => void;
  updateStudent: (id: string, updates: Partial<Student>) => void;
  bulkAddStudents: (newStudents: Omit<Student, 'id'>[]) => void;
  deleteStudent: (id: string) => void;
  updateGeofence: (data: Geofence) => void;
  addAdmin: (admin: Omit<AdminUser, 'id'>) => void;
  updateAdmin: (id: string, updates: Partial<AdminUser>) => void;
  deleteAdmin: (id: string) => void;
  updateFormat: (id: number, template: string) => void;
  forceManualSync: () => Promise<void>;
  studentLogout: () => void;
}

const defaultSessions: Session[] = [
  { id: '1', name: 'Shift Pagi', startTime: '07:00', endTime: '09:00', toleranceMinutes: 15, isActive: true },
  { id: '2', name: 'Shift Siang', startTime: '12:00', endTime: '13:30', toleranceMinutes: 15, isActive: true },
];
const defaultGeofence: Geofence = { lat: -6.200000, lng: 106.816666, radius: 500, name: 'Gedung Kampus Pusat' };
const defaultClusters: Cluster[] = [{ id: 'c1', name: 'Angkatan 2024' }, { id: 'c2', name: 'Angkatan 2025' }];

const defaultFormats: ChatFormat[] = [
  { id: 1, title: "Pembukaan Sesi", description: "Dikirim saat jam shift dimulai.", template: "🔔 *NOTIFIKASI ABSENSI DIBUKA* 🔔\n\nHalo *[Nama Lengkap]*, sesi absensi untuk *[Shift]* Dept. RKG hari ini telah resmi *DIBUKA*.\n\n📋 *Detail Sesi Absensi:*\n• Kelompok: *[Kelompok]*\n• Jam Tepat Waktu: *[Jam Sesi]* WITA\n• Batas Tutup Sesi: *[Jam Tutup]* WITA\n\nYuk, segera lakukan validasi kehadiran Anda sekarang melalui portal resmi kami:\n[Link]\n\nSelamat bertugas! 🏥" },
  { id: 2, title: "Pengingat Sisa Waktu", description: "Mengingatkan MHS yang belum absen.", template: "⚠️ *PENGINGAT TERAKHIR ABSENSI* ⚠️\n\nPanggilan kepada *[Nama Lengkap]*! Sistem mendeteksi Anda *BELUM* melakukan absensi untuk *[Shift]* hari ini.\n\nWaktu absensi Anda hampir habis. Sesi ini akan ditutup secara permanen pada pukul *[Jam Tutup]* WITA.\n\nMohon segera selesaikan absen Anda di sini:\n[Link]" },
  { id: 3, title: "Berhasil Absen", description: "Dikirim real-time setelah absen.", template: "✅ *ABSENSI BERHASIL DITERIMA* ✅\n\nTerima kasih *[Nama Lengkap]*! Data kehadiran Anda untuk *[Shift]* telah diamankan ke Database Dept. RKG.\n\n📌 *Bukti Rekam Kehadiran:*\n• Waktu Absen: *[Jam Absen]* WITA\n• Kelompok: *[Kelompok]*\n• Keamanan: Tervalidasi (Face ID + GPS)\n\nSilakan cek ringkasan kehadiran di dashboard Anda:\n[Link]" },
  { id: 25, title: "Onboarding Baru", description: "Pesan welcome & penyebaran password akun.", template: "🎉 *SELAMAT DATANG DI DEPT. RKG!* 🎉\n\nHalo *[Nama Lengkap]*, selamat bergabung! \nData Anda telah didaftarkan ke Sistem Absensi Digital Dept. RKG.\n\nKredensial Akses Anda:\n👤 *Nama:* [Nama Lengkap]\n🆔 *NIM:* [NIM]\n👥 *Kelompok:* [Kelompok]\n🗓️ *Masa Stase:* [Tanggal Mulai] - [Tanggal Akhir]\n🔑 *Kata Sandi:* [Password]\n\nAgar Anda dapat mulai absen, ikuti langkah berikut:\n1. Buka portal di: [Link]\n2. Klik tombol \"Mulai Absensi\" di layar utama.\n3. Masukkan NIM dan Kata Sandi Anda.\n\n⚠️ PENTING: Perangkat/HP pertama yang Anda gunakan login akan DIKUNCI (Terikat) dengan akun Anda.\nJaga kerahasiaan sandi Anda!" }
];

const AppContext = createContext<AppContextType | null>(null);

// ==========================================
// APP PROVIDER
// ==========================================
const AppProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [isAppLoading, setIsAppLoading] = useState(true);
  const [isCloudSync, setIsCloudSync] = useState(false);
  const [syncStatus, setSyncStatus] = useState<SyncStatus>('offline');
  
  const [clusters, setClusters] = useState<Cluster[]>([]);
  const [sessions, setSessions] = useState<Session[]>([]);
  const [logs, setLogs] = useState<Log[]>([]);
  const [students, setStudents] = useState<Student[]>([]);
  const [geofence, setGeofence] = useState<Geofence>(defaultGeofence);
  const [admins, setAdmins] = useState<AdminUser[]>([]);
  const [formats, setFormats] = useState<ChatFormat[]>([]);

  useEffect(() => {
    const initData = async () => {
      const cloudAvailable = CloudStore.isAvailable();
      setIsCloudSync(cloudAvailable);
      setSyncStatus(cloudAvailable ? 'synced' : 'offline');

      let c = null, s = null, l = null, st = null, gf = null, ad = null, fm = null;

      if (cloudAvailable) {
        c = await CloudStore.get('axaxyz_clusters');
        s = await CloudStore.get('axaxyz_sessions');
        l = await CloudStore.get('axaxyz_logs');
        st = await CloudStore.get('axaxyz_students');
        gf = await CloudStore.get('axaxyz_geofence');
        ad = await CloudStore.get('axaxyz_admins');
        fm = await CloudStore.get('axaxyz_formats');
      }

      if (!c) c = JSON.parse(localStorage.getItem('axaxyz_clusters') || 'null');
      if (!s) s = JSON.parse(localStorage.getItem('axaxyz_sessions') || 'null');
      if (!l) l = JSON.parse(localStorage.getItem('axaxyz_logs') || 'null');
      if (!st) st = JSON.parse(localStorage.getItem('axaxyz_students') || 'null');
      if (!gf) gf = JSON.parse(localStorage.getItem('axaxyz_geofence') || 'null');
      if (!ad) ad = JSON.parse(localStorage.getItem('axaxyz_admins') || 'null');
      if (!fm) fm = JSON.parse(localStorage.getItem('axaxyz_formats') || 'null');

      setClusters(c || defaultClusters);
      setSessions(s || defaultSessions);
      setLogs(l || []);
      setStudents(st || []);
      setGeofence(gf || defaultGeofence);
      setAdmins(ad || []);
      setFormats(fm || defaultFormats);
      
      setIsAppLoading(false);
    };
    initData();
  }, []);

  const syncToCloud = async (key: string, data: any) => {
    if (!CloudStore.isAvailable()) return;
    setSyncStatus('syncing');
    try {
      await CloudStore.set(key, JSON.stringify(data));
      setSyncStatus('synced');
    } catch (e) {
      setSyncStatus('error');
    }
  };

  const forceManualSync = async () => {
    if (!CloudStore.isAvailable()) {
      alert("❌ Gagal Sinkronisasi: Pengaturan Cloud Database tidak ditemukan.");
      return;
    }
    setSyncStatus('syncing');
    try {
      await CloudStore.set('axaxyz_clusters', JSON.stringify(clusters));
      await CloudStore.set('axaxyz_sessions', JSON.stringify(sessions));
      await CloudStore.set('axaxyz_logs', JSON.stringify(logs));
      await CloudStore.set('axaxyz_students', JSON.stringify(students));
      await CloudStore.set('axaxyz_geofence', JSON.stringify(geofence));
      await CloudStore.set('axaxyz_admins', JSON.stringify(admins));
      await CloudStore.set('axaxyz_formats', JSON.stringify(formats));
      setSyncStatus('synced');
      alert("✅ Seluruh data berhasil dicadangkan ke Cloud Database!");
    } catch (e: any) {
      setSyncStatus('error');
      alert("❌ Terjadi kesalahan saat menyimpan data: " + e.message);
    }
  };

  const saveClusters = (d: Cluster[]) => { setClusters(d); localStorage.setItem('axaxyz_clusters', JSON.stringify(d)); syncToCloud('axaxyz_clusters', d); };
  const saveSessions = (d: Session[]) => { setSessions(d); localStorage.setItem('axaxyz_sessions', JSON.stringify(d)); syncToCloud('axaxyz_sessions', d); };
  const saveLogs = (d: Log[]) => { setLogs(d); localStorage.setItem('axaxyz_logs', JSON.stringify(d)); syncToCloud('axaxyz_logs', d); };
  const saveStudents = (d: Student[]) => { setStudents(d); localStorage.setItem('axaxyz_students', JSON.stringify(d)); syncToCloud('axaxyz_students', d); };
  const saveGeofence = (d: Geofence) => { setGeofence(d); localStorage.setItem('axaxyz_geofence', JSON.stringify(d)); syncToCloud('axaxyz_geofence', d); };
  const saveAdmins = (d: AdminUser[]) => { setAdmins(d); localStorage.setItem('axaxyz_admins', JSON.stringify(d)); syncToCloud('axaxyz_admins', d); };
  const saveFormats = (d: ChatFormat[]) => { setFormats(d); localStorage.setItem('axaxyz_formats', JSON.stringify(d)); syncToCloud('axaxyz_formats', d); };

  const addCluster = (data: Omit<Cluster, 'id'>) => saveClusters([...clusters, { ...data, id: Math.random().toString(36).substr(2, 9) }]);
  const updateCluster = (id: string, data: Omit<Cluster, 'id'>) => saveClusters(clusters.map(c => c.id === id ? { ...c, ...data } : c));
  const deleteCluster = (id: string) => saveClusters(clusters.filter(c => c.id !== id));

  const addLog = (logData: Omit<Log, 'id' | 'timestamp'>) => saveLogs([{ ...logData, id: Math.random().toString(36).substr(2, 9), timestamp: new Date().toISOString() }, ...logs]);
  const deleteLog = (id: string) => saveLogs(logs.filter(l => l.id !== id));
  
  const updateSession = (id: string, updates: Partial<Session>) => saveSessions(sessions.map(s => s.id === id ? { ...s, ...updates } : s));
  const addSession = (sessionData: Omit<Session, 'id'>) => saveSessions([...sessions, { ...sessionData, id: Math.random().toString(36).substr(2, 9) }]);
  const deleteSession = (id: string) => saveSessions(sessions.filter(s => s.id !== id));
  
  const addStudent = (studentData: Omit<Student, 'id'>) => saveStudents([...students, { ...studentData, id: Math.random().toString(36).substr(2, 9) }]);
  const updateStudent = (id: string, updates: Partial<Student>) => saveStudents(students.map(s => s.id === id ? { ...s, ...updates } : s));
  const bulkAddStudents = (newStudents: Omit<Student, 'id'>[]) => {
    const formatted = newStudents.map(s => ({ ...s, id: Math.random().toString(36).substr(2, 9) }));
    saveStudents([...students, ...formatted]);
  };
  const deleteStudent = (id: string) => saveStudents(students.filter(s => s.id !== id));
  const updateGeofence = (data: Geofence) => saveGeofence(data);

  const addAdmin = (adminData: Omit<AdminUser, 'id'>) => saveAdmins([...admins, { ...adminData, id: Math.random().toString(36).substr(2, 9) }]);
  const updateAdmin = (id: string, updates: Partial<AdminUser>) => saveAdmins(admins.map(a => a.id === id ? { ...a, ...updates } : a));
  const deleteAdmin = (id: string) => saveAdmins(admins.filter(a => a.id !== id));

  const updateFormat = (id: number, template: string) => {
    saveFormats(formats.map(f => f.id === id ? { ...f, template } : f));
  };

  const studentLogout = () => {
     if(!confirm('Apakah Anda yakin ingin logout / keluar dari akun mahasiswa di perangkat ini?')) return;
     const ownerNim = localStorage.getItem('axaxyz_device_owner');
     if (ownerNim) {
        const student = students.find(s => s.nim === ownerNim);
        if (student) updateStudent(student.id, { deviceId: null });
     }
     localStorage.removeItem('axaxyz_device_id');
     localStorage.removeItem('axaxyz_device_owner');
     alert('Anda telah berhasil keluar dari sistem absensi di perangkat ini.');
     window.location.reload();
  };

  if (isAppLoading) {
    return (
      <div className="min-h-screen bg-[#020617] flex flex-col items-center justify-center p-4 radiology-bg">
        <div className="relative mb-6">
          <div className="absolute inset-0 bg-cyan-500/30 blur-[50px] rounded-full animate-pulse"></div>
          <div className="w-24 h-24 bg-[#0A1628] border border-cyan-500/50 rounded-3xl flex items-center justify-center shadow-[0_0_30px_rgba(6,182,212,0.4)] relative z-10 overflow-hidden p-4">
             <ScanFace className="w-full h-full text-cyan-400 opacity-80" />
          </div>
        </div>
        <Loader2 className="w-8 h-8 text-cyan-400 animate-spin mb-4 drop-shadow-[0_0_10px_rgba(6,182,212,0.8)]" />
        <h2 className="text-xl font-black text-cyan-100 tracking-widest uppercase mb-2">Memuat Sistem...</h2>
        <p className="text-cyan-600/80 text-xs text-center max-w-xs font-mono uppercase">Menyiapkan portal absensi mahasiswa</p>
      </div>
    );
  }

  return (
    <AppContext.Provider value={{ 
      isCloudSync, syncStatus, clusters, sessions, logs, students, geofence, admins, formats,
      addCluster, updateCluster, deleteCluster, addLog, deleteLog, updateSession, addSession, deleteSession, 
      addStudent, updateStudent, bulkAddStudents, deleteStudent, updateGeofence, forceManualSync, studentLogout,
      addAdmin, updateAdmin, deleteAdmin, updateFormat
    }}>
      {children}
    </AppContext.Provider>
  );
};

const useAppContext = () => {
  const context = useContext(AppContext);
  if (!context) throw new Error("useAppContext must be used within AppProvider");
  return context;
};

// ==========================================
// DATE FILTER UTILITY HOOK
// ==========================================
const useDateFilter = () => {
  const todayStr = getLocalYYYYMMDD(new Date());
  const [datePreset, setDatePreset] = useState('today');
  const [customStart, setCustomStart] = useState(todayStr);
  const [customEnd, setCustomEnd] = useState(todayStr);

  const { startObj, endObj } = useMemo(() => {
    const now = new Date();
    let s = startOfDay(now);
    let e = endOfDay(now);

    if (datePreset === 'week') {
      const past = new Date(now);
      past.setDate(past.getDate() - 7);
      s = startOfDay(past);
    } else if (datePreset === 'month') {
      const past = new Date(now);
      past.setDate(past.getDate() - 30);
      s = startOfDay(past);
    } else if (datePreset === 'custom') {
      s = startOfDay(new Date(customStart));
      e = endOfDay(new Date(customEnd));
    }
    return { startObj: s, endObj: e };
  }, [datePreset, customStart, customEnd]);

  const FilterUI = () => (
    <div className="flex flex-col sm:flex-row gap-3 w-full sm:w-auto items-start sm:items-end">
      <div className="flex flex-col gap-1 w-full sm:w-auto">
         <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Filter Tanggal</label>
         <div className="flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl px-2 h-11 w-full sm:w-auto focus-within:border-cyan-400 transition-colors">
            <select value={datePreset} onChange={e=>setDatePreset(e.target.value)} className="bg-transparent text-cyan-50 text-xs font-bold uppercase outline-none cursor-pointer px-3 w-full sm:w-40 h-full">
              <option value="today">Hari Ini</option>
              <option value="week">1 Minggu Terakhir</option>
              <option value="month">1 Bulan Terakhir</option>
              <option value="custom">Custom Tanggal</option>
            </select>
         </div>
      </div>
      {datePreset === 'custom' && (
        <div className="flex flex-col sm:flex-row gap-3 w-full sm:w-auto items-start sm:items-end">
          <div className="flex flex-col gap-1 w-full sm:w-auto">
             <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Mulai Tanggal</label>
             <input type="date" value={customStart} onChange={e => setCustomStart(e.target.value)} className="bg-[#050B14] border border-cyan-500/30 text-cyan-50 text-xs font-mono px-3 h-11 rounded-lg outline-none focus:border-cyan-400 w-full" />
          </div>
          <div className="flex flex-col gap-1 w-full sm:w-auto">
             <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Hingga Tanggal</label>
             <input type="date" value={customEnd} onChange={e => setCustomEnd(e.target.value)} className="bg-[#050B14] border border-cyan-500/30 text-cyan-50 text-xs font-mono px-3 h-11 rounded-lg outline-none focus:border-cyan-400 w-full" />
          </div>
        </div>
      )}
    </div>
  );

  return { startObj, endObj, FilterUI };
};


// ==========================================
// STUDENT DASHBOARD (SUPER UPGRADE)
// ==========================================
const StudentDashboard: React.FC<{ onStartAbsen: () => void, linkedNim: string | null }> = ({ onStartAbsen, linkedNim }) => {
  const { logs, students, clusters, sessions } = useAppContext();

  const student = students.find(s => s.nim === linkedNim);
  const studentName = student?.name;
  const firstName = studentName ? studentName.split(' ')[0] : 'Mahasiswa Baru';
  const myCluster = clusters.find(c => c.id === student?.clusterId);
  const activeSessions = useMemo(() => sessions.filter(s => s.isActive), [sessions]);
  
  const myLogs = useMemo(() => {
    if (!linkedNim) return [];
    return logs.filter(l => l.nim === linkedNim).sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());
  }, [logs, linkedNim]);

  const totalHadir = myLogs.length;
  const onTime = myLogs.filter(l => l.status === 'Hadir').length;
  const late = myLogs.filter(l => l.status === 'Terlambat').length;
  const onTimePercent = totalHadir > 0 ? Math.round((onTime / totalHadir) * 100) : 0;

  const todayStr = getLocalYYYYMMDD(new Date());
  const todayLogs = myLogs.filter(l => getLocalYYYYMMDD(l.timestamp) === todayStr);
  const hasClockedInToday = todayLogs.length > 0;
  const lastClockInTime = hasClockedInToday ? new Date(todayLogs[0].timestamp).toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }) : null;

  const { chartData: staseDataList, alpha: myAlpha, belumAbsen: myBelumAbsen } = useMemo(() => {
     if (!linkedNim) {
         return { chartData: [], alpha: 0, belumAbsen: 0 };
     }

     let startD = new Date(); 
     startD.setDate(startD.getDate() - 6); 
     let endD = new Date();

     if (myCluster?.startDate && myCluster?.endDate) {
         startD = new Date(myCluster.startDate);
         const staseEnd = new Date(myCluster.endDate);
         endD = staseEnd > new Date() ? new Date() : staseEnd; 
     }

     const data = [];
     const todayLocal = getLocalYYYYMMDD(new Date());
     let tempAlpha = 0;
     let tempBelumAbsen = 0;

     for (let d = new Date(startD); d <= endD; d.setDate(d.getDate() + 1)) {
         if (d > new Date()) break;

         const dateStrLocal = getLocalYYYYMMDD(d);
         const isToday = dateStrLocal === todayLocal;
         const label = d.toLocaleDateString('id-ID', { day: '2-digit', month: 'short' });
         
         let dayHadir = 0;
         let dayTerlambat = 0;
         let dayAlpha = 0;
         let dayBelumAbsen = 0;

         activeSessions.forEach(sess => {
             const log = myLogs.find(l => getLocalYYYYMMDD(l.timestamp) === dateStrLocal && l.sessionName === sess.name);
             if (log) {
                 if (log.status === 'Hadir') dayHadir++;
                 else if (log.status === 'Terlambat') dayTerlambat++;
             } else {
                 if (isToday) {
                     const currentMinutes = new Date().getHours() * 60 + new Date().getMinutes();
                     const [endH, endM] = sess.endTime.split(':').map(Number);
                     const endTotal = endH * 60 + endM;
                     const endWithTol = endTotal + sess.toleranceMinutes;
                     
                     if (currentMinutes > endWithTol) { dayAlpha++; tempAlpha++; }
                     else { dayBelumAbsen++; tempBelumAbsen++; }
                 } else {
                     if (d < new Date(new Date().setHours(0,0,0,0))) {
                        dayAlpha++; tempAlpha++;
                     }
                 }
             }
         });

         data.push({ 
            day: label, 
            Hadir: dayHadir, 
            Terlambat: dayTerlambat, 
            Alpha: dayAlpha, 
            BelumAbsen: dayBelumAbsen 
         });
     }
     return { chartData: data, alpha: tempAlpha, belumAbsen: tempBelumAbsen };
  }, [myLogs, myCluster, activeSessions, linkedNim]);

  const recentLogs = myLogs.slice(0, 4);

  const getInitials = (name: string) => {
     const parts = name.split(' ');
     if (parts.length > 1) return (parts[0][0] + parts[1][0]).toUpperCase();
     return name.substring(0, 2).toUpperCase();
  }

  const getGreeting = () => {
     const hour = new Date().getHours();
     if (hour < 12) return 'Selamat Pagi';
     if (hour < 15) return 'Selamat Siang';
     if (hour < 18) return 'Selamat Sore';
     return 'Selamat Malam';
  };

  return (
    <div className="animate-in fade-in duration-700 w-full space-y-6">
      {/* Header */}
      <div className="mb-8">
        <h2 className="text-3xl md:text-4xl font-black text-white tracking-wide">{getGreeting()}, {firstName}</h2>
        <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-2">{new Date().toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })} • Ringkasan kehadiran live</p>
      </div>

      {/* Stats Cards (SUPER UPGRADE TO 5 WIDGETS) */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-4">
        {/* Card 1: Total Hadir */}
        <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-5 rounded-[1.5rem] shadow-lg relative overflow-hidden group hover:border-cyan-400/50 transition-all flex flex-col justify-between">
           <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 md:w-10 md:h-10 rounded-full bg-cyan-500/10 flex items-center justify-center text-cyan-400 shrink-0"><Clock className="w-4 h-4 md:w-5 md:h-5"/></div>
              <span className="text-[10px] md:text-xs font-black uppercase tracking-widest text-cyan-500 leading-tight">Total Hadir</span>
           </div>
           <div className="flex items-end gap-2 mt-auto">
              <h3 className="text-4xl md:text-5xl font-black text-white leading-none">{totalHadir}</h3>
              <span className="text-cyan-400 text-[9px] md:text-xs font-bold mb-1 tracking-wider bg-cyan-500/10 px-2 py-1 rounded-md whitespace-nowrap">Total Sesi</span>
           </div>
        </div>
        
        {/* Card 2: Tepat Waktu */}
        <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-5 rounded-[1.5rem] shadow-lg relative overflow-hidden group hover:border-emerald-400/50 transition-all flex flex-col justify-between">
           <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 md:w-10 md:h-10 rounded-full bg-emerald-500/10 flex items-center justify-center text-emerald-400 shrink-0"><CheckCircle2 className="w-4 h-4 md:w-5 md:h-5"/></div>
              <span className="text-[10px] md:text-xs font-black uppercase tracking-widest text-emerald-500 leading-tight">Tepat Waktu</span>
           </div>
           <div className="flex items-end gap-2 mt-auto">
              <h3 className="text-4xl md:text-5xl font-black text-white leading-none">{onTimePercent}<span className="text-xl md:text-2xl">%</span></h3>
              <span className="text-emerald-400 text-[9px] md:text-xs font-bold mb-1 tracking-wider bg-emerald-500/10 px-2 py-1 rounded-md whitespace-nowrap">Rata-rata</span>
           </div>
        </div>

        {/* Card 3: Terlambat (Purple Theme) */}
        <div className="bg-gradient-to-br from-[#4c1d95]/90 to-[#7e22ce]/90 backdrop-blur-xl border border-purple-500/50 p-5 rounded-[1.5rem] shadow-[0_10px_40px_rgba(126,34,206,0.3)] relative overflow-hidden group flex flex-col justify-between col-span-2 md:col-span-1">
           <div className="absolute top-0 right-0 w-32 h-32 bg-white/5 rounded-full blur-2xl pointer-events-none"></div>
           <div className="flex items-center gap-3 mb-4 relative z-10">
              <div className="w-8 h-8 md:w-10 md:h-10 rounded-full bg-white/20 flex items-center justify-center text-white shrink-0"><AlertCircle className="w-4 h-4 md:w-5 md:h-5"/></div>
              <span className="text-[10px] md:text-xs font-black uppercase tracking-widest text-purple-100 leading-tight">Terlambat</span>
           </div>
           <div className="flex items-end gap-2 relative z-10 mt-auto">
              <h3 className="text-4xl md:text-5xl font-black text-white leading-none">{late}</h3>
              <span className="text-purple-200 text-[9px] md:text-xs font-bold mb-1 tracking-wider bg-white/20 px-2 py-1 rounded-md whitespace-nowrap">Sesi</span>
           </div>
        </div>

        {/* Card 4: Tidak Hadir (Alpha) (Rose Theme) */}
        <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-rose-500/20 p-5 rounded-[1.5rem] shadow-lg relative overflow-hidden group hover:border-rose-400/50 transition-all flex flex-col justify-between">
           <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 md:w-10 md:h-10 rounded-full bg-rose-500/10 flex items-center justify-center text-rose-400 shrink-0"><UserX className="w-4 h-4 md:w-5 md:h-5"/></div>
              <span className="text-[10px] md:text-xs font-black uppercase tracking-widest text-rose-500 leading-tight">Tidak Hadir (Alpha)</span>
           </div>
           <div className="flex items-end gap-2 mt-auto">
              <h3 className="text-4xl md:text-5xl font-black text-white leading-none">{myAlpha}</h3>
              <span className="text-rose-400 text-[9px] md:text-[10px] font-bold mb-1 tracking-wider bg-rose-500/10 px-2 py-1 rounded-md whitespace-nowrap">Total Sesi</span>
           </div>
        </div>

        {/* Card 5: Belum Absen (Blue/Amber Theme) */}
        <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-blue-500/20 p-5 rounded-[1.5rem] shadow-lg relative overflow-hidden group hover:border-blue-400/50 transition-all flex flex-col justify-between">
           <div className="flex items-center gap-3 mb-4">
              <div className="w-8 h-8 md:w-10 md:h-10 rounded-full bg-blue-500/10 flex items-center justify-center text-blue-400 shrink-0"><ActivitySquare className="w-4 h-4 md:w-5 md:h-5"/></div>
              <span className="text-[10px] md:text-xs font-black uppercase tracking-widest text-blue-500 leading-tight">Belum Absen</span>
           </div>
           <div className="flex items-end gap-2 mt-auto">
              <h3 className="text-4xl md:text-5xl font-black text-white leading-none">{myBelumAbsen}</h3>
              <span className="text-blue-400 text-[9px] md:text-[10px] font-bold mb-1 tracking-wider bg-blue-500/10 px-2 py-1 rounded-md whitespace-nowrap">Sesi Hari Ini</span>
           </div>
        </div>
      </div>

      {/* Middle Section */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5 mt-8">
        {/* Chart */}
        <div className="lg:col-span-2 bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-[2rem] shadow-lg flex flex-col overflow-hidden">
          <h3 className="text-base font-black text-white mb-1 tracking-widest uppercase truncate">{myCluster?.startDate ? 'Kehadiran Periode Stase' : 'Kehadiran Mingguan'}</h3>
          <p className="text-[10px] md:text-xs text-cyan-500/70 font-mono uppercase mb-6">Jam tervalidasi, sinkronisasi otomatis ke sistem</p>
          <div className="flex-1 w-full min-h-[200px] overflow-x-auto custom-scrollbar">
             <div className="min-w-[400px] h-full">
                 {staseDataList.length > 0 ? (
                   <ResponsiveContainer width="100%" height="100%">
                     <BarChart data={staseDataList} margin={{ top: 0, right: 0, left: 0, bottom: 0 }}>
                       <XAxis dataKey="day" axisLine={false} tickLine={false} tick={{fill: '#64748b', fontSize: 10, fontWeight: 'bold'}} dy={10} interval="preserveStartEnd" minTickGap={20} />
                       <Tooltip cursor={{fill: '#1e293b', opacity: 0.4}} contentStyle={{backgroundColor: '#050B14', borderColor: '#8b5cf6', color: '#f8fafc', borderRadius: '1rem', fontSize: '12px'}} />
                       <Bar dataKey="Hadir" stackId="a" fill="#10b981" barSize={staseDataList.length > 15 ? 15 : 40} />
                       <Bar dataKey="Terlambat" stackId="a" fill="#f59e0b" />
                       <Bar dataKey="Alpha" stackId="a" fill="#f43f5e" />
                       <Bar dataKey="BelumAbsen" stackId="a" fill="#3b82f6" />
                     </BarChart>
                   </ResponsiveContainer>
                 ) : (
                   <div className="h-full flex items-center justify-center text-cyan-800 font-mono text-xs uppercase tracking-widest">Data Kosong (Belum Login)</div>
                 )}
             </div>
          </div>
          <div className="flex justify-center gap-4 mt-6 flex-wrap">
             <div className="flex items-center gap-1.5 text-[9px] font-bold uppercase tracking-wider text-slate-300">
                <div className="w-2.5 h-2.5 rounded-full bg-emerald-500"></div> Hadir
             </div>
             <div className="flex items-center gap-1.5 text-[9px] font-bold uppercase tracking-wider text-slate-300">
                <div className="w-2.5 h-2.5 rounded-full bg-amber-500"></div> Terlambat
             </div>
             <div className="flex items-center gap-1.5 text-[9px] font-bold uppercase tracking-wider text-slate-300">
                <div className="w-2.5 h-2.5 rounded-full bg-rose-500"></div> Alpha
             </div>
             <div className="flex items-center gap-1.5 text-[9px] font-bold uppercase tracking-wider text-slate-300">
                <div className="w-2.5 h-2.5 rounded-full bg-blue-500"></div> Belum Absen
             </div>
          </div>
        </div>

        {/* Clock-in Card */}
        <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-[2rem] shadow-lg flex flex-col items-center justify-center text-center relative overflow-hidden">
           <div className="absolute top-0 right-0 w-48 h-48 bg-purple-600/10 rounded-full blur-3xl pointer-events-none"></div>
           
           <h3 className="text-base font-black text-white mb-1 tracking-widest uppercase relative z-10">Panel Absensi</h3>
           <p className="text-[10px] md:text-xs text-cyan-500/70 font-mono uppercase mb-8 relative z-10">Face ID + GPS verified</p>

           {/* Segmented Circle Avatar (Matching the reference UI) */}
           <div className="relative w-32 h-32 md:w-36 md:h-36 mb-8 z-10">
              <svg className="absolute inset-0 w-full h-full" viewBox="0 0 100 100">
                 <circle cx="50" cy="50" r="46" fill="none" stroke={hasClockedInToday ? "#8b5cf6" : "#1e293b"} strokeWidth="3" strokeDasharray="60 12" strokeLinecap="round" className={hasClockedInToday ? "animate-[spin_20s_linear_infinite]" : ""} />
              </svg>
              <div className="absolute inset-0 flex items-center justify-center">
                 <div className="w-20 h-20 md:w-24 md:h-24 bg-[#050B14] rounded-full border border-purple-500/30 flex items-center justify-center shadow-[inset_0_0_20px_rgba(147,51,234,0.2)]">
                    <User className={cn("w-10 h-10", hasClockedInToday ? "text-purple-400" : "text-slate-600")} />
                 </div>
              </div>
           </div>

           <div className="relative z-10 w-full">
             {hasClockedInToday ? (
                <div className="w-full py-4 bg-gradient-to-r from-purple-700 to-purple-500 rounded-2xl shadow-[0_10px_30px_rgba(147,51,234,0.4)] border border-purple-400/50">
                   <span className="text-white font-black tracking-widest text-lg md:text-xl">{lastClockInTime}</span>
                </div>
             ) : (
                <button onClick={onStartAbsen} className="w-full py-4 bg-gradient-to-r from-cyan-600 to-purple-600 hover:from-cyan-500 hover:to-purple-500 text-white rounded-2xl transition-all duration-300 shadow-[0_10px_30px_rgba(147,51,234,0.4)] font-black tracking-widest text-sm active:scale-95 border border-purple-400/50">
                   MULAI ABSENSI
                </button>
             )}
             {!linkedNim && (
               <p className="text-[9px] text-cyan-400 font-mono mt-4 uppercase">Tautkan perangkat ini dengan melakukan absensi perdana.</p>
             )}
           </div>
        </div>
      </div>

      {/* Bottom Section */}
      <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-[2rem] shadow-lg">
         <h3 className="text-base font-black text-white mb-6 tracking-widest uppercase">Absen Terakhir</h3>
         <div className="flex flex-wrap gap-4 overflow-x-auto pb-2 custom-scrollbar">
            {recentLogs.map((log, idx) => (
               <div key={log.id + idx} className="flex items-center gap-4 bg-[#050B14] p-3 md:p-4 rounded-2xl border border-cyan-500/20 min-w-[220px] shadow-sm hover:border-cyan-500/50 transition-colors">
                  <div className="w-10 h-10 md:w-12 md:h-12 rounded-full bg-purple-500/20 flex items-center justify-center text-purple-400 font-black text-xs md:text-sm border border-purple-500/30 shrink-0">
                     {getInitials(log.name)}
                  </div>
                  <div className="flex-1">
                     <p className="text-xs md:text-sm font-bold text-white uppercase truncate max-w-[120px]">{log.name}</p>
                     <p className="text-[10px] text-cyan-500/70 font-mono">{new Date(log.timestamp).toLocaleTimeString('id-ID', {hour: '2-digit', minute:'2-digit'})}</p>
                  </div>
                  <div className="shrink-0 w-6 h-6 rounded-full flex items-center justify-center bg-[#0A1628] border border-cyan-900/50">
                     {log.status === 'Hadir' ? <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400 drop-shadow-[0_0_5px_rgba(16,185,129,0.8)]"/> : <AlertCircle className="w-3.5 h-3.5 text-amber-400 drop-shadow-[0_0_5px_rgba(245,158,11,0.8)]"/>}
                  </div>
               </div>
            ))}
            {recentLogs.length === 0 && (
              <div className="w-full text-center py-6 border-2 border-dashed border-cyan-900/50 rounded-2xl">
                 <p className="text-xs text-cyan-700 font-mono uppercase tracking-widest">Belum ada riwayat absensi tercatat di perangkat ini.</p>
              </div>
            )}
         </div>
      </div>
    </div>
  );
};


// ==========================================
// PORTAL MAHASISWA WIZARD
// ==========================================
const TimeCheck: React.FC<{ onComplete: (data: { sessionName: string; status: 'Hadir' | 'Terlambat' }) => void }> = ({ onComplete }) => {
  const { sessions } = useAppContext();
  const [currentTime, setCurrentTime] = useState(new Date());
  
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  const activeSession = useMemo(() => {
    const currentMinutes = currentTime.getHours() * 60 + currentTime.getMinutes();
    for (const session of sessions) {
      if (!session.isActive) continue;
      const [startH, startM] = session.startTime.split(':').map(Number);
      const [endH, endM] = session.endTime.split(':').map(Number);
      const startTotal = startH * 60 + startM;
      const endTotal = endH * 60 + endM;
      
      // LOGIKA UPGRADE: Toleransi terlambat ditambahkan DARI JAM BERAKHIR
      const endWithTolerance = endTotal + session.toleranceMinutes;
      
      // Jika waktu saat ini berada di dalam rentang awal hingga batas akhir + toleransi
      if (currentMinutes >= startTotal && currentMinutes <= endWithTolerance) {
        // Terlambat dihitung jika waktu absen sudah LEWAT DARI jam berakhir normal
        const isLate = currentMinutes > endTotal;
        return { session, status: isLate ? 'Terlambat' : 'Hadir', endTotal, endWithTolerance };
      }
    }
    return null;
  }, [currentTime, sessions]);

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-8 space-y-6 md:space-y-8 animate-in fade-in zoom-in duration-500 w-full">
      <div className="relative animate-[pulse_4s_ease-in-out_infinite]">
        <div className="absolute inset-0 bg-cyan-600/30 blur-[40px] rounded-full"></div>
        <Clock className="w-20 h-20 md:w-24 md:h-24 text-cyan-400 relative z-10 drop-shadow-[0_0_20px_rgba(6,182,212,0.8)]" />
      </div>
      <div className="text-center space-y-2 relative z-10">
        <h2 className="text-4xl md:text-5xl font-black text-white tracking-widest drop-shadow-md font-mono">
          {format(currentTime, 'HH:mm:ss')}
        </h2>
        <p className="text-sm md:text-base text-cyan-400/70 font-medium uppercase tracking-widest">{currentTime.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
      </div>

      <div className="w-full max-w-md bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-5 md:p-6 rounded-3xl shadow-[0_20px_50px_rgba(0,0,0,0.5)] transition-all duration-300 relative z-10">
        {activeSession ? (
          <div className="space-y-6">
            <div className="flex items-center justify-between p-4 bg-[#050B14]/80 border border-cyan-500/10 rounded-2xl relative overflow-hidden">
              <div className="absolute left-0 top-0 w-1 h-full bg-cyan-500"></div>
              <div className="pl-2">
                <p className="text-cyan-50 font-black text-lg md:text-xl tracking-wide">{activeSession.session.name}</p>
                <p className="text-xs md:text-sm text-cyan-400/80 font-mono mt-1">{activeSession.session.startTime} - {activeSession.session.endTime}</p>
                <p className="text-[10px] md:text-xs text-slate-500 mt-2 font-mono uppercase">
                  Batas Hadir: {activeSession.session.endTime} | Tutup Sesi: {Math.floor(activeSession.endWithTolerance / 60).toString().padStart(2, '0')}:{(activeSession.endWithTolerance % 60).toString().padStart(2, '0')}
                </p>
              </div>
              <span className={cn("px-4 py-2 text-xs font-black uppercase tracking-widest rounded-xl shadow-lg border", activeSession.status === 'Hadir' ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/30" : "bg-amber-500/10 text-amber-400 border-amber-500/30")}>
                {activeSession.status}
              </span>
            </div>
            <button onClick={() => onComplete({ sessionName: activeSession.session.name, status: activeSession.status as 'Hadir' | 'Terlambat' })} className="w-full py-4 bg-cyan-600 hover:bg-cyan-500 text-white font-black uppercase tracking-widest rounded-2xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.4)] hover:shadow-[0_0_30px_rgba(6,182,212,0.6)] flex items-center justify-center gap-3 active:scale-[0.98]">
              Lanjut Cek Lokasi <ScanFace className="w-5 h-5" />
            </button>
          </div>
        ) : (
          <div className="space-y-4 text-center">
            <div className="p-6 bg-rose-950/30 border border-rose-500/30 rounded-2xl text-rose-400 flex flex-col items-center gap-3">
              <AlertCircle className="w-12 h-12 animate-pulse drop-shadow-[0_0_15px_rgba(244,63,94,0.5)]" />
              <div>
                 <p className="font-black text-lg uppercase tracking-widest text-rose-200">Belum Waktunya</p>
                 <p className="text-xs mt-2 font-mono text-rose-400/80">Saat ini tidak ada jadwal absensi yang sedang aktif (atau sesi telah tertutup).</p>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

const LocationCheck: React.FC<{ onComplete: (loc: {lat: number, lng: number}) => void }> = ({ onComplete }) => {
  const { geofence } = useAppContext();
  const [status, setStatus] = useState<'loading' | 'success' | 'error'>('loading');
  const [distance, setDistance] = useState<number | null>(null);
  const [errorMsg, setErrorMsg] = useState('');

  const checkLocation = useCallback(() => {
    setStatus('loading');
    if (!navigator.geolocation) {
      setStatus('error'); setErrorMsg('Akses lokasi (GPS) pada browser Anda tidak diizinkan.'); return;
    }
    navigator.geolocation.getCurrentPosition(
      (position) => {
        const { latitude, longitude } = position.coords;
        const dist = calculateDistance(latitude, longitude, geofence.lat, geofence.lng);
        setDistance(dist);
        if (dist <= geofence.radius) {
          setStatus('success');
          setTimeout(() => onComplete({ lat: latitude, lng: longitude }), 1500);
        } else {
          setStatus('error');
          setErrorMsg(`Anda berada di luar jangkauan area absen (${geofence.name || 'Lokasi Kampus'}).`);
        }
      },
      (error) => { setStatus('error'); setErrorMsg('Sinyal GPS lemah atau belum diaktifkan. Pastikan GPS perangkat menyala.'); },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  }, [onComplete, geofence]);

  useEffect(() => { checkLocation(); }, [checkLocation]);

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-8 space-y-6 max-w-md mx-auto animate-in slide-in-from-right duration-500 w-full">
      <div className="w-32 h-32 bg-[#0A1628] rounded-full flex items-center justify-center border-2 border-cyan-500/20 relative overflow-hidden shadow-[0_0_40px_rgba(6,182,212,0.15)] z-10">
        {status === 'loading' && <div className="absolute inset-0 border-[6px] border-cyan-500/20 border-t-cyan-500 rounded-full animate-spin"></div>}
        {status === 'loading' && <div className="absolute inset-0 bg-cyan-500/10 rounded-full animate-pulse opacity-50"></div>}
        
        {/* Radar Line Effect */}
        {status === 'loading' && <div className="absolute top-1/2 left-1/2 w-[150%] h-[2px] bg-gradient-to-r from-transparent via-cyan-400 to-transparent origin-left animate-[spin_2s_linear_infinite] opacity-60"></div>}

        <Navigation className={cn("w-12 h-12 relative z-10 transition-colors duration-500", status === 'error' ? 'text-rose-400 drop-shadow-[0_0_15px_rgba(244,63,94,0.8)]' : 'text-cyan-400 drop-shadow-[0_0_15px_rgba(6,182,212,0.8)]')} />
      </div>

      <div className="text-center space-y-2 w-full bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-3xl shadow-[0_20px_50px_rgba(0,0,0,0.5)] transition-all duration-300 relative z-10">
        <h3 className="text-2xl font-black text-white tracking-widest uppercase">Cek Lokasi</h3>
        {status === 'loading' && <p className="text-cyan-400/60 font-mono text-xs uppercase tracking-widest mt-2">Mencari lokasi perangkat Anda...</p>}
        
        {status === 'success' && (
          <div className="text-emerald-400 space-y-3 animate-in fade-in zoom-in mt-6">
            <ActivitySquare className="w-16 h-16 mx-auto drop-shadow-[0_0_20px_rgba(16,185,129,0.8)]" />
            <div>
               <p className="font-black text-xl uppercase tracking-widest text-emerald-300">Lokasi Sesuai</p>
               <p className="text-xs text-emerald-500/80 font-mono mt-1">Jarak Anda: {Math.round(distance || 0)}m dari titik pusat absen.</p>
            </div>
          </div>
        )}
        
        {status === 'error' && (
          <div className="space-y-6 animate-in fade-in zoom-in mt-4">
            <div className="p-5 bg-rose-950/40 border border-rose-500/40 rounded-2xl text-center relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-1 bg-rose-500"></div>
              <p className="text-xs font-black text-rose-200 uppercase tracking-wide leading-relaxed">{errorMsg}</p>
              {distance && <p className="text-[10px] text-rose-400 mt-3 font-mono bg-rose-950 inline-block px-3 py-1.5 rounded-lg border border-rose-500/20">Jarak saat ini: {Math.round(distance)}m (Batas Maksimal: {geofence.radius}m)</p>}
            </div>
            <button onClick={checkLocation} className="w-full py-4 bg-transparent border border-cyan-500/50 hover:bg-cyan-500/10 text-cyan-400 rounded-2xl text-xs uppercase tracking-widest font-black flex items-center justify-center gap-3 transition-all duration-300 active:scale-[0.98]">
              <RefreshCcw className="w-4 h-4" /> Cek Ulang Lokasi
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

const QRScanner: React.FC<{ activeSessionName: string, onComplete: (data: {nim: string, name: string, deviceId: string, clusterName?: string}) => void }> = ({ activeSessionName, onComplete }) => {
  const { students, updateStudent, clusters, logs } = useAppContext();
  const [nimInput, setNimInput] = useState('');
  const [passInput, setPassInput] = useState('');
  const [error, setError] = useState('');
  const [isScanning, setIsScanning] = useState(false);
  const [isAutoLoggingIn, setIsAutoLoggingIn] = useState(true);
  
  const qrScannerRef = useRef<any>(null);

  // AUTO LOGIN MECHANISM & 1x ATTENDANCE PREVENTION
  useEffect(() => {
     const checkAutoLogin = () => {
        const ownerNim = localStorage.getItem('axaxyz_device_owner');
        const finalDeviceId = localStorage.getItem('axaxyz_device_id');
        
        if (ownerNim && finalDeviceId && students.length > 0) {
           const foundStudent = students.find(s => s.nim === ownerNim);
           if (foundStudent && foundStudent.deviceId === finalDeviceId) {
              const todayLocal = getLocalYYYYMMDD(new Date());
              const alreadyAttended = logs.some(l => l.nim === ownerNim && l.sessionName === activeSessionName && getLocalYYYYMMDD(l.timestamp) === todayLocal);
              
              if (alreadyAttended) {
                 setError(`⚠️ Halo ${foundStudent.name}, Anda sudah melakukan absensi pada sesi ${activeSessionName} untuk hari ini.`);
                 setIsAutoLoggingIn(false);
                 return;
              }

              // Auto-proceed
              const clusterName = foundStudent.clusterId ? clusters.find(c => c.id === foundStudent.clusterId)?.name : '';
              onComplete({ nim: ownerNim, name: foundStudent.name, deviceId: finalDeviceId, clusterName });
              return;
           }
        }
        setIsAutoLoggingIn(false);
     };

     // Slight delay to allow state hydration
     const timer = setTimeout(checkAutoLogin, 800);
     return () => clearTimeout(timer);
  }, [students, logs, activeSessionName, clusters, onComplete]);


  const handleVerify = (scannedNim?: string) => {
    setError('');
    const targetNim = scannedNim || nimInput;
    if (!targetNim) { setError('Mohon masukkan atau scan NIM Anda terlebih dahulu.'); return; }

    const todayLocal = getLocalYYYYMMDD(new Date());
    const alreadyAttended = logs.some(l => l.nim === targetNim && l.sessionName === activeSessionName && getLocalYYYYMMDD(l.timestamp) === todayLocal);
    
    let studentName = 'Mahasiswa Belum Terdaftar';
    
    if (students.length > 0) {
      const foundStudent = students.find(s => s.nim === targetNim);
      if (foundStudent) studentName = foundStudent.name;
    }

    if (alreadyAttended) {
      setError(`⚠️ Gagal: ${studentName} sudah tercatat hadir pada sesi ${activeSessionName} hari ini.`);
      return;
    }

    let clusterName = '';
    let finalDeviceId = localStorage.getItem('axaxyz_device_id');
    
    if (!finalDeviceId) {
      finalDeviceId = Math.random().toString(36).substring(2, 15);
      localStorage.setItem('axaxyz_device_id', finalDeviceId);
    }
    
    if (students.length > 0) {
      if (!passInput && !scannedNim) { setError('Kata sandi diperlukan untuk melanjutkan.'); return; }
      const foundStudent = students.find(s => s.nim === targetNim);
      if (!foundStudent) {
        setError('Akses ditolak: NIM Anda belum terdaftar di database.'); return;
      }
      
      if (!scannedNim && foundStudent.password !== passInput) {
        setError('Akses ditolak: Kata sandi yang dimasukkan salah.'); return;
      }
      studentName = foundStudent.name;
      if (foundStudent.clusterId) {
         clusterName = clusters.find(c => c.id === foundStudent.clusterId)?.name || '';
      }

      if (foundStudent.deviceId && foundStudent.deviceId !== finalDeviceId) {
        setError('⚠️ Keamanan Sistem: Akun NIM ini sudah terikat (login) di HP/Perangkat lain. Hubungi Admin jika Anda mengganti HP.');
        return;
      }
      
      if (!foundStudent.deviceId) {
        updateStudent(foundStudent.id, { deviceId: finalDeviceId });
      }
    } else {
      studentName = 'Mode Bebas (Belum ada data di database)'; 
      let deviceOwner = localStorage.getItem('axaxyz_device_owner');
      if (!deviceOwner) {
        localStorage.setItem('axaxyz_device_owner', targetNim); 
      } else if (deviceOwner !== targetNim) {
        setError('⚠️ Keamanan Sistem: Perangkat/HP ini sudah digunakan untuk login akun mahasiswa lain.');
        return;
      }
    }

    // Bind owner to localstorage to show NIM in header
    localStorage.setItem('axaxyz_device_owner', targetNim);

    onComplete({ nim: targetNim, name: studentName, deviceId: finalDeviceId, clusterName });
  };

  const startScanner = async () => {
    setIsScanning(true);
    setError('');
    
    try {
      const { Html5Qrcode } = await import('html5-qrcode');
      
      setTimeout(() => {
        try {
          const html5QrCode = new Html5Qrcode("qr-reader-box");
          qrScannerRef.current = html5QrCode;
          
          html5QrCode.start(
            { facingMode: "environment" },
            { fps: 10, qrbox: { width: 250, height: 250 } },
            (decodedText: string) => {
              setNimInput(decodedText);
              html5QrCode.stop().then(() => {
                setIsScanning(false);
                handleVerify(decodedText); 
              });
            },
            () => {} 
          ).catch((err: any) => {
             setError('Gagal membuka kamera scanner. Pastikan izin kamera diberikan.');
             setIsScanning(false);
          });
        } catch (err) {
          setError('Terjadi masalah pada modul kamera perangkat Anda.');
          setIsScanning(false);
        }
      }, 100);
    } catch (error) {
      setError('Scanner QR belum siap. Coba muat ulang halaman.');
      setIsScanning(false);
    }
  };

  const stopScanner = () => {
    if (qrScannerRef.current) {
      qrScannerRef.current.stop().catch(() => {});
    }
    setIsScanning(false);
  };

  if (isAutoLoggingIn) {
     return (
        <div className="flex flex-col items-center justify-center p-8 space-y-6 w-full max-w-md mx-auto z-10 relative">
           <Loader2 className="w-12 h-12 text-cyan-400 animate-spin drop-shadow-[0_0_15px_rgba(6,182,212,0.8)]" />
           <p className="text-cyan-300 font-mono uppercase tracking-widest text-xs animate-pulse text-center">Menyiapkan sistem untuk Anda...</p>
        </div>
     );
  }

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-8 space-y-6 max-w-md mx-auto animate-in slide-in-from-right duration-500 w-full">
      <div className="w-full bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-3xl shadow-[0_20px_50px_rgba(0,0,0,0.5)] transition-all duration-300 z-10 relative">
        <div className="text-center mb-8">
          <div className="w-20 h-20 bg-[#050B14] border border-cyan-500/50 rounded-2xl flex items-center justify-center mx-auto mb-5 shadow-[0_0_30px_rgba(6,182,212,0.3)]">
            {isScanning ? <Camera className="w-10 h-10 text-cyan-400 animate-pulse drop-shadow-[0_0_10px_rgba(6,182,212,0.8)]" /> : <QrCode className="w-10 h-10 text-cyan-400 drop-shadow-[0_0_10px_rgba(6,182,212,0.8)]" />}
          </div>
          <h3 className="text-2xl font-black text-white mb-2 tracking-widest uppercase">Masuk Akun</h3>
          <p className="text-cyan-500/70 text-xs font-mono uppercase tracking-wide">Pindai QR Code atau Input Manual</p>
        </div>

        {isScanning ? (
          <div className="space-y-4 animate-in fade-in zoom-in">
             <div className="relative w-full rounded-2xl overflow-hidden border border-cyan-500 bg-black aspect-square shadow-[0_0_40px_rgba(6,182,212,0.3)]">
                <div id="qr-reader-box" className="w-full h-full opacity-80 mix-blend-screen"></div>
                
                {/* Scanner Grid Overlay */}
                <div className="absolute inset-0 pointer-events-none bg-[linear-gradient(transparent_50%,rgba(6,182,212,0.1)_50%)] bg-[length:100%_4px]"></div>
                
                <div className="absolute inset-0 pointer-events-none">
                   <div className="w-full h-full bg-gradient-to-b from-transparent via-cyan-500/30 to-transparent animate-[scan_2s_ease-in-out_infinite] border-b-2 border-cyan-400"></div>
                </div>
                
                {/* Corner Accents */}
                <div className="absolute top-4 left-4 w-8 h-8 border-t-4 border-l-4 border-cyan-400 rounded-tl-lg shadow-[0_0_10px_rgba(6,182,212,0.8)]"></div>
                <div className="absolute top-4 right-4 w-8 h-8 border-t-4 border-r-4 border-cyan-400 rounded-tr-lg shadow-[0_0_10px_rgba(6,182,212,0.8)]"></div>
                <div className="absolute bottom-4 left-4 w-8 h-8 border-b-4 border-l-4 border-cyan-400 rounded-bl-lg shadow-[0_0_10px_rgba(6,182,212,0.8)]"></div>
                <div className="absolute bottom-4 right-4 w-8 h-8 border-b-4 border-r-4 border-cyan-400 rounded-br-lg shadow-[0_0_10px_rgba(6,182,212,0.8)]"></div>
             </div>
             <p className="text-[10px] text-center text-cyan-400 font-mono uppercase tracking-widest animate-pulse">Menunggu pindaian QR Code...</p>
             <button onClick={stopScanner} className="w-full py-4 bg-transparent border border-rose-500/50 hover:bg-rose-500/10 text-rose-400 rounded-xl text-xs uppercase tracking-widest font-black transition-all duration-300 active:scale-[0.98]">Batalkan Kamera</button>
          </div>
        ) : (
          <div className="space-y-6">
            <button onClick={startScanner} className="w-full py-4 bg-cyan-600/20 hover:bg-cyan-600/40 border border-cyan-500/50 text-cyan-400 font-black tracking-widest uppercase text-xs rounded-2xl flex justify-center items-center gap-3 transition-all duration-300 active:scale-[0.98] shadow-[0_0_15px_rgba(6,182,212,0.2)]">
              <Camera className="w-5 h-5" /> Buka Kamera QR Scanner
            </button>
            
            <div className="relative flex items-center py-2 opacity-40">
               <div className="flex-grow border-t border-cyan-500/50"></div>
               <span className="flex-shrink-0 mx-4 text-cyan-300 text-[9px] font-mono tracking-[0.3em] uppercase">Atau Input Manual</span>
               <div className="flex-grow border-t border-cyan-500/50"></div>
            </div>

            <div className="space-y-1.5">
              <label className="text-[9px] text-cyan-500/80 font-bold uppercase tracking-[0.2em] ml-1">Nomor Induk Mahasiswa (NIM)</label>
              <div className="flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl overflow-hidden focus-within:border-cyan-400 focus-within:shadow-[0_0_15px_rgba(6,182,212,0.3)] transition-all duration-300">
                <div className="pl-4 pr-2 text-cyan-600"><Fingerprint className="w-5 h-5"/></div>
                <input type="text" placeholder="Ketik NIM..." className="w-full bg-transparent py-4 pr-4 text-cyan-50 font-mono outline-none placeholder-cyan-900/50" value={nimInput} onChange={(e) => setNimInput(e.target.value)} />
              </div>
            </div>
            
            {students.length > 0 && (
              <div className="space-y-1.5">
                <label className="text-[9px] text-cyan-500/80 font-bold uppercase tracking-[0.2em] ml-1">Kata Sandi (Password)</label>
                <div className="flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl overflow-hidden focus-within:border-cyan-400 focus-within:shadow-[0_0_15px_rgba(6,182,212,0.3)] transition-all duration-300">
                  <div className="pl-4 pr-2 text-cyan-600"><Key className="w-5 h-5"/></div>
                  <input type="password" placeholder="Ketik Sandi..." className="w-full bg-transparent py-4 pr-4 text-cyan-50 font-mono outline-none placeholder-cyan-900/50" value={passInput} onChange={(e) => setPassInput(e.target.value)} />
                </div>
              </div>
            )}
            
            {error && (
              <div className="p-4 bg-rose-950/40 border border-rose-500/40 rounded-xl flex items-start gap-3 animate-in shake">
                <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
                <p className="text-xs font-mono text-rose-200 leading-relaxed uppercase tracking-wide">{error}</p>
              </div>
            )}

            <button onClick={() => handleVerify()} className="w-full py-4 mt-4 bg-cyan-600 hover:bg-cyan-500 text-white font-black tracking-widest uppercase text-xs rounded-2xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.4)] active:scale-[0.98]">
              Lanjut Verifikasi
            </button>
          </div>
        )}
      </div>
      <p className="text-[9px] text-cyan-600/60 text-center max-w-xs uppercase tracking-[0.2em] font-mono z-10 relative">
        Data dilindungi dengan sistem Device Fingerprinting
      </p>

      <style>{`
        @keyframes scan {
          0% { transform: translateY(-100%); }
          50% { transform: translateY(100%); }
          100% { transform: translateY(-100%); }
        }
      `}</style>
    </div>
  );
};

const SelfieCapture: React.FC<{ onComplete: (base64: string) => void }> = ({ onComplete }) => {
  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const [image, setImage] = useState<string | null>(null);

  const startCamera = async () => {
    try {
      const mediaStream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "user", aspectRatio: 16 / 9 }
      });
      streamRef.current = mediaStream;
      if (videoRef.current) {
        videoRef.current.srcObject = mediaStream;
      }
    } catch (err) {
      console.error("Camera access failed:", err);
    }
  };

  useEffect(() => {
    startCamera();
    return () => {
      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
      }
    };
  }, []);

  const capture = useCallback(() => {
    if (videoRef.current) {
      const canvas = document.createElement('canvas');
      canvas.width = videoRef.current.videoWidth;
      canvas.height = videoRef.current.videoHeight;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        // PERBAIKAN MIRRORING: Hanya draw image tanpa scale(-1, 1) agar foto asli
        ctx.drawImage(videoRef.current, 0, 0, canvas.width, canvas.height);
        setImage(canvas.toDataURL('image/jpeg', 0.8));
      }
    }
  }, []);

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-6 space-y-6 w-full max-w-md mx-auto animate-in slide-in-from-right duration-500 z-10 relative">
      <div className="text-center">
        <h3 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Foto Bukti Hadir</h3>
        <p className="text-cyan-500/70 text-xs font-mono uppercase tracking-wide mt-2">Posisikan wajah Anda di tengah kotak kamera</p>
      </div>

      <div className="w-full bg-[#050B14] rounded-3xl overflow-hidden border-2 border-cyan-500/50 relative shadow-[0_0_40px_rgba(6,182,212,0.3)] aspect-[3/4] md:aspect-video flex items-center justify-center transition-all duration-500 group">
        
        {!image ? (
          <video
            ref={videoRef}
            autoPlay
            playsInline
            muted
            className="w-full h-full object-cover transform scale-x-[-1]"
          />
        ) : (
          <img src={image} alt="Selfie Absen" className="w-full h-full object-cover" />
        )}
        
        {/* Overlay X-Ray Style Grid */}
        <div className="absolute inset-0 pointer-events-none bg-[linear-gradient(transparent_95%,rgba(6,182,212,0.1)_100%),linear-gradient(90deg,transparent_95%,rgba(6,182,212,0.1)_100%)] bg-[length:40px_40px]"></div>

        {!image && (
          <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
             
             {/* Face Tracking HUD */}
             <div className="relative w-56 h-72">
                {/* HUD Corners */}
                <div className="absolute top-0 left-0 w-8 h-8 border-t-2 border-l-2 border-cyan-400"></div>
                <div className="absolute top-0 right-0 w-8 h-8 border-t-2 border-r-2 border-cyan-400"></div>
                <div className="absolute bottom-0 left-0 w-8 h-8 border-b-2 border-l-2 border-cyan-400"></div>
                <div className="absolute bottom-0 right-0 w-8 h-8 border-b-2 border-r-2 border-cyan-400"></div>
                
                {/* Scanning Line */}
                <div className="absolute w-full h-[1px] bg-cyan-400 shadow-[0_0_10px_rgba(6,182,212,1)] animate-[scan_3s_ease-in-out_infinite]"></div>
                
                {/* Crosshair */}
                <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-4 h-4">
                   <div className="absolute top-1/2 left-0 w-full h-[1px] bg-cyan-500/50"></div>
                   <div className="absolute top-0 left-1/2 w-[1px] h-full bg-cyan-500/50"></div>
                </div>
             </div>
          </div>
        )}
      </div>

      <div className="w-full mt-4">
        {!image ? (
          <button onClick={capture} className="w-full py-5 bg-transparent border-2 border-cyan-500 hover:bg-cyan-500/20 text-cyan-400 font-black tracking-widest uppercase text-sm rounded-2xl transition-all duration-300 flex items-center justify-center gap-4 active:scale-[0.98] shadow-[0_0_20px_rgba(6,182,212,0.3)] inset-shadow">
            <div className="w-6 h-6 rounded-full border-[3px] border-cyan-400 flex items-center justify-center">
              <div className="w-2 h-2 rounded-full bg-cyan-400 animate-ping"></div>
            </div>
            Ambil Foto Selfie
          </button>
        ) : (
          <div className="flex gap-4">
            <button onClick={() => { setImage(null); startCamera(); }} className="flex-1 py-4 bg-transparent border border-rose-500/50 hover:bg-rose-500/10 text-rose-400 font-bold uppercase tracking-widest text-xs rounded-xl transition-all duration-300 active:scale-[0.95]">Ulangi Foto</button>
            <button onClick={() => onComplete(image)} className="flex-[2] py-4 bg-cyan-600 hover:bg-cyan-500 text-white font-black tracking-widest uppercase text-xs rounded-xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.4)] flex items-center justify-center gap-3 active:scale-[0.95]">
              <CheckCircle2 className="w-5 h-5" /> Gunakan Foto Ini
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

const SuccessScreen: React.FC<{ reset: () => void }> = ({ reset }) => (
  <div className="flex flex-col items-center justify-center p-8 space-y-8 text-center animate-in zoom-in duration-500 w-full z-10 relative">
    <div className="relative">
      <div className="absolute inset-0 bg-emerald-500/20 blur-[50px] rounded-full"></div>
      <div className="w-36 h-36 bg-[#050B14] border border-emerald-500/50 rounded-full flex items-center justify-center relative z-10 animate-bounce shadow-[0_0_40px_rgba(16,185,129,0.3)]">
         <ShieldCheck className="w-20 h-20 text-emerald-400 drop-shadow-[0_0_15px_rgba(16,185,129,0.8)]" />
      </div>
    </div>
    <div className="space-y-4">
      <h2 className="text-3xl md:text-4xl font-black text-emerald-300 tracking-widest uppercase">Absensi Berhasil!</h2>
      <p className="text-cyan-500/70 text-xs font-mono uppercase tracking-wide max-w-sm mx-auto leading-relaxed">
        Data kehadiran, jam, lokasi, dan foto selfie Anda telah berhasil disimpan ke dalam sistem.
      </p>
    </div>
    <button onClick={reset} className="px-10 py-4 bg-transparent border border-cyan-500 hover:bg-cyan-500/20 text-cyan-400 rounded-2xl transition-all duration-300 font-black tracking-widest uppercase text-sm mt-8 shadow-[0_0_20px_rgba(6,182,212,0.2)] active:scale-95">
      Selesai
    </button>
  </div>
);

const AttendanceWizard: React.FC = () => {
  const { addLog, studentLogout } = useAppContext();
  const [step, setStep] = useState(0); // 0 = Dashboard
  const [data, setData] = useState<Partial<Log>>({});
  const [linkedNim, setLinkedNim] = useState<string | null>(null);

  const activeSessionRef = useRef<string>('');

  useEffect(() => {
     setLinkedNim(localStorage.getItem('axaxyz_device_owner'));
  }, [step]); // re-check on step change

  const reset = () => { setStep(0); setData({}); activeSessionRef.current = ''; };
  const steps = ['Waktu', 'Lokasi', 'Identitas', 'Verifikasi'];

  return (
    <div className="min-h-screen flex flex-col font-sans text-slate-100 overflow-hidden relative radiology-bg">
      {/* Dark Theme Global Styling */}
      <style>{`
        .radiology-bg {
           background-color: #020617;
           background-image: 
             radial-gradient(circle at 15% 50%, rgba(6, 182, 212, 0.08), transparent 25%),
             radial-gradient(circle at 85% 30%, rgba(59, 130, 246, 0.08), transparent 25%),
             linear-gradient(rgba(6, 182, 212, 0.03) 1px, transparent 1px),
             linear-gradient(90deg, rgba(6, 182, 212, 0.03) 1px, transparent 1px);
           background-size: 100% 100%, 100% 100%, 30px 30px, 30px 30px;
           background-position: 0 0, 0 0, 0 0, 0 0;
           animation: pulse-bg 10s ease-in-out infinite alternate;
        }
        @keyframes pulse-bg {
           0% { background-color: #020617; }
           100% { background-color: #050b14; }
        }
      `}</style>
      
      <header className="w-full p-4 md:p-6 flex justify-between items-center relative z-20 border-b border-cyan-500/20 bg-[#0A1628]/80 backdrop-blur-xl shadow-[0_4px_30px_rgba(0,0,0,0.5)]">
        <div className="flex items-center gap-4">
          <div className="w-12 h-12 bg-[#050B14] border border-cyan-500/50 rounded-xl flex items-center justify-center shadow-[0_0_15px_rgba(6,182,212,0.4)] overflow-hidden p-2">
             <img src="/axalogo.png" alt="ABSENSI DEPT. RKG" className="w-full h-full object-contain filter drop-shadow-[0_0_5px_rgba(6,182,212,0.8)]" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
             <ActivitySquare className="text-cyan-400 w-full h-full hidden" />
          </div>
          <div className="flex flex-col">
             <span className="font-black text-lg md:text-2xl tracking-[0.2em] text-cyan-50 uppercase drop-shadow-[0_0_10px_rgba(6,182,212,0.5)]">DEPT. RKG</span>
             <span className="text-[8px] md:text-[10px] text-cyan-400 font-mono tracking-widest uppercase mt-0.5">Sistem Absensi Mahasiswa</span>
          </div>
        </div>
        
        <div className="flex items-center gap-3">
           {linkedNim && (
              <div className="hidden md:flex items-center gap-2 px-3 py-1.5 bg-[#050B14] border border-cyan-500/30 rounded-lg">
                 <User className="w-3.5 h-3.5 text-cyan-400" />
                 <span className="text-[10px] font-mono text-cyan-300 tracking-widest">{linkedNim}</span>
              </div>
           )}
           {linkedNim ? (
              <button onClick={studentLogout} className="text-[9px] md:text-[10px] font-black px-4 py-2 bg-rose-950/50 border border-rose-500/50 hover:bg-rose-500 hover:text-white rounded-lg text-rose-400 tracking-[0.15em] uppercase shadow-[0_0_10px_rgba(244,63,94,0.2)] transition-all flex items-center gap-2">
                 <LogOut className="w-3 h-3" /> Keluar
              </button>
           ) : (
              <div className="text-[9px] md:text-[10px] font-black px-4 py-2 bg-cyan-950/50 border border-cyan-500/50 rounded-lg text-cyan-300 tracking-[0.15em] uppercase shadow-[0_0_10px_rgba(6,182,212,0.2)]">Portal Mahasiswa</div>
           )}
        </div>
      </header>

      <main className="flex-1 flex flex-col relative z-10 w-full max-w-[1400px] mx-auto px-4 py-6 md:py-10 overflow-y-auto custom-scrollbar">
        {step > 0 && step < 5 && (
          <div className="mb-8 md:mb-16 max-w-2xl mx-auto w-full px-2 relative z-20">
            <div className="flex justify-between relative">
              <div className="absolute top-1/2 -translate-y-1/2 left-0 w-full h-[2px] bg-cyan-950"></div>
              <div className="absolute top-1/2 -translate-y-1/2 left-0 h-[2px] bg-cyan-400 transition-all duration-700 ease-in-out shadow-[0_0_10px_rgba(6,182,212,0.8)]" style={{ width: `${((step - 1) / (steps.length - 1)) * 100}%` }}></div>
              {steps.map((label, idx) => {
                const isActive = step === idx + 1; const isPassed = step > idx + 1;
                return (
                  <div key={label} className="relative z-10 flex flex-col items-center gap-3">
                    <div className={cn("w-10 h-10 md:w-12 md:h-12 rounded-full flex items-center justify-center text-xs md:text-sm font-black border-2 transition-all duration-500 bg-[#050B14]", isActive ? "border-cyan-400 text-cyan-400 shadow-[0_0_20px_rgba(6,182,212,0.6)] scale-110" : isPassed ? "border-cyan-600 text-cyan-500" : "border-cyan-900 text-cyan-800")}>
                      <div>{isPassed ? <CheckCircle2 className="w-5 h-5 md:w-6 md:h-6" /> : idx + 1}</div>
                    </div>
                    <span className={cn("text-[9px] md:text-[10px] font-mono absolute -bottom-7 w-max tracking-widest uppercase", isActive ? "text-cyan-400 font-bold" : isPassed ? "text-cyan-600" : "text-cyan-900")}>{label}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        <div className="flex-1 flex items-center justify-center w-full pt-4">
          {step === 0 && <StudentDashboard onStartAbsen={() => setStep(1)} linkedNim={linkedNim} />}
          {step === 1 && <TimeCheck onComplete={(d) => { activeSessionRef.current = d.sessionName; setData(prev => ({...prev, ...d})); setStep(2); }} />}
          {step === 2 && <LocationCheck onComplete={(d) => { setData(prev => ({...prev, location: d})); setStep(3); }} />}
          {step === 3 && <QRScanner activeSessionName={activeSessionRef.current} onComplete={(d) => { setData(prev => ({...prev, ...d})); setStep(4); }} />}
          {step === 4 && <SelfieCapture onComplete={(photo) => { addLog({ ...data, photoBase64: photo } as Omit<Log, 'id' | 'timestamp'>); setStep(5); }} />}
          {step === 5 && <SuccessScreen reset={reset} />}
        </div>
      </main>

      <footer className="text-center py-4 text-[10px] md:text-xs text-cyan-600/60 font-mono tracking-widest relative z-50 w-full bg-black/20">
        <a href="/ourteam" className="hover:text-cyan-400 hover:drop-shadow-[0_0_8px_rgba(6,182,212,0.8)] transition-all duration-300 cursor-pointer">
          Copyright © 2026 DEPT. RKG RSIGM UMI— All Rights Reserved. Made with ❤️
        </a>
      </footer>
      
      <style>{`
        .custom-scrollbar::-webkit-scrollbar { width: 4px; height: 4px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #050B14; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: rgba(6, 182, 212, 0.3); border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: rgba(6, 182, 212, 0.6); }
      `}</style>
    </div>
  );
};


// ==========================================
// ADMIN AREA
// ==========================================
const AdminLogin: React.FC<{ onLogin: () => void }> = ({ onLogin }) => {
  const { admins } = useAppContext();
  const [user, setUser] = useState('');
  const [pass, setPass] = useState('');
  const [err, setErr] = useState('');
  const [isLoading, setIsLoading] = useState(false);
  const [showPass, setShowPass] = useState(false);
  const [attempts, setAttempts] = useState(0);
  const [lockoutTimer, setLockoutTimer] = useState(0);

  useEffect(() => {
    let timer: NodeJS.Timeout;
    if (lockoutTimer > 0) timer = setTimeout(() => setLockoutTimer(lockoutTimer - 1), 1000);
    return () => clearTimeout(timer);
  }, [lockoutTimer]);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (lockoutTimer > 0) {
      setErr(`Sistem terkunci. Silakan coba lagi dalam ${lockoutTimer} detik.`);
      return;
    }
    setIsLoading(true); setErr('');
    await new Promise(resolve => setTimeout(resolve, 800));

    // Verify Env Fallback OR Custom Admin Logic
    const envUser = process.env.NEXT_PUBLIC_ADMIN_USER;
    const envPass = process.env.NEXT_PUBLIC_ADMIN_PASS;

    const isEnvMatch = (envUser && envPass && user === envUser && pass === envPass);
    const isDbMatch = admins.some(a => a.username === user && a.password === pass);

    if (isEnvMatch || isDbMatch) {
      setAttempts(0);
      localStorage.setItem('axaxyz_admin_auth', 'true');
      onLogin();
    } else {
      const newAttempts = attempts + 1;
      setAttempts(newAttempts);
      if (newAttempts >= 3) {
        setLockoutTimer(30); 
        setErr('❌ Akses ditolak. Anda diblokir sementara (30 detik).');
      } else {
        setErr(`❌ Username atau Password salah. (Sisa percobaan: ${3 - newAttempts})`);
      }
    }
    setIsLoading(false);
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-4 relative overflow-hidden w-full radiology-bg">
      <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[400px] h-[400px] bg-cyan-600/10 rounded-full blur-[100px] pointer-events-none"></div>

      <div className="w-full max-w-md bg-[#0A1628]/90 backdrop-blur-3xl border border-cyan-500/20 p-8 md:p-10 rounded-3xl shadow-[0_20px_60px_rgba(0,0,0,0.8)] relative z-10 animate-in slide-in-from-bottom-8 fade-in duration-700">
        <div className="flex flex-col items-center mb-10">
          <div className="relative mb-6">
            <div className="w-24 h-24 bg-[#050B14] border-2 border-cyan-500/50 rounded-2xl flex items-center justify-center shadow-[0_0_40px_rgba(6,182,212,0.4)] p-4 overflow-hidden transform rotate-45">
              <div className="-rotate-45 w-full h-full">
                 <img src="/axalogo.png" alt="RKG" className="w-full h-full object-contain filter drop-shadow-[0_0_8px_rgba(6,182,212,0.8)]" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
                 <ShieldCheck className="w-full h-full text-cyan-400 hidden" />
              </div>
            </div>
            <div className="absolute -bottom-2 -right-2 w-10 h-10 bg-[#0A1628] rounded-full border-2 border-cyan-500/50 flex items-center justify-center shadow-[0_0_15px_rgba(6,182,212,0.5)]">
              <Lock className="w-4 h-4 text-cyan-400" />
            </div>
          </div>
          <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-[0.2em] uppercase">Login Admin</h2>
          <p className="text-cyan-500/80 text-[10px] md:text-xs mt-2 uppercase tracking-[0.3em] font-mono">Panel Kelola Absensi</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-6">
          {err && (
            <div className="p-4 bg-rose-950/40 border border-rose-500/40 text-rose-300 text-xs font-mono rounded-xl flex items-start gap-3 animate-in shake duration-300 uppercase tracking-wider">
              <AlertCircle className="w-5 h-5 shrink-0" />
              <p className="leading-tight mt-0.5">{err}</p>
            </div>
          )}

          <div className="space-y-2">
            <label className="text-[9px] text-cyan-500/80 font-bold uppercase tracking-[0.2em] ml-1">Username</label>
            <div className="relative flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl overflow-hidden focus-within:border-cyan-400 transition-all duration-300 shadow-[inset_0_2px_10px_rgba(0,0,0,0.5)]">
              <div className="pl-4 pr-3 text-cyan-600"><User className="w-4 h-4"/></div>
              <input type="text" value={user} onChange={e=>setUser(e.target.value)} disabled={lockoutTimer > 0 || isLoading} className="w-full bg-transparent py-4 pr-4 text-cyan-50 font-mono outline-none placeholder-cyan-900/50 disabled:opacity-50 text-sm" placeholder="Ketik Username..." required />
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-[9px] text-cyan-500/80 font-bold uppercase tracking-[0.2em] ml-1">Password</label>
            <div className="relative flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl overflow-hidden focus-within:border-cyan-400 transition-all duration-300 shadow-[inset_0_2px_10px_rgba(0,0,0,0.5)]">
              <div className="pl-4 pr-3 text-cyan-600"><Key className="w-4 h-4"/></div>
              <input type={showPass ? 'text' : 'password'} value={pass} onChange={e=>setPass(e.target.value)} disabled={lockoutTimer > 0 || isLoading} className="w-full bg-transparent py-4 pr-12 text-cyan-50 font-mono outline-none placeholder-cyan-900/50 disabled:opacity-50 text-sm" placeholder="••••••••" required />
              <button type="button" onClick={() => setShowPass(!showPass)} className="absolute right-4 text-cyan-600 hover:text-cyan-400 transition-colors">
                {showPass ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
              </button>
            </div>
          </div>

          <button type="submit" disabled={lockoutTimer > 0 || isLoading} className="w-full py-4 mt-8 bg-cyan-600 hover:bg-cyan-500 disabled:bg-cyan-900 disabled:text-cyan-700 disabled:cursor-not-allowed text-white font-black tracking-[0.2em] uppercase text-xs rounded-xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.3)] hover:shadow-[0_0_30px_rgba(6,182,212,0.6)] flex justify-center items-center gap-3 active:scale-95 border border-cyan-400/50">
            {isLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : (
               <>Masuk ke Dashboard <ChevronRight className="w-4 h-4" /></>
            )}
          </button>
        </form>
      </div>

      <footer className="text-center py-4 text-[10px] md:text-xs text-cyan-600/60 font-mono tracking-widest relative z-50 w-full mt-auto">
        <a href="/ourteam" className="hover:text-cyan-400 hover:drop-shadow-[0_0_8px_rgba(6,182,212,0.8)] transition-all duration-300 cursor-pointer">
          Copyright © 2026 DEPT. RKG RSIGM UMI— All Rights Reserved. Made with ❤️
        </a>
      </footer>
    </div>
  );
};

const AdminDashboardHome: React.FC = () => {
  const { logs, students, clusters, sessions } = useAppContext();
  
  // Date Range and Filters
  const { startObj, endObj, FilterUI } = useDateFilter();
  const [selectedCluster, setSelectedCluster] = useState('All');

  // Dynamic Total Days Calculation for robust Alpha metrics
  const diffTime = Math.abs(endObj.getTime() - startObj.getTime());
  const totalDaysInRange = Math.ceil(diffTime / (1000 * 60 * 60 * 24)) || 1;

  // Filter Logs based on Date & Cluster (Waktu Lokal)
  const filteredLogs = logs.filter(l => {
    const logDate = new Date(l.timestamp);
    const inDateRange = logDate >= startObj && logDate <= endObj;
    
    let matchCluster = true;
    if (selectedCluster !== 'All') {
       matchCluster = l.clusterName === clusters.find(c => c.id === selectedCluster)?.name;
    }
    return inDateRange && matchCluster;
  });

  const filteredStudents = selectedCluster === 'All' ? students : students.filter(s => s.clusterId === selectedCluster);

  // Advanced Student Stats Detail Calculation (For Table)
  const activeSessions = sessions.filter(s => s.isActive);
  
  const studentStats = filteredStudents.map(student => {
     const studentLogs = filteredLogs.filter(l => l.nim === student.nim);
     let hadir = 0;
     let terlambat = 0;
     let alpha = 0;
     let belumAbsen = 0;
     
     // Evaluate each day in the date range
     const rangeStart = new Date(startObj);
     const rangeEnd = new Date(endObj); 
     const todayLocal = getLocalYYYYMMDD(new Date());

     for (let d = new Date(rangeStart); d <= rangeEnd; d.setDate(d.getDate() + 1)) {
         // Stop checking if day is in the future
         if (d > new Date()) break;

         const dateStrLocal = getLocalYYYYMMDD(d);
         const isToday = dateStrLocal === todayLocal;
         
         activeSessions.forEach(sess => {
             // Sinkronisasi Bug Fix: Validasi dengan konversi ke YYYY-MM-DD Lokal
             const log = studentLogs.find(l => getLocalYYYYMMDD(l.timestamp) === dateStrLocal && l.sessionName === sess.name);
             if (log) {
                 if (log.status === 'Hadir') hadir++;
                 else if (log.status === 'Terlambat') terlambat++;
             } else {
                 if (isToday) {
                     const currentMinutes = new Date().getHours() * 60 + new Date().getMinutes();
                     const [endH, endM] = sess.endTime.split(':').map(Number);
                     const endTotal = endH * 60 + endM;
                     const endWithTol = endTotal + sess.toleranceMinutes;
                     
                     if (currentMinutes > endWithTol) alpha++; // completely missed
                     else belumAbsen++; // Masih punya waktu untuk absen
                 } else {
                     // Check specifically if the loop date is strictly in the past
                     if (d < new Date(new Date().setHours(0,0,0,0))) {
                        alpha++; 
                     }
                 }
             }
         });
     }
     
     return { ...student, hadir, terlambat, alpha, belumAbsen };
  });

  // Basic Stats for Top Cards
  const totalLogsCount = filteredLogs.length;
  const onTimeCount = filteredLogs.filter(l => l.status === 'Hadir').length;
  const lateCount = filteredLogs.filter(l => l.status === 'Terlambat').length;
  const totalAlphaCount = studentStats.reduce((acc, curr) => acc + curr.alpha, 0);
  const totalBelumAbsenCount = studentStats.reduce((acc, curr) => acc + curr.belumAbsen, 0);

  // Chart 1: Daily Trend (Area Chart)
  const dailyDataMap: Record<string, { date: string; Hadir: number; Terlambat: number }> = {};
  filteredLogs.forEach(log => {
     const dateStr = new Date(log.timestamp).toLocaleDateString('id-ID', {day: 'numeric', month: 'short'});
     if (!dailyDataMap[dateStr]) dailyDataMap[dateStr] = { date: dateStr, Hadir: 0, Terlambat: 0 };
     if (log.status === 'Hadir') dailyDataMap[dateStr].Hadir++;
     else if (log.status === 'Terlambat') dailyDataMap[dateStr].Terlambat++;
  });
  const trendData = Object.values(dailyDataMap);
  
  // Chart 2: Pie Chart Overall
  const pieData = [
     { name: 'Tepat Waktu', value: onTimeCount, color: '#10b981' }, 
     { name: 'Terlambat', value: lateCount, color: '#f59e0b' },
     { name: 'Tidak Absen', value: totalAlphaCount, color: '#f43f5e' }
  ].filter(d => d.value > 0);

  return (
    <div className="space-y-6 animate-in fade-in duration-500 pb-10">
      
      {/* FILTER SECTION */}
      <div className="bg-[#0A1628]/80 backdrop-blur-md border border-cyan-500/20 p-5 md:p-6 rounded-[1.5rem] flex flex-col xl:flex-row gap-5 justify-between items-start xl:items-end shadow-lg">
         <div>
            <h2 className="text-xl md:text-2xl font-black text-cyan-50 tracking-widest uppercase">Dashboard Absensi</h2>
            <p className="text-cyan-500/70 text-xs font-mono uppercase mt-1">Ringkasan data kehadiran mahasiswa</p>
         </div>
         <div className="flex flex-col sm:flex-row gap-3 w-full xl:w-auto items-start sm:items-end">
            <FilterUI />
            <div className="flex flex-col gap-1 w-full sm:w-auto">
               <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Filter Kelompok / Angkatan</label>
               <div className="flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl px-2 h-11 w-full sm:w-auto focus-within:border-cyan-400 transition-colors">
                  <select value={selectedCluster} onChange={e => setSelectedCluster(e.target.value)} className="bg-transparent text-cyan-50 text-xs font-bold uppercase outline-none cursor-pointer px-3 w-full sm:min-w-[150px] h-full">
                     <option value="All">Semua Kelompok</option>
                     {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
               </div>
            </div>
         </div>
      </div>

      {/* STATS WIDGETS DENGAN 5 CARD */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-5 gap-4 md:gap-5">
        {[
          { title: 'Total Rekam Absen', val: totalLogsCount, icon: ActivitySquare, color: 'text-cyan-400', bg: 'bg-cyan-500/10', border: 'border-cyan-500/30' },
          { title: 'Tepat Waktu', val: onTimeCount, icon: CheckCircle2, color: 'text-emerald-400', bg: 'bg-emerald-500/10', border: 'border-emerald-500/30' },
          { title: 'Terlambat Hadir', val: lateCount, icon: Clock, color: 'text-amber-400', bg: 'bg-amber-500/10', border: 'border-amber-500/30' },
          { title: 'Data Kosong (Alpha)', val: totalAlphaCount, icon: UserX, color: 'text-rose-400', bg: 'bg-rose-500/10', border: 'border-rose-500/30' },
          { title: 'Belum Absen (Hari Ini)', val: totalBelumAbsenCount, icon: Clock, color: 'text-purple-400', bg: 'bg-purple-500/10', border: 'border-purple-500/30' }
        ].map((stat, i) => (
          <div key={i} className={`bg-[#0A1628]/60 backdrop-blur-md border ${stat.border} p-4 sm:p-5 rounded-[1.5rem] flex items-center justify-between transition-all duration-300 hover:bg-[#0A1628] hover:-translate-y-1 shadow-lg`}>
            <div>
               <p className="text-cyan-500/70 text-[8px] sm:text-[9px] font-bold uppercase tracking-[0.2em] mb-2 pr-2 leading-tight">{stat.title}</p>
               <h3 className="text-2xl sm:text-3xl font-black text-white font-mono">{stat.val}</h3>
            </div>
            <div className={cn("w-10 h-10 md:w-12 md:h-12 rounded-xl flex items-center justify-center shadow-inner border border-white/5 shrink-0", stat.bg)}>
               <stat.icon className={cn("w-5 h-5 md:w-6 md:h-6", stat.color)} />
            </div>
          </div>
        ))}
      </div>

      {/* CHARTS SECTION */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-5 min-h-[350px]">
        
        {/* TREND CHART */}
        <div className="lg:col-span-2 bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 p-6 rounded-[1.5rem] flex flex-col shadow-lg relative overflow-hidden">
          <div className="absolute top-0 right-0 w-32 h-32 bg-cyan-600/10 rounded-bl-[100px] pointer-events-none"></div>
          <h3 className="text-sm font-black text-cyan-50 mb-6 tracking-widest uppercase flex items-center gap-2"><Activity className="w-4 h-4 text-cyan-400"/> Tren Absensi Harian</h3>
          <div className="flex-1 w-full min-h-[250px] relative z-10">
            {trendData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={trendData} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                  <defs>
                    <linearGradient id="colorHadir" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#06b6d4" stopOpacity={0.8}/>
                      <stop offset="95%" stopColor="#06b6d4" stopOpacity={0}/>
                    </linearGradient>
                    <linearGradient id="colorTelat" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.8}/>
                      <stop offset="95%" stopColor="#f59e0b" stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                  <XAxis dataKey="date" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                  <YAxis stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                  <Tooltip cursor={{stroke: '#334155', strokeWidth: 2, fill: 'transparent'}} contentStyle={{backgroundColor: '#050B14', borderColor: '#06b6d4', color: '#f8fafc', borderRadius: '0.75rem', fontSize: '12px'}} />
                  <Area type="monotone" dataKey="Hadir" stroke="#06b6d4" strokeWidth={3} fillOpacity={1} fill="url(#colorHadir)" />
                  <Area type="monotone" dataKey="Terlambat" stroke="#f59e0b" strokeWidth={3} fillOpacity={1} fill="url(#colorTelat)" />
                </AreaChart>
              </ResponsiveContainer>
            ) : <div className="h-full flex items-center justify-center text-cyan-800 font-mono text-xs uppercase">Grafik Kosong (Tidak ada data)</div>}
          </div>
        </div>

        {/* PIE CHART */}
        <div className="bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 p-6 rounded-[1.5rem] flex flex-col shadow-lg">
          <h3 className="text-sm font-black text-cyan-50 mb-6 tracking-widest uppercase">Komposisi Kehadiran</h3>
          <div className="flex-1 w-full min-h-[250px]">
             {pieData.length > 0 ? (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={pieData} cx="50%" cy="50%" innerRadius="65%" outerRadius="85%" paddingAngle={5} dataKey="value">
                      {pieData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} stroke="rgba(0,0,0,0.5)" strokeWidth={2}/>)}
                    </Pie>
                    <Tooltip contentStyle={{backgroundColor: '#050B14', borderColor: '#1e293b', color: '#f8fafc', borderRadius: '0.75rem', fontSize: '12px'}} itemStyle={{color: '#fff'}} />
                  </PieChart>
                </ResponsiveContainer>
             ) : <div className="h-full flex items-center justify-center text-cyan-800 font-mono text-xs uppercase">Grafik Kosong</div>}
          </div>
          
          {/* Custom Legend */}
          <div className="flex justify-center gap-4 mt-2">
             {pieData.map(d => (
                <div key={d.name} className="flex items-center gap-1.5 text-[9px] font-bold uppercase tracking-wider text-slate-300">
                   <div className="w-2.5 h-2.5 rounded-full" style={{backgroundColor: d.color}}></div>
                   {d.name}
                </div>
             ))}
          </div>
        </div>
        
        {/* REKAPITULASI KEHADIRAN MAHASISWA (REPLACING BAR CHART BY SESSION) */}
        <div className="lg:col-span-3 bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 p-6 rounded-[1.5rem] flex flex-col shadow-lg relative overflow-hidden">
          <div className="absolute top-0 right-0 w-64 h-64 bg-cyan-600/5 rounded-bl-[100px] pointer-events-none"></div>
          
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-6 relative z-10 gap-3">
             <h3 className="text-sm font-black text-cyan-50 tracking-widest uppercase">Rincian Kehadiran Mahasiswa</h3>
             <div className="text-[10px] text-cyan-400 bg-cyan-950/50 px-3 py-1.5 rounded-lg border border-cyan-500/30 font-mono font-bold tracking-widest uppercase">
                {studentStats.length} Entitas Mahasiswa
             </div>
          </div>
          
          <div className="flex-1 w-full overflow-x-auto relative z-10 custom-scrollbar">
             <table className="w-full text-left border-collapse min-w-[800px]">
                <thead>
                   <tr className="bg-[#050B14]/80 border-b border-cyan-500/30 text-cyan-500 text-[10px] tracking-[0.2em] uppercase font-black">
                      <th className="p-4">NIM</th>
                      <th className="p-4">Nama Mahasiswa</th>
                      <th className="p-4">Kelompok</th>
                      <th className="p-4 text-center">Tepat Waktu</th>
                      <th className="p-4 text-center">Terlambat</th>
                      <th className="p-4 text-center">Tidak Hadir (Alpha)</th>
                      <th className="p-4 text-center">Belum Absen (Hari Ini)</th>
                   </tr>
                </thead>
                <tbody className="divide-y divide-cyan-900/30">
                   {studentStats.map((st, idx) => (
                      <tr key={st.id || idx} className="hover:bg-cyan-900/20 transition-colors duration-200 text-cyan-50 group">
                         <td className="p-4 font-mono text-sm tracking-wider">{st.nim}</td>
                         <td className="p-4 font-bold text-sm uppercase">{st.name}</td>
                         <td className="p-4">
                            <span className="text-[9px] uppercase font-bold tracking-widest text-cyan-300 bg-cyan-950/50 border border-cyan-500/30 px-2 py-1 rounded-md shadow-sm">
                               {clusters.find(c => c.id === st.clusterId)?.name || 'TANPA KELOMPOK'}
                            </span>
                         </td>
                         <td className="p-4 text-center">
                            <span className="inline-flex items-center justify-center w-8 h-8 rounded-full bg-emerald-500/10 text-emerald-400 font-bold text-sm border border-emerald-500/30 group-hover:bg-emerald-500/20">{st.hadir}</span>
                         </td>
                         <td className="p-4 text-center">
                            <span className="inline-flex items-center justify-center w-8 h-8 rounded-full bg-amber-500/10 text-amber-400 font-bold text-sm border border-amber-500/30 group-hover:bg-amber-500/20">{st.terlambat}</span>
                         </td>
                         <td className="p-4 text-center">
                            <span className="inline-flex items-center justify-center w-8 h-8 rounded-full bg-rose-500/10 text-rose-400 font-bold text-sm border border-rose-500/30 group-hover:bg-rose-500/20">{st.alpha}</span>
                         </td>
                         <td className="p-4 text-center">
                            <span className="inline-flex items-center justify-center w-8 h-8 rounded-full bg-slate-500/10 text-slate-300 font-bold text-sm border border-slate-500/30 group-hover:bg-slate-500/20">{st.belumAbsen}</span>
                         </td>
                      </tr>
                   ))}
                   {studentStats.length === 0 && (
                      <tr><td colSpan={7} className="p-12 text-center text-cyan-800 font-mono text-sm uppercase tracking-widest">Tidak ada data mahasiswa untuk filter ini.</td></tr>
                   )}
                </tbody>
             </table>
          </div>
        </div>

      </div>
    </div>
  );
};

const AdminClusters: React.FC = () => {
  const { clusters, addCluster, updateCluster, deleteCluster } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newC, setNewC] = useState({ name: '', startDate: '', endDate: '' });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editData, setEditData] = useState({ name: '', startDate: '', endDate: '' });

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if(newC.name.trim()) addCluster(newC);
    setIsAdding(false); setNewC({ name: '', startDate: '', endDate: '' });
  };

  const handleUpdate = (e: React.FormEvent) => {
     e.preventDefault();
     if(editingId && editData.name.trim()) updateCluster(editingId, editData);
     setEditingId(null); setEditData({ name: '', startDate: '', endDate: '' });
  }

  return (
    <div className="space-y-6 animate-in fade-in duration-500 pb-10">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Data Kelompok / Angkatan</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono uppercase mt-1">Kelola Pengelompokan Mahasiswa & Masa Stase</p>
        </div>
        <button onClick={() => setIsAdding(!isAdding)} className="flex items-center justify-center gap-2 px-5 py-3 bg-cyan-600/20 text-cyan-400 hover:bg-cyan-500/30 border border-cyan-500/50 rounded-xl transition-all duration-300 font-black uppercase tracking-widest text-xs shadow-[0_0_15px_rgba(6,182,212,0.2)] w-full md:w-auto">
          <Plus className="w-4 h-4" /> Tambah Kelompok Baru
        </button>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-[#0A1628]/80 backdrop-blur-md border border-cyan-500/30 p-5 md:p-6 rounded-2xl flex flex-col lg:flex-row gap-4 items-end shadow-xl animate-in slide-in-from-top-4">
          <div className="flex-1 space-y-1.5 w-full">
            <label className="text-[10px] md:text-xs text-cyan-500 font-bold uppercase tracking-widest ml-1">Nama Kelompok</label>
            <input required type="text" value={newC.name} onChange={e=>setNewC({...newC, name: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors text-sm font-mono" placeholder="Contoh: Angkatan 2025" />
          </div>
          <div className="flex-1 space-y-1.5 w-full">
            <label className="text-[10px] md:text-xs text-cyan-500 font-bold uppercase tracking-widest ml-1">Mulai Stase RKG</label>
            <input type="date" value={newC.startDate} onChange={e=>setNewC({...newC, startDate: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors text-sm font-mono" />
          </div>
          <div className="flex-1 space-y-1.5 w-full">
            <label className="text-[10px] md:text-xs text-cyan-500 font-bold uppercase tracking-widest ml-1">Akhir Stase RKG</label>
            <input type="date" value={newC.endDate} onChange={e=>setNewC({...newC, endDate: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors text-sm font-mono" />
          </div>
          <button type="submit" className="w-full lg:w-auto px-8 py-3.5 bg-cyan-600 hover:bg-cyan-500 text-white font-black uppercase tracking-widest text-xs rounded-xl transition-all duration-300 shadow-lg active:scale-95">Simpan</button>
        </form>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-3 gap-5">
         {clusters.map(c => (
            <div key={c.id} className="bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 p-5 md:p-6 rounded-2xl flex flex-col gap-4 group hover:border-cyan-500/50 transition-all duration-300 shadow-lg relative overflow-hidden">
               {editingId === c.id ? (
                  <form onSubmit={handleUpdate} className="flex flex-col gap-3 relative z-10">
                     <input autoFocus required type="text" placeholder="Nama Kelompok" value={editData.name} onChange={e=>setEditData({...editData, name: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/50 rounded-lg px-3 py-2 text-white outline-none text-sm font-mono" />
                     <div className="flex gap-2">
                        <input type="date" value={editData.startDate} onChange={e=>setEditData({...editData, startDate: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/50 rounded-lg px-3 py-2 text-white outline-none text-[10px] font-mono" />
                        <input type="date" value={editData.endDate} onChange={e=>setEditData({...editData, endDate: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/50 rounded-lg px-3 py-2 text-white outline-none text-[10px] font-mono" />
                     </div>
                     <div className="flex gap-2">
                        <button type="submit" className="flex-1 bg-emerald-500/20 text-emerald-400 p-2 rounded-lg border border-emerald-500/30 flex justify-center"><CheckCircle2 className="w-4 h-4"/></button>
                        <button type="button" onClick={()=>setEditingId(null)} className="flex-1 bg-rose-500/20 text-rose-400 p-2 rounded-lg border border-rose-500/30 flex justify-center"><X className="w-4 h-4"/></button>
                     </div>
                  </form>
               ) : (
                  <>
                     <div className="absolute top-0 right-0 w-24 h-24 bg-cyan-600/5 rounded-bl-[60px] pointer-events-none"></div>
                     <div className="flex items-start justify-between relative z-10">
                        <div className="flex items-center gap-3">
                           <div className="w-10 h-10 bg-cyan-950/50 rounded-xl flex items-center justify-center border border-cyan-500/30 shrink-0"><Network className="w-5 h-5 text-cyan-400" /></div>
                           <div>
                              <h3 className="font-bold text-white text-base tracking-wide truncate max-w-[150px]">{c.name}</h3>
                              <p className="text-[9px] text-cyan-500 font-mono uppercase tracking-widest mt-0.5">
                                 {c.startDate && c.endDate ? `${new Date(c.startDate).toLocaleDateString('id-ID',{day:'2-digit', month:'short'})} - ${new Date(c.endDate).toLocaleDateString('id-ID',{day:'2-digit', month:'short'})}` : 'Tanggal Stase Belum Diatur'}
                              </p>
                           </div>
                        </div>
                        <div className="flex gap-2 opacity-100 lg:opacity-0 lg:group-hover:opacity-100 transition-opacity shrink-0">
                           <button onClick={()=>{setEditingId(c.id); setEditData({name: c.name, startDate: c.startDate || '', endDate: c.endDate || ''});}} className="p-2 bg-blue-500/10 hover:bg-blue-500/20 text-blue-400 border border-blue-500/20 rounded-lg"><Edit className="w-4 h-4"/></button>
                           <button onClick={()=>{if(confirm(`Hapus kelompok ${c.name}? Data mahasiswa terkait akan terpengaruh.`)) deleteCluster(c.id);}} className="p-2 bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 border border-rose-500/20 rounded-lg"><Trash2 className="w-4 h-4"/></button>
                        </div>
                     </div>
                  </>
               )}
            </div>
         ))}
         {clusters.length === 0 && <div className="col-span-full p-8 text-center border-2 border-dashed border-cyan-900 rounded-2xl text-cyan-700 font-mono text-sm uppercase">Belum Ada Kelompok Terdaftar</div>}
      </div>
    </div>
  );
};

const AdminStudents: React.FC = () => {
  const { students, addStudent, updateStudent, bulkAddStudents, deleteStudent, clusters } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newS, setNewS] = useState({ name: '', nim: '', password: '', clusterId: '' });
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  
  const [search, setSearch] = useState('');
  const [selectedClusterForBulk, setSelectedClusterForBulk] = useState('');
  const [filterClusterDisplay, setFilterClusterDisplay] = useState('All'); 
  const [defaultBulkPassword, setDefaultBulkPassword] = useState('123'); // SUPER UPGRADE
  
  const [selectedStudentForKTM, setSelectedStudentForKTM] = useState<Student | null>(null);

  useEffect(() => {
     const savedPass = localStorage.getItem('axaxyz_default_bulk_pass');
     if (savedPass) setDefaultBulkPassword(savedPass);
  }, []);

  const handleDefaultPassChange = (e: React.ChangeEvent<HTMLInputElement>) => {
     setDefaultBulkPassword(e.target.value);
     localStorage.setItem('axaxyz_default_bulk_pass', e.target.value);
  }

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    addStudent({ ...newS, password: newS.password || defaultBulkPassword });
    setIsAdding(false);
    setNewS({ name: '', nim: '', password: '', clusterId: '' });
  };

  const handleUpdate = (e: React.FormEvent) => {
    e.preventDefault();
    if(editingStudent) {
       updateStudent(editingStudent.id, { name: editingStudent.name, nim: editingStudent.nim, password: editingStudent.password, clusterId: editingStudent.clusterId });
       setEditingStudent(null);
    }
  };

  const handleUnlinkDevice = (id: string, name: string) => {
     if(confirm(`Konfirmasi Pelepasan Akses Perangkat (Logout HP) untuk ${name}?`)) {
        updateStudent(id, { deviceId: null });
     }
  };

  // EXCEL (XLSX) BULK UPLOAD SUPPORT - DYNAMIC LOAD (SUPER UPGRADE)
  const handleBulkUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!selectedClusterForBulk) {
       alert("Pilih Kelompok (Angkatan) terlebih dahulu di tombol 'PILIH KELOMPOK (IMPORT)' sebelum upload file Excel.");
       e.target.value = ''; // reset input
       return;
    }

    try {
      const XLSX = await loadXlsx();
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      // Expecting standard columns, row 1 as header. (Nama, NIM, Kelompok)
      const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet);
      
      const newSt: Omit<Student, 'id'>[] = [];
      let missingClusterCount = 0;

      jsonData.forEach(row => {
         // Flexible matching for typical column names in Indonesia
         const name = row['Nama'] || row['NAMA'] || row['nama'] || row['Nama Lengkap'] || row['Name'];
         const nim = row['NIM'] || row['NIM '] || row['nim'] || row['Nomor Induk'];
         const clusterCol = row['Kelompok'] || row['KELOMPOK'] || row['Cluster'] || row['Angkatan'];
         
         if (name && nim) {
            let finalClusterId = selectedClusterForBulk;
            
            // Prioritaskan Kelompok dari File Excel jika kolom tersedia
            if (clusterCol) {
                const found = clusters.find(c => c.name.toLowerCase().trim() === String(clusterCol).toLowerCase().trim());
                if (found) {
                   finalClusterId = found.id;
                }
            }

            if (!finalClusterId) {
                missingClusterCount++;
            }

            newSt.push({ 
               name: String(name).trim(), 
               nim: String(nim).trim(), 
               password: defaultBulkPassword, // Menggunakan state dari input box
               clusterId: finalClusterId || ''
            });
         }
      });

      if (newSt.length > 0) {
        if (missingClusterCount > 0 && !selectedClusterForBulk) {
           if(!confirm(`⚠️ Peringatan Sistem: Terdapat ${missingClusterCount} data mahasiswa tanpa informasi kelompok (tidak ada di Excel & tidak dipilih di dropdown). Lanjutkan import?`)) {
               e.target.value = ''; return;
           }
        }
        bulkAddStudents(newSt);
        alert(`✅ Sistem Berhasil mengimpor ${newSt.length} mahasiswa dengan default password '${defaultBulkPassword}'.`);
      } else {
        alert('❌ Gagal mendeteksi data. Pastikan format kolom baris pertama memiliki header "Nama" dan "NIM".');
      }
    } catch (err) {
       console.error("Bulk Upload Error:", err);
       alert("❌ Terjadi kesalahan saat membaca file .xlsx");
    }
    
    e.target.value = ''; // reset input
  };

  const filtered = students.filter(s => {
      const matchesSearch = s.name.toLowerCase().includes(search.toLowerCase()) || s.nim.includes(search);
      const matchesCluster = filterClusterDisplay === 'All' || s.clusterId === filterClusterDisplay;
      return matchesSearch && matchesCluster;
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col w-full relative pb-10">
      
      {/* HEADER SECTION - REARRANGED TO BE BELOW TITLE */}
      <div className="flex flex-col gap-6">
        <div className="shrink-0">
          <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Data Mahasiswa</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1 uppercase">Kelola Data Mahasiswa dan Kartu Absen (QR)</p>
        </div>
        
        <div className="flex flex-col xl:flex-row gap-4 w-full items-start">
          {/* Action Buttons Row */}
          <div className="flex flex-col sm:flex-row gap-3 w-full xl:w-auto bg-[#0A1628]/80 p-3 rounded-2xl border border-cyan-500/20 shadow-lg items-center shrink-0">
             
             {/* KELOMPOK DROPDOWN */}
             <div className="flex flex-col w-full sm:w-auto gap-1">
               <div className="flex items-center bg-[#050B14] border border-purple-500/30 rounded-xl px-2 h-11 w-full sm:w-44 focus-within:border-purple-400 transition-colors">
                  <select value={selectedClusterForBulk} onChange={e=>setSelectedClusterForBulk(e.target.value)} className="bg-transparent text-purple-100 text-xs font-bold uppercase outline-none w-full cursor-pointer appearance-none px-2 text-center sm:text-left">
                     <option value="" disabled>PILIH KELOMPOK (DEFAULT)</option>
                     {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                  </select>
               </div>
             </div>

             {/* SUPER UPGRADE: INPUT DEFAULT PASSWORD */}
             <div className="flex flex-col w-full sm:w-auto gap-1">
               <div className="flex items-center bg-[#050B14] border border-purple-500/30 rounded-xl px-3 h-11 w-full sm:w-32 focus-within:border-purple-400 transition-colors" title="Sandi ini otomatis jadi password akun yang di-import dari Excel">
                  <Key className="w-3.5 h-3.5 text-purple-400 mr-2 shrink-0" />
                  <input 
                    type="text" 
                    value={defaultBulkPassword} 
                    onChange={handleDefaultPassChange} 
                    className="bg-transparent text-purple-100 text-xs font-bold w-full outline-none placeholder-purple-500/50"
                    placeholder="Sandi Default"
                  />
               </div>
             </div>
             
             <label className="flex flex-1 sm:flex-none w-full sm:w-auto justify-center items-center gap-2 px-6 py-3 sm:py-2.5 bg-gradient-to-r from-purple-600/30 to-fuchsia-600/30 text-purple-300 hover:from-purple-500/50 hover:to-fuchsia-500/50 border border-purple-400/50 rounded-xl transition-all duration-300 font-black uppercase text-[10px] md:text-xs cursor-pointer active:scale-95 shadow-[0_0_15px_rgba(147,51,234,0.3)] whitespace-nowrap">
                <Upload className="w-4 h-4" /> Import Excel
                <input type="file" accept=".xlsx, .xls" className="hidden" onChange={handleBulkUpload} />
             </label>
             
             <button onClick={() => setIsAdding(!isAdding)} className="flex flex-1 sm:flex-none w-full sm:w-auto justify-center items-center gap-2 px-6 py-3 sm:py-2.5 bg-cyan-600/20 text-cyan-400 hover:bg-cyan-600/40 border border-cyan-500/50 rounded-xl transition-all duration-300 font-black uppercase text-[10px] md:text-xs active:scale-95 shadow-[0_0_10px_rgba(6,182,212,0.3)] whitespace-nowrap">
               <Plus className="w-4 h-4" /> Input Manual
             </button>
          </div>

          {/* COMPREHENSIVE SYSTEMATIC GUIDE FOR EXCEL IMPORT */}
          <div className="bg-gradient-to-br from-[#050B14]/90 to-[#0A1628]/90 p-4 rounded-2xl border border-purple-500/40 flex flex-col sm:flex-row items-start sm:items-center gap-3 md:gap-4 shadow-[0_10px_30px_rgba(147,51,234,0.15)] relative overflow-hidden group w-full xl:flex-1">
             <div className="absolute top-0 right-0 w-32 h-32 bg-purple-600/10 rounded-bl-[100px] pointer-events-none transition-transform group-hover:scale-110"></div>
             <div className="bg-purple-950/60 p-2.5 rounded-xl border border-purple-500/50 shrink-0 relative z-10 shadow-[inset_0_0_15px_rgba(147,51,234,0.3)]">
                <FileText className="w-5 h-5 text-purple-400" />
             </div>
             <div className="flex-1 relative z-10 w-full">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center mb-2 gap-2">
                   <p className="text-[11px] md:text-xs text-purple-200 font-black uppercase tracking-[0.15em] flex items-center gap-2 drop-shadow-md">
                      Panduan Format Excel 
                   </p>
                   <span className="bg-gradient-to-r from-purple-600 to-fuchsia-600 text-white border border-purple-400/50 px-2.5 py-1 rounded-lg text-[8px] font-black tracking-[0.2em] shadow-[0_0_15px_rgba(192,38,211,0.5)]">
                      ✨ SUPER UPGRADE
                   </span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-x-4 gap-y-2 text-[9px] md:text-[10px] text-cyan-100/80 font-mono leading-relaxed">
                   <div className="space-y-1">
                      <p><span className="text-purple-400 font-bold bg-purple-950/50 px-1 rounded">Kolom A</span> Nama Lengkap <span className="text-purple-300 font-bold">*Wajib</span></p>
                   </div>
                   <div className="space-y-1">
                      <p><span className="text-purple-400 font-bold bg-purple-950/50 px-1 rounded">Kolom B</span> NIM / Stambuk <span className="text-purple-300 font-bold">*Wajib</span></p>
                   </div>
                   <div className="space-y-1 sm:col-span-2 md:col-span-1">
                      <p><span className="text-emerald-400 font-bold bg-emerald-950/50 px-1 rounded">Kolom C</span> Kelompok <span className="text-emerald-300 italic">(Opsional)</span></p>
                   </div>
                </div>
                <p className="text-[8px] text-cyan-500 italic mt-2.5 border-t border-cyan-900/50 pt-2">*Pastikan Baris 1 pada file Excel diisi Header. Jika kolom C kosong, data otomatis masuk ke kelompok dropdown. <strong className="text-emerald-400">Password default menyesuaikan input box di atas.</strong></p>
             </div>
          </div>
        </div>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-[#0A1628]/90 backdrop-blur-md border border-cyan-500/50 p-5 md:p-6 rounded-[1.5rem] grid grid-cols-1 md:grid-cols-5 gap-4 items-end animate-in slide-in-from-top-4 shadow-[0_10px_30px_rgba(0,0,0,0.5)]">
          <div className="space-y-1.5 md:col-span-2">
            <label className="text-[9px] md:text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Nama Mahasiswa</label>
            <input required type="text" value={newS.name} onChange={e=>setNewS({...newS, name: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm" placeholder="Nama Lengkap..." />
          </div>
          <div className="space-y-1.5">
            <label className="text-[9px] md:text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">NIM</label>
            <input required type="text" value={newS.nim} onChange={e=>setNewS({...newS, nim: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm" placeholder="Nomor Induk Mahasiswa..." />
          </div>
          <div className="space-y-1.5">
            <label className="text-[9px] md:text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Kelompok</label>
            <select required value={newS.clusterId} onChange={e=>setNewS({...newS, clusterId: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3 text-cyan-50 outline-none focus:border-cyan-400 font-bold text-xs uppercase appearance-none cursor-pointer">
               <option value="" disabled>Pilih Kelompok</option>
               {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
          <button type="submit" className="w-full py-3 bg-cyan-600 hover:bg-cyan-500 text-white font-black uppercase tracking-widest text-xs rounded-xl transition-all duration-300 shadow-[0_0_15px_rgba(6,182,212,0.4)] active:scale-95">Simpan Data</button>
        </form>
      )}

      {/* FILTER TAMPILAN TABEL */}
      <div className="flex flex-col sm:flex-row gap-3 w-full max-w-2xl items-end">
         <div className="flex flex-col gap-1 flex-1 w-full">
            <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Pencarian Mahasiswa</label>
            <div className="relative">
               <SearchIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-cyan-600" />
               <input type="text" placeholder="Cari Nama atau NIM..." value={search} onChange={e=>setSearch(e.target.value)} className="w-full bg-[#0A1628]/80 border border-cyan-500/30 rounded-2xl pl-11 pr-4 h-11 text-cyan-50 outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" />
            </div>
         </div>
         <div className="flex flex-col gap-1 w-full sm:w-56">
            <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Filter Kelompok</label>
            <select value={filterClusterDisplay} onChange={e=>setFilterClusterDisplay(e.target.value)} className="bg-[#0A1628]/80 border border-cyan-500/30 rounded-2xl px-5 h-11 text-cyan-50 outline-none focus:border-cyan-400 transition-colors shadow-inner w-full font-bold text-xs uppercase cursor-pointer appearance-none">
              <option value="All">Filter: Semua Kelompok</option>
              {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
         </div>
      </div>

      <div className="flex-1 bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 rounded-[1.5rem] overflow-hidden flex flex-col shadow-[0_15px_40px_rgba(0,0,0,0.5)] relative">
        <div className="overflow-x-auto custom-scrollbar">
          <table className="w-full text-left border-collapse min-w-[800px]">
            <thead>
              <tr className="bg-[#050B14]/80 border-b border-cyan-500/20 text-cyan-500 text-[10px] tracking-[0.2em] uppercase font-black">
                <th className="p-4 md:p-5 whitespace-nowrap">NIM</th>
                <th className="p-4 md:p-5 whitespace-nowrap">Nama Mahasiswa</th>
                <th className="p-4 md:p-5 whitespace-nowrap">Kelompok</th>
                <th className="p-4 md:p-5 text-center whitespace-nowrap">Status Login HP</th>
                <th className="p-4 md:p-5 text-right whitespace-nowrap">Pilihan</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-cyan-900/30">
              {filtered.map(st => (
                <tr key={st.id} className="hover:bg-cyan-900/20 transition-colors duration-200 text-cyan-50 group">
                  <td className="p-4 md:p-5 font-mono text-sm tracking-wider">{st.nim}</td>
                  <td className="p-4 md:p-5 font-bold text-sm uppercase max-w-[200px] truncate">{st.name}</td>
                  <td className="p-4 md:p-5">
                     <span className="text-[10px] uppercase font-bold tracking-widest text-cyan-300 bg-cyan-950/50 border border-cyan-500/30 px-3 py-1.5 rounded-md shadow-sm">
                        {clusters.find(c => c.id === st.clusterId)?.name || 'BELUM ADA KELOMPOK'}
                     </span>
                  </td>
                  <td className="p-4 md:p-5 text-center">
                    {st.deviceId ? (
                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-950/50 text-emerald-400 text-[10px] font-black uppercase tracking-widest border border-emerald-500/30 shadow-[0_0_10px_rgba(16,185,129,0.2)]"><CheckCircle2 className="w-3.5 h-3.5"/> Terhubung</span>
                    ) : (
                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-900 text-slate-500 text-[10px] font-black uppercase tracking-widest border border-white/10">Kosong</span>
                    )}
                  </td>
                  <td className="p-4 md:p-5 text-right flex justify-end gap-2">
                    {st.deviceId && (
                      <button onClick={() => handleUnlinkDevice(st.id, st.name)} title="Lepas Otoritas Perangkat" className="p-2 md:p-2.5 text-amber-500 hover:text-amber-300 rounded-xl transition-all duration-300 border border-amber-500/30 bg-amber-950/40 hover:bg-amber-900 hover:-translate-y-0.5 active:scale-95 shadow-sm">
                         <RefreshCcw className="w-4 h-4" />
                      </button>
                    )}
                    <button onClick={() => setEditingStudent(st)} title="Edit Data Mahasiswa" className="p-2 md:p-2.5 text-blue-500 hover:text-blue-300 rounded-xl transition-all duration-300 border border-blue-500/30 bg-blue-950/40 hover:bg-blue-900 hover:-translate-y-0.5 active:scale-95 shadow-sm">
                       <Settings className="w-4 h-4" />
                    </button>
                    {/* BUTTON KIRIM WA */}
                    <button onClick={() => { alert(`Sistem segera memanggil REST API (app/api/wa/route.ts) untuk mengirim Kredensial Onboarding (Skenario 25) ke ${st.name}...`); }} title="Kirim Akses ke WA" className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-emerald-400 bg-emerald-950/50 border border-emerald-500/40 hover:bg-emerald-600 hover:text-white rounded-xl transition-all duration-300 flex items-center gap-2 hover:-translate-y-0.5 active:scale-95 shadow-[0_0_10px_rgba(16,185,129,0.2)]">
                      <MessageCircle className="w-4 h-4" /> WA
                    </button>
                    <button onClick={() => setSelectedStudentForKTM(st)} title="Cetak Kartu Absen (QR)" className="px-3 py-2 text-[10px] font-black uppercase tracking-widest text-cyan-400 bg-cyan-950/50 border border-cyan-500/40 hover:bg-cyan-600 hover:text-white rounded-xl transition-all duration-300 flex items-center gap-2 hover:-translate-y-0.5 active:scale-95 shadow-[0_0_10px_rgba(6,182,212,0.2)]">
                      <ScanFace className="w-4 h-4"/> Cetak Kartu
                    </button>
                    <button onClick={() => {if(confirm(`Hapus permanen mahasiswa ${st.name}?`)) deleteStudent(st.id);}} title="Hapus Mahasiswa" className="p-2 md:p-2.5 text-rose-500 hover:text-white hover:bg-rose-600 rounded-xl transition-all duration-300 border border-rose-500/30 bg-rose-950/40 hover:-translate-y-0.5 active:scale-95 shadow-sm">
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && <tr><td colSpan={5} className="p-12 text-center text-cyan-800 font-mono text-sm uppercase tracking-widest">Tidak ada data mahasiswa ditemukan.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {/* EDIT MODAL */}
      {editingStudent && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 backdrop-blur-xl p-4 animate-in fade-in zoom-in-95 duration-200">
           <form onSubmit={handleUpdate} className="bg-[#0A1628] border border-cyan-500/40 p-6 md:p-8 rounded-3xl w-full max-w-md shadow-[0_0_50px_rgba(6,182,212,0.3)] relative radiology-bg">
              <div className="flex justify-between items-center mb-8">
                 <h3 className="text-xl md:text-2xl font-black text-cyan-50 tracking-widest uppercase">Edit Data Mahasiswa</h3>
                 <button type="button" onClick={() => setEditingStudent(null)} className="p-2 bg-rose-950/50 hover:bg-rose-500 hover:text-white border border-rose-500/30 rounded-xl transition-colors text-rose-400"><X className="w-5 h-5"/></button>
              </div>
              <div className="space-y-5">
                 <div className="space-y-1.5">
                    <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Nama Lengkap</label>
                    <input required type="text" value={editingStudent.name} onChange={e=>setEditingStudent({...editingStudent, name: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm shadow-inner" />
                 </div>
                 <div className="space-y-1.5">
                    <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">NIM</label>
                    <input required type="text" value={editingStudent.nim} onChange={e=>setEditingStudent({...editingStudent, nim: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm shadow-inner" />
                 </div>
                 <div className="space-y-1.5">
                    <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Ubah Sandi</label>
                    <input required type="text" value={editingStudent.password || ''} onChange={e=>setEditingStudent({...editingStudent, password: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm shadow-inner" />
                 </div>
                 <div className="space-y-1.5">
                    <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Pindah Kelompok</label>
                    <select required value={editingStudent.clusterId || ''} onChange={e=>setEditingStudent({...editingStudent, clusterId: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-bold text-xs uppercase appearance-none cursor-pointer">
                       <option value="" disabled>Pilih Kelompok</option>
                       {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                    </select>
                 </div>
                 <button type="submit" className="w-full py-4 mt-6 bg-cyan-600 hover:bg-cyan-500 text-white font-black tracking-widest uppercase text-xs rounded-xl transition-all duration-300 shadow-[0_10px_20px_rgba(6,182,212,0.4)] active:scale-95 border border-cyan-400/50">
                    Simpan Perubahan
                 </button>
              </div>
           </form>
        </div>
      )}

      {/* KTM PRINT MODAL */}
      {selectedStudentForKTM && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/95 backdrop-blur-xl p-4 animate-in fade-in zoom-in-95 duration-200">
          <style>{`
            @media print {
              body * { visibility: hidden; }
              #ktm-print-area, #ktm-print-area * { visibility: visible !important; }
              #ktm-print-area { position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); margin: 0; background-color: #050B14; }
            }
          `}</style>
          <div className="bg-[#0A1628] border border-cyan-500/30 p-6 md:p-8 rounded-[2rem] w-full max-w-[450px] shadow-[0_0_50px_rgba(6,182,212,0.3)] relative z-50">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-lg font-black text-cyan-50 tracking-widest uppercase">Cetak Kartu Absen</h3>
              <button onClick={() => setSelectedStudentForKTM(null)} className="p-2 bg-rose-950/50 hover:bg-rose-500 hover:text-white border border-rose-500/30 rounded-xl transition-colors text-rose-400"><X className="w-4 h-4"/></button>
            </div>
            
            <div id="ktm-print-area" className="w-[320px] md:w-[340px] h-[500px] md:h-[540px] mx-auto bg-[#050B14] rounded-[2rem] p-6 relative overflow-hidden shadow-2xl flex flex-col items-center justify-between border-[4px] border-cyan-500/50">
               
               {/* Graphics for Card */}
               <div className="absolute top-0 left-0 w-full h-2 bg-cyan-400"></div>
               <div className="absolute top-[-50px] right-[-50px] w-48 h-48 bg-cyan-600/20 rounded-full blur-[40px]"></div>
               <div className="absolute bottom-[-50px] left-[-50px] w-48 h-48 bg-purple-600/20 rounded-full blur-[40px]"></div>
               <div className="absolute inset-0 pointer-events-none bg-[linear-gradient(transparent_95%,rgba(6,182,212,0.1)_100%),linear-gradient(90deg,transparent_95%,rgba(6,182,212,0.1)_100%)] bg-[length:20px_20px]"></div>
               
               <div className="text-center relative z-10 w-full mt-4">
                 <div className="w-16 h-16 bg-[#0A1628] border-2 border-cyan-400 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-[0_0_20px_rgba(6,182,212,0.5)] p-2">
                   <img src="/axalogo.png" alt="Logo" className="w-full h-full object-contain filter drop-shadow-[0_0_5px_rgba(6,182,212,0.8)]" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
                   <ActivitySquare className="text-cyan-400 w-full h-full hidden" />
                 </div>
                 <h2 className="text-cyan-50 font-black tracking-[0.2em] text-lg drop-shadow-md">DEPT. RKG</h2>
                 <p className="text-cyan-400 text-[8px] tracking-[0.3em] font-bold uppercase mt-1">Sistem Absensi Mahasiswa</p>
                 <p className="text-cyan-600 text-[7px] tracking-[0.2em] uppercase mt-1">Kartu Akses Absen</p>
               </div>

               <div className="bg-white p-3 rounded-2xl relative z-10 shadow-[0_0_30px_rgba(6,182,212,0.6)] border-4 border-[#0A1628]">
                  <img src={`https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${selectedStudentForKTM.nim}&margin=0`} alt="QR Code" className="w-36 h-36" />
               </div>

               <div className="text-center relative z-10 w-full bg-[#0A1628]/80 p-5 rounded-[1.5rem] backdrop-blur-md border border-cyan-500/40 mb-2">
                 <h1 className="text-sm md:text-base font-black text-white uppercase leading-tight mb-2 truncate px-2 tracking-widest">{selectedStudentForKTM.name}</h1>
                 <div className="h-[2px] w-16 bg-cyan-500 mx-auto my-2 rounded-full shadow-[0_0_10px_rgba(6,182,212,0.8)]"></div>
                 <p className="text-cyan-300 font-mono text-lg tracking-[0.2em] font-bold mt-2">{selectedStudentForKTM.nim}</p>
                 <p className="text-[#050B14] bg-cyan-500 inline-block px-3 py-1 rounded-md text-[8px] font-black uppercase tracking-[0.2em] mt-3">
                    {clusters.find(c => c.id === selectedStudentForKTM.clusterId)?.name || 'BELUM ADA KELOMPOK'}
                 </p>
               </div>
            </div>

            <button onClick={() => window.print()} className="w-full mt-8 py-4 bg-transparent border-2 border-cyan-500 hover:bg-cyan-500/20 text-cyan-400 font-black tracking-widest uppercase text-xs rounded-2xl flex items-center justify-center gap-3 transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.3)] active:scale-95">
              <Printer className="w-4 h-4" /> Cetak Sekarang (Print)
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

const AdminReports: React.FC = () => {
  const { logs, sessions, clusters, deleteLog } = useAppContext();
  const [search, setSearch] = useState('');
  const [filterSession, setFilterSession] = useState('All');
  const [filterCluster, setFilterCluster] = useState('All');
  const { startObj, endObj, FilterUI } = useDateFilter();
  
  const [previewImage, setPreviewImage] = useState<string | null>(null);

  const filteredLogs = logs.filter(log => {
    const matchSearch = log.name.toLowerCase().includes(search.toLowerCase()) || log.nim.includes(search);
    const matchSession = filterSession === 'All' || log.sessionName === filterSession;
    const clusterName = clusters.find(c => c.id === filterCluster)?.name;
    const matchCluster = filterCluster === 'All' || log.clusterName === clusterName;
    
    // Add Date Filter from Dropdown
    const logDate = new Date(log.timestamp);
    const inDateRange = logDate >= startObj && logDate <= endObj;

    return matchSearch && matchSession && matchCluster && inDateRange;
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col relative w-full pb-10">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
           <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Riwayat Kehadiran</h2>
           <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1 uppercase">Data waktu, lokasi, dan foto absensi mahasiswa</p>
        </div>
        <button onClick={() => exportToExcel(filteredLogs)} className="w-full md:w-auto flex items-center justify-center gap-3 px-8 py-3.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white rounded-xl transition-all duration-300 font-black tracking-widest uppercase text-xs shadow-[0_0_20px_rgba(16,185,129,0.4)] active:scale-95">
           <Download className="w-4 h-4" /> Download Laporan (Excel)
        </button>
      </div>

      <div className="flex flex-col md:flex-row gap-3 md:gap-4 bg-[#0A1628]/60 p-4 rounded-2xl border border-cyan-500/20 shadow-lg items-end">
        <FilterUI />
        
        <div className="flex flex-col gap-1 flex-1 w-full md:w-auto">
           <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Pencarian Data</label>
           <div className="relative w-full">
              <SearchIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-cyan-600" />
              <input type="text" placeholder="Cari Nama atau NIM..." value={search} onChange={e=>setSearch(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl pl-11 pr-4 h-11 text-cyan-50 outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" />
           </div>
        </div>
        
        <div className="flex flex-col gap-1 w-full sm:w-auto">
           <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Filter Kelompok</label>
           <div className="flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl px-2 h-11 w-full sm:w-auto focus-within:border-cyan-400 transition-colors">
              <select value={filterCluster} onChange={e=>setFilterCluster(e.target.value)} className="bg-transparent text-cyan-50 text-xs font-bold uppercase outline-none cursor-pointer px-3 w-full sm:w-40 h-full">
                <option value="All">Semua Kelompok</option>
                {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
           </div>
        </div>

        <div className="flex flex-col gap-1 w-full sm:w-auto">
           <label className="text-[9px] text-cyan-500 uppercase tracking-widest font-bold">Jadwal Shift</label>
           <div className="flex items-center bg-[#050B14] border border-cyan-500/30 rounded-xl px-2 h-11 w-full sm:w-auto focus-within:border-cyan-400 transition-colors">
              <select value={filterSession} onChange={e=>setFilterSession(e.target.value)} className="bg-transparent text-cyan-50 text-xs font-bold uppercase outline-none cursor-pointer px-3 w-full sm:w-40 h-full">
                <option value="All">Semua Shift</option>
                {sessions.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
              </select>
           </div>
        </div>
      </div>

      <div className="flex-1 bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 rounded-[1.5rem] overflow-hidden flex flex-col shadow-xl">
        <div className="overflow-x-auto custom-scrollbar">
          <table className="w-full text-left border-collapse min-w-[1000px]">
            <thead>
              <tr className="bg-[#050B14]/80 border-b border-cyan-500/30 text-cyan-500 text-[10px] tracking-[0.2em] uppercase font-black">
                <th className="p-4 md:p-5">Foto Absen</th>
                <th className="p-4 md:p-5">Data Mahasiswa</th>
                <th className="p-4 md:p-5">Waktu Kehadiran</th>
                <th className="p-4 md:p-5">Jadwal Shift</th>
                <th className="p-4 md:p-5">Lokasi Absen</th>
                <th className="p-4 md:p-5 text-right">Opsi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-cyan-900/30">
              {filteredLogs.map(log => (
                <tr key={log.id} className="hover:bg-cyan-900/20 transition-colors duration-200">
                  <td className="p-4 md:p-5">
                    <div onClick={() => setPreviewImage(log.photoBase64)} className="w-16 h-16 rounded-xl overflow-hidden border-2 border-cyan-500/40 bg-black relative group cursor-pointer shadow-md hover:shadow-[0_0_15px_rgba(6,182,212,0.6)] hover:border-cyan-300 transition-all duration-300">
                      <img src={log.photoBase64} alt="Selfie" className="w-full h-full object-cover" />
                      <div className="absolute inset-0 bg-[#0A1628]/60 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity backdrop-blur-sm">
                        <Maximize className="w-5 h-5 text-cyan-400" />
                      </div>
                    </div>
                  </td>
                  <td className="p-4 md:p-5">
                     <p className="font-bold text-cyan-50 text-sm uppercase tracking-wide truncate max-w-[200px] mb-1">{log.name}</p>
                     <p className="text-xs text-cyan-400/80 font-mono tracking-widest">{log.nim}</p>
                     <p className="text-[9px] mt-2 inline-block px-2 py-0.5 bg-cyan-950 text-cyan-300 rounded border border-cyan-500/20 font-bold uppercase tracking-wider">{log.clusterName || 'Tanpa Kelompok'}</p>
                  </td>
                  <td className="p-4 md:p-5">
                     <p className="text-cyan-50 font-black font-mono text-base tracking-wider mb-1 drop-shadow-md">{new Date(log.timestamp).toLocaleTimeString('id-ID')}</p>
                     <p className="text-[10px] md:text-xs text-cyan-500/80 font-mono uppercase tracking-widest">{new Date(log.timestamp).toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'short' })}</p>
                  </td>
                  <td className="p-4 md:p-5">
                     <p className="text-cyan-200 text-xs font-bold uppercase tracking-widest mb-2">{log.sessionName}</p>
                     <span className={cn("px-3 py-1 text-[9px] font-black uppercase tracking-[0.2em] rounded-md border shadow-sm", log.status === 'Hadir' ? "bg-emerald-950/50 text-emerald-400 border-emerald-500/40 shadow-[0_0_10px_rgba(16,185,129,0.2)]" : "bg-amber-950/50 text-amber-400 border-amber-500/40 shadow-[0_0_10px_rgba(245,158,11,0.2)]")}>{log.status}</span>
                  </td>
                  <td className="p-4 md:p-5">
                    <a href={`https://www.google.com/maps?q=${log.location.lat},${log.location.lng}`} target="_blank" rel="noreferrer" className="inline-flex items-center gap-2 px-3 py-2 bg-cyan-950/50 hover:bg-cyan-600 hover:text-white text-cyan-400 text-[9px] font-black uppercase tracking-[0.2em] rounded-lg border border-cyan-500/40 transition-all duration-300 shadow-sm active:scale-95">
                      <MapPin className="w-3 h-3" /> Buka Peta
                    </a>
                    <p className="text-[9px] text-cyan-600/70 mt-2.5 font-mono uppercase tracking-widest bg-[#050B14] inline-block px-2 py-1 rounded-md border border-cyan-900/50">{log.location.lat.toFixed(5)}, {log.location.lng.toFixed(5)}</p>
                  </td>
                  <td className="p-4 md:p-5 text-right">
                    <button onClick={() => { if(confirm(`Yakin ingin menghapus riwayat kehadiran ${log.name}?`)) deleteLog(log.id); }} title="Hapus Riwayat" className="p-2.5 text-rose-500 hover:text-white hover:bg-rose-600 rounded-xl transition-all duration-300 border border-transparent hover:border-rose-500/50 hover:shadow-[0_0_15px_rgba(244,63,94,0.4)] active:scale-95">
                      <Trash2 className="w-5 h-5" />
                    </button>
                  </td>
                </tr>
              ))}
              {filteredLogs.length === 0 && <tr><td colSpan={6} className="p-16 text-center text-cyan-800 font-mono text-sm uppercase tracking-widest">Belum ada riwayat absensi.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
      
      {/* MODAL FULLSCREEN PREVIEW IMAGE RESPONSIVE */}
      {previewImage && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-[#050B14]/95 backdrop-blur-2xl p-4 animate-in fade-in zoom-in-95 duration-300" onClick={() => setPreviewImage(null)}>
          <div className="relative max-w-3xl w-full flex flex-col items-center justify-center">
            <button onClick={() => setPreviewImage(null)} className="absolute -top-14 md:-top-16 right-0 md:-right-8 p-3 bg-rose-950/50 hover:bg-rose-500 hover:text-white rounded-xl transition-all duration-300 text-rose-400 shadow-[0_0_15px_rgba(244,63,94,0.2)] active:scale-90 border border-rose-500/30">
              <X className="w-6 h-6"/>
            </button>
            <div className="relative w-full overflow-hidden rounded-[2rem] border-[4px] md:border-[8px] border-cyan-500/30 shadow-[0_0_80px_rgba(6,182,212,0.4)] bg-black">
                {/* HUD Overlay for fullscreen */}
                <div className="absolute inset-0 pointer-events-none bg-[linear-gradient(transparent_95%,rgba(6,182,212,0.2)_100%),linear-gradient(90deg,transparent_95%,rgba(6,182,212,0.2)_100%)] bg-[length:40px_40px] mix-blend-screen opacity-50"></div>
                <img src={previewImage} alt="Preview Foto Absen" className="max-w-full max-h-[75vh] md:max-h-[85vh] w-full object-contain mx-auto" onClick={e => e.stopPropagation()} />
            </div>
            <p className="mt-5 text-cyan-400 text-[10px] font-mono tracking-[0.2em] bg-[#0A1628] px-4 py-2 rounded-lg border border-cyan-500/20 uppercase">Ketuk area luar untuk menutup foto</p>
          </div>
        </div>
      )}
    </div>
  );
};

const AdminGeofence: React.FC = () => {
  const { geofence, updateGeofence } = useAppContext();
  const [lat, setLat] = useState(geofence.lat.toString());
  const [lng, setLng] = useState(geofence.lng.toString());
  const [radius, setRadius] = useState(geofence.radius.toString());
  const [locationName, setLocationName] = useState(geofence.name || 'Gedung Kampus Pusat');

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    updateGeofence({ lat: parseFloat(lat), lng: parseFloat(lng), radius: parseInt(radius), name: locationName });
    alert('Pengaturan lokasi absensi berhasil disimpan!');
  };

  const getMyLocation = () => {
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (pos) => { setLat(pos.coords.latitude.toString()); setLng(pos.coords.longitude.toString()); },
        () => alert('Gagal mendeteksi lokasi GPS Anda saat ini.')
      );
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 max-w-3xl pb-10">
      <div>
        <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Pengaturan Lokasi Absen</h2>
        <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1 uppercase">Tentukan batas area kampus atau tempat kerja</p>
      </div>

      <form onSubmit={handleSave} className="bg-[#0A1628]/80 backdrop-blur-md border border-cyan-500/30 p-6 md:p-8 rounded-[2rem] space-y-6 md:space-y-8 shadow-[0_15px_40px_rgba(0,0,0,0.5)]">
        <div className="p-5 bg-cyan-950/30 border border-cyan-500/30 rounded-2xl flex items-start gap-4 shadow-inner relative overflow-hidden">
          <div className="absolute left-0 top-0 w-1 h-full bg-cyan-500"></div>
          <Navigation className="w-7 h-7 text-cyan-400 mt-1 shrink-0 drop-shadow-[0_0_8px_rgba(6,182,212,0.8)]" />
          <p className="text-xs text-cyan-100/90 leading-relaxed font-mono uppercase tracking-wide">Mahasiswa hanya bisa melakukan absen jika lokasi GPS mereka berada dalam jangkauan jarak (<b>Batas Radius Maksimal</b>) dari koordinat lokasi yang Anda tentukan di bawah ini.</p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-5 md:gap-6">
          <div className="space-y-1.5 md:col-span-2">
             <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Nama Lokasi Absen</label>
             <input required type="text" value={locationName} onChange={e=>setLocationName(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm font-mono" placeholder="Contoh: Gedung Rektorat" />
          </div>
          <div className="space-y-1.5">
            <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Latitude</label>
            <input required type="number" step="any" value={lat} onChange={e=>setLat(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" />
          </div>
          <div className="space-y-1.5">
            <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Longitude</label>
            <input required type="number" step="any" value={lng} onChange={e=>setLng(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" />
          </div>
        </div>

        <div className="space-y-1.5">
           <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Batas Jarak Radius (Meter)</label>
           <input required type="number" min="10" value={radius} onChange={e=>setRadius(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-4 text-cyan-400 outline-none focus:border-cyan-400 transition-colors shadow-inner font-black text-lg md:text-xl text-center tracking-widest" />
        </div>

        <div className="flex flex-col md:flex-row gap-3 md:gap-4 pt-8 border-t border-cyan-900/50">
          <button type="button" onClick={getMyLocation} className="w-full md:w-auto px-6 py-4 bg-[#050B14] hover:bg-cyan-950/40 border border-cyan-500/40 text-cyan-400 font-black tracking-widest uppercase text-xs rounded-xl transition-all duration-300 flex items-center justify-center gap-3 active:scale-95 shadow-sm">
            <MapPin className="w-4 h-4" /> Gunakan Lokasi Saya Saat Ini
          </button>
          <button type="submit" className="w-full md:flex-1 py-4 bg-cyan-600 hover:bg-cyan-500 text-white font-black tracking-[0.15em] uppercase text-xs rounded-xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.4)] active:scale-95 border border-cyan-400/50">
            Simpan Lokasi
          </button>
        </div>
      </form>
    </div>
  );
};

const AdminSettings: React.FC = () => {
  const { sessions, updateSession, addSession, deleteSession } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [editingSessId, setEditingSessId] = useState<string | null>(null);
  
  const [formSess, setFormSess] = useState({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault(); 
    if (editingSessId) {
       updateSession(editingSessId, { ...formSess });
    } else {
       addSession({ ...formSess, isActive: true }); 
    }
    setIsAdding(false); setEditingSessId(null); 
    setFormSess({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });
  };

  const startEdit = (sess: Session) => {
     setIsAdding(true);
     setEditingSessId(sess.id);
     setFormSess({ name: sess.name, startTime: sess.startTime, endTime: sess.endTime, toleranceMinutes: sess.toleranceMinutes });
  }

  const cancelForm = () => {
     setIsAdding(false); setEditingSessId(null);
     setFormSess({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });
  }

  return (
    <div className="space-y-6 animate-in fade-in duration-500 pb-10">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
           <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Pengaturan Jadwal Shift</h2>
           <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1 uppercase">Kelola jadwal jam kehadiran mahasiswa</p>
        </div>
        <button onClick={() => {cancelForm(); setIsAdding(true);}} className="flex items-center gap-2 px-6 py-3 bg-cyan-600/20 text-cyan-400 hover:bg-cyan-500/30 border border-cyan-500/50 rounded-xl transition-all duration-300 font-black uppercase tracking-widest text-xs shadow-[0_0_15px_rgba(6,182,212,0.2)] w-full md:w-auto justify-center">
           <Plus className="w-4 h-4" /> Tambah Jadwal Baru
        </button>
      </div>

      {isAdding && (
        <form onSubmit={handleSave} className="bg-[#0A1628]/90 backdrop-blur-md border border-cyan-500/50 p-6 md:p-8 rounded-3xl grid grid-cols-1 md:grid-cols-5 gap-5 md:gap-6 items-end animate-in slide-in-from-top-4 shadow-[0_15px_40px_rgba(0,0,0,0.5)] relative">
          
          <div className="absolute top-4 right-4 cursor-pointer text-cyan-600 hover:text-cyan-400" onClick={cancelForm}><X className="w-5 h-5"/></div>

          <div className="space-y-1.5 md:col-span-2">
             <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Nama Shift</label>
             <input required type="text" value={formSess.name} onChange={e=>setFormSess({...formSess, name: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm" placeholder="Contoh: Shift Pagi / Kelas A" />
          </div>
          <div className="space-y-1.5">
             <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Jam Mulai</label>
             <input required type="time" value={formSess.startTime} onChange={e=>setFormSess({...formSess, startTime: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm" />
          </div>
          <div className="space-y-1.5">
             <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Jam Berakhir</label>
             <input required type="time" value={formSess.endTime} onChange={e=>setFormSess({...formSess, endTime: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm" />
          </div>
          <div className="space-y-1.5">
             <label className="text-[10px] text-cyan-500 font-bold uppercase tracking-widest ml-1">Toleransi Tutup Sesi (+ Dari Jam Berakhir)</label>
             <div className="relative">
                <input required type="number" min="0" value={formSess.toleranceMinutes} onChange={e=>setFormSess({...formSess, toleranceMinutes: parseInt(e.target.value)})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl pl-4 pr-12 py-3.5 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-sm" />
                <span className="absolute right-4 top-1/2 -translate-y-1/2 text-cyan-600 text-xs font-mono font-bold">MENIT</span>
             </div>
          </div>
          <button type="submit" className="md:col-span-5 w-full py-4 bg-cyan-600 hover:bg-cyan-500 text-white font-black uppercase tracking-widest text-xs rounded-xl transition-all duration-300 shadow-[0_0_15px_rgba(6,182,212,0.4)] active:scale-95 mt-2">
             {editingSessId ? 'Simpan Perubahan' : 'Buat Jadwal Shift'}
          </button>
        </form>
      )}

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 md:gap-6">
        {sessions.map(session => (
          <div key={session.id} className={cn("p-6 rounded-[2rem] border transition-all duration-300 hover:shadow-lg group relative overflow-hidden", session.isActive ? "bg-[#0A1628]/80 border-cyan-500/30 hover:border-cyan-400/60 backdrop-blur-md shadow-[0_10px_30px_rgba(0,0,0,0.5)]" : "bg-[#050B14]/80 opacity-60 border-cyan-900 hover:opacity-100")}>
            
            <div className="absolute top-0 right-0 w-24 h-24 bg-cyan-600/5 rounded-bl-full pointer-events-none"></div>

            <div className="flex justify-between items-start mb-6 relative z-10">
              <h3 className="text-xl font-black text-cyan-50 uppercase tracking-widest max-w-[60%]">{session.name}</h3>
              <div className="flex gap-2">
                <button onClick={() => updateSession(session.id, { isActive: !session.isActive })} className={cn("px-4 py-2 text-[9px] font-black uppercase tracking-widest rounded-lg border transition-all duration-300 shadow-sm active:scale-95 flex items-center gap-1", session.isActive ? "bg-emerald-950/50 text-emerald-400 border-emerald-500/40 shadow-[0_0_10px_rgba(16,185,129,0.2)]" : "bg-slate-900/80 text-slate-500 border-slate-700")}>
                   {session.isActive ? <><Activity className="w-3 h-3"/> Aktif</> : 'Nonaktif'}
                </button>
              </div>
            </div>
            
            <div className="space-y-4 text-xs font-mono bg-[#050B14] p-5 rounded-2xl border border-cyan-500/20 relative z-10 shadow-inner">
              <div className="flex justify-between items-center text-cyan-400/80">
                 <div className="flex items-center gap-3"><Clock className="w-4 h-4 text-cyan-500"/> Jam Tepat Waktu</div>
                 <span className="text-cyan-50 font-bold bg-[#0A1628] px-3 py-1.5 rounded-lg border border-cyan-500/20">{session.startTime} - {session.endTime}</span>
              </div>
              <div className="flex justify-between items-center text-cyan-400/80">
                 <div className="flex items-center gap-3"><ActivitySquare className="w-4 h-4 text-purple-500"/> Toleransi (Tutup Sesi)</div>
                 <span className="text-purple-300 font-bold bg-purple-950/40 px-3 py-1.5 rounded-lg border border-purple-500/30">+{session.toleranceMinutes} Menit</span>
              </div>
            </div>

            <div className="flex justify-end gap-2 mt-5 relative z-10">
               <button onClick={() => startEdit(session)} className="flex items-center gap-2 px-4 py-2 text-[10px] font-black uppercase tracking-widest bg-blue-950/50 text-blue-400 hover:bg-blue-600 hover:text-white border border-blue-500/30 rounded-xl transition-all duration-300 active:scale-95">
                 <Edit className="w-3.5 h-3.5" /> Edit
               </button>
               <button onClick={() => {if(confirm(`Yakin ingin menghapus jadwal ${session.name}?`)) deleteSession(session.id);}} className="flex items-center gap-2 px-4 py-2 text-[10px] font-black uppercase tracking-widest bg-rose-950/50 text-rose-400 hover:bg-rose-600 hover:text-white border border-rose-500/30 rounded-xl transition-all duration-300 active:scale-95">
                 <Trash2 className="w-3.5 h-3.5" /> Hapus
               </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// ==========================================
// ADMIN FORMATS (CRUD REALTIME TEMPLATE WA)
// ==========================================
const AdminFormats: React.FC = () => {
  const { formats, updateFormat } = useAppContext();
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editTemplate, setEditTemplate] = useState('');

  const handleEdit = (f: ChatFormat) => {
    setEditingId(f.id);
    setEditTemplate(f.template);
  };

  const handleSave = (id: number) => {
    updateFormat(id, editTemplate);
    setEditingId(null);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 pb-10">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Manajemen Format WA</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono uppercase mt-1">Ubah Template Pesan Bot WhatsApp secara Real-Time</p>
        </div>
      </div>

      <div className="bg-[#0A1628]/80 backdrop-blur-xl border border-cyan-500/20 p-5 rounded-2xl shadow-lg mb-6">
         <h3 className="text-xs font-black uppercase text-cyan-400 mb-3 tracking-widest flex items-center gap-2"><FileText className="w-4 h-4"/> Variabel Dinamis (Gunakan ini di dalam teks):</h3>
         <div className="flex flex-wrap gap-2 text-[10px] font-mono font-bold">
            {['[Nama Lengkap]', '[NIM]', '[Kelompok]', '[Shift]', '[Jam Sesi]', '[Jam Tutup]', '[Jam Absen]', '[Tanggal Mulai]', '[Tanggal Akhir]', '[Password]', '[Link]'].map(v => (
              <span key={v} className="bg-[#050B14] text-purple-300 px-2 py-1 rounded border border-purple-500/30">{v}</span>
            ))}
         </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        {formats.map((f) => (
          <div key={f.id} className="bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 p-5 rounded-2xl flex flex-col gap-4 group hover:border-cyan-500/50 transition-all shadow-lg relative overflow-hidden">
             <div className="flex justify-between items-start">
                <div>
                  <span className="bg-cyan-900 text-cyan-300 px-2 py-1 rounded-md text-[9px] font-black tracking-widest uppercase border border-cyan-500/40">ID Skenario: {f.id}</span>
                  <h3 className="text-base font-black text-white uppercase mt-3 tracking-wider">{f.title}</h3>
                  <p className="text-[10px] text-cyan-500/80 font-mono leading-relaxed mt-1">{f.description}</p>
                </div>
                {editingId !== f.id && (
                  <button onClick={() => handleEdit(f)} className="p-2 bg-blue-500/10 hover:bg-blue-500/20 text-blue-400 border border-blue-500/20 rounded-lg transition-all active:scale-95">
                    <Edit className="w-4 h-4" />
                  </button>
                )}
             </div>

             {editingId === f.id ? (
                <div className="flex flex-col gap-3 mt-2">
                  <textarea 
                    className="w-full bg-[#050B14] border border-cyan-500/50 rounded-xl p-3 text-cyan-50 outline-none focus:border-cyan-400 font-mono text-xs leading-relaxed min-h-[150px] custom-scrollbar" 
                    value={editTemplate} 
                    onChange={(e) => setEditTemplate(e.target.value)} 
                  />
                  <div className="flex gap-2 justify-end">
                    <button onClick={() => setEditingId(null)} className="px-4 py-2 bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 border border-rose-500/20 rounded-lg text-[10px] font-black uppercase tracking-wider transition-all">Batal</button>
                    <button onClick={() => handleSave(f.id)} className="px-4 py-2 bg-emerald-500/20 hover:bg-emerald-500/30 text-emerald-400 border border-emerald-500/30 rounded-lg text-[10px] font-black uppercase tracking-wider transition-all flex items-center gap-2"><CheckCircle2 className="w-3.5 h-3.5"/> Simpan Ke Cloud</button>
                  </div>
                </div>
             ) : (
                <div className="bg-[#050B14] border border-cyan-900/50 p-4 rounded-xl mt-2 relative">
                   <pre className="text-[10px] text-cyan-100 font-mono whitespace-pre-wrap leading-relaxed">{f.template}</pre>
                </div>
             )}
          </div>
        ))}
      </div>
    </div>
  );
};

const AdminManagement: React.FC = () => {
  const { admins, addAdmin, updateAdmin, deleteAdmin } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newAd, setNewAd] = useState({ username: '', password: '' });
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editUser, setEditUser] = useState('');
  const [editPass, setEditPass] = useState('');

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if(newAd.username.trim() && newAd.password.trim()) {
       addAdmin(newAd);
       setIsAdding(false); 
       setNewAd({ username: '', password: '' });
    }
  };

  const handleUpdate = (e: React.FormEvent) => {
     e.preventDefault();
     if(editingId && editUser.trim()) {
        updateAdmin(editingId, { username: editUser, password: editPass });
        setEditingId(null); 
        setEditUser('');
        setEditPass('');
     }
  }

  return (
    <div className="space-y-6 animate-in fade-in duration-500 pb-10">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl md:text-3xl font-black text-cyan-50 tracking-widest uppercase">Kelola Akun Admin</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono uppercase mt-1">Tambah atau atur akses masuk ke Dashboard</p>
        </div>
        <button onClick={() => setIsAdding(!isAdding)} className="flex items-center gap-2 px-5 py-3 bg-cyan-600/20 text-cyan-400 hover:bg-cyan-500/30 border border-cyan-500/50 rounded-xl transition-all duration-300 font-black uppercase tracking-widest text-xs shadow-[0_0_15px_rgba(6,182,212,0.2)]">
          <Plus className="w-4 h-4" /> Tambah Admin Baru
        </button>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-[#0A1628]/80 backdrop-blur-md border border-cyan-500/30 p-5 md:p-6 rounded-2xl flex flex-col md:flex-row gap-4 items-end shadow-xl animate-in slide-in-from-top-4">
          <div className="flex-1 space-y-1.5 w-full">
            <label className="text-[10px] md:text-xs text-cyan-500 font-bold uppercase tracking-widest ml-1">Username Admin Baru</label>
            <input required type="text" value={newAd.username} onChange={e=>setNewAd({...newAd, username: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors text-sm font-mono" placeholder="Ketik Username..." />
          </div>
          <div className="flex-1 space-y-1.5 w-full">
            <label className="text-[10px] md:text-xs text-cyan-500 font-bold uppercase tracking-widest ml-1">Password Baru</label>
            <input required type="text" value={newAd.password} onChange={e=>setNewAd({...newAd, password: e.target.value})} className="w-full bg-[#050B14] border border-cyan-500/30 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors text-sm font-mono" placeholder="Ketik Password..." />
          </div>
          <button type="submit" className="w-full md:w-auto px-8 py-3.5 bg-cyan-600 hover:bg-cyan-500 text-white font-black uppercase tracking-widest text-xs rounded-xl transition-all duration-300 shadow-lg active:scale-95">Simpan Admin</button>
        </form>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-5">
         {admins.map(a => (
            <div key={a.id} className="bg-[#0A1628]/60 backdrop-blur-md border border-cyan-500/20 p-5 rounded-2xl flex flex-col gap-4 group hover:border-cyan-500/50 transition-all duration-300 shadow-lg">
               {editingId === a.id ? (
                  <form onSubmit={handleUpdate} className="flex flex-col gap-3">
                     <div className="flex gap-2">
                         <User className="w-5 h-5 text-cyan-600" />
                         <input autoFocus required placeholder="Username" type="text" value={editUser} onChange={e=>setEditUser(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/50 rounded-lg px-3 py-2 text-white outline-none text-sm font-mono" />
                     </div>
                     <div className="flex gap-2">
                         <Key className="w-5 h-5 text-cyan-600" />
                         <input required placeholder="Password" type="text" value={editPass} onChange={e=>setEditPass(e.target.value)} className="w-full bg-[#050B14] border border-cyan-500/50 rounded-lg px-3 py-2 text-white outline-none text-sm font-mono" />
                     </div>
                     <div className="flex gap-2 justify-end mt-2">
                         <button type="submit" className="bg-emerald-500/20 text-emerald-400 p-2 rounded-lg border border-emerald-500/30 flex-1 flex justify-center"><CheckCircle2 className="w-4 h-4"/></button>
                         <button type="button" onClick={()=>setEditingId(null)} className="bg-rose-500/20 text-rose-400 p-2 rounded-lg border border-rose-500/30 flex-1 flex justify-center"><X className="w-4 h-4"/></button>
                     </div>
                  </form>
               ) : (
                  <>
                     <div className="flex items-center gap-3">
                        <div className="w-10 h-10 bg-cyan-950/50 rounded-xl flex items-center justify-center border border-cyan-500/30"><ShieldCheck className="w-5 h-5 text-cyan-400" /></div>
                        <div>
                           <h3 className="font-bold text-white text-base tracking-wide font-mono">{a.username}</h3>
                           <p className="text-[10px] text-cyan-500 tracking-widest uppercase mt-0.5">Admin Dashboard</p>
                        </div>
                     </div>
                     <div className="flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                        <button onClick={()=>{setEditingId(a.id); setEditUser(a.username); setEditPass(a.password || '');}} className="flex-1 flex justify-center items-center gap-2 py-2 bg-blue-500/10 hover:bg-blue-500/20 text-blue-400 border border-blue-500/20 rounded-lg text-[10px] uppercase font-bold tracking-wider"><Edit className="w-3.5 h-3.5"/> Edit</button>
                        <button onClick={()=>{if(confirm(`Yakin ingin menghapus Admin ${a.username}?`)) deleteAdmin(a.id);}} className="flex-1 flex justify-center items-center gap-2 py-2 bg-rose-500/10 hover:bg-rose-500/20 text-rose-400 border border-rose-500/20 rounded-lg text-[10px] uppercase font-bold tracking-wider"><Trash2 className="w-3.5 h-3.5"/> Hapus</button>
                     </div>
                  </>
               )}
            </div>
         ))}
         {admins.length === 0 && <div className="col-span-full p-8 text-center border-2 border-dashed border-cyan-900 rounded-2xl text-cyan-700 font-mono text-sm uppercase">Belum Ada Admin Terdaftar</div>}
      </div>
    </div>
  );
};


const SearchIcon = ({ className }: { className?: string }) => (
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></svg>
);

const AdminLayout: React.FC<{ children: React.ReactNode, activeRoute: string, setRoute: (r:string)=>void }> = ({ children, activeRoute, setRoute }) => {
  const { syncStatus } = useAppContext();
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const handleLogout = () => { localStorage.removeItem('axaxyz_admin_auth'); setRoute('admin-login'); };

  const navItems = [
    { id: 'admin-dashboard', icon: ActivitySquare, label: 'Dashboard Utama' },
    { id: 'admin-clusters', icon: Network, label: 'Data Kelompok' },
    { id: 'admin-students', icon: Database, label: 'Data Mahasiswa' },
    { id: 'admin-reports', icon: FileText, label: 'Riwayat Absensi' },
    { id: 'admin-geofence', icon: Map, label: 'Pengaturan Lokasi' },
    { id: 'admin-settings', icon: Calendar, label: 'Jadwal Absen' },
    { id: 'admin-formats', icon: MessageSquare, label: 'Manajemen Format' },
    { id: 'admin-management', icon: ShieldCheck, label: 'Kelola Admin' },
  ];

  return (
    <div className="min-h-screen bg-[#050B14] flex text-cyan-50 font-sans w-full overflow-hidden relative radiology-bg">
      
      {/* MOBILE MENU OVERLAY */}
      {isMobileMenuOpen && (
         <div className="fixed inset-0 bg-[#050B14]/90 z-40 md:hidden backdrop-blur-md animate-in fade-in duration-300" onClick={() => setIsMobileMenuOpen(false)}></div>
      )}

      {/* RESPONSIVE SIDEBAR */}
      <aside className={cn(
         "fixed inset-y-0 left-0 z-50 w-[280px] md:w-72 bg-[#0A1628]/95 border-r border-cyan-500/20 flex flex-col backdrop-blur-3xl transition-transform duration-500 ease-[cubic-bezier(0.16,1,0.3,1)] md:relative md:translate-x-0 shadow-[20px_0_50px_rgba(0,0,0,0.8)] md:shadow-none",
         isMobileMenuOpen ? "translate-x-0" : "-translate-x-full"
      )}>
        <div className="p-6 md:p-8 border-b border-cyan-900/50 flex items-center justify-between">
          <div className="flex items-center gap-4">
             <div className="w-12 h-12 bg-[#050B14] border border-cyan-500/50 rounded-xl flex items-center justify-center shadow-[0_0_15px_rgba(6,182,212,0.4)] overflow-hidden p-2">
                <img src="/axalogo.png" alt="Logo" className="w-full h-full object-contain filter drop-shadow-[0_0_5px_rgba(6,182,212,0.8)]" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
                <ActivitySquare className="text-cyan-400 w-full h-full hidden" />
             </div>
             <div className="flex flex-col">
                <span className="font-black text-lg md:text-xl tracking-[0.2em] text-cyan-50 uppercase drop-shadow-[0_0_10px_rgba(6,182,212,0.5)]">DEPT. RKG</span>
                <span className="text-[7px] md:text-[8px] text-cyan-400 font-mono tracking-[0.3em] uppercase mt-1">Admin Panel</span>
             </div>
          </div>
          <button className="md:hidden p-2 bg-cyan-950/50 border border-cyan-500/30 rounded-xl text-cyan-400 transition-colors" onClick={() => setIsMobileMenuOpen(false)}>
             <X className="w-5 h-5"/>
          </button>
        </div>
        
        <nav className="flex-1 p-4 md:p-5 space-y-3 overflow-y-auto custom-scrollbar">
          {navItems.map(item => (
            <button key={item.id} onClick={() => { setRoute(item.id); setIsMobileMenuOpen(false); }} className={cn("w-full flex items-center gap-4 px-5 py-4 rounded-xl transition-all duration-300 text-xs font-black tracking-widest uppercase active:scale-[0.98] border", activeRoute === item.id ? "bg-cyan-950/60 text-cyan-300 border-cyan-500/50 shadow-[inset_0_0_20px_rgba(6,182,212,0.2)] shadow-[0_0_15px_rgba(6,182,212,0.2)]" : "text-cyan-600/70 hover:bg-[#050B14] hover:text-cyan-400 border-transparent hover:border-cyan-900/50")}>
              <item.icon className={cn("w-5 h-5 transition-transform duration-300", activeRoute === item.id && "scale-110 drop-shadow-[0_0_8px_rgba(6,182,212,0.8)]")} /> {item.label}
            </button>
          ))}
        </nav>
        
        <div className="p-4 md:p-5 border-t border-cyan-900/50">
          <button onClick={handleLogout} className="w-full flex items-center justify-center gap-3 px-5 py-4 rounded-xl text-rose-400 bg-rose-950/30 hover:bg-rose-600 hover:text-white transition-all duration-300 text-[10px] font-black uppercase tracking-[0.2em] border border-rose-500/30 active:scale-95 shadow-sm hover:shadow-[0_0_15px_rgba(244,63,94,0.4)]">
             <LogOut className="w-4 h-4" /> Keluar (Logout)
          </button>
        </div>
      </aside>

      <main className="flex-1 flex flex-col relative overflow-y-auto w-full h-screen custom-scrollbar">
        {/* RESPONSIVE HEADER & STATUS BADGE */}
        <header className="sticky top-0 p-4 md:p-6 flex justify-between md:justify-end items-center z-30 w-full bg-[#0A1628]/80 backdrop-blur-xl border-b border-cyan-900/50 shadow-[0_4px_30px_rgba(0,0,0,0.5)]">
           <button className="md:hidden p-2.5 bg-[#050B14] border border-cyan-500/30 rounded-xl text-cyan-400 transition-colors active:scale-95 shadow-[0_0_10px_rgba(6,182,212,0.2)]" onClick={() => setIsMobileMenuOpen(true)}>
              <Menu className="w-5 h-5" />
           </button>
           
           <div className="flex items-center gap-2.5 px-4 md:px-5 py-2 md:py-2.5 bg-[#050B14] border border-cyan-500/30 rounded-xl text-[9px] md:text-[10px] font-bold shadow-[0_0_20px_rgba(0,0,0,0.5)] transition-all font-mono">
               {syncStatus === 'syncing' && <><RefreshCcw className="w-3.5 h-3.5 md:w-4 md:h-4 animate-spin text-cyan-400"/> <span className="text-cyan-400 tracking-[0.2em] uppercase">Menyimpan...</span></>}
               {syncStatus === 'synced' && <><Cloud className="w-3.5 h-3.5 md:w-4 md:h-4 text-emerald-400 drop-shadow-[0_0_5px_rgba(16,185,129,0.8)]"/> <span className="text-emerald-400 tracking-[0.2em] uppercase">Tersimpan Online</span></>}
               {syncStatus === 'error' && <><CloudOff className="w-3.5 h-3.5 md:w-4 md:h-4 text-rose-400"/> <span className="text-rose-400 tracking-[0.2em] uppercase">Gagal Simpan</span></>}
               {syncStatus === 'offline' && <><ServerCrash className="w-3.5 h-3.5 md:w-4 md:h-4 text-amber-500"/> <span className="text-amber-500 tracking-[0.2em] uppercase">Mode Offline</span></>}
           </div>
        </header>

        {/* Global Lighting Effects */}
        <div className="absolute top-[-10%] right-[-10%] w-[50%] h-[50%] bg-cyan-600/10 rounded-full blur-[120px] pointer-events-none"></div>
        <div className="absolute bottom-[-10%] left-[-10%] w-[50%] h-[50%] bg-blue-600/10 rounded-full blur-[120px] pointer-events-none"></div>
        
        <div className="p-4 md:p-8 max-w-7xl mx-auto relative z-10 flex-1 w-full">{children}</div>

        <footer className="text-center py-6 text-[10px] md:text-xs text-cyan-600/60 font-mono tracking-widest mt-auto relative z-50 w-full">
          <a href="/ourteam" className="hover:text-cyan-400 hover:drop-shadow-[0_0_8px_rgba(6,182,212,0.8)] transition-all duration-300 cursor-pointer">
            Copyright © 2026 DEPT. RKG RSIGM UMI— All Rights Reserved. Made with ❤️
          </a>
        </footer>
      </main>
      
      {/* GLOBAL SCROLLBAR STYLING */}
      <style>{`
        .custom-scrollbar::-webkit-scrollbar { width: 4px; height: 4px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #050B14; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: rgba(6, 182, 212, 0.3); border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: rgba(6, 182, 212, 0.6); }
      `}</style>
    </div>
  );
};

export default function App() {
  const [route, setRoute] = useState<string>('student');

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
      document.title = "Sistem Absensi Mahasiswa - DEPT. RKG";

      // SEO Google Site Verification (Gold Standard GSC)
      let metaGsc = document.querySelector("meta[name='google-site-verification']");
      if (!metaGsc) {
        metaGsc = document.createElement('meta');
        metaGsc.setAttribute('name', 'google-site-verification');
        metaGsc.setAttribute('content', 'AAKLVErwuFUspLpKD6XZwRxIZ5XqaTwy1BEK6-Rl0Ig');
        document.head.appendChild(metaGsc);
      }
    }
  }, []);

  useEffect(() => {
    const isAdminAuthed = localStorage.getItem('axaxyz_admin_auth') === 'true';
    if (route.startsWith('admin-') && route !== 'admin-login' && !isAdminAuthed) setRoute('admin-login');
  }, [route]);

  return (
    <AppProvider>
      <div className="fixed bottom-4 md:bottom-6 right-4 md:right-6 z-[999] flex gap-2 md:gap-3 bg-[#0A1628]/90 backdrop-blur-xl p-2.5 rounded-2xl border border-cyan-500/30 shadow-[0_20px_50px_rgba(0,0,0,0.8)]">
        <button onClick={() => setRoute('student')} className={cn("px-4 md:px-6 py-2.5 md:py-3.5 rounded-xl text-[9px] md:text-[10px] font-black uppercase tracking-[0.2em] transition-all duration-300 active:scale-95 shadow-sm border", route === 'student' ? "bg-cyan-600 text-white shadow-[0_0_15px_rgba(6,182,212,0.4)] border-cyan-400" : "bg-[#050B14] text-cyan-600 hover:bg-cyan-950/50 hover:text-cyan-400 border-transparent hover:border-cyan-900/50")}>Portal Mahasiswa</button>
        <button onClick={() => setRoute(typeof window !== 'undefined' && localStorage.getItem('axaxyz_admin_auth') === 'true' ? 'admin-dashboard' : 'admin-login')} className={cn("px-4 md:px-6 py-2.5 md:py-3.5 rounded-xl text-[9px] md:text-[10px] font-black uppercase tracking-[0.2em] transition-all duration-300 active:scale-95 shadow-sm border", route.startsWith('admin') ? "bg-blue-600 text-white shadow-[0_0_15px_rgba(37,99,235,0.4)] border-blue-400" : "bg-[#050B14] text-cyan-600 hover:bg-cyan-950/50 hover:text-cyan-400 border-transparent hover:border-cyan-900/50")}>Portal Admin</button>
      </div>

      {route === 'student' && <AttendanceWizard />}
      {route === 'admin-login' && <AdminLogin onLogin={() => setRoute('admin-dashboard')} />}
      
      {['admin-dashboard', 'admin-students', 'admin-clusters', 'admin-settings', 'admin-formats', 'admin-reports', 'admin-geofence', 'admin-management'].includes(route) && (
        <AdminLayout activeRoute={route} setRoute={setRoute}>
          {route === 'admin-dashboard' && <AdminDashboardHome />}
          {route === 'admin-students' && <AdminStudents />}
          {route === 'admin-clusters' && <AdminClusters />}
          {route === 'admin-geofence' && <AdminGeofence />}
          {route === 'admin-settings' && <AdminSettings />}
          {route === 'admin-formats' && <AdminFormats />}
          {route === 'admin-reports' && <AdminReports />}
          {route === 'admin-management' && <AdminManagement />}
        </AdminLayout>
      )}
    </AppProvider>
  );
}
