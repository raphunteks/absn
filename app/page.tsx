"use client";

import React, { useState, useEffect, createContext, useContext, useRef, useCallback, useMemo } from 'react';
import { 
  Camera, MapPin, Clock, QrCode, CheckCircle2, AlertCircle, 
  BarChart3, Settings, FileText, LogOut, Users, Download, Plus, Trash2,
  RefreshCcw, ChevronRight, Fingerprint, Map, Activity, Key, Upload, Database, Navigation,
  Printer, X, CreditCard, Eye, EyeOff, Lock, ShieldCheck, Loader2, User, CloudOff,
  Maximize, Menu, Edit, UserMinus, Calendar, Layers, ActivitySquare, ScanFace
} from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  PieChart, Pie, Cell, AreaChart, Area, ComposedChart, Legend, Line
} from 'recharts';

import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { format, parseISO, isWithinInterval, startOfDay, endOfDay } from 'date-fns';

// ==========================================
// UPSTASH REDIS CLOUD CLIENT (REST API POST)
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
      console.error(`Redis GET Error [${key}]:`, e);
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
      console.error(`Redis SET Error [${key}]:`, e);
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

// Fungsi pembantu untuk memuat modul XLSX secara dinamis melalui CDN untuk menghindari error kompilasi
const loadXLSX = async (): Promise<any> => {
  if (typeof window !== 'undefined' && (window as any).XLSX) return (window as any).XLSX;
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.onload = () => resolve((window as any).XLSX);
    script.onerror = reject;
    document.head.appendChild(script);
  });
};

// Poin 3: Export Data ke XLSX
const exportToXLSX = async (logs: Log[], students: Student[], clusters: Cluster[]) => {
  try {
    const XLSX = await loadXLSX();
    const data = logs.map(log => {
      const student = students.find(s => s.nim === log.nim);
      const cluster = clusters.find(c => c.id === student?.clusterId)?.name || 'Tanpa Cluster';
      const date = new Date(log.timestamp).toLocaleDateString('id-ID');
      const time = new Date(log.timestamp).toLocaleTimeString('id-ID');
      const mapsLink = `https://www.google.com/maps?q=${log.location.lat},${log.location.lng}`;
      
      return {
        'ID Log': log.id,
        'NIM': log.nim,
        'Nama': log.name,
        'Cluster / Kategori': cluster,
        'Tanggal': date,
        'Waktu': time,
        'Sesi': log.sessionName,
        'Status': log.status,
        'Latitude': log.location.lat,
        'Longitude': log.location.lng,
        'Link Maps': mapsLink
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Riwayat Absensi");
    XLSX.writeFile(workbook, `Rekap_Absensi_RKG_${new Date().toISOString().split('T')[0]}.xlsx`);
  } catch (err) {
    alert("Gagal memuat modul Excel. Pastikan koneksi internet stabil.");
  }
};

const redis = Redis.fromEnv();

const CloudStore = {
  isAvailable: () => { 
    return redis !== null && redis.url !== '' && redis.token !== ''; 
  },
  getCredentials: () => {
    return { url: redis.url, token: redis.token };
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

// Interfaces
interface Cluster { id: string; name: string; }
interface Session { id: string; name: string; startTime: string; endTime: string; toleranceMinutes: number; isActive: boolean; }
interface Log { id: string; nim: string; name: string; timestamp: string; sessionName: string; status: 'Hadir' | 'Terlambat'; location: { lat: number; lng: number }; photoBase64: string; deviceId: string; }
interface Student { id: string; nim: string; name: string; password?: string; deviceId?: string | null; clusterId?: string; }
interface Geofence { lat: number; lng: number; radius: number; name?: string; }

type SyncStatus = 'offline' | 'synced' | 'syncing' | 'error';

interface AppContextType {
  sessions: Session[];
  clusters: Cluster[];
  logs: Log[];
  students: Student[];
  geofence: Geofence;
  isCloudSync: boolean;
  syncStatus: SyncStatus;
  addLog: (log: Omit<Log, 'id' | 'timestamp'>) => void;
  deleteLog: (id: string) => void;
  updateSession: (id: string, updates: Partial<Session>) => void;
  addSession: (session: Omit<Session, 'id'>) => void;
  deleteSession: (id: string) => void;
  addCluster: (cluster: Omit<Cluster, 'id'>) => void;
  updateCluster: (id: string, updates: Partial<Cluster>) => void;
  deleteCluster: (id: string) => void;
  addStudent: (student: Omit<Student, 'id'>) => void;
  updateStudent: (id: string, updates: Partial<Student>) => void;
  bulkAddStudents: (newStudents: Omit<Student, 'id'>[]) => void;
  deleteStudent: (id: string) => void;
  unlinkMyDevice: (nim: string) => void;
  updateGeofence: (data: Geofence) => void;
  forceManualSync: () => Promise<void>;
}

const defaultSessions: Session[] = [
  { id: '1', name: 'Pagi', startTime: '07:00', endTime: '09:00', toleranceMinutes: 15, isActive: true },
  { id: '2', name: 'Siang', startTime: '12:00', endTime: '13:30', toleranceMinutes: 15, isActive: true },
];
const defaultClusters: Cluster[] = [
  { id: 'c1', name: 'Cluster I 2025' },
  { id: 'c2', name: 'Cluster II 2025' }
];
const defaultGeofence: Geofence = { lat: -6.200000, lng: 106.816666, radius: 500, name: 'Kampus Utama' };

const AppContext = createContext<AppContextType | null>(null);

const AppProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [isAppLoading, setIsAppLoading] = useState(true);
  const [isCloudSync, setIsCloudSync] = useState(false);
  const [syncStatus, setSyncStatus] = useState<SyncStatus>('offline');
  
  const [sessions, setSessions] = useState<Session[]>([]);
  const [clusters, setClusters] = useState<Cluster[]>([]);
  const [logs, setLogs] = useState<Log[]>([]);
  const [students, setStudents] = useState<Student[]>([]);
  const [geofence, setGeofence] = useState<Geofence>(defaultGeofence);

  useEffect(() => {
    const initData = async () => {
      const cloudAvailable = CloudStore.isAvailable();
      setIsCloudSync(cloudAvailable);
      setSyncStatus(cloudAvailable ? 'synced' : 'offline');

      let s = null, c = null, l = null, st = null, gf = null;

      if (cloudAvailable) {
        s = await CloudStore.get('axaxyz_sessions');
        c = await CloudStore.get('axaxyz_clusters');
        l = await CloudStore.get('axaxyz_logs');
        st = await CloudStore.get('axaxyz_students');
        gf = await CloudStore.get('axaxyz_geofence');
      }

      if (!s) s = JSON.parse(localStorage.getItem('axaxyz_sessions') || 'null');
      if (!c) c = JSON.parse(localStorage.getItem('axaxyz_clusters') || 'null');
      if (!l) l = JSON.parse(localStorage.getItem('axaxyz_logs') || 'null');
      if (!st) st = JSON.parse(localStorage.getItem('axaxyz_students') || 'null');
      if (!gf) gf = JSON.parse(localStorage.getItem('axaxyz_geofence') || 'null');

      setSessions(s || defaultSessions);
      setClusters(c || defaultClusters);
      setLogs(l || []);
      setStudents(st || []);
      setGeofence(gf || defaultGeofence);
      
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
      console.error(`Sync Engine Error [${key}]:`, e);
      setSyncStatus('error');
    }
  };

  const forceManualSync = async () => {
    if (!CloudStore.isAvailable()) {
      alert("❌ Sinkronisasi Gagal: Konfigurasi NEXT_PUBLIC_... Upstash tidak terbaca.");
      return;
    }
    setSyncStatus('syncing');
    try {
      await CloudStore.set('axaxyz_sessions', JSON.stringify(sessions));
      await CloudStore.set('axaxyz_clusters', JSON.stringify(clusters));
      await CloudStore.set('axaxyz_logs', JSON.stringify(logs));
      await CloudStore.set('axaxyz_students', JSON.stringify(students));
      await CloudStore.set('axaxyz_geofence', JSON.stringify(geofence));
      setSyncStatus('synced');
      alert("✅ Data berhasil disinkronisasi paksa ke Cloud Upstash.");
    } catch (e: any) {
      console.error(e);
      setSyncStatus('error');
      alert("❌ Error saat sinkronisasi: " + e.message);
    }
  };

  const saveSessions = (d: Session[]) => { setSessions(d); localStorage.setItem('axaxyz_sessions', JSON.stringify(d)); syncToCloud('axaxyz_sessions', d); };
  const saveClusters = (d: Cluster[]) => { setClusters(d); localStorage.setItem('axaxyz_clusters', JSON.stringify(d)); syncToCloud('axaxyz_clusters', d); };
  const saveLogs = (d: Log[]) => { setLogs(d); localStorage.setItem('axaxyz_logs', JSON.stringify(d)); syncToCloud('axaxyz_logs', d); };
  const saveStudents = (d: Student[]) => { setStudents(d); localStorage.setItem('axaxyz_students', JSON.stringify(d)); syncToCloud('axaxyz_students', d); };
  const saveGeofence = (d: Geofence) => { setGeofence(d); localStorage.setItem('axaxyz_geofence', JSON.stringify(d)); syncToCloud('axaxyz_geofence', d); };

  const addLog = (logData: Omit<Log, 'id' | 'timestamp'>) => saveLogs([{ ...logData, id: Math.random().toString(36).substr(2, 9), timestamp: new Date().toISOString() }, ...logs]);
  const deleteLog = (id: string) => saveLogs(logs.filter(l => l.id !== id));
  
  const updateSession = (id: string, updates: Partial<Session>) => saveSessions(sessions.map(s => s.id === id ? { ...s, ...updates } : s));
  const addSession = (sessionData: Omit<Session, 'id'>) => saveSessions([...sessions, { ...sessionData, id: Math.random().toString(36).substr(2, 9) }]);
  const deleteSession = (id: string) => saveSessions(sessions.filter(s => s.id !== id));
  
  const addCluster = (clusterData: Omit<Cluster, 'id'>) => saveClusters([...clusters, { ...clusterData, id: Math.random().toString(36).substr(2, 9) }]);
  const updateCluster = (id: string, updates: Partial<Cluster>) => saveClusters(clusters.map(c => c.id === id ? { ...c, ...updates } : c));
  const deleteCluster = (id: string) => saveClusters(clusters.filter(c => c.id !== id));

  const addStudent = (studentData: Omit<Student, 'id'>) => saveStudents([...students, { ...studentData, id: Math.random().toString(36).substr(2, 9) }]);
  const updateStudent = (id: string, updates: Partial<Student>) => saveStudents(students.map(s => s.id === id ? { ...s, ...updates } : s));
  const bulkAddStudents = (newStudents: Omit<Student, 'id'>[]) => {
    const formatted = newStudents.map(s => ({ ...s, id: Math.random().toString(36).substr(2, 9) }));
    saveStudents([...students, ...formatted]);
  };
  const deleteStudent = (id: string) => saveStudents(students.filter(s => s.id !== id));
  
  // Poin 5: Unlink device untuk mahasiswa
  const unlinkMyDevice = (nim: string) => {
    const student = students.find(s => s.nim === nim);
    if(student) {
      updateStudent(student.id, { deviceId: null });
    }
  };

  const updateGeofence = (data: Geofence) => saveGeofence(data);

  if (isAppLoading) {
    return (
      <div className="min-h-screen bg-slate-950 flex flex-col items-center justify-center p-4">
        <div className="relative mb-6">
          <div className="absolute inset-0 bg-cyan-500/20 blur-3xl rounded-full"></div>
          <div className="w-20 h-20 bg-gradient-to-br from-cyan-900 to-slate-800 border-2 border-cyan-500/50 rounded-3xl flex items-center justify-center shadow-[0_0_30px_rgba(6,182,212,0.5)] animate-[pulse_2s_ease-in-out_infinite] relative z-10 overflow-hidden p-3">
             <ScanFace className="w-10 h-10 text-cyan-400" />
          </div>
        </div>
        <Loader2 className="w-8 h-8 text-cyan-400 animate-spin mb-4" />
        <h2 className="text-xl font-bold text-white mb-2 tracking-widest">INITIALIZING X-RAY CORE...</h2>
        <p className="text-cyan-500/70 text-sm text-center max-w-xs font-mono">Connecting to Secure Cloud Database</p>
      </div>
    );
  }

  return (
    <AppContext.Provider value={{ 
      isCloudSync, syncStatus, sessions, clusters, logs, students, geofence, 
      addLog, deleteLog, updateSession, addSession, deleteSession, 
      addCluster, updateCluster, deleteCluster,
      addStudent, updateStudent, bulkAddStudents, deleteStudent, unlinkMyDevice,
      updateGeofence, forceManualSync 
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
// COMPONENT: STUDENT ATTENDANCE WIZARD
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
      if (currentMinutes >= startTotal && currentMinutes <= endTotal) {
        const isLate = currentMinutes > (startTotal + session.toleranceMinutes);
        return { session, status: isLate ? 'Terlambat' : 'Hadir' };
      }
    }
    return null;
  }, [currentTime, sessions]);

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-8 space-y-6 md:space-y-8 animate-in fade-in zoom-in duration-500 w-full">
      <div className="relative animate-[bounce_3s_ease-in-out_infinite]">
        <div className="absolute inset-0 bg-cyan-500/20 blur-3xl rounded-full"></div>
        <Clock className="w-20 h-20 md:w-24 md:h-24 text-cyan-400 relative z-10 drop-shadow-[0_0_15px_rgba(6,182,212,0.8)]" />
      </div>
      <div className="text-center space-y-2">
        <h2 className="text-5xl md:text-6xl font-black bg-clip-text text-transparent bg-gradient-to-r from-cyan-300 via-cyan-100 to-purple-300 tracking-tighter drop-shadow-sm font-mono">
          {format(currentTime, 'HH.mm.ss')}
        </h2>
        <p className="text-sm md:text-base text-cyan-500/70 font-medium tracking-widest uppercase">{currentTime.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
      </div>

      <div className="w-full max-w-md bg-slate-900/60 backdrop-blur-xl border border-cyan-500/20 p-5 md:p-6 rounded-3xl shadow-[0_0_40px_rgba(6,182,212,0.1)] transition-all duration-300">
        {activeSession ? (
          <div className="space-y-6">
            <div className="flex items-center justify-between p-4 bg-black/40 border border-cyan-500/10 rounded-2xl relative overflow-hidden">
              <div className="absolute top-0 left-0 w-1 h-full bg-cyan-500"></div>
              <div className="pl-2">
                <p className="text-white font-bold text-lg md:text-xl tracking-wide">{activeSession.session.name}</p>
                <p className="text-xs md:text-sm text-cyan-200/60 font-medium mt-0.5">{activeSession.session.startTime} - {activeSession.session.endTime}</p>
                <p className="text-[10px] md:text-xs text-slate-500 mt-1.5 font-mono">
                  Batas Waktu: {activeSession.session.startTime.split(':')[0]}:{String(parseInt(activeSession.session.startTime.split(':')[1]) + activeSession.session.toleranceMinutes).padStart(2, '0')}
                </p>
              </div>
              <span className={cn("px-3 py-1.5 text-xs font-black rounded-full shadow-lg border", activeSession.status === 'Hadir' ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/30 shadow-emerald-500/20" : "bg-orange-500/10 text-orange-400 border-orange-500/30 shadow-orange-500/20")}>
                {activeSession.status}
              </span>
            </div>
            <button onClick={() => onComplete({ sessionName: activeSession.session.name, status: activeSession.status as 'Hadir' | 'Terlambat' })} className="w-full py-4 bg-gradient-to-r from-cyan-600 to-cyan-800 hover:from-cyan-500 hover:to-cyan-700 text-white font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.4)] hover:shadow-[0_0_30px_rgba(6,182,212,0.6)] flex items-center justify-center gap-2 active:scale-[0.98] border border-cyan-400/30">
              Inisiasi Sesi <ChevronRight className="w-5 h-5" />
            </button>
          </div>
        ) : (
          <div className="space-y-4 text-center">
            <div className="p-6 bg-rose-950/40 border border-rose-500/20 rounded-2xl text-rose-400 flex flex-col items-center gap-3">
              <AlertCircle className="w-10 h-10 animate-pulse drop-shadow-[0_0_10px_rgba(244,63,94,0.8)]" />
              <div>
                 <p className="font-bold text-lg tracking-wider">OFFLINE</p>
                 <p className="text-sm mt-1 opacity-80">Tidak ada sesi aktif untuk saat ini.</p>
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
      setStatus('error'); setErrorMsg('Sensor Geolocation tidak didukung sistem.'); return;
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
          setErrorMsg(`Target berada di luar zona radiasi aman (${geofence.name || 'Pusat'}).`);
        }
      },
      (error) => { setStatus('error'); setErrorMsg('Gagal mendeteksi sinyal GPS. Pastikan izin akses lokasi aktif.'); },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  }, [onComplete, geofence]);

  useEffect(() => { checkLocation(); }, [checkLocation]);

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-8 space-y-6 max-w-md mx-auto animate-in slide-in-from-right duration-500 w-full">
      <div className="w-32 h-32 bg-slate-900/50 rounded-full flex items-center justify-center border-2 border-cyan-500/20 relative overflow-hidden shadow-[0_0_40px_rgba(6,182,212,0.15)]">
        {status === 'loading' && (
           <>
             <div className="absolute inset-0 border-[4px] border-cyan-500/10 border-t-cyan-400 rounded-full animate-spin"></div>
             <div className="absolute inset-4 border-[2px] border-purple-500/10 border-b-purple-400 rounded-full animate-spin-reverse"></div>
             <div className="absolute inset-0 bg-cyan-500/5 rounded-full animate-ping opacity-50"></div>
           </>
        )}
        <Navigation className={cn("w-12 h-12 relative z-10 transition-colors duration-500", status === 'error' ? 'text-rose-400 drop-shadow-[0_0_10px_rgba(244,63,94,0.8)]' : 'text-cyan-400 drop-shadow-[0_0_10px_rgba(6,182,212,0.8)]')} />
      </div>

      <div className="text-center space-y-2 w-full bg-slate-900/60 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-3xl shadow-2xl transition-all duration-300">
        <h3 className="text-2xl font-black text-white tracking-widest uppercase">Geo-Scan</h3>
        {status === 'loading' && <p className="text-cyan-200/60 text-sm font-mono animate-pulse">Memindai koordinat satelit...</p>}
        
        {status === 'success' && (
          <div className="text-cyan-400 space-y-3 animate-in fade-in zoom-in mt-4">
            <CheckCircle2 className="w-14 h-14 mx-auto drop-shadow-[0_0_15px_rgba(6,182,212,0.8)]" />
            <div>
               <p className="font-bold text-xl tracking-widest">AKSES DIBERIKAN</p>
               <p className="text-sm text-cyan-200/60 font-mono mt-1">Deviasi: {Math.round(distance || 0)}m dari inti.</p>
            </div>
          </div>
        )}
        
        {status === 'error' && (
          <div className="space-y-5 animate-in fade-in zoom-in mt-4">
            <div className="p-4 bg-rose-950/40 border border-rose-500/20 rounded-2xl text-center">
              <p className="text-sm font-bold text-rose-400 leading-relaxed">{errorMsg}</p>
              {distance && <p className="text-[10px] text-rose-300 mt-2 font-mono bg-rose-900/30 inline-block px-3 py-1 rounded-full border border-rose-500/20">Deviasi saat ini: {Math.round(distance)}m (Maks: {geofence.radius}m)</p>}
            </div>
            <button onClick={checkLocation} className="w-full py-4 bg-slate-800/50 hover:bg-slate-800 text-white rounded-2xl text-sm font-bold tracking-widest uppercase flex items-center justify-center gap-2 transition-all duration-300 border border-cyan-500/20 hover:border-cyan-400/50 active:scale-[0.98]">
              <RefreshCcw className="w-5 h-5" /> Pindai Ulang
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

const QRScanner: React.FC<{ onComplete: (data: {nim: string, name: string, deviceId: string}) => void }> = ({ onComplete }) => {
  const { students, updateStudent } = useAppContext();
  const [nimInput, setNimInput] = useState('');
  const [passInput, setPassInput] = useState('');
  const [error, setError] = useState('');
  const [isScanning, setIsScanning] = useState(false);
  
  const qrScannerRef = useRef<any>(null);

  const handleVerify = (scannedNim?: string) => {
    setError('');
    const targetNim = scannedNim || nimInput;
    if (!targetNim) { setError('Otorisasi ditolak. Masukkan atau Scan NIM.'); return; }

    let studentName = 'Subjek Tidak Dikenal';
    let finalDeviceId = localStorage.getItem('axaxyz_device_id');
    
    if (!finalDeviceId) {
      finalDeviceId = Math.random().toString(36).substring(2, 15);
      localStorage.setItem('axaxyz_device_id', finalDeviceId);
    }
    
    if (students.length > 0) {
      if (!passInput && !scannedNim) { setError('Otorisasi ditolak. Masukkan Password (PIN).'); return; }
      const foundStudent = students.find(s => s.nim === targetNim);
      if (!foundStudent) {
        setError('Data NIM tidak ditemukan di server.'); return;
      }
      
      if (!scannedNim && foundStudent.password !== passInput) {
        setError('Kredensial tidak valid.'); return;
      }
      studentName = foundStudent.name;

      if (foundStudent.deviceId && foundStudent.deviceId !== finalDeviceId) {
        setError('⚠️ FRAUD ALERT: Kredensial telah digunakan di perangkat/instrumen lain. Hubungi Admin.');
        return;
      }
      
      if (!foundStudent.deviceId) {
        updateStudent(foundStudent.id, { deviceId: finalDeviceId });
      }
    } else {
      studentName = 'Subjek Mode Bypass'; 
      let deviceOwner = localStorage.getItem('axaxyz_device_owner');
      if (!deviceOwner) {
        localStorage.setItem('axaxyz_device_owner', targetNim); 
      } else if (deviceOwner !== targetNim) {
        setError('⚠️ FRAUD ALERT: Perangkat terkunci untuk NIM lain.');
        return;
      }
    }

    // Poin 5: Pastikan axaxyz_device_owner diset untuk fitur Unlink Device
    localStorage.setItem('axaxyz_device_owner', targetNim);

    onComplete({ nim: targetNim, name: studentName, deviceId: finalDeviceId });
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
             setError('Optik gagal dimuat. Periksa izin akses instrumen kamera.');
             setIsScanning(false);
          });
        } catch (err) {
          setError('Kesalahan fatal pada modul optik.');
          setIsScanning(false);
        }
      }, 100);
    } catch (error) {
      setError('Modul QR Scanner corrupt.');
      setIsScanning(false);
    }
  };

  const stopScanner = () => {
    if (qrScannerRef.current) {
      qrScannerRef.current.stop().catch(() => {});
    }
    setIsScanning(false);
  };

  return (
    <div className="flex flex-col items-center justify-center p-4 md:p-8 space-y-6 max-w-md mx-auto animate-in slide-in-from-right duration-500 w-full">
      <div className="w-full bg-slate-900/60 backdrop-blur-xl border border-cyan-500/20 p-6 md:p-8 rounded-[2rem] shadow-[0_0_40px_rgba(6,182,212,0.1)] transition-all duration-300">
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-slate-800 border border-cyan-500/50 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-[0_0_20px_rgba(6,182,212,0.4)]">
            {isScanning ? <ScanFace className="w-8 h-8 text-cyan-400 animate-pulse" /> : <QrCode className="w-8 h-8 text-cyan-400 drop-shadow-[0_0_5px_rgba(6,182,212,0.8)]" />}
          </div>
          <h3 className="text-2xl font-black text-white mb-2 tracking-widest uppercase">Identifikasi</h3>
          <p className="text-cyan-500/70 text-xs font-mono">Pindai KTM / Input Kredensial Manual</p>
        </div>

        {isScanning ? (
          <div className="space-y-4 animate-in fade-in zoom-in">
             <div className="relative w-full rounded-2xl overflow-hidden border border-cyan-500/50 bg-black aspect-square shadow-[0_0_30px_rgba(6,182,212,0.3)]">
                <div id="qr-reader-box" className="w-full h-full opacity-80 mix-blend-screen"></div>
                <div className="absolute inset-0 pointer-events-none">
                   <div className="w-full h-[15%] bg-gradient-to-b from-transparent via-cyan-500/30 to-transparent animate-[scan_2s_ease-in-out_infinite] border-y border-cyan-400/50"></div>
                </div>
                {/* Crosshair Overlay */}
                <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
                    <div className="w-48 h-48 border-2 border-cyan-500/30 rounded-3xl"></div>
                </div>
             </div>
             <p className="text-[10px] text-center text-cyan-400 font-mono tracking-widest uppercase animate-pulse">Menyelaraskan Optik...</p>
             <button onClick={stopScanner} className="w-full py-3.5 bg-rose-950/40 hover:bg-rose-900/60 border border-rose-500/30 text-rose-400 rounded-xl text-xs font-bold tracking-widest uppercase transition-colors duration-300 active:scale-[0.98]">Batalkan Pemindaian</button>
          </div>
        ) : (
          <div className="space-y-5">
            <button onClick={startScanner} className="w-full py-4 bg-cyan-950/40 hover:bg-cyan-900/60 border border-cyan-500/50 text-cyan-400 font-bold tracking-widest uppercase rounded-2xl flex justify-center items-center gap-2 transition-all duration-300 active:scale-[0.98] shadow-[inset_0_0_20px_rgba(6,182,212,0.1)]">
              <Camera className="w-5 h-5" /> Aktifkan Optik KTM
            </button>
            
            <div className="relative flex items-center py-2 opacity-60">
               <div className="flex-grow border-t border-cyan-500/30"></div>
               <span className="flex-shrink-0 mx-4 text-cyan-500/50 text-[10px] font-mono tracking-widest uppercase">Atau Override Manual</span>
               <div className="flex-grow border-t border-cyan-500/30"></div>
            </div>

            <div className="space-y-2">
              <label className="text-[10px] text-cyan-200/60 font-bold uppercase tracking-widest ml-1 font-mono">Nomor Induk Mahasiswa</label>
              <div className="flex items-center bg-black/40 border border-cyan-500/20 rounded-2xl overflow-hidden focus-within:border-cyan-400 focus-within:ring-1 focus-within:ring-cyan-400 transition-all duration-300 shadow-inner">
                <div className="pl-4 pr-2 text-cyan-500/50"><Fingerprint className="w-5 h-5"/></div>
                <input type="text" placeholder="Masukkan NIM..." className="w-full bg-transparent py-3.5 pr-4 text-white font-mono outline-none placeholder-slate-700 text-sm" value={nimInput} onChange={(e) => setNimInput(e.target.value)} />
              </div>
            </div>
            
            {students.length > 0 && (
              <div className="space-y-2">
                <label className="text-[10px] text-cyan-200/60 font-bold uppercase tracking-widest ml-1 font-mono">Kode Akses (PIN)</label>
                <div className="flex items-center bg-black/40 border border-cyan-500/20 rounded-2xl overflow-hidden focus-within:border-cyan-400 focus-within:ring-1 focus-within:ring-cyan-400 transition-all duration-300 shadow-inner">
                  <div className="pl-4 pr-2 text-cyan-500/50"><Key className="w-5 h-5"/></div>
                  <input type="password" placeholder="••••••••" className="w-full bg-transparent py-3.5 pr-4 text-white font-mono outline-none placeholder-slate-700 text-sm tracking-widest" value={passInput} onChange={(e) => setPassInput(e.target.value)} />
                </div>
              </div>
            )}
            
            {error && (
              <div className="p-4 bg-rose-950/40 border border-rose-500/30 rounded-2xl flex items-start gap-3 animate-in shake">
                <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
                <p className="text-[11px] font-mono text-rose-300 leading-relaxed uppercase tracking-wider">{error}</p>
              </div>
            )}

            <button onClick={() => handleVerify()} className="w-full py-4 mt-2 bg-slate-800 hover:bg-slate-700 border border-cyan-500/30 text-cyan-300 font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 shadow-[0_0_15px_rgba(6,182,212,0.1)] hover:shadow-[0_0_25px_rgba(6,182,212,0.3)] active:scale-[0.98]">
              Verifikasi Data
            </button>
          </div>
        )}
      </div>
      <style>{`
        @keyframes scan {
          0%, 100% { transform: translateY(-100%); }
          50% { transform: translateY(600%); }
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
    <div className="flex flex-col items-center justify-center p-4 md:p-6 space-y-6 w-full max-w-md mx-auto animate-in slide-in-from-right duration-500">
      <div className="text-center">
        <h3 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Visual Scan</h3>
        <p className="text-cyan-500/70 text-xs font-mono mt-1.5">Posisikan struktur wajah pada grid pandang.</p>
      </div>

      <div className="w-full bg-slate-950 rounded-[2rem] overflow-hidden border border-cyan-500/50 p-2 relative shadow-[0_0_40px_rgba(6,182,212,0.2)] flex items-center justify-center transition-all duration-500 group aspect-[3/4] md:aspect-video">
        <div className="w-full h-full rounded-[1.5rem] overflow-hidden relative bg-black">
          {!image ? (
            <video
              ref={videoRef}
              autoPlay
              playsInline
              muted
              className="w-full h-full object-cover transform scale-x-[-1] opacity-90 sepia-[.3] hue-rotate-[180deg] contrast-125 saturate-50"
            />
          ) : (
            <img src={image} alt="Selfie" className="w-full h-full object-cover sepia-[.3] hue-rotate-[180deg] contrast-125 saturate-50" />
          )}
          
          {/* X-Ray / Medical Overlay Effects */}
          <div className="absolute inset-0 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/scan-lines-light.png')] opacity-30 mix-blend-overlay"></div>
          
          {!image && (
            <div className="absolute inset-0 pointer-events-none flex items-center justify-center">
               <div className="absolute inset-0 bg-cyan-900/10"></div>
               {/* Center Focus Target */}
               <div className="relative w-48 h-64 md:w-56 md:h-72 border border-cyan-400/30 rounded-[3rem] shadow-[inset_0_0_30px_rgba(6,182,212,0.2)] flex flex-col items-center justify-between py-4">
                  <div className="w-full flex justify-between px-4">
                    <div className="w-4 h-4 border-t-2 border-l-2 border-cyan-300"></div>
                    <div className="w-4 h-4 border-t-2 border-r-2 border-cyan-300"></div>
                  </div>
                  <div className="w-16 h-[1px] bg-cyan-400/50"></div>
                  <div className="w-full flex justify-between px-4">
                    <div className="w-4 h-4 border-b-2 border-l-2 border-cyan-300"></div>
                    <div className="w-4 h-4 border-b-2 border-r-2 border-cyan-300"></div>
                  </div>
               </div>
            </div>
          )}
        </div>
      </div>

      <div className="w-full">
        {!image ? (
          <button onClick={capture} className="w-full py-4 bg-cyan-950/40 hover:bg-cyan-900/60 border border-cyan-500/50 text-cyan-400 font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 flex items-center justify-center gap-3 active:scale-[0.98] shadow-[0_0_20px_rgba(6,182,212,0.1)]">
            <div className="w-6 h-6 rounded-full border-[2px] border-cyan-400 flex items-center justify-center">
              <div className="w-2 h-2 rounded-full bg-cyan-300 animate-pulse"></div>
            </div>
            Ekstrak Biometrik
          </button>
        ) : (
          <div className="flex gap-4">
            <button onClick={() => { setImage(null); startCamera(); }} className="flex-1 py-4 bg-slate-800/50 hover:bg-slate-800 border border-cyan-500/20 text-slate-300 font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 active:scale-[0.95] text-xs">Ulangi</button>
            <button onClick={() => onComplete(image)} className="flex-[2] py-4 bg-gradient-to-r from-cyan-600 to-cyan-800 hover:from-cyan-500 hover:to-cyan-700 border border-cyan-400/50 text-white font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 shadow-[0_0_30px_rgba(6,182,212,0.4)] flex items-center justify-center gap-2 active:scale-[0.95] text-xs">
              <CheckCircle2 className="w-4 h-4" /> Autentikasi
            </button>
          </div>
        )}
      </div>
    </div>
  );
};

const SuccessScreen: React.FC<{ reset: () => void }> = ({ reset }) => (
  <div className="flex flex-col items-center justify-center p-8 space-y-8 text-center animate-in zoom-in duration-500 w-full">
    <div className="relative">
      <div className="absolute inset-0 bg-cyan-500/20 blur-3xl rounded-full"></div>
      <div className="w-32 h-32 bg-cyan-950/50 border-2 border-cyan-500/50 rounded-full flex items-center justify-center relative z-10 animate-bounce shadow-[0_0_40px_rgba(6,182,212,0.4)]">
         <ShieldCheck className="w-16 h-16 text-cyan-300 drop-shadow-[0_0_10px_rgba(6,182,212,0.8)]" />
      </div>
    </div>
    <div className="space-y-3">
      <h2 className="text-3xl md:text-4xl font-black text-white tracking-widest uppercase">Data Terekam</h2>
      <p className="text-cyan-500/70 font-mono text-xs md:text-sm max-w-sm mx-auto leading-relaxed">Matriks kehadiran dan struktur biometrik berhasil dienkripsi ke dalam database pusat.</p>
    </div>
    <button onClick={reset} className="px-10 py-4 bg-slate-800/50 hover:bg-slate-800 border border-cyan-500/30 text-cyan-300 font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 mt-8 shadow-xl active:scale-95 text-xs">Terminasi Sesi</button>
  </div>
);

const AttendanceWizard: React.FC = () => {
  const { addLog, unlinkMyDevice } = useAppContext();
  const [step, setStep] = useState(1);
  const [data, setData] = useState<Partial<Log>>({});
  
  // Poin 5: Mendapatkan NIM dari localStorage untuk fitur Logout Device
  const [deviceOwnerNIM, setDeviceOwnerNIM] = useState<string | null>(null);

  useEffect(() => {
     if(typeof window !== 'undefined') {
        setDeviceOwnerNIM(localStorage.getItem('axaxyz_device_owner'));
     }
  }, [step]); // re-check on step change

  const reset = () => { setStep(1); setData({}); };
  const steps = ['Waktu', 'Lokasi', 'Identitas', 'Verifikasi'];

  const handleDeviceLogout = () => {
     if(!deviceOwnerNIM) return;
     if(confirm(`PERHATIAN!\n\nAnda akan menghapus tautan instrumen ini dari NIM: ${deviceOwnerNIM}.\n\nLanjutkan Unlink Device?`)) {
        unlinkMyDevice(deviceOwnerNIM);
        localStorage.removeItem('axaxyz_device_owner');
        localStorage.removeItem('axaxyz_device_id');
        setDeviceOwnerNIM(null);
        alert('Tautan instrumen berhasil diputus. Sistem siap untuk subjek baru.');
        window.location.reload();
     }
  };

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col font-sans text-slate-100 overflow-hidden relative selection:bg-cyan-500/30">
      {/* Background futuristik medis radiologi */}
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_top,_var(--tw-gradient-stops))] from-slate-900 via-slate-950 to-slate-950"></div>
      <div className="absolute inset-0 bg-[url('https://www.transparenttextures.com/patterns/carbon-fibre.png')] opacity-5 mix-blend-overlay"></div>
      <div className="absolute top-[-10%] left-[-10%] w-[60%] h-[60%] bg-cyan-900/20 rounded-full blur-[150px] pointer-events-none"></div>
      <div className="absolute bottom-[-10%] right-[-10%] w-[60%] h-[60%] bg-purple-900/10 rounded-full blur-[150px] pointer-events-none"></div>
      
      <header className="w-full p-4 md:p-6 flex justify-between items-center relative z-10 border-b border-cyan-500/10 bg-slate-950/80 backdrop-blur-xl shadow-lg">
        <div className="flex items-center gap-3 md:gap-4">
          <div className="w-10 h-10 md:w-12 md:h-12 bg-slate-900 border border-cyan-500/30 rounded-xl flex items-center justify-center shadow-[0_0_15px_rgba(6,182,212,0.3)] overflow-hidden p-1.5 md:p-2 transition-transform duration-300 hover:scale-105 cursor-pointer">
             <img src="/axalogo.png" alt="ABSENSI DEPT. RKG" className="w-full h-full object-contain drop-shadow-md brightness-150 grayscale contrast-125" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
             <ActivitySquare className="w-6 h-6 text-cyan-400 hidden" />
          </div>
          <div className="flex flex-col">
             <span className="font-black text-lg md:text-2xl tracking-[0.2em] text-cyan-50 drop-shadow-[0_0_5px_rgba(6,182,212,0.5)]">DEPT. RKG</span>
             <span className="text-[8px] md:text-[10px] text-cyan-500/80 font-mono tracking-widest uppercase">Radiology Core System</span>
          </div>
        </div>
        
        {/* Poin 5: Tombol Logout Device MHS */}
        <div className="flex flex-col md:flex-row items-end md:items-center gap-2 md:gap-4">
           {deviceOwnerNIM && (
              <div className="flex items-center gap-2 bg-slate-900 border border-cyan-500/20 px-3 py-1.5 rounded-lg shadow-inner">
                 <span className="text-[9px] md:text-[10px] text-cyan-500/70 font-mono tracking-widest uppercase">ID: {deviceOwnerNIM}</span>
                 <button onClick={handleDeviceLogout} title="Unlink Instrumen" className="text-rose-400/70 hover:text-rose-400 transition-colors ml-1 p-1 bg-rose-950/30 rounded hover:bg-rose-900/50">
                    <UserMinus className="w-3.5 h-3.5" />
                 </button>
              </div>
           )}
           <div className="text-[9px] md:text-[10px] font-bold px-4 py-2 bg-cyan-950/50 border border-cyan-500/30 rounded-full text-cyan-300 tracking-[0.2em] shadow-[0_0_10px_rgba(6,182,212,0.2)] uppercase">PORTAL MHS</div>
        </div>
      </header>

      <main className="flex-1 flex flex-col relative z-10 w-full max-w-5xl mx-auto px-4 py-6 md:py-12 overflow-y-auto">
        {step < 5 && (
          <div className="mb-8 md:mb-16 max-w-2xl mx-auto w-full px-2">
            <div className="flex justify-between relative">
              <div className="absolute top-1/2 -translate-y-1/2 left-0 w-full h-1 bg-slate-800/80 rounded-full shadow-inner"></div>
              <div className="absolute top-1/2 -translate-y-1/2 left-0 h-1 bg-gradient-to-r from-cyan-600 to-cyan-300 rounded-full transition-all duration-700 ease-in-out shadow-[0_0_15px_rgba(6,182,212,0.8)]" style={{ width: `${((step - 1) / (steps.length - 1)) * 100}%` }}></div>
              {steps.map((label, idx) => {
                const isActive = step === idx + 1; const isPassed = step > idx + 1;
                return (
                  <div key={label} className="relative z-10 flex flex-col items-center gap-2 md:gap-3">
                    <div className={cn("w-8 h-8 md:w-10 md:h-10 rounded-full flex items-center justify-center text-xs md:text-sm font-black border-[2px] transition-all duration-500 font-mono", isActive ? "bg-slate-950 border-cyan-400 text-cyan-300 shadow-[0_0_20px_rgba(6,182,212,0.6)] scale-110" : isPassed ? "bg-cyan-900 border-cyan-600 text-cyan-100" : "bg-slate-900 border-slate-700 text-slate-600")}>
                      {isPassed ? <CheckCircle2 className="w-4 h-4 md:w-5 md:h-5" /> : `0${idx + 1}`}
                    </div>
                    <span className={cn("text-[9px] md:text-[10px] font-bold absolute -bottom-6 md:-bottom-7 w-max tracking-widest uppercase font-mono", isActive ? "text-cyan-300 drop-shadow-[0_0_5px_rgba(6,182,212,0.8)]" : isPassed ? "text-cyan-600" : "text-slate-600")}>{label}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        <div className="flex-1 flex items-center justify-center w-full">
          {step === 1 && <TimeCheck onComplete={(d) => { setData(prev => ({...prev, ...d})); setStep(2); }} />}
          {step === 2 && <LocationCheck onComplete={(d) => { setData(prev => ({...prev, location: d})); setStep(3); }} />}
          {step === 3 && <QRScanner onComplete={(d) => { setData(prev => ({...prev, ...d})); setStep(4); }} />}
          {step === 4 && <SelfieCapture onComplete={(photo) => { addLog({ ...data, photoBase64: photo } as Omit<Log, 'id' | 'timestamp'>); setStep(5); }} />}
          {step === 5 && <SuccessScreen reset={reset} />}
        </div>
      </main>
    </div>
  );
};

// ==========================================
// COMPONENT: ADMIN LOGIN
// ==========================================

const AdminLogin: React.FC<{ onLogin: () => void }> = ({ onLogin }) => {
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
      setErr(`Sistem terkunci. Terminasi: ${lockoutTimer} detik.`);
      return;
    }
    setIsLoading(true); setErr('');
    await new Promise(resolve => setTimeout(resolve, 800));

    let ADMIN_USER = 'admin';
    let ADMIN_PASS = 'admin123';

    if (typeof process !== 'undefined' && process.env) {
      if (process.env.NEXT_PUBLIC_ADMIN_USER) ADMIN_USER = process.env.NEXT_PUBLIC_ADMIN_USER;
      if (process.env.NEXT_PUBLIC_ADMIN_PASS) ADMIN_PASS = process.env.NEXT_PUBLIC_ADMIN_PASS;
    }

    if (user === ADMIN_USER && pass === ADMIN_PASS) {
      setAttempts(0);
      localStorage.setItem('axaxyz_admin_auth', 'true');
      onLogin();
    } else {
      const newAttempts = attempts + 1;
      setAttempts(newAttempts);
      if (newAttempts >= 3) {
        setLockoutTimer(30); 
        setErr('❌ BLOKIR: Batas upaya akses terlampaui (30s).');
      } else {
        setErr(`❌ Kredensial tidak valid. (Sisa upaya: ${3 - newAttempts})`);
      }
    }
    setIsLoading(false);
  };

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col items-center justify-center p-4 relative overflow-hidden w-full selection:bg-cyan-500/30">
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_center,_var(--tw-gradient-stops))] from-slate-900 via-slate-950 to-slate-950"></div>
      <div className="absolute inset-0 bg-[url('https://www.transparenttextures.com/patterns/carbon-fibre.png')] opacity-5 mix-blend-overlay"></div>
      <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[300px] h-[300px] md:w-[600px] md:h-[600px] bg-cyan-900/10 rounded-full blur-[100px] pointer-events-none border border-cyan-500/5"></div>

      <div className="w-full max-w-md bg-slate-900/60 backdrop-blur-3xl border border-cyan-500/20 p-6 md:p-10 rounded-[2rem] shadow-[0_0_50px_rgba(6,182,212,0.1)] relative z-10 animate-in slide-in-from-bottom-8 fade-in duration-700">
        <div className="flex flex-col items-center mb-8">
          <div className="relative mb-5">
            <div className="w-20 h-20 md:w-24 md:h-24 bg-slate-950 border border-cyan-500/50 rounded-[1.5rem] flex items-center justify-center shadow-[0_0_30px_rgba(6,182,212,0.3)] p-3 overflow-hidden">
              <img src="/axalogo.png" alt="ABSENSI DEPT. RKG" className="w-full h-full object-contain drop-shadow-md brightness-150 grayscale contrast-125" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
              <ShieldCheck className="w-10 h-10 md:w-12 md:h-12 text-cyan-400 hidden" />
            </div>
            <div className="absolute -bottom-2 -right-2 w-10 h-10 bg-slate-950 rounded-full border-[2px] border-cyan-500/50 flex items-center justify-center shadow-lg">
              <Lock className="w-4 h-4 text-cyan-400" />
            </div>
          </div>
          <h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Admin Core</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm mt-1.5 font-mono tracking-widest uppercase opacity-80">Secured Access Protocol</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-5">
          {err && (
            <div className="p-4 bg-rose-950/40 border border-rose-500/30 text-rose-400 text-xs font-mono tracking-wider rounded-2xl flex items-start gap-3 animate-in shake duration-300">
              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
              <p className="leading-tight uppercase">{err}</p>
            </div>
          )}

          <div className="space-y-2">
            <label className="text-[10px] text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">ID Pengguna</label>
            <div className="relative flex items-center bg-black/40 border border-cyan-500/20 rounded-2xl overflow-hidden focus-within:border-cyan-400 focus-within:ring-1 focus-within:ring-cyan-400 transition-all duration-300 shadow-inner">
              <div className="pl-4 pr-3 text-cyan-500/50"><User className="w-5 h-5"/></div>
              <input type="text" value={user} onChange={e=>setUser(e.target.value)} disabled={lockoutTimer > 0 || isLoading} className="w-full bg-transparent py-4 pr-4 text-white font-mono outline-none placeholder-slate-700 disabled:opacity-50 text-sm" placeholder="Input ID Admin" required />
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-[10px] text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Kunci Akses</label>
            <div className="relative flex items-center bg-black/40 border border-cyan-500/20 rounded-2xl overflow-hidden focus-within:border-cyan-400 focus-within:ring-1 focus-within:ring-cyan-400 transition-all duration-300 shadow-inner">
              <div className="pl-4 pr-3 text-cyan-500/50"><Key className="w-5 h-5"/></div>
              <input type={showPass ? 'text' : 'password'} value={pass} onChange={e=>setPass(e.target.value)} disabled={lockoutTimer > 0 || isLoading} className="w-full bg-transparent py-4 pr-12 text-white font-mono outline-none placeholder-slate-700 disabled:opacity-50 text-sm tracking-widest" placeholder="••••••••" required />
              <button type="button" onClick={() => setShowPass(!showPass)} className="absolute right-4 text-cyan-500/50 hover:text-cyan-400 transition-colors">
                {showPass ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
              </button>
            </div>
          </div>

          <button type="submit" disabled={lockoutTimer > 0 || isLoading} className="w-full py-4 mt-6 bg-cyan-950/40 hover:bg-cyan-900/60 border border-cyan-500/50 text-cyan-300 disabled:border-slate-700 disabled:text-slate-600 disabled:cursor-not-allowed font-bold tracking-widest uppercase rounded-2xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.1)] hover:shadow-[0_0_30px_rgba(6,182,212,0.3)] flex justify-center items-center gap-2 group active:scale-95 text-xs md:text-sm">
            {isLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : (
               <>Otorisasi <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" /></>
            )}
          </button>
        </form>
      </div>
    </div>
  );
};

// ==========================================
// COMPONENT: ADMIN DASHBOARD (Poin 2 & Poin 4)
// ==========================================
const AdminDashboardHome: React.FC = () => {
  const { logs, students, clusters } = useAppContext();
  
  // Poin 2: Filter Range Tanggal dan Cluster
  const [startDate, setStartDate] = useState(new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString().split('T')[0]); // Default awal bulan ini
  const [endDate, setEndDate] = useState(new Date().toISOString().split('T')[0]);
  const [selectedCluster, setSelectedCluster] = useState('All');

  const parsedStart = startOfDay(parseISO(startDate));
  const parsedEnd = endOfDay(parseISO(endDate));

  // Filter logs berdasarkan range tanggal dan cluster
  const filteredLogs = logs.filter(log => {
    const logDate = new Date(log.timestamp);
    const inDateRange = isWithinInterval(logDate, { start: parsedStart, end: parsedEnd });
    
    const student = students.find(s => s.nim === log.nim);
    const inCluster = selectedCluster === 'All' || student?.clusterId === selectedCluster;
    
    return inDateRange && inCluster;
  });

  // Kalkulasi Status Hari Ini (Berdasarkan filter Cluster)
  const todayStr = new Date().toISOString().split('T')[0];
  const todayLogs = filteredLogs.filter(l => l.timestamp.startsWith(todayStr));
  
  const targetStudents = selectedCluster === 'All' ? students : students.filter(s => s.clusterId === selectedCluster);
  const totalStudents = targetStudents.length;
  
  const uniqueAttendeesToday = new Set(todayLogs.map(l => l.nim)).size;
  const onTimeToday = todayLogs.filter(l => l.status === 'Hadir').length;
  const lateToday = todayLogs.filter(l => l.status === 'Terlambat').length;
  
  // Poin 2: "Jumlah Tidak Absen Di Hari Ini"
  const notAttendedToday = Math.max(0, totalStudents - uniqueAttendeesToday);

  // Data for Area Chart (Trend over time)
  const logsByDate = filteredLogs.reduce((acc, log) => {
    const date = log.timestamp.split('T')[0];
    if(!acc[date]) acc[date] = { date, Hadir: 0, Terlambat: 0 };
    acc[date][log.status]++;
    return acc;
  }, {} as Record<string, any>);
  const areaData = Object.values(logsByDate).sort((a: any, b: any) => a.date.localeCompare(b.date));

  // Data for Composed Chart (Cluster Comparison)
  const clusterStats = clusters.map(cluster => {
     const clusterStudents = students.filter(s => s.clusterId === cluster.id).map(s => s.nim);
     const cLogs = filteredLogs.filter(l => clusterStudents.includes(l.nim));
     return {
        name: cluster.name,
        Hadir: cLogs.filter(l => l.status === 'Hadir').length,
        Terlambat: cLogs.filter(l => l.status === 'Terlambat').length,
        Total: cLogs.length
     };
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      
      {/* HEADER & FILTERS */}
      <div className="flex flex-col lg:flex-row justify-between items-start lg:items-end gap-4 bg-slate-900/40 p-5 rounded-[2rem] border border-cyan-500/10">
        <div>
          <h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Data Analitik</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1">Pemantauan matriks kehadiran terpadu.</p>
        </div>
        
        <div className="flex flex-col sm:flex-row gap-3 w-full lg:w-auto">
           <div className="flex items-center bg-black/40 border border-cyan-500/20 rounded-xl overflow-hidden px-3 py-2 shadow-inner">
              <Calendar className="w-4 h-4 text-cyan-500/50 mr-2"/>
              <input type="date" value={startDate} onChange={e => setStartDate(e.target.value)} className="bg-transparent text-xs text-white font-mono outline-none" />
              <span className="text-cyan-500/50 mx-2">-</span>
              <input type="date" value={endDate} onChange={e => setEndDate(e.target.value)} className="bg-transparent text-xs text-white font-mono outline-none" />
           </div>
           
           <div className="flex items-center bg-black/40 border border-cyan-500/20 rounded-xl overflow-hidden px-3 py-2 shadow-inner">
              <Layers className="w-4 h-4 text-cyan-500/50 mr-2"/>
              <select value={selectedCluster} onChange={e => setSelectedCluster(e.target.value)} className="bg-transparent text-xs text-white font-mono outline-none w-full min-w-[120px] appearance-none cursor-pointer">
                 <option value="All">Semua Cluster</option>
                 {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
              </select>
           </div>
        </div>
      </div>

      {/* METRICS CARDS */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
        {[
          { title: 'Subjek Terdaftar', val: totalStudents, icon: Users, color: 'text-blue-400', border: 'border-blue-500/30' },
          { title: 'Hadir (Hari Ini)', val: onTimeToday, icon: CheckCircle2, color: 'text-emerald-400', border: 'border-emerald-500/30' },
          { title: 'Terlambat (Hari Ini)', val: lateToday, icon: Clock, color: 'text-amber-400', border: 'border-amber-500/30' },
          { title: 'Belum Absen (Hari Ini)', val: notAttendedToday, icon: UserMinus, color: 'text-rose-400', border: 'border-rose-500/30' } // Poin 2: Ganti Upstash Redis dgn Total Tidak Absen
        ].map((stat, i) => (
          <div key={i} className={cn("bg-slate-900/60 border p-5 rounded-[1.5rem] flex items-center justify-between transition-all duration-300 hover:bg-slate-800/80 shadow-[0_0_15px_rgba(0,0,0,0.2)]", stat.border)}>
            <div>
               <p className="text-slate-400 text-[10px] md:text-xs font-mono uppercase tracking-widest mb-1.5">{stat.title}</p>
               <h3 className="text-2xl md:text-3xl font-black text-white font-mono">{stat.val}</h3>
            </div>
            <div className="w-12 h-12 rounded-xl flex items-center justify-center bg-black/40 shadow-inner border border-white/5"><stat.icon className={cn("w-6 h-6", stat.color)} /></div>
          </div>
        ))}
      </div>

      {/* Poin 4: Cloud Diagnostic UI (GBR 2) telah dihilangkan sepenuhnya dari blok ini agar rapi. */}

      {/* ADVANCED CHARTS (Poin 2) */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 min-h-[350px] md:min-h-[400px]">
        {/* Trend Area Chart */}
        <div className="bg-slate-900/60 border border-cyan-500/20 p-6 rounded-[2rem] flex flex-col shadow-[0_0_30px_rgba(6,182,212,0.05)] relative overflow-hidden">
          <div className="absolute top-0 right-0 p-4 opacity-10 pointer-events-none"><ActivitySquare className="w-32 h-32 text-cyan-500"/></div>
          <h3 className="text-sm font-bold text-cyan-100 mb-6 tracking-widest uppercase font-mono border-b border-cyan-500/20 pb-4">Tren Kehadiran (Rentang Waktu)</h3>
          <div className="flex-1 w-full min-h-[250px] z-10">
            {areaData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <AreaChart data={areaData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                  <defs>
                    <linearGradient id="colorHadir" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#10b981" stopOpacity={0.8}/>
                      <stop offset="95%" stopColor="#10b981" stopOpacity={0}/>
                    </linearGradient>
                    <linearGradient id="colorTelat" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="5%" stopColor="#f59e0b" stopOpacity={0.8}/>
                      <stop offset="95%" stopColor="#f59e0b" stopOpacity={0}/>
                    </linearGradient>
                  </defs>
                  <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                  <XAxis dataKey="date" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} tickFormatter={(val) => val.substring(5)} />
                  <YAxis stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                  <Tooltip cursor={{stroke: '#334155', strokeWidth: 1}} contentStyle={{backgroundColor: '#020617', borderColor: '#06b6d4', color: '#f8fafc', borderRadius: '1rem', fontSize: '12px', fontFamily: 'monospace'}} />
                  <Legend iconType="circle" wrapperStyle={{fontSize: '11px', fontFamily: 'monospace', paddingTop: '10px'}}/>
                  <Area type="monotone" dataKey="Hadir" stroke="#10b981" fillOpacity={1} fill="url(#colorHadir)" />
                  <Area type="monotone" dataKey="Terlambat" stroke="#f59e0b" fillOpacity={1} fill="url(#colorTelat)" />
                </AreaChart>
              </ResponsiveContainer>
            ) : <div className="h-full flex items-center justify-center text-slate-500 font-mono text-xs uppercase tracking-widest">Data Tidak Tersedia</div>}
          </div>
        </div>
        
        {/* Cluster Comparison Composed Chart */}
        <div className="bg-slate-900/60 border border-cyan-500/20 p-6 rounded-[2rem] flex flex-col shadow-[0_0_30px_rgba(6,182,212,0.05)] relative overflow-hidden">
          <div className="absolute top-0 right-0 p-4 opacity-10 pointer-events-none"><Layers className="w-32 h-32 text-purple-500"/></div>
          <h3 className="text-sm font-bold text-cyan-100 mb-6 tracking-widest uppercase font-mono border-b border-cyan-500/20 pb-4">Komparasi Per Cluster</h3>
          <div className="flex-1 w-full min-h-[250px] z-10">
             {clusterStats.length > 0 ? (
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={clusterStats} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#1e293b" vertical={false} />
                    <XAxis dataKey="name" stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                    <YAxis stroke="#64748b" fontSize={10} tickLine={false} axisLine={false} />
                    <Tooltip cursor={{fill: '#1e293b', opacity: 0.4}} contentStyle={{backgroundColor: '#020617', borderColor: '#06b6d4', color: '#f8fafc', borderRadius: '1rem', fontSize: '12px', fontFamily: 'monospace'}} />
                    <Legend iconType="circle" wrapperStyle={{fontSize: '11px', fontFamily: 'monospace', paddingTop: '10px'}}/>
                    <Bar dataKey="Hadir" stackId="a" fill="#0ea5e9" radius={[0, 0, 4, 4]} barSize={30} />
                    <Bar dataKey="Terlambat" stackId="a" fill="#8b5cf6" radius={[4, 4, 0, 0]} />
                    <Line type="monotone" dataKey="Total" stroke="#f43f5e" strokeWidth={2} dot={{ r: 4, fill: '#f43f5e' }} />
                  </ComposedChart>
                </ResponsiveContainer>
             ) : <div className="h-full flex items-center justify-center text-slate-500 font-mono text-xs uppercase tracking-widest">Data Tidak Tersedia</div>}
          </div>
        </div>
      </div>
    </div>
  );
};

// ==========================================
// COMPONENT: ADMIN STUDENTS (Poin 1, Kategori Cluster & XLSX)
// ==========================================
const AdminStudents: React.FC = () => {
  const { students, clusters, addStudent, updateStudent, bulkAddStudents, deleteStudent } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newS, setNewS] = useState({ name: '', nim: '', password: '', clusterId: '' });
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  
  const [search, setSearch] = useState('');
  const [filterCluster, setFilterCluster] = useState('All');
  
  const [selectedStudentForKTM, setSelectedStudentForKTM] = useState<Student | null>(null);

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    addStudent({ ...newS, password: newS.password || `${newS.nim}123`, clusterId: newS.clusterId || undefined });
    setIsAdding(false);
    setNewS({ name: '', nim: '', password: '', clusterId: '' });
  };

  const handleUpdate = (e: React.FormEvent) => {
    e.preventDefault();
    if(editingStudent) {
       updateStudent(editingStudent.id, { 
          name: editingStudent.name, 
          nim: editingStudent.nim, 
          password: editingStudent.password,
          clusterId: editingStudent.clusterId || undefined
       });
       setEditingStudent(null);
    }
  };

  const handleUnlinkDevice = (id: string, name: string) => {
     if(confirm(`Yakin memutus tautan instrumen untuk subjek ${name}?`)) {
        updateStudent(id, { deviceId: null });
     }
  };

  // Poin 1 & 2: Bulk Upload Format Excel (.XLSX)
  const handleBulkUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    try {
      const XLSX: any = await loadXLSX();
      const reader = new FileReader();
      
      reader.onload = (event) => {
        try {
          const data = event.target?.result;
          const workbook = XLSX.read(data, { type: 'binary' });
          const sheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[sheetName];
          const json: any[] = XLSX.utils.sheet_to_json(worksheet);
          
          const newSt: Omit<Student, 'id'>[] = [];
          json.forEach(row => {
            // Asumsi penamaan header pada excel
            const name = row['Nama'] || row['NAMA'] || row['name'];
            const nim = String(row['NIM'] || row['nim']);
            const clusterName = row['Cluster'] || row['CLUSTER'] || row['Kategori'];
            
            let clusterId = undefined;
            if (clusterName) {
              const foundCluster = clusters.find(c => c.name.toLowerCase() === String(clusterName).toLowerCase());
              if (foundCluster) clusterId = foundCluster.id;
            }

            if (name && nim) {
               newSt.push({ name: String(name), nim, password: `${nim}123`, clusterId });
            }
          });
          
          if (newSt.length > 0) {
            bulkAddStudents(newSt);
            alert(`Berhasil menginjeksi ${newSt.length} data subjek ke database.`);
          } else {
            alert('Gagal mendeteksi kolom yang valid (Pastikan ada header "Nama" dan "NIM" dalam Excel Anda).');
          }
        } catch (err) {
          console.error(err);
          alert('Format file korup atau tidak didukung.');
        }
      };
      reader.readAsBinaryString(file);
    } catch (err) {
      console.error(err);
      alert('Gagal memuat parser XLSX. Pastikan koneksi internet berjalan.');
    } finally {
      e.target.value = ''; // Reset input agar bisa memuat file yang sama
    }
  };

  const filtered = students.filter(s => {
     const matchSearch = s.name.toLowerCase().includes(search.toLowerCase()) || s.nim.includes(search);
     const matchCluster = filterCluster === 'All' || s.clusterId === filterCluster;
     return matchSearch && matchCluster;
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col w-full relative">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-slate-900/40 p-5 rounded-[2rem] border border-cyan-500/10">
        <div>
          <h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Data Subjek</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1">Registrasi dan Manajemen Identitas Biometrik KTM</p>
        </div>
        <div className="flex flex-col sm:flex-row gap-2 md:gap-3 w-full md:w-auto">
           <label className="flex flex-1 md:flex-none justify-center items-center gap-2 px-5 py-3 bg-purple-900/40 text-purple-300 hover:bg-purple-800/60 border border-purple-500/30 rounded-xl transition-all duration-300 font-bold cursor-pointer active:scale-95 shadow-sm uppercase tracking-widest text-xs">
              <Upload className="w-4 h-4" /> Inject Data (XLSX)
              <input type="file" accept=".xlsx, .xls, .csv" className="hidden" onChange={handleBulkUpload} />
           </label>
           <button onClick={() => setIsAdding(!isAdding)} className="flex flex-1 md:flex-none justify-center items-center gap-2 px-5 py-3 bg-cyan-900/40 text-cyan-300 hover:bg-cyan-800/60 border border-cyan-500/30 rounded-xl transition-all duration-300 font-bold active:scale-95 shadow-sm uppercase tracking-widest text-xs">
             <Plus className="w-4 h-4" /> Registrasi Manual
           </button>
        </div>
      </div>

      {/* FILTERING AREA */}
      <div className="flex flex-col md:flex-row gap-3 md:gap-4">
        <div className="relative flex-1">
           <SearchIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-cyan-500/50" />
           <input type="text" placeholder="Cari Identitas / NIM..." value={search} onChange={e=>setSearch(e.target.value)} className="w-full bg-slate-900/60 border border-cyan-500/20 rounded-xl pl-10 pr-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm placeholder-slate-600" />
        </div>
        <select value={filterCluster} onChange={e=>setFilterCluster(e.target.value)} className="bg-slate-900/60 border border-cyan-500/20 rounded-xl px-5 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner w-full md:w-56 cursor-pointer appearance-none text-sm">
          <option value="All">Semua Cluster</option>
          {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-slate-900/80 backdrop-blur-md border border-cyan-500/30 p-5 md:p-6 rounded-[1.5rem] grid grid-cols-1 md:grid-cols-5 gap-4 items-end animate-in slide-in-from-top-4 shadow-[0_0_30px_rgba(6,182,212,0.1)]">
          <div className="space-y-1.5 md:col-span-2">
            <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Nama Lengkap</label>
            <input required type="text" value={newS.name} onChange={e=>setNewS({...newS, name: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-400 font-mono text-sm" placeholder="Input Nama..." />
          </div>
          <div className="space-y-1.5">
            <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">NIM</label>
            <input required type="text" value={newS.nim} onChange={e=>setNewS({...newS, nim: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-400 font-mono text-sm" placeholder="Input NIM..." />
          </div>
          <div className="space-y-1.5">
            <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Cluster</label>
            <select value={newS.clusterId} onChange={e=>setNewS({...newS, clusterId: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-400 font-mono text-sm appearance-none cursor-pointer">
               <option value="">-- Kosong --</option>
               {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
            </select>
          </div>
          <button type="submit" className="w-full py-3 bg-cyan-950/50 hover:bg-cyan-900 border border-cyan-500/50 text-cyan-300 font-bold rounded-xl transition-all duration-300 shadow-lg active:scale-95 uppercase tracking-widest text-xs">Simpan Data</button>
        </form>
      )}

      <div className="flex-1 bg-slate-900/60 backdrop-blur-md border border-cyan-500/20 rounded-[1.5rem] overflow-hidden flex flex-col shadow-[0_0_20px_rgba(0,0,0,0.5)] relative">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[900px]">
            <thead>
              <tr className="bg-black/40 border-b border-cyan-500/20 text-cyan-500/70 text-xs tracking-widest uppercase font-mono">
                <th className="p-4 md:p-5 whitespace-nowrap font-bold">NIM</th>
                <th className="p-4 md:p-5 whitespace-nowrap font-bold">Nama Subjek</th>
                <th className="p-4 md:p-5 whitespace-nowrap font-bold">Cluster</th>
                <th className="p-4 md:p-5 whitespace-nowrap font-bold">Kredensial</th>
                <th className="p-4 md:p-5 text-center whitespace-nowrap font-bold">Status Instrumen</th>
                <th className="p-4 md:p-5 text-right whitespace-nowrap font-bold">Manajemen</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-cyan-500/10">
              {filtered.map(st => {
                const clusterName = clusters.find(c => c.id === st.clusterId)?.name || '-';
                return (
                <tr key={st.id} className="hover:bg-cyan-950/30 transition-colors duration-200 text-slate-200 group">
                  <td className="p-4 md:p-5 font-mono text-sm text-cyan-100">{st.nim}</td>
                  <td className="p-4 md:p-5 font-bold text-sm md:text-base tracking-wide">{st.name}</td>
                  <td className="p-4 md:p-5 font-mono text-xs text-purple-300">{clusterName}</td>
                  <td className="p-4 md:p-5"><span className="text-[10px] md:text-xs bg-black/60 px-2.5 py-1.5 rounded-md border border-cyan-500/20 font-mono text-cyan-500">{st.password}</span></td>
                  <td className="p-4 md:p-5 text-center">
                    {st.deviceId ? (
                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-emerald-950/40 text-emerald-400 text-[10px] md:text-xs font-mono tracking-widest border border-emerald-500/30 shadow-sm"><CheckCircle2 className="w-3.5 h-3.5"/> TERTAUT</span>
                    ) : (
                      <span className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full bg-slate-900 text-slate-500 text-[10px] md:text-xs font-mono tracking-widest border border-slate-700 shadow-sm">BEBAS</span>
                    )}
                  </td>
                  <td className="p-4 md:p-5 text-right flex justify-end gap-2 md:gap-3">
                    {st.deviceId && (
                      <button onClick={() => handleUnlinkDevice(st.id, st.name)} title="Unlink Instrumen" className="p-2 md:p-2.5 text-amber-400 hover:text-white rounded-xl transition-all duration-300 border border-amber-500/30 bg-amber-950/40 hover:bg-amber-700 shadow-sm hover:shadow-lg active:scale-95">
                         <RefreshCcw className="w-4 h-4" />
                      </button>
                    )}
                    <button onClick={() => setEditingStudent(st)} title="Modifikasi Data" className="p-2 md:p-2.5 text-blue-400 hover:text-white rounded-xl transition-all duration-300 border border-blue-500/30 bg-blue-950/40 hover:bg-blue-700 shadow-sm hover:shadow-lg active:scale-95">
                       <Settings className="w-4 h-4" />
                    </button>
                    <button onClick={() => setSelectedStudentForKTM(st)} title="Render KTM" className="px-3 md:px-4 py-2 md:py-2.5 text-[10px] md:text-xs font-mono tracking-widest font-bold text-cyan-300 bg-cyan-950/40 border border-cyan-500/30 hover:bg-cyan-700 hover:text-white rounded-xl transition-all duration-300 flex items-center gap-2 shadow-sm hover:shadow-lg active:scale-95 uppercase">
                      <CreditCard className="w-4 h-4"/> <span className="hidden md:inline">KTM</span>
                    </button>
                    <button onClick={() => deleteStudent(st.id)} title="Hapus Permanen" className="p-2 md:p-2.5 text-slate-500 hover:text-white hover:bg-rose-900 rounded-xl transition-all duration-300 border border-transparent hover:border-rose-500/50 hover:shadow-[0_0_15px_rgba(244,63,94,0.5)] active:scale-95">
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </td>
                </tr>
              )})}
              {filtered.length === 0 && <tr><td colSpan={6} className="p-10 text-center text-cyan-500/50 font-mono tracking-widest uppercase">Database Kosong / Tidak Ditemukan.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {editingStudent && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 backdrop-blur-md p-4 animate-in fade-in zoom-in-95 duration-200">
           <form onSubmit={handleUpdate} className="bg-slate-900/90 border border-cyan-500/30 p-6 md:p-8 rounded-[2rem] w-full max-w-md shadow-[0_0_50px_rgba(6,182,212,0.15)] relative">
              <div className="flex justify-between items-center mb-6 md:mb-8">
                 <h3 className="text-xl md:text-2xl font-black text-white tracking-widest uppercase">Modifikasi Data</h3>
                 <button type="button" onClick={() => setEditingStudent(null)} className="p-2 bg-white/5 hover:bg-rose-900/60 hover:text-rose-400 rounded-full transition-colors text-slate-400 border border-transparent hover:border-rose-500/30"><X className="w-5 h-5"/></button>
              </div>
              <div className="space-y-4 md:space-y-5">
                 <div className="space-y-1.5">
                    <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Nama Lengkap</label>
                    <input required type="text" value={editingStudent.name} onChange={e=>setEditingStudent({...editingStudent, name: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm" />
                 </div>
                 <div className="space-y-1.5">
                    <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">NIM</label>
                    <input required type="text" value={editingStudent.nim} onChange={e=>setEditingStudent({...editingStudent, nim: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm" />
                 </div>
                 <div className="space-y-1.5">
                    <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Cluster</label>
                    <select value={editingStudent.clusterId || ''} onChange={e=>setEditingStudent({...editingStudent, clusterId: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm appearance-none cursor-pointer">
                       <option value="">-- Kosong --</option>
                       {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
                    </select>
                 </div>
                 <div className="space-y-1.5">
                    <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Kredensial / Password</label>
                    <input required type="text" value={editingStudent.password || ''} onChange={e=>setEditingStudent({...editingStudent, password: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm" />
                 </div>
                 <button type="submit" className="w-full py-4 mt-4 bg-cyan-950/60 hover:bg-cyan-900/80 border border-cyan-500/50 text-cyan-300 font-bold uppercase tracking-widest rounded-2xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.1)] active:scale-95 text-xs">
                    Injeksi Perubahan Database
                 </button>
              </div>
           </form>
        </div>
      )}

      {selectedStudentForKTM && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 backdrop-blur-md p-4 animate-in fade-in zoom-in-95 duration-200">
          <style>{`
            @media print {
              body * { visibility: hidden; }
              #ktm-print-area, #ktm-print-area * { visibility: visible; }
              #ktm-print-area { position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); margin: 0; }
            }
          `}</style>
          <div className="bg-slate-900 border border-cyan-500/30 p-6 md:p-8 rounded-[2rem] w-full max-w-[450px] shadow-[0_0_50px_rgba(6,182,212,0.15)] relative z-50">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-xl md:text-2xl font-black text-white tracking-widest uppercase">Preview KTM Cetak</h3>
              <button onClick={() => setSelectedStudentForKTM(null)} className="p-2 bg-white/5 hover:bg-rose-900/60 hover:text-rose-400 rounded-full transition-colors text-white border border-transparent hover:border-rose-500/30"><X className="w-5 h-5"/></button>
            </div>
            
            {/* Desain KTM bernuansa Sci-Fi Medis */}
            <div id="ktm-print-area" className="w-[320px] md:w-[340px] h-[500px] md:h-[540px] mx-auto bg-slate-950 rounded-[2rem] p-6 relative overflow-hidden shadow-2xl flex flex-col items-center justify-between border-[2px] border-cyan-500/50">
               <div className="absolute top-0 left-0 w-full h-[30%] bg-gradient-to-b from-cyan-900/40 to-transparent"></div>
               <div className="absolute bottom-0 right-0 w-full h-[30%] bg-gradient-to-t from-purple-900/20 to-transparent"></div>
               <div className="absolute inset-0 bg-[url('https://www.transparenttextures.com/patterns/carbon-fibre.png')] opacity-10 mix-blend-overlay"></div>
               
               <div className="text-center relative z-10 w-full mt-2 md:mt-4">
                 <div className="w-14 h-14 md:w-16 md:h-16 bg-slate-900 border border-cyan-500/50 rounded-2xl flex items-center justify-center mx-auto mb-3 shadow-[0_0_20px_rgba(6,182,212,0.3)] p-2 md:p-2.5 overflow-hidden">
                   <img src="/axalogo.png" alt="ABSENSI DEPT. RKGo" className="w-full h-full object-contain brightness-150 grayscale contrast-125" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
                 </div>
                 <h2 className="text-white font-black tracking-[0.2em] text-base md:text-lg drop-shadow-md">DEPT. RKG</h2>
                 <p className="text-cyan-400 text-[8px] md:text-[9px] tracking-[0.25em] font-mono mt-1 opacity-90 uppercase">Kartu Tanda Mahasiswa</p>
               </div>

               <div className="bg-white p-3 md:p-4 rounded-2xl relative z-10 shadow-[0_0_40px_rgba(6,182,212,0.4)] border-4 border-cyan-500/20">
                  <img src={`https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${selectedStudentForKTM.nim}&margin=0`} alt="QR Code" className="w-36 h-36 md:w-40 md:h-40" />
               </div>

               <div className="text-center relative z-10 w-full bg-black/60 p-4 md:p-5 rounded-[1.5rem] backdrop-blur-md border border-cyan-500/30 mb-2 shadow-inner">
                 <h1 className="text-lg md:text-xl font-black text-white uppercase leading-tight mb-1 truncate px-2 tracking-wide">{selectedStudentForKTM.name}</h1>
                 <div className="h-1 w-16 bg-gradient-to-r from-cyan-600 to-cyan-300 mx-auto my-2.5 rounded-full"></div>
                 <p className="text-cyan-300 font-mono text-lg md:text-xl tracking-[0.15em] font-bold">{selectedStudentForKTM.nim}</p>
                 <p className="text-purple-300/70 font-mono text-[9px] md:text-[10px] tracking-widest uppercase mt-2">{clusters.find(c => c.id === selectedStudentForKTM.clusterId)?.name || 'NO CLUSTER'}</p>
               </div>
            </div>

            <button onClick={() => window.print()} className="w-full mt-8 py-4 bg-cyan-950/60 hover:bg-cyan-900/80 border border-cyan-500/50 text-cyan-300 font-black rounded-2xl flex items-center justify-center gap-3 transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.2)] active:scale-95 text-xs uppercase tracking-widest">
              <Printer className="w-4 h-4" /> Render & Cetak Fisik
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

// ==========================================
// COMPONENT: ADMIN CLUSTERS (Poin 1: Kategori Cluster)
// ==========================================
const AdminClusters: React.FC = () => {
  const { clusters, addCluster, updateCluster, deleteCluster } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newCluster, setNewCluster] = useState({ name: '' });
  const [editingCluster, setEditingCluster] = useState<Cluster | null>(null);

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    if(newCluster.name.trim() === '') return;
    addCluster({ name: newCluster.name });
    setIsAdding(false);
    setNewCluster({ name: '' });
  };

  const handleUpdate = (e: React.FormEvent) => {
    e.preventDefault();
    if(editingCluster && editingCluster.name.trim() !== '') {
       updateCluster(editingCluster.id, { name: editingCluster.name });
       setEditingCluster(null);
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col w-full relative max-w-4xl mx-auto">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-slate-900/40 p-5 rounded-[2rem] border border-cyan-500/10">
        <div>
          <h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Manajemen Cluster</h2>
          <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1">Kelola Kategori / Kelompok Mahasiswa</p>
        </div>
        <button onClick={() => setIsAdding(!isAdding)} className="flex w-full md:w-auto justify-center items-center gap-2 px-5 py-3 bg-cyan-900/40 text-cyan-300 hover:bg-cyan-800/60 border border-cyan-500/30 rounded-xl transition-all duration-300 font-bold active:scale-95 shadow-sm uppercase tracking-widest text-xs">
          <Plus className="w-4 h-4" /> Tambah Cluster
        </button>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-slate-900/80 backdrop-blur-md border border-cyan-500/30 p-5 md:p-6 rounded-[1.5rem] flex flex-col md:flex-row gap-4 items-end animate-in slide-in-from-top-4 shadow-[0_0_30px_rgba(6,182,212,0.1)]">
          <div className="space-y-1.5 flex-1 w-full">
            <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Nama Cluster</label>
            <input required type="text" value={newCluster.name} onChange={e=>setNewCluster({name: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 font-mono text-sm" placeholder="Contoh: Cluster I 2025" />
          </div>
          <button type="submit" className="w-full md:w-auto px-8 py-3.5 bg-cyan-950/50 hover:bg-cyan-900 border border-cyan-500/50 text-cyan-300 font-bold rounded-xl transition-all duration-300 shadow-lg active:scale-95 uppercase tracking-widest text-xs whitespace-nowrap">Simpan</button>
        </form>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
         {clusters.map(cluster => (
            <div key={cluster.id} className="bg-slate-900/60 border border-cyan-500/20 p-5 rounded-[1.5rem] flex flex-col justify-between shadow-lg hover:border-cyan-500/50 transition-colors group">
               {editingCluster?.id === cluster.id ? (
                  <form onSubmit={handleUpdate} className="flex flex-col gap-3">
                     <input autoFocus required type="text" value={editingCluster.name} onChange={e=>setEditingCluster({...editingCluster, name: e.target.value})} className="w-full bg-black/60 border border-cyan-400/50 rounded-lg px-3 py-2 text-white outline-none font-mono text-sm" />
                     <div className="flex gap-2">
                        <button type="button" onClick={() => setEditingCluster(null)} className="flex-1 py-1.5 bg-slate-800 text-slate-400 rounded border border-slate-700 text-[10px] uppercase font-bold tracking-wider hover:bg-slate-700">Batal</button>
                        <button type="submit" className="flex-1 py-1.5 bg-cyan-900/60 text-cyan-300 rounded border border-cyan-500/50 text-[10px] uppercase font-bold tracking-wider hover:bg-cyan-800">Save</button>
                     </div>
                  </form>
               ) : (
                  <>
                     <div className="flex items-center gap-3 mb-4">
                        <Layers className="w-6 h-6 text-cyan-500/50" />
                        <h3 className="text-base font-bold text-white font-mono tracking-wide">{cluster.name}</h3>
                     </div>
                     <div className="flex justify-end gap-2 border-t border-cyan-500/10 pt-3">
                        <button onClick={() => setEditingCluster(cluster)} className="p-1.5 text-blue-400 hover:text-white bg-blue-950/30 hover:bg-blue-800 border border-blue-500/20 rounded transition-colors" title="Edit Kategori">
                           <Edit className="w-4 h-4"/>
                        </button>
                        <button onClick={() => {if(confirm(`Hapus permanen cluster ${cluster.name}?`)) deleteCluster(cluster.id);}} className="p-1.5 text-rose-400 hover:text-white bg-rose-950/30 hover:bg-rose-800 border border-rose-500/20 rounded transition-colors" title="Hapus Kategori">
                           <Trash2 className="w-4 h-4"/>
                        </button>
                     </div>
                  </>
               )}
            </div>
         ))}
         {clusters.length === 0 && <div className="col-span-full p-8 text-center text-cyan-500/50 font-mono tracking-widest uppercase">Belum ada cluster terdaftar.</div>}
      </div>
    </div>
  );
};

const AdminGeofence: React.FC = () => {
  const { geofence, updateGeofence } = useAppContext();
  const [lat, setLat] = useState(geofence.lat.toString());
  const [lng, setLng] = useState(geofence.lng.toString());
  const [radius, setRadius] = useState(geofence.radius.toString());
  const [locationName, setLocationName] = useState(geofence.name || 'Kampus Utama');

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    updateGeofence({ lat: parseFloat(lat), lng: parseFloat(lng), radius: parseInt(radius), name: locationName });
    alert('Konfigurasi matriks geospasial berhasil direkam ke Core Database.');
  };

  const getMyLocation = () => {
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (pos) => { setLat(pos.coords.latitude.toString()); setLng(pos.coords.longitude.toString()); },
        () => alert('Sensor GPS tidak memadai. Input manual diperlukan.')
      );
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 max-w-3xl mx-auto">
      <div className="bg-slate-900/40 p-5 rounded-[2rem] border border-cyan-500/10">
        <h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Parameter Geospasial</h2>
        <p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1">Kalibrasi zona radiasi/absensi aktif.</p>
      </div>

      <form onSubmit={handleSave} className="bg-slate-900/60 backdrop-blur-md border border-cyan-500/20 p-6 md:p-8 rounded-[2rem] space-y-6 md:space-y-8 shadow-[0_0_30px_rgba(6,182,212,0.05)]">
        <div className="p-4 md:p-5 bg-cyan-950/40 border border-cyan-500/30 rounded-2xl flex items-start gap-3 md:gap-4 shadow-inner">
          <Navigation className="w-6 h-6 text-cyan-400 mt-0.5 shrink-0 animate-pulse" />
          <p className="text-xs md:text-sm text-cyan-100/80 leading-relaxed font-mono">Restriksi aktif: Subjek hanya dapat menginisiasi sesi absensi jika koordinat GPS *real-time* berada di dalam <b>Radius Toleransi</b> yang dikalkulasi dari episentrum (Lat/Lng) di bawah.</p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-5 md:gap-6">
          <div className="space-y-1.5 md:col-span-2">
             <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Label Episentrum (Display Error)</label>
             <input required type="text" value={locationName} onChange={e=>setLocationName(e.target.value)} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" placeholder="Contoh: Gedung A Kampus" />
          </div>
          <div className="space-y-1.5">
            <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Latitude (Lintang)</label>
            <input required type="number" step="any" value={lat} onChange={e=>setLat(e.target.value)} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" />
          </div>
          <div className="space-y-1.5">
            <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Longitude (Bujur)</label>
            <input required type="number" step="any" value={lng} onChange={e=>setLng(e.target.value)} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-sm" />
          </div>
        </div>

        <div className="space-y-1.5">
           <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Radius Maksimal (Meter)</label>
           <input required type="number" min="10" value={radius} onChange={e=>setRadius(e.target.value)} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-cyan-300 font-bold outline-none focus:border-cyan-400 transition-colors shadow-inner font-mono text-base" />
           <p className="text-[9px] md:text-[10px] text-slate-500 mt-1.5 ml-1 font-mono uppercase">Rekomendasi keamanan: 500 - 1000 meter.</p>
        </div>

        <div className="flex flex-col md:flex-row gap-3 md:gap-4 pt-6 border-t border-cyan-500/10">
          <button type="button" onClick={getMyLocation} className="w-full md:w-auto px-6 py-4 bg-slate-800/50 hover:bg-slate-800 border border-cyan-500/20 text-cyan-100/80 font-bold rounded-xl transition-all duration-300 flex items-center justify-center gap-2 active:scale-95 shadow-sm uppercase tracking-widest text-xs">
            <MapPin className="w-4 h-4" /> Kalibrasi GPS Saat Ini
          </button>
          <button type="submit" className="w-full md:flex-1 py-4 bg-cyan-950/60 hover:bg-cyan-900/80 border border-cyan-500/50 text-cyan-300 font-bold rounded-xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.1)] active:scale-95 uppercase tracking-widest text-xs">
            Injeksi & Sinkronisasi
          </button>
        </div>
      </form>
    </div>
  );
};

// ==========================================
// COMPONENT: ADMIN SETTINGS (Poin 1 CRUD Sesi)
// ==========================================
const AdminSettings: React.FC = () => {
  const { sessions, updateSession, addSession, deleteSession } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newSess, setNewSess] = useState({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });
  const [editingSession, setEditingSession] = useState<Session | null>(null);

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault(); addSession({ ...newSess, isActive: true }); setIsAdding(false); setNewSess({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });
  };

  const handleUpdate = (e: React.FormEvent) => {
     e.preventDefault();
     if(editingSession) {
        updateSession(editingSession.id, {
           name: editingSession.name,
           startTime: editingSession.startTime,
           endTime: editingSession.endTime,
           toleranceMinutes: editingSession.toleranceMinutes
        });
        setEditingSession(null);
     }
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 w-full relative">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-slate-900/40 p-5 rounded-[2rem] border border-cyan-500/10">
        <div><h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Konfigurasi Sesi</h2><p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1">Parameter Shift Waktu & Toleransi Keterlambatan</p></div>
        <button onClick={() => setIsAdding(!isAdding)} className="flex w-full md:w-auto justify-center items-center gap-2 px-5 py-3 bg-cyan-900/40 text-cyan-300 hover:bg-cyan-800/60 border border-cyan-500/30 rounded-xl transition-all duration-300 font-bold active:scale-95 shadow-sm uppercase tracking-widest text-xs"><Plus className="w-4 h-4" /> Inisiasi Sesi Baru</button>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-slate-900/80 backdrop-blur-md border border-cyan-500/30 p-5 md:p-6 rounded-[1.5rem] grid grid-cols-1 md:grid-cols-5 gap-4 md:gap-5 items-end animate-in slide-in-from-top-4 shadow-[0_0_30px_rgba(6,182,212,0.1)]">
          <div className="space-y-1.5"><label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Label Sesi</label><input required type="text" value={newSess.name} onChange={e=>setNewSess({...newSess, name: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" placeholder="e.g. Kuliah Pagi" /></div>
          <div className="space-y-1.5"><label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Jam Start</label><input required type="time" value={newSess.startTime} onChange={e=>setNewSess({...newSess, startTime: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" /></div>
          <div className="space-y-1.5"><label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Jam End</label><input required type="time" value={newSess.endTime} onChange={e=>setNewSess({...newSess, endTime: e.target.value})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" /></div>
          <div className="space-y-1.5"><label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Toleransi (Menit)</label><input required type="number" min="0" value={newSess.toleranceMinutes} onChange={e=>setNewSess({...newSess, toleranceMinutes: parseInt(e.target.value)})} className="w-full bg-black/40 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" /></div>
          <button type="submit" className="w-full py-3.5 bg-cyan-950/60 hover:bg-cyan-900 border border-cyan-500/50 text-cyan-300 font-bold rounded-xl transition-all duration-300 shadow-lg active:scale-95 uppercase tracking-widest text-xs">Simpan Data</button>
        </form>
      )}

      {editingSession && (
         <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 backdrop-blur-md p-4 animate-in fade-in zoom-in-95 duration-200">
            <form onSubmit={handleUpdate} className="bg-slate-900/90 border border-cyan-500/30 p-6 md:p-8 rounded-[2rem] w-full max-w-md shadow-[0_0_50px_rgba(6,182,212,0.15)] relative">
               <div className="flex justify-between items-center mb-6">
                  <h3 className="text-xl md:text-2xl font-black text-white tracking-widest uppercase">Modifikasi Sesi</h3>
                  <button type="button" onClick={() => setEditingSession(null)} className="p-2 bg-white/5 hover:bg-rose-900/60 hover:text-rose-400 rounded-full transition-colors text-slate-400 border border-transparent hover:border-rose-500/30"><X className="w-5 h-5"/></button>
               </div>
               <div className="space-y-4">
                  <div className="space-y-1.5">
                     <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Label Sesi</label>
                     <input required type="text" value={editingSession.name} onChange={e=>setEditingSession({...editingSession, name: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" />
                  </div>
                  <div className="grid grid-cols-2 gap-4">
                     <div className="space-y-1.5">
                        <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Jam Start</label>
                        <input required type="time" value={editingSession.startTime} onChange={e=>setEditingSession({...editingSession, startTime: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" />
                     </div>
                     <div className="space-y-1.5">
                        <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Jam End</label>
                        <input required type="time" value={editingSession.endTime} onChange={e=>setEditingSession({...editingSession, endTime: e.target.value})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" />
                     </div>
                  </div>
                  <div className="space-y-1.5">
                     <label className="text-[10px] md:text-xs text-cyan-500/70 font-mono font-bold uppercase tracking-widest ml-1">Toleransi (Menit)</label>
                     <input required type="number" min="0" value={editingSession.toleranceMinutes} onChange={e=>setEditingSession({...editingSession, toleranceMinutes: parseInt(e.target.value)})} className="w-full bg-black/50 border border-cyan-500/20 rounded-xl px-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 text-sm" />
                  </div>
                  <button type="submit" className="w-full py-4 mt-4 bg-cyan-950/60 hover:bg-cyan-900/80 border border-cyan-500/50 text-cyan-300 font-bold uppercase tracking-widest rounded-2xl transition-all duration-300 shadow-[0_0_20px_rgba(6,182,212,0.1)] active:scale-95 text-xs">Simpan Perubahan</button>
               </div>
            </form>
         </div>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 md:gap-5">
        {sessions.map(session => (
          <div key={session.id} className={cn("p-5 md:p-6 rounded-[1.5rem] border transition-all duration-300 shadow-lg group relative overflow-hidden", session.isActive ? "bg-slate-900/60 border-cyan-500/20 hover:border-cyan-400/50 backdrop-blur-md" : "bg-black/40 opacity-70 border-slate-700/50 hover:opacity-100")}>
            <div className="absolute top-0 left-0 w-1 h-full bg-cyan-500/30"></div>
            <div className="flex justify-between items-start mb-5 pl-2">
              <h3 className="text-xl font-black text-white tracking-wider">{session.name}</h3>
              <div className="flex gap-2">
                <button onClick={() => updateSession(session.id, { isActive: !session.isActive })} className={cn("px-2.5 py-1 text-[9px] md:text-[10px] font-mono font-bold uppercase tracking-widest rounded border transition-all duration-300 shadow-sm active:scale-95", session.isActive ? "bg-emerald-950/40 text-emerald-400 border-emerald-500/30 hover:bg-emerald-900/60" : "bg-slate-800 text-slate-400 border-slate-700 hover:bg-slate-700")}>{session.isActive ? 'AKTIF' : 'NONAKTIF'}</button>
              </div>
            </div>
            <div className="space-y-3 text-xs md:text-sm text-cyan-100/70 bg-black/40 p-4 rounded-xl border border-cyan-500/10 font-mono pl-3">
              <div className="flex items-center gap-3"><Clock className="w-4 h-4 text-cyan-500/50"/> Jendela Waktu: <span className="text-cyan-300 font-bold bg-cyan-950/30 px-2 py-0.5 rounded border border-cyan-500/20">{session.startTime} - {session.endTime}</span></div>
              <div className="flex items-center gap-3"><Activity className="w-4 h-4 text-purple-500/50"/> Batas Deviasi: <span className="text-purple-300 font-bold">{session.toleranceMinutes} min</span></div>
            </div>
            
            {/* Action Buttons */}
            <div className="flex justify-end gap-2 mt-4 pt-3 border-t border-cyan-500/10">
               <button onClick={() => setEditingSession(session)} className="p-1.5 text-blue-400 hover:text-white bg-blue-950/30 hover:bg-blue-800 border border-blue-500/20 rounded transition-colors" title="Edit Sesi">
                  <Edit className="w-4 h-4"/>
               </button>
               <button onClick={() => {if(confirm(`Hapus sesi ${session.name}?`)) deleteSession(session.id);}} className="p-1.5 text-rose-400 hover:text-white bg-rose-950/30 hover:bg-rose-800 border border-rose-500/20 rounded transition-colors" title="Hapus Sesi">
                  <Trash2 className="w-4 h-4" />
               </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

const AdminReports: React.FC = () => {
  const { logs, sessions, clusters, students, deleteLog } = useAppContext();
  const [search, setSearch] = useState('');
  const [filterSession, setFilterSession] = useState('All');
  const [filterCluster, setFilterCluster] = useState('All');
  
  const [previewImage, setPreviewImage] = useState<string | null>(null);

  const filteredLogs = logs.filter(log => {
    const student = students.find(s => s.nim === log.nim);
    const matchSearch = log.name.toLowerCase().includes(search.toLowerCase()) || log.nim.includes(search);
    const matchSession = filterSession === 'All' || log.sessionName === filterSession;
    const matchCluster = filterCluster === 'All' || student?.clusterId === filterCluster;
    return matchSearch && matchSession && matchCluster;
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col relative w-full">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 bg-slate-900/40 p-5 rounded-[2rem] border border-cyan-500/10">
        <div><h2 className="text-2xl md:text-3xl font-black text-white tracking-widest uppercase">Master Log</h2><p className="text-cyan-500/70 text-xs md:text-sm font-mono mt-1">Basis data histori biometrik, waktu, & geospasial.</p></div>
        {/* Poin 3: Export Excel XLSX */}
        <button onClick={() => exportToXLSX(filteredLogs, students, clusters)} className="w-full md:w-auto flex items-center justify-center gap-2 px-6 py-3 bg-emerald-950/60 hover:bg-emerald-900/80 border border-emerald-500/50 text-emerald-300 rounded-xl transition-all duration-300 font-bold shadow-[0_0_15px_rgba(16,185,129,0.2)] active:scale-95 whitespace-nowrap uppercase tracking-widest text-xs"><Download className="w-4 h-4" /> Export Report (XLSX)</button>
      </div>

      <div className="flex flex-col md:flex-row gap-3 md:gap-4">
        <div className="relative flex-1">
           <SearchIcon className="absolute left-4 top-1/2 -translate-y-1/2 w-4 h-4 text-cyan-500/50" />
           <input type="text" placeholder="Cari ID / NIM..." value={search} onChange={e=>setSearch(e.target.value)} className="w-full bg-slate-900/60 border border-cyan-500/20 rounded-xl pl-10 pr-4 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner text-sm placeholder-slate-600" />
        </div>
        <select value={filterCluster} onChange={e=>setFilterCluster(e.target.value)} className="bg-slate-900/60 border border-cyan-500/20 rounded-xl px-5 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner w-full md:w-48 cursor-pointer appearance-none text-sm">
          <option value="All">Semua Cluster</option>
          {clusters.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
        </select>
        <select value={filterSession} onChange={e=>setFilterSession(e.target.value)} className="bg-slate-900/60 border border-cyan-500/20 rounded-xl px-5 py-3.5 text-white font-mono outline-none focus:border-cyan-400 transition-colors shadow-inner w-full md:w-48 cursor-pointer appearance-none text-sm">
          <option value="All">Semua Shift</option>
          {sessions.map(s => <option key={s.id} value={s.name}>Sesi: {s.name}</option>)}
        </select>
      </div>

      <div className="flex-1 bg-slate-900/60 backdrop-blur-md border border-cyan-500/20 rounded-[1.5rem] overflow-hidden flex flex-col shadow-[0_0_20px_rgba(0,0,0,0.3)]">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse min-w-[1000px]">
            <thead>
              <tr className="bg-black/40 border-b border-cyan-500/20 text-cyan-500/70 text-xs tracking-widest uppercase font-mono">
                <th className="p-4 md:p-5 font-bold">Biometrik</th>
                <th className="p-4 md:p-5 font-bold">Data Subjek & Kategori</th>
                <th className="p-4 md:p-5 font-bold">Timestamp Server</th>
                <th className="p-4 md:p-5 font-bold">Sesi & Indikator</th>
                <th className="p-4 md:p-5 font-bold">Radar Geospasial</th>
                <th className="p-4 md:p-5 text-right font-bold">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-cyan-500/10">
              {filteredLogs.map(log => {
                const student = students.find(s => s.nim === log.nim);
                const clusterName = clusters.find(c => c.id === student?.clusterId)?.name || 'Tanpa Cluster';
                
                return (
                <tr key={log.id} className="hover:bg-cyan-950/30 transition-colors duration-200">
                  <td className="p-4 md:p-5">
                    <div onClick={() => setPreviewImage(log.photoBase64)} className="w-14 h-14 md:w-16 md:h-16 rounded-xl overflow-hidden border-2 border-cyan-500/30 bg-black relative group cursor-pointer shadow-md hover:shadow-[0_0_15px_rgba(6,182,212,0.5)] hover:border-cyan-400 transition-all duration-300">
                      <img src={log.photoBase64} alt="Selfie" className="w-full h-full object-cover sepia-[.2] hue-rotate-[180deg]" />
                      <div className="absolute inset-0 bg-black/60 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity backdrop-blur-sm">
                        <Maximize className="w-5 h-5 text-cyan-300" />
                      </div>
                    </div>
                  </td>
                  <td className="p-4 md:p-5">
                     <p className="font-bold text-white text-sm md:text-base truncate max-w-[200px] tracking-wide">{log.name}</p>
                     <p className="text-xs text-cyan-200/60 font-mono mt-0.5">{log.nim}</p>
                     <p className="text-[9px] text-purple-400/80 font-mono uppercase tracking-widest mt-1.5">{clusterName}</p>
                  </td>
                  <td className="p-4 md:p-5">
                     <p className="text-cyan-100 font-bold font-mono text-base">{new Date(log.timestamp).toLocaleTimeString('id-ID')}</p>
                     <p className="text-[10px] md:text-xs text-cyan-500/60 font-mono uppercase tracking-widest mt-0.5">{new Date(log.timestamp).toLocaleDateString('id-ID', { weekday: 'short', day: '2-digit', month: 'short' })}</p>
                  </td>
                  <td className="p-4 md:p-5">
                     <p className="text-slate-200 text-xs md:text-sm font-bold mb-2 tracking-wide uppercase">{log.sessionName}</p>
                     <span className={cn("px-3 py-1 text-[9px] md:text-[10px] font-mono font-bold uppercase tracking-widest rounded border shadow-sm", log.status === 'Hadir' ? "bg-emerald-950/40 text-emerald-400 border-emerald-500/30" : "bg-amber-950/40 text-amber-400 border-amber-500/30")}>{log.status}</span>
                  </td>
                  <td className="p-4 md:p-5">
                    <a href={`https://www.google.com/maps?q=${log.location.lat},${log.location.lng}`} target="_blank" rel="noreferrer" className="inline-flex items-center gap-2 px-3 py-2 bg-cyan-950/40 hover:bg-cyan-900/60 hover:text-white text-cyan-400 text-[9px] md:text-[10px] font-bold uppercase tracking-widest rounded-lg border border-cyan-500/30 transition-all duration-300 shadow-sm active:scale-95">
                      <MapPin className="w-3.5 h-3.5" /> Lacak Sinyal GPS
                    </a>
                    <p className="text-[9px] md:text-[10px] text-cyan-500/50 mt-2 font-mono bg-black/40 inline-block px-2 py-0.5 rounded border border-cyan-500/10">{log.location.lat.toFixed(5)}, {log.location.lng.toFixed(5)}</p>
                  </td>
                  <td className="p-4 md:p-5 text-right">
                    <button onClick={() => { if(confirm(`Peringatan: Aksi ini menghapus log absensi secara permanen. Lanjutkan?`)) deleteLog(log.id); }} title="Purge Record" className="p-2.5 text-slate-500 hover:text-rose-400 hover:bg-rose-950/50 rounded-xl transition-all duration-300 border border-transparent hover:border-rose-500/30 active:scale-95">
                      <Trash2 className="w-5 h-5" />
                    </button>
                  </td>
                </tr>
              )})}
              {filteredLogs.length === 0 && <tr><td colSpan={6} className="p-12 text-center text-cyan-500/50 font-mono tracking-widest uppercase text-sm">Tidak ada matriks data pada log.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
      
      {/* MODAL FULLSCREEN PREVIEW IMAGE RESPONSIVE */}
      {previewImage && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/95 backdrop-blur-xl p-4 animate-in fade-in zoom-in-95 duration-300" onClick={() => setPreviewImage(null)}>
          <div className="relative max-w-3xl w-full flex flex-col items-center justify-center">
            <button onClick={() => setPreviewImage(null)} className="absolute -top-14 md:-top-16 right-0 md:-right-8 p-3 bg-white/5 hover:bg-rose-900/50 hover:text-rose-400 rounded-full transition-all duration-300 text-slate-500 border border-transparent hover:border-rose-500/30">
              <X className="w-6 h-6 md:w-8 md:h-8"/>
            </button>
            <div className="relative w-full overflow-hidden rounded-[2rem] border-[2px] border-cyan-500/50 shadow-[0_0_60px_rgba(6,182,212,0.3)] bg-black p-2">
                <div className="w-full h-full rounded-[1.5rem] overflow-hidden relative">
                   <img src={previewImage} alt="Preview Selfie Fullscreen" className="max-w-full max-h-[75vh] md:max-h-[85vh] w-full object-contain mx-auto sepia-[.2] hue-rotate-[180deg]" onClick={e => e.stopPropagation()} />
                   <div className="absolute inset-0 pointer-events-none bg-[url('https://www.transparenttextures.com/patterns/scan-lines-light.png')] opacity-20 mix-blend-overlay"></div>
                   {/* Grid UI on preview */}
                   <div className="absolute inset-0 pointer-events-none border-[1px] border-cyan-500/20">
                      <div className="absolute top-1/2 w-full h-[1px] bg-cyan-500/20"></div>
                      <div className="absolute left-1/2 h-full w-[1px] bg-cyan-500/20"></div>
                   </div>
                </div>
            </div>
            <p className="mt-5 text-cyan-500/50 text-[10px] md:text-xs font-mono tracking-widest uppercase bg-black/60 px-4 py-2 rounded-full border border-cyan-500/20 shadow-inner">Ketuk area luar untuk terminasi pratinjau</p>
          </div>
        </div>
      )}
    </div>
  );
};

const SearchIcon = ({ className }: { className?: string }) => (
  <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className={className}><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></svg>
);

const AdminLayout: React.FC<{ children: React.ReactNode, activeRoute: string, setRoute: (r:string)=>void }> = ({ children, activeRoute, setRoute }) => {
  const { syncStatus, forceManualSync } = useAppContext();
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const handleLogout = () => { localStorage.removeItem('axaxyz_admin_auth'); setRoute('admin-login'); };

  const navItems = [
    { id: 'admin-dashboard', icon: BarChart3, label: 'Data Analitik' },
    { id: 'admin-students', icon: Database, label: 'Data Subjek / MHS' },
    { id: 'admin-clusters', icon: Layers, label: 'Manajemen Cluster' },
    { id: 'admin-reports', icon: FileText, label: 'Master Log Laporan' },
    { id: 'admin-geofence', icon: Map, label: 'Parameter Geospasial' },
    { id: 'admin-settings', icon: Settings, label: 'Konfigurasi Sesi' },
  ];

  return (
    <div className="min-h-screen bg-slate-950 flex text-slate-200 font-sans w-full overflow-hidden relative selection:bg-cyan-500/30">
      
      {/* RADIOLOGY BACKGROUND */}
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_top,_var(--tw-gradient-stops))] from-slate-900 via-slate-950 to-slate-950"></div>
      <div className="absolute inset-0 bg-[url('https://www.transparenttextures.com/patterns/carbon-fibre.png')] opacity-5 mix-blend-overlay"></div>
      
      {/* MOBILE MENU OVERLAY */}
      {isMobileMenuOpen && (
         <div className="fixed inset-0 bg-black/90 z-40 md:hidden backdrop-blur-md animate-in fade-in duration-300" onClick={() => setIsMobileMenuOpen(false)}></div>
      )}

      {/* RESPONSIVE SIDEBAR */}
      <aside className={cn(
         "fixed inset-y-0 left-0 z-50 w-[280px] md:w-72 bg-slate-950/95 md:bg-slate-950/60 border-r border-cyan-500/10 flex flex-col backdrop-blur-2xl transition-transform duration-500 ease-[cubic-bezier(0.16,1,0.3,1)] md:relative md:translate-x-0 shadow-[20px_0_50px_rgba(0,0,0,0.5)] md:shadow-none",
         isMobileMenuOpen ? "translate-x-0" : "-translate-x-full"
      )}>
        <div className="p-6 md:p-8 border-b border-cyan-500/10 flex items-center justify-between bg-black/20">
          <div className="flex items-center gap-3 md:gap-4">
             <div className="w-10 h-10 md:w-12 md:h-12 bg-slate-900 rounded-xl md:rounded-2xl flex items-center justify-center shadow-[0_0_15px_rgba(6,182,212,0.3)] overflow-hidden p-1.5 md:p-2 border border-cyan-500/30">
                <img src="/axalogo.png" alt="ABSENSI DEPT. RKG" className="w-full h-full object-contain brightness-150 grayscale contrast-125" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
             </div>
             <div>
                <h1 className="font-black text-xl md:text-2xl tracking-[0.15em] text-cyan-50 drop-shadow-[0_0_5px_rgba(6,182,212,0.5)]">R.K.G</h1>
                <p className="text-[8px] md:text-[9px] uppercase tracking-[0.2em] text-cyan-500/80 font-mono mt-0.5">Admin Security Core</p>
             </div>
          </div>
          <button className="md:hidden p-2 bg-white/5 hover:bg-white/10 rounded-xl text-slate-400 transition-colors" onClick={() => setIsMobileMenuOpen(false)}>
             <X className="w-5 h-5"/>
          </button>
        </div>
        
        <nav className="flex-1 p-4 md:p-5 space-y-2 overflow-y-auto custom-scrollbar">
          {navItems.map(item => (
            <button key={item.id} onClick={() => { setRoute(item.id); setIsMobileMenuOpen(false); }} className={cn("w-full flex items-center gap-4 px-5 py-3.5 rounded-2xl transition-all duration-300 text-xs font-bold tracking-widest uppercase font-mono active:scale-[0.98]", activeRoute === item.id ? "bg-cyan-950/40 text-cyan-300 border border-cyan-500/40 shadow-[inset_0_0_20px_rgba(6,182,212,0.15)]" : "text-cyan-100/50 hover:bg-white/5 hover:text-cyan-100/90 border border-transparent hover:border-cyan-500/10")}>
              <item.icon className={cn("w-4 h-4 transition-transform duration-300", activeRoute === item.id && "scale-110 text-cyan-400")} /> {item.label}
            </button>
          ))}
        </nav>
        
        <div className="p-4 md:p-5 border-t border-cyan-500/10 bg-black/20">
          <button onClick={handleLogout} className="w-full flex items-center justify-center gap-3 px-5 py-3.5 rounded-2xl text-rose-400/80 bg-rose-950/20 hover:bg-rose-900/60 hover:text-rose-300 transition-all duration-300 text-xs tracking-widest uppercase font-mono font-bold border border-rose-500/20 active:scale-95 shadow-sm hover:border-rose-500/50 hover:shadow-[0_0_15px_rgba(244,63,94,0.3)]"><LogOut className="w-4 h-4" /> Terminasi Akses</button>
        </div>
      </aside>

      <main className="flex-1 relative overflow-y-auto w-full h-screen custom-scrollbar bg-slate-950/50">
        
        {/* RESPONSIVE HEADER & STATUS BADGE */}
        <header className="sticky top-0 p-4 md:p-5 flex justify-between md:justify-end items-center z-30 w-full bg-slate-950/80 backdrop-blur-xl border-b border-cyan-500/10 shadow-lg">
           <button className="md:hidden p-2.5 bg-slate-900 border border-cyan-500/20 hover:bg-slate-800 rounded-xl text-cyan-400 transition-colors active:scale-95 shadow-sm" onClick={() => setIsMobileMenuOpen(true)}>
              <Menu className="w-5 h-5" />
           </button>
           
           <div className="flex items-center gap-3">
               <button onClick={forceManualSync} className="p-2 md:px-4 md:py-2 bg-slate-900 hover:bg-cyan-950/50 border border-cyan-500/20 hover:border-cyan-500/50 text-cyan-400 rounded-full md:rounded-xl transition-all duration-300 flex items-center gap-2 shadow-sm active:scale-95" title="Force Sync ke Cloud">
                  <Upload className="w-3.5 h-3.5 md:w-4 md:h-4" />
                  <span className="hidden md:inline text-[9px] md:text-[10px] font-mono tracking-widest uppercase font-bold">Injeksi Cloud</span>
               </button>
               
               <div className="flex items-center gap-2 px-3 md:px-4 py-2 bg-black/60 border border-cyan-500/20 rounded-full md:rounded-xl text-[9px] md:text-[10px] font-mono font-bold shadow-inner">
                   {syncStatus === 'syncing' && <><RefreshCcw className="w-3.5 h-3.5 md:w-4 md:h-4 animate-spin text-cyan-400"/> <span className="text-cyan-400 tracking-widest uppercase hidden md:inline">Syncing...</span></>}
                   {syncStatus === 'synced' && <><CheckCircle2 className="w-3.5 h-3.5 md:w-4 md:h-4 text-emerald-400"/> <span className="text-emerald-400 tracking-widest uppercase hidden md:inline">Cloud Terhubung</span></>}
                   {syncStatus === 'error' && <><CloudOff className="w-3.5 h-3.5 md:w-4 md:h-4 text-rose-400"/> <span className="text-rose-400 tracking-widest uppercase hidden md:inline">Sync Gagal</span></>}
                   {syncStatus === 'offline' && <><CloudOff className="w-3.5 h-3.5 md:w-4 md:h-4 text-slate-500"/> <span className="text-slate-500 tracking-widest uppercase hidden md:inline">Mode Lokal</span></>}
               </div>
           </div>
        </header>

        <div className="absolute top-[10%] right-[-10%] w-[50%] h-[50%] bg-cyan-900/10 rounded-full blur-[120px] pointer-events-none"></div>
        <div className="p-4 md:p-8 max-w-7xl mx-auto relative z-10 min-h-full pb-20">{children}</div>
      </main>
      
      {/* GLOBAL SCROLLBAR STYLING */}
      <style>{`
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: transparent; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: rgba(6, 182, 212, 0.2); border-radius: 10px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: rgba(6, 182, 212, 0.4); }
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
      document.title = "DEPT. RKG - Absensi Biometrik";
    }
  }, []);

  useEffect(() => {
    const isAdminAuthed = localStorage.getItem('axaxyz_admin_auth') === 'true';
    if (route.startsWith('admin-') && route !== 'admin-login' && !isAdminAuthed) setRoute('admin-login');
  }, [route]);

  return (
    <AppProvider>
      <div className="fixed bottom-4 md:bottom-6 right-4 md:right-6 z-[999] flex gap-2 md:gap-3 bg-black/80 backdrop-blur-xl p-2.5 rounded-[1.5rem] border border-cyan-500/20 shadow-[0_0_30px_rgba(0,0,0,0.8)]">
        <button onClick={() => setRoute('student')} className={cn("px-4 md:px-5 py-2.5 md:py-3 rounded-xl text-[9px] md:text-[10px] font-mono font-bold uppercase tracking-widest transition-all duration-300 active:scale-95 shadow-sm", route === 'student' ? "bg-cyan-900/60 text-cyan-300 border border-cyan-500/50 shadow-[0_0_15px_rgba(6,182,212,0.3)]" : "bg-white/5 text-slate-400 hover:bg-cyan-950/40 hover:text-cyan-300 border border-transparent hover:border-cyan-500/30")}>Portal MHS</button>
        <button onClick={() => setRoute(typeof window !== 'undefined' && localStorage.getItem('axaxyz_admin_auth') === 'true' ? 'admin-dashboard' : 'admin-login')} className={cn("px-4 md:px-5 py-2.5 md:py-3 rounded-xl text-[9px] md:text-[10px] font-mono font-bold uppercase tracking-widest transition-all duration-300 active:scale-95 shadow-sm", route.startsWith('admin') ? "bg-purple-900/60 text-purple-300 border border-purple-500/50 shadow-[0_0_15px_rgba(168,85,247,0.3)]" : "bg-white/5 text-slate-400 hover:bg-purple-950/40 hover:text-purple-300 border border-transparent hover:border-purple-500/30")}>Admin Core</button>
      </div>

      {route === 'student' && <AttendanceWizard />}
      {route === 'admin-login' && <AdminLogin onLogin={() => setRoute('admin-dashboard')} />}
      
      {['admin-dashboard', 'admin-students', 'admin-clusters', 'admin-settings', 'admin-reports', 'admin-geofence'].includes(route) && (
        <AdminLayout activeRoute={route} setRoute={setRoute}>
          {route === 'admin-dashboard' && <AdminDashboardHome />}
          {route === 'admin-students' && <AdminStudents />}
          {route === 'admin-clusters' && <AdminClusters />}
          {route === 'admin-geofence' && <AdminGeofence />}
          {route === 'admin-settings' && <AdminSettings />}
          {route === 'admin-reports' && <AdminReports />}
        </AdminLayout>
      )}
    </AppProvider>
  );
}
