"use client";

import React, { useState, useEffect, createContext, useContext, useRef, useCallback, useMemo } from 'react';
import { 
  Camera, MapPin, Clock, QrCode, CheckCircle2, AlertCircle, 
  BarChart3, Settings, FileText, LogOut, Users, Download, Plus, Trash2,
  RefreshCcw, ChevronRight, Fingerprint, Map, Activity, Key, Upload, Database, Navigation,
  Printer, X, CreditCard, Eye, EyeOff, Lock, ShieldCheck, Loader2, User, Cloud, CloudOff,
  Server, ServerCrash, DatabaseZap, Maximize
} from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell } from 'recharts';

import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { format } from 'date-fns';

// HAPUS TANDA KOMENTAR (//) PADA BARIS DI BAWAH INI SAAT DI-DEPLOY KE VERCEL
// import { SpeedInsights } from "@vercel/speed-insights/next";

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

const exportToCSV = (logs: Log[]) => {
  const headers = ['ID,NIM,Name,Date,Time,Session,Status,Lat,Lng,Maps Link'];
  const rows = logs.map(log => {
    const date = new Date(log.timestamp).toLocaleDateString('id-ID');
    const time = new Date(log.timestamp).toLocaleTimeString('id-ID');
    const mapsLink = `https://www.google.com/maps?q=${log.location.lat},${log.location.lng}`;
    return `${log.id},${log.nim},${log.name},${date},${time},${log.sessionName},${log.status},${log.location.lat},${log.location.lng},${mapsLink}`;
  });
  const csvContent = "data:text/csv;charset=utf-8," + [headers, ...rows].join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement("a");
  link.setAttribute("href", encodedUri);
  link.setAttribute("download", `axaxyz_report_${new Date().toISOString().split('T')[0]}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
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

interface Session { id: string; name: string; startTime: string; endTime: string; toleranceMinutes: number; isActive: boolean; }
interface Log { id: string; nim: string; name: string; timestamp: string; sessionName: string; status: 'Hadir' | 'Terlambat'; location: { lat: number; lng: number }; photoBase64: string; deviceId: string; }
interface Student { id: string; nim: string; name: string; password?: string; deviceId?: string | null; }
// UPDATED: Geofence now supports Location Name
interface Geofence { lat: number; lng: number; radius: number; name?: string; }

type SyncStatus = 'offline' | 'synced' | 'syncing' | 'error';

interface AppContextType {
  sessions: Session[];
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
  addStudent: (student: Omit<Student, 'id'>) => void;
  updateStudent: (id: string, updates: Partial<Student>) => void;
  bulkAddStudents: (newStudents: Omit<Student, 'id'>[]) => void;
  deleteStudent: (id: string) => void;
  updateGeofence: (data: Geofence) => void;
  forceManualSync: () => Promise<void>;
}

const defaultSessions: Session[] = [
  { id: '1', name: 'Pagi', startTime: '07:00', endTime: '09:00', toleranceMinutes: 15, isActive: true },
  { id: '2', name: 'Siang', startTime: '12:00', endTime: '13:30', toleranceMinutes: 15, isActive: true },
  { id: '3', name: 'Sore', startTime: '16:00', endTime: '17:30', toleranceMinutes: 15, isActive: true },
];

const defaultGeofence: Geofence = { lat: -6.200000, lng: 106.816666, radius: 500, name: 'Kampus Utama' };

const AppContext = createContext<AppContextType | null>(null);

const AppProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const [isAppLoading, setIsAppLoading] = useState(true);
  const [isCloudSync, setIsCloudSync] = useState(false);
  const [syncStatus, setSyncStatus] = useState<SyncStatus>('offline');
  
  const [sessions, setSessions] = useState<Session[]>([]);
  const [logs, setLogs] = useState<Log[]>([]);
  const [students, setStudents] = useState<Student[]>([]);
  const [geofence, setGeofence] = useState<Geofence>(defaultGeofence);

  useEffect(() => {
    const initData = async () => {
      const cloudAvailable = CloudStore.isAvailable();
      setIsCloudSync(cloudAvailable);
      setSyncStatus(cloudAvailable ? 'synced' : 'offline');

      let s = null, l = null, st = null, gf = null;

      if (cloudAvailable) {
        s = await CloudStore.get('axaxyz_sessions');
        l = await CloudStore.get('axaxyz_logs');
        st = await CloudStore.get('axaxyz_students');
        gf = await CloudStore.get('axaxyz_geofence');
      }

      if (!s) s = JSON.parse(localStorage.getItem('axaxyz_sessions') || 'null');
      if (!l) l = JSON.parse(localStorage.getItem('axaxyz_logs') || 'null');
      if (!st) st = JSON.parse(localStorage.getItem('axaxyz_students') || 'null');
      if (!gf) gf = JSON.parse(localStorage.getItem('axaxyz_geofence') || 'null');

      setSessions(s || defaultSessions);
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
      alert("❌ Sinkronisasi Gagal: Konfigurasi NEXT_PUBLIC_... Upstash tidak terbaca di Environment Variables.");
      return;
    }
    setSyncStatus('syncing');
    try {
      await CloudStore.set('axaxyz_sessions', JSON.stringify(sessions));
      await CloudStore.set('axaxyz_logs', JSON.stringify(logs));
      await CloudStore.set('axaxyz_students', JSON.stringify(students));
      await CloudStore.set('axaxyz_geofence', JSON.stringify(geofence));
      setSyncStatus('synced');
      alert("✅ Sinkronisasi Force Sync ke Upstash Redis berhasil!");
    } catch (e: any) {
      console.error(e);
      setSyncStatus('error');
      alert("❌ Error saat sinkronisasi: " + e.message);
    }
  };

  const saveSessions = (d: Session[]) => { setSessions(d); localStorage.setItem('axaxyz_sessions', JSON.stringify(d)); syncToCloud('axaxyz_sessions', d); };
  const saveLogs = (d: Log[]) => { setLogs(d); localStorage.setItem('axaxyz_logs', JSON.stringify(d)); syncToCloud('axaxyz_logs', d); };
  const saveStudents = (d: Student[]) => { setStudents(d); localStorage.setItem('axaxyz_students', JSON.stringify(d)); syncToCloud('axaxyz_students', d); };
  const saveGeofence = (d: Geofence) => { setGeofence(d); localStorage.setItem('axaxyz_geofence', JSON.stringify(d)); syncToCloud('axaxyz_geofence', d); };

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

  if (isAppLoading) {
    return (
      <div className="min-h-screen bg-slate-950 flex flex-col items-center justify-center p-4">
        <div className="relative mb-6">
          <div className="absolute inset-0 bg-cyan-500/20 blur-3xl rounded-full"></div>
          <div className="w-20 h-20 bg-gradient-to-br from-cyan-400 to-purple-600 rounded-3xl flex items-center justify-center shadow-lg animate-pulse relative z-10 overflow-hidden p-3">
             <img src="/axalogo.png" alt="AXAXYZ Logo" className="w-full h-full object-contain drop-shadow-md" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
             <span className="font-bold text-white text-3xl hidden">A.</span>
          </div>
        </div>
        <Loader2 className="w-8 h-8 text-cyan-400 animate-spin mb-4" />
        <h2 className="text-xl font-bold text-white mb-2">Menyinkronkan Database...</h2>
        <p className="text-slate-400 text-sm text-center max-w-xs">Memverifikasi koneksi cloud Upstash Redis.</p>
      </div>
    );
  }

  return (
    <AppContext.Provider value={{ 
      isCloudSync, syncStatus, sessions, logs, students, geofence, 
      addLog, deleteLog, updateSession, addSession, deleteSession, addStudent, updateStudent, bulkAddStudents, deleteStudent, updateGeofence, forceManualSync 
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
    <div className="flex flex-col items-center justify-center p-8 space-y-8 animate-in fade-in zoom-in duration-500 w-full">
      <div className="relative">
        <div className="absolute inset-0 bg-cyan-500/20 blur-3xl rounded-full"></div>
        <Clock className="w-24 h-24 text-cyan-400 relative z-10" />
      </div>
      <div className="text-center space-y-2">
        <h2 className="text-4xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-cyan-400 to-purple-400">
          {format(currentTime, 'HH.mm.ss')}
        </h2>
        <p className="text-slate-400">{currentTime.toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
      </div>

      <div className="w-full max-w-md bg-white/5 backdrop-blur-xl border border-white/10 p-6 rounded-2xl shadow-2xl">
        {activeSession ? (
          <div className="space-y-6">
            <div className="flex items-center justify-between p-4 bg-slate-900/50 border border-white/5 rounded-xl">
              <div>
                <p className="text-white font-semibold">Sesi Aktif: {activeSession.session.name}</p>
                <p className="text-sm text-slate-400">{activeSession.session.startTime} - {activeSession.session.endTime}</p>
                <p className="text-xs text-slate-500 mt-1">
                  Batas Tepat Waktu: {activeSession.session.startTime.split(':')[0]}:{String(parseInt(activeSession.session.startTime.split(':')[1]) + activeSession.session.toleranceMinutes).padStart(2, '0')}
                </p>
              </div>
              <span className={cn("px-3 py-1 text-xs font-bold rounded-full", activeSession.status === 'Hadir' ? "bg-emerald-500/20 text-emerald-400" : "bg-orange-500/20 text-orange-400")}>
                {activeSession.status}
              </span>
            </div>
            <button onClick={() => onComplete({ sessionName: activeSession.session.name, status: activeSession.status as 'Hadir' | 'Terlambat' })} className="w-full py-4 bg-gradient-to-r from-cyan-500 to-purple-600 hover:from-cyan-400 hover:to-purple-500 text-white font-bold rounded-xl transition-all shadow-[0_0_20px_rgba(6,182,212,0.3)] hover:shadow-[0_0_30px_rgba(168,85,247,0.5)] flex items-center justify-center gap-2">
              Mulai Absensi <ChevronRight className="w-5 h-5" />
            </button>
          </div>
        ) : (
          <div className="space-y-4 text-center">
            <div className="p-4 bg-rose-500/10 border border-rose-500/20 rounded-xl text-rose-400 flex flex-col items-center gap-2">
              <AlertCircle className="w-8 h-8" />
              <p className="font-semibold">Absensi Ditutup</p>
              <p className="text-sm">Tidak ada sesi absensi yang aktif saat ini.</p>
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
      setStatus('error'); setErrorMsg('Geolocation tidak didukung di browser ini.'); return;
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
          // DYNAMIC NAME UPGRADE: Message now uses Geofence Location Name
          setErrorMsg(`Anda berada di luar radius area ${geofence.name || 'kampus'}.`);
        }
      },
      (error) => { setStatus('error'); setErrorMsg('Gagal mendapatkan lokasi. Pastikan GPS aktif dan diizinkan.'); },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  }, [onComplete, geofence]);

  useEffect(() => { checkLocation(); }, [checkLocation]);

  return (
    <div className="flex flex-col items-center justify-center p-8 space-y-6 max-w-md mx-auto animate-in slide-in-from-right duration-500 w-full">
      <div className="w-20 h-20 bg-white/5 rounded-full flex items-center justify-center border border-white/10 relative overflow-hidden">
        {status === 'loading' && <div className="absolute inset-0 border-4 border-cyan-500/30 border-t-cyan-500 rounded-full animate-spin"></div>}
        <MapPin className={cn("w-10 h-10", status === 'error' ? 'text-rose-400' : 'text-cyan-400')} />
      </div>

      <div className="text-center space-y-2 w-full bg-white/5 backdrop-blur-xl border border-white/10 p-6 rounded-2xl shadow-xl">
        <h3 className="text-xl font-bold text-white">Validasi Lokasi</h3>
        {status === 'loading' && <p className="text-slate-400">Mengambil koordinat GPS Anda...</p>}
        {status === 'success' && (
          <div className="text-emerald-400 space-y-2 animate-in fade-in zoom-in">
            <CheckCircle2 className="w-12 h-12 mx-auto" />
            <p className="font-semibold">Lokasi Valid!</p>
            <p className="text-sm text-emerald-500/80">Jarak: {Math.round(distance || 0)}m dari pusat.</p>
          </div>
        )}
        {status === 'error' && (
          <div className="space-y-4 animate-in fade-in zoom-in">
            <div className="p-4 bg-rose-500/10 border border-rose-500/20 rounded-xl text-center">
              <p className="text-sm font-semibold text-rose-400 leading-relaxed">{errorMsg}</p>
              {distance && <p className="text-xs text-rose-500/70 mt-2">Jarak saat ini: {Math.round(distance)}m (Maks: {geofence.radius}m)</p>}
            </div>
            {/* UPGRADE ANTI-BYPASS: Bypass button entirely removed. User is forced to be in location. */}
            <button onClick={checkLocation} className="w-full py-4 bg-white/10 hover:bg-white/20 text-white rounded-xl text-sm font-bold flex items-center justify-center gap-2 transition-all">
              <RefreshCcw className="w-5 h-5" /> Coba Lagi
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
    if (!targetNim) { setError('Masukkan atau Scan NIM Anda.'); return; }

    let studentName = 'Mahasiswa Tidak Dikenal';
    let finalDeviceId = localStorage.getItem('axaxyz_device_id');
    
    if (!finalDeviceId) {
      finalDeviceId = Math.random().toString(36).substring(2, 15);
      localStorage.setItem('axaxyz_device_id', finalDeviceId);
    }
    
    if (students.length > 0) {
      if (!passInput && !scannedNim) { setError('Masukkan Password Anda.'); return; }
      const foundStudent = students.find(s => s.nim === targetNim);
      if (!foundStudent) {
        setError('NIM tidak terdaftar di sistem.'); return;
      }
      
      if (!scannedNim && foundStudent.password !== passInput) {
        setError('Password salah.'); return;
      }
      studentName = foundStudent.name;

      // VERIFIKASI DEVICE KE DATABASE (ANTI-FRAUD SUPER)
      if (foundStudent.deviceId && foundStudent.deviceId !== finalDeviceId) {
        setError('⚠️ Fraud Alert: Akun (NIM) ini sudah tertaut pada perangkat/HP lain. Hubungi Admin untuk melakukan Reset Device.');
        return;
      }
      
      // BIND DEVICE OTOMATIS JIKA BELUM TERTAUT
      if (!foundStudent.deviceId) {
        updateStudent(foundStudent.id, { deviceId: finalDeviceId });
      }
    } else {
      studentName = 'Mahasiswa Mode Bypass'; 
      // LOKAL BINDING JIKA DATABASE KOSONG
      let deviceOwner = localStorage.getItem('axaxyz_device_owner');
      if (!deviceOwner) {
        localStorage.setItem('axaxyz_device_owner', targetNim); 
      } else if (deviceOwner !== targetNim) {
        setError('⚠️ Fraud Alert: Perangkat ini sudah terdaftar untuk NIM lain.');
        return;
      }
    }

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
             setError('Gagal mengakses kamera. Gunakan input manual atau periksa izin browser.');
             setIsScanning(false);
          });
        } catch (err) {
          setError('Terjadi kesalahan sistem kamera.');
          setIsScanning(false);
        }
      }, 100);
    } catch (error) {
      setError('Library QR Scanner gagal dimuat. Pastikan modul terinstall dengan benar.');
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
    <div className="flex flex-col items-center justify-center p-8 space-y-6 max-w-md mx-auto animate-in slide-in-from-right duration-500 w-full">
      <div className="w-full bg-white/5 backdrop-blur-xl border border-white/10 p-8 rounded-3xl shadow-xl">
        <div className="text-center mb-8">
          <div className="w-16 h-16 bg-gradient-to-tr from-cyan-500 to-purple-500 rounded-2xl flex items-center justify-center mx-auto mb-4 shadow-lg">
            {isScanning ? <Camera className="w-8 h-8 text-white animate-pulse" /> : <QrCode className="w-8 h-8 text-white" />}
          </div>
          <h3 className="text-2xl font-bold text-white mb-2">Identitas Mahasiswa</h3>
          <p className="text-slate-400 text-sm">Validasi akses absensi dengan KTM Anda.</p>
        </div>

        {isScanning ? (
          <div className="space-y-4 animate-in fade-in zoom-in">
             <div id="qr-reader-box" className="w-full rounded-2xl overflow-hidden border-2 border-dashed border-cyan-500 bg-black aspect-square"></div>
             <p className="text-xs text-center text-cyan-400">Arahkan kamera ke QR Code di KTM Anda...</p>
             <button onClick={stopScanner} className="w-full py-3 bg-white/10 text-white rounded-xl text-sm font-bold">Batalkan Scan</button>
          </div>
        ) : (
          <div className="space-y-5">
            <button onClick={startScanner} className="w-full py-4 bg-cyan-500/20 hover:bg-cyan-500/30 border border-cyan-500/50 text-cyan-400 font-bold rounded-xl flex justify-center items-center gap-2 transition-all">
              <Camera className="w-5 h-5" /> Buka Kamera Scanner KTM
            </button>
            
            <div className="relative flex items-center py-2">
               <div className="flex-grow border-t border-white/10"></div>
               <span className="flex-shrink-0 mx-4 text-slate-500 text-xs">ATAU INPUT MANUAL</span>
               <div className="flex-grow border-t border-white/10"></div>
            </div>

            <div className="space-y-2">
              <label className="text-xs text-slate-400 font-semibold uppercase tracking-wider">NIM / Nomor Induk</label>
              <div className="flex items-center bg-slate-900/50 border border-white/10 rounded-xl overflow-hidden focus-within:border-cyan-500 transition-colors">
                <div className="pl-4 pr-2 text-slate-500"><Fingerprint className="w-5 h-5"/></div>
                <input type="text" placeholder="Masukkan NIM..." className="w-full bg-transparent py-3 pr-4 text-white outline-none placeholder-slate-600" value={nimInput} onChange={(e) => setNimInput(e.target.value)} />
              </div>
            </div>
            
            {students.length > 0 && (
              <div className="space-y-2">
                <label className="text-xs text-slate-400 font-semibold uppercase tracking-wider">Password</label>
                <div className="flex items-center bg-slate-900/50 border border-white/10 rounded-xl overflow-hidden focus-within:border-cyan-500 transition-colors">
                  <div className="pl-4 pr-2 text-slate-500"><Key className="w-5 h-5"/></div>
                  <input type="password" placeholder="Masukkan Password..." className="w-full bg-transparent py-3 pr-4 text-white outline-none placeholder-slate-600" value={passInput} onChange={(e) => setPassInput(e.target.value)} />
                </div>
              </div>
            )}
            
            {error && (
              <div className="p-3 bg-rose-500/10 border border-rose-500/20 rounded-lg flex items-start gap-3 animate-in shake">
                <AlertCircle className="w-5 h-5 text-rose-400 shrink-0 mt-0.5" />
                <p className="text-sm text-rose-400 leading-tight">{error}</p>
              </div>
            )}

            <button onClick={() => handleVerify()} className="w-full py-3.5 bg-gradient-to-r from-cyan-500 to-purple-600 hover:from-cyan-400 hover:to-purple-500 text-white font-bold rounded-xl transition-all shadow-[0_0_15px_rgba(6,182,212,0.2)]">
              Lanjutkan
            </button>
          </div>
        )}
      </div>
      <p className="text-xs text-slate-500 text-center max-w-xs">
        Sistem dilengkapi Device Fingerprinting. Satu perangkat hanya dapat digunakan untuk satu identitas NIM.
      </p>
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
        // UPGRADE NO-MIRRORING: 
        // Menghapus fungsi flip (ctx.scale & translate) agar foto disimpan natural apa adanya.
        ctx.drawImage(videoRef.current, 0, 0, canvas.width, canvas.height);
        setImage(canvas.toDataURL('image/jpeg', 0.8));
      }
    }
  }, []);

  return (
    <div className="flex flex-col items-center justify-center p-6 space-y-6 w-full max-w-md mx-auto animate-in slide-in-from-right duration-500">
      <div className="text-center">
        <h3 className="text-2xl font-bold text-white">Verifikasi Wajah</h3>
        <p className="text-slate-400 text-sm mt-1">Posisikan wajah Anda di tengah layar.</p>
      </div>

      <div className="w-full bg-slate-900 rounded-3xl overflow-hidden border-4 border-white/10 relative shadow-2xl aspect-[3/4] md:aspect-video flex items-center justify-center bg-black">
        {!image ? (
          <video
            ref={videoRef}
            autoPlay
            playsInline
            muted
            // UX Standard: Viewfinder tetap ter-mirror sebagai cermin untuk memudahkan posisi muka
            className="w-full h-full object-cover transform scale-x-[-1]"
          />
        ) : (
          // FOTO TERSIMPAN: Tidak lagi menggunakan scale-x-[-1] agar sesuai dengan asli
          <img src={image} alt="Selfie" className="w-full h-full object-cover" />
        )}
        
        {!image && (
          <div className="absolute inset-0 pointer-events-none border-[40px] border-black/30 flex items-center justify-center">
             <div className="w-48 h-64 border-2 border-dashed border-white/50 rounded-[4rem]"></div>
          </div>
        )}
      </div>

      <div className="w-full">
        {!image ? (
          <button onClick={capture} className="w-full py-4 bg-white/10 hover:bg-white/20 border border-white/20 text-white font-bold rounded-2xl transition-all flex items-center justify-center gap-2 group">
            <div className="w-6 h-6 rounded-full border-2 border-white flex items-center justify-center group-hover:bg-white transition-colors">
              <div className="w-2 h-2 rounded-full bg-white group-hover:bg-slate-900"></div>
            </div>
            Ambil Foto
          </button>
        ) : (
          <div className="flex gap-4">
            <button onClick={() => { setImage(null); startCamera(); }} className="flex-1 py-4 bg-white/5 hover:bg-white/10 border border-white/10 text-white font-semibold rounded-2xl transition-all">Ulangi</button>
            <button onClick={() => onComplete(image)} className="flex-1 py-4 bg-gradient-to-r from-cyan-500 to-purple-600 hover:from-cyan-400 hover:to-purple-500 text-white font-bold rounded-2xl transition-all shadow-lg flex items-center justify-center gap-2">
              <CheckCircle2 className="w-5 h-5" /> Konfirmasi
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
      <div className="absolute inset-0 bg-emerald-500/20 blur-3xl rounded-full"></div>
      <CheckCircle2 className="w-32 h-32 text-emerald-400 relative z-10 animate-bounce" style={{ animationDuration: '2s' }} />
    </div>
    <div className="space-y-3">
      <h2 className="text-3xl font-bold text-white">Absensi Berhasil!</h2>
      <p className="text-slate-400 max-w-xs mx-auto">Data kehadiran Anda telah tercatat dengan aman di dalam sistem.</p>
    </div>
    <button onClick={reset} className="px-8 py-3 bg-white/10 hover:bg-white/20 border border-white/10 text-white rounded-xl transition-all font-medium mt-8">Kembali ke Beranda</button>
  </div>
);

const AttendanceWizard: React.FC = () => {
  const { addLog } = useAppContext();
  const [step, setStep] = useState(1);
  const [data, setData] = useState<Partial<Log>>({});

  const reset = () => { setStep(1); setData({}); };
  const steps = ['Waktu', 'Lokasi', 'Identitas', 'Verifikasi'];

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col font-sans text-slate-100 overflow-hidden relative">
      <div className="absolute top-[-10%] left-[-10%] w-[40%] h-[40%] bg-purple-600/20 rounded-full blur-[120px] pointer-events-none"></div>
      <div className="absolute bottom-[-10%] right-[-10%] w-[40%] h-[40%] bg-cyan-600/20 rounded-full blur-[120px] pointer-events-none"></div>
      <header className="w-full p-6 flex justify-between items-center relative z-10 border-b border-white/5 bg-slate-950/50 backdrop-blur-md">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-gradient-to-br from-cyan-400 to-purple-600 rounded-xl flex items-center justify-center shadow-lg overflow-hidden p-1.5">
             <img src="/axalogo.png" alt="AXAXYZ Logo" className="w-full h-full object-contain drop-shadow-sm" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
             <span className="font-bold text-white text-xl tracking-tighter hidden">A.</span>
          </div>
          <span className="font-bold text-xl tracking-widest text-transparent bg-clip-text bg-gradient-to-r from-white to-slate-400">AXAXYZ</span>
        </div>
        <div className="text-xs font-medium px-4 py-1.5 bg-white/5 border border-white/10 rounded-full text-slate-300">Portal Mahasiswa</div>
      </header>
      <main className="flex-1 flex flex-col relative z-10 w-full max-w-4xl mx-auto px-4 py-8">
        {step < 5 && (
          <div className="mb-12">
            <div className="flex justify-between relative">
              <div className="absolute top-1/2 -translate-y-1/2 left-0 w-full h-1 bg-white/10 rounded-full"></div>
              <div className="absolute top-1/2 -translate-y-1/2 left-0 h-1 bg-gradient-to-r from-cyan-500 to-purple-500 rounded-full transition-all duration-500" style={{ width: `${((step - 1) / (steps.length - 1)) * 100}%` }}></div>
              {steps.map((label, idx) => {
                const isActive = step === idx + 1; const isPassed = step > idx + 1;
                return (
                  <div key={label} className="relative z-10 flex flex-col items-center gap-2">
                    <div className={cn("w-8 h-8 rounded-full flex items-center justify-center text-sm font-bold border-2 transition-colors duration-300", isActive ? "bg-slate-900 border-cyan-400 text-cyan-400 shadow-[0_0_15px_rgba(6,182,212,0.5)]" : isPassed ? "bg-cyan-500 border-cyan-500 text-white" : "bg-slate-900 border-white/20 text-white/30")}>
                      {isPassed ? <CheckCircle2 className="w-4 h-4" /> : idx + 1}
                    </div>
                    <span className={cn("text-xs font-medium absolute -bottom-6 w-max", isActive ? "text-cyan-400" : isPassed ? "text-slate-300" : "text-slate-600")}>{label}</span>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        <div className="flex-1 flex items-center justify-center">
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
      setErr(`Sistem terkunci. Silakan coba lagi dalam ${lockoutTimer} detik.`);
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
        setErr('❌ Akses diblokir sementara (30s) karena terlalu banyak percobaan gagal.');
      } else {
        setErr(`❌ Username atau password salah. (Sisa percobaan: ${3 - newAttempts})`);
      }
    }
    setIsLoading(false);
  };

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col items-center justify-center p-4 relative overflow-hidden w-full">
      <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_center,_var(--tw-gradient-stops))] from-slate-900 via-slate-950 to-slate-950"></div>
      <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[600px] bg-cyan-500/5 rounded-full blur-[100px] pointer-events-none"></div>

      <div className="w-full max-w-md bg-white/5 backdrop-blur-2xl border border-white/10 p-8 rounded-[2rem] shadow-2xl relative z-10 animate-in slide-in-from-bottom-8 fade-in duration-700">
        <div className="flex flex-col items-center mb-8">
          <div className="relative mb-5">
            <div className="w-20 h-20 bg-gradient-to-br from-cyan-400 to-purple-600 rounded-3xl flex items-center justify-center shadow-[0_0_30px_rgba(6,182,212,0.3)] p-3 overflow-hidden">
              <img src="/axalogo.png" alt="AXAXYZ Logo" className="w-full h-full object-contain drop-shadow-md" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
              <ShieldCheck className="w-10 h-10 text-white hidden" />
            </div>
            <div className="absolute -bottom-2 -right-2 w-9 h-9 bg-slate-900 rounded-full border-2 border-slate-800 flex items-center justify-center">
              <Lock className="w-4 h-4 text-cyan-400" />
            </div>
          </div>
          <h2 className="text-2xl font-bold text-white tracking-wide">AXAXYZ Admin</h2>
          <p className="text-slate-400 text-sm mt-1">Enterprise Security Portal</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-5">
          {err && (
            <div className="p-4 bg-rose-500/10 border border-rose-500/20 text-rose-400 text-sm rounded-2xl flex items-start gap-3 animate-in shake duration-300">
              <AlertCircle className="w-5 h-5 shrink-0 mt-0.5" />
              <p className="leading-tight">{err}</p>
            </div>
          )}

          <div className="space-y-1.5">
            <label className="text-xs text-slate-400 font-semibold uppercase tracking-wider ml-1">Username</label>
            <div className="relative flex items-center bg-slate-900/50 border border-white/10 rounded-2xl overflow-hidden focus-within:border-cyan-500 focus-within:ring-1 focus-within:ring-cyan-500 transition-all">
              <div className="pl-4 pr-3 text-slate-500"><User className="w-5 h-5"/></div>
              <input type="text" value={user} onChange={e=>setUser(e.target.value)} disabled={lockoutTimer > 0 || isLoading} className="w-full bg-transparent py-3.5 pr-4 text-white outline-none placeholder-slate-600 disabled:opacity-50" placeholder="Masukkan username admin" required />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-xs text-slate-400 font-semibold uppercase tracking-wider ml-1">Password</label>
            <div className="relative flex items-center bg-slate-900/50 border border-white/10 rounded-2xl overflow-hidden focus-within:border-cyan-500 focus-within:ring-1 focus-within:ring-cyan-500 transition-all">
              <div className="pl-4 pr-3 text-slate-500"><Key className="w-5 h-5"/></div>
              <input type={showPass ? 'text' : 'password'} value={pass} onChange={e=>setPass(e.target.value)} disabled={lockoutTimer > 0 || isLoading} className="w-full bg-transparent py-3.5 pr-12 text-white outline-none placeholder-slate-600 disabled:opacity-50" placeholder="••••••••" required />
              <button type="button" onClick={() => setShowPass(!showPass)} className="absolute right-4 text-slate-400 hover:text-cyan-400 transition-colors">
                {showPass ? <EyeOff className="w-5 h-5" /> : <Eye className="w-5 h-5" />}
              </button>
            </div>
          </div>

          <button type="submit" disabled={lockoutTimer > 0 || isLoading} className="w-full py-4 mt-4 bg-gradient-to-r from-cyan-500 to-purple-600 hover:from-cyan-400 hover:to-purple-500 disabled:from-slate-700 disabled:to-slate-800 disabled:text-slate-500 disabled:cursor-not-allowed text-white font-bold rounded-2xl transition-all shadow-[0_0_20px_rgba(6,182,212,0.2)] flex justify-center items-center gap-2 group">
            {isLoading ? <Loader2 className="w-5 h-5 animate-spin" /> : (
               <>Masuk Sistem <ChevronRight className="w-5 h-5 group-hover:translate-x-1 transition-transform" /></>
            )}
          </button>
        </form>
      </div>
    </div>
  );
};

const AdminDashboardHome: React.FC = () => {
  const { logs, isCloudSync, forceManualSync } = useAppContext();
  const dbCreds = CloudStore.getCredentials();
  
  const today = new Date().toISOString().split('T')[0];
  const todayLogs = logs.filter(l => l.timestamp.startsWith(today));
  const total = todayLogs.length;
  const onTime = todayLogs.filter(l => l.status === 'Hadir').length;
  const late = todayLogs.filter(l => l.status === 'Terlambat').length;
  const sessionCounts = todayLogs.reduce((acc, log) => { acc[log.sessionName] = (acc[log.sessionName] || 0) + 1; return acc; }, {} as Record<string, number>);
  const barData = Object.entries(sessionCounts).map(([name, count]) => ({ name, Kehadiran: count }));
  const pieData = [{ name: 'Tepat Waktu', value: onTime, color: '#10b981' }, { name: 'Terlambat', value: late, color: '#f59e0b' }].filter(d => d.value > 0);

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <div className="flex justify-between items-end">
        <div>
          <h2 className="text-2xl font-bold text-white">Ringkasan Hari Ini</h2>
          <p className="text-slate-400 text-sm">{new Date().toLocaleDateString('id-ID', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}</p>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
        {[
          { title: 'Total Kehadiran', val: total, icon: Users, color: 'text-cyan-400', bg: 'bg-cyan-500/10' },
          { title: 'Tepat Waktu', val: onTime, icon: CheckCircle2, color: 'text-emerald-400', bg: 'bg-emerald-500/10' },
          { title: 'Terlambat', val: late, icon: Clock, color: 'text-amber-400', bg: 'bg-amber-500/10' }
        ].map((stat, i) => (
          <div key={i} className="bg-white/5 border border-white/10 p-6 rounded-2xl flex items-center justify-between">
            <div><p className="text-slate-400 text-sm font-medium mb-1">{stat.title}</p><h3 className="text-3xl font-bold text-white">{stat.val}</h3></div>
            <div className={cn("w-12 h-12 rounded-full flex items-center justify-center", stat.bg)}><stat.icon className={cn("w-6 h-6", stat.color)} /></div>
          </div>
        ))}
        
        <div className="bg-slate-900 border border-white/10 p-5 rounded-2xl flex flex-col items-center justify-center text-center relative overflow-hidden">
           {isCloudSync ? (
             <>
               <div className="w-10 h-10 rounded-full bg-emerald-500/20 border border-emerald-500/30 flex items-center justify-center mb-2"><Cloud className="w-5 h-5 text-emerald-400" /></div>
               <h3 className="text-white font-bold text-sm">Upstash Redis AKTIF</h3>
               <p className="text-emerald-400/80 text-[10px] mt-1 leading-tight">Database terhubung sempurna.</p>
             </>
           ) : (
             <>
               <div className="w-10 h-10 rounded-full bg-rose-500/20 border border-rose-500/30 flex items-center justify-center mb-2"><ServerCrash className="w-5 h-5 text-rose-400" /></div>
               <h3 className="text-white font-bold text-sm">Mode LOKAL Saja</h3>
               <p className="text-rose-400/80 text-[10px] mt-1 leading-tight">Gagal mendeteksi NEXT_PUBLIC ENV.</p>
             </>
           )}
        </div>
      </div>
      
      <div className="bg-slate-900 border border-indigo-500/30 p-6 rounded-2xl">
         <div className="flex items-center gap-3 mb-4">
            <DatabaseZap className="w-6 h-6 text-indigo-400" />
            <h3 className="text-lg font-bold text-white">Cloud Database Diagnostic</h3>
         </div>
         <div className="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            <div className="bg-black/50 border border-white/5 p-4 rounded-xl">
               <p className="text-xs text-slate-500 font-bold uppercase mb-1">Detected REST URL (NEXT_PUBLIC_...)</p>
               <p className="text-sm font-mono text-cyan-400 break-all">{dbCreds.url ? dbCreds.url : <span className="text-rose-500">❌ URL Tidak Terdeteksi di .env</span>}</p>
            </div>
            <div className="bg-black/50 border border-white/5 p-4 rounded-xl">
               <p className="text-xs text-slate-500 font-bold uppercase mb-1">Detected REST TOKEN</p>
               <p className="text-sm font-mono text-purple-400 break-all">{dbCreds.token ? `${dbCreds.token.substring(0, 10)}••••••••••••••••` : <span className="text-rose-500">❌ Token Tidak Terdeteksi di .env</span>}</p>
            </div>
         </div>
         <div className="flex items-center justify-between bg-indigo-500/10 p-4 rounded-xl border border-indigo-500/20">
            <p className="text-sm text-indigo-200">Jika Data di Upstash kosong namun URL terdeteksi, tekan tombol Force Sync untuk mengunggah semua data lokal Anda.</p>
            <button onClick={forceManualSync} className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white font-bold rounded-lg transition-colors flex gap-2 items-center">
              <Upload className="w-4 h-4" /> Force Sync
            </button>
         </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 h-[400px]">
        <div className="bg-white/5 border border-white/10 p-6 rounded-2xl flex flex-col">
          <h3 className="text-lg font-semibold text-white mb-6">Kehadiran per Sesi</h3>
          <div className="flex-1">
            {barData.length > 0 ? (
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={barData} margin={{ top: 0, right: 0, left: -20, bottom: 0 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#334155" vertical={false} />
                  <XAxis dataKey="name" stroke="#94a3b8" fontSize={12} tickLine={false} axisLine={false} />
                  <YAxis stroke="#94a3b8" fontSize={12} tickLine={false} axisLine={false} />
                  <Tooltip cursor={{fill: '#334155', opacity: 0.4}} contentStyle={{backgroundColor: '#0f172a', borderColor: '#334155', color: '#f8fafc'}} />
                  <Bar dataKey="Kehadiran" fill="#06b6d4" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            ) : <div className="h-full flex items-center justify-center text-slate-500">Belum ada data hari ini</div>}
          </div>
        </div>
        <div className="bg-white/5 border border-white/10 p-6 rounded-2xl flex flex-col">
          <h3 className="text-lg font-semibold text-white mb-6">Rasio Keterlambatan</h3>
          <div className="flex-1">
             {pieData.length > 0 ? (
                <ResponsiveContainer width="100%" height="100%">
                  <PieChart>
                    <Pie data={pieData} cx="50%" cy="50%" innerRadius={80} outerRadius={120} paddingAngle={5} dataKey="value">
                      {pieData.map((entry, index) => <Cell key={`cell-${index}`} fill={entry.color} />)}
                    </Pie>
                    <Tooltip contentStyle={{backgroundColor: '#0f172a', borderColor: '#334155', color: '#f8fafc'}} />
                  </PieChart>
                </ResponsiveContainer>
             ) : <div className="h-full flex items-center justify-center text-slate-500">Belum ada data hari ini</div>}
          </div>
        </div>
      </div>
    </div>
  );
};

const AdminStudents: React.FC = () => {
  const { students, addStudent, updateStudent, bulkAddStudents, deleteStudent } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newS, setNewS] = useState({ name: '', nim: '', password: '' });
  const [editingStudent, setEditingStudent] = useState<Student | null>(null);
  const [search, setSearch] = useState('');
  
  const [selectedStudentForKTM, setSelectedStudentForKTM] = useState<Student | null>(null);

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    addStudent({ ...newS, password: newS.password || `${newS.nim}123` });
    setIsAdding(false);
    setNewS({ name: '', nim: '', password: '' });
  };

  const handleUpdate = (e: React.FormEvent) => {
    e.preventDefault();
    if(editingStudent) {
       updateStudent(editingStudent.id, { name: editingStudent.name, nim: editingStudent.nim, password: editingStudent.password });
       setEditingStudent(null);
    }
  };

  const handleUnlinkDevice = (id: string, name: string) => {
     if(confirm(`Yakin ingin mereset/logout perangkat terikat untuk mahasiswa ${name}?`)) {
        updateStudent(id, { deviceId: null });
     }
  };

  const handleBulkUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const text = event.target?.result as string;
      const lines = text.split('\n');
      const newSt: Omit<Student, 'id'>[] = [];
      lines.forEach(line => {
        const parts = line.split(/[,;\t]/);
        if (parts.length >= 2) {
          const name = parts[0].trim();
          const nim = parts[1].trim();
          if (name && nim) newSt.push({ name, nim, password: `${nim}123` });
        }
      });
      if (newSt.length > 0) {
        bulkAddStudents(newSt);
        alert(`Berhasil mengimpor ${newSt.length} mahasiswa.`);
      }
    };
    reader.readAsText(file);
  };

  const filtered = students.filter(s => s.name.toLowerCase().includes(search.toLowerCase()) || s.nim.includes(search));

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col w-full relative">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-white">Data Mahasiswa</h2>
          <p className="text-slate-400 text-sm">Kelola daftar dan cetak Kartu QR Mahasiswa (KTM)</p>
        </div>
        <div className="flex gap-2">
           <label className="flex items-center gap-2 px-4 py-2 bg-purple-500/20 text-purple-400 hover:bg-purple-500/30 border border-purple-500/50 rounded-xl transition-colors font-medium cursor-pointer">
              <Upload className="w-4 h-4" /> Bulk Upload (CSV)
              <input type="file" accept=".csv, .txt" className="hidden" onChange={handleBulkUpload} />
           </label>
           <button onClick={() => setIsAdding(!isAdding)} className="flex items-center gap-2 px-4 py-2 bg-cyan-500/20 text-cyan-400 hover:bg-cyan-500/30 border border-cyan-500/50 rounded-xl transition-colors font-medium">
             <Plus className="w-4 h-4" /> Tambah Manual
           </button>
        </div>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-white/5 border border-cyan-500/30 p-6 rounded-2xl grid grid-cols-1 md:grid-cols-4 gap-4 items-end animate-in slide-in-from-top-4">
          <div className="space-y-1">
            <label className="text-xs text-slate-400">Nama Lengkap</label>
            <input required type="text" value={newS.name} onChange={e=>setNewS({...newS, name: e.target.value})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" placeholder="Contoh: Budi Santoso" />
          </div>
          <div className="space-y-1">
            <label className="text-xs text-slate-400">NIM</label>
            <input required type="text" value={newS.nim} onChange={e=>setNewS({...newS, nim: e.target.value})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" placeholder="Nomor Induk..." />
          </div>
          <div className="space-y-1">
            <label className="text-xs text-slate-400">Password (Opsional)</label>
            <input type="text" value={newS.password} onChange={e=>setNewS({...newS, password: e.target.value})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" placeholder="Default: [NIM]123" />
          </div>
          <button type="submit" className="w-full py-2.5 bg-cyan-600 hover:bg-cyan-500 text-white font-bold rounded-xl transition-all">Simpan Mhs</button>
        </form>
      )}

      <input type="text" placeholder="Cari Nama / NIM..." value={search} onChange={e=>setSearch(e.target.value)} className="w-full max-w-md bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" />

      <div className="flex-1 bg-white/5 border border-white/10 rounded-2xl overflow-hidden flex flex-col">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-white/5 border-b border-white/10 text-slate-400 text-sm">
                <th className="p-4 font-medium">NIM</th>
                <th className="p-4 font-medium">Nama Lengkap</th>
                <th className="p-4 font-medium">Password Aktif</th>
                <th className="p-4 font-medium text-center">Status Device</th>
                <th className="p-4 font-medium text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5">
              {filtered.map(st => (
                <tr key={st.id} className="hover:bg-white/5 transition-colors text-slate-200">
                  <td className="p-4 font-mono">{st.nim}</td>
                  <td className="p-4 font-medium">{st.name}</td>
                  <td className="p-4"><span className="text-xs bg-slate-800 px-2 py-1 rounded border border-white/10 font-mono text-slate-400">{st.password}</span></td>
                  <td className="p-4 text-center">
                    {st.deviceId ? (
                      <span className="inline-flex items-center gap-1.5 px-2 py-1 rounded bg-emerald-500/10 text-emerald-400 text-[10px] font-bold border border-emerald-500/20"><CheckCircle2 className="w-3 h-3"/> Tertaut</span>
                    ) : (
                      <span className="inline-flex items-center gap-1.5 px-2 py-1 rounded bg-slate-800 text-slate-400 text-[10px] font-bold border border-white/10">Belum Tertaut</span>
                    )}
                  </td>
                  <td className="p-4 text-right flex justify-end gap-2">
                    {st.deviceId && (
                      <button onClick={() => handleUnlinkDevice(st.id, st.name)} title="Reset/Logout Perangkat Tertaut" className="p-1.5 text-amber-400 hover:bg-amber-400/10 rounded-lg transition-colors border border-amber-400/20 bg-amber-400/5">
                         <RefreshCcw className="w-4 h-4" />
                      </button>
                    )}
                    <button onClick={() => setEditingStudent(st)} title="Edit Data Akun" className="p-1.5 text-blue-400 hover:bg-blue-400/10 rounded-lg transition-colors border border-blue-400/20 bg-blue-400/5">
                       <Settings className="w-4 h-4" />
                    </button>
                    <button onClick={() => setSelectedStudentForKTM(st)} title="Cetak KTM" className="px-3 py-1.5 text-xs text-cyan-400 bg-cyan-400/10 border border-cyan-400/20 hover:bg-cyan-400/20 rounded-lg transition-colors flex items-center gap-1">
                      <CreditCard className="w-3 h-3"/> KTM
                    </button>
                    <button onClick={() => deleteStudent(st.id)} title="Hapus Data" className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-rose-400/10 rounded-lg transition-colors border border-transparent hover:border-rose-400/20"><Trash2 className="w-4 h-4" /></button>
                  </td>
                </tr>
              ))}
              {filtered.length === 0 && <tr><td colSpan={5} className="p-8 text-center text-slate-500">Belum ada data mahasiswa terdaftar.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>

      {editingStudent && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/80 backdrop-blur-sm p-4 animate-in fade-in zoom-in-95 duration-200">
           <form onSubmit={handleUpdate} className="bg-slate-900 border border-white/10 p-6 rounded-3xl w-full max-w-md shadow-2xl relative">
              <div className="flex justify-between items-center mb-6">
                 <h3 className="text-xl font-bold text-white">Edit Akun Mahasiswa</h3>
                 <button type="button" onClick={() => setEditingStudent(null)} className="p-1 text-slate-400 hover:text-white transition-colors"><X className="w-5 h-5"/></button>
              </div>
              <div className="space-y-4">
                 <div className="space-y-1">
                    <label className="text-xs text-slate-400">Nama Lengkap</label>
                    <input required type="text" value={editingStudent.name} onChange={e=>setEditingStudent({...editingStudent, name: e.target.value})} className="w-full bg-slate-950 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-blue-500" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs text-slate-400">NIM</label>
                    <input required type="text" value={editingStudent.nim} onChange={e=>setEditingStudent({...editingStudent, nim: e.target.value})} className="w-full bg-slate-950 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-blue-500" />
                 </div>
                 <div className="space-y-1">
                    <label className="text-xs text-slate-400">Password</label>
                    <input required type="text" value={editingStudent.password || ''} onChange={e=>setEditingStudent({...editingStudent, password: e.target.value})} className="w-full bg-slate-950 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-blue-500" />
                 </div>
                 <button type="submit" className="w-full py-3.5 mt-2 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded-xl transition-all shadow-[0_0_15px_rgba(59,130,246,0.3)]">
                    Simpan Perubahan
                 </button>
              </div>
           </form>
        </div>
      )}

      {selectedStudentForKTM && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
          <style>{`
            @media print {
              body * { visibility: hidden; }
              #ktm-print-area, #ktm-print-area * { visibility: visible; }
              #ktm-print-area { position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); margin: 0; }
            }
          `}</style>
          <div className="bg-slate-900 border border-white/10 p-6 rounded-3xl w-[400px] shadow-2xl relative z-50">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-lg font-bold text-white">Preview Kartu KTM</h3>
              <button onClick={() => setSelectedStudentForKTM(null)} className="p-1 bg-white/10 hover:bg-white/20 rounded-full transition-colors text-white"><X className="w-5 h-5"/></button>
            </div>
            
            <div id="ktm-print-area" className="w-[340px] h-[540px] mx-auto bg-gradient-to-br from-cyan-600 to-purple-800 rounded-[2rem] p-6 relative overflow-hidden shadow-2xl flex flex-col items-center justify-between border-4 border-white/10">
               <div className="absolute top-[-50px] right-[-50px] w-48 h-48 bg-white/10 rounded-full blur-2xl"></div>
               <div className="absolute bottom-[-50px] left-[-50px] w-48 h-48 bg-black/20 rounded-full blur-2xl"></div>
               
               <div className="text-center relative z-10 w-full mt-4">
                 <div className="w-16 h-16 bg-white rounded-2xl flex items-center justify-center mx-auto mb-3 shadow-[0_10px_20px_rgba(0,0,0,0.3)] p-2 overflow-hidden">
                   <img src="/axalogo.png" alt="AXAXYZ Logo" className="w-full h-full object-contain" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
                   <span className="font-bold text-slate-900 text-3xl hidden">A.</span>
                 </div>
                 <h2 className="text-white font-black tracking-widest text-lg drop-shadow-md">AXAXYZ UNIVERSITY</h2>
                 <p className="text-cyan-200 text-[10px] tracking-[0.2em] font-bold uppercase mt-1 opacity-90">Kartu Tanda Mahasiswa</p>
               </div>

               <div className="bg-white p-3.5 rounded-2xl relative z-10 shadow-[0_10px_30px_rgba(0,0,0,0.5)]">
                  <img src={`https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${selectedStudentForKTM.nim}&margin=0`} alt="QR Code" className="w-40 h-40" />
               </div>

               <div className="text-center relative z-10 w-full bg-black/40 p-5 rounded-2xl backdrop-blur-md border border-white/10 mb-2">
                 <h1 className="text-xl font-bold text-white uppercase leading-tight mb-1">{selectedStudentForKTM.name}</h1>
                 <div className="h-px w-12 bg-cyan-500 mx-auto my-2"></div>
                 <p className="text-cyan-300 font-mono text-xl tracking-[0.1em] font-bold">{selectedStudentForKTM.nim}</p>
               </div>
            </div>

            <button onClick={() => window.print()} className="w-full mt-8 py-3.5 bg-white hover:bg-slate-200 text-slate-900 font-bold rounded-xl flex items-center justify-center gap-2 transition-colors">
              <Printer className="w-5 h-5" /> Cetak (Print) Sekarang
            </button>
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
  // UPGRADE: Location Name State
  const [locationName, setLocationName] = useState(geofence.name || 'Kampus Utama');

  const handleSave = (e: React.FormEvent) => {
    e.preventDefault();
    updateGeofence({ lat: parseFloat(lat), lng: parseFloat(lng), radius: parseInt(radius), name: locationName });
    alert('Pengaturan lokasi berhasil disimpan dan disinkronisasikan ke Server.');
  };

  const getMyLocation = () => {
    if (navigator.geolocation) {
      navigator.geolocation.getCurrentPosition(
        (pos) => { setLat(pos.coords.latitude.toString()); setLng(pos.coords.longitude.toString()); },
        () => alert('Gagal mendapatkan lokasi Admin.')
      );
    }
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 max-w-2xl">
      <div>
        <h2 className="text-2xl font-bold text-white">Pengaturan Lokasi Kampus</h2>
        <p className="text-slate-400 text-sm">Tentukan pusat koordinat absensi dan batas radius (Geofencing).</p>
      </div>

      <form onSubmit={handleSave} className="bg-white/5 border border-white/10 p-6 rounded-2xl space-y-6">
        <div className="p-4 bg-cyan-500/10 border border-cyan-500/20 rounded-xl flex items-start gap-3">
          <Navigation className="w-5 h-5 text-cyan-400 mt-0.5" />
          <p className="text-sm text-cyan-200">Mahasiswa hanya dapat melakukan absensi jika jarak GPS mereka berada di dalam <b>Radius Absensi</b> yang dihitung dari titik Latitude & Longitude di bawah ini.</p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
          <div className="space-y-1 md:col-span-2">
             <label className="text-xs text-slate-400 font-semibold uppercase">Nama Area Absensi</label>
             <input required type="text" value={locationName} onChange={e=>setLocationName(e.target.value)} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-500" placeholder="Contoh: Gedung A Kampus" />
             <p className="text-xs text-slate-500 mt-1">Nama ini akan muncul pada notifikasi error saat Mahasiswa berada di luar batas.</p>
          </div>
          <div className="space-y-1">
            <label className="text-xs text-slate-400 font-semibold uppercase">Latitude (Garis Lintang)</label>
            <input required type="number" step="any" value={lat} onChange={e=>setLat(e.target.value)} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-500 font-mono" />
          </div>
          <div className="space-y-1">
            <label className="text-xs text-slate-400 font-semibold uppercase">Longitude (Garis Bujur)</label>
            <input required type="number" step="any" value={lng} onChange={e=>setLng(e.target.value)} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-500 font-mono" />
          </div>
        </div>

        <div className="space-y-1">
           <label className="text-xs text-slate-400 font-semibold uppercase">Radius Absensi (Dalam Meter)</label>
           <input required type="number" min="10" value={radius} onChange={e=>setRadius(e.target.value)} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-3 text-white outline-none focus:border-cyan-500 font-mono" />
           <p className="text-xs text-slate-500 mt-1">Rekomendasi: 500 - 1000 meter untuk area kampus.</p>
        </div>

        <div className="flex gap-4 pt-4 border-t border-white/5">
          <button type="button" onClick={getMyLocation} className="px-6 py-3 bg-white/5 hover:bg-white/10 border border-white/10 text-white font-medium rounded-xl transition-all flex items-center gap-2">
            <MapPin className="w-4 h-4" /> Gunakan Lokasi Saya Saat Ini
          </button>
          <button type="submit" className="flex-1 py-3 bg-gradient-to-r from-cyan-500 to-purple-600 hover:from-cyan-400 hover:to-purple-500 text-white font-bold rounded-xl transition-all shadow-[0_0_15px_rgba(6,182,212,0.2)]">
            Simpan & Sync Konfigurasi
          </button>
        </div>
      </form>
    </div>
  );
};

const AdminSettings: React.FC = () => {
  const { sessions, updateSession, addSession, deleteSession } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  const [newSess, setNewSess] = useState({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault(); addSession({ ...newSess, isActive: true }); setIsAdding(false); setNewSess({ name: '', startTime: '', endTime: '', toleranceMinutes: 15 });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <div className="flex justify-between items-center">
        <div><h2 className="text-2xl font-bold text-white">Manajemen Sesi</h2><p className="text-slate-400 text-sm">Atur jadwal shift absensi.</p></div>
        <button onClick={() => setIsAdding(!isAdding)} className="flex items-center gap-2 px-4 py-2 bg-cyan-500/20 text-cyan-400 hover:bg-cyan-500/30 border border-cyan-500/50 rounded-xl transition-colors font-medium"><Plus className="w-4 h-4" /> Tambah Sesi</button>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-white/5 border border-cyan-500/30 p-6 rounded-2xl grid grid-cols-1 md:grid-cols-5 gap-4 items-end animate-in slide-in-from-top-4">
          <div className="space-y-1"><label className="text-xs text-slate-400">Nama Sesi</label><input required type="text" value={newSess.name} onChange={e=>setNewSess({...newSess, name: e.target.value})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" placeholder="e.g. Kuliah Malam" /></div>
          <div className="space-y-1"><label className="text-xs text-slate-400">Jam Mulai</label><input required type="time" value={newSess.startTime} onChange={e=>setNewSess({...newSess, startTime: e.target.value})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" /></div>
          <div className="space-y-1"><label className="text-xs text-slate-400">Jam Selesai</label><input required type="time" value={newSess.endTime} onChange={e=>setNewSess({...newSess, endTime: e.target.value})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" /></div>
          <div className="space-y-1"><label className="text-xs text-slate-400">Toleransi (Menit)</label><input required type="number" min="0" value={newSess.toleranceMinutes} onChange={e=>setNewSess({...newSess, toleranceMinutes: parseInt(e.target.value)})} className="w-full bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" /></div>
          <button type="submit" className="w-full py-2.5 bg-cyan-600 hover:bg-cyan-500 text-white font-bold rounded-xl transition-all">Simpan & Sync</button>
        </form>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
        {sessions.map(session => (
          <div key={session.id} className={cn("p-5 rounded-2xl border transition-colors", session.isActive ? "bg-white/5 border-white/10" : "bg-white/5 opacity-50 border-white/5")}>
            <div className="flex justify-between items-start mb-4">
              <h3 className="text-lg font-bold text-white">{session.name}</h3>
              <div className="flex gap-2">
                <button onClick={() => updateSession(session.id, { isActive: !session.isActive })} className={cn("px-3 py-1 text-xs font-semibold rounded-full border", session.isActive ? "bg-emerald-500/20 text-emerald-400 border-emerald-500/30" : "bg-slate-500/20 text-slate-400 border-slate-500/30")}>{session.isActive ? 'Aktif' : 'Nonaktif'}</button>
                <button onClick={() => deleteSession(session.id)} className="p-1 text-slate-400 hover:text-rose-400 transition-colors"><Trash2 className="w-4 h-4" /></button>
              </div>
            </div>
            <div className="space-y-2 text-sm text-slate-400">
              <div className="flex items-center gap-2"><Clock className="w-4 h-4"/> Waktu: <span className="text-slate-200">{session.startTime} - {session.endTime}</span></div>
              <div className="flex items-center gap-2"><Activity className="w-4 h-4"/> Toleransi: <span className="text-slate-200">{session.toleranceMinutes} menit</span></div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

const AdminReports: React.FC = () => {
  const { logs, sessions, deleteLog } = useAppContext();
  const [search, setSearch] = useState('');
  const [filterSession, setFilterSession] = useState('All');
  
  const [previewImage, setPreviewImage] = useState<string | null>(null);

  const filteredLogs = logs.filter(log => {
    const matchSearch = log.name.toLowerCase().includes(search.toLowerCase()) || log.nim.includes(search);
    const matchSession = filterSession === 'All' || log.sessionName === filterSession;
    return matchSearch && matchSession;
  });

  return (
    <div className="space-y-6 animate-in fade-in duration-500 h-full flex flex-col relative">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div><h2 className="text-2xl font-bold text-white">Laporan Kehadiran</h2><p className="text-slate-400 text-sm">Data histori absensi mahasiswa.</p></div>
        <button onClick={() => exportToCSV(filteredLogs)} className="flex items-center gap-2 px-4 py-2 bg-gradient-to-r from-emerald-500 to-emerald-600 hover:from-emerald-400 hover:to-emerald-500 text-white rounded-xl transition-all font-medium shadow-lg"><Download className="w-4 h-4" /> Export CSV</button>
      </div>

      <div className="flex flex-col md:flex-row gap-4">
        <input type="text" placeholder="Cari Nama / NIM..." value={search} onChange={e=>setSearch(e.target.value)} className="flex-1 bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500" />
        <select value={filterSession} onChange={e=>setFilterSession(e.target.value)} className="bg-slate-900/50 border border-white/10 rounded-xl px-4 py-2.5 text-white outline-none focus:border-cyan-500 w-full md:w-48 appearance-none">
          <option value="All">Semua Sesi</option>
          {sessions.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
        </select>
      </div>

      <div className="flex-1 bg-white/5 border border-white/10 rounded-2xl overflow-hidden flex flex-col">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-white/5 border-b border-white/10 text-slate-400 text-sm">
                <th className="p-4 font-medium">Profil</th>
                <th className="p-4 font-medium">NIM / Nama</th>
                <th className="p-4 font-medium">Waktu Absen</th>
                <th className="p-4 font-medium">Sesi & Status</th>
                <th className="p-4 font-medium">Lokasi Absen</th>
                <th className="p-4 font-medium text-right">Aksi</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5">
              {filteredLogs.map(log => (
                <tr key={log.id} className="hover:bg-white/5 transition-colors">
                  <td className="p-4">
                    <div onClick={() => setPreviewImage(log.photoBase64)} className="w-12 h-12 rounded-lg overflow-hidden border border-white/10 bg-black relative group cursor-pointer shadow-lg hover:border-cyan-400 transition-colors">
                      <img src={log.photoBase64} alt="Selfie" className="w-full h-full object-cover" />
                      <div className="absolute inset-0 bg-black/60 flex items-center justify-center opacity-0 group-hover:opacity-100 transition-opacity">
                        <Maximize className="w-4 h-4 text-white" />
                      </div>
                    </div>
                  </td>
                  <td className="p-4"><p className="font-semibold text-white">{log.name}</p><p className="text-xs text-slate-400 font-mono">{log.nim}</p></td>
                  <td className="p-4"><p className="text-slate-200">{new Date(log.timestamp).toLocaleTimeString('id-ID')}</p><p className="text-xs text-slate-400">{new Date(log.timestamp).toLocaleDateString('id-ID')}</p></td>
                  <td className="p-4">
                     <p className="text-slate-300 text-sm mb-1">{log.sessionName}</p>
                     <span className={cn("px-2.5 py-0.5 text-[10px] font-bold rounded-full", log.status === 'Hadir' ? "bg-emerald-500/20 text-emerald-400" : "bg-amber-500/20 text-amber-400")}>{log.status}</span>
                  </td>
                  <td className="p-4">
                    <a href={`https://www.google.com/maps?q=${log.location.lat},${log.location.lng}`} target="_blank" rel="noreferrer" className="inline-flex items-center gap-1.5 px-3 py-1.5 bg-cyan-500/10 hover:bg-cyan-500/20 text-cyan-400 text-xs font-semibold rounded-lg border border-cyan-500/20 transition-colors shadow-sm">
                      <MapPin className="w-3 h-3" /> Buka G-Maps
                    </a>
                    <p className="text-[10px] text-slate-500 mt-1.5 font-mono">{log.location.lat.toFixed(5)}, {log.location.lng.toFixed(5)}</p>
                  </td>
                  <td className="p-4 text-right">
                    <button onClick={() => { if(confirm(`Yakin ingin menghapus data absensi untuk ${log.name}?`)) deleteLog(log.id); }} title="Hapus Log Absensi" className="p-1.5 text-slate-400 hover:text-rose-400 hover:bg-rose-400/10 rounded-lg transition-colors border border-transparent hover:border-rose-400/20">
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </td>
                </tr>
              ))}
              {filteredLogs.length === 0 && <tr><td colSpan={6} className="p-8 text-center text-slate-500">Tidak ada data ditemukan.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
      
      {previewImage && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-black/90 backdrop-blur-sm p-4 animate-in fade-in zoom-in-95 duration-300" onClick={() => setPreviewImage(null)}>
          <div className="relative max-w-2xl max-h-[90vh] w-full flex flex-col items-center justify-center">
            <button onClick={() => setPreviewImage(null)} className="absolute -top-12 right-0 p-2 bg-white/10 hover:bg-rose-500/80 hover:text-white rounded-full transition-colors text-slate-300">
              <X className="w-6 h-6"/>
            </button>
            <img src={previewImage} alt="Preview Selfie Fullscreen" className="max-w-full max-h-[85vh] object-contain rounded-2xl border-4 border-white/10 shadow-2xl" onClick={e => e.stopPropagation()} />
            <p className="mt-4 text-slate-400 text-sm">Klik di luar area gambar untuk menutup</p>
          </div>
        </div>
      )}
    </div>
  );
};

const AdminLayout: React.FC<{ children: React.ReactNode, activeRoute: string, setRoute: (r:string)=>void }> = ({ children, activeRoute, setRoute }) => {
  const { syncStatus } = useAppContext();
  const handleLogout = () => { localStorage.removeItem('axaxyz_admin_auth'); setRoute('admin-login'); };

  const navItems = [
    { id: 'admin-dashboard', icon: BarChart3, label: 'Dashboard' },
    { id: 'admin-students', icon: Database, label: 'Data Mahasiswa' },
    { id: 'admin-reports', icon: FileText, label: 'Laporan Absensi' },
    { id: 'admin-geofence', icon: Map, label: 'Pengaturan Lokasi' },
    { id: 'admin-settings', icon: Settings, label: 'Pengaturan Sesi' },
  ];

  return (
    <div className="min-h-screen bg-slate-950 flex text-slate-200 font-sans w-full">
      <aside className="w-72 bg-slate-900/50 border-r border-white/10 flex flex-col backdrop-blur-xl shrink-0">
        <div className="p-6 border-b border-white/10 flex items-center gap-3">
          <div className="w-10 h-10 bg-gradient-to-br from-cyan-400 to-purple-600 rounded-xl flex items-center justify-center shadow-lg overflow-hidden p-1.5">
             <img src="/axalogo.png" alt="AXAXYZ Logo" className="w-full h-full object-contain drop-shadow-sm" onError={(e) => { e.currentTarget.style.display = 'none'; e.currentTarget.nextElementSibling?.classList.remove('hidden'); }} />
             <span className="font-bold text-white text-xl tracking-tighter hidden">A.</span>
          </div>
          <div><h1 className="font-bold text-xl tracking-widest text-transparent bg-clip-text bg-gradient-to-r from-white to-slate-400">AXAXYZ</h1><p className="text-[10px] uppercase tracking-widest text-cyan-500 font-bold">Admin Portal</p></div>
        </div>
        <nav className="flex-1 p-4 space-y-2 mt-4 overflow-y-auto">
          {navItems.map(item => (
            <button key={item.id} onClick={() => setRoute(item.id)} className={cn("w-full flex items-center gap-3 px-4 py-3 rounded-xl transition-all text-sm font-medium", activeRoute === item.id ? "bg-gradient-to-r from-cyan-500/20 to-purple-500/20 text-cyan-400 border border-cyan-500/30 shadow-[inset_0_0_20px_rgba(6,182,212,0.1)]" : "text-slate-400 hover:bg-white/5 hover:text-slate-200")}>
              <item.icon className="w-5 h-5" /> {item.label}
            </button>
          ))}
        </nav>
        <div className="p-4 border-t border-white/10">
          <button onClick={handleLogout} className="w-full flex items-center gap-3 px-4 py-3 rounded-xl text-rose-400 hover:bg-rose-500/10 transition-colors text-sm font-medium"><LogOut className="w-5 h-5" /> Keluar</button>
        </div>
      </aside>

      <main className="flex-1 relative overflow-y-auto w-full">
        <header className="absolute top-0 right-0 p-6 flex justify-end z-50 w-full pointer-events-none">
           <div className="pointer-events-auto flex items-center gap-2 px-4 py-2 bg-slate-900/80 backdrop-blur-md border border-white/10 rounded-full text-xs font-bold shadow-xl transition-all">
               {syncStatus === 'syncing' && <><RefreshCcw className="w-3.5 h-3.5 animate-spin text-cyan-400"/> <span className="text-cyan-400">Syncing to Cloud...</span></>}
               {syncStatus === 'synced' && <><CheckCircle2 className="w-3.5 h-3.5 text-emerald-400"/> <span className="text-emerald-400">Database Synced</span></>}
               {syncStatus === 'error' && <><CloudOff className="w-3.5 h-3.5 text-rose-400"/> <span className="text-rose-400">Sync Error</span></>}
               {syncStatus === 'offline' && <><CloudOff className="w-3.5 h-3.5 text-slate-500"/> <span className="text-slate-500">Local Mode Only</span></>}
           </div>
        </header>

        <div className="absolute top-[-20%] right-[-10%] w-[50%] h-[50%] bg-cyan-600/10 rounded-full blur-[150px] pointer-events-none"></div>
        <div className="absolute bottom-[-20%] left-[-10%] w-[50%] h-[50%] bg-purple-600/10 rounded-full blur-[150px] pointer-events-none"></div>
        <div className="p-8 max-w-7xl mx-auto relative z-10 min-h-full mt-10">{children}</div>
      </main>
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
      document.title = "AXAXYZ - Portal Kehadiran Pintar";
    }
  }, []);

  useEffect(() => {
    const isAdminAuthed = localStorage.getItem('axaxyz_admin_auth') === 'true';
    if (route.startsWith('admin-') && route !== 'admin-login' && !isAdminAuthed) setRoute('admin-login');
  }, [route]);

  return (
    <AppProvider>
      <div className="fixed bottom-4 right-4 z-[999] flex gap-2 bg-slate-900/80 backdrop-blur-md p-2 rounded-xl border border-white/10 shadow-2xl">
        <button onClick={() => setRoute('student')} className={cn("px-4 py-2 rounded-lg text-xs font-bold transition-colors", route === 'student' ? "bg-cyan-500 text-white" : "bg-white/10 text-slate-300 hover:bg-white/20")}>Mode Mahasiswa</button>
        <button onClick={() => setRoute(typeof window !== 'undefined' && localStorage.getItem('axaxyz_admin_auth') === 'true' ? 'admin-dashboard' : 'admin-login')} className={cn("px-4 py-2 rounded-lg text-xs font-bold transition-colors", route.startsWith('admin') ? "bg-purple-500 text-white" : "bg-white/10 text-slate-300 hover:bg-white/20")}>Mode Admin</button>
      </div>

      {route === 'student' && <AttendanceWizard />}
      {route === 'admin-login' && <AdminLogin onLogin={() => setRoute('admin-dashboard')} />}
      
      {['admin-dashboard', 'admin-students', 'admin-settings', 'admin-reports', 'admin-geofence'].includes(route) && (
        <AdminLayout activeRoute={route} setRoute={setRoute}>
          {route === 'admin-dashboard' && <AdminDashboardHome />}
          {route === 'admin-students' && <AdminStudents />}
          {route === 'admin-geofence' && <AdminGeofence />}
          {route === 'admin-settings' && <AdminSettings />}
          {route === 'admin-reports' && <AdminReports />}
        </AdminLayout>
      )}

      {/* HAPUS TANDA KOMENTAR PADA BARIS DI BAWAH INI SAAT DI-DEPLOY KE VERCEL */}
      {/<SpeedInsights />}
    </AppProvider>
  );
}
