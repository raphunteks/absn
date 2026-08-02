import React, { useState, useEffect, createContext, useContext, useRef, useMemo, useCallback } from 'react';
import { format, parse, addMinutes, isAfter, isBefore, isSameDay } from 'date-fns';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, 
  PieChart, Pie, Cell, ResponsiveContainer
} from 'recharts';
import {
  Camera, MapPin, Clock, QrCode, CheckCircle2, AlertTriangle, LayoutDashboard, 
  Settings, FileText, LogOut, Menu, X, Trash2, Power, Download, ChevronRight, User, Map
} from 'lucide-react';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- UTILS ---
export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const CAMPUS_COORDS = { lat: -6.200000, lng: 106.816666 };
const RADIUS_M = 500;

function getDistance(lat1: number, lon1: number, lat2: number, lon2: number) {
  const R = 6371e3; // Earth radius in meters
  const φ1 = lat1 * Math.PI/180;
  const φ2 = lat2 * Math.PI/180;
  const Δφ = (lat2-lat1) * Math.PI/180;
  const Δλ = (lon2-lon1) * Math.PI/180;
  const a = Math.sin(Δφ/2) * Math.sin(Δφ/2) +
            Math.cos(φ1) * Math.cos(φ2) *
            Math.sin(Δλ/2) * Math.sin(Δλ/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c; // in meters
}

function exportToCSV(data: any[], filename: string) {
  const csvRows = [];
  const headers = Object.keys(data[0] || {});
  csvRows.push(headers.join(','));
  
  for (const row of data) {
    const values = headers.map(header => {
      const val = row[header];
      const escaped = ('' + val).replace(/"/g, '\\"');
      return `"${escaped}"`;
    });
    csvRows.push(values.join(','));
  }
  
  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.setAttribute('hidden', '');
  a.setAttribute('href', url);
  a.setAttribute('download', filename);
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// --- TYPES ---
interface Session {
  id: string;
  name: string;
  startTime: string; // HH:mm
  endTime: string;   // HH:mm
  toleranceMinutes: number;
  isActive: boolean;
}

interface Log {
  id: string;
  nim: string;
  name: string;
  timestamp: string; // ISO String
  sessionName: string;
  status: 'Hadir' | 'Terlambat';
  location: { lat: number, lng: number };
  photoBase64: string;
  deviceId: string;
}

// --- CONTEXT ---
interface AppContextType {
  sessions: Session[];
  logs: Log[];
  isAdminAuthenticated: boolean;
  addLog: (log: Omit<Log, 'id' | 'timestamp'>) => void;
  addSession: (session: Omit<Session, 'id'>) => void;
  updateSession: (id: string, updates: Partial<Session>) => void;
  deleteSession: (id: string) => void;
  loginAdmin: (user: string, pass: string) => boolean;
  logoutAdmin: () => void;
}

const AppContext = createContext<AppContextType | null>(null);

const DEFAULT_SESSIONS: Session[] = [
  { id: '1', name: 'Pagi', startTime: '07:00', endTime: '09:00', toleranceMinutes: 15, isActive: true },
  { id: '2', name: 'Siang', startTime: '12:00', endTime: '13:30', toleranceMinutes: 15, isActive: true },
  { id: '3', name: 'Sore', startTime: '16:00', endTime: '17:30', toleranceMinutes: 15, isActive: true },
];

export function AppProvider({ children }: { children: React.ReactNode }) {
  const [sessions, setSessions] = useState<Session[]>([]);
  const [logs, setLogs] = useState<Log[]>([]);
  const [isAdminAuthenticated, setIsAdminAuthenticated] = useState(false);

  useEffect(() => {
    const s = localStorage.getItem('axaxyz_sessions');
    const l = localStorage.getItem('axaxyz_logs');
    const auth = localStorage.getItem('axaxyz_admin_auth');
    
    if (s) setSessions(JSON.parse(s));
    else setSessions(DEFAULT_SESSIONS);
    
    if (l) setLogs(JSON.parse(l));
    if (auth === 'true') setIsAdminAuthenticated(true);
  }, []);

  useEffect(() => {
    if (sessions.length) localStorage.setItem('axaxyz_sessions', JSON.stringify(sessions));
  }, [sessions]);

  useEffect(() => {
    localStorage.setItem('axaxyz_logs', JSON.stringify(logs));
  }, [logs]);

  const addLog = (logData: Omit<Log, 'id' | 'timestamp'>) => {
    const newLog: Log = {
      ...logData,
      id: Math.random().toString(36).substr(2, 9),
      timestamp: new Date().toISOString()
    };
    setLogs(prev => [newLog, ...prev]);
  };

  const addSession = (s: Omit<Session, 'id'>) => {
    setSessions(prev => [...prev, { ...s, id: Math.random().toString(36).substr(2, 9) }]);
  };

  const updateSession = (id: string, updates: Partial<Session>) => {
    setSessions(prev => prev.map(s => s.id === id ? { ...s, ...updates } : s));
  };

  const deleteSession = (id: string) => {
    setSessions(prev => prev.filter(s => s.id !== id));
  };

  const loginAdmin = (u: string, p: string) => {
    // Hardcoded demo credentials
    if (u === 'admin' && p === 'admin123') {
      setIsAdminAuthenticated(true);
      localStorage.setItem('axaxyz_admin_auth', 'true');
      return true;
    }
    return false;
  };

  const logoutAdmin = () => {
    setIsAdminAuthenticated(false);
    localStorage.removeItem('axaxyz_admin_auth');
  };

  return (
    <AppContext.Provider value={{ sessions, logs, isAdminAuthenticated, addLog, addSession, updateSession, deleteSession, loginAdmin, logoutAdmin }}>
      {children}
    </AppContext.Provider>
  );
}

const useAppContext = () => {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error("useAppContext must be used within AppProvider");
  return ctx;
};

// --- STUDENT COMPONENTS ---

function TimeCheck({ onNext, sessions }: { onNext: (data: any) => void, sessions: Session[] }) {
  const [error, setError] = useState<string | null>(null);
  
  const handleCheck = () => {
    const now = new Date();
    let activeSession = null;
    let status = '';

    for (const session of sessions) {
      if (!session.isActive) continue;
      const start = parse(session.startTime, 'HH:mm', now);
      const end = parse(session.endTime, 'HH:mm', now);
      const endWithTolerance = addMinutes(end, session.toleranceMinutes);

      if (now >= start && now <= end) {
        activeSession = session;
        status = 'Hadir';
        break;
      } else if (now > end && now <= endWithTolerance) {
        activeSession = session;
        status = 'Terlambat';
        break;
      }
    }

    if (activeSession) {
      onNext({ sessionName: activeSession.name, status });
    } else {
      setError("Absensi Ditutup saat ini. Silakan cek jadwal aktif.");
    }
  };

  return (
    <div className="flex flex-col items-center justify-center p-8 space-y-6 animate-in fade-in zoom-in duration-500">
      <div className="w-24 h-24 rounded-full bg-gradient-to-br from-cyan-500/20 to-purple-500/20 flex items-center justify-center border border-white/10">
        <Clock className="w-12 h-12 text-cyan-400" />
      </div>
      <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-cyan-400 to-purple-400">Validasi Waktu</h2>
        <p className="text-slate-400 max-w-xs">Sistem akan memeriksa jadwal absensi aktif saat ini.</p>
      </div>
      
      {error && (
        <div className="flex items-start space-x-3 bg-red-500/10 border border-red-500/30 p-4 rounded-xl text-red-400 text-sm max-w-sm">
          <AlertTriangle className="w-5 h-5 shrink-0" />
          <p>{error}</p>
        </div>
      )}

      <button onClick={handleCheck} className="w-full max-w-xs py-3 px-6 rounded-xl font-semibold bg-gradient-to-r from-cyan-600 to-purple-600 hover:from-cyan-500 hover:to-purple-500 transition-all shadow-[0_0_20px_rgba(6,182,212,0.3)] hover:shadow-[0_0_30px_rgba(6,182,212,0.5)]">
        Periksa Jadwal
      </button>
    </div>
  );
}

function LocationCheck({ onNext }: { onNext: (data: any) => void }) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const checkLocation = () => {
    setLoading(true);
    setError(null);

    if (!navigator.geolocation) {
      setError("Geolocation tidak didukung oleh browser Anda.");
      setLoading(false);
      return;
    }

    navigator.geolocation.getCurrentPosition(
      (position) => {
        const userLat = position.coords.latitude;
        const userLng = position.coords.longitude;
        const dist = getDistance(userLat, userLng, CAMPUS_COORDS.lat, CAMPUS_COORDS.lng);

        if (dist <= RADIUS_M) {
          onNext({ location: { lat: userLat, lng: userLng } });
        } else {
          setError(`Anda berada di luar area kampus. Jarak Anda: ${Math.round(dist)}m. Maksimal: ${RADIUS_M}m.`);
        }
        setLoading(false);
      },
      (err) => {
        setError(`Gagal mendapatkan lokasi: ${err.message}. Pastikan GPS aktif.`);
        setLoading(false);
      },
      { enableHighAccuracy: true, timeout: 10000, maximumAge: 0 }
    );
  };

  return (
    <div className="flex flex-col items-center justify-center p-8 space-y-6 animate-in fade-in zoom-in duration-500">
      <div className="w-24 h-24 rounded-full bg-gradient-to-br from-emerald-500/20 to-teal-500/20 flex items-center justify-center border border-white/10">
        <MapPin className="w-12 h-12 text-emerald-400" />
      </div>
      <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold text-white">Verifikasi Lokasi</h2>
        <p className="text-slate-400 max-w-xs">Pastikan Anda berada di dalam area Kampus ({RADIUS_M}m).</p>
      </div>

      {error && (
        <div className="flex items-start space-x-3 bg-red-500/10 border border-red-500/30 p-4 rounded-xl text-red-400 text-sm max-w-sm">
          <AlertTriangle className="w-5 h-5 shrink-0 mt-0.5" />
          <p>{error}</p>
        </div>
      )}

      <button 
        onClick={checkLocation} 
        disabled={loading}
        className="w-full max-w-xs py-3 px-6 rounded-xl font-semibold bg-white/10 hover:bg-white/20 border border-white/10 transition-all flex items-center justify-center space-x-2"
      >
        {loading ? (
          <span className="w-5 h-5 border-2 border-white/20 border-t-white rounded-full animate-spin" />
        ) : (
          <>
            <Map className="w-5 h-5" />
            <span>{error ? 'Coba Lagi' : 'Cek Lokasi Sekarang'}</span>
          </>
        )}
      </button>
    </div>
  );
}

function QRScanner({ onNext }: { onNext: (data: any) => void }) {
  const [error, setError] = useState<string | null>(null);

  return (
    <div className="flex flex-col items-center justify-center p-6 space-y-6 animate-in slide-in-from-right duration-500 w-full">
      <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold text-white flex items-center justify-center space-x-2">
          <QrCode className="w-6 h-6 text-purple-400" />
          <span>Scan ID Card</span>
        </h2>
        <p className="text-slate-400 text-sm">Posisikan QR Code pada tengah kotak.</p>
      </div>

      <div className="w-full max-w-sm rounded-2xl overflow-hidden bg-black/50 border border-white/10 p-4 shadow-2xl relative flex flex-col items-center justify-center min-h-[250px]">
         
         <QrCode className="w-16 h-16 text-slate-600 mb-4 animate-pulse" />
         <p className="text-slate-500 text-sm mb-6 text-center">Kamera Scanner<br/>(Mode Simulasi)</p>
         
         <button 
           onClick={() => {
              const dummyNim = localStorage.getItem('axaxyz_device_owner') || "10112233";
              const savedNIM = localStorage.getItem('axaxyz_device_owner');
              if (savedNIM && savedNIM !== dummyNim) {
                setError(`⚠️ FRAUD ALERT: Perangkat terdaftar untuk NIM ${savedNIM}.`);
                return;
              }
              if (!savedNIM) localStorage.setItem('axaxyz_device_owner', dummyNim);
              onNext({ nim: dummyNim, name: "Mahasiswa Demo", deviceId: "demo-device" });
           }}
           className="bg-purple-600/80 hover:bg-purple-500 px-4 py-2 rounded-lg text-white font-medium shadow-lg transition-all"
         >
           Simulasi Scan (Dev Mode)
         </button>
      </div>

      {error && (
        <div className="flex items-start space-x-3 bg-red-500/10 border border-red-500/30 p-4 rounded-xl text-red-400 text-sm max-w-sm w-full">
          <AlertTriangle className="w-5 h-5 shrink-0 mt-0.5" />
          <p className="leading-tight">{error}</p>
        </div>
      )}
    </div>
  );
}

function SelfieCapture({ onNext }: { onNext: (data: any) => void }) {
  const videoRef = useRef<HTMLVideoElement>(null);
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const [imgSrc, setImgSrc] = useState<string | null>(null);

  useEffect(() => {
    let stream: MediaStream | null = null;
    if (!imgSrc) {
      navigator.mediaDevices.getUserMedia({ video: { facingMode: "user" } })
        .then(s => {
          stream = s;
          if (videoRef.current) {
            videoRef.current.srcObject = s;
          }
        })
        .catch(err => console.error("Camera error:", err));
    }
    return () => {
      if (stream) {
        stream.getTracks().forEach(t => t.stop());
      }
    };
  }, [imgSrc]);

  const capture = useCallback(() => {
    if (videoRef.current && canvasRef.current) {
      const video = videoRef.current;
      const canvas = canvasRef.current;
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
        setImgSrc(canvas.toDataURL('image/jpeg'));
      }
    }
  }, []);

  return (
    <div className="flex flex-col items-center justify-center p-6 space-y-6 animate-in slide-in-from-right duration-500 w-full">
       <div className="text-center space-y-2">
        <h2 className="text-2xl font-bold text-white flex items-center justify-center space-x-2">
          <Camera className="w-6 h-6 text-pink-400" />
          <span>Selfie Kehadiran</span>
        </h2>
        <p className="text-slate-400 text-sm">Ambil foto selfie di area kampus.</p>
      </div>

      <div className="w-full max-w-sm rounded-2xl overflow-hidden bg-black border border-white/10 shadow-2xl relative min-h-[300px] flex items-center justify-center">
        {imgSrc ? (
          <img src={imgSrc} alt="Selfie preview" className="w-full h-auto aspect-[3/4] object-cover" />
        ) : (
          <>
            <video ref={videoRef} autoPlay playsInline className="w-full h-auto object-cover aspect-[3/4]" />
            <canvas ref={canvasRef} className="hidden" />
          </>
        )}
      </div>

      <div className="flex w-full max-w-sm space-x-4">
        {imgSrc ? (
          <>
            <button onClick={() => setImgSrc(null)} className="flex-1 py-3 px-4 rounded-xl font-semibold bg-white/5 hover:bg-white/10 border border-white/10 transition-all text-slate-300">
              Ulangi
            </button>
            <button onClick={() => onNext({ photoBase64: imgSrc })} className="flex-1 py-3 px-4 rounded-xl font-semibold bg-gradient-to-r from-cyan-600 to-purple-600 hover:from-cyan-500 hover:to-purple-500 transition-all shadow-[0_0_20px_rgba(6,182,212,0.3)]">
              Konfirmasi
            </button>
          </>
        ) : (
          <button onClick={capture} className="w-full py-4 rounded-xl font-semibold bg-white text-slate-900 hover:bg-slate-200 transition-all flex justify-center items-center space-x-2">
            <Camera className="w-5 h-5" />
            <span>Ambil Foto</span>
          </button>
        )}
      </div>
    </div>
  );
}

function AttendanceWizard({ onComplete }: { onComplete: () => void }) {
  const { sessions, addLog } = useAppContext();
  const [step, setStep] = useState(1);
  const [wizardData, setWizardData] = useState<Partial<Log>>({});

  const handleNext = (data: Partial<Log>) => {
    const updatedData = { ...wizardData, ...data };
    setWizardData(updatedData);
    
    if (step === 4) {
      // Final Submission
      addLog(updatedData as Omit<Log, 'id' | 'timestamp'>);
      setStep(5);
    } else {
      setStep(step + 1);
    }
  };

  return (
    <div className="w-full max-w-md mx-auto min-h-[600px] bg-white/5 backdrop-blur-xl border border-white/10 shadow-2xl rounded-3xl overflow-hidden relative flex flex-col">
      {/* Header Progress */}
      {step < 5 && (
        <div className="p-6 border-b border-white/5 flex items-center justify-between">
          <div className="flex space-x-2">
            {[1, 2, 3, 4].map(i => (
              <div key={i} className={cn("h-1.5 rounded-full transition-all duration-300", 
                i === step ? "w-8 bg-cyan-400 shadow-[0_0_10px_rgba(34,211,238,0.8)]" : 
                i < step ? "w-4 bg-purple-500/50" : "w-4 bg-white/10"
              )} />
            ))}
          </div>
          <span className="text-xs font-medium text-slate-500 uppercase tracking-wider">Langkah {step}/4</span>
        </div>
      )}

      {/* Content Area */}
      <div className="flex-1 flex items-center justify-center p-4">
        {step === 1 && <TimeCheck onNext={handleNext} sessions={sessions} />}
        {step === 2 && <LocationCheck onNext={handleNext} />}
        {step === 3 && <QRScanner onNext={handleNext} />}
        {step === 4 && <SelfieCapture onNext={handleNext} />}
        {step === 5 && (
          <div className="flex flex-col items-center justify-center space-y-6 text-center animate-in zoom-in duration-500 p-8">
            <div className="w-24 h-24 bg-emerald-500/20 rounded-full flex items-center justify-center border border-emerald-500/30">
              <CheckCircle2 className="w-12 h-12 text-emerald-400" />
            </div>
            <div className="space-y-2">
              <h2 className="text-3xl font-bold text-white">Berhasil!</h2>
              <p className="text-slate-400">Absensi Anda telah tercatat ke dalam sistem.</p>
            </div>
            <div className="bg-black/30 w-full p-4 rounded-xl border border-white/5 text-left space-y-2 text-sm text-slate-300">
              <div className="flex justify-between"><span className="text-slate-500">NIM:</span> <span>{wizardData.nim}</span></div>
              <div className="flex justify-between"><span className="text-slate-500">Nama:</span> <span>{wizardData.name}</span></div>
              <div className="flex justify-between"><span className="text-slate-500">Sesi:</span> <span>{wizardData.sessionName}</span></div>
              <div className="flex justify-between"><span className="text-slate-500">Status:</span> 
                <span className={wizardData.status === 'Hadir' ? 'text-emerald-400' : 'text-amber-400 font-bold'}>{wizardData.status}</span>
              </div>
            </div>
            <button onClick={onComplete} className="w-full py-3 rounded-xl font-semibold bg-white/10 hover:bg-white/20 transition-all text-white border border-white/10">
              Kembali ke Awal
            </button>
          </div>
        )}
      </div>
    </div>
  );
}


// --- ADMIN COMPONENTS ---

function AdminLogin({ onNavigate }: { onNavigate: (path: string) => void }) {
  const { loginAdmin } = useAppContext();
  const [user, setUser] = useState('');
  const [pass, setPass] = useState('');
  const [err, setErr] = useState('');

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (loginAdmin(user, pass)) {
      onNavigate('admin-dashboard');
    } else {
      setErr('Kredensial tidak valid.');
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4">
      <div className="w-full max-w-sm bg-white/5 backdrop-blur-xl border border-white/10 p-8 rounded-3xl shadow-2xl space-y-8">
        <div className="text-center space-y-2">
          <div className="inline-flex h-12 w-12 bg-gradient-to-br from-cyan-500 to-purple-600 rounded-xl items-center justify-center shadow-lg mb-2">
            <Settings className="w-6 h-6 text-white" />
          </div>
          <h1 className="text-2xl font-bold text-white">Portal Admin</h1>
          <p className="text-slate-400 text-sm">Akses dashboard AXAXYZ</p>
        </div>

        <form onSubmit={handleLogin} className="space-y-4">
          {err && <div className="text-red-400 text-sm bg-red-500/10 p-3 rounded-lg border border-red-500/20 text-center">{err}</div>}
          <div className="space-y-1">
            <label className="text-xs font-medium text-slate-400 uppercase tracking-wider ml-1">Username</label>
            <input 
              type="text" value={user} onChange={e => setUser(e.target.value)} required
              className="w-full bg-black/40 border border-white/10 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-cyan-500 focus:ring-1 focus:ring-cyan-500 transition-all"
              placeholder="admin"
            />
          </div>
          <div className="space-y-1">
            <label className="text-xs font-medium text-slate-400 uppercase tracking-wider ml-1">Password</label>
            <input 
              type="password" value={pass} onChange={e => setPass(e.target.value)} required
              className="w-full bg-black/40 border border-white/10 rounded-xl px-4 py-3 text-white focus:outline-none focus:border-cyan-500 focus:ring-1 focus:ring-cyan-500 transition-all"
              placeholder="••••••••"
            />
          </div>
          <button type="submit" className="w-full py-3 mt-4 rounded-xl font-semibold bg-gradient-to-r from-cyan-600 to-purple-600 hover:from-cyan-500 hover:to-purple-500 text-white shadow-lg transition-all">
            Masuk ke Sistem
          </button>
        </form>
        
        <button onClick={() => onNavigate('student')} className="w-full text-sm text-slate-500 hover:text-white transition-colors">
          &larr; Kembali ke Mode Mahasiswa
        </button>
      </div>
    </div>
  );
}

function AdminLayout({ children, currentPath, onNavigate }: { children: React.ReactNode, currentPath: string, onNavigate: (path: string) => void }) {
  const { logoutAdmin } = useAppContext();
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

  const navItems = [
    { path: 'admin-dashboard', icon: LayoutDashboard, label: 'Dashboard' },
    { path: 'admin-reports', icon: FileText, label: 'Laporan & Ekspor' },
    { path: 'admin-settings', icon: Settings, label: 'Pengaturan Sesi' },
  ];

  const handleLogout = () => {
    logoutAdmin();
    onNavigate('admin-login');
  };

  return (
    <div className="min-h-screen bg-slate-950 flex flex-col md:flex-row text-slate-200">
      {/* Mobile Header */}
      <div className="md:hidden flex items-center justify-between p-4 bg-white/5 border-b border-white/10">
        <div className="font-bold text-xl tracking-wider bg-clip-text text-transparent bg-gradient-to-r from-cyan-400 to-purple-500">
          AXAXYZ
        </div>
        <button onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)} className="p-2 bg-white/10 rounded-lg">
          {isMobileMenuOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
        </button>
      </div>

      {/* Sidebar */}
      <div className={cn(
        "fixed md:static inset-y-0 left-0 z-50 w-72 bg-slate-950 md:bg-white/5 backdrop-blur-2xl border-r border-white/10 transform transition-transform duration-300 flex flex-col",
        isMobileMenuOpen ? "translate-x-0" : "-translate-x-full md:translate-x-0"
      )}>
        <div className="p-6 hidden md:block">
          <div className="font-black text-3xl tracking-widest bg-clip-text text-transparent bg-gradient-to-r from-cyan-400 to-purple-500">
            AXAXYZ
          </div>
          <p className="text-xs text-slate-500 mt-1 uppercase tracking-wider font-semibold">Enterprise Attendance</p>
        </div>

        <nav className="flex-1 px-4 py-8 md:py-4 space-y-2">
          {navItems.map((item) => {
            const Icon = item.icon;
            const isActive = currentPath === item.path;
            return (
              <button
                key={item.path}
                onClick={() => { onNavigate(item.path); setIsMobileMenuOpen(false); }}
                className={cn(
                  "w-full flex items-center space-x-3 px-4 py-3 rounded-xl transition-all text-sm font-medium",
                  isActive 
                    ? "bg-gradient-to-r from-cyan-500/20 to-purple-500/20 border border-white/10 text-white shadow-[0_0_15px_rgba(6,182,212,0.15)]" 
                    : "text-slate-400 hover:text-white hover:bg-white/5"
                )}
              >
                <Icon className={cn("w-5 h-5", isActive ? "text-cyan-400" : "")} />
                <span>{item.label}</span>
              </button>
            )
          })}
        </nav>

        <div className="p-4 border-t border-white/10">
           <button onClick={handleLogout} className="w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-red-400 hover:bg-red-500/10 hover:text-red-300 transition-all text-sm font-medium">
             <LogOut className="w-5 h-5" />
             <span>Keluar Sistem</span>
           </button>
        </div>
      </div>

      {/* Main Content */}
      <div className="flex-1 overflow-auto bg-slate-950 p-4 md:p-8 relative">
         <div className="max-w-6xl mx-auto">
            {children}
         </div>
      </div>
    </div>
  );
}

function AdminDashboard() {
  const { logs, sessions } = useAppContext();
  
  const today = new Date();
  const todayLogs = logs.filter(log => isSameDay(new Date(log.timestamp), today));
  
  const totalToday = todayLogs.length;
  const onTimeCount = todayLogs.filter(l => l.status === 'Hadir').length;
  const lateCount = todayLogs.filter(l => l.status === 'Terlambat').length;

  const sessionData = sessions.map(s => {
    return {
      name: s.name,
      Hadir: todayLogs.filter(l => l.sessionName === s.name && l.status === 'Hadir').length,
      Terlambat: todayLogs.filter(l => l.sessionName === s.name && l.status === 'Terlambat').length,
    }
  });

  const pieData = [
    { name: 'Tepat Waktu', value: onTimeCount, color: '#34d399' },
    { name: 'Terlambat', value: lateCount, color: '#fbbf24' },
  ];

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <h2 className="text-3xl font-bold text-white">Ikhtisar Hari Ini</h2>
      
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {[
          { label: 'Total Kehadiran', value: totalToday, icon: User, color: 'text-cyan-400', bg: 'bg-cyan-400/10' },
          { label: 'Tepat Waktu (Hadir)', value: onTimeCount, icon: CheckCircle2, color: 'text-emerald-400', bg: 'bg-emerald-400/10' },
          { label: 'Terlambat', value: lateCount, icon: Clock, color: 'text-amber-400', bg: 'bg-amber-400/10' },
        ].map((stat, i) => (
          <div key={i} className="bg-white/5 border border-white/10 rounded-2xl p-6 flex items-center space-x-4">
            <div className={cn("p-4 rounded-xl", stat.bg)}>
              <stat.icon className={cn("w-6 h-6", stat.color)} />
            </div>
            <div>
              <p className="text-slate-400 text-sm font-medium">{stat.label}</p>
              <h3 className="text-3xl font-bold text-white mt-1">{stat.value}</h3>
            </div>
          </div>
        ))}
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <div className="lg:col-span-2 bg-white/5 border border-white/10 rounded-2xl p-6">
           <h3 className="text-lg font-semibold text-white mb-6">Kehadiran per Sesi (Hari ini)</h3>
           <div className="h-72 w-full">
             <ResponsiveContainer width="100%" height="100%">
               <BarChart data={sessionData}>
                 <CartesianGrid strokeDasharray="3 3" stroke="#334155" vertical={false} />
                 <XAxis dataKey="name" stroke="#94a3b8" tick={{ fill: '#94a3b8' }} />
                 <YAxis stroke="#94a3b8" tick={{ fill: '#94a3b8' }} allowDecimals={false} />
                 <RechartsTooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '8px', color: '#fff' }} />
                 <Legend wrapperStyle={{ paddingTop: '20px' }} />
                 <Bar dataKey="Hadir" fill="#34d399" radius={[4, 4, 0, 0]} />
                 <Bar dataKey="Terlambat" fill="#fbbf24" radius={[4, 4, 0, 0]} />
               </BarChart>
             </ResponsiveContainer>
           </div>
        </div>
        
        <div className="bg-white/5 border border-white/10 rounded-2xl p-6">
           <h3 className="text-lg font-semibold text-white mb-6">Rasio Keterlambatan</h3>
           <div className="h-64 w-full">
             {totalToday === 0 ? (
               <div className="w-full h-full flex items-center justify-center text-slate-500">Belum ada data</div>
             ) : (
               <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={pieData} cx="50%" cy="50%" innerRadius={60} outerRadius={80} paddingAngle={5} dataKey="value">
                    {pieData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.color} stroke="transparent" />
                    ))}
                  </Pie>
                  <RechartsTooltip contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '8px' }} itemStyle={{ color: '#fff' }} />
                  <Legend verticalAlign="bottom" height={36} />
                </PieChart>
              </ResponsiveContainer>
             )}
           </div>
        </div>
      </div>
    </div>
  );
}

function AdminSettings() {
  const { sessions, addSession, updateSession, deleteSession } = useAppContext();
  const [isAdding, setIsAdding] = useState(false);
  
  const [newSess, setNewSess] = useState({ name: '', startTime: '07:00', endTime: '09:00', toleranceMinutes: 15 });

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    addSession({ ...newSess, isActive: true });
    setIsAdding(false);
    setNewSess({ name: '', startTime: '07:00', endTime: '09:00', toleranceMinutes: 15 });
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500 max-w-4xl">
      <div className="flex justify-between items-center">
        <h2 className="text-3xl font-bold text-white">Pengaturan Sesi</h2>
        <button onClick={() => setIsAdding(!isAdding)} className="px-4 py-2 bg-white/10 hover:bg-white/20 border border-white/20 rounded-lg text-sm font-medium transition-all">
          {isAdding ? 'Batal' : '+ Tambah Sesi Baru'}
        </button>
      </div>

      {isAdding && (
        <form onSubmit={handleAdd} className="bg-white/5 border border-white/10 rounded-2xl p-6 grid grid-cols-1 md:grid-cols-5 gap-4 items-end animate-in slide-in-from-top-4">
          <div className="space-y-2">
            <label className="text-xs text-slate-400 uppercase">Nama Sesi</label>
            <input type="text" required value={newSess.name} onChange={e => setNewSess({...newSess, name: e.target.value})} className="w-full bg-black/40 border border-white/10 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-cyan-500" placeholder="e.g. Kuliah Malam" />
          </div>
          <div className="space-y-2">
            <label className="text-xs text-slate-400 uppercase">Mulai (HH:mm)</label>
            <input type="time" required value={newSess.startTime} onChange={e => setNewSess({...newSess, startTime: e.target.value})} className="w-full bg-black/40 border border-white/10 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-cyan-500 [&::-webkit-calendar-picker-indicator]:filter-[invert(1)]" />
          </div>
          <div className="space-y-2">
            <label className="text-xs text-slate-400 uppercase">Selesai (HH:mm)</label>
            <input type="time" required value={newSess.endTime} onChange={e => setNewSess({...newSess, endTime: e.target.value})} className="w-full bg-black/40 border border-white/10 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-cyan-500 [&::-webkit-calendar-picker-indicator]:filter-[invert(1)]" />
          </div>
          <div className="space-y-2">
            <label className="text-xs text-slate-400 uppercase">Toleransi (Menit)</label>
            <input type="number" required min="0" value={newSess.toleranceMinutes} onChange={e => setNewSess({...newSess, toleranceMinutes: parseInt(e.target.value)})} className="w-full bg-black/40 border border-white/10 rounded-lg px-3 py-2 text-white focus:outline-none focus:border-cyan-500" />
          </div>
          <button type="submit" className="py-2.5 px-4 bg-cyan-600 hover:bg-cyan-500 text-white rounded-lg font-medium shadow-lg transition-all">Simpan</button>
        </form>
      )}

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {sessions.map(s => (
          <div key={s.id} className={cn("relative p-6 rounded-2xl border transition-all duration-300", s.isActive ? "bg-white/5 border-white/10 shadow-lg" : "bg-black/20 border-white/5 opacity-60")}>
            <div className="flex justify-between items-start mb-4">
              <h3 className="text-xl font-bold text-white">{s.name}</h3>
              <button 
                onClick={() => updateSession(s.id, { isActive: !s.isActive })}
                className={cn("w-12 h-6 rounded-full flex items-center p-1 cursor-pointer transition-colors", s.isActive ? "bg-emerald-500" : "bg-slate-700")}
              >
                <div className={cn("bg-white w-4 h-4 rounded-full shadow-md transform transition-transform", s.isActive ? "translate-x-6" : "")} />
              </button>
            </div>
            
            <div className="space-y-2 text-sm text-slate-300">
              <div className="flex items-center space-x-2"><Clock className="w-4 h-4 text-slate-500" /> <span>{s.startTime} - {s.endTime}</span></div>
              <div className="flex items-center space-x-2"><AlertTriangle className="w-4 h-4 text-slate-500" /> <span>Toleransi: {s.toleranceMinutes} menit</span></div>
            </div>

            <button onClick={() => deleteSession(s.id)} className="absolute bottom-4 right-4 p-2 text-slate-500 hover:text-red-400 hover:bg-red-500/10 rounded-lg transition-all">
              <Trash2 className="w-4 h-4" />
            </button>
          </div>
        ))}
      </div>
    </div>
  );
}

function AdminReports() {
  const { logs, sessions } = useAppContext();
  const [filterSession, setFilterSession] = useState('All');
  const [search, setSearch] = useState('');

  const filteredLogs = useMemo(() => {
    return logs.filter(log => {
      const matchSess = filterSession === 'All' || log.sessionName === filterSession;
      const matchSearch = log.nim.includes(search) || log.name.toLowerCase().includes(search.toLowerCase());
      return matchSess && matchSearch;
    });
  }, [logs, filterSession, search]);

  const handleExport = () => {
    if (!filteredLogs.length) return alert("Tidak ada data untuk diekspor");
    const data = filteredLogs.map(l => ({
      NIM: l.nim,
      Nama: l.name,
      Tanggal: format(new Date(l.timestamp), 'yyyy-MM-dd'),
      Waktu: format(new Date(l.timestamp), 'HH:mm:ss'),
      Sesi: l.sessionName,
      Status: l.status,
      Latitude: l.location.lat,
      Longitude: l.location.lng,
      'Device Fingerprint': l.deviceId
    }));
    exportToCSV(data, `Laporan_Kehadiran_AXAXYZ_${format(new Date(), 'yyyyMMdd_HHmm')}.csv`);
  };

  return (
    <div className="space-y-6 animate-in fade-in duration-500">
      <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <h2 className="text-3xl font-bold text-white">Laporan Kehadiran</h2>
        <button onClick={handleExport} className="flex items-center space-x-2 px-4 py-2 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 rounded-lg text-white font-medium shadow-[0_0_15px_rgba(16,185,129,0.3)] transition-all">
          <Download className="w-4 h-4" />
          <span>Ekspor CSV</span>
        </button>
      </div>

      <div className="flex flex-col sm:flex-row gap-4 bg-white/5 p-4 rounded-2xl border border-white/10">
         <input 
            type="text" placeholder="Cari NIM atau Nama..." value={search} onChange={e => setSearch(e.target.value)}
            className="flex-1 bg-black/40 border border-white/10 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-cyan-500" 
         />
         <select 
            value={filterSession} onChange={e => setFilterSession(e.target.value)}
            className="bg-black/40 border border-white/10 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-cyan-500 min-w-[200px]"
         >
            <option value="All">Semua Sesi</option>
            {sessions.map(s => <option key={s.id} value={s.name}>{s.name}</option>)}
         </select>
      </div>

      <div className="bg-white/5 border border-white/10 rounded-2xl overflow-hidden shadow-xl">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-sm whitespace-nowrap">
            <thead className="bg-black/40 text-slate-400 uppercase text-xs font-semibold">
              <tr>
                <th className="px-6 py-4">Foto</th>
                <th className="px-6 py-4">Mahasiswa</th>
                <th className="px-6 py-4">Waktu</th>
                <th className="px-6 py-4">Sesi</th>
                <th className="px-6 py-4">Lokasi (Lat, Lng)</th>
                <th className="px-6 py-4">Status</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-white/5">
              {filteredLogs.length === 0 ? (
                <tr><td colSpan={6} className="px-6 py-8 text-center text-slate-500">Data tidak ditemukan.</td></tr>
              ) : (
                filteredLogs.map(log => (
                  <tr key={log.id} className="hover:bg-white/[0.02] transition-colors">
                    <td className="px-6 py-3">
                       <div className="w-10 h-10 rounded-lg overflow-hidden border border-white/20 relative group">
                          {log.photoBase64 ? (
                            <>
                              <img src={log.photoBase64} alt="Selfie" className="w-full h-full object-cover" />
                              {/* Hover Image Enlarge */}
                              <div className="hidden group-hover:block absolute top-0 left-12 w-32 h-40 bg-black border border-white/20 rounded-lg z-50 shadow-2xl overflow-hidden">
                                <img src={log.photoBase64} alt="Selfie Zoom" className="w-full h-full object-cover" />
                              </div>
                            </>
                          ) : (
                            <div className="w-full h-full bg-slate-800 flex items-center justify-center"><User className="w-5 h-5 text-slate-500"/></div>
                          )}
                       </div>
                    </td>
                    <td className="px-6 py-3">
                      <div className="font-semibold text-white">{log.name}</div>
                      <div className="text-xs text-slate-500 font-mono">{log.nim}</div>
                    </td>
                    <td className="px-6 py-3">
                      <div className="text-white">{format(new Date(log.timestamp), 'dd MMM yyyy')}</div>
                      <div className="text-xs text-slate-400">{format(new Date(log.timestamp), 'HH:mm:ss')}</div>
                    </td>
                    <td className="px-6 py-3 text-slate-300">{log.sessionName}</td>
                    <td className="px-6 py-3 text-xs text-slate-400 font-mono">
                      {log.location.lat.toFixed(5)}, {log.location.lng.toFixed(5)}
                    </td>
                    <td className="px-6 py-3">
                      <span className={cn(
                        "px-2.5 py-1 rounded-full text-xs font-semibold border",
                        log.status === 'Hadir' ? "bg-emerald-500/10 text-emerald-400 border-emerald-500/20" : "bg-amber-500/10 text-amber-400 border-amber-500/20"
                      )}>
                        {log.status}
                      </span>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}


// --- MAIN APP ROUTER (Simulating Next.js App Router) ---

function AppContent() {
  const { isAdminAuthenticated } = useAppContext();
  const [currentPath, setCurrentPath] = useState('student'); // 'student', 'admin-login', 'admin-dashboard', etc.

  // Protected route enforcement
  useEffect(() => {
    if (currentPath.startsWith('admin-') && currentPath !== 'admin-login' && !isAdminAuthenticated) {
      setCurrentPath('admin-login');
    }
  }, [currentPath, isAdminAuthenticated]);

  return (
    <div className="min-h-screen bg-slate-950 font-sans selection:bg-cyan-500/30">
      {currentPath === 'student' && (
        <div className="min-h-screen flex flex-col relative overflow-hidden">
          {/* Decorative Background Elements */}
          <div className="absolute top-0 inset-x-0 h-[500px] bg-gradient-to-b from-cyan-900/20 to-transparent pointer-events-none" />
          <div className="absolute -top-40 -right-40 w-96 h-96 bg-purple-600/20 rounded-full blur-[100px] pointer-events-none" />
          <div className="absolute top-40 -left-40 w-96 h-96 bg-cyan-600/20 rounded-full blur-[100px] pointer-events-none" />

          {/* Student Header */}
          <header className="relative z-10 w-full p-6 flex justify-between items-center max-w-7xl mx-auto">
            <div className="flex items-center space-x-3">
              <div className="w-10 h-10 bg-gradient-to-br from-cyan-400 to-purple-600 rounded-xl flex items-center justify-center shadow-[0_0_20px_rgba(6,182,212,0.4)]">
                <span className="font-black text-white text-lg tracking-tighter">AX</span>
              </div>
              <span className="font-bold text-xl text-white tracking-widest">AXAXYZ</span>
            </div>
            <button onClick={() => setCurrentPath(isAdminAuthenticated ? 'admin-dashboard' : 'admin-login')} className="text-sm font-medium text-slate-400 hover:text-white transition-colors flex items-center space-x-1 border border-transparent hover:border-white/10 px-3 py-1.5 rounded-full hover:bg-white/5">
              <span>Admin Portal</span>
              <ChevronRight className="w-4 h-4" />
            </button>
          </header>

          <main className="flex-1 relative z-10 flex flex-col items-center justify-center p-4">
             <div className="w-full max-w-md text-center mb-8 space-y-2">
                <h1 className="text-4xl md:text-5xl font-black text-white tracking-tight">Presensi <span className="bg-clip-text text-transparent bg-gradient-to-r from-cyan-400 to-purple-400">Pintar</span></h1>
                <p className="text-slate-400 text-sm md:text-base">Gunakan identitas dan lokasi yang valid untuk mencatat kehadiran Anda.</p>
             </div>
             <AttendanceWizard onComplete={() => setCurrentPath('student')} />
          </main>
          
          <footer className="relative z-10 text-center p-6 text-xs text-slate-600">
            &copy; {new Date().getFullYear()} AXAXYZ Enterprise Systems. Built for strict compliance.
          </footer>
        </div>
      )}

      {currentPath === 'admin-login' && <AdminLogin onNavigate={setCurrentPath} />}
      
      {currentPath.startsWith('admin-') && currentPath !== 'admin-login' && (
        <AdminLayout currentPath={currentPath} onNavigate={setCurrentPath}>
          {currentPath === 'admin-dashboard' && <AdminDashboard />}
          {currentPath === 'admin-settings' && <AdminSettings />}
          {currentPath === 'admin-reports' && <AdminReports />}
        </AdminLayout>
      )}
    </div>
  );
}

export default function App() {
  return (
    <AppProvider>
      <AppContent />
    </AppProvider>
  );
}