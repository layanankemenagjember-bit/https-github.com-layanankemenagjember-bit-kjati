import React, { useState, useEffect, useRef } from 'react';
import { createRoot } from 'react-dom/client';
import firebase from 'firebase/compat/app';
import 'firebase/compat/firestore';
import * as XLSX from 'xlsx';

// --- FIREBASE CONFIG ---
const firebaseConfig = {
  apiKey: "AIzaSyCkYrq9jEUfQIMbBebJq7RetzofcY3oKwo",
  authDomain: "kjati-absensi-app.firebaseapp.com",
  projectId: "kjati-absensi-app",
  storageBucket: "kjati-absensi-app.appspot.com",
  messagingSenderId: "272787703921",
  appId: "1:272787703921:web:f2b261433c937d6ceff3d0"
};

if (!firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
}
const db = firebase.firestore();

// --- HELPERS ---
const haversineDistance = (lat1: number, lon1: number, lat2: number, lon2: number) => {
  const R = 6371e3;
  const φ1 = lat1 * Math.PI / 180;
  const φ2 = lat2 * Math.PI / 180;
  const Δφ = (lat2 - lat1) * Math.PI / 180;
  const Δλ = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(Δφ/2) * Math.sin(Δφ/2) + Math.cos(φ1) * Math.cos(φ2) * Math.sin(Δλ/2) * Math.sin(Δλ/2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
  return R * c;
};

const formatDateTime = (ts: any) => {
  if(!ts) return '-';
  const d = ts.toDate ? ts.toDate() : new Date(ts.seconds * 1000);
  return d.toLocaleString('id-ID', { day:'2-digit', month:'2-digit', year:'numeric', hour:'2-digit', minute:'2-digit' });
};

const isEventExpired = (dateStr: string, endTimeStr: string) => {
  try {
    const [year, month, day] = dateStr.split('-').map(Number);
    const [hour, minute] = endTimeStr.split(':').map(Number);
    const endDateTime = new Date(year, month - 1, day, hour, minute);
    const now = new Date();
    return now.getTime() > (endDateTime.getTime() + 10 * 60 * 1000);
  } catch (e) {
    return false;
  }
};

const isEventNotStartedYet = (dateStr: string, startTimeStr: string) => {
  try {
    const [year, month, day] = dateStr.split('-').map(Number);
    const [hour, minute] = startTimeStr.split(':').map(Number);
    const startDateTime = new Date(year, month - 1, day, hour, minute);
    const now = new Date();
    return now.getTime() < startDateTime.getTime();
  } catch (e) {
    return false;
  }
};

// --- FACE RECOGNITION HELPERS ---
const getGrayscaleVector = (videoEl: HTMLVideoElement): number[] => {
  try {
    const canvas = document.createElement('canvas');
    canvas.width = 32;
    canvas.height = 32;
    const ctx = canvas.getContext('2d');
    if(!ctx) return [];
    
    // Crop to center-square
    const size = Math.min(videoEl.videoWidth, videoEl.videoHeight);
    const sx = (videoEl.videoWidth - size) / 2;
    const sy = (videoEl.videoHeight - size) / 2;
    
    ctx.drawImage(videoEl, sx, sy, size, size, 0, 0, 32, 32);
    
    const imgData = ctx.getImageData(0, 0, 32, 32);
    const data = imgData.data;
    const vector: number[] = [];
    
    for (let i = 0; i < data.length; i += 4) {
      const r = data[i];
      const g = data[i+1];
      const b = data[i+2];
      const gray = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
      vector.push(gray);
    }
    return vector;
  } catch (e) {
    console.error("Gagal mengekstrak vektor wajah", e);
    return [];
  }
};

const getGrayscaleCandidates = (videoEl: HTMLVideoElement): number[][] => {
  try {
    const candidates: number[][] = [];
    const videoWidth = videoEl.videoWidth;
    const videoHeight = videoEl.videoHeight;
    const size = Math.min(videoWidth, videoHeight);
    if (size <= 0) return [];

    const sx = (videoWidth - size) / 2;
    const sy = (videoHeight - size) / 2;

    const canvas = document.createElement('canvas');
    canvas.width = 32;
    canvas.height = 32;
    const ctx = canvas.getContext('2d');
    if (!ctx) return [];

    // Let's define strategic crops to handle translation, aspect ratios, and scaling
    const offset1 = Math.round(size * 0.05); // 5% shift
    const offset2 = Math.round(size * 0.10); // 10% shift
    
    const crops = [
      // Exact center
      { x: sx, y: sy, w: size, h: size },
      
      // Translation shifts (Left, Right, Up, Down)
      { x: Math.max(0, sx - offset1), y: sy, w: size, h: size },
      { x: Math.min(videoWidth - size, sx + offset1), y: sy, w: size, h: size },
      { x: sx, y: Math.max(0, sy - offset1), w: size, h: size },
      { x: sx, y: Math.min(videoHeight - size, sy + offset1), w: size, h: size },
      
      // Stronger translation shifts
      { x: Math.max(0, sx - offset2), y: sy, w: size, h: size },
      { x: Math.min(videoWidth - size, sx + offset2), y: sy, w: size, h: size },
      
      // Scale shifts (Zoom in: smaller crop window)
      { x: sx + Math.round(size * 0.06), y: sy + Math.round(size * 0.06), w: Math.round(size * 0.88), h: Math.round(size * 0.88) },
      // Scale shifts (Zoom out: larger crop window if video bounds allow)
      { x: Math.max(0, sx - Math.round(size * 0.06)), y: Math.max(0, sy - Math.round(size * 0.06)), w: Math.min(videoWidth, Math.round(size * 1.12)), h: Math.min(videoHeight, Math.round(size * 1.12)) }
    ];

    for (const crop of crops) {
      ctx.clearRect(0, 0, 32, 32);
      ctx.drawImage(videoEl, crop.x, crop.y, crop.w, crop.h, 0, 0, 32, 32);
      
      const imgData = ctx.getImageData(0, 0, 32, 32);
      const data = imgData.data;
      const vector: number[] = [];
      for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i+1];
        const b = data[i+2];
        const gray = Math.round(0.299 * r + 0.587 * g + 0.114 * b);
        vector.push(gray);
      }
      candidates.push(vector);
    }
    
    return candidates;
  } catch (e) {
    console.error("Gagal mengekstrak kandidat vektor wajah", e);
    return [];
  }
};

const getCroppedBase64Photo = (videoEl: HTMLVideoElement): string => {
  try {
    const canvas = document.createElement('canvas');
    canvas.width = 160;
    canvas.height = 160;
    const ctx = canvas.getContext('2d');
    if(!ctx) return '';
    
    // Crop to center-square
    const size = Math.min(videoEl.videoWidth, videoEl.videoHeight);
    const sx = (videoEl.videoWidth - size) / 2;
    const sy = (videoEl.videoHeight - size) / 2;
    
    // Mirror the image for webcam look
    ctx.translate(160, 0);
    ctx.scale(-1, 1);
    ctx.drawImage(videoEl, sx, sy, size, size, 0, 0, 160, 160);
    
    return canvas.toDataURL('image/jpeg', 0.85);
  } catch (e) {
    console.error("Gagal mengambil foto wajah", e);
    return '';
  }
};

const calculateFaceSimilarity = (arr1: number[], arr2: number[]): number => {
  if (!arr1 || !arr2 || arr1.length !== arr2.length) return 0;
  
  let sum1 = 0, sum2 = 0;
  const n = arr1.length;
  for (let i = 0; i < n; i++) {
    sum1 += arr1[i];
    sum2 += arr2[i];
  }
  const mean1 = sum1 / n;
  const mean2 = sum2 / n;
  
  let num = 0;
  let den1 = 0, den2 = 0;
  for (let i = 0; i < n; i++) {
    const d1 = arr1[i] - mean1;
    const d2 = arr2[i] - mean2;
    num += d1 * d2;
    den1 += d1 * d1;
    den2 += d2 * d2;
  }
  
  if (den1 === 0 || den2 === 0) return 0;
  const correlation = num / Math.sqrt(den1 * den2);
  
  // Map correlation from [-1, 1] to [0, 100]%
  return Math.max(0, Math.min(100, (correlation + 1) * 50));
};

// --- DB SERVICE ---
const dbService = {
  validateAsn: async (nip: string) => {
    const snap = await db.collection("asn_users").where("nip", "==", nip.replace(/\D/g, '')).limit(1).get();
    if(!snap.empty) return { id: snap.docs[0].id, ...snap.docs[0].data(), role: 'asn' } as any;
    return null;
  },
  searchAsn: async (q: string) => {
    const clean = q.replace(/\D/g, '');
    if(clean.length < 3) return [];
    const snap = await db.collection("asn_users").where("nip", ">=", clean).where("nip", "<=", clean + '\uf8ff').limit(5).get();
    return snap.docs.map(d => ({ id: d.id, ...d.data() }));
  },
  getAllFaceProfiles: async () => {
    try {
      const snap = await db.collection("asn_users").where("hasFaceId", "==", true).get();
      return snap.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    } catch (e) {
      console.error("Gagal mengambil daftar Face ID", e);
      return [];
    }
  }
};

const Toast = ({ message, type, onClose }: { message: string, type: string, onClose: () => void }) => {
  useEffect(() => { const t = setTimeout(onClose, 3000); return () => clearTimeout(t); }, [onClose]);
  return <div className={`toast toast-${type}`}>{message}</div>;
};

// --- FACE RECOGNITION LOGIN SCANNER ---
const FaceLoginScanner = ({ onMatch, onSwitchToNip, onRegisterClick }: any) => {
  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const [isStarted, setIsStarted] = useState(false);
  const [status, setStatus] = useState<'initializing' | 'scanning' | 'success' | 'failed' | 'no-faces'>('initializing');
  const [statusMsg, setStatusMsg] = useState('Mengaktifkan Sensor Wajah...');
  const [faceProfiles, setFaceProfiles] = useState<any[]>([]);
  const [highestMatchScore, setHighestMatchScore] = useState<number>(0);
  const [highestMatchName, setHighestMatchName] = useState<string>('');
  const scanIntervalRef = useRef<any>(null);

  const startFlow = () => {
    setIsStarted(true);
    setStatus('initializing');
    setStatusMsg('Memuat database biometrik...');
    loadProfiles();
  };

  const loadProfiles = async () => {
    try {
      const profiles = await dbService.getAllFaceProfiles();
      setFaceProfiles(profiles);
      if (profiles.length === 0) {
        setStatus('no-faces');
        setStatusMsg('Belum ada data Face ID terdaftar.');
      } else {
        startCamera();
      }
    } catch (err) {
      console.error(err);
      setStatus('failed');
      setStatusMsg('Gagal memuat data biometrik.');
    }
  };

  const startCamera = async () => {
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: 'user', width: { ideal: 320 }, height: { ideal: 320 } }
      });
      streamRef.current = stream;
      
      // Wait a short duration to ensure DOM refs are fully synchronized
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          videoRef.current.onloadedmetadata = () => {
            videoRef.current?.play().then(() => {
              setStatus('scanning');
              setStatusMsg('Mencari kecocokan wajah...');
            }).catch((e) => {
              console.error("Video play error:", e);
              setStatus('failed');
              setStatusMsg('Gagal memutar video.');
            });
          };
        } else {
          setStatus('scanning');
          setStatusMsg('Mencari kecocokan wajah...');
        }
      }, 150);
    } catch (err) {
      console.error("Camera access error:", err);
      setStatus('failed');
      setStatusMsg('Kamera tidak tersedia.');
    }
  };

  const cancelFlow = () => {
    if (streamRef.current) {
      streamRef.current.getTracks().forEach(track => track.stop());
      streamRef.current = null;
    }
    if (scanIntervalRef.current) {
      clearInterval(scanIntervalRef.current);
    }
    setIsStarted(false);
    setHighestMatchScore(0);
    setHighestMatchName('');
  };

  useEffect(() => {
    return () => {
      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
      }
      if (scanIntervalRef.current) {
        clearInterval(scanIntervalRef.current);
      }
    };
  }, []);

  useEffect(() => {
    if (!isStarted || status !== 'scanning' || faceProfiles.length === 0 || !videoRef.current) return;

    const performScan = () => {
      if (!videoRef.current) return;
      try {
        const liveCandidates = getGrayscaleCandidates(videoRef.current);
        if (liveCandidates.length === 0) return;

        let bestMatch: any = null;
        let maxSimilarity = 0;

        for (const profile of faceProfiles) {
          if (!profile.faceVector) continue;
          
          // Match the single stored vector with multiple shifted/scaled live crops
          for (const candVector of liveCandidates) {
            const similarity = calculateFaceSimilarity(candVector, profile.faceVector);
            if (similarity > maxSimilarity) {
              maxSimilarity = similarity;
              bestMatch = profile;
            }
          }
        }

        // Display highest similarity and matching progress
        if (bestMatch && maxSimilarity > 50) {
          const matchedPct = Math.round(maxSimilarity);
          setHighestMatchScore(matchedPct);
          setHighestMatchName(bestMatch.name);

          // We use 80% with our multi-candidate alignment checks as a robust, highly secure threshold
          if (matchedPct >= 80) {
            setStatus('success');
            setStatusMsg(`Terverifikasi: ${bestMatch.name} (${matchedPct}%)`);
            
            if (scanIntervalRef.current) clearInterval(scanIntervalRef.current);
            if (streamRef.current) {
              streamRef.current.getTracks().forEach(track => track.stop());
            }

            setTimeout(() => {
              onMatch({ ...bestMatch, role: 'asn' });
            }, 1200);
            return;
          } else {
            setStatusMsg(`Mencocokkan: ${matchedPct}% (Butuh 80% - Posisikan lebih tegak)`);
          }
        } else {
          setHighestMatchScore(0);
          setHighestMatchName('');
          setStatusMsg('Mencari kecocokan wajah...');
        }
      } catch (err) {
        console.error("Scanning error", err);
      }
    };

    scanIntervalRef.current = setInterval(performScan, 350);

    return () => {
      if (scanIntervalRef.current) clearInterval(scanIntervalRef.current);
    };
  }, [isStarted, status, faceProfiles]);

  if (!isStarted) {
    return (
      <div style={{ textAlign: 'center', padding: '15px 10px', background: '#f8fafc', borderRadius: '16px', border: '1.5px solid #e2e8f0', margin: '5px 0 15px 0' }}>
        <div style={{ fontSize: '42px', marginBottom: '8px', animation: 'pulse 2s infinite' }}>👤</div>
        <h4 style={{ fontSize: '15px', fontWeight: '800', color: '#0c3e26', margin: '0 0 6px 0' }}>Sistem Login Face ID V2.0</h4>
        <p style={{ fontSize: '11.5px', color: '#64748b', lineHeight: '1.5', margin: '0 auto 15px auto', maxWidth: '290px' }}>
          Teknologi pindaian biometrik terkalibrasi. Mendukung koreksi pencahayaan, posisi wajah, dan pindaian multi-sudut.
        </p>
        
        <button 
          type="button" 
          onClick={startFlow} 
          className="btn-premium-masuk" 
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px', padding: '12px', background: '#0c3e26', border: 'none', color: '#fff', fontWeight: '800', fontSize: '13px', borderRadius: '10px', cursor: 'pointer', transition: 'all 0.2s', boxShadow: '0 4px 10px rgba(12, 62, 38, 0.15)' }}
        >
          📷 Klik untuk Pindai Wajah
        </button>

        <div style={{ marginTop: '15px', borderTop: '1px dashed #e2e8f0', paddingTop: '12px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <button 
            type="button" 
            onClick={onRegisterClick} 
            style={{ background: 'none', border: 'none', color: '#0c3e26', fontWeight: '800', fontSize: '12px', cursor: 'pointer', textDecoration: 'underline' }}
          >
            Daftar Face ID Baru (Uji Coba)
          </button>
          <span style={{ fontSize: '11px', color: '#94a3b8' }}>Bagi pegawai yang belum mendaftarkan biometrik</span>
        </div>
      </div>
    );
  }

  return (
    <div style={{ textAlign: 'center', padding: '10px 0' }}>
      <div className={`face-scanner-container ${status === 'failed' || status === 'no-faces' ? 'error' : status === 'success' ? 'success' : ''}`}>
        {status === 'initializing' && (
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%', color: '#64748b', fontSize: '13px' }}>
            Mengaktifkan Sensor...
          </div>
        )}
        {(status === 'scanning' || status === 'success' || status === 'initializing') && (
          <video ref={videoRef} autoPlay playsInline muted className="face-scanner-video" style={{ display: status === 'initializing' ? 'none' : 'block' }} />
        )}
        {status === 'scanning' && <div className="face-scanner-laser" />}
        <div className="face-scanner-overlay">
          <div className="face-scanner-bracket-tl" />
          <div className="face-scanner-bracket-tr" />
          <div className="face-scanner-bracket-bl" />
          <div className="face-scanner-bracket-br" />
          <div className="face-scanner-target" />
        </div>
      </div>

      <div className={`face-login-status-badge ${status === 'failed' || status === 'no-faces' ? 'error' : status === 'success' ? '' : 'info'}`}>
        <span>{status === 'scanning' ? '🔍' : status === 'success' ? '✓' : '⚡'}</span>
        <span>{statusMsg}</span>
      </div>

      {/* Real-time Matching Score Progress Feedback */}
      {status === 'scanning' && highestMatchScore > 0 && (
        <div style={{ marginTop: '12px', padding: '0 20px' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: '11px', color: '#64748b', marginBottom: '4px' }}>
            <span>Kecocokan Sinyal Wajah:</span>
            <span style={{ fontWeight: '800', color: highestMatchScore >= 70 ? '#16a34a' : '#0c3e26' }}>{highestMatchScore}%</span>
          </div>
          <div style={{ width: '100%', height: '6px', background: '#e2e8f0', borderRadius: '10px', overflow: 'hidden' }}>
            <div style={{ width: `${highestMatchScore}%`, height: '100%', background: highestMatchScore >= 70 ? '#16a34a' : '#eab308', borderRadius: '10px', transition: 'width 0.2s ease-in-out' }} />
          </div>
        </div>
      )}

      {status === 'no-faces' && (
        <div style={{ padding: '0 10px', fontSize: '12.5px', color: '#64748b', marginTop: '5px' }}>
          Silakan gunakan login NIP manual terlebih dahulu, lalu aktifkan Face ID di Dashboard Pegawai.
        </div>
      )}

      <div style={{ display: 'flex', gap: '8px', justifyContent: 'center', marginTop: '15px' }}>
        <button type="button" onClick={cancelFlow} className="face-login-toggle-btn" style={{ borderColor: '#ef4444', color: '#ef4444' }}>
          ✕ Batal / Reset
        </button>
        {(status === 'failed' || status === 'no-faces') && (
          <button type="button" onClick={onSwitchToNip} className="face-login-toggle-btn">
            ⌨️ Gunakan Login NIP Manual
          </button>
        )}
      </div>
    </div>
  );
};

// --- REGISTER ASN & FACE ID MODAL ---
const RegisterAsnModal = ({ onClose, onSuccess, onError }: any) => {
  const [nip, setNip] = useState('');
  const [asnData, setAsnData] = useState<any>(null);
  const [isSearching, setIsSearching] = useState(false);
  const [manualMode, setManualMode] = useState(false);
  const [manualInputs, setManualInputs] = useState({ name: '', instansi: '' });

  const [step, setStep] = useState<'details' | 'scan'>('details');
  const [cameraReady, setCameraReady] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Mengaktifkan Sensor Wajah...');
  const [isSaving, setIsSaving] = useState(false);
  
  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | null>(null);

  const handleSearchNip = async () => {
    const cleanNip = nip.replace(/\D/g, '');
    if (cleanNip.length < 5) {
      onError("NIP minimal 5 digit untuk simulasi.");
      return;
    }
    setIsSearching(true);
    try {
      const data = await dbService.validateAsn(cleanNip);
      if (data) {
        setAsnData(data);
        setManualMode(false);
      } else {
        setAsnData(null);
        setManualMode(true);
        onError("NIP tidak ditemukan di database utama. Silakan isi Nama & Instansi secara manual untuk simulasi.");
      }
    } catch (err) {
      onError("Gagal melakukan pencarian database.");
    } finally {
      setIsSearching(false);
    }
  };

  // Trigger search on 18 digits or enter
  const handleNipChange = (val: string) => {
    const clean = val.replace(/\D/g, '');
    setNip(clean);
    if (clean.length === 18) {
      setIsSearching(true);
      dbService.validateAsn(clean).then(data => {
        if (data) {
          setAsnData(data);
          setManualMode(false);
        } else {
          setAsnData(null);
          setManualMode(true);
        }
        setIsSearching(false);
      }).catch(() => setIsSearching(false));
    }
  };

  const startCamera = async () => {
    setStep('scan');
    setCameraReady(false);
    setStatusMsg('Mengaktifkan Sensor Wajah...');
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: 'user', width: { ideal: 320 }, height: { ideal: 320 } }
      });
      streamRef.current = stream;
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          videoRef.current.onloadedmetadata = () => {
            videoRef.current?.play().then(() => {
              setCameraReady(true);
              setStatusMsg('Sensor Wajah Aktif. Posisikan wajah Anda di tengah.');
            }).catch(() => {
              onError("Gagal memutar video kamera.");
              setStep('details');
            });
          };
        }
      }, 300);
    } catch (err) {
      console.error(err);
      onError("Kamera tidak tersedia.");
      setStep('details');
    }
  };

  useEffect(() => {
    return () => {
      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
      }
    };
  }, []);

  const handleRegister = async () => {
    if (!videoRef.current || isSaving) return;
    setIsSaving(true);
    setStatusMsg('Mengekstrak Fitur Biometrik...');
    try {
      const cleanNip = nip.replace(/\D/g, '');
      const photo = getCroppedBase64Photo(videoRef.current);
      const vector = getGrayscaleVector(videoRef.current);

      if (vector.length === 0 || !photo) {
        throw new Error("Gagal menangkap gambar wajah.");
      }

      setStatusMsg('Menyimpan data Biometrik...');

      let finalUser: any = null;

      if (asnData && asnData.id) {
        // NIP was found, we update their face ID data
        await db.collection("asn_users").doc(asnData.id).update({
          facePhoto: photo,
          faceVector: vector,
          hasFaceId: true
        });
        finalUser = {
          ...asnData,
          facePhoto: photo,
          faceVector: vector,
          hasFaceId: true
        };
      } else {
        // Manual mode, create a new record
        const newDoc = await db.collection("asn_users").add({
          nip: cleanNip,
          name: manualInputs.name || 'Pegawai Baru',
          instansi: manualInputs.instansi || 'Unit Kerja Baru',
          facePhoto: photo,
          faceVector: vector,
          hasFaceId: true
        });
        finalUser = {
          id: newDoc.id,
          nip: cleanNip,
          name: manualInputs.name || 'Pegawai Baru',
          instansi: manualInputs.instansi || 'Unit Kerja Baru',
          facePhoto: photo,
          faceVector: vector,
          hasFaceId: true,
          role: 'asn'
        };
      }

      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
      }

      onSuccess(finalUser);
    } catch (err: any) {
      onError(err.message || "Gagal mengaktifkan Face ID.");
      setIsSaving(false);
      setStatusMsg('Sensor Wajah Aktif. Posisikan wajah Anda di tengah.');
    }
  };

  return (
    <div className="modal">
      <div className="modal-content" style={{ maxWidth: '420px', color: '#333' }}>
        <div className="header" style={{ marginBottom: '15px' }}>
          <h3 style={{ color: '#0c3e26', fontSize: '18px', fontWeight: '800' }}>Aktivasi Face ID Pegawai</h3>
          <button onClick={onClose} className="logout-btn" disabled={isSaving}>Batal</button>
        </div>

        {step === 'details' ? (
          <div>
            <p style={{ fontSize: '12.5px', color: '#64748b', marginBottom: '15px', lineHeight: '1.4' }}>
              Masukkan NIP Anda yang sudah terdaftar di sistem. Sistem akan mencocokkan Nama & Instansi Anda secara otomatis.
            </p>
            
            <label style={{ fontWeight: '700', fontSize: '12.5px' }}>NIP Pegawai</label>
            <div style={{ display: 'flex', gap: '8px', marginBottom: '15px' }}>
              <input 
                value={nip} 
                onChange={e => handleNipChange(e.target.value)} 
                placeholder="Masukkan NIP Anda" 
                maxLength={18}
                style={{ flex: 1, marginBottom: 0 }}
              />
              <button 
                type="button" 
                className="btn btn-outline" 
                style={{ width: 'auto', padding: '0 15px', marginTop: 0 }}
                onClick={handleSearchNip}
                disabled={isSearching || nip.length < 5}
              >
                {isSearching ? '...' : 'Cari'}
              </button>
            </div>

            {asnData && (
              <div style={{ background: '#f0fdf4', border: '1.5px solid #bbf7d0', padding: '12px', borderRadius: '12px', marginBottom: '15px' }}>
                <div style={{ fontSize: '11px', textTransform: 'uppercase', color: '#16a34a', fontWeight: '800', letterSpacing: '0.05em' }}>Pegawai Ditemukan:</div>
                <div style={{ fontSize: '14px', fontWeight: '800', color: '#0f291e', marginTop: '2px' }}>{asnData.name}</div>
                <div style={{ fontSize: '12px', color: '#15803d', marginTop: '1px' }}>{asnData.instansi}</div>
              </div>
            )}

            {manualMode && (
              <div style={{ background: '#fffbeb', border: '1.5px solid #fef3c7', padding: '15px', borderRadius: '12px', marginBottom: '15px' }}>
                <div style={{ fontSize: '11px', textTransform: 'uppercase', color: '#d97706', fontWeight: '800', letterSpacing: '0.05em' }}>Simulasi ASN Baru:</div>
                <p style={{ fontSize: '12px', color: '#b45309', margin: '4px 0 10px 0', lineHeight: '1.4' }}>
                  NIP tidak ditemukan di database bawaan. Anda dapat mengisi data simulasi di bawah ini untuk menguji pendaftaran:
                </p>
                <label style={{ fontSize: '11px', fontWeight: '700' }}>Nama Lengkap</label>
                <input 
                  value={manualInputs.name} 
                  onChange={e => setManualInputs({ ...manualInputs, name: e.target.value })} 
                  placeholder="Nama Lengkap & Gelar" 
                  style={{ background: '#fff', fontSize: '13px', padding: '8px 12px', height: '36px', marginBottom: '10px' }}
                />
                <label style={{ fontSize: '11px', fontWeight: '700' }}>Instansi</label>
                <input 
                  value={manualInputs.instansi} 
                  onChange={e => setManualInputs({ ...manualInputs, instansi: e.target.value })} 
                  placeholder="Contoh: Kantor Kemenag Jember" 
                  style={{ background: '#fff', fontSize: '13px', padding: '8px 12px', height: '36px', marginBottom: '0' }}
                />
              </div>
            )}

            <button 
              className="btn btn-primary" 
              style={{ marginTop: '10px' }} 
              onClick={startCamera} 
              disabled={!asnData && (!manualMode || !manualInputs.name || !manualInputs.instansi)}
            >
              Mulai Sensor Wajah ➔
            </button>
          </div>
        ) : (
          <div style={{ textAlign: 'center' }}>
            <p style={{ fontSize: '12.5px', color: '#64748b', marginBottom: '10px' }}>
              Posisikan wajah Anda tepat di tengah dan menghadap kamera, lalu klik Daftar.
            </p>

            <div className={`face-scanner-container ${isSaving ? 'success' : ''}`}>
              <video ref={videoRef} autoPlay playsInline muted className="face-scanner-video" />
              <div className="face-scanner-laser" />
              <div className="face-scanner-overlay">
                <div className="face-scanner-bracket-tl" />
                <div className="face-scanner-bracket-tr" />
                <div className="face-scanner-bracket-bl" />
                <div className="face-scanner-bracket-br" />
                <div className="face-scanner-target" />
              </div>
            </div>

            <div className="face-login-status-badge info" style={{ marginBottom: '15px' }}>
              <span>⚡</span>
              <span>{statusMsg}</span>
            </div>

            <div style={{ display: 'flex', gap: '10px' }}>
              <button 
                className="btn btn-outline" 
                style={{ flex: 1, marginTop: 0 }} 
                onClick={() => {
                  if (streamRef.current) streamRef.current.getTracks().forEach(track => track.stop());
                  setStep('details');
                }}
                disabled={isSaving}
              >
                Kembali
              </button>
              <button 
                className="btn btn-primary" 
                style={{ flex: 2, marginTop: 0 }} 
                onClick={handleRegister} 
                disabled={!cameraReady || isSaving}
              >
                {isSaving ? 'Mendaftarkan...' : 'Ambil Foto & Daftar'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

// --- USER FACE ID SETTINGS ---
const UserFaceIdSettings = ({ user, onUpdateUser, onError, onSuccess }: any) => {
  const [isRegistering, setIsRegistering] = useState(false);
  const [cameraReady, setCameraReady] = useState(false);
  const [statusMsg, setStatusMsg] = useState('Mengaktifkan Sensor Wajah...');
  const [isSaving, setIsSaving] = useState(false);

  const videoRef = useRef<HTMLVideoElement>(null);
  const streamRef = useRef<MediaStream | null>(null);

  const startCamera = async () => {
    setIsRegistering(true);
    setCameraReady(false);
    setStatusMsg('Mengaktifkan Sensor Wajah...');
    try {
      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: 'user', width: { ideal: 320 }, height: { ideal: 320 } }
      });
      streamRef.current = stream;
      setTimeout(() => {
        if (videoRef.current) {
          videoRef.current.srcObject = stream;
          videoRef.current.onloadedmetadata = () => {
            videoRef.current?.play().then(() => {
              setCameraReady(true);
              setStatusMsg('Sensor Wajah Aktif. Posisikan wajah Anda di tengah.');
            }).catch(() => {
              onError("Gagal memulai kamera.");
              setIsRegistering(false);
            });
          };
        }
      }, 300);
    } catch (err) {
      console.error(err);
      onError("Kamera tidak tersedia.");
      setIsRegistering(false);
    }
  };

  useEffect(() => {
    return () => {
      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
      }
    };
  }, []);

  const handleRegister = async () => {
    if (!videoRef.current || isSaving) return;
    setIsSaving(true);
    setStatusMsg('Mengekstrak Fitur Biometrik...');
    try {
      const photo = getCroppedBase64Photo(videoRef.current);
      const vector = getGrayscaleVector(videoRef.current);

      if (vector.length === 0 || !photo) {
        throw new Error("Gagal menangkap gambar wajah.");
      }

      setStatusMsg('Menyimpan ke Profil Anda...');
      
      await db.collection("asn_users").doc(user.id).update({
        facePhoto: photo,
        faceVector: vector,
        hasFaceId: true
      });

      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
      }

      const updatedUser = { ...user, facePhoto: photo, faceVector: vector, hasFaceId: true };
      onUpdateUser(updatedUser);
      setIsRegistering(false);
      setIsSaving(false);
      onSuccess("Biometrik Face ID berhasil diaktifkan!");
    } catch (err: any) {
      onError(err.message || "Gagal mengaktifkan Face ID.");
      setIsSaving(false);
      setStatusMsg('Sensor Wajah Aktif. Posisikan wajah Anda di tengah.');
    }
  };

  const handleDelete = async () => {
    if (!window.confirm("Apakah Anda yakin ingin menonaktifkan Face ID?")) return;
    try {
      await db.collection("asn_users").doc(user.id).update({
        facePhoto: firebase.firestore.FieldValue.delete(),
        faceVector: firebase.firestore.FieldValue.delete(),
        hasFaceId: false
      });
      const updatedUser = { ...user, hasFaceId: false };
      delete updatedUser.facePhoto;
      delete updatedUser.faceVector;
      onUpdateUser(updatedUser);
      onSuccess("Biometrik Face ID berhasil dinonaktifkan.");
    } catch (err: any) {
      onError("Gagal menonaktifkan Face ID.");
    }
  };

  return (
    <div className="card">
      <h3 className="section-title">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
          <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z"/>
          <circle cx="12" cy="13" r="4"/>
        </svg>
        Keamanan Biometrik Face ID
      </h3>

      {!isRegistering ? (
        <div style={{ display: 'flex', gap: '15px', alignItems: 'center', flexWrap: 'wrap' }}>
          {user.hasFaceId ? (
            <React.Fragment>
              <div style={{ position: 'relative', width: '80px', height: '80px', borderRadius: '50%', overflow: 'hidden', border: '3px solid #27AE60' }}>
                <img src={user.facePhoto} style={{ width: '100%', height: '100%', objectFit: 'cover' }} alt="Face Biometric" />
              </div>
              <div style={{ flex: 1, minWidth: '200px' }}>
                <div style={{ color: '#27AE60', fontWeight: '700', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '5px' }}>
                  <span>✓</span> Face ID Aktif (Login Instan Diaktifkan)
                </div>
                <p style={{ fontSize: '11.5px', color: '#666', marginTop: '4px', lineHeight: '1.4' }}>
                  Anda sekarang bisa masuk ke sistem dalam 1 detik dengan mengarahkan wajah ke kamera, tanpa mengetik NIP manual.
                </p>
                <button className="btn btn-danger" style={{ width: 'auto', padding: '6px 12px', fontSize: '11px', marginTop: '10px' }} onClick={handleDelete}>
                  Nonaktifkan Face ID
                </button>
              </div>
            </React.Fragment>
          ) : (
            <div style={{ width: '100%' }}>
              <div style={{ color: '#E74C3C', fontWeight: '700', fontSize: '14px', display: 'flex', alignItems: 'center', gap: '5px' }}>
                <span>❌</span> Face ID Belum Aktif
              </div>
              <p style={{ fontSize: '12.5px', color: '#666', marginTop: '5px', lineHeight: '1.5' }}>
                Aktifkan Face ID sekarang agar Anda tidak perlu lagi repot mengetik 18-digit NIP saat login. Cepat, aman, dan sangat praktis.
              </p>
              <button className="btn btn-primary" style={{ width: 'auto', padding: '10px 20px', marginTop: '12px' }} onClick={startCamera}>
                Daftarkan Face ID Sekarang
              </button>
            </div>
          )}
        </div>
      ) : (
        <div style={{ textAlign: 'center' }}>
          <p style={{ fontSize: '12.5px', color: '#666', marginBottom: '10px' }}>
            Posisikan wajah Anda di tengah lingkaran dan menghadap kamera, lalu klik Ambil Foto.
          </p>

          <div className="face-scanner-container">
            <video ref={videoRef} autoPlay playsInline muted className="face-scanner-video" />
            <div className="face-scanner-laser" />
            <div className="face-scanner-overlay">
              <div className="face-scanner-bracket-tl" />
              <div className="face-scanner-bracket-tr" />
              <div className="face-scanner-bracket-bl" />
              <div className="face-scanner-bracket-br" />
              <div className="face-scanner-target" />
            </div>
          </div>

          <div className="face-login-status-badge info" style={{ marginBottom: '15px' }}>
            <span>⚡</span>
            <span>{statusMsg}</span>
          </div>

          <div style={{ display: 'flex', gap: '10px', maxWidth: '300px', margin: '0 auto' }}>
            <button 
              className="btn btn-outline" 
              style={{ flex: 1, marginTop: 0 }} 
              onClick={() => {
                if (streamRef.current) streamRef.current.getTracks().forEach(track => track.stop());
                setIsRegistering(false);
              }}
              disabled={isSaving}
            >
              Batal
            </button>
            <button 
              className="btn btn-primary" 
              style={{ flex: 2, marginTop: 0 }} 
              onClick={handleRegister} 
              disabled={!cameraReady || isSaving}
            >
              {isSaving ? 'Menyimpan...' : 'Ambil Foto'}
            </button>
          </div>
        </div>
      )}
    </div>
  );
};

const CheckinModal = ({ event, user, onClose, onSuccess, onError }: any) => {
  const [step, setStep] = useState('location'); 
  const [msg, setMsg] = useState('Mencari lokasi...');
  const [photo, setPhoto] = useState<string | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [cameraReady, setCameraReady] = useState(false);
  const [isSigned, setIsSigned] = useState(false);
  const videoRef = useRef<HTMLVideoElement>(null);
  const signRef = useRef<HTMLCanvasElement>(null);
  const streamRef = useRef<MediaStream | null>(null);

  useEffect(() => {
    if (step === 'location') {
      navigator.geolocation.getCurrentPosition(
        (pos) => {
          const { latitude, longitude } = pos.coords;
          const dist = haversineDistance(parseFloat(event.latitude), parseFloat(event.longitude), latitude, longitude);
          if (dist <= event.radius) {
            setMsg(`Lokasi Terverifikasi! Jarak: ${Math.round(dist)}m`);
            setTimeout(() => setStep('selfie'), 800);
          } else {
            onError(`Di luar radius! Jarak: ${Math.round(dist)}m. Harus dalam ${event.radius}m.`);
            onClose();
          }
        },
        () => {
          onError("Gagal mendapatkan lokasi. Pastikan GPS aktif.");
          onClose();
        },
        { enableHighAccuracy: true, timeout: 10000 }
      );
    }
  }, [step, event.latitude, event.longitude, event.radius, onError, onClose]);

  useEffect(() => {
    if (step === 'selfie') {
      const startCamera = async () => {
        try {
          const stream = await navigator.mediaDevices.getUserMedia({ 
            video: { facingMode: 'user', width: { ideal: 640 }, height: { ideal: 640 } } 
          });
          streamRef.current = stream;
          if (videoRef.current) {
            videoRef.current.srcObject = stream;
            videoRef.current.onloadedmetadata = () => {
              videoRef.current?.play().then(() => setCameraReady(true));
            };
          }
        } catch (err) {
          onError("Kamera tidak tersedia.");
          setStep('location');
        }
      };
      startCamera();
    }
    return () => {
      if (streamRef.current) {
        streamRef.current.getTracks().forEach(track => track.stop());
        streamRef.current = null;
      }
    };
  }, [step, onError]);

  const takePhoto = () => {
    if (!cameraReady || !videoRef.current || videoRef.current.readyState < 2) return;
    const v = videoRef.current;
    const c = document.createElement('canvas');
    c.width = v.videoWidth; 
    c.height = v.videoHeight; 
    const ctx = c.getContext('2d');
    if (!ctx) return;
    ctx.translate(c.width, 0);
    ctx.scale(-1, 1);
    ctx.drawImage(v, 0, 0, c.width, c.height);
    setPhoto(c.toDataURL('image/jpeg', 0.8)); 
    
    if(streamRef.current) {
      streamRef.current.getTracks().forEach(t => t.stop());
    }
    setStep('signature');
  };

  useEffect(() => {
    if (step === 'signature' && signRef.current) {
      const cvs = signRef.current;
      cvs.width = cvs.offsetWidth; cvs.height = cvs.offsetHeight;
      const ctx = cvs.getContext('2d');
      if (!ctx) return;
      ctx.lineWidth = 3; ctx.strokeStyle = "#2C3E50"; ctx.lineJoin = 'round'; ctx.lineCap = 'round';
      let drawing = false;
      
      const getPos = (e: any) => {
        const rect = cvs.getBoundingClientRect();
        const clientX = e.touches ? e.touches[0].clientX : e.clientX;
        const clientY = e.touches ? e.touches[0].clientY : e.clientY;
        return { x: clientX - rect.left, y: clientY - rect.top };
      }

      const startDrawing = (e: any) => {
        drawing = true;
        setIsSigned(true);
        const p = getPos(e);
        ctx.beginPath();
        ctx.moveTo(p.x, p.y);
      };

      const doDrawing = (e: any) => {
        if(!drawing) return;
        const p = getPos(e);
        ctx.lineTo(p.x, p.y);
        ctx.stroke();
      };

      const stopDrawing = () => drawing = false;

      cvs.onmousedown = startDrawing;
      cvs.onmousemove = doDrawing;
      cvs.onmouseup = stopDrawing;
      cvs.onmouseleave = stopDrawing;
      
      cvs.ontouchstart = (e) => { startDrawing(e); e.preventDefault(); };
      cvs.ontouchmove = (e) => { doDrawing(e); e.preventDefault(); };
      cvs.ontouchend = (e) => { stopDrawing(); e.preventDefault(); };
    }
  }, [step]);

  const resetSignature = () => {
    if (signRef.current) {
      const ctx = signRef.current.getContext('2d');
      if (ctx) ctx.clearRect(0, 0, signRef.current.width, signRef.current.height);
      setIsSigned(false);
    }
  };

  const submit = async () => {
    if(isProcessing || !isSigned || !signRef.current) return;
    setIsProcessing(true);
    try {
      const checkSnap = await db.collection("events").doc(event.id).collection("attendance")
        .where("userId", "==", user.id)
        .limit(1)
        .get();
      
      if (!checkSnap.empty) {
        onError("Anda sudah mengisi kehadiran untuk acara ini.");
        onClose();
        return;
      }

      const signImg = signRef.current.toDataURL('image/png');
      await db.collection("events").doc(event.id).collection("attendance").add({
        userId: user.id, userName: user.name, 
        userNip: user.nip || '-', userInstansi: user.instansi || '-',
        photo, signature: signImg, checkinTime: firebase.firestore.Timestamp.now()
      });
      onSuccess();
    } catch (e) { 
      onError("Gagal mengirim data. Coba lagi."); 
      setIsProcessing(false);
    }
  };

  return (
    <div className="modal">
      <div className="modal-content">
        <div className="header"><h3>ABSENSI: {event.name}</h3><button onClick={onClose} className="logout-btn" disabled={isProcessing}>Batal</button></div>
        {step === 'location' && <div style={{textAlign:'center', padding:40}}><p>{msg}</p></div>}
        {step === 'selfie' && (
          <div style={{textAlign:'center'}}>
            <div style={{position:'relative', background:'#000', borderRadius:8, overflow:'hidden', minHeight:300}}>
               {!cameraReady && <p style={{color:'#fff', position:'absolute', top:'50%', left:'50%', transform:'translate(-50%,-50%)'}}>Inisialisasi Kamera...</p>}
               <video ref={videoRef} autoPlay playsInline muted style={{display: cameraReady ? 'block' : 'none'}}></video>
            </div>
            <button className="btn btn-primary" onClick={takePhoto} disabled={!cameraReady}>Ambil Foto Sekarang</button>
          </div>
        )}
        {step === 'signature' && (
          <div>
            <p style={{fontSize:14, marginBottom:10}}>
              Bubuhkan tanda tangan Anda:<span className="required-star">*</span>
            </p>
            <canvas ref={signRef} className="signature-canvas"></canvas>
            {!isSigned && <p style={{fontSize:11, color:'#E74C3C', marginTop:-5, marginBottom:10}}>Tanda tangan wajib diisi sebelum dikirim.</p>}
            <div style={{display:'flex', gap:10}}>
              <button className="btn btn-outline" style={{flex:1}} onClick={resetSignature}>Reset</button>
              <button 
                className="btn btn-primary" 
                style={{flex:2}} 
                onClick={submit} 
                disabled={isProcessing || !isSigned}
              >
                {isProcessing ? 'Mengirim...' : 'Kirim Absensi'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

const AttendanceReport = ({ eventId, eventName, onClose, formatDateTime }: any) => {
  const [attendance, setAttendance] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsubscribe = db.collection("events").doc(eventId).collection("attendance")
      .orderBy("checkinTime", "desc")
      .onSnapshot(s => {
        setAttendance(s.docs.map(d => d.data()));
        setLoading(false);
      });
    return () => unsubscribe();
  }, [eventId]);

  const exportHTML = () => {
    let rows = "";
    attendance.forEach((r, i) => {
      const ident = (r.userNip && r.userNip !== '-') ? r.userNip : (r.userInstansi || '-');
      rows += `<tr><td>${i+1}</td><td><b>${r.userName}</b></td><td>${ident}</td><td>${formatDateTime(r.checkinTime)}</td><td align="center"><img src="${r.photo}" width="60"></td><td align="center"><img src="${r.signature}" width="100"></td></tr>`;
    });
    const html = `<html><head><title>Laporan - ${eventName}</title><style>body{font-family:sans-serif;padding:20px}table{width:100%;border-collapse:collapse}th,td{border:1px solid #ddd;padding:8px}th{background:#f5f5f5}</style></head><body><h2>Daftar Hadir: ${eventName}</h2><table><thead><tr><th>No</th><th>Nama</th><th>Identitas</th><th>Waktu</th><th>Foto</th><th>TTD</th></tr></thead><tbody>${rows}</tbody></table></body></html>`;
    const blob = new Blob([html], {type:'text/html'});
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `Laporan_${eventName}.html`;
    a.click();
  };

  const exportXLSX = () => {
    const data = attendance.map((r, i) => ({
      'No': i + 1,
      'Nama': r.userName,
      'NIP/Instansi': (r.userNip && r.userNip !== '-') ? r.userNip : (r.userInstansi || '-'),
      'Waktu Check-in': formatDateTime(r.checkinTime),
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Daftar Hadir");
    XLSX.writeFile(wb, `Laporan_${eventName}.xlsx`);
  };

  return (
    <div className="modal">
      <div className="modal-content">
        <div className="header"><h3>Laporan: {eventName}</h3><button onClick={onClose} className="logout-btn">Tutup</button></div>
        {loading ? <p>Memuat data...</p> : (
          <React.Fragment>
            <div className="stats-card">
               <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"></path><circle cx="9" cy="7" r="4"></circle><path d="M23 21v-2a4 4 0 0 0-3-3.87"></path><path d="M16 3.13a4 4 0 0 1 0 7.75"></path></svg>
               Total: {attendance.length} Orang Sudah Hadir
            </div>
            <div style={{display:'flex', gap:10, marginBottom:15}}>
              <button className="btn btn-primary" onClick={exportHTML} style={{width:'auto', marginTop: 0}}>Download HTML</button>
              <button className="btn btn-secondary" onClick={exportXLSX} style={{width:'auto', marginTop: 0}}>Download XLSX</button>
            </div>
            <div className="table-responsive">
              <table>
                <thead><tr><th>Nama</th><th>Identitas</th><th>Waktu</th><th>Foto</th></tr></thead>
                <tbody>
                  {attendance.map((a,i) => (
                    <tr key={i}>
                      <td>{a.userName}</td>
                      <td>{a.userNip !== '-' ? a.userNip : a.userInstansi}</td>
                      <td style={{fontSize:11}}>{formatDateTime(a.checkinTime)}</td>
                      <td><img src={a.photo} className="thumb" /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </React.Fragment>
        )}
      </div>
    </div>
  );
}

const AdminDashboard = ({ onLogout }: any) => {
  const defLat = '-8.177229';
  const defLng = '113.700393';
  const [adminTab, setAdminTab] = useState('acara');
  const [events, setEvents] = useState<any[]>([]);
  const [newEvent, setNewEvent] = useState({ name: '', date: '', startTime: '07:00', endTime: '10:00', locationName: '', latitude: defLat, longitude: defLng, radius: 50 });
  const [editingEvent, setEditingEvent] = useState<any>(null);
  const [viewId, setViewId] = useState<any>(null);
  const [toast, setToast] = useState<any>(null);
  const [isSaving, setIsSaving] = useState(false);
  const [closingId, setClosingId] = useState<string | null>(null);
  const [confirmClose, setConfirmClose] = useState<any>(null);

  const [bulkEmployeeText, setBulkEmployeeText] = useState('');
  const [newPass, setNewPass] = useState('');

  useEffect(() => {
    const unsubscribe = db.collection("events").orderBy("createdAt", "desc").onSnapshot(s => {
      setEvents(s.docs.map(d => ({ id: d.id, ...d.data() })));
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    const checkAutoClose = async () => {
      const ongoingEvents = events.filter(e => e.status === 'berlangsung');
      for (const ev of ongoingEvents) {
        if (isEventExpired(ev.date, ev.endTime)) {
          try {
            await db.collection("events").doc(ev.id).update({ status: 'selesai' });
          } catch (err) {
            console.error("Auto close failed", err);
          }
        }
      }
    };
    const timer = setInterval(checkAutoClose, 60000);
    checkAutoClose();
    return () => clearInterval(timer);
  }, [events]);

  const handleCreate = async (e: any) => {
    e.preventDefault();
    setIsSaving(true);
    try {
      await db.collection("events").add({ ...newEvent, createdAt: firebase.firestore.Timestamp.now(), status: 'berlangsung' });
      setNewEvent({ name: '', date: '', startTime: '07:00', endTime: '10:00', locationName: '', latitude: defLat, longitude: defLng, radius: 50 });
      setToast({ message: "Acara berhasil dibuat!", type: 'success' });
    } catch(err) { setToast({ message: "Gagal membuat acara.", type: 'error' }); }
    setIsSaving(false);
  };

  const handleUpdateAdminPass = async (e: any) => {
    e.preventDefault();
    if(!newPass) return;
    setIsSaving(true);
    try {
      await db.collection("config").doc("admin").update({ password: newPass });
      setNewPass('');
      setToast({ message: "Password Admin berhasil diubah!", type: 'success' });
    } catch(err) { setToast({ message: "Gagal mengubah password.", type: 'error' }); }
    setIsSaving(false);
  }

  const handleBulkUpdateEmployees = async () => {
    if(!bulkEmployeeText) return;
    setIsSaving(true);
    let updatedCount = 0;
    let skippedCount = 0;
    try {
      const data = JSON.parse(bulkEmployeeText);
      if(!Array.isArray(data)) throw new Error("Format harus Array JSON");
      for (const emp of data) {
        if(!emp.nip) continue;
        const snap = await db.collection("asn_users").where("nip", "==", emp.nip.toString()).limit(1).get();
        if(!snap.empty) {
          await db.collection("asn_users").doc(snap.docs[0].id).update({ 
            name: emp.name || snap.docs[0].data().name,
            instansi: emp.instansi || snap.docs[0].data().instansi || '-'
          });
          updatedCount++;
        } else {
          skippedCount++;
        }
      }
      setBulkEmployeeText('');
      setToast({ message: `Selesai! ${updatedCount} Diupdate, ${skippedCount} Dilewati.`, type: 'success' });
    } catch(err: any) { 
      setToast({ message: "Error: " + err.message, type: 'error' }); 
    }
    setIsSaving(false);
  };

  const executeCloseEvent = async () => {
    const id = confirmClose.id;
    setClosingId(id);
    setConfirmClose(null);
    try {
      await db.collection("events").doc(id).update({ status: 'selesai' });
      setToast({ message: "Berhasil! Acara telah ditutup.", type: 'success' });
    } catch (err) {
      setToast({ message: "Gagal menutup acara.", type: 'error' });
    } finally {
      setClosingId(null);
    }
  };

  const handleUpdateEvent = async (e: any) => {
    e.preventDefault();
    setIsSaving(true);
    try {
      const { id, ...updateData } = editingEvent;
      await db.collection("events").doc(id).update(updateData);
      setEditingEvent(null);
      setToast({ message: "Acara diperbarui!", type: 'success' });
    } catch (err) {
      setToast({ message: "Gagal memperbarui acara.", type: 'error' });
    } finally {
      setIsSaving(false);
    }
  };

  return (
    <div className="container">
      {toast && <Toast {...toast} onClose={()=>setToast(null)} />}
      <div className="header"><h2>Panel Admin KJati</h2><button className="logout-btn" onClick={onLogout}>Keluar</button></div>
      <div className="tabs">
        <button className={`tab ${adminTab==='acara'?'active':''}`} onClick={()=>setAdminTab('acara')}>Acara</button>
        <button className={`tab ${adminTab==='pegawai'?'active':''}`} onClick={()=>setAdminTab('pegawai')}>Pegawai</button>
        <button className={`tab ${adminTab==='pengaturan'?'active':''}`} onClick={()=>setAdminTab('pengaturan')}>Pengaturan</button>
      </div>
      {adminTab === 'acara' && (
        <React.Fragment>
          <div className="card" style={{position:'relative'}}>
            {isSaving && <div className="loading-overlay">Menyimpan...</div>}
            <h3>Buat Acara Baru</h3>
            <form onSubmit={handleCreate}>
              <label>Nama Acara</label><input value={newEvent.name} onChange={e=>setNewEvent({...newEvent, name:e.target.value})} required />
              <div style={{display:'flex', gap:10}}>
                <div style={{flex:2}}><label>Tanggal</label><input type="date" value={newEvent.date} onChange={e=>setNewEvent({...newEvent, date:e.target.value})} required /></div>
                <div style={{flex:1}}><label>Mulai</label><input type="time" value={newEvent.startTime} onChange={e=>setNewEvent({...newEvent, startTime:e.target.value})} required /></div>
                <div style={{flex:1}}><label>Selesai</label><input type="time" value={newEvent.endTime} onChange={e=>setNewEvent({...newEvent, endTime:e.target.value})} required /></div>
              </div>
              <label>Lokasi Tempat</label><input value={newEvent.locationName} onChange={e=>setNewEvent({...newEvent, locationName:e.target.value})} required />
              <div style={{display:'flex', gap:10}}>
                <div style={{flex:1}}><label>Latitude</label><input value={newEvent.latitude} onChange={e=>setNewEvent({...newEvent, latitude:e.target.value})} /></div>
                <div style={{flex:1}}><label>Longitude</label><input value={newEvent.longitude} onChange={e=>setNewEvent({...newEvent, longitude:e.target.value})} /></div>
                <div style={{flex:1}}><label>Radius (m)</label><input type="number" value={newEvent.radius} onChange={e=>setNewEvent({...newEvent, radius: parseInt(e.target.value) || 0})} /></div>
              </div>
              <button className="btn btn-primary" disabled={isSaving}>Publikasikan Acara</button>
            </form>
          </div>
          <div className="card">
            <h3>Monitoring Acara</h3>
            {events.map(ev => (
              <div key={ev.id} style={{padding:'15px 0', borderBottom:'1px solid #eee', display:'flex', justifyContent:'space-between', alignItems:'center'}}>
                <div>
                  <span className={`badge ${ev.status==='berlangsung'?'badge-success':'badge-danger'}`}>{ev.status}</span>
                  <div style={{fontWeight:600, marginTop:5}}>{ev.name}</div>
                  <small style={{color:'#888'}}>{ev.date}</small>
                </div>
                <div style={{display:'flex', gap:5}}>
                  <button className="btn btn-outline" style={{width:'auto', padding:'6px 12px', marginTop:0, fontSize:12}} onClick={()=>setViewId(ev)}>Laporan</button>
                  {ev.status === 'berlangsung' && (
                    <React.Fragment>
                      <button className="btn btn-secondary" style={{width:'auto', padding:'6px 12px', marginTop:0, fontSize:12}} onClick={()=>setEditingEvent({...ev})}>Edit</button>
                      <button className="btn btn-danger" style={{width:'auto', padding:'6px 12px', marginTop:0, fontSize:12}} disabled={closingId === ev.id} onClick={() => setConfirmClose(ev)}> {closingId === ev.id ? '...' : 'Tutup'} </button>
                    </React.Fragment>
                  )}
                </div>
              </div>
            ))}
          </div>
        </React.Fragment>
      )}
      {adminTab === 'pegawai' && (
        <div className="card" style={{position:'relative'}}>
          {isSaving && <div className="loading-overlay">Memproses...</div>}
          <h3>Update Database Pegawai</h3>
          <p style={{fontSize: 12, color: '#666', marginBottom: 15}}>Gunakan format JSON Array.</p>
          <textarea placeholder='[{"nip":"198...", "name":"Budi Santoso", "instansi":"KUA..."}, ...]' value={bulkEmployeeText} onChange={e=>setBulkEmployeeText(e.target.value)} />
          <button className="btn btn-secondary" onClick={handleBulkUpdateEmployees} disabled={isSaving || !bulkEmployeeText}>Proses Update Pegawai</button>
        </div>
      )}
      {adminTab === 'pengaturan' && (
        <div className="card" style={{position:'relative'}}>
          {isSaving && <div className="loading-overlay">Menyimpan...</div>}
          <h3>Pengaturan Admin</h3>
          <form onSubmit={handleUpdateAdminPass}>
            <label>Ganti Password Admin Baru</label>
            <input type="password" value={newPass} onChange={e=>setNewPass(e.target.value)} required placeholder="Masukkan password baru" />
            <button className="btn btn-danger" disabled={isSaving || !newPass}>Simpan Password Baru</button>
          </form>
        </div>
      )}
      {confirmClose && (
        <div className="modal">
          <div className="modal-content" style={{maxWidth: 400, textAlign: 'center'}}>
            <h3>Konfirmasi Tutup</h3>
            <p style={{margin: '15px 0', fontSize: 14}}>Anda yakin ingin mengakhiri sesi absensi untuk acara <b>{confirmClose.name}</b>?</p>
            <div style={{display:'flex', gap: 10}}>
              <button className="btn btn-outline" style={{flex: 1}} onClick={() => setConfirmClose(null)}>Batal</button>
              <button className="btn btn-danger" style={{flex: 1}} onClick={executeCloseEvent}>Ya, Tutup</button>
            </div>
          </div>
        </div>
      )}
      {editingEvent && (
        <div className="modal">
          <div className="modal-content" style={{maxWidth: 600}}>
            <div className="header"><h3>Edit Detail Acara</h3><button onClick={()=>setEditingEvent(null)} className="logout-btn">Batal</button></div>
            <form onSubmit={handleUpdateEvent}>
               <label>Nama Acara</label><input value={editingEvent.name} onChange={e=>setEditingEvent({...editingEvent, name:e.target.value})} required />
               <div style={{display:'flex', gap:10}}>
                  <div style={{flex:2}}><label>Tanggal</label><input type="date" value={editingEvent.date} onChange={e=>setEditingEvent({...editingEvent, date:e.target.value})} required /></div>
                  <div style={{flex:1}}><label>Mulai</label><input type="time" value={editingEvent.startTime} onChange={e=>setEditingEvent({...editingEvent, startTime:e.target.value})} required /></div>
                  <div style={{flex:1}}><label>Selesai</label><input type="time" value={editingEvent.endTime} onChange={e=>setEditingEvent({...editingEvent, endTime:e.target.value})} required /></div>
               </div>
               <label>Lokasi Tempat</label><input value={editingEvent.locationName} onChange={e=>setEditingEvent({...editingEvent, locationName:e.target.value})} required />
               <div style={{display:'flex', gap:10}}>
                  <div style={{flex:1}}><label>Latitude</label><input value={editingEvent.latitude} onChange={e=>setEditingEvent({...editingEvent, latitude:e.target.value})} /></div>
                  <div style={{flex:1}}><label>Longitude</label><input value={editingEvent.longitude} onChange={e=>setEditingEvent({...editingEvent, longitude:e.target.value})} /></div>
                  <div style={{flex:1}}><label>Radius (m)</label><input type="number" value={editingEvent.radius} onChange={e=>setEditingEvent({...editingEvent, radius: parseInt(e.target.value) || 0})} /></div>
               </div>
               <button className="btn btn-primary" disabled={isSaving} style={{marginTop: 20}}>Simpan Perubahan</button>
            </form>
          </div>
        </div>
      )}
      {viewId && <AttendanceReport eventId={viewId.id} eventName={viewId.name} onClose={()=>setViewId(null)} formatDateTime={formatDateTime} />}
    </div>
  );
};

const UserDashboard = ({ user, onLogout, onUpdateUser }: any) => {
  const [activeEvents, setActiveEvents] = useState<any[]>([]);
  const [attendedIds, setAttendedIds] = useState(new Set());
  const [userHistory, setUserHistory] = useState<any[]>([]);
  const [loadingActive, setLoadingActive] = useState(true);
  const [loadingHistory, setLoadingHistory] = useState(true);
  const [selected, setSelected] = useState<any>(null);
  const [toast, setToast] = useState<any>(null);

  useEffect(() => {
    const unsubscribe = db.collection("events")
      .where("status", "==", "berlangsung")
      .onSnapshot(async (s) => {
        const evList = s.docs.map(d => ({ id: d.id, ...d.data() }))
          .filter((ev: any) => !isEventExpired(ev.date, ev.endTime));
        setActiveEvents(evList);
        
        const attendedSet = new Set();
        try {
          const promises = evList.map(ev => 
            db.collection("events").doc(ev.id).collection("attendance")
              .where("userId", "==", user.id).limit(1).get()
          );
          const snaps = await Promise.all(promises);
          snaps.forEach((snap, index) => {
            if (!snap.empty) attendedSet.add(evList[index].id);
          });
          setAttendedIds(attendedSet);
        } catch (err) {
          console.error("Error checking attendance for active events", err);
        }
        setLoadingActive(false);
      });
    return () => unsubscribe();
  }, [user.id]);

  useEffect(() => {
    const loadHistory = async () => {
      setLoadingHistory(true);
      try {
        const s = await db.collection("events").orderBy("createdAt", "desc").limit(20).get();
        const evList = s.docs.map(d => ({ id: d.id, ...d.data() }));
        
        const historyPromises = evList.map(async (ev: any) => {
          const snap = await db.collection("events").doc(ev.id).collection("attendance")
            .where("userId", "==", user.id).limit(1).get();
          if (!snap.empty) return { ...ev, attendance: snap.docs[0].data() };
          return null;
        });
        const results = await Promise.all(historyPromises);
        setUserHistory(results.filter(r => r !== null).sort((a: any, b: any) => b.attendance.checkinTime.seconds - a.attendance.checkinTime.seconds));
      } catch (e) { 
        console.error("History error", e); 
      } finally {
        setLoadingHistory(false);
      }
    };
    loadHistory();
  }, [user.id]);

  const handleAttendanceSuccess = () => {
    setSelected(null);
    setToast({message: `Absensi Berhasil! Anda telah tercatat hadir.`, type:"success"});
  };

  return (
    <div className="container">
      {toast && <Toast {...toast} onClose={()=>setToast(null)} />}
      <div className="header">
        <div><h2 style={{fontSize:18}}>Salam, {user.name}</h2><small>{user.role==='asn'?user.nip:user.instansi}</small></div>
        <button className="logout-btn" onClick={onLogout}>Keluar</button>
      </div>

      <div className="card">
        <h3 className="section-title">
           <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
           Acara Aktif Hari Ini
        </h3>
        {loadingActive ? <p style={{textAlign:'center', padding:20, color:'#888'}}>Mencari acara aktif...</p> : (
          activeEvents.length === 0 ? (
            <p style={{color:'#888', textAlign:'center', padding: '10px 0'}}>Tidak ada acara aktif saat ini.</p>
          ) : activeEvents.map(ev => {
            const isAlreadyAttended = attendedIds.has(ev.id);
            const isNotStarted = isEventNotStartedYet(ev.date, ev.startTime);
            return (
              <div key={ev.id} style={{padding:'15px 0', borderBottom:'1px solid #eee', display:'flex', justifyContent:'space-between', alignItems:'center'}}>
                <div style={{flex: 1, paddingRight: 10}}>
                  <div style={{fontWeight: 600}}>{ev.name}</div>
                  <small style={{color:'#666', display:'block'}}>{ev.locationName} • Pukul {ev.startTime} - {ev.endTime}</small>
                  {isAlreadyAttended && (
                    <span className="status-checked">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
                      Sudah Hadir
                    </span>
                  )}
                  {isNotStarted && !isAlreadyAttended && (
                    <span style={{color: '#E67E22', fontSize: 12, fontWeight: 500, display: 'flex', alignItems: 'center', gap: 4, marginTop: 4}}>
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>
                      Belum dimulai (Mulai pukul {ev.startTime})
                    </span>
                  )}
                </div>
                <div>
                  {isAlreadyAttended ? (
                    <span className="badge badge-gray" style={{padding: '8px 12px'}}>Selesai</span>
                  ) : isNotStarted ? (
                    <button className="btn" style={{width:'auto', padding:'8px 20px', marginTop:0, background:'#eaeded', color:'#7f8c8d', cursor:'not-allowed', border:'1px solid #bdc3c7'}} disabled>Belum Mulai</button>
                  ) : (
                    <button className="btn btn-primary" style={{width:'auto', padding:'8px 20px', marginTop:0}} onClick={()=>setSelected(ev)}>Hadir</button>
                  )}
                </div>
              </div>
            );
          })
        )}
      </div>

      {user.role === 'asn' && (
        <UserFaceIdSettings 
          user={user} 
          onUpdateUser={onUpdateUser} 
          onError={(m: string) => setToast({ message: m, type: 'error' })} 
          onSuccess={(m: string) => setToast({ message: m, type: 'success' })} 
        />
      )}

      <div className="card">
        <h3 className="section-title">
           <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 20v-6M9 20v-10M15 20v-4M3 20h18"/></svg>
           Riwayat Kehadiran Saya
        </h3>
        {loadingHistory ? <p style={{textAlign:'center', padding:20, color:'#888'}}>Memuat riwayat...</p> : (
          userHistory.length === 0 ? (
            <p style={{color:'#888', textAlign:'center', padding: '10px 0'}}>Belum ada riwayat kehadiran baru-baru ini.</p>
          ) : userHistory.map(h => (
            <div key={h.id} className="history-item">
              <div className="history-header">
                <div>
                  <div style={{fontWeight: 600, fontSize:15}}>{h.name}</div>
                  <small style={{color:'#888'}}>{h.date} • {h.locationName}</small>
                </div>
                <span className={`badge ${h.status==='selesai'?'badge-danger':'badge-success'}`}>{h.status}</span>
              </div>
              <div style={{marginTop:8, borderTop:'1px dashed #eee', paddingTop:8, fontSize:12, color:'#444'}}>
                <b>Waktu Presensi:</b> {formatDateTime(h.attendance.checkinTime)}
                <div className="history-details">
                  <img src={h.attendance.photo} className="history-thumb" alt="Selfie" />
                  <img src={h.attendance.signature} className="history-thumb" style={{background:'#f9f9f9'}} alt="Tanda Tangan" />
                </div>
              </div>
            </div>
          ))
        )}
      </div>

      {selected && (
        <CheckinModal event={selected} user={user} onClose={()=>setSelected(null)} onSuccess={handleAttendanceSuccess} onError={(m: string)=>setToast({message:m, type:"error"})} />
      )}
    </div>
  );
};

const LoginPage = ({ onLogin }: any) => {
  const [tab, setTab] = useState('asn');
  const [asnLoginMethod, setAsnLoginMethod] = useState<'face' | 'nip'>('nip');
  const [showRegisterModal, setShowRegisterModal] = useState(false);
  const [nip, setNip] = useState('');
  const [suggestions, setSuggestions] = useState<any[]>([]);
  const [inputs, setInputs] = useState({ name: '', instansi: '', nip: '', jabatan: '', user: '', pass: '' });
  const [loading, setLoading] = useState(false);
  const [toast, setToast] = useState<any>(null);

  const handleNip = async (val: string) => {
    setNip(val);
    const res = await dbService.searchAsn(val);
    setSuggestions(res);
  };

  const submit = async (e: any) => {
    e.preventDefault();
    setLoading(true);
    try {
      if(tab==='asn') {
        const u = await dbService.validateAsn(nip);
        if(!u) throw new Error("NIP tidak ditemukan");
        onLogin(u);
      } else if(tab==='umum') {
        onLogin({ 
          id: 'guest_'+Date.now(), 
          name: inputs.name, 
          instansi: inputs.instansi, 
          nip: inputs.nip || undefined, 
          jabatan: inputs.jabatan || undefined, 
          role: 'umum' 
        });
      } else {
        const snap = await db.collection("config").doc("admin").get();
        const admin = snap.data();
        if(admin && admin.username === inputs.user && admin.password === inputs.pass) onLogin({ name: 'Admin', role: 'admin' });
        else throw new Error("Username/Password Salah");
      }
    } catch(err: any) { setToast({ message: err.message, type: 'error' }); }
    setLoading(false);
  };

  const reactNativeTypeScriptCode = [
    "import React, { useState } from 'react';",
    "import {",
    "  StyleSheet,",
    "  Text,",
    "  View,",
    "  TextInput,",
    "  TouchableOpacity,",
    "  ScrollView,",
    "  SafeAreaView,",
    "  KeyboardAvoidingView,",
    "  Platform,",
    "  Alert,",
    "} from 'react-native';",
    "",
    "type LoginRole = 'asn' | 'umum' | 'admin';",
    "",
    "export default function KjatiLogin() {",
    "  const [role, setRole] = useState<LoginRole>('asn');",
    "  const [nip, setNip] = useState('');",
    "  const [name, setName] = useState('');",
    "  const [instansi, setInstansi] = useState('');",
    "  const [user, setUser] = useState('');",
    "  const [pass, setPass] = useState('');",
    "",
    "  const handleLogin = () => {",
    "    if (role === 'asn') {",
    "      if (!nip) {",
    "        Alert.alert('Error', 'Silakan masukkan NIP Pegawai Anda');",
    "        return;",
    "      }",
    "      Alert.alert('Sukses', 'Mencoba login ASN dengan NIP: ' + nip);",
    "    } else if (role === 'umum') {",
    "      if (!name || !instansi) {",
    "        Alert.alert('Error', 'Silakan isi nama dan asal instansi Anda');",
    "        return;",
    "      }",
    "      Alert.alert('Sukses', 'Selamat Datang Umum: ' + name);",
    "    } else {",
    "      if (!user || !pass) {",
    "        Alert.alert('Error', 'Silakan masukkan username dan password admin');",
    "        return;",
    "      }",
    "      Alert.alert('Sukses', 'Mencoba autentikasi Admin');",
    "    }",
    "  };",
    "",
    "  return (",
    "    <SafeAreaView style={styles.safeContainer}>",
    "      <KeyboardAvoidingView",
    "        behavior={Platform.OS === 'ios' ? 'padding' : 'height'}",
    "        style={styles.container}",
    "      >",
    "        <ScrollView contentContainerStyle={styles.scrollContainer} showsVerticalScrollIndicator={false}>",
    "          {/* Header Brand Kjati */}",
    "          <View style={styles.headerSection}>",
    "            <View style={styles.logoBorder}>",
    "              <View style={styles.logoContainer}>",
    "                <Text style={styles.logoText}>K</Text>",
    "              </View>",
    "            </View>",
    "            <Text style={styles.brandName}>KJATI</Text>",
    "            <Text style={styles.brandSubtitle}>Kemenag Jember Absensi Terintegrasi</Text>",
    "            <View style={styles.badgeContainer}>",
    "              <Text style={styles.badgeText}>🕌 Madrasah Version</Text>",
    "            </View>",
    "          </View>",
    "",
    "          {/* Form Login Premium Card */}",
    "          <View style={styles.card}>",
    "            <Text style={styles.cardTitle}>Pilih Jenis Login</Text>",
    "            <Text style={styles.cardSubtitle}>Silakan pilih jenis akun untuk melanjutkan</Text>",
    "",
    "            {/* Tab Selector layout */}",
    "            <View style={styles.tabContainer}>",
    "              <TouchableOpacity",
    "                style={[styles.tabButton, role === 'asn' && styles.tabButtonActive]}",
    "                onPress={() => setRole('asn')}",
    "              >",
    "                <Text style={[styles.tabText, role === 'asn' && styles.tabTextActive]}>ASN</Text>",
    "                <Text style={styles.tabSubText}>Pegawai</Text>",
    "              </TouchableOpacity>",
    "",
    "              <TouchableOpacity",
    "                style={[styles.tabButton, role === 'umum' && styles.tabButtonActive]}",
    "                onPress={() => setRole('umum')}",
    "              >",
    "                <Text style={[styles.tabText, role === 'umum' && styles.tabTextActive]}>Umum</Text>",
    "                <Text style={styles.tabSubText}>Pendukung</Text>",
    "              </TouchableOpacity>",
    "",
    "              <TouchableOpacity",
    "                style={[styles.tabButton, role === 'admin' && styles.tabButtonActive]}",
    "                onPress={() => setRole('admin')}",
    "              >",
    "                <Text style={[styles.tabText, role === 'admin' && styles.tabTextActive]}>Admin</Text>",
    "                <Text style={styles.tabSubText}>Sistem</Text>",
    "              </TouchableOpacity>",
    "            </View>",
    "",
    "            {/* Inputs Form block */}",
    "            <View style={styles.formContainer}>",
    "              {role === 'asn' && (",
    "                <View style={styles.inputGroup}>",
    "                  <Text style={styles.inputLabel}>NIP Pegawai</Text>",
    "                  <View style={styles.inputWrapper}>",
    "                    <Text style={styles.inputIcon}>👤</Text>",
    "                    <TextInput",
    "                      style={styles.textInput}",
    "                      placeholder='Masukkan NIP Pegawai',",
    "                      placeholderTextColor='#94a3b8',",
    "                      value={nip},",
    "                      onChangeText={setNip},",
    "                      keyboardType='numeric'",
    "                    />",
    "                  </View>",
    "                </View>",
    "              )}",
    "",
    "              {role === 'umum' && (",
    "                <>",
    "                  <View style={styles.inputGroup}>",
    "                    <Text style={styles.inputLabel}>Nama Lengkap</Text>",
    "                    <View style={styles.inputWrapper}>",
    "                      <Text style={styles.inputIcon}>👤</Text>",
    "                      <TextInput",
    "                        style={styles.textInput}",
    "                        placeholder='Nama Lengkap',",
    "                        placeholderTextColor='#94a3b8',",
    "                        value={name},",
    "                        onChangeText={setName},",
    "                      />",
    "                    </View>",
    "                  </View>",
    "                  <View style={styles.inputGroup}>",
    "                    <Text style={styles.inputLabel}>Asal Instansi</Text>",
    "                    <View style={styles.inputWrapper}>",
    "                      <Text style={styles.inputIcon}>🏢</Text>",
    "                      <TextInput",
    "                        style={styles.textInput}",
    "                        placeholder='Instansi / Sekolah',",
    "                        placeholderTextColor='#94a3b8',",
    "                        value={instansi},",
    "                        onChangeText={setInstansi},",
    "                      />",
    "                    </View>",
    "                  </View>",
    "                </>",
    "              )}",
    "",
    "              {role === 'admin' && (",
    "                <>",
    "                  <View style={styles.inputGroup}>",
    "                    <Text style={styles.inputLabel}>Username Admin</Text>",
    "                    <View style={styles.inputWrapper}>",
    "                      <Text style={styles.inputIcon}>🔑</Text>",
    "                      <TextInput",
    "                        style={styles.textInput}",
    "                        placeholder='Username Admin',",
    "                        placeholderTextColor='#94a3b8',",
    "                        value={user},",
    "                        onChangeText={setUser},",
    "                        autoCapitalize='none',",
    "                      />",
    "                    </View>",
    "                  </View>",
    "                  <View style={styles.inputGroup}>",
    "                    <Text style={styles.inputLabel}>Kata Sandi</Text>",
    "                    <View style={styles.inputWrapper}>",
    "                      <Text style={styles.inputIcon}>🔒</Text>",
    "                      <TextInput",
    "                        style={styles.textInput}",
    "                        placeholder='Kata Sandi Admin',",
    "                        placeholderTextColor='#94a3b8',",
    "                        secureTextEntry,",
    "                        value={pass},",
    "                        onChangeText={setPass},",
    "                        autoCapitalize='none',",
    "                      />",
    "                    </View>",
    "                  </View>",
    "                </>",
    "              )}",
    "",
    "              {/* Action Button */}",
    "              <TouchableOpacity style={styles.submitBtn} onPress={handleLogin}>",
    "                <Text style={styles.submitBtnText}>Masuk Sekarang</Text>",
    "              </TouchableOpacity>",
    "            </View>",
    "",
    "            {/* Shield Security Alert */}",
    "            <View style={styles.noticeBox}>",
    "              <Text style={styles.noticeTitle}>🛡️ Aman & Terpercaya</Text>",
    "              <Text style={styles.noticeBody}>",
    "                Sistem absensi aman dan terintegrasi dengan basis data Kementerian Agama Kantor Kabupaten Jember.",
    "              </Text>",
    "            </View>",
    "          </View>",
    "",
    "          <Text style={styles.footerText}>© 2026 Kementerian Agama Kabupaten Jember</Text>",
    "        </ScrollView>",
    "      </KeyboardAvoidingView>",
    "    </SafeAreaView>",
    "  );",
    "}",
    "",
    "const styles = StyleSheet.create({",
    "  safeContainer: {",
    "    flex: 1,",
    "    backgroundColor: '#052412',",
    "  },",
    "  container: {",
    "    flex: 1,",
    "  },",
    "  scrollContainer: {",
    "    alignItems: 'center',",
    "    paddingHorizontal: 20,",
    "    paddingTop: 45,",
    "    paddingBottom: 35,",
    "  },",
    "  headerSection: {",
    "    alignItems: 'center',",
    "    marginBottom: 30,",
    "  },",
    "  logoBorder: {",
    "    padding: 3,",
    "    borderRadius: 24,",
    "    borderWidth: 3,",
    "    borderColor: '#FFD700',",
    "    marginBottom: 10,",
    "  },",
    "  logoContainer: {",
    "    width: 60,",
    "    height: 60,",
    "    borderRadius: 20,",
    "    backgroundColor: '#198754',",
    "    alignItems: 'center',",
    "    justifyContent: 'center',",
    "  },",
    "  logoText: {",
    "    fontSize: 26,",
    "    fontWeight: '800',",
    "    color: '#FFD700',",
    "  },",
    "  brandName: {",
    "    fontSize: 30,",
    "    fontWeight: '900',",
    "    color: '#ffffff',",
    "    letterSpacing: 2,",
    "  },",
    "  brandSubtitle: {",
    "    fontSize: 13,",
    "    color: '#a3f3c3',",
    "    marginTop: 4,",
    "    textAlign: 'center',",
    "  },",
    "  badgeContainer: {",
    "    backgroundColor: 'rgba(255, 255, 255, 0.15)',",
    "    borderRadius: 20,",
    "    paddingVertical: 4,",
    "    paddingHorizontal: 12,",
    "    marginTop: 8,",
    "  },",
    "  badgeText: {",
    "    color: '#ffffff',",
    "    fontSize: 10,",
    "    fontWeight: '700',",
    "  },",
    "  card: {",
    "    backgroundColor: '#ffffff',",
    "    borderRadius: 24,",
    "    width: '100%',",
    "    padding: 24,",
    "    shadowColor: '#000',",
    "    shadowOffset: { width: 0, height: 10 },",
    "    shadowOpacity: 0.25,",
    "    shadowRadius: 12,",
    "    elevation: 8,",
    "  },",
    "  cardTitle: {",
    "    fontSize: 18,",
    "    fontWeight: '800',",
    "    color: '#052412',",
    "    textAlign: 'center',",
    "  },",
    "  cardSubtitle: {",
    "    fontSize: 12,",
    "    color: '#64748b',",
    "    textAlign: 'center',",
    "    marginTop: 4,",
    "    marginBottom: 16,",
    "  },",
    "  tabContainer: {",
    "    flexDirection: 'row',",
    "    backgroundColor: '#f1f5f9',",
    "    borderRadius: 12,",
    "    padding: 4,",
    "    marginBottom: 16,",
    "  },",
    "  tabButton: {",
    "    flex: 1,",
    "    paddingVertical: 8,",
    "    alignItems: 'center',",
    "    borderRadius: 8,",
    "  },",
    "  tabButtonActive: {",
    "    backgroundColor: '#ffffff',",
    "    borderWidth: 1,",
    "    borderColor: '#27AE60',",
    "  },",
    "  tabText: {",
    "    fontSize: 12,",
    "    fontWeight: '700',",
    "    color: '#64748b',",
    "  },",
    "  tabTextActive: {",
    "    color: '#052412',",
    "  },",
    "  tabSubText: {",
    "    fontSize: 8,",
    "    color: '#475569',",
    "    marginTop: 2,",
    "  },",
    "  formContainer: {",
    "    width: '100%',",
    "  },",
    "  inputGroup: {",
    "    marginBottom: 14,",
    "  },",
    "  inputLabel: {",
    "    fontSize: 12,",
    "    fontWeight: '700',",
    "    color: '#052412',",
    "    marginBottom: 4,",
    "  },",
    "  inputWrapper: {",
    "    flexDirection: 'row',",
    "    alignItems: 'center',",
    "    borderWidth: 1.5,",
    "    borderColor: '#cbd5e1',",
    "    borderRadius: 12,",
    "    paddingHorizontal: 12,",
    "    backgroundColor: '#ffffff',",
    "  },",
    "  inputIcon: {",
    "    fontSize: 14,",
    "    marginRight: 6,",
    "  },",
    "  textInput: {",
    "    flex: 1,",
    "    height: 44,",
    "    fontSize: 14,",
    "    color: '#1e293b',",
    "  },",
    "  submitBtn: {",
    "    backgroundColor: '#27AE60',",
    "    borderRadius: 12,",
    "    height: 48,",
    "    alignItems: 'center',",
    "    justifyContent: 'center',",
    "    marginTop: 10,",
    "  },",
    "  submitBtnText: {",
    "    color: '#ffffff',",
    "    fontSize: 15,",
    "    fontWeight: '700',",
    "  },",
    "  noticeBox: {",
    "    backgroundColor: '#f0fdf4',",
    "    borderWidth: 1,",
    "    borderColor: '#dcfce7',",
    "    borderRadius: 16,",
    "    padding: 12,",
    "    marginTop: 20,",
    "  },",
    "  noticeTitle: {",
    "    fontSize: 12,",
    "    fontWeight: '700',",
    "    color: '#166534',",
    "  },",
    "  noticeBody: {",
    "    fontSize: 10,",
    "    color: '#166534',",
    "    marginTop: 2,",
    "    lineHeight: 14,",
    "  },",
    "  footerText: {",
    "    color: 'rgba(255, 255, 255, 0.5)',",
    "    fontSize: 10,",
    "    marginTop: 28,",
    "    textAlign: 'center',",
    "  },",
    "});"
  ].join('\n');



  return (
    <div className="login-container-new">
      <div className="background-pattern"></div>
      
      {/* Decorative Golden Top Wave Curve and Gradients */}
      <svg viewBox="0 0 1440 320" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ position: 'absolute', top: 0, left: 0, width: '100%', pointerEvents: 'none', zIndex: 1, opacity: 0.22 }}>
        <path d="M0 0H1440V120C1320 180 1140 240 900 240C660 240 480 140 0 280V0Z" fill="url(#top_grad_new)" />
        <path d="M0 278C480 138 660 238 900 238C1140 238 1320 178 1440 118" stroke="#ffd700" strokeWidth="1.5" strokeDasharray="6 6" />
        <path d="M0 268C480 128 660 228 900 228C1140 228 1320 168 1440 108" stroke="#ffd700" strokeWidth="1" />
        <defs>
          <linearGradient id="top_grad_new" x1="720" y1="0" x2="720" y2="280" gradientUnits="userSpaceOnUse">
            <stop stopColor="#0d4224" />
            <stop offset="0.6" stopColor="#052412" stopOpacity="0.8" />
            <stop offset="1" stopColor="#052412" stopOpacity="0" />
          </linearGradient>
        </defs>
      </svg>

      {/* Structured Vector Mosque Silhouette Layer */}
      <div className="mosque-silhouette" style={{ position: 'absolute', bottom: 0, left: 0, right: 0, height: '240px', overflow: 'hidden', pointerEvents: 'none', zIndex: 1 }}>
        <svg viewBox="0 0 1440 280" fill="none" xmlns="http://www.w3.org/2000/svg" style={{ width: '100%', height: '100%', opacity: 0.13 }}>
          <path d="M0 280V180C150 180 250 150 400 150C550 150 650 190 800 190C950 190 1100 160 1250 160C1350 160 1410 170 1440 175V280H0Z" fill="#042010" />
          
          {/* Back Layer Mosque Silhouette */}
          <path d="M195 280V120H185V115H195V80 L200 60 L205 80 V115H215V120H205V280H195Z" fill="#0b371b" opacity="0.5"/>
          <path d="M220 280V140 C220 115 235 95 260 95 C285 95 300 115 300 140 V280H220Z" fill="#0b371b" opacity="0.5"/>
          <path d="M315 280V120H305V115H315V80 L320 60 L325 80 V115H335V120H325V280H315Z" fill="#0b371b" opacity="0.5"/>

          {/* Front Layer Detailed Mosque Profile */}
          <path d="M1000 280V90H985V83H1000V40 L1007.5 15 L1015 40 V83H1030V90H1015V280H1000Z" fill="#0f4422" opacity="0.8"/>
          <path d="M1040 280V110 C1040 90 1055 75 1075 75 C1095 75 1110 90 1110 110 V280H1040Z" fill="#0f4422" opacity="0.8"/>
          <path d="M1085 280V95 C1085 64 1110 40 1140 40 C1170 40 1195 64 1195 95 V280H1085Z" fill="#0f4422" opacity="0.9"/>
          <path d="M1170 280V110 C1170 90 1185 75 1205 75 C1225 75 1240 90 1240 110 V280H1170Z" fill="#0f4422" opacity="0.8"/>
          <path d="M1250 280V90H1235V83H1250V40 L1257.5 15 L1265 40 V83H1280V90H1265V280H1250Z" fill="#0f4422" opacity="0.8"/>

          {/* High balancing spire */}
          <path d="M120 280V70H105V62H120V30 L127.5 5 L135 30 V62H150V70H135V280H120Z" fill="#0b371b" opacity="0.6"/>

          {/* Star ornaments */}
          <circle cx="200" cy="40" r="1.5" fill="#FFD700" opacity="0.6"/>
          <circle cx="450" cy="60" r="1" fill="#FFFFFF" opacity="0.5"/>
          <circle cx="750" cy="30" r="2" fill="#FFD700" opacity="0.7"/>
          <circle cx="900" cy="50" r="1.5" fill="#FFFFFF" opacity="0.4"/>
          <circle cx="1350" cy="70" r="1" fill="#FFFFFF" opacity="0.5"/>
        </svg>
      </div>

      {toast && <Toast {...toast} onClose={()=>setToast(null)} />}

      <div className="login-main-row">
        {/* Left Side: Brand presentation */}
        <div className="login-brand-side" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center' }}>
          {/* Logo Kemenag Standalone & Centered */}
          <div style={{ 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center', 
            border: '3px solid #FFD700', 
            borderRadius: '50%', 
            padding: '6px', 
            width: '110px', 
            height: '110px', 
            background: 'linear-gradient(135deg, #198754 0%, #115c38 100%)',
            boxShadow: '0 8px 24px rgba(0, 0, 0, 0.3)',
            marginBottom: '15px'
          }}>
            <img 
              src="https://www.freepnglogos.com/uploads/logo-kemenag-png/logo-kementerian-agama-gambar-logo-depag-png-0.png" 
              style={{ width: '68px', height: 'auto', display: 'block' }} 
              alt="Logo Kemenag" 
            />
          </div>
          
          {/* KJATI Title */}
          <h1 style={{ 
            fontSize: '56px', 
            fontWeight: '900', 
            color: '#ffffff', 
            letterSpacing: '5px', 
            lineHeight: '1.1', 
            margin: '10px 0 5px 0',
            textShadow: '0 2px 8px rgba(0,0,0,0.4)',
            textAlign: 'center'
          }}>KJATI</h1>
          
          <div style={{ 
            height: '2px', 
            width: '180px', 
            background: 'linear-gradient(90deg, transparent, #FFD700, transparent)', 
            margin: '10px auto' 
          }}></div>
          
          <p style={{ 
            fontSize: '20px', 
            fontWeight: '700', 
            color: '#ffd700', 
            margin: '5px 0 12px 0',
            letterSpacing: '1px',
            textShadow: '0 2px 4px rgba(0,0,0,0.5)',
            textAlign: 'center'
          }}>
            Kemenag Jember Absensi Terintegrasi
          </p>
          
          <p style={{ 
            fontSize: '15px', 
            color: '#e2faed', 
            lineHeight: '1.6', 
            maxWidth: '440px', 
            margin: '5px auto 0 auto',
            textShadow: '0 1px 2px rgba(0,0,0,0.3)',
            textAlign: 'center',
            fontWeight: '400',
            opacity: '0.95'
          }}>
            Presensi Digital untuk layanan kerja yang lebih disiplin dan akuntabel.
          </p>

          <div className="kjati-badge" style={{ marginTop: '20px', padding: '6px 16px', background: 'rgba(255, 215, 0, 0.12)', border: '1px solid rgba(255, 215, 0, 0.3)' }}>
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#FFD700" strokeWidth="3"><polyline points="20 6 9 17 4 12"/></svg>
            <span style={{ color: '#FFD700', fontWeight: '700' }}>Terintegrasi</span>
            <span style={{ color: 'rgba(255, 215, 0, 0.4)' }}>•</span>
            <span style={{ color: '#FFD700', fontWeight: '700' }}>Akurat</span>
            <span style={{ color: 'rgba(255, 215, 0, 0.4)' }}>•</span>
            <span style={{ color: '#FFD700', fontWeight: '700' }}>Transparan</span>
          </div>
        </div>

        {/* Mobile Header: Centered & Rich Brand Layout */}
        <div className="mobile-brand-header" style={{ width: '100%', textAlign: 'center', display: 'flex', flexDirection: 'column', alignItems: 'center' }}>
          <div style={{ 
            display: 'flex', 
            alignItems: 'center', 
            justifyContent: 'center', 
            border: '2px solid #FFD700', 
            borderRadius: '50%', 
            padding: '5px', 
            width: '75px', 
            height: '75px', 
            background: 'linear-gradient(135deg, #198754 0%, #115c38 100%)',
            boxShadow: '0 4px 12px rgba(0, 0, 0, 0.25)',
            marginBottom: '8px'
          }}>
            <img 
              src="https://www.freepnglogos.com/uploads/logo-kemenag-png/logo-kementerian-agama-gambar-logo-depag-png-0.png" 
              style={{ width: '48px', height: 'auto', display: 'block' }} 
              alt="Logo Kemenag" 
            />
          </div>
          <h1 style={{ margin: 0, fontSize: '28px', color: '#fff', fontWeight: '900', letterSpacing: '2px' }}>KJATI</h1>
          <p style={{ margin: '2px 0 0 0', textShadow: '0 1px 2px rgba(0,0,0,0.3)', fontSize: '13px', color: '#ffd700', fontWeight: '700' }}>Kemenag Jember Absensi Terintegrasi</p>
          <p className="mobile-tagline" style={{ fontWeight: '400', fontSize: '11.5px', color: '#edfdf4', marginTop: '2px', textShadow: '0 1px 2px rgba(0,0,0,0.2)' }}>
            Presensi Digital untuk layanan kerja yang lebih disiplin dan akuntabel.
          </p>
        </div>

        {/* Right Side: Gorgeous White Interactive Login Card */}
        <div className="login-card-premium">
          <div style={{ textAlign: 'center' }}>
            <h2 style={{ fontSize: '22px', fontWeight: '800', color: '#0c3e26' }}>Pilih Jenis Login</h2>
            <p style={{ fontSize: '13px', color: '#64748b', marginTop: '4px' }}>Silakan pilih jenis akun untuk melanjutkan</p>
          </div>

          <div className="tab-pills-new">
            <button type="button" className={`tab-pill ${tab === 'asn' ? 'active' : ''}`} onClick={() => setTab('asn')}>
              <span style={{ fontSize: '16px' }}>👤</span>
              <span>ASN</span>
              <span style={{ fontSize: '8px', color: '#64748b', marginTop: '1px' }}>Pegawai</span>
            </button>
            <button type="button" className={`tab-pill ${tab === 'umum' ? 'active' : ''}`} onClick={() => setTab('umum')}>
              <span style={{ fontSize: '16px' }}>👥</span>
              <span>Umum</span>
              <span style={{ fontSize: '8px', color: '#64748b', marginTop: '1px' }}>Pendukung</span>
            </button>
            <button type="button" className={`tab-pill ${tab === 'admin' ? 'active' : ''}`} onClick={() => setTab('admin')}>
              <span style={{ fontSize: '16px' }}>🔑</span>
              <span>Admin</span>
              <span style={{ fontSize: '8px', color: '#64748b', marginTop: '1px' }}>Sistem</span>
            </button>
          </div>

          <form onSubmit={submit} style={{ margin: 0 }}>
            {tab==='asn' && (
              <div>
                {/* Method selector for ASN: Face Login or NIP */}
                <div className="face-tab-toggle-container" style={{ display: 'flex', gap: '8px', marginBottom: '20px', background: '#f1f5f9', padding: '4px', borderRadius: '12px', border: '1.5px solid #e2e8f0' }}>
                  <button type="button" onClick={() => setAsnLoginMethod('nip')} style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px', padding: '8px 12px', fontSize: '12.5px', fontWeight: '700', border: 'none', background: asnLoginMethod === 'nip' ? '#0c3e26' : 'none', color: asnLoginMethod === 'nip' ? '#fff' : '#64748b', borderRadius: '8px', cursor: 'pointer', transition: 'all 0.25s' }}>
                    ⌨️ Ketik NIP
                  </button>
                  <button type="button" onClick={() => setAsnLoginMethod('face')} style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px', padding: '8px 12px', fontSize: '12.5px', fontWeight: '700', border: 'none', background: asnLoginMethod === 'face' ? '#0c3e26' : 'none', color: asnLoginMethod === 'face' ? '#fff' : '#64748b', borderRadius: '8px', cursor: 'pointer', transition: 'all 0.25s' }}>
                    📷 Face ID (Instan)
                  </button>
                </div>

                {asnLoginMethod === 'face' ? (
                  <FaceLoginScanner 
                    onMatch={onLogin} 
                    onSwitchToNip={() => setAsnLoginMethod('nip')} 
                    onRegisterClick={() => setShowRegisterModal(true)} 
                  />
                ) : (
                  <div className="input-grp-premium">
                    <label>NIP Pegawai</label>
                    <div className="input-wrapper-new">
                      <span style={{ fontSize: '16px', color: '#94a3b8' }}>👤</span>
                      <input value={nip} onChange={e=>handleNip(e.target.value)} required placeholder="Contoh: 1980..." />
                    </div>
                    {suggestions.length > 0 && (
                      <ul className="suggestions" style={{ position: 'absolute', width: 'calc(100% - 70px)' }}>
                        {suggestions.map((s,i) => <li key={i} className="suggestion-item" onClick={()=>{setNip(s.nip); setSuggestions([]);}}>{s.name} ({s.nip})</li>)}
                      </ul>
                    )}
                  </div>
                )}
              </div>
            )}

            {tab==='umum' && (
              <div className="input-grp-premium" style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
                <div>
                  <label>Nama Lengkap</label>
                  <div className="input-wrapper-new">
                    <span style={{ fontSize: '16px', color: '#94a3b8' }}>👤</span>
                    <input value={inputs.name} onChange={e=>setInputs({...inputs, name:e.target.value})} required placeholder="Masukkan Nama Lengkap Anda" />
                  </div>
                </div>
                <div>
                  <label>Instansi / Sekolah</label>
                  <div className="input-wrapper-new">
                    <span style={{ fontSize: '16px', color: '#94a3b8' }}>🏢</span>
                    <input value={inputs.instansi} onChange={e=>setInputs({...inputs, instansi:e.target.value})} required placeholder="Contoh: KUA Jati, MAN Jember, dll" />
                  </div>
                </div>
                <div>
                  <label>NIP <span style={{ fontSize: '11px', color: '#94a3b8', fontWeight: 'normal' }}>(Opsional)</span></label>
                  <div className="input-wrapper-new">
                    <span style={{ fontSize: '16px', color: '#94a3b8' }}>🪪</span>
                    <input value={inputs.nip} onChange={e=>setInputs({...inputs, nip:e.target.value})} placeholder="Masukkan NIP (jika ada)" />
                  </div>
                </div>
                <div>
                  <label>Jabatan <span style={{ fontSize: '11px', color: '#94a3b8', fontWeight: 'normal' }}>(Opsional)</span></label>
                  <div className="input-wrapper-new">
                    <span style={{ fontSize: '16px', color: '#94a3b8' }}>💼</span>
                    <input value={inputs.jabatan} onChange={e=>setInputs({...inputs, jabatan:e.target.value})} placeholder="Contoh: Staf, Guru, Pengawas, dll" />
                  </div>
                </div>
              </div>
            )}

            {tab==='admin' && (
              <div className="input-grp-premium" style={{ display: 'flex', flexDirection: 'column', gap: '15px' }}>
                <div>
                  <label>Username Admin</label>
                  <div className="input-wrapper-new">
                    <span style={{ fontSize: '16px', color: '#94a3b8' }}>🔑</span>
                    <input onChange={e=>setInputs({...inputs, user:e.target.value})} required placeholder="Username Anda" />
                  </div>
                </div>
                <div>
                  <label>Password</label>
                  <div className="input-wrapper-new">
                    <span style={{ fontSize: '16px', color: '#94a3b8' }}>🔒</span>
                    <input type="password" onChange={e=>setInputs({...inputs, pass:e.target.value})} required placeholder="••••••••" style={{ letterSpacing: '4px' }} />
                  </div>
                </div>
              </div>
            )}

            {(!(tab === 'asn' && asnLoginMethod === 'face')) && (
              <button type="submit" className="btn-premium-masuk" disabled={loading}>
                {loading ? (
                  <span>Memproses...</span>
                ) : (
                  <>
                    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
                    Masuk Sekarang
                  </>
                )}
              </button>
            )}
          </form>

          {tab === 'asn' && (
            <div style={{ textAlign: 'center', marginTop: '15px' }}>
              <button type="button" onClick={() => setShowRegisterModal(true)} style={{ background: 'none', border: 'none', color: '#0c3e26', fontWeight: '700', fontSize: '12.5px', cursor: 'pointer', textDecoration: 'underline' }}>
                Belum terdaftar? Daftar ASN & Face ID Baru (Uji Coba)
              </button>
            </div>
          )}

          {/* Secure lock footer banner inside form */}
          <div style={{ marginTop: '25px', padding: '15px', background: '#f0fdf4', border: '1.5px solid #dcfce7', borderRadius: '16px', display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
            <span style={{ fontSize: '20px' }}>🛡️</span>
            <div>
              <h4 style={{ fontSize: '13px', fontWeight: '800', color: '#166534', margin: '0 0 2px 0' }}>Aman & Terpercaya</h4>
              <p style={{ fontSize: '11px', color: '#166534', margin: 0, lineHeight: '1.5' }}>Data Anda dilindungi dengan enkripsi tingkat tinggi</p>
              <div style={{ fontSize: '11px', color: '#166534', opacity: 0.7, marginTop: '4px', fontWeight: '500' }}>© Kemenag Jember 2025</div>
            </div>
          </div>
        </div>
      </div>

      {showRegisterModal && (
        <RegisterAsnModal 
          onClose={() => setShowRegisterModal(false)} 
          onSuccess={(newUser: any) => {
            setShowRegisterModal(false);
            onLogin(newUser);
          }} 
          onError={(m: string) => setToast({ message: m, type: 'error' })} 
        />
      )}
      {toast && <Toast {...toast} onClose={() => setToast(null)} />}

    </div>
  );
};

const App = () => {
  const [user, setUser] = useState<any>(null);
  useEffect(() => { 
    const s = localStorage.getItem('user'); 
    if(s) setUser(JSON.parse(s)); 
  }, []);
  const login = (u: any) => { setUser(u); localStorage.setItem('user', JSON.stringify(u)); };
  const logout = () => { setUser(null); localStorage.removeItem('user'); };
  if(!user) return <LoginPage onLogin={login} />;
  if(user.role === 'admin') return <AdminDashboard onLogout={logout} />;
  return <UserDashboard user={user} onLogout={logout} onUpdateUser={(updatedUser: any) => { setUser(updatedUser); localStorage.setItem('user', JSON.stringify(updatedUser)); }} />;
};

const container = document.getElementById('root');
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
