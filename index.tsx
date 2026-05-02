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
    return snap.docs.map(d => d.data());
  }
};

const Toast = ({ message, type, onClose }: { message: string, type: string, onClose: () => void }) => {
  useEffect(() => { const t = setTimeout(onClose, 3000); return () => clearTimeout(t); }, [onClose]);
  return <div className={`toast toast-${type}`}>{message}</div>;
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
                <div style={{flex:1}}><label>Radius (m)</label><input type="number" value={newEvent.radius} onChange={e=>setNewEvent({...newEvent, radius:e.target.value})} /></div>
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
                  <div style={{flex:1}}><label>Radius (m)</label><input type="number" value={editingEvent.radius} onChange={e=>setEditingEvent({...editingEvent, radius:e.target.value})} /></div>
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

const UserDashboard = ({ user, onLogout }: any) => {
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
            return (
              <div key={ev.id} style={{padding:'15px 0', borderBottom:'1px solid #eee', display:'flex', justifyContent:'space-between', alignItems:'center'}}>
                <div style={{flex: 1, paddingRight: 10}}>
                  <div style={{fontWeight: 600}}>{ev.name}</div>
                  <small style={{color:'#666', display:'block'}}>{ev.locationName}</small>
                  {isAlreadyAttended && (
                    <span className="status-checked">
                      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
                      Sudah Hadir
                    </span>
                  )}
                </div>
                <div>
                  {isAlreadyAttended ? (
                    <span className="badge badge-gray" style={{padding: '8px 12px'}}>Selesai</span>
                  ) : (
                    <button className="btn btn-primary" style={{width:'auto', padding:'8px 20px', marginTop:0}} onClick={()=>setSelected(ev)}>Hadir</button>
                  )}
                </div>
              </div>
            );
          })
        )}
      </div>

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
  const [nip, setNip] = useState('');
  const [suggestions, setSuggestions] = useState<any[]>([]);
  const [inputs, setInputs] = useState({ name: '', instansi: '', user: '', pass: '' });
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
        onLogin({ id: 'guest_'+Date.now(), name: inputs.name, instansi: inputs.instansi, role: 'umum' });
      } else {
        const snap = await db.collection("config").doc("admin").get();
        const admin = snap.data();
        if(admin && admin.username === inputs.user && admin.password === inputs.pass) onLogin({ name: 'Admin', role: 'admin' });
        else throw new Error("Username/Password Salah");
      }
    } catch(err: any) { setToast({ message: err.message, type: 'error' }); }
    setLoading(false);
  };

  return (
    <div style={{minHeight:'100vh', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', padding:20}}>
      {toast && <Toast {...toast} onClose={()=>setToast(null)} />}
      <img src="https://www.freepnglogos.com/uploads/logo-kemenag-png/logo-kementerian-agama-gambar-logo-depag-png-0.png" className="main-logo" />
      <h1 style={{color:'#27AE60'}}>KJati</h1>
      <p style={{marginBottom:25, color:'#666', fontSize:13}}>Kemenag Jember Absensi Terintegrasi</p>
      <div className="card" style={{width:'100%', maxWidth:400}}>
        <div className="tabs">
          <button className={`tab ${tab==='asn'?'active':''}`} onClick={()=>setTab('asn')}>ASN</button>
          <button className={`tab ${tab==='umum'?'active':''}`} onClick={()=>setTab('umum')}>UMUM</button>
          <button className={`tab ${tab==='admin'?'active':''}`} onClick={()=>setTab('admin')}>ADMIN</button>
        </div>
        <form onSubmit={submit}>
          {tab==='asn' && (
            <div className="relative">
              <label>NIP Pegawai</label>
              <input value={nip} onChange={e=>handleNip(e.target.value)} required placeholder="Contoh: 1980..." />
              {suggestions.length > 0 && (
                <ul className="suggestions">
                  {suggestions.map((s,i) => <li key={i} className="suggestion-item" onClick={()=>{setNip(s.nip); setSuggestions([]);}}>{s.name} ({s.nip})</li>)}
                </ul>
              )}
            </div>
          )}
          {tab==='umum' && (
            <React.Fragment>
              <label>Nama Lengkap</label><input onChange={e=>setInputs({...inputs, name:e.target.value})} required placeholder="Nama Peserta" />
              <label>Instansi</label><input onChange={e=>setInputs({...inputs, instansi:e.target.value})} required placeholder="Asal Instansi" />
            </React.Fragment>
          )}
          {tab==='admin' && (
            <React.Fragment>
              <label>Username</label><input onChange={e=>setInputs({...inputs, user:e.target.value})} required />
              <label>Password</label><input type="password" onChange={e=>setInputs({...inputs, pass:e.target.value})} required />
            </React.Fragment>
          )}
          <button className="btn btn-primary" disabled={loading} style={{marginTop:20}}>{loading ? 'Memproses...' : 'Masuk'}</button>
        </form>
      </div>
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
  return <UserDashboard user={user} onLogout={logout} />;
};

const container = document.getElementById('root');
if (container) {
  const root = createRoot(container);
  root.render(<App />);
}
