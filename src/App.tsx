/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/*
  -------------------------------------------------------------------------
  🚀 วิธีย้ายโค้ดไป GitHub และเชื่อมต่อ Google Sheets (ภาษาไทย)
  -------------------------------------------------------------------------
  1. ใน AI Studio: เลือกเมนู 'Settings' -> 'Export to GitHub' เพื่อสร้าง Repository
  2. ใน GitHub: ไปที่ Settings > Pages เลือก 'GitHub Actions' เป็นแหล่งการ Deploy
  3. ใน Google Apps Script (GAS): 
     - สร้างไฟล์ชื่อ 'Dashboard.gs' และวางโค้ดด้านล่างนี้
     - กด 'Deploy' -> 'New Deployment' เลือก 'Web App'
     - ตั้งค่า 'Who has access' เป็น 'Anyone' 
     - ก๊อปปี้ 'Web App URL' มาใส่ในช่อง Input บนหน้าเว็บของคุณ

  --- [คัดลอกส่วนนี้ไปใส่ใน Apps Script] ---
  function doGet(e) {
    if (e.parameter.action === "getData") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("GeneralInfo") || ss.getSheets()[0];
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const result = data.slice(1).map(row => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  function updateMemberData(rowObj) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("GeneralInfo") || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const idKey = "ID Number"; // หรือชื่อหัวตารางที่คุณใช้
    const idIndex = data[0].indexOf(idKey);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] == rowObj[idKey]) {
        const headers = data[0];
        const newRow = headers.map(h => rowObj[h]);
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([newRow]);
        return "SUCCESS";
      }
    }
  }
*/

import React, { useState, useEffect, useMemo } from 'react';
import { 
  LayoutDashboard, 
  Settings, 
  UserRound, 
  BarChart3, 
  Lock, 
  LogOut, 
  ShieldCheck,
  Activity,
  CalendarDays,
  Search,
  CheckCircle2,
  HeartPulse,
  BookOpen,
  GraduationCap,
  Globe,
  Database,
  ExternalLink,
  ChevronRight,
  Stethoscope,
  Save,
  RefreshCw,
  Link2
} from 'lucide-react';
import { 
  BarChart, 
  Bar, 
  XAxis, 
  YAxis, 
  CartesianGrid, 
  Tooltip, 
  ResponsiveContainer,
  PieChart, 
  Pie, 
  Cell
} from 'recharts';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

const BLUE_PALETTE = ['#0f172a', '#1e40af', '#3b82f6', '#60a5fa', '#93c5fd'];

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [password, setPassword] = useState('');
  const [showLogin, setShowLogin] = useState(false);
  const [view, setView] = useState<'dashboard' | 'admin'>('dashboard');
  const [sheetData, setSheetData] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [gasUrl, setGasUrl] = useState(localStorage.getItem('gas_url') || ''); 
  const [isRefreshing, setIsRefreshing] = useState(false);

  // ฟังก์ชันดึงข้อมูลจาก GAS Web App
  const fetchData = async (url: string) => {
    if (!url) {
      setLoading(false);
      // Mock data ถ้าไม่มี URL
      setSheetData([
        { 'ID Number': '001', 'Name': 'คุณยายสมศรี ใจดี', 'Age': 68, 'Health Conditions': 'เบาหวาน', 'Phone': '081-xxx-xxxx' },
        { 'ID Number': '002', 'Name': 'คุณตาบุญมี มีสุข', 'Age': 72, 'Health Conditions': 'ปกติ', 'Phone': '089-xxx-xxxx' },
      ]);
      return;
    }

    setLoading(true);
    try {
      const response = await fetch(`${url}?action=getData`);
      const result = await response.json();
      setSheetData(result);
      localStorage.setItem('gas_url', url);
    } catch (err) {
      console.error("Fetch Error:", err);
      alert("ไม่สามารถเชื่อมต่อกับ Google Sheets ได้ โปรดตรวจสอบ Web App URL");
    } finally {
      setLoading(false);
      setIsRefreshing(false);
    }
  };

  useEffect(() => {
    const win = window as any;
    if (win.google?.script?.run) {
      // ถ้ารันใน GAS โดยตรง (doGet)
      win.google.script.run
        .withSuccessHandler((result: any[]) => {
          setSheetData(result);
          setLoading(false);
        })
        .getSheetData();
    } else {
      // ถ้ารันภายนอก (GitHub Pages)
      fetchData(gasUrl);
    }
  }, []);

  const stats = useMemo(() => {
    if (!sheetData.length) return null;
    const total = sheetData.length;
    const chronic = sheetData.filter(d => d['Health Conditions'] || d['โรคประจำตัว'] || d['สุขภาพ']).length;
    
    const ageGroups = { '60-70 ปี': 0, '71-80 ปี': 0, '81 ปีขึ้นไป': 0 };
    sheetData.forEach(d => {
      const age = Number(d['Age'] || d['อายุ']);
      if (age >= 81) ageGroups['81 ปีขึ้นไป']++;
      else if (age >= 71) ageGroups['71-80 ปี']++;
      else ageGroups['60-70 ปี']++;
    });

    return { 
      total, 
      chronic, 
      ageChart: Object.entries(ageGroups).map(([name, value]) => ({ name, value })) 
    };
  }, [sheetData]);

  const handleUpdate = (row: any) => {
    const win = window as any;
    if (win.google?.script?.run) {
      win.google.script.run.updateMemberData(row);
    } else {
      alert("⚠️ ขณะนี้ระบบรองรับการแก้ไขผ่านหน้า Web App ของ Google Apps Script โดยตรงเท่านั้น (GitHub Mode: Read-Only)");
    }
    setSheetData(prev => prev.map(r => r['ID Number'] === row['ID Number'] ? row : r));
  };

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-slate-900 font-sans antialiased">
      {/* Sidebar - Elite Dark Blue */}
      <aside className="fixed left-0 top-0 h-full w-24 bg-[#0F172A] flex flex-col items-center py-10 z-50">
        <div className="mb-14">
          <div className="w-14 h-14 bg-gradient-to-br from-blue-500 to-indigo-700 rounded-2xl flex items-center justify-center shadow-2xl shadow-blue-500/20">
            <GraduationCap className="text-white" size={32} />
          </div>
        </div>
        
        <div className="flex-1 flex flex-col gap-10">
          <button onClick={() => setView('dashboard')} className={cn("p-4 rounded-2xl transition-all", view === 'dashboard' ? "bg-white/10 text-blue-400 shadow-inner" : "text-slate-500 hover:text-white")}>
            <LayoutDashboard size={28} />
          </button>
          <button onClick={() => isAuthenticated ? setView('admin') : setShowLogin(true)} className={cn("p-4 rounded-2xl transition-all", view === 'admin' ? "bg-white/10 text-blue-400 shadow-inner" : "text-slate-500 hover:text-white")}>
            <Database size={28} />
          </button>
        </div>

        <button onClick={() => isAuthenticated ? setIsAuthenticated(false) : setShowLogin(true)} className="mt-auto p-4 text-slate-500 hover:text-rose-400 transition-colors">
          {isAuthenticated ? <LogOut size={28} /> : <Lock size={28} />}
        </button>
      </aside>

      {/* Main Content */}
      <main className="pl-24 min-h-screen pb-20">
        <header className="px-12 py-10 bg-white/80 backdrop-blur-xl border-b border-slate-200 sticky top-0 z-40 flex justify-between items-end">
          <div className="space-y-1">
            <div className="flex items-center gap-3">
               <h1 className="text-4xl font-black text-slate-900 tracking-tighter">ELITE AGING</h1>
               <span className="text-[10px] font-bold bg-blue-600 text-white px-3 py-1 rounded-full uppercase tracking-widest">Dashboard v3</span>
            </div>
            <p className="text-slate-400 text-sm font-medium flex items-center gap-2">
               <Globe className="text-blue-500" size={16} /> 
               ระบบบริหารจัดการข้อมูลโรงเรียนผู้สูงอายุ
            </p>
          </div>
          
          <div className="flex gap-6 items-center">
             <div className="flex items-center gap-3 bg-slate-50 border border-slate-200 p-3 rounded-2xl">
                <Link2 className="text-slate-400" size={16} />
                <input 
                  className="bg-transparent text-xs font-bold outline-none w-48 text-slate-600"
                  placeholder="URL ของ Web App จาก GAS"
                  value={gasUrl}
                  onChange={(e) => setGasUrl(e.target.value)}
                  onBlur={() => fetchData(gasUrl)}
                />
                <button onClick={() => fetchData(gasUrl)} className="p-1 hover:text-blue-600">
                  <RefreshCw size={14} className={isRefreshing ? "animate-spin" : ""} />
                </button>
             </div>
             <div className="text-right">
                <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">การเชื่อมต่อ</span>
                <div className="flex items-center gap-2 justify-end mt-1">
                  <div className={cn("w-2.5 h-2.5 rounded-full shadow-[0_0_10px_rgba(16,185,129,0.7)]", gasUrl ? "bg-emerald-500 animate-pulse" : "bg-slate-300")} />
                  <span className="text-sm font-bold text-slate-900">{gasUrl ? "เชื่อมต่อข้อมูลจริง" : "โหมดทดสอบ"}</span>
                </div>
             </div>
          </div>
        </header>

        {loading ? (
          <div className="flex flex-col items-center justify-center h-[60vh] gap-6">
             <div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin shadow-xl" />
             <p className="text-xs font-bold text-slate-400 uppercase tracking-widest animate-pulse">กำลังดึงข้อมูลจาก GOOGLE SHEETS...</p>
          </div>
        ) : (
          <div className="p-12 max-w-7xl mx-auto space-y-12">
            {view === 'dashboard' ? (
              <>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-8">
                  <div className="bg-white p-10 rounded-[2.5rem] border border-slate-200 shadow-sm transition-all hover:shadow-xl hover:-translate-y-1 group">
                     <p className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">จำนวนนักเรียนทั้งหมด</p>
                     <div className="flex justify-between items-end">
                        <h2 className="text-5xl font-black text-slate-900">{stats?.total} <span className="text-lg font-bold text-slate-400">คน</span></h2>
                        <UserRound className="text-blue-500 opacity-20 group-hover:opacity-100 transition-opacity" size={40} />
                     </div>
                  </div>
                  <div className="bg-white p-10 rounded-[2.5rem] border border-slate-200 shadow-sm transition-all hover:shadow-xl hover:-translate-y-1">
                     <p className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">กลุ่มอายุหลัก</p>
                     <div className="flex justify-between items-end">
                        <h2 className="text-5xl font-black text-slate-900">{stats?.ageChart[0]?.value}</h2>
                        <span className="text-xs font-bold text-slate-400 mb-2">ช่วง 60-70 ปี</span>
                     </div>
                  </div>
                  <div className="bg-slate-900 p-10 rounded-[2.5rem] text-white shadow-2xl relative overflow-hidden group">
                     <div className="absolute top-0 right-0 p-6 opacity-10">
                        <Activity size={80} />
                     </div>
                     <p className="text-xs font-bold text-blue-400 uppercase tracking-widest mb-2">ภาวะสุขภาพที่ต้องเฝ้าระวัง</p>
                     <h2 className="text-5xl font-black">{stats?.chronic} <span className="text-lg font-bold text-slate-500">ราย</span></h2>
                     <p className="text-[10px] font-bold text-slate-500 mt-2 uppercase">Detected Chronic Conditions</p>
                  </div>
                </div>

                <div className="grid grid-cols-1 lg:grid-cols-2 gap-10">
                   <div className="bg-white p-10 rounded-[3rem] border border-slate-200 shadow-sm">
                      <div className="flex justify-between items-center mb-10">
                         <h3 className="text-xl font-bold flex items-center gap-3">
                           <BarChart3 className="text-blue-600" size={24} />
                           สถิติจำนวนนักเรียนตามช่วงอายุ
                         </h3>
                         <ChevronRight className="text-slate-300" />
                      </div>
                      <div className="h-[300px]">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={stats?.ageChart}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="name" axisLine={false} tickLine={false} fontSize={12} dy={10} />
                            <YAxis axisLine={false} tickLine={false} fontSize={12} hide />
                            <Tooltip cursor={{fill: '#f8fafc'}} contentStyle={{borderRadius: '20px', border: 'none', boxShadow: '0 25px 50px -12px rgb(0 0 0 / 0.1)'}} />
                            <Bar dataKey="value" fill="#2563eb" radius={[12, 12, 12, 12]} barSize={60} />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                   </div>

                   <div className="bg-[#0F172A] p-10 rounded-[3rem] text-white relative">
                      <h3 className="text-xl font-bold mb-10 flex items-center gap-3">
                        <HeartPulse className="text-blue-400" size={24} />
                        สัดส่วนข้อมูลเชิงสุขภาพ
                      </h3>
                      <div className="flex items-center">
                        <div className="h-[250px] flex-1">
                          <ResponsiveContainer width="100%" height="100%">
                            <PieChart>
                              <Pie 
                                data={stats?.ageChart} 
                                innerRadius={60} 
                                outerRadius={90} 
                                paddingAngle={10} 
                                dataKey="value"
                                stroke="none"
                              >
                                {stats?.ageChart.map((_, i) => <Cell key={i} fill={BLUE_PALETTE[i % BLUE_PALETTE.length]} />)}
                              </Pie>
                              <Tooltip contentStyle={{backgroundColor: '#1e293b', border: 'none', borderRadius: '12px'}} />
                            </PieChart>
                          </ResponsiveContainer>
                        </div>
                        <div className="space-y-4 pr-10">
                           {stats?.ageChart.map((entry, i) => (
                             <div key={i} className="flex items-center gap-3">
                                <div className="w-3 h-3 rounded-full" style={{backgroundColor: BLUE_PALETTE[i % BLUE_PALETTE.length]}} />
                                <span className="text-xs font-bold text-slate-400">{entry.name}</span>
                             </div>
                           ))}
                        </div>
                      </div>
                   </div>
                </div>

                <div className="bg-white rounded-[3rem] border border-slate-200 overflow-hidden shadow-sm">
                   <div className="px-12 py-10 bg-slate-50/50 border-b border-slate-100 flex justify-between items-center">
                      <h3 className="text-xl font-bold flex items-center gap-3 text-slate-900">
                         <BookOpen className="text-blue-600" size={24} />
                         รายชื่อนักเรียนล่าสุด (จาก Google Sheets)
                      </h3>
                      <button onClick={() => setView('admin')} className="text-xs font-bold text-blue-600 uppercase tracking-widest hover:underline flex items-center gap-2">
                        จัดการข้อมูลหลังบ้าน <ExternalLink size={14} />
                      </button>
                   </div>
                   <div className="overflow-x-auto">
                     <table className="w-full text-left">
                       <thead>
                         <tr className="bg-white border-b border-slate-100 text-[10px] font-black text-slate-400 uppercase tracking-[0.2em]">
                           <th className="px-12 py-6">ชื่อ-นามสกุล</th>
                           <th className="px-12 py-6">อายุ</th>
                           <th className="px-12 py-6">รหัสประจำตัว</th>
                           <th className="px-12 py-6">ภาวะสุขภาพ</th>
                         </tr>
                       </thead>
                       <tbody className="divide-y divide-slate-50">
                         {sheetData.slice(0, 8).map((row, i) => (
                           <tr key={i} className="hover:bg-blue-50/30 transition-colors">
                             <td className="px-12 py-6">
                                <p className="text-sm font-black text-slate-900">{row['Name'] || row['ชื่อ-นามสกุล']}</p>
                                <p className="text-[10px] text-slate-400 font-bold">{row['Phone'] || row['เบอร์โทรศัพท์'] || "ไม่มีข้อมูลติดต่อ"}</p>
                             </td>
                             <td className="px-12 py-6 text-sm font-bold text-slate-500">{row['Age'] || row['อายุ']} ปี</td>
                             <td className="px-12 py-6 text-xs font-mono text-slate-400">{row['ID Number'] || row['รหัส']}</td>
                             <td className="px-12 py-6">
                                {row['Health Conditions'] || row['โรคประจำตัว'] || row['สุขภาพ'] ? (
                                   <div className="bg-rose-50 border border-rose-100 text-rose-600 text-[10px] font-black px-3 py-1 rounded-full inline-block">
                                      {row['Health Conditions'] || row['โรคประจำตัว'] || row['สุขภาพ']}
                                   </div>
                                ) : (
                                   <div className="bg-emerald-50 border border-emerald-100 text-emerald-600 text-[10px] font-black px-3 py-1 rounded-full inline-block">
                                      ปกติ
                                   </div>
                                )}
                             </td>
                           </tr>
                         ))}
                       </tbody>
                     </table>
                   </div>
                </div>
              </>
            ) : (
              <div className="space-y-10 animate-in fade-in duration-700">
                <div className="flex justify-between items-end">
                   <div className="space-y-2">
                      <h2 className="text-5xl font-black text-slate-900 tracking-tighter uppercase italic">Admin Terminal</h2>
                      <p className="text-slate-400 font-medium">แก้ไขและจัดการข้อมูลดิบจาก GeneralInfo</p>
                   </div>
                   <div className="relative group">
                      <Search className="absolute left-5 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-600 transition-colors" size={20} />
                      <input 
                         className="pl-14 pr-10 py-5 bg-white border border-slate-200 rounded-[2rem] w-[400px] shadow-sm outline-none focus:ring-4 focus:ring-blue-600/5 transition-all text-sm font-bold"
                         placeholder="ค้นหาตามชื่อ หรือรหัส..."
                         value={search}
                         onChange={(e) => setSearch(e.target.value)}
                      />
                   </div>
                </div>

                <div className="bg-white rounded-[3rem] border border-slate-200 shadow-2xl overflow-hidden">
                   <div className="overflow-x-auto max-h-[650px]">
                      <table className="w-full text-left">
                         <thead className="bg-[#0F172A] text-white sticky top-0 z-10">
                            <tr>
                               {sheetData[0] && Object.keys(sheetData[0]).map(h => (
                                 <th key={h} className="px-10 py-8 text-[10px] font-black uppercase tracking-[0.2em] opacity-40 whitespace-nowrap">{h}</th>
                               ))}
                               <th className="px-10 py-8 text-[10px] font-black uppercase tracking-[0.2em] opacity-40">บันทึก</th>
                            </tr>
                         </thead>
                         <tbody className="divide-y divide-slate-100">
                            {sheetData.filter(r => Object.values(r).some(v => String(v).includes(search))).map((row, i) => (
                              <tr key={i} className="hover:bg-slate-50 transition-colors">
                                 {Object.entries(row).map(([k, v]: [string, any], j) => (
                                   <td key={j} className="px-10 py-6">
                                      <input 
                                         className="bg-transparent border-b border-transparent hover:border-slate-300 focus:border-blue-600 outline-none text-sm w-full py-2 font-bold text-slate-700"
                                         defaultValue={v}
                                         onBlur={(e) => { row[k] = e.target.value; }}
                                      />
                                   </td>
                                 ))}
                                 <td className="px-10 py-6">
                                    <button 
                                       onClick={() => handleUpdate(row)}
                                       className="p-4 bg-blue-50 text-blue-600 rounded-2xl hover:bg-blue-600 hover:text-white transition-all shadow-lg shadow-blue-500/10"
                                    >
                                       <Save size={24} />
                                    </button>
                                 </td>
                              </tr>
                            ))}
                         </tbody>
                      </table>
                   </div>
                </div>
              </div>
            )}
          </div>
        )}
      </main>

      {/* Login Shield */}
      {showLogin && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-3xl z-[100] flex items-center justify-center p-8">
           <div className="bg-white rounded-[4rem] p-16 max-w-md w-full shadow-2xl relative overflow-hidden">
              <div className="absolute top-0 right-0 p-10 opacity-5 -rotate-12 translate-x-4 -translate-y-4">
                 <ShieldCheck size={200} />
              </div>
              <div className="relative z-10 flex flex-col items-center">
                 <div className="w-24 h-24 bg-blue-600 rounded-[2rem] flex items-center justify-center mb-10 shadow-2xl shadow-blue-600/30">
                   <Lock className="text-white" size={40} />
                 </div>
                 <h2 className="text-3xl font-black mb-3 tracking-tighter">เข้าสู่ระบบจัดการ</h2>
                 <p className="text-slate-400 text-sm text-center mb-12 font-bold uppercase tracking-widest">กรุณายืนยันรหัสเข้าถึงแผงควบคุม</p>
                 <form 
                   onSubmit={(e) => { 
                      e.preventDefault(); 
                      if(password === '1111') { setIsAuthenticated(true); setShowLogin(false); setView('admin'); } 
                      else alert('รหัสผ่านไม่ถูกต้อง'); 
                   }} 
                   className="w-full space-y-6"
                  >
                    <input 
                      type="password" placeholder="••••" autoFocus
                      className="w-full px-8 py-6 bg-slate-100 border border-slate-200 rounded-[2.5rem] text-center text-5xl font-black tracking-[0.5em] focus:ring-8 focus:ring-blue-600/5 outline-none transition-all placeholder:tracking-normal"
                      onChange={(e) => setPassword(e.target.value)}
                    />
                    <div className="flex gap-4 pt-4">
                       <button type="button" onClick={() => setShowLogin(false)} className="flex-1 py-5 font-black text-slate-400 uppercase text-xs tracking-[0.2em] hover:text-slate-600 transition-colors">ยกเลิก</button>
                       <button type="submit" className="flex-1 py-6 bg-[#0F172A] text-white rounded-[2.5rem] font-black shadow-2xl hover:bg-black transition-all text-sm uppercase tracking-[0.3em]">ยืนยัน</button>
                    </div>
                 </form>
              </div>
           </div>
        </div>
      )}
    </div>
  );
}
