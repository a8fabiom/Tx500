import React, { useState, useEffect, useMemo, useRef } from 'react';
import { initializeApp } from 'firebase/app';
import {
  getFirestore,
  collection,
  doc,
  addDoc,
  onSnapshot,
  updateDoc,
  deleteDoc,
  serverTimestamp,
  query,
  orderBy
} from 'firebase/firestore';
import {
  getAuth,
  signInAnonymously,
  signInWithCustomToken,
  onAuthStateChanged
} from 'firebase/auth';
import {
  Plus,
  Trash2,
  ExternalLink,
  X,
  Edit2,
  Image as ImageIcon,
  Download,
  Upload,
  Euro,
  Terminal,
  StickyNote,
  Search,
  ArrowUpRight,
  Loader2,
  CheckCircle2,
  Circle,
  FileSpreadsheet,
  Calendar,
  Type,
  Coins,
  ListChecks,
  Sparkles,
  Save,
  Filter,
  Store,
  Zap,
  Link2
} from 'lucide-react';

// --- FIREBASE CONFIGURATION ---
const firebaseConfig = JSON.parse(__firebase_config);
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const appId = typeof __app_id !== 'undefined' ? __app_id : 'efs-rx500-restauro-v9';

const CATEGORIES = ["Lacagem", "Decapagem", "Cromagem", "Reparação", "Novo", "Concluído", "Pendente"];

const FIELD_MAP = {
  name: { label: 'Nome', keys: ['peça', 'nome', 'item', 'title', 'produto'] },
  vendor: { label: 'Fornecedor', keys: ['vendedor', 'loja', 'fornecedor', 'publisher', 'vendor'] },
  url: { label: 'Link URL', keys: ['link', 'url', 'sítio', 'web'] },
  imageUrl: { label: 'Foto URL', keys: ['imagem', 'foto', 'image', 'thumb', 'photo'] },
  price: { label: 'Preço', keys: ['preço', 'valor', '€', 'price', 'cost'] },
  category: { label: 'Categoria', keys: ['cat', 'estado', 'fase', 'category'] },
  notes: { label: 'Notas', keys: ['nota', 'obs', 'descri', 'notes'] }
};

// --- HELPER COMPONENTS ---

const InlineInput = ({ value, onSave, type = "text", className = "", prefix = "" }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [val, setVal] = useState(value);

  const handleBlur = () => {
    setIsEditing(false);
    if (val !== value) onSave(val);
  };

  const handleKeyDown = (e) => {
    if (e.key === 'Enter') handleBlur();
    if (e.key === 'Escape') {
      setVal(value);
      setIsEditing(false);
    }
  };

  if (isEditing) {
    return (
      <input
        autoFocus
        type={type}
        className={`bg-blue-50 border-b border-blue-500 outline-none px-1 w-full ${className}`}
        value={val}
        onChange={(e) => setVal(e.target.value)}
        onBlur={handleBlur}
        onKeyDown={handleKeyDown}
        onClick={(e) => e.stopPropagation()}
      />
    );
  }

  return (
    <div
      onClick={(e) => { e.stopPropagation(); setIsEditing(true); }}
      className={`cursor-pointer hover:bg-slate-50 rounded px-1 transition-colors truncate ${className}`}
    >
      {type === "number" ? `${prefix}${Number(value || 0).toFixed(2)}€` : (value || '---')}
    </div>
  );
};

// --- MAIN APPLICATION ---

export default function App() {
  const [user, setUser] = useState(null);
  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedItems, setSelectedItems] = useState([]);
  const [editingItem, setEditingItem] = useState(null);
  const [isAdding, setIsAdding] = useState(false);
  const [expandedImage, setExpandedImage] = useState(null);
  const [noteEditor, setNoteEditor] = useState(null);
  const fileInputRef = useRef(null);

  const [showLogs, setShowLogs] = useState(false);
  const [logs, setLogs] = useState([]);

  const [searchTerm, setSearchTerm] = useState('');
  const [filterCategory, setFilterCategory] = useState('Todas');
  const [filterVendor, setFilterVendor] = useState('Todos');
  const [sortBy, setSortBy] = useState('date');
  const [sortOrder, setSortOrder] = useState('desc');

  // Import State
  const [importModalOpen, setImportModalOpen] = useState(false);
  const [importMode, setImportMode] = useState('excel');
  const [importType, setImportType] = useState('simple');
  const [bulkLinks, setBulkLinks] = useState('');
  const [importData, setImportData] = useState([]);
  const [importHeaders, setImportHeaders] = useState([]);
  const [mapping, setMapping] = useState({
    name: '', vendor: '', url: '', imageUrl: '', price: '', category: '', notes: ''
  });
  const [isProcessingImport, setIsProcessingImport] = useState(false);

  const [newItem, setNewItem] = useState({
    name: '', vendor: '', url: '', imageUrl: '', price: '', category: 'Pendente', notes: ''
  });
  const [isAutoFilling, setIsAutoFilling] = useState(false);

  const addLog = (message, type = 'info') => {
    const newLog = { id: Math.random().toString(36), timestamp: new Date().toLocaleTimeString(), message, type };
    setLogs(prev => [...prev.slice(-15), newLog]);
  };

  // External Lib Load (XLSX)
  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
    script.async = true;
    script.onload = () => addLog("SheetJS Engine Loaded", "success");
    document.head.appendChild(script);
  }, []);

  // Auth Effect
  useEffect(() => {
    const initAuth = async () => {
      try {
        if (typeof __initial_auth_token !== 'undefined' && __initial_auth_token) {
          await signInWithCustomToken(auth, __initial_auth_token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (err) { addLog(`Auth Error: ${err.message}`, "error"); }
    };
    initAuth();
    const unsubscribe = onAuthStateChanged(auth, setUser);
    return () => unsubscribe();
  }, []);

  // Data Sync Effect
  useEffect(() => {
    if (!user) return;
    const colRef = collection(db, 'artifacts', appId, 'public', 'data', 'peças');
    const q = query(colRef);
    const unsubscribe = onSnapshot(q, (snapshot) => {
      setItems(snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() })));
      setLoading(false);
    }, (error) => {
      addLog(`Firestore Sync Error: ${error.message}`, "error");
      setLoading(false);
    });
    return () => unsubscribe();
  }, [user]);

  // Derived Values
  const vendorsList = useMemo(() => {
    const list = items
      .map(i => i.vendor)
      .filter(v => v && String(v).trim() !== '')
      .map(v => String(v).trim());
    return ['Todos', ...new Set(list)].sort();
  }, [items]);

  const processedItems = useMemo(() => {
    let result = items.filter(item => {
      const catMatch = filterCategory === 'Todas' || item.category === filterCategory;
      const vendorMatch = filterVendor === 'Todos' || item.vendor === filterVendor;
      const searchMatch = !searchTerm ||
        (item.name || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
        (item.vendor || '').toLowerCase().includes(searchTerm.toLowerCase());
      return catMatch && vendorMatch && searchMatch;
    });

    return result.sort((a, b) => {
      let valA, valB;
      if (sortBy === 'price') {
        valA = Number(a.price) || 0;
        valB = Number(b.price) || 0;
      } else if (sortBy === 'name') {
        valA = (a.name || '').toLowerCase();
        valB = (b.name || '').toLowerCase();
      } else {
        valA = a.createdAt?.seconds || 0;
        valB = b.createdAt?.seconds || 0;
      }
      if (valA < valB) return sortOrder === 'asc' ? -1 : 1;
      if (valA > valB) return sortOrder === 'asc' ? 1 : -1;
      return 0;
    });
  }, [items, filterCategory, filterVendor, searchTerm, sortBy, sortOrder]);

  const stats = useMemo(() => {
    const targetItems = selectedItems.length > 0 ? items.filter(i => selectedItems.includes(i.id)) : processedItems;
    const total = targetItems.reduce((acc, curr) => acc + (Number(curr.price) || 0), 0);
    return { total, count: targetItems.length };
  }, [items, selectedItems, processedItems]);

  // Actions
  const toggleSelectAll = () => {
    if (selectedItems.length > 0) {
      setSelectedItems([]);
    } else {
      setSelectedItems(processedItems.map(i => i.id));
    }
  };

  const getMetadata = async (url) => {
    if (!url || !url.startsWith('http')) return null;
    try {
      const response = await fetch(`https://api.microlink.io?url=${encodeURIComponent(url)}`);
      const data = await response.json();
      if (data.status === 'success') {
        return {
          name: data.data.title || '',
          vendor: data.data.publisher || new URL(url).hostname.replace('www.', ''),
          imageUrl: data.data.image?.url || data.data.logo?.url || '',
        };
      }
    } catch (err) { addLog(`Metadata extraction failed for: ${url}`, "error"); }
    return null;
  };

  const handleUrlChange = async (url) => {
    if (isAdding) {
      setNewItem(prev => ({ ...prev, url }));
      if (url.startsWith('http')) {
        setIsAutoFilling(true);
        const meta = await getMetadata(url);
        if (meta) {
          setNewItem(prev => ({
            ...prev,
            name: prev.name || meta.name,
            vendor: prev.vendor || meta.vendor,
            imageUrl: prev.imageUrl || meta.imageUrl
          }));
        }
        setIsAutoFilling(false);
      }
    } else {
      setEditingItem(prev => ({ ...prev, url }));
    }
  };

  const saveNewItem = async () => {
    if (!user) return;
    try {
      await addDoc(collection(db, 'artifacts', appId, 'public', 'data', 'peças'), {
        ...newItem,
        price: Number(newItem.price) || 0,
        createdAt: serverTimestamp()
      });
      setNewItem({ name: '', vendor: '', url: '', imageUrl: '', price: '', category: 'Pendente', notes: '' });
      setIsAdding(false);
      addLog("Item added successfully.");
    } catch (err) { addLog(err.message, "error"); }
  };

  const updateItem = async (id, fields) => {
    if (!user) return;
    try {
      const { id: _, createdAt: __, ...data } = fields;
      await updateDoc(doc(db, 'artifacts', appId, 'public', 'data', 'peças', id), {
        ...data,
        price: Number(data.price) || 0
      });
      addLog(`Item ${id} updated.`);
    } catch (err) { addLog(err.message, "error"); }
  };

  const deleteItems = async () => {
    if (!user || selectedItems.length === 0) return;
    try {
      for (const id of selectedItems) {
        await deleteDoc(doc(db, 'artifacts', appId, 'public', 'data', 'peças', id));
      }
      setSelectedItems([]);
      addLog(`${selectedItems.length} items removed.`);
    } catch (err) { addLog(err.message, "error"); }
  };

  const exportToExcel = () => {
    if (!window.XLSX) return;
    const dataToExport = processedItems.map(i => ({
      Nome: i.name, Fornecedor: i.vendor, Preço: i.price, Categoria: i.category, Link: i.url, Notas: i.notes
    }));
    const ws = window.XLSX.utils.json_to_sheet(dataToExport);
    const wb = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(wb, ws, "Inventory");
    window.XLSX.writeFile(wb, `RX500_Export_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file || !window.XLSX) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = window.XLSX.read(bstr, { type: 'binary' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = window.XLSX.utils.sheet_to_json(ws);
      if (data.length > 0) {
        setImportData(data);
        const headers = Object.keys(data[0]);
        setImportHeaders(headers);
        const newMapping = {};
        Object.keys(FIELD_MAP).forEach(k => {
          const match = headers.find(h => FIELD_MAP[k].keys.some(pk => h.toLowerCase().includes(pk.toLowerCase())));
          newMapping[k] = match || '';
        });
        setMapping(newMapping);
        setImportMode('excel');
        setImportModalOpen(true);
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = null;
  };

  const processImport = async () => {
    setIsProcessingImport(true);
    let count = 0;
    const colRef = collection(db, 'artifacts', appId, 'public', 'data', 'peças');

    if (importMode === 'excel') {
      for (const row of importData) {
        try {
          let baseData = {
            name: row[mapping.name] || 'Unnamed Piece',
            vendor: row[mapping.vendor] || '',
            url: row[mapping.url] || '',
            imageUrl: row[mapping.imageUrl] || '',
            price: parseFloat(String(row[mapping.price] || '0').replace(',', '.')) || 0,
            category: row[mapping.category] || 'Pendente',
            notes: row[mapping.notes] || '',
            createdAt: serverTimestamp()
          };
          if (importType === 'auto' && baseData.url) {
            const meta = await getMetadata(baseData.url);
            if (meta) {
              baseData.name = meta.name || baseData.name;
              baseData.vendor = meta.vendor || baseData.vendor;
              baseData.imageUrl = meta.imageUrl || baseData.imageUrl;
            }
          }
          await addDoc(colRef, baseData);
          count++;
        } catch (e) { addLog(`Error importing row: ${e.message}`, "error"); }
      }
    } else {
      const links = bulkLinks.split('\n').map(l => l.trim()).filter(l => l.startsWith('http'));
      for (const link of links) {
        try {
          const meta = await getMetadata(link);
          await addDoc(colRef, {
            name: meta?.name || 'Item from Link',
            vendor: meta?.vendor || '',
            url: link,
            imageUrl: meta?.imageUrl || '',
            price: 0,
            category: 'Pendente',
            notes: '',
            createdAt: serverTimestamp()
          });
          count++;
        } catch (e) { addLog(`Error fetching link: ${link}`, "error"); }
      }
    }
    setImportModalOpen(false);
    setIsProcessingImport(false);
    setBulkLinks('');
    addLog(`Import complete: ${count} items added.`);
  };

  if (loading) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-[#FBFBFB] gap-4">
        <Loader2 className="animate-spin text-slate-950" size={32} />
        <p className="text-[10px] font-black uppercase tracking-widest text-slate-400 animate-pulse">Initializing RX500 Engine...</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#FBFBFB] text-slate-900 font-sans selection:bg-blue-100">

      {/* --- HEADER --- */}
      <header className="sticky top-0 z-50 bg-white/95 backdrop-blur-md border-b border-slate-100 px-4 py-3">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-slate-950 rounded-xl flex items-center justify-center shadow-lg shadow-slate-900/10">
              <span className="text-white text-[10px] font-black tracking-tighter">RX</span>
            </div>
            <div>
              <h1 className="text-[11px] font-black uppercase tracking-[0.25em] text-slate-950 leading-none">RX500 SPORT</h1>
              <p className="text-[7px] font-bold text-slate-400 uppercase tracking-widest mt-1">Restoration Manager v9.2</p>
            </div>
          </div>

          <div className="flex items-center gap-1.5">
            <button onClick={() => setShowLogs(!showLogs)} className={`p-2 rounded-xl transition-all ${showLogs ? 'text-blue-600 bg-blue-50' : 'text-slate-300 hover:text-slate-950'}`} title="System Logs"><Terminal size={16}/></button>
            <button onClick={exportToExcel} className="p-2 text-slate-300 hover:text-emerald-600 rounded-xl" title="Export Excel"><Download size={16}/></button>
            <button onClick={() => { setImportMode('links'); setImportModalOpen(true); }} className="p-2 text-slate-300 hover:text-blue-600 rounded-xl" title="Import via Links"><Link2 size={16}/></button>
            <label className="p-2 text-slate-300 hover:text-slate-950 cursor-pointer rounded-xl" title="Import Excel">
              <Upload size={16}/>
              <input ref={fileInputRef} type="file" accept=".xlsx, .xls" className="hidden" onChange={handleFileUpload} />
            </label>
            <div className="w-px h-6 bg-slate-100 mx-2" />
            <button onClick={() => setIsAdding(true)} className="bg-slate-950 text-white px-5 py-2.5 rounded-2xl text-[10px] font-black hover:bg-slate-800 transition-all flex items-center gap-2 shadow-xl shadow-slate-900/20 active:scale-95">
              <Plus size={12} strokeWidth={3} /> NEW ENTRY
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto p-4 pb-36">

        {/* LOG PANEL */}
        {showLogs && (
          <div className="mb-6 bg-slate-900 rounded-2xl p-4 font-mono text-[10px] text-slate-400 border border-slate-800 shadow-inner animate-in slide-in-from-top-4">
            <div className="flex items-center justify-between mb-3 border-b border-white/5 pb-2">
              <span className="flex items-center gap-2 text-blue-400 font-bold"><Terminal size={12}/> SYSTEM LOGS</span>
              <button onClick={() => setLogs([])} className="text-[8px] hover:text-white uppercase font-black">Clear</button>
            </div>
            {logs.length === 0 && <div className="italic opacity-30">Waiting for system events...</div>}
            <div className="max-h-[150px] overflow-y-auto space-y-1 custom-scrollbar">
              {logs.map(log => (
                <div key={log.id} className="flex gap-2">
                  <span className="opacity-30 shrink-0">[{log.timestamp}]</span>
                  <span className={log.type === 'error' ? 'text-red-400' : log.type === 'success' ? 'text-emerald-400' : 'text-slate-300'}>
                    {log.message}
                  </span>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* --- TOOLBAR --- */}
        <section className="mb-8 sticky top-20 z-40 bg-white/80 backdrop-blur-md p-2 rounded-3xl border border-slate-200 shadow-sm flex flex-col md:flex-row items-center gap-3">
          <div className="flex items-center bg-slate-50 rounded-2xl px-4 py-2 flex-grow min-w-0 border border-slate-100 group focus-within:ring-2 ring-slate-950/5 transition-all">
            <button
              onClick={toggleSelectAll}
              className={`p-1.5 rounded-lg transition-all ${selectedItems.length > 0 ? 'bg-slate-950 text-white' : 'text-slate-300 hover:text-slate-950'}`}
            >
              <ListChecks size={16} />
            </button>
            <div className="w-px h-4 bg-slate-200 mx-3" />
            <Search size={16} className="text-slate-300 group-focus-within:text-slate-950" />
            <input
              type="text"
              placeholder="Search parts, vendors, categories..."
              className="w-full pl-3 pr-1 py-1 bg-transparent border-none text-[11px] focus:ring-0 outline-none font-bold placeholder:font-normal placeholder:text-slate-300"
              value={searchTerm}
              onChange={e => setSearchTerm(e.target.value)}
            />
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <div className="relative">
              <div className={`p-3 rounded-2xl transition-all flex items-center gap-2 ${filterCategory !== 'Todas' ? 'bg-blue-50 text-blue-600 ring-1 ring-blue-100' : 'text-slate-400 bg-slate-50 border border-slate-100'}`}>
                <Filter size={16} />
                <span className="text-[10px] font-black uppercase tracking-tight">{filterCategory === 'Todas' ? 'Categories' : filterCategory}</span>
                <select className="absolute inset-0 opacity-0 cursor-pointer" value={filterCategory} onChange={e => setFilterCategory(e.target.value)}>
                  <option value="Todas">All Categories</option>
                  {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            </div>

            <div className="relative">
              <div className={`p-3 rounded-2xl transition-all flex items-center gap-2 ${filterVendor !== 'Todos' ? 'bg-purple-50 text-purple-600 ring-1 ring-purple-100' : 'text-slate-400 bg-slate-50 border border-slate-100'}`}>
                <Store size={16} />
                <span className="text-[10px] font-black uppercase tracking-tight truncate max-w-[80px]">{filterVendor === 'Todos' ? 'Vendors' : filterVendor}</span>
                <select className="absolute inset-0 opacity-0 cursor-pointer" value={filterVendor} onChange={e => setFilterVendor(e.target.value)}>
                  <option value="Todos">All Vendors</option>
                  {vendorsList.filter(v => v !== 'Todos').map(v => <option key={v} value={v}>{v}</option>)}
                </select>
              </div>
            </div>

            <div className="w-px h-6 bg-slate-100 mx-1" />

            <div className="flex items-center gap-1 bg-slate-50 p-1 rounded-2xl border border-slate-100">
              {[ { id: 'date', icon: Calendar }, { id: 'name', icon: Type }, { id: 'price', icon: Coins }].map(opt => (
                <button
                  key={opt.id}
                  onClick={() => setSortBy(opt.id)}
                  className={`p-2 rounded-xl transition-all ${sortBy === opt.id ? 'bg-white shadow-sm text-slate-950 ring-1 ring-slate-100' : 'text-slate-300 hover:text-slate-500'}`}
                >
                  <opt.icon size={14} />
                </button>
              ))}
              <button onClick={() => setSortOrder(prev => prev === 'asc' ? 'desc' : 'asc')} className="p-2 text-slate-400 hover:text-slate-950 transition-all">
                <ArrowUpRight size={14} strokeWidth={3} className={`transition-transform duration-500 ${sortOrder === 'desc' ? 'rotate-180' : ''}`} />
              </button>
            </div>

            {selectedItems.length > 0 && (
              <button onClick={deleteItems} className="p-3 bg-red-50 text-red-500 rounded-2xl hover:bg-red-500 hover:text-white transition-all shadow-sm border border-red-100 group">
                <Trash2 size={16} className="group-active:scale-90 transition-transform" />
              </button>
            )}
          </div>
        </section>

        {/* --- GRID --- */}
        <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-5 xl:grid-cols-6 gap-6">
          {processedItems.length === 0 && (
            <div className="col-span-full py-32 text-center border-2 border-dashed border-slate-100 rounded-[3rem]">
              <div className="inline-flex flex-col items-center gap-4">
                <div className="p-6 bg-slate-50 rounded-full">
                  <Search size={48} className="text-slate-200" strokeWidth={1} />
                </div>
                <div className="space-y-1">
                  <p className="text-[12px] font-black uppercase text-slate-400 tracking-widest">No pieces found</p>
                  <p className="text-[9px] font-bold text-slate-300 uppercase">Try adjusting your filters or search terms</p>
                </div>
              </div>
            </div>
          )}

          {processedItems.map(item => (
            <div
              key={item.id}
              className={`group relative bg-white rounded-[2.5rem] border transition-all duration-500 flex flex-col h-full overflow-hidden ${selectedItems.includes(item.id) ? 'border-slate-950 ring-4 ring-slate-950/5 shadow-2xl' : 'border-slate-100 hover:border-slate-300 hover:shadow-xl hover:-translate-y-1'}`}
            >
              {/* Image Section */}
              <div className="relative h-[150px] bg-[#FDFDFD] flex items-center justify-center p-6 cursor-zoom-in" onClick={() => item.imageUrl && setExpandedImage(item.imageUrl)}>
                {item.imageUrl ? (
                  <img src={item.imageUrl} className="max-w-full max-h-full object-contain mix-blend-multiply transition-transform group-hover:scale-110 duration-700" alt={item.name} />
                ) : (
                  <div className="flex flex-col items-center gap-2">
                    <ImageIcon size={32} className="text-slate-100" strokeWidth={1} />
                    <span className="text-[7px] font-black uppercase text-slate-200 tracking-widest">No Image</span>
                  </div>
                )}

                {/* Status Badge */}
                <div className="absolute top-4 left-4 flex flex-col gap-1 pointer-events-none">
                  <span className={`px-2 py-1 rounded-xl text-[7px] font-black uppercase tracking-tight shadow-sm border ${
                    item.category === 'Concluído' ? 'bg-emerald-500 text-white border-emerald-400' :
                    item.category === 'Novo' ? 'bg-blue-600 text-white border-blue-500' :
                    item.category === 'Reparação' ? 'bg-amber-500 text-white border-amber-400' :
                    'bg-slate-950 text-white border-slate-800'
                  }`}>
                    {item.category}
                  </span>
                </div>

                {/* Selection Overlay */}
                <button
                  onClick={(e) => { e.stopPropagation(); setSelectedItems(prev => prev.includes(item.id) ? prev.filter(i => i !== item.id) : [...prev, item.id]); }}
                  className="absolute top-4 right-4 p-2 bg-white/80 backdrop-blur rounded-xl border border-slate-100 shadow-sm opacity-0 group-hover:opacity-100 transition-all duration-300"
                >
                   {selectedItems.includes(item.id) ? <CheckCircle2 size={16} className="text-slate-950" /> : <Circle size={16} className="text-slate-200" />}
                </button>
              </div>

              {/* Data Section */}
              <div className="p-5 flex flex-col flex-grow gap-3">
                <div className="space-y-1">
                  <InlineInput
                    value={item.name}
                    className="text-[11px] font-black text-slate-900 leading-tight min-h-[1.2em]"
                    onSave={(newVal) => updateItem(item.id, { ...item, name: newVal })}
                  />
                  <div className="flex items-center gap-1">
                    <Store size={10} className="text-slate-200" />
                    <InlineInput
                      value={item.vendor || 'Unknown Vendor'}
                      className="text-[8px] text-slate-400 font-bold uppercase tracking-widest truncate"
                      onSave={(newVal) => updateItem(item.id, { ...item, vendor: newVal })}
                    />
                  </div>
                </div>

                <div className="mt-auto pt-3 border-t border-slate-50 flex items-end justify-between">
                  <div className="space-y-0.5">
                    <p className="text-[6px] font-black uppercase text-slate-300 tracking-widest">Price Point</p>
                    <InlineInput
                      value={item.price}
                      type="number"
                      className="text-[14px] font-black text-slate-950 tabular-nums"
                      onSave={(newVal) => updateItem(item.id, { ...item, price: newVal })}
                    />
                  </div>

                  <div className="flex items-center gap-1.5">
                    {item.url && (
                      <a
                        href={item.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        onClick={e => e.stopPropagation()}
                        className="p-2 bg-slate-50 hover:bg-blue-50 text-slate-400 hover:text-blue-600 rounded-xl transition-all"
                      >
                        <ExternalLink size={14} />
                      </a>
                    )}
                    <button
                      onClick={(e) => { e.stopPropagation(); setNoteEditor({ id: item.id, notes: item.notes || '' }); }}
                      className={`p-2 rounded-xl transition-all ${item.notes ? 'bg-amber-50 text-amber-600 border border-amber-100' : 'bg-slate-50 text-slate-400 opacity-0 group-hover:opacity-100'}`}
                    >
                      <StickyNote size={14} />
                    </button>
                    <button
                      onClick={(e) => { e.stopPropagation(); setEditingItem(item); }}
                      className="p-2 bg-slate-50 text-slate-400 hover:text-slate-950 hover:bg-slate-100 rounded-xl transition-all opacity-0 group-hover:opacity-100"
                    >
                      <Edit2 size={14} />
                    </button>
                  </div>
                </div>
              </div>
            </div>
          ))}
        </div>
      </main>

      {/* --- TOTALS BAR --- */}
      <div className="fixed bottom-8 left-1/2 -translate-x-1/2 z-50 w-full max-w-sm px-6">
        <div className="bg-slate-950/95 backdrop-blur-2xl text-white rounded-[2rem] p-4 flex items-center justify-between shadow-2xl border border-white/10 ring-1 ring-black/20 animate-in slide-in-from-bottom-8 duration-700">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-emerald-500/20 rounded-2xl flex items-center justify-center text-emerald-400 shadow-inner">
              <Euro size={20} strokeWidth={3} />
            </div>
            <div>
              <p className="text-[7px] font-black uppercase text-white/30 tracking-[0.3em] leading-none mb-1.5">TOTAL INVESTMENT</p>
              <h4 className="text-lg font-black tabular-nums tracking-tighter">
                {stats.total.toLocaleString('pt-PT', { style: 'currency', currency: 'EUR' })}
              </h4>
            </div>
          </div>
          <div className="text-right pl-6 border-l border-white/5 pr-2">
            <p className="text-[7px] font-black uppercase text-white/30 tracking-[0.3em] leading-none mb-1.5">PIECES</p>
            <p className="text-lg font-black tabular-nums">{stats.count}</p>
          </div>
        </div>
      </div>

      {/* --- MODALS --- */}

      {/* NOTE EDITOR */}
      {noteEditor && (
        <div className="fixed inset-0 z-[150] flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-sm animate-in fade-in" onClick={() => setNoteEditor(null)}>
          <div className="bg-white w-full max-w-md rounded-[2.5rem] p-8 shadow-2xl scale-in-center" onClick={e => e.stopPropagation()}>
            <div className="flex justify-between items-center mb-6">
              <div className="flex items-center gap-3">
                <div className="p-3 bg-amber-50 rounded-2xl text-amber-600"><StickyNote size={20} /></div>
                <div>
                  <h3 className="text-xs font-black uppercase text-slate-800 tracking-tight">Technical Notes</h3>
                  <p className="text-[8px] font-bold text-slate-400 uppercase tracking-widest mt-0.5">Specifications & Reminders</p>
                </div>
              </div>
              <button onClick={() => setNoteEditor(null)} className="p-2 hover:bg-slate-100 rounded-full transition-colors"><X size={20} /></button>
            </div>
            <textarea
              className="w-full h-48 p-5 bg-slate-50 rounded-3xl text-xs border border-slate-100 outline-none focus:border-amber-300 focus:ring-8 focus:ring-amber-50/50 transition-all resize-none font-medium leading-relaxed"
              placeholder="Enter technical details, part numbers, or fitting notes..."
              value={noteEditor.notes}
              onChange={(e) => setNoteEditor({ ...noteEditor, notes: e.target.value })}
              autoFocus
            />
            <button
              onClick={() => {
                updateItem(noteEditor.id, { notes: noteEditor.notes });
                setNoteEditor(null);
              }}
              className="w-full mt-6 bg-slate-950 text-white py-4 rounded-2xl text-[10px] font-black uppercase tracking-widest flex items-center justify-center gap-2 shadow-xl shadow-slate-950/30 hover:bg-slate-800 transition-all"
            >
              <Save size={16} /> Save Notes
            </button>
          </div>
        </div>
      )}

      {/* EDIT/ADD MODAL */}
      {(isAdding || editingItem) && (
        <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-950/60 backdrop-blur-md overflow-y-auto">
          <div className="bg-white w-full max-w-2xl rounded-[3rem] p-10 shadow-2xl my-8">
            <div className="flex justify-between items-center mb-10">
              <div className="flex items-center gap-4">
                <div className="w-12 h-12 bg-slate-100 rounded-[1.25rem] flex items-center justify-center text-slate-950 shadow-sm">
                  {isAdding ? <Plus size={24} strokeWidth={2.5}/> : <Edit2 size={22} />}
                </div>
                <div>
                  <h2 className="text-base font-black uppercase tracking-tight">{isAdding ? 'New Component' : 'Edit Registry'}</h2>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-[0.2em] mt-1">EFS RX500 Sport Inventory System</p>
                </div>
              </div>
              <button onClick={() => { setIsAdding(false); setEditingItem(null); }} className="p-3 hover:bg-slate-100 rounded-full transition-colors"><X size={24}/></button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
              <div className="space-y-6">
                <div className="space-y-2.5 relative">
                  <div className="flex items-center justify-between">
                    <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Product Link</label>
                    {isAutoFilling && <div className="flex items-center gap-2 animate-pulse text-blue-500"><Sparkles size={12} /><span className="text-[8px] font-black uppercase">Analysing Store...</span></div>}
                  </div>
                  <div className="relative">
                    <input
                      type="url"
                      placeholder="https://shop.com/part-link"
                      className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] border border-slate-100 focus:border-blue-400 outline-none transition-all pr-12"
                      value={isAdding ? newItem.url : editingItem.url}
                      onChange={(e) => handleUrlChange(e.target.value)}
                    />
                    <Link2 size={16} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-300" />
                  </div>
                </div>

                <div className="space-y-2.5">
                  <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Part Name</label>
                  <input type="text" className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] font-bold border border-slate-100 focus:border-slate-300 outline-none transition-all" value={isAdding ? newItem.name : editingItem.name} onChange={(e) => isAdding ? setNewItem({...newItem, name: e.target.value}) : setEditingItem({...editingItem, name: e.target.value})} />
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-2.5">
                    <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Unit Price</label>
                    <div className="relative">
                      <input type="number" step="0.01" className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] font-black border border-slate-100 outline-none pr-10" value={isAdding ? newItem.price : editingItem.price} onChange={(e) => isAdding ? setNewItem({...newItem, price: e.target.value}) : setEditingItem({...editingItem, price: e.target.value})} />
                      <Euro size={14} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-300" />
                    </div>
                  </div>
                  <div className="space-y-2.5">
                    <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Status / Phase</label>
                    <select className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] font-bold border border-slate-100 outline-none appearance-none cursor-pointer" value={isAdding ? newItem.category : editingItem.category} onChange={(e) => isAdding ? setNewItem({...newItem, category: e.target.value}) : setEditingItem({...editingItem, category: e.target.value})}>
                      {CATEGORIES.map(c => <option key={c} value={c}>{c}</option>)}
                    </select>
                  </div>
                </div>
              </div>

              <div className="space-y-6">
                <div className="space-y-2.5">
                  <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Store / Publisher</label>
                  <input type="text" className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] border border-slate-100 outline-none" value={isAdding ? newItem.vendor : editingItem.vendor} onChange={(e) => isAdding ? setNewItem({...newItem, vendor: e.target.value}) : setEditingItem({...editingItem, vendor: e.target.value})} />
                </div>
                <div className="space-y-2.5">
                  <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Image URL</label>
                  <input type="url" className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] border border-slate-100 outline-none font-mono text-[9px]" value={isAdding ? newItem.imageUrl : editingItem.imageUrl} onChange={(e) => isAdding ? setNewItem({...newItem, imageUrl: e.target.value}) : setEditingItem({...editingItem, imageUrl: e.target.value})} />
                </div>
                <div className="space-y-2.5">
                  <label className="text-[9px] font-black uppercase text-slate-400 tracking-widest">Quick Notes</label>
                  <textarea rows={3} className="w-full p-4 bg-slate-50 rounded-2xl text-[11px] border border-slate-100 outline-none resize-none leading-relaxed" value={isAdding ? newItem.notes : editingItem.notes} onChange={(e) => isAdding ? setNewItem({...newItem, notes: e.target.value}) : setEditingItem({...editingItem, notes: e.target.value})} />
                </div>
              </div>
            </div>

            <div className="flex gap-4 mt-12">
               <button onClick={() => { setIsAdding(false); setEditingItem(null); }} className="flex-1 py-5 rounded-2xl font-black text-[10px] uppercase tracking-[0.2em] text-slate-400 bg-slate-50 hover:bg-slate-100 transition-colors">Dismiss</button>
               <button onClick={isAdding ? saveNewItem : () => { updateItem(editingItem.id, editingItem); setEditingItem(null); }} className="flex-[2] bg-slate-950 text-white py-5 rounded-2xl font-black text-[10px] uppercase tracking-[0.2em] shadow-2xl shadow-slate-900/30 hover:bg-slate-800 transition-all flex items-center justify-center gap-3">
                <CheckCircle2 size={18} /> {isAdding ? 'Commit Entry' : 'Update Record'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IMPORT MODAL */}
      {importModalOpen && (
        <div className="fixed inset-0 z-[130] flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xl animate-in fade-in">
          <div className="bg-white w-full max-w-xl rounded-[3rem] p-10 shadow-2xl border border-white/10 scale-in-center">
            <div className="flex items-center justify-between mb-8">
              <div className="flex items-center gap-4">
                <div className="p-3 bg-emerald-50 rounded-[1.25rem] text-emerald-600 shadow-sm"><FileSpreadsheet size={24} /></div>
                <div>
                  <h3 className="text-sm font-black uppercase tracking-tight">Mass Import Center</h3>
                  <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest mt-1">Populate registry via Links or Files</p>
                </div>
              </div>
              <button onClick={() => setImportModalOpen(false)} className="p-2 hover:bg-slate-50 rounded-full transition-colors"><X size={24}/></button>
            </div>

            {/* Mode Switcher */}
            <div className="flex bg-slate-100 p-2 rounded-2xl mb-8 ring-1 ring-slate-200 shadow-inner">
              <button
                onClick={() => setImportMode('links')}
                className={`flex-1 py-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all flex items-center justify-center gap-3 ${importMode === 'links' ? 'bg-white text-blue-600 shadow-md ring-1 ring-blue-50' : 'text-slate-400 hover:text-slate-600'}`}
              >
                <Link2 size={14}/> Auto-Link Scan
              </button>
              <button
                onClick={() => setImportMode('excel')}
                className={`flex-1 py-4 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all flex items-center justify-center gap-3 ${importMode === 'excel' ? 'bg-white text-emerald-600 shadow-md ring-1 ring-emerald-50' : 'text-slate-400 hover:text-slate-600'}`}
              >
                <FileSpreadsheet size={14}/> Excel Mapper
              </button>
            </div>

            {importMode === 'links' ? (
              <div className="space-y-6 animate-in slide-in-from-top-2">
                <div className="bg-blue-50/50 p-5 rounded-[1.5rem] border border-blue-100 flex items-start gap-4">
                  <Sparkles size={20} className="text-blue-500 shrink-0 mt-0.5" />
                  <div>
                    <p className="text-[11px] font-black text-blue-900 mb-1 uppercase tracking-tight">AI Metadata Analysis</p>
                    <p className="text-[9px] text-blue-700 font-medium leading-relaxed">Paste product links (one per line). We'll crawl each URL to extract high-res images, pricing, and titles automatically.</p>
                  </div>
                </div>
                <textarea
                  className="w-full h-56 p-5 bg-slate-50 rounded-[1.5rem] text-[10px] font-mono border border-slate-100 focus:border-blue-300 outline-none transition-all resize-none shadow-inner"
                  placeholder="https://motorcycle-shop.com/engine-gasket&#10;https://part-vendor.pt/brake-disk&#10;..."
                  value={bulkLinks}
                  onChange={(e) => setBulkLinks(e.target.value)}
                />
              </div>
            ) : (
              <div className="space-y-6 animate-in slide-in-from-top-2">
                 <div className="flex bg-slate-50 p-1.5 rounded-2xl border border-slate-200">
                  <button onClick={() => setImportType('simple')} className={`flex-1 py-3 rounded-xl text-[9px] font-black transition-all ${importType === 'simple' ? 'bg-white shadow-sm text-slate-950' : 'text-slate-400'}`}>STATIC IMPORT</button>
                  <button onClick={() => setImportType('auto')} className={`flex-1 py-3 rounded-xl text-[9px] font-black transition-all ${importType === 'auto' ? 'bg-white shadow-sm text-emerald-600' : 'text-slate-400'}`}>HYBRID SCAN</button>
                </div>

                <div className="bg-emerald-50/50 p-5 rounded-[1.5rem] border border-emerald-100 flex items-start gap-4">
                  <Zap size={20} className="text-emerald-500 shrink-0 mt-0.5" />
                  <p className="text-[9px] text-emerald-800 font-medium leading-relaxed">
                    {importType === 'simple'
                      ? "Direct Excel mapping. We'll import columns exactly as matched."
                      : "We'll read links from your Excel to verify live pricing and fetch updated product images via Microlink API."}
                  </p>
                </div>

                <div className="grid grid-cols-2 gap-4 max-h-[35vh] overflow-y-auto pr-2 custom-scrollbar p-1">
                    {Object.keys(FIELD_MAP).map(k => (
                      <div key={k} className="space-y-1.5">
                        <label className="text-[8px] font-black uppercase text-slate-400 tracking-widest">{FIELD_MAP[k].label}</label>
                        <select
                          className="w-full p-3.5 bg-slate-50 border border-slate-100 rounded-xl text-[10px] font-bold outline-none hover:bg-white focus:bg-white transition-all shadow-sm"
                          value={mapping[k]}
                          onChange={e => setMapping({...mapping, [k]: e.target.value})}
                        >
                          <option value="">(Skip Field)</option>
                          {importHeaders.map(h => <option key={h} value={h}>{h}</option>)}
                        </select>
                      </div>
                    ))}
                </div>
              </div>
            )}

            <button
              onClick={processImport}
              disabled={isProcessingImport || (importMode === 'links' && !bulkLinks.trim())}
              className="w-full mt-8 p-5 bg-slate-950 text-white font-black text-[11px] uppercase tracking-[0.2em] rounded-2xl flex items-center justify-center gap-4 disabled:opacity-30 shadow-2xl shadow-slate-900/30 active:scale-95 transition-all"
            >
              {isProcessingImport ? (
                <>
                  <Loader2 className="animate-spin" size={18}/>
                  Processing Batch...
                </>
              ) : (
                <>
                  <CheckCircle2 size={18}/>
                  {importMode === 'excel' ? `Commit ${importData.length} records` : 'Analyse & Inject Links'}
                </>
              )}
            </button>
          </div>
        </div>
      )}

      {/* FULLSCREEN LIGHTBOX */}
      {expandedImage && (
        <div className="fixed inset-0 z-[200] bg-slate-950/98 backdrop-blur-3xl flex items-center justify-center p-8 animate-in fade-in duration-300" onClick={() => setExpandedImage(null)}>
           <img src={expandedImage} className="max-w-full max-h-full rounded-[3rem] object-contain shadow-[0_0_100px_rgba(255,255,255,0.05)] ring-1 ring-white/10" alt="Fullscreen View" />
           <div className="absolute top-10 right-10 flex flex-col gap-4">
              <button onClick={() => setExpandedImage(null)} className="p-4 bg-white/10 hover:bg-white/20 text-white rounded-full transition-all backdrop-blur-md">
                <X size={28}/>
              </button>
           </div>
        </div>
      )}

      {/* --- STYLES --- */}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

        :root {
          font-family: 'Inter', sans-serif;
        }

        .custom-scrollbar::-webkit-scrollbar {
          width: 5px;
        }

        .custom-scrollbar::-webkit-scrollbar-thumb {
          background: #E2E8F0;
          border-radius: 10px;
        }

        .custom-scrollbar::-webkit-scrollbar-track {
          background: transparent;
        }

        @keyframes scale-in-center {
          0% { transform: scale(0.9); opacity: 0; }
          100% { transform: scale(1); opacity: 1; }
        }

        .scale-in-center {
          animation: scale-in-center 0.3s cubic-bezier(0.25, 0.46, 0.45, 0.94) both;
        }

        input[type="number"]::-webkit-inner-spin-button,
        input[type="number"]::-webkit-outer-spin-button {
          -webkit-appearance: none;
          margin: 0;
        }
      `}</style>
    </div>
  );
}
