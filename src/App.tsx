import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import { 
  Settings, 
  Sliders, 
  Database, 
  UploadCloud, 
  Trash2, 
  Search, 
  Car, 
  FileText, 
  Printer, 
  Download,
  Shuffle, 
  Plus, 
  X, 
  Clock, 
  User, 
  MapPin, 
  Gauge, 
  Key, 
  Tag, 
  CheckCircle2, 
  AlertTriangle,
  FileCheck,
  Award,
  ChevronRight,
  LogOut,
  SlidersHorizontal,
  ChevronLeft
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { DeliveryRecord, ColumnMapping, BrandType, ChecklistSection } from './types';
import { FIELD_ALIASES_MAP, DEFAULT_CHECKLIST, ACCESSORIES_LIST, GENERAL_OBSERVATIONS_TEMPLATES } from './data';

const DEFAULT_DEVENTRY: DeliveryRecord = {
  CLIENTE: 'JEHENER JOSUE ALVARADO CRUZ',
  CEDULA: '001-020626-0000A',
  NO_PROFORMA: '00002986',
  CHASIS: 'JAANPR71HV7100023',
  FECHA_FACTURA: 'martes, 2 de junio de 2026',
  FECHA_ENTREGA: 'martes, 2 de junio de 2026',
  VENDEDOR: 'DAVID MELENDEZ',
  NO_TELEFONO: '+505 8888-8888',
  DIRECCION: 'Managua, Nicaragua',
  DEPARTAMENTO: 'MANAGUA',
  MODELO: 'NPR71HL',
  COLOR: 'BLANCO',
  MOTOR: '4HG1-1FK063',
  TIPO_VEHICULO: 'CAMION CABINA Y CHASIS',
  ANIO: '2027',
};

// Local storage keys
const STORAGE_RECORDS_KEY = 'samsa_delivery_records';
const STORAGE_MAPPING_KEY = 'samsa_delivery_column_mapping';
const STORAGE_BRAND_KEY = 'samsa_delivery_preferred_brand';

export default function App() {
  // Navigation State
  const [activeTab, setActiveTab] = useState<'config' | 'operacion' | 'documento'>('config');

  // Excel & File State
  const [excelData, setExcelData] = useState<DeliveryRecord[]>([DEFAULT_DEVENTRY]);
  const [excelColumns, setExcelColumns] = useState<string[]>(Object.keys(DEFAULT_DEVENTRY));
  const [columnMapping, setColumnMapping] = useState<ColumnMapping>({
    CLIENTE: 'CLIENTE',
    CEDULA: 'CEDULA',
    NO_PROFORMA: 'NO_PROFORMA',
    CHASIS: 'CHASIS',
    FECHA_FACTURA: 'FECHA_FACTURA',
    FECHA_ENTREGA: 'FECHA_ENTREGA',
    VENDEDOR: 'VENDEDOR',
    NO_TELEFONO: 'NO_TELEFONO',
    DIRECCION: 'DIRECCION',
    DEPARTAMENTO: 'DEPARTAMENTO',
    MODELO: 'MODELO',
    COLOR: 'COLOR',
    MOTOR: 'MOTOR',
    TIPO_VEHICULO: 'TIPO_VEHICULO',
    ANIO: 'ANIO',
  });

  // Search & Selector State
  const [searchQuery, setSearchQuery] = useState('JEHENER JOSUE ALVARADO CRUZ');
  const [isDropdownOpen, setIsDropdownOpen] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState<DeliveryRecord | null>(DEFAULT_DEVENTRY);
  const [selectedBrand, setSelectedBrand] = useState<BrandType>('isuzu');
  const [letterheadStyle, setLetterheadStyle] = useState<'standard' | 'blank' | 'custom'>('blank');
  const [customLetterheadA, setCustomLetterheadA] = useState<string[] | null>(null);
  const [customLetterheadB, setCustomLetterheadB] = useState<string[] | null>(null);
  const [activeCustomLetterheadSlot, setActiveCustomLetterheadSlot] = useState<'A' | 'B'>('A');

  const customLetterhead = activeCustomLetterheadSlot === 'A' ? customLetterheadA : customLetterheadB;
  const [isProcessingPdf, setIsProcessingPdf] = useState(false);

  // Interactive Checklist State (Mutable inside preview for custom live adjustments before printing!)
  const [checklist, setChecklist] = useState<ChecklistSection[]>(DEFAULT_CHECKLIST);
  const [accessories, setAccessories] = useState<{ id: string; label: string; checked: boolean }[]>(
    ACCESSORIES_LIST.map(acc => ({ ...acc, checked: true }))
  );

  // Manual Editing Form Overrides
  const [formModelo, setFormModelo] = useState('NPR71HL');
  const [formColor, setFormColor] = useState('BLANCO');
  const [formAnio, setFormAnio] = useState('2027');
  const [formMotor, setFormMotor] = useState('4HG1-1FK063');
  const [formTipo, setFormTipo] = useState('CAMION CABINA Y CHASIS');
  const [formDepartamento, setFormDepartamento] = useState('MANAGUA');
  const [formKm, setFormKm] = useState('12');
  const [formLlave, setFormLlave] = useState('7');
  const [formObs, setFormObs] = useState(GENERAL_OBSERVATIONS_TEMPLATES[0]);

  // UI States
  const [excelFileName, setExcelFileName] = useState('');
  const [alertMessage, setAlertMessage] = useState<{ text: string; type: 'success' | 'error' | 'info' | 'warning' } | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [showFullTable, setShowFullTable] = useState(false);
  const [tableSearch, setTableSearch] = useState('');
  const [tablePage, setTablePage] = useState(0);
  const itemsPerPage = 8;
  const [currentTime, setCurrentTime] = useState('');

  // Refs
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Time Effect
  useEffect(() => {
    const updateTime = () => {
      const now = new Date();
      setCurrentTime(now.toLocaleTimeString('es-NI', { hour: '2-digit', minute: '2-digit', second: '2-digit' }));
    };
    updateTime();
    const interval = setInterval(updateTime, 1000);
    return () => clearInterval(interval);
  }, []);

  // Keyboard events & Close Click Outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        setIsDropdownOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleCustomLetterheadUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      if (file.type !== 'application/pdf') {
        showToast('El formato para subir hoja membrete debe ser PDF.', 'error');
        return;
      }
      if (file.size > 25 * 1024 * 1024) {
        showToast('El PDF es demasiado grande. Por favor elija un archivo menor a 25MB.', 'error');
        return;
      }

      setIsProcessingPdf(true);
      showToast(`Procesando PDF para el Membrete ${activeCustomLetterheadSlot}...`, 'info');

      try {
        const reader = new FileReader();
        reader.onload = async (event) => {
          try {
            const arrayBuffer = event.target?.result as ArrayBuffer;
            if (!arrayBuffer) {
              throw new Error('No se pudo leer el archivo PDF.');
            }

            // Set worker dynamically matching the loaded library version to avoid mismatch errors
            const pdfjsVersion = (pdfjsLib as any).version || '6.0.227';
            pdfjsLib.GlobalWorkerOptions.workerSrc = `https://unpkg.com/pdfjs-dist@${pdfjsVersion}/build/pdf.worker.min.mjs`;

            const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(arrayBuffer) });
            const pdf = await loadingTask.promise;
            
            const pages: string[] = [];
            const numPages = Math.min(pdf.numPages, 4); // Extract up to 4 pages

            for (let pageNum = 1; pageNum <= numPages; pageNum++) {
              const page = await pdf.getPage(pageNum);
              
              // Standard Letter size viewport (scale 1.5 matches 918x1188px, highly detailed yet optimized for localStorage)
              const viewport = page.getViewport({ scale: 1.5 });
              const canvas = document.createElement('canvas');
              const context = canvas.getContext('2d');
              canvas.width = viewport.width;
              canvas.height = viewport.height;

              if (context) {
                // Background color white
                context.fillStyle = '#ffffff';
                context.fillRect(0, 0, canvas.width, canvas.height);

                await page.render({
                  canvasContext: context,
                  viewport: viewport,
                } as any).promise;

                // High quality JPEG compressed image to stay inside localStorage limits
                const dataUrl = canvas.toDataURL('image/jpeg', 0.85);
                pages.push(dataUrl);
              }
            }

            if (pages.length > 0) {
              if (activeCustomLetterheadSlot === 'A') {
                setCustomLetterheadA(pages);
              } else {
                setCustomLetterheadB(pages);
              }
              setLetterheadStyle('custom');
              showToast(`✓ PDF procesado con éxito para Membrete ${activeCustomLetterheadSlot} (${pages.length} página/s).`, 'success');
            } else {
              throw new Error('No se pudieron extraer páginas válidas del documento PDF.');
            }
          } catch (pdfError: any) {
            console.error('PDF parsing error:', pdfError);
            showToast(`Error al parsear páginas del PDF: ${pdfError.message || pdfError}`, 'error');
          } finally {
            setIsProcessingPdf(false);
          }
        };

        reader.onerror = () => {
          showToast('Error al leer el archivo PDF.', 'error');
          setIsProcessingPdf(false);
        };

        reader.readAsArrayBuffer(file);
      } catch (err: any) {
        console.error('File reading error:', err);
        showToast('Error de entrada/salida de archivo.', 'error');
        setIsProcessingPdf(false);
      }
    }
  };

  const clearCustomLetterhead = () => {
    if (activeCustomLetterheadSlot === 'A') {
      setCustomLetterheadA(null);
    } else {
      setCustomLetterheadB(null);
    }
    setLetterheadStyle('blank');
    showToast(`Membrete personalizado ${activeCustomLetterheadSlot} restablecido.`, 'info');
  };

  // Load from Local Storage on Mount
  useEffect(() => {
    try {
      const savedRecords = localStorage.getItem(STORAGE_RECORDS_KEY);
      const savedMapping = localStorage.getItem(STORAGE_MAPPING_KEY);
      const savedBrand = localStorage.getItem(STORAGE_BRAND_KEY);

      if (savedRecords) {
        const parsed = JSON.parse(savedRecords);
        if (parsed && parsed.length > 0) {
          setExcelData(parsed);
          // Derive existing keys as headers
          const cols = Object.keys(parsed[0]);
          setExcelColumns(cols);
          setExcelFileName('Restaurado de Sesión Anterior');
          if (parsed[0]) {
            const first = parsed[0];
            setSelectedRecord(first);
            setFormModelo(first.MODELO || 'NPR71HL');
            setFormColor(first.COLOR || 'BLANCO');
            setFormAnio(first.ANIO || '2027');
            setFormMotor(first.MOTOR || '4HG1-1FK063');
            setFormTipo(first.TIPO_VEHICULO || 'CAMION CABINA Y CHASIS');
            setFormDepartamento(first.DEPARTAMENTO || 'MANAGUA');
          }
          showToast('✓ Base de datos recuperada de la última sesión.', 'success');
        }
      } else {
        setSelectedRecord(DEFAULT_DEVENTRY);
        setFormModelo(DEFAULT_DEVENTRY.MODELO || 'NPR71HL');
        setFormColor(DEFAULT_DEVENTRY.COLOR || 'BLANCO');
        setFormAnio(DEFAULT_DEVENTRY.ANIO || '2027');
        setFormMotor(DEFAULT_DEVENTRY.MOTOR || '4HG1-1FK063');
        setFormTipo(DEFAULT_DEVENTRY.TIPO_VEHICULO || 'CAMION CABINA Y CHASIS');
        setFormDepartamento(DEFAULT_DEVENTRY.DEPARTAMENTO || 'MANAGUA');
      }

      if (savedMapping) {
        setColumnMapping(JSON.parse(savedMapping));
      }

      if (savedBrand) {
        setSelectedBrand(savedBrand as BrandType);
      }

      const savedLStyle = localStorage.getItem('samsa_letterhead_style');
      const savedCustomLH = localStorage.getItem('samsa_custom_letterhead');
      const savedCustomLHA = localStorage.getItem('samsa_custom_letterhead_a');
      const savedCustomLHB = localStorage.getItem('samsa_custom_letterhead_b');
      const savedActiveSlot = localStorage.getItem('samsa_active_letterhead_slot');

      if (savedLStyle) {
        setLetterheadStyle(savedLStyle === 'standard' ? 'blank' : savedLStyle as any);
      } else {
        setLetterheadStyle('blank');
      }

      // Restore slot A or migrate old single slot
      if (savedCustomLHA) {
        try {
          const parsed = JSON.parse(savedCustomLHA);
          setCustomLetterheadA(Array.isArray(parsed) ? parsed : [savedCustomLHA]);
        } catch (e) {
          setCustomLetterheadA([savedCustomLHA]);
        }
      } else if (savedCustomLH) {
        try {
          const parsed = JSON.parse(savedCustomLH);
          setCustomLetterheadA(Array.isArray(parsed) ? parsed : [savedCustomLH]);
        } catch (e) {
          setCustomLetterheadA([savedCustomLH]);
        }
      }

      // Restore slot B
      if (savedCustomLHB) {
        try {
          const parsed = JSON.parse(savedCustomLHB);
          setCustomLetterheadB(Array.isArray(parsed) ? parsed : [savedCustomLHB]);
        } catch (e) {
          setCustomLetterheadB([savedCustomLHB]);
        }
      }

      if (savedActiveSlot === 'A' || savedActiveSlot === 'B') {
        setActiveCustomLetterheadSlot(savedActiveSlot);
      }
    } catch (e) {
      console.error('Error recovering state from localStorage:', e);
    }
  }, []);

  // Save changes to localStorage
  useEffect(() => {
    localStorage.setItem('samsa_letterhead_style', letterheadStyle);
  }, [letterheadStyle]);

  useEffect(() => {
    if (customLetterheadA) {
      try {
        localStorage.setItem('samsa_custom_letterhead_a', JSON.stringify(customLetterheadA));
      } catch (err) {
        console.warn('Could not save large base64 image A into localStorage:', err);
      }
    } else {
      localStorage.removeItem('samsa_custom_letterhead_a');
    }
  }, [customLetterheadA]);

  useEffect(() => {
    if (customLetterheadB) {
      try {
        localStorage.setItem('samsa_custom_letterhead_b', JSON.stringify(customLetterheadB));
      } catch (err) {
        console.warn('Could not save large base64 image B into localStorage:', err);
      }
    } else {
      localStorage.removeItem('samsa_custom_letterhead_b');
    }
  }, [customLetterheadB]);

  useEffect(() => {
    localStorage.setItem('samsa_active_letterhead_slot', activeCustomLetterheadSlot);
  }, [activeCustomLetterheadSlot]);

  // Custom Toast helper
  const showToast = (text: string, type: 'success' | 'error' | 'info' | 'warning' = 'info') => {
    setAlertMessage({ text, type });
    setTimeout(() => {
      setAlertMessage(prev => prev?.text === text ? null : prev);
    }, 5000);
  };

  // Automated Mapping logic based on aliases
  const autoMapColumns = (cols: string[]) => {
    const mapping: ColumnMapping = {
      CLIENTE: '',
      CEDULA: '',
      NO_PROFORMA: '',
      CHASIS: '',
      FECHA_FACTURA: '',
      FECHA_ENTREGA: '',
      VENDEDOR: '',
      NO_TELEFONO: '',
      DIRECCION: '',
      DEPARTAMENTO: '',
      MODELO: '',
      COLOR: '',
      MOTOR: '',
      TIPO_VEHICULO: '',
      ANIO: '',
    };

    const colsCleaned = cols.map(c => ({
      original: c,
      normalized: c.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s\-\/\\]/g, '_')
    }));

    for (const [key, aliases] of Object.entries(FIELD_ALIASES_MAP)) {
      // Look for exact match
      const exactMatch = colsCleaned.find(c => aliases.includes(c.normalized));
      if (exactMatch) {
        mapping[key as keyof ColumnMapping] = exactMatch.original;
        continue;
      }
      
      // Look for fuzzy substring match
      const fuzzyMatch = colsCleaned.find(c => 
        aliases.some(alias => c.normalized.includes(alias) || alias.includes(c.normalized))
      );
      if (fuzzyMatch) {
        mapping[key as keyof ColumnMapping] = fuzzyMatch.original;
      }
    }

    setColumnMapping(mapping);
    localStorage.setItem(STORAGE_MAPPING_KEY, JSON.stringify(mapping));
  };

  // Excel Upload Parser
  const handleExcelUpload = (file: File) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const ab = e.target?.result;
        if (!ab) throw new Error('No se pudo leer el archivo de entrada');
        
        const wb = XLSX.read(ab, { type: 'array', cellDates: true });
        const firstSheetName = wb.SheetNames[0];
        const ws = wb.Sheets[firstSheetName];
        
        const rawJson = XLSX.utils.sheet_to_json<any[]>(ws, { header: 1, defval: '' });
        if (rawJson.length < 2) {
          throw new Error('El archivo no contiene suficientes filas. Asegúrate de tener una fila de encabezados.');
        }

        const headers = rawJson[0].map((h: any) => String(h).trim()).filter((h: string) => h !== '');
        
        const formattedData: DeliveryRecord[] = rawJson.slice(1)
          .filter((row: any) => row.some((cell: any) => cell !== null && cell !== ''))
          .map((row: any) => {
            const record: any = {};
            headers.forEach((col: string, idx: number) => {
              const val = row[idx];
              if (val instanceof Date) {
                record[col] = val.toLocaleDateString('es-NI');
              } else {
                record[col] = val !== undefined && val !== null ? String(val).trim() : '';
              }
            });
            return record as DeliveryRecord;
          });

        setExcelColumns(headers);
        setExcelData(formattedData);
        setExcelFileName(file.name);
        
        // Auto Map
        autoMapColumns(headers);
        
        // Persist to offline localStorage
        localStorage.setItem(STORAGE_RECORDS_KEY, JSON.stringify(formattedData));
        
        showToast(`✅ ¡Archivo de ventas importado con éxito! Se cargaron ${formattedData.length} registros.`, 'success');
      } catch (err: any) {
        showToast(`❌ Error al procesar el archivo Excel: ${err.message}`, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Drag-and-Drop Handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = () => {
    setDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) {
      handleExcelUpload(file);
    }
  };

  const clearDatabase = () => {
    if (window.confirm('¿Estás seguro de que deseas vaciar los datos cargados? Esto eliminará la base de datos de ventas de esta pestaña.')) {
      setExcelData([]);
      setExcelColumns([]);
      setExcelFileName('');
      localStorage.removeItem(STORAGE_RECORDS_KEY);
      localStorage.removeItem(STORAGE_MAPPING_KEY);
      setSelectedRecord(null);
      setSearchQuery('');
      showToast('🗑️ Base de datos eliminada.', 'info');
    }
  };

  // Manual column mapping change handler
  const handleMapChange = (field: keyof ColumnMapping, val: string) => {
    const updated = { ...columnMapping, [field]: val };
    setColumnMapping(updated);
    localStorage.setItem(STORAGE_MAPPING_KEY, JSON.stringify(updated));
    showToast(`✓ Mapeo de [${field}] actualizado.`, 'success');
  };

  // Retrieve value safely based on mapped headings
  const getMappedValue = (record: DeliveryRecord | null, field: keyof ColumnMapping): string => {
    if (!record) return '';
    const columnHeader = columnMapping[field];
    if (columnHeader && record[columnHeader] !== undefined) {
      return String(record[columnHeader]);
    }
    // Fallback search
    const aliases = FIELD_ALIASES_MAP[field];
    for (const [key, val] of Object.entries(record)) {
      const normalizedKey = key.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[\s_\-]/g, '');
      if (aliases.some(alias => normalizedKey === alias.replace(/_/g, '') || normalizedKey.includes(alias.replace(/_/g, '')))) {
        return String(val);
      }
    }
    return '';
  };

  // Client Search filtering
  const getSearchResults = () => {
    if (!searchQuery || searchQuery.length < 2) return [];
    
    const query = searchQuery.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    return excelData.filter(record => {
      const cliente = getMappedValue(record, 'CLIENTE').toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      const chasis = getMappedValue(record, 'CHASIS').toLowerCase();
      const factura = getMappedValue(record, 'NO_PROFORMA').toLowerCase();
      
      return cliente.includes(query) || chasis.includes(query) || factura.includes(query);
    }).slice(0, 8);
  };

  // Quick statistical summaries
  const isuzuCount = excelData.filter(r => {
    const model = getMappedValue(r, 'MODELO').toUpperCase();
    return !model.includes('GAC') && !model.includes('GS') && !model.includes('EMKOO') && !model.includes('TRUMPCHI');
  }).length;
  
  const gacCount = excelData.length - isuzuCount;

  // Manual record generation triggers
  const handleCreateManualRecord = () => {
    const t = new Date();
    const mockRecord: DeliveryRecord = {
      CLIENTE: 'PRODUCTOS DE NICARAGUA S.A.',
      CEDULA: 'J1210000346332',
      NO_PROFORMA: 'PRO-2026' + String(Math.floor(Math.random() * 8999) + 1000),
      CHASIS: '8ATB1DPR7HL53' + String(Math.floor(Math.random() * 89999) + 10000),
      FECHA_FACTURA: t.toLocaleDateString('es-NI'),
      FECHA_ENTREGA: t.toLocaleDateString('es-NI'),
      VENDEDOR: 'Lic. Javier Zelaya',
      NO_TELEFONO: '+505 8899-7722',
      DIRECCION: 'Carretera Norte KM 4.5, Frente a Cervecería',
      DEPARTAMENTO: 'Managua',
      MODELO: 'NPR71HL CABINA',
      COLOR: 'Blanco Isuzu S01',
      MOTOR: '4HG1-923812',
      TIPO_VEHICULO: 'Camión Liviano Chasis',
      ANIO: '2027'
    };
    
    // Set headers if empty
    if (excelColumns.length === 0) {
      setExcelColumns(Object.keys(mockRecord));
    }
    
    const newData = [mockRecord, ...excelData];
    setExcelData(newData);
    localStorage.setItem(STORAGE_RECORDS_KEY, JSON.stringify(newData));

    // Force exact column mapping matching
    const identityMap = { ...columnMapping };
    Object.keys(mockRecord).forEach(k => {
      if (!identityMap[k as keyof ColumnMapping]) {
        identityMap[k as keyof ColumnMapping] = k;
      }
    });
    setColumnMapping(identityMap);
    localStorage.setItem(STORAGE_MAPPING_KEY, JSON.stringify(identityMap));

    // Select immediately
    handleSelectRecord(0, newData);
    setActiveTab('operacion');
    showToast('✓ Registro rápido de ejemplo creado. Modifique los campos directamente.', 'success');
  };

  const handleSelectRecord = (index: number, currentList = excelData) => {
    const record = currentList[index];
    setSelectedRecord(record);
    setIsDropdownOpen(false);
    setSearchQuery(getMappedValue(record, 'CLIENTE'));

    // Initialize custom overridden fields
    const mod = getMappedValue(record, 'MODELO');
    const col = getMappedValue(record, 'COLOR');
    const an = getMappedValue(record, 'ANIO');
    const mot = getMappedValue(record, 'MOTOR');
    const tp = getMappedValue(record, 'TIPO_VEHICULO');
    const dep = getMappedValue(record, 'DEPARTAMENTO');

    setFormModelo(mod);
    setFormColor(col);
    setFormAnio(an);
    setFormMotor(mot);
    setFormTipo(tp || 'Unidad Comercial');
    setFormDepartamento(dep || 'Managua');
    setFormKm('12'); // Default delivery kilometer check 
    setFormLlave(String(Math.floor(Math.random() * 89) + 10));

    // Choose brand based on model name keywords
    const isGac = mod.toUpperCase().includes('GAC') || 
                  mod.toUpperCase().includes('GS') || 
                  mod.toUpperCase().includes('EMKOO') || 
                  mod.toUpperCase().includes('TRUMPCHI') ||
                  mod.toUpperCase().includes('EMPOW') ||
                  mod.toUpperCase().includes('M8');
    setSelectedBrand(isGac ? 'gac' : 'isuzu');

    // Default template suggestion
    setFormObs(GENERAL_OBSERVATIONS_TEMPLATES[0]);
  };

  const startEmptyManualRecord = () => {
    const freshRecord: DeliveryRecord = {
      CLIENTE: '',
      CEDULA: '',
      NO_PROFORMA: '',
      CHASIS: '',
      FECHA_FACTURA: new Date().toLocaleDateString('es-NI'),
      FECHA_ENTREGA: new Date().toLocaleDateString('es-NI'),
      VENDEDOR: 'Asesor de Ventas',
      NO_TELEFONO: '',
      DIRECCION: '',
      DEPARTAMENTO: 'Managua',
      MODELO: '',
      COLOR: 'Blanco',
      MOTOR: '',
      TIPO_VEHICULO: 'Sedán / SUV',
      ANIO: String(new Date().getFullYear() + 1)
    };
    
    setSelectedRecord(freshRecord);
    setSearchQuery('');
    setFormModelo('');
    setFormColor('Blanco');
    setFormAnio(String(new Date().getFullYear() + 1));
    setFormMotor('');
    setFormTipo('Camioneta SUV');
    setFormDepartamento('Managua');
    setFormKm('0');
    setFormLlave('');
    setFormObs(GENERAL_OBSERVATIONS_TEMPLATES[0]);
    
    setActiveTab('operacion');
    showToast('📝 Formulario listo para entrada manual directa.', 'info');
  };

  // Compile final print variables
  const getCompiledVariables = () => {
    if (!selectedRecord) return null;
    return {
      CLIENTE: getMappedValue(selectedRecord, 'CLIENTE') || 'COOPERATIVA DE TRANSPORTE R.L.',
      CEDULA: getMappedValue(selectedRecord, 'CEDULA') || 'J0310000122233',
      NO_PROFORMA: getMappedValue(selectedRecord, 'NO_PROFORMA') || 'FAC-001092',
      CHASIS: getMappedValue(selectedRecord, 'CHASIS') || 'WAV19828HBDL29',
      FECHA_FACTURA: getMappedValue(selectedRecord, 'FECHA_FACTURA') || new Date().toLocaleDateString('es-NI'),
      FECHA_ENTREGA: getMappedValue(selectedRecord, 'FECHA_ENTREGA') || new Date().toLocaleDateString('es-NI'),
      VENDEDOR: getMappedValue(selectedRecord, 'VENDEDOR') || 'Carlos J. Morales U.',
      NO_TELEFONO: getMappedValue(selectedRecord, 'NO_TELEFONO') || 'N/D',
      DIRECCION: getMappedValue(selectedRecord, 'DIRECCION') || 'Managua, Nicaragua',
      DEPARTAMENTO: formDepartamento,
      MODELO: formModelo,
      COLOR: formColor,
      MOTOR: formMotor,
      TIPO_VEHICULO: formTipo,
      ANIO: formAnio,
      KM: formKm,
      LLAVE: formLlave,
      OBS: formObs
    };
  };

  const handleGenerateDocuments = () => {
    if (!selectedRecord && !formModelo) {
      showToast('⚠️ Por favor selecciona o ingresa los datos mínimos de un vehículo primero.', 'warning');
      return;
    }
    
    // In case user hasn't selected a formal record but filled form, make a mock record
    if (!selectedRecord) {
      const mock: DeliveryRecord = {
        CLIENTE: 'CLIENTE INDIVIDUAL',
        CEDULA: '001-201090-0021K',
        NO_PROFORMA: 'FAC-' + String(Math.floor(Math.random() * 899)),
        CHASIS: '8ATB1DPR7HL500012',
        FECHA_FACTURA: new Date().toLocaleDateString('es-NI'),
        FECHA_ENTREGA: new Date().toLocaleDateString('es-NI'),
        VENDEDOR: 'ASESOR SAMSA',
        NO_TELEFONO: 'Managua',
        DIRECCION: 'Managua, Nicaragua',
        DEPARTAMENTO: formDepartamento || 'Managua',
        MODELO: formModelo,
        COLOR: formColor,
        MOTOR: formMotor,
        TIPO_VEHICULO: formTipo,
        ANIO: formAnio
      };
      setSelectedRecord(mock);
    }
    
    setActiveTab('documento');
    showToast('✨ 4 Documentos de entrega listos para revisión e impresión.', 'success');
  };

  // Helper trigger for window printing
  const triggerPrint = () => {
    window.print();
  };

  // Switch branding easily
  const toggleBrandBranding = () => {
    const nextBrand = selectedBrand === 'isuzu' ? 'gac' : 'isuzu';
    setSelectedBrand(nextBrand);
    localStorage.setItem(STORAGE_BRAND_KEY, nextBrand);
  };

  // Filters within checkout list inside configuration tab
  const getFilteredSheetData = () => {
    if (!tableSearch) return excelData;
    const query = tableSearch.toLowerCase();
    return excelData.filter(row => 
      Object.values(row).some(v => String(v).toLowerCase().includes(query))
    );
  };

  const filteredRows = getFilteredSheetData();
  const pageCount = Math.ceil(filteredRows.length / itemsPerPage);
  const displayedRows = filteredRows.slice(tablePage * itemsPerPage, (tablePage + 1) * itemsPerPage);

  const compVars = getCompiledVariables();

  return (
    <div className="min-h-screen bg-[#090b0f] text-[#dee2f0] font-sans flex flex-col antialiased">
      
      {/* GLOBAL TOAST BANNER */}
      <AnimatePresence>
        {alertMessage && (
          <motion.div 
            initial={{ opacity: 0, y: -20, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: -20, scale: 0.95 }}
            className="fixed top-18 right-6 z-250 pointer-events-auto"
          >
            <div className={`flex items-center gap-3 px-5 py-3.5 rounded-lg border shadow-xl max-w-[420px] transition-all bg-[#12151d] ${
              alertMessage.type === 'success' ? 'border-emerald-500/30 text-emerald-300 shadow-emerald-950/20' : 
              alertMessage.type === 'error' ? 'border-red-500/30 text-red-300 shadow-red-950/20' : 
              alertMessage.type === 'warning' ? 'border-amber-500/30 text-amber-300 shadow-amber-950/20' : 
              'border-[#2e364f] text-[#d1d7eb]'
            }`}>
              {alertMessage.type === 'success' && <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />}
              {alertMessage.type === 'error' && <AlertTriangle className="w-5 h-5 text-red-400 shrink-0" />}
              {alertMessage.type === 'warning' && <AlertTriangle className="w-5 h-5 text-amber-400 shrink-0" />}
              <div className="text-sm font-medium leading-relaxed">{alertMessage.text}</div>
              <button onClick={() => setAlertMessage(null)} className="hover:text-white ml-2 text-slate-500 self-start transition-colors">
                <X className="w-4 h-4" />
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* HEADER BAR */}
      <header className="bg-[#0f111a] border-b border-[#212739] px-6 h-16 flex items-center justify-between sticky top-0 z-100 no-print backdrop-blur-md bg-opacity-95">
        <div className="flex items-center gap-3">
          <div className="w-9 h-9 bg-gradient-to-tr from-[#d4922a] to-[#ffd060] rounded-lg flex items-center justify-center text-[#090b0f] font-bold text-lg tracking-tighter shrink-0 shadow-lg shadow-[#d4922a]/10">
            S
          </div>
          <div>
            <div className="font-extrabold text-[#f1f3f9] text-sm tracking-widest uppercase flex items-center gap-2">
              SAMSA DELIVERIES
              <span className="text-[9px] font-sans font-medium bg-[#1e2335] text-[#d4922a] px-1.5 py-0.5 rounded border border-[#2e364f] tracking-normal uppercase">
                AUTOMÁNTICA
              </span>
            </div>
            <div className="text-[10px] text-slate-400 tracking-wider">
              Servicio Automotriz Mántica S.A. — ISUZU & GAC MOTOR
            </div>
          </div>
        </div>

        {/* NAVIGATION PIPES */}
        <nav className="flex items-center gap-1.5 bg-[#0a0c12] p-1.5 rounded-lg border border-[#212739]">
          <button 
            onClick={() => setActiveTab('config')}
            className={`nav-btn px-4 py-2 font-medium text-xs rounded-md transition-all flex items-center gap-2 ${
              activeTab === 'config' 
                ? 'bg-[#181c2b] text-[#ffd060] border border-[#d4922a]/20 shadow-md' 
                : 'text-slate-400 hover:text-slate-100 hover:bg-[#11141f]'
            }`}
          >
            <Database className="w-3.5 h-3.5" />
            <span>⚙ Base de Ventas</span>
            {excelData.length > 0 && <span className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse"></span>}
          </button>
          
          <button 
            onClick={() => setActiveTab('operacion')}
            className={`nav-btn px-4 py-2 font-medium text-xs rounded-md transition-all flex items-center gap-2 ${
              activeTab === 'operacion' 
                ? 'bg-[#181c2b] text-[#ffd060] border border-[#d4922a]/20 shadow-md' 
                : 'text-slate-400 hover:text-slate-100 hover:bg-[#11141f]'
            }`}
          >
            <Car className="w-3.5 h-3.5" />
            <span>📊 Operación Diaria</span>
          </button>

          <button 
            onClick={() => setActiveTab('documento')}
            className={`nav-btn px-4 py-2 font-medium text-xs rounded-md transition-all flex items-center gap-2 ${
              activeTab === 'documento' 
                ? 'bg-[#181c2b] text-[#ffd060] border border-[#d4922a]/20 shadow-md' 
                : 'text-slate-400 hover:text-slate-100 hover:bg-[#11141f]'
            }`}
          >
            <FileText className="w-3.5 h-3.5" />
            <span>📄 Panel de Impresión</span>
            {selectedRecord && <span className="bg-[#d4922a]/15 text-[#ffd060] text-[9px] px-1.5 py-0.2 rounded-full border border-[#d4922a]/30 font-bold">4 Pág</span>}
          </button>
        </nav>
      </header>

      {/* CORE FRAMEWORK */}
      <main className="flex-1 overflow-y-auto no-print">
        
        {/* TAB 1: CONFIGRACIÓN (DATOS SOURCE) */}
        {activeTab === 'config' && (
          <motion.div 
            initial={{ opacity: 0, y: 5 }} 
            animate={{ opacity: 1, y: 0 }} 
            className="p-8 max-w-7xl mx-auto space-y-8"
          >
            {/* HERO HERO COMPONENT */}
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 p-6 rounded-2xl bg-gradient-to-r from-[#121522] to-[#0c0e16] border border-[#1d2338]">
              <div>
                <h1 className="text-xl md:text-2xl font-bold tracking-tight text-white flex items-center gap-2">
                  Gestión de <span className="text-[#ffd060]">Base de Ventas</span>
                </h1>
                <p className="text-xs text-slate-400 mt-1 max-w-2xl leading-relaxed">
                  Cargue registros directamente desde archivos de Excel generados en su sistema administrativo SAMSA. Los campos clave de facturación y chasis se autodetectarán para simplificar su flujo diario.
                </p>
              </div>
              <div className="flex gap-2">
                <button 
                  onClick={startEmptyManualRecord}
                  className="px-4 py-2 bg-slate-900 border border-slate-700 hover:border-[#d4922a] text-xs font-semibold rounded-lg flex items-center gap-2 hover:bg-slate-800 transition-all text-slate-300"
                >
                  <Plus className="w-4 h-4 text-[#ffd060]" />
                  Registro Individual
                </button>
                <button 
                  onClick={handleCreateManualRecord}
                  className="px-4 py-2 bg-gradient-to-r from-[#d4922a] to-[#ebb252] text-[#090b0f] text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-lg shadow-[#d4922a]/10 hover:brightness-110 active:scale-95 transition-all"
                >
                  <FileCheck className="w-4 h-4" />
                  Cargar Ejemplo Rápido
                </button>
              </div>
            </div>

            {/* QUICK OVERVIEV METRICS */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
              <div className="bg-[#0f111a] border border-[#1c2132] p-4 rounded-xl flex items-center justify-between">
                <div>
                  <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold">Unidades Cargadas</div>
                  <div className="text-2xl font-extrabold text-white mt-1">{excelData.length}</div>
                </div>
                <div className="w-10 h-10 rounded-lg bg-emerald-500/10 flex items-center justify-center text-emerald-400">
                  <Database className="w-5 h-5" />
                </div>
              </div>

              <div className="bg-[#0f111a] border border-[#1c2132] p-4 rounded-xl flex items-center justify-between">
                <div>
                  <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold">Unidades ISUZU</div>
                  <div className="text-2xl font-extrabold text-red-400 mt-1">{isuzuCount}</div>
                </div>
                <div className="w-10 h-10 rounded-lg bg-red-500/10 flex items-center justify-center text-red-400">
                  <Award className="w-5 h-5" />
                </div>
              </div>

              <div className="bg-[#0f111a] border border-[#1c2132] p-4 rounded-xl flex items-center justify-between">
                <div>
                  <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold">Unidades GAC MOTOR</div>
                  <div className="text-2xl font-extrabold text-blue-400 mt-1">{gacCount}</div>
                </div>
                <div className="w-10 h-10 rounded-lg bg-blue-500/10 flex items-center justify-center text-blue-400">
                  <Car className="w-5 h-5" />
                </div>
              </div>

              <div className="bg-[#0f111a] border border-[#1c2132] p-4 rounded-xl flex items-center justify-between">
                <div>
                  <div className="text-[10px] uppercase tracking-widest text-slate-500 font-bold">Estado Mapeo</div>
                  <div className="text-sm font-bold text-slate-300 mt-1.5 flex items-center gap-1.5">
                    {Object.values(columnMapping).filter(Boolean).length}/15 Campos Activos
                  </div>
                </div>
                <div className="w-10 h-10 rounded-lg bg-amber-500/10 flex items-center justify-center text-amber-400">
                  <Sliders className="w-5 h-5" />
                </div>
              </div>
            </div>

            {/* DROP AND DROP PANEL */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
              
              {/* UPLOADER */}
              <div className="lg:col-span-2 space-y-6">
                <div 
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  className={`border-2 border-dashed rounded-2xl p-10 text-center transition-all cursor-pointer relative bg-[#0e1017] ${
                    dragOver ? 'border-[#ffd060] bg-[#1d1b15]' : 'border-[#242c44] hover:border-[#ffd060]/45'
                  } ${excelData.length > 0 ? 'border-emerald-500/35 bg-[#0f1412]' : ''}`}
                >
                  <input 
                    type="file" 
                    accept=".xlsx,.xls,.csv" 
                    onChange={(e) => {
                      const file = e.target.files?.[0];
                      if (file) handleExcelUpload(file);
                    }}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                  />
                  <div className="flex flex-col items-center justify-center">
                    <div className={`w-14 h-14 rounded-full flex items-center justify-center mb-4 ${
                      excelData.length > 0 ? 'bg-emerald-500/10 text-emerald-400' : 'bg-slate-900 text-[#ffd060]'
                    }`}>
                      <UploadCloud className="w-7 h-7" />
                    </div>
                    {excelData.length > 0 ? (
                      <div>
                        <h4 className="text-base font-bold text-emerald-300">¡Hoja de Ventas Importada exitosamente!</h4>
                        <p className="text-xs text-slate-400 mt-1">Archivo actual: <span className="font-mono text-white underline">{excelFileName}</span></p>
                        <div className="mt-4 px-3 py-1 bg-emerald-500/10 text-emerald-400 text-xs rounded-full inline-block font-semibold">
                          ✓ {excelData.length} registros listos para operar
                        </div>
                      </div>
                    ) : (
                      <div>
                        <h4 className="text-base font-bold text-slate-100">Cargar Archivo Excel de Ventas / Facturación</h4>
                        <p className="text-xs text-slate-400 mt-1 max-w-md mx-auto">
                          Arrastre su archivo Excel de entregas aquí, o haga clic para explorar. Se procesa completamente de forma segura en local.
                        </p>
                        <div className="mt-4 text-[11px] font-semibold text-slate-500 flex items-center justify-center gap-3">
                          <span>Soportado: XLS, XLSX, CSV</span>
                        </div>
                      </div>
                    )}
                  </div>
                </div>

                {/* VISUAL DATAGRID FOR VERIFYING DATA */}
                {excelData.length > 0 && (
                  <div className="bg-[#0f111a] border border-[#1b2133] rounded-2xl p-6 space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <h3 className="text-sm font-bold text-slate-100 flex items-center gap-1.5">
                          <Database className="w-4 h-4 text-[#ffd060]" />
                          Visor Interactivo de Hojas Cargadas
                        </h3>
                        <p className="text-[11px] text-slate-500">Mostrando registros reales parseados del Excel.</p>
                      </div>
                      
                      <div className="flex items-center gap-2">
                        <div className="relative">
                          <Search className="absolute left-3 top-2.5 w-3.5 h-3.5 text-slate-400" />
                          <input 
                            type="text" 
                            placeholder="Buscar en la grilla..."
                            value={tableSearch}
                            onChange={(e) => {
                              setTableSearch(e.target.value);
                              setTablePage(0);
                            }}
                            className="bg-[#06080d] border border-[#21283a] rounded-lg pl-9 pr-3 py-1.5 text-xs text-slate-300 placeholder-slate-500 focus:outline-none focus:border-[#d4922a] w-52"
                          />
                        </div>
                        <button 
                          onClick={clearDatabase}
                          className="p-1.5 bg-[#1a0f0f] border border-red-500/20 text-red-400 hover:bg-red-950/20 rounded-lg hover:border-red-500/40 transition-all"
                          title="Vaciar base de datos"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>

                    <div className="overflow-x-auto border border-[#22293c] rounded-xl">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="bg-[#141824] border-b border-[#22293c] text-slate-400 font-semibold">
                            <th className="p-3">Cliente</th>
                            <th className="p-3 font-mono">Factura</th>
                            <th className="p-3 font-mono">Chasis / VIN</th>
                            <th className="p-3">Modelo</th>
                            <th className="p-3">Color</th>
                            <th className="p-3">Vendedor</th>
                            <th className="p-3 text-center">Acciones</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-[#1e2436] text-slate-300">
                          {displayedRows.length === 0 ? (
                            <tr>
                              <td colSpan={7} className="p-8 text-center text-slate-500">
                                Ningún registro coincide con el filtro de búsqueda.
                              </td>
                            </tr>
                          ) : (
                            displayedRows.map((row, idx) => {
                              const listIdx = excelData.indexOf(row);
                              return (
                                <tr key={idx} className="hover:bg-[#121624]">
                                  <td className="p-3 font-medium text-slate-100 max-w-[180px] break-words truncate">
                                    {getMappedValue(row, 'CLIENTE') || 'N/A'}
                                  </td>
                                  <td className="p-3 font-mono text-amber-500">
                                    {getMappedValue(row, 'NO_PROFORMA') || 'N/A'}
                                  </td>
                                  <td className="p-3 font-mono text-[#ffd060]">
                                    {getMappedValue(row, 'CHASIS') || 'N/A'}
                                  </td>
                                  <td className="p-3 font-semibold text-slate-200">
                                    {getMappedValue(row, 'MODELO') || 'N/A'}
                                  </td>
                                  <td className="p-3">{getMappedValue(row, 'COLOR') || 'N/A'}</td>
                                  <td className="p-3 text-slate-400">{getMappedValue(row, 'VENDEDOR') || 'N/A'}</td>
                                  <td className="p-3 text-center">
                                    <button 
                                      onClick={() => {
                                        handleSelectRecord(listIdx);
                                        setActiveTab('operacion');
                                      }}
                                      className="px-2 py-1 bg-gradient-to-r from-[#d4922a]/10 to-[#ffd060]/10 hover:from-[#d4922a]/20 hover:to-[#ffd060]/20 text-[#ffd060] font-bold text-[10px] rounded border border-[#d4922a]/30 transition-all"
                                    >
                                      Operar
                                    </button>
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>

                    {/* PAGINATION */}
                    {pageCount > 1 && (
                      <div className="flex items-center justify-between text-xs text-slate-400 p-2">
                        <div>
                          Mostrando {tablePage * itemsPerPage + 1} - {Math.min((tablePage + 1) * itemsPerPage, filteredRows.length)} de {filteredRows.length} registros
                        </div>
                        <div className="flex items-center gap-1">
                          <button 
                            onClick={() => setTablePage(p => Math.max(0, p - 1))}
                            disabled={tablePage === 0}
                            className="p-1.5 bg-[#131722] hover:bg-[#1b2132] disabled:opacity-40 rounded-lg border border-[#21283c]"
                          >
                            <ChevronLeft className="w-4 h-4" />
                          </button>
                          <span className="px-3">Página {tablePage + 1} de {pageCount}</span>
                          <button 
                            onClick={() => setTablePage(p => Math.min(pageCount - 1, p + 1))}
                            disabled={tablePage === pageCount - 1}
                            className="p-1.5 bg-[#131722] hover:bg-[#1b2132] disabled:opacity-40 rounded-lg border border-[#21283c]"
                          >
                            <ChevronRight className="w-4 h-4" />
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>

              {/* COLUMN MAPPING SIDEBAR */}
              <div className="space-y-6">
                <div className="bg-[#0f111a] border border-[#1c2132] rounded-2xl p-6 space-y-4">
                  <div className="flex items-center justify-between border-b border-[#21283d] pb-3">
                    <h3 className="text-xs font-bold text-slate-150 uppercase tracking-wider flex items-center gap-1.5">
                      <SlidersHorizontal className="w-4 h-4 text-[#ffd060]" />
                      Correspondencia
                    </h3>
                    <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#ffd060]/10 text-[#ffd060] font-semibold border border-[#d4922a]/20">
                      Auto-detectar activo
                    </span>
                  </div>
                  
                  <p className="text-[11px] text-slate-400 leading-relaxed">
                    Alinee las columnas del Excel con los campos esperados en los formularios impresos oficiales.
                  </p>

                  <div className="space-y-2.5 max-h-[460px] overflow-y-auto pr-1">
                    {Object.keys(FIELD_ALIASES_MAP).map((fieldKey) => {
                      const typedKey = fieldKey as keyof ColumnMapping;
                      return (
                        <div key={fieldKey} className="flex flex-col gap-1.5 bg-[#0a0c12] p-2.5 rounded-lg border border-[#1e2335]">
                          <div className="flex items-center justify-between">
                            <span className="text-[10px] font-mono font-bold text-[#ffd060]">
                              {`{{${fieldKey}}}`}
                            </span>
                            <span className="text-[9px] text-slate-500 font-medium uppercase">
                              Campo Destino
                            </span>
                          </div>
                          
                          <select 
                            value={columnMapping[typedKey] || ''}
                            onChange={(e) => handleMapChange(typedKey, e.target.value)}
                            className="w-full bg-[#111422] border border-[#252a3f] rounded px-2 py-1 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                          >
                            <option value="">— Ninguno (Omitir) —</option>
                            {excelColumns.map((col, cIdx) => (
                              <option key={cIdx} value={col}>{col}</option>
                            ))}
                          </select>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

            </div>
          </motion.div>
        )}

        {/* TAB 2: OPERACIÓN DIARIA */}
        {activeTab === 'operacion' && (
          <motion.div 
            initial={{ opacity: 0, y: 5 }} 
            animate={{ opacity: 1, y: 0 }} 
            className="p-8 max-w-5xl mx-auto space-y-6"
          >
            {/* SEARCH BANNER */}
            <div className="bg-[#0f111a] border border-[#1b2133] rounded-2xl p-6 space-y-4">
              <div>
                <h2 className="text-base font-bold text-white flex items-center gap-2">
                  <Search className="w-5 h-5 text-[#ffd060]" />
                  Buscador de Clientes en Base de Ventas
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  Ingrese el nombre completo del cliente, el número de chasis (VIN) de la unidad, o el número de factura.
                </p>
              </div>

              {excelData.length === 0 ? (
                <div className="bg-[#1a140c] border border-amber-500/20 p-4 rounded-xl flex items-start gap-3">
                  <AlertTriangle className="w-5 h-5 text-amber-500 shrink-0 mt-0.5" />
                  <div>
                    <h5 className="text-xs font-bold text-amber-300">Base de Datos vacía</h5>
                    <p className="text-[11px] text-slate-400 mt-0.5 leading-relaxed">
                      No has cargado ningún Excel de ventas en la pestaña de Configuración. Puedes ingresar los datos del vehículo de forma totalmente manual escribiéndolos a continuación, o ir a Configuración para cargar un Excel de ejemplo rápido.
                    </p>
                    <div className="mt-3 flex gap-2">
                      <button 
                        onClick={() => setActiveTab('config')}
                        className="px-3 py-1.5 bg-amber-500 text-slate-950 text-[10px] font-bold rounded-md hover:bg-amber-400 transition-all"
                      >
                        Ir a Cargar Base de Datos
                      </button>
                      <button 
                        onClick={() => {
                          const mock: DeliveryRecord = {
                            CLIENTE: 'CORPORACION FINANCIERA NACIONAL',
                            CEDULA: 'J0310000213948',
                            NO_PROFORMA: 'FAC-100231',
                            CHASIS: '8ATB1DPR7HL000192',
                            FECHA_FACTURA: new Date().toLocaleDateString('es-NI'),
                            FECHA_ENTREGA: new Date().toLocaleDateString('es-NI'),
                            VENDEDOR: 'Ing. Mercedes Centeno',
                            NO_TELEFONO: '+505 8899-1022',
                            DIRECCION: 'Managua, Semáforos El Dorado 2c al Norte',
                            DEPARTAMENTO: 'Managua',
                            MODELO: 'GS CLÁSSICO 1.5T',
                            COLOR: 'Plateado',
                            MOTOR: 'G4GB-102931',
                            TIPO_VEHICULO: 'SUV Familiar Premium',
                            ANIO: '2027'
                          };
                          setExcelData([mock]);
                          autoMapColumns(Object.keys(mock));
                          handleSelectRecord(0, [mock]);
                        }}
                        className="px-3 py-1.5 bg-slate-900 border border-slate-700 text-slate-200 text-[10px] font-semibold rounded-md hover:bg-slate-800 transition-all"
                      >
                        Cargar Datos Rápidos Manuales
                      </button>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="relative" ref={dropdownRef}>
                  <Search className="absolute left-3.5 top-3.5 w-5 h-5 text-slate-400" />
                  <input 
                    type="text" 
                    placeholder="Escriba Cliente, Chasis o N° de Factura..."
                    value={searchQuery}
                    onChange={(e) => {
                      setSearchQuery(e.target.value);
                      setIsDropdownOpen(true);
                    }}
                    onFocus={() => setIsDropdownOpen(true)}
                    className="w-full bg-[#06080d] border border-[#232a3d] rounded-xl pl-11 pr-4 py-3 text-sm text-slate-200 focus:outline-none focus:border-[#d4922a] placeholder-slate-500 focus:ring-1 focus:ring-[#d4922a]/20"
                  />

                  {/* DROP DOWN AUTO COMPLETE HIGHLIGHTS */}
                  {isDropdownOpen && searchQuery.length >= 2 && (
                    <div className="absolute top-full left-0 right-0 bg-[#0e111a] border border-[#21283d] rounded-xl mt-1.5 shadow-2xl max-h-72 overflow-y-auto overflow-x-hidden z-200 divide-y divide-[#1e2335]">
                      {getSearchResults().length === 0 ? (
                        <div className="p-4 text-center text-xs text-slate-500">
                          Ningún resultado coincide con &quot;{searchQuery}&quot;
                        </div>
                      ) : (
                        getSearchResults().map((result, idx) => {
                          const originalIndex = excelData.indexOf(result);
                          const client = getMappedValue(result, 'CLIENTE');
                          const chasis = getMappedValue(result, 'CHASIS');
                          const factura = getMappedValue(result, 'NO_PROFORMA');
                          const model = getMappedValue(result, 'MODELO') || 'Comercial';
                          const isGac = model.toUpperCase().includes('GAC') || 
                                        model.toUpperCase().includes('GS') || 
                                        model.toUpperCase().includes('EMKOO') || 
                                        model.toUpperCase().includes('TRUMPCHI');

                          return (
                            <button
                              key={idx}
                              onClick={() => handleSelectRecord(originalIndex)}
                              className="w-full text-left p-3 flex items-center justify-between text-xs hover:bg-[#141825] transition-all focus:outline-none"
                            >
                              <div className="space-y-1 max-w-[70%]">
                                <div className="font-bold text-slate-100 truncate">{client}</div>
                                <div className="text-slate-400 flex items-center gap-2">
                                  <span>Factura: <strong className="text-[#ffd060] font-mono">{factura}</strong></span>
                                  <span>•</span>
                                  <span>VIN: <strong className="font-mono">{chasis}</strong></span>
                                </div>
                              </div>
                              <div className="flex flex-col items-end gap-1.5 shrink-0">
                                <span className="text-[10px] text-slate-300 font-semibold truncate max-w-[120px]">{model}</span>
                                <span className={`px-2 py-0.5 rounded-full text-[9px] font-extrabold uppercase ${
                                  isGac ? 'bg-blue-500/10 text-blue-400 border border-blue-500/20' : 'bg-red-500/10 text-red-400 border border-red-500/20'
                                }`}>
                                  {isGac ? 'GAC MOTOR' : 'ISUZU'}
                                </span>
                              </div>
                            </button>
                          );
                        })
                      )}
                    </div>
                  )}
                </div>
              )}
            </div>

            {/* SELECTION PREVIEW AND DYNAMIC FORM */}
            {(selectedRecord || formModelo) && (
              <motion.div 
                initial={{ opacity: 0, scale: 0.99 }} 
                animate={{ opacity: 1, scale: 1 }}
                className="space-y-6"
              >
                {/* HEAD DETAILS */}
                <div className="bg-[#0f111a] border border-[#1b2133] rounded-2xl p-6">
                  <div className="flex flex-wrap items-center justify-between gap-4 border-b border-[#21283d] pb-4 mb-5">
                    <div>
                      <div className="text-[10px] text-slate-500 uppercase tracking-widest font-extrabold">Operando Venta Vigente</div>
                      <h3 className="text-base font-extrabold text-[#ffd060] mt-1">
                        {getMappedValue(selectedRecord, 'CLIENTE') || 'ENTRADA MANUAL'}
                      </h3>
                      <div className="text-[11px] text-slate-400 mt-0.5 flex gap-3">
                        <span>Cédula: <strong>{getMappedValue(selectedRecord, 'CEDULA') || 'J0000000000'}</strong></span>
                        <span>•</span>
                        <span>Factura: <strong>{getMappedValue(selectedRecord, 'NO_PROFORMA') || 'N/D'}</strong></span>
                      </div>
                    </div>

                    {/* TARGET BRAND FOR PRINT - DYNAMIC COLORS */}
                    <div className="flex items-center gap-1 bg-[#06080d] p-1 rounded-xl border border-[#21283d]">
                      <button 
                        onClick={() => setSelectedBrand('isuzu')}
                        className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 ${
                          selectedBrand === 'isuzu' 
                            ? 'bg-red-500 text-white shadow-lg shadow-red-950/20' 
                            : 'text-slate-400 hover:text-slate-100'
                        }`}
                      >
                        <Award className="w-3.5 h-3.5" />
                        ISUZU
                      </button>
                      <button 
                        onClick={() => setSelectedBrand('gac')}
                        className={`px-3 py-1.5 text-xs font-bold rounded-lg transition-all flex items-center gap-1.5 ${
                          selectedBrand === 'gac' 
                            ? 'bg-blue-500 text-white shadow-lg shadow-blue-950/20' 
                            : 'text-slate-400 hover:text-slate-100'
                        }`}
                      >
                        <Car className="w-3.5 h-3.5" />
                        GAC MOTOR
                      </button>
                    </div>
                  </div>

                  {/* FORM INPUT GRIDS */}
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <Tag className="w-3 h-3 text-[#ffd060]" />
                        Modelo Vehículo
                      </label>
                      <input 
                        type="text" 
                        value={formModelo}
                        onChange={(e) => setFormModelo(e.target.value)}
                        placeholder="NPR71HL"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <Sliders className="w-3 h-3 text-[#ffd060]" />
                        Color Exterior
                      </label>
                      <input 
                        type="text" 
                        value={formColor}
                        onChange={(e) => setFormColor(e.target.value)}
                        placeholder="Blanco / Plateado"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <Clock className="w-3 h-3 text-[#ffd060]" />
                        Año Modelo
                      </label>
                      <input 
                        type="text" 
                        value={formAnio}
                        onChange={(e) => setFormAnio(e.target.value)}
                        placeholder="2027"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <SlidersHorizontal className="w-3 h-3 text-[#ffd060]" />
                        Identificador de Motor
                      </label>
                      <input 
                        type="text" 
                        value={formMotor}
                        onChange={(e) => setFormMotor(e.target.value)}
                        placeholder="4HG1-A109"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a] font-mono font-bold"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <Car className="w-3 h-3 text-[#ffd060]" />
                        Tipo de Unidad
                      </label>
                      <input 
                        type="text" 
                        value={formTipo}
                        onChange={(e) => setFormTipo(e.target.value)}
                        placeholder="Camión de Chasis y Cabina / SUV"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <MapPin className="w-3 h-3 text-[#ffd060]" />
                        Dirección / Depto
                      </label>
                      <input 
                        type="text" 
                        value={formDepartamento}
                        onChange={(e) => setFormDepartamento(e.target.value)}
                        placeholder="Managua"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <Gauge className="w-3 h-3 text-[#ffd060]" />
                        Kilómetros de Recorrido
                      </label>
                      <input 
                        type="text" 
                        value={formKm}
                        onChange={(e) => setFormKm(e.target.value)}
                        placeholder="12"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>

                    <div className="flex flex-col gap-1.5">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <Key className="w-3 h-3 text-[#ffd060]" />
                        Número de Llaves Entregadas
                      </label>
                      <input 
                        type="text" 
                        value={formLlave}
                        onChange={(e) => setFormLlave(e.target.value)}
                        placeholder="Duplicado N° 2"
                        className="bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a]"
                      />
                    </div>
                  </div>

                  {/* OBSERVATIONS AND TEMPLATES SHORTCUTS */}
                  <div className="mt-5 space-y-2">
                    <div className="flex items-center justify-between">
                      <label className="text-[10px] font-bold text-slate-500 uppercase tracking-wider flex items-center gap-1">
                        <FileCheck className="w-3.5 h-3.5 text-[#ffd060]" />
                        Observaciones y Cláusulas del acta
                      </label>
                      <span className="text-[9px] text-[#ffd060]/75">Sugerencias rápidas:</span>
                    </div>

                    {/* SHORTCUT PILLS */}
                    <div className="flex flex-wrap gap-1.5 pb-2">
                      {GENERAL_OBSERVATIONS_TEMPLATES.map((tmpl, tIdx) => (
                        <button
                          key={tIdx}
                          type="button"
                          onClick={() => setFormObs(tmpl)}
                          className={`text-[9px] font-medium px-2.5 py-1 rounded-full border text-left truncate max-w-sm transition-all ${
                            formObs === tmpl 
                              ? 'bg-[#ffd060]/10 text-[#ffd060] border-[#d4922a]/35' 
                              : 'bg-[#06080d] text-slate-400 border-[#21283d] hover:text-white hover:border-slate-600'
                          }`}
                        >
                          {tmpl}
                        </button>
                      ))}
                    </div>

                    <textarea 
                      value={formObs}
                      onChange={(e) => setFormObs(e.target.value)}
                      rows={3}
                      className="w-full bg-[#06080d] border border-[#21283d] rounded-lg px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-[#d4922a] leading-relaxed resize-y"
                      placeholder="Escriba observaciones adicionales..."
                    />
                  </div>

                  {/* BUTTON ACTION ACTION */}
                  <div className="mt-8 pt-5 border-t border-[#21283d] flex justify-end gap-3">
                    <button 
                      onClick={() => {
                        setSelectedRecord(null);
                        setSearchQuery('');
                      }}
                      className="px-4 py-2 bg-slate-900 border border-slate-700 text-xs font-semibold rounded-lg hover:bg-slate-800 transition-all text-slate-300"
                    >
                      Cancelar Selección
                    </button>
                    <button 
                      onClick={handleGenerateDocuments}
                      className="px-6 py-2.5 bg-gradient-to-r from-[#d4922a] to-[#ffd060] text-slate-950 font-extrabold text-sm rounded-lg hover:brightness-110 active:scale-95 transition-all flex items-center gap-2 shadow-lg shadow-[#d4922a]/20"
                    >
                      Generar 4 Hojas de Impresión
                      <ChevronRight className="w-4 h-4 text-slate-900" />
                    </button>
                  </div>
                </div>
              </motion.div>
            )}

            {/* IF NO RECORD SELECTED YET */}
            {!selectedRecord && !formModelo && (
              <div className="p-16 border border-dashed border-[#23293b] rounded-2xl bg-[#0b0c13] text-center max-w-lg mx-auto">
                <div className="w-12 h-12 rounded-full bg-slate-900 text-[#ffd060] flex items-center justify-center mx-auto mb-4">
                  <User className="w-6 h-6" />
                </div>
                <h4 className="text-slate-200 font-bold text-sm">Ningún Cliente Seleccionado</h4>
                <p className="text-xs text-slate-500 mt-1.5 leading-relaxed">
                  Para proceder con la impresión, busque y seleccione un cliente en el buscador de arriba. También puede iniciar una entrada manual limpia para crear un acta puntual.
                </p>
                <div className="mt-5 flex gap-2 justify-center">
                  <button 
                    onClick={startEmptyManualRecord}
                    className="px-4 py-2 bg-[#d4922a]/10 hover:bg-[#d4922a]/20 border border-[#d4922a]/40 text-[#ffd060] font-bold text-xs rounded-lg transition-all"
                  >
                    Crear Nueva Entrega Manual
                  </button>
                </div>
              </div>
            )}
          </motion.div>
        )}

        {/* TAB 3: DOCUMENTOS (PROCESO IMPRESION) - SCREEN PREVIEW */}
        {activeTab === 'documento' && (
          <div className="p-8 max-w-7xl mx-auto space-y-6">
            
            {/* ACTION BANNER - SIMPLIFIED ACTION BAR */}
            <div className="bg-[#0f111a] border border-[#1b2133] rounded-2xl p-4 flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-center gap-2">
                <FileText className="w-5 h-5 text-[#ffd060]" />
                <span className="text-sm font-bold text-slate-100">Visor de Documentos (Carta)</span>
              </div>

              <div className="flex flex-wrap items-center gap-2.5 shrink-0">
                <button 
                  onClick={triggerPrint}
                  className="px-5 py-2.5 bg-gradient-to-r from-emerald-500 to-teal-500 text-white text-xs font-bold rounded-lg hover:brightness-110 active:scale-95 transition-all flex items-center gap-1.5 shadow-lg shadow-emerald-500/20"
                >
                  <Download className="w-4 h-4 text-white" />
                  Descargar o Imprimir PDF
                </button>
                
                <button 
                  onClick={() => setActiveTab('operacion')}
                  className="px-3 py-2 bg-slate-900 border border-slate-700 text-slate-300 text-xs font-bold rounded-lg hover:bg-slate-800 transition-all"
                >
                  Modificar Datos
                </button>
              </div>
            </div>

            {/* SCREEN LAYOUT WITH PAGES ON THE DESKTOP */}
            <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
              
              {/* INTERACTIVE CONTROLS ON THE SIDEBAR */}
              <div className="lg:col-span-1 space-y-6 animate-fadeIn">
                
                {/* SELECCIÓN Y CONFIGURACIÓN DE MEMBRETE */}
                <div className="bg-[#0f111a] border border-[#1c2132] rounded-2xl p-5 space-y-4">
                  <h4 className="text-xs font-extrabold uppercase text-[#ffd060] tracking-wider flex items-center gap-1.5">
                    <FileText className="w-4 h-4 text-[#ffd060]" />
                    Configuración de Membrete
                  </h4>
                  <p className="text-[11px] text-slate-400 leading-relaxed">
                    Personalice la presentación del membrete para la impresión física o exportación PDF.
                  </p>
                  <div className="space-y-3">
                    {/* STYLE OPTION CONTROLS */}
                    <div className="grid grid-cols-1 gap-2">
                      <button
                        onClick={() => {
                          setLetterheadStyle('blank');
                          showToast('✓ Modo Papel Membretado Físico activado.', 'info');
                        }}
                        className={`text-left rounded-lg text-[11px] font-bold transition-all border p-2.5 flex flex-col gap-1 w-full ${
                          letterheadStyle === 'blank'
                            ? 'bg-[#d4922a]/10 border-[#d4922a]/40 text-[#ffd060]'
                            : 'bg-[#06080d]/50 border-[#1e2335] text-slate-400 hover:bg-[#121623]'
                        }`}
                      >
                        <span className="flex items-center gap-1.5">
                          <span className={`w-2 h-2 rounded-full ${letterheadStyle === 'blank' ? 'bg-[#ffd060]' : 'bg-slate-700'}`} />
                          Papel Físico Membretado (Limpio)
                        </span>
                        <span className="text-[9.5px] text-slate-500 font-normal leading-normal">
                          Oculta los membretes en pantalla para imprimir directamente en sus propias hojas ya impresas.
                        </span>
                      </button>

                      <button
                        onClick={() => {
                          setLetterheadStyle('custom');
                          showToast('✓ Modo PDF de Membrete Propio activado.', 'info');
                        }}
                        className={`text-left rounded-lg text-[11px] font-bold transition-all border p-2.5 flex flex-col gap-1 w-full ${
                          letterheadStyle === 'custom'
                            ? 'bg-[#d4922a]/10 border-[#d4922a]/40 text-[#ffd060]'
                            : 'bg-[#06080d]/50 border-[#1e2335] text-slate-400 hover:bg-[#121623]'
                        }`}
                      >
                        <span className="flex items-center gap-1.5">
                          <span className={`w-2 h-2 rounded-full ${letterheadStyle === 'custom' ? 'bg-[#ffd060]' : 'bg-slate-700'}`} />
                          Subir Membrete en PDF
                        </span>
                        <span className="text-[9.5px] text-slate-500 font-normal leading-normal">
                          Cargue un archivo PDF de membrete para aplicar de forma personalizada pág. por pág.
                        </span>
                      </button>
                    </div>

                    {/* UPLOAD CONTROLS */}
                    {letterheadStyle === 'custom' && (
                      <div className="pt-2 border-t border-[#1e2335] space-y-3">
                        <span className="text-[10px] font-bold text-slate-400 block uppercase tracking-wider">
                          Ranura de Membrete
                        </span>

                        {/* Slot Switcher */}
                        <div className="flex gap-2 p-1 bg-[#0a0d16] border border-[#1e2335] rounded-xl">
                          <button
                            type="button"
                            onClick={() => {
                              setActiveCustomLetterheadSlot('A');
                              showToast('✓ Seleccionado Ranura Membrete A.', 'success');
                            }}
                            className={`flex-1 py-1.5 px-2 text-center text-[10px] font-extrabold rounded-lg transition-all flex items-center justify-center gap-1 ${
                              activeCustomLetterheadSlot === 'A'
                                ? 'bg-[#ffd060] text-slate-950 shadow font-black'
                                : 'text-slate-400 hover:text-white'
                            }`}
                          >
                            <span>Membrete A</span>
                            <span className={`w-1.5 h-1.5 rounded-full ${customLetterheadA ? 'bg-emerald-400' : 'bg-rose-500/50'}`}></span>
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              setActiveCustomLetterheadSlot('B');
                              showToast('✓ Seleccionado Ranura Membrete B.', 'success');
                            }}
                            className={`flex-1 py-1.5 px-2 text-center text-[10px] font-extrabold rounded-lg transition-all flex items-center justify-center gap-1 ${
                              activeCustomLetterheadSlot === 'B'
                                ? 'bg-[#ffd060] text-slate-950 shadow font-black'
                                : 'text-slate-400 hover:text-white'
                            }`}
                          >
                            <span>Membrete B</span>
                            <span className={`w-1.5 h-1.5 rounded-full ${customLetterheadB ? 'bg-emerald-400' : 'bg-rose-500/50'}`}></span>
                          </button>
                        </div>

                        <span className="text-[10px] font-bold text-slate-400 block uppercase tracking-wider pt-1">
                          Archivo Membrete de Fondo (PDF) - Slot {activeCustomLetterheadSlot}
                        </span>

                        {isProcessingPdf ? (
                          <div className="border border-[#ffd060]/30 rounded-xl p-4 text-center bg-slate-950/40 flex flex-col items-center gap-2 animate-pulse">
                            <span className="text-lg animate-spin">⏳</span>
                            <span className="text-[11px] font-bold text-[#ffd060]">Renderizando Páginas...</span>
                            <span className="text-[8.5px] text-slate-500">Convirtiendo PDF a alta definición</span>
                          </div>
                        ) : !customLetterhead ? (
                          <div className="relative border border-dashed border-[#23293b] hover:border-[#ffd060]/50 rounded-xl p-4 text-center cursor-pointer bg-slate-950/40 hover:bg-slate-950/80 transition-all">
                            <input
                              type="file"
                              accept="application/pdf"
                              onChange={handleCustomLetterheadUpload}
                              className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                            />
                            <div className="flex flex-col items-center gap-1.5">
                              <span className="text-lg">📄</span>
                              <span className="text-[11px] font-bold text-slate-300">Seleccionar PDF ({activeCustomLetterheadSlot})</span>
                              <span className="text-[8.5px] text-slate-500">Soporta PDFs de 1 a 4 páginas</span>
                            </div>
                          </div>
                        ) : (
                          <div className="bg-[#06080d] border border-[#212739] p-2.5 rounded-xl flex items-center justify-between gap-3 animate-fadeIn">
                            <div className="flex items-center gap-2 min-w-0">
                              <div className="w-10 h-10 border border-[#202538] rounded overflow-hidden bg-white shrink-0 flex items-center justify-center">
                                <span className="text-xs font-bold text-slate-700">PDF</span>
                              </div>
                              <div className="min-w-0">
                                <span className="text-[10px] font-mono font-bold text-slate-200 block truncate">
                                  {customLetterhead.length} Página(s) Listas
                                </span>
                                <span className="text-[8px] text-emerald-400 font-extrabold uppercase block tracking-wider">
                                  Membrete {activeCustomLetterheadSlot} Activo
                                </span>
                              </div>
                            </div>
                            <button
                              onClick={clearCustomLetterhead}
                              className="w-7 h-7 rounded-lg bg-red-500/10 hover:bg-red-500/20 text-red-400 hover:text-red-500 flex items-center justify-center border border-red-500/20 active:scale-95 transition-all shrink-0 text-xs font-bold"
                              title="Restablecer Membrete"
                            >
                              ✕
                            </button>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                </div>

                <div className="bg-[#0f111a] border border-[#1c2132] rounded-2xl p-5 space-y-4">
                  <h4 className="text-xs font-extrabold uppercase text-slate-400 tracking-wider">
                    Herramientas de Llenado
                  </h4>
                  <p className="text-[11px] text-slate-500 leading-relaxed">
                    Marque o desmarque los items del checklist general aquí. Se verá reflejado inmediatamente en los documentos impresos oficiales.
                  </p>

                  <div className="space-y-4">
                    {/* KEY ACCESORIES IN CLUSTERS */}
                    <div className="space-y-2">
                      <span className="text-[10px] font-bold text-[#ffd060] uppercase block">
                        Inventario Accesorios (Pág 2)
                      </span>
                      <div className="space-y-1.5 max-h-[220px] overflow-y-auto pr-1">
                        {accessories.map((acc, aIdx) => (
                          <label key={acc.id} className="flex items-center gap-2 bg-[#06080d] border border-[#1e2335] p-2 rounded cursor-pointer hover:border-slate-600 transition-colors">
                            <input 
                              type="checkbox" 
                              checked={acc.checked}
                              onChange={() => {
                                const copy = [...accessories];
                                copy[aIdx].checked = !copy[aIdx].checked;
                                setAccessories(copy);
                              }}
                              className="w-3.5 h-3.5 rounded bg-slate-900 border-slate-700 text-emerald-500 focus:ring-0 cursor-pointer"
                            />
                            <span className="text-[10px] text-slate-300">{acc.label}</span>
                          </label>
                        ))}
                      </div>
                    </div>

                    <div className="border-t border-[#21283d] pt-3">
                      <span className="text-[10px] font-bold text-slate-500 uppercase block mb-1">
                        Estatus Vehículo
                      </span>
                      <div className="text-[10px] bg-[#111422] p-3 rounded-lg border border-[#21283a] space-y-2">
                        <div className="flex justify-between">
                          <span className="text-slate-400">Cliente:</span>
                          <span className="font-semibold text-white truncate max-w-[120px]">{compVars?.CLIENTE}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-slate-400">VIN:</span>
                          <span className="font-mono font-semibold text-[#ffd060] text-[9px]">{compVars?.CHASIS}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-slate-400">Factura:</span>
                          <span className="font-semibold text-amber-500 text-[9px]">{compVars?.NO_PROFORMA}</span>
                        </div>
                        <div className="flex justify-between">
                          <span className="text-slate-400">Año:</span>
                          <span className="font-semibold text-white">{compVars?.ANIO}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* LIVE DOCK VIEWER GRID (Centered matching letter page shadows) */}
              <div id="preview-scroll-container" className="lg:col-span-3 space-y-10 flex flex-col items-center">
                
                {/* PAGE 1 */}
                <div className="bg-white border rounded shadow-2xl relative select-text" style={{ width: '816px', minHeight: '1056px', padding: '0px' }}>
                  {compVars && <DocumentPage1 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} checklist={checklist} setChecklist={setChecklist} />}
                </div>

                {/* PAGE 2 */}
                <div className="bg-white border rounded shadow-2xl relative select-text" style={{ width: '816px', minHeight: '1056px', padding: '0px' }}>
                  {compVars && <DocumentPage2 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} accessories={accessories} setAccessories={setAccessories} />}
                </div>

                {/* PAGE 3 */}
                <div className="bg-white border rounded shadow-2xl relative select-text" style={{ width: '816px', minHeight: '1056px', padding: '0px' }}>
                  {compVars && <DocumentPage3 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} />}
                </div>

                {/* PAGE 4 */}
                <div className="bg-white border rounded shadow-2xl relative select-text" style={{ width: '816px', minHeight: '1056px', padding: '0px' }}>
                  {compVars && <DocumentPage4 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} />}
                </div>

              </div>

            </div>
          </div>
        )}

      </main>

      {/* PRINT-ONLY EMBEDDED RENDER PIPELINE */}
      <div className="hidden print:block bg-white text-black select-text">
        {compVars && (
          <>
            <div className="relative" style={{ width: '100%', minHeight: '100vh', pageBreakAfter: 'always' }}>
              <DocumentPage1 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} checklist={checklist} setChecklist={setChecklist} />
            </div>
            <div className="relative" style={{ width: '100%', minHeight: '100vh', pageBreakAfter: 'always' }}>
              <DocumentPage2 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} accessories={accessories} setAccessories={setAccessories} />
            </div>
            <div className="relative" style={{ width: '100%', minHeight: '100vh', pageBreakAfter: 'always' }}>
              <DocumentPage3 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} />
            </div>
            <div className="relative" style={{ width: '100%', minHeight: '100vh' }}>
              <DocumentPage4 compVars={compVars} selectedBrand={selectedBrand} letterheadStyle={letterheadStyle} customLetterhead={customLetterhead} />
            </div>
          </>
        )}
      </div>

      {/* STATUS METRICS BAR */}
      <footer className="h-9 bg-[#0b0c13] border-t border-[#202538] px-6 text-xs text-slate-400 font-medium flex items-center justify-between no-print mt-auto shrink-0 z-100">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-1.5">
            <span className={`w-1.5 h-1.5 rounded-full ${excelData.length > 0 ? 'bg-emerald-500 animate-pulse' : 'bg-red-500'}`}></span>
            <span>Excel: {excelData.length > 0 ? `${excelData.length} registros cargados` : 'Base vacía'}</span>
          </div>
          <span>•</span>
          <div>Diseño: <strong className="text-white">Editorial A4 Legal Nicaragua (Letter Sized)</strong></div>
        </div>
        
        <div className="flex items-center gap-3">
          <span>Hora local: <span className="font-mono text-emerald-400">{currentTime || '12:00:00'}</span></span>
          <span className="text-slate-600">|</span>
          <span className="font-semibold text-slate-300">SAMSA AutoMántica v4.0</span>
        </div>
      </footer>

    </div>
  );
}

// =========================================================================
// DOCUMENT SUB-PAGES COMPONENTS FOR HIGH-FIDELITY LAYOUT (Inches/Letter size)
// =========================================================================

interface DocProps {
  compVars: {
    CLIENTE: string;
    CEDULA: string;
    NO_PROFORMA: string;
    CHASIS: string;
    FECHA_FACTURA: string;
    FECHA_ENTREGA: string;
    VENDEDOR: string;
    NO_TELEFONO: string;
    DIRECCION: string;
    DEPARTAMENTO: string;
    MODELO: string;
    COLOR: string;
    MOTOR: string;
    TIPO_VEHICULO: string;
    ANIO: string;
    KM: string;
    LLAVE: string;
    OBS: string;
  };
  selectedBrand: BrandType;
  letterheadStyle: 'standard' | 'blank' | 'custom';
  customLetterhead: string[] | null;
}

const DocHeader = ({ brand, letterheadStyle }: { brand: BrandType; letterheadStyle: 'standard' | 'blank' | 'custom' }) => {
  if (letterheadStyle !== 'standard') {
    return <div className="h-20 w-full select-none pointer-events-none mb-3"></div>;
  }
  const isGac = brand === 'gac';
  if (isGac) {
    return (
      <div className="flex flex-col select-none text-black">
        <div className="flex items-center justify-between pb-1">
          <div className="flex items-center gap-2">
            {/* Highly detailed, professional vector custom GAC logo icon with metallic gradient */}
            <svg className="w-11 h-7 shrink-0" viewBox="0 0 100 60" fill="none" xmlns="http://www.w3.org/2000/svg">
              <defs>
                <linearGradient id="gacSilverGradient" x1="0" y1="0" x2="100" y2="60" gradientUnits="userSpaceOnUse">
                  <stop offset="0%" stopColor="#FFFFFF" />
                  <stop offset="25%" stopColor="#DADADA" />
                  <stop offset="50%" stopColor="#7B7B7B" />
                  <stop offset="75%" stopColor="#EEEEEE" />
                  <stop offset="100%" stopColor="#3A3A3A" />
                </linearGradient>
              </defs>
              <ellipse cx="50" cy="30" rx="42" ry="22" stroke="url(#gacSilverGradient)" strokeWidth="6.5" fill="none" />
              <path d="M72 30 C72 19 61 14 48 14 C32 14 23 21 23 30 C23 39 32 46 48 46 C58 46 67 41 71 35" stroke="url(#gacSilverGradient)" strokeWidth="6.5" strokeLinecap="round" strokeLinejoin="round" />
              <path d="M44 30 H71" stroke="url(#gacSilverGradient)" strokeWidth="6.5" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            <span className="text-[#000000] font-sans font-black tracking-[0.14em] text-xl leading-none" style={{ fontFamily: 'system-ui, -apple-system, sans-serif' }}>
              GAC MOTOR
            </span>
          </div>
          <div className="text-right text-[9px] font-sans text-black leading-tight font-medium">
            Galerías Santo Domingo,<br />
            Frente a Hotel Hyatt Place, Managua<br />
            +505 7886 5566<br />
            www.gacnicaragua.com
          </div>
        </div>
        {/* Thin grey line under GAC header as shown in the second image */}
        <div className="h-[1.5px] bg-[#cdd1db] w-full mt-0.5 mb-2"></div>
      </div>
    );
  } else {
    return (
      <div className="flex flex-col select-none text-black">
        <div className="flex items-end justify-between pb-0.5">
          {/* SAMSA AUTOMOTRIZ Logo (Row 1: SAMSA, Row 2: <Red bar> AUTOMOTRIZ) as in image 1 */}
          <div className="flex flex-col items-start leading-none shrink-0 w-[112px]">
            <span className="text-2xl font-black text-black tracking-[0.04em] leading-none text-left w-full" style={{ fontFamily: 'system-ui, -apple-system, sans-serif', fontWeight: 950 }}>
              SAMSA
            </span>
            <div className="flex items-center w-full mt-1">
              <div className="h-[3.2px] bg-[#df1010] flex-grow mr-1.5 rounded-sm"></div>
              <span className="text-[7.2px] font-black text-black tracking-[0.14em] leading-none uppercase shrink-0">
                AUTOMOTRIZ
              </span>
            </div>
          </div>
          {/* Company address and RUC info */}
          <div className="text-right text-[10px] text-black font-sans font-bold tracking-tight mb-0.5 leading-none">
            SERVICIO AUTOMOTRIZ MANTICA, S.A. RUC:J0510000034981
          </div>
        </div>
        {/* Thick solid grey bar under header as shown in the first image */}
        <div className="h-[5px] bg-[#aaaaaa] w-full mt-1 mb-2.5"></div>
      </div>
    );
  }
};

const DocFooter = ({ brand, pageNum, letterheadStyle }: { brand: BrandType; pageNum: number; letterheadStyle: 'standard' | 'blank' | 'custom' }) => {
  if (letterheadStyle !== 'standard') {
    return (
      <div className="absolute bottom-6 left-8 right-8 select-none text-black text-[8.5px] font-sans font-semibold flex justify-end">
        Página {pageNum} de 4
      </div>
    );
  }
  const isGac = brand === 'gac';
  if (isGac) {
    return (
      <div className="absolute bottom-6 left-8 right-8 select-none text-black animate-fadeIn">
        <div className="flex items-end justify-between">
          {/* Bottom-left: SAMSA AUTOMOTRIZ Logo as shown in Image 2 */}
          <div className="flex flex-col items-start leading-none shrink-0 w-[112px]">
            <span className="text-2xl font-black text-black tracking-[0.04em] leading-none text-left w-full" style={{ fontFamily: 'system-ui, -apple-system, sans-serif', fontWeight: 950 }}>
              SAMSA
            </span>
            <div className="flex items-center w-full mt-1">
              <div className="h-[3.2px] bg-[#df1010] flex-grow mr-1.5 rounded-sm"></div>
              <span className="text-[7.2px] font-black text-black tracking-[0.14em] leading-none uppercase shrink-0">
                AUTOMOTRIZ
              </span>
            </div>
          </div>
          
          {/* Multi-layered custom colored horizontal letterhead accent line as in image 2 */}
          <div className="flex-1 flex items-end h-[6px] relative ml-4 mb-[3px]">
            <div className="absolute inset-x-0 h-[1.5px] bg-[#cdd1db] bottom-0"></div>
            {/* Perfectly measured red horizontal segment embedded on top of the grey rule */}
            <div className="absolute left-[15%] w-[35%] h-[3px] bg-[#df1010] bottom-0"></div>
          </div>
        </div>
      </div>
    );
  } else {
    return (
      <div className="absolute bottom-6 left-8 right-8 select-none text-black">
        <div className="flex items-end justify-between pb-2.5">
          {/* Left layout: contact info row with red icons and dark text */}
          <div className="flex items-center gap-5 text-[9.5px] font-bold text-black font-sans">
            <div className="flex items-center gap-1.5">
              <svg className="w-3.5 h-3.5 text-[#df1010]" fill="currentColor" viewBox="0 0 24 24">
                <path d="M6.62 10.79c1.44 2.83 3.76 5.14 6.59 6.59l2.2-2.2c.27-.27.67-.36 1.02-.24 1.12.37 2.33.57 3.57.57.55 0 1 .45 1 1V20c0 .55-.45 1-1 1-9.39 0-17-7.61-17-17 0-.55.45-1 1-1h3.5c.55 0 1 .45 1 1 0 1.25.2 2.45.57 3.57.11.35.03.74-.25 1.02l-2.2 2.2z" />
              </svg>
              <span className="font-sans text-black">2277-0177</span>
            </div>
            
            <div className="flex items-center gap-1.5">
              <svg className="w-3.5 h-3.5 text-[#df1010]" fill="currentColor" viewBox="0 0 24 24">
                <path d="M12 2C8.13 2 5 5.13 5 9c0 5.25 7 13 7 13s7-7.75 7-13c0-3.87-3.13-7-7-7zm0 9.5c-1.38 0-2.5-1.12-2.5-2.5s1.12-2.5 2.5-2.5 2.5 1.12 2.5 2.5-1.12 2.5-2.5 2.5z" />
              </svg>
              <span className="font-sans text-black">Rotonda universitaria</span>
            </div>
            
            <div className="flex items-center gap-1.5">
              <svg className="w-3.5 h-3.5 text-[#df1010]" fill="currentColor" viewBox="0 0 24 24">
                <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z" />
              </svg>
              <span className="font-sans text-black">www.isuzunicaragua.com</span>
            </div>
          </div>
          
          {/* Right layout: wide bold ISUZU corporate logo and page index counter */}
          <div className="flex items-baseline gap-4 select-none">
            <span className="text-[8px] text-black font-sans font-medium">Página {pageNum} de 4</span>
            {/* Phenomenally accurate vector recreation of the ISUZU red block wordmark */}
            <svg className="h-[13px] w-auto" viewBox="0 0 135 28" fill="#df1010" xmlns="http://www.w3.org/2000/svg">
              <rect x="0" y="0" width="10" height="28" />
              <rect x="18" y="0" width="24" height="8" />
              <rect x="18" y="8" width="8" height="6" />
              <rect x="18" y="11" width="24" height="6" />
              <rect x="34" y="14" width="8" height="6" />
              <rect x="18" y="20" width="24" height="8" />
              <rect x="49" y="0" width="8" height="21" />
              <rect x="65" y="0" width="8" height="21" />
              <rect x="49" y="20" width="24" height="8" />
              <rect x="80" y="0" width="24" height="8" />
              <polygon points="104,8 95,8 80,20 89,20" />
              <rect x="80" y="20" width="24" height="8" />
              <rect x="111" y="0" width="8" height="21" />
              <rect x="127" y="0" width="8" height="21" />
              <rect x="111" y="20" width="24" height="8" />
            </svg>
          </div>
        </div>
        {/* Thick solid red bar spanning the entire width at the absolute bottom */}
        <div className="h-[5px] bg-[#df1010] w-full"></div>
      </div>
    );
  }
};

// PÁGINA 1: ACTA DE INSPECCIÓN Y RECEPCIÓN
const DocumentPage1 = ({ compVars, selectedBrand, letterheadStyle, customLetterhead, checklist, setChecklist }: DocProps & { checklist: ChecklistSection[], setChecklist: React.Dispatch<React.SetStateAction<ChecklistSection[]>> }) => {
  const isGac = selectedBrand === 'gac';
  const brandName = isGac ? 'GAC MOTOR' : 'ISUZU';
  const brandColor = '#df1010';

  const toggleCheck = (sectionIdx: number, itemIdx: number) => {
    const copy = [...checklist];
    copy[sectionIdx].items[itemIdx].checked = !copy[sectionIdx].items[itemIdx].checked;
    setChecklist(copy);
  };

  return (
    <div className="pt-3 px-8 pb-36 min-h-[1056px] relative text-[#000] font-serif select-text text-left">
      {/* Dynamic custom background letterhead banner if custom is chosen */}
      {letterheadStyle === 'custom' && customLetterhead && customLetterhead.length > 0 && (
        <div className="absolute inset-0 z-0 pointer-events-none overflow-hidden">
          <img 
            src={customLetterhead[0]} 
            className="w-full h-full object-fill opacity-100" 
            alt="Fondo membrete personalizado Pág 1" 
            referrerPolicy="no-referrer"
          />
        </div>
      )}

      {/* Brand Watermarks for Page 1 */}
      {letterheadStyle === 'standard' && selectedBrand === 'isuzu' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.03] z-0">
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">MU-X</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">D-MAX</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">TRUCKS</div>
        </div>
      )}
      {letterheadStyle === 'standard' && selectedBrand === 'gac' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.025] z-0">
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMZOOM</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMKOO</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">GS8</div>
        </div>
      )}

      <DocHeader brand={selectedBrand} letterheadStyle={letterheadStyle} />
      
      <div className="text-center mt-3 mb-3 pb-0.5" style={{ borderBottom: `2px solid ${brandColor}` }}>
        <h2 className="text-sm font-extrabold uppercase tracking-widest m-0 text-black" style={{ fontFamily: 'system-ui, sans-serif' }}>
          ENTREGA DE VEHICULO AL CLIENTE
        </h2>
      </div>

      {/* METRIC ROW */}
      <div className="grid grid-cols-2 gap-x-6 gap-y-1 bg-[#fdfdfd] border border-slate-200 py-1.5 px-3 mb-1.5 text-[9px] relative z-10 rounded shadow-sm">
        <div className="space-y-1">
          <div className="flex items-end"><span className="font-bold text-black w-24 uppercase tracking-wider text-[8px] font-sans">CLIENTE:</span><span className="border-b border-slate-300 pb-0.5 flex-1 font-sans font-bold text-black truncate">{compVars.CLIENTE}</span></div>
          <div className="flex items-end"><span className="font-bold text-black w-24 uppercase tracking-wider text-[8px] font-sans">TIPO DE VEHÍCULO:</span><span className="border-b border-slate-300 pb-0.5 flex-1 font-sans font-bold text-black truncate">{compVars.MODELO} {compVars.TIPO_VEHICULO ? `(${compVars.TIPO_VEHICULO})` : ''}</span></div>
          <div className="flex items-end"><span className="font-bold text-black w-24 uppercase tracking-wider text-[8px] font-sans">FECHA:</span><span className="border-b border-slate-300 pb-0.5 flex-1 font-sans text-black">{compVars.FECHA_ENTREGA}</span></div>
        </div>
        <div className="space-y-1">
          <div className="flex items-end"><span className="font-bold text-black w-20 uppercase tracking-wider text-[8px] font-sans">CHASIS:</span><span className="border-b border-slate-300 pb-0.5 flex-1 font-mono font-bold text-black">{compVars.CHASIS}</span></div>
          <div className="flex items-end"><span className="font-bold text-black w-20 uppercase tracking-wider text-[8px] font-sans">MOTOR:</span><span className="border-b border-slate-300 pb-0.5 flex-1 font-mono text-black">{compVars.MOTOR}</span></div>
          <div className="flex items-end"><span className="font-bold text-black w-20 uppercase tracking-wider text-[8px] font-sans">COLOR:</span><span className="border-b border-slate-300 pb-0.5 flex-1 font-sans text-black">{compVars.COLOR}</span></div>
        </div>
      </div>

      <p className="text-[9px] leading-relaxed mb-1.5 text-justify relative z-10 text-black">
        El presente documento detalla las inspecciones conjuntas realizadas sobre la unidad previo a su entrega oficial:
      </p>

      {/* CHECKLIST GRID IN COLUMNS - LEFT BLANK AS REQUESTED */}
      <div className="grid grid-cols-2 gap-x-6 gap-y-1.5 mt-0.5 text-[9px] relative z-10">
        
        {checklist.slice(0, 3).map((section, sIdx) => (
          <div key={section.title} className="space-y-0.5">
            <div className="font-bold uppercase text-[8.5px] border-b pb-[1px]" style={{ color: 'black', borderColor: brandColor }}>
              {section.title}
            </div>
            <div className="space-y-0.5">
              {section.items.map((item, iIdx) => (
                <div 
                  key={item.id} 
                  onClick={() => toggleCheck(sIdx, iIdx)}
                  className="flex items-center justify-between py-[1px] cursor-pointer hover:bg-slate-50 transition-colors select-none"
                >
                  <span className="text-[8.5px] text-black leading-tight pr-2">{item.label}</span>
                  <div className="w-3.5 h-3.5 border border-black rounded flex items-center justify-center shrink-0 bg-white font-sans font-bold text-[8px]">
                    {/* Kept totally blank for manual fill out as requested! */}
                  </div>
                </div>
              ))}
            </div>
          </div>
        ))}

        {checklist.slice(3).map((section, sIdx) => {
          const actualSectionIdx = sIdx + 3;
          return (
            <div key={section.title} className="space-y-0.5">
              <div className="font-bold uppercase text-[8.5px] border-b pb-[1px]" style={{ color: 'black', borderColor: brandColor }}>
                {section.title}
              </div>
              <div className="space-y-0.5">
                {section.items.map((item, iIdx) => (
                  <div 
                    key={item.id} 
                    onClick={() => toggleCheck(actualSectionIdx, iIdx)}
                    className="flex items-center justify-between py-[1px] cursor-pointer hover:bg-slate-50 transition-colors select-none"
                  >
                    <span className="text-[8.5px] text-black leading-tight pr-2">{item.label}</span>
                    <div className="w-3.5 h-3.5 border border-black rounded flex items-center justify-center shrink-0 bg-white font-sans font-bold text-[8px]">
                      {/* Kept totally blank for manual fill out as requested! */}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          );
        })}

      </div>

      <div className="mt-1.5 py-1 px-2 bg-slate-50 border border-slate-300 rounded text-[9px] text-black text-center font-bold tracking-wide uppercase leading-normal relative z-10 font-sans">
        HAGO CONSTAR QUE HE RECIBIDO A MI ENTERA SASTIFACCION EL VEHICULO ARRIBA DESCRITO.
      </div>

      {/* SIGN BLOCK */}
      <div className="grid grid-cols-2 gap-12 mt-16 relative z-10">
        <div className="text-center pt-3 border-t border-slate-400">
          <div className="font-extrabold text-[10.5px] uppercase text-black leading-tight truncate">{compVars.CLIENTE}</div>
          <div className="text-[10px] text-black font-sans tracking-wider mt-0.5 uppercase font-semibold">Recibi Conforme</div>
          <div className="text-[9.5px] mt-0.5 text-black font-sans font-medium">Identificación: {compVars.CEDULA}</div>
        </div>
        <div className="text-center pt-3 border-t border-slate-400">
          <div className="font-extrabold text-[10.5px] uppercase text-black leading-tight truncate">{compVars.VENDEDOR}</div>
          <div className="text-[10px] text-black font-sans tracking-wider mt-0.5 uppercase font-semibold">Entregue</div>
          <div className="text-[9.5px] mt-0.5 text-black font-sans font-medium">Servicio Automotriz Mántica S.A.</div>
        </div>
      </div>

      <DocFooter brand={selectedBrand} pageNum={1} letterheadStyle={letterheadStyle} />
    </div>
  );
};

// PÁGINA 2: CONSTANCIA DE RECEPCIÓN DE VEHÍCULO Y ACCESORIOS
const DocumentPage2 = ({ compVars, selectedBrand, letterheadStyle, customLetterhead, accessories, setAccessories }: DocProps & { accessories: { id: string; label: string; checked: boolean }[], setAccessories: React.Dispatch<React.SetStateAction<{ id: string; label: string; checked: boolean }[]>> }) => {
  const isGac = selectedBrand === 'gac';
  const brandName = isGac ? 'GAC MOTOR' : 'ISUZU';
  const brandColor = '#df1010';

  const toggleAccessory = (idx: number) => {
    const copy = [...accessories];
    copy[idx].checked = !copy[idx].checked;
    setAccessories(copy);
  };

  return (
    <div className="pt-3 px-8 pb-36 min-h-[1056px] relative text-[#000] font-serif select-text text-left">
      {/* Dynamic custom background letterhead banner if custom is chosen */}
      {letterheadStyle === 'custom' && customLetterhead && customLetterhead.length > 0 && (
        <div className="absolute inset-0 z-0 pointer-events-none overflow-hidden">
          <img 
            src={customLetterhead[1] || customLetterhead[0]} 
            className="w-full h-full object-fill opacity-100" 
            alt="Fondo membrete personalizado Pág 2" 
            referrerPolicy="no-referrer"
          />
        </div>
      )}

      {/* Brand Watermarks for Page 2 */}
      {letterheadStyle === 'standard' && selectedBrand === 'isuzu' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.03] z-0">
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">MU-X</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">D-MAX</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">TRUCKS</div>
        </div>
      )}
      {letterheadStyle === 'standard' && selectedBrand === 'gac' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.025] z-0">
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMZOOM</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMKOO</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">GS8</div>
        </div>
      )}

      <DocHeader brand={selectedBrand} letterheadStyle={letterheadStyle} />
      
      <div className="text-center mt-3 mb-3 pb-0.5" style={{ borderBottom: `2px solid ${brandColor}` }}>
        <h2 className="text-base font-extrabold uppercase tracking-widest m-0 text-black" style={{ fontFamily: 'system-ui, sans-serif' }}>
          ENTREGA DE VEHICULO NUEVO
        </h2>
      </div>

      <p className="text-[10.5px] leading-relaxed mb-4 text-justify text-black relative z-10">
        Hago constar que en esta fecha estoy recibiendo de Sres. <strong>SERVICIO AUTOMOTRIZ MANTICA, S. A.</strong>, (1) Vehículo abajo descrito, a mi entera satisfacción y en perfecto estado de funcionamiento, habiendo pasado inspección por él, tanto en su estructura interior and exterior así como la constatación de los accesorios que me entregan:
      </p>

      {/* VEHICLE MAIN SPEC TABLE */}
      <div className="relative z-10 mb-4 select-text">
        <table className="w-full border-collapse border-2 border-slate-800 text-[10.5px] text-black">
          <thead>
            <tr className="bg-slate-100">
              <th className="border border-slate-700 px-3 py-1.5 uppercase font-sans font-bold tracking-wider text-black text-left">MODELO</th>
              <th className="border border-slate-700 px-3 py-1.5 uppercase font-sans font-bold tracking-wider text-black text-left">CHASIS</th>
              <th className="border border-slate-700 px-3 py-1.5 uppercase font-sans font-bold tracking-wider text-black text-left">MOTOR</th>
              <th className="border border-slate-700 px-3 py-1.5 uppercase font-sans font-bold tracking-wider text-black text-left">COLOR</th>
            </tr>
          </thead>
          <tbody>
            <tr className="bg-white">
              <td className="border border-slate-700 px-3 py-1.5 font-bold font-sans text-black">{compVars.MODELO}</td>
              <td className="border border-slate-700 px-3 py-1.5 font-bold font-mono text-black">{compVars.CHASIS}</td>
              <td className="border border-slate-700 px-3 py-1.5 font-mono text-black">{compVars.MOTOR}</td>
              <td className="border border-slate-700 px-3 py-1.5 uppercase font-sans text-black">{compVars.COLOR}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <div className="col-span-1 space-y-2 mb-4 relative z-10">
        <div className="text-[10px] font-sans font-bold uppercase tracking-wider mb-1 text-black" style={{ color: 'black' }}>
          LISTA DE ACCESORIOS
        </div>

        {/* COMPACT TABLE VIEW - WITH CLEAN BLANK CHECKBOXES FOR MANUAL TICKING */}
        <div className="grid grid-cols-2 gap-x-6 gap-y-1.5 border border-slate-200 rounded p-3 text-[10.5px] bg-[#fafafa]">
          {accessories.map((acc, idx) => (
            <div 
              key={acc.id} 
              onClick={() => toggleAccessory(idx)}
              className="flex items-center justify-between py-1 border-b border-dashed border-slate-200 cursor-pointer select-none hover:bg-white"
            >
              <span className="font-medium text-black">{acc.label}</span>
              <div className="flex items-center gap-2">
                <div className="w-3.5 h-3.5 border border-slate-500 rounded flex items-center justify-center shrink-0 bg-white">
                  {/* Kept totally blank for manual fill out as requested! */}
                </div>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* METRIC ROWS SIDE-BY-SIDE */}
      <div className="grid grid-cols-3 gap-4 mb-4 relative z-10 text-[10px] text-black">
        <div className="col-span-1 border border-slate-200 bg-[#fafafa] rounded px-2 py-3 text-center flex flex-col justify-between min-h-[56px]">
          <span className="font-bold text-black uppercase text-[8.5px] block font-sans leading-none">LLAVE No</span>
          <span className="border-b border-dashed border-slate-400 w-16 mx-auto h-4 block"></span>
        </div>
        <div className="col-span-1 border border-slate-200 bg-[#fafafa] rounded px-2 py-3 text-center flex flex-col justify-between min-h-[56px]">
          <span className="font-bold text-black uppercase text-[8.5px] block font-sans leading-none">KM RECORRIDOS</span>
          <span className="border-b border-dashed border-slate-400 w-20 mx-auto h-4 block"></span>
        </div>
        <div className="col-span-1 border border-slate-200 bg-[#fafafa] rounded px-2 py-3 text-center flex flex-col justify-between min-h-[56px]">
          <span className="font-bold text-black uppercase text-[8.5px] block font-sans leading-none">Código Factura</span>
          <span className="font-extrabold text-black text-[11px] font-mono leading-none">{compVars.NO_PROFORMA || 'N/D'}</span>
        </div>
      </div>

      {/* OBSERVATIONS AREA */}
      <div className="border border-slate-300 bg-white rounded p-4 text-[10.5px] relative z-10 space-y-1.5">
        <div className="font-sans font-bold text-[9px] text-black uppercase tracking-wider">
          OBSERVACIONES:
        </div>
        <p className="leading-relaxed text-black text-justify font-sans whitespace-pre-line text-[10.5px]">
          {compVars.OBS}
        </p>
      </div>

      <p className="text-[9.5px] text-black mt-4 leading-normal relative z-10 font-sans">
        FECHA DE CONTROL DE ENTREGA: {compVars.FECHA_ENTREGA}
      </p>

      {/* SIGN BLOCKS - 3 COMPONENT SIZES */}
      <div className="grid grid-cols-3 gap-6 mt-16 text-center text-[10px] relative z-10">
        <div className="pt-4 border-t border-slate-400">
          <div className="font-extrabold truncate text-black text-[10.5px]">{compVars.CLIENTE}</div>
          <div className="text-[10px] text-black font-sans tracking-wide mt-1 uppercase font-semibold">Recibi Conforme</div>
        </div>
        <div className="pt-4 border-t border-slate-400">
          <div className="font-extrabold text-black text-[10.5px]">Carlos J. Morales U.</div>
          <div className="text-[10px] text-black font-sans tracking-wide mt-1 uppercase font-semibold">Vo. Bo.</div>
        </div>
        <div className="pt-4 border-t border-slate-400">
          <div className="font-extrabold truncate text-black text-[10.5px]">{compVars.VENDEDOR}</div>
          <div className="text-[10px] text-black font-sans tracking-wide mt-1 uppercase font-semibold">Entregue</div>
        </div>
      </div>

      <DocFooter brand={selectedBrand} pageNum={2} letterheadStyle={letterheadStyle} />
    </div>
  );
};

// PÁGINA 3: CONDICIONES DE GARANTÍA LEGAL Y TALLER
const DocumentPage3 = ({ compVars, selectedBrand, letterheadStyle, customLetterhead }: DocProps) => {
  const isGac = selectedBrand === 'gac';
  const brandName = isGac ? 'GAC MOTOR' : 'ISUZU';
  const brandColor = '#df1010';

  return (
    <div className="pt-3 px-8 pb-36 min-h-[1056px] relative text-[#000] font-serif select-text text-left">
      {/* Dynamic custom background letterhead banner if custom is chosen */}
      {letterheadStyle === 'custom' && customLetterhead && customLetterhead.length > 0 && (
        <div className="absolute inset-0 z-0 pointer-events-none overflow-hidden">
          <img 
            src={customLetterhead[2] || customLetterhead[0]} 
            className="w-full h-full object-fill opacity-100" 
            alt="Fondo membrete personalizado Pág 3" 
            referrerPolicy="no-referrer"
          />
        </div>
      )}

      {/* Brand Watermarks for Page 3 */}
      {letterheadStyle === 'standard' && selectedBrand === 'isuzu' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.03] z-0">
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">MU-X</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">D-MAX</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">TRUCKS</div>
        </div>
      )}
      {letterheadStyle === 'standard' && selectedBrand === 'gac' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.025] z-0">
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMZOOM</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMKOO</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">GS8</div>
        </div>
      )}

      <DocHeader brand={selectedBrand} letterheadStyle={letterheadStyle} />
      
      <div className="text-center mt-3 mb-3 pb-0.5" style={{ borderBottom: `2px solid ${brandColor}` }}>
        <h2 className="text-base font-extrabold uppercase tracking-widest m-0 text-black" style={{ fontFamily: 'system-ui, sans-serif' }}>
          CONDICIONES DE GARANTIA DE FABRICACION
        </h2>
      </div>

      {/* PAGE 3 METADATA AREA */}
      <div className="grid grid-cols-2 gap-x-6 gap-y-1.5 border border-slate-300 p-3 mb-4 text-[10px] bg-slate-50 rounded select-text z-10 relative text-black">
        <div className="space-y-1">
          <div className="flex"><span className="font-bold text-black w-28 font-sans text-[8.5px] tracking-wider">SEÑOR(ES):</span><strong className="flex-1 text-black">{compVars.CLIENTE}</strong></div>
          <div className="flex"><span className="font-bold text-black w-28 font-sans text-[8.5px] tracking-wider">DIRECCIÓN:</span><span className="flex-1 uppercase font-semibold text-black">{compVars.DEPARTAMENTO}</span></div>
          <div className="flex"><span className="font-bold text-black w-28 font-sans text-[8.5px] tracking-wider">No. DE FACTURA:</span><strong className="flex-1 text-black font-mono">{compVars.NO_PROFORMA}</strong></div>
        </div>
        <div className="space-y-1">
          <div className="flex"><span className="font-bold text-black w-20 font-sans text-[8.5px] tracking-wider">MODELO:</span><span className="flex-1 font-semibold text-black">{compVars.MODELO}</span></div>
          <div className="flex"><span className="font-bold text-black w-20 font-sans text-[8.5px] tracking-wider">CHASIS:</span><strong className="flex-1 font-mono text-black">{compVars.CHASIS}</strong></div>
          <div className="flex"><span className="font-bold text-black w-20 font-sans text-[8.5px] tracking-wider">MOTOR:</span><span className="flex-1 font-mono text-black">{compVars.MOTOR}</span></div>
          <div className="flex"><span className="font-bold text-black w-20 font-sans text-[8.5px] tracking-wider">COLOR:</span><span className="flex-1 uppercase font-semibold text-black">{compVars.COLOR}</span></div>
        </div>
      </div>

      <div className="space-y-3.5 text-[10.5px] leading-relaxed text-black text-justify relative z-10 font-sans">
        <p>
          Todos los vehículos nuevos cuentan con una garantía de fábrica que a continuación le describimos.
        </p>
        <p>
          Todos los chequeos de mantenimiento posteriores de acuerdo a recomendaciones del fabricante (cada 5,000 Kilómetros) son pagados en su totalidad. (Mano de Obra y Aceite y Lubricantes)
        </p>

        <p className="font-semibold p-2.5 border-l-2 bg-[#fdfdfd] text-black text-[10.5px]" style={{ borderColor: brandColor }}>
          {isGac ? (
            <span>La garantía de funcionamiento de los vehículos GAC MOTOR y distribuido por SERVICIO AUTOMOTRIZ MANTICA S.A. será por (96) Noventa y seis meses o hasta que recorran una distancia de CIENTO CINCUENTA MIL (150,000) kilómetros, lo que se venza primero después de la fecha de entrega.</span>
          ) : (
            <span>La garantía de funcionamiento de los vehículos ISUZU y distribuido por SERVICIO AUTOMOTRIZ MANTICA S.A. será por (36) Treinta y seis meses o hasta que recorran una distancia de CIEN MIL (100,000) kilómetros, lo que se venza primero después de la fecha de entrega.</span>
          )}
        </p>

        <p>
          Asimismo, nos comprometemos a reponer gratuitamente, cualquier pieza o conjunto de piezas que fallaran, se rompan o desgasten prematuramente debido al diseño, material o fabricación defectuosa o mal montaje, dentro del período de garantía siempre que el uso y servicio del vehículo haya sido normal y que estos defectos nos sean comunicados por ustedes oportunamente.
        </p>

        <p>
          Para que la garantía se mantenga vigente es necesario dar el mantenimiento recomendado por el fabricante cada 5,000 km. En nuestros talleres de Auto Mantica como distribuidor autorizado por el fabricante, ya que debido a la falta de este la garantía se pierde.
        </p>

        <div className="space-y-1 bg-[#fafafa] border border-slate-200 rounded p-2.5 text-[10px]">
          <p className="font-bold text-black font-sans uppercase tracking-wider text-[9px]">La garantía no será aplicable en los siguientes casos:</p>
          <ul className="list-disc pl-5 space-y-1 text-black text-[9.5px]">
            <li>Averías y/o defectos en los productos debido a cualquier servicio realizado en un taller de reparaciones que no sea el de AUTOMOTRIZ MANTICA S.A.</li>
            <li>Averías y/o defectos en los productos ocasionado por el hecho de que el usuario realice algún cambio, modificaciones o alteraciones de los productos sin haber obtenido antes el consentimiento por escrito o las apropiadas instrucciones de parte nuestra.</li>
            <li>Averías y/o defectos en los productos debido a la operación que exceda de las capacidades nominales estipulada por el fabricante.</li>
            <li>Averías y/o defectos ocasionados por el mal uso en la operación diaria de la unidad.</li>
            <li>Averías y/o defectos que sean consecuencias del mal uso, accidentes o negligencia.</li>
            <li>La garantía no incluye los componentes cuyo desgaste es normal durante la operación dentro y fuera del período de garantía: Pastillas de frenos, Clutch, Llantas, Batería, Aceite y Lubricantes, Filtros, Chisperos o Toberas de Inyección, amortiguadores, tricos, etc.</li>
            <li>La fábrica no se hace responsable por componentes que no sean de fabricación propia sino; de otros concesionarios que se dedican a ellos en gran escala Por Ejem: Baterías, Llantas y Equipo de sonido.</li>
            <li>Tenga presente que las reparaciones causadas por el uso de combustible o aceites indebidos, sucios o contaminados no es cubierta por la garantía.</li>
          </ul>
        </div>
      </div>

      <p className="text-[9.5px] text-black mt-4 leading-normal relative z-10 font-sans">
        FECHA: {compVars.FECHA_ENTREGA}
      </p>

      {/* SIGNBLOCK */}
      <div className="grid grid-cols-2 gap-12 mt-16 text-center text-[10px] relative z-10 font-sans">
        <div className="pt-4 border-t border-slate-400">
          <div className="font-extrabold text-black text-[10.5px]">Lic. Carlos Morales</div>
          <div className="text-[10px] text-black uppercase tracking-wider mt-0.5 font-semibold">Departamento de Vehículos</div>
          <div className="text-[9.5px] text-black mt-1 font-mono">Auto Mántica S.A.</div>
        </div>
        <div className="pt-4 border-t border-slate-400">
          <div className="font-extrabold text-black text-[10.5px] truncate">{compVars.CLIENTE}</div>
          <div className="text-[10px] text-black uppercase tracking-wider mt-0.5 font-semibold">Recibo Conforme</div>
          <div className="text-[9.5px] text-black mt-1 font-mono">Identificación: {compVars.CEDULA}</div>
        </div>
      </div>

      <DocFooter brand={selectedBrand} pageNum={3} letterheadStyle={letterheadStyle} />
    </div>
  );
};

// PÁGINA 4: ORDEN DE SALIDA DE VEHÍCULO (CONTROL DE PORTÓN / VIGILANTE)
const DocumentPage4 = ({ compVars, selectedBrand, letterheadStyle, customLetterhead }: DocProps) => {
  const isGac = selectedBrand === 'gac';
  const brandName = isGac ? 'GAC MOTOR' : 'ISUZU';
  const brandColor = '#df1010';

  return (
    <div className="pt-3 px-8 pb-36 min-h-[1056px] relative text-[#000] font-serif select-text text-left">
      {/* Dynamic custom background letterhead banner if custom is chosen */}
      {letterheadStyle === 'custom' && customLetterhead && customLetterhead.length > 0 && (
        <div className="absolute inset-0 z-0 pointer-events-none overflow-hidden">
          <img 
            src={customLetterhead[3] || customLetterhead[0]} 
            className="w-full h-full object-fill opacity-100" 
            alt="Fondo membrete personalizado Pág 4" 
            referrerPolicy="no-referrer"
          />
        </div>
      )}

      {/* Brand Watermarks for Page 4 */}
      {letterheadStyle === 'standard' && selectedBrand === 'isuzu' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.03] z-0">
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">MU-X</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">D-MAX</div>
          <div className="text-[110px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">TRUCKS</div>
        </div>
      )}
      {letterheadStyle === 'standard' && selectedBrand === 'gac' && (
        <div className="absolute inset-x-0 top-[180px] bottom-[180px] flex flex-col justify-around items-center select-none pointer-events-none opacity-[0.025] z-0">
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMZOOM</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">EMKOO</div>
          <div className="text-[100px] font-sans font-black tracking-[0.15em] text-[#000] leading-none">GS8</div>
        </div>
      )}

      <DocHeader brand={selectedBrand} letterheadStyle={letterheadStyle} />
      
      <div className="text-center mt-3 mb-3 pb-0.5" style={{ borderBottom: `2px solid ${brandColor}` }}>
        <h2 className="text-base font-extrabold uppercase tracking-widest m-0 text-black" style={{ fontFamily: 'system-ui, sans-serif' }}>
          ORDEN DE SALIDA
        </h2>
      </div>

      <div className="space-y-4 my-4 text-[10.5px] leading-relaxed relative z-10 font-sans text-black">
        
        <div className="space-y-0.5 border-l-2 pl-3" style={{ borderColor: brandColor }}>
          <div className="font-bold text-black text-[11px]">Señores:</div>
          <div className="uppercase font-extrabold text-black text-[10px] tracking-wider">Encargado de Seguridad</div>
        </div>

        <p className="text-justify font-sans text-black leading-relaxed text-[11px]">
          Se autoriza a <strong className="font-bold uppercase text-black font-sans">{compVars.CLIENTE}</strong> la salida del vehiculo con la descripcion siguiente:
        </p>

        {/* DOUBLE COLUMN BOX LISTING THE SPEC VALUES */}
        <div className="border border-slate-300 p-5 rounded bg-[#fcfcfc] bg-opacity-95 shadow-sm text-black">
          <div className="grid grid-cols-2 gap-y-3.5 gap-x-8 text-[11px] font-sans">
            <div className="flex justify-between border-b border-slate-200 pb-1.5"><span className="font-bold text-black uppercase text-[9px] tracking-wider font-sans">MARCA:</span><strong className="text-black uppercase">{brandName}</strong></div>
            <div className="flex justify-between border-b border-slate-200 pb-1.5"><span className="font-bold text-black uppercase text-[9px] tracking-wider font-sans">MODELO:</span><strong className="text-black uppercase">{compVars.MODELO}</strong></div>
            
            <div className="flex justify-between border-b border-slate-200 pb-1.5"><span className="font-bold text-black uppercase text-[9px] tracking-wider font-sans">CHASIS:</span><strong className="text-black font-mono font-extrabold text-[11.5px]">{compVars.CHASIS}</strong></div>
            <div className="flex justify-between border-b border-slate-200 pb-1.5"><span className="font-bold text-black uppercase text-[9px] tracking-wider font-sans">MOTOR:</span><strong className="text-black font-mono">{compVars.MOTOR}</strong></div>
            
            <div className="flex justify-between border-b border-slate-200 pb-1.5"><span className="font-bold text-black uppercase text-[9px] tracking-wider font-sans">COLOR:</span><strong className="text-black uppercase">{compVars.COLOR}</strong></div>
            <div className="flex justify-between border-b border-slate-200 pb-1.5"><span className="font-bold text-black uppercase text-[9px] tracking-wider font-sans">AÑO:</span><strong className="text-black font-mono">{compVars.ANIO}</strong></div>
          </div>
        </div>

        {/* FACTURACION LINKAGE */}
        <div className="bg-[#fafafa] border border-slate-300 rounded p-4 text-[10.5px] space-y-2 relative z-10 font-sans shadow-sm text-black">
          <div className="flex gap-2 items-baseline"><span className="font-bold text-black uppercase text-[9px] tracking-wider w-48 shrink-0">Entrega por compra amparada en factura No.</span><strong className="text-black font-mono font-bold text-[11px]">{compVars.NO_PROFORMA}</strong></div>
          <div className="flex gap-2 items-baseline"><span className="font-bold text-black uppercase text-[9px] tracking-wider w-48 shrink-0">A nombre de:</span><strong className="text-black text-[11px]">{compVars.CLIENTE}</strong></div>
          <div className="flex gap-2 items-baseline"><span className="font-bold text-black uppercase text-[9px] tracking-wider w-48 shrink-0">fecha de factura:</span><strong className="text-black font-mono text-[11px]">{compVars.FECHA_FACTURA}</strong></div>
        </div>

      </div>

      {/* SIGNBLOCK */}
      <div className="grid grid-cols-2 gap-12 mt-20 text-center text-[10.5px] relative z-10 font-sans">
        <div className="pt-6 border-t border-slate-400">
          <div className="font-extrabold text-black text-[11px]">Lic. Carlos Morales</div>
          <div className="text-[10px] text-black uppercase tracking-wider mt-0.5 font-semibold">Departamento de Vehículos</div>
          <div className="text-[9.5px] text-black mt-1 font-mono">Firma Autorizada</div>
        </div>
        <div className="pt-6 border-t border-slate-400">
          <div className="font-extrabold text-black text-[11px] truncate">{compVars.VENDEDOR}</div>
          <div className="text-[10px] text-black uppercase tracking-wider mt-0.5 font-semibold">Recibo Conforme</div>
        </div>
      </div>

      <DocFooter brand={selectedBrand} pageNum={4} letterheadStyle={letterheadStyle} />
    </div>
  );
};

