import { useState, useEffect, useMemo } from 'react'
import * as XLSX from 'xlsx'
import {
  Users,
  UserCheck,
  UserX,
  BarChart2,
  RefreshCw,
  LinkIcon,
  AlertCircle,
  Loader2,
  Calendar,
  CalendarRange,
  Clock,
  TrendingUp,
  Filter,
  Info,
  ChevronDown,
} from 'lucide-react'

const FUENTES = [
  'Página Web',
  'Facebook',
  'Amistad',
  'Anuncio',
  'Instagram',
  'Google',
  'Ubicación',
]

const mockCsvData = [
  { Id: '1', Nombre: 'María García', Email: 'maria.garcia@email.com', Teléfono: '555-0101', Género: 'Female', 'Fecha de nacimiento': '1990-05-15', Membresía: 'Paquete 8 Sesiones', Estatus: 'active', 'Fecha inicio': '2025-01-10', 'Fecha fin': '2025-04-10', 'Fecha registro': '10-01-2025', 'Horario registro': '09:30', Sucursal: 'Pilates Center', Fuente: 'Página Web' },
  { Id: '2', Nombre: 'Carlos López', Email: 'carlos.lopez@email.com', Teléfono: '555-0102', Género: 'Male', 'Fecha de nacimiento': '1985-08-22', Membresía: 'Clase de Prueba', Estatus: 'active', 'Fecha inicio': '2025-01-12', 'Fecha fin': '2025-01-12', 'Fecha registro': '12-01-2025', 'Horario registro': '14:00', Sucursal: 'Pilates Center', Fuente: 'Facebook' },
  { Id: '3', Nombre: 'Ana Martínez', Email: 'ana.martinez@email.com', Teléfono: '555-0103', Género: 'Female', 'Fecha de nacimiento': '1992-03-08', Membresía: 'Paquete 12 Sesiones', Estatus: 'active', 'Fecha inicio': '2025-01-15', 'Fecha fin': '2025-04-15', 'Fecha registro': '15-01-2025', 'Horario registro': '11:45', Sucursal: 'Pilates Center', Fuente: 'Amistad' },
  { Id: '4', Nombre: 'Pedro Sánchez', Email: 'pedro.sanchez@email.com', Teléfono: '555-0104', Género: 'Male', 'Fecha de nacimiento': '1988-11-30', Membresía: 'Clase Individual', Estatus: 'inactive', 'Fecha inicio': '2024-06-01', 'Fecha fin': '2024-09-01', 'Fecha registro': '01-06-2024', 'Horario registro': '16:20', Sucursal: 'Pilates Center', Fuente: 'Instagram' },
  { Id: '5', Nombre: 'Laura Rodríguez', Email: 'laura.rodriguez@email.com', Teléfono: '555-0105', Género: 'Female', 'Fecha de nacimiento': '1995-07-19', Membresía: 'Paquete 16 Sesiones', Estatus: 'active', 'Fecha inicio': '2025-02-01', 'Fecha fin': '2025-06-01', 'Fecha registro': '01-02-2025', 'Horario registro': '10:00', Sucursal: 'Pilates Center', Fuente: 'Google' },
  { Id: '6', Nombre: 'Miguel Torres', Email: 'miguel.torres@email.com', Teléfono: '555-0106', Género: 'Male', 'Fecha de nacimiento': '1982-12-05', Membresía: 'Paquete 8 Sesiones', Estatus: 'inactive', 'Fecha inicio': '2024-08-15', 'Fecha fin': '2024-11-15', 'Fecha registro': '15-08-2024', 'Horario registro': '18:30', Sucursal: 'Pilates Center', Fuente: 'Anuncio' },
  { Id: '7', Nombre: 'Sofia Hernández', Email: 'sofia.hernandez@email.com', Teléfono: '555-0107', Género: 'Female', 'Fecha de nacimiento': '1993-09-12', Membresía: 'Paquete 4 Sesiones', Estatus: 'active', 'Fecha inicio': '2025-02-10', 'Fecha fin': '2025-03-10', 'Fecha registro': '10-02-2025', 'Horario registro': '12:15', Sucursal: 'Pilates Center', Fuente: 'Ubicación' },
  { Id: '8', Nombre: 'Jorge Díaz', Email: 'jorge.diaz@email.com', Teléfono: '555-0108', Género: 'Male', 'Fecha de nacimiento': '1987-04-25', Membresía: 'Clase de Prueba', Estatus: 'active', 'Fecha inicio': '2025-02-12', 'Fecha fin': '2025-02-12', 'Fecha registro': '12-02-2025', 'Horario registro': '15:45', Sucursal: 'Pilates Center', Fuente: 'Página Web' },
  { Id: '9', Nombre: 'Carmen Ruiz', Email: 'carmen.ruiz@email.com', Teléfono: '555-0109', Género: 'Female', 'Fecha de nacimiento': '1991-06-18', Membresía: 'Paquete 12 Sesiones', Estatus: 'active', 'Fecha inicio': '2025-02-20', 'Fecha fin': '2025-05-20', 'Fecha registro': '20-02-2025', 'Horario registro': '09:00', Sucursal: 'Pilates Center', Fuente: 'Facebook' },
  { Id: '10', Nombre: 'Antonio Gómez', Email: 'antonio.gomez@email.com', Teléfono: '555-0110', Género: 'Male', 'Fecha de nacimiento': '1979-01-10', Membresía: 'Paquete 8 Sesiones', Estatus: 'inactive', 'Fecha inicio': '2024-10-01', 'Fecha fin': '2025-01-01', 'Fecha registro': '01-10-2024', 'Horario registro': '17:00', Sucursal: 'Pilates Center', Fuente: 'Amistad' },
  { Id: '11', Nombre: 'Isabel Flores', Email: 'isabel.flores@email.com', Teléfono: '555-0111', Género: 'Female', 'Fecha de nacimiento': '1994-10-22', Membresía: 'Clase Individual', Estatus: 'active', 'Fecha inicio': '2025-03-01', 'Fecha fin': '2025-03-01', 'Fecha registro': '01-03-2025', 'Horario registro': '11:30', Sucursal: 'Pilates Center', Fuente: 'Instagram' },
  { Id: '12', Nombre: 'Roberto Castro', Email: 'roberto.castro@email.com', Teléfono: '555-0112', Género: 'Male', 'Fecha de nacimiento': '1986-02-14', Membresía: 'Paquete 16 Sesiones', Estatus: 'active', 'Fecha inicio': '2025-03-05', 'Fecha fin': '2025-07-05', 'Fecha registro': '05-03-2025', 'Horario registro': '14:20', Sucursal: 'Pilates Center', Fuente: 'Google' },
]

function parseFechaRegistro(str) {
  if (!str) return null
  const parts = str.trim().split('-')
  if (parts.length !== 3) return null
  const [dia, mes, anio] = parts.map(Number)
  return new Date(anio, mes - 1, dia)
}

function parseFechaISO(str) {
  if (!str) return null
  const parts = str.trim().split('-')
  if (parts.length !== 3) return null
  const [anio, mes, dia] = parts.map(Number)
  return new Date(anio, mes - 1, dia)
}

function isSameDay(d1, d2) {
  if (!d1 || !d2) return false
  return d1.getFullYear() === d2.getFullYear() &&
    d1.getMonth() === d2.getMonth() &&
    d1.getDate() === d2.getDate()
}

function getWeekStart(d) {
  const fecha = new Date(d)
  const dia = fecha.getDay()
  const diff = fecha.getDate() - dia + (dia === 0 ? -6 : 1)
  return new Date(fecha.setDate(diff))
}

function isSameWeek(d1, d2) {
  if (!d1 || !d2) return false
  const w1 = getWeekStart(d1)
  const w2 = getWeekStart(d2)
  return isSameDay(w1, w2)
}

function isSameMonth(d1, d2) {
  if (!d1 || !d2) return false
  return d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth()
}

function extractFileId(url) {
  if (!url) return null
  
  const sheetsRegex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]{25,100})/
  const sheetsMatch = url.match(sheetsRegex)
  if (sheetsMatch) return sheetsMatch[1]
  
  const driveRegex = /\/file\/d\/([a-zA-Z0-9-_]{25,100})/
  const driveMatch = url.match(driveRegex)
  if (driveMatch) return driveMatch[1]
  
  const driveIdRegex = /[?&]id=([a-zA-Z0-9-_]{25,100})/
  const driveIdMatch = url.match(driveIdRegex)
  if (driveIdMatch) return driveIdMatch[1]
  
  return null
}

function normalizeKeys(row) {
  const normalized = {}
  Object.entries(row).forEach(([key, value]) => {
    const normKey = key.trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .toLowerCase();
    
    if (normKey === 'id') normalized['Id'] = value;
    else if (normKey === 'nombre') normalized['Nombre'] = value;
    else if (normKey === 'email' || normKey === 'correo') normalized['Email'] = value;
    else if (normKey === 'telefono' || normKey === 'celular') normalized['Teléfono'] = value;
    else if (normKey === 'genero' || normKey === 'sexo') normalized['Género'] = value;
    else if (normKey === 'fecha de nacimiento' || normKey === 'nacimiento' || normKey === 'fecha_nacimiento') normalized['Fecha de nacimiento'] = value;
    else if (normKey === 'membresia' || normKey === 'plan') normalized['Membresía'] = value;
    else if (normKey === 'estatus' || normKey === 'estado') normalized['Estatus'] = value;
    else if (normKey === 'fecha inicio' || normKey === 'inicio' || normKey === 'fecha_inicio') normalized['Fecha inicio'] = value;
    else if (normKey === 'fecha fin' || normKey === 'fin' || normKey === 'fecha_fin') normalized['Fecha fin'] = value;
    else if (normKey === 'fecha registro' || normKey === 'registro' || normKey === 'fecha_registro') normalized['Fecha registro'] = value;
    else if (normKey === 'horario registro' || normKey === 'horario' || normKey === 'hora_registro' || normKey === 'horario_registro') normalized['Horario registro'] = value;
    else if (normKey === 'sucursal') normalized['Sucursal'] = value;
    else if (normKey === 'fuente' || normKey === 'origen') normalized['Fuente'] = value;
    else {
      normalized[key.trim()] = value;
    }
  });
  return normalized;
}

function fixEncoding(str) {
  if (typeof str !== 'string') return str;
  try {
    const decoded = decodeURIComponent(escape(str));
    return decoded;
  } catch (e) {
    let cleaned = str;
    cleaned = cleaned
      .replace(/Ã¡/g, 'á')
      .replace(/Ã©/g, 'é')
      .replace(/Ã­/g, 'í')
      .replace(/Ã³/g, 'ó')
      .replace(/Ãº/g, 'ú')
      .replace(/Ã±/g, 'ñ')
      .replace(/Ã/g, 'Á')
      .replace(/Ã‰/g, 'É')
      .replace(/Ã/g, 'Í')
      .replace(/Ã“/g, 'Ó')
      .replace(/Ãš/g, 'Ú')
      .replace(/Ã‘/g, 'Ñ');
    return cleaned;
  }
}

function formatRowDate(val, expectedFormat = 'YYYY-MM-DD') {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    if (expectedFormat === 'DD-MM-YYYY') {
      return `${d}-${m}-${y}`;
    }
    return `${y}-${m}-${d}`;
  }
  
  if (typeof val === 'number') {
    const date = new Date(Math.round((val - 25569) * 86400 * 1000));
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    if (expectedFormat === 'DD-MM-YYYY') {
      return `${d}-${m}-${y}`;
    }
    return `${y}-${m}-${d}`;
  }
  
  let str = String(val).trim();
  str = fixEncoding(str);
  str = str.replace(/\//g, '-');
  if (str.includes(' ')) {
    str = str.split(' ')[0];
  }
  return str;
}

const fuenteColors = {
  'Página Web': 'bg-blue-100 text-blue-700',
  'Facebook': 'bg-indigo-100 text-indigo-700',
  'Amistad': 'bg-green-100 text-green-700',
  'Anuncio': 'bg-orange-100 text-orange-700',
  'Instagram': 'bg-pink-100 text-pink-700',
  'Google': 'bg-yellow-100 text-yellow-700',
  'Ubicación': 'bg-purple-100 text-purple-700',
}

export default function App() {
  const [csvUrl, setCsvUrl] = useState('')
  const [rawData, setRawData] = useState([])
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState(null)
  const [filterFuente, setFilterFuente] = useState('Todas')
  const [filterEstatus, setFilterEstatus] = useState('Todos')
  const [periodoMode, setPeriodoMode] = useState('dia')
  const [periodoFecha, setPeriodoFecha] = useState(new Date())
  const [periodoDesde, setPeriodoDesde] = useState('')
  const [periodoHasta, setPeriodoHasta] = useState('')
  const [periodoFuente, setPeriodoFuente] = useState('Todas')
  const [showTooltip, setShowTooltip] = useState(false)

  // Local state for bottom "Socios" table filters
  const [sociosFilterFuente, setSociosFilterFuente] = useState('Todas')
  const [sociosFilterEstatus, setSociosFilterEstatus] = useState('Todos')
  const [sociosDesde, setSociosDesde] = useState('')
  const [sociosHasta, setSociosHasta] = useState('')

  useEffect(() => {
    setRawData(mockCsvData)
  }, [])

  const handleFetch = async () => {
    if (!csvUrl.trim()) {
      setRawData(mockCsvData)
      return
    }
    setLoading(true)
    setError(null)

    const fileId = extractFileId(csvUrl);
    let urlsToTry = [csvUrl];

    if (fileId) {
      if (csvUrl.includes('spreadsheets')) {
        urlsToTry = [
          `https://docs.google.com/spreadsheets/d/${fileId}/export?format=csv`,
          `https://docs.google.com/uc?export=download&id=${fileId}`
        ];
      } else {
        urlsToTry = [
          `https://docs.google.com/uc?export=download&id=${fileId}`,
          `https://docs.google.com/spreadsheets/d/${fileId}/export?format=csv`
        ];
      }
    }

    let success = false;
    let lastError = null;

    for (const url of urlsToTry) {
      try {
        const response = await fetch(url)
        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`)
        
        const buffer = await response.arrayBuffer()
        const workbook = XLSX.read(new Uint8Array(buffer), {
          type: 'array',
          codepage: 65001,
          cellDates: true,
          cellNF: false,
          cellText: true
        })

        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" })

        if (!jsonData || jsonData.length === 0) {
          throw new Error('No se encontraron registros en el archivo')
        }

        const processedRows = jsonData.map(row => {
          const norm = normalizeKeys(row)
          
          Object.keys(norm).forEach(key => {
            if (typeof norm[key] === 'string') {
              norm[key] = fixEncoding(norm[key])
            }
          })
          
          if (norm['Estatus']) {
            const estVal = String(norm['Estatus']).trim().toLowerCase()
            if (estVal.includes('act') || estVal === 'active' || estVal === 'activo') {
              norm['Estatus'] = 'active'
            } else {
              norm['Estatus'] = 'inactive'
            }
          } else {
            norm['Estatus'] = 'inactive'
          }
          
          norm['Fecha de nacimiento'] = formatRowDate(norm['Fecha de nacimiento'], 'YYYY-MM-DD')
          norm['Fecha inicio'] = formatRowDate(norm['Fecha inicio'], 'YYYY-MM-DD')
          norm['Fecha fin'] = formatRowDate(norm['Fecha fin'], 'YYYY-MM-DD')
          norm['Fecha registro'] = formatRowDate(norm['Fecha registro'], 'DD-MM-YYYY')
          
          if (norm['Horario registro']) {
            const hrVal = norm['Horario registro']
            if (hrVal instanceof Date) {
              const h = String(hrVal.getHours()).padStart(2, '0')
              const m = String(hrVal.getMinutes()).padStart(2, '0')
              norm['Horario registro'] = `${h}:${m}`
            } else if (typeof hrVal === 'number' && hrVal < 1) {
              const totalSeconds = Math.round(hrVal * 24 * 3600)
              const h = String(Math.floor(totalSeconds / 3600)).padStart(2, '0')
              const m = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0')
              norm['Horario registro'] = `${h}:${m}`
            } else {
              norm['Horario registro'] = String(hrVal).trim()
            }
          }
          
          if (norm['Fuente']) {
            const fVal = String(norm['Fuente']).trim().toLowerCase()
            const match = FUENTES.find(f => f.toLowerCase() === fVal)
            norm['Fuente'] = match || norm['Fuente']
          }
          
          return norm
        })

        setRawData(processedRows)
        success = true
        break;
      } catch (err) {
        lastError = err;
        console.warn(`Error al intentar cargar desde ${url}:`, err);
      }
    }

    if (!success) {
      setError(lastError ? lastError.message : 'Error desconocido al cargar los datos. Verifica que el archivo sea público y accesible.')
    }
    setLoading(false)
  }

  // Global filters for KPIs and charts
  const filteredData = useMemo(() => {
    return rawData.filter(item => {
      const matchFuente = filterFuente === 'Todas' || item.Fuente === filterFuente
      const matchEstatus = filterEstatus === 'Todos' ||
        (filterEstatus === 'active' && item.Estatus === 'active') ||
        (filterEstatus === 'inactive' && item.Estatus === 'inactive')
      return matchFuente && matchEstatus
    })
  }, [rawData, filterFuente, filterEstatus])

  // Local filtered list for bottom "Socios" table with independent controls
  const filteredSocios = useMemo(() => {
    return rawData.filter(item => {
      const matchFuente = sociosFilterFuente === 'Todas' || item.Fuente === sociosFilterFuente
      const matchEstatus = sociosFilterEstatus === 'Todos' ||
        (sociosFilterEstatus === 'active' && item.Estatus === 'active') ||
        (sociosFilterEstatus === 'inactive' && item.Estatus === 'inactive')
      
      let matchFecha = true
      if (sociosDesde || sociosHasta) {
        const fechaInicio = parseFechaISO(item['Fecha inicio'])
        if (fechaInicio) {
          if (sociosDesde) {
            const desde = parseFechaISO(sociosDesde)
            if (desde && fechaInicio < desde) matchFecha = false
          }
          if (sociosHasta) {
            const hasta = parseFechaISO(sociosHasta)
            if (hasta && fechaInicio > hasta) matchFecha = false
          }
        } else {
          matchFecha = false
        }
      }
      return matchFuente && matchEstatus && matchFecha
    })
  }, [rawData, sociosFilterFuente, sociosFilterEstatus, sociosDesde, sociosHasta])

  const kpis = useMemo(() => {
    const total = filteredData.length
    const activos = filteredData.filter(d => d.Estatus === 'active').length
    const inactivos = filteredData.filter(d => d.Estatus === 'inactive').length
    const tasa = total > 0 ? Math.round((activos / total) * 100) : 0
    return { total, activos, inactivos, tasa }
  }, [filteredData])

  const adquisicionPorFuente = useMemo(() => {
    const counts = {}
    filteredData.forEach(item => {
      const f = item.Fuente || 'Sin fuente'
      counts[f] = (counts[f] || 0) + 1
    })
    const total = filteredData.length
    return Object.entries(counts)
      .map(([fuente, count]) => ({
        fuente,
        count,
        percentage: total > 0 ? Math.round((count / total) * 100) : 0
      }))
      .sort((a, b) => b.count - a.count)
  }, [filteredData])

  const topMembresias = useMemo(() => {
    const counts = {}
    filteredData.forEach(item => {
      const m = item.Membresía || 'Sin membresía'
      counts[m] = (counts[m] || 0) + 1
    })
    return Object.entries(counts)
      .map(([nombre, count]) => ({ nombre, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5)
  }, [filteredData])

  const registrosPeriodo = useMemo(() => {
    let data = rawData
    if (periodoFuente !== 'Todas') {
      data = data.filter(d => d.Fuente === periodoFuente)
    }
    return data.filter(item => {
      const fechaReg = parseFechaRegistro(item['Fecha registro'])
      if (!fechaReg) return false
      switch (periodoMode) {
        case 'dia':
          return isSameDay(fechaReg, periodoFecha)
        case 'semana':
          return isSameWeek(fechaReg, periodoFecha)
        case 'mes':
          return isSameMonth(fechaReg, periodoFecha)
        case 'custom':
          if (!periodoDesde || !periodoHasta) return false
          const desde = parseFechaISO(periodoDesde)
          const hasta = parseFechaISO(periodoHasta)
          if (!desde || !hasta) return false
          return fechaReg >= desde && fechaReg <= hasta
        default:
          return true
      }
    })
  }, [rawData, periodoMode, periodoFecha, periodoDesde, periodoHasta, periodoFuente])

  const conteoPeriodo = registrosPeriodo.length

  const conteoPorFuentePeriodo = useMemo(() => {
    const counts = {}
    registrosPeriodo.forEach(item => {
      const f = item.Fuente || 'Sin fuente'
      counts[f] = (counts[f] || 0) + 1
    })
    return Object.entries(counts)
      .map(([fuente, count]) => ({ fuente, count }))
      .sort((a, b) => b.count - a.count)
  }, [registrosPeriodo])

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-100 px-6 py-4">
        <div className="max-w-7xl mx-auto flex items-center gap-3">
          <div className="w-10 h-10 bg-teal-600 rounded-xl flex items-center justify-center">
            <Users className="w-6 h-6 text-white" />
          </div>
          <h1 className="text-xl font-semibold text-slate-800">Pilates Center</h1>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-6 py-6 space-y-6">
        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-4">
          <div className="flex flex-col md:flex-row gap-4 items-start md:items-center">
            <div className="relative flex-1 w-full">
              <div className="flex items-center gap-2">
                <LinkIcon className="w-4 h-4 text-slate-400" />
                <input
                  type="text"
                  placeholder="Pega aquí la URL de compartir de Google Sheets o Google Drive"
                  value={csvUrl}
                  onChange={(e) => setCsvUrl(e.target.value)}
                  className="flex-1 px-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 focus:border-transparent"
                />
                <button
                  onClick={() => setShowTooltip(!showTooltip)}
                  className="p-2 text-slate-400 hover:text-slate-600"
                >
                  <Info className="w-4 h-4" />
                </button>
              </div>
              {showTooltip && (
                <div className="absolute top-full left-0 mt-2 w-full md:w-96 bg-slate-800 text-white text-xs p-4 rounded-lg z-10">
                  <p className="font-medium mb-2">Cómo compartir el archivo:</p>
                  <ol className="list-decimal list-inside space-y-1 opacity-90">
                    <li>Sube tu archivo (.xlsx o Google Sheets) a Drive</li>
                    <li>Haz clic derecho → Compartir → Compartir</li>
                    <li>Cambia el acceso general a <b>"Cualquier persona con el enlace"</b> como Lector</li>
                    <li>Copia el enlace y pégalo aquí</li>
                  </ol>
                </div>
              )}
            </div>
            <button
              onClick={handleFetch}
              disabled={loading}
              className="flex items-center gap-2 px-4 py-2 bg-teal-600 text-white rounded-lg hover:bg-teal-700 transition-colors disabled:opacity-50"
            >
              {loading ? <Loader2 className="w-4 h-4 animate-spin" /> : <RefreshCw className="w-4 h-4" />}
              Actualizar
            </button>
          </div>
          {error && (
            <div className="mt-3 flex items-center gap-2 text-red-600 text-sm">
              <AlertCircle className="w-4 h-4" />
              {error}
            </div>
          )}
        </div>

        <div className="flex flex-col md:flex-row gap-4">
          <select
            value={filterFuente}
            onChange={(e) => setFilterFuente(e.target.value)}
            className="px-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
          >
            <option value="Todas">Todas las fuentes</option>
            {FUENTES.map(f => (
              <option key={f} value={f}>{f}</option>
            ))}
          </select>
          <select
            value={filterEstatus}
            onChange={(e) => setFilterEstatus(e.target.value)}
            className="px-4 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
          >
            <option value="Todos">Todos los estatus</option>
            <option value="active">Activos</option>
            <option value="inactive">Inactivos</option>
          </select>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-teal-100 rounded-xl flex items-center justify-center">
                <Users className="w-5 h-5 text-teal-600" />
              </div>
              <div>
                <p className="text-sm text-slate-500">Total Socios</p>
                <p className="text-2xl font-semibold text-slate-800">{kpis.total}</p>
              </div>
            </div>
          </div>
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-green-100 rounded-xl flex items-center justify-center">
                <UserCheck className="w-5 h-5 text-green-600" />
              </div>
              <div>
                <p className="text-sm text-slate-500">Socios Activos</p>
                <p className="text-2xl font-semibold text-slate-800">{kpis.activos}</p>
              </div>
            </div>
          </div>
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-red-100 rounded-xl flex items-center justify-center">
                <UserX className="w-5 h-5 text-red-600" />
              </div>
              <div>
                <p className="text-sm text-slate-500">Socios Inactivos</p>
                <p className="text-2xl font-semibold text-slate-800">{kpis.inactivos}</p>
              </div>
            </div>
          </div>
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-purple-100 rounded-xl flex items-center justify-center">
                <TrendingUp className="w-5 h-5 text-purple-600" />
              </div>
              <div>
                <p className="text-sm text-slate-500">Tasa de Retención</p>
                <p className="text-2xl font-semibold text-slate-800">{kpis.tasa}%</p>
              </div>
            </div>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <h3 className="text-sm font-medium text-slate-600 mb-4 flex items-center gap-2">
              <BarChart2 className="w-4 h-4" />
              Adquisición por Fuente
            </h3>
            <div className="space-y-3">
              {adquisicionPorFuente.map(item => (
                <div key={item.fuente}>
                  <div className="flex justify-between text-sm mb-1">
                    <span className="text-slate-700">{item.fuente}</span>
                    <span className="text-slate-500">{item.count} ({item.percentage}%)</span>
                  </div>
                  <div className="w-full bg-slate-100 rounded-full h-2">
                    <div
                      className="bg-teal-500 h-2 rounded-full transition-all"
                      style={{ width: `${item.percentage}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
            <h3 className="text-sm font-medium text-slate-600 mb-4 flex items-center gap-2">
              <Filter className="w-4 h-4" />
              Membresías Populares
            </h3>
            <div className="space-y-3">
              {topMembresias.map((item, idx) => (
                <div key={item.nombre} className="flex items-center gap-3">
                  <span className="w-6 h-6 bg-teal-100 text-teal-700 rounded-full flex items-center justify-center text-xs font-medium">
                    {idx + 1}
                  </span>
                  <span className="flex-1 text-sm text-slate-700 truncate">{item.nombre}</span>
                  <span className="text-sm font-medium text-slate-800">{item.count}</span>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
          <h3 className="text-sm font-medium text-slate-600 mb-4 flex items-center gap-2">
            <CalendarRange className="w-4 h-4" />
            Nuevos Registros por Período
          </h3>

          <div className="flex flex-wrap gap-2 mb-4">
            {[
              { id: 'dia', label: 'Día', icon: Calendar },
              { id: 'semana', label: 'Semana', icon: CalendarRange },
              { id: 'mes', label: 'Mes', icon: Calendar },
              { id: 'custom', label: 'Custom', icon: Clock },
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => setPeriodoMode(tab.id)}
                className={`flex items-center gap-2 rounded-lg text-sm font-medium transition-colors ${
                  periodoMode === tab.id
                    ? 'bg-teal-600 text-white px-4 py-2'
                    : 'bg-slate-100 text-slate-600 hover:bg-slate-200 px-4 py-2'
                }`}
              >
                <tab.icon className="w-4 h-4" />
                {tab.label}
              </button>
            ))}
          </div>

          <div className="flex flex-wrap gap-4 mb-4 items-end">
            {periodoMode === 'custom' ? (
              <>
                <div>
                  <label className="block text-xs text-slate-500 mb-1">Desde</label>
                  <input
                    type="date"
                    value={periodoDesde}
                    onChange={(e) => setPeriodoDesde(e.target.value)}
                    className="px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500"
                  />
                </div>
                <div>
                  <label className="block text-xs text-slate-500 mb-1">Hasta</label>
                  <input
                    type="date"
                    value={periodoHasta}
                    onChange={(e) => setPeriodoHasta(e.target.value)}
                    className="px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500"
                  />
                </div>
              </>
            ) : (
              <div>
                <label className="block text-xs text-slate-500 mb-1">
                  {periodoMode === 'semana' ? 'Semana' : periodoMode === 'mes' ? 'Mes' : 'Fecha'}
                </label>
                <input
                  type="date"
                  value={periodoFecha.toISOString().split('T')[0]}
                  onChange={(e) => setPeriodoFecha(new Date(e.target.value))}
                  className="px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500"
                />
              </div>
            )}
            <div>
              <label className="block text-xs text-slate-500 mb-1">Fuente</label>
              <select
                value={periodoFuente}
                onChange={(e) => setPeriodoFuente(e.target.value)}
                className="px-3 py-2 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
              >
                <option value="Todas">Todas</option>
                {FUENTES.map(f => (
                  <option key={f} value={f}>{f}</option>
                ))}
              </select>
            </div>
          </div>

          <div className="bg-teal-50 rounded-xl p-4 mb-4">
            <p className="text-sm text-slate-600">Total registros en el período</p>
            <p className="text-3xl font-bold text-teal-700">{conteoPeriodo}</p>
          </div>

          {conteoPorFuentePeriodo.length > 0 && (
            <div className="flex flex-wrap gap-2 mb-4">
              {conteoPorFuentePeriodo.map(item => (
                <span
                  key={item.fuente}
                  className={`px-3 py-1 rounded-full text-xs font-medium ${fuenteColors[item.fuente] || 'bg-slate-100 text-slate-700'}`}
                >
                  {item.fuente}: {item.count}
                </span>
              ))}
            </div>
          )}

          {registrosPeriodo.length > 0 ? (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100">
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Nombre</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Membresía</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Fuente</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Fecha Registro</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Horario</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Estatus</th>
                  </tr>
                </thead>
                <tbody>
                  {registrosPeriodo.slice(0, 15).map((item, idx) => (
                    <tr key={idx} className="border-b border-slate-50 hover:bg-slate-50">
                      <td className="py-3 px-2 text-slate-800">{item.Nombre}</td>
                      <td className="py-3 px-2 text-slate-600">{item.Membresía}</td>
                      <td className="py-3 px-2">
                        <span className={`px-2 py-1 rounded-full text-xs font-medium ${fuenteColors[item.Fuente] || 'bg-slate-100 text-slate-700'}`}>
                          {item.Fuente}
                        </span>
                      </td>
                      <td className="py-3 px-2 text-slate-600">{item['Fecha registro']}</td>
                      <td className="py-3 px-2 text-slate-600">{item['Horario registro']}</td>
                      <td className="py-3 px-2">
                        <span className={`px-2 py-1 rounded-full text-xs font-medium ${
                          item.Estatus === 'active' ? 'bg-teal-100 text-teal-700' : 'bg-red-100 text-red-600'
                        }`}>
                          {item.Estatus === 'active' ? 'Activo' : 'Inactivo'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-sm text-slate-500 text-center py-4">No hay registros en el período seleccionado</p>
          )}
        </div>

        <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-5">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4 mb-6">
            <h3 className="text-sm font-medium text-slate-600 flex items-center gap-2">
              <Users className="w-4 h-4" />
              Socios
            </h3>
            
            {/* Filter controls inside the card */}
            <div className="flex flex-wrap items-end gap-3">
              <div>
                <label className="block text-xs text-slate-500 mb-1">Fecha de Inicio (Desde)</label>
                <input
                  type="date"
                  value={sociosDesde}
                  onChange={(e) => setSociosDesde(e.target.value)}
                  className="px-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
                />
              </div>
              <div>
                <label className="block text-xs text-slate-500 mb-1">Fecha de Inicio (Hasta)</label>
                <input
                  type="date"
                  value={sociosHasta}
                  onChange={(e) => setSociosHasta(e.target.value)}
                  className="px-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
                />
              </div>
              <div>
                <label className="block text-xs text-slate-500 mb-1">Fuente</label>
                <select
                  value={sociosFilterFuente}
                  onChange={(e) => setSociosFilterFuente(e.target.value)}
                  className="px-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
                >
                  <option value="Todas">Todas las fuentes</option>
                  {FUENTES.map(f => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs text-slate-500 mb-1">Estatus</label>
                <select
                  value={sociosFilterEstatus}
                  onChange={(e) => setSociosFilterEstatus(e.target.value)}
                  className="px-3 py-1.5 border border-slate-200 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-teal-500 bg-white"
                >
                  <option value="Todos">Todos</option>
                  <option value="active">Activos</option>
                  <option value="inactive">Inactivos</option>
                </select>
              </div>
              {(sociosDesde || sociosHasta || sociosFilterFuente !== 'Todas' || sociosFilterEstatus !== 'Todos') && (
                <button
                  onClick={() => {
                    setSociosDesde('')
                    setSociosHasta('')
                    setSociosFilterFuente('Todas')
                    setSociosFilterEstatus('Todos')
                  }}
                  className="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-lg text-xs font-medium transition-colors"
                >
                  Limpiar
                </button>
              )}
            </div>
          </div>

          {filteredSocios.length > 0 ? (
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-slate-100">
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Nombre</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Membresía</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Fecha Inicio</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Fuente</th>
                    <th className="text-left py-3 px-2 font-medium text-slate-600">Estatus</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSocios.slice(0, 100).map((item, idx) => (
                    <tr key={idx} className="border-b border-slate-50 hover:bg-slate-50">
                      <td className="py-3 px-2 text-slate-800">{item.Nombre}</td>
                      <td className="py-3 px-2 text-slate-600">{item.Membresía}</td>
                      <td className="py-3 px-2 text-slate-600">{item['Fecha inicio']}</td>
                      <td className="py-3 px-2">
                        <span className={`px-2 py-1 rounded-full text-xs font-medium ${fuenteColors[item.Fuente] || 'bg-slate-100 text-slate-700'}`}>
                          {item.Fuente}
                        </span>
                      </td>
                      <td className="py-3 px-2">
                        <span className={`px-2 py-1 rounded-full text-xs font-medium ${
                          item.Estatus === 'active' ? 'bg-teal-100 text-teal-700' : 'bg-red-100 text-red-600'
                        }`}>
                          {item.Estatus === 'active' ? 'Activo' : 'Inactivo'}
                        </span>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ) : (
            <p className="text-sm text-slate-500 text-center py-8">No se encontraron socios que coincidan con los filtros seleccionados</p>
          )}
          
          {filteredSocios.length > 100 ? (
            <p className="text-xs text-slate-500 text-center mt-4 border-t border-slate-50 pt-3">
              Mostrando primeros 100 de {filteredSocios.length} socios filtrados
            </p>
          ) : (
            <p className="text-xs text-slate-500 text-center mt-4 border-t border-slate-50 pt-3">
              Mostrando {filteredSocios.length} socios
            </p>
          )}
        </div>
      </main>
    </div>
  )
}
