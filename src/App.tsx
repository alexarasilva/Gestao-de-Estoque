import React, { useState, useMemo, useEffect } from 'react';
import {
  ResponsiveContainer,
  BarChart,
  Bar,
  XAxis,
  YAxis,
  Tooltip as RechartsTooltip,
  Legend as RechartsLegend,
  Cell,
  PieChart,
  Pie
} from 'recharts';
import {
  LayoutDashboard,
  ClipboardList,
  BaggageClaim,
  BarChart3,
  Search,
  Plus,
  ArrowRightLeft,
  Download,
  AlertTriangle,
  CheckCircle,
  Clock,
  XCircle,
  FileDown,
  ChevronRight,
  List,
  LayoutGrid,
  Filter,
  User,
  MapPin,
  Calendar,
  X,
  History,
  TrendingDown,
  Info,
  Layers,
  ChevronLeft,
  Check,
  Lock,
  Unlock,
  Shield,
  Users,
  Settings,
  Edit,
  Trash2,
  FileSpreadsheet,
  Upload,
  RefreshCw,
  AlertCircle,
  Database,
  ArrowUp,
  ArrowDown,
  ArrowUpDown
} from 'lucide-react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// Custom LogoIcon component reflecting the corporate design identity (stylized Delta/A in a yellow circle)
export function LogoIcon({ className = "w-9 h-9" }: { className?: string }) {
  const [hasError, setHasError] = useState(false);

  if (hasError) {
    return (
      <svg
        viewBox="0 0 100 100"
        className={`${className} select-none shrink-0 rounded-full`}
        fill="none"
        xmlns="http://www.w3.org/2000/svg"
      >
        <circle cx="50" cy="50" r="48" fill="#FFC800" />
        {/* Inner solid upward-pointing triangle (drawn first so black chevron masks it cleanly) */}
        <path
          d="M50 56 L59 75 H41 Z"
          fill="#FFFFFF"
        />
        {/* Stylized upward delta chevron outer contour */}
        <path
          d="M50 19 L73 75 H65 L50 43 L35 75 H27 Z"
          fill="#000000"
        />
      </svg>
    );
  }

  return (
    <img
      src="https://construtoraabsoluta.com.br/wp-content/uploads/2021/04/cropped-favicon-absoluta-192x192.png"
      alt="Logo Absoluta"
      className={`${className} select-none shrink-0 rounded-full object-contain`}
      onError={() => setHasError(true)}
      referrerPolicy="no-referrer"
    />
  );
}

// Corporate Logo Wide (without circular background) to display beautifully at the bottom of the sidebar
export function CorporateLogoWide({ collapsed = false }: { collapsed?: boolean }) {
  const [hasError, setHasError] = useState(false);

  if (collapsed) {
    if (hasError) {
      return (
        <div className="flex justify-center py-4 px-2 w-full border-t border-slate-800/40 mt-4">
          <svg viewBox="20 15 60 70" className="w-8 h-8 select-none shrink-0" fill="none" xmlns="http://www.w3.org/2000/svg">
            {/* Inner solid upward-pointing triangle */}
            <path
              d="M50 56 L59 75 H41 Z"
              fill="#FFFFFF"
            />
            {/* Yellow delta chevron logo */}
            <path
              d="M50 19 L73 75 H65 L50 43 L35 75 H27 Z"
              fill="#FFC800"
            />
          </svg>
        </div>
      );
    }

    return (
      <div className="flex justify-center py-4 px-2 w-full border-t border-slate-800/40 mt-4">
        <img
          src="https://construtoraabsoluta.com.br/wp-content/uploads/2021/04/cropped-favicon-absoluta-192x192.png"
          alt="Logo Absoluta"
          className="w-10 h-10 select-none shrink-0 object-contain rounded-full"
          onError={() => setHasError(true)}
          referrerPolicy="no-referrer"
        />
      </div>
    );
  }

  if (hasError) {
    return (
      <div className="px-4 py-4 mt-4 border-t border-slate-800/40 flex flex-col items-center justify-center w-full">
        <div className="w-full flex items-center justify-center gap-3 hover:opacity-95 transition-all">
          <svg viewBox="20 15 60 70" className="w-12 h-12 select-none shrink-0" fill="none" xmlns="http://www.w3.org/2000/svg">
            {/* Inner solid upward-pointing triangle */}
            <path
              d="M50 56 L59 75 H41 Z"
              fill="#FFFFFF"
            />
            {/* Yellow delta chevron logo */}
            <path
              d="M50 19 L73 75 H65 L50 43 L35 75 H27 Z"
              fill="#FFC800"
            />
          </svg>
          <div className="flex flex-col">
            <span className="text-sm font-sans font-black tracking-widest text-white leading-none">ABSOLUTA</span>
            <span className="text-[9px] font-mono font-extrabold tracking-[0.25em] text-[#FFC800] mt-1 leading-none">CONSTRUTORA</span>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="px-4 py-4 mt-4 border-t border-slate-800/40 flex items-center justify-center w-full">
      <div className="w-full flex items-center justify-center hover:opacity-95 transition-all px-2">
        <img
          src="https://construtoraabsoluta.com.br/wp-content/uploads/2021/04/logo-absoluta-construtora-topo.png"
          alt="Absoluta Construtora Logo"
          className="w-48 h-auto select-none shrink-0 object-contain"
          onError={() => setHasError(true)}
          referrerPolicy="no-referrer"
        />
      </div>
    </div>
  );
}


// Interfaces
export interface Obra {
  id: string;
  nome: string;
}

export interface Pedido {
  id: string;
  insumo: string;
  codigo: string;
  obra: string;
  obraId?: string;
  qtdSolicitada: number;
  qtdRecebida: number;
  unidade: string;
  status: 'Pendente' | 'Parcial' | 'Entregue' | 'Cancelado';
  dataPedido: string;
  dataChegada?: string;
  
  // Sienge excel fields
  codComprador?: string;
  fornecedor?: string;
  codDetalhe?: string;
  descricaoDetalhe?: string;
  marca?: string;
  descricaoUnidade?: string;
  qtdPendenteImportada?: number;
}

export function calculateDaysElapsed(dataPedidoStr: any, statusText: string, dataChegadaStr?: any): number {
  if (!dataPedidoStr) return 0;
  
  const parseDate = (val: any): Date | null => {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
      const d = new Date((val - 25569) * 86400 * 1000);
      return isNaN(d.getTime()) ? null : d;
    }
    const str = String(val).trim();
    const parts = str.split('/');
    if (parts.length === 3) {
      const day = parseInt(parts[0], 10);
      const month = parseInt(parts[1], 10) - 1; // 0-indexed
      const year = parseInt(parts[2], 10);
      if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
        return new Date(year, month, day);
      }
    }
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
  };

  const start = parseDate(dataPedidoStr);
  if (!start) return 1;

  let end: Date;
  if (statusText === 'Entregue') {
    end = (dataChegadaStr ? parseDate(dataChegadaStr) : null) || new Date();
  } else {
    end = new Date(); // Currently in progress
  }

  start.setHours(0, 0, 0, 0);
  end.setHours(0, 0, 0, 0);

  const diffTime = Math.max(0, end.getTime() - start.getTime());
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays === 0 ? 1 : diffDays;
}

export interface Baixa {
  id: string;
  insumo: string;
  codigo: string;
  obra: string;
  obraId?: string;
  quantidade: number;
  unidade: string;
  colaborador: string;
  data: string;
  destino?: string;
  autorizado?: string;
}

export interface Usuario {
  id: string;
  nome: string;
  login: string;
  senha: string;
  role: 'Administrador' | 'Colaborador';
  podeCriarPedido: boolean;
  podeDarBaixa: boolean;
  podeReceberMercadoria: boolean;
  podeExcluirPedido: boolean;
  podeExcluirObra: boolean;
}

function parseDateTimeString(str: string | undefined): Date | null {
  if (!str) return null;
  const cleaned = str.trim();
  const dateTimeParts = cleaned.split(' ');
  const dateStr = dateTimeParts[0];
  const timeStr = dateTimeParts[1] || '00:00';

  const dateParts = dateStr.split('/');
  if (dateParts.length === 3) {
    const day = parseInt(dateParts[0], 10);
    const month = parseInt(dateParts[1], 10) - 1;
    const year = parseInt(dateParts[2], 10);
    const timeParts = timeStr.split(':');
    const hour = parseInt(timeParts[0] || '0', 10);
    const minute = parseInt(timeParts[1] || '0', 10);
    const second = parseInt(timeParts[2] || '0', 10);
    const d = new Date(year, month, day, hour, minute, second);
    return isNaN(d.getTime()) ? null : d;
  }

  const d = new Date(cleaned);
  return isNaN(d.getTime()) ? null : d;
}

function formatDateTimeToISO(date: Date | null): string {
  if (!date) return '-';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${day}/${month}/${year} ${hours}:${minutes}`;
}

export default function App() {
  // Current Navigation View Tab
  const [currentTab, setCurrentTab] = useState<'painel' | 'pedidos' | 'estoque' | 'relatorios' | 'admin'>('painel');

  // Reconciliation states
  const [showReconcileModal, setShowReconcileModal] = useState<boolean>(false);
  const [reconcileItem, setReconcileItem] = useState<any>(null);
  const [reconcileNewPhysicalQty, setReconcileNewPhysicalQty] = useState<number | ''>('');
  const [reconcileJustificativa, setReconcileJustificativa] = useState<string>('');
  const [reconcileResponsavel, setReconcileResponsavel] = useState<string>('');

  // Currently Selected Construction Project (Obra Filter)
  const [selectedObra, setSelectedObra] = useState<string>('Todas as Obras');

  // Search filter query
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Table Status Filter (Supports selecting multiple status values)
  const [statusFilter, setStatusFilter] = useState<string[]>(['Pendente']);

  const handleStatusFilterToggle = (st: string) => {
    if (st === 'Todos') {
      setStatusFilter(['Todos']);
    } else {
      let nextFilters = statusFilter.filter(item => item !== 'Todos');
      if (nextFilters.includes(st)) {
        nextFilters = nextFilters.filter(item => item !== st);
      } else {
        nextFilters.push(st);
      }
      if (nextFilters.length === 0) {
        nextFilters = ['Todos'];
      }
      setStatusFilter(nextFilters);
    }
  };

  // General Transaction Count (Movimentações Hoje)
  const [movementsCount, setMovementsCount] = useState<number>(142);

  // Last Sienge Import Date and Time
  const [lastImportDate, setLastImportDate] = useState<string>(() => {
    return localStorage.getItem('last_import_date') || '19/06/2026 às 08:30';
  });

  // Dynamic list of active Construction sites (Obras)
  // Google Sheets Integration URL and Settings persisted via LocalStorage
  const [googleSheetsUrl, setGoogleSheetsUrl] = useState<string>(() => {
    return localStorage.getItem('construmais_gs_url') || '';
  });
  const [autoSync, setAutoSync] = useState<boolean>(() => {
    return localStorage.getItem('construmais_auto_sync') === 'true';
  });
  const [isSyncing, setIsSyncing] = useState<boolean>(false);

  // Dynamic list of active Construction sites (Obras)
  const [obras, setObras] = useState<Obra[]>(() => {
    const saved = localStorage.getItem('construmais_obras_v3');
    if (saved) return JSON.parse(saved);
    
    // Migração de dados legados do localStorage antigo de string[] para o novo Obra[]
    const legacySaved = localStorage.getItem('construmais_obras');
    if (legacySaved) {
      try {
        const parsed = JSON.parse(legacySaved);
        if (Array.isArray(parsed)) {
          if (typeof parsed[0] === 'string') {
            return parsed.map((o, idx) => ({ id: `OB-${101 + idx}`, nome: o }));
          } else if (parsed[0] && parsed[0].id) {
            return parsed;
          }
        }
      } catch (e) {}
    }
    return [
      { id: 'OB-101', nome: 'Residencial Alvorada' },
      { id: 'OB-102', nome: 'Torre Infinito' },
      { id: 'OB-103', nome: 'Complexo Hospitalar' }
    ];
  });

  // Backward compatibility bridge
  const listObras = obras.map(o => o.nome);

  // Active Users database and custom permissions state
  const [usuarios, setUsuarios] = useState<Usuario[]>(() => {
    const saved = localStorage.getItem('construmais_usuarios');
    return saved ? JSON.parse(saved) : [
      {
        id: 'usr-1',
        nome: 'Carlos Roberto (Admin)',
        login: 'carlos@empresa.com.br',
        senha: 'c123',
        role: 'Administrador',
        podeCriarPedido: true,
        podeDarBaixa: true,
        podeReceberMercadoria: true,
        podeExcluirPedido: true,
        podeExcluirObra: true
      },
      {
        id: 'usr-2',
        nome: 'Mestre José',
        login: 'jose@empresa.com.br',
        senha: 'j123',
        role: 'Colaborador',
        podeCriarPedido: true,
        podeDarBaixa: true,
        podeReceberMercadoria: false,
        podeExcluirPedido: false,
        podeExcluirObra: false
      },
      {
        id: 'usr-3',
        nome: 'Cleiton Silva (Estagiário)',
        login: 'cleiton@empresa.com.br',
        senha: 'cl123',
        role: 'Colaborador',
        podeCriarPedido: false,
        podeDarBaixa: true,
        podeReceberMercadoria: false,
        podeExcluirPedido: false,
        podeExcluirObra: false
      }
    ];
  });

  // Authenticated/Operating active profile context switcher
  const [currentUser, setCurrentUser] = useState<Usuario>(() => {
    const savedUser = localStorage.getItem('construmais_current_user');
    if (savedUser) return JSON.parse(savedUser);
    return {
      id: 'usr-1',
      nome: 'Carlos Roberto (Admin)',
      login: 'carlos@empresa.com.br',
      senha: 'c123',
      role: 'Administrador',
      podeCriarPedido: true,
      podeDarBaixa: true,
      podeReceberMercadoria: true,
      podeExcluirPedido: true,
      podeExcluirObra: true
    };
  });

  // Initial State: Purchase Orders (Pedidos)
  const [pedidos, setPedidos] = useState<Pedido[]>(() => {
    const saved = localStorage.getItem('construmais_pedidos');
    return saved ? JSON.parse(saved) : [
      { id: 'PC-2024-0891', insumo: 'Cimento CP-II 50kg', codigo: 'CIM-001', obra: 'Residencial Alvorada', qtdSolicitada: 500, qtdRecebida: 350, unidade: 'un', status: 'Parcial', dataPedido: '14/05/2026' },
      { id: 'PC-2024-0902', insumo: 'Vergalhão CA-50 10mm', codigo: 'VER-002', obra: 'Torre Infinito', qtdSolicitada: 1200, qtdRecebida: 0, unidade: 'kg', status: 'Pendente', dataPedido: '13/05/2026' },
      { id: 'PC-2024-0844', insumo: 'Areia Lavada Média', codigo: 'ARE-003', obra: 'Complexo Hospitalar', qtdSolicitada: 40, qtdRecebida: 40, unidade: 'm³', status: 'Entregue', dataPedido: '10/05/2026' },
      { id: 'PC-2024-0915', insumo: 'Piso Porcelanato 80x80', codigo: 'PIS-004', obra: 'Residencial Alvorada', qtdSolicitada: 340, qtdRecebida: 0, unidade: 'm²', status: 'Cancelado', dataPedido: '12/05/2026' },
      { id: 'PC-2024-0922', insumo: 'Tinta Acrílica Premium 18L', codigo: 'TIN-005', obra: 'Torre Infinito', qtdSolicitada: 80, qtdRecebida: 80, unidade: 'un', status: 'Entregue', dataPedido: '11/05/2026' },
      { id: 'PC-2024-0935', insumo: 'Cabo Flexível 2.5mm²', codigo: 'CAB-006', obra: 'Residencial Alvorada', qtdSolicitada: 1000, qtdRecebida: 400, unidade: 'm', status: 'Parcial', dataPedido: '14/05/2026' },
      { id: 'PC-2024-0941', insumo: 'Tubo de PVC Esgoto 100mm', codigo: 'TUB-007', obra: 'Complexo Hospitalar', qtdSolicitada: 120, qtdRecebida: 0, unidade: 'un', status: 'Pendente', dataPedido: '15/05/2026' },
      { id: 'PC-2024-0948', insumo: 'Aço CA-60 5.0mm', codigo: 'ACO-008', obra: 'Torre Infinito', qtdSolicitada: 500, qtdRecebida: 500, unidade: 'kg', status: 'Entregue', dataPedido: '12/05/2026' },
    ];
  });

  // Initial State: Material Withdrawals (Baixas de Material)
  const [baixas, setBaixas] = useState<Baixa[]>(() => {
    const saved = localStorage.getItem('construmais_baixas');
    return saved ? JSON.parse(saved) : [
      { id: 'BX-2024-0012', insumo: 'Cimento CP-II 50kg', codigo: 'CIM-001', obra: 'Residencial Alvorada', quantidade: 120, unidade: 'un', colaborador: 'Mestre José', data: '16/06/2026 09:15' },
      { id: 'BX-2024-0013', insumo: 'Areia Lavada Média', codigo: 'ARE-003', obra: 'Complexo Hospitalar', quantidade: 15, unidade: 'm³', colaborador: 'Carlos Roberto', data: '16/06/2026 10:30' },
      { id: 'BX-2024-0014', insumo: 'Tinta Acrílica Premium 18L', codigo: 'TIN-005', obra: 'Torre Infinito', quantidade: 30, unidade: 'un', colaborador: 'Pintor Antenor', data: '16/06/2026 11:00' },
      { id: 'BX-2024-0015', insumo: 'Cimento CP-II 50kg', codigo: 'CIM-001', obra: 'Residencial Alvorada', quantidade: 80, unidade: 'un', colaborador: 'Pedreiro Cleiton', data: '16/06/2026 13:45' }
    ];
  });

  // LocalStorage Persistence effect triggers on any change
  useEffect(() => {
    localStorage.setItem('construmais_obras_v3', JSON.stringify(obras));
    localStorage.setItem('construmais_obras', JSON.stringify(obras.map(o => o.nome)));
  }, [obras]);

  useEffect(() => {
    localStorage.setItem('construmais_usuarios', JSON.stringify(usuarios));
  }, [usuarios]);

  useEffect(() => {
    localStorage.setItem('construmais_current_user', JSON.stringify(currentUser));
  }, [currentUser]);

  useEffect(() => {
    localStorage.setItem('construmais_pedidos', JSON.stringify(pedidos));
  }, [pedidos]);

  useEffect(() => {
    localStorage.setItem('construmais_baixas', JSON.stringify(baixas));
  }, [baixas]);

  useEffect(() => {
    localStorage.setItem('construmais_gs_url', googleSheetsUrl);
  }, [googleSheetsUrl]);

  useEffect(() => {
    localStorage.setItem('construmais_auto_sync', String(autoSync));
  }, [autoSync]);

  // Debounced auto-sync when local state changes and autoSync is enabled
  useEffect(() => {
    if (autoSync && googleSheetsUrl) {
      const timer = setTimeout(() => {
        handleSheetsSync('push', true);
      }, 3500); // 3.5 Segundos de debounce para agrupar clicks rápidos
      return () => clearTimeout(timer);
    }
  }, [pedidos, baixas, obras, usuarios, autoSync, googleSheetsUrl]);

  // Modals Visibility
  const [showOrderModal, setShowOrderModal] = useState<boolean>(false);
  const [showWithdrawModal, setShowWithdrawModal] = useState<boolean>(false);
  const [showWithdrawSuccessScreen, setShowWithdrawSuccessScreen] = useState<boolean>(false);
  const [showReconcileSuccessScreen, setShowReconcileSuccessScreen] = useState<boolean>(false);
  const [showReceiveModal, setShowReceiveModal] = useState<boolean>(false);

  // Administrative Modals & States
  const [editingPedido, setEditingPedido] = useState<Pedido | null>(null);
  const [showEditPedidoModal, setShowEditPedidoModal] = useState<boolean>(false);
  const [lockModalReason, setLockModalReason] = useState<string | null>(null);

  // User management details modal edit contexts
  const [editingUsuario, setEditingUsuario] = useState<Usuario | null>(null);
  const [showUserEditModal, setShowUserEditModal] = useState<boolean>(false);
  const [newUserOpen, setNewUserOpen] = useState<boolean>(false);

  // User switcher modal
  const [showUserSwitcherModal, setShowUserSwitcherModal] = useState<boolean>(false);

  // Obra management variables
  const [showAddObraModal, setShowAddObraModal] = useState<boolean>(false);
  const [newObraName, setNewObraName] = useState<string>('');
  const [editingObraName, setEditingObraName] = useState<string | null>(null);

  // Modal State Variables
  const [selectedPedidoId, setSelectedPedidoId] = useState<string>('');
  const [selectedPedidoCodigo, setSelectedPedidoCodigo] = useState<string>('');
  const [qtdReceiveInput, setQtdReceiveInput] = useState<number | ''>('');
  const [receiptNotaFiscal, setReceiptNotaFiscal] = useState<string>('');
  const [receiptResponsavel, setReceiptResponsavel] = useState<string>('');
  const [receiptStatusAlteracao, setReceiptStatusAlteracao] = useState<string>('deduzir_auto');

  // Sienge excel/spreadsheet import state
  const [pedidosViewMode, setPedidosViewMode] = useState<'cards' | 'table'>('table');
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState<boolean>(false);
  const [sortConfig, setSortConfig] = useState<{
    key: keyof Pedido | 'progress' | 'id_numeric' | '';
    direction: 'asc' | 'desc' | '';
  }>({
    key: '',
    direction: '',
  });
  const [inventorySortConfig, setInventorySortConfig] = useState<{
    key: 'obra' | 'insumo' | 'unidade' | 'recebido' | 'baixado' | 'saldo' | 'ultimaMovimentacao' | '';
    direction: 'asc' | 'desc' | '';
  }>({
    key: 'ultimaMovimentacao',
    direction: 'desc',
  });
  const [showClearConfirmModal, setShowClearConfirmModal] = useState<boolean>(false);
  const [showImportModal, setShowImportModal] = useState<boolean>(false);
  const [parsedImportItems, setParsedImportItems] = useState<Pedido[]>([]);
  const [importedFileName, setImportedFileName] = useState<string>('');
  const [importStats, setImportStats] = useState<{ newOrders: number; newItems: number; updatedQtds: number } | null>(null);
  
  // New Order Form States
  const [newOrderInsumo, setNewOrderInsumo] = useState<string>('');
  const [newOrderCodigo, setNewOrderCodigo] = useState<string>('');
  const [newOrderObra, setNewOrderObra] = useState<string>('Residencial Alvorada');
  const [newOrderQtd, setNewOrderQtd] = useState<number>(100);
  const [newOrderUnidade, setNewOrderUnidade] = useState<string>('un');

  // New Withdrawal Form States
  const [newWithdrawObra, setNewWithdrawObra] = useState<string>('Residencial Alvorada');
  const [newWithdrawInsumo, setNewWithdrawInsumo] = useState<string>('');
  const [newWithdrawQtd, setNewWithdrawQtd] = useState<number>(10);
  const [newWithdrawColab, setNewWithdrawColab] = useState<string>('');
  const [newWithdrawDestino, setNewWithdrawDestino] = useState<string>('');
  const [newWithdrawAutorizado, setNewWithdrawAutorizado] = useState<string>('');

  // Report tab filters dynamic local state variables
  const [reportStartDate, setReportStartDate] = useState<string>('');
  const [reportEndDate, setReportEndDate] = useState<string>('');
  const [reportSelectedObra, setReportSelectedObra] = useState<string>('Todas as Obras');
  const [reportSelectedUser, setReportSelectedUser] = useState<string>('Todos os Usuários');
  const [reportSelectedType, setReportSelectedType] = useState<string>('Ambos');
  const [reportSelectedMaterial, setReportSelectedMaterial] = useState<string>('Todos os Materiais');
  const [reportSelectedSupplier, setReportSelectedSupplier] = useState<string>('Todos os Fornecedores');

  // Toast System
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'warn' | 'info' | null }>({ message: '', type: null });

  const triggerToast = (message: string, type: 'success' | 'warn' | 'info') => {
    setToast({ message, type });
    setTimeout(() => {
      setToast({ message: '', type: null });
    }, 2250);
  };

  const handleReconcileWeight = () => {
    if (!reconcileItem) return;
    const newQty = Number(reconcileNewPhysicalQty);
    if (isNaN(newQty) || newQty < 0) {
      triggerToast('Por favor, informe uma quantidade física válida.', 'warn');
      return;
    }

    if (!reconcileJustificativa.trim()) {
      triggerToast('O preenchimento do campo "Justificativa da Divergência" é obrigatório.', 'warn');
      return;
    }

    if (!reconcileResponsavel.trim()) {
      triggerToast('O preenchimento do campo "Responsável Técnico / Almoxarife" é obrigatório.', 'warn');
      return;
    }

    const difference = reconcileItem.saldo - newQty;
    if (difference === 0) {
      triggerToast('A quantidade informada é idêntica ao saldo atual.', 'info');
      setShowReconcileModal(false);
      return;
    }

    const newBaixa: Baixa = {
      id: `RE-2026-${Math.floor(1000 + Math.random() * 9000)}`,
      insumo: reconcileItem.insumo,
      codigo: reconcileItem.codigo,
      obra: reconcileItem.obra,
      quantidade: difference,
      unidade: reconcileItem.unidade,
      colaborador: reconcileResponsavel.trim(),
      data: new Date().toLocaleString('pt-BR', { hour12: false }).replace(',', ''),
      destino: `Ajuste: ${reconcileJustificativa.trim()}`,
      autorizado: reconcileResponsavel.trim()
    };

    setBaixas(prev => {
      const next = [newBaixa, ...prev];
      localStorage.setItem('construmais_baixas', JSON.stringify(next));
      return next;
    });

    setShowReconcileModal(false);
    setShowReconcileSuccessScreen(true);
    
    // Clear states
    const oldItem = reconcileItem;
    const oldQty = newQty;
    
    setReconcileNewPhysicalQty('');
    setReconcileJustificativa('');
    setReconcileResponsavel('');

    setTimeout(() => {
      setShowReconcileSuccessScreen(false);
      setReconcileItem(null);
      setCurrentTab('estoque');
      triggerToast(`Estoque reconciliado de ${oldItem.saldo} para ${oldQty} ${oldItem.unidade}!`, 'success');
    }, 1800);
  };

  const handleSortToggle = (key: keyof Pedido | 'progress' | 'id_numeric') => {
    let newDirection: 'asc' | 'desc' | '' = 'asc';
    if (sortConfig.key === key) {
      if (sortConfig.direction === 'asc') {
        newDirection = 'desc';
      } else {
        newDirection = '';
      }
    }
    setSortConfig({ key: newDirection ? key : '', direction: newDirection });
    
    if (newDirection === '') {
      triggerToast('Ordenação desativada', 'info');
    } else {
      const orderLabel = newDirection === 'asc' ? 'A-Z / Crescente' : 'Z-A / Decrescente';
      const keyLabels: Record<string, string> = {
        id_numeric: 'Código',
        obra: 'Obra / C.C',
        fornecedor: 'Fornecedor / Comprador',
        insumo: 'Insumo / Descrição',
        progress: 'Pedido / Recebido',
        status: 'Status Logístico'
      };
      const label = keyLabels[key] || String(key);
      triggerToast(`Ordenando por ${label}: ${orderLabel}`, 'success');
    }
  };

  const handleInventorySortToggle = (key: 'obra' | 'insumo' | 'unidade' | 'recebido' | 'baixado' | 'saldo' | 'ultimaMovimentacao') => {
    let newDirection: 'asc' | 'desc' | '' = 'asc';
    if (inventorySortConfig.key === key) {
      if (inventorySortConfig.direction === 'asc') {
        newDirection = 'desc';
      } else {
        newDirection = '';
      }
    }
    setInventorySortConfig({ key: newDirection ? key : '', direction: newDirection });
    
    if (newDirection === '') {
      triggerToast('Ordenação do estoque desativada', 'info');
    } else {
      const orderLabel = newDirection === 'asc' ? 'Crescente' : 'Decrescente';
      const keyLabels: Record<string, string> = {
        obra: 'Obra / Centro de Custo',
        insumo: 'Material / Insumo',
        unidade: 'Unidade',
        recebido: 'Entradas Acumuladas',
        baixado: 'Baixas Registradas',
        saldo: 'Saldo Físico Atual',
        ultimaMovimentacao: 'Última Movimentação'
      };
      const label = keyLabels[key] || String(key);
      triggerToast(`Ordenando estoque por ${label}: ${orderLabel}`, 'success');
    }
  };


  // Calculate Net Available Stock for each (Insumo + Obra)
  // Formula: Available = Sum(qtdRecebida from Pedidos with same insumo & obra) - Sum(quantidade from Baixas with same insumo & obra)
  const stockInventory = useMemo(() => {
    const map: { 
      [key: string]: { 
        insumo: string; 
        codigo: string; 
        obra: string; 
        recebido: number; 
        baixado: number; 
        saldo: number; 
        unidade: string;
        obraId?: string;
        fornecedor?: string;
        marca?: string;
        codDetalhe?: string;
        descricaoDetalhe?: string;
        descricaoUnidade?: string;
        ultimaMovimentacao?: string;
        rawUltimaMovimentacao?: number;
        dataUltimaEntrada?: string;
        dataUltimaBaixa?: string;
      } 
    } = {};

    // 1. Gather all received amounts
    pedidos.forEach((p) => {
      if (p.status === 'Cancelado') return;
      const key = `${p.obra}:::${p.codigo || ''}:::${p.insumo}`;
      if (!map[key]) {
        map[key] = { 
          insumo: p.insumo, 
          codigo: p.codigo || '', 
          obra: p.obra, 
          recebido: 0, 
          baixado: 0, 
          saldo: 0, 
          unidade: p.unidade,
          obraId: p.obraId,
          fornecedor: p.fornecedor,
          marca: p.marca,
          codDetalhe: p.codDetalhe,
          descricaoDetalhe: p.descricaoDetalhe,
          descricaoUnidade: p.descricaoUnidade,
          ultimaMovimentacao: '-'
        };
      }
      map[key].recebido += p.qtdRecebida;
      map[key].saldo += p.qtdRecebida;
      
      // Keep last imported / updated data
      if (p.fornecedor) map[key].fornecedor = p.fornecedor;
      if (p.marca) map[key].marca = p.marca;
      if (p.codDetalhe) map[key].codDetalhe = p.codDetalhe;
      if (p.descricaoDetalhe) map[key].descricaoDetalhe = p.descricaoDetalhe;
      if (p.descricaoUnidade) map[key].descricaoUnidade = p.descricaoUnidade;
      if (p.obraId) map[key].obraId = p.obraId;
    });

    // 2. Subtract all withdrawn amounts
    baixas.forEach((b) => {
      const key = `${b.obra}:::${b.codigo || ''}:::${b.insumo}`;
      if (!map[key]) {
        map[key] = { 
          insumo: b.insumo, 
          codigo: b.codigo || '', 
          obra: b.obra, 
          recebido: 0, 
          baixado: 0, 
          saldo: 0, 
          unidade: b.unidade,
          obraId: b.obraId,
          ultimaMovimentacao: '-'
        };
      }
      map[key].baixado += b.quantidade;
      map[key].saldo -= b.quantidade;
    });

    // 3. Find latest movement date-time for each item
    Object.keys(map).forEach((key) => {
      const item = map[key];
      let latestDate: Date | null = null;
      let latestEntradaDate: Date | null = null;
      let latestBaixaDate: Date | null = null;

      // Check matching deliveries in pedidos
      pedidos.forEach((p) => {
        if (p.status === 'Cancelado') return;
        if (p.qtdRecebida <= 0) return;
        const matches = p.obra === item.obra && (p.codigo || '') === item.codigo && p.insumo === item.insumo;
        if (matches) {
          const d = parseDateTimeString(p.dataChegada || p.dataPedido);
          if (d && (!latestDate || d > latestDate)) {
            latestDate = d;
          }
          if (d && (!latestEntradaDate || d > latestEntradaDate)) {
            latestEntradaDate = d;
          }
        }
      });

      // Check matching baixas
      baixas.forEach((b) => {
        const matches = b.obra === item.obra && (b.codigo || '') === item.codigo && b.insumo === item.insumo;
        if (matches) {
          const d = parseDateTimeString(b.data);
          if (d && (!latestDate || d > latestDate)) {
            latestDate = d;
          }
          if (d && (!latestBaixaDate || d > latestBaixaDate)) {
            latestBaixaDate = d;
          }
        }
      });

      // Default fallback if no date found yet
      if (!latestDate) {
        pedidos.forEach((p) => {
          const matches = p.obra === item.obra && (p.codigo || '') === item.codigo && p.insumo === item.insumo;
          if (matches) {
            const d = parseDateTimeString(p.dataPedido);
            if (d && (!latestDate || d > latestDate)) {
              latestDate = d;
            }
          }
        });
      }

      // Default fallback if no entry date found yet
      if (!latestEntradaDate) {
        pedidos.forEach((p) => {
          const matches = p.obra === item.obra && (p.codigo || '') === item.codigo && p.insumo === item.insumo;
          if (matches) {
            const d = parseDateTimeString(p.dataPedido);
            if (d && (!latestEntradaDate || d > latestEntradaDate)) {
              latestEntradaDate = d;
            }
          }
        });
      }

      item.ultimaMovimentacao = latestDate ? formatDateTimeToISO(latestDate) : '-';
      item.rawUltimaMovimentacao = latestDate ? latestDate.getTime() : 0;
      
      item.dataUltimaEntrada = latestEntradaDate ? formatDateTimeToISO(latestEntradaDate) : '-';
      item.dataUltimaBaixa = latestBaixaDate ? formatDateTimeToISO(latestBaixaDate) : '-';
    });

    return Object.values(map);
  }, [pedidos, baixas]);

  // Derived filtered inventories / orders lists
  const filteredPedidos = useMemo(() => {
    const list = pedidos.filter((p) => {
      const pObra = p.obra || '';
      const matchObra = selectedObra === 'Todas as Obras' || pObra === selectedObra;
      
      const q = searchQuery.toLowerCase().trim();
      const matchSearch =
        (p.insumo || '').toLowerCase().includes(q) ||
        (p.codigo || '').toLowerCase().includes(q) ||
        (p.id || '').toLowerCase().includes(q) ||
        (pObra).toLowerCase().includes(q) ||
        (p.obraId || '').toLowerCase().includes(q) ||
        (p.fornecedor || '').toLowerCase().includes(q) ||
        (p.codComprador || '').toLowerCase().includes(q) ||
        (p.codDetalhe || '').toLowerCase().includes(q) ||
        (p.descricaoDetalhe || '').toLowerCase().includes(q) ||
        (p.marca || '').toLowerCase().includes(q);
        
      const matchStatus = statusFilter.includes('Todos') || statusFilter.includes(p.status);
      return matchObra && matchSearch && matchStatus;
    });

    // Default sorting: Newest first (by date dataPedido or parsed date)
    const parseDateHelper = (val: any): number => {
      if (!val) return 0;
      const str = String(val).trim();
      const parts = str.split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0], 10);
        const month = parseInt(parts[1], 10) - 1; // 0-indexed
        const year = parseInt(parts[2], 10);
        if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
          return new Date(year, month, day).getTime();
        }
      }
      const d = new Date(str);
      return isNaN(d.getTime()) ? 0 : d.getTime();
    };

    list.sort((a, b) => {
      const timeA = parseDateHelper(a.dataPedido);
      const timeB = parseDateHelper(b.dataPedido);
      if (timeA !== timeB) {
        return timeB - timeA; // Descending (most recent dates first)
      }
      // If dates match, sort by numeric part of the ID descending
      const numA = parseInt(a.id.replace(/[^\d]/g, ''), 10) || 0;
      const numB = parseInt(b.id.replace(/[^\d]/g, ''), 10) || 0;
      return numB - numA;
    });

    if (sortConfig.key && sortConfig.direction) {
      list.sort((a, b) => {
        let valA: any = '';
        let valB: any = '';

        if (sortConfig.key === 'progress') {
          // Progress ratio
          const progressA = a.qtdSolicitada > 0 ? (a.qtdRecebida / a.qtdSolicitada) : 0;
          const progressB = b.qtdSolicitada > 0 ? (b.qtdRecebida / b.qtdSolicitada) : 0;
          valA = progressA;
          valB = progressB;
        } else if (sortConfig.key === 'id_numeric') {
          // Try numerical comparison on code or ID
          const numA = parseInt(a.id.replace(/[^\d]/g, ''), 10);
          const numB = parseInt(b.id.replace(/[^\d]/g, ''), 10);
          valA = isNaN(numA) ? a.id : numA;
          valB = isNaN(numB) ? b.id : numB;
        } else {
          valA = a[sortConfig.key as keyof Pedido];
          valB = b[sortConfig.key as keyof Pedido];
        }

        if (valA === undefined || valA === null) valA = '';
        if (valB === undefined || valB === null) valB = '';

        if (typeof valA === 'string' && typeof valB === 'string') {
          return sortConfig.direction === 'asc'
            ? valA.localeCompare(valB, 'pt-BR', { sensitivity: 'base', numeric: true })
            : valB.localeCompare(valA, 'pt-BR', { sensitivity: 'base', numeric: true });
        }

        if (valA < valB) return sortConfig.direction === 'asc' ? -1 : 1;
        if (valA > valB) return sortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }

    return list;
  }, [pedidos, selectedObra, searchQuery, statusFilter, sortConfig]);

  const filteredStock = useMemo(() => {
    let list = stockInventory.filter((item) => {
      const itemObra = item.obra || '';
      const matchObra = selectedObra === 'Todas as Obras' || itemObra === selectedObra;
      
      const q = searchQuery.toLowerCase().trim();
      const matchSearch =
        (item.insumo || '').toLowerCase().includes(q) ||
        (item.codigo || '').toLowerCase().includes(q) ||
        (itemObra).toLowerCase().includes(q) ||
        (item.fornecedor || '').toLowerCase().includes(q) ||
        (item.marca || '').toLowerCase().includes(q) ||
        (item.descricaoDetalhe || '').toLowerCase().includes(q);
        
      return matchObra && matchSearch;
    });

    if (inventorySortConfig.key && inventorySortConfig.direction) {
      list = [...list].sort((a, b) => {
        let valA: any;
        let valB: any;

        const k = inventorySortConfig.key;
        if (k === 'recebido' || k === 'baixado' || k === 'saldo') {
          valA = Number(a[k]) || 0;
          valB = Number(b[k]) || 0;
        } else if (k === 'ultimaMovimentacao') {
          valA = a.rawUltimaMovimentacao || 0;
          valB = b.rawUltimaMovimentacao || 0;
        } else {
          valA = String(a[k as keyof typeof a] || '').toLowerCase();
          valB = String(b[k as keyof typeof b] || '').toLowerCase();
        }

        if (valA < valB) return inventorySortConfig.direction === 'asc' ? -1 : 1;
        if (valA > valB) return inventorySortConfig.direction === 'asc' ? 1 : -1;
        return 0;
      });
    }

    return list;
  }, [stockInventory, selectedObra, searchQuery, inventorySortConfig]);

  // Calculated Stats Indicators based on "selectedObra"
  const stats = useMemo(() => {
    const ordersSubset = pedidos.filter((p) => selectedObra === 'Todas as Obras' || p.obra === selectedObra);
    
    // 1. Pending (Pendente) orders total
    const pendentes = ordersSubset.filter((p) => p.status === 'Pendente').length;
    
    // 2. Partially Delivered (Parcial) orders total
    const parciais = ordersSubset.filter((p) => p.status === 'Parcial').length;
    
    // 3. Critical Stock: Count materials assigned where available stock balance is low relative to their incoming expectations,
    // or when the stock of a needed item is exactly 0 but has active demand.
    // Let's count items where saldo < (recebido * 0.15) or saldo <= 0, but total solicitado is high.
    const subsetStock = stockInventory.filter((item) => selectedObra === 'Todas as Obras' || item.obra === selectedObra);
    const estoqueCritico = subsetStock.filter(item => {
      // Stock is critical if there's been some incoming, but available is 0, or is less than 15% of received
      return item.recebido > 0 && (item.saldo <= (item.recebido * 0.15));
    }).length;

    return {
      pendentes,
      parciais,
      estoqueCritico,
    };
  }, [pedidos, stockInventory, selectedObra]);

  const dashboardChartData = useMemo(() => {
    const ordersSubset = pedidos.filter((p) => selectedObra === 'Todas as Obras' || p.obra === selectedObra);
    const totalEntregues = ordersSubset.filter((p) => p.status === 'Entregue').length;
    const totalParciais = ordersSubset.filter((p) => p.status === 'Parcial').length;
    const totalPendentes = ordersSubset.filter((p) => p.status === 'Pendente').length;
    
    return [
      { name: 'Totalmente Entregue', value: totalEntregues, fill: '#10b981' },
      { name: 'Parcialmente Entregue', value: totalParciais, fill: '#3b82f6' },
      { name: 'Pendente', value: totalPendentes, fill: '#f59e0b' }
    ];
  }, [pedidos, selectedObra]);

  const dashboardCriticalStockList = useMemo(() => {
    const subsetStock = stockInventory.filter((item) => selectedObra === 'Todas as Obras' || item.obra === selectedObra);
    return subsetStock.filter(item => item.recebido > 0 && (item.saldo <= (item.recebido * 0.15)));
  }, [stockInventory, selectedObra]);

  // Distinct list of insumos belonging to a specific construction site for withdrawal dropdown
  const insumosForWithdrawal = useMemo(() => {
    return stockInventory
      .filter((item) => item.obra === newWithdrawObra && item.saldo > 0)
      .map((item) => ({ insumo: item.insumo, saldo: item.saldo, unidade: item.unidade }));
  }, [stockInventory, newWithdrawObra]);

  // Dynamic received materials efficiency statistics for the real/actual obras
  const obrasRecebimentoStats = useMemo(() => {
    const statsMap: { [obra: string]: { recebido: number; solicitado: number } } = {};
    
    // Initialize for all active obras
    obras.forEach(o => {
      statsMap[o.nome] = { recebido: 0, solicitado: 0 };
    });

    pedidos.forEach(p => {
      if (p.status === 'Cancelado') return;
      if (!statsMap[p.obra]) {
        statsMap[p.obra] = { recebido: 0, solicitado: 0 };
      }
      statsMap[p.obra].recebido += p.qtdRecebida || 0;
      statsMap[p.obra].solicitado += p.qtdSolicitada || 0;
    });

    return Object.keys(statsMap).map(obraName => {
      const { recebido, solicitado } = statsMap[obraName];
      const percent = solicitado > 0 ? Math.round((recebido / solicitado) * 100) : 0;
      return {
        name: obraName,
        percent,
        recebido,
        solicitado
      };
    });
  }, [pedidos, obras]);

  // Unified list of all entries (Entradas) and withdrawals (Saídas)
  const unifiedMovements = useMemo(() => {
    const list: any[] = [];
    
    // 1. Add received entries from pedidos
    pedidos.forEach((p) => {
      if (p.status === 'Cancelado') return;
      if (p.qtdRecebida > 0) {
        list.push({
          id: p.id,
          codigo: p.codigo || '',
          insumo: p.insumo,
          codDetalhe: p.codDetalhe || '',
          descricaoDetalhe: p.descricaoDetalhe || '',
          obra: p.obra,
          tipo: 'ENTRADA',
          quantidade: p.qtdRecebida,
          unidade: p.unidade,
          documento: p.notaFiscal ? `NF-${p.notaFiscal} (PED-${p.id})` : `NF-Recebimento (PED-${p.id})`,
          responsavel: p.responsavelRecebimento || 'Almoxarife / Portaria',
          data: p.dataChegada || p.dataPedido,
          fornecedor: p.fornecedor || '-',
          isEstornoable: true,
          isPedidoType: true,
          rawTimestamp: parseDateTimeString(p.dataChegada || p.dataPedido)?.getTime() || 0
        });
      }
    });

    // 2. Add withdrawals from baixas
    baixas.forEach((b) => {
      const refPedido = pedidos.find(p => (p.codigo && p.codigo === b.codigo) || (p.insumo && p.insumo === b.insumo));
      const isRecons = b.id.startsWith('RE-') || b.destino?.startsWith('Ajuste:');
      list.push({
        id: b.id,
        codigo: b.codigo || '',
        insumo: b.insumo,
        codDetalhe: refPedido?.codDetalhe || '',
        descricaoDetalhe: refPedido?.descricaoDetalhe || '',
        obra: b.obra,
        tipo: isRecons ? 'AJUSTE' : 'SAIDA',
        quantidade: b.quantidade,
        unidade: b.unidade,
        documento: b.destino || 'Retirada de Obra',
        responsavel: b.colaborador || 'Mestre de Obras',
        data: b.data,
        fornecedor: '-',
        isEstornoable: true,
        isPedidoType: false,
        rawTimestamp: parseDateTimeString(b.data)?.getTime() || 0
      });
    });

    // Sort by descending timestamp (latest first)
    return list.sort((a, b) => b.rawTimestamp - a.rawTimestamp);
  }, [pedidos, baixas]);

  // Distinct lists extracted dynamically from data for filter select-boxes
  const reportUsersList = useMemo(() => {
    const set = new Set<string>();
    usuarios.forEach(u => { if (u.nome) set.add(u.nome.trim()); });
    unifiedMovements.forEach(m => { if (m.responsavel) set.add(m.responsavel.trim()); });
    return Array.from(set).sort();
  }, [usuarios, unifiedMovements]);

  const reportMaterialsList = useMemo(() => {
    const set = new Set<string>();
    unifiedMovements.forEach(m => { if (m.insumo) set.add(m.insumo.trim()); });
    return Array.from(set).sort();
  }, [unifiedMovements]);

  const reportSuppliersList = useMemo(() => {
    const set = new Set<string>();
    pedidos.forEach(p => { if (p.fornecedor) set.add(p.fornecedor.trim()); });
    return Array.from(set).sort();
  }, [pedidos]);

  // Dynamically filtered movements based on selected user filters
  const filteredMovements = useMemo(() => {
    return unifiedMovements.filter((m) => {
      // 1. Obra filter
      if (reportSelectedObra !== 'Todas as Obras' && m.obra !== reportSelectedObra) {
        return false;
      }
      // 2. User/Responsável filter
      if (reportSelectedUser !== 'Todos os Usuários' && m.responsavel.trim() !== reportSelectedUser) {
        return false;
      }
      // 3. Movement Type filter
      if (reportSelectedType !== 'Ambos') {
        if (reportSelectedType === 'Entrada' && m.tipo !== 'ENTRADA') return false;
        if (reportSelectedType === 'Saída' && m.tipo !== 'SAIDA') return false;
        if (reportSelectedType === 'Ajuste' && m.tipo !== 'AJUSTE') return false;
      }
      // 4. Material/Insumo filter
      if (reportSelectedMaterial !== 'Todos os Materiais' && m.insumo.trim() !== reportSelectedMaterial) {
        return false;
      }
      // 5. Supplier/Fornecedor filter
      if (reportSelectedSupplier !== 'Todos os Fornecedores' && m.fornecedor.trim() !== reportSelectedSupplier) {
        return false;
      }
      // 6. Date period filter (between start date and end date)
      if (reportStartDate) {
        const itemDate = parseDateTimeString(m.data);
        if (itemDate) {
          const startDate = new Date(reportStartDate + 'T00:00:00');
          if (itemDate < startDate) return false;
        }
      }
      if (reportEndDate) {
        const itemDate = parseDateTimeString(m.data);
        if (itemDate) {
          const endDate = new Date(reportEndDate + 'T23:59:59');
          if (itemDate > endDate) return false;
        }
      }
      return true;
    });
  }, [
    unifiedMovements,
    reportSelectedObra,
    reportSelectedUser,
    reportSelectedType,
    reportSelectedMaterial,
    reportSelectedSupplier,
    reportStartDate,
    reportEndDate
  ]);

  // Handler for estorno/deletion of movements from diary
  const handleEstornoMovement = (item: any) => {
    if (item.tipo === 'SAIDA' || item.tipo === 'AJUSTE') {
      const actionName = item.tipo === 'AJUSTE' ? 'Estornar Ajuste de Reconciliação' : 'Estornar Movimentação de Saída';
      executeGuardedAction('dar_baixa', actionName, () => {
        setBaixas(prev => prev.filter(b => b.id !== item.id));
        triggerToast(item.tipo === 'AJUSTE' ? 'Reconciliação estornada! Saldo original restaurado.' : 'Lançamento de saída estornado com sucesso! Saldo reorganizado.', 'success');
      });
    } else {
      executeGuardedAction('receber_mercadoria', 'Estornar Recebimento de Entrada', () => {
        setPedidos(prev => prev.map(p => {
          if (p.id === item.id && (p.codigo || '') === item.codigo) {
            return {
              ...p,
              qtdRecebida: 0,
              status: 'Pendente',
              dataChegada: undefined,
              responsavelRecebimento: undefined,
              notaFiscal: undefined
            };
          }
          return p;
        }));
        triggerToast('Lançamento de entrada estornado! Quantidade deduzida do estoque.', 'success');
      });
    }
  };


  // Permission validator helper
  const checkPermission = (action: 'criar_pedido' | 'dar_baixa' | 'receber_mercadoria' | 'excluir_pedido' | 'excluir_obra' | 'admin'): boolean => {
    if (currentUser.role === 'Administrador') return true;
    
    switch (action) {
      case 'criar_pedido':
        return currentUser.podeCriarPedido;
      case 'dar_baixa':
        return currentUser.podeDarBaixa;
      case 'receber_mercadoria':
        return currentUser.podeReceberMercadoria;
      case 'excluir_pedido':
        return currentUser.podeExcluirPedido;
      case 'excluir_obra':
        return currentUser.podeExcluirObra;
      case 'admin':
        return false; // Only Admins can modify accounts and global settings
      default:
        return false;
    }
  };

  const executeGuardedAction = (
    action: 'criar_pedido' | 'dar_baixa' | 'receber_mercadoria' | 'excluir_pedido' | 'excluir_obra' | 'admin', 
    description: string, 
    callback: () => void
  ) => {
    if (checkPermission(action)) {
      callback();
    } else {
      let requiredPermissionName = "";
      switch (action) {
        case 'criar_pedido': requiredPermissionName = "Solicitar/Criar Pedidos de Compra"; break;
        case 'dar_baixa': requiredPermissionName = "Registrar Saídas/Baixas de Materiais"; break;
        case 'receber_mercadoria': requiredPermissionName = "Dar Entrada / Confirmar Recebimento de Cargas"; break;
        case 'excluir_pedido': requiredPermissionName = "Excluir Pedidos do Histórico"; break;
        case 'excluir_obra': requiredPermissionName = "Excluir Obras/Projetos"; break;
        case 'admin': requiredPermissionName = "Acesso de Administrador Geral (Controle de Usuários e Obras)"; break;
      }
      setLockModalReason(
        `Sua conta de usuário ativa (perfil "${currentUser.nome}") não possui autorização para: "${description}".\n\nRequisito necessário: Habilitar a permissão em "Administração" > "Usuários e Permissões" estando logado como Administrador corporativo.`
      );
    }
  };

  // Handler For Posting New Purchase Order/Insumo Request
  const handleCreateOrder = (e: React.FormEvent) => {
    e.preventDefault();
    executeGuardedAction('criar_pedido', 'Criar novo Pedido de Compra', () => {
      if (!newOrderInsumo.trim()) {
        triggerToast('Por favor, informe o nome do insumo.', 'warn');
        return;
      }
      if (newOrderQtd <= 0) {
        triggerToast('A quantidade solicitada deve ser maior que zero.', 'warn');
        return;
      }

      const randomId = `PC-2024-0${Math.floor(Math.random() * 900) + 100}`;
      const resolvedCodigo = newOrderCodigo.trim().toUpperCase() || `INS-${Math.floor(Math.random() * 900) + 100}`;
      const matchedObra = obras.find(o => o.nome === newOrderObra);
      const newPedido: Pedido = {
        id: randomId,
        insumo: newOrderInsumo.trim(),
        codigo: resolvedCodigo,
        obra: newOrderObra,
        obraId: matchedObra ? matchedObra.id : undefined,
        qtdSolicitada: newOrderQtd,
        qtdRecebida: 0,
        unidade: newOrderUnidade,
        status: 'Pendente',
        dataPedido: new Date().toLocaleDateString('pt-BR'),
      };

      setPedidos([newPedido, ...pedidos]);
      setMovementsCount(prev => prev + 1);
      setShowOrderModal(false);
      setNewOrderInsumo('');
      setNewOrderCodigo('');
      triggerToast(`Pedido ${randomId} cadastrado com sucesso para ${newOrderObra}!`, 'success');
    });
  };

  // Handler For Posting Material Withdrawal (Baixa)
  const handleCreateWithdrawal = (e: React.FormEvent) => {
    e.preventDefault();
    executeGuardedAction('dar_baixa', 'Registrar Baixa de Material em Obra', () => {
      if (!newWithdrawInsumo) {
        triggerToast('Informe o insumo a ser retirado.', 'warn');
        return;
      }

      if (!newWithdrawDestino.trim()) {
        triggerToast('O preenchimento do campo "4. Destino de Aplicação / Local" é obrigatório.', 'warn');
        return;
      }

      if (!newWithdrawAutorizado.trim()) {
        triggerToast('O preenchimento do campo "5. Autorizado por" é obrigatório.', 'warn');
        return;
      }

      const targetStockItem = stockInventory.find(
        (item) => item.obra === newWithdrawObra && item.insumo === newWithdrawInsumo
      );

      if (!targetStockItem) {
        triggerToast('Material não encontrado no estoque.', 'warn');
        return;
      }

      if (newWithdrawQtd > targetStockItem.saldo) {
        triggerToast(`A quantidade consumida (${newWithdrawQtd}) não pode ser maior que o saldo disponível (${targetStockItem.saldo}).`, 'warn');
        return;
      }

      if (newWithdrawQtd <= 0) {
        triggerToast('A quantidade consumida deve ser maior que zero.', 'warn');
        return;
      }

      const randomId = `BX-2024-0${Math.floor(Math.random() * 900) + 100}`;
      const matchedObra = obras.find(o => o.nome === newWithdrawObra);
      const newBaixa: Baixa = {
        id: randomId,
        insumo: newWithdrawInsumo,
        codigo: targetStockItem.codigo || '',
        obra: newWithdrawObra,
        obraId: matchedObra ? matchedObra.id : undefined,
        quantidade: newWithdrawQtd,
        unidade: targetStockItem.unidade,
        colaborador: newWithdrawAutorizado.trim(),
        data: new Date().toLocaleString('pt-BR', { hour12: false }).replace(',', ''),
        destino: newWithdrawDestino.trim(),
        autorizado: newWithdrawAutorizado.trim()
      };

      setBaixas([newBaixa, ...baixas]);
      setMovementsCount(prev => prev + 1);
      setShowWithdrawModal(false);
      
      // Show center screen overlay "Baixa realizada"
      setShowWithdrawSuccessScreen(true);
      
      setTimeout(() => {
        setShowWithdrawSuccessScreen(false);
        setCurrentTab('estoque');
        
        // Reset states
        setNewWithdrawColab('');
        setNewWithdrawDestino('');
        setNewWithdrawAutorizado('');
        setNewWithdrawInsumo('');
      }, 1800);
    });
  };

  // Prepares the delivery reception modal
  const openReceiveModal = (id: string, codigo: string) => {
    executeGuardedAction('receber_mercadoria', 'Registrar Entrada de Materiais', () => {
      const ped = pedidos.find(p => p.id === id && p.codigo === codigo);
      if (ped) {
        setSelectedPedidoId(id);
        setSelectedPedidoCodigo(codigo);
        setQtdReceiveInput(''); // Começa não preenchido (vazio) conforme solicitado
        setReceiptNotaFiscal('');
        setReceiptResponsavel(currentUser.nome || '');
        setReceiptStatusAlteracao('deduzir_auto');
        setShowReceiveModal(true);
      }
    });
  };

  // Processes delivery check-in
  const handleConfirmReceipt = (e: React.FormEvent) => {
    e.preventDefault();
    if (!receiptResponsavel || !receiptResponsavel.trim()) {
      triggerToast('O Responsável de Almoxarifado é obrigatório.', 'warn');
      return;
    }
    const inputVal = Number(qtdReceiveInput);
    if (!qtdReceiveInput || inputVal <= 0) {
      triggerToast('A quantidade recebida deve ser positiva.', 'warn');
      return;
    }

    const currentPed = pedidos.find(p => p.id === selectedPedidoId && p.codigo === selectedPedidoCodigo);
    if (currentPed) {
      const remaining = currentPed.qtdSolicitada - currentPed.qtdRecebida;
      if (inputVal > remaining) {
        triggerToast(`A quantidade informada (${inputVal}) não pode exceder o saldo pendente de ${remaining} ${currentPed.unidade.toUpperCase()}`, 'warn');
        return;
      }
    }

    setPedidos(
      pedidos.map((p) => {
        if (p.id === selectedPedidoId && p.codigo === selectedPedidoCodigo) {
          const newRec = Number(p.qtdRecebida) + inputVal;
           let newStatus: 'Pendente' | 'Parcial' | 'Entregue' | 'Cancelado' = 'Parcial';
           let dataCheg = p.dataChegada;

          if (receiptStatusAlteracao === 'entregue') {
            newStatus = 'Entregue';
            dataCheg = new Date().toLocaleDateString('pt-BR');
          } else if (receiptStatusAlteracao === 'parcial') {
            newStatus = 'Parcial';
          } else if (receiptStatusAlteracao === 'pendente') {
            newStatus = 'Pendente';
          } else if (receiptStatusAlteracao === 'cancelado') {
            newStatus = 'Cancelado';
          } else {
            // deduzir_auto
            if (newRec >= p.qtdSolicitada) {
              newStatus = 'Entregue';
              dataCheg = new Date().toLocaleDateString('pt-BR'); // Record arrival date today!
            } else if (newRec > 0) {
              newStatus = 'Parcial';
            } else {
              newStatus = 'Pendente';
            }
          }

          return {
            ...p,
            qtdRecebida: Math.min(newRec, p.qtdSolicitada), // clamp to target
            status: newStatus,
            dataChegada: dataCheg
          };
        }
        return p;
      })
    );

    setMovementsCount(prev => prev + 1);
    setShowReceiveModal(false);
    triggerToast('PEDIDO RECEBIDO COM SUCESSO', 'success');
  };

  // Function to process imported Sienge items (UPSERT / MERGE same-pedido items)
  const processImportedPedidos = (importedItems: Pedido[]) => {
    let newOrdersCount = 0;
    let newItemsCount = 0;
    let updatedQuantitiesCount = 0;
    
    let updatedList = [...pedidos];
    const existingOrderIds = new Set(pedidos.map(p => p.id));
    
    // Register any new construction sites referenced in the spreadsheet
    const uniqueObrasInImport = Array.from(new Set(importedItems.map(r => r.obra)));
    const finalObrasList = [...obras];
    const activeObrasNames = obras.map(o => o.nome.toLowerCase().trim());
    const newObrasToRegister = uniqueObrasInImport.filter(o => o && !activeObrasNames.includes(o.toLowerCase().trim()));
    if (newObrasToRegister.length > 0) {
      const startId = obras.length + 101;
      newObrasToRegister.forEach((name, idx) => {
        // Look up corresponding spreadsheet codObra (obraId) if available
        const matchedItem = importedItems.find(item => item.obra && item.obra.toLowerCase().trim() === name.toLowerCase().trim());
        finalObrasList.push({
          id: matchedItem?.obraId || `OB-${startId + idx}`,
          nome: name
        });
      });
      setObras(finalObrasList);
    }
    
    importedItems.forEach((newItem) => {
      // Find matching item in current orders list: same order ID and same insumo (or code)
      const exactMatchIndex = updatedList.findIndex(
        (p) => p.id === newItem.id && 
               ((p.codigo && newItem.codigo && p.codigo === newItem.codigo) || 
                (p.insumo || '').toLowerCase().trim() === (newItem.insumo || '').toLowerCase().trim())
      );
      
      const newItemObraNormalized = (newItem.obra || '').toLowerCase().trim();
      const matchedObra = finalObrasList.find(o => (o.nome || '').toLowerCase().trim() === newItemObraNormalized);
      const resolvedObraId = matchedObra ? matchedObra.id : newItem.obraId;
      
      if (exactMatchIndex !== -1) {
        // Exists: update details if changed
        const existingItem = updatedList[exactMatchIndex];
        const newQtdSolicitada = newItem.qtdSolicitada;
        const currentQtdRecebida = existingItem.qtdRecebida;
        
        // If the spreadsheet provides status entrega, use it responsibly
        let newStatus = newItem.status !== 'Pendente' ? newItem.status : existingItem.status;
        if (newStatus !== 'Cancelado') {
          if (currentQtdRecebida >= newQtdSolicitada) {
            newStatus = 'Entregue';
          } else if (currentQtdRecebida > 0) {
            newStatus = 'Parcial';
          } else if (newStatus === 'Entregue' && currentQtdRecebida < newQtdSolicitada) {
            newStatus = 'Parcial';
          }
        }
        
        updatedList[exactMatchIndex] = {
          ...existingItem,
          obra: newItem.obra || existingItem.obra,
          obraId: resolvedObraId,
          unidade: newItem.unidade || existingItem.unidade,
          qtdSolicitada: newQtdSolicitada,
          status: newStatus,
          codigo: newItem.codigo || existingItem.codigo,
          insumo: newItem.insumo || existingItem.insumo,
          
          // New Sienge advanced fields mapping - keep last imported supplier
          codComprador: newItem.codComprador || existingItem.codComprador,
          fornecedor: newItem.fornecedor || existingItem.fornecedor, // "considerar sempre o ultimo"
          codDetalhe: newItem.codDetalhe || existingItem.codDetalhe,
          descricaoDetalhe: newItem.descricaoDetalhe || existingItem.descricaoDetalhe,
          marca: newItem.marca || existingItem.marca,
          descricaoUnidade: newItem.descricaoUnidade || existingItem.descricaoUnidade,
          qtdPendenteImportada: newItem.qtdPendenteImportada !== undefined ? newItem.qtdPendenteImportada : existingItem.qtdPendenteImportada,
          dataPedido: newItem.dataPedido || existingItem.dataPedido,
          dataChegada: newStatus === 'Entregue' ? (existingItem.dataChegada || newItem.dataChegada || new Date().toLocaleDateString('pt-BR')) : undefined
        };
        updatedQuantitiesCount++;
      } else {
        // Brand new item
        if (existingOrderIds.has(newItem.id)) {
          newItemsCount++;
        } else {
          newOrdersCount++;
          existingOrderIds.add(newItem.id);
        }
        
        updatedList.push({
          ...newItem,
          obraId: resolvedObraId,
          qtdRecebida: newItem.qtdRecebida || 0,
          status: newItem.status || 'Pendente',
          dataChegada: newItem.status === 'Entregue' ? (newItem.dataChegada || new Date().toLocaleDateString('pt-BR')) : undefined
        });
      }
    });

    // Map all final items to their resolved or newly built obraId
    const finalMappedList = updatedList.map(item => {
      const itemObraNormalized = (item.obra || '').toLowerCase().trim();
      const matched = finalObrasList.find(o => (o.nome || '').toLowerCase().trim() === itemObraNormalized);
      return {
        ...item,
        obraId: matched ? matched.id : item.obraId
      };
    });
    
    setPedidos(finalMappedList);
    setImportStats({
      newOrders: newOrdersCount,
      newItems: newItemsCount,
      updatedQtds: updatedQuantitiesCount
    });

    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const hours = String(now.getHours()).padStart(2, '0');
    const minutes = String(now.getMinutes()).padStart(2, '0');
    const formattedDate = `${day}/${month}/${year} às ${hours}:${minutes}`;
    setLastImportDate(formattedDate);
    localStorage.setItem('last_import_date', formattedDate);
    
    triggerToast(
      `Verificação concluída! ${newOrdersCount} novos pedidos, ${newItemsCount} itens adicionados e ${updatedQuantitiesCount} quantidades atualizadas!`,
      'success'
    );
  };

  // Parse Sienge spreadsheet columns and rows with grouping support
  const parseSiengeSheet = (sheetData: any[][]) => {
    if (sheetData.length <= 1) return [];
    
    // Normalization helper to strip accents, keep only lowercase alphanumeric and spaces
    const cleanStr = (val: any): string => {
      if (!val) return '';
      return String(val)
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '') // remove accents
        .replace(/[^a-z0-9 ]/g, '') // strip punctuation & special symbols, keep spaces
        .trim();
    };

    let headerRowIndex = 0;
    let foundHeader = false;
    
    // Step 1: Let's search for a row containing at least 3 distinct Sienge header keywords
    for (let r = 0; r < Math.min(50, sheetData.length); r++) {
      const row = sheetData[r];
      if (row && Array.isArray(row)) {
        let matchCount = 0;
        row.forEach(cell => {
          const s = cleanStr(cell);
          if (
            s.includes('pedido') ||
            s.includes('insumo') ||
            s.includes('obra') ||
            s.includes('fornecedor') ||
            s.includes('comprador') ||
            s.includes('marca') ||
            s.includes('quant') ||
            s.includes('unidade') ||
            s.includes('emiss') ||
            s.includes('detalhe')
          ) {
            matchCount++;
          }
        });
        if (matchCount >= 3) {
          headerRowIndex = r;
          foundHeader = true;
          break;
        }
      }
    }
    
    // Step 2: Fallback to any row with at least 1 typical match if no robust header row was found
    if (!foundHeader) {
      for (let r = 0; r < Math.min(25, sheetData.length); r++) {
        const row = sheetData[r];
        if (row && Array.isArray(row) && row.some(cell => {
          const s = cleanStr(cell);
          return s.includes('n pedido') || s.includes('no pedido') || s.includes('data pedido') || s.includes('descricao insumo') || s.includes('desc insumo');
        })) {
          headerRowIndex = r;
          foundHeader = true;
          break;
        }
      }
    }
    
    const headers = sheetData[headerRowIndex].map(h => cleanStr(h));
    
    let colId = -1;
    let colInsumo = -1;
    let colCodigo = -1;
    let colObra = -1;
    let colCodObra = -1;
    let colQtd = -1;
    let colUnidade = -1;
    let colData = -1;
    
    // New Sienge columns
    let colCodComprador = -1;
    let colFornecedor = -1;
    let colCodDetalhe = -1;
    let colDescDetalhe = -1;
    let colMarca = -1;
    let colDescUnidade = -1;
    let colStatusEntrega = -1;
    let colQtdPendente = -1;
    
    headers.forEach((s, idx) => {
      // s is already cleanStr normalized (lowercased, no accents, no symbols except spaces)
      
      // 1. Nº Pedido
      if (s === 'n pedido' || s === 'no pedido' || s === 'npedido' || s === 'numero pedido' || s.includes('n ped') || s.includes('no ped') || s.includes('numped') || s === 'pedido' || s === 'id' || s === 'pc') {
        colId = idx;
      }
      // 2. Data pedido
      else if (s === 'data pedido' || s === 'data do pedido' || s === 'data emissao' || s.includes('data') || s.includes('emissao') || s === 'dt pedido' || s === 'dtpedido') {
        colData = idx;
      }
      // 3. Cód. Obra
      else if (s === 'cod obra' || s === 'codigo obra' || (s.includes('obra') && (s.includes('cod') || s.includes('codigo')))) {
        colCodObra = idx;
      }
      // 4. Obra
      else if (s === 'obra' || s === 'nome obra' || s.includes('empreendimento') || s.includes('filial') || s.includes('local')) {
        colObra = idx;
      }
      // 5. Cód. Comprador
      else if (s === 'cod comprador' || s === 'codigo comprador' || s.includes('comprador')) {
        colCodComprador = idx;
      }
      // 6. Fornecedor
      else if (s === 'fornecedor' || s === 'forn' || s === 'parceiro' || s.includes('fornecedor')) {
        colFornecedor = idx;
      }
      // 7. Cód. Insumo
      else if (s === 'cod insumo' || s === 'codigo insumo' || s === 'codinsumo' || s === 'insumo codigo') {
        colCodigo = idx;
      }
      // 8. Descrição insumo
      else if (s === 'descricao insumo' || s === 'desc insumo' || s === 'insumo' || s === 'material' || s === 'item' || s === 'produto' || (s.includes('desc') && s.includes('insumo'))) {
        colInsumo = idx;
      }
      // 9. Cód. Detalhe
      else if (s === 'cod detalhe' || s === 'codigo detalhe' || (s.includes('detalhe') && s.includes('cod'))) {
        colCodDetalhe = idx;
      }
      // 10. Descrição detalhe
      else if (s === 'descricao detalhe' || s === 'desc detalhe' || s === 'detalhe' || (s.includes('desc') && s.includes('detalhe'))) {
        colDescDetalhe = idx;
      }
      // 11. Descrição marca
      else if (s === 'descricao marca' || s === 'desc marca' || s === 'marca' || s.includes('marca')) {
        colMarca = idx;
      }
      // 12. Símbolo unidade medida
      else if (s === 'simbolo unidade medida' || s === 'simbolo unidade' || s === 'simb unidade' || s === 'unid' || s === 'un' || s === 'unidade' || s === 'um') {
        colUnidade = idx;
      }
      // 13. Descrição unidade medida
      else if (s === 'descricao unidade medida' || (s.includes('desc') && s.includes('unidade'))) {
        colDescUnidade = idx;
      }
      // 14. Status entrega
      else if (s === 'status entrega' || s === 'statusentrega' || s === 'status do pedido' || s === 'status' || s.includes('status')) {
        colStatusEntrega = idx;
      }
      // 15. Quant. pendente
      else if (s === 'quant pendente' || s === 'quantidade pendente' || s === 'qtd pendente' || s === 'saldo pendente' || s === 'pendente' || s.includes('pendente')) {
        colQtdPendente = idx;
      }
      // Quantidade comprada / solicitada fallback
      else if (s === 'quant solicitada' || s === 'quantidade solicitada' || s === 'qtd solicitada' || s === 'quantidade comprada' || s === 'qtd comprada' || s === 'quantidade' || s === 'qtd' || s === 'volume') {
        colQtd = idx;
      }
    });
    
    // Safety automatic index fallbacks matching exact columns of the Sienge spreadsheet:
    // A=0: Nº Pedido, B=1: Data pedido, C=2: Cód. Obra, D=3: Obra, E=4: Cód. Comprador,
    // F=5: Fornecedor, G=6: Cód. Insumo, H=7: Descrição insumo, I=8: Cód. Detalhe,
    // J=9: Descrição detalhe, K=10: Descrição marca, L=11: Símbolo unidade medida,
    // M=12: Descrição unidade medida, N=13: Quant. solicitada, O=14: Status entrega, P=15: Quant. pendente
    if (colId === -1) colId = 0;
    if (colData === -1) colData = 1;
    if (colCodObra === -1) colCodObra = 2;
    if (colObra === -1) colObra = 3;
    if (colCodComprador === -1) colCodComprador = 4;
    if (colFornecedor === -1) colFornecedor = 5;
    if (colCodigo === -1) colCodigo = 6;
    if (colInsumo === -1) colInsumo = 7;
    if (colCodDetalhe === -1) colCodDetalhe = 8;
    if (colDescDetalhe === -1) colDescDetalhe = 9;
    if (colMarca === -1) colMarca = 10;
    if (colUnidade === -1) colUnidade = 11;
    if (colDescUnidade === -1) colDescUnidade = 12;
    if (colQtd === -1) colQtd = 13;
    if (colStatusEntrega === -1) colStatusEntrega = 14;
    if (colQtdPendente === -1) colQtdPendente = 15;
    
    const results: Pedido[] = [];
    let lastOrderId = '';
    let lastObra = listObras[0] || 'Residencial Alvorada';
    let lastObraId = '';
    let lastData = new Date().toLocaleDateString('pt-BR');
    
    for (let r = headerRowIndex + 1; r < sheetData.length; r++) {
      const row = sheetData[r];
      if (!row || row.length === 0 || row.every(cell => cell === null || cell === '')) continue;
      
      let rawIdVal = colId !== -1 && row[colId] !== undefined ? row[colId] : '';
      let rawId = '';
      if (typeof rawIdVal === 'number') {
        rawId = String(Math.floor(rawIdVal));
      } else {
        rawId = String(rawIdVal).trim();
      }

      let rawInsumo = colInsumo !== -1 && row[colInsumo] !== undefined ? String(row[colInsumo]).trim() : '';
      let rawCodigo = colCodigo !== -1 && row[colCodigo] !== undefined ? String(row[colCodigo]).trim() : '';
      let rawObra = colObra !== -1 && row[colObra] !== undefined ? String(row[colObra]).trim() : '';
      let rawCodObra = colCodObra !== -1 && row[colCodObra] !== undefined ? String(row[colCodObra]).trim() : '';
      let rawQtd = colQtd !== -1 && row[colQtd] !== undefined ? parseFloat(String(row[colQtd]).replace(/[^\d.-]/g, '')) : 0;
      let rawUnidade = colUnidade !== -1 && row[colUnidade] !== undefined ? String(row[colUnidade]).trim() : 'un';
      
      let rawDataVal = colData !== -1 && row[colData] !== undefined ? row[colData] : '';
      let rawData = '';
      if (typeof rawDataVal === 'number') {
        const dateObj = new Date((rawDataVal - 25569) * 86400 * 1000);
        if (!isNaN(dateObj.getTime())) {
          const dy = String(dateObj.getUTCDate()).padStart(2, '0');
          const mn = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
          const yr = dateObj.getUTCFullYear();
          rawData = `${dy}/${mn}/${yr}`;
        } else {
          rawData = String(rawDataVal).trim();
        }
      } else if (rawDataVal instanceof Date) {
        const dy = String(rawDataVal.getUTCDate()).padStart(2, '0');
        const mn = String(rawDataVal.getUTCMonth() + 1).padStart(2, '0');
        const yr = rawDataVal.getUTCFullYear();
        rawData = `${dy}/${mn}/${yr}`;
      } else {
        const trimmed = String(rawDataVal).trim();
        if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
          const parts = trimmed.split('T')[0].split('-');
          rawData = `${parts[2]}/${parts[1]}/${parts[0]}`;
        } else {
          rawData = trimmed;
        }
      }
      
      // New fields parsing
      let rawCodComprador = colCodComprador !== -1 && row[colCodComprador] !== undefined ? String(row[colCodComprador]).trim() : '';
      let rawFornecedor = colFornecedor !== -1 && row[colFornecedor] !== undefined ? String(row[colFornecedor]).trim() : '';
      let rawCodDetalhe = colCodDetalhe !== -1 && row[colCodDetalhe] !== undefined ? String(row[colCodDetalhe]).trim() : '';
      let rawDescDetalhe = colDescDetalhe !== -1 && row[colDescDetalhe] !== undefined ? String(row[colDescDetalhe]).trim() : '';
      let rawMarca = colMarca !== -1 && row[colMarca] !== undefined ? String(row[colMarca]).trim() : '';
      let rawDescUnidade = colDescUnidade !== -1 && row[colDescUnidade] !== undefined ? String(row[colDescUnidade]).trim() : '';
      let rawStatusEntrega = colStatusEntrega !== -1 && row[colStatusEntrega] !== undefined ? String(row[colStatusEntrega]).trim() : '';
      let rawQtdPendente = colQtdPendente !== -1 && row[colQtdPendente] !== undefined ? parseFloat(String(row[colQtdPendente]).replace(/[^\d.-]/g, '')) : undefined;
      
      if (!rawInsumo || rawInsumo.toLowerCase() === 'insumo' || rawInsumo.toLowerCase() === 'descrição') continue;
      
      if (rawId) lastOrderId = rawId;
      if (rawObra) lastObra = rawObra;
      if (rawCodObra) lastObraId = rawCodObra;
      if (rawData) lastData = rawData;
      
      const finalId = lastOrderId || 'PC-IMPORTADO';
      const finalObra = lastObra || 'Todas as Obras';
      const finalObraId = rawCodObra || lastObraId || undefined;
      const finalData = lastData || new Date().toLocaleDateString('pt-BR');
      
      let finalCodigo = rawCodigo;
      if (!finalCodigo) {
        const prefix = rawInsumo.substring(0, 3).toUpperCase().replace(/[^A-Z]/g, 'GEN');
        finalCodigo = `${prefix}-${Math.floor(100 + Math.random() * 900)}`;
      }
      
      let statusMapped: 'Pendente' | 'Parcial' | 'Entregue' | 'Cancelado' = 'Pendente';
      if (rawStatusEntrega) {
        const lowerStatus = rawStatusEntrega.toLowerCase();
        if (lowerStatus.includes('entregue') || lowerStatus.includes('concluido') || lowerStatus.includes('concluído') || lowerStatus.includes('chegou') || lowerStatus.includes('recebido') || lowerStatus === 'e') {
          statusMapped = 'Entregue';
        } else if (lowerStatus.includes('parcial') || lowerStatus === 'p') {
          statusMapped = 'Parcial';
        } else if (lowerStatus.includes('cancelado') || lowerStatus === 'c') {
          statusMapped = 'Cancelado';
        }
      } else if (rawQtdPendente !== undefined && rawQtd > 0) {
        if (rawQtdPendente <= 0) {
          statusMapped = 'Entregue';
        } else if (rawQtdPendente < rawQtd) {
          statusMapped = 'Parcial';
        }
      }
      
      // Calculate parsed quantity request
      let computedQtdSolicitada = isNaN(rawQtd) || rawQtd <= 0 ? 1 : rawQtd;
      let computedQtdRecebida = 0;
      if (statusMapped === 'Entregue') {
        computedQtdRecebida = computedQtdSolicitada;
      } else if (rawQtdPendente !== undefined && rawQtdPendente < computedQtdSolicitada) {
        computedQtdRecebida = computedQtdSolicitada - rawQtdPendente;
      }
      
      const matched = obras.find(o => o.nome.toLowerCase().trim() === finalObra.toLowerCase().trim());
      results.push({
        id: finalId,
        insumo: rawInsumo,
        codigo: finalCodigo,
        obra: finalObra,
        obraId: finalObraId || (matched ? matched.id : undefined),
        qtdSolicitada: computedQtdSolicitada,
        qtdRecebida: computedQtdRecebida,
        unidade: rawUnidade || 'un',
        status: statusMapped,
        dataPedido: finalData,
        dataChegada: statusMapped === 'Entregue' ? new Date().toLocaleDateString('pt-BR') : undefined,
        
        // Extended Sienge properties
        codComprador: rawCodComprador || undefined,
        fornecedor: rawFornecedor || undefined,
        codDetalhe: rawCodDetalhe || undefined,
        descricaoDetalhe: rawDescDetalhe || undefined,
        marca: rawMarca || undefined,
        descricaoUnidade: rawDescUnidade || undefined,
        qtdPendenteImportada: rawQtdPendente
      });
    }
    
    return results;
  };

  const preprocessSemicolonCsv = (rows: any[][]): any[][] => {
    if (!rows || rows.length === 0) return [];
    
    // Check if rows are single-column and contain semicolons or commas
    const testRows = rows.filter(r => r && r.length > 0).slice(0, 8);
    const hasSemicolon = testRows.some(r => r.length === 1 && String(r[0]).includes(';'));
    const hasComma = !hasSemicolon && testRows.some(r => r.length === 1 && String(r[0]).includes(','));
    
    if (hasSemicolon) {
      return rows.map(row => {
        if (row && row.length === 1) {
          const line = String(row[0]);
          return line.split(';').map(cell => cell.replace(/^["']|["']$/g, '').trim());
        }
        return row;
      });
    } else if (hasComma) {
      return rows.map(row => {
        if (row && row.length === 1) {
          const line = String(row[0]);
          return line.split(',').map(cell => cell.replace(/^["']|["']$/g, '').trim());
        }
        return row;
      });
    }
    return rows;
  };

  const processParsedData = (jsonData: any[][]) => {
    const parsed = parseSiengeSheet(jsonData);
    if (parsed.length === 0) {
      triggerToast('Nenhum item válido foi encontrado na planilha Sienge. Verifique as colunas.', 'warn');
    } else {
      setParsedImportItems(parsed);
      triggerToast(`${parsed.length} itens detectados e prontos para sincronização!`, 'success');
    }
  };

  const handleUploadSienge = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    
    setImportedFileName(file.name);

    if (file.name.endsWith('.csv')) {
      const csvReader = new FileReader();
      csvReader.onload = (csvEvent) => {
        try {
          const text = csvEvent.target?.result as string;
          // read as string directly
          const wb = XLSX.read(text, { type: 'string' });
          const firstSheet = wb.SheetNames[0];
          const ws = wb.Sheets[firstSheet];
          let rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
          rows = preprocessSemicolonCsv(rows);
          processParsedData(rows);
        } catch (err) {
          console.error("String CSV parsing failed, fallback to array buffer: ", err);
          // Fallback to array buffer but using ISO-8859-1 conversion check or similar
          const fallbackReader = new FileReader();
          fallbackReader.onload = (fEvent) => {
            try {
              const data = new Uint8Array(fEvent.target?.result as ArrayBuffer);
              const wb = XLSX.read(data, { type: 'array' });
              const firstSheet = wb.SheetNames[0];
              const ws = wb.Sheets[firstSheet];
              let rows = XLSX.utils.sheet_to_json(ws, { header: 1 }) as any[][];
              rows = preprocessSemicolonCsv(rows);
              processParsedData(rows);
            } catch (innerErr) {
              console.error(innerErr);
              triggerToast('Erro ao ler arquivo CSV de salvamento.', 'warn');
            }
          };
          fallbackReader.readAsArrayBuffer(file);
        }
      };
      // Sienge exports in Portuguese locales are famously coded in Windows-1252 / ISO-8859-1
      csvReader.readAsText(file, 'ISO-8859-1');
    } else {
      // Excel workbook binary reading
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
          processParsedData(jsonData);
        } catch (err: any) {
          console.error(err);
          triggerToast('Erro ao ler arquivo Excel. Verifique a formato.', 'warn');
        }
      };
      reader.readAsArrayBuffer(file);
    }
  };

  const handleLoadDemoSienge = () => {
    const demoItems: Pedido[] = [
      { 
        id: 'PC-2026-9040', 
        insumo: 'Cimento CP-II 50kg', 
        codigo: 'CIM-001', 
        obra: 'Residencial Alvorada', 
        obraId: 'OB-101',
        qtdSolicitada: 750, 
        qtdRecebida: 0, 
        unidade: 'SC', 
        status: 'Pendente', 
        dataPedido: '10/06/2026',
        codComprador: 'COMP-A',
        fornecedor: 'Votorantim Metais',
        codDetalhe: 'DET-CIM',
        descricaoDetalhe: 'Premium de Secagem Rápida',
        marca: 'Votoran',
        descricaoUnidade: 'Saco 50kg'
      },
      { 
        id: 'PC-2026-9040', 
        insumo: 'Pedra Britada Zero', 
        codigo: 'PED-102', 
        obra: 'Residencial Alvorada', 
        obraId: 'OB-101',
        qtdSolicitada: 25, 
        qtdRecebida: 0, 
        unidade: 'm³', 
        status: 'Pendente', 
        dataPedido: '12/06/2026',
        codComprador: 'COMP-B',
        fornecedor: 'Pedreira Guarulhos Ltda',
        codDetalhe: 'DET-PED',
        descricaoDetalhe: 'Material Lavado de Alta Pureza',
        marca: 'Nacional',
        descricaoUnidade: 'Metro Cúbico'
      },
      { 
        id: 'PC-2026-2510', 
        insumo: 'Tubo PVC Soldável 25mm', 
        codigo: 'TUB-404', 
        obra: 'Vista Parl', 
        obraId: 'OB-202',
        qtdSolicitada: 150, 
        qtdRecebida: 0, 
        unidade: 'm', 
        status: 'Pendente', 
        dataPedido: '14/06/2026',
        codComprador: 'COMP-A',
        fornecedor: 'Tigre S/A Tubos',
        codDetalhe: 'DET-TUB',
        descricaoDetalhe: 'Série Reforçada Predial Classe A',
        marca: 'Tigre',
        descricaoUnidade: 'Metros lineares'
      },
      { 
        id: 'PC-2026-2510', 
        insumo: 'Curva PVC 90 graus 25mm', 
        codigo: 'CUR-405', 
        obra: 'Vista Parl', 
        obraId: 'OB-202',
        qtdSolicitada: 80, 
        qtdRecebida: 0, 
        unidade: 'un', 
        status: 'Pendente', 
        dataPedido: '15/06/2026',
        codComprador: 'COMP-A',
        fornecedor: 'Amanco Brasil',
        codDetalhe: 'DET-CUR',
        descricaoDetalhe: 'Conexão Predial Água Fria',
        marca: 'Amanco',
        descricaoUnidade: 'Unidade'
      }
    ];
    setParsedImportItems(demoItems);
    triggerToast('Planilha modelo Sienge avançada (4 itens) carregada para testes!', 'info');
  };

  // Synchronizes state with Google Sheets App Script URL
  const handleSheetsSync = async (direction: 'pull' | 'push', isSilent: boolean = false) => {
    if (!googleSheetsUrl) {
      if (!isSilent) triggerToast('Insira a URL do Apps Script primeiro para realizar a sincronização.', 'warn');
      return;
    }
    
    if (!isSilent) setIsSyncing(true);
    try {
      if (direction === 'pull') {
        const response = await fetch(`${googleSheetsUrl}?action=getData`);
        if (!response.ok) throw new Error('Não foi possível conectar ao Google Sheets.');
        const result = await response.json();
        
        if (result.status === 'success' || result.status === 'ok' || result.obras) {
          if (result.pedidos) setPedidos(result.pedidos);
          if (result.baixas) setBaixas(result.baixas);
          if (result.obras) setObras(result.obras);
          if (result.usuarios) setUsuarios(result.usuarios);
          if (!isSilent) triggerToast('Dados sincronizados da planilha com sucesso!', 'success');
        } else {
          throw new Error(result.message || 'Erro na resposta do servidor.');
        }
      } else {
        // Send state to sheets
        const payload = {
          action: 'saveData',
          pedidos,
          baixas,
          obras,
          usuarios
        };
        const response = await fetch(googleSheetsUrl, {
          method: 'POST',
          mode: 'no-cors', // standard Apps Script POST bypass
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify(payload)
        });
        
        if (!isSilent) triggerToast('Dados enviados para a planilha com sucesso! (Deploy atualizado)', 'success');
      }
    } catch (err: any) {
      console.error(err);
      if (!isSilent) triggerToast(`Erro na sincronização: ${err.message || 'Verifique se a URL já foi publicada para acesso "Qualquer pessoa/Anyone".'}`, 'warn');
    } finally {
      if (!isSilent) setIsSyncing(false);
    }
  };

  // ZIP Exporter using JSZip
  const handleExportProjectZip = async () => {
    const zip = new JSZip();
    
    const addFileToZip = async (filePath: string, zipPath: string, fallbackContent: string = '') => {
      try {
        const response = await fetch(filePath);
        if (response.ok) {
          const content = await response.text();
          zip.file(zipPath, content);
        } else {
          zip.file(zipPath, fallbackContent);
        }
      } catch (e) {
        zip.file(zipPath, fallbackContent);
      }
    };

    triggerToast('Empacotando projeto no arquivo ZIP...', 'info');

    // Package key files
    await addFileToZip('/package.json', 'package.json');
    await addFileToZip('/index.html', 'index.html');
    await addFileToZip('/vite.config.ts', 'vite.config.ts');
    await addFileToZip('/code.gs', 'code.gs');
    await addFileToZip('/src/main.tsx', 'src/main.tsx');
    await addFileToZip('/src/App.tsx', 'src/App.tsx');
    await addFileToZip('/src/index.css', 'src/index.css');
    await addFileToZip('/tailwind.config.js', 'tailwind.config.js');
    await addFileToZip('/tsconfig.json', 'tsconfig.json');

    // Generate custom README deployment guide
    const readmeContent = `# Construmais Construtora - Guia de Implantação do Sistema

Parabéns! Você baixou os arquivos completos do painel integrado.

## Arquivos inclusos:
- \`code.gs\`: Script de integração do Google Sheets (Apps Script).
- \`src/App.tsx\`: Motor principal com interface web corporativa em React.
- \`src/index.css\` & \`vite.config.ts\`: Design system Tailwind CSS.
- \`package.json\`: Dependências do projeto.

---

## 1º PASSO: Configurar o Google Sheets (Banco de Dados)
Como as planilhas do Google Sheets funcionarão como seu banco de dados, você deve linká-las da seguinte forma:

1. **Crie uma nova Planilha Google**.
2. No menu superior da planilha, clique em **Extensões** > **Apps Script**.
3. Apague todo o código gerado por padrão e **cole as instruções descritas no arquivo \`code.gs\`** fornecido neste pacote.
4. Clique no botão de **Salvar** (ícone de disquete).
5. Clique em **Implantar** (Deploy) > **Nova Implantação**.
6. Em "Selecionar tipo", escolha **Editor da Web (Web App)**.
7. Altere as configurações de acesso:
   - *Executar como*: **Eu** (Sua conta Google).
   - *Quem tem acesso*: **Qualquer Pessoa** (Anyone - para permitir que o front-end envie e leia dados com segurança).
8. Clique em **Implantar**. Autorize as permissões necessárias na sua conta Google.
9. Copiar a **URL do Aplicativo Web** gerada (ex: \`https://script.google.com/macros/s/.../exec\`).

---

## 2º PASSO: Conectar a URL no Aplicativo
No Painel Construmais:
1. Vá na aba **Administração (Painel de Administração Corporal & Permissões)**.
2. Na seção **Conexão em Nuvem Google Sheets**, insira a URL copiada do Apps Script.
3. Seus dados estarão integrados e persistidos em tempo real na nuvem do Google Sheets!

---

## 3º PASSO: Hospedagem Gratuita na Vercel (Front-end)
Você pode subir o código do Front-end na Vercel de forma ultra rápida:
1. Extraia o conteúdo deste arquivo ZIP.
2. Crie um repositório no seu **Github** e suba todos os arquivos extraídos (exceção da pasta node_modules).
3. Entre na [Vercel](https://vercel.com/) e faça login com seu Github.
4. Clique em **Add New** > **Project** e selecione o repositório criado.
5. As configurações padrão de Vite/React já estarão preenchidas. Clique em **Deploy**.
6. Pronto! Seu aplicativo estará publicado mundialmente de forma segura e rápida.
`;

    zip.file('README_GUIA_DE_IMPLANTACAO.md', readmeContent);

    try {
      const contentBlob = await zip.generateAsync({ type: 'blob' });
      const downloadUrl = URL.createObjectURL(contentBlob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = `construmais_projeto_completo.zip`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      triggerToast('Seu arquivo ZIP de implantação foi baixado com sucesso!', 'success');
    } catch (err) {
      console.error(err);
      triggerToast('Ocorreu um erro ao gerar o pacote ZIP.', 'warn');
    }
  };

  // Administrative functions for Obras management
  const handleCreateObra = (name: string, idStr?: string) => {
    executeGuardedAction('admin', 'Criar Nova Instalação/Obra', () => {
      const trimmed = name.trim();
      const rawId = idStr ? idStr.trim() : "";
      const generatedId = `OB-${101 + obras.length}`;
      const finalId = rawId ? rawId.toUpperCase() : generatedId;

      if (!trimmed) {
        triggerToast("Nome da obra não pode estar vazio.", "warn");
        return;
      }
      if (obras.map(o => o.nome.toLowerCase()).includes(trimmed.toLowerCase())) {
        triggerToast("Esta obra já está cadastrada no sistema.", "warn");
        return;
      }
      if (obras.map(o => o.id.toLowerCase()).includes(finalId.toLowerCase())) {
        triggerToast("Este ID de obra já está cadastrado no sistema. Escolha outro.", "warn");
        return;
      }

      const newObraItem: Obra = { id: finalId, nome: trimmed };
      setObras([...obras, newObraItem]);
      setNewObraName('');
      setShowAddObraModal(false);
      triggerToast(`Obra "${trimmed}" (ID: ${finalId}) cadastrada com êxito!`, "success");
    });
  };

  const handleRenameObra = (oldName: string, newName: string, newId?: string) => {
    executeGuardedAction('admin', 'Editar Obra', () => {
      const trimmedNew = newName.trim();
      const trimmedId = newId ? newId.trim().toUpperCase() : '';
      if (!trimmedNew) {
        triggerToast("O novo nome não pode estar vazio.", "warn");
        return;
      }
      if (obras.map(o => o.nome.toLowerCase()).includes(trimmedNew.toLowerCase()) && trimmedNew.toLowerCase() !== oldName.toLowerCase()) {
        triggerToast("Este nome de obra já está em uso.", "warn");
        return;
      }

      const oldObraObj = obras.find(o => o.nome === oldName);
      if (trimmedId && oldObraObj && oldObraObj.id.toLowerCase() !== trimmedId.toLowerCase()) {
        if (obras.map(o => o.id.toLowerCase()).includes(trimmedId.toLowerCase())) {
          triggerToast("Este ID de obra já está cadastrado no sistema. Escolha outro.", "warn");
          return;
        }
      }
      
      const targetId = trimmedId || (oldObraObj ? oldObraObj.id : `OB-${101 + obras.length}`);

      // Update obras array
      setObras(obras.map(o => o.nome === oldName ? { ...o, nome: trimmedNew, id: targetId } : o));
      
      // Cascade update pedidos and baixas
      setPedidos(pedidos.map(p => p.obra === oldName ? { ...p, obra: trimmedNew, obraId: targetId } : p));
      setBaixas(baixas.map(b => b.obra === oldName ? { ...b, obra: trimmedNew, obraId: targetId } : b));
      
      // Keep UI filters aligned
      if (selectedObra === oldName) setSelectedObra(trimmedNew);
      if (newOrderObra === oldName) setNewOrderObra(trimmedNew);
      if (newWithdrawObra === oldName) setNewWithdrawObra(trimmedNew);
      
      setEditingObraName(null);
      triggerToast(`Obra editada para "${trimmedNew}" (ID: ${targetId})!`, "success");
    });
  };

  const handleDeleteObra = (obraName: string) => {
    executeGuardedAction('excluir_obra', `Excluir a Obra "${obraName}"`, () => {
      if (obras.length <= 1) {
        triggerToast("O sistema exige no mínimo uma obra ativa cadastrada.", "warn");
        return;
      }
      
      const updatedObras = obras.filter(o => o.nome !== obraName);
      setObras(updatedObras);
      
      if (selectedObra === obraName) setSelectedObra("Todas as Obras");
      if (newOrderObra === obraName) setNewOrderObra(updatedObras[0]?.nome || '');
      if (newWithdrawObra === obraName) setNewWithdrawObra(updatedObras[0]?.nome || '');
      
      triggerToast(`Obra "${obraName}" removida do sistema.`, "success");
    });
  };

  // Administrative functions for Pedidos configuration (CRUD)
  const handleDeletePedido = (id: string) => {
    executeGuardedAction('excluir_pedido', `Excluir o Pedido ${id}`, () => {
      setPedidos(pedidos.filter(p => p.id !== id));
      triggerToast(`Pedido ${id} removido definitivamente com sucesso!`, "success");
    });
  };

  const handleClearAllData = () => {
    executeGuardedAction('admin', 'Limpar Banco de Dados (Zerar Pedidos e Baixas)', () => {
      setPedidos([]);
      setBaixas([]);
      triggerToast('A Base de dados de pedidos e o histórico de baixas foram apagados! Pronto para importar Sienge!', 'success');
      setShowClearConfirmModal(false);
    });
  };

  const handleOpenEditPedidoModal = (p: Pedido) => {
    executeGuardedAction('admin', `Editar especificações do Pedido ${p.id}`, () => {
      setEditingPedido({ ...p });
      setShowEditPedidoModal(true);
    });
  };

  const handleSavePedidoEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingPedido) return;
    if (!editingPedido.insumo.trim()) {
      triggerToast("Nome do insumo é obrigatório.", "warn");
      return;
    }
    if (editingPedido.qtdSolicitada <= 0) {
      triggerToast("A quantidade solicitada deve ser maior que zero.", "warn");
      return;
    }

    setPedidos(pedidos.map(p => (p.id === editingPedido.id && p.codigo === editingPedido.codigo) ? {
      ...editingPedido,
      insumo: editingPedido.insumo.trim()
    } : p));
    setShowEditPedidoModal(false);
    setEditingPedido(null);
    triggerToast(`Pedido ${editingPedido.id} atualizado com êxito!`, "success");
  };

  // Administrative functions for Users configuration (CRUD & Passwords)
  const handleOpenEditUserModal = (u: Usuario) => {
    executeGuardedAction('admin', `Editar permissões do usuário ${u.nome}`, () => {
      setEditingUsuario({ ...u });
      setShowUserEditModal(true);
    });
  };

  const handleSaveUsuarioEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingUsuario) return;
    if (!editingUsuario.nome.trim() || !editingUsuario.login.trim() || !editingUsuario.senha.trim()) {
      triggerToast("Nome, Login e Senha são obrigatórios.", "warn");
      return;
    }
    const duplicate = usuarios.find(u => u.login.toLowerCase() === editingUsuario.login.toLowerCase() && u.id !== editingUsuario.id);
    if (duplicate) {
      triggerToast("Este login/e-mail já está em uso por outro colaborador.", "warn");
      return;
    }

    const updatedUsers = usuarios.map(u => u.id === editingUsuario.id ? editingUsuario : u);
    setUsuarios(updatedUsers);

    // If edited current operating simulator user, sync immediately
    if (editingUsuario.id === currentUser.id) {
      setCurrentUser(editingUsuario);
    }

    setShowUserEditModal(false);
    setEditingUsuario(null);
    triggerToast("Dados do usuário e permissões salvos com êxito!", "success");
  };

  const handleCreateUsuario = (name: string, loginStr: string, pass: string, roleType: 'Administrador' | 'Colaborador') => {
    executeGuardedAction('admin', 'Criar Novo Usuário de Gestão', () => {
      if (!name.trim() || !loginStr.trim() || !pass.trim()) {
        triggerToast("Preencha todos os campos obrigatórios.", "warn");
        return;
      }
      const duplicate = usuarios.find(u => u.login.toLowerCase() === loginStr.toLowerCase().trim());
      if (duplicate) {
        triggerToast("Este login de acesso já está em uso.", "warn");
        return;
      }

      const newUser: Usuario = {
        id: `usr-${Math.floor(Math.random() * 90000) + 10000}`,
        nome: name.trim(),
        login: loginStr.trim(),
        senha: pass.trim(),
        role: roleType,
        podeCriarPedido: roleType === 'Administrador',
        podeDarBaixa: true,
        podeReceberMercadoria: roleType === 'Administrador',
        podeExcluirPedido: roleType === 'Administrador',
        podeExcluirObra: roleType === 'Administrador'
      };

      setUsuarios([...usuarios, newUser]);
      setNewUserOpen(false);
      triggerToast(`Usuário "${name}" registrado no banco com senha ativa!`, "success");
    });
  };

  const handleDeleteUsuario = (id: string) => {
    executeGuardedAction('admin', 'Excluir Usuário', () => {
      if (id === currentUser.id) {
        triggerToast("Não é permitido excluir o usuário que está operando no momento.", "warn");
        return;
      }
      setUsuarios(usuarios.filter(u => u.id !== id));
      triggerToast("O cadastro do colaborador foi removido do sistema.", "success");
    });
  };

  const handleEditCredential = (userId: string, field: 'login' | 'senha', value: string) => {
    setUsuarios(prev => prev.map(u => {
      if (u.id === userId) {
        const updated = { ...u, [field]: value };
        if (currentUser.id === userId) {
          setCurrentUser(updated);
        }
        return updated;
      }
      return u;
    }));
  };

  const handleTogglePermission = (userId: string, field: keyof Usuario, checked: boolean) => {
    setUsuarios(prev => prev.map(u => {
      if (u.id === userId) {
        const updated = { ...u, [field]: checked };
        if (currentUser.id === userId) {
          setCurrentUser(updated);
        }
        return updated;
      }
      return u;
    }));
  };

  const handleSwitchSimulatorUser = (u: Usuario) => {
    setCurrentUser(u);
    // If the non-admin tab is selected but we switch away from admin tabs, make sure we redirect to panel
    if (u.role !== 'Administrador' && currentTab === 'admin') {
      setCurrentTab('painel');
    }
    triggerToast(`Sessão operante alterada para: ${u.nome} (${u.role})`, 'info');
  };

  const handleWithdrawObraChange = (val: string) => {
    setNewWithdrawObra(val);
    setNewWithdrawInsumo('');
  };

  // Export CSV of filtered unified movements (Diário de Movimentações)
  const handleExportCSV = () => {
    if (filteredMovements.length === 0) {
      triggerToast('Nenhum dado filtrado para exportação.', 'warn');
      return;
    }

    const headers = "Data/Hora,Tipo,Obra Destino,Insumo,Quantidade,Unidade,Documento/Justificativa,Responsavel\n";
    const csvRows = filteredMovements.map(m => {
      // Safe string escaping for commas or quotes
      const safeInsumo = `"${m.insumo.replace(/"/g, '""')}"`;
      const safeObra = `"${m.obra.replace(/"/g, '""')}"`;
      const safeDoc = `"${m.documento.replace(/"/g, '""')}"`;
      const safeResp = `"${m.responsavel.replace(/"/g, '""')}"`;
      return `${m.data},${m.tipo},${safeObra},${safeInsumo},${m.quantidade},${m.unidade},${safeDoc},${safeResp}`;
    });
    
    const csvContent = headers + csvRows.join("\n");
    
    // Inject UTF-8 BOM so Excel opens accented characters correctly
    const blob = new Blob(["\uFEFF" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `diario_movimentacoes_filtrado.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    triggerToast('Relatório exportado com sucesso.', 'success');
  };

  return (
    <div className="flex h-screen w-full bg-[#0F1115] font-sans text-slate-300 overflow-hidden select-none">
      
      {/* Toast Notification Bar */}
      {toast.message && (
        toast.message === 'PEDIDO RECEBIDO COM SUCESSO' ? (
          <div className="fixed inset-0 bg-black/70 backdrop-blur-xs flex items-center justify-center z-50 animate-fade-in">
            <div id="status-toast" className="bg-[#152e24] border border-emerald-500/80 text-emerald-300 px-8 py-6 rounded-2xl shadow-2xl flex flex-col items-center gap-4 text-center max-w-sm animate-bounce">
              <CheckCircle size={44} className="text-emerald-400" />
              <div className="flex flex-col gap-1.5">
                <span className="font-bold text-lg tracking-wider text-white">PEDIDO RECEBIDO COM SUCESSO</span>
                <span className="text-xs text-emerald-300/80">O recebimento do pedido de compra foi concluído e o saldo atualizado.</span>
              </div>
            </div>
          </div>
        ) : (
          <div id="status-toast" className={`fixed top-6 right-6 z-50 flex items-center gap-3 px-5 py-3 rounded-lg shadow-2xl border text-sm animate-bounce 
            ${toast.type === 'success' ? 'bg-[#152e24] border-emerald-500/40 text-emerald-300' : ''}
            ${toast.type === 'warn' ? 'bg-[#31231a] border-amber-500/40 text-amber-300' : ''}
            ${toast.type === 'info' ? 'bg-[#192330] border-blue-500/40 text-blue-300' : ''}
          `}>
            {toast.type === 'success' && <CheckCircle size={18} className="text-emerald-400" />}
            {toast.type === 'warn' && <AlertTriangle size={18} className="text-amber-400" />}
            {toast.type === 'info' && <Info size={18} className="text-blue-400" />}
            <span className="font-medium">{toast.message}</span>
          </div>
        )
      )}

      {/* Sidebar Navigation */}
      <aside id="sidebar-panel" className={`transition-all duration-300 ${isSidebarCollapsed ? 'w-20' : 'w-[260px]'} bg-[#161920] border-r border-slate-800 flex flex-col justify-between shrink-0`}>
        <div className={isSidebarCollapsed ? 'p-3' : 'p-6'}>
          {/* Logo / Branding */}
          {!isSidebarCollapsed ? (
            <div className="flex items-center justify-between gap-2.5 mb-6">
              <div className="flex items-center gap-3">
                <LogoIcon className="w-12 h-12 rounded-full shadow-lg shadow-yellow-500/20 shrink-0 transition-transform duration-300 hover:scale-105" />
                <div>
                  <h1 className="text-base font-sans font-black tracking-tight text-white uppercase leading-none">Gestão de <span className="text-[#FFC800]">Estoque</span></h1>
                  <p className="text-[10px] text-yellow-505 font-mono tracking-widest uppercase font-extrabold mt-1">ABSOLUTA CONSTRUTORA</p>
                </div>
              </div>
              <button 
                type="button" 
                onClick={() => setIsSidebarCollapsed(true)}
                className="p-1.5 rounded bg-slate-850 hover:bg-slate-800 text-slate-400 hover:text-white border border-slate-800 transition-colors shrink-0"
                title="Recolher menu"
              >
                <ChevronLeft size={16} />
              </button>
            </div>
          ) : (
            <div className="flex flex-col items-center gap-4 mb-6">
              <LogoIcon className="w-11 h-11 rounded-full shadow-lg shadow-yellow-500/20 transition-transform duration-300 hover:scale-105" />
              <button 
                type="button" 
                onClick={() => setIsSidebarCollapsed(false)}
                className="p-1.5 rounded-full bg-slate-800 hover:bg-emerald-650 text-slate-300 hover:text-white border border-slate-705 transition-all shadow-md active:scale-95"
                title="Expandir menu"
              >
                <ChevronRight size={16} />
              </button>
            </div>
          )}

          {/* User Simulator Switcher Box */}
          {!isSidebarCollapsed ? (
            <div className="mb-6 p-3 bg-[#0F1115] border border-slate-800 rounded-lg">
              <div className="flex justify-between items-center mb-1.5">
                <span className="text-[10px] font-mono uppercase tracking-wider text-slate-500 flex items-center gap-1">
                  <Shield size={10} className="text-yellow-500" />
                  Operando Como:
                </span>
                <span className={`text-[9px] font-bold uppercase px-1 rounded ${
                  currentUser.role === 'Administrador' ? 'bg-purple-500/20 text-purple-400' : 'bg-blue-500/20 text-blue-400'
                }`}>
                  {currentUser.role}
                </span>
              </div>
              
              <div className="relative">
                <select
                  value={currentUser.id}
                  onChange={(e) => {
                    const selectedId = e.target.value;
                    const matchedUser = usuarios.find(u => u.id === selectedId);
                    if (matchedUser) {
                      handleSwitchSimulatorUser(matchedUser);
                    }
                  }}
                  className="w-full bg-[#161920] border border-slate-700 text-xs text-white rounded p-1.5 focus:outline-none focus:border-yellow-500 font-medium cursor-pointer"
                >
                  {usuarios.map(u => (
                    <option key={u.id} value={u.id}>
                      {u.nome} ({u.role === 'Administrador' ? 'Admin' : 'Colaborador'})
                    </option>
                  ))}
                </select>
              </div>
              <p className="text-[9px] text-slate-500 mt-1.5 leading-tight italic">
                Altere o perfil para testar as travas de ações e as permissões de acesso.
              </p>
            </div>
          ) : (
            <div className="mb-6 flex justify-center">
              <div className="relative group/user">
                <div 
                  className={`w-9 h-9 rounded-full border flex items-center justify-center cursor-pointer transition-all ${
                    currentUser.role === 'Administrador' ? 'bg-purple-950/40 border-purple-500/30 text-purple-400' : 'bg-blue-950/40 border-blue-500/30 text-blue-400'
                  }`}
                  title={`Operando Como: ${currentUser.nome} (${currentUser.role})`}
                >
                  <Shield size={14} />
                </div>
                <div className="invisible group-hover/user:visible absolute left-12 top-0 z-50 bg-[#161920] border border-slate-800 p-2.5 rounded-lg shadow-xl w-48 transition-all">
                  <p className="text-[9px] text-slate-500 font-mono uppercase mb-1.5">Perfil rápido:</p>
                  <select
                    value={currentUser.id}
                    onChange={(e) => {
                      const selectedId = e.target.value;
                      const matchedUser = usuarios.find(u => u.id === selectedId);
                      if (matchedUser) {
                        handleSwitchSimulatorUser(matchedUser);
                      }
                    }}
                    className="w-full bg-[#0F1115] border border-slate-700 text-xs text-white rounded p-1"
                  >
                    {usuarios.map(u => (
                      <option key={u.id} value={u.id}>
                        {u.nome}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          )}

          {/* Navigation Links */}
          <nav className="space-y-1.5 w-full" id="navigation-root">
            <button
              id="nav-tab-painel"
              onClick={() => setCurrentTab('painel')}
              className={`w-full flex items-center ${isSidebarCollapsed ? 'justify-center p-2.5 relative' : 'justify-between px-4 py-2.5'} rounded-md text-sm font-medium transition-all ${
                currentTab === 'painel'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
              title="Painel Geral"
            >
              <span className="flex items-center gap-3">
                <LayoutDashboard size={18} className="shrink-0" />
                {!isSidebarCollapsed && <span>Painel Geral</span>}
              </span>
              {!isSidebarCollapsed && currentTab === 'painel' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
              {isSidebarCollapsed && currentTab === 'painel' && <span className="absolute left-1 w-1 h-6 bg-emerald-400 rounded-r-md"></span>}
            </button>

            <button
              id="nav-tab-pedidos"
              onClick={() => {
                setCurrentTab('pedidos');
                setStatusFilter(['Pendente']);
              }}
              className={`w-full flex items-center ${isSidebarCollapsed ? 'justify-center p-2.5 relative' : 'justify-between px-4 py-2.5'} rounded-md text-sm font-medium transition-all ${
                currentTab === 'pedidos'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
              title="Pedidos de Compra"
            >
              <span className="flex items-center gap-3">
                <ClipboardList size={18} className="shrink-0" />
                {!isSidebarCollapsed && <span>Pedidos de Compra</span>}
              </span>
              {!isSidebarCollapsed && currentTab === 'pedidos' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
              {isSidebarCollapsed && currentTab === 'pedidos' && <span className="absolute left-1 w-1 h-6 bg-emerald-400 rounded-r-md"></span>}
            </button>

            <button
              id="nav-tab-estoque"
              onClick={() => setCurrentTab('estoque')}
              className={`w-full flex items-center ${isSidebarCollapsed ? 'justify-center p-2.5 relative' : 'justify-between px-4 py-2.5'} rounded-md text-sm font-medium transition-all ${
                currentTab === 'estoque'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
              title="Estoque por Obra"
            >
              <span className="flex items-center gap-3">
                <BaggageClaim size={18} className="shrink-0" />
                {!isSidebarCollapsed && <span>Estoque por Obra</span>}
              </span>
              {!isSidebarCollapsed && currentTab === 'estoque' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
              {isSidebarCollapsed && currentTab === 'estoque' && <span className="absolute left-1 w-1 h-6 bg-emerald-400 rounded-r-md"></span>}
            </button>

            <button
              id="nav-tab-relatorios"
              onClick={() => setCurrentTab('relatorios')}
              className={`w-full flex items-center ${isSidebarCollapsed ? 'justify-center p-2.5 relative' : 'justify-between px-4 py-2.5'} rounded-md text-sm font-medium transition-all ${
                currentTab === 'relatorios'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
              title="Relatórios e Baixas"
            >
              <span className="flex items-center gap-3">
                <BarChart3 size={18} className="shrink-0" />
                {!isSidebarCollapsed && <span>Relatórios e Baixas</span>}
              </span>
              {!isSidebarCollapsed && currentTab === 'relatorios' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
              {isSidebarCollapsed && currentTab === 'relatorios' && <span className="absolute left-1 w-1 h-6 bg-emerald-400 rounded-r-md"></span>}
            </button>

            <button
              id="nav-tab-admin"
              onClick={() => {
                executeGuardedAction('admin', 'Acessar painel administrativo', () => {
                  setCurrentTab('admin');
                });
              }}
              className={`w-full flex items-center ${isSidebarCollapsed ? 'justify-center p-2.5 relative' : 'justify-between px-4 py-2.5'} rounded-md text-sm font-medium transition-all ${
                currentTab === 'admin'
                  ? 'bg-purple-600/10 text-purple-400 border border-purple-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
              title="Administração"
            >
              <span className="flex items-center gap-3">
                <Shield size={18} className={`shrink-0 ${currentUser.role === 'Administrador' ? 'text-purple-400' : 'text-slate-500'}`} />
                {!isSidebarCollapsed && <span>Administração</span>}
              </span>
              {!isSidebarCollapsed ? (
                currentUser.role !== 'Administrador' ? (
                  <Lock size={12} className="text-slate-500 shrink-0" />
                ) : currentTab === 'admin' ? (
                  <span className="w-1.5 h-1.5 bg-purple-400 rounded-full shrink-0"></span>
                ) : null
              ) : (
                currentTab === 'admin' && <span className="absolute left-1 w-1 h-6 bg-purple-400 rounded-r-md"></span>
              )}
            </button>
          </nav>
          
          {/* Logo Corporativo sem Fundo */}
          <CorporateLogoWide collapsed={isSidebarCollapsed} />
        </div>

        {/* User profile / Footer */}
        <div id="sidebar-footer" className={`border-t border-slate-800 bg-[#12151B] transition-all duration-300 ${isSidebarCollapsed ? 'p-3 flex justify-center' : 'p-5'}`}>
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center text-xs font-bold text-yellow-500 uppercase shrink-0" title={currentUser.nome}>
              {currentUser.nome.substring(0, 2)}
            </div>
            {!isSidebarCollapsed && (
              <div className="flex-1 min-w-0">
                <p className="text-xs font-semibold text-white truncate">{currentUser.nome}</p>
                <p className="text-[10px] text-slate-500 truncate">{currentUser.login}</p>
              </div>
            )}
          </div>
        </div>
      </aside>

      {/* Main Container */}
      <main className="flex-1 flex flex-col min-w-0" id="main-content">
        
        {/* Top Header */}
        <header id="main-header" className="h-16 bg-[#161920]/50 border-b border-slate-800 px-8 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-4">
            <div className="relative">
              <select
                id="header-obra-select"
                value={selectedObra}
                onChange={(e) => setSelectedObra(e.target.value)}
                className="bg-[#1F242D] border border-slate-700 text-sm rounded-lg pl-3 pr-8 py-1.5 text-white font-medium focus:outline-none focus:ring-1 focus:ring-emerald-500 appearance-none cursor-pointer"
              >
                <option value="Todas as Obras">Todas as Obras</option>
                {listObras.map((o) => (
                  <option key={o} value={o}>{o}</option>
                ))}
              </select>
              <div className="absolute right-2.5 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400">
                <ChevronRight size={14} className="rotate-90" />
              </div>
            </div>
            
            <span className="text-slate-700 font-sans italic">|</span>
            <div className="text-[10px] text-slate-500 font-medium uppercase tracking-widest hidden sm:block">
              Base atualizada em {lastImportDate} • Movimentações: {movementsCount}
            </div>
          </div>

          <div className="flex gap-2.5">
            {currentTab !== 'pedidos' && currentTab !== 'painel' && (
              <>
                <button
                  id="btn-baixa-material"
                  onClick={() => {
                    // Prepopulate current work selection if valid
                    if (listObras.includes(selectedObra)) {
                      setNewWithdrawObra(selectedObra);
                    } else {
                      setNewWithdrawObra('Residencial Alvorada');
                    }
                    setNewWithdrawInsumo('');
                    setNewWithdrawQtd(10);
                    setShowWithdrawModal(true);
                  }}
                  className="px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-white rounded-md text-sm font-medium border border-slate-700 transition-all flex items-center gap-1.5"
                >
                  <ArrowRightLeft size={15} className="text-slate-300" />
                  Baixa de Material
                </button>
                
                <button
                  id="btn-nova-entrada"
                  onClick={() => {
                    if (listObras.includes(selectedObra)) {
                      setNewOrderObra(selectedObra);
                    }
                    setShowOrderModal(true);
                  }}
                  className="px-3.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-md text-sm font-medium transition-all flex items-center gap-1.5 shadow-md shadow-emerald-900/10"
                >
                  <Plus size={16} />
                  Nova Entrada
                </button>
              </>
            )}
          </div>
        </header>

        {/* Central Content Canvas */}
        <div className="flex-1 p-8 overflow-y-auto" id="view-canvas-container">
          
          {/* Welcome Dashboard Notification Banner */}
          <div className="mb-6 p-4 bg-gradient-to-r from-emerald-950/30 to-[#161920] border border-emerald-500/10 rounded-xl flex items-start gap-3">
            <div className="p-2 bg-emerald-500/10 rounded-lg text-emerald-400">
              <Info size={16} />
            </div>
            <div>
              <h3 className="text-xs font-semibold text-emerald-400">Gestão Integrada de Obras Ativas</h3>
              <p className="text-xs text-slate-400 mt-0.5">
                Utilize o menu superior para filtrar estatísticas e acompanhar as entregas locais em tempo real. Adicione novos pedidos de insumos clicando em <strong>Nova Entrada</strong>, ou gaste o estoque com <strong>Baixa de Material</strong>.
              </p>
            </div>
          </div>

          {/* MAIN PAGE VIEW: PAINEL GERAL */}
          {currentTab === 'painel' && (
            <div className="space-y-6" id="dashboard-view-pane">
              
              {/* Stats Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4" id="stats-summary-grid">
                
                <div className="bg-[#161920] border border-slate-800 p-5 rounded-xl flex justify-between items-center relative overflow-hidden group hover:border-slate-700/55 transition-all">
                  <div>
                    <p className="text-slate-500 text-[10px] font-semibold uppercase tracking-wider mb-1">Pedidos Pendentes</p>
                    <p className="text-3xl md:text-3xl lg:text-4xl font-sans font-bold text-amber-500 transition-all tracking-tight" id="stat-pedidos-pendentes">{stats.pendentes}</p>
                    <span className="text-[10px] text-slate-500 mt-1 block">Aguardando transporte</span>
                  </div>
                  <div className="p-3 bg-amber-500/5 text-amber-500 rounded-lg group-hover:scale-105 transition-all">
                    <Clock size={20} />
                  </div>
                  <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-amber-500/20"></div>
                </div>

                <div className="bg-[#161920] border border-slate-800 p-5 rounded-xl flex justify-between items-center relative overflow-hidden group hover:border-slate-700/55 transition-all">
                  <div>
                    <p className="text-slate-500 text-[10px] font-semibold uppercase tracking-wider mb-1">Entregas Parciais</p>
                    <p className="text-3xl md:text-3xl lg:text-4xl font-sans font-bold text-blue-400 tracking-tight" id="stat-pedidos-parciais">{stats.parciais}</p>
                    <span className="text-[10px] text-slate-500 mt-1 block">Recebidos em partes</span>
                  </div>
                  <div className="p-3 bg-blue-500/5 text-blue-400 rounded-lg group-hover:scale-105 transition-all">
                    <Layers size={21} />
                  </div>
                  <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-blue-500/20"></div>
                </div>

                <div className="bg-[#161920] border border-slate-800 p-5 rounded-xl flex justify-between items-center relative overflow-hidden group hover:border-slate-700/55 transition-all">
                  <div>
                    <p className="text-slate-500 text-[10px] font-semibold uppercase tracking-wider mb-1">Alerta: Estoque Crítico</p>
                    <p className="text-3xl md:text-3xl lg:text-4xl font-sans font-bold text-rose-500 tracking-tight" id="stat-estoque-critico">{stats.estoqueCritico}</p>
                    <span className="text-[10px] text-rose-500/70 mt-1 block">Estoque local no limite</span>
                  </div>
                  <div className="p-3 bg-rose-500/5 text-rose-500 rounded-lg group-hover:scale-105 transition-all animate-pulse">
                    <AlertTriangle size={20} />
                  </div>
                  <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-rose-500/20"></div>
                </div>

                <div className="bg-[#161920] border border-slate-800 p-5 rounded-xl flex justify-between items-center relative overflow-hidden group hover:border-slate-700/55 transition-all">
                  <div>
                    <p className="text-slate-500 text-[10px] font-semibold uppercase tracking-wider mb-1">Movimentações Hoje</p>
                    <p className="text-3xl md:text-3xl lg:text-4xl font-sans font-bold text-emerald-400 tracking-tight" id="stat-movimentos-hoje">{movementsCount}</p>
                    <span className="text-[10px] text-slate-500 mt-1 block">Atividades registradas</span>
                  </div>
                  <div className="p-3 bg-emerald-50/5 text-emerald-400 rounded-lg group-hover:scale-105 transition-all">
                    <TrendingDown size={21} className="rotate-180" />
                  </div>
                  <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-emerald-500/20"></div>
                </div>
              </div>

              {/* Main Dashboard Advanced Visual Statistics & Critical Inventory */}
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6" id="dashboard-statistics-container">
                
                {/* Left Card: Pedidos Statistics Chart & KPI */}
                <div className="lg:col-span-6 bg-[#161920] border border-slate-800 rounded-xl p-6 flex flex-col justify-between" id="chart-pedidos-stats">
                  <div>
                    <div className="flex justify-between items-start mb-6">
                      <div>
                        <h2 className="text-base font-sans font-bold text-white uppercase tracking-tight flex items-center gap-2">
                          <BarChart3 size={16} className="text-emerald-400" />
                          Estatísticas de Pedidos
                        </h2>
                        <p className="text-xs text-slate-500 mt-0.5">Métricas de entrega da obra: <strong className="text-emerald-400">{selectedObra}</strong></p>
                      </div>
                      <span className="text-[10px] bg-slate-800 text-slate-300 font-mono px-2 py-0.5 rounded-full">
                        {selectedObra === 'Todas as Obras' ? 'Painel Geral' : 'Obra Selecionada'}
                      </span>
                    </div>

                    {/* Chart Render */}
                    <div className="w-full bg-[#0F1115] border border-slate-800/80 rounded-lg p-4 mb-5" style={{ minHeight: '260px' }}>
                      {dashboardChartData.every(item => item.value === 0) ? (
                        <div className="h-[230px] flex flex-col items-center justify-center text-slate-500">
                          <ClipboardList size={32} className="text-slate-700 mb-2" />
                          <p className="text-xs">Nenhum pedido cadastrado para esta obra.</p>
                        </div>
                      ) : (
                        <div style={{ width: '100%', height: 230 }}>
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dashboardChartData} margin={{ top: 15, right: 10, left: -20, bottom: 5 }}>
                              <XAxis 
                                dataKey="name" 
                                stroke="#64748b" 
                                fontSize={10} 
                                tickLine={false} 
                                axisLine={false} 
                              />
                              <YAxis 
                                stroke="#64748b" 
                                fontSize={10} 
                                tickLine={false} 
                                axisLine={false} 
                                allowDecimals={false} 
                              />
                              <RechartsTooltip 
                                contentStyle={{ backgroundColor: '#0f172a', borderColor: '#334155', borderRadius: '6px' }} 
                                itemStyle={{ color: '#f8fafc', fontSize: '11px' }}
                                labelStyle={{ color: '#94a3b8', fontWeight: 'bold', fontSize: '10px' }}
                              />
                              <Bar dataKey="value" stroke="none" radius={[4, 4, 0, 0]} maxBarSize={45}>
                                {dashboardChartData.map((entry, index) => (
                                  <Cell key={`cell-${index}`} fill={entry.fill} />
                                ))}
                              </Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Pertinent custom KPI summary: Supplies Efficiency Rate */}
                  <div className="bg-[#1C2028] border border-slate-800 p-4.5 rounded-lg flex flex-col sm:flex-row items-center justify-between gap-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-emerald-500/10 text-emerald-400 rounded-lg">
                        <CheckCircle size={18} />
                      </div>
                      <div>
                        <h4 className="text-xs font-bold text-slate-200">Eficiência de Atendimento</h4>
                        <p className="text-[10px] text-slate-500 mt-0.5">Proporção ponderada de pedidos atendidos, entregas e andamento.</p>
                      </div>
                    </div>
                    <div className="text-right flex flex-row sm:flex-col items-center sm:items-end gap-2 sm:gap-0">
                      <span className="text-2xl font-black font-sans text-emerald-400">
                        {(() => {
                          const totEntregas = dashboardChartData[0]?.value || 0;
                          const totParciais = dashboardChartData[1]?.value || 0;
                          const totPendentes = dashboardChartData[2]?.value || 0;
                          const total = totEntregas + totParciais + totPendentes;
                          if (total === 0) return '100%';
                          return `${Math.round(((totEntregas * 100) + (totParciais * 50)) / total)}%`;
                        })()}
                      </span>
                      <span className="text-[10.5px] text-slate-400 font-mono">Índice Conclusão</span>
                    </div>
                  </div>
                </div>

                {/* Right Card: Real Obras Received Material Progress Chart */}
                <div className="lg:col-span-6 bg-[#161920] border border-slate-800 rounded-xl p-6 flex flex-col justify-between" id="chart-obras-recebimento">
                  <div>
                    <div className="flex justify-between items-start mb-6">
                      <div>
                        <h2 className="text-base font-sans font-bold text-white uppercase tracking-tight flex items-center gap-2">
                          <Layers size={16} className="text-blue-400" />
                          Recebimento por Obra (Real)
                        </h2>
                        <p className="text-xs text-slate-500 mt-0.5">Progresso ponderado de materiais recebidos págos (Obras Importadas)</p>
                      </div>
                      <span className="text-[10px] bg-slate-800 text-slate-300 font-mono px-2 py-0.5 rounded-full">
                        Estatística Dinâmica
                      </span>
                    </div>

                    {/* Dynamic Real Obras Bar Chart */}
                    <div className="space-y-4 pt-1" id="real-works-received-statistics" style={{ minHeight: '260px' }}>
                      {obrasRecebimentoStats.length === 0 ? (
                        <div className="h-[230px] flex flex-col items-center justify-center text-slate-500 bg-[#0F1115] border border-slate-800 rounded-lg">
                          <ClipboardList size={32} className="text-slate-700 mb-2" />
                          <p className="text-xs">Não há dados de obras reais cadastrados.</p>
                        </div>
                      ) : (
                        <div className="space-y-3.5 max-h-[250px] overflow-y-auto pr-1">
                          {obrasRecebimentoStats.map((item) => {
                            const percent = item.percent;
                            const barColor = percent >= 75 ? 'bg-emerald-500' : percent >= 45 ? 'bg-amber-500' : 'bg-blue-500';
                            return (
                              <div key={item.name} className="space-y-1.5 p-2.5 rounded-lg bg-[#0F1115] border border-slate-800/80 hover:border-slate-700 transition-all">
                                <div className="flex justify-between text-xs font-sans text-slate-300">
                                  <span className="font-extrabold truncate max-w-[170px] sm:max-w-[220px]">{item.name}</span>
                                  <span className="font-bold shrink-0">{percent}% Recebido ({item.recebido} un)</span>
                                </div>
                                <div className="w-full bg-slate-950 rounded-xl h-2 overflow-hidden border border-slate-800/50">
                                  <div className={`h-full ${barColor} rounded-xl transition-all duration-500`} style={{ width: `${percent}%` }}></div>
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Footer Detail Card */}
                  <div className="bg-[#1C2028] border border-slate-800 p-4.5 rounded-lg flex flex-col sm:flex-row items-center justify-between gap-4 mt-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2.5 bg-blue-500/10 text-blue-400 rounded-lg">
                        <MapPin size={18} />
                      </div>
                      <div>
                        <h4 className="text-xs font-bold text-slate-200">Total de Instalações</h4>
                        <p className="text-[10px] text-slate-500 mt-0.5">Canteiros rastreados no sistema de almoxarifado local.</p>
                      </div>
                    </div>
                    <div className="text-right flex flex-row sm:flex-col items-center sm:items-end gap-2 sm:gap-0">
                      <span className="text-2.5xl font-black font-sans text-blue-400">{obras.length}</span>
                      <span className="text-[10.5px] text-slate-400 font-mono">Obras Ativas</span>
                    </div>
                  </div>
                </div>

                {/* Row 2: Critical Stock List (<15%) spanning 12 columns */}
                <div className="lg:col-span-12 bg-[#161920] border border-slate-800 rounded-xl p-6 flex flex-col" id="chart-insumos-criticos">
                  <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800">
                    <h2 className="text-base font-sans font-bold text-white uppercase tracking-tight flex items-center gap-2">
                      <AlertTriangle size={16} className="text-rose-500" />
                      Alerta: Estoque Crítico
                    </h2>
                    <span className="text-[10px] text-slate-500 font-mono font-bold uppercase tracking-wider">Saldo &lt; 15%</span>
                  </div>

                  <p className="text-xs text-slate-400 mb-4 leading-relaxed">
                    Lista de insumos ativos recebidos que estão com saldo de estoque baixo em relação ao fornecido, necessitando de reabastecimento.
                  </p>

                  {dashboardCriticalStockList.length === 0 ? (
                    <div className="h-full min-h-[140px] flex flex-col items-center justify-center text-center p-6 bg-[#0F1115] border border-dashed border-slate-800 rounded-lg">
                      <div className="w-10 h-10 rounded-full bg-emerald-500/10 text-emerald-400 flex items-center justify-center mb-3">
                        <CheckCircle size={20} />
                      </div>
                      <p className="text-xs font-bold text-white">Nenhum Insumo Crítico</p>
                      <p className="text-[10.5px] text-slate-500 mt-1 max-w-[400px]">Todos os insumos possuem saldo estável ou acima de 15% do recebido.</p>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4 overflow-y-auto max-h-[350px]">
                      {dashboardCriticalStockList.map((item) => {
                        const ratio = item.recebido > 0 ? (item.saldo / item.recebido) : 0;
                        const percent = Math.max(0, Math.min(100, Math.round(ratio * 100)));
                        
                        return (
                          <div key={`${item.obra}:::${item.codigo}:::${item.insumo}`} className="p-3 bg-[#0F1115] border border-slate-800 rounded-lg hover:border-slate-700 transition-all flex flex-col justify-between gap-2.5 group notranslate" translate="no">
                            <div className="flex justify-between items-start">
                              <div className="max-w-[70%]">
                                <h4 className="text-xs font-bold text-white truncate group-hover:text-purple-300 transition-colors">{item.insumo}</h4>
                                <div className="flex gap-1.5 mt-0.5 items-center">
                                  <span className="text-[9px] font-mono text-slate-500">COD: {item.codigo || 'S/C'}</span>
                                  {selectedObra === 'Todas as Obras' && (
                                    <>
                                      <span className="text-[#3b82f6]/40">•</span>
                                      <span className="text-[9.5px] text-slate-450 font-semibold truncate max-w-[120px]">{item.obra}</span>
                                    </>
                                  )}
                                </div>
                              </div>
                              <div className="text-right">
                                <span className="text-xs font-bold font-mono text-rose-500">{item.saldo}</span>
                                <span className="text-[10px] text-slate-500"> / {item.recebido} {item.unidade}</span>
                              </div>
                            </div>

                            {/* visual progress gauge */}
                            <div className="w-full">
                              <div className="flex justify-between text-[9px] text-slate-500 font-mono mb-1">
                                <span>Percentual Disponível</span>
                                <span className="font-bold text-rose-400">{percent}%</span>
                              </div>
                              <div className="w-full bg-slate-900 rounded-full h-1.5 overflow-hidden">
                                <div 
                                  className="bg-gradient-to-r from-red-650 to-rose-500 h-full rounded-full" 
                                  style={{ width: `${percent}%` }}
                                ></div>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                  
                  {/* Insight and Recommendations */}
                  <div className="mt-4 p-3 bg-rose-500/5 border border-rose-500/10 rounded-lg text-[11px] text-rose-400 flex items-start gap-2">
                    <AlertCircle size={14} className="shrink-0 mt-0.5" />
                    <span>
                      {dashboardCriticalStockList.length > 0 
                        ? `Atenção: É recomendado solicitar novas ordens de compras no Sienge para os ${dashboardCriticalStockList.length} insumos de estoque críticos.`
                        : 'Recomendação: Estoques regulares. Monitore diários para antecipar consumo de insumos pesados.'}
                    </span>
                  </div>
                </div>

              </div>
            </div>
          )}

          {/* VIEW TAB 2: PEDIDOS DE COMPRA DETALHADO */}
          {currentTab === 'pedidos' && (
            <div className="space-y-6" id="pedidos-tab-view">
              <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4">
                <div>
                  <h2 className="text-xl font-sans font-extrabold text-white">Gestão Integrada de Pedidos</h2>
                  <p className="text-sm text-slate-400 mt-1">Acompanhe, filtre e configure o status das compras ativas de materiais e ferragens para as obras.</p>
                </div>
                <div className="flex flex-wrap gap-2.5">
                  <button
                    onClick={() => setShowOrderModal(true)}
                    className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded-md text-sm font-semibold transition-all flex items-center gap-2"
                  >
                    <Plus size={16} /> Solicitando Novo Pedido
                  </button>
                  <button
                    onClick={() => {
                      setParsedImportItems([]);
                      setImportedFileName('');
                      setImportStats(null);
                      setShowImportModal(true);
                    }}
                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-705 rounded-md text-sm font-semibold transition-all flex items-center gap-2"
                  >
                    <FileSpreadsheet size={16} className="text-emerald-450" />
                    Importar Sienge Sincronizado
                  </button>
                  <button
                    onClick={() => {
                      executeGuardedAction('admin', 'Zerar Banco de Dados de Pedidos e Baixas', () => {
                        setShowClearConfirmModal(true);
                      });
                    }}
                    className="px-4 py-2 bg-rose-500/10 hover:bg-rose-600 text-rose-400 hover:text-white border border-rose-550/25 rounded-md text-sm font-semibold transition-all flex items-center gap-2"
                    title="Excluir todos os pedidos para carregar uma nova planilha limpa"
                  >
                    <Trash2 size={16} />
                    Limpar Base (Zerar)
                  </button>
                </div>
              </div>

              {/* Status breakdown metrics */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl p-6 flex flex-col md:flex-row justify-around gap-6 divide-y md:divide-y-0 md:divide-x divide-slate-800">
                <div className="flex-1 pb-4 md:pb-0 md:pr-4 text-center">
                  <p className="text-[11px] text-slate-400 uppercase font-bold tracking-wider mb-1">Total Solicitados</p>
                  <p className="text-2xl font-sans font-bold text-white tracking-tight">{pedidos.length} Pedidos</p>
                </div>
                <div className="flex-1 py-4 md:py-0 md:px-4 text-center">
                  <p className="text-[11px] text-amber-400 uppercase font-bold tracking-wider mb-1">Pendentes de Entrega</p>
                  <p className="text-2xl font-sans font-bold text-amber-500 tracking-tight">
                    {pedidos.filter(p => p.status === 'Pendente').length} Pedidos
                  </p>
                </div>
                <div className="flex-1 py-4 md:py-0 md:px-4 text-center">
                  <p className="text-[11px] text-blue-400 uppercase font-bold tracking-wider mb-1">Entregas Parciais</p>
                  <p className="text-2xl font-sans font-bold text-blue-400 tracking-tight">
                    {pedidos.filter(p => p.status === 'Parcial').length} Pedidos
                  </p>
                </div>
                <div className="flex-1 pt-4 md:pt-0 md:pl-4 text-center">
                  <p className="text-[11px] text-emerald-400 uppercase font-bold tracking-wider mb-1 font-semibold">Entregue Totalmente</p>
                  <p className="text-2xl font-sans font-bold text-emerald-400 tracking-tight">
                    {pedidos.filter(p => p.status === 'Entregue').length} Pedidos
                  </p>
                </div>
              </div>

              {/* Search and Filters Header */}
              <div className="bg-[#161920] border border-slate-800 p-4.5 rounded-xl flex flex-col md:flex-row gap-3 justify-between items-center bg-[#1C2028]">
                <div className="relative w-full md:w-80">
                  <input
                    type="text"
                    placeholder="Pesquisar pedidos (insumo, código, obra)..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full bg-[#0F1115]/95 border border-slate-700 rounded-lg py-1.5 pl-9 pr-3 text-xs text-white placeholder:text-slate-500 focus:outline-none focus:border-emerald-500 transition-colors"
                  />
                  <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-500" />
                </div>

                <div className="flex flex-col sm:flex-row gap-3 w-full md:w-auto justify-end items-center">
                  <div className="flex bg-[#0F1115] p-1 rounded-md border border-slate-700 text-xs text-slate-400">
                    {['Todos', 'Pendente', 'Parcial', 'Entregue'].map((st) => {
                      const isActive = statusFilter.includes(st);
                      return (
                        <button
                          key={st}
                          onClick={() => handleStatusFilterToggle(st)}
                          className={`px-2.5 py-1 rounded transition-colors ${
                            isActive ? 'bg-emerald-600/20 text-emerald-400 font-semibold' : 'hover:text-slate-200'
                          }`}
                          title={st === 'Todos' ? 'Mostrar todos os status' : `Filtrar por ${st} (Clique para selecionar múltiplos)`}
                        >
                          {st}
                        </button>
                      );
                    })}
                  </div>

                  <div className="flex bg-[#0F1115] p-1 rounded-md border border-slate-700 text-xs text-slate-400">
                    <button
                      type="button"
                      onClick={() => setPedidosViewMode('cards')}
                      className={`px-2.5 py-1 rounded transition-colors flex items-center gap-1.5 ${
                        pedidosViewMode === 'cards' ? 'bg-emerald-600/20 text-emerald-400 font-semibold' : 'hover:text-slate-200'
                      }`}
                      title="Visualização em Colunas"
                    >
                      <LayoutGrid size={13} />
                      <span className="text-[11px] font-medium">Colunas</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setPedidosViewMode('table')}
                      className={`px-2.5 py-1 rounded transition-colors flex items-center gap-1.5 ${
                        pedidosViewMode === 'table' ? 'bg-emerald-600/20 text-emerald-400 font-semibold' : 'hover:text-slate-200'
                      }`}
                      title="Visualização em Lista / Linhas"
                    >
                      <List size={13} />
                      <span className="text-[11px] font-medium">Listado</span>
                    </button>
                  </div>
                </div>
              </div>

              {/* Advanced Orders List - Dual view (Cards or Table listado) */}
              {pedidosViewMode === 'table' ? (
                <div className="overflow-x-auto border border-slate-800 rounded-xl bg-[#161920]">
                  <table className="w-full text-left text-xs text-slate-300 border-collapse">
                    <thead className="bg-[#1C2028] text-slate-400 text-[10px] uppercase font-sans font-semibold tracking-wider border-b border-slate-800">
                      <tr>
                        {/* PEDIDO DE COMPRA / EMISSÃO */}
                        <th 
                          className="px-4 py-3.5 min-w-[140px] border-r border-slate-850 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleSortToggle('id_numeric')}
                          title="Clique para ordenar por Pedido de Compra"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Pedido de Compra / Emissão</span>
                            <span className="inline-flex shrink-0">
                              {sortConfig.key === 'id_numeric' ? (
                                sortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* C. CUSTO / OBRA */}
                        <th 
                          className="px-4 py-3.5 min-w-[140px] border-r border-slate-850 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleSortToggle('obra')}
                          title="Clique para ordenar por Obra"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Obra / C.C</span>
                            <span className="inline-flex shrink-0">
                              {sortConfig.key === 'obra' ? (
                                sortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* FORNECEDOR */}
                        <th 
                          className="px-4 py-3.5 min-w-[140px] border-r border-slate-850 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleSortToggle('fornecedor')}
                          title="Clique para ordenar por Fornecedor"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Fornecedor / Comprador</span>
                            <span className="inline-flex shrink-0">
                              {sortConfig.key === 'fornecedor' ? (
                                sortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* INSUMO SIENGE */}
                        <th 
                          className="px-4 py-3.5 min-w-[160px] border-r border-slate-850 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleSortToggle('insumo')}
                          title="Clique para ordenar por Insumo"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Insumo / Descrição</span>
                            <span className="inline-flex shrink-0">
                              {sortConfig.key === 'insumo' ? (
                                sortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* PEDIDO / RECEBIDO */}
                        <th 
                          className="px-4 py-3.5 min-w-[150px] border-r border-slate-855 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleSortToggle('progress')}
                          title="Clique para ordenar por Quantidade Pedida / Recebida"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Pedido / Recebido</span>
                            <span className="inline-flex shrink-0 font-sans normal-case">
                              {sortConfig.key === 'progress' ? (
                                sortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* STATUS LOGÍSTICO */}
                        <th 
                          className="px-4 py-3.5 min-w-[140px] border-r border-slate-850 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleSortToggle('status')}
                          title="Clique para ordenar por Status Logístico"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider">Status Logístico</span>
                            <span className="inline-flex shrink-0">
                              {sortConfig.key === 'status' ? (
                                sortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        <th className="px-4 py-3.5 text-right font-sans font-semibold text-[10px] uppercase tracking-wider text-slate-400 select-none">
                          <span className="text-[10px] text-slate-400 font-semibold uppercase tracking-wider block">Painel de Ações</span>
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/60 bg-[#161920]">
                      {filteredPedidos.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="text-center py-12 text-slate-500">
                            <Layers size={36} className="mx-auto text-slate-700 mb-2.5" />
                            Nenhum pedido atende aos filtros definidos nesta aba.
                          </td>
                        </tr>
                      ) : (
                        filteredPedidos.map((p) => {
                          const progressVal = p.qtdSolicitada && p.qtdSolicitada > 0 ? Math.round((p.qtdRecebida / p.qtdSolicitada) * 100) : 0;
                          const cleanCodeLine = (() => {
                            const codInsuStr = p.codigo !== undefined && p.codigo !== null ? String(p.codigo) : '';
                            const codDetStr = p.codDetalhe !== undefined && p.codDetalhe !== null ? String(p.codDetalhe) : '';
                            const codInsuNum = codInsuStr ? codInsuStr.replace(/[^\d]/g, '') : '';
                            const codDetNum = codDetStr ? codDetStr.replace(/[^\d]/g, '') : '';
                            return [codInsuNum, codDetNum].filter(Boolean).join(' / ');
                          })();
                          return (
                            <tr key={`${p.id}-${p.codigo}`} className="hover:bg-slate-800/20 transition-all font-sans">
                              {/* CÓDIGO / EMISSÃO */}
                              <td className="px-4 py-4 whitespace-nowrap">
                                <div className="text-white font-bold text-sm tracking-wide">{p.id}</div>
                                <div className="text-slate-500 text-[11px] mt-1 font-mono">{p.dataPedido}</div>
                              </td>

                              {/* C. CUSTO / OBRA */}
                              <td className="px-4 py-4 whitespace-normal max-w-[200px]">
                                <div className="font-semibold text-slate-200 text-sm whitespace-nowrap overflow-hidden text-ellipsis" title={p.obra}>
                                  {p.obra}
                                </div>
                                {p.obraId && (
                                  <div className="text-[10px] text-slate-500 font-mono mt-0.5">Cód: {p.obraId}</div>
                                )}
                              </td>

                              {/* FORNECEDOR */}
                              <td className="px-4 py-4 whitespace-normal max-w-[200px]">
                                <div className="font-medium text-slate-200 text-sm whitespace-nowrap overflow-hidden text-ellipsis" title={p.fornecedor || 'Não Especificado'}>
                                  {p.fornecedor || 'Não especificado'}
                                </div>
                                {p.codComprador && (
                                  <div className="text-[10.5px] text-slate-500 mt-0.5">
                                    Comprador: <span className="font-mono">{p.codComprador}</span>
                                  </div>
                                )}
                              </td>

                              {/* INSUMO SIENGE */}
                              <td className="px-4 py-4">
                                <div className="font-bold text-slate-100 text-[14px]">
                                  {p.insumo}
                                </div>
                                {p.descricaoDetalhe && (
                                  <div className="text-slate-400 text-xs mt-0.5 italic leading-tight">
                                    {p.descricaoDetalhe}
                                  </div>
                                )}
                                <div className="text-[11px] text-slate-500 mt-1 flex items-center gap-1.5 flex-wrap">
                                  <span>Unid: <strong className="text-slate-400 uppercase font-mono">{p.unidade}</strong></span>
                                  {cleanCodeLine && (
                                    <>
                                      <span className="text-slate-700">•</span>
                                      <span className="font-mono text-slate-400">Cód: {cleanCodeLine}</span>
                                    </>
                                  )}
                                  {p.marca && (
                                    <>
                                      <span className="text-slate-700">•</span>
                                      <span className="text-purple-400/90 font-medium">Marca: {p.marca}</span>
                                    </>
                                  )}
                                </div>
                              </td>

                              {/* PEDIDO / RECEBIDO */}
                              <td className="px-4 py-4">
                                <div className="flex flex-col items-center">
                                  <div className="font-mono font-bold flex items-center gap-1 text-[13px] text-slate-200">
                                    <span>{p.qtdSolicitada}</span>
                                    <span className="text-slate-500">/</span>
                                    <span className={p.qtdRecebida === p.qtdSolicitada ? 'text-emerald-400 font-extrabold' : p.qtdRecebida > 0 ? 'text-amber-400 font-extrabold' : 'text-slate-300 font-extrabold'}>
                                      {p.qtdRecebida}
                                    </span>
                                    <span className="text-[10px] text-slate-400 uppercase font-bold ml-1">{p.unidade}</span>
                                  </div>
                                  <div className="w-28 bg-[#0F1115] rounded-full h-1.5 overflow-hidden mt-2">
                                    <div
                                      className={`h-full rounded-full transition-all duration-300 ${
                                        p.status === 'Entregue' ? 'bg-emerald-500' :
                                        p.status === 'Parcial' ? 'bg-amber-500' :
                                        p.status === 'Cancelado' ? 'bg-rose-500' : 'bg-slate-700'
                                      }`}
                                      style={{ width: `${Math.min(progressVal, 100)}%` }}
                                    ></div>
                                  </div>
                                </div>
                              </td>

                              {/* STATUS LOGÍSTICO */}
                              <td className="px-4 py-4 whitespace-nowrap">
                                <div className="flex flex-col items-start gap-1.5">
                                  <span className={`px-3 py-1 rounded-md text-[11px] tracking-wide font-bold min-w-[145px] text-center inline-block ${
                                    p.status === 'Entregue' ? 'border border-emerald-555 text-emerald-400 bg-emerald-500/5' :
                                    p.status === 'Parcial' ? 'border border-amber-555 text-amber-400 bg-amber-500/5' :
                                    p.status === 'Cancelado' ? 'border border-rose-500/20 text-rose-450 bg-rose-500/5' :
                                    'border border-slate-700/60 text-slate-300 bg-slate-800/40'
                                  }`}>
                                    {p.status === 'Entregue' ? 'Totalmente entregue' :
                                     p.status === 'Parcial' ? 'Parcialmente entregue' :
                                     p.status === 'Cancelado' ? 'Cancelado' : 'Pendente'}
                                  </span>
                                  
                                  <select
                                    value={p.status}
                                    onChange={(e) => {
                                      const newStatus = e.target.value as 'Pendente' | 'Parcial' | 'Entregue' | 'Cancelado';
                                      setPedidos(pedidos.map(item => {
                                        if (item.id === p.id && item.codigo === p.codigo) {
                                          let updatedRecebida = item.qtdRecebida;
                                          if (newStatus === 'Entregue') {
                                            updatedRecebida = item.qtdSolicitada;
                                          } else if (newStatus === 'Pendente') {
                                            updatedRecebida = 0;
                                          }
                                          return {
                                            ...item,
                                            status: newStatus,
                                            qtdRecebida: updatedRecebida,
                                            dataChegada: newStatus === 'Entregue' ? (item.dataChegada || new Date().toLocaleDateString('pt-BR')) : undefined
                                          };
                                        }
                                        return item;
                                      }));
                                      triggerToast(`Status alterado para ${newStatus}.`, 'success');
                                    }}
                                    className="w-full max-w-[145px] bg-[#0F1115] border border-slate-800 text-slate-400 text-[10px] rounded px-2 py-1 focus:border-purple-500 focus:outline-none cursor-pointer"
                                  >
                                    <option value="Pendente">Marcar Pendente</option>
                                    <option value="Parcial">Marcar Parcial</option>
                                    <option value="Entregue">Marcar Completo</option>
                                    <option value="Cancelado">Marcar Cancelado</option>
                                  </select>

                                  {p.status === 'Entregue' ? (
                                    <span className="text-[10px] text-purple-400/90 font-medium tracking-wide mt-0.5">⏱️ Entregue em {calculateDaysElapsed(p.dataPedido, p.status, p.dataChegada)} dias</span>
                                  ) : p.status === 'Cancelado' ? (
                                    <span className="text-[10px] text-slate-500 tracking-wide mt-0.5">⏱️ Cancelado</span>
                                  ) : (
                                    <span className="text-[10px] text-amber-500 font-medium tracking-wide mt-0.5">⏱️ Ativo há {calculateDaysElapsed(p.dataPedido, p.status, p.dataChegada)} dias</span>
                                  )}
                                </div>
                              </td>

                              {/* PAINEL DE AÇÕES */}
                              <td className="px-4 py-4 text-right">
                                <div className="flex gap-2 items-center justify-end">
                                  {/* Edit button */}
                                  <button
                                    onClick={() => handleOpenEditPedidoModal(p)}
                                    className="p-1.5 h-8 w-8 text-slate-400 hover:text-white bg-slate-800/25 hover:bg-slate-800 rounded border border-slate-700/60 inline-flex items-center justify-center transition-all"
                                    title="Editar Pedido"
                                  >
                                    <Edit size={14} />
                                  </button>

                                  {/* Principal action button */}
                                  {p.status === 'Entregue' ? (
                                    <div className="px-3 py-1.5 h-8 bg-slate-900/60 text-slate-500 text-[11px] font-semibold rounded border border-slate-800/80 flex items-center gap-1 cursor-default select-none">
                                      <Check size={12} className="text-slate-500" />
                                      Finalizado
                                    </div>
                                  ) : p.status === 'Cancelado' ? (
                                    <div className="px-3 py-1.5 h-8 bg-slate-900/40 text-slate-650 text-[11px] font-semibold rounded border border-slate-900/60 flex items-center gap-1 cursor-default select-none">
                                      Cancelado
                                    </div>
                                  ) : (
                                    <button
                                      onClick={() => openReceiveModal(p.id, p.codigo || '')}
                                      className="px-3 py-1.5 h-8 bg-[#4F46E5] hover:bg-[#5850EC] text-white rounded text-[11px] font-extrabold tracking-wide transition-all shadow-sm flex items-center justify-center whitespace-nowrap"
                                    >
                                      RECEBIMENTO
                                    </button>
                                  )}

                                  {/* Delete button */}
                                  <button
                                    onClick={() => handleDeletePedido(p.id)}
                                    className="p-1.5 h-8 w-8 text-rose-500/70 hover:text-rose-400 bg-rose-955/5 hover:bg-rose-955/20 rounded border border-rose-900/20 inline-flex items-center justify-center transition-all"
                                    title="Excluir Pedido"
                                  >
                                    <Trash2 size={13} />
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  {filteredPedidos.length === 0 ? (
                    <div className="col-span-full text-center py-12 bg-[#161920] border border-slate-800 rounded-xl text-slate-500">
                      <Layers size={36} className="mx-auto text-slate-700 mb-2.5" />
                      Nenhum pedido atende aos filtros definidos nesta aba.
                    </div>
                  ) : (
                    filteredPedidos.map((p) => {
                      const progressVal = p.qtdSolicitada && p.qtdSolicitada > 0 ? Math.round((p.qtdRecebida / p.qtdSolicitada) * 100) : 0;
                      return (
                        <div key={`${p.id}-${p.codigo}`} className="bg-[#161920] border border-slate-800 rounded-xl p-5 hover:border-slate-700 transition-all flex flex-col justify-between">
                          
                          <div>
                            <div className="flex justify-between items-start mb-3">
                              <span className="text-xs font-semibold text-slate-400 font-mono flex flex-wrap gap-1.5 items-center">
                                <span className="text-slate-500">{p.id}</span>
                                <span className="text-slate-600">•</span>
                                <span className="bg-[#1D2028] text-yellow-500/90 text-[10px] px-1.5 py-0.5 rounded font-bold border border-yellow-500/10">
                                  {p.codigo || '-'}
                                </span>
                                <span className="text-slate-600">•</span>
                                <span className="text-slate-500">{p.dataPedido}</span>
                              </span>
                              <span className={`px-2 py-0.5 rounded text-[10px] tracking-wider uppercase font-extrabold ${
                                p.status === 'Entregue' ? 'bg-emerald-500/15 text-emerald-400' :
                                p.status === 'Parcial' ? 'bg-blue-500/15 text-blue-400' :
                                p.status === 'Cancelado' ? 'bg-rose-500/15 text-rose-400' : 'bg-amber-500/15 text-amber-500'
                              }`}>
                                {p.status}
                              </span>
                            </div>

                            <div className="text-[10px] text-purple-400/95 font-sans mb-1.5 inline-block font-medium">
                              {p.status === 'Entregue' ? (
                                <span className="bg-purple-500/5 px-2 py-0.5 rounded border border-purple-500/10 block">⏱️ Entregue em {calculateDaysElapsed(p.dataPedido, p.status, p.dataChegada)} dias</span>
                              ) : p.status === 'Cancelado' ? (
                                <span className="text-slate-500 block">⏱️ Cancelado</span>
                              ) : (
                                <span className="bg-amber-500/5 text-amber-400 border border-amber-500/10 px-2 py-0.5 rounded block">⏱️ Ativo há {calculateDaysElapsed(p.dataPedido, p.status, p.dataChegada)} dias</span>
                              )}
                            </div>

                            <h4 className="text-base font-semibold text-white mb-1">{p.insumo}</h4>
                            
                            <p className="text-xs text-slate-400 flex items-center gap-1 mb-2">
                              <MapPin size={12} className="text-slate-500" />
                              {p.obra} {p.obraId ? `[${p.obraId}]` : ''}
                            </p>

                            <div className="flex flex-wrap items-center gap-x-2 gap-y-1 text-[10px] text-slate-500 mb-3.5 border-t border-slate-800/40 pt-2">
                              <span>Unid: <strong>{p.unidade}</strong>{p.descricaoUnidade ? ` (${p.descricaoUnidade})` : ''}</span>
                              {p.marca && <span className="px-1.5 py-0.2 bg-slate-800 text-slate-300 rounded font-mono">marca: {p.marca}</span>}
                              {p.fornecedor && <span className="text-purple-400">forn: <strong>{p.fornecedor}</strong></span>}
                              {p.codComprador && <span className="text-blue-400 font-mono">comprador: {p.codComprador}</span>}
                            </div>

                            {p.descricaoDetalhe && (
                              <div className="text-xs text-slate-450 italic font-sans leading-relaxed mb-4 bg-slate-900/40 p-2 rounded border border-slate-800/40">
                                “{p.descricaoDetalhe}” {p.codDetalhe && <span className="font-mono text-[9px] text-slate-500">[{p.codDetalhe}]</span>}
                              </div>
                            )}
                          </div>

                          {/* Order Progress Details */}
                          <div className="space-y-3 mt-auto">
                            <div className="flex justify-between text-xs">
                              <span className="text-slate-500">Quantidade Solicitada / Recebida:</span>
                              <span className="text-slate-300 font-semibold">{p.qtdSolicitada} {p.unidade} / <span className="text-emerald-400">{p.qtdRecebida} {p.unidade}</span></span>
                            </div>

                            <div className="w-full bg-slate-900 rounded-full h-2 overflow-hidden shadow-inner">
                              <div
                                className={`h-full rounded-full transition-all duration-300 ${
                                  p.status === 'Entregue' ? 'bg-emerald-500' :
                                  p.status === 'Parcial' ? 'bg-blue-400' :
                                  p.status === 'Cancelado' ? 'bg-rose-500' : 'bg-amber-500'
                                }`}
                                style={{ width: `${progressVal}%` }}
                              ></div>
                            </div>

                            <div className="flex items-center justify-between text-xs pt-2">
                              <div className="text-slate-400 font-mono text-[11px]">Falta: <span className="text-slate-300 font-bold">{p.qtdSolicitada - p.qtdRecebida} {p.unidade}</span></div>
                              <div className="flex gap-2 items-center">
                                {p.status !== 'Entregue' && p.status !== 'Cancelado' && (
                                  <>
                                    <button
                                      onClick={() => {
                                        setPedidos(pedidos.map(item => item.id === p.id && item.codigo === p.codigo ? { ...item, status: 'Cancelado' as const } : item));
                                        triggerToast(`Pedido ${p.id} cancelado com sucesso.`, 'info');
                                      }}
                                      className="px-2.5 py-1 text-rose-500 hover:text-white hover:bg-rose-950/40 rounded transition-all text-xs font-semibold"
                                    >
                                      Cancelar
                                    </button>
                                    <button
                                      onClick={() => openReceiveModal(p.id, p.codigo || '')}
                                      className="px-3 py-1 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-xs font-semibold transition-colors"
                                    >
                                      RECEBIMENTO
                                    </button>
                                  </>
                                )}

                                {/* Administrative actions */}
                                <button
                                  onClick={() => handleOpenEditPedidoModal(p)}
                                  className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded transition-all"
                                  title="Editar Especificações"
                                >
                                  <Edit size={14} />
                                </button>
                                <button
                                  onClick={() => handleDeletePedido(p.id)}
                                  className="p-1.5 text-rose-500/70 hover:text-rose-400 hover:bg-slate-800 rounded transition-all"
                                  title="Excluir Definitivamente"
                                >
                                  <Trash2 size={14} />
                                </button>
                              </div>
                            </div>
                          </div>

                        </div>
                      );
                    })
                  )}
                </div>
              )}
            </div>
          )}

          {/* VIEW TAB 3: ESTOQUE POR OBRA */}
          {currentTab === 'estoque' && (
            <div className="space-y-6" id="estoque-tab-view">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-xl font-sans font-extrabold text-white">Disponibilidade e Saldos de Estoque</h2>
                  <p className="text-sm text-slate-400 mt-0.5">Visão do inventário em lista considerando as entradas recebidas menos as baixas locais em obra.</p>
                </div>
              </div>

              {/* Search & Stats Header */}
              <div className="bg-[#161920] border border-slate-800 p-4.5 rounded-xl flex flex-col md:flex-row gap-3 justify-between items-center bg-[#1C2028]">
                <div className="relative w-full md:w-80">
                  <input
                    type="text"
                    placeholder="Pesquisar estoque (insumo, código, obra)..."
                    value={searchQuery}
                    onChange={(e) => setSearchQuery(e.target.value)}
                    className="w-full bg-[#0F1115]/95 border border-slate-700 rounded-lg py-1.5 pl-9 pr-3 text-xs text-white placeholder:text-slate-500 focus:outline-none focus:border-emerald-500 transition-colors"
                  />
                  <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-500" />
                </div>
                
                <div className="flex flex-wrap items-center gap-4 text-xs font-medium">
                  <span className="text-slate-400">
                    Filtrando por: <strong className="text-emerald-400 font-bold">{selectedObra}</strong>
                  </span>
                  <span className="bg-slate-800 px-3 py-1 rounded text-slate-300">
                    Total Itens: {filteredStock.length}
                  </span>
                </div>
              </div>

              {/* Table List View */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl overflow-hidden shadow-xl" id="inventory-table-container">
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm border-collapse" id="inventory-list-table">
                    <thead>
                      <tr className="text-slate-450 font-semibold border-b border-slate-800 bg-[#12151C] text-[10px] uppercase tracking-wider">
                        {/* CENTRO DE CUSTO / OBRA */}
                        <th 
                          className="px-6 py-3.5 text-left text-slate-400 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleInventorySortToggle('obra')}
                          title="Clique para ordenar por Centro de Custo / Obra"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span>Centro de Custo / Obra</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'obra' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* MATERIAL / INSUMO */}
                        <th 
                          className="px-6 py-3.5 text-left text-slate-400 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleInventorySortToggle('insumo')}
                          title="Clique para ordenar por Material / Insumo"
                        >
                          <div className="flex items-center justify-between gap-1.5">
                            <span>Material / Insumo</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'insumo' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* UNIDADE */}
                        <th 
                          className="px-6 py-3.5 text-center text-slate-400 cursor-pointer select-none group hover:bg-slate-800/40 transition-all ml-auto"
                          onClick={() => handleInventorySortToggle('unidade')}
                          title="Clique para ordenar por Unidade"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span>Unidade</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'unidade' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* ENTRADAS ACUMULADAS (+) */}
                        <th 
                          className="px-6 py-3.5 text-center text-[#10B981] cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleInventorySortToggle('recebido')}
                          title="Clique para ordenar por Entradas Acumuladas"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span>Entradas Acumuladas (+)</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'recebido' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-[#10B981]/50 group-hover:text-[#10B981] transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* BAIXAS REGISTRADAS (-) */}
                        <th 
                          className="px-6 py-3.5 text-center text-[#F43F5E] cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleInventorySortToggle('baixado')}
                          title="Clique para ordenar por Baixas Registradas"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span>Baixas Registradas (-)</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'baixado' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-[#F43F5E]/50 group-hover:text-[#F43F5E] transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* SALDO FÍSICO ATUAL */}
                        <th 
                          className="px-6 py-3.5 text-center text-[#FFC800] cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleInventorySortToggle('saldo')}
                          title="Clique para ordenar por Saldo Físico"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span>Saldo Físico Atual</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'saldo' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-[#FFC800]/50 group-hover:text-[#FFC800] transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* ÚLTIMA MOVIMENTAÇÃO */}
                        <th 
                          className="px-6 py-3.5 text-center text-slate-400 cursor-pointer select-none group hover:bg-slate-800/40 transition-all"
                          onClick={() => handleInventorySortToggle('ultimaMovimentacao')}
                          title="Clique para ordenar por Última Movimentação"
                        >
                          <div className="flex items-center justify-center gap-1.5">
                            <span>Última Movimentação</span>
                            <span className="inline-flex shrink-0">
                              {inventorySortConfig.key === 'ultimaMovimentacao' ? (
                                inventorySortConfig.direction === 'asc' ? (
                                  <ArrowUp size={13} className="text-emerald-400 font-bold" />
                                ) : (
                                  <ArrowDown size={13} className="text-emerald-400 font-bold" />
                                )
                              ) : (
                                <ArrowUpDown size={13} className="text-slate-600 group-hover:text-slate-450 transition-colors opacity-40" />
                              )}
                            </span>
                          </div>
                        </th>

                        {/* AJUSTES / INVENTÁRIO */}
                        <th className="px-6 py-3.5 text-center text-slate-400 font-semibold select-none">
                          Ajustes / Inventário
                        </th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/85">
                      {filteredStock.length === 0 ? (
                        <tr>
                          <td colSpan={8} className="text-center py-12 text-slate-500">
                            <Layers size={36} className="mx-auto text-slate-700 mb-2.5" />
                            Nenhum insumo encontrado no estoque com os filtros aplicados.
                          </td>
                        </tr>
                      ) : (
                        filteredStock.map((item, idx) => {
                          return (
                            <tr key={idx} className="hover:bg-slate-800/20 transition-colors border-b border-slate-800/40">
                              {/* 1. Centro de Custo / Obra */}
                              <td className="px-6 py-2">
                                <div className="text-white font-bold text-sm tracking-wide uppercase">{item.obra}</div>
                              </td>

                              {/* 2. Material / Insumo */}
                              <td className="px-6 py-2">
                                <div className="text-white font-bold text-sm tracking-wide uppercase">{item.insumo}</div>
                                {(item.codDetalhe || item.descricaoDetalhe) && (
                                  <div className="mt-0.5 text-xs text-slate-400 max-w-sm font-sans flex items-center gap-1.5 leading-tight">
                                    {item.codDetalhe && item.descricaoDetalhe ? (
                                      <>
                                        <span className="font-mono font-semibold text-slate-350">{item.codDetalhe}</span>
                                        <span className="text-slate-500">-</span>
                                        <span className="text-slate-400 font-medium">{item.descricaoDetalhe}</span>
                                      </>
                                    ) : (
                                      <>
                                        {item.codDetalhe && <span className="font-mono font-semibold text-slate-350">{item.codDetalhe}</span>}
                                        {item.descricaoDetalhe && <span className="text-slate-400 font-medium">{item.descricaoDetalhe}</span>}
                                      </>
                                    )}
                                  </div>
                                )}
                              </td>

                              {/* 3. Unidade */}
                              <td className="px-6 py-2 text-center font-bold text-xs uppercase text-slate-400 font-sans">
                                {item.unidade || '-'}
                              </td>

                              {/* 4. Entradas Acumuladas */}
                              <td className="px-6 py-2 text-center font-sans">
                                <div className="text-[#10B981] font-extrabold text-base">
                                  {item.recebido.toLocaleString('pt-BR')}
                                </div>
                                <div className="text-[10px] text-slate-500 font-mono mt-0.5 whitespace-nowrap">
                                  {item.dataUltimaEntrada && item.dataUltimaEntrada !== '-' ? (
                                    item.dataUltimaEntrada
                                  ) : (
                                    '-'
                                  )}
                                </div>
                              </td>

                              {/* 5. Baixas Registradas */}
                              <td className="px-6 py-2 text-center font-sans">
                                <div className="text-[#F43F5E] font-extrabold text-base">
                                  {item.baixado.toLocaleString('pt-BR')}
                                </div>
                                <div className="text-[10px] text-slate-500 font-mono mt-0.5 whitespace-nowrap">
                                  {item.dataUltimaBaixa && item.dataUltimaBaixa !== '-' ? (
                                    item.dataUltimaBaixa
                                  ) : (
                                    '-'
                                  )}
                                </div>
                              </td>

                              {/* 6. Saldo Físico Atual */}
                              <td className="px-6 py-2 text-center">
                                <div className="inline-block border border-[#FFC800]/40 rounded-lg text-[#FFC800] bg-[#FFC800]/5 px-3 py-1 font-black text-sm tracking-wider font-mono min-w-[70px]">
                                  {item.saldo.toLocaleString('pt-BR')}
                                </div>
                              </td>

                              {/* 7. Última Movimentação */}
                              <td className="px-6 py-2 text-center font-mono text-xs text-slate-500 whitespace-nowrap">
                                {item.ultimaMovimentacao || '-'}
                              </td>

                              {/* 8. Ajustes / Inventário */}
                              <td className="px-6 py-2 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  <button
                                    onClick={() => {
                                      setReconcileItem(item);
                                      setReconcileNewPhysicalQty(item.saldo);
                                      setReconcileJustificativa('');
                                      setReconcileResponsavel('');
                                      setShowReconcileModal(true);
                                    }}
                                    className="px-3.5 py-1.5 bg-transparent hover:bg-[#FFC800]/5 border border-[#FFC800]/30 hover:border-[#FFC800] text-[#FFC800] text-[10px] font-black rounded uppercase tracking-wider transition-colors inline-flex items-center gap-1"
                                  >
                                    Reconciliar Saldo
                                  </button>
                                  <button
                                    onClick={() => {
                                      setNewWithdrawObra(item.obra);
                                      setNewWithdrawInsumo(item.insumo);
                                      setNewWithdrawQtd(Math.min(10, item.saldo));
                                      setShowWithdrawModal(true);
                                    }}
                                    className="px-3.5 py-1.5 bg-[#F43F5E] hover:bg-rose-500 text-white text-[10px] font-black rounded uppercase tracking-wider transition-colors inline-flex items-center gap-1"
                                  >
                                    Baixar Consumo
                                  </button>
                                </div>
                              </td>
                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* VIEW TAB 4: ADVANCED REPORTS & WITHDRAWAL LOGS */}
          {currentTab === 'relatorios' && (
            <div className="space-y-6 animate-in fade-in duration-200" id="reports-tab-view">
              
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                <div>
                  <h2 className="text-xl font-sans font-extrabold text-white">Relatórios e Diários de Obras</h2>
                  <p className="text-sm text-slate-400 mt-1">Gere relatórios customizados, valide consumos físicos e filtre históricos consolidados em tempo real.</p>
                </div>
                <button
                  onClick={handleExportCSV}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-md text-sm font-semibold border border-slate-705 transition-all flex items-center gap-2 self-start cursor-pointer shadow-md"
                >
                  <FileDown size={15} className="text-emerald-400" /> Exportar Planilha Filtrada
                </button>
              </div>

              {/* Advanced Interactive Filter Box */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl p-5" id="report-filters-panel">
                <div className="flex items-center gap-2 mb-4 text-white">
                  <Filter size={16} className="text-emerald-400" />
                  <h3 className="text-xs font-bold uppercase tracking-wider">Painel Dinâmico de Filtros do Relatório</h3>
                </div>
                
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 xl:grid-cols-7 gap-4">
                  
                  {/* 1. Start Date */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">De (Período):</label>
                    <input
                      type="date"
                      value={reportStartDate}
                      onChange={(e) => setReportStartDate(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                    />
                  </div>

                  {/* 2. End Date */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">Até (Período):</label>
                    <input
                      type="date"
                      value={reportEndDate}
                      onChange={(e) => setReportEndDate(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                    />
                  </div>

                  {/* 3. Obra Selection */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">Obra Destino:</label>
                    <select
                      value={reportSelectedObra}
                      onChange={(e) => setReportSelectedObra(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                    >
                      <option value="Todas as Obras">Todas as Obras</option>
                      {obras.map((o) => (
                        <option key={o.id} value={o.nome}>{o.nome}</option>
                      ))}
                    </select>
                  </div>

                  {/* 4. User / Operador */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">Responsável:</label>
                    <select
                      value={reportSelectedUser}
                      onChange={(e) => setReportSelectedUser(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer text-ellipsis"
                    >
                      <option value="Todos os Usuários">Todos os Usuários</option>
                      {reportUsersList.map((user) => (
                        <option key={user} value={user}>{user}</option>
                      ))}
                    </select>
                  </div>

                  {/* 5. Type (Entrada vs Saída) */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">Tipo:</label>
                    <select
                      value={reportSelectedType}
                      onChange={(e) => setReportSelectedType(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                    >
                      <option value="Ambos">Ambos (Entrada/Saída/Ajuste)</option>
                      <option value="Entrada">Entrada (Abastecimento)</option>
                      <option value="Saída">Saída (Consumo/Gasto)</option>
                      <option value="Ajuste">Ajuste (Reconciliação)</option>
                    </select>
                  </div>

                  {/* 6. Material / Insumo */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">Material / Insumo:</label>
                    <select
                      value={reportSelectedMaterial}
                      onChange={(e) => setReportSelectedMaterial(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer text-ellipsis"
                    >
                      <option value="Todos os Materiais">Todos os Materiais</option>
                      {reportMaterialsList.map((material) => (
                        <option key={material} value={material}>{material}</option>
                      ))}
                    </select>
                  </div>

                  {/* 7. Supplier / Fornecedor */}
                  <div className="space-y-1.5">
                    <label className="text-[11px] font-semibold text-slate-400 block">Fornecedor:</label>
                    <select
                      value={reportSelectedSupplier}
                      onChange={(e) => setReportSelectedSupplier(e.target.value)}
                      className="w-full bg-[#1F242D] border border-slate-700 rounded px-2.5 py-1.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer text-ellipsis"
                    >
                      <option value="Todos os Fornecedores">Todos os Fornecedores</option>
                      {reportSuppliersList.map((sup) => (
                        <option key={sup} value={sup}>{sup}</option>
                      ))}
                    </select>
                  </div>

                </div>

                {/* Clear Filters Action */}
                {(reportStartDate || reportEndDate || reportSelectedObra !== 'Todas as Obras' || reportSelectedUser !== 'Todos os Usuários' || reportSelectedType !== 'Ambos' || reportSelectedMaterial !== 'Todos os Materiais' || reportSelectedSupplier !== 'Todos os Fornecedores') && (
                  <div className="flex justify-end mt-4">
                    <button
                      onClick={() => {
                        setReportStartDate('');
                        setReportEndDate('');
                        setReportSelectedObra('Todas as Obras');
                        setReportSelectedUser('Todos os Usuários');
                        setReportSelectedType('Ambos');
                        setReportSelectedMaterial('Todos os Materiais');
                        setReportSelectedSupplier('Todos os Fornecedores');
                        triggerToast('Filtros do relatório redefinidos!', 'info');
                      }}
                      className="px-3 py-1 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded text-xs transition-all flex items-center gap-1.5 cursor-pointer border border-slate-700"
                    >
                      Restaurar Todos os Filtros
                    </button>
                  </div>
                )}
              </div>

              {/* Dynamic KPI summary card in search scope */}
              <div className="grid grid-cols-1 sm:grid-cols-4 gap-4" id="filtered-report-summary-cards">
                <div className="p-4 bg-[#161920] border border-slate-800 rounded-xl">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">Registros Encontrados</span>
                  <div className="mt-1 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-sans text-white">{filteredMovements.length}</span>
                    <span className="text-xs text-slate-500">movimentações</span>
                  </div>
                </div>
                <div className="p-4 bg-[#161920] border border-slate-800 rounded-xl">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">Entradas Filtradas</span>
                  <div className="mt-1 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-sans text-emerald-400">
                      {filteredMovements.filter(m => m.tipo === 'ENTRADA').length}
                    </span>
                    <span className="text-xs text-slate-500">lotes abastecidos</span>
                  </div>
                </div>
                <div className="p-4 bg-[#161920] border border-slate-800 rounded-xl">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">Saídas / Consumo</span>
                  <div className="mt-1 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-sans text-rose-400">
                      {filteredMovements.filter(m => m.tipo === 'SAIDA').length}
                    </span>
                    <span className="text-xs text-slate-500">baixas registradas</span>
                  </div>
                </div>
                <div className="p-4 bg-[#161920] border border-slate-800 rounded-xl">
                  <span className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block">Ajustes Realizados</span>
                  <div className="mt-1 flex items-baseline gap-2">
                    <span className="text-2xl font-black font-sans text-[#FFC800]">
                      {filteredMovements.filter(m => m.tipo === 'AJUSTE').length}
                    </span>
                    <span className="text-xs text-slate-500">reconciliações</span>
                  </div>
                </div>
              </div>

              {/* Log List of historical withdrawals (Historico de Baixas / Diário de Retiradas) resembling user image */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl overflow-hidden" id="withdraw-logs-panel">
                <div className="p-5 border-b border-slate-800 bg-[#1C2028] flex justify-between items-center">
                  <div>
                    <h3 className="text-base font-sans font-bold text-white uppercase tracking-tight">Diário de Retiradas e Consumo</h3>
                    <p className="text-xs text-slate-500 mt-0.5">Logs sequenciais de movimentações registradas por mestre de obras, almoxarifes e estagiários.</p>
                  </div>
                  <History size={18} className="text-slate-400" />
                </div>

                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse" id="reports-diary-table">
                    <thead>
                      <tr className="bg-[#0F1115] border-b border-slate-800/80 text-[10px] font-sans font-extrabold text-slate-455 uppercase tracking-widest leading-none">
                        <th className="py-4.5 px-5">Data / Hora</th>
                        <th className="py-4.5 px-4 text-center">Tipo</th>
                        <th className="py-4.5 px-4">Obra Destino</th>
                        <th className="py-4.5 px-4">Código</th>
                        <th className="py-4.5 px-4">Insumo</th>
                        <th className="py-4.5 px-4 text-right">QTD.</th>
                        <th className="py-4.5 px-5">Documento / Justificativa de Uso</th>
                        <th className="py-4.5 px-4">Responsável Técnico</th>
                        <th className="py-4.5 px-5 text-center">Estorno</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/50 font-sans text-xs">
                      {filteredMovements.length === 0 ? (
                        <tr>
                          <td colSpan={9} className="py-14 text-center text-slate-500 bg-[#161920]">
                            <ClipboardList size={36} className="mx-auto text-slate-800 mb-3" />
                            <p className="text-sm font-semibold text-white">Nenhuma movimentação de estoque localizada.</p>
                            <p className="text-[11px] text-slate-500 mt-1">Limpe os filtros de pesquisa acima para reexibir todos os lançamentos.</p>
                          </td>
                        </tr>
                      ) : (
                        filteredMovements.map((item) => {
                          const isEntrada = item.tipo === 'ENTRADA';
                          const isAjuste = item.tipo === 'AJUSTE';

                          // Determine displayed sign, value, and color class for quantity
                          let qtyDisplay = '';
                          let qtyColorClass = 'text-rose-400';

                          if (isEntrada) {
                            qtyDisplay = `${item.quantidade}`;
                            qtyColorClass = 'text-emerald-400';
                          } else if (isAjuste) {
                            if (item.quantidade < 0) {
                              qtyDisplay = `${Math.abs(item.quantidade)}`;
                              qtyColorClass = 'text-emerald-400';
                            } else {
                              qtyDisplay = `-${Math.abs(item.quantidade)}`;
                              qtyColorClass = 'text-rose-400';
                            }
                          } else {
                            qtyDisplay = `-${item.quantidade}`;
                            qtyColorClass = 'text-rose-400';
                          }

                          return (
                            <tr key={`${item.tipo}:::${item.id}:::${item.codigo}`} className="hover:bg-slate-800/10 transition-colors">
                              
                              {/* DATA / HORA */}
                              <td className="py-4 px-5 font-mono text-slate-450 whitespace-nowrap">
                                {item.data}
                              </td>

                              {/* TIPO badge in layout format */}
                              <td className="py-4 px-4 text-center whitespace-nowrap">
                                {isEntrada ? (
                                  <span className="inline-block px-3 py-1 text-[9px] font-extrabold uppercase bg-[#004D40]/30 text-emerald-400 border border-emerald-500/20 rounded">
                                    Entrada
                                  </span>
                                ) : isAjuste ? (
                                  <span className="inline-block px-3 py-1 text-[9px] font-extrabold uppercase bg-[#FFC800]/10 text-[#FFC800] border border-[#FFC800]/20 rounded">
                                    Ajuste
                                  </span>
                                ) : (
                                  <span className="inline-block px-3 py-1 text-[9px] font-extrabold uppercase bg-[#4A1525]/30 text-rose-400 border border-rose-500/20 rounded">
                                    Saída
                                  </span>
                                )}
                              </td>

                              {/* OBRA DESTINO */}
                              <td className="py-4 px-4 font-bold text-slate-200">
                                {item.obra}
                              </td>

                              {/* CÓDIGO DO INSUMO */}
                              <td className="py-4 px-4 font-mono text-slate-400">
                                {item.codigo || '-'}
                              </td>

                              {/* INSUMO */}
                              <td className="py-4 px-4" title={item.insumo}>
                                <span className="block font-bold text-slate-200">{item.insumo}</span>
                                {(item.codDetalhe || item.descricaoDetalhe) ? (
                                  <span className="block text-[11px] text-slate-400 font-medium font-sans mt-0.5">
                                    {`${item.codDetalhe || ''} ${item.descricaoDetalhe || ''}`.trim()}
                                  </span>
                                ) : (
                                  <span className="block text-[11px] text-slate-500 italic mt-0.5">
                                    Sem detalhamento adicional
                                  </span>
                                )}
                              </td>

                              {/* QUANTIDADE */}
                              <td className={`py-4 px-4 text-right font-bold font-mono whitespace-nowrap ${qtyColorClass}`}>
                                {qtyDisplay} <span className="text-[10px] font-medium text-slate-500">{item.unidade}</span>
                              </td>

                              {/* DOCUMENTO / JUSTIFICATIVA DE USO */}
                              <td className="py-4 px-5 text-slate-400" title={item.documento}>
                                {item.documento}
                              </td>

                              {/* RESPONSÁVEL TÉCNICO */}
                              <td className="py-4 px-4 text-slate-350 font-semibold whitespace-nowrap">
                                {item.responsavel}
                              </td>

                              {/* ESTORNO */}
                              <td className="py-4 px-5 text-center">
                                <button
                                  onClick={() => handleEstornoMovement(item)}
                                  className="p-1.5 duration-150 hover:bg-rose-500/10 hover:text-rose-450 text-rose-500/60 rounded-full cursor-pointer transition-all border border-transparent hover:border-rose-500/20 shadow-sm"
                                  title={`Estornar esta ${isEntrada ? 'entrada (devolução)' : 'baixa'} de material`}
                                >
                                  <Trash2 size={13} />
                                </button>
                              </td>

                            </tr>
                          );
                        })
                      )}
                    </tbody>
                  </table>
                </div>

              </div>

            </div>
          )}

          {/* VIEW TAB 5: ADMIN - CONFIGURAÇÕES, OBRAS E PERMISSÕES */}
          {currentTab === 'admin' && (
            <div className="space-y-6 animate-in fade-in duration-200" id="admin-tab-pane">
              
              {/* Adm Banner Info */}
              <div className="p-5 bg-gradient-to-r from-purple-950/20 to-[#161920] border border-purple-500/20 rounded-xl flex items-start gap-3.5">
                <div className="p-2.5 bg-[#4c1d95]/20 rounded-lg text-purple-400 border border-purple-500/20 shrink-0">
                  <Shield size={20} />
                </div>
                <div>
                  <h3 className="text-sm font-extrabold text-white font-sans">Painel de Administração Corporal & Permissões</h3>
                  <p className="text-xs text-slate-400 mt-1 leading-relaxed">
                    Operando como <strong>{currentUser.nome}</strong>. Modifique autonomias de ação do time, crie instalações e altere senhas de mestre de obras. Use o menu suspenso operador na barra lateral para trocar de perfil instantaneamente e testar suas regras.
                  </p>
                </div>
              </div>

              {/* SECTOR: GOOGLE SHEETS & DEPLOYMENT UTILITIES */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6" id="google-sheets-configs">
                
                {/* Panel A: Google Sheets Sync Setup */}
                <div className="bg-[#161920] border border-slate-800 rounded-xl p-5 flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800">
                      <h3 className="text-sm font-sans font-extrabold text-white flex items-center gap-2">
                        <Database size={16} className="text-emerald-400" />
                        Conexão em Nuvem Google Sheets
                      </h3>
                      <span className="text-[10px] bg-emerald-500/10 text-emerald-400 px-1.5 py-0.5 rounded font-mono uppercase">API Ativa</span>
                    </div>

                    <p className="text-xs text-slate-400 mb-4 leading-relaxed">
                      Sincronize todo o banco de dados (Pedidos, Baixas, Obras, Usuários) diretamente com as abas do seu Google Planilhas via Apps Script.
                    </p>

                    <div className="space-y-3.5 mb-4">
                      <div>
                        <label className="text-[10px] text-slate-500 font-bold uppercase tracking-wider block mb-1.5">URL do Google Web App:</label>
                        <input
                          type="text"
                          value={googleSheetsUrl}
                          onChange={(e) => setGoogleSheetsUrl(e.target.value)}
                          placeholder="https://script.google.com/macros/s/.../exec"
                          className="w-full bg-[#0F1115] border border-slate-700 rounded-lg py-2 px-3 text-xs text-slate-300 placeholder-slate-600 focus:border-emerald-500 focus:outline-none font-mono"
                        />
                      </div>

                      <div className="flex items-center justify-between p-3 bg-[#0F1115] rounded-lg border border-slate-800">
                        <div className="flex items-center gap-2">
                          <input
                            type="checkbox"
                            checked={autoSync}
                            onChange={(e) => setAutoSync(e.target.checked)}
                            id="auto-sync-checkbox"
                            className="rounded text-emerald-600 bg-slate-905 border-slate-800 focus:ring-emerald-500 focus:ring-0"
                          />
                          <label htmlFor="auto-sync-checkbox" className="text-xs text-slate-300 font-medium cursor-pointer">Sincronização em tempo real</label>
                        </div>
                        <span className="text-[9px] text-slate-500 font-mono">Auto Save</span>
                      </div>
                    </div>
                  </div>

                  <div className="flex gap-2.5 pt-4 border-t border-slate-800">
                    <button
                      onClick={() => handleSheetsSync('pull')}
                      disabled={isSyncing}
                      className="flex-1 py-1.5 px-3 bg-slate-800 hover:bg-slate-700 disabled:opacity-55 text-slate-200 border border-slate-705 rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1.5"
                    >
                      <RefreshCw size={13} className={isSyncing ? "animate-spin" : ""} />
                      Puxar Dados (Pull)
                    </button>
                    <button
                      onClick={() => handleSheetsSync('push')}
                      disabled={isSyncing}
                      className="flex-1 py-1.5 px-3 bg-emerald-650 hover:bg-emerald-600 disabled:opacity-55 text-white rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-1.5"
                    >
                      <Upload size={13} />
                      Enviar Dados (Push)
                    </button>
                  </div>
                </div>

                {/* Panel B: Vercel & Zip Packager */}
                <div className="bg-[#161920] border border-slate-800 rounded-xl p-5 flex flex-col justify-between">
                  <div>
                    <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800">
                      <h3 className="text-sm font-sans font-extrabold text-white flex items-center gap-2">
                        <FileDown size={16} className="text-purple-400" />
                        Exportar Pacote para GitHub / Vercel
                      </h3>
                      <span className="text-[10px] bg-purple-500/10 text-purple-400 px-1.5 py-0.5 rounded font-mono uppercase">Vercel Ready</span>
                    </div>

                    <p className="text-xs text-slate-400 mb-4 leading-relaxed">
                      Faça o download de todos os arquivos do aplicativo compactados em uma pasta ZIP. O pacote contém o código principal com o <strong className="text-purple-450 font-semibold">code.gs</strong> para o Apps Script e um manual passo-a-passo.
                    </p>

                    <div className="p-3.5 bg-purple-950/10 border border-purple-900/30 rounded-lg text-xs text-slate-300 space-y-2">
                      <div className="font-semibold text-purple-400 flex items-center gap-1.5">
                        <Info size={14} />
                        Pronto para Produção
                      </div>
                      <p className="text-[11px] leading-relaxed text-slate-400">
                        A estrutura de arquivos gerada está 100% livre de dados soltos e otimizada para ser conectada diretamente a um repositório no Github com deploy automático na Vercel de forma gratuita.
                      </p>
                    </div>
                  </div>

                  <div className="pt-4 border-t border-slate-800">
                    <button
                      onClick={handleExportProjectZip}
                      className="w-full py-2.5 px-4 bg-purple-650 hover:bg-purple-600 text-white rounded-lg text-xs font-bold transition-all flex items-center justify-center gap-2 shadow-lg shadow-purple-950/20"
                    >
                      <Download size={14} />
                      Baixar Projeto Completo em ZIP
                    </button>
                  </div>
                </div>

              </div>

              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
                
                {/* 1. SECTOR: GESTÃO DAS OBRAS (5 Columns) */}
                <div className="lg:col-span-5 bg-[#161920] border border-slate-800 rounded-xl p-5 flex flex-col justify-between" id="admin-obras-panel">
                  <div>
                    <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800">
                      <h3 className="text-sm font-sans font-extrabold text-white flex items-center gap-2">
                        <MapPin size={16} className="text-purple-400" />
                        Obras em Andamento ({obras.length})
                      </h3>
                      <span className="text-[10px] text-slate-500 font-mono select-none">Global</span>
                    </div>

                    <p className="text-xs text-slate-400 mb-4 leading-normal">
                      Exclua ou renomeie as obras cadastradas. As modificações atualizam pedidos de compras e diários de saídas correspondentes.
                    </p>

                    {/* Obras lists */}
                    <div className="space-y-2 max-h-[300px] overflow-y-auto pr-1">
                      {obras.map((obrObj) => (
                        <div key={obrObj.id} className="flex items-center justify-between p-3 bg-[#0F1115] rounded-lg border border-slate-800 group hover:border-slate-700 transition-all">
                          {editingObraName === obrObj.nome ? (
                            <div className="flex flex-col gap-2 w-full p-1 bg-[#0F1115] rounded">
                              <div className="grid grid-cols-3 gap-2">
                                <div className="col-span-1">
                                  <label className="text-[9px] text-slate-500 font-mono block uppercase mb-1">ID</label>
                                  <input
                                    type="text"
                                    defaultValue={obrObj.id}
                                    id={`input-rename-id-${obrObj.id}`}
                                    placeholder="Ex: OB-101"
                                    className="bg-slate-900 border border-slate-700 text-xs text-white rounded px-2 py-1 w-full focus:outline-none uppercase font-mono"
                                  />
                                </div>
                                <div className="col-span-2">
                                  <label className="text-[9px] text-slate-500 font-mono block uppercase mb-1">Nome da Obra</label>
                                  <input
                                    type="text"
                                    defaultValue={obrObj.nome}
                                    id={`input-rename-obra-${obrObj.id}`}
                                    className="bg-slate-900 border border-slate-700 text-xs text-white rounded px-2 py-1 w-full focus:outline-none"
                                    onKeyDown={(e) => {
                                      if (e.key === 'Enter') {
                                        const idVal = (document.getElementById(`input-rename-id-${obrObj.id}`) as HTMLInputElement)?.value;
                                        handleRenameObra(obrObj.nome, (e.target as HTMLInputElement).value, idVal);
                                      } else if (e.key === 'Escape') {
                                        setEditingObraName(null);
                                      }
                                    }}
                                  />
                                </div>
                              </div>
                              <div className="flex justify-end gap-2 mt-1">
                                <button
                                  type="button"
                                  onClick={() => setEditingObraName(null)}
                                  className="text-[10px] text-slate-400 hover:text-slate-200 transition-colors py-0.5 px-2"
                                >
                                  Cancelar
                                </button>
                                <button
                                  type="button"
                                  onClick={() => {
                                    const idVal = (document.getElementById(`input-rename-id-${obrObj.id}`) as HTMLInputElement)?.value || '';
                                    const val = (document.getElementById(`input-rename-obra-${obrObj.id}`) as HTMLInputElement)?.value || '';
                                    handleRenameObra(obrObj.nome, val, idVal);
                                  }}
                                  className="text-white bg-purple-600 hover:bg-purple-500 text-[10px] font-bold px-2.5 py-1 rounded transition-all"
                                >
                                  Salvar
                                </button>
                              </div>
                            </div>
                          ) : (
                            <>
                              <div className="flex flex-col truncate max-w-[180px]">
                                <span className="text-xs font-semibold text-white truncate">{obrObj.nome}</span>
                                <span className="text-[10px] text-slate-500 font-mono">ID: {obrObj.id}</span>
                              </div>
                              <div className="flex items-center gap-1.5 opacity-80 group-hover:opacity-100 transition-all">
                                <button
                                  type="button"
                                  onClick={() => setEditingObraName(obrObj.nome)}
                                  className="p-1 text-slate-400 hover:text-white rounded hover:bg-slate-800 transition-all"
                                  title="Renomear Empreendimento"
                                >
                                  <Edit size={12} />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteObra(obrObj.nome)}
                                  className="p-1 text-rose-450 hover:text-rose-400 rounded hover:bg-slate-800 transition-all"
                                  title="Remover Empreendimento"
                                >
                                  <Trash2 size={12} />
                                </button>
                              </div>
                            </>
                          )}
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Add Obra Form */}
                  <div className="mt-6 pt-4 border-t border-slate-800">
                    <label className="text-[11px] text-slate-400 font-semibold block mb-2">Cadastrar Novo Empreendimento/Obra:</label>
                    <div className="space-y-2">
                      <div className="grid grid-cols-3 gap-2">
                        <input
                          type="text"
                          placeholder="ID Obra (Ex: OB-201, PETR)"
                          id="new-obra-id-input-field"
                          className="bg-[#0F1115] border border-slate-700 rounded-lg py-1.5 px-3 text-xs text-slate-200 placeholder-slate-600 focus:border-purple-500 focus:outline-none col-span-1 font-mono uppercase"
                        />
                        <input
                          type="text"
                          placeholder="Nome da Obra..."
                          id="new-obra-name-input-field"
                          className="bg-[#0F1115] border border-slate-700 rounded-lg py-1.5 px-3 text-xs text-slate-200 placeholder-slate-600 focus:border-purple-500 focus:outline-none col-span-2"
                        />
                      </div>
                      <button
                        type="button"
                        onClick={() => {
                          const idVal = (document.getElementById('new-obra-id-input-field') as HTMLInputElement)?.value || '';
                          const nameVal = (document.getElementById('new-obra-name-input-field') as HTMLInputElement)?.value || '';
                          handleCreateObra(nameVal, idVal);
                          const inpId = document.getElementById('new-obra-id-input-field') as HTMLInputElement;
                          const inpName = document.getElementById('new-obra-name-input-field') as HTMLInputElement;
                          if (inpId) inpId.value = '';
                          if (inpName) inpName.value = '';
                        }}
                        className="w-full bg-purple-650 hover:bg-purple-600 text-white rounded-lg py-2 text-xs font-bold transition-all flex items-center justify-center gap-1.5 shrink-0"
                      >
                        <Plus size={14} />
                        Cadastrar Empreendimento com ID
                      </button>
                    </div>
                  </div>

                </div>

                {/* 2. SECTOR: USUÁRIOS E PERMISSÕES (7 Columns) */}
                <div className="lg:col-span-7 bg-[#161920] border border-slate-800 rounded-xl p-5 flex flex-col justify-between" id="admin-usuarios-panel">
                  <div>
                    <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800">
                      <h3 className="text-sm font-sans font-extrabold text-white flex items-center gap-2">
                        <Users size={16} className="text-purple-400" />
                        Lista de Colaboradores e Autonomias ({usuarios.length})
                      </h3>
                      <button
                        type="button"
                        onClick={() => setNewUserOpen(true)}
                        className="text-[11px] bg-purple-550/15 text-purple-400 hover:bg-purple-600 hover:text-white px-2.5 py-1 rounded border border-purple-500/20 font-bold transition-all flex items-center gap-1"
                      >
                        <Plus size={12} />
                        Novo Colaborador
                      </button>
                    </div>

                    <p className="text-xs text-slate-400 mb-4 leading-normal">
                      Dica: Selecione as permissões específicas para cada profile ou altere credenciais de login e senha provisória de acesso.
                    </p>

                    {/* Users list with checkboxes */}
                    <div className="space-y-4 max-h-[400px] overflow-y-auto pr-1">
                      {usuarios.map((usr) => (
                        <div key={usr.id} className="p-4 bg-[#0F1115] rounded-xl border border-slate-800 flex flex-col gap-3">
                          
                          {/* Profile details */}
                          <div className="flex items-center justify-between">
                            <div className="flex items-center gap-2.5">
                              <div className="w-8 h-8 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center text-xs font-bold text-yellow-500">
                                {usr.nome.substring(0, 2).toUpperCase()}
                              </div>
                              <div>
                                <h4 className="text-xs font-bold text-white flex items-center gap-1.5">
                                  {usr.nome}
                                  <span className={`text-[9px] px-1.5 py-0.2 rounded font-bold uppercase ${
                                    usr.role === 'Administrador' ? 'bg-purple-500/20 text-purple-400' : 'bg-blue-500/20 text-blue-400'
                                  }`}>
                                    {usr.role}
                                  </span>
                                </h4>
                                <span className="text-[10px] text-slate-500 font-mono">ID: {usr.id}</span>
                              </div>
                            </div>

                            {/* Self login blocker */}
                            {currentUser.id !== usr.id && (
                              <button
                                type="button"
                                onClick={() => handleSwitchSimulatorUser(usr)}
                                className="text-[10px] text-amber-500 hover:text-white hover:bg-amber-600/10 px-2 py-0.5 rounded border border-amber-500/20 font-medium transition-all"
                              >
                                Logar Como
                              </button>
                            )}
                          </div>

                          {/* Editable login and password */}
                          <div className="grid grid-cols-2 gap-3 pt-2 border-t border-slate-800/60">
                            <div>
                              <span className="text-[9px] text-slate-500 font-bold block mb-1">E-mail / Login:</span>
                              <input
                                type="text"
                                value={usr.login}
                                onChange={(e) => handleEditCredential(usr.id, 'login', e.target.value)}
                                className="w-full bg-[#161920] border border-slate-800 rounded px-2.5 py-1 text-xs text-slate-300 font-mono focus:border-purple-500 focus:outline-none"
                              />
                            </div>
                            <div>
                              <span className="text-[9px] text-slate-500 font-bold block mb-1">Senha de Acesso:</span>
                              <input
                                type="text"
                                value={usr.senha}
                                onChange={(e) => handleEditCredential(usr.id, 'senha', e.target.value)}
                                className="w-full bg-[#161920] border border-slate-800 rounded px-2.5 py-1 text-xs text-slate-300 focus:border-purple-500 focus:outline-none"
                              />
                            </div>
                          </div>

                          {/* Checkboxes matrix */}
                          <div className="p-3 bg-[#161920]/40 rounded-lg border border-slate-800/80 mt-1">
                            <span className="text-[9px] text-purple-400 font-bold uppercase tracking-wider block mb-2">Permissões Específicas do Perfil:</span>
                            
                            <div className="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                              <label className="flex items-center gap-2 cursor-pointer text-xs text-slate-300 hover:text-white transition-all select-none">
                                <input
                                  type="checkbox"
                                  checked={usr.podeCriarPedido}
                                  onChange={(e) => handleTogglePermission(usr.id, 'podeCriarPedido', e.target.checked)}
                                  className="rounded text-purple-600 bg-slate-905 border-slate-800 focus:ring-purple-500 focus:ring-0"
                                />
                                Pode Criar Pedidos
                              </label>

                              <label className="flex items-center gap-2 cursor-pointer text-xs text-slate-300 hover:text-white transition-all select-none">
                                <input
                                  type="checkbox"
                                  checked={usr.podeDarBaixa}
                                  onChange={(e) => handleTogglePermission(usr.id, 'podeDarBaixa', e.target.checked)}
                                  className="rounded text-purple-600 bg-slate-905 border-slate-800 focus:ring-purple-500 focus:ring-0"
                                />
                                Pode Gastar / Baixar
                              </label>

                              <label className="flex items-center gap-2 cursor-pointer text-xs text-slate-300 hover:text-white transition-all select-none">
                                <input
                                  type="checkbox"
                                  checked={usr.podeReceberMercadoria}
                                  onChange={(e) => handleTogglePermission(usr.id, 'podeReceberMercadoria', e.target.checked)}
                                  className="rounded text-purple-600 bg-slate-905 border-slate-800 focus:ring-purple-500 focus:ring-0"
                                />
                                Registrar Entrada Física
                              </label>

                              <label className="flex items-center gap-2 cursor-pointer text-xs text-slate-300 hover:text-white transition-all select-none">
                                <input
                                  type="checkbox"
                                  checked={usr.podeExcluirPedido}
                                  onChange={(e) => handleTogglePermission(usr.id, 'podeExcluirPedido', e.target.checked)}
                                  className="rounded text-purple-600 bg-slate-905 border-slate-800 focus:ring-purple-500 focus:ring-0"
                                />
                                Excluir Pedidos de Obra
                              </label>

                              <label className="flex items-center gap-2 cursor-pointer text-xs text-rose-400 hover:text-rose-350 transition-all select-none col-span-1 sm:col-span-2 mt-1">
                                <input
                                  type="checkbox"
                                  checked={usr.podeExcluirObra}
                                  onChange={(e) => handleTogglePermission(usr.id, 'podeExcluirObra', e.target.checked)}
                                  className="rounded text-rose-600 bg-slate-905 border-slate-800 focus:ring-rose-550 focus:ring-0"
                                />
                                Excluir Instalações Globais
                              </label>
                            </div>
                          </div>

                          {/* Delete operator button */}
                          {usr.id !== currentUser.id && (
                            <div className="flex justify-end mt-1">
                              <button
                                type="button"
                                onClick={() => handleDeleteUsuario(usr.id)}
                                className="text-rose-500 hover:text-rose-450 hover:underline text-[10px] font-semibold"
                              >
                                Apagar Registro de Operador
                              </button>
                            </div>
                          )}

                        </div>
                      ))}
                    </div>
                  </div>

                </div>

              </div>

            </div>
          )}

        </div>

      </main>

      {/* MODAL: IMPORTADOR DE FLUXO SIENGE EXCEL */}
      {showImportModal && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#161920] border border-slate-700 w-full max-w-2xl rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {/* Heading */}
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="text-emerald-500 w-5 h-5" />
                <div>
                  <h3 className="text-sm font-bold text-white uppercase tracking-wider">Importador de Pedidos Sienge ERP</h3>
                  <span className="text-[10px] text-slate-500 mt-0.5">Sincronização e inteligência de merge</span>
                </div>
              </div>
              <button
                onClick={() => setShowImportModal(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-6 space-y-5">
              
              {/* Box 1: Instructions & Quick Demo Loader */}
              <div className="p-4 bg-[#0F1115] rounded-lg border border-slate-850 text-xs text-slate-400 space-y-2.5">
                <p className="leading-relaxed">
                  Para importar, arraste ou selecione sua planilha de compra do <strong>Sienge</strong> (.xls, .xlsx) ou utilize o modelo de demonstração pré-configurado com itens repetidos e novos no botão abaixo para testar a congruência.
                </p>
                <div className="flex items-center justify-between pt-1">
                  <span className="text-[10px] text-slate-500 font-mono">Formato Aceito: Excel / CSV</span>
                  <button
                    type="button"
                    onClick={handleLoadDemoSienge}
                    className="text-emerald-400 hover:text-emerald-300 font-semibold underline text-xs flex items-center gap-1"
                  >
                    💡 Carregar Planilha Modelo de Teste
                  </button>
                </div>
              </div>

              {/* Box 2: Dropzone / Picker */}
              <div className="border border-dashed border-slate-700 hover:border-emerald-500/50 bg-[#0F1115]/30 rounded-xl p-6 flex flex-col items-center justify-center gap-2.5 transition-all relative">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={handleUploadSienge}
                  className="absolute inset-0 opacity-0 cursor-pointer w-full h-full z-20"
                />
                {parsedImportItems.length > 0 ? (
                  <div className="flex flex-col items-center gap-2 z-10 text-center">
                    <CheckCircle className="text-emerald-500 w-10 h-10 animate-bounce" />
                    <div>
                      <span className="text-xs text-emerald-400 font-bold block">✓ Arquivo Carregado com Sucesso!</span>
                      {importedFileName && (
                        <span className="text-[11px] text-slate-350 font-mono inline-block mt-1.5 bg-slate-800/80 px-2.5 py-1 rounded border border-slate-700">
                          📄 {importedFileName}
                        </span>
                      )}
                      <span className="text-[10px] text-slate-500 block mt-1">
                        Sugerindo <strong className="text-white">{parsedImportItems.length} itens</strong> para importar/sincronizar.
                      </span>
                    </div>
                  </div>
                ) : (
                  <>
                    <Upload className="text-slate-500 w-8 h-8 pointer-events-none" />
                    <div className="text-center pointer-events-none">
                      <span className="text-xs text-slate-300 font-semibold block">Escolher Planilha de Compras Sienge</span>
                      <span className="text-[10px] text-slate-500 block mt-0.5">Arraste para cá ou clique para explorar seu computador</span>
                    </div>
                  </>
                )}
              </div>

              {/* Botão de Enviar Requisitado diretamente após Anexar */}
              {parsedImportItems.length > 0 && (
                <div className="p-4 bg-emerald-950/25 border border-emerald-500/30 rounded-xl flex flex-col md:flex-row items-center justify-between gap-4 animate-in slide-in-from-top-4 duration-200">
                  <div className="space-y-1 text-center md:text-left">
                    <h4 className="text-xs font-bold text-emerald-400 flex items-center justify-center md:justify-start gap-1">
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" /> Pronto para Processar!
                    </h4>
                    <p className="text-[10px] text-slate-400 leading-relaxed">
                      Sincronize novas ordens e recalcule as quantidades e saldos sem alterar as entradas físicas já registradas.
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      processImportedPedidos(parsedImportItems);
                      setShowImportModal(false);
                    }}
                    className="w-full md:w-auto px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-black shadow-lg shadow-emerald-950/50 transition-all flex items-center justify-center gap-2 hover:scale-105 active:scale-95"
                  >
                    <CheckCircle size={14} />
                    Sincronizar e Enviar Agora
                  </button>
                </div>
              )}

              {/* Box 3: Table Preview of Parsed Items */}
              {parsedImportItems.length > 0 && (
                <div className="space-y-2">
                  <span className="text-[11px] text-slate-400 font-semibold block tracking-wider uppercase">Itens Identificados na Planilha ({parsedImportItems.length}):</span>
                  <div className="max-h-[180px] overflow-y-auto border border-slate-800 rounded-lg overflow-hidden">
                    <table className="w-full text-left text-xs text-slate-300">
                      <thead className="bg-[#1C2028] text-slate-400 text-[10px] uppercase font-mono tracking-wider border-b border-slate-800">
                        <tr>
                          <th className="px-3 py-2">Cód. Pedido</th>
                          <th className="px-3 py-2">Insumo</th>
                          <th className="px-3 py-2">Obra</th>
                          <th className="px-3 py-2 text-right">Qtd</th>
                          <th className="px-3 py-2">Unid</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-800 bg-[#0F1115]">
                        {parsedImportItems.map((item, idx) => {
                          const isNewOrder = !pedidos.some(p => p.id === item.id);
                          const isSameOrderNewItem = !isNewOrder && !pedidos.some(p => p.id === item.id && p.codigo === item.codigo);
                          const isQtdUpdate = !isNewOrder && !isSameOrderNewItem && pedidos.some(p => p.id === item.id && p.codigo === item.codigo && p.qtdSolicitada !== item.qtdSolicitada);
                          
                          return (
                            <tr key={idx} className="hover:bg-slate-800/40">
                              <td className="px-3 py-2 font-mono text-[10px]">
                                {item.id}
                                {isNewOrder && <span className="ml-1.5 px-1 py-0.2 text-[8px] bg-purple-500/20 text-purple-400 font-sans font-bold rounded uppercase">Novo</span>}
                              </td>
                              <td className="px-3 py-2 font-medium truncate max-w-[150px]">
                                {item.insumo}
                                {isSameOrderNewItem && <span className="ml-1.5 px-1 py-0.2 text-[8px] bg-blue-500/20 text-blue-400 font-sans font-bold rounded uppercase">Item Novo</span>}
                              </td>
                              <td className="px-3 py-2 truncate max-w-[100px] text-slate-400">{item.obra}</td>
                              <td className="px-3 py-2 text-right font-semibold text-slate-200">
                                {item.qtdSolicitada}
                                {isQtdUpdate && <span className="ml-1 text-[9px] text-amber-500">🔄</span>}
                              </td>
                              <td className="px-3 py-2 text-slate-500 font-mono">{item.unidade}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                  
                  {/* Legend guide */}
                  <div className="flex gap-4 text-[10px] text-slate-500 justify-end pt-1">
                    <span className="flex items-center gap-1">
                      <span className="w-2 h-2 rounded bg-purple-500/20 inline-block"></span> Novo Pedido
                    </span>
                    <span className="flex items-center gap-1">
                      <span className="w-2 h-2 rounded bg-blue-500/20 inline-block"></span> Novo Item na mesma PO
                    </span>
                    <span className="flex items-center gap-1">
                      <span>🔄</span> Atualização de Qtd Comprada
                    </span>
                  </div>
                </div>
              )}

              {/* Box 4: Warning and Upsert Rules info */}
              <div className="flex gap-2.5 p-3.5 bg-amber-950/15 border border-amber-900/30 rounded-lg text-xs leading-relaxed text-slate-400">
                <AlertCircle className="text-amber-500 shrink-0 w-4 h-4 mt-0.5" />
                <p>
                  <strong>Regra de Unicidade:</strong> O sistema reconhecerá automaticamente se um ID de compra já existe no banco de dados corporativo, listará novos fluxos adicionados e recalculará as quantidades solicitadas sem alterar as entradas físicas já recebidas da obra!
                </p>
              </div>

            </div>

            {/* Footer */}
            <div className="p-4 bg-[#1C2028] border-t border-slate-800 flex justify-between gap-3">
              <button
                type="button"
                onClick={() => setShowImportModal(false)}
                className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-lg text-xs font-bold transition-all"
              >
                Cancelar
              </button>
              <button
                type="button"
                onClick={() => {
                  processImportedPedidos(parsedImportItems);
                  setShowImportModal(false);
                }}
                disabled={parsedImportItems.length === 0}
                className={`px-4 py-2 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 ${
                  parsedImportItems.length === 0
                    ? 'bg-slate-800 text-slate-600 cursor-not-allowed'
                    : 'bg-emerald-600 hover:bg-emerald-500 text-white shadow-md shadow-emerald-900/10'
                }`}
              >
                <CheckCircle size={14} />
                Confirmar & Consolidar Sincronismo
              </button>
            </div>

          </div>
        </div>
      )}

      {/* MODAL 1: REGISTRAR NOVA ENTRADA / PEDIDO */}
      {showOrderModal && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#161920] border border-slate-700 w-full max-w-md rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {/* Modal Heading */}
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <h3 className="text-base font-sans font-extrabold text-white">Nova Entrada: Comprar Insumo</h3>
              <button
                onClick={() => setShowOrderModal(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            {/* Modal Body / Form */}
            <form onSubmit={handleCreateOrder} className="p-6 space-y-4">
              
              {/* Construction Site Selection */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Destinar para Obra:</label>
                <select
                  value={newOrderObra}
                  onChange={(e) => setNewOrderObra(e.target.value)}
                  className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-sm rounded-lg p-2.5 focus:border-emerald-500 focus:outline-none"
                >
                  {listObras.map((o) => (
                    <option key={o} value={o}>{o}</option>
                  ))}
                </select>
              </div>

              {/* Material/Insumo Name input */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Nome do Insumo (especificar tipo/medida):</label>
                <input
                  type="text"
                  placeholder="Ex: Cimento CP-II 50kg, Areia Fina..."
                  value={newOrderInsumo}
                  onChange={(e) => setNewOrderInsumo(e.target.value)}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-sm text-slate-300 focus:border-emerald-500 focus:outline-none placeholder-slate-600"
                />
              </div>

              {/* Código do Insumo (opcional / auto-gerado) */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Código do Insumo (Opcional):</label>
                <input
                  type="text"
                  placeholder="Ex: CIM-001 (Vazio para auto-gerar)"
                  value={newOrderCodigo}
                  onChange={(e) => setNewOrderCodigo(e.target.value)}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-sm text-slate-300 focus:border-emerald-500 focus:outline-none placeholder-slate-600 font-mono"
                />
              </div>

              {/* Quantity fields & Units */}
              <div className="grid grid-cols-2 gap-3">
                
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Quantidade:</label>
                  <input
                    type="number"
                    min="1"
                    value={newOrderQtd}
                    onChange={(e) => setNewOrderQtd(Math.max(1, parseInt(e.target.value) || 0))}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-sm text-slate-300 focus:border-emerald-500 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Unidade de Medida:</label>
                  <select
                    value={newOrderUnidade}
                    onChange={(e) => setNewOrderUnidade(e.target.value)}
                    className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-sm rounded-lg p-2.5 focus:border-emerald-500 focus:outline-none"
                  >
                    <option value="un">un (unidade)</option>
                    <option value="kg">kg (quilograma)</option>
                    <option value="m³">m³ (metro cúbico)</option>
                    <option value="m²">m² (metro quadrado)</option>
                    <option value="m">m (metro linear)</option>
                  </select>
                </div>

              </div>

              {/* Footer Form row */}
              <div className="pt-4 border-t border-slate-800 flex justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setShowOrderModal(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-sm transition-all"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-sm font-semibold transition-all"
                >
                  Adicionar Pedido
                </button>
              </div>

            </form>
          </div>
        </div>
      )}

      {/* MODAL 2: REGISTRAR BAIXA DE MATERIAL */}
      {showWithdrawModal && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#0B0D12] border border-slate-800 w-full max-w-md rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {/* Header */}
            <div className="p-5 border-b border-slate-850 flex justify-between items-start bg-[#0D1017]">
              <div>
                <h3 className="text-white font-bold text-sm tracking-wide">
                  Baixa de material / Saida de estoque
                </h3>
                <p className="text-slate-400 text-[11px] mt-0.5 leading-tight">
                  Dedução direta no saldo físico do almoxarifado
                </p>
              </div>
              <button
                type="button"
                onClick={() => setShowWithdrawModal(false)}
                className="text-slate-400 hover:text-white text-xs font-semibold uppercase tracking-wider transition-all pt-0.5"
              >
                Fechar
              </button>
            </div>

            {(() => {
              const selectedStockItem = stockInventory.find(
                item => item.obra === newWithdrawObra && item.insumo === newWithdrawInsumo
              );
              const maxAvailable = selectedStockItem ? selectedStockItem.saldo : 0;

              return (
                <form onSubmit={handleCreateWithdrawal} className="p-5 space-y-4">
                  
                  {/* 1. Centro de Custo / Obra Solicitante */}
                  <div>
                    <label className="text-xs text-slate-350 font-bold block mb-1.5 font-sans">
                      1. Centro de Custo / Obra Solicitante:
                    </label>
                    <div className="w-full bg-[#07090D]/80 border border-slate-850 text-slate-400 text-sm rounded-lg p-2.5 font-sans font-medium flex items-center justify-between">
                      <span className="font-semibold text-slate-300">{newWithdrawObra}</span>
                      <span className="text-[9px] text-slate-500 uppercase tracking-widest font-black bg-slate-900 px-1.5 py-0.5 rounded border border-slate-850">
                        Preenchido
                      </span>
                    </div>
                  </div>

                  {/* 2. Material / Insumo Requisitado */}
                  <div>
                    <label className="text-xs text-slate-350 font-bold block mb-1.5 font-sans">
                      2. Material / Insumo Requisitado:
                    </label>
                    <div className="w-full bg-[#07090D]/80 border border-slate-850 text-slate-300 text-sm rounded-lg p-3 font-sans flex flex-col gap-1">
                      <span className="font-bold text-white text-sm">{newWithdrawInsumo}</span>
                      {selectedStockItem && (selectedStockItem.codDetalhe || selectedStockItem.descricaoDetalhe) && (
                        <div className="text-xs text-slate-400 mt-1 flex items-center flex-wrap gap-1 border-t border-slate-850/60 pt-1.5">
                          {selectedStockItem.codDetalhe && selectedStockItem.descricaoDetalhe ? (
                            <>
                              <span className="font-mono font-semibold text-emerald-400">{selectedStockItem.codDetalhe}</span>
                              <span className="text-slate-500">-</span>
                              <span className="text-slate-400 font-medium">{selectedStockItem.descricaoDetalhe}</span>
                            </>
                          ) : (
                            <span className="font-mono font-semibold text-emerald-400">
                              {selectedStockItem.codDetalhe || selectedStockItem.descricaoDetalhe}
                            </span>
                          )}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Banner: Saldo para Liberação Física */}
                  {selectedStockItem && (
                    <div className="border border-slate-800 rounded-lg bg-[#07090D] px-4 py-2.5 flex items-center justify-between">
                      <span className="text-slate-450 text-xs font-semibold font-sans">
                        Saldo para Liberação Física:
                      </span>
                      <span className="text-[#FFC800] font-black text-sm tracking-wide font-mono">
                        {selectedStockItem.saldo} {selectedStockItem.unidade}
                      </span>
                    </div>
                  )}

                  {/* 3. Quantidade Consumida */}
                  <div>
                    <label className="text-xs text-slate-300 font-bold block mb-1.5 font-sans">
                      3. Quantidade Consumida: <span className="text-rose-500 font-bold ml-0.5">*</span>
                    </label>
                    <input
                      type="number"
                      min="1"
                      max={maxAvailable}
                      placeholder="Ex: 25"
                      value={newWithdrawQtd || ''}
                      onChange={(e) => {
                        const val = parseInt(e.target.value);
                        if (isNaN(val)) {
                          setNewWithdrawQtd(0);
                        } else {
                          // Imposes maximum stock limit strictly on input
                          setNewWithdrawQtd(Math.min(maxAvailable, Math.max(0, val)));
                        }
                      }}
                      required
                      className="w-full bg-[#07090D] border border-slate-805 rounded-lg p-2.5 text-sm text-slate-200 placeholder-slate-500 focus:border-slate-705 focus:outline-none transition-all font-sans"
                    />
                    {newWithdrawQtd > maxAvailable && (
                      <p className="text-[11px] text-rose-500 mt-1 font-sans">
                        A quantidade consumida não pode ser maior do que o saldo total ({maxAvailable}).
                      </p>
                    )}
                  </div>

                  {/* 4. Destino de Aplicação / Local */}
                  <div>
                    <label className="text-xs text-slate-300 font-bold block mb-1.5 font-sans flex items-center gap-1">
                      <span>4. Destino de Aplicação / Local:</span>
                      <span className="text-rose-500 font-bold">*</span>
                    </label>
                    <input
                      type="text"
                      placeholder="Ex: Alvenaria estrutural do 3º pavimento do bloco B"
                      value={newWithdrawDestino}
                      onChange={(e) => setNewWithdrawDestino(e.target.value)}
                      required
                      className="w-full bg-[#07090D] border border-slate-805 rounded-lg p-2.5 text-sm text-slate-200 placeholder-slate-500 focus:border-slate-705 focus:outline-none transition-all font-sans"
                    />
                  </div>

                  {/* 5. Autorizado por (Mestre / Encarregado) */}
                  <div>
                    <label className="text-xs text-slate-300 font-bold block mb-1.5 font-sans flex items-center gap-1">
                      <span>5. Autorizado por (Mestre / Encarregado):</span>
                      <span className="text-rose-500 font-bold">*</span>
                    </label>
                    <input
                      type="text"
                      placeholder="Ex: Eng. Ricardo Albuquerque"
                      value={newWithdrawAutorizado}
                      onChange={(e) => setNewWithdrawAutorizado(e.target.value)}
                      required
                      className="w-full bg-[#07090D] border border-slate-805 rounded-lg p-2.5 text-sm text-slate-200 placeholder-slate-500 focus:border-slate-705 focus:outline-none transition-all font-sans"
                    />
                  </div>

                  {/* Footer Action Buttons */}
                  <div className="pt-4 border-t border-slate-850 flex justify-end gap-3 bg-[#0B0D12]">
                    <button
                      type="button"
                      onClick={() => setShowWithdrawModal(false)}
                      className="px-5 py-2.5 bg-[#1E2530] hover:bg-[#252E3B] text-slate-300 text-sm font-bold rounded-lg transition-all font-sans"
                    >
                      Cancelar
                    </button>
                    <button
                      type="submit"
                      disabled={maxAvailable === 0 || !newWithdrawInsumo || newWithdrawQtd <= 0 || newWithdrawQtd > maxAvailable || !newWithdrawDestino.trim() || !newWithdrawAutorizado.trim()}
                      className={`px-5 py-2.5 rounded-lg text-sm font-bold transition-all font-sans ${
                        maxAvailable === 0 || !newWithdrawInsumo || newWithdrawQtd <= 0 || newWithdrawQtd > maxAvailable || !newWithdrawDestino.trim() || !newWithdrawAutorizado.trim()
                          ? 'bg-[#1E2530]/50 text-slate-600 cursor-not-allowed border border-slate-900/40'
                          : 'bg-emerald-600 hover:bg-emerald-500 text-white shadow-md shadow-emerald-950/25'
                      }`}
                    >
                      Confirmar Saída
                    </button>
                  </div>

                </form>
              );
            })()}
          </div>
        </div>
      )}

      {/* OVERLAY DE SUCESSO DA BAIXA (CENTERED POPUP) */}
      {showWithdrawSuccessScreen && (
        <div className="fixed inset-0 bg-black/95 backdrop-blur-md z-50 flex flex-col items-center justify-center p-4">
          <div className="bg-[#0B0D12] border border-emerald-500/30 p-8 rounded-2xl shadow-2xl flex flex-col items-center justify-center max-w-xs w-full text-center scale-in-center animate-in fade-in duration-300">
            <div className="w-16 h-16 bg-emerald-500/15 rounded-full flex items-center justify-center mb-4 border border-emerald-500/40 animate-bounce">
              <Check size={32} className="text-emerald-400 stroke-[3px]" />
            </div>
            <h3 className="text-xl font-black text-white tracking-wide">
              Baixa realizada
            </h3>
            <p className="text-slate-400 text-xs mt-2 leading-relaxed font-sans">
              O estoque da obra foi atualizado e as alterações foram registradas.
            </p>
          </div>
        </div>
      )}

      {/* OVERLAY DE SUCESSO DA RECONCILIAÇÃO / AJUSTE DE ESTOQUE */}
      {showReconcileSuccessScreen && (
        <div className="fixed inset-0 bg-black/95 backdrop-blur-md z-50 flex flex-col items-center justify-center p-4">
          <div className="bg-[#0B0D12] border border-emerald-500/30 p-8 rounded-2xl shadow-2xl flex flex-col items-center justify-center max-w-xs w-full text-center scale-in-center animate-in fade-in duration-300">
            <div className="w-16 h-16 bg-emerald-500/15 rounded-full flex items-center justify-center mb-4 border border-emerald-500/40 animate-bounce">
              <Check size={32} className="text-emerald-400 stroke-[3px]" />
            </div>
            <h3 className="text-xl font-black text-white tracking-wide">
              Ajuste realizado
            </h3>
            <p className="text-slate-400 text-xs mt-2 leading-relaxed font-sans">
              O estoque da obra foi reconciliado e o lançamento de ajuste foi registrado com sucesso.
            </p>
          </div>
        </div>
      )}

      {/* MODAL 3: INTEGRAR ENTRADA RECEBIDA (RECEIVE DELIVERED CONSIGNMENT) */}
      {showReceiveModal && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-sm z-50 flex items-center justify-center p-4" id="receive-entry-modal">
          <div className="bg-[#161920] border border-slate-700 w-full max-w-xl rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {(() => {
              const currPed = pedidos.find(p => p.id === selectedPedidoId && p.codigo === selectedPedidoCodigo);
              if (!currPed) return null;

              const progressVal = currPed.qtdSolicitada > 0 ? Math.round((currPed.qtdRecebida / currPed.qtdSolicitada) * 100) : 0;
              const additionalVal = currPed.qtdSolicitada > 0 && qtdReceiveInput > 0 ? Math.round((qtdReceiveInput / currPed.qtdSolicitada) * 100) : 0;
              const totalPreviewVal = Math.min(100, progressVal + additionalVal);

              return (
                <>
                  {/* Header */}
                  <div className="p-6 border-b border-slate-800/80 flex justify-between items-center bg-[#171b26]/50">
                    <div>
                      <h3 className="text-md md:text-lg font-bold text-white tracking-tight">RECEBIMENTO DE PEDIDO DE COMPRA</h3>
                      <span className="text-xs text-slate-400 mt-1 block">Vínculo: Pedido #{currPed.id}</span>
                    </div>
                    <button
                      type="button"
                      onClick={() => setShowReceiveModal(false)}
                      className="text-xs font-bold text-slate-400 hover:text-white transition-colors uppercase tracking-wider"
                    >
                      Fechar
                    </button>
                  </div>

                  <form onSubmit={handleConfirmReceipt} className="p-6 space-y-5">
                    
                    {/* Top 2-Column Info Grid */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      {/* Obra Card */}
                      <div className="bg-[#0F1115]/60 border border-slate-800 p-4.5 rounded-xl flex flex-col gap-1.5 shadow-sm min-w-0">
                        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">OBRA DE ABASTECIMENTO</span>
                        <span className="text-sm font-extrabold text-[#ECEEF2] truncate" title={currPed.obra}>{currPed.obra}</span>
                      </div>

                      {/* Insumo Card */}
                      <div className="bg-[#0F1115]/60 border border-slate-800 p-4.5 rounded-xl flex flex-col gap-1.5 shadow-sm min-w-0">
                        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">INSUMO</span>
                        <span className="text-sm font-extrabold text-[#ECEEF2] truncate" title={currPed.insumo}>{currPed.insumo}</span>
                      </div>
                    </div>

                    {/* Numeric Tracking Panel */}
                    <div className="bg-[#0F1115] border border-slate-800 rounded-xl p-4.5 space-y-3">
                      <div className="flex justify-between items-center text-sm">
                        <span className="text-slate-400 font-medium">Total do Pedido de Compra:</span>
                        <span className="text-white font-extrabold">{currPed.qtdSolicitada} {currPed.unidade.toUpperCase()}</span>
                      </div>
                      <div className="flex justify-between items-center text-sm">
                        <span className="text-slate-400 font-medium">Entregue:</span>
                        <span className="text-[#10B981] font-extrabold">{currPed.qtdRecebida} {currPed.unidade.toUpperCase()}</span>
                      </div>

                      <div className="border-t border-slate-800 my-1"></div>

                      <div className="flex justify-between items-center text-sm pt-0.5">
                        <span className="text-slate-350 font-semibold">Saldo Pendente:</span>
                        <span className="text-[#FFC800] font-extrabold">{Math.max(0, currPed.qtdSolicitada - currPed.qtdRecebida)} {currPed.unidade.toUpperCase()}</span>
                      </div>

                      {/* Dynamic filler progress bar */}
                      <div className="space-y-1.5 pt-2">
                        <div className="w-full bg-[#090b0e] rounded-full h-2.5 overflow-hidden border border-slate-900 shadow-inner relative">
                          {/* Highlight background of future prediction */}
                          <div
                            className="h-full bg-indigo-500/40 rounded-full transition-all duration-300 absolute left-0 top-0"
                            style={{ width: `${totalPreviewVal}%` }}
                          ></div>
                          {/* Current verified fill */}
                          <div
                            className="h-full bg-[#059669] rounded-full transition-all duration-200 absolute left-0 top-0"
                            style={{ width: `${progressVal}%` }}
                          ></div>
                        </div>
                        <div className="flex justify-between text-[9px] text-slate-500 font-medium">
                          <span>Entregue: {progressVal}%</span>
                          {additionalVal > 0 && <span className="text-indigo-400 font-semibold">Simulado +{additionalVal}%</span>}
                          <span>Total: {totalPreviewVal}%</span>
                        </div>
                      </div>
                    </div>

                    {/* Quantity input */}
                    <div>
                      <label className="text-xs font-semibold text-slate-300 block mb-2">Quantidade Entregue neste Recebimento: *</label>
                      <div className="flex gap-3">
                        <input
                          type="number"
                          placeholder="Ex: 150"
                          min="1"
                          value={qtdReceiveInput}
                          onChange={(e) => {
                            const inputValStr = e.target.value;
                            if (inputValStr === '') {
                              setQtdReceiveInput('');
                            } else {
                              const pending = currPed.qtdSolicitada - currPed.qtdRecebida;
                              const val = Number(inputValStr);
                              if (val > pending) {
                                setQtdReceiveInput(pending);
                                triggerToast(`A quantidade não pode exceder o saldo pendente de ${pending} ${currPed.unidade.toUpperCase()}`, 'warn');
                              } else if (val < 0) {
                                setQtdReceiveInput(0);
                              } else {
                                setQtdReceiveInput(val);
                              }
                            }
                          }}
                          className="flex-1 bg-[#0F1115] border border-slate-750 rounded-xl p-3 text-sm text-white placeholder-slate-700 font-semibold focus:border-[#4F46E5] focus:ring-1 focus:ring-[#4F46E5]/40 focus:outline-none transition-colors"
                          required
                        />
                        <div className="bg-[#0F1115] border border-slate-750 text-slate-400 text-xs font-bold uppercase rounded-xl px-4 flex items-center justify-center shrink-0 min-w-[55px]">
                          {currPed.unidade.toUpperCase()}
                        </div>
                      </div>
                    </div>

                    {/* Dual-column Nota Fiscal / Responsibility */}
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                      <div>
                        <label className="text-xs font-semibold text-slate-350 block mb-1.5">Nota Fiscal / Documento:</label>
                        <input
                          type="text"
                          placeholder="Ex: NF-12891"
                          value={receiptNotaFiscal}
                          onChange={(e) => setReceiptNotaFiscal(e.target.value)}
                          className="w-full bg-[#0F1115] border border-slate-750 rounded-xl p-3 text-xs text-white placeholder-slate-700 focus:border-[#4F46E5] focus:outline-none transition-colors"
                        />
                      </div>
                      <div>
                        <label className="text-xs font-semibold text-slate-350 block mb-1.5">Responsável de Almox: *</label>
                        <input
                          type="text"
                          placeholder="Ex: Carlos Silva"
                          value={receiptResponsavel}
                          onChange={(e) => setReceiptResponsavel(e.target.value)}
                          className="w-full bg-[#0F1115] border border-slate-750 rounded-xl p-3 text-xs text-white placeholder-slate-700 focus:border-[#4F46E5] focus:outline-none transition-colors"
                          required
                        />
                      </div>
                    </div>

                    {/* Status Dropdown selector */}
                    <div>
                      <label className="text-xs font-semibold text-slate-350 block mb-1.5">Alteração de Status do Pedido:</label>
                      <div className="relative">
                        <select
                          value={receiptStatusAlteracao}
                          onChange={(e) => setReceiptStatusAlteracao(e.target.value)}
                          className="w-full bg-[#0F1115] border border-slate-750 rounded-xl p-3 text-xs text-slate-200 focus:border-[#4F46E5] focus:outline-none appearance-none cursor-pointer pr-10 font-medium"
                        >
                          <option value="deduzir_auto">Dedução Automática pelo Saldo (Entregue/Parcial)</option>
                          <option value="entregue">Forçar Status como Entregue (Manual)</option>
                          <option value="parcial">Forçar Status como Parcial (Manual)</option>
                          <option value="pendente">Forçar Status como Pendente (Manual)</option>
                          <option value="cancelado">Forçar Status como Cancelado (Manual)</option>
                        </select>
                        <div className="absolute right-4 top-1/2 -translate-y-1/2 pointer-events-none text-slate-400 text-xs">
                          ▼
                        </div>
                      </div>
                      {receiptStatusAlteracao === 'deduzir_auto' && (
                        <div className="mt-2 text-[11px] text-slate-400 flex items-center gap-1 bg-[#0F1115] p-2 rounded-lg border border-slate-800">
                          <span>Status previsto (Automático):</span>
                          {(() => {
                            const projectedTotal = currPed.qtdRecebida + Number(qtdReceiveInput || 0);
                            if (projectedTotal >= currPed.qtdSolicitada) {
                              return <span className="text-emerald-400 font-extrabold bg-emerald-500/5 px-2 py-0.5 rounded border border-emerald-500/20">Totalmente Entregue</span>;
                            } else if (projectedTotal > 0) {
                              return <span className="text-amber-400 font-extrabold bg-amber-500/5 px-2 py-0.5 rounded border border-amber-500/20">Parcialmente Entregue</span>;
                            } else {
                              return <span className="text-slate-300 font-extrabold bg-slate-800/40 px-2 py-0.5 rounded border border-slate-700/60">Pendente</span>;
                            }
                          })()}
                        </div>
                      )}
                    </div>

                    {/* Action buttons */}
                    <div className="flex justify-end gap-3 pt-4 border-t border-slate-800/80 mt-6">
                      <button
                        type="button"
                        onClick={() => setShowReceiveModal(false)}
                        className="px-5 py-2.5 bg-slate-800 hover:bg-[#1E2330] text-slate-300 rounded-xl text-xs font-bold transition-all border border-slate-700/50"
                      >
                        Cancelar
                      </button>
                      <button
                        type="submit"
                        className="px-5 py-2.5 bg-[#4F46E5] hover:bg-[#5850EC] text-white rounded-xl text-xs font-bold transition-all shadow-lg shadow-indigo-650/10"
                      >
                        Confirmar Recebimento
                      </button>
                    </div>

                  </form>
                </>
              );
            })()}

          </div>
        </div>
      )}

      {/* SECURITY LOCK / PERMISSION REFUSED MODAL */}
      {lockModalReason && (
        <div className="fixed inset-0 bg-black/90 backdrop-blur-md z-50 flex items-center justify-center p-4">
          <div className="bg-[#1C161D] border border-red-500/30 w-full max-w-sm rounded-xl shadow-2xl p-6 text-center animate-in fade-in zoom-in-95 duration-200">
            <div className="w-14 h-14 bg-red-500/10 border border-red-500/30 rounded-full flex items-center justify-center mx-auto mb-4 text-red-400">
              <Lock size={24} />
            </div>
            <h3 className="text-lg font-sans font-extrabold text-white mb-2">Acesso Negado</h3>
            <p className="text-xs text-slate-400 mb-4 leading-relaxed">
              Você tentou <strong>{lockModalReason}</strong>. Esta operação é restrita para perfis com autonomias específicas autorizadas.
            </p>
            
            <div className="bg-[#120F14] border border-red-500/10 rounded-lg p-3.5 text-left mb-6">
              <span className="text-[10px] font-mono text-slate-500 uppercase block mb-1">Perfil Operador Atual:</span>
              <span className="text-xs font-semibold text-rose-450 flex items-center gap-1.5 font-mono">
                ● {currentUser.nome} l {currentUser.role}
              </span>
              <p className="text-[10px] text-slate-500 mt-2 leading-relaxed">
                Dica de Teste: mude o usuário operador na barra lateral esquerda (Simulador de Perfil) para <strong>"Carlos Roberto (Administrador)"</strong> para liberar e validar todas as ações.
              </p>
            </div>
            
            <button
              onClick={() => setLockModalReason(null)}
              className="w-full bg-slate-800 hover:bg-slate-700 text-white font-semibold py-2 rounded text-sm transition-all border border-slate-700"
            >
              Entendido
            </button>
          </div>
        </div>
      )}

      {/* MODAL 4: EDITAR ESPECIFICAÇÕES DO PEDIDO (ADMIN ONLY) */}
      {showEditPedidoModal && editingPedido && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#161920] border border-slate-700 w-full max-w-md rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <div>
                <h3 className="text-base font-sans font-extrabold text-white">Editar Especificações do Pedido</h3>
                <span className="text-xs text-slate-500 mt-0.5 font-mono">ID: {editingPedido.id}</span>
              </div>
              <button
                type="button"
                onClick={() => setShowEditPedidoModal(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            <form onSubmit={handleSavePedidoEdit} className="p-6 space-y-4">
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Destinar para Obra:</label>
                <select
                  value={editingPedido.obra}
                  onChange={(e) => {
                    const selectedName = e.target.value;
                    const matchedObra = obras.find(o => o.nome === selectedName);
                    setEditingPedido({ 
                      ...editingPedido, 
                      obra: selectedName,
                      obraId: matchedObra ? matchedObra.id : undefined
                    });
                  }}
                  className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-xs rounded-lg p-2.5 focus:border-purple-500 focus:outline-none"
                >
                  {obras.map((o) => (
                    <option key={o.id} value={o.nome}>{(o.id && o.id !== 'undefined') ? `[${o.id}] ${o.nome}` : o.nome}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Nome do Insumo (Destaque):</label>
                <input
                  type="text"
                  value={editingPedido.insumo}
                  onChange={(e) => setEditingPedido({ ...editingPedido, insumo: e.target.value })}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0 font-medium text-purple-200"
                />
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Código do Insumo:</label>
                  <input
                    type="text"
                    value={editingPedido.codigo}
                    onChange={(e) => setEditingPedido({ ...editingPedido, codigo: e.target.value.toUpperCase() })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0 font-mono"
                  />
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Marca / Fabricante:</label>
                  <input
                    type="text"
                    placeholder="Ex: Tigre, Votoran..."
                    value={editingPedido.marca || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, marca: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Fornecedor:</label>
                  <input
                    type="text"
                    placeholder="Nome do Fornecedor..."
                    value={editingPedido.fornecedor || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, fornecedor: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0"
                  />
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Cód. Comprador:</label>
                  <input
                    type="text"
                    placeholder="Ex: COMP-A"
                    value={editingPedido.codComprador || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, codComprador: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0 font-mono"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Cód. Detalhe:</label>
                  <input
                    type="text"
                    placeholder="Ex: DET-CIM"
                    value={editingPedido.codDetalhe || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, codDetalhe: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0 font-mono"
                  />
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Descrição do Detalhe:</label>
                  <input
                    type="text"
                    placeholder="Ex: Secagem Rápida"
                    value={editingPedido.descricaoDetalhe || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, descricaoDetalhe: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Qtd. Solicitada:</label>
                  <input
                    type="number"
                    min="1"
                    value={editingPedido.qtdSolicitada}
                    onChange={(e) => setEditingPedido({ ...editingPedido, qtdSolicitada: Math.max(1, Number(e.target.value)) })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none"
                  />
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Qtd. Recebida:</label>
                  <input
                    type="number"
                    min="0"
                    max={editingPedido.qtdSolicitada}
                    value={editingPedido.qtdRecebida}
                    onChange={(e) => setEditingPedido({ ...editingPedido, qtdRecebida: Math.min(editingPedido.qtdSolicitada, Math.max(0, Number(e.target.value))) })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Unid. Medida (Símbolo):</label>
                  <select
                    value={editingPedido.unidade}
                    onChange={(e) => setEditingPedido({ ...editingPedido, unidade: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-xs rounded-lg p-2.5"
                  >
                    <option value="un">un (unidade)</option>
                    <option value="SC">SC (Saco 50kg)</option>
                    <option value="kg">kg (quilograma)</option>
                    <option value="m³">m³ (metro cúbico)</option>
                    <option value="m²">m² (metro quadrado)</option>
                    <option value="m">m (metro linear)</option>
                  </select>
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5 font-mono">Descrição Unidade:</label>
                  <input
                    type="text"
                    placeholder="Ex: Saco de 50 Quilos"
                    value={editingPedido.descricaoUnidade || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, descricaoUnidade: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Status:</label>
                  <select
                    value={editingPedido.status}
                    onChange={(e) => setEditingPedido({ ...editingPedido, status: e.target.value as any })}
                    className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-xs rounded-lg p-2.5"
                  >
                    <option value="Pendente">Pendente</option>
                    <option value="Parcial">Parcial</option>
                    <option value="Entregue">Entregue</option>
                    <option value="Cancelado">Cancelado</option>
                  </select>
                </div>

                <div>
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Data Pedido:</label>
                  <input
                    type="text"
                    value={editingPedido.dataPedido || ''}
                    onChange={(e) => setEditingPedido({ ...editingPedido, dataPedido: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0 font-mono"
                  />
                </div>
              </div>

              <div className="pt-4 border-t border-slate-800 flex justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setShowEditPedidoModal(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-400 rounded text-xs transition-all"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  className="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded text-xs font-semibold transition-all"
                >
                  Confirmar Alterações
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL 5: CADASTRAR NOVO OPERADOR/USUÁRIO */}
      {newUserOpen && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#161920] border border-slate-700 w-full max-w-sm rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <div>
                <h3 className="text-base font-sans font-extrabold text-white">Cadastrar Novo Colaborador</h3>
                <span className="text-xs text-slate-500 mt-0.5 font-mono">Credenciais Provisórias</span>
              </div>
              <button
                type="button"
                onClick={() => setNewUserOpen(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Nome Completo:</label>
                <input
                  type="text"
                  placeholder="Ex: Engenheiro Roberto..."
                  id="new-user-fullname"
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-550 focus:outline-none placeholder-slate-700"
                />
              </div>
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">E-mail / Login de Acesso:</label>
                <input
                  type="text"
                  placeholder="Ex: roberto@empresa.com.br"
                  id="new-user-loginname"
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-550 focus:outline-none placeholder-slate-700"
                />
              </div>
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Senha Provisória:</label>
                <input
                  type="text"
                  placeholder="Ex: r123..."
                  id="new-user-passcode"
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-205 focus:border-purple-550 focus:outline-none placeholder-slate-705"
                />
              </div>
              <div>
                <label className="text-xs text-slate-400 block mb-1.5 font-sans font-extrabold">Função Atribuída:</label>
                <select
                  id="new-user-role-select"
                  className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-xs rounded-lg p-2.5 focus:border-purple-550 focus:outline-none"
                >
                  <option value="Colaborador">Colaborador (Permissões Customizáveis)</option>
                  <option value="Administrador">Administrador (Acesso Total)</option>
                </select>
              </div>

              <div className="pt-4 border-t border-slate-800 flex justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setNewUserOpen(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-400 rounded text-xs transition-all"
                >
                  Cancelar
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const nameVal = (document.getElementById('new-user-fullname') as HTMLInputElement)?.value || '';
                    const loginVal = (document.getElementById('new-user-loginname') as HTMLInputElement)?.value || '';
                    const passVal = (document.getElementById('new-user-passcode') as HTMLInputElement)?.value || '';
                    const roleVal = (document.getElementById('new-user-role-select') as HTMLSelectElement)?.value as any;
                    handleCreateUsuario(nameVal, loginVal, passVal, roleVal);
                  }}
                  className="px-4 py-2 bg-purple-600 hover:bg-purple-500 text-white rounded text-xs font-semibold transition-all shadow-md"
                >
                  Salvar Cadastro
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* CONFIRMAÇÃO DE APAGAMENTO TOTAL (ZERAR BANCO DE DADOS) */}
      {showClearConfirmModal && (
        <div className="fixed inset-0 bg-black/80 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#161920] border border-red-955/30 w-full max-w-md rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <div className="flex items-center gap-2">
                <AlertTriangle size={18} className="text-rose-500 animate-pulse" />
                <h3 className="text-base font-sans font-extrabold text-white">Confirmar Apagamento Geral</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowClearConfirmModal(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-6 space-y-4">
              <div className="bg-rose-500/10 border border-rose-500/20 p-4 rounded-lg">
                <p className="text-xs text-rose-300 font-medium leading-relaxed">
                  Atenção: Esta ação é irreversível e apagará todos os dados de <strong>Pedidos de Compra</strong> e o histórico de <strong>Baixas de Estoque</strong> salvos localmente.
                </p>
                <p className="text-[11px] text-slate-400 mt-2">
                  Use esta operação para limpar as informações fictícias de teste e importar a planilha de dados da sua própria empresa.
                </p>
              </div>

              <div className="pt-2 flex justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setShowClearConfirmModal(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-400 rounded text-xs font-semibold transition-all"
                >
                  Cancelar, manter dados
                </button>
                <button
                  type="button"
                  onClick={handleClearAllData}
                  className="px-4 py-2 bg-rose-600 hover:bg-rose-500 text-white rounded text-xs font-semibold transition-all flex items-center gap-1.5 shadow-lg shadow-rose-900/20"
                >
                  <Trash2 size={14} />
                  Sim, apagar tudo e zerar
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL: RECONCILIAÇÃO DE SALDO FÍSICO (INVENTÁRIO) */}
      {showReconcileModal && reconcileItem && (
        <div className="fixed inset-0 bg-black/85 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-[#0B0F19] border border-[#232D42] w-full max-w-md rounded-[16px] shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            {/* Header */}
            <div className="p-6 pb-4 flex justify-between items-start">
              <div>
                <h3 className="text-white font-extrabold text-base tracking-wide font-sans">Reconciliação de Estoque Físico</h3>
                <span className="text-xs text-slate-400 font-sans block mt-1">{reconcileItem.obra}</span>
              </div>
              <button
                type="button"
                onClick={() => {
                  setShowReconcileModal(false);
                  setReconcileItem(null);
                }}
                className="text-slate-400 hover:text-white transition-all text-sm font-semibold pt-1"
              >
                Fechar
              </button>
            </div>

            {/* Subtle Divider Line */}
            <div className="border-b border-[#1F2636] mx-6"></div>

            <div className="p-6 space-y-4">
              
              {/* Insumo Box / Framed Details */}
              <div className="border border-[#2D384E] rounded-xl bg-[#111726] p-4 space-y-1.5">
                <div className="text-xs text-slate-400">
                  Insumo: <span className="font-extrabold text-white text-sm block mt-0.5">{reconcileItem.insumo}</span>
                </div>
                <div className="text-xs text-slate-400">
                  Descrição / Detalhe do Insumo: <span className="font-semibold text-slate-300 block mt-0.5">
                    {reconcileItem.codDetalhe || reconcileItem.descricaoDetalhe ? (
                      `${reconcileItem.codDetalhe || ''} ${reconcileItem.descricaoDetalhe || ''}`.trim()
                    ) : (
                      'Sem detalhamento adicional'
                    )}
                  </span>
                </div>
                <div className="text-xs text-slate-400 pt-1 border-t border-[#1F2636]/50 mt-2">
                  Saldo Atual do Sistema: <span className="font-black text-[#FFC800] uppercase font-mono">{reconcileItem.saldo} {reconcileItem.unidade}</span>
                </div>
              </div>

              {/* Input Area 1: Quantidade Física */}
              <div className="space-y-1.5">
                <label className="text-xs text-slate-300 font-bold font-sans block">
                  Novo Saldo Real Contado Fisicamente no Pátio: *
                </label>
                <div className="flex gap-2">
                  <input
                    type="number"
                    step="any"
                    value={reconcileNewPhysicalQty}
                    onChange={(e) => setReconcileNewPhysicalQty(e.target.value === '' ? '' : Number(e.target.value))}
                    className="flex-1 bg-[#07090F] border border-[#232D42] focus:border-[#FFC800] text-white p-2.5 rounded-lg text-sm font-bold font-mono focus:outline-none placeholder-slate-700"
                    placeholder="Ex: 195"
                    required
                  />
                  <div className="bg-[#07090F] border border-[#232D42] text-slate-400 font-extrabold font-mono px-4 rounded-lg flex items-center justify-center text-xs uppercase min-w-[55px] select-none">
                    {reconcileItem.unidade || 'SC'}
                  </div>
                </div>
              </div>

              {/* Input Area 2: Justificativa */}
              <div className="space-y-1.5">
                <label className="text-xs text-slate-300 font-bold font-sans block">
                  Justificativa da Divergência (Quebra, perda ou erro): *
                </label>
                <textarea
                  rows={2.5}
                  placeholder="Ex: Perdas no manuseio de sacarias / Ajuste anual de balanço físico"
                  value={reconcileJustificativa}
                  onChange={(e) => setReconcileJustificativa(e.target.value)}
                  className="w-full bg-[#07090F] border border-[#232D42] focus:border-[#FFC800] text-white text-xs p-2.5 rounded-lg focus:outline-none placeholder-slate-600 resize-none font-sans leading-relaxed"
                  required
                />
              </div>

              {/* Input Area 3: Responsável */}
              <div className="space-y-1.5">
                <label className="text-xs text-slate-300 font-bold font-sans block">
                  Responsável Técnico / Almoxarife: *
                </label>
                <input
                  type="text"
                  placeholder="Ex: Eng. Ricardo Albuquerque"
                  value={reconcileResponsavel}
                  onChange={(e) => setReconcileResponsavel(e.target.value)}
                  className="w-full bg-[#07090F] border border-[#232D42] focus:border-[#FFC800] text-white text-xs p-2.5 rounded-lg focus:outline-none placeholder-slate-600 font-sans"
                  required
                />
              </div>

              {/* Action Buttons */}
              <div className="border-t border-[#1F2636] pt-4 flex gap-3 justify-end font-sans">
                <button
                  type="button"
                  onClick={() => {
                    setShowReconcileModal(false);
                    setReconcileItem(null);
                  }}
                  className="px-5 py-2.5 bg-[#1F2937]/70 hover:bg-[#202938] text-slate-300 rounded-lg text-xs font-extrabold uppercase tracking-wider transition-all"
                >
                  Cancelar
                </button>
                <button
                  type="button"
                  onClick={handleReconcileWeight}
                  className="px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-black uppercase tracking-wider transition-all"
                >
                  Aplicar reajuste
                </button>
              </div>

            </div>
          </div>
        </div>
      )}

    </div>
  );
}

