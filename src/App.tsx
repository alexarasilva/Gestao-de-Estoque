import React, { useState, useMemo, useEffect } from 'react';
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
  Database
} from 'lucide-react';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

// Custom LogoIcon component reflecting the corporate design identity (stylized Delta/A in a yellow circle)
export function LogoIcon({ className = "w-9 h-9" }: { className?: string }) {
  return (
    <svg
      viewBox="0 0 100 100"
      className={`${className} select-none shrink-0`}
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
    >
      <circle cx="50" cy="50" r="48" fill="#FFC800" />
      {/* Stylized upward delta chevron outer contour */}
      <path
        d="M50 19 L73 75 H65 L50 43 L35 75 H27 Z"
        fill="#000000"
      />
      {/* Inner solid upward-pointing triangle */}
      <path
        d="M50 54 L58.3 75 H41.7 Z"
        fill="#000000"
      />
    </svg>
  );
}

// Interfaces
export interface Pedido {
  id: string;
  insumo: string;
  codigo: string;
  obra: string;
  qtdSolicitada: number;
  qtdRecebida: number;
  unidade: string;
  status: 'Pendente' | 'Parcial' | 'Entregue' | 'Cancelado';
  dataPedido: string;
}

export interface Baixa {
  id: string;
  insumo: string;
  codigo: string;
  obra: string;
  quantidade: number;
  unidade: string;
  colaborador: string;
  data: string;
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

export default function App() {
  // Current Navigation View Tab
  const [currentTab, setCurrentTab] = useState<'painel' | 'pedidos' | 'estoque' | 'relatorios' | 'admin'>('painel');

  // Currently Selected Construction Project (Obra Filter)
  const [selectedObra, setSelectedObra] = useState<string>('Todas as Obras');

  // Search filter query
  const [searchQuery, setSearchQuery] = useState<string>('');

  // Table Status Filter
  const [statusFilter, setStatusFilter] = useState<string>('Todos');

  // General Transaction Count (Movimentações Hoje)
  const [movementsCount, setMovementsCount] = useState<number>(142);

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
  const [obras, setObras] = useState<string[]>(() => {
    const saved = localStorage.getItem('construmais_obras');
    return saved ? JSON.parse(saved) : [
      'Residencial Alvorada',
      'Torre Infinito',
      'Complexo Hospitalar'
    ];
  });

  // Backward compatibility bridge
  const listObras = obras;

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
    localStorage.setItem('construmais_obras', JSON.stringify(obras));
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

  // Modals Visibility
  const [showOrderModal, setShowOrderModal] = useState<boolean>(false);
  const [showWithdrawModal, setShowWithdrawModal] = useState<boolean>(false);
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
  const [qtdReceiveInput, setQtdReceiveInput] = useState<number>(0);

  // Sienge excel/spreadsheet import state
  const [showImportModal, setShowImportModal] = useState<boolean>(false);
  const [parsedImportItems, setParsedImportItems] = useState<Pedido[]>([]);
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

  // Toast System
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'warn' | 'info' | null }>({ message: '', type: null });

  const triggerToast = (message: string, type: 'success' | 'warn' | 'info') => {
    setToast({ message, type });
    setTimeout(() => {
      setToast({ message: '', type: null });
    }, 4500);
  };


  // Calculate Net Available Stock for each (Insumo + Obra)
  // Formula: Available = Sum(qtdRecebida from Pedidos with same insumo & obra) - Sum(quantidade from Baixas with same insumo & obra)
  const stockInventory = useMemo(() => {
    const map: { [key: string]: { insumo: string; codigo: string; obra: string; recebido: number; baixado: number; saldo: number; unidade: string } } = {};

    // 1. Gather all received amounts
    pedidos.forEach((p) => {
      if (p.status === 'Cancelado') return;
      const key = `${p.obra}:::${p.codigo || ''}:::${p.insumo}`;
      if (!map[key]) {
        map[key] = { insumo: p.insumo, codigo: p.codigo || '', obra: p.obra, recebido: 0, baixado: 0, saldo: 0, unidade: p.unidade };
      }
      map[key].recebido += p.qtdRecebida;
      map[key].saldo += p.qtdRecebida;
    });

    // 2. Subtract all withdrawn amounts
    baixas.forEach((b) => {
      const key = `${b.obra}:::${b.codigo || ''}:::${b.insumo}`;
      if (!map[key]) {
        map[key] = { insumo: b.insumo, codigo: b.codigo || '', obra: b.obra, recebido: 0, baixado: 0, saldo: 0, unidade: b.unidade };
      }
      map[key].baixado += b.quantidade;
      map[key].saldo -= b.quantidade;
    });

    return Object.values(map);
  }, [pedidos, baixas]);

  // Derived filtered inventories / orders lists
  const filteredPedidos = useMemo(() => {
    return pedidos.filter((p) => {
      const matchObra = selectedObra === 'Todas as Obras' || p.obra === selectedObra;
      const matchSearch =
        p.insumo.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (p.codigo || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.id.toLowerCase().includes(searchQuery.toLowerCase()) ||
        p.obra.toLowerCase().includes(searchQuery.toLowerCase());
      const matchStatus = statusFilter === 'Todos' || p.status === statusFilter;
      return matchObra && matchSearch && matchStatus;
    });
  }, [pedidos, selectedObra, searchQuery, statusFilter]);

  const filteredStock = useMemo(() => {
    return stockInventory.filter((item) => {
      const matchObra = selectedObra === 'Todas as Obras' || item.obra === selectedObra;
      const matchSearch =
        item.insumo.toLowerCase().includes(searchQuery.toLowerCase()) ||
        (item.codigo || '').toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.obra.toLowerCase().includes(searchQuery.toLowerCase());
      return matchObra && matchSearch;
    });
  }, [stockInventory, selectedObra, searchQuery]);

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

  // Distinct list of insumos belonging to a specific construction site for withdrawal dropdown
  const insumosForWithdrawal = useMemo(() => {
    return stockInventory
      .filter((item) => item.obra === newWithdrawObra && item.saldo > 0)
      .map((item) => ({ insumo: item.insumo, saldo: item.saldo, unidade: item.unidade }));
  }, [stockInventory, newWithdrawObra]);


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
      const newPedido: Pedido = {
        id: randomId,
        insumo: newOrderInsumo.trim(),
        codigo: resolvedCodigo,
        obra: newOrderObra,
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
      if (!newWithdrawColab.trim()) {
        triggerToast('Digite o nome do colaborador responsável pela retira.', 'warn');
        return;
      }

      const targetStockItem = stockInventory.find(
        (item) => item.obra === newWithdrawObra && item.insumo === newWithdrawInsumo
      );

      if (!targetStockItem || targetStockItem.saldo < newWithdrawQtd) {
        triggerToast('Saldo insuficiente em estoque local para realizar esta baixa.', 'warn');
        return;
      }

      const randomId = `BX-2024-0${Math.floor(Math.random() * 900) + 100}`;
      const newBaixa: Baixa = {
        id: randomId,
        insumo: newWithdrawInsumo,
        codigo: targetStockItem.codigo || '',
        obra: newWithdrawObra,
        quantidade: newWithdrawQtd,
        unidade: targetStockItem.unidade,
        colaborador: newWithdrawColab.trim(),
        data: new Date().toLocaleString('pt-BR', { hour12: false }).replace(',', '')
      };

      setBaixas([newBaixa, ...baixas]);
      setMovementsCount(prev => prev + 1);
      setShowWithdrawModal(false);
      setNewWithdrawColab('');
      setNewWithdrawInsumo('');
      triggerToast(`Baixa ${randomId} de ${newWithdrawQtd} ${targetStockItem.unidade} de ${newWithdrawInsumo} registrada com sucesso!`, 'success');
    });
  };

  // Prepares the delivery reception modal
  const openReceiveModal = (id: string, codigo: string) => {
    executeGuardedAction('receber_mercadoria', 'Registrar Entrada de Materiais', () => {
      const ped = pedidos.find(p => p.id === id && p.codigo === codigo);
      if (ped) {
        setSelectedPedidoId(id);
        setSelectedPedidoCodigo(codigo);
        setQtdReceiveInput(Number(ped.qtdSolicitada - ped.qtdRecebida));
        setShowReceiveModal(true);
      }
    });
  };

  // Processes delivery check-in
  const handleConfirmReceipt = (e: React.FormEvent) => {
    e.preventDefault();
    if (qtdReceiveInput <= 0) {
      triggerToast('A quantidade recebida deve ser positiva.', 'warn');
      return;
    }

    setPedidos(
      pedidos.map((p) => {
        if (p.id === selectedPedidoId && p.codigo === selectedPedidoCodigo) {
          const newRec = Number(p.qtdRecebida) + Number(qtdReceiveInput);
          let newStatus: 'Pendente' | 'Parcial' | 'Entregue' | 'Cancelado' = 'Parcial';
          if (newRec >= p.qtdSolicitada) {
            newStatus = 'Entregue';
          }
          return {
            ...p,
            qtdRecebida: Math.min(newRec, p.qtdSolicitada), // clamp to target
            status: newStatus
          };
        }
        return p;
      })
    );

    setMovementsCount(prev => prev + 1);
    setShowReceiveModal(false);
    triggerToast(`Entrada registrada com sucesso! Estoque atualizado.`, 'success');
  };

  // Function to process imported Sienge items (UPSERT / MERGE same-pedido items)
  const processImportedPedidos = (importedItems: Pedido[]) => {
    let newOrdersCount = 0;
    let newItemsCount = 0;
    let updatedQuantitiesCount = 0;
    
    let updatedList = [...pedidos];
    const existingOrderIds = new Set(pedidos.map(p => p.id));
    
    importedItems.forEach((newItem) => {
      // Find matching item in current orders list: same order ID and same insumo (or code)
      const exactMatchIndex = updatedList.findIndex(
        (p) => p.id === newItem.id && 
               ((p.codigo && newItem.codigo && p.codigo === newItem.codigo) || 
                p.insumo.toLowerCase().trim() === newItem.insumo.toLowerCase().trim())
      );
      
      if (exactMatchIndex !== -1) {
        // Exists: update details if changed
        const existingItem = updatedList[exactMatchIndex];
        const newQtdSolicitada = newItem.qtdSolicitada;
        const currentQtdRecebida = existingItem.qtdRecebida;
        
        let newStatus = existingItem.status;
        if (newStatus !== 'Cancelado') {
          if (currentQtdRecebida >= newQtdSolicitada) {
            newStatus = 'Entregue';
          } else if (currentQtdRecebida > 0) {
            newStatus = 'Parcial';
          } else {
            newStatus = 'Pendente';
          }
        }
        
        updatedList[exactMatchIndex] = {
          ...existingItem,
          obra: newItem.obra || existingItem.obra,
          unidade: newItem.unidade || existingItem.unidade,
          qtdSolicitada: newQtdSolicitada,
          status: newStatus,
          codigo: newItem.codigo || existingItem.codigo,
          insumo: newItem.insumo || existingItem.insumo,
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
          qtdRecebida: 0,
          status: 'Pendente'
        });
      }
    });
    
    // Register any new construction sites referenced in the spreadsheet
    const uniqueObrasInImport = Array.from(new Set(importedItems.map(r => r.obra)));
    const newObrasToRegister = uniqueObrasInImport.filter(o => o && !obras.includes(o));
    if (newObrasToRegister.length > 0) {
      setObras([...obras, ...newObrasToRegister]);
    }
    
    setPedidos(updatedList);
    setImportStats({
      newOrders: newOrdersCount,
      newItems: newItemsCount,
      updatedQtds: updatedQuantitiesCount
    });
    
    triggerToast(
      `Verificação concluída! ${newOrdersCount} novos pedidos, ${newItemsCount} itens adicionados e ${updatedQuantitiesCount} quantidades atualizadas!`,
      'success'
    );
  };

  // Parse Sienge spreadsheet columns and rows with grouping support
  const parseSiengeSheet = (sheetData: any[][]) => {
    if (sheetData.length <= 1) return [];
    
    let headerRowIndex = 0;
    for (let r = 0; r < Math.min(10, sheetData.length); r++) {
      const row = sheetData[r];
      if (row && row.some(cell => {
        const s = String(cell).toLowerCase();
        return s.includes('pedido') || s.includes('insumo') || s.includes('obra') || s.includes('quantidade');
      })) {
        headerRowIndex = r;
        break;
      }
    }
    
    const headers = sheetData[headerRowIndex].map(h => String(h || '').trim());
    
    let colId = -1, colInsumo = -1, colCodigo = -1, colObra = -1, colQtd = -1, colUnidade = -1, colData = -1;
    
    headers.forEach((h, idx) => {
      const s = h.toLowerCase();
      if (s.includes('pedido') || s.includes('número') || s.includes('numero') || s === 'id' || s === 'pc') colId = idx;
      else if (s.includes('insumo') || s.includes('descri') || s.includes('nome') || s.includes('item') || s.includes('material')) colInsumo = idx;
      else if (s.includes('código') || s.includes('codigo') || s === 'cod' || s === 'cód') colCodigo = idx;
      else if (s.includes('obra') || s.includes('local') || s.includes('filial') || s.includes('empreendimento')) colObra = idx;
      else if (s.includes('solicitada') || s.includes('quantidade') || s.includes('qtd') || s.includes('volume') || s.includes('comprada')) colQtd = idx;
      else if (s.includes('unidade') || s === 'un' || s === 'um' || s === 'medida' || s === 'unid') colUnidade = idx;
      else if (s.includes('data') || s.includes('emissão') || s.includes('emissao') || s.includes('compra')) colData = idx;
    });
    
    if (colId === -1) colId = headers.findIndex(h => /id|num/i.test(h)) !== -1 ? headers.findIndex(h => /id|num/i.test(h)) : 0;
    if (colInsumo === -1) colInsumo = headers.findIndex(h => /ins|des|nom|itm/i.test(h)) !== -1 ? headers.findIndex(h => /ins|des|nom|itm/i.test(h)) : 1;
    if (colCodigo === -1) colCodigo = headers.findIndex(h => /cod/i.test(h)) !== -1 ? headers.findIndex(h => /cod/i.test(h)) : 2;
    if (colObra === -1) colObra = headers.findIndex(h => /obr|loc|fil/i.test(h)) !== -1 ? headers.findIndex(h => /obr|loc|fil/i.test(h)) : 3;
    if (colQtd === -1) colQtd = headers.findIndex(h => /qtd|sol|qua/i.test(h)) !== -1 ? headers.findIndex(h => /qtd|sol|qua/i.test(h)) : 4;
    if (colUnidade === -1) colUnidade = headers.findIndex(h => /un|um|med/i.test(h)) !== -1 ? headers.findIndex(h => /un|um|med/i.test(h)) : 5;
    if (colData === -1) colData = headers.findIndex(h => /dat|emi/i.test(h)) !== -1 ? headers.findIndex(h => /dat|emi/i.test(h)) : 6;
    
    const results: Pedido[] = [];
    let lastOrderId = '';
    let lastObra = listObras[0] || 'Residencial Alvorada';
    let lastData = new Date().toLocaleDateString('pt-BR');
    
    for (let r = headerRowIndex + 1; r < sheetData.length; r++) {
      const row = sheetData[r];
      if (!row || row.length === 0 || row.every(cell => cell === null || cell === '')) continue;
      
      let rawId = colId !== -1 && row[colId] !== undefined ? String(row[colId]).trim() : '';
      let rawInsumo = colInsumo !== -1 && row[colInsumo] !== undefined ? String(row[colInsumo]).trim() : '';
      let rawCodigo = colCodigo !== -1 && row[colCodigo] !== undefined ? String(row[colCodigo]).trim() : '';
      let rawObra = colObra !== -1 && row[colObra] !== undefined ? String(row[colObra]).trim() : '';
      let rawQtd = colQtd !== -1 && row[colQtd] !== undefined ? parseFloat(String(row[colQtd]).replace(/[^\d.-]/g, '')) : 0;
      let rawUnidade = colUnidade !== -1 && row[colUnidade] !== undefined ? String(row[colUnidade]).trim() : 'un';
      let rawData = colData !== -1 && row[colData] !== undefined ? String(row[colData]).trim() : '';
      
      if (!rawInsumo || rawInsumo.toLowerCase() === 'insumo' || rawInsumo.toLowerCase() === 'descrição') continue;
      
      if (rawId) lastOrderId = rawId;
      if (rawObra) lastObra = rawObra;
      if (rawData) lastData = rawData;
      
      const finalId = lastOrderId || 'PC-IMPORTADO';
      const finalObra = lastObra || 'Todas as Obras';
      const finalData = lastData || new Date().toLocaleDateString('pt-BR');
      
      let finalCodigo = rawCodigo;
      if (!finalCodigo) {
        const prefix = rawInsumo.substring(0, 3).toUpperCase().replace(/[^A-Z]/g, 'GEN');
        finalCodigo = `${prefix}-${Math.floor(100 + Math.random() * 900)}`;
      }
      
      results.push({
        id: finalId,
        insumo: rawInsumo,
        codigo: finalCodigo,
        obra: finalObra,
        qtdSolicitada: isNaN(rawQtd) || rawQtd <= 0 ? 1 : rawQtd,
        qtdRecebida: 0,
        unidade: rawUnidade || 'un',
        status: 'Pendente',
        dataPedido: finalData
      });
    }
    
    return results;
  };

  const handleUploadSienge = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        const parsed = parseSiengeSheet(jsonData);
        if (parsed.length === 0) {
          triggerToast('Nenhum item válido foi encontrado na planilha Sienge.', 'warn');
        } else {
          setParsedImportItems(parsed);
          triggerToast(`${parsed.length} itens detectados e prontos para verificação!`, 'success');
        }
      } catch (err: any) {
        console.error(err);
        triggerToast('Houve um erro ao ler o arquivo Excel. Verifique o formato.', 'warn');
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleLoadDemoSienge = () => {
    const demoItems: Pedido[] = [
      { id: 'PC-2024-0891', insumo: 'Cimento CP-II 50kg', codigo: 'CIM-001', obra: 'Residencial Alvorada', qtdSolicitada: 750, qtdRecebida: 0, unidade: 'un', status: 'Pendente', dataPedido: '16/06/2026' },
      { id: 'PC-2024-0891', insumo: 'Torneira Esfera Brass 1/2', codigo: 'TOR-009', obra: 'Residencial Alvorada', qtdSolicitada: 45, qtdRecebida: 0, unidade: 'un', status: 'Pendente', dataPedido: '16/06/2026' },
      { id: 'PC-2026-1010', insumo: 'Gesso Acartonado Tabica 3m', codigo: 'GES-012', obra: 'Residencial Alvorada', qtdSolicitada: 120, qtdRecebida: 0, unidade: 'un', status: 'Pendente', dataPedido: '16/06/2026' },
      { id: 'PC-2026-1010', insumo: 'Parafuso Drywall GN 25', codigo: 'PAR-013', obra: 'Residencial Alvorada', qtdSolicitada: 2000, qtdRecebida: 0, unidade: 'un', status: 'Pendente', dataPedido: '16/06/2026' },
    ];
    setParsedImportItems(demoItems);
    triggerToast('Planilha modelo Sienge (4 itens) carregada para testes!', 'info');
  };

  // Synchronizes state with Google Sheets App Script URL
  const handleSheetsSync = async (direction: 'pull' | 'push') => {
    if (!googleSheetsUrl) {
      triggerToast('Insira a URL do Apps Script primeiro para realizar a sincronização.', 'warn');
      return;
    }
    
    setIsSyncing(true);
    try {
      if (direction === 'pull') {
        const response = await fetch(`${googleSheetsUrl}?action=getData`);
        if (!response.ok) throw new Error('Não foi possível conectar ao Google Sheets.');
        const result = await response.json();
        
        if (result.status === 'success') {
          if (result.pedidos) setPedidos(result.pedidos);
          if (result.baixas) setBaixas(result.baixas);
          if (result.obras) setObras(result.obras);
          if (result.usuarios) setUsuarios(result.usuarios);
          triggerToast('Dados sincronizados da planilha com sucesso!', 'success');
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
        
        triggerToast('Dados enviados para a planilha com sucesso! (Deploy atualizado)', 'success');
      }
    } catch (err: any) {
      console.error(err);
      triggerToast(`Erro na sincronização: ${err.message || 'Verifique se a URL já foi publicada para acesso "Qualquer pessoa/Anyone".'}`, 'warn');
    } finally {
      setIsSyncing(false);
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
  const handleCreateObra = (name: string) => {
    executeGuardedAction('admin', 'Criar Nova Instalação/Obra', () => {
      const trimmed = name.trim();
      if (!trimmed) {
        triggerToast("Nome da obra não pode estar vazio.", "warn");
        return;
      }
      if (obras.map(o => o.toLowerCase()).includes(trimmed.toLowerCase())) {
        triggerToast("Esta obra já está cadastrada no sistema.", "warn");
        return;
      }
      setObras([...obras, trimmed]);
      setNewObraName('');
      setShowAddObraModal(false);
      triggerToast(`Obra "${trimmed}" cadastrada com êxito!`, "success");
    });
  };

  const handleRenameObra = (oldName: string, newName: string) => {
    executeGuardedAction('admin', 'Editar/Renomear Obra', () => {
      const trimmedNew = newName.trim();
      if (!trimmedNew) {
        triggerToast("O novo nome não pode estar vazio.", "warn");
        return;
      }
      if (obras.map(o => o.toLowerCase()).includes(trimmedNew.toLowerCase()) && trimmedNew.toLowerCase() !== oldName.toLowerCase()) {
        triggerToast("Este nome de obra já está em uso.", "warn");
        return;
      }
      
      // Update obras array
      setObras(obras.map(o => o === oldName ? trimmedNew : o));
      
      // Cascade update pedidos and baixas
      setPedidos(pedidos.map(p => p.obra === oldName ? { ...p, obra: trimmedNew } : p));
      setBaixas(baixas.map(b => b.obra === oldName ? { ...b, obra: trimmedNew } : b));
      
      // Keep UI filters aligned
      if (selectedObra === oldName) setSelectedObra(trimmedNew);
      if (newOrderObra === oldName) setNewOrderObra(trimmedNew);
      if (newWithdrawObra === oldName) setNewWithdrawObra(trimmedNew);
      
      setEditingObraName(null);
      triggerToast(`Obra renomeada para "${trimmedNew}"!`, "success");
    });
  };

  const handleDeleteObra = (obraName: string) => {
    executeGuardedAction('excluir_obra', `Excluir a Obra "${obraName}"`, () => {
      if (obras.length <= 1) {
        triggerToast("O sistema exige no mínimo uma obra ativa cadastrada.", "warn");
        return;
      }
      
      setObras(obras.filter(o => o !== obraName));
      
      if (selectedObra === obraName) setSelectedObra("Todas as Obras");
      if (newOrderObra === obraName) setNewOrderObra(obras.filter(o => o !== obraName)[0]);
      if (newWithdrawObra === obraName) setNewWithdrawObra(obras.filter(o => o !== obraName)[0]);
      
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

    setPedidos(pedidos.map(p => p.id === editingPedido.id ? {
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

  // Mock Export CSV logic
  const handleExportCSV = () => {
    const csvContent = "ID,Insumo,Obra,Solicitado,Recebido,Unidade,Status\n" + 
      pedidos.map(p => `${p.id},${p.insumo},${p.obra},${p.qtdSolicitada},${p.qtdRecebida},${p.unidade},${p.status}`).join("\n");
    
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", `relatorio_estoque_${selectedObra.replace(/\s+/g, '_')}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    triggerToast('Relatório exportado com sucesso.', 'success');
  };

  return (
    <div className="flex h-screen w-full bg-[#0F1115] font-sans text-slate-300 overflow-hidden select-none">
      
      {/* Toast Notification Bar */}
      {toast.message && (
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
      )}

      {/* Sidebar Navigation */}
      <aside id="sidebar-panel" className="w-[260px] bg-[#161920] border-r border-slate-800 flex flex-col justify-between shrink-0">
        <div className="p-6">
          {/* Logo / Branding */}
          <div className="flex items-center gap-2.5 mb-6">
            <LogoIcon className="w-9 h-9 shadow-lg shadow-yellow-500/10" />
            <div>
              <h1 className="text-lg font-serif font-semibold tracking-tight text-white leading-tight">GestãoEstoque</h1>
              <p className="text-[10px] text-slate-500 font-mono tracking-widest uppercase">Controle local</p>
            </div>
          </div>

          {/* User Simulator Switcher Box */}
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

          {/* Navigation Links */}
          <nav className="space-y-1.5" id="navigation-root">
            <button
              id="nav-tab-painel"
              onClick={() => setCurrentTab('painel')}
              className={`w-full flex items-center justify-between px-4 py-2.5 rounded-md text-sm font-medium transition-all ${
                currentTab === 'painel'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
            >
              <span className="flex items-center gap-3">
                <LayoutDashboard size={18} />
                Painel Geral
              </span>
              {currentTab === 'painel' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
            </button>

            <button
              id="nav-tab-pedidos"
              onClick={() => setCurrentTab('pedidos')}
              className={`w-full flex items-center justify-between px-4 py-2.5 rounded-md text-sm font-medium transition-all ${
                currentTab === 'pedidos'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
            >
              <span className="flex items-center gap-3">
                <ClipboardList size={18} />
                Pedidos de Compra
              </span>
              {currentTab === 'pedidos' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
            </button>

            <button
              id="nav-tab-estoque"
              onClick={() => setCurrentTab('estoque')}
              className={`w-full flex items-center justify-between px-4 py-2.5 rounded-md text-sm font-medium transition-all ${
                currentTab === 'estoque'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
            >
              <span className="flex items-center gap-3">
                <BaggageClaim size={18} />
                Estoque por Obra
              </span>
              {currentTab === 'estoque' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
            </button>

            <button
              id="nav-tab-relatorios"
              onClick={() => setCurrentTab('relatorios')}
              className={`w-full flex items-center justify-between px-4 py-2.5 rounded-md text-sm font-medium transition-all ${
                currentTab === 'relatorios'
                  ? 'bg-emerald-600/10 text-emerald-400 border border-emerald-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
            >
              <span className="flex items-center gap-3">
                <BarChart3 size={18} />
                Relatórios e Baixas
              </span>
              {currentTab === 'relatorios' && <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full"></span>}
            </button>

            <button
              id="nav-tab-admin"
              onClick={() => {
                executeGuardedAction('admin', 'Acessar painel administrativo', () => {
                  setCurrentTab('admin');
                });
              }}
              className={`w-full flex items-center justify-between px-4 py-2.5 rounded-md text-sm font-medium transition-all ${
                currentTab === 'admin'
                  ? 'bg-purple-600/10 text-purple-400 border border-purple-500/20'
                  : 'text-slate-400 hover:text-white hover:bg-slate-800'
              }`}
            >
              <span className="flex items-center gap-3">
                <Shield size={18} className={currentUser.role === 'Administrador' ? 'text-purple-400' : 'text-slate-500'} />
                Administração
              </span>
              {currentUser.role !== 'Administrador' ? (
                <Lock size={12} className="text-slate-500" />
              ) : currentTab === 'admin' ? (
                <span className="w-1.5 h-1.5 bg-purple-400 rounded-full"></span>
              ) : null}
            </button>
          </nav>
        </div>

        {/* User profile / Footer */}
        <div id="sidebar-footer" className="p-5 border-t border-slate-800 bg-[#12151B]">
          <div className="flex items-center gap-3">
            <div className="w-9 h-9 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center text-xs font-bold text-yellow-500 uppercase">
              {currentUser.nome.substring(0, 2)}
            </div>
            <div className="flex-1 min-w-0">
              <p className="text-xs font-semibold text-white truncate">{currentUser.nome}</p>
              <p className="text-[10px] text-slate-500 truncate">{currentUser.login}</p>
            </div>
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
            
            <span className="text-slate-700 font-serif italic">|</span>
            <div className="text-[10px] text-slate-500 font-medium uppercase tracking-widest hidden sm:block">
              Base atualizada hoje às 08:30 • Movimentações: {movementsCount}
            </div>
          </div>

          <div className="flex gap-2.5">
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
                    <p className="text-3xl font-serif text-amber-500 transition-all font-semibold" id="stat-pedidos-pendentes">{stats.pendentes}</p>
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
                    <p className="text-3xl font-serif text-blue-400 font-semibold" id="stat-pedidos-parciais">{stats.parciais}</p>
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
                    <p className="text-3xl font-serif text-rose-500 font-semibold" id="stat-estoque-critico">{stats.estoqueCritico}</p>
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
                    <p className="text-3xl font-serif text-emerald-400 font-semibold" id="stat-movimentos-hoje">{movementsCount}</p>
                    <span className="text-[10px] text-slate-500 mt-1 block">Atividades registradas</span>
                  </div>
                  <div className="p-3 bg-emerald-50/5 text-emerald-400 rounded-lg group-hover:scale-105 transition-all">
                    <TrendingDown size={21} className="rotate-180" />
                  </div>
                  <div className="absolute bottom-0 left-0 right-0 h-[2px] bg-emerald-500/20"></div>
                </div>
              </div>

              {/* Main Dashboard Interactive Table */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl overflow-hidden flex flex-col" id="dashboard-table-container">
                
                {/* Search and Filters Header */}
                <div className="px-6 py-4.5 border-b border-slate-800 flex flex-col sm:flex-row gap-3 justify-between items-center bg-[#1C2028]">
                  <div>
                    <h2 className="text-base font-serif font-semibold text-white">Acompanhamento de Pedidos</h2>
                    <p className="text-[11px] text-slate-500 mt-0.5">Foque na entrada, verificação e atualização dos insumos contratados</p>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                    {/* Search Field */}
                    <div className="relative w-full sm:w-60">
                      <input
                        id="search-input"
                        type="text"
                        placeholder="Filtrar insumo ou pedido..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full bg-[#0F1115] border border-slate-700 rounded-md py-1.5 pl-8 pr-3 text-sm text-slate-300 placeholder-slate-500 focus:border-emerald-500 focus:outline-none focus:ring-1 focus:ring-emerald-500"
                      />
                      <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-500" />
                    </div>

                    {/* Status Filter Tab Group */}
                    <div className="flex bg-[#0F1115] p-1 rounded-md border border-slate-700 text-xs text-slate-400">
                      {['Todos', 'Pendente', 'Parcial', 'Entregue'].map((st) => (
                        <button
                          key={st}
                          onClick={() => setStatusFilter(st)}
                          className={`px-2.5 py-1 rounded transition-colors ${
                            statusFilter === st ? 'bg-emerald-600/20 text-emerald-400 font-semibold' : 'hover:text-slate-200'
                          }`}
                        >
                          {st}
                        </button>
                      ))}
                    </div>

                    {/* Report Download Trigger */}
                    <button
                      id="btn-export-csv"
                      onClick={handleExportCSV}
                      title="Exportar CSV"
                      className="p-1.5 bg-slate-800 hover:bg-slate-700 border border-slate-700 rounded-md text-slate-400 hover:text-white transition-all shrink-0"
                    >
                      <Download size={16} />
                    </button>
                  </div>
                </div>

                {/* Database Table Rendering */}
                <div className="overflow-x-auto">
                  <table className="w-full text-left text-sm border-collapse" id="orders-dashboard-table">
                    <thead>
                      <tr className="text-slate-400 font-medium border-b border-slate-800 bg-[#12151C] text-xs">
                        <th className="px-6 py-3.5">ID Pedido</th>
                        <th className="px-6 py-3.5">Obra / Insumo</th>
                        <th className="px-6 py-3.5 text-right">Qtd. Solicitada</th>
                        <th className="px-6 py-3.5 text-right">Qtd. Recebida</th>
                        <th className="px-6 py-3.5">Progresso</th>
                        <th className="px-6 py-3.5">Status</th>
                        <th className="px-6 py-3.5 text-center">Ações</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/80">
                      {filteredPedidos.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="text-center py-10 text-slate-500">
                            <Layers size={36} className="mx-auto text-slate-700 mb-2.5" />
                            Nenhum pedido atende aos filtros definidos.
                          </td>
                        </tr>
                      ) : (
                        filteredPedidos.map((ped) => {
                          const percent = Math.min(
                            100,
                            Math.round((ped.qtdRecebida / ped.qtdSolicitada) * 100)
                          );
                          return (
                            <tr key={ped.id} className="hover:bg-slate-800/20 transition-colors">
                              {/* Order ID */}
                              <td className="px-6 py-4 font-mono text-xs text-slate-500 font-bold">
                                {ped.id}
                              </td>

                              {/* Material name & Construction Site */}
                              <td className="px-6 py-4">
                                <div className="text-white font-medium text-sm">{ped.insumo}</div>
                                <div className="text-xs text-slate-500 flex items-center gap-1 mt-0.5">
                                  <MapPin size={10} className="text-slate-600" />
                                  {ped.obra}
                                </div>
                              </td>

                              {/* Requested Quantity */}
                              <td className="px-6 py-4 text-right font-medium text-slate-300">
                                {ped.qtdSolicitada.toLocaleString('pt-BR')} {ped.unidade}
                              </td>

                              {/* Received Quantity */}
                              <td className="px-6 py-4 text-right">
                                <span className={`font-semibold ${ped.qtdRecebida > 0 ? 'text-emerald-400' : 'text-slate-500'}`}>
                                  {ped.qtdRecebida.toLocaleString('pt-BR')}
                                </span>
                                <span className="text-xs text-slate-500"> /{ped.unidade}</span>
                              </td>

                              {/* Small Meter bar */}
                              <td className="px-6 py-4 min-w-[120px]">
                                <div className="flex items-center gap-2">
                                  <div className="flex-1 bg-slate-900 rounded-full h-1.5 overflow-hidden">
                                    <div
                                      className={`h-full rounded-full ${
                                        ped.status === 'Entregue'
                                          ? 'bg-emerald-500'
                                          : ped.status === 'Parcial'
                                          ? 'bg-blue-400'
                                          : ped.status === 'Cancelado'
                                          ? 'bg-red-500'
                                          : 'bg-amber-500/40'
                                      }`}
                                      style={{ width: `${percent}%` }}
                                    ></div>
                                  </div>
                                  <span className="text-[10px] font-mono text-slate-400 w-7 text-right">
                                    {percent}%
                                  </span>
                                </div>
                              </td>

                              {/* Status Badge */}
                              <td className="px-6 py-4">
                                <span
                                  className={`inline-block px-2.5 py-0.5 rounded text-[10px] uppercase font-bold tracking-wider border ${
                                    ped.status === 'Entregue'
                                      ? 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20'
                                      : ped.status === 'Parcial'
                                      ? 'bg-blue-500/10 text-blue-400 border-blue-500/20'
                                      : ped.status === 'Cancelado'
                                      ? 'bg-rose-500/10 text-rose-400 border-rose-500/20'
                                      : 'bg-amber-500/10 text-amber-500 border-amber-500/20'
                                  }`}
                                >
                                  {ped.status === 'Entregue' ? 'Entregue' : ped.status === 'Parcial' ? 'Parcial' : ped.status === 'Cancelado' ? 'Cancelado' : 'Pendente'}
                                </span>
                              </td>

                              {/* Action Call */}
                              <td className="px-6 py-4 text-center">
                                <div className="flex items-center justify-center gap-2">
                                  {ped.status !== 'Entregue' && ped.status !== 'Cancelado' ? (
                                    <button
                                      onClick={() => openReceiveModal(ped.id, ped.codigo || '')}
                                      className="px-2.5 py-1 text-xs bg-emerald-600/10 hover:bg-emerald-600 text-emerald-400 hover:text-white rounded border border-emerald-500/30 transition-all font-semibold"
                                      title="Dar Entrada Física"
                                    >
                                      Entrada
                                    </button>
                                  ) : (
                                    <span className="text-xs text-slate-500">Concluído</span>
                                  )}
                                  
                                  {/* Administrative Pencil Action */}
                                  <button
                                    onClick={() => handleOpenEditPedidoModal(ped)}
                                    className="p-1 text-slate-400 hover:text-white hover:bg-slate-800 rounded transition-all"
                                    title="Editar Especificações"
                                  >
                                    <Edit size={14} />
                                  </button>
                                  
                                  {/* Administrative Trash Action */}
                                  <button
                                    onClick={() => handleDeletePedido(ped.id)}
                                    className="p-1 text-rose-550/70 hover:text-rose-400 hover:bg-slate-800 rounded transition-all"
                                    title="Excluir Definitivamente"
                                  >
                                    <Trash2 size={14} />
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

                {/* Footer and summary of data table */}
                <div className="p-4.5 border-t border-slate-800 bg-[#161920] flex flex-col sm:flex-row gap-4 items-center justify-between text-xs text-slate-500">
                  <div>
                    Exibindo <span className="text-slate-300 font-semibold">{filteredPedidos.length}</span> de <span className="text-slate-300 font-semibold">{pedidos.length}</span> registros totais
                  </div>
                  
                  <div className="flex flex-col sm:flex-row gap-4 items-center">
                    <span className="font-medium text-slate-400">
                      Filtro Atual:{' '}
                      <span className="text-white ml-1 font-semibold underline decoration-emerald-500">
                        {selectedObra}
                      </span>
                    </span>
                    
                    <div className="flex gap-1" id="pagination-set">
                      <button className="w-7 h-7 flex items-center justify-center rounded border border-slate-700 hover:bg-slate-800 transition-all cursor-not-allowed">
                        ‹
                      </button>
                      <button className="w-7 h-7 flex items-center justify-center rounded border border-emerald-500/40 bg-emerald-500/10 text-emerald-400 font-bold">
                        1
                      </button>
                      <button className="w-7 h-7 flex items-center justify-center rounded border border-slate-700 hover:bg-slate-800 transition-all text-slate-400">
                        2
                      </button>
                      <button className="w-7 h-7 flex items-center justify-center rounded border border-slate-700 hover:bg-slate-800 transition-all">
                        ›
                      </button>
                    </div>
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
                  <h2 className="text-xl font-serif font-semibold text-white">Gestão Integrada de Pedidos</h2>
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
                      setImportStats(null);
                      setShowImportModal(true);
                    }}
                    className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-705 rounded-md text-sm font-semibold transition-all flex items-center gap-2"
                  >
                    <FileSpreadsheet size={16} className="text-emerald-450" />
                    Importar Sienge Sincronizado
                  </button>
                </div>
              </div>

              {/* Status breakdown metrics */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl p-6 flex flex-col md:flex-row justify-around gap-6 divide-y md:divide-y-0 md:divide-x divide-slate-800">
                <div className="flex-1 pb-4 md:pb-0 md:pr-4 text-center">
                  <p className="text-[11px] text-slate-400 uppercase font-bold tracking-wider mb-1">Total Solicitados</p>
                  <p className="text-2xl font-serif text-white">{pedidos.length} Pedidos</p>
                </div>
                <div className="flex-1 py-4 md:py-0 md:px-4 text-center">
                  <p className="text-[11px] text-amber-400 uppercase font-bold tracking-wider mb-1">Pendentes de Entrega</p>
                  <p className="text-2xl font-serif text-amber-500 font-semibold">
                    {pedidos.filter(p => p.status === 'Pendente').length} Pedidos
                  </p>
                </div>
                <div className="flex-1 py-4 md:py-0 md:px-4 text-center">
                  <p className="text-[11px] text-blue-400 uppercase font-bold tracking-wider mb-1">Entregas Parciais</p>
                  <p className="text-2xl font-serif text-blue-400 font-semibold">
                    {pedidos.filter(p => p.status === 'Parcial').length} Pedidos
                  </p>
                </div>
                <div className="flex-1 pt-4 md:pt-0 md:pl-4 text-center">
                  <p className="text-[11px] text-emerald-400 uppercase font-bold tracking-wider mb-1 font-semibold">Entregue Totalmente</p>
                  <p className="text-2xl font-serif text-emerald-400 font-semibold">
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

                <div className="flex bg-[#0F1115] p-1 rounded-md border border-slate-700 text-xs text-slate-400">
                  {['Todos', 'Pendente', 'Parcial', 'Entregue'].map((st) => (
                    <button
                      key={st}
                      onClick={() => setStatusFilter(st)}
                      className={`px-2.5 py-1 rounded transition-colors ${
                        statusFilter === st ? 'bg-emerald-600/20 text-emerald-400 font-semibold' : 'hover:text-slate-200'
                      }`}
                    >
                      {st}
                    </button>
                  ))}
                </div>
              </div>

              {/* Advanced Orders List Cards */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {filteredPedidos.length === 0 ? (
                  <div className="col-span-full text-center py-12 bg-[#161920] border border-slate-800 rounded-xl text-slate-500">
                    <Layers size={36} className="mx-auto text-slate-700 mb-2.5" />
                    Nenhum pedido atende aos filtros definidos nesta aba.
                  </div>
                ) : (
                  filteredPedidos.map((p) => {
                    const progressVal = Math.round((p.qtdRecebida / p.qtdSolicitada) * 100);
                    return (
                      <div key={p.id} className="bg-[#161920] border border-slate-800 rounded-xl p-5 hover:border-slate-700 transition-all flex flex-col justify-between">
                        
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

                          <h4 className="text-base font-semibold text-white mb-1">{p.insumo}</h4>
                          <p className="text-xs text-slate-400 flex items-center gap-1 mb-4">
                            <MapPin size={12} className="text-slate-500" />
                            {p.obra}
                          </p>
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
                                      setPedidos(pedidos.map(item => item.id === p.id ? { ...item, status: 'Cancelado' as const } : item));
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
                                    Dar Entrada
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
            </div>
          )}

          {/* VIEW TAB 3: ESTOQUE POR OBRA */}
          {currentTab === 'estoque' && (
            <div className="space-y-6" id="estoque-tab-view">
              <div className="flex items-center justify-between">
                <div>
                  <h2 className="text-xl font-serif font-semibold text-white">Disponibilidade e Saldos de Estoque</h2>
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
                      <tr className="text-slate-450 font-medium border-b border-slate-800 bg-[#12151C] text-xs uppercase tracking-wider text-slate-400">
                        <th className="px-6 py-3.5">Código</th>
                        <th className="px-6 py-3.5">Nome do Insumo / Obra</th>
                        <th className="px-6 py-3.5 text-right">Saldo Disponível</th>
                        <th className="px-6 py-3.5 text-right">Recebido Total</th>
                        <th className="px-6 py-3.5 text-right">Utilizado (Baixas)</th>
                        <th className="px-6 py-3.5 text-center min-w-[120px]">Consumo</th>
                        <th className="px-6 py-3.5 text-center">Ações</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-800/85">
                      {filteredStock.length === 0 ? (
                        <tr>
                          <td colSpan={7} className="text-center py-12 text-slate-500">
                            <Layers size={36} className="mx-auto text-slate-700 mb-2.5" />
                            Nenhum insumo encontrado no estoque com os filtros aplicados.
                          </td>
                        </tr>
                      ) : (
                        filteredStock.map((item, idx) => {
                          const percent = Math.min(100, Math.round((item.baixado / (item.recebido || 1)) * 100));
                          const isCritical = item.recebido > 0 && item.saldo <= item.recebido * 0.15;
                          
                          return (
                            <tr key={idx} className="hover:bg-slate-800/20 transition-colors">
                              {/* Código do Insumo */}
                              <td className="px-6 py-4 font-mono text-xs font-semibold">
                                <span className="bg-zinc-800/80 text-yellow-500 px-2.5 py-1 rounded border border-yellow-500/10 shadow-sm">
                                  {item.codigo || '-'}
                                </span>
                              </td>

                              {/* Insumo Name and Site */}
                              <td className="px-6 py-4">
                                <div className="text-white font-medium text-sm">{item.insumo}</div>
                                <div className="text-xs text-slate-500 flex items-center gap-1 mt-0.5">
                                  <MapPin size={10} className="text-slate-600" />
                                  {item.obra}
                                </div>
                              </td>

                              {/* Saldo Disponível */}
                              <td className="px-6 py-4 text-right font-serif">
                                <span className={`text-sm font-semibold tracking-wide ${
                                  item.saldo <= 0 ? 'text-red-400' :
                                  isCritical ? 'text-rose-400' : 'text-emerald-400'
                                }`}>
                                  {item.saldo.toLocaleString('pt-BR')}
                                </span>
                                <span className="text-xs text-slate-500 ml-1 font-sans">{item.unidade}</span>
                                {isCritical && item.saldo > 0 && (
                                  <span className="block text-[9px] uppercase tracking-wider text-rose-400 font-sans mt-0.5 font-bold">Estoque Crítico</span>
                                )}
                              </td>

                              {/* Recebido Total */}
                              <td className="px-6 py-4 text-right text-slate-300 font-medium">
                                {item.recebido.toLocaleString('pt-BR')} <span className="text-xs text-slate-500">{item.unidade}</span>
                              </td>

                              {/* Utilizado (Baixas) */}
                              <td className="px-6 py-4 text-right text-orange-400/90 font-medium">
                                {item.baixado.toLocaleString('pt-BR')} <span className="text-xs text-slate-500">{item.unidade}</span>
                              </td>

                              {/* Consumo Bar */}
                              <td className="px-6 py-4">
                                <div className="flex items-center gap-2 max-w-[150px] mx-auto">
                                  <div className="flex-1 bg-slate-900 rounded-full h-1.5 overflow-hidden">
                                    <div
                                      className={`h-full rounded-full ${
                                        percent >= 90 ? 'bg-rose-500' :
                                        percent >= 50 ? 'bg-amber-500' : 'bg-blue-400'
                                      }`}
                                      style={{ width: `${percent}%` }}
                                    ></div>
                                  </div>
                                  <span className="text-[10px] font-mono text-slate-400 w-8 text-right">
                                    {percent}%
                                  </span>
                                </div>
                              </td>

                              {/* Ações */}
                              <td className="px-6 py-4 text-center">
                                {item.saldo > 0 ? (
                                  <button
                                    onClick={() => {
                                      setNewWithdrawObra(item.obra);
                                      setNewWithdrawInsumo(item.insumo);
                                      setNewWithdrawQtd(Math.min(10, item.saldo));
                                      setShowWithdrawModal(true);
                                    }}
                                    className="px-3 py-1 bg-amber-600 hover:bg-amber-500 text-slate-900 font-bold rounded text-xs transition-colors flex items-center gap-1.5 mx-auto"
                                  >
                                    <TrendingDown size={12} className="rotate-180 text-black" /> Dar Baixa
                                  </button>
                                ) : (
                                  <span className="text-xs text-slate-600 italic">Sem Saldo</span>
                                )}
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
            <div className="space-y-6" id="reports-tab-view">
              
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                <div>
                  <h2 className="text-xl font-serif font-semibold text-white">Relatórios e Registro de Baixas</h2>
                  <p className="text-sm text-slate-400 mt-1">Inspeção detalhada de consumo de insumos na ponta, perdas estimadas e logs de retiradas.</p>
                </div>
                <button
                  onClick={handleExportCSV}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-md text-sm font-semibold border border-slate-700 transition-all flex items-center gap-2 self-start"
                >
                  <FileDown size={16} /> Exportar Planilha Completa
                </button>
              </div>

              {/* Minimalist analytics graphical representation */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* SVG Mock Chart card */}
                <div className="bg-[#161920] border border-slate-800 rounded-xl p-6 lg:col-span-2">
                  <div className="flex justify-between items-center mb-4">
                    <h3 className="text-sm font-bold text-white uppercase tracking-wider">Distribuição do Recebimento de Materiais</h3>
                    <span className="text-xs text-slate-500">Saldos Globais por Obra</span>
                  </div>

                  {/* Simulated full Responsive SVG Chart */}
                  <div className="h-60 flex flex-col justify-end gap-3 pt-6" id="mock-chart-canvas">
                    
                    {/* Alvorada Bar */}
                    <div className="space-y-1">
                      <div className="flex justify-between text-xs text-slate-400">
                        <span>Residencial Alvorada</span>
                        <span>71% Recebido</span>
                      </div>
                      <div className="w-full bg-[#0F1115] rounded-lg h-3 overflow-hidden">
                        <div className="h-full bg-emerald-500 rounded-lg" style={{ width: '71%' }}></div>
                      </div>
                    </div>

                    {/* Infinito Bar */}
                    <div className="space-y-1 mt-2">
                      <div className="flex justify-between text-xs text-slate-400">
                        <span>Torre Infinito</span>
                        <span>42% Recebido</span>
                      </div>
                      <div className="w-full bg-[#0F1115] rounded-lg h-3 overflow-hidden">
                        <div className="h-full bg-amber-500 rounded-lg" style={{ width: '42%' }}></div>
                      </div>
                    </div>

                    {/* Hospitalar Bar */}
                    <div className="space-y-1 mt-2">
                      <div className="flex justify-between text-xs text-slate-400">
                        <span>Complexo Hospitalar</span>
                        <span>67% Recebido</span>
                      </div>
                      <div className="w-full bg-[#0F1115] rounded-lg h-3 overflow-hidden">
                        <div className="h-full bg-blue-500 rounded-lg" style={{ width: '67%' }}></div>
                      </div>
                    </div>

                    <div className="flex justify-between text-[10px] text-slate-500 pt-3 border-t border-slate-800/80">
                      <span>0% Solicitado</span>
                      <span>25%</span>
                      <span>50%</span>
                      <span>75%</span>
                      <span>100% Entregue</span>
                    </div>

                  </div>
                </div>

                {/* Efficiency metrics overview */}
                <div className="bg-[#161920] border border-slate-800 rounded-xl p-6">
                  <h3 className="text-sm font-bold text-white uppercase tracking-wider mb-4">Status de Insumos</h3>
                  <div className="space-y-4">
                    <div className="flex justify-between items-center p-2.5 bg-slate-900/60 rounded border border-slate-800">
                      <span className="text-xs text-slate-400">Cimento CP-II</span>
                      <span className="text-xs font-semibold text-emerald-400">Bom Estado (300 un)</span>
                    </div>
                    <div className="flex justify-between items-center p-2.5 bg-slate-900/60 rounded border border-slate-800">
                      <span className="text-xs text-slate-400">Vergalhão CA-50</span>
                      <span className="text-xs font-semibold text-rose-400">Alerta Crítico (0 kg)</span>
                    </div>
                    <div className="flex justify-between items-center p-2.5 bg-slate-900/60 rounded border border-slate-800">
                      <span className="text-xs text-slate-400">Areia Lavada</span>
                      <span className="text-xs font-semibold text-emerald-400">Lote Reservado (25 m³)</span>
                    </div>
                    <div className="flex justify-between items-center p-2.5 bg-slate-900/60 rounded border border-slate-800">
                      <span className="text-xs text-slate-400">Piso Porcelanato</span>
                      <span className="text-xs font-semibold text-slate-500">Inativo / Cancelado</span>
                    </div>
                  </div>
                </div>

              </div>

              {/* Log List of historical withdrawals (Historico de Baixas) */}
              <div className="bg-[#161920] border border-slate-800 rounded-xl overflow-hidden" id="withdraw-logs-panel">
                <div className="p-5 border-b border-slate-800 bg-[#1C2028] flex justify-between items-center">
                  <div>
                    <h3 className="text-base font-serif font-semibold text-white">Diário de Retiradas e Consumo</h3>
                    <p className="text-xs text-slate-500 mt-0.5">Logs sequenciais de movimentações registradas por mestre de obras e estagiários.</p>
                  </div>
                  <History size={18} className="text-slate-400" />
                </div>

                <div className="divide-y divide-slate-800/80">
                  {baixas
                    .filter((b) => selectedObra === 'Todas as Obras' || b.obra === selectedObra)
                    .map((item) => (
                      <div key={item.id} className="p-4.5 hover:bg-slate-800/10 transition-colors flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-sm">
                        
                        <div className="flex items-start gap-3">
                          <div className="w-8 h-8 rounded bg-orange-950/20 border border-orange-500/20 text-orange-400 flex items-center justify-center shrink-0">
                            <TrendingDown size={14} />
                          </div>
                          <div>
                            <span className="text-xs text-slate-500 italic block">{item.data} • Responsável: <strong className="text-slate-300">{item.colaborador}</strong></span>
                            <span className="text-white font-medium">{item.insumo}</span> na obra <span className="text-slate-400 font-semibold">{item.obra}</span>
                          </div>
                        </div>

                        <div className="text-right sm:text-right flex sm:flex-col justify-between sm:justify-start items-center sm:items-end">
                          <span className="text-xs text-slate-500 block sm:hidden">Quantidade retirada:</span>
                          <span className="text-sm font-bold text-orange-400 font-mono">-{item.quantidade} {item.unidade}</span>
                          <span className="text-[10px] text-slate-500 font-mono block mt-0.5">{item.id}</span>
                        </div>

                      </div>
                    ))}
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
                  <h3 className="text-sm font-bold text-white font-serif">Painel de Administração Corporal & Permissões</h3>
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
                      <h3 className="text-sm font-serif font-semibold text-white flex items-center gap-2">
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
                      <h3 className="text-sm font-serif font-semibold text-white flex items-center gap-2">
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
                      <h3 className="text-sm font-serif font-semibold text-white flex items-center gap-2">
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
                      {obras.map((obr) => (
                        <div key={obr} className="flex items-center justify-between p-3 bg-[#0F1115] rounded-lg border border-slate-800 group hover:border-slate-700 transition-all">
                          {editingObraName === obr ? (
                            <div className="flex items-center gap-2 w-full">
                              <input
                                type="text"
                                defaultValue={obr}
                                id={`input-rename-obra-${obr}`}
                                className="bg-slate-900 border border-purple-500 text-xs text-white rounded px-2 py-1 w-full focus:outline-none"
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter') {
                                    handleRenameObra(obr, (e.target as HTMLInputElement).value);
                                  } else if (e.key === 'Escape') {
                                    setEditingObraName(null);
                                  }
                                }}
                              />
                              <button
                                type="button"
                                onClick={() => {
                                  const val = (document.getElementById(`input-rename-obra-${obr}`) as HTMLInputElement)?.value;
                                  handleRenameObra(obr, val || '');
                                }}
                                className="text-emerald-400 hover:text-emerald-300 text-xs font-semibold px-2"
                              >
                                Salvar
                              </button>
                            </div>
                          ) : (
                            <>
                              <span className="text-xs font-semibold text-white truncate max-w-[180px]">{obr}</span>
                              <div className="flex items-center gap-1.5 opacity-80 group-hover:opacity-100 transition-all">
                                <button
                                  type="button"
                                  onClick={() => setEditingObraName(obr)}
                                  className="p-1 text-slate-400 hover:text-white rounded hover:bg-slate-800 transition-all"
                                  title="Renomear Empreendimento"
                                >
                                  <Edit size={12} />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteObra(obr)}
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
                    <div className="flex gap-2">
                      <input
                        type="text"
                        placeholder="Ex: Condomínio Bella Vue..."
                        id="new-obra-input-value"
                        className="bg-[#0F1115] border border-slate-700 rounded-lg py-1.5 px-3 text-xs text-slate-200 placeholder-slate-600 focus:border-purple-500 focus:outline-none flex-1"
                      />
                      <button
                        type="button"
                        onClick={() => {
                          const val = (document.getElementById('new-obra-input-value') as HTMLInputElement)?.value || '';
                          handleCreateObra(val);
                          const inp = document.getElementById('new-obra-input-value') as HTMLInputElement;
                          if (inp) inp.value = '';
                        }}
                        className="bg-purple-650 hover:bg-purple-600 text-white rounded-lg px-3.5 py-1.5 text-xs font-bold transition-all flex items-center gap-1 shrink-0"
                      >
                        <Plus size={14} />
                        Adicionar
                      </button>
                    </div>
                  </div>

                </div>

                {/* 2. SECTOR: USUÁRIOS E PERMISSÕES (7 Columns) */}
                <div className="lg:col-span-7 bg-[#161920] border border-slate-800 rounded-xl p-5 flex flex-col justify-between" id="admin-usuarios-panel">
                  <div>
                    <div className="flex justify-between items-center mb-4 pb-2 border-b border-slate-800">
                      <h3 className="text-sm font-serif font-semibold text-white flex items-center gap-2">
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
                  className="absolute inset-0 opacity-0 cursor-pointer w-full h-full"
                />
                <Upload className="text-slate-500 w-8 h-8" />
                <div className="text-center">
                  <span className="text-xs text-slate-300 font-semibold block">Escolher Planilha de Compras Sienge</span>
                  <span className="text-[10px] text-slate-500 block mt-0.5">Arraste para cá ou clique para explorar seu computador</span>
                </div>
              </div>

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
              <h3 className="text-base font-serif font-semibold text-white">Nova Entrada: Comprar Insumo</h3>
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
          <div className="bg-[#161920] border border-slate-700 w-full max-w-md rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <h3 className="text-base font-serif font-semibold text-white">Registrar Saída / Baixa em Obra</h3>
              <button
                onClick={() => setShowWithdrawModal(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            <form onSubmit={handleCreateWithdrawal} className="p-6 space-y-4">
              
              {/* Construction Site Select (Obra origin) */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Selecionar Obra:</label>
                <select
                  value={newWithdrawObra}
                  onChange={(e) => handleWithdrawObraChange(e.target.value)}
                  className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-sm rounded-lg p-2.5 focus:border-emerald-500 focus:outline-none"
                >
                  {listObras.map((o) => (
                    <option key={o} value={o}>{o}</option>
                  ))}
                </select>
              </div>

              {/* Insumo Origin of Withdrawal */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Insumo com Saldo Disponível:</label>
                {insumosForWithdrawal.length === 0 ? (
                  <div className="bg-[#150a0a] border border-red-950 p-3 rounded text-xs text-red-400 italic">
                    Não existem insumos com saldo em estoque para retirada nesta obra no momento.
                  </div>
                ) : (
                  <select
                    value={newWithdrawInsumo}
                    onChange={(e) => {
                      setNewWithdrawInsumo(e.target.value);
                      const matched = insumosForWithdrawal.find(i => i.insumo === e.target.value);
                      if (matched) {
                        setNewWithdrawQtd(Math.min(10, matched.saldo));
                      }
                    }}
                    required
                    className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-sm rounded-lg p-2.5 focus:border-emerald-500 focus:outline-none"
                  >
                    <option value="">-- Selecionar Insumo --</option>
                    {insumosForWithdrawal.map((item, idx) => (
                      <option key={idx} value={item.insumo}>
                        {item.insumo} (Saldo: {item.saldo} {item.unidade})
                      </option>
                    ))}
                  </select>
                )}
              </div>

              {/* Outgoing quantity to log */}
              {newWithdrawInsumo && (
                <div>
                  <div className="flex justify-between items-center mb-1">
                    <label className="text-xs text-slate-400 font-semibold block">Quantidade a Devolver / Retirar:</label>
                    <span className="text-[10px] text-emerald-400 font-mono">
                      Máx: {insumosForWithdrawal.find(i => i.insumo === newWithdrawInsumo)?.saldo || 0}
                    </span>
                  </div>
                  <input
                    type="number"
                    min="1"
                    max={insumosForWithdrawal.find(i => i.insumo === newWithdrawInsumo)?.saldo || 9999}
                    value={newWithdrawQtd}
                    onChange={(e) => {
                      const maxAvailable = insumosForWithdrawal.find(i => i.insumo === newWithdrawInsumo)?.saldo || 0;
                      setNewWithdrawQtd(Math.min(maxAvailable, Math.max(1, parseInt(e.target.value) || 0)));
                    }}
                    className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-sm text-slate-300 focus:border-emerald-500 focus:outline-none"
                  />
                </div>
              )}

              {/* Collaborator (who is the person withdrawing) */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Colaborador / Mestre Responsável:</label>
                <input
                  type="text"
                  placeholder="Ex: Mestre de Obras Antonio, Eng. Silva..."
                  value={newWithdrawColab}
                  onChange={(e) => setNewWithdrawColab(e.target.value)}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-sm text-slate-300 focus:border-emerald-500 focus:outline-none placeholder-slate-600"
                />
              </div>

              {/* Footer */}
              <div className="pt-4 border-t border-slate-800 flex justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => setShowWithdrawModal(false)}
                  className="px-4 py-2 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded text-sm transition-all"
                >
                  Cancelar
                </button>
                <button
                  type="submit"
                  disabled={insumosForWithdrawal.length === 0 || !newWithdrawInsumo}
                  className={`px-4 py-2 rounded text-sm font-semibold transition-all ${
                    insumosForWithdrawal.length === 0 || !newWithdrawInsumo
                      ? 'bg-slate-800 text-slate-600 cursor-not-allowed border border-slate-800'
                      : 'bg-emerald-600 hover:bg-emerald-500 text-white'
                  }`}
                >
                  Registrar Saída
                </button>
              </div>

            </form>
          </div>
        </div>
      )}

      {/* MODAL 3: INTEGRAR ENTRADA RECEBIDA (RECEIVE DELIVERED CONSIGNMENT) */}
      {showReceiveModal && (
        <div className="fixed inset-0 bg-black/75 backdrop-blur-sm z-50 flex items-center justify-center p-4" id="receive-entry-modal">
          <div className="bg-[#161920] border border-slate-700 w-full max-w-sm rounded-xl shadow-2xl overflow-hidden animate-in fade-in zoom-in-95 duration-200">
            
            <div className="p-5 border-b border-slate-800 flex justify-between items-center bg-[#1C2028]">
              <div>
                <h3 className="text-sm font-bold text-white uppercase tracking-wider">Confirmar Entrada Física</h3>
                <span className="text-xs text-slate-500 mt-0.5">Lote recebido por transportadora</span>
              </div>
              <button
                onClick={() => setShowReceiveModal(false)}
                className="text-slate-400 hover:text-white p-1 hover:bg-slate-800 rounded transition-all"
              >
                <X size={18} />
              </button>
            </div>

            <form onSubmit={handleConfirmReceipt} className="p-6 space-y-4">
              
              {/* Detailed Context */}
              <div className="bg-[#0F1115] p-3.5 rounded-lg border border-slate-800 text-xs">
                {(() => {
                  const currPed = pedidos.find(p => p.id === selectedPedidoId);
                  if (!currPed) return null;
                  return (
                    <div className="space-y-1.5">
                      <div className="text-slate-500">ID Pedido: <span className="font-bold text-slate-300">{currPed.id}</span></div>
                      <div className="text-white font-medium">{currPed.insumo}</div>
                      <div className="text-slate-400">Obra Alinhada: <span className="font-semibold text-slate-300">{currPed.obra}</span></div>
                      <div className="flex justify-between pt-1.5 border-t border-slate-850 mt-1">
                        <span>Esperado Total:</span>
                        <strong className="text-slate-300">{currPed.qtdSolicitada} {currPed.unidade}</strong>
                      </div>
                      <div className="flex justify-between">
                        <span>Pendente Restante:</span>
                        <strong className="text-amber-500">{currPed.qtdSolicitada - currPed.qtdRecebida} {currPed.unidade}</strong>
                      </div>
                    </div>
                  );
                })()}
              </div>

              {/* Delivery Input Quantity */}
              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Quantidade Entregue Agora:</label>
                <input
                  type="number"
                  min="1"
                  value={qtdReceiveInput}
                  onChange={(e) => setQtdReceiveInput(Math.max(1, parseInt(e.target.value) || 0))}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-sm text-slate-300 focus:border-emerald-500 focus:outline-none"
                />
              </div>

              {/* Auto full check */}
              <div className="flex justify-end gap-2.5 pt-4 border-t border-slate-800">
                <button
                  type="button"
                  onClick={() => {
                    const currentPed = pedidos.find(p => p.id === selectedPedidoId);
                    if (currentPed) {
                      setQtdReceiveInput(currentPed.qtdSolicitada - currentPed.qtdRecebida);
                    }
                  }}
                  className="px-3 py-1.5 text-xs text-emerald-400 hover:text-white hover:bg-emerald-950/20 border border-emerald-500/20 rounded transition-all font-semibold"
                >
                  Receber Tudo
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded text-xs font-semibold transition-all"
                >
                  Confirmar Entrega
                </button>
              </div>

            </form>
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
            <h3 className="text-lg font-serif font-bold text-white mb-2">Acesso Negado</h3>
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
                <h3 className="text-base font-serif font-semibold text-white">Editar Especificações do Pedido</h3>
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
                  onChange={(e) => setEditingPedido({ ...editingPedido, obra: e.target.value })}
                  className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-xs rounded-lg p-2.5 focus:border-purple-500 focus:outline-none"
                >
                  {obras.map((o) => (
                    <option key={o} value={o}>{o}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Nome do Insumo:</label>
                <input
                  type="text"
                  value={editingPedido.insumo}
                  onChange={(e) => setEditingPedido({ ...editingPedido, insumo: e.target.value })}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0"
                />
              </div>

              <div>
                <label className="text-xs text-slate-400 font-semibold block mb-1.5">Código do Insumo:</label>
                <input
                  type="text"
                  value={editingPedido.codigo}
                  onChange={(e) => setEditingPedido({ ...editingPedido, codigo: e.target.value.toUpperCase() })}
                  className="w-full bg-[#0F1115] border border-slate-700 rounded-lg p-2.5 text-xs text-slate-200 focus:border-purple-500 focus:outline-none focus:ring-0 font-mono"
                />
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
                  <label className="text-xs text-slate-400 font-semibold block mb-1.5">Unidade de Medida:</label>
                  <select
                    value={editingPedido.unidade}
                    onChange={(e) => setEditingPedido({ ...editingPedido, unidade: e.target.value })}
                    className="w-full bg-[#0F1115] border border-slate-700 text-slate-300 text-xs rounded-lg p-2.5"
                  >
                    <option value="un">un (unidade)</option>
                    <option value="kg">kg (quilograma)</option>
                    <option value="m³">m³ (metro cúbico)</option>
                    <option value="m²">m² (metro quadrado)</option>
                    <option value="m">m (metro linear)</option>
                  </select>
                </div>

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
                <h3 className="text-base font-serif font-semibold text-white">Cadastrar Novo Colaborador</h3>
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
                <label className="text-xs text-slate-400 block mb-1.5 font-serif font-medium">Função Atribuída:</label>
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

    </div>
  );
}

