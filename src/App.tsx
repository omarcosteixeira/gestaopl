import React, { useState, useEffect, useMemo } from "react";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";
import { initializeApp, getApp } from "firebase/app";
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  createUserWithEmailAndPassword,
  signOut,
  sendPasswordResetEmail,
  updatePassword,
  updateProfile,
  getAuth,
  User,
} from "firebase/auth";
import {
  collection,
  query,
  onSnapshot,
  addDoc,
  updateDoc,
  doc,
  deleteDoc,
  serverTimestamp,
  where,
  or,
  limit,
  orderBy,
  getDoc,
  setDoc,
  getDocs,
  writeBatch,
} from "firebase/firestore";
import {
  LayoutDashboard,
  UserPlus,
  History,
  Database,
  GraduationCap,
  Settings,
  LogOut,
  Plus,
  Trash2,
  ExternalLink,
  CheckCircle2,
  XCircle,
  Phone,
  Search,
  Users,
  User as UserIcon,
  TrendingUp,
  Calendar,
  Download,
  Upload,
  Menu,
  X,
  Check,
  ChevronRight,
  AlertCircle,
  FileText,
  Clock,
  Calculator,
  LayoutGrid,
  List,
  ShieldCheck,
  Megaphone,
  Sun,
  Edit2,
  Share2,
  Edit,
  Save,
  MapPin,
  Lock,
  Unlock,
  Circle,
  KeyRound,
  Building2,
  MessageSquare,
  PhoneOutgoing,
  Mail,
  Globe,
  Copy,
  Bot,
  Send,
  Bell,
  Monitor,
  Maximize,
  Cloud,
  RefreshCw,
  Play,
  Pause,
  ChevronUp,
  ChevronDown,
  Target,
  Cake,
  CheckSquare,
  Square,
  Coins,
  BookOpen,
  Briefcase,
  Boxes,
  Smartphone,
  Chrome,
  BarChart3,
  Eye,
  EyeOff,
  UserMinus,
  Wrench,
  Hash,
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import {
  auth,
  db,
  COLLECTIONS,
  handleFirestoreError,
  OperationType,
  secondaryAuth,
  firebaseConfigPrincipal,
  firebaseConfigComercial,
} from "./firebase";
import {
  cn,
  formatPhone,
  getWhatsAppUrl,
  validateCPF,
  formatCPF,
} from "./lib/utils";
import * as XLSX from "xlsx";
import { EmailMarketingView } from "./components/EmailMarketingView";
import { RelatoriosView } from "./components/RelatoriosView";
import { ControleConcorrenciaView } from "./components/ControleConcorrenciaView";
import Mapa3D from "./components/Mapa3D";
import {
  UserProfile,
  SalesContact,
  Lead,
  BaseEntry,
  GapEntry,
  PlannerTask,
  LinkUtil,
  UserRole,
  FiesProuniEntry,
  FiesProuniVaga,
  Campanha,
  BomDiaCaptacao,
  ForecastCaptacao,
  BomDiaMetrics,
  PeriodoCaptacao,
  CalendarioAcao,
  EmpresaParceira,
  WhatsAppMessage,
  MapaoAcademicoEntry,
  BaseDisparoEntry,
  BotConfig,
  MetaDia,
  MetaSM,
  MetaCurso,
  QgLigacao,
  SolicitacaoFolga,
  CursoDisponivel,
  InsumoPedido,
  InsumoEstoque,
  InsumoBaixa,
  InsumoPedidoComercial,
  InsumoEstoqueComercial,
  IsencaoEntry,
  ControleConcorrencia,
  PedidoCursoEntry,
  Ligacao,
  AnalysisScheme,
  PeriodAnalysis,
  SolicitacaoManutencao
} from "./types";
import { OPENROUTER_MODELS } from "./ai-config";
import CrescimentoAnualAdmin from "./components/CrescimentoAnualAdmin";
import { ProfileModal } from "./components/ProfileModal";
import { PublicRegistrationForm } from "./components/PublicRegistrationForm";
import { FormulariosView } from "./components/FormulariosView";
import { PublicCustomForm } from "./components/PublicCustomForm";
import { PublicInsumoForm } from "./components/PublicInsumoForm";
import { PublicMaintenanceForm } from "./components/PublicMaintenanceForm";
import { PublicPedidoCursoForm } from "./components/PublicPedidoCursoForm";
import { MessageTemplateModal } from "./components/MessageTemplateModal";
import { CursosDisponiveisView } from "./components/CursosDisponiveisView";
import { ControleInsumosView } from "./components/ControleInsumosView";
import { SolicitacoesManutencaoView } from "./components/SolicitacoesManutencaoView";
import { ControleInsumosComercialView } from "./components/ControleInsumosComercialView";
import { WhatsAppMessageEditor } from "./components/WhatsAppMessageEditor";
import { AdminFuncionariosView } from "./components/AdminFuncionariosView";
import { IsencoesView } from "./components/IsencoesView";
import { WhatsAppMessageSelector } from "./components/WhatsAppMessageSelector";
import { MultiSelect } from "./components/MultiSelect";
import { EvasaoView } from "./components/EvasaoView";
import NovasOportunidadesView from "./components/NovasOportunidadesView";
import ControleLigacoesView from "./components/ControleLigacoesView";
import CRMView from "./components/CRMView";
import MetaSMView from "./components/MetaSMView";
import MetaCursosView from "./components/MetaCursosView";

// --- Helpers ---
export const replaceMessageVariables = (
  template: string,
  lead: any,
): string => {
  if (!template) return "";
  let text = template;
  text = text.replace(/\[nome\]/gi, lead.nome || "");
  text = text.replace(/\[curso\]/gi, lead.curso || lead.cursoInteresse || "");
  text = text.replace(/\[matr[ií]cula\]/gi, lead.numeroMatricula || "");

  // Novas variáveis
  text = text.replace(
    /\[unidade\]/gi,
    lead.unidade || lead.nome_unidade || "nossa unidade",
  );
  text = text.replace(
    /\[data_contato\]/gi,
    new Date().toLocaleDateString("pt-BR"),
  );

  const hour = new Date().getHours();
  const saudacao =
    hour < 12 ? "Bom dia" : hour < 18 ? "Boa tarde" : "Boa noite";
  text = text.replace(/\[saudacao\]/gi, saudacao);

  if (lead.missingDocs) {
    text = text.replace(
      /\[pendencias\]/gi,
      Array.isArray(lead.missingDocs)
        ? lead.missingDocs.join(", ")
        : lead.missingDocs,
    );
  }

  return text;
};

const exportToExcel = (data: any[], fileName: string) => {
  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Dados");
  XLSX.writeFile(workbook, `${fileName}.xlsx`);
};

const exportToCSV = (data: any[], fileName: string) => {
  if (data.length === 0) return;
  const headers = Object.keys(data[0]);
  const csvContent = [
    headers.join(","),
    ...data.map((row) =>
      headers.map((header) => JSON.stringify(row[header])).join(","),
    ),
  ].join("\n");

  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const link = document.createElement("a");
  const url = URL.createObjectURL(blob);
  link.setAttribute("href", url);
  link.setAttribute("download", `${fileName}.csv`);
  link.style.visibility = "hidden";
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
};

const importFromExcel = (file: File, callback: (data: any[]) => void) => {
  const reader = new FileReader();
  reader.onload = (e) => {
    const bstr = e.target?.result;
    const workbook = XLSX.read(bstr, { type: "binary" });
    const worksheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    callback(data);
  };
  reader.readAsBinaryString(file);
};

// Component WhatsAppMessageSelector moved to src/components/WhatsAppMessageSelector.tsx

// --- Constants ---
const HOLIDAYS = [
  "2024-01-01",
  "2024-03-29",
  "2024-04-21",
  "2024-05-01",
  "2024-05-30",
  "2024-07-09",
  "2024-09-07",
  "2024-10-12",
  "2024-11-02",
  "2024-11-15",
  "2024-11-20",
  "2024-12-25",
  "2025-01-01",
  "2025-04-18",
  "2025-04-21",
  "2025-05-01",
  "2025-06-19",
  "2025-09-07",
  "2025-10-12",
  "2025-11-02",
  "2025-11-15",
  "2025-11-20",
  "2025-12-25",
  "2026-01-01",
  "2026-04-03",
  "2026-04-21",
  "2026-05-01",
  "2026-06-04",
  "2026-09-07",
  "2026-10-12",
  "2026-11-02",
  "2026-11-15",
  "2026-11-20",
  "2026-12-25",
];

const getWorkingDaysRemaining = (endDateStr: string) => {
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const dataFim = new Date(endDateStr);
  dataFim.setHours(0, 0, 0, 0);

  if (dataFim < hoje) return 0;

  let count = 0;
  let curDate = new Date(hoje.getTime());
  // Start counting from today if it's a working day
  while (curDate <= dataFim) {
    const dayOfWeek = curDate.getDay(); // 0 = Sunday
    const dateString = curDate.toISOString().split("T")[0];
    const isSunday = dayOfWeek === 0;
    const isHoliday = HOLIDAYS.includes(dateString);

    if (!isSunday && !isHoliday) {
      count++;
    }
    curDate.setDate(curDate.getDate() + 1);
  }
  return count;
};

const getWorkingDaysBetween = (startDateStr: string, endDateStr: string) => {
  const start = new Date(startDateStr);
  start.setHours(0, 0, 0, 0);
  const end = new Date(endDateStr);
  end.setHours(0, 0, 0, 0);

  if (end < start) return 0;

  let count = 0;
  let curDate = new Date(start.getTime());
  while (curDate <= end) {
    const dayOfWeek = curDate.getDay();
    const dateString = curDate.toISOString().split("T")[0];
    const isSunday = dayOfWeek === 0;
    const isHoliday = HOLIDAYS.includes(dateString);

    if (!isSunday && !isHoliday) {
      count++;
    }
    curDate.setDate(curDate.getDate() + 1);
  }
  return count;
};

const formatLocalDateString = (dateStr: string) => {
  if (!dateStr) return "";
  const dateOnly = dateStr.split("T")[0];
  if (dateOnly.includes("-")) {
    const parts = dateOnly.split("-");
    if (parts.length === 3) {
      return `${parts[2]}/${parts[1]}/${parts[0]}`;
    }
  }
  return dateStr;
};

export const ROLES: Record<string, UserRole> = {
  ADMIN_MASTER: "Admin Master",
  PROMOTOR: "Promotor",
  FDV: "FDV",
  SALA_MATRICULA: "Sala de Matrícula",
  QG: "QG",
  LIDER_FDV: "Líder/FDV",
  SSA: "SSA",
  GESTOR_UNIDADE: "Gestor Unidade",
  GESTOR_COMERCIAL: "Gestor Comercial",
  ACADEMICO: "Acadêmico",
  PROMOTOR_RUA: "Promotor/rua",
  GESTOR_COMERCIAL_COMERCIAL: "Gerente Comercial (Comercial)",
  FDV_COMERCIAL: "FDV (Comercial)",
  FINANCEIRO: "Financeiro",
  TECNICO: "Técnico",
  REGIONAL: "Regional",
};

const VIEW_PERMISSIONS: Record<string, UserRole[]> = {
  dashboard: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  formularios: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  relatorios: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  cadastro: [
    ROLES.ADMIN_MASTER,
    ROLES.PROMOTOR,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  historico: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  bases: [
    ROLES.ADMIN_MASTER,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_UNIDADE,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
  ],
  gap: [ROLES.ADMIN_MASTER, ROLES.SALA_MATRICULA, ROLES.LIDER_FDV, ROLES.GESTOR_UNIDADE],
  fiesProuni: [
    ROLES.ADMIN_MASTER,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
  ],
  campanhas: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  calendario: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  empresas: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.QG,
    ROLES.SSA,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  calculo: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.PROMOTOR,
    ROLES.SSA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  mapao: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  basesDisparo: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
  ],
  basesRenovacao: [ROLES.ADMIN_MASTER, ROLES.LIDER_FDV, ROLES.SSA],
  avisos: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.PROMOTOR,
    ROLES.ACADEMICO,
  ],
  emailMarketing: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  controleConcorrencia: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_UNIDADE,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FINANCEIRO,
  ],
  admin: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
  ],
  crm: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
  ],
  controlePagamentos: [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_UNIDADE,
  ],
  evasao: [
    ROLES.ADMIN_MASTER,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.LIDER_FDV,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.SALA_MATRICULA,
  ],
  cursos: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
    ROLES.LIDER_FDV,
    ROLES.SSA,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.ACADEMICO,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.FDV_COMERCIAL,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
  ],
  solicitacaoManutencao: Object.values(ROLES),
  controleInsumos: [
    ROLES.ADMIN_MASTER,
    ROLES.ACADEMICO,
    ROLES.FINANCEIRO,
    ROLES.TECNICO,
    ROLES.GESTOR_UNIDADE,
    ROLES.LIDER_FDV,
  ],
  controleInsumosComercial: [
    ROLES.ADMIN_MASTER,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.LIDER_FDV,
    ROLES.FINANCEIRO,
  ],
  isencoes: [
    ROLES.ADMIN_MASTER,
    ROLES.SALA_MATRICULA,
    ROLES.LIDER_FDV,
    ROLES.FDV,
    ROLES.GESTOR_UNIDADE,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.FDV_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ],
  controleLigacoes: [
    ROLES.ADMIN_MASTER,
    ROLES.GESTOR_COMERCIAL,
    ROLES.REGIONAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
    ROLES.LIDER_FDV,
    ROLES.FDV,
    ROLES.FDV_COMERCIAL,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.SALA_MATRICULA,
    ROLES.QG,
  ],
};

// --- Components ---
function PasswordChangeModal({ onComplete }: { onComplete: () => void }) {
  const [newPassword, setNewPassword] = useState("");
  const [confirmPassword, setConfirmPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (newPassword.length < 6) {
      setError("A senha deve ter pelo menos 6 caracteres.");
      return;
    }
    if (newPassword !== confirmPassword) {
      setError("As senhas não coincidem.");
      return;
    }

    setLoading(true);
    setError("");
    try {
      if (auth.currentUser) {
        await updatePassword(auth.currentUser, newPassword);
        onComplete();
      }
    } catch (err: any) {
      setError("Erro ao atualizar senha. Tente sair e entrar novamente.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-900/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
      <motion.div
        initial={{ opacity: 0, y: 20 }}
        animate={{ opacity: 1, y: 0 }}
        className="bg-white p-8 rounded-3xl shadow-2xl border border-slate-100 max-w-md w-full"
      >
        <div className="w-16 h-16 bg-blue-100 text-blue-600 rounded-2xl flex items-center justify-center mb-6">
          <KeyRound size={32} />
        </div>
        <h2 className="text-2xl font-bold text-slate-900 mb-2">
          Troca de Senha Obrigatória
        </h2>
        <p className="text-slate-500 mb-6">
          Para sua segurança, você deve alterar sua senha padrão antes de
          continuar.
        </p>

        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-xs font-bold text-slate-500 mb-1">
              Nova Senha
            </label>
            <input
              type="password"
              required
              value={newPassword}
              onChange={(e) => setNewPassword(e.target.value)}
              className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
              placeholder="Mínimo 6 caracteres"
            />
          </div>
          <div>
            <label className="block text-xs font-bold text-slate-500 mb-1">
              Confirmar Nova Senha
            </label>
            <input
              type="password"
              required
              value={confirmPassword}
              onChange={(e) => setConfirmPassword(e.target.value)}
              className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
            />
          </div>

          {error && (
            <div className="p-3 bg-rose-50 text-rose-600 text-xs font-bold rounded-xl flex items-center space-x-2">
              <AlertCircle size={14} />
              <span>{error}</span>
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-blue-600 text-white py-4 rounded-2xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 disabled:opacity-50"
          >
            {loading ? "Atualizando..." : "Atualizar Senha"}
          </button>
        </form>
      </motion.div>
    </div>
  );
}

const Toast = ({
  message,
  type,
  onClose,
}: {
  message: string;
  type: "success" | "error";
  onClose: () => void;
}) => (
  <motion.div
    initial={{ x: 100, opacity: 0 }}
    animate={{ x: 0, opacity: 1 }}
    exit={{ x: 100, opacity: 0 }}
    className={cn(
      "fixed top-5 right-5 z-50 p-4 rounded-lg shadow-lg flex items-center space-x-2 text-white",
      type === "success" ? "bg-emerald-600" : "bg-rose-600",
    )}
  >
    {type === "success" ? (
      <CheckCircle2 size={20} />
    ) : (
      <AlertCircle size={20} />
    )}
    <span className="font-medium">{message}</span>
    <button onClick={onClose} className="ml-2 hover:opacity-80">
      <X size={16} />
    </button>
  </motion.div>
);

function MapaoAcademicoView({
  mapao,
  onToast,
  profile,
}: {
  mapao: MapaoAcademicoEntry[];
  onToast: (m: string, t?: "success" | "error") => void;
  profile: UserProfile;
}) {
  const [showModal, setShowModal] = useState(false);
  const [editingEntry, setEditingEntry] = useState<MapaoAcademicoEntry | null>(
    null,
  );

  const defaultDisciplina = {
    codDisc: "",
    disciplina: "",
    dia: "Segunda-feira",
    horario: "",
    turma: "",
    tipoDisciplina: "PRESENCIAL",
    professor: "",
    matricula: "",
    observacao: "",
    linkAula: "",
  };

  const [formData, setFormData] = useState<Partial<MapaoAcademicoEntry>>({
    modalidade: "Presencial",
    tipoCurso: "GRADUACAO",
    periodo: "",
    semestre: "",
    disciplinas: [{ ...defaultDisciplina }],
  });
  const [expandedCards, setExpandedCards] = useState<Record<string, boolean>>({});
  
  // Filters
  const [filterCurso, setFilterCurso] = useState("");
  const [filterPeriodo, setFilterPeriodo] = useState("");
  const [filterSemestre, setFilterSemestre] = useState("");
  const [filterTipoCurso, setFilterTipoCurso] = useState("");

  const toggleExpand = (id: string) => {
    setExpandedCards(prev => ({ ...prev, [id]: !prev[id] }));
  };


  const handleShareCard = async (cardId: string, cardName: string) => {
    const element = document.getElementById(`mapao-card-${cardId}`);
    if (!element) return;
    try {
      onToast("Gerando PDF, aguarde...", "success");
      const canvas = await html2canvas(element, { backgroundColor: '#ffffff', scale: 1, useCORS: true, logging: false });
      if (canvas.width === 0 || canvas.height === 0) {
        throw new Error("Canvas gerado está vazio (tamanho 0). O elemento pode estar oculto.");
      }
      const imgData = canvas.toDataURL('image/jpeg', 0.9);
      if (typeof jsPDF !== "function") throw new Error("jsPDF não carregado corretamente");
      // Fallback for jspdf format
      const pdf = new jsPDF({
        orientation: canvas.width > canvas.height ? 'landscape' : 'portrait',
        unit: 'px',
        format: [canvas.width, canvas.height]
      });
      pdf.addImage(imgData, 'JPEG', 0, 0, canvas.width, canvas.height);
      const safeName = cardName.replace(/[^a-zA-Z0-9]/g, '_');
      pdf.save(`mapao-${safeName}.pdf`);
      onToast("PDF gerado com sucesso!", "success");
    } catch (error: any) {
      console.error('Failed to generate PDF', error);
      onToast(`Erro ao gerar PDF: ${error.message || "Erro desconhecido"}`, "error");
    }
  };

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();

    const isDuplicate = mapao.some(
      (m) =>
        m.curso?.toLowerCase() === formData.curso?.toLowerCase() &&
        m.modalidade === formData.modalidade &&
        m.periodo === formData.periodo &&
        m.id !== editingEntry?.id,
    );

    if (isDuplicate) {
      onToast(
        "Este curso/modalidade/período já está cadastrado no Mapão.",
        "error",
      );
      return;
    }

    try {
      if (editingEntry) {
        await updateDoc(doc(db, COLLECTIONS.MAPAO_ACADEMICO, editingEntry.id), {
          ...formData,
          createdAt: serverTimestamp(),
        });
        onToast("Registro atualizado!");
      } else {
        await addDoc(collection(db, COLLECTIONS.MAPAO_ACADEMICO), {
          ...formData,
          createdAt: serverTimestamp(),
        });
        onToast("Registro cadastrado!");
      }
      setShowModal(false);
      setEditingEntry(null);
      setFormData({
        modalidade: "Presencial",
        tipoCurso: "GRADUACAO",
        periodo: "",
        semestre: "",
        disciplinas: [{ ...defaultDisciplina }],
      });
    } catch (err: any) {
      onToast("Erro ao salvar.", "error");
    }
  };

  const handleDuplicate = async (entry: MapaoAcademicoEntry) => {
    try {
      const { id, ...data } = entry;
      await addDoc(collection(db, COLLECTIONS.MAPAO_ACADEMICO), {
        ...data,
        createdAt: serverTimestamp(),
      });
      onToast("Registro duplicado!");
    } catch (err: any) {
      onToast("Erro ao duplicar.", "error");
    }
  };

  const handleAddDisciplina = () => {
    if (formData.disciplinas && formData.disciplinas.length < 7) {
      setFormData((prev) => ({
        ...prev,
        disciplinas: [...(prev.disciplinas || []), { ...defaultDisciplina }],
      }));
    }
  };

  const handleRemoveDisciplina = (index: number) => {
    const newDisciplinas = [...(formData.disciplinas || [])];
    newDisciplinas.splice(index, 1);
    setFormData((prev) => ({ ...prev, disciplinas: newDisciplinas }));
  };

  const handleChangeDisciplina = (
    index: number,
    field: string,
    value: string,
  ) => {
    const newDisciplinas: any = [...(formData.disciplinas || [])];
    newDisciplinas[index][field] = value;
    if (field === "dia" && value === "Virtual") {
      newDisciplinas[index].horario = "";
    }
    setFormData((prev) => ({ ...prev, disciplinas: newDisciplinas }));
  };

  const handleExport = () => {
    const exportData: any[] = [];
    mapao.forEach(m => {
      const disciplinas = m.disciplinas || [];
      if (disciplinas.length === 0) {
        exportData.push({
          Modalidade: m.modalidade,
          Curso: m.curso,
          Período: m.periodo,
          "Tipo Curso": m.tipoCurso,
          "Cód. Disciplina": "",
          Disciplina: "",
          Dia: "",
          Horário: "",
          Turma: "",
          "Tipo Disciplina": "",
          Professor: "",
          Matrícula: "",
          Observação: "",
          "Link Aula": ""
        });
      } else {
        disciplinas.forEach(d => {
          exportData.push({
            Modalidade: m.modalidade,
            Curso: m.curso,
            Período: m.periodo,
            "Tipo Curso": m.tipoCurso,
            "Cód. Disciplina": d.codDisc,
            Disciplina: d.disciplina,
            Dia: d.dia,
            Horário: d.horario,
            Turma: d.turma,
            "Tipo Disciplina": d.tipoDisciplina,
            Professor: d.professor,
            Matrícula: d.matricula,
            Observação: d.observacao,
            "Link Aula": d.linkAula || ""
          });
        });
      }
    });
    exportToExcel(exportData, "Mapao_Academico");
  };

  const handleImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return "";
        };

        const map = new Map<string, any>();
        
        data.forEach((row: any) => {
          const curso = String(getVal(row, "curso") || "").trim();
          const modalidade = String(getVal(row, "modalidade") || "").trim();
          const periodo = String(getVal(row, "período", "periodo") || "").trim();
          
          if (!curso) return; 
          
          const key = `${curso.toLowerCase()}|${modalidade.toLowerCase()}|${periodo.toLowerCase()}`;
          
          if (!map.has(key)) {
            map.set(key, {
              curso,
              modalidade: modalidade || "Presencial",
              periodo,
              tipoCurso: String(getVal(row, "tipo curso", "tipocurso") || "GRADUACAO").trim(),
              disciplinas: []
            });
          }
          
          const entry = map.get(key);
          const disciplina = String(getVal(row, "disciplina") || "").trim();
          
          if (disciplina) {
            entry.disciplinas.push({
              codDisc: String(getVal(row, "cód. disciplina", "cod disciplina", "coddisc") || "").trim(),
              disciplina,
              dia: String(getVal(row, "dia") || "").trim() || "Segunda-feira",
              horario: String(getVal(row, "horário", "horario") || "").trim(),
              turma: String(getVal(row, "turma") || "").trim(),
              tipoDisciplina: String(getVal(row, "tipo disciplina", "tipodisciplina") || "").trim() || "PRESENCIAL",
              professor: String(getVal(row, "professor") || "").trim(),
              matricula: String(getVal(row, "matrícula", "matricula") || "").trim(),
              observacao: String(getVal(row, "observação", "observacao") || "").trim(),
              linkAula: String(getVal(row, "link aula", "linkaula") || "").trim()
            });
          }
        });
        
        let importedCount = 0;
        const newEntries = Array.from(map.values());
        
        for (const item of newEntries) {
          const isDuplicate = mapao.some(
            (m) =>
              m.curso?.toLowerCase() === item.curso?.toLowerCase() &&
              m.modalidade?.toLowerCase() === item.modalidade?.toLowerCase() &&
              m.periodo?.toLowerCase() === item.periodo?.toLowerCase()
          );

          if (!isDuplicate) {
            await addDoc(collection(db, COLLECTIONS.MAPAO_ACADEMICO), {
              ...item,
              createdAt: serverTimestamp(),
            });
            importedCount++;
          }
        }
                
        onToast(`${importedCount} novos registros importados com sucesso!`, "success");
      } catch (err: any) {
        onToast("Erro ao importar arquivo.", "error");
      }
    });
    
    e.target.value = ''; // Reset input
  };

  const canEdit = true;

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
            Mapão Acadêmico
          </h2>
          <p className="text-sm text-slate-500">
            Gestão de cursos, disciplinas e horários
          </p>
        </div>

        {canEdit && (
          <div className="flex gap-2">
            <input
              type="file"
              accept=".xlsx, .xls"
              className="hidden"
              id="import-mapao-file"
              onChange={handleImport}
            />
            <label
              htmlFor="import-mapao-file"
              className="cursor-pointer bg-slate-100 text-slate-700 px-4 py-2.5 rounded-2xl font-bold hover:bg-slate-200 transition-all flex items-center space-x-2"
            >
              <Upload size={18} />
              <span className="hidden sm:inline">Importar</span>
            </label>
            <button
              onClick={handleExport}
              className="bg-slate-100 text-slate-700 px-4 py-2.5 rounded-2xl font-bold hover:bg-slate-200 transition-all flex items-center space-x-2"
            >
              <Download size={18} />
              <span className="hidden sm:inline">Exportar</span>
            </button>
            <button
              onClick={() => {
                setEditingEntry(null);
                setFormData({
                  modalidade: "Presencial",
                  tipoCurso: "GRADUACAO",
                  disciplinas: [{ ...defaultDisciplina }],
                });
                setShowModal(true);
              }}
              className="bg-blue-600 text-white px-6 py-2.5 rounded-2xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 flex items-center space-x-2"
            >
              <Plus size={20} />
              <span>Novo Cadastro</span>
            </button>
          </div>
        )}

      </div>

      {/* Filter Bar */}
      <div className="bg-white p-4 rounded-3xl border border-slate-100 shadow-sm flex flex-col md:flex-row gap-4 mb-6">
        <div className="flex-1 relative">
          <Search size={18} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400" />
          <input
            type="text"
            placeholder="Filtrar por Curso"
            value={filterCurso}
            onChange={(e) => setFilterCurso(e.target.value)}
            className="w-full pl-11 pr-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          />
        </div>
        <div className="flex-1">
          <input
            type="text"
            placeholder="Filtrar por Período"
            value={filterPeriodo}
            onChange={(e) => setFilterPeriodo(e.target.value)}
            className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          />
        </div>
        <div className="w-full md:w-48 shrink-0">
          <select
            value={filterSemestre}
            onChange={(e) => setFilterSemestre(e.target.value)}
            className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          >
            <option value="">Qualquer Semestre</option>
            <option value="1º">1º</option>
            <option value="2º">2º</option>
            <option value="3º">3º</option>
            <option value="4º">4º</option>
            <option value="5º">5º</option>
            <option value="6º">6º</option>
            <option value="7º">7º</option>
            <option value="8º">8º</option>
            <option value="9º">9º</option>
            <option value="10º">10º</option>
          </select>
        </div>
        <div className="w-full md:w-48 shrink-0">
          <select
            value={filterTipoCurso}
            onChange={(e) => setFilterTipoCurso(e.target.value)}
            className="w-full px-4 py-3 rounded-2xl border border-slate-200 outline-none font-bold text-sm text-slate-700 focus:ring-2 focus:ring-blue-500"
          >
            <option value="">Qualquer Tipo</option>
            <option value="GRADUACAO">GRADUAÇÃO</option>
            <option value="TECNICO">TÉCNICO</option>
          </select>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 xl:grid-cols-3 gap-6">
        {mapao.filter((entry) => {
          const matchCurso = !filterCurso || entry.curso?.toLowerCase().includes(filterCurso.toLowerCase());
          const matchPeriodo = !filterPeriodo || entry.periodo?.toLowerCase().includes(filterPeriodo.toLowerCase());
          const matchSemestre = !filterSemestre || entry.semestre === filterSemestre;
          const matchTipo = !filterTipoCurso || entry.tipoCurso === filterTipoCurso;
          return matchCurso && matchPeriodo && matchSemestre && matchTipo;
        }).map((entry) => {
          const disciplinasList = entry.disciplinas || [];
          const isExpanded = expandedCards[entry.id];
          return (
            <motion.div
              id={`mapao-card-${entry.id}`}
              key={entry.id}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              className={cn(
                "p-6 rounded-[2rem] border shadow-sm transition-all relative group flex flex-col gap-4 bg-white",
                entry.tipoCurso === "GRADUACAO"
                  ? "bg-white border-blue-100"
                  : "bg-white border-emerald-100",
              )}
            >
              <div className="flex justify-between items-start">
                <div className="flex flex-wrap gap-2">
                  <span
                    className={cn(
                      "px-2.5 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider",
                      entry.tipoCurso === "GRADUACAO"
                        ? "bg-blue-100 text-blue-600"
                        : "bg-emerald-100 text-emerald-600",
                    )}
                  >
                    {entry.tipoCurso}
                  </span>
                  <span className="px-2.5 py-1 rounded-full text-[10px] font-bold uppercase tracking-wider bg-slate-100 text-slate-600">
                    {entry.modalidade}
                  </span>
                </div>
                <div className="flex space-x-1 shrink-0 bg-white/50 p-1 rounded-xl shadow-sm border border-slate-100/50 backdrop-blur-sm">
                  <button
                    onClick={() => handleShareCard(entry.id, entry.curso || "Curso")}
                    className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                    title="Baixar como PDF"
                  >
                    <FileText size={14} />
                  </button>
                  {canEdit && (
                    <>
                      <button
                        onClick={() => {
                          setEditingEntry(entry);
                          setFormData({
                            ...entry,
                            semestre: entry.semestre || "",
                            disciplinas:
                              disciplinasList.length > 0
                                ? disciplinasList
                                : [{ ...defaultDisciplina }],
                          });
                          setShowModal(true);
                        }}
                        className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                        title="Editar"
                      >
                        <Edit2 size={14} />
                      </button>
                      <button
                        onClick={() => handleDuplicate(entry)}
                        className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-lg transition-all"
                        title="Duplicar"
                      >
                        <Copy size={14} />
                      </button>
                      <button
                        onClick={async () => {
                          if (window.confirm("Excluir?"))
                            await deleteDoc(
                              doc(db, COLLECTIONS.MAPAO_ACADEMICO, entry.id),
                            );
                        }}
                        className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-all"
                        title="Excluir"
                      >
                        <Trash2 size={14} />
                      </button>
                    </>
                  )}
                </div>
              </div>

              <div className="flex justify-between items-center gap-4 cursor-pointer" onClick={() => toggleExpand(entry.id)}>
                <div>
                  <h3 className="text-xl font-bold text-slate-900 leading-tight mb-1">
                    {entry.curso}
                  </h3>
                  <p className="text-xs text-slate-500 font-bold uppercase tracking-widest">
                    {entry.periodo} {entry.semestre ? ` - ${entry.semestre}` : ""}
                  </p>
                </div>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    toggleExpand(entry.id);
                  }}
                  className="bg-slate-50 hover:bg-slate-100 p-2 rounded-full text-slate-500 transition-colors shrink-0"
                >
                  {isExpanded ? <ChevronUp size={20} /> : <ChevronDown size={20} />}
                </button>
              </div>

              {isExpanded && (
              <div className="flex-1 space-y-3 mt-2">
                {disciplinasList.map((disc, idx) => (
                  <div
                    key={idx}
                    className="bg-slate-50 p-4 rounded-2xl border border-slate-100 flex flex-col gap-2 relative overflow-hidden"
                  >
                    <div className="flex justify-between items-start gap-2">
                      <div>
                        <p className="text-[9px] font-bold text-slate-400 uppercase tracking-widest leading-none mb-1">
                          {disc.codDisc}
                        </p>
                        <p className="text-sm font-bold text-slate-800 leading-tight">
                          {disc.disciplina}
                        </p>
                      </div>
                      <span className={cn(
                        "px-2 py-1 text-[9px] font-bold uppercase tracking-wider rounded-lg shrink-0",
                        disc.tipoDisciplina === "PRESENCIAL" ? "bg-purple-100 text-purple-600" :
                        disc.tipoDisciplina === "TEAMS" || disc.tipoDisciplina === "ONLINE" ? "bg-blue-100 text-blue-600" :
                        "bg-orange-100 text-orange-600"
                      )}>
                        {disc.tipoDisciplina || "PRESENCIAL"}
                      </span>
                    </div>

                    <p className="text-xs text-slate-600 font-medium mt-1">
                      <span className="text-slate-400 font-normal">Prof:</span> {disc.professor || "-"}
                    </p>

                    <div className="flex flex-wrap gap-2 mt-2">
                      <div className="flex items-center space-x-1.5 text-slate-600 bg-white px-2.5 py-1.5 rounded-lg border border-slate-100 shadow-sm">
                        <Calendar size={12} className="text-blue-500" />
                        <span className="text-[10px] font-bold">
                          {disc.dia || "-"}
                        </span>
                      </div>
                      <div className="flex items-center space-x-1.5 text-slate-600 bg-white px-2.5 py-1.5 rounded-lg border border-slate-100 shadow-sm">
                        <Clock size={12} className="text-amber-500" />
                        <span className="text-[10px] font-bold">
                          {disc.horario || "-"}
                        </span>
                      </div>
                      <div className="flex items-center space-x-1.5 text-slate-600 bg-white px-2.5 py-1.5 rounded-lg border border-slate-100 shadow-sm">
                        <Users size={12} className="text-emerald-500" />
                        <span className="text-[10px] font-bold">
                          {disc.turma || "-"}

                        </span>
                      </div>
                    </div>

                    {disc.observacao && (
                      <div className="mt-3 text-xs text-slate-600 bg-white p-2.5 rounded-xl border border-slate-200 border-dashed">
                        {disc.observacao}
                      </div>
                    )}
                    {disc.linkAula && (
                      <a href={disc.linkAula} target="_blank" rel="noopener noreferrer" className="mt-2 flex items-center gap-1.5 text-xs font-bold text-blue-600 hover:text-blue-700 transition-colors w-fit bg-blue-50 px-3 py-1.5 rounded-lg">
                        <ExternalLink size={14} /> Link da Aula
                      </a>
                    )}
                  </div>
                ))}
                {disciplinasList.length === 0 && (
                  <div className="bg-slate-50 border border-slate-100 border-dashed rounded-2xl py-8 flex flex-col items-center justify-center text-slate-400">
                    <GraduationCap size={24} className="mb-2 opacity-50" />
                    <p className="text-xs italic">Nenhuma disciplina cadastrada.</p>
                  </div>
                )}
              </div>
              )}
            </motion.div>
          );
        })}
      </div>

      {showModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            className="bg-white rounded-[2.5rem] shadow-2xl max-w-3xl w-full p-8 overflow-y-auto max-h-[90vh] custom-scrollbar"
          >
            <div className="flex justify-between items-center mb-6 sticky top-0 bg-white py-2 z-10 border-b border-slate-100">
              <h3 className="text-2xl font-bold text-slate-900">
                {editingEntry ? "Editar Curso" : "Novo Cadastro Acadêmico"}
              </h3>
              <button
                type="button"
                onClick={() => setShowModal(false)}
                className="p-2 hover:bg-slate-100 rounded-full transition-all text-slate-400"
              >
                <X size={24} />
              </button>
            </div>

            <form onSubmit={handleSave} className="space-y-8">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pb-6 border-b border-slate-100">
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Período
                  </label>
                  <input
                    type="text"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.periodo || ""}
                    onChange={(e) =>
                      setFormData({ ...formData, periodo: e.target.value })
                    }
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Semestre
                  </label>
                  <select
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.semestre || ""}
                    onChange={(e) =>
                      setFormData({ ...formData, semestre: e.target.value })
                    }
                    required
                  >
                    <option value="">Selecione...</option>
                    <option value="1º">1º</option>
                    <option value="2º">2º</option>
                    <option value="3º">3º</option>
                    <option value="4º">4º</option>
                    <option value="5º">5º</option>
                    <option value="6º">6º</option>
                    <option value="7º">7º</option>
                    <option value="8º">8º</option>
                    <option value="9º">9º</option>
                    <option value="10º">10º</option>
                  </select>
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Tipo de Curso
                  </label>
                  <select
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.tipoCurso}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        tipoCurso: e.target.value as any,
                      })
                    }
                    required
                  >
                    <option value="GRADUACAO">GRADUAÇÃO</option>
                    <option value="TECNICO">TÉCNICO</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Modalidade
                  </label>
                  <select
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.modalidade}
                    onChange={(e) =>
                      setFormData({ ...formData, modalidade: e.target.value })
                    }
                    required
                  >
                    <option value="Presencial">Presencial</option>
                    <option value="EAD">EAD</option>
                    <option value="Semipresencial">Semipresencial</option>
                  </select>
                </div>

                <div className="md:col-span-2">
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Nome do Curso
                  </label>
                  <input
                    type="text"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.curso || ""}
                    onChange={(e) =>
                      setFormData({ ...formData, curso: e.target.value })
                    }
                    required
                  />
                </div>
              </div>

              <div>
                <div className="flex justify-between items-center mb-4">
                  <h4 className="font-bold text-slate-800 text-lg">
                    Disciplinas do Curso
                  </h4>
                  {(formData.disciplinas?.length || 0) < 7 && (
                    <button
                      type="button"
                      onClick={handleAddDisciplina}
                      className="text-sm flex items-center gap-1 text-blue-600 hover:text-blue-700 font-bold bg-blue-50 px-3 py-1.5 rounded-xl"
                    >
                      <Plus size={16} /> Adicionar (
                      {formData.disciplinas?.length || 0}/7)
                    </button>
                  )}
                </div>

                <div className="space-y-4">
                  {formData.disciplinas?.map((disc, idx) => (
                    <div
                      key={idx}
                      className="bg-slate-50 border border-slate-200 p-5 rounded-2xl relative"
                    >
                      {formData.disciplinas &&
                        formData.disciplinas.length > 1 && (
                          <button
                            type="button"
                            onClick={() => handleRemoveDisciplina(idx)}
                            className="absolute top-4 right-4 text-rose-400 hover:text-rose-600 hover:bg-rose-50 p-1.5 rounded-lg transition-colors"
                          >
                            <Trash2 size={16} />
                          </button>
                        )}
                      <h5 className="text-xs font-bold uppercase text-slate-400 mb-4">
                        Disciplina {idx + 1}
                      </h5>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Código
                          </label>
                          <input
                            type="text"
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.codDisc}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "codDisc",
                                e.target.value,
                              )
                            }
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Disciplina
                          </label>
                          <input
                            type="text"
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.disciplina}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "disciplina",
                                e.target.value,
                              )
                            }
                            required
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Dia da Semana
                          </label>
                          <select
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.dia}
                            onChange={(e) =>
                              handleChangeDisciplina(idx, "dia", e.target.value)
                            }
                            required
                          >
                            {[
                              "Segunda-feira",
                              "Terça-feira",
                              "Quarta-feira",
                              "Quinta-feira",
                              "Sexta-feira",
                              "Sábado",
                              "Virtual",
                            ].map((d) => (
                              <option key={d} value={d}>
                                {d}
                              </option>
                            ))}
                          </select>
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Horário{" "}
                            {disc.dia === "Virtual" ? "(Não se aplica)" : ""}
                          </label>
                          <input
                            type="text"
                            placeholder={
                              disc.dia === "Virtual"
                                ? "Virtual"
                                : "Ex: 19:00 - 22:00"
                            }
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500 disabled:opacity-50 disabled:bg-slate-100"
                            value={disc.horario}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "horario",
                                e.target.value,
                              )
                            }
                            required={disc.dia !== "Virtual"}
                            disabled={disc.dia === "Virtual"}
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Turma
                          </label>
                          <input
                            type="text"
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.turma}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "turma",
                                e.target.value,
                              )
                            }
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-slate-500 mb-1">
                            Tipo Disciplina
                          </label>
                          <select
                            className="w-full px-3 py-2.5 rounded-xl border border-slate-200 outline-none text-sm font-bold text-slate-700 focus:ring-2 focus:ring-blue-500"
                            value={disc.tipoDisciplina}
                            onChange={(e) =>
                              handleChangeDisciplina(
                                idx,
                                "tipoDisciplina",
                                e.target.value,
                              )
                            }
                            required
                          >
                            <option value="PRESENCIAL">Presencial</option>
                            <option value="ONLINE">Online</option>
                            <option value="TEAMS">Teams</option>
                          </select>
                        </div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="flex space-x-4 pt-4 sticky bottom-0 bg-white py-4 border-t border-slate-100 z-10">
                <button
                  type="submit"
                  className="flex-1 bg-blue-600 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-100 hover:bg-blue-700 transition-all"
                >
                  {editingEntry ? "Salvar Alterações" : "Cadastrar"}
                </button>
                <button
                  type="button"
                  onClick={() => setShowModal(false)}
                  className="px-8 bg-slate-100 text-slate-600 font-bold py-4 rounded-2xl hover:bg-slate-200 transition-all border border-slate-200"
                >
                  Cancelar
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
}

function BasesDisparoView({
  bases,
  onToast,
}: {
  bases: BaseDisparoEntry[];
  onToast: (m: string, t?: "success" | "error") => void;
}) {
  const [showModal, setShowModal] = useState(false);
  const [filterDate, setFilterDate] = useState(
    new Date().toISOString().split("T")[0],
  );
  const [formData, setFormData] = useState<Partial<BaseDisparoEntry>>({
    data: new Date().toISOString().split("T")[0],
    totalDisparos: 0,
    positivos: 0,
    negativos: 0,
  });

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      await addDoc(collection(db, COLLECTIONS.BASES_DISPARO), {
        ...formData,
        createdAt: serverTimestamp(),
      });
      onToast("Base registrada!");
      setShowModal(false);
      setFormData({
        data: new Date().toISOString().split("T")[0],
        totalDisparos: 0,
        positivos: 0,
        negativos: 0,
      });
    } catch (err: any) {
      onToast("Erro ao registrar.", "error");
    }
  };

  const filteredBases = bases.filter((b) => b.data === filterDate);

  const totalDisparos = filteredBases.reduce(
    (acc, b) => acc + b.totalDisparos,
    0,
  );
  const totalPositivos = filteredBases.reduce((acc, b) => acc + b.positivos, 0);
  const totalNegativos = filteredBases.reduce((acc, b) => acc + b.negativos, 0);

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 tracking-tight">
            Bases de Disparo
          </h2>
          <p className="text-sm text-slate-500">
            Métricas diárias de disparos e conversão
          </p>
        </div>
        <div className="flex items-center space-x-4">
          <input
            type="date"
            className="px-4 py-2.5 rounded-2xl border border-slate-200 text-sm font-bold text-slate-700"
            value={filterDate}
            onChange={(e) => setFilterDate(e.target.value)}
          />
          <button
            onClick={() => setShowModal(true)}
            className="bg-blue-600 text-white px-6 py-2.5 rounded-2xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-100 flex items-center space-x-2"
          >
            <Plus size={20} />
            <span>Registrar Base</span>
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-sm">
          <p className="text-xs font-bold text-slate-400 uppercase mb-1">
            Total de Disparos
          </p>
          <p className="text-3xl font-black text-blue-600">{totalDisparos}</p>
        </div>
        <div className="bg-emerald-50 p-6 rounded-3xl border border-emerald-100 shadow-sm">
          <p className="text-xs font-bold text-emerald-500 uppercase mb-1">
            Total Positivos
          </p>
          <p className="text-3xl font-black text-emerald-600">
            {totalPositivos}
          </p>
          <p className="text-xs font-bold text-emerald-500 mt-2">
            Taxa:{" "}
            {totalDisparos > 0
              ? ((totalPositivos / totalDisparos) * 100).toFixed(1)
              : 0}
            %
          </p>
        </div>
        <div className="bg-rose-50 p-6 rounded-3xl border border-rose-100 shadow-sm">
          <p className="text-xs font-bold text-rose-500 uppercase mb-1">
            Total Negativos
          </p>
          <p className="text-3xl font-black text-rose-600">{totalNegativos}</p>
        </div>
      </div>

      <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
        <div className="p-6 border-b border-slate-100">
          <h3 className="text-xl font-bold text-slate-900">Listagem Diária</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs">
            <thead>
              <tr className="bg-slate-50 text-slate-500 font-bold uppercase tracking-wider">
                <th className="px-6 py-4">Nome da Base</th>
                <th className="px-6 py-4">Total Disparos</th>
                <th className="px-6 py-4">Positivos</th>
                <th className="px-6 py-4">Negativos</th>
                <th className="px-6 py-4">Conversão</th>
                <th className="px-6 py-4">Ação</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 text-sm">
              {filteredBases.map((b) => (
                <tr key={b.id} className="hover:bg-slate-50 transition-colors">
                  <td className="px-6 py-4 font-bold text-slate-900">
                    {b.nomeBase}
                  </td>
                  <td className="px-6 py-4 font-bold text-blue-600">
                    {b.totalDisparos}
                  </td>
                  <td className="px-6 py-4 font-bold text-emerald-600">
                    {b.positivos}
                  </td>
                  <td className="px-6 py-4 font-bold text-rose-600">
                    {b.negativos}
                  </td>
                  <td className="px-6 py-4 font-bold text-slate-700">
                    {b.totalDisparos > 0
                      ? ((b.positivos / b.totalDisparos) * 100).toFixed(1)
                      : 0}
                    %
                  </td>
                  <td className="px-6 py-4">
                    <button
                      onClick={async () => {
                        if (window.confirm("Excluir?"))
                          await deleteDoc(
                            doc(db, COLLECTIONS.BASES_DISPARO, b.id),
                          );
                      }}
                      className="text-rose-500 hover:bg-rose-100 p-2 rounded-xl transition-all"
                    >
                      <Trash2 size={18} />
                    </button>
                  </td>
                </tr>
              ))}
              {filteredBases.length === 0 && (
                <tr>
                  <td
                    colSpan={6}
                    className="px-6 py-12 text-center text-slate-400 italic"
                  >
                    Nenhum registro para esta data.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {showModal && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            className="bg-white rounded-[2.5rem] shadow-2xl max-w-lg w-full p-8"
          >
            <div className="flex justify-between items-center mb-8">
              <h3 className="text-2xl font-bold text-slate-900">
                Registrar Métricas da Base
              </h3>
              <button
                onClick={() => setShowModal(false)}
                className="p-2 hover:bg-slate-100 rounded-full transition-all text-slate-400"
              >
                <X size={24} />
              </button>
            </div>

            <form onSubmit={handleSave} className="space-y-6">
              <div>
                <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                  Data do Disparo
                </label>
                <input
                  type="date"
                  className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                  value={formData.data}
                  onChange={(e) =>
                    setFormData({ ...formData, data: e.target.value })
                  }
                  required
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                  Nome da Base
                </label>
                <input
                  type="text"
                  className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                  value={formData.nomeBase}
                  onChange={(e) =>
                    setFormData({ ...formData, nomeBase: e.target.value })
                  }
                  required
                />
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Total
                  </label>
                  <input
                    type="number"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.totalDisparos}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        totalDisparos: Number(e.target.value),
                      })
                    }
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Positivos
                  </label>
                  <input
                    type="number"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.positivos}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        positivos: Number(e.target.value),
                      })
                    }
                    required
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-2 uppercase">
                    Negativos
                  </label>
                  <input
                    type="number"
                    className="w-full px-4 py-3 rounded-2xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none font-bold text-slate-700"
                    value={formData.negativos}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        negativos: Number(e.target.value),
                      })
                    }
                    required
                  />
                </div>
              </div>

              <div className="flex space-x-4 pt-4">
                <button
                  type="submit"
                  className="flex-1 bg-blue-600 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-100 hover:bg-blue-700 transition-all"
                >
                  Registrar
                </button>
                <button
                  type="button"
                  onClick={() => setShowModal(false)}
                  className="px-8 bg-slate-100 text-slate-600 font-bold py-4 rounded-2xl hover:bg-slate-200 transition-all"
                >
                  Cancelar
                </button>
              </div>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
}

const StatCard = ({
  title,
  value,
  icon: Icon,
  color,
  trend,
}: {
  title: string;
  value: string | number;
  icon: any;
  color: string;
  trend?: string;
}) => (
  <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-100 flex items-center justify-between">
    <div>
      <p className="text-sm font-semibold text-slate-500 uppercase tracking-wider mb-1">
        {title}
      </p>
      <p className="text-3xl font-bold text-slate-900">{value}</p>
      {trend && (
        <p className="text-xs text-emerald-600 font-medium mt-1 flex items-center">
          <TrendingUp size={12} className="mr-1" /> {trend}
        </p>
      )}
    </div>
    <div className={cn("p-4 rounded-2xl", color)}>
      <Icon size={24} className="text-white" />
    </div>
  </div>
);

// --- Main App ---

function CampanhasView({
  campanhas,
  onToast,
}: {
  campanhas: Campanha[];
  onToast: (m: string, t?: "success" | "error") => void;
}) {
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isDetailModalOpen, setIsDetailModalOpen] = useState(false);
  const [editingCampanha, setEditingCampanha] = useState<Campanha | null>(null);
  const [selectedCampanha, setSelectedCampanha] = useState<Campanha | null>(
    null,
  );
  const [searchTerm, setSearchTerm] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [startDateFilter, setStartDateFilter] = useState("");
  const [endDateFilter, setEndDateFilter] = useState("");

  const getEffectiveStatus = (camp: Campanha) => {
    const today = new Date().toISOString().split("T")[0];
    if (today < camp.dataInicio) return "Pendente";
    if (today > camp.dataFim) return "Finalizada";
    return "Ativa";
  };

  const filteredCampanhas = useMemo(() => {
    return campanhas.filter((camp) => {
      const effectiveStatus = getEffectiveStatus(camp);
      const matchesSearch = camp.nome
        .toLowerCase()
        .includes(searchTerm.toLowerCase());
      const matchesStatus = !statusFilter || effectiveStatus === statusFilter;

      let matchesDate = true;
      if (startDateFilter && endDateFilter) {
        matchesDate =
          camp.dataInicio <= endDateFilter && camp.dataFim >= startDateFilter;
      } else if (startDateFilter) {
        matchesDate = camp.dataFim >= startDateFilter;
      } else if (endDateFilter) {
        matchesDate = camp.dataInicio <= endDateFilter;
      }
      return matchesSearch && matchesStatus && matchesDate;
    });
  }, [campanhas, searchTerm, statusFilter, startDateFilter, endDateFilter]);

  const handleSave = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const payload = {
      nome: formData.get("nome") as string,
      dataInicio: formData.get("dataInicio") as string,
      dataFim: formData.get("dataFim") as string,
      objetivo: formData.get("objetivo") as string,
      updatedAt: serverTimestamp(),
    };

    const isDuplicate = campanhas.some(
      (c) =>
        c.nome.toLowerCase() === payload.nome.toLowerCase() &&
        c.id !== editingCampanha?.id,
    );
    if (isDuplicate) {
      onToast("Já existe uma campanha com este nome.", "error");
      return;
    }

    try {
      if (editingCampanha) {
        await updateDoc(
          doc(db, COLLECTIONS.CAMPANHAS, editingCampanha.id),
          payload,
        );
        onToast("Campanha atualizada!");
      } else {
        await addDoc(collection(db, COLLECTIONS.CAMPANHAS), {
          ...payload,
          createdAt: serverTimestamp(),
        });
        onToast("Campanha criada!");
      }
      setIsModalOpen(false);
      setEditingCampanha(null);
    } catch (err: any) {
      handleFirestoreError(err, OperationType.WRITE, COLLECTIONS.CAMPANHAS);
      onToast("Erro ao salvar campanha.", "error");
    }
  };

  const handleExport = () => {
    const data = filteredCampanhas.map((c) => ({
      Nome: c.nome,
      "Data Início": c.dataInicio,
      "Data Fim": c.dataFim,
      Status: c.status,
      Objetivo: c.objetivo,
    }));
    exportToExcel(data, "Campanhas");
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return undefined;
        };

        const batch = data.map((item) => {
          const rawStatus = String(getVal(item, "Status", "status") || "")
            .trim()
            .toLowerCase();
          const finalStatus =
            rawStatus === "ativa"
              ? "Ativa"
              : rawStatus === "inativa"
                ? "Inativa"
                : rawStatus === "pendente"
                  ? "Pendente"
                  : "Ativa";

          return {
            nome: String(getVal(item, "Nome", "nome") || "").trim(),
            dataInicio: String(
              getVal(item, "Data Início", "dataInicio", "data_inicio") || "",
            ).trim(),
            dataFim: String(
              getVal(item, "Data Fim", "dataFim", "data_fim") || "",
            ).trim(),
            status: finalStatus,
            objetivo: String(getVal(item, "Objetivo", "objetivo") || "").trim(),
            createdAt: serverTimestamp(),
          };
        });

        let imported = 0;
        let skipped = 0;
        const inserted = new Set();
        for (const entry of batch) {
          if (!entry.nome) continue;
          const isDup =
            campanhas.some(
              (c) => c.nome.trim().toLowerCase() === entry.nome.toLowerCase(),
            ) ||
            Array.from(inserted).some(
              (name: any) =>
                String(name).toLowerCase() === entry.nome.toLowerCase(),
            );
          if (!isDup) {
            await addDoc(collection(db, COLLECTIONS.CAMPANHAS), entry);
            inserted.add(entry.nome);
            imported++;
          } else {
            skipped++;
          }
        }
        onToast(
          `${imported} campanhas importadas! ${skipped > 0 ? `${skipped} ignoradas.` : ""}`,
        );
      } catch (err: any) {
        onToast("Erro ao importar campanhas.", "error");
      }
    });
  };

  return (
    <div className="space-y-6">
      <div className="flex justify-between items-center">
        <h2 className="text-2xl font-bold text-slate-800">Campanhas</h2>
        <div className="flex space-x-2">
          <button
            onClick={() => {
              setEditingCampanha(null);
              setIsModalOpen(true);
            }}
            className="bg-blue-600 text-white px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-700 transition-all shadow-lg shadow-blue-200"
          >
            <Plus size={20} />
            <span>Nova Campanha</span>
          </button>
          <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
            <Upload size={18} />
            <span>Importar</span>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleImport}
              className="hidden"
            />
          </label>
          <button
            onClick={handleExport}
            className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
          >
            <Download size={18} />
            <span>Exportar</span>
          </button>
        </div>
      </div>

      <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex flex-wrap gap-4 items-center">
        <div className="relative flex-1 min-w-[200px]">
          <Search
            className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
            size={18}
          />
          <input
            type="text"
            placeholder="Buscar por nome..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm"
          />
        </div>
        <select
          value={statusFilter}
          onChange={(e) => setStatusFilter(e.target.value)}
          className="px-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm"
        >
          <option value="">Todos os Status</option>
          <option value="Ativa">Ativa</option>
          <option value="Pendente">Pendente</option>
          <option value="Finalizada">Finalizada</option>
        </select>
        <div className="flex items-center gap-2 flex-wrap">
          <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">
            Período:
          </span>
          <input
            type="date"
            value={startDateFilter}
            onChange={(e) => setStartDateFilter(e.target.value)}
            className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm font-semibold text-slate-600"
            title="Data de Início"
          />
          <span className="text-slate-400 text-xs font-bold">até</span>
          <input
            type="date"
            value={endDateFilter}
            onChange={(e) => setEndDateFilter(e.target.value)}
            className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm font-semibold text-slate-600"
            title="Data de Fim"
          />
          {(startDateFilter || endDateFilter) && (
            <button
              onClick={() => {
                setStartDateFilter("");
                setEndDateFilter("");
              }}
              className="text-xs font-bold text-red-500 hover:text-red-700 hover:underline transition-all cursor-pointer px-2"
            >
              Limpar
            </button>
          )}
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
        {filteredCampanhas.map((camp) => {
          const effectiveStatus = getEffectiveStatus(camp);
          return (
            <div
              key={camp.id}
              className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 flex flex-col justify-between cursor-pointer hover:shadow-md transition-all"
              onClick={() => {
                setSelectedCampanha(camp);
                setIsDetailModalOpen(true);
              }}
            >
              <div>
                <div className="flex justify-between items-start mb-4">
                  <h3 className="text-lg font-bold text-slate-900">
                    {camp.nome}
                  </h3>
                  <span
                    className={cn(
                      "px-2 py-1 rounded-full text-[10px] font-bold uppercase",
                      effectiveStatus === "Ativa"
                        ? "bg-emerald-100 text-emerald-600"
                        : effectiveStatus === "Pendente"
                          ? "bg-blue-100 text-blue-600"
                          : "bg-slate-100 text-slate-600",
                    )}
                  >
                    {effectiveStatus}
                  </span>
                </div>
                <p className="text-sm text-slate-500 mb-4 line-clamp-2">
                  {camp.objetivo}
                </p>
                <div className="flex items-center space-x-4 text-xs text-slate-400">
                  <div className="flex items-center space-x-1">
                    <Calendar size={14} />
                    <span>
                      {camp.dataInicio} - {camp.dataFim}
                    </span>
                  </div>
                </div>
              </div>
            </div>
          );
        })}
        {filteredCampanhas.length === 0 && (
          <div className="col-span-full py-12 text-center text-slate-400 italic">
            Nenhuma campanha encontrada.
          </div>
        )}
      </div>

      <AnimatePresence>
        {isDetailModalOpen && selectedCampanha && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white p-8 rounded-3xl shadow-xl w-full max-w-lg space-y-6"
            >
              <h2 className="text-2xl font-bold text-slate-900">
                {selectedCampanha.nome}
              </h2>
              <div className="space-y-4">
                <div>
                  <p className="text-xs font-bold text-slate-500 uppercase">
                    Período
                  </p>
                  <p className="text-sm text-slate-700">
                    {selectedCampanha.dataInicio} - {selectedCampanha.dataFim}
                  </p>
                </div>
                <div>
                  <p className="text-xs font-bold text-slate-500 uppercase">
                    Objetivo
                  </p>
                  <p className="text-sm text-slate-700">
                    {selectedCampanha.objetivo}
                  </p>
                </div>
              </div>
              <div className="flex space-x-4 pt-4">
                <button
                  onClick={() => setIsDetailModalOpen(false)}
                  className="flex-1 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold hover:bg-slate-200 transition-all"
                >
                  Fechar
                </button>
                <button
                  onClick={() => {
                    setEditingCampanha(selectedCampanha);
                    setIsDetailModalOpen(false);
                    setIsModalOpen(true);
                  }}
                  className="flex-1 py-3 bg-blue-600 text-white rounded-xl font-bold hover:bg-blue-700 transition-all"
                >
                  Editar Campanha
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-3xl shadow-xl w-full max-w-md overflow-hidden"
            >
              <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-xl font-bold text-slate-900">
                  {editingCampanha ? "Editar Campanha" : "Nova Campanha"}
                </h3>
                <button
                  onClick={() => setIsModalOpen(false)}
                  className="text-slate-400 hover:text-slate-600"
                >
                  <X size={24} />
                </button>
              </div>
              <form onSubmit={handleSave} className="p-6 space-y-4">
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Nome da Campanha
                  </label>
                  <input
                    name="nome"
                    defaultValue={editingCampanha?.nome}
                    required
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">
                      Início
                    </label>
                    <input
                      type="date"
                      name="dataInicio"
                      defaultValue={editingCampanha?.dataInicio}
                      required
                      className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-bold text-slate-700 mb-1">
                      Fim
                    </label>
                    <input
                      type="date"
                      name="dataFim"
                      defaultValue={editingCampanha?.dataFim}
                      required
                      className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Status
                  </label>
                  <select
                    name="status"
                    defaultValue={editingCampanha?.status || "Ativa"}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                  >
                    <option value="Ativa">Ativa</option>
                    <option value="Pausada">Pausada</option>
                    <option value="Finalizada">Finalizada</option>
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Objetivo
                  </label>
                  <textarea
                    name="objetivo"
                    defaultValue={editingCampanha?.objetivo}
                    rows={3}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <button
                  type="submit"
                  className="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-200"
                >
                  {editingCampanha ? "Salvar Alterações" : "Criar Campanha"}
                </button>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function FiesProuniView({
  data,
  vagas = [],
  onToast,
  profile,
  whatsappMessages,
  periodos,
  botConfig,
  onSendBot,
  onMassSendBot,
}: {
  data: FiesProuniEntry[];
  vagas?: FiesProuniVaga[];
  onToast: (m: string, t?: "success" | "error") => void;
  profile: UserProfile;
  whatsappMessages: WhatsAppMessage[];
  periodos: PeriodoCaptacao[];
  botConfig: BotConfig;
  onSendBot: (tel: string, msg: string, contactName?: string) => void;
  onMassSendBot: (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => void;
}) {
  const [activeTab, setActiveTab] = useState<"lista" | "informacoes">(
    "informacoes",
  );
  const [searchTerm, setSearchTerm] = useState("");
  const [periodoFilter, setPeriodoFilter] = useState("");
  const [tipoFilter, setTipoFilter] = useState("");
  const [listaFilter, setListaFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");
  const [bolsaFilter, setBolsaFilter] = useState("");
  const [cursoFilter, setCursoFilter] = useState("");
  const [situacaoFilter, setSituacaoFilter] = useState("");
  const [vagasPeriodoFilter, setVagasPeriodoFilter] = useState("");
  const [vagasMetodologiaFilter, setVagasMetodologiaFilter] = useState("");
  const [vagasBolsaFilter, setVagasBolsaFilter] = useState("");
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingEntry, setEditingEntry] = useState<FiesProuniEntry | null>(
    null,
  );
  const [isVagaModalOpen, setIsVagaModalOpen] = useState(false);
  const [editingVaga, setEditingVaga] = useState<FiesProuniVaga | null>(null);
  const [cpfInput, setCpfInput] = useState("");

  const isAdmin = true;

  useEffect(() => {
    if (editingEntry) {
      setCpfInput(formatCPF(editingEntry.cpf));
    } else {
      setCpfInput("");
    }
  }, [editingEntry, isModalOpen]);

  const filteredData = data.filter((item) => {
    // Restrict visibility to entries from the same unit, unless admin/gestor
    if (
      profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
      profile.role !== ROLES.GESTOR_COMERCIAL &&
      profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
      profile.role !== ROLES.SSA
    ) {
      if (profile.unidade && item.unidade && item.unidade !== profile.unidade) {
        return false;
      }
    }

    const matchesSearch =
      (item.nome || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.cpf || "").includes(searchTerm) ||
      (item.curso || "").toLowerCase().includes(searchTerm.toLowerCase()) ||
      (item.lista &&
        item.lista.toLowerCase().includes(searchTerm.toLowerCase())) ||
      (item.responsavelEntrevista &&
        item.responsavelEntrevista
          .toLowerCase()
          .includes(searchTerm.toLowerCase())) ||
      (item.status &&
        item.status.toLowerCase().includes(searchTerm.toLowerCase()));
    const matchesPeriodo = !periodoFilter || item.periodo === periodoFilter;
    const matchesTipo = !tipoFilter || item.tipo === tipoFilter;
    const matchesLista = !listaFilter || item.lista === listaFilter;
    const matchesStatus = !statusFilter || item.status === statusFilter;
    const matchesBolsa = !bolsaFilter || item.bolsa === bolsaFilter;
    const matchesCurso = !cursoFilter || item.curso === cursoFilter;
    const matchesSituacao = !situacaoFilter || item.situacao === situacaoFilter;
    return (
      matchesSearch &&
      matchesPeriodo &&
      matchesTipo &&
      matchesLista &&
      matchesStatus &&
      matchesBolsa &&
      matchesCurso &&
      matchesSituacao
    );
  });

  const uniqueListas = Array.from(
    new Set(data.map((i) => i.lista).filter(Boolean)),
  ).sort();
  const uniqueStatuses = Array.from(
    new Set(data.map((i) => i.status).filter(Boolean)),
  ).sort();
  const uniquePeriodos = Array.from(
    new Set(data.map((i) => i.periodo).filter(Boolean)),
  ).sort();
  const uniqueCursos = Array.from(
    new Set(data.map((i) => i.curso).filter(Boolean)),
  ).sort();
  const uniqueSituacoes = Array.from(
    new Set(data.map((i) => i.situacao).filter(Boolean)),
  ).sort();

  const stats = {
    total: filteredData.length,
    pendentes: filteredData.filter((i) => i.docsEntreguesStatus === "Pendente")
      .length,
    parcial: filteredData.filter((i) => i.docsEntreguesStatus === "Parcial")
      .length,
    entregaram: filteredData.filter((i) => i.docsEntreguesStatus === "Sim")
      .length,
    comInscricao: filteredData.filter((i) => i.inscricaoSales).length,
    comMatricula: filteredData.filter((i) => i.numeroMatricula).length,
    emAnalise: filteredData.filter((i) => i.digitalizaStatus === "Em Análise")
      .length,
    concluido: filteredData.filter((i) => i.digitalizaStatus === "Concluído")
      .length,
  };

  const safeVagas = Array.isArray(vagas) ? vagas : [];

  const filteredVagas = safeVagas.filter((item) => {
    if (
      profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
      profile.role !== ROLES.GESTOR_COMERCIAL &&
      profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
      profile.role !== ROLES.SSA
    ) {
      if (profile.unidade && item.unidade && item.unidade !== profile.unidade) {
        return false;
      }
    }
    const matchesPeriodo =
      !vagasPeriodoFilter || item.periodo === vagasPeriodoFilter;
    const matchesMetodologia =
      !vagasMetodologiaFilter || item.metodologia === vagasMetodologiaFilter;
    const matchesBolsa = !vagasBolsaFilter || item.bolsa === vagasBolsaFilter;

    return matchesPeriodo && matchesMetodologia && matchesBolsa;
  });

  const uniqueMetodologias = Array.from(
    new Set(safeVagas.map((i) => i.metodologia).filter(Boolean)),
  ).sort();

  const vagasStats = {
    totalVagas: filteredVagas.reduce(
      (acc, curr) => acc + (Number(curr?.vagas) || 0),
      0,
    ),
    total100: filteredVagas
      .filter((v) => v?.bolsa === "100%")
      .reduce((acc, curr) => acc + (Number(curr?.vagas) || 0), 0),
    total50: filteredVagas
      .filter((v) => v?.bolsa === "50%")
      .reduce((acc, curr) => acc + (Number(curr?.vagas) || 0), 0),
  };

  const handleSaveVaga = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const payload = {
      periodo: formData.get("periodo") as string,
      codCurso: formData.get("codCurso") as string,
      curso: formData.get("curso") as string,
      turno: formData.get("turno") as string,
      metodologia: formData.get("metodologia") as string,
      bolsa: formData.get("bolsa") as "50%" | "100%",
      vagas: parseInt(formData.get("vagas") as string, 10) || 0,
      unidade: (formData.get("unidade") as string) || "",
    };

    try {
      if (editingVaga) {
        await updateDoc(
          doc(db, COLLECTIONS.FIES_PROUNI_VAGAS, editingVaga.id),
          {
            ...payload,
            updatedAt: serverTimestamp(),
          },
        );
        onToast("Vaga atualizada com sucesso!", "success");
      } else {
        await addDoc(collection(db, COLLECTIONS.FIES_PROUNI_VAGAS), {
          ...payload,
          createdAt: serverTimestamp(),
        });
        onToast("Vaga cadastrada com sucesso!", "success");
      }
      setIsVagaModalOpen(false);
      setEditingVaga(null);
    } catch (err) {
      handleFirestoreError(
        err,
        editingVaga ? OperationType.UPDATE : OperationType.CREATE,
        COLLECTIONS.FIES_PROUNI_VAGAS,
      );
    }
  };

  const handleDeleteVaga = async (id: string) => {
    if (!window.confirm("Deseja realmente excluir esta vaga?")) return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.FIES_PROUNI_VAGAS, id));
      onToast("Vaga excluída com sucesso!");
    } catch (err) {
      handleFirestoreError(
        err,
        OperationType.DELETE,
        COLLECTIONS.FIES_PROUNI_VAGAS,
      );
    }
  };

  const handleSave = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const cpf = formData.get("cpf") as string;

    if (!validateCPF(cpf)) {
      onToast("CPF inválido. Por favor, verifique os 11 dígitos.", "error");
      return;
    }

    const cleanCpf = cpf.replace(/\D/g, "");
    const isDuplicate = data.some(
      (item) => item.cpf === cleanCpf && item.id !== editingEntry?.id,
    );
    if (isDuplicate) {
      onToast("Este CPF já está cadastrado no FIES/Prouni.", "error");
      return;
    }

    const payload = {
      nome: formData.get("nome") as string,
      cpf: cpf.replace(/\D/g, ""), // Store only digits
      telefone: formData.get("telefone") as string,
      email: formData.get("email") as string,
      endereco: formData.get("endereco") as string,
      status: formData.get("status") as string,
      tipo: formData.get("tipo") as "FIES" | "PROUNI",
      bolsa: formData.get("bolsa") as "PARCIAL" | "INTEGRAL",
      situacao: formData.get("situacao") as "Candidato" | "Aluno (mesmo curso)" | "Aluno (outro curso)",
      cotaPPI: formData.get("cotaPPI") as "Sim" | "Não",
      metodologia: formData.get("metodologia") as string,
      curso: formData.get("curso") as string,
      inscricaoSales: formData.get("inscricaoSales") as string,
      numeroMatricula: formData.get("numeroMatricula") as string,
      tcbAssinado: formData.get("tcbAssinado") === "on",
      digitalizaStatus: formData.get("digitalizaStatus") as any,
      docsEntreguesStatus: formData.get("docsEntreguesStatus") as any,
      sisprouniStatus: formData.get("sisprouniStatus") as any,
      responsavelEntrevista: formData.get("responsavelEntrevista") as string,
      dataEntrevista: formData.get("dataEntrevista") as string,
      observacao: formData.get("observacao") as string,
      periodo: formData.get("periodo") as string,
      lista: formData.get("lista") as string,
      posicaoRanking: formData.get("posicaoRanking") as string,
      documentosEntregues:
        (formData.get("documentos") as string)
          ?.split(",")
          .map((s) => s.trim())
          .filter(Boolean) || [],
      updatedAt: serverTimestamp(),
    };

    try {
      if (editingEntry) {
        await updateDoc(
          doc(db, COLLECTIONS.FIES_PROUNI, editingEntry.id),
          payload,
        );
        onToast("Registro atualizado!");
      } else {
        await addDoc(collection(db, COLLECTIONS.FIES_PROUNI), {
          ...payload,
          unidade: profile.unidade || "",
          createdAt: serverTimestamp(),
        });
        onToast("Registro cadastrado!");
      }
      setIsModalOpen(false);
      setEditingEntry(null);
    } catch (err: any) {
      handleFirestoreError(err, OperationType.WRITE, COLLECTIONS.FIES_PROUNI);
      onToast("Erro ao salvar registro.", "error");
    }
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      onToast("Importando registros...");
      let successCount = 0;
      let errorCount = 0;

      const getVal = (row: any, ...keys: string[]) => {
        const rowKeys = Object.keys(row);
        for (const key of keys) {
          const foundKey = rowKeys.find(
            (k) => k.toLowerCase() === key.toLowerCase(),
          );
          if (foundKey && row[foundKey] !== undefined) return row[foundKey];
        }
        return undefined;
      };

      for (const row of data) {
        try {
          const rawCpf = String(getVal(row, "CPF", "cpf") || "");
          const cpf = rawCpf.replace(/\D/g, "");
          if (!cpf) continue;

          const payload = {
            nome: String(getVal(row, "Nome", "nome") || ""),
            cpf,
            telefone: String(getVal(row, "Telefone", "telefone") || ""),
            email: String(getVal(row, "Email", "email") || ""),
            endereco: String(getVal(row, "Endereço", "Endereco", "endereco") || ""),
            status: String(getVal(row, "Status", "status") || "Pendente"),
            tipo: (String(getVal(row, "Tipo", "tipo") || "PROUNI").toUpperCase() === "FIES"
              ? "FIES"
              : "PROUNI") as "FIES" | "PROUNI",
            bolsa: String(getVal(row, "Bolsa", "bolsa") || "INTEGRAL").toUpperCase().includes("PARCIAL") ? "PARCIAL" : "INTEGRAL",
            situacao: String(getVal(row, "Situação", "Situacao", "situação", "situacao") || "Candidato")
              .toLowerCase()
              .includes("outro curso")
              ? "Aluno (outro curso)"
              : String(getVal(row, "Situação", "Situacao", "situação", "situacao") || "")
                  .toLowerCase()
                  .includes("mesmo curso")
              ? "Aluno (mesmo curso)"
              : "Candidato",
            cotaPPI: String(getVal(row, "Cota PPI", "cota ppi", "Cota_PPI", "cotappi") || "").toLowerCase().includes("sim") ? "Sim" : "Não",
            curso: String(getVal(row, "Curso", "curso") || ""),
            posicaoRanking: String(getVal(row, "Ranking", "ranking") || ""),
            lista: String(getVal(row, "Lista", "lista") || ""),
            periodo: String(getVal(row, "Periodo", "Período", "periodo", "período") || ""),
            metodologia: String(getVal(row, "Metodologia", "metodologia") || ""),
            responsavelEntrevista: String(getVal(row, "Responsável Entrevista", "Responsavel Entrevista", "responsavel entrevista") || ""),
            dataEntrevista: String(getVal(row, "Data Entrevista", "data entrevista") || ""),
            docsEntreguesStatus: String(
              getVal(row, "Status Docs", "status docs") || "Pendente",
            ) as any,
            inscricaoSales: String(getVal(row, "Inscrição Sales", "Inscricao Sales", "inscricao sales") || ""),
            numeroMatricula: String(getVal(row, "Número Matrícula", "Numero Matricula", "numero matricula") || ""),
            digitalizaStatus: String(
              getVal(row, "Status Digitaliza", "status digitaliza") || "Pendente",
            ) as any,
            sisprouniStatus: String(getVal(row, "SISPROUNI", "sisprouni") || "Pendente") as any,
            tcbAssinado: String(getVal(row, "TCB Assinado", "tcb assinado")).toLowerCase() === "sim",
            documentosEntregues: String(getVal(row, "Documentos Entregues", "documentos entregues") || "")
              .split(",")
              .map((s) => s.trim())
              .filter(Boolean),
            observacao: String(getVal(row, "Observação", "Observacao", "observacao") || ""),
            unidade: profile.unidade || "",
            createdAt: serverTimestamp(),
            updatedAt: serverTimestamp(),
          };

          await addDoc(collection(db, COLLECTIONS.FIES_PROUNI), payload);
          successCount++;
        } catch (err) {
          console.error("Erro ao importar registro Fies/Prouni:", err);
          errorCount++;
        }
      }

      onToast(
        `Importação concluída: ${successCount} sucesso, ${errorCount} erros`,
        successCount > 0 ? "success" : "error",
      );
    });
    e.target.value = "";
  };

  const handleExport = () => {
    const exportData = filteredData.map((item) => ({
      Nome: item.nome,
      CPF: item.cpf,
      Telefone: item.telefone || "",
      Email: item.email || "",
      Endereço: item.endereco || "",
      Status: item.status || "",
      Tipo: item.tipo,
      Bolsa: item.bolsa,
      Situação: item.situacao || "",
      "Cota PPI": item.cotaPPI || "",
      Curso: item.curso,
      Ranking: item.posicaoRanking || "",
      Lista: item.lista || "",
      Periodo: item.periodo || "",
      Metodologia: item.metodologia || "",
      "Responsável Entrevista": item.responsavelEntrevista || "",
      "Data Entrevista": item.dataEntrevista || "",
      "Status Docs": item.docsEntreguesStatus || "",
      "Inscrição Sales": item.inscricaoSales || "",
      "Número Matrícula": item.numeroMatricula || "",
      "Status Digitaliza": item.digitalizaStatus,
      SISPROUNI: item.sisprouniStatus || "Pendente",
      "TCB Assinado": item.tcbAssinado ? "Sim" : "Não",
      "Documentos Entregues": item.documentosEntregues?.join(", ") || "",
      Observação: item.observacao || "",
    }));
    exportToExcel(exportData, "Fies_Prouni");
  };

  const handleExportVagas = () => {
    const exportData = vagas.map((v) => ({
      Período: v.periodo || "",
      "Cod. Curso": v.codCurso || "",
      Curso: v.curso || "",
      Turno: v.turno || "",
      Metodologia: v.metodologia || "",
      Bolsa: v.bolsa || "",
      Vagas: v.vagas || 0,
      Unidade: v.unidade || "",
    }));
    exportToExcel(exportData, "Fies_Prouni_Vagas");
  };

  const handleImportVagas = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      onToast("Importando vagas...");
      let successCount = 0;
      let errorCount = 0;

      const getVal = (row: any, ...keys: string[]) => {
        const rowKeys = Object.keys(row);
        for (const key of keys) {
          const foundKey = rowKeys.find(
            (k) => k.toLowerCase() === key.toLowerCase(),
          );
          if (foundKey && row[foundKey] !== undefined) return row[foundKey];
        }
        return undefined;
      };

      for (const row of data) {
        try {
          const payload = {
            periodo: String(getVal(row, "Período", "Periodo", "período", "periodo") || ""),
            codCurso: String(getVal(row, "Cod. Curso", "cod curso", "cod. curso", "codCurso") || ""),
            curso: String(getVal(row, "Curso", "curso") || ""),
            turno: String(getVal(row, "Turno", "turno") || ""),
            metodologia: String(getVal(row, "Metodologia", "metodologia") || ""),
            bolsa: String(getVal(row, "Bolsa", "bolsa") || "") as "50%" | "100%",
            vagas: parseInt(String(getVal(row, "Vagas", "vagas")), 10) || 0,
            unidade: String(getVal(row, "Unidade", "unidade") || ""),
            createdAt: serverTimestamp(),
          };

          if (payload.curso && payload.periodo && payload.bolsa) {
            await addDoc(
              collection(db, COLLECTIONS.FIES_PROUNI_VAGAS),
              payload,
            );
            successCount++;
          }
        } catch (err) {
          console.error("Erro ao importar vaga:", err);
          errorCount++;
        }
      }

      onToast(
        `Importação concluída: ${successCount} sucesso, ${errorCount} erros`,
        successCount > 0 ? "success" : "error",
      );
    });
    e.target.value = "";
  };

  const [selectedEntries, setSelectedEntries] = useState<string[]>([]);

  const handleBulkDelete = async () => {
    if (selectedEntries.length === 0) return;
    if (
      window.confirm(
        `Deseja excluir ${selectedEntries.length} registros Fies/Prouni selecionados?`,
      )
    ) {
      try {
        for (const id of selectedEntries) {
          await deleteDoc(doc(db, COLLECTIONS.FIES_PROUNI, id));
        }
        onToast(`${selectedEntries.length} registros removidos.`);
        setSelectedEntries([]);
      } catch (err: any) {
        onToast("Erro ao excluir registros.", "error");
      }
    }
  };

  const handleDeleteIndividual = async (id: string) => {
    if (window.confirm("Deseja excluir este registro?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.FIES_PROUNI, id));
        onToast("Registro removido.");
      } catch (err: any) {
        onToast("Erro ao excluir registro.", "error");
      }
    }
  };

  const toggleSelect = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedEntries([...selectedEntries, id]);
    } else {
      setSelectedEntries(selectedEntries.filter((s) => s !== id));
    }
  };

  const toggleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedEntries(filteredData.map((b) => b.id));
    } else {
      setSelectedEntries([]);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div className="flex items-center gap-4">
          <h2 className="text-2xl font-bold text-gray-800">
            Acompanhamento Fies/Prouni
          </h2>
          <div className="flex bg-slate-100 p-1 rounded-xl">
            <button
              onClick={() => setActiveTab("informacoes")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === "informacoes" ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}
            >
              Informações
            </button>
            <button
              onClick={() => setActiveTab("lista")}
              className={`px-4 py-2 rounded-lg text-sm font-bold transition-all ${activeTab === "lista" ? "bg-white text-blue-600 shadow-sm" : "text-slate-500 hover:text-slate-700"}`}
            >
              Lista
            </button>
          </div>
        </div>
        <div className="flex space-x-2">
          {activeTab === "lista" ? (
            <>
              <button
                onClick={() => {
                  setEditingEntry(null);
                  setIsModalOpen(true);
                }}
                className="bg-indigo-600 text-white px-4 py-2 rounded-lg flex items-center space-x-2 hover:bg-indigo-700 transition-colors"
              >
                <Plus size={20} />
                <span>Novo Cadastro</span>
              </button>
              <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
                <Upload size={18} />
                <span>Importar Lista</span>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  className="hidden"
                  onChange={handleImport}
                />
              </label>
              <button
                onClick={handleExport}
                className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
              >
                <Download size={18} />
                <span>Exportar Excel</span>
              </button>
            </>
          ) : (
            <>
              <button
                onClick={() => {
                  setEditingVaga(null);
                  setIsVagaModalOpen(true);
                }}
                className="bg-indigo-600 text-white px-4 py-2 rounded-lg flex items-center space-x-2 hover:bg-indigo-700 transition-colors"
              >
                <Plus size={20} />
                <span>Nova Vaga</span>
              </button>
              <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
                <Upload size={18} />
                <span>Importar Vagas</span>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  className="hidden"
                  onChange={handleImportVagas}
                />
              </label>
              <button
                onClick={handleExportVagas}
                className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
              >
                <Download size={18} />
                <span>Exportar Excel</span>
              </button>
            </>
          )}
        </div>
      </div>

      {activeTab === "lista" && (
        <>
          {/* Dashboard Cards */}
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
            <StatCard
              title="Total Candidatos"
              value={stats.total}
              icon={Users}
              color="bg-blue-500"
            />
            <StatCard
              title="Pendentes Doc"
              value={stats.pendentes}
              icon={AlertCircle}
              color="bg-red-500"
            />
            <StatCard
              title="Docs Parciais"
              value={stats.parcial}
              icon={Clock}
              color="bg-amber-500"
            />
            <StatCard
              title="Docs Entregues"
              value={stats.entregaram}
              icon={CheckCircle2}
              color="bg-green-500"
            />
            <StatCard
              title="Com Inscrição"
              value={stats.comInscricao}
              icon={FileText}
              color="bg-indigo-500"
            />
            <StatCard
              title="Com Matrícula"
              value={stats.comMatricula}
              icon={GraduationCap}
              color="bg-purple-500"
            />
            <StatCard
              title="Em Análise"
              value={stats.emAnalise}
              icon={Clock}
              color="bg-amber-500"
            />
            <StatCard
              title="Docs OK"
              value={stats.concluido}
              icon={ShieldCheck}
              color="bg-emerald-500"
            />
          </div>

          {/* Filters */}
          <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex flex-wrap gap-4">
            <div className="flex-1 min-w-[200px] relative">
              <Search
                className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
                size={18}
              />
              <input
                type="text"
                placeholder="Pesquisar por nome, CPF ou curso..."
                className="w-full pl-10 pr-4 py-2 border border-gray-200 rounded-lg focus:ring-2 focus:ring-indigo-500 outline-none"
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
              />
            </div>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={periodoFilter}
              onChange={(e) => setPeriodoFilter(e.target.value)}
            >
              <option value="">Todos os Períodos</option>
              {uniquePeriodos.map((p) => (
                <option key={p} value={p}>
                  {p}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 max-w-[200px]"
              value={cursoFilter}
              onChange={(e) => setCursoFilter(e.target.value)}
            >
              <option value="">Todos os Cursos</option>
              {uniqueCursos.map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={tipoFilter}
              onChange={(e) => setTipoFilter(e.target.value)}
            >
              <option value="">Fies & Prouni</option>
              <option value="FIES">Apenas FIES</option>
              <option value="PROUNI">Apenas PROUNI</option>
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={bolsaFilter}
              onChange={(e) => setBolsaFilter(e.target.value)}
            >
              <option value="">Todas as Bolsas</option>
              <option value="INTEGRAL">INTEGRAL</option>
              <option value="PARCIAL">PARCIAL</option>
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={situacaoFilter}
              onChange={(e) => setSituacaoFilter(e.target.value)}
            >
              <option value="">Todas as Situações</option>
              {uniqueSituacoes.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={listaFilter}
              onChange={(e) => setListaFilter(e.target.value)}
            >
              <option value="">Todas as Listas</option>
              {uniqueListas.map((l) => (
                <option key={l} value={l}>
                  {l}
                </option>
              ))}
            </select>
            <select
              className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="">Todos os Status</option>
              {uniqueStatuses.map((s) => (
                <option key={s} value={s}>
                  {s}
                </option>
              ))}
            </select>
          </div>

          {/* Table */}
          <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead className="bg-gray-50 border-b border-gray-100">
                  <tr>
                    <th className="px-6 py-4">
                      <input
                        type="checkbox"
                        checked={
                          selectedEntries.length === filteredData.length &&
                          filteredData.length > 0
                        }
                        onChange={(e) => toggleSelectAll(e.target.checked)}
                      />
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Candidato
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Lista/Status
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Tipo/Bolsa
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Curso/Metodologia
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Documentação
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      Digitaliza
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600">
                      TCB
                    </th>
                    <th className="px-6 py-4 text-sm font-semibold text-gray-600 flex items-center gap-4">
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkDelete}
                          className="text-rose-600 font-bold hover:underline"
                        >
                          excluir selecionados
                        </button>
                      )}
                      {selectedEntries.length > 0 && botConfig.url && (
                        <button
                          onClick={() => {
                            const selectedObjs = data.filter((g) =>
                              selectedEntries.includes(g.id),
                            );
                            const payloads = selectedObjs.map((item) => {
                              const isMatAcadOk =
                                item.numeroMatricula &&
                                item.numeroMatricula.trim().length > 0;
                              const type = isMatAcadOk
                                ? "fiesProuni_1"
                                : "fiesProuni_0";
                              const msgTemplate = whatsappMessages.find(
                                (m) =>
                                  m.tipo === type || m.tipo === "fiesProuni",
                              );
                              const text = msgTemplate
                                ? replaceMessageVariables(
                                    msgTemplate.texto,
                                    item,
                                  )
                                : `Olá ${item.nome}, tudo bem?`;
                              return {
                                telefone: item.telefone,
                                message: text,
                                nome: item.nome,
                              };
                            });
                            onMassSendBot(payloads);
                            setSelectedEntries([]);
                          }}
                          className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1"
                        >
                          <Bot size={14} /> Em Massa
                        </button>
                      )}
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-50">
                  {filteredData.map((item) => (
                    <tr
                      key={item.id}
                      className="hover:bg-gray-50/50 transition-colors"
                    >
                      <td className="px-6 py-4">
                        <input
                          type="checkbox"
                          checked={selectedEntries.includes(item.id)}
                          onChange={(e) =>
                            toggleSelect(item.id, e.target.checked)
                          }
                        />
                      </td>
                      <td className="px-6 py-4">
                        <div className="font-medium text-gray-900">
                          {item.nome || "Sem nome"}
                        </div>
                        <div className="text-[10px] font-bold text-indigo-500">
                          Ranking: {item.posicaoRanking || "-"}
                        </div>
                        <div className="text-xs text-gray-500">
                          {formatCPF(item.cpf || "")}
                        </div>
                        <div className="text-xs text-gray-400">
                          {item.periodo}
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <div className="text-sm font-bold text-indigo-600">
                          {item.lista || "-"}
                        </div>
                        <div className="text-[10px] text-slate-500 uppercase font-bold">
                          {item.status || "Sem Status"}
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2 py-1 rounded-full text-xs font-bold ${item.tipo === "FIES" ? "bg-blue-100 text-blue-700" : "bg-orange-100 text-orange-700"}`}
                        >
                          {item.tipo}
                        </span>
                        <div className="text-xs text-gray-500 mt-1 font-bold">
                          {item.bolsa}
                        </div>
                        {item.situacao && (
                          <div className="text-[10px] text-gray-400 mt-1">
                            Sit.: {item.situacao}
                          </div>
                        )}
                        {item.cotaPPI && (
                          <div className="text-[10px] text-gray-400">
                            PPI: {item.cotaPPI}
                          </div>
                        )}
                      </td>
                      <td className="px-6 py-4">
                        <div className="text-sm text-gray-700">
                          {item.curso}
                        </div>
                        <div className="text-xs text-gray-500">
                          {item.metodologia}
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2 py-1 rounded-full text-[10px] font-bold ${
                            item.docsEntreguesStatus === "Sim"
                              ? "bg-green-100 text-green-700"
                              : item.docsEntreguesStatus === "Parcial"
                                ? "bg-amber-100 text-amber-700"
                                : "bg-red-100 text-red-700"
                          }`}
                        >
                          {item.docsEntreguesStatus || "Pendente"}
                        </span>
                        <div className="text-[10px] text-slate-400 mt-1">
                          {item.documentosEntregues?.length || 0} docs
                        </div>
                      </td>
                      <td className="px-6 py-4">
                        <span
                          className={`px-2 py-1 rounded-full text-xs font-medium ${
                            item.digitalizaStatus === "Concluído"
                              ? "bg-green-100 text-green-700"
                              : item.digitalizaStatus === "Em Análise"
                                ? "bg-amber-100 text-amber-700"
                                : "bg-gray-100 text-gray-600"
                          }`}
                        >
                          {item.digitalizaStatus}
                        </span>
                      </td>
                      <td className="px-6 py-4">
                        {item.tcbAssinado ? (
                          <CheckCircle2 className="text-green-500" size={20} />
                        ) : (
                          <Clock className="text-gray-300" size={20} />
                        )}
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center space-x-2">
                          <button
                            onClick={() => {
                              setEditingEntry(item);
                              setIsModalOpen(true);
                            }}
                            className="text-indigo-600 hover:text-indigo-800 font-medium text-sm p-2 hover:bg-indigo-50 rounded-lg transition-all"
                            title="Editar"
                          >
                            <Edit2 size={18} />
                          </button>
                          {item.telefone && (
                            <>
                              {botConfig.url && (
                                <button
                                  onClick={() => {
                                    const isMatAcadOk =
                                      item.numeroMatricula &&
                                      item.numeroMatricula.trim().length > 0;
                                    const type = isMatAcadOk
                                      ? "fiesProuni_1"
                                      : "fiesProuni_0";
                                    const msgObj = whatsappMessages.find(
                                      (m) =>
                                        m.tipo === type ||
                                        m.tipo === "fiesProuni",
                                    );
                                    const msg = replaceMessageVariables(
                                      msgObj
                                        ? msgObj.texto
                                        : `Olá [nome], tudo bem?`,
                                      item,
                                    );
                                    onSendBot(item.telefone, msg);
                                  }}
                                  className="text-blue-600 hover:text-blue-800 p-2 hover:bg-blue-50 rounded-lg transition-all"
                                  title="Enviar pelo Bot ARGO'S"
                                >
                                  <Bot size={18} />
                                </button>
                              )}
                              <a
                                href={getWhatsAppUrl(
                                  item.telefone,
                                  (() => {
                                    const isMatAcadOk =
                                      item.numeroMatricula &&
                                      item.numeroMatricula.trim().length > 0;
                                    const type = isMatAcadOk
                                      ? "fiesProuni_1"
                                      : "fiesProuni_0";
                                    const msg = whatsappMessages.find(
                                      (m) =>
                                        m.tipo === type ||
                                        m.tipo === "fiesProuni",
                                    );
                                    if (msg)
                                      return replaceMessageVariables(
                                        msg.texto,
                                        item,
                                      );
                                    return `Olá ${item.nome}, tudo bem?`;
                                  })(),
                                )}
                                target="_blank"
                                rel="noopener noreferrer"
                                className="text-emerald-600 hover:text-emerald-800 p-2 hover:bg-emerald-50 rounded-lg transition-all"
                                title="Enviar WhatsApp"
                              >
                                <MessageSquare size={18} />
                              </a>
                            </>
                          )}
                          <button
                            onClick={() => handleDeleteIndividual(item.id)}
                            className="text-rose-400 hover:text-rose-600 p-2 hover:bg-rose-50 rounded-lg transition-all"
                            title="Excluir"
                          >
                            <Trash2 size={18} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {activeTab === "informacoes" && (
        <div className="space-y-6">
          {/* Filters for Vagas */}
          <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 flex flex-wrap gap-4">
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">
                Semestre / Período
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasPeriodoFilter}
                onChange={(e) => setVagasPeriodoFilter(e.target.value)}
              >
                <option value="">Todos os Semestres</option>
                {periodos.map((p) => (
                  <option key={p.id} value={p.nome}>
                    {p.nome}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">
                Metodologia
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasMetodologiaFilter}
                onChange={(e) => setVagasMetodologiaFilter(e.target.value)}
              >
                <option value="">Todas as Metodologias</option>
                {uniqueMetodologias.map((m) => (
                  <option key={m} value={m}>
                    {m}
                  </option>
                ))}
              </select>
            </div>
            <div className="flex flex-col gap-1">
              <label className="text-[10px] font-bold text-slate-400 uppercase ml-1">
                Bolsa
              </label>
              <select
                className="px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 text-sm"
                value={vagasBolsaFilter}
                onChange={(e) => setVagasBolsaFilter(e.target.value)}
              >
                <option value="">Todas as Bolsas</option>
                <option value="50%">50%</option>
                <option value="100%">100%</option>
              </select>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4">
              <div className="p-3 bg-blue-500 rounded-xl text-white">
                <FileText size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">
                  Total de Vagas
                </p>
                <h3 className="text-2xl font-bold text-slate-900">
                  {vagasStats.totalVagas}
                </h3>
              </div>
            </div>
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4 border-l-4 border-l-emerald-500">
              <div className="p-3 bg-emerald-500 rounded-xl text-white">
                <ShieldCheck size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">
                  Bolsas 100%
                </p>
                <h3 className="text-2xl font-bold text-emerald-600">
                  {vagasStats.total100}
                </h3>
              </div>
            </div>
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm flex items-center space-x-4 border-l-4 border-l-blue-500">
              <div className="p-3 bg-blue-500 rounded-xl text-white">
                <BookOpen size={24} />
              </div>
              <div>
                <p className="text-sm font-medium text-slate-500">Bolsas 50%</p>
                <h3 className="text-2xl font-bold text-blue-600">
                  {vagasStats.total50}
                </h3>
              </div>
            </div>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider">
                    <th className="p-4 font-bold border-b border-slate-100">
                      Período
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Cod. Curso
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Curso
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Turno
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Metodologia
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Bolsa
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-center">
                      Vagas
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100 text-right">
                      Ações
                    </th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-50">
                  {filteredVagas.length === 0 ? (
                    <tr>
                      <td
                        colSpan={8}
                        className="p-8 text-center text-slate-400"
                      >
                        Nenhuma vaga cadastrada.
                      </td>
                    </tr>
                  ) : (
                    filteredVagas.map((vaga) => (
                      <tr
                        key={vaga.id}
                        className="hover:bg-slate-50/50 transition-colors"
                      >
                        <td className="p-4 font-medium text-slate-700">
                          {vaga.periodo}
                        </td>
                        <td className="p-4 font-mono text-slate-500">
                          {vaga.codCurso}
                        </td>
                        <td className="p-4 font-medium text-slate-800">
                          {vaga.curso}
                        </td>
                        <td className="p-4 text-center text-slate-600">
                          {vaga.turno}
                        </td>
                        <td className="p-4 text-center text-slate-600">
                          {vaga.metodologia}
                        </td>
                        <td className="p-4 text-center">
                          <span
                            className={`inline-block px-2.5 py-1 rounded-full text-xs font-bold ${vaga.bolsa === "100%" ? "bg-emerald-100 text-emerald-700" : "bg-blue-100 text-blue-700"}`}
                          >
                            {vaga.bolsa}
                          </span>
                        </td>
                        <td className="p-4 text-center font-bold text-slate-800">
                          {vaga.vagas}
                        </td>
                        <td className="p-4 text-right">
                          <div className="flex justify-end gap-2">
                            <button
                              onClick={() => {
                                setEditingVaga(vaga);
                                setIsVagaModalOpen(true);
                              }}
                              className="text-slate-400 hover:text-blue-600 p-2 hover:bg-blue-50 rounded-lg transition-all"
                            >
                              <Edit2 size={18} />
                            </button>
                            <button
                              onClick={() => handleDeleteVaga(vaga.id)}
                              className="text-rose-400 hover:text-rose-600 p-2 hover:bg-rose-50 rounded-lg transition-all"
                            >
                              <Trash2 size={18} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* Modal */}
      <AnimatePresence>
        {isModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-2xl max-h-[90vh] overflow-y-auto"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center sticky top-0 bg-white z-10">
                <h3 className="text-xl font-bold text-gray-800">
                  {editingEntry
                    ? "Editar Registro"
                    : "Novo Cadastro Fies/Prouni"}
                </h3>
                <button
                  onClick={() => setIsModalOpen(false)}
                  className="text-gray-400 hover:text-gray-600"
                >
                  <X size={24} />
                </button>
              </div>
              <form onSubmit={handleSave} className="p-6 space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Nome Completo
                    </label>
                    <input
                      name="nome"
                      defaultValue={editingEntry?.nome}
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      CPF
                    </label>
                    <input
                      name="cpf"
                      value={cpfInput}
                      onChange={(e) =>
                        setCpfInput(formatCPF(e.target.value || ""))
                      }
                      required
                      placeholder="000.000.000-00"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Telefone
                    </label>
                    <input
                      name="telefone"
                      defaultValue={editingEntry?.telefone}
                      onChange={(e) => {
                        e.target.value = formatPhone(e.target.value);
                      }}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Email
                    </label>
                    <input
                      name="email"
                      type="email"
                      defaultValue={editingEntry?.email}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Endereço
                    </label>
                    <input
                      name="endereco"
                      defaultValue={editingEntry?.endereco}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Status
                    </label>
                    <select
                      name="status"
                      defaultValue={editingEntry?.status || "Pendente"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Aprovado">Aprovado</option>
                      <option value="Reprovado">Reprovado</option>
                      <option value="Em Análise">Em Análise</option>
                      <option value="Desistente">Desistente</option>
                      <option value="Não compareceu">Não compareceu</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Tipo
                    </label>
                    <select
                      name="tipo"
                      defaultValue={editingEntry?.tipo || "PROUNI"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="FIES">FIES</option>
                      <option value="PROUNI">PROUNI</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Bolsa
                    </label>
                    <select
                      name="bolsa"
                      defaultValue={editingEntry?.bolsa || "INTEGRAL"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="INTEGRAL">INTEGRAL</option>
                      <option value="PARCIAL">PARCIAL</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Situação
                    </label>
                    <select
                      name="situacao"
                      defaultValue={editingEntry?.situacao || "Candidato"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Candidato">Candidato</option>
                      <option value="Aluno (mesmo curso)">Aluno (mesmo curso)</option>
                      <option value="Aluno (outro curso)">Aluno (outro curso)</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Cota PPI
                    </label>
                    <select
                      name="cotaPPI"
                      defaultValue={editingEntry?.cotaPPI || "Não"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Sim">Sim</option>
                      <option value="Não">Não</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Período
                    </label>
                    <input
                      name="periodo"
                      defaultValue={editingEntry?.periodo}
                      placeholder="Ex: 2025.1"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Lista
                    </label>
                    <input
                      name="lista"
                      defaultValue={editingEntry?.lista}
                      placeholder="Ex: Lista 1"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Posição no Ranking
                    </label>
                    <input
                      name="posicaoRanking"
                      defaultValue={editingEntry?.posicaoRanking}
                      placeholder="Ex: 15º"
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Curso
                    </label>
                    <input
                      name="curso"
                      defaultValue={editingEntry?.curso}
                      required
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Metodologia
                    </label>
                    <input
                      name="metodologia"
                      defaultValue={editingEntry?.metodologia}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Inscrição Sales
                    </label>
                    <input
                      name="inscricaoSales"
                      defaultValue={editingEntry?.inscricaoSales}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Número Matrícula
                    </label>
                    <input
                      name="numeroMatricula"
                      defaultValue={editingEntry?.numeroMatricula}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Status Digitaliza
                    </label>
                    <select
                      name="digitalizaStatus"
                      defaultValue={
                        editingEntry?.digitalizaStatus || "Não Postado"
                      }
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Não Postado">Não Postado</option>
                      <option value="Em Análise">Em Análise</option>
                      <option value="Concluído">Concluído</option>
                      <option value="Documento reprovado">
                        Documento reprovado
                      </option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Status Documentos
                    </label>
                    <select
                      name="docsEntreguesStatus"
                      defaultValue={
                        editingEntry?.docsEntreguesStatus || "Pendente"
                      }
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Parcial">Parcial</option>
                      <option value="Sim">Sim (Tudo Entregue)</option>
                      <option value="Não compareceu">Não compareceu</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      SISPROUNI
                    </label>
                    <select
                      name="sisprouniStatus"
                      defaultValue={editingEntry?.sisprouniStatus || "Pendente"}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                    >
                      <option value="Pendente">Pendente</option>
                      <option value="Aprovado">Aprovado</option>
                      <option value="Reprovado">Reprovado</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Responsável Entrevista
                    </label>
                    <input
                      name="responsavelEntrevista"
                      defaultValue={
                        editingEntry?.responsavelEntrevista || profile.name
                      }
                      readOnly={!isAdmin}
                      className={`w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 ${!isAdmin ? "bg-slate-50 text-slate-500" : ""}`}
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Data Entrevista
                    </label>
                    <input
                      name="dataEntrevista"
                      type="date"
                      defaultValue={
                        editingEntry?.dataEntrevista ||
                        new Date().toISOString().split("T")[0]
                      }
                      readOnly={!isAdmin}
                      className={`w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500 ${!isAdmin ? "bg-slate-50 text-slate-500" : ""}`}
                    />
                  </div>
                  <div className="flex items-center space-x-2 pt-8">
                    <input
                      type="checkbox"
                      name="tcbAssinado"
                      defaultChecked={editingEntry?.tcbAssinado}
                      className="w-5 h-5 text-indigo-600 rounded focus:ring-indigo-500"
                    />
                    <label className="text-sm font-medium text-gray-700">
                      TCB Assinado
                    </label>
                  </div>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Documentos Entregues (separados por vírgula)
                  </label>
                  <input
                    name="documentos"
                    defaultValue={editingEntry?.documentosEntregues?.join(", ")}
                    placeholder="Ex: RG, CPF, Diploma"
                    className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Observações / O que falta
                  </label>
                  <textarea
                    name="observacao"
                    defaultValue={editingEntry?.observacao}
                    rows={3}
                    className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
                <div className="pt-4">
                  <button
                    type="submit"
                    className="w-full bg-indigo-600 text-white py-3 rounded-xl font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-200"
                  >
                    {editingEntry ? "Salvar Alterações" : "Cadastrar Candidato"}
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}

        {isVagaModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-lg"
            >
              <div className="p-6 border-b border-gray-100 flex justify-between items-center bg-white rounded-t-2xl">
                <h3 className="text-xl font-bold text-gray-800">
                  {editingVaga ? "Editar Vaga" : "Nova Vaga"}
                </h3>
                <button
                  onClick={() => setIsVagaModalOpen(false)}
                  className="text-gray-400 hover:text-gray-600 transition-colors"
                >
                  <X size={24} />
                </button>
              </div>

              <form onSubmit={handleSaveVaga} className="p-6 space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Período
                    </label>
                    <input
                      name="periodo"
                      required
                      defaultValue={editingVaga?.periodo}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Cod. Curso
                    </label>
                    <input
                      name="codCurso"
                      required
                      defaultValue={editingVaga?.codCurso}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div className="md:col-span-2">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Curso
                    </label>
                    <input
                      name="curso"
                      required
                      defaultValue={editingVaga?.curso}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Turno
                    </label>
                    <select
                      name="turno"
                      required
                      defaultValue={editingVaga?.turno}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione</option>
                      <option value="Manhã">Manhã</option>
                      <option value="Tarde">Tarde</option>
                      <option value="Noite">Noite</option>
                      <option value="Integral">Integral</option>
                      <option value="EAD">EAD</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Metodologia
                    </label>
                    <select
                      name="metodologia"
                      required
                      defaultValue={editingVaga?.metodologia}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione</option>
                      <option value="Presencial">Presencial</option>
                      <option value="EAD">EAD</option>
                      <option value="Híbrido">Híbrido</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Bolsa
                    </label>
                    <select
                      name="bolsa"
                      required
                      defaultValue={editingVaga?.bolsa}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    >
                      <option value="">Selecione</option>
                      <option value="50%">50%</option>
                      <option value="100%">100%</option>
                    </select>
                  </div>
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                      Vagas
                    </label>
                    <input
                      name="vagas"
                      type="number"
                      required
                      min="0"
                      defaultValue={editingVaga?.vagas}
                      className="w-full px-4 py-2 border border-gray-200 rounded-lg outline-none focus:ring-2 focus:ring-blue-500"
                    />
                  </div>
                </div>
                <div className="pt-4">
                  <button
                    type="submit"
                    className="w-full bg-blue-600 text-white py-3 rounded-xl font-bold hover:bg-blue-700 transition-colors shadow-lg shadow-blue-200"
                  >
                    {editingVaga ? "Salvar Alterações" : "Cadastrar Vaga"}
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

export default function App() {
  const [analysisSchemes, setAnalysisSchemes] = useState<AnalysisScheme[]>([]);
  const [user, setUser] = useState<User | null>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [currentView, setCurrentView] = useState(() => {
    const params = new URLSearchParams(window.location.search);
    return params.get("view") || "cadastro";
  });
  const [toast, setToast] = useState<{
    message: string;
    type: "success" | "error";
  } | null>(null);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [insumosBaixas, setInsumosBaixas] = useState<InsumoBaixa[]>([]);
  const [isOnline, setIsOnline] = useState<boolean>(
    typeof navigator !== "undefined" ? navigator.onLine : true,
  );

  useEffect(() => {
    if (typeof window === "undefined") return;
    const handleOnline = () => setIsOnline(true);
    const handleOffline = () => setIsOnline(false);

    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);

    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
    };
  }, []);

  // Data States
  const [salesContacts, setSalesContacts] = useState<SalesContact[]>([]);
  const [leads, setLeads] = useState<Lead[]>([]);
  const [users, setUsers] = useState<UserProfile[]>([]);
  const [bases, setBases] = useState<BaseEntry[]>([]);
  const [gap, setGap] = useState<GapEntry[]>([]);
  const [isencoes, setIsencoes] = useState<IsencaoEntry[]>([]);
  const [solicitacoesManutencao, setSolicitacoesManutencao] = useState<SolicitacaoManutencao[]>([]);
  const [fiesProuni, setFiesProuni] = useState<FiesProuniEntry[]>([]);
  const [fiesProuniVagas, setFiesProuniVagas] = useState<FiesProuniVaga[]>([]);
  const [campanhas, setCampanhas] = useState<Campanha[]>([]);
  const [bomDia, setBomDia] = useState<BomDiaCaptacao[]>([]);
  const [forecast, setForecast] = useState<ForecastCaptacao[]>([]);
  const [metaDia, setMetaDia] = useState<MetaDia[]>([]);
  const [metaSM, setMetaSM] = useState<MetaSM[]>([]);
  const [metaCursos, setMetaCursos] = useState<MetaCurso[]>([]);
  const [qgLigacoes, setQgLigacoes] = useState<QgLigacao[]>([]);
  const [planner, setPlanner] = useState<PlannerTask[]>([]);
  const [periodos, setPeriodos] = useState<PeriodoCaptacao[]>([]);
  const [calendarioAcoes, setCalendarioAcoes] = useState<CalendarioAcao[]>([]);
  const [empresasParceiras, setEmpresasParceiras] = useState<EmpresaParceira[]>(
    [],
  );
  const [controleConcorrencia, setControleConcorrencia] = useState<
    ControleConcorrencia[]
  >([]);
  const [whatsappMessages, setWhatsappMessages] = useState<WhatsAppMessage[]>(
    [],
  );
  const [activeWhatsappTemplates, setActiveWhatsappTemplates] = useState<
    Record<string, string>
  >({});
  const [links, setLinks] = useState<LinkUtil[]>([]);
  const [mapao, setMapao] = useState<MapaoAcademicoEntry[]>([]);
  const [basesDisparo, setBasesDisparo] = useState<BaseDisparoEntry[]>([]);
  const [basesRenovacao, setBasesRenovacao] = useState<BaseEntry[]>([]);
  const [cursos, setCursos] = useState<CursoDisponivel[]>([]);
  const uniqueUnidades = useMemo(() => {
    return Array.from(
      new Set((cursos || []).map((c) => c.nomeUnidade).filter(Boolean)),
    ).sort();
  }, [cursos]);
  const [pedidosCursos, setPedidosCursos] = useState<PedidoCursoEntry[]>([]);
  const [ligacoes, setLigacoes] = useState<Ligacao[]>([]);
  const [insumosPedidos, setInsumosPedidos] = useState<InsumoPedido[]>([]);
  const [insumosEstoque, setInsumosEstoque] = useState<InsumoEstoque[]>([]);
  const [insumosPedidosComercial, setInsumosPedidosComercial] = useState<
    InsumoPedidoComercial[]
  >([]);
  const [insumosEstoqueComercial, setInsumosEstoqueComercial] = useState<
    InsumoEstoqueComercial[]
  >([]);
  const [botConfig, setBotConfig] = useState<BotConfig>({
    url: "",
    active: false,
  });
  const [botStatuses, setBotStatuses] = useState<
    Record<
      string,
      {
        status: string;
        pairingCode?: string;
        qrCode?: string;
        qrUrl?: string;
        active?: boolean;
      }
    >
  >({});
  const [showInjectModal, setShowInjectModal] = useState(false);
  const [injectBotNumber, setInjectBotNumber] = useState("");
  const [injectSessionData, setInjectSessionData] = useState("");
  const [initialActionData, setInitialActionData] =
    useState<Partial<CalendarioAcao> | null>(null);
  const [activePopup, setActivePopup] = useState<{
    title: string;
    message: string;
  } | null>(null);
  const [isProfileModalOpen, setIsProfileModalOpen] = useState(false);
  const [massSendProgress, setMassSendProgress] = useState<{
    total: number;
    sent: number;
    active: boolean;
    info: string;
  }>({ total: 0, sent: 0, active: false, info: "" });
  const [isMassSendPaused, setIsMassSendPaused] = useState(false);
  const massSendControlRef = React.useRef({ paused: false, cancelled: false });

  const showToast = (
    message: string,
    type: "success" | "error" = "success",
  ) => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  const showPopup = (title: string, message: string) => {
    setActivePopup({ title, message });
  };

  const canView = (view: string) => {
    if (!profile) return false;
    if (
      profile.email === "canaldonutri@gmail.com" ||
      profile.email === "marcos.teixeira@estacio.br" ||
      profile.role === "Admin Master"
    ) {
      return true;
    }
    const isComercial =
      localStorage.getItem("servidor_selected") === "comercial";
    if (profile.role === ROLES.FINANCEIRO) {
      if (isComercial) {
        return ["controlePagamentos", "controleInsumosComercial"].includes(view);
      } else {
        return VIEW_PERMISSIONS[view]?.includes(profile.role) || false;
      }
    }
    return VIEW_PERMISSIONS[view]?.includes(profile.role) || false;
  };

  const callBotApi = async (
    path: string,
    options: { method?: "GET" | "POST"; body?: any } = {},
  ) => {
    // Determine the exact URL to fetch from, using the requested Railway API directly for send actions
    const directUrl =
      path === "/api/send"
        ? "https://argoscliente-production-170b.up.railway.app/api/send"
        : botConfig.url
          ? `${botConfig.url.endsWith("/") ? botConfig.url.slice(0, -1) : botConfig.url}${path}`
          : `https://argoscliente-production-170b.up.railway.app${path}`;

    const fetchOptions: RequestInit = {
      method: options.method || "GET",
      headers: {
        "Content-Type": "application/json",
      },
    };
    if (options.method === "POST" && options.body) {
      fetchOptions.body = JSON.stringify(options.body);
    }

    const response = await fetch(directUrl, fetchOptions);
    if (!response.ok) {
      const isJson = response.headers
        .get("content-type")
        ?.includes("application/json");
      const json = isJson ? await response.json().catch(() => ({})) : {};
      throw new Error(
        json.error ||
          json.message ||
          `Erro ao conectar ao Bot (${response.status})`,
      );
    }

    const contentType = response.headers.get("content-type") || "";
    if (!contentType.includes("application/json")) {
      const text = await response.text();
      throw new Error(
        `O Bot no Railway retornou uma resposta inesperada (formato não-JSON). O bot pode estar offline ou em reinicialização.`,
      );
    }

    const resData = await response.json();

    // Support either direct raw JSON responses or wrapper structures with { success: boolean, data?: any }
    if (
      resData !== null &&
      typeof resData === "object" &&
      "success" in resData
    ) {
      if (!resData.success) {
        throw new Error(resData.data?.error || resData.error || `Falha no bot`);
      }
      return "data" in resData ? resData.data : resData;
    }

    return resData;
  };

  const sendAppWhatsApp = async (recipientPhone: string, message: string) => {
    let rawPhone = recipientPhone.replace(/\D/g, "");
    if (rawPhone.startsWith("0")) rawPhone = rawPhone.substring(1);
    if (rawPhone.length === 10 || rawPhone.length === 11) {
      rawPhone = `55${rawPhone}`;
    }
    if (!rawPhone) return;

    try {
      const finalMessage =
        message +
        "\n\nPor favor não responder nesse whatsapp. Pois ele é apenas um numero de assistência de envio.";
      await callBotApi("/api/send", {
        method: "POST",
        body: {
          botNumber: "5524993346717",
          number: rawPhone,
          message: finalMessage,
          force: true,
          manual: true,
        },
      });
      console.log(`WhatsApp sent to ${rawPhone} via bot 5524993346717`);
    } catch (err) {
      console.error("Error sending WhatsApp notification:", err);
    }
  };

  const sendAppTelegram = async (
    telegramHandleOrId: string,
    message: string,
  ) => {
    if (!telegramHandleOrId) return;
    const targetUrl = botConfig?.telegramBotUrl || "";
    const apiKey = botConfig?.telegramApiKey || "";
    if (!targetUrl) {
      console.log("Telegram Bot URL not configured in botConfig.");
      return;
    }
    try {
      const chatId = telegramHandleOrId.trim();
      const response = await fetch("/api/bot-proxy", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          targetUrl,
          method: "POST",
          headers: {
            "x-api-key": apiKey,
          },
          body: {
            chatId,
            mensagem: message,
          },
        }),
      });
      const resData = await response.json();
      if (!response.ok || !resData.success) {
        console.error(
          "Failed to send Telegram message at root:",
          resData.error || "Unknown error",
        );
      } else {
        console.log(`Telegram message sent to ${chatId}`);
      }
    } catch (err) {
      console.error("Error calling Telegram bot-proxy at root:", err);
    }
  };

  const handleSendBotMessage = async (
    telefone: string,
    message: string,
    contactName?: string,
  ) => {
    const currentBotNumber = profile?.botNumber;
    let safeBotNumber = currentBotNumber
      ? currentBotNumber.replace(/\D/g, "")
      : "";

    // Auto-fallback: if the user's personal bot number is offline, not active,
    // or not set, look for any online bot session in the system to route the dispatch.
    const isUserBotOnline =
      safeBotNumber && (botStatuses as any)[safeBotNumber]?.status === "online";

    if (!isUserBotOnline) {
      const firstOnlineBot = Object.entries(botStatuses).find(
        ([_, info]) => (info as any)?.status === "online",
      )?.[0];
      if (firstOnlineBot) {
        console.log(
          `Fallback bot activated: Routing message via active online session: ${firstOnlineBot}`,
        );
        safeBotNumber = firstOnlineBot;
      } else if (!safeBotNumber) {
        showToast(
          "Você ainda não tem um número de WhatsApp configurado (Administração -> GestãoPro) e nenhum bot está ativo no momento.",
          "error",
        );
        return;
      }
    }

    // Format phone: remove non-numeric, strip leading zero if present
    let rawPhone = telefone.replace(/\D/g, "");
    if (rawPhone.startsWith("0")) rawPhone = rawPhone.substring(1);
    // Add country code if not present and has standard length
    if (rawPhone.length === 10 || rawPhone.length === 11) {
      rawPhone = `55${rawPhone}`;
    }

    try {
      const isTargetBot = safeBotNumber === "5524993346717";
      const finalMessage = isTargetBot
        ? message +
          "\n\nPor favor não responder nesse whatsapp. Pois ele é apenas um numero de assistência de envio."
        : message;

      await callBotApi("/api/send", {
        method: "POST",
        body: {
          botNumber: safeBotNumber,
          number: rawPhone,
          message: finalMessage,
          contactName: contactName || "",
          force: true,
          manual: true,
        },
      });
      showToast("Mensagem enviada com sucesso pelo Bot ARGO'S!");

      // Log to CRM
      try {
        let sentiment = "Neutro";
        try {
          const res = await fetch("/api/crm/sentiment", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ text: message }),
          });
          const data = await res.json();
          if (data.success && data.sentiment) {
            sentiment = data.sentiment;
          }
        } catch (e) {
          console.error("Error analyzing sentiment:", e);
        }

        const msgData = {
          text: message,
          senderId: profile?.uid || "system",
          senderName: profile?.name || "System",
          senderRole: profile?.role || "",
          receiverPhone: rawPhone,
          timestamp: serverTimestamp(),
          type: "sent",
          status: "sent",
        };
        await addDoc(collection(db, COLLECTIONS.MESSAGES), msgData);

        // Update or Create conversation
        await setDoc(
          doc(db, COLLECTIONS.CONVERSATIONS, rawPhone),
          {
            contactPhone: rawPhone,
            contactName: contactName || rawPhone,
            lastMessage: message,
            lastMessageTimestamp: serverTimestamp(),
            unreadCount: 0,
            unidade: profile?.unidade || "",
            sentiment,
          },
          { merge: true },
        );
      } catch (err) {
        console.error("Error logging message to CRM:", err);
      }

      // Automatic Status Transition Logic upon message sent
      try {
        const phonesMatch = (p1?: string, p2?: string): boolean => {
          if (!p1 || !p2) return false;
          const c1 = p1.replace(/\D/g, "");
          const c2 = p2.replace(/\D/g, "");
          if (c1 === c2) return true;
          const s1 = c1.startsWith("55")
            ? c1.substring(2)
            : c1.startsWith("0")
              ? c1.substring(1)
              : c1;
          const s2 = c2.startsWith("55")
            ? c2.substring(2)
            : c2.startsWith("0")
              ? c2.substring(1)
              : c2;
          if (s1 === s2) return true;
          if (s1.length >= 8 && s2.length >= 8) {
            const last8_1 = s1.slice(-8);
            const last8_2 = s2.slice(-8);
            const ddd1 = s1.substring(0, 2);
            const ddd2 = s2.substring(0, 2);
            if (last8_1 === last8_2 && ddd1 === ddd2) return true;
          }
          return false;
        };

        const matchedLeads = leads.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );
        const matchedBases = bases.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );
        const matchedBasesRenovacao = basesRenovacao.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );
        const matchedFiesProuni = fiesProuni.filter((item) =>
          phonesMatch(item.telefone, telefone),
        );

        const existsInGap = gap.some((g) => {
          if (phonesMatch(g.telefone, telefone)) return true;
          const matchedCpf =
            matchedLeads.find((l) => l.cpf)?.cpf ||
            matchedBases.find((b) => b.cpf)?.cpf ||
            matchedBasesRenovacao.find((br) => br.cpf)?.cpf ||
            matchedFiesProuni.find((fp) => fp.cpf)?.cpf;
          if (matchedCpf && g.cpf) {
            const c1 = matchedCpf.replace(/\D/g, "");
            const c2 = g.cpf.replace(/\D/g, "");
            if (c1 && c1 === c2) return true;
          }
          return false;
        });

        // 1. Process matched LEADS
        for (const lead of matchedLeads) {
          if (existsInGap) {
            if (lead.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.LEADS, lead.id), {
                status: "Convertido",
              });
            }
          } else if (lead.status.toLowerCase() === "pendente") {
            await updateDoc(doc(db, COLLECTIONS.LEADS, lead.id), {
              status: "Sem retorno",
            });
          }
        }

        // 2. Process matched BASES
        for (const entry of matchedBases) {
          if (existsInGap) {
            if (entry.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.BASES, entry.id), {
                status: "Convertido",
              });
            }
          } else if (entry.status.toLowerCase() === "pendente") {
            await updateDoc(doc(db, COLLECTIONS.BASES, entry.id), {
              status: "Sem retorno",
            });
          }
        }

        // 3. Process matched BASES_RENOVACAO
        for (const entry of matchedBasesRenovacao) {
          if (existsInGap) {
            if (entry.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.BASES_RENOVACAO, entry.id), {
                status: "Convertido",
              });
            }
          } else if (entry.status.toLowerCase() === "pendente") {
            await updateDoc(doc(db, COLLECTIONS.BASES_RENOVACAO, entry.id), {
              status: "Sem retorno",
            });
          }
        }

        // 4. Process matched FIES_PROUNI
        for (const entry of matchedFiesProuni) {
          if (existsInGap) {
            if (entry.status !== "Convertido") {
              await updateDoc(doc(db, COLLECTIONS.FIES_PROUNI, entry.id), {
                status: "Convertido",
              });
            }
          } else if (
            entry.status &&
            entry.status.toLowerCase() === "pendente"
          ) {
            await updateDoc(doc(db, COLLECTIONS.FIES_PROUNI, entry.id), {
              status: "Sem retorno",
            });
          }
        }

        let tipoContato = "outro";
        let baseName = "";
        if (matchedLeads.length > 0) {
          tipoContato = "leads";
        } else if (matchedBases.length > 0) {
          tipoContato = "bases";
          baseName = matchedBases[0].nomeBase;
        } else if (matchedBasesRenovacao.length > 0) {
          tipoContato = "bases_renovacao";
          baseName = matchedBasesRenovacao[0].nomeBase;
        } else if (matchedFiesProuni.length > 0) {
          tipoContato = "fies_prouni";
        } else if (existsInGap) {
          tipoContato = "gap";
        }

        await addDoc(collection(db, COLLECTIONS.BOT_REPORTS), {
          userId: profile?.uid || "unknown",
          userName: profile?.nome || "Usuário Desconhecido",
          userRole: profile?.role || "unknown",
          telefone,
          tipoContato,
          baseName,
          sentAt: serverTimestamp(),
        });
      } catch (statusErr: any) {
        console.error(
          "[Auto Status Update] Failed to update statuses or log report:",
          statusErr,
        );
      }
    } catch (err: any) {
      showToast(`Erro ao enviar mensagem: ${err.message}`, "error");
    }
  };

  const sendSilentWhatsApp = async (telefone: string, message: string) => {
    const currentBotNumber = profile?.botNumber;
    let safeBotNumber = currentBotNumber
      ? currentBotNumber.replace(/\D/g, "")
      : "";

    const isUserBotOnline =
      safeBotNumber && (botStatuses as any)[safeBotNumber]?.status === "online";

    if (!isUserBotOnline) {
      const firstOnlineBot = Object.entries(botStatuses).find(
        ([_, info]) => (info as any)?.status === "online",
      )?.[0];
      if (firstOnlineBot) {
        safeBotNumber = firstOnlineBot;
      } else if (!safeBotNumber) {
        return;
      }
    }

    let rawPhone = telefone.replace(/\D/g, "");
    if (rawPhone.startsWith("0")) rawPhone = rawPhone.substring(1);
    if (rawPhone.length === 10 || rawPhone.length === 11) {
      rawPhone = `55${rawPhone}`;
    }

    try {
      const isTargetBot = safeBotNumber === "5524993346717";
      const finalMessage = isTargetBot
        ? message +
          "\n\nPor favor não responder nesse whatsapp. Pois ele é apenas um numero de assistência de envio."
        : message;

      await callBotApi("/api/send", {
        method: "POST",
        body: {
          botNumber: safeBotNumber,
          number: rawPhone,
          message: finalMessage,
          force: true,
          manual: true,
        },
      });
    } catch (err) {
      console.error(err);
    }
  };

  const handleMassSendBotMessages = async (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => {
    if (massSendProgress.active) {
      showToast("Já existe um envio em massa em andamento.", "error");
      return;
    }

    if (messages.length === 0) return;
    if (
      !window.confirm(
        `Deseja iniciar o envio em massa via bot para ${messages.length} contatos?`,
      )
    )
      return;

    massSendControlRef.current = { paused: false, cancelled: false };
    setIsMassSendPaused(false);

    setMassSendProgress({
      total: messages.length,
      sent: 0,
      active: true,
      info: "Iniciando...",
    });

    const waitWithCheck = async (seconds: number, labelPrefix: string) => {
      for (let s = 0; s < seconds; s++) {
        if (massSendControlRef.current.cancelled) return;
        while (
          massSendControlRef.current.paused &&
          !massSendControlRef.current.cancelled
        ) {
          setMassSendProgress((prev) => ({
            ...prev,
            info: `Robô Pausado... (${prev.sent}/${messages.length})`,
          }));
          await new Promise((resolve) => setTimeout(resolve, 500));
        }
        if (massSendControlRef.current.cancelled) return;
        const remaining = seconds - s;
        setMassSendProgress((prev) => ({
          ...prev,
          info: `${labelPrefix} (${remaining}s restantes)... (${prev.sent}/${messages.length})`,
        }));
        await new Promise((resolve) => setTimeout(resolve, 1000));
      }
    };

    let sentCount = 0;
    for (let i = 0; i < messages.length; i++) {
      if (massSendControlRef.current.cancelled) {
        break;
      }

      while (
        massSendControlRef.current.paused &&
        !massSendControlRef.current.cancelled
      ) {
        setMassSendProgress((prev) => ({
          ...prev,
          info: `Robô Pausado... (${sentCount}/${messages.length})`,
        }));
        await new Promise((resolve) => setTimeout(resolve, 500));
      }

      if (massSendControlRef.current.cancelled) {
        break;
      }

      if (i > 0) {
        if (sentCount % 5 === 0) {
          await waitWithCheck(120, "Pausa de 2 min");
        } else {
          await waitWithCheck(30, "Aguardando cooldown");
        }
      }

      if (massSendControlRef.current.cancelled) {
        break;
      }

      while (
        massSendControlRef.current.paused &&
        !massSendControlRef.current.cancelled
      ) {
        setMassSendProgress((prev) => ({
          ...prev,
          info: `Robô Pausado... (${sentCount}/${messages.length})`,
        }));
        await new Promise((resolve) => setTimeout(resolve, 500));
      }

      if (massSendControlRef.current.cancelled) {
        break;
      }

      setMassSendProgress((prev) => ({
        ...prev,
        sent: sentCount,
        info: `Enviando... (${sentCount + 1}/${messages.length})`,
      }));

      try {
        await handleSendBotMessage(
          messages[i].telefone,
          messages[i].message,
          messages[i].nome,
        );
      } catch (e) {
        console.error("Error sending bot message in mass: ", e);
      }
      sentCount++;
      setMassSendProgress((prev) => ({
        ...prev,
        sent: sentCount,
      }));
    }

    const wasCancelled = massSendControlRef.current.cancelled;
    setMassSendProgress({ total: 0, sent: 0, active: false, info: "" });
    setIsMassSendPaused(false);

    if (wasCancelled) {
      showToast("Envio em massa cancelado pelo usuário.", "error");
    } else {
      showToast("Envio em massa concluído!", "success");
    }
  };

  // Subscribe to Analysis Schemes
  useEffect(() => {
    const q = collection(db, COLLECTIONS.CRESCIMENTO_ANUAL);
    const unsubscribe = onSnapshot(q, (snapshot) => {
      const list = snapshot.docs.map((doc) => ({
        id: doc.id,
        ...doc.data(),
      })) as AnalysisScheme[];
      setAnalysisSchemes(list);
    }, (err) => {
      console.log("Crescimento Anual snapshot error:", err);
    });
    return () => unsubscribe();
  }, []);

  const handleSaveAnalysisScheme = async (scheme: Partial<AnalysisScheme>) => {
    try {
      if (scheme.id) {
        const { id, ...data } = scheme;
        await updateDoc(doc(db, COLLECTIONS.CRESCIMENTO_ANUAL, id), {
          ...data,
          updatedAt: serverTimestamp(),
        });
        showToast("Análise salva com sucesso!");
      } else {
        await addDoc(collection(db, COLLECTIONS.CRESCIMENTO_ANUAL), {
          ...scheme,
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
        });
        showToast("Análise criada com sucesso!");
      }
    } catch (err: any) {
      console.error("Error saving analysis scheme:", err);
      showToast("Erro ao salvar análise.", "error");
    }
  };

  const handleDeleteAnalysisScheme = async (id: string) => {
    if (window.confirm("Deseja excluir esta análise?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.CRESCIMENTO_ANUAL, id));
        showToast("Análise excluída.");
      } catch (err: any) {
        console.error("Error deleting analysis scheme:", err);
        showToast("Erro ao excluir análise.", "error");
      }
    }
  };

  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      if (user) {
        try {
          // 1. Try to get profile by UID
          let userDoc = await getDoc(doc(db, COLLECTIONS.USERS, user.uid));

          if (!userDoc.exists()) {
            // 2. If not found by UID, try to find by email (for pre-registered users)
            const q = query(
              collection(db, COLLECTIONS.USERS),
              where("email", "==", user.email),
            );
            const querySnap = await getDocs(q);

            if (!querySnap.empty) {
              // Found by email, use this document
              const existingDoc = querySnap.docs[0];
              const data = existingDoc.data();

              // If the document ID is not the UID, we should ideally migrate it
              // but for now we'll just use it. Wait, if we use it, rules might fail
              // because rules expect path/.../users/{uid}.
              // So we MUST migrate it to a document with UID as ID.
              await setDoc(doc(db, COLLECTIONS.USERS, user.uid), {
                ...data,
                uid: user.uid,
                updatedAt: serverTimestamp(),
              });

              // Delete the old document if it had a different ID
              if (existingDoc.id !== user.uid) {
                try {
                  await deleteDoc(doc(db, COLLECTIONS.USERS, existingDoc.id));
                } catch (e) {
                  console.warn(
                    "Could not delete old user document, likely due to rules. Skipping.",
                    e,
                  );
                }
              }

              userDoc = await getDoc(doc(db, COLLECTIONS.USERS, user.uid));
            } else {
              // 3. Create default profile if not exists at all
              let role = ROLES.PROMOTOR;
              let servidor: "principal" | "comercial" = "principal";
              let name = user.email!.split("@")[0];

              if (user.displayName) {
                const parts = user.displayName.split("|");
                name = parts[0] || name;
                if (parts.length > 1 && parts[1] === "comercial") {
                  servidor = "comercial";
                  role = "Promotor/rua" as any;
                }
              }

              if (
                user.email === "marcos.teixeira@estacio.br" ||
                user.email === "canaldonutri@gmail.com"
              ) {
                role = ROLES.ADMIN_MASTER;
              } else {
                const allUsers = await getDocs(
                  query(collection(db, COLLECTIONS.USERS), limit(1)),
                );
                if (allUsers.empty) {
                  role = (
                    servidor === "comercial"
                      ? "Gerente Comercial (Comercial)"
                      : ROLES.LIDER_FDV
                  ) as any;
                }
              }

              const newProfile = {
                uid: user.uid,
                email: user.email!,
                name,
                role,
                servidor,
                mustChangePassword: false, // Default for self-signup
                createdAt: serverTimestamp(),
                dashboardWidgets: { stats: true, links: true, planner: true },
              };
              await setDoc(doc(db, COLLECTIONS.USERS, user.uid), newProfile);
              userDoc = await getDoc(doc(db, COLLECTIONS.USERS, user.uid));
            }
          }

          if (userDoc.exists()) {
            const data = userDoc.data() as UserProfile;
            if (
              data.email === "marcos.teixeira@estacio.br" ||
              data.email === "canaldonutri@gmail.com"
            ) {
              data.role = ROLES.ADMIN_MASTER;
            }
            setProfile({ uid: user.uid, ...data } as UserProfile);
          }
          setUser(user);
        } catch (error: any) {
          console.error("Error fetching/creating profile details:", {
            code: error.code,
            message: error.message,
            stack: error.stack,
          });
          showToast(`Erro ao carregar perfil: ${error.message}`, "error");
          setUser(null);
          setProfile(null);
        }
      } else {
        setUser(null);
        setProfile(null);
      }
      setLoading(false);
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    if (!user) return;

    // Listeners for users require auth
    let unsubUsers = () => {};
    if (user && profile) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let usersQuery = query(collection(db, COLLECTIONS.USERS));
      if (isRestricted) {
        usersQuery = query(
          usersQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubUsers = onSnapshot(
        usersQuery,
        (snap) => {
          setUsers(
            snap.docs.map((d) => ({ uid: d.id, ...d.data() }) as UserProfile),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.USERS),
      );
    }

    let unsubPlanner = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubPlanner = onSnapshot(
        collection(db, COLLECTIONS.PLANNER),
        (snap) => {
          setPlanner(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as PlannerTask),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.PLANNER),
      );
    }

    let unsubLinks = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubLinks = onSnapshot(
        collection(db, COLLECTIONS.LINKS),
        (snap) => {
          setLinks(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as LinkUtil),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.LINKS),
      );
    }

    let unsubSalesContacts = () => {};
    let unsubLeads = () => {};
    if (profile) {
      let leadsQuery;
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      if (
        [
          ROLES.ADMIN_MASTER,
          ROLES.LIDER_FDV,
          ROLES.SALA_MATRICULA,
          ROLES.QG,
        ].includes(profile.role)
      ) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.GESTOR_UNIDADE) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("unidade", "==", profile.unidade || ""),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.GESTOR_COMERCIAL_COMERCIAL) {
        // Gerente Comercial (Comercial) ver everything in Comercial
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("servidor", "==", "comercial"),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.FDV_COMERCIAL) {
        // FDV (Comercial) sees their own leads or those from their linked promontors.
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          or(
            where("promotorId", "==", user!.uid),
            where("linkadoA", "==", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.FDV) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          or(
            where("promotorId", "==", user!.uid),
            where("linkadoA", "==", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (profile.role === ROLES.GESTOR_COMERCIAL) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          or(
            where("promotorId", "==", user!.uid),
            where("promotorRole", "in", [ROLES.PROMOTOR, ROLES.FDV]),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (
        profile.role === ROLES.PROMOTOR ||
        profile.role === ROLES.PROMOTOR_RUA
      ) {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("promotorId", "==", user!.uid),
          orderBy("createdAt", "desc"),
        );
      } else {
        leadsQuery = query(
          collection(db, COLLECTIONS.LEADS),
          where("promotorId", "==", "none"),
          orderBy("createdAt", "desc"),
        );
      }

      if (isRestricted) {
        if (
          profile.role === ROLES.FDV ||
          profile.role === ROLES.FDV_COMERCIAL
        ) {
          // Already restricted by promotorId/linkadoA above, but we keep unit filter to be safe or bypass it.
          // The user said ONLY what they or linked promotor filled.
          // If we add the unit filter, it might exclude their own leads if they are in a different unit (unlikely).
          // But to be strict with "SÓ PODERÁ VE", we keep the current query which is already restricted to UID.
        } else {
          leadsQuery = query(
            leadsQuery,
            where("unidade", "==", profile.unidade || "Matriz"),
          );
        }
      }

      unsubLeads = onSnapshot(
        leadsQuery,
        (snap) => {
          setLeads(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as Lead));
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.LEADS),
      );
      
      unsubSalesContacts = onSnapshot(
        query(collection(db, COLLECTIONS.SALES_CONTACTS), orderBy("createdAt", "desc")),
        (snap) => {
          setSalesContacts(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as SalesContact));
        },
        (err) => console.error("Error loading sales contacts:", err)
      );
    }

    let unsubBases = () => {};
    if (profile && VIEW_PERMISSIONS.bases.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let basesQuery;
      if (isRestricted) {
        if (
          profile.role === ROLES.FDV ||
          profile.role === ROLES.FDV_COMERCIAL
        ) {
          basesQuery = query(
            collection(db, COLLECTIONS.BASES),
            or(
              where("promotorId", "==", user!.uid),
              where("linkadoA", "==", user!.uid),
            ),
            orderBy("createdAt", "desc"),
          );
        } else {
          basesQuery = query(
            collection(db, COLLECTIONS.BASES),
            where("unidade", "==", profile.unidade || "Matriz"),
            orderBy("createdAt", "desc"),
          );
        }
      } else {
        basesQuery = query(
          collection(db, COLLECTIONS.BASES),
          orderBy("createdAt", "desc"),
        );
      }

      unsubBases = onSnapshot(
        basesQuery,
        (snap) => {
          setBases(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as BaseEntry),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.BASES),
      );
    }

    let unsubGap = () => {};
    if (profile && VIEW_PERMISSIONS.gap.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let gapQuery = query(
        collection(db, COLLECTIONS.GAP),
        orderBy("createdAt", "desc"),
      );
      if (isRestricted) {
        gapQuery = query(
          gapQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubGap = onSnapshot(
        gapQuery,
        (snap) => {
          setGap(snap.docs.map((d) => ({ id: d.id, ...d.data() }) as GapEntry));
        },
        (err) => handleFirestoreError(err, OperationType.LIST, COLLECTIONS.GAP),
      );
    }

    let unsubIsencoes = () => {};
    let unsubSolicitacoesManutencao = () => {};
    if (profile && VIEW_PERMISSIONS.isencoes.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let isencoesQuery = query(
        collection(db, COLLECTIONS.ISENCOES),
        orderBy("createdAt", "desc"),
      );
      if (isRestricted) {
        isencoesQuery = query(
          isencoesQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubIsencoes = onSnapshot(
        isencoesQuery,
        (snap) => {
          setIsencoes(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as IsencaoEntry),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.ISENCOES),
      );

      unsubSolicitacoesManutencao = onSnapshot(
        collection(db, COLLECTIONS.SOLICITACOES_MANUTENCAO),
        (snap) => {
          setSolicitacoesManutencao(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as SolicitacaoManutencao),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.SOLICITACOES_MANUTENCAO),
      );
    }

    let unsubPedidosCursos = () => {};
    if (
      profile &&
      (VIEW_PERMISSIONS.historico.includes(profile.role) ||
        VIEW_PERMISSIONS.relatorios.includes(profile.role))
    ) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let pcQuery = query(
        collection(db, COLLECTIONS.PEDIDO_CURSOS),
        orderBy("createdAt", "desc"),
      );
      if (isRestricted) {
        pcQuery = query(
          pcQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubPedidosCursos = onSnapshot(
        pcQuery,
        (snap) => {
          setPedidosCursos(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as PedidoCursoEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.PEDIDO_CURSOS,
          ),
      );
    }

    let unsubFiesProuni = () => {};
    let unsubFiesProuniVagas = () => {};
    if (profile && VIEW_PERMISSIONS.fiesProuni.includes(profile.role)) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        profile.role !== ROLES.SSA &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let fpQuery = query(
        collection(db, COLLECTIONS.FIES_PROUNI),
        orderBy("createdAt", "desc"),
      );
      let fpvQuery = query(
        collection(db, COLLECTIONS.FIES_PROUNI_VAGAS),
        orderBy("createdAt", "desc"),
      );

      if (isRestricted) {
        fpQuery = query(
          fpQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
        fpvQuery = query(
          fpvQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubFiesProuni = onSnapshot(
        fpQuery,
        (snap) => {
          setFiesProuni(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as FiesProuniEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.FIES_PROUNI,
          ),
      );
      unsubFiesProuniVagas = onSnapshot(
        fpvQuery,
        (snap) => {
          setFiesProuniVagas(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as FiesProuniVaga),
          );
        },
        (err) => console.error("Error fetching FIES_PROUNI_VAGAS:", err),
      );
    }

    let unsubCampanhas = () => {};
    if (profile && VIEW_PERMISSIONS.campanhas.includes(profile.role)) {
      unsubCampanhas = onSnapshot(
        collection(db, COLLECTIONS.CAMPANHAS),
        (snap) => {
          setCampanhas(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as Campanha),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.CAMPANHAS),
      );
    }

    let unsubBomDia = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubBomDia = onSnapshot(
        collection(db, COLLECTIONS.BOM_DIA),
        (snap) => {
          setBomDia(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as BomDiaCaptacao),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.BOM_DIA),
      );
    }

    let unsubForecast = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubForecast = onSnapshot(
        collection(db, COLLECTIONS.FORECAST),
        (snap) => {
          setForecast(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as ForecastCaptacao,
            ),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.FORECAST),
      );
    }

    let unsubMetaDia = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubMetaDia = onSnapshot(
        collection(db, COLLECTIONS.META_DIA),
        (snap) => {
          setMetaDia(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MetaDia),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.META_DIA),
      );
    }

    let unsubMetaSM = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubMetaSM = onSnapshot(
        collection(db, COLLECTIONS.META_SM),
        (snap) => {
          setMetaSM(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MetaSM),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.META_SM),
      );
    }

    let unsubMetaCursos = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubMetaCursos = onSnapshot(
        collection(db, COLLECTIONS.META_CURSOS),
        (snap) => {
          setMetaCursos(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as MetaCurso),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.META_CURSOS),
      );
    }

    let unsubQgLigacoes = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubQgLigacoes = onSnapshot(
        collection(db, COLLECTIONS.QG_LIGACOES),
        (snap) => {
          setQgLigacoes(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as QgLigacao),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.QG_LIGACOES,
          ),
      );
    }

    let unsubPeriodos = () => {};
    if (profile && VIEW_PERMISSIONS.dashboard.includes(profile.role)) {
      unsubPeriodos = onSnapshot(
        collection(db, COLLECTIONS.PERIODO_CAPTACAO),
        (snap) => {
          setPeriodos(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as PeriodoCaptacao,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.PERIODO_CAPTACAO,
          ),
      );
    }

    let unsubCalendario = () => {};
    if (
      profile &&
      (VIEW_PERMISSIONS.calendario.includes(profile.role) ||
        VIEW_PERMISSIONS.controlePagamentos.includes(profile.role) ||
        canView("controlePagamentos"))
    ) {
      const isRestricted =
        profile.role !== ROLES.ADMIN_MASTER && profile.role !== ROLES.FINANCEIRO &&
        profile.role !== ROLES.GESTOR_COMERCIAL &&
        profile.role !== ROLES.GESTOR_COMERCIAL_COMERCIAL &&
        profile.role !== ROLES.LIDER_FDV &&
        !["canaldonutri@gmail.com", "marcos.teixeira@estacio.br"].includes(
          user?.email || "",
        );

      let calendarioQuery;
      if (
        [
          ROLES.ADMIN_MASTER,
          ROLES.LIDER_FDV,
          ROLES.SALA_MATRICULA,
          ROLES.GESTOR_UNIDADE,
          ROLES.GESTOR_COMERCIAL,
          ROLES.FINANCEIRO,
          ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
        ].includes(profile.role)
      ) {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          orderBy("createdAt", "desc"),
        );
      } else if (
        profile.role === ROLES.FDV ||
        profile.role === ROLES.FDV_COMERCIAL
      ) {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          or(
            where("creatorId", "==", user!.uid),
            where("creatorRole", "==", ROLES.PROMOTOR),
            where("creatorRole", "==", ROLES.PROMOTOR_RUA),
            where("colaboradorId", "==", user!.uid),
            where("colaboradoresIds", "array-contains", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else if (
        profile.role === ROLES.PROMOTOR ||
        profile.role === ROLES.PROMOTOR_RUA
      ) {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          or(
            where("creatorId", "==", user!.uid),
            where("colaboradorId", "==", user!.uid),
            where("promotoresSelecionados", "array-contains", user!.uid),
          ),
          orderBy("createdAt", "desc"),
        );
      } else {
        calendarioQuery = query(
          collection(db, COLLECTIONS.CALENDARIO_ACOES),
          where("creatorId", "==", "none"),
          orderBy("createdAt", "desc"),
        );
      }

      if (isRestricted) {
        calendarioQuery = query(
          calendarioQuery,
          where("unidade", "==", profile.unidade || "Matriz"),
        );
      }

      unsubCalendario = onSnapshot(
        calendarioQuery,
        (snap) => {
          setCalendarioAcoes(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as CalendarioAcao),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.CALENDARIO_ACOES,
          ),
      );
    }

    let unsubEmpresas = () => {};
    if (profile && VIEW_PERMISSIONS.empresas.includes(profile.role)) {
      let empresasQuery = query(collection(db, COLLECTIONS.EMPRESAS_PARCEIRAS));

      const isRestricted = ![
        ROLES.ADMIN_MASTER,
        ROLES.GESTOR_COMERCIAL,
        ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
        ROLES.FINANCEIRO,
      ].includes(profile.role);

      if (isRestricted) {
        if (profile.role === ROLES.GESTOR_UNIDADE) {
          empresasQuery = query(
            empresasQuery,
            where(
              "unidadesVinculadas",
              "array-contains",
              profile.unidade || "",
            ),
          );
        } else if (
          profile.role === ROLES.FDV ||
          profile.role === ROLES.FDV_COMERCIAL
        ) {
          empresasQuery = query(
            empresasQuery,
            or(
              where(
                "unidadesVinculadas",
                "array-contains",
                profile.unidade || "",
              ),
              where("consultorId", "==", user!.uid),
              where("creatorId", "==", user!.uid),
            ),
          );
        }
      }

      unsubEmpresas = onSnapshot(
        empresasQuery,
        (snap) => {
          setEmpresasParceiras(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as EmpresaParceira,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.EMPRESAS_PARCEIRAS,
          ),
      );
    }

    let unsubControleConcorrencia = () => {};
    if (
      profile &&
      VIEW_PERMISSIONS.controleConcorrencia.includes(profile.role)
    ) {
      unsubControleConcorrencia = onSnapshot(
        collection(db, COLLECTIONS.CONTROLE_CONCORRENCIA),
        (snap) => {
          setControleConcorrencia(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as ControleConcorrencia,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.CONTROLE_CONCORRENCIA,
          ),
      );
    }

    let unsubWhatsApp = () => {};
    if (user) {
      unsubWhatsApp = onSnapshot(
        collection(db, COLLECTIONS.WHATSAPP_MESSAGES),
        (snap) => {
          setWhatsappMessages(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as WhatsAppMessage,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.WHATSAPP_MESSAGES,
          ),
      );
    }

    let unsubMapao = () => {};
    if (profile && VIEW_PERMISSIONS.mapao.includes(profile.role)) {
      unsubMapao = onSnapshot(
        collection(db, COLLECTIONS.MAPAO_ACADEMICO),
        (snap) => {
          setMapao(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as MapaoAcademicoEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.MAPAO_ACADEMICO,
          ),
      );
    }

    let unsubBasesDisparo = () => {};
    if (profile && VIEW_PERMISSIONS.basesDisparo.includes(profile.role)) {
      unsubBasesDisparo = onSnapshot(
        collection(db, COLLECTIONS.BASES_DISPARO),
        (snap) => {
          setBasesDisparo(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as BaseDisparoEntry,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.BASES_DISPARO,
          ),
      );
    }

    let unsubBasesRenovacao = () => {};
    if (profile && VIEW_PERMISSIONS.basesRenovacao.includes(profile.role)) {
      unsubBasesRenovacao = onSnapshot(
        collection(db, COLLECTIONS.BASES_RENOVACAO),
        (snap) => {
          setBasesRenovacao(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as BaseEntry),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.BASES_RENOVACAO,
          ),
      );
    }

    let unsubCursos = () => {};
    if (profile && VIEW_PERMISSIONS.cursos.includes(profile.role)) {
      unsubCursos = onSnapshot(
        collection(db, COLLECTIONS.CURSOS),
        (snap) => {
          setCursos(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as CursoDisponivel,
            ),
          );
        },
        (err) =>
          handleFirestoreError(err, OperationType.LIST, COLLECTIONS.CURSOS),
      );
    }

    let unsubInsumosPedidos = () => {};
    let unsubInsumosEstoque = () => {};
    if (profile && VIEW_PERMISSIONS.controleInsumos.includes(profile.role)) {
      unsubInsumosPedidos = onSnapshot(
        collection(db, COLLECTIONS.INSUMOS_PEDIDOS),
        (snap) => {
          setInsumosPedidos(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as InsumoPedido),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_PEDIDOS,
          ),
      );

      unsubInsumosEstoque = onSnapshot(
        collection(db, COLLECTIONS.INSUMOS_ESTOQUE),
        (snap) => {
          setInsumosEstoque(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as InsumoEstoque),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_ESTOQUE,
          ),
      );
    }

    let unsubInsumosPedidosComercial = () => {};
    let unsubInsumosEstoqueComercial = () => {};
    let unsubInsumosBaixas = () => {};
    if (
      profile &&
      VIEW_PERMISSIONS.controleInsumosComercial.includes(profile.role)
    ) {
      const isGerenteOrAdmin =
        profile.role === ROLES.ADMIN_MASTER ||
        profile.role === "Admin Master" ||
        profile.role === "Gerente Comercial (Comercial)" ||
        (profile.role === "Gestor Comercial" || profile.role === "Regional");

      const qPedidosComercial = isGerenteOrAdmin
        ? collection(db, COLLECTIONS.INSUMOS_PEDIDOS_COMERCIAL)
        : query(
            collection(db, COLLECTIONS.INSUMOS_PEDIDOS_COMERCIAL),
            where("solicitanteId", "==", profile.uid),
          );

      const qEstoqueComercial = isGerenteOrAdmin
        ? collection(db, COLLECTIONS.INSUMOS_ESTOQUE_COMERCIAL)
        : query(
            collection(db, COLLECTIONS.INSUMOS_ESTOQUE_COMERCIAL),
            where("ownerId", "==", profile.uid),
          );

      unsubInsumosPedidosComercial = onSnapshot(
        qPedidosComercial,
        (snap) => {
          setInsumosPedidosComercial(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as InsumoPedidoComercial,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_PEDIDOS_COMERCIAL,
          ),
      );

      unsubInsumosEstoqueComercial = onSnapshot(
        qEstoqueComercial,
        (snap) => {
          setInsumosEstoqueComercial(
            snap.docs.map(
              (d) => ({ id: d.id, ...d.data() }) as InsumoEstoqueComercial,
            ),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_ESTOQUE_COMERCIAL,
          ),
      );

      unsubInsumosBaixas = onSnapshot(
        collection(db, COLLECTIONS.INSUMOS_BAIXAS),
        (snap) => {
          setInsumosBaixas(
            snap.docs.map((d) => ({ id: d.id, ...d.data() }) as InsumoBaixa),
          );
        },
        (err) =>
          handleFirestoreError(
            err,
            OperationType.LIST,
            COLLECTIONS.INSUMOS_BAIXAS,
          ),
      );
    }

    const unsubLigacoes = onSnapshot(
      collection(db, COLLECTIONS.CONTROLE_LIGACOES),
      (snap) => {
        setLigacoes(
          snap.docs.map((d) => ({ id: d.id, ...d.data() }) as Ligacao),
        );
      },
      (err) =>
        handleFirestoreError(
          err,
          OperationType.LIST,
          COLLECTIONS.CONTROLE_LIGACOES,
        ),
    );

    return () => {
      unsubUsers();
      unsubPlanner();
      unsubLinks();
      unsubLeads();
      unsubBases();
      unsubGap();
      unsubIsencoes();
      unsubSolicitacoesManutencao();
      unsubPedidosCursos();
      unsubFiesProuni();
      unsubFiesProuniVagas();
      unsubCampanhas();
      unsubBomDia();
      unsubForecast();
      unsubMetaDia();
      unsubMetaSM();
      unsubMetaCursos();
      unsubQgLigacoes();
      unsubPeriodos();
      unsubCalendario();
      unsubEmpresas();
      unsubWhatsApp();
      unsubMapao();
      unsubBasesDisparo();
      unsubBasesRenovacao();
      unsubCursos();
      unsubControleConcorrencia();
      unsubInsumosPedidos();
      unsubInsumosEstoque();
      unsubInsumosPedidosComercial();
      unsubInsumosEstoqueComercial();
      unsubInsumosBaixas();
      unsubLigacoes();
    };
  }, [user, profile]);

  useEffect(() => {
    const unsubBotConfig = onSnapshot(
      doc(db, COLLECTIONS.BOT_CONFIG, "main"),
      (snap) => {
        if (snap.exists()) {
          setBotConfig({ id: snap.id, ...snap.data() } as BotConfig);
        } else {
          setBotConfig({ url: "", active: false });
        }
      },
      (err) => {
        console.warn("Could not load botConfig publicly:", err);
      },
    );
    return () => unsubBotConfig();
  }, []);

  useEffect(() => {
    // Test connection to Firestore as per instructions
    const testConnection = async () => {
      try {
        const { getDocFromServer, doc } = await import("firebase/firestore");
        await getDocFromServer(
          doc(db, COLLECTIONS.BOT_CONFIG, "connection_test"),
        );
        console.log("Firestore connection test: OK");
      } catch (err) {
        console.warn(
          "Firestore connection test check (expected error if doc doesn't exist):",
          err,
        );
      }
    };
    testConnection();
  }, []);

  useEffect(() => {
    let intervalId: NodeJS.Timeout;

    const checkBotStatus = async () => {
      try {
        const data = await callBotApi("/api/status");
        if (data && data.bots) {
          setBotStatuses(data.bots);
        }
      } catch (e: any) {
        console.debug("Bot check fail via proxy:", e.message);
      }
    };

    checkBotStatus();
    intervalId = setInterval(checkBotStatus, 3000);
    return () => clearInterval(intervalId);
  }, [botConfig.url]);

  const knownLeadsRef = React.useRef<Set<string> | null>(null);
  const knownCampanhasRef = React.useRef<Set<string> | null>(null);

  useEffect(() => {
    if (!profile) return;
    if (
      profile.role !== ROLES.LIDER_FDV &&
      profile.role !== ROLES.SALA_MATRICULA
    )
      return;

    if (knownLeadsRef.current === null) {
      knownLeadsRef.current = new Set(leads.map((l) => l.id!));
      return;
    }

    let hasNew = false;
    leads.forEach((l) => {
      if (!knownLeadsRef.current!.has(l.id!)) {
        knownLeadsRef.current!.add(l.id!);
        hasNew = true;
      }
    });

    if (hasNew) {
      showPopup("Novo Lead!", "Um novo lead foi adicionado no Histórico.");
    }
  }, [leads, profile]);

  useEffect(() => {
    if (!profile) return;
    if (
      profile.role !== ROLES.LIDER_FDV &&
      profile.role !== ROLES.SALA_MATRICULA
    )
      return;

    if (knownCampanhasRef.current === null) {
      knownCampanhasRef.current = new Set(campanhas.map((c) => c.id!));
      return;
    }

    let hasNew = false;
    campanhas.forEach((c) => {
      if (!knownCampanhasRef.current!.has(c.id!)) {
        knownCampanhasRef.current!.add(c.id!);
        hasNew = true;
      }
    });

    if (hasNew) {
      showPopup("Nova Campanha!", "Uma nova campanha foi adicionada.");
    }
  }, [campanhas, profile]);

  useEffect(() => {
    if (profile && !canView(currentView)) {
      const availableViews = [
        "dashboard",
        "cadastro",
        "historico",
        "bases",
        "gap",
        "fiesProuni",
        "mapao",
        "cursos",
        "basesDisparo",
        "campanhas",
        "calendario",
        "empresas",
        "calculo",
        "emailMarketing",
        "admin",
        "controlePagamentos",
      ];
      const firstAvailable = availableViews.find((v) => canView(v));
      if (firstAvailable) {
        setCurrentView(firstAvailable);
      }
    }
  }, [profile, currentView]);

  const handleSaveLigacao = async (ligacao: Partial<Ligacao>) => {
    try {
      await addDoc(collection(db, COLLECTIONS.CONTROLE_LIGACOES), {
        ...ligacao,
        createdAt: serverTimestamp(),
      });
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, COLLECTIONS.CONTROLE_LIGACOES);
      throw err;
    }
  };

  if (loading) {
    return (
      <div className="min-h-screen bg-[#01112c] flex items-center justify-center">
        <motion.div
          animate={{ rotate: 360 }}
          transition={{ repeat: Infinity, duration: 1, ease: "linear" }}
          className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full"
        />
      </div>
    );
  }

  const searchParams = new URLSearchParams(window.location.search);
  const publicFormId = searchParams.get("formId");

  if (publicFormId) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicCustomForm onToast={showToast} />
      </div>
    );
  }

  if (currentView === "pedido-insumos") {
    return (
      <div className="min-h-screen bg-[#01112c] flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicInsumoForm onToast={showToast} />
      </div>
    );
  }

  if (currentView === "manutencao-publica") {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicMaintenanceForm />
      </div>
    );
  }

  if (currentView === "desconto") {
    return (
      <div className="min-h-screen bg-[#01112c] flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicRegistrationForm onToast={showToast} />
      </div>
    );
  }

  if (currentView === "pedido-curso") {
    return (
      <div className="min-h-screen bg-[#01112c] flex flex-col justify-between">
        <AnimatePresence>
          {toast && (
            <Toast
              message={toast.message}
              type={toast.type}
              onClose={() => setToast(null)}
            />
          )}
        </AnimatePresence>
        <PublicPedidoCursoForm onToast={showToast} />
      </div>
    );
  }

  if (!user) {
    return <AuthScreen onToast={showToast} botConfig={botConfig} />;
  }

  if (profile?.blocked) {
    return (
      <div className="min-h-screen bg-[#01112c] flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-3xl shadow-xl border border-rose-100 text-center max-w-md">
          <XCircle size={64} className="text-rose-500 mx-auto mb-4" />
          <h2 className="text-2xl font-bold text-slate-900">
            Acesso Bloqueado
          </h2>
          <p className="text-slate-500 mt-2">
            Sua conta foi suspensa. Entre em contato com o administrador para
            mais informações.
          </p>
          <button
            onClick={() => signOut(auth)}
            className="mt-6 w-full bg-slate-900 text-white py-3 rounded-xl font-bold hover:bg-slate-800 transition-all"
          >
            Sair
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-[#01112c] flex">
      <AnimatePresence>
        {toast && (
          <Toast
            message={toast.message}
            type={toast.type}
            onClose={() => setToast(null)}
          />
        )}
      </AnimatePresence>

      <AnimatePresence>
        {activePopup && (
          <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[200] flex items-center justify-center p-4">
            <motion.div
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="bg-white p-6 md:p-8 rounded-3xl shadow-xl border border-slate-100 max-w-sm w-full text-center relative"
            >
              <button
                onClick={() => setActivePopup(null)}
                className="absolute top-4 right-4 text-slate-400 hover:text-slate-600 transition-colors"
              >
                <X size={20} />
              </button>
              <div className="mx-auto w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-4 border-4 border-blue-50">
                <Bell size={32} className="text-blue-600" />
              </div>
              <h3 className="text-xl font-bold text-slate-900 mb-2">
                {activePopup.title}
              </h3>
              <p className="text-sm text-slate-600 mb-6">
                {activePopup.message}
              </p>
              <button
                onClick={() => setActivePopup(null)}
                className="w-full bg-blue-600 text-white font-bold py-3 rounded-xl hover:bg-blue-700 transition shadow-lg shadow-blue-200"
              >
                Ciente
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      <AnimatePresence>
        {massSendProgress.active && (
          <motion.div
            initial={{ y: 100, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: 100, opacity: 0 }}
            className="fixed bottom-6 left-1/2 -translate-x-1/2 bg-white px-6 py-4 rounded-2xl shadow-2xl border border-slate-200 z-[300] flex flex-col items-center gap-2 max-w-sm w-[90%]"
          >
            <div className="flex items-center gap-3 w-full">
              <div className="p-2 bg-blue-100 text-blue-600 rounded-full animate-pulse">
                <Bot size={20} />
              </div>
              <div className="flex-1">
                <h4 className="font-bold text-slate-800 text-sm">
                  Disparo em Massa (Bot)
                </h4>
                <p className="text-xs text-slate-500">
                  {massSendProgress.info}
                </p>
              </div>
              <div className="font-bold text-blue-600">
                {(
                  (massSendProgress.sent / (massSendProgress.total || 1)) *
                  100
                ).toFixed(0)}
                %
              </div>
            </div>
            <div className="w-full bg-slate-100 rounded-full h-2 mt-2 overflow-hidden">
              <div
                className="bg-blue-600 h-2 rounded-full transition-all duration-300"
                style={{
                  width: `${(massSendProgress.sent / (massSendProgress.total || 1)) * 100}%`,
                }}
              />
            </div>

            <div className="flex gap-2 w-full mt-2">
              <button
                type="button"
                onClick={() => {
                  const newPaused = !isMassSendPaused;
                  massSendControlRef.current.paused = newPaused;
                  setIsMassSendPaused(newPaused);
                  showToast(newPaused ? "Robô pausado!" : "Robô retomado!");
                }}
                className={`flex-1 flex items-center justify-center gap-1.5 py-1.5 px-3 rounded-xl text-xs font-bold transition-all border ${
                  isMassSendPaused
                    ? "bg-emerald-50 text-emerald-700 border-emerald-200 hover:bg-emerald-100"
                    : "bg-amber-50 text-amber-700 border-amber-200 hover:bg-amber-100"
                }`}
              >
                {isMassSendPaused ? (
                  <>
                    <Play size={14} /> Retomar
                  </>
                ) : (
                  <>
                    <Pause size={14} /> Pausar
                  </>
                )}
              </button>
              <button
                type="button"
                onClick={() => {
                  if (
                    window.confirm(
                      "Deseja realmente cancelar o envio em massa?",
                    )
                  ) {
                    massSendControlRef.current.cancelled = true;
                    massSendControlRef.current.paused = false;
                    setIsMassSendPaused(false);
                    showToast("Cancelando envio em massa...");
                  }
                }}
                className="bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100 py-1.5 px-3 rounded-xl text-xs font-bold transition-all flex items-center justify-center gap-1.5"
              >
                <X size={14} /> Cancelar
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {profile?.mustChangePassword && (
        <PasswordChangeModal
          onComplete={async () => {
            try {
              if (user) {
                await updateDoc(doc(db, COLLECTIONS.USERS, user.uid), {
                  mustChangePassword: false,
                  updatedAt: serverTimestamp(),
                });
                setProfile((prev) =>
                  prev ? { ...prev, mustChangePassword: false } : null,
                );
                showToast("Senha atualizada com sucesso!");
              }
            } catch (err: any) {
              showToast("Erro ao atualizar status do perfil.", "error");
            }
          }}
        />
      )}

      <AnimatePresence>
        {isProfileModalOpen && profile && (
          <ProfileModal
            isOpen={isProfileModalOpen}
            onClose={() => setIsProfileModalOpen(false)}
            profile={profile}
            setProfile={setProfile}
            botConfig={botConfig}
            botStatuses={botStatuses}
            onToast={showToast}
          />
        )}
      </AnimatePresence>

      <AnimatePresence>
        {showInjectModal && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 z-[9999]"
          >
            <motion.div
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="bg-white rounded-3xl shadow-2xl w-full max-w-lg overflow-hidden"
            >
              <div className="bg-blue-600 p-6 text-white flex justify-between items-center">
                <div className="flex items-center gap-3">
                  <div className="bg-white/20 p-2 rounded-xl">
                    <MessageSquare size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold leading-none">Conectar via Injeção</h3>
                    <p className="text-blue-100 text-[10px] mt-1 font-medium">MODO AVANÇADO / ANTI-BLOQUEIO</p>
                  </div>
                </div>
                <button
                  onClick={() => setShowInjectModal(false)}
                  className="hover:bg-white/20 p-2 rounded-full transition-colors"
                >
                  <X size={20} />
                </button>
              </div>
              
              <div className="p-6 space-y-4 max-h-[70vh] overflow-y-auto custom-scrollbar">
                <div className="bg-amber-50 border border-amber-200 p-4 rounded-2xl text-[11px] text-amber-800 leading-relaxed shadow-sm">
                  <p className="font-bold flex items-center gap-2 mb-2 text-amber-900">
                    <AlertCircle size={14} />
                    COMO USAR (SOP):
                  </p>
                  <ol className="list-decimal ml-4 space-y-2">
                    <li>Abra o <strong>WhatsApp Web oficial</strong> em uma aba anônima do Chrome.</li>
                    <li>Faça o login normal pelo celular (QR Code ou Número).</li>
                    <li>Com o WhatsApp aberto, clique na extensão <strong>PESK Linker</strong> e copie o JSON.</li>
                    <li>Cole o código abaixo e clique em Injetar.</li>
                    <li><strong>IMPORTANTE:</strong> Feche a aba do WhatsApp Web imediatamente após o sucesso.</li>
                  </ol>
                </div>
                
                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5 ml-1">Número do Bot</label>
                  <div className="relative">
                    <div className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400">
                      <Hash size={16} />
                    </div>
                    <input 
                      type="text" 
                      value={injectBotNumber} 
                      onChange={e => setInjectBotNumber(e.target.value.replace(/\D/g, ''))}
                      className="w-full bg-slate-50 border border-slate-200 pl-10 pr-4 py-3 rounded-2xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all text-sm font-bold"
                      placeholder="5524999999999"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[10px] font-bold text-slate-500 uppercase tracking-widest mb-1.5 ml-1">Dados da Sessão (JSON)</label>
                  <textarea 
                    rows={6}
                    value={injectSessionData} 
                    onChange={e => setInjectSessionData(e.target.value)}
                    className="w-full bg-slate-50 border border-slate-200 px-4 py-3 rounded-2xl focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none transition-all font-mono text-[10px] resize-none"
                    placeholder='Cole aqui o JSON gerado pela extensão...'
                  />
                </div>

                <button
                  onClick={async () => {
                    if (!injectBotNumber || !injectSessionData) {
                      showToast("Por favor, preencha o número e cole o JSON.", "error");
                      return;
                    }
                    try {
                      let sessionDataObj;
                      try {
                        sessionDataObj = JSON.parse(injectSessionData);
                      } catch (e) {
                        showToast("JSON inválido! Copie novamente da extensão.", "error");
                        return;
                      }

                      await callBotApi("/api/inject", {
                        method: "POST",
                        body: { 
                          botNumber: injectBotNumber,
                          sessionData: sessionDataObj
                        }
                      });
                      
                      showToast("Sucesso! Sessão injetada. O bot está iniciando...", "success");
                      setShowInjectModal(false);
                      setInjectSessionData("");
                      
                      // Refresh status after injection
                      setTimeout(async () => {
                        try {
                          const data = await callBotApi("/api/status");
                          if (data && data.bots) setBotStatuses(data.bots);
                        } catch (e) {}
                      }, 4000);
                    } catch (err: any) {
                      showToast(`Erro na injeção: ${err.message}`, "error");
                    }
                  }}
                  className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-600/20 transition-all flex items-center justify-center gap-2 group"
                >
                  <span>Injetar Sessão e Conectar</span>
                  <ChevronRight size={18} className="group-hover:translate-x-1 transition-transform" />
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Sidebar */}
      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-40 w-64 bg-[#011a3c] border-r border-[#092e5c] transform transition-transform duration-300 ease-in-out lg:translate-x-0 lg:static lg:inset-0",
          isSidebarOpen ? "translate-x-0" : "-translate-x-full",
        )}
      >
        <div className="h-full flex flex-col">
          <div className="p-6 flex items-center space-x-3">
            {botConfig?.loginLogo ? (
              <img
                src={botConfig.loginLogo}
                alt="Logo"
                className="w-full max-h-12 object-contain drop-shadow-md"
                referrerPolicy="no-referrer"
              />
            ) : (
              <>
                <div className="w-10 h-10 bg-blue-600 rounded-xl flex items-center justify-center text-white shadow-lg shadow-blue-500/20">
                  <TrendingUp size={24} />
                </div>
                <h1 className="text-xl font-bold text-white tracking-tight">
                  Gestão Oeste pro
                </h1>
              </>
            )}
          </div>

          <nav className="flex-1 px-4 space-y-1 overflow-y-auto pr-2">
            {[
              { id: "dashboard", label: "Rotina", icon: LayoutDashboard },
              { id: "relatorios", label: "Relatórios", icon: BarChart3 },
              { id: "cadastro", label: "Novo Lead", icon: UserPlus },
              { id: "historico", label: "Histórico", icon: History },
              { id: "crm", label: "CRM WhatsApp", icon: MessageSquare },
              { id: "bases", label: "Bases", icon: Database },
              { id: "gap", label: "GAP Acadêmico", icon: GraduationCap },
              {
                id: "isencoes",
                label: "Acompanhamento de Isenções",
                icon: ShieldCheck,
              },
              { id: "fiesProuni", label: "Fies/Prouni", icon: FileText },
              { id: "mapao", label: "Mapão Acadêmico", icon: MapPin },
              { id: "cursos", label: "Cursos Disponíveis", icon: BookOpen },
              { id: "basesDisparo", label: "Bases de Disparo", icon: Globe },
              { id: "basesRenovacao", label: "Base Líquida", icon: Database },
              { id: "campanhas", label: "Campanhas", icon: Megaphone },
              { id: "calendario", label: "Plano de Ação", icon: Calendar },
              { id: "empresas", label: "Empresas Parceiras", icon: Building2 },
              {
                id: "controleConcorrencia",
                label: "Controle de Concorrência",
                icon: Target,
              },
              { id: "controleLigacoes", label: "Controle de Ligações", icon: Phone },
              { id: "evasao", label: "Evasão", icon: UserMinus },
              {
                id: "calculo",
                label: "Cálculo de Remuneração",
                icon: Calculator,
              },
              {
                id: "controlePagamentos",
                label: "Controle de Pagamentos",
                icon: Coins,
              },
              {
                id: "solicitacaoManutencao",
                label: "Solicitação de Manutenção",
                icon: Wrench,
              },
              {
                id: "controleInsumos",
                label: "Controle de Insumos",
                icon: Boxes,
              },
              {
                id: "controleInsumosComercial",
                label: "Controle de Insumos (Comercial)",
                icon: Boxes,
              },
              {
                id: "emailMarketing",
                label: "Envio de e-mail Marketing",
                icon: Mail,
              },
              { id: "admin", label: "Administração", icon: Settings },
            ].map(
              (item) =>
                canView(item.id) && (
                  <button
                    key={item.id}
                    onClick={() => {
                      setCurrentView(item.id);
                      setIsSidebarOpen(false);
                    }}
                    className={cn(
                      "w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-sm font-semibold transition-all",
                      currentView === item.id
                        ? "bg-blue-500/10 text-white"
                        : "text-slate-400 hover:bg-[#082a5c] hover:text-white",
                    )}
                  >
                    <item.icon size={20} />
                    <span>{item.label}</span>
                  </button>
                ),
            )}
          </nav>

          <div className="p-4 border-t border-[#092e5c]">
            <div className="bg-[#082a5c]/50 p-4 rounded-2xl mb-4">
              <p className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-1">
                Usuário
              </p>
              <p className="text-sm font-bold text-white truncate">
                {profile?.name}
              </p>
              <span className="inline-block mt-1 px-2 py-0.5 bg-blue-500/20 text-blue-400 text-[10px] font-bold rounded-full">
                {profile?.role}
              </span>
            </div>

            <div className="space-y-1">
              <button
                onClick={async () => {
                  if (user?.email) {
                    try {
                      await sendPasswordResetEmail(auth, user.email);
                      showToast("E-mail de redefinição enviado!");
                    } catch (err: any) {
                      showToast("Erro ao enviar e-mail.", "error");
                    }
                  }
                }}
                className="w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-sm font-semibold text-slate-400 hover:bg-[#082a5c] hover:text-white transition-all"
              >
                <KeyRound size={20} />
                <span>Trocar Senha</span>
              </button>

              <button
                onClick={() => signOut(auth)}
                className="w-full flex items-center space-x-3 px-4 py-3 rounded-xl text-sm font-semibold text-rose-400 hover:bg-rose-500/10 transition-all"
              >
                <LogOut size={20} />
                <span>Sair do Sistema</span>
              </button>
            </div>
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 flex flex-col min-w-0 overflow-hidden">
        <header className="h-16 bg-[#011a3c] border-b border-[#092e5c] flex items-center justify-between px-4 lg:px-8 shrink-0">
          <button
            onClick={() => setIsSidebarOpen(true)}
            className="lg:hidden p-2 text-slate-400 hover:bg-[#082a5c] rounded-lg"
          >
            <Menu size={24} />
          </button>
          <div className="flex-1 lg:flex-none flex items-center space-x-3 flex-wrap gap-y-1">
            <h2 className="text-lg font-bold text-white capitalize ml-2 lg:ml-0">
              {currentView.replace("-", " ")}
            </h2>
            <span className="px-2.5 py-1 bg-gradient-to-r from-blue-600 to-sky-500 text-white text-[10px] font-extrabold rounded-md shadow-sm uppercase tracking-wider">
              Servidor:{" "}
              {localStorage.getItem("servidor_selected") === "comercial"
                ? "Comercial"
                : "SM"}
            </span>
            {isOnline ? (
              <span className="flex items-center space-x-1.5 px-2.5 py-1 bg-emerald-500/10 text-emerald-400 text-[10px] font-extrabold rounded-md border border-emerald-500/20 shadow-sm uppercase tracking-wider">
                <span className="w-1.5 h-1.5 bg-emerald-400 rounded-full animate-pulse"></span>
                <span>Online / Sincronizado</span>
              </span>
            ) : (
              <span className="flex items-center space-x-1.5 px-2.5 py-1 bg-amber-500/10 text-amber-400 text-[10px] font-extrabold rounded-md border border-amber-500/20 shadow-sm uppercase tracking-wider animate-pulse">
                <span className="w-1.5 h-1.5 bg-amber-400 rounded-full"></span>
                <span>Sem Conexão (Modo Cache Offline)</span>
              </span>
            )}
          </div>
          <div className="flex items-center space-x-4">
            <div className="hidden md:flex items-center space-x-2 text-sm text-slate-400">
              <Calendar size={16} />
              <span>{new Date().toLocaleDateString("pt-BR")}</span>
            </div>
            <button
              onClick={() => setIsProfileModalOpen(true)}
              className="flex items-center space-x-1.5 bg-[#082a5c]/50 hover:bg-[#082a5c] text-white px-3 py-1.5 rounded-xl border border-[#092e5c] text-sm font-semibold transition-all active:scale-95 cursor-pointer"
            >
              <UserIcon size={15} className="text-slate-300" />
              <span>Perfil</span>
            </button>
          </div>
        </header>

        <div className="flex-1 overflow-y-auto p-4 lg:p-8">
          <AnimatePresence mode="wait">
            <motion.div
              key={currentView}
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.2 }}
            >
              {currentView === "dashboard" && (
                <DashboardView
                  leads={leads}
                  planner={[...planner]}
                  links={links}
                  profile={profile!}
                  onToast={showToast}
                  campanhas={campanhas}
                  bomDia={bomDia}
                  forecast={forecast}
                  periodos={periodos}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  qgLigacoes={qgLigacoes}
                  users={users}
                />
              )}
              {currentView === "relatorios" && (
                <RelatoriosView
                  leads={leads}
                  bases={bases}
                  fiesProuni={fiesProuni}
                  calendarioAcoes={calendarioAcoes}
                  pedidosCursos={pedidosCursos}
                  empresasParceiras={empresasParceiras}
                  insumosPedidos={insumosPedidos}
                  insumosEstoque={insumosEstoque}
                  insumosBaixas={insumosBaixas}
                  isencoes={isencoes}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  ligacoes={ligacoes}
                  solicitacoesManutencao={solicitacoesManutencao}
                  salesContacts={salesContacts}
                  analysisSchemes={analysisSchemes}
                  profile={profile!}
                  onToast={showToast}
                />
              )}
              {currentView === "cadastro" && (
                <CadastroView
                  onToast={showToast}
                  profile={profile!}
                  calendarioAcoes={calendarioAcoes}
                  uniqueUnidades={uniqueUnidades}
                />
              )}
              {currentView === "historico" && (
                <HistoricoView
                  leads={leads}
                  profile={profile!}
                  onToast={showToast}
                  users={users}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                  gap={gap}
                  basesRenovacao={basesRenovacao}
                  calendarioAcoes={calendarioAcoes}
                  pedidosCursos={pedidosCursos}
                />
              )}
              {currentView === "crm" && (
                ["Admin Master", "Líder/FDV"].includes(profile?.role || "") ? (
                  <CRMView
                    leads={leads}
                    bases={bases}
                    fiesProuni={fiesProuni}
                    gap={gap}
                    profile={profile!}
                    onSendBot={handleSendBotMessage}
                    onToast={showToast}
                  />
                ) : (
                  <div className="p-12 text-center bg-white rounded-3xl border border-slate-100 shadow-sm max-w-2xl mx-auto mt-10">
                    <div className="w-20 h-20 bg-slate-50 text-slate-300 rounded-full flex items-center justify-center mx-auto mb-6">
                      <MessageSquare size={40} />
                    </div>
                    <h3 className="text-2xl font-bold text-slate-800">Acesso Restrito</h3>
                    <p className="text-slate-500 mt-3 text-lg">A aba CRM está disponível apenas para Perfis Líder/FDV e Administradores enquanto estiver em fase de testes.</p>
                    <button 
                      onClick={() => setCurrentView('dashboard')}
                      className="mt-8 bg-blue-600 text-white px-8 py-3 rounded-xl font-bold hover:bg-blue-700 transition-all shadow-lg shadow-blue-500/20"
                    >
                      Voltar ao Início
                    </button>
                  </div>
                )
              )}
              {currentView === "bases" && (
                <BasesView
                  bases={bases}
                  profile={profile!}
                  onToast={showToast}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                  gap={gap}
                  basesRenovacao={basesRenovacao}
                />
              )}
              {currentView === "gap" && (
                <GapView
                  gap={gap}
                  onToast={showToast}
                  profile={profile}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                  calendarioAcoes={calendarioAcoes}
                />
              )}
              {currentView === "isencoes" && (
                <IsencoesView
                  isencoes={isencoes}
                  gap={gap}
                  onToast={showToast}
                  profile={profile!}
                />
              )}
              {currentView === "fiesProuni" && (
                <FiesProuniView
                  data={fiesProuni}
                  vagas={fiesProuniVagas}
                  onToast={showToast}
                  profile={profile!}
                  whatsappMessages={whatsappMessages}
                  periodos={periodos}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                />
              )}
              {currentView === "mapao" && (
                <MapaoAcademicoView
                  mapao={mapao}
                  onToast={showToast}
                  profile={profile!}
                />
              )}
              {currentView === "cursos" && (
                <CursosDisponiveisView
                  cursos={cursos}
                  onToast={showToast}
                  profile={profile!}
                />
              )}
              {currentView === "basesDisparo" && (
                <BasesDisparoView bases={basesDisparo} onToast={showToast} />
              )}
              {currentView === "basesRenovacao" && (
                <BasesRenovacaoView
                  bases={basesRenovacao}
                  onToast={showToast}
                  profile={profile}
                  whatsappMessages={whatsappMessages}
                  botConfig={botConfig}
                  onSendBot={handleSendBotMessage}
                  onMassSendBot={handleMassSendBotMessages}
                />
              )}
              {currentView === "campanhas" && (
                <CampanhasView campanhas={campanhas} onToast={showToast} />
              )}
              {currentView === "calculo" && <CalculoRemuneracaoView />}
              {currentView === "emailMarketing" && (
                <EmailMarketingView onToast={showToast} />
              )}
              {currentView === "controlePagamentos" && (
                <ControlePagamentosView
                  calendarioAcoes={calendarioAcoes}
                  users={users}
                  onToast={showToast}
                  profile={profile}
                />
              )}
              {currentView === "solicitacaoManutencao" && (
                <SolicitacoesManutencaoView
                  profile={profile}
                  onToast={showToast}
                  users={users}
                />
              )}
              {currentView === "controleInsumos" && (
                <ControleInsumosView
                  pedidos={insumosPedidos}
                  estoque={insumosEstoque}
                  profile={profile!}
                  onToast={showToast}
                  botConfig={botConfig}
                />
              )}
              {currentView === "controleInsumosComercial" && (
                <ControleInsumosComercialView
                  pedidos={insumosPedidosComercial}
                  estoque={insumosEstoqueComercial}
                  profile={profile!}
                  onToast={showToast}
                  botConfig={botConfig}
                />
              )}
              {currentView === "calendario" && (
                <CalendarioAcoesView
                  data={calendarioAcoes}
                  onToast={showToast}
                  profile={profile!}
                  initialData={initialActionData}
                  onClearInitialData={() => setInitialActionData(null)}
                  users={users}
                  callBotApi={callBotApi}
                  leads={leads}
                  gap={gap}
                  onSendWhatsApp={sendAppWhatsApp}
                />
              )}
              {currentView === "empresas" && (
                <EmpresasParceirasView
                  data={empresasParceiras}
                  leads={leads}
                  acoes={calendarioAcoes}
                  onToast={showToast}
                  cursos={cursos}
                  users={users}
                  onSendWhatsApp={sendAppWhatsApp}
                  botConfig={botConfig}
                  uniqueUnidades={uniqueUnidades}
                  profile={profile!}
                  onGenerateAction={(empresa) => {
                    setInitialActionData({
                      nome: `Ação na empresa ${empresa.nome}`,
                      local: empresa.endereco,
                      observacao: `Responsável: ${empresa.responsavel}\nTelefone: ${empresa.telefone}`,
                    });
                    setCurrentView("calendario");
                  }}
                />
              )}
              {currentView === "controleConcorrencia" && (
                <ControleConcorrenciaView
                  data={controleConcorrencia}
                  onToast={showToast}
                />
              )}
              {currentView === "controleLigacoes" && (
              <ControleLigacoesView
                leads={leads}
                bases={bases}
                acoes={calendarioAcoes}
                ligacoes={ligacoes}
                fiesProuni={fiesProuni}
                gap={gap}
                profile={profile!}
                onSaveLigacao={handleSaveLigacao}
                onToast={showToast}
              />
            )}

            {currentView === "evasao" && (
                <EvasaoView profile={profile} onToast={showToast} />
              )}
              {currentView === "admin" && (
                <AdminView
                  profile={profile}
                  users={users}
                  links={links}
                  onToast={showToast}
                  leads={leads}
                  bases={bases}
                  gap={gap}
                  planner={[...planner]}
                  campanhas={campanhas}
                  bomDia={bomDia}
                  forecast={forecast}
                  periodos={periodos}
                  whatsappMessages={whatsappMessages}
                  activeWhatsappTemplates={activeWhatsappTemplates}
                  setActiveWhatsappTemplates={setActiveWhatsappTemplates}
                  empresasParceiras={empresasParceiras}
                  botConfig={botConfig}
                  botStatuses={botStatuses}
                  setBotStatuses={setBotStatuses}
                  callBotApi={callBotApi}
                  metaDia={metaDia}
                  metaSM={metaSM}
                  metaCursos={metaCursos}
                  qgLigacoes={qgLigacoes}
                  cursos={cursos}
                  uniqueUnidades={uniqueUnidades}
                  analysisSchemes={analysisSchemes}
                  onSaveAnalysisScheme={handleSaveAnalysisScheme}
                  onDeleteAnalysisScheme={handleDeleteAnalysisScheme}
                  setShowInjectModal={setShowInjectModal}
                />
              )}
            </motion.div>
          </AnimatePresence>

          <footer className="mt-12 py-6 border-t border-slate-200 text-center">
            <p className="text-sm text-slate-500 font-medium">
              Sistema Criado por{" "}
              <span className="font-bold text-slate-900">Agencia Argo's</span> -
              <a
                href={getWhatsAppUrl(
                  "24992777019",
                  "Gostaria de realizar um orçamento para um sistema",
                )}
                target="_blank"
                rel="noopener noreferrer"
                className="ml-1 text-blue-600 hover:underline font-bold"
              >
                Telefone: (24) 99277-7019
              </a>
            </p>
          </footer>
        </div>
      </main>

      {/* Mobile Sidebar Overlay */}
      {isSidebarOpen && (
        <div
          className="fixed inset-0 bg-slate-900/50 z-30 lg:hidden backdrop-blur-sm"
          onClick={() => setIsSidebarOpen(false)}
        />
      )}
    </div>
  );
}

// --- View Components ---

function AvisosView() {
  return null;
}

function AuthScreen({
  onToast,
  botConfig,
}: {
  onToast: (m: string, t?: "success" | "error") => void;
  botConfig?: BotConfig;
}) {
  const [isLogin, setIsLogin] = useState(true);
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [name, setName] = useState("");
  const [servidor, setServidor] = useState<"principal" | "comercial">(
    (localStorage.getItem("servidor_selected") as "principal" | "comercial") ||
      "principal",
  );
  const [loading, setLoading] = useState(false);

  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [showInstallGuide, setShowInstallGuide] = useState(false);
  const [showApkModal, setShowApkModal] = useState(false);
  const [isAppInstalled, setIsAppInstalled] = useState(false);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const isStandalone =
      window.matchMedia("(display-mode: standalone)").matches ||
      (window.navigator as any).standalone ||
      document.referrer.includes("android-app://");
    setIsAppInstalled(isStandalone);

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
    return () => {
      window.removeEventListener(
        "beforeinstallprompt",
        handleBeforeInstallPrompt,
      );
    };
  }, []);

  const handleInstallClick = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: { outcome: string }) => {
        if (choiceResult.outcome === "accepted") {
          console.log("User accepted the PWA install prompt");
          setIsAppInstalled(true);
        }
        setDeferredPrompt(null);
      });
    } else {
      setShowInstallGuide((prev) => !prev);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    if (!isLogin && password.length < 6) {
      onToast("A senha deve ter pelo menos 6 caracteres.", "error");
      setLoading(false);
      return;
    }
    try {
      if (isLogin) {
        try {
          // Attempt login on the CURRENTLY SELECTED server
          await signInWithEmailAndPassword(auth, email, password);
        } catch (err: any) {
          // If user login fails with invalid credentials or user not found,
          // let's check programmatically if the credentials are valid on the OTHER server!
          const isUserNotFound =
            err.code === "auth/user-not-found" ||
            err.code === "auth/invalid-credential";
          if (isUserNotFound) {
            const currentSelected = servidor;
            const alternativeServer =
              currentSelected === "principal" ? "comercial" : "principal";

            // Build the alternative config
            const altConfig =
              alternativeServer === "comercial"
                ? firebaseConfigComercial
                : firebaseConfigPrincipal;

            // Resolve alternative app
            let altApp;
            try {
              altApp = getApp("alternative_login_check");
            } catch {
              altApp = initializeApp(altConfig, "alternative_login_check");
            }
            const altAuth = getAuth(altApp);

            try {
              // Attempt login on the ALTERNATIVE server
              await signInWithEmailAndPassword(altAuth, email, password);
              // SUCCESS on the other server! Let's update localStorage and reload to apply the active configuration
              localStorage.setItem("servidor_selected", alternativeServer);
              onToast(
                `Login bem sucedido! Redirecionando para o Servidor ${alternativeServer === "principal" ? "Principal" : "Comercial"}...`,
                "success",
              );
              setTimeout(() => {
                window.location.reload();
              }, 1200);
              return;
            } catch (altErr) {
              // Failed on both servers, throw the original authentication error
              throw err;
            }
          } else {
            throw err;
          }
        }
      } else {
        const userCred = await createUserWithEmailAndPassword(
          auth,
          email,
          password,
        );
        // Pack the chosen servidor into displayName so App.tsx can extract it
        await updateProfile(userCred.user, {
          displayName: `${name}|${servidor}`,
        });
        onToast("Conta criada com sucesso!");
      }
    } catch (err: any) {
      console.error("Auth error details (AuthScreen):", {
        code: err.code,
        message: err.message,
        stack: err.stack,
      });
      let friendlyMessage = err.message;
      if (
        err.code === "auth/user-not-found" ||
        err.code === "auth/invalid-credential"
      ) {
        friendlyMessage =
          "E-mail ou senha inválidos em ambos os servidores (Principal / Comercial).";
      } else if (err.code === "auth/wrong-password") {
        friendlyMessage = "Senha incorreta.";
      }
      onToast(`Erro: ${friendlyMessage}`, "error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#011430] flex flex-col md:flex-row relative overflow-hidden font-sans text-white">
      {/* Absolute Ambient Background Lights */}
      <div className="absolute top-0 right-0 w-[500px] h-[500px] rounded-full bg-blue-500/15 blur-[120px] pointer-events-none z-0" />
      <div className="absolute bottom-0 left-0 w-[500px] h-[500px] rounded-full bg-cyan-500/10 blur-[130px] pointer-events-none z-0" />

      {/* LEFT COLUMN: Login panel container */}
      <div className="w-full md:w-[42%] lg:w-[38%] xl:w-[34%] bg-[#011a3c] border-r border-[#092e5c] p-8 sm:p-12 md:p-16 flex flex-col justify-between relative z-10 shadow-2xl min-h-screen">
        <div className="my-auto space-y-8">
          <div>
            {botConfig?.loginLogo ? (
              <div className="mb-6 flex">
                <img
                  src={botConfig.loginLogo}
                  alt="Logo"
                  className="max-h-32 w-full object-contain drop-shadow-lg"
                  referrerPolicy="no-referrer"
                />
              </div>
            ) : (
              <>
                <div className="w-16 h-16 bg-gradient-to-tr from-sky-500 to-blue-600 rounded-2xl flex items-center justify-center text-white shadow-lg shadow-sky-500/20 mb-6">
                  <TrendingUp size={32} />
                </div>
                <h2 className="text-3xl font-extrabold text-white tracking-tight">
                  Gestão Oeste pro
                </h2>
              </>
            )}
            <p className="text-slate-400 mt-2 text-sm">
              {isLogin
                ? "Bem-vindo de volta! Insira suas credenciais:"
                : "Preencha os dados e crie sua conta agora:"}
            </p>
          </div>

          {/* Servidor Selector (Principal vs Comercial) */}
          <div className="flex bg-[#032554] p-1.5 rounded-2xl border border-[#0b3c7c] shadow-inner">
            <button
              type="button"
              onClick={() => {
                if (servidor !== "principal") {
                  localStorage.setItem("servidor_selected", "principal");
                  window.location.reload();
                }
              }}
              className={`flex-1 py-3 text-xs font-extrabold rounded-xl transition-all cursor-pointer ${servidor === "principal" ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20" : "text-slate-400 hover:text-white"}`}
            >
              Principal
            </button>
            <button
              type="button"
              onClick={() => {
                if (servidor !== "comercial") {
                  localStorage.setItem("servidor_selected", "comercial");
                  window.location.reload();
                }
              }}
              className={`flex-1 py-3 text-xs font-extrabold rounded-xl transition-all cursor-pointer ${servidor === "comercial" ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20" : "text-slate-400 hover:text-white"}`}
            >
              Comercial
            </button>
          </div>

          {/* Alerta de Configuração se os projetos forem idênticos no modo Principal */}
          {servidor === "principal" &&
            firebaseConfigPrincipal.apiKey === "AIzaSyBexxjzDAuNSgY90rlVqpz4AQZDE-QwSG4" && (
              <div className="p-3 bg-amber-500/10 border border-amber-500/20 rounded-xl text-amber-200 text-[10px] font-bold leading-relaxed flex items-start gap-2 animate-pulse">
                <AlertCircle size={14} className="flex-shrink-0 mt-0.5" />
                <p>
                  O servidor Principal está usando as credenciais do servidor
                  Comercial. Configure as variáveis de ambiente (API Key e App
                  ID) do projeto gestaopro-761e1 para habilitar o acesso.
                </p>
              </div>
            )}

          <form onSubmit={handleSubmit} className="space-y-4 font-sans">
            {!isLogin && (
              <div>
                <label className="block text-xs font-bold text-slate-300 uppercase tracking-wider mb-1.5">
                  Nome Completo
                </label>
                <input
                  type="text"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  required
                  className="w-full bg-[#032654] border border-[#0d4182] text-white px-4 py-3.5 rounded-xl focus:ring-2 focus:ring-sky-500 outline-none transition-all placeholder-slate-500 text-sm font-medium"
                  placeholder="Seu nome"
                />
              </div>
            )}
            <div>
              <label className="block text-xs font-bold text-slate-300 uppercase tracking-wider mb-1.5">
                Email
              </label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                required
                className="w-full bg-[#032654] border border-[#0d4182] text-white px-4 py-3.5 rounded-xl focus:ring-2 focus:ring-sky-500 outline-none transition-all placeholder-slate-500 text-sm font-medium"
                placeholder="seu@email.com"
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-300 uppercase tracking-wider mb-1.5">
                Senha
              </label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                required
                className="w-full bg-[#032654] border border-[#0d4182] text-white px-4 py-3.5 rounded-xl focus:ring-2 focus:ring-sky-500 outline-none transition-all placeholder-slate-500 text-sm font-medium"
                placeholder="••••••••"
              />
            </div>

            {isLogin && (
              <div className="flex justify-end">
                <button
                  type="button"
                  onClick={async () => {
                    let resetEmail = email;
                    if (!resetEmail) {
                      const inputEmail = window.prompt(
                        "Por favor, digite seu e-mail para receber o link de redefinição de senha:",
                      );
                      if (!inputEmail) return;
                      resetEmail = inputEmail;
                    }
                    try {
                      await sendPasswordResetEmail(auth, resetEmail);
                      onToast(
                        "E-mail de redefinição enviado com sucesso! Verifique sua caixa de entrada.",
                        "success",
                      );
                    } catch (err: any) {
                      onToast(
                        "Erro ao enviar e-mail. Verifique se o endereço é válido.",
                        "error",
                      );
                    }
                  }}
                  className="text-xs font-bold text-sky-400 hover:text-sky-300 hover:underline transition-colors cursor-pointer"
                >
                  Esqueci minha senha
                </button>
              </div>
            )}

            <button
              type="submit"
              disabled={loading}
              className="w-full bg-gradient-to-tr from-sky-500 to-blue-600 hover:from-sky-600 hover:to-blue-700 text-white font-bold py-4 rounded-xl shadow-lg transition-all disabled:opacity-50 disabled:cursor-not-allowed active:scale-[0.98] cursor-pointer"
            >
              {loading
                ? "Processando..."
                : isLogin
                  ? "Entrar no Sistema"
                  : "Criar Minha Conta"}
            </button>
          </form>

          <div className="mt-8 text-center pt-2 border-t border-[#092e5c]">
            <button
              onClick={() => setIsLogin(!isLogin)}
              className="text-sm font-semibold text-sky-400 hover:text-sky-300 hover:underline cursor-pointer"
            >
              {isLogin
                ? "Não tem uma conta? Cadastre-se"
                : "Já tem uma conta? Faça login"}
            </button>
          </div>

          {/* Android App Promotion Card on Login */}
          {!isAppInstalled && (
            <div className="mt-8 pt-6 border-t border-[#092e5c] space-y-4">
              <div className="bg-[#032554]/60 p-5 rounded-2xl border border-sky-500/10 text-white relative overflow-hidden transition-all duration-300">
                <div className="flex items-start space-x-3.5">
                  <div className="p-2.5 bg-sky-950/80 rounded-xl border border-sky-500/20 text-emerald-400 flex items-center justify-center shadow shrink-0">
                    <Smartphone size={24} />
                  </div>
                  <div className="space-y-1">
                    <h4 className="text-sm font-extrabold text-white leading-tight">
                      Instalar Aplicativo (Android)
                    </h4>
                    <p className="text-xs text-slate-300 leading-relaxed font-semibold">
                      Deseja usar no celular? Instale o App para usar{" "}
                      <strong className="text-emerald-400 font-extrabold">
                        com ou sem internet
                      </strong>
                      . Sincroniza automático ao conectar.
                    </p>
                  </div>
                </div>

                <div className="flex items-center space-x-2 mt-4">
                  <button
                    onClick={handleInstallClick}
                    className="flex-1 flex items-center justify-center space-x-1.5 px-3 py-2.5 bg-emerald-500 hover:bg-emerald-600 text-slate-950 font-extrabold text-xs rounded-lg shadow-md hover:scale-[1.01] transition-all cursor-pointer"
                  >
                    <Download size={14} />
                    <span>Instalar no Aparelho</span>
                  </button>
                  <button
                    onClick={() => setShowInstallGuide(!showInstallGuide)}
                    className="px-3 py-2.5 bg-white/10 hover:bg-white/15 text-slate-100 font-bold text-xs rounded-lg transition-all cursor-pointer flex-1"
                  >
                    Instruções
                  </button>
                </div>

                {showInstallGuide && (
                  <div className="mt-4 pt-4 border-t border-[#092e5c] space-y-3 text-xs">
                    <div className="space-y-1">
                      <span className="text-[10px] font-black uppercase text-emerald-400">
                        Passo 1:
                      </span>
                      <p className="text-slate-300 font-semibold leading-relaxed">
                        Abra este endereço no{" "}
                        <strong className="text-white font-bold">
                          Google Chrome
                        </strong>{" "}
                        do seu Android.
                      </p>
                    </div>
                    <div className="space-y-1">
                      <span className="text-[10px] font-black uppercase text-emerald-400">
                        Passo 2:
                      </span>
                      <p className="text-slate-300 font-semibold leading-relaxed">
                        Toque nos{" "}
                        <strong className="text-white font-bold">
                          três pontinhos (⋮)
                        </strong>{" "}
                        no canto superior direito.
                      </p>
                    </div>
                    <div className="space-y-1">
                      <span className="text-[10px] font-black uppercase text-emerald-400">
                        Passo 3:
                      </span>
                      <p className="text-slate-300 font-semibold leading-relaxed">
                        Selecione{" "}
                        <strong className="text-emerald-400 font-extrabold">
                          "Instalar aplicativo"
                        </strong>{" "}
                        ou{" "}
                        <strong className="text-emerald-400 font-extrabold">
                          "Adicionar à tela inicial"
                        </strong>
                        .
                      </p>
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>

        {/* Humble system credits info */}
        <div className="text-center text-[10px] text-slate-500 font-mono tracking-widest mt-6">
          OESTE HUNTER © {new Date().getFullYear()}
        </div>
      </div>

      {/* RIGHT COLUMN: The majestic interactive Oeste Hunter logo artwork or Custom Logo */}
      <div className="hidden md:flex flex-1 items-center justify-center bg-[#01112c] p-12 relative overflow-hidden z-0">
        {/* Subtle grid mesh backdrop */}
        <div className="absolute inset-0 bg-[radial-gradient(#082a5c_1px,transparent_1px)] [background-size:24px_24px] opacity-25" />

        {/* Animated ambient outer halo circles */}
        <div className="absolute w-[600px] h-[600px] border border-[#0d4182]/20 rounded-full animate-pulse" />
        <div className="absolute w-[800px] h-[800px] border border-[#0d4182]/10 rounded-full opacity-60" />

        {/* SVG ART Container */}
        <div className="relative z-10 w-full flex justify-center">
          {botConfig?.loginLogo ? (
            <img
              src={botConfig.loginLogo}
              alt="Logo Promocional"
              className="w-full max-w-[560px] aspect-square rounded-3xl object-contain drop-shadow-[0_35px_60px_rgba(14,116,253,0.35)] border border-slate-700/40 p-12 bg-[#011a3c]/50 animate-fade-in"
              referrerPolicy="no-referrer"
            />
          ) : (
            /* Oeste Hunter Badge SVG */
            <svg
              viewBox="0 0 1000 1000"
              className="w-full max-w-[560px] aspect-square drop-shadow-[0_25px_60px_rgba(14,116,253,0.35)] select-none"
            >
              <defs>
                <linearGradient
                  id="blueRingGrad"
                  x1="0%"
                  y1="0%"
                  x2="100%"
                  y2="100%"
                >
                  <stop offset="0%" stopColor="#0a397a" />
                  <stop offset="50%" stopColor="#125cb5" />
                  <stop offset="100%" stopColor="#082c5f" />
                </linearGradient>
                <linearGradient
                  id="wolfEyeGrad"
                  x1="0%"
                  y1="0%"
                  x2="100%"
                  y2="0%"
                >
                  <stop offset="0%" stopColor="#00f0ff" />
                  <stop offset="100%" stopColor="#00a8ff" />
                </linearGradient>
                <linearGradient
                  id="muzzleGrad"
                  x1="0%"
                  y1="0%"
                  x2="0%"
                  y2="100%"
                >
                  <stop offset="0%" stopColor="#ffffff" />
                  <stop offset="100%" stopColor="#cfd8dc" />
                </linearGradient>
                <linearGradient
                  id="bannerGrad"
                  x1="0%"
                  y1="0%"
                  x2="100%"
                  y2="0%"
                >
                  <stop offset="0%" stopColor="#010f24" />
                  <stop offset="50%" stopColor="#051c3d" />
                  <stop offset="100%" stopColor="#010d21" />
                </linearGradient>
                <filter id="glow" x="-20%" y="-20%" width="140%" height="140%">
                  <feGaussianBlur stdDeviation="8" result="blur" />
                  <feComposite in="SourceGraphic" in2="blur" operator="over" />
                </filter>
                <filter
                  id="eyeGlow"
                  x="-50%"
                  y="-50%"
                  width="200%"
                  height="200%"
                >
                  <feGaussianBlur stdDeviation="6" result="blur" />
                  <feComposite in="SourceGraphic" in2="blur" operator="over" />
                </filter>
              </defs>

              {/* Target Crosshairs Reticle */}
              <g stroke="#ffffff" strokeWidth="2.5" opacity="0.3">
                {/* Vertical Crosshair Line */}
                <line x1="500" y1="20" x2="500" y2="980" />
                {/* Horizontal Crosshair Line */}
                <line x1="20" y1="500" x2="980" y2="500" />

                {/* Target ticks (top, bottom, left, right) */}
                <line x1="500" y1="80" x2="520" y2="80" />
                <line x1="500" y1="140" x2="515" y2="140" />
                <line x1="500" y1="200" x2="520" y2="200" />

                <line x1="500" y1="920" x2="520" y2="920" />
                <line x1="500" y1="860" x2="515" y2="860" />
                <line x1="500" y1="800" x2="520" y2="800" />

                <line x1="80" y1="500" x2="80" y2="520" />
                <line x1="140" y1="500" x2="140" y2="515" />
                <line x1="200" y1="500" x2="200" y2="520" />

                <line x1="920" y1="500" x2="920" y2="520" />
                <line x1="860" y1="500" x2="860" y2="515" />
                <line x1="800" y1="500" x2="800" y2="520" />
              </g>

              {/* Target Reticle Outer Box ticks */}
              <rect
                x="480"
                y="40"
                width="40"
                height="20"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />
              <rect
                x="480"
                y="940"
                width="40"
                height="20"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />
              <rect
                x="40"
                y="480"
                width="20"
                height="40"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />
              <rect
                x="940"
                y="480"
                width="20"
                height="40"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                opacity="0.4"
              />

              {/* 1. Outer target circle with dashes */}
              <circle
                cx="500"
                cy="500"
                r="445"
                fill="none"
                stroke="#ffffff"
                strokeWidth="3"
                strokeDasharray="16 20"
                opacity="0.35"
              />

              {/* 2. Concentric circle borders */}
              <circle
                cx="500"
                cy="500"
                r="415"
                fill="none"
                stroke="#ffffff"
                strokeWidth="2.5"
                opacity="0.4"
              />

              {/* 3. Main Thick Ring Outer Rim */}
              <circle
                cx="500"
                cy="500"
                r="400"
                fill="none"
                stroke="#ffffff"
                strokeWidth="4"
              />

              {/* 4. The Mighty Blue Ring Body */}
              <circle
                cx="500"
                cy="500"
                r="348"
                fill="none"
                stroke="url(#blueRingGrad)"
                strokeWidth="100"
              />

              {/* 5. Inner Rim of the Blue Ring */}
              <circle
                cx="500"
                cy="500"
                r="298"
                fill="none"
                stroke="#ffffff"
                strokeWidth="4"
              />

              {/* 6. Main Inner Graphic Backdrop (Turquoise circle) */}
              <circle
                cx="500"
                cy="500"
                r="294"
                fill="#009be1"
                stroke="#ffffff"
                strokeWidth="2"
              />
              <circle cx="500" cy="500" r="275" fill="#0388c7" />

              {/* Curves for circular text alignment */}
              {/* Path for 'OESTE' arched on top (left-to-right) */}
              <path
                id="topArchPath"
                d="M 160,500 A 340,340 0 0,1 840,500"
                fill="none"
              />

              {/* Path for 'OESTE HUNTER' arched on bottom (right-to-left) */}
              <path
                id="bottomArchPath"
                d="M 840,500 A 340,340 0 0,1 160,500"
                fill="none"
              />

              {/* Arched Texts */}
              <text
                fontFamily="'Inter', sans-serif"
                fontWeight="900"
                fontSize="75"
                fill="#ffffff"
                letterSpacing="18"
              >
                <textPath
                  href="#topArchPath"
                  startOffset="50%"
                  textAnchor="middle"
                >
                  OESTE
                </textPath>
              </text>

              <text
                fontFamily="'Inter', sans-serif"
                fontWeight="900"
                fontSize="44"
                fill="#ffffff"
                letterSpacing="14"
              >
                <textPath
                  href="#bottomArchPath"
                  startOffset="50%"
                  textAnchor="middle"
                >
                  OESTE HUNTER
                </textPath>
              </text>

              {/* ======================================= */}
              {/* WOLF HEAD INTERIOR ELEMENT MASCOT ART   */}
              {/* ======================================= */}
              <g id="wolfMascot" transform="translate(0, -35)">
                {/* Wolf Ears Behind */}
                {/* Left ear dark back */}
                <polygon
                  points="350,330 435,420 380,480"
                  fill="#020f2b"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                {/* Left ear internal blue */}
                <polygon points="365,345 425,415 385,465" fill="#0096e6" />

                {/* Right ear dark back */}
                <polygon
                  points="650,330 565,420 620,480"
                  fill="#020f2b"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                {/* Right ear internal blue */}
                <polygon points="635,345 575,415 615,465" fill="#0096e6" />

                {/* Wolf Forehead and Cheek structure */}
                {/* Base Head polygon */}
                <polygon
                  points="500,380 340,500 370,625 500,680 630,625 660,500"
                  fill="#03112b"
                />

                {/* White outer framing highlights (Cheek fur) */}
                {/* Left cheek outer fluff */}
                <polygon
                  points="340,500 310,560 385,585"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />
                <polygon
                  points="310,560 330,630 400,610"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />

                {/* Right cheek outer fluff */}
                <polygon
                  points="660,500 690,560 615,585"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />
                <polygon
                  points="690,560 670,630 600,610"
                  fill="#ffffff"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />

                {/* Intermediate Blue Shadows Cheeks */}
                <polygon points="340,500 385,585 410,515" fill="#0a3c7c" />
                <polygon points="660,500 615,585 590,515" fill="#0a3c7c" />

                {/* Side Dark fur shades */}
                <polygon
                  points="410,515 385,585 440,590 460,530"
                  fill="#00183b"
                />
                <polygon
                  points="590,515 615,585 560,590 540,530"
                  fill="#00183b"
                />

                {/* Center forehead wolf shield (cyan core) */}
                <polygon points="500,380 460,470 500,510" fill="#00bdff" />
                <polygon points="500,380 540,470 500,510" fill="#00bdff" />

                <polygon points="500,395 470,470 500,500" fill="#ffffff" />
                <polygon points="500,395 530,470 500,500" fill="#ffffff" />

                {/* Wolf Eyes Areas (Black framing masks) */}
                <polygon
                  points="420,490 470,510 460,535 410,515"
                  fill="#010614"
                />
                <polygon
                  points="580,490 530,510 540,535 590,515"
                  fill="#010614"
                />

                {/* Fierce Cyan Eyes */}
                <polygon
                  points="432,498 462,510 452,525 430,512"
                  fill="url(#wolfEyeGrad)"
                  filter="url(#eyeGlow)"
                />
                <polygon
                  points="568,498 538,510 548,525 570,512"
                  fill="url(#wolfEyeGrad)"
                  filter="url(#eyeGlow)"
                />

                {/* Wolf Nose Bridge */}
                <polygon
                  points="500,510 460,530 475,590 500,610"
                  fill="#020f26"
                />
                <polygon
                  points="500,510 540,530 525,590 500,610"
                  fill="#020f26"
                />

                <polygon
                  points="500,510 480,530 485,585 500,600"
                  fill="#0080cf"
                />
                <polygon
                  points="500,510 520,530 515,585 500,600"
                  fill="#0080cf"
                />

                {/* Muzzle (White muzzle side facets) */}
                <polygon
                  points="500,610 440,590 445,635 500,665"
                  fill="url(#muzzleGrad)"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />
                <polygon
                  points="500,610 560,590 555,635 500,665"
                  fill="url(#muzzleGrad)"
                  stroke="#ffffff"
                  strokeWidth="2.5"
                />

                {/* Black Nose Tip */}
                <polygon points="500,620 475,605 525,605" fill="#010614" />
                <polygon points="500,620 485,635 515,635" fill="#010614" />
                <polygon
                  points="475,605 525,605 515,635 485,635"
                  fill="#010614"
                />
                {/* Nose shine */}
                <circle cx="500" cy="612" r="3" fill="#ffffff" />
              </g>

              {/* Left and Right Side Banners - NPS and CAPTAÇÃO DE ALUNOS */}
              {/* LEFT BANNER (NPS) */}
              <g id="leftBanner">
                <polygon
                  points="6,505 130,505 110,615 6,615 30,560"
                  fill="#07336e"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                <polygon
                  points="12,515 120,515 104,605 12,605"
                  fill="#0b4594"
                />

                <text
                  x="63"
                  y="578"
                  textAnchor="middle"
                  fontFamily="'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="46"
                  fill="#ffffff"
                  letterSpacing="1"
                >
                  NPS
                </text>
              </g>

              {/* RIGHT BANNER (CAPTAÇÃO DE ALUNOS) */}
              <g id="rightBanner">
                <polygon
                  points="994,505 870,505 890,615 994,615 970,560"
                  fill="#07336e"
                  stroke="#ffffff"
                  strokeWidth="3"
                />
                <polygon
                  points="988,515 880,515 896,605 988,605"
                  fill="#0b4594"
                />

                <text
                  x="934"
                  y="555"
                  textAnchor="middle"
                  fontFamily="'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="20"
                  fill="#ffffff"
                  letterSpacing="2"
                >
                  CAPTAÇÃO
                </text>
                <text
                  x="934"
                  y="583"
                  textAnchor="middle"
                  fontFamily="'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="18"
                  fill="#ffffff"
                  letterSpacing="1"
                >
                  DE ALUNOS
                </text>
              </g>

              {/* Giant Horizontal Bottom Ribbon - HUNTER */}
              <g id="hunterBanner" transform="translate(0, 10)">
                {/* Banner Ribbon shadow back folds */}
                <polygon points="180,685 240,685 220,625" fill="#01050e" />
                <polygon points="820,685 760,685 780,625" fill="#01050e" />

                {/* Front Main Banner Body */}
                <polygon
                  points="180,625 820,625 790,735 210,735"
                  fill="url(#bannerGrad)"
                  stroke="#ffffff"
                  strokeWidth="6"
                />

                {/* Inner stroke accent */}
                <polygon
                  points="195,635 805,635 780,725 220,725"
                  fill="none"
                  stroke="#2575fc"
                  strokeWidth="3.5"
                  opacity="0.8"
                />

                {/* Bold Athletics display text - HUNTER */}
                <text
                  x="500"
                  y="702"
                  textAnchor="middle"
                  fontFamily="'Impact', 'Arial Black', 'Inter', sans-serif"
                  fontWeight="900"
                  fontSize="105"
                  fill="#ffffff"
                  letterSpacing="5"
                  filter="url(#glow)"
                >
                  HUNTER
                </text>
              </g>
            </svg>
          )}
        </div>
      </div>

      {/* APK Information Modal */}
      {showApkModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-900/80 backdrop-blur-sm animate-fade-in">
          <div className="bg-white rounded-3xl p-8 max-w-md w-full shadow-2xl relative overflow-hidden">
            {/* Header / Icon */}
            <div className="mx-auto w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-6 shadow-inner">
              <Download size={32} className="text-blue-600" />
            </div>

            <h3 className="text-xl font-black text-slate-800 text-center mb-2">
              Download do Arquivo APK
            </h3>
            <p className="text-sm text-slate-500 text-center font-medium leading-relaxed mb-6">
              O projeto nativo Android foi gerado e configurado usando
              Capacitor.
            </p>

            <div className="space-y-4 bg-slate-50 p-5 rounded-2xl border border-slate-100 mb-6">
              <div className="flex items-start gap-3">
                <div className="bg-emerald-100 text-emerald-600 p-1.5 rounded-lg shrink-0 mt-0.5">
                  <Smartphone size={18} />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-slate-800">
                    1. Instalação Imediata (Via Chrome)
                  </h4>
                  <p className="text-xs text-slate-600 mt-1 font-medium">
                    Recomendado: Feche esta aba e clique em{" "}
                    <strong className="text-emerald-600">
                      "Instalar no Aparelho"
                    </strong>{" "}
                    na tela de login (usando o Google Chrome no seu celular)
                    para instalação automática PWA/WebAPK direta no aparelho.
                  </p>
                </div>
              </div>

              <div className="w-full h-px bg-slate-200"></div>

              <div className="flex items-start gap-3">
                <div className="bg-amber-100 text-amber-600 p-1.5 rounded-lg shrink-0 mt-0.5">
                  <Download size={18} />
                </div>
                <div>
                  <h4 className="text-sm font-bold text-slate-800">
                    2. Desenvolvedores (Compilação Nativa)
                  </h4>
                  <p className="text-xs text-slate-600 mt-1 font-medium">
                    Devido as limitações do ambiente Cloud, o arquivo{" "}
                    <strong className="font-bold">.apk</strong> real precisa ser
                    compilado localmente: Exporte os arquivos do app, abra a
                    pasta{" "}
                    <code className="bg-slate-200 px-1 py-0.5 rounded font-mono text-amber-800">
                      android/
                    </code>{" "}
                    no Android Studio e compile o APK.
                  </p>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-3">
              <button
                onClick={() => setShowApkModal(false)}
                className="flex-1 px-4 py-3 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl font-bold transition-all cursor-pointer"
              >
                Voltar à Tela Inicial
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function DashboardView({
  leads,
  planner,
  links,
  profile,
  onToast,
  campanhas,
  bomDia,
  forecast,
  periodos,
  metaDia,
  metaSM,
  metaCursos,
  qgLigacoes,
  users,
}: {
  leads: Lead[];
  planner: PlannerTask[];
  links: LinkUtil[];
  profile: UserProfile;
  onToast: (m: string, t?: "success" | "error") => void;
  campanhas: Campanha[];
  bomDia: BomDiaCaptacao[];
  forecast: ForecastCaptacao[];
  periodos: PeriodoCaptacao[];
  metaDia: MetaDia[];
  metaSM: MetaSM[];
  metaCursos: MetaCurso[];
  qgLigacoes: QgLigacao[];
  users: UserProfile[];
}) {
  const [isCustomizing, setIsCustomizing] = useState(false);
  const [linksSearchTerm, setLinksSearchTerm] = useState("");
  const [linksFilterLocal, setLinksFilterLocal] = useState("");
  const [deferredPrompt, setDeferredPrompt] = useState<any>(null);
  const [showInstallGuide, setShowInstallGuide] = useState(false);
  const [isAppInstalled, setIsAppInstalled] = useState(false);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const isStandalone =
      window.matchMedia("(display-mode: standalone)").matches ||
      (window.navigator as any).standalone ||
      document.referrer.includes("android-app://");
    setIsAppInstalled(isStandalone);

    const handleBeforeInstallPrompt = (e: Event) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener("beforeinstallprompt", handleBeforeInstallPrompt);
    return () => {
      window.removeEventListener(
        "beforeinstallprompt",
        handleBeforeInstallPrompt,
      );
    };
  }, []);

  const handleInstallClick = () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      deferredPrompt.userChoice.then((choiceResult: { outcome: string }) => {
        if (choiceResult.outcome === "accepted") {
          console.log("User accepted the PWA install prompt");
          setIsAppInstalled(true);
        }
        setDeferredPrompt(null);
      });
    } else {
      setShowInstallGuide(true);
    }
  };

  const defaultWidgets = {
    stats: false,
    links: true,
    planner: true,
    campanhas: false,
    bomDia: true,
    forecast: true,
    periodo: true,
    qgLigacoes: true,
    metaSM: true,
    metaCursos: true,
    aniversarios: true,
  };
  const widgets = profile?.dashboardWidgets
    ? { ...defaultWidgets, ...profile.dashboardWidgets }
    : defaultWidgets;

  const currentMonthNum = new Date().getMonth() + 1; // 1-12
  const monthNamesPt = [
    "Janeiro",
    "Fevereiro",
    "Março",
    "Abril",
    "Maio",
    "Junho",
    "Julho",
    "Agosto",
    "Setembro",
    "Outubro",
    "Novembro",
    "Dezembro",
  ];
  const currentMonthName = monthNamesPt[currentMonthNum - 1];

  const currentDayNum = new Date().getDate();
  const checkIsToday = (dob: string) => {
    const parts = dob.split("-");
    if (parts.length !== 3) return false;
    return (
      parseInt(parts[2], 10) === currentDayNum &&
      parseInt(parts[1], 10) === currentMonthNum
    );
  };

  const birthdaysThisMonth = (users || [])
    .filter((u) => {
      if (u.blocked) return false;
      if (!u.dataNascimento) return false;
      const dateParts = u.dataNascimento.split("-");
      if (dateParts.length !== 3) return false;
      const birthMonth = parseInt(dateParts[1], 10);
      return birthMonth === currentMonthNum;
    })
    .sort((a, b) => {
      const dayA = parseInt(a.dataNascimento!.split("-")[2], 10);
      const dayB = parseInt(b.dataNascimento!.split("-")[2], 10);
      return dayA - dayB;
    });

  const today = new Date().toISOString().split("T")[0];
  const activePeriod = periodos.find(
    (p) => today >= p.inicioInscricao && today <= p.fimMatFin,
  );

  // Find meta for today, or find the latest meta as a fallback
  const todayEntry = metaDia.find((m) => m.data === today);
  const latestEntry =
    metaDia.length > 0
      ? [...metaDia].sort((a, b) => b.data.localeCompare(a.data))[0]
      : null;
  const activeMeta = todayEntry || latestEntry;

  const days = [
    "Segunda-feira",
    "Terça-feira",
    "Quarta-feira",
    "Quinta-feira",
    "Sexta-feira",
    "Sábado",
    "Domingo",
  ];

  const toggleWidget = async (
    key: keyof NonNullable<UserProfile["dashboardWidgets"]>,
  ) => {
    try {
      const newWidgets = { ...widgets, [key]: !widgets[key] };
      await updateDoc(doc(db, COLLECTIONS.USERS, profile.uid), {
        dashboardWidgets: newWidgets,
      });
      onToast("Preferências salvas!");
    } catch (err: any) {
      onToast("Erro ao salvar preferências.", "error");
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex justify-between items-center">
        <h2 className="text-2xl font-bold text-slate-800">Dashboard</h2>
        <div className="flex items-center space-x-4">
          <button
            onClick={() => setIsCustomizing(true)}
            className="flex items-center space-x-2 px-4 py-2 bg-white border border-slate-200 rounded-xl text-sm font-bold text-slate-600 hover:bg-slate-50 transition-all"
          >
            <Settings size={18} />
            <span>Personalizar</span>
          </button>
        </div>
      </div>

      {/* Android App Promotion Card */}
      {!isAppInstalled && (
        <div
          id="android-app-prompt-card"
          className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 p-6 rounded-3xl text-white shadow-xl relative overflow-hidden border border-slate-700/60 transition-all duration-300"
        >
          {/* Decorative design bubbles */}
          <div className="absolute -top-16 -right-16 w-32 h-32 bg-emerald-500/10 rounded-full blur-2xl"></div>
          <div className="absolute -bottom-16 -left-16 w-32 h-32 bg-blue-500/10 rounded-full blur-2xl"></div>

          <div className="flex flex-col md:flex-row md:items-center justify-between gap-6 relative z-10">
            <div className="flex items-start space-x-4">
              <div className="p-3.5 bg-slate-800/80 rounded-2xl border border-slate-700/80 flex items-center justify-center shadow-lg transform hover:scale-105 transition-all shrink-0">
                <img
                  src="/icon.svg"
                  alt="Gestão Oeste"
                  className="w-12 h-12 rounded-xl object-contain"
                />
              </div>
              <div className="space-y-1">
                <div className="flex items-center space-x-2">
                  <span className="px-2.5 py-0.5 bg-emerald-500/20 text-emerald-400 text-[10px] font-black uppercase tracking-wider rounded-md border border-emerald-500/30">
                    Instalação Android
                  </span>
                  <span className="text-xs text-slate-400 font-bold uppercase tracking-wider">
                    Suporte Offline Completo
                  </span>
                </div>
                <h3 className="text-xl font-black tracking-tight leading-none text-white">
                  Instalar Aplicativo Gestão Oeste no Android
                </h3>
                <p className="text-sm text-slate-300 max-w-2xl mt-1.5 leading-relaxed font-semibold">
                  Trabalhe de qualquer lugar! Faça pedidos de insumos e
                  visualize dados{" "}
                  <strong className="text-emerald-400 font-bold">
                    com ou sem internet
                  </strong>
                  . Ao voltar a ter conexão, o sistema sincroniza
                  automaticamente com o servidor.
                </p>
              </div>
            </div>

            <div className="flex items-center space-x-3 shrink-0 self-end md:self-center">
              <button
                onClick={handleInstallClick}
                className="flex items-center space-x-2 px-5 py-3 bg-emerald-500 hover:bg-emerald-600 active:transform active:scale-95 text-slate-950 font-extrabold text-sm rounded-xl shadow-lg shadow-emerald-500/20 transition-all cursor-pointer"
              >
                <Smartphone size={18} />
                <span>Instalar Aplicativo</span>
              </button>
              <button
                onClick={() => setShowInstallGuide(!showInstallGuide)}
                className="flex items-center space-x-2 px-4 py-3 bg-white/10 hover:bg-white/15 text-slate-100 font-bold text-sm rounded-xl transition-all cursor-pointer"
              >
                <span>Instruções</span>
              </button>
            </div>
          </div>

          {/* Expanded Step-by-Step Installation Guide */}
          {showInstallGuide && (
            <div className="mt-6 pt-6 border-t border-slate-700/60 grid grid-cols-1 md:grid-cols-3 gap-4 text-sm animate-fade-in">
              <div className="bg-slate-800/40 p-4 rounded-2xl border border-slate-700/40 space-y-2">
                <span className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-400 font-extrabold flex items-center justify-center text-xs">
                  1
                </span>
                <h4 className="font-extrabold text-white flex items-center gap-1.5 text-xs uppercase tracking-wider">
                  <Chrome size={14} className="text-emerald-400" /> No Google
                  Chrome
                </h4>
                <p className="text-slate-300 text-xs leading-relaxed font-semibold">
                  Abra este site no seu aparelho Android utilizando o navegador{" "}
                  <strong className="text-emerald-400 font-bold">
                    Google Chrome
                  </strong>
                  .
                </p>
              </div>

              <div className="bg-slate-800/40 p-4 rounded-2xl border border-slate-700/40 space-y-2">
                <span className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-400 font-extrabold flex items-center justify-center text-xs">
                  2
                </span>
                <h4 className="font-extrabold text-white flex items-center gap-1.5 text-xs uppercase tracking-wider">
                  <Smartphone size={14} className="text-emerald-400" /> Menu de
                  Opções
                </h4>
                <p className="text-slate-300 text-xs leading-relaxed font-semibold">
                  Toque nos{" "}
                  <strong className="text-white font-bold">
                    três pontinhos (⋮)
                  </strong>{" "}
                  localizados no canto superior direito do navegador Chrome.
                </p>
              </div>

              <div className="bg-slate-800/40 p-4 rounded-2xl border border-slate-700/40 space-y-2">
                <span className="w-6 h-6 rounded-full bg-emerald-500/20 text-emerald-400 font-extrabold flex items-center justify-center text-xs">
                  3
                </span>
                <h4 className="font-extrabold text-white flex items-center gap-1.5 text-xs uppercase tracking-wider">
                  <Download size={14} className="text-emerald-400" /> Instalar
                  App
                </h4>
                <p className="text-slate-300 text-xs leading-relaxed font-semibold">
                  Selecione{" "}
                  <strong className="text-emerald-400 font-bold">
                    "Instalar aplicativo"
                  </strong>{" "}
                  ou{" "}
                  <strong className="text-emerald-400 font-bold">
                    "Adicionar à tela de início"
                  </strong>
                  . Um atalho oficial será criado no seu telefone!
                </p>
              </div>

              <div className="col-span-1 md:col-span-3 flex justify-end mt-2 animate-fade-in">
                <button
                  onClick={() => setShowInstallGuide(false)}
                  className="px-3.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-xs text-slate-300 rounded-lg hover:text-white transition-all font-bold cursor-pointer border border-slate-700"
                >
                  Fechar Instruções
                </button>
              </div>
            </div>
          )}
        </div>
      )}

      {(() => {
        const todayDateObj = new Date(today + "T12:00:00Z");
        const dayOfWeek = todayDateObj.getUTCDay();
        const startOfWeek = new Date(todayDateObj);
        startOfWeek.setUTCDate(todayDateObj.getUTCDate() - dayOfWeek);
        const endOfWeek = new Date(startOfWeek);
        endOfWeek.setUTCDate(startOfWeek.getUTCDate() + 6);
        const startOfWeekStr = startOfWeek.toISOString().split("T")[0];
        const endOfWeekStr = endOfWeek.toISOString().split("T")[0];

        const thisWeekMetas = metaDia.filter(m => m.data >= startOfWeekStr && m.data <= endOfWeekStr);
        if (thisWeekMetas.length === 0) return null;

        const weekTotYTD = thisWeekMetas.reduce((acc, item) => acc + item.ytdPresencial + item.ytdSemipresencial + item.ytdDigital, 0);
        const weekTotReal = thisWeekMetas.reduce((acc, item) => acc + item.realizadoPresencial + item.realizadoSemipresencial + item.realizadoDigital, 0);
        const weekTotAA = thisWeekMetas.reduce((acc, item) => acc + item.aaPresencial + item.aaSemipresencial + item.aaDigital, 0);

        let statusText = "Abaixo da Meta";
        let statusColor = "bg-rose-50 text-rose-600 border-rose-100";
        if (weekTotYTD > 0 && weekTotReal > weekTotYTD) {
          statusText = "Meta Superada!";
          statusColor = "bg-emerald-50 text-emerald-600 border-emerald-100";
        } else if (weekTotYTD > 0 && weekTotReal === weekTotYTD) {
          statusText = "Meta Atingida";
          statusColor = "bg-blue-50 text-blue-600 border-blue-100";
        }

        return (
          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 space-y-4 mb-6">
            <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center">
              <div>
                <div className="flex items-center space-x-2 text-slate-900">
                  <Target size={20} className="text-indigo-600" />
                  <h3 className="text-lg font-bold">
                    Acompanhamento de Meta Semanal
                  </h3>
                </div>
                <p className="text-xs text-slate-400 font-medium mt-1">
                  Semana: <span className="font-bold">{new Date(startOfWeekStr + "T12:00:00Z").toLocaleDateString("pt-BR")}</span> a <span className="font-bold">{new Date(endOfWeekStr + "T12:00:00Z").toLocaleDateString("pt-BR")}</span>
                </p>
              </div>
              <span className={cn("px-3 py-1.5 rounded-full text-xs font-bold border mt-2 sm:mt-0", statusColor)}>
                {statusText}
              </span>
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4 pt-2">
              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-slate-500 mb-1">Ano Anterior (Semana)</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-slate-700">{weekTotAA}</span>
                </div>
              </div>
              <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-blue-600 mb-1">Meta (Semana)</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-blue-700">{weekTotYTD}</span>
                </div>
              </div>
              <div className="bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-emerald-600 mb-1">Realizado (Semana)</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-emerald-700">{weekTotReal}</span>
                </div>
              </div>
              <div className="bg-amber-50/50 p-4 rounded-2xl border border-amber-100/50">
                <p className="text-[10px] uppercase font-bold tracking-wider text-amber-600 mb-1">Atingimento</p>
                <div className="flex items-end gap-2">
                  <span className="text-2xl font-black text-amber-700">
                    {weekTotYTD > 0 ? ((weekTotReal / weekTotYTD) * 100).toFixed(1) : 0}%
                  </span>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {activeMeta && (
        <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100 space-y-4">
          <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center">
            <div>
              <div className="flex items-center space-x-2 text-slate-900">
                <Target size={20} className="text-blue-600" />
                <h3 className="text-lg font-bold">
                  Acompanhamento de Meta Diária
                </h3>
              </div>
              <p className="text-xs text-slate-400 font-medium mt-1">
                Referente ao dia:{" "}
                <span className="font-bold">
                  {new Date(activeMeta.data + "T00:00:00").toLocaleDateString(
                    "pt-BR",
                  )}
                </span>
                {activeMeta.data === today
                  ? " (Hoje)"
                  : " (Última meta registrada)"}
              </p>
            </div>

            {(() => {
              const totYTD =
                activeMeta.ytdPresencial +
                activeMeta.ytdSemipresencial +
                activeMeta.ytdDigital;
              const totReal =
                activeMeta.realizadoPresencial +
                activeMeta.realizadoSemipresencial +
                activeMeta.realizadoDigital;

              let statusText = "Abaixo da Meta";
              let statusColor = "bg-rose-50 text-rose-600 border-rose-100";
              if (totReal > totYTD) {
                statusText = "Meta Superada!";
                statusColor =
                  "bg-emerald-50 text-emerald-600 border-emerald-100";
              } else if (totReal === totYTD) {
                statusText = "Meta Atingida";
                statusColor = "bg-blue-50 text-blue-600 border-blue-100";
              }

              return (
                <span
                  className={cn(
                    "px-3 py-1.5 rounded-full text-xs font-bold border mt-2 sm:mt-0",
                    statusColor,
                  )}
                >
                  {statusText}
                </span>
              );
            })()}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-4 gap-4 pt-2">
            <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
              <span className="text-[10px] font-bold text-slate-400 uppercase">
                Boletos Necessários (YTD)
              </span>
              <span className="text-2xl font-black text-slate-800 mt-2">
                {activeMeta.ytdPresencial +
                  activeMeta.ytdSemipresencial +
                  activeMeta.ytdDigital}
              </span>
            </div>

            {(() => {
              const totYTD =
                activeMeta.ytdPresencial +
                activeMeta.ytdSemipresencial +
                activeMeta.ytdDigital;
              const totReal =
                activeMeta.realizadoPresencial +
                activeMeta.realizadoSemipresencial +
                activeMeta.realizadoDigital;

              let color = "text-rose-600";
              if (totReal > totYTD) color = "text-emerald-600";
              else if (totReal === totYTD) color = "text-blue-600";

              return (
                <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-400 uppercase">
                    Total Realizado
                  </span>
                  <span className={cn("text-2xl font-black mt-2", color)}>
                    {totReal}
                  </span>
                </div>
              );
            })()}

            <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
              <span className="text-[10px] font-bold text-slate-400 uppercase">
                Ano Anterior (A.A)
              </span>
              <span className="text-2xl font-black text-slate-500 mt-2">
                {activeMeta.aaPresencial +
                  activeMeta.aaSemipresencial +
                  activeMeta.aaDigital}
              </span>
            </div>

            {(() => {
              const totYTD =
                activeMeta.ytdPresencial +
                activeMeta.ytdSemipresencial +
                activeMeta.ytdDigital;
              const totReal =
                activeMeta.realizadoPresencial +
                activeMeta.realizadoSemipresencial +
                activeMeta.realizadoDigital;
              const pct = totYTD > 0 ? (totReal / totYTD) * 100 : 0;

              let pctBg = "bg-rose-50 text-rose-700";
              if (totReal > totYTD) pctBg = "bg-emerald-50 text-emerald-700";
              else if (totReal === totYTD) pctBg = "bg-blue-50 text-blue-700";

              return (
                <div className="bg-slate-50/50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
                  <span className="text-[10px] font-bold text-slate-400 uppercase">
                    Aproveitamento
                  </span>
                  <div className="flex items-baseline space-x-2 mt-2">
                    <span
                      className={cn(
                        "text-xl font-extrabold px-2.5 py-0.5 rounded-lg",
                        pctBg,
                      )}
                    >
                      {pct.toFixed(0)}%
                    </span>
                  </div>
                </div>
              );
            })()}
          </div>

          <div className="grid grid-cols-2 md:grid-cols-4 gap-4 pt-2 border-t border-slate-100">
            {[
              {
                label: "Presencial",
                ytd: activeMeta.ytdPresencial,
                real: activeMeta.realizadoPresencial,
                aa: activeMeta.aaPresencial,
                accent: "border-l-4 border-l-blue-500",
              },
              {
                label: "Semipresencial",
                ytd: activeMeta.ytdSemipresencial,
                real: activeMeta.realizadoSemipresencial,
                aa: activeMeta.aaSemipresencial,
                accent: "border-l-4 border-l-orange-500",
              },
              {
                label: "Digital",
                ytd: activeMeta.ytdDigital,
                real: activeMeta.realizadoDigital,
                aa: activeMeta.aaDigital,
                accent: "border-l-4 border-l-indigo-500",
              },
              {
                label: "Curso Técnico",
                ytd: activeMeta.ytdTecnico || 0,
                real: activeMeta.realizadoTecnico || 0,
                aa: activeMeta.aaTecnico || 0,
                accent: "border-l-4 border-l-emerald-500",
              },
            ].map((modal, idx) => {
              let color = "text-rose-600";
              if (modal.real > modal.ytd) color = "text-emerald-600";
              else if (modal.real === modal.ytd) color = "text-blue-600";

              return (
                <div
                  key={idx}
                  className={cn(
                    "bg-slate-50/30 p-3 rounded-xl border border-slate-100 flex justify-between items-center",
                    modal.accent,
                  )}
                >
                  <div>
                    <span className="text-xs font-bold text-slate-700">
                      {modal.label}
                    </span>
                    <span className="block text-[10px] text-slate-400">
                      Ano Ant: {modal.aa}
                    </span>
                  </div>
                  <div className="text-right">
                    <span className="text-[10px] font-bold text-slate-400 block uppercase">
                      Meta / Real
                    </span>
                    <span className="text-xs font-semibold text-slate-600">
                      {modal.ytd}
                    </span>
                    <span className="mx-1 text-slate-300">/</span>
                    <span className={cn("text-xs font-bold", color)}>
                      {modal.real}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Aniversariantes do Mês Widget */}
      {widgets.aniversarios !== false && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-rose-500 mb-6">
            <Cake size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Aniversariantes do Mês ({currentMonthName})
            </h3>
          </div>
          {birthdaysThisMonth.length > 0 ? (
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 xl:grid-cols-4 gap-4">
              {birthdaysThisMonth.map((u) => {
                const bday = parseInt(u.dataNascimento!.split("-")[2], 10);
                const isToday = checkIsToday(u.dataNascimento!);
                return (
                  <div
                    key={u.uid}
                    className={cn(
                      "p-4 rounded-2xl border transition-all flex items-center justify-between",
                      isToday
                        ? "bg-rose-50/50 border-rose-200 shadow-sm shadow-rose-50"
                        : "bg-slate-50/50 border-slate-100 hover:border-slate-200",
                    )}
                  >
                    <div className="flex items-center space-x-3 overflow-hidden">
                      <div
                        className={cn(
                          "w-10 h-10 rounded-full flex items-center justify-center font-bold text-sm shrink-0",
                          isToday
                            ? "bg-rose-600 text-white animate-bounce"
                            : "bg-blue-50 text-blue-600",
                        )}
                      >
                        {u.name.charAt(0).toUpperCase()}
                      </div>
                      <div className="overflow-hidden">
                        <h4 className="font-bold text-slate-800 text-sm truncate">
                          {u.name}
                        </h4>
                        <p className="text-[10px] text-slate-400 font-semibold truncate uppercase tracking-wider">
                          {u.role}
                        </p>
                      </div>
                    </div>
                    <div className="text-right shrink-0">
                      {isToday ? (
                        <span className="inline-block px-2 py-1 bg-amber-100 text-amber-800 text-[10px] font-black rounded-lg uppercase tracking-wide animate-pulse">
                          Hoje! 🎉
                        </span>
                      ) : (
                        <span className="text-xs font-black text-slate-500">
                          Dia {bday}
                        </span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          ) : (
            <div className="text-center py-8 bg-slate-50/50 border border-dashed border-slate-200 rounded-2xl">
              <Cake size={32} className="mx-auto text-slate-300 mb-2" />
              <p className="text-sm text-slate-400 font-semibold">
                Nenhum aniversariante registrado neste mês de {currentMonthName}
                .
              </p>
            </div>
          )}
        </section>
      )}

      {/* Meta SM Dashboard Card */}
      {widgets.metaSM && (metaSM && metaSM.length > 0) && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-indigo-600 mb-6">
            <Target size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Acompanhamento de Meta SM
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
            {[...metaSM].sort((a,b) => (b.semestre || "").localeCompare(a.semestre || "")).map(m => (
              <React.Fragment key={m.id}>
                <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100 flex flex-col justify-between">
                  <div>
                    <p className="text-xs font-bold text-slate-500 uppercase">Semestre {m.semestre}</p>
                    <p className="text-lg font-black text-slate-900">{m.realizado}</p>
                    <p className="text-xs text-slate-400">Realizado SM</p>
                  </div>
                  <div className="w-full text-right mt-2 pt-2 border-t border-slate-200">
                    <p className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                      Atualizado: {m.createdAt?.seconds ? new Date(m.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : m.createdAt ? new Date(m.createdAt).toLocaleDateString("pt-BR") : "-"}
                    </p>
                  </div>
                </div>
                <div className="bg-blue-50/50 p-4 rounded-2xl border border-blue-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP Meta Dia</p>
                  <p className="text-lg font-black">{(() => {
                    const gap = m.realizado - (m.metaDia || 0);
                    return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                  })()}</p>
                  <div className="mt-2 w-full bg-blue-100 rounded-full h-1.5">
                    <div className="bg-blue-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, (m.realizado / (m.metaDia || 1)) * 100)}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center mt-1">
                    <p className="text-xs text-blue-400">Meta: {m.metaDia}</p>
                    <p className="text-xs font-bold text-blue-500">{((m.realizado / (m.metaDia || 1)) * 100).toFixed(1)}%</p>
                  </div>
                </div>
                <div className="bg-emerald-50/50 p-4 rounded-2xl border border-emerald-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP A.A</p>
                  <p className="text-lg font-black">{(() => {
                    const gap = m.realizado - (m.metaAA || 0);
                    return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                  })()}</p>
                  <div className="mt-2 w-full bg-emerald-100 rounded-full h-1.5">
                    <div className="bg-emerald-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, (m.realizado / (m.metaAA || 1)) * 100)}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center mt-1">
                    <p className="text-xs text-emerald-400">Ano Anterior: {m.metaAA}</p>
                    <p className="text-xs font-bold text-emerald-500">{((m.realizado / (m.metaAA || 1)) * 100).toFixed(1)}%</p>
                  </div>
                </div>
                <div className="bg-purple-50/50 p-4 rounded-2xl border border-purple-100/50">
                  <p className="text-xs font-bold text-slate-500 uppercase">GAP Final</p>
                  <p className="text-lg font-black">{(() => {
                    const gap = m.realizado - (m.metaFinal || 0);
                    return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                  })()}</p>
                  <div className="mt-2 w-full bg-purple-100 rounded-full h-1.5">
                    <div className="bg-purple-500 h-1.5 rounded-full" style={{ width: `${Math.min(100, (m.realizado / (m.metaFinal || 1)) * 100)}%` }}></div>
                  </div>
                  <div className="flex justify-between items-center mt-1">
                    <p className="text-xs text-purple-400">Meta: {m.metaFinal}</p>
                    <p className="text-xs font-bold text-purple-500">{((m.realizado / (m.metaFinal || 1)) * 100).toFixed(1)}%</p>
                  </div>
                </div>
              </React.Fragment>
            ))}
          </div>
        </section>
      )}

      {/* Meta Cursos Dashboard Card */}
      {widgets.metaCursos && (metaCursos && metaCursos.length > 0) && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center space-x-2 text-indigo-600 mb-6">
            <Target size={24} />
            <h3 className="text-xl font-bold text-slate-900">
              Acompanhamento de Meta Cursos
            </h3>
          </div>
          <div className="grid grid-cols-1 xl:grid-cols-2 2xl:grid-cols-3 gap-6">
            {[...metaCursos].sort((a,b) => (b.semestre || "").localeCompare(a.semestre || "") || (a.curso || "").localeCompare(b.curso || "")).map(m => (
              <div key={m.id} className="bg-slate-50 border border-slate-200 rounded-2xl overflow-hidden shadow-sm">
                <div className="bg-indigo-600 p-4 flex justify-between items-center">
                  <h4 className="font-bold text-white text-sm uppercase tracking-wider">{m.curso}</h4>
                  <span className="bg-white/20 text-white text-[10px] font-bold px-2 py-1 rounded-full uppercase">{m.semestre}</span>
                </div>
                <div className="p-4">
                  <table className="w-full text-xs">
                    <thead>
                      <tr className="text-slate-400 font-bold uppercase tracking-tighter">
                        <th className="text-left pb-2">Indicador</th>
                        <th className="text-center pb-2">INSC</th>
                        <th className="text-center pb-2">FINANC</th>
                        <th className="text-center pb-2">ACAD</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {[
                        { label: "Meta Final", key: "metaFinal", color: "text-slate-600" },
                        { label: "Meta Dia", key: "metaDia", color: "text-slate-600" },
                        { label: "Ano Anterior", key: "metaAA", color: "text-slate-400" },
                        { label: "Realizado", key: "realizado", color: "text-emerald-600 font-bold" }
                      ].map((row, idx) => (
                        <tr key={idx} className="hover:bg-white/50 transition-colors">
                          <td className="py-2 font-semibold text-slate-500">{row.label}</td>
                          <td className={`py-2 text-center ${row.color}`}>{m.inscritos?.[row.key as keyof typeof m.inscritos] ?? m[row.key as keyof typeof m] ?? 0}</td>
                          <td className={`py-2 text-center ${row.color}`}>{m.financeiro?.[row.key as keyof typeof m.financeiro] ?? "-"}</td>
                          <td className={`py-2 text-center ${row.color}`}>{m.academico?.[row.key as keyof typeof m.academico] ?? "-"}</td>
                        </tr>
                      ))}
                      
                      {/* Calculated Gaps */}
                      {[
                        { label: "Gap Meta Dia", metaKey: "metaDia" },
                        { label: "Gap Ano Ant.", metaKey: "metaAA" },
                        { label: "Gap Final", metaKey: "metaFinal" }
                      ].map((row, idx) => (
                        <tr key={`calc-${idx}`} className="bg-slate-100/50">
                          <td className="py-1.5 font-bold text-[9px] text-slate-400 uppercase">{row.label}</td>
                          <td className="py-1.5 text-center text-[10px] font-bold">
                            {(() => {
                              const real = m.inscritos?.realizado ?? m.realizado ?? 0;
                              const meta = m.inscritos?.[row.metaKey as keyof typeof m.inscritos] ?? m[row.metaKey as keyof typeof m] ?? 0;
                              const gap = real - meta;
                              if (gap === 0 && !m.inscritos) return <span className="text-slate-600">0</span>;
                              return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                            })()}
                          </td>
                          <td className="py-1.5 text-center text-[10px] font-bold">
                            {(() => {
                              if (!m.financeiro) return <span className="text-slate-400">-</span>;
                              const real = m.financeiro.realizado || 0;
                              const meta = m.financeiro[row.metaKey as keyof typeof m.financeiro] || 0;
                              const gap = real - meta;
                              return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                            })()}
                          </td>
                          <td className="py-1.5 text-center text-[10px] font-bold">
                            {(() => {
                              if (!m.academico) return <span className="text-slate-400">-</span>;
                              const real = m.academico.realizado || 0;
                              const meta = m.academico[row.metaKey as keyof typeof m.academico] || 0;
                              const gap = real - meta;
                              return <span className={gap > 0 ? "text-emerald-600" : gap < 0 ? "text-rose-600" : "text-slate-600"}>{gap > 0 ? "+" : ""}{gap}</span>;
                            })()}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="bg-slate-100 p-2 text-right">
                  <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                    Atualizado: {m.createdAt?.seconds ? new Date(m.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : m.createdAt ? new Date(m.createdAt).toLocaleDateString("pt-BR") : "-"}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      
      {/* Bom Dia Captação (Complete - All cards) */}
      {widgets.bomDia && bomDia.filter((b) => !b.oculto).length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <div className="flex items-center space-x-2 text-emerald-600">
              <Sun size={24} />
              <h3 className="text-xl font-bold text-slate-900">
                Bom Dia Captação
              </h3>
            </div>
          </div>
          <div className="grid grid-cols-1 xl:grid-cols-2 2xl:grid-cols-3 gap-6">
            {bomDia
              .filter((b) => !b.oculto)
              .map((card) => (
                <div
                  key={card.id}
                  className="bg-slate-50 rounded-2xl border border-slate-100 overflow-hidden"
                >
                  <div className="bg-emerald-600 p-4">
                    <h4 className="font-bold text-white text-sm uppercase tracking-wider">
                      {card.titulo}
                    </h4>
                  </div>
                  <div className="p-4">
                    <table className="w-full text-xs">
                      <thead>
                        <tr className="text-slate-400 font-bold uppercase tracking-tighter">
                          <th className="text-left pb-2">Indicador</th>
                          <th className="text-center pb-2">INSC</th>
                          <th className="text-center pb-2">MAT FIN</th>
                          <th className="text-center pb-2">MAT ACAD</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-200">
                        {[
                          {
                            label: "Meta Final",
                            data: card.metaFinal,
                            color: "text-slate-600",
                          },
                          {
                            label: "Meta Dia",
                            data: card.metaDia,
                            color: "text-slate-600",
                          },
                          {
                            label: "Ano Anterior",
                            data: card.anoAnterior,
                            color: "text-slate-400",
                          },
                          {
                            label: "Real",
                            data: card.real,
                            color: "text-emerald-600 font-bold",
                          },
                        ].map((row, idx) => (
                          <tr
                            key={idx}
                            className="hover:bg-white/50 transition-colors"
                          >
                            <td className="py-2 font-semibold text-slate-500">
                              {row.label}
                            </td>
                            <td className={cn("py-2 text-center", row.color)}>
                              {row.data?.insc ?? 0}
                            </td>
                            <td className={cn("py-2 text-center", row.color)}>
                              {row.data?.matFin ?? 0}
                            </td>
                            <td className={cn("py-2 text-center", row.color)}>
                              {row.data?.matAcad ?? 0}
                            </td>
                          </tr>
                        ))}
                        {/* Calculated Rows */}
                        {[
                          {
                            label: "% Meta Dia",
                            calc: (m: keyof BomDiaMetrics) =>
                              card.metaDia && card.metaDia[m] > 0 && card.real
                                ? `${((card.real[m] / card.metaDia[m]) * 100).toFixed(0)}%`
                                : "0%",
                            color: "text-blue-600 font-bold",
                          },
                          {
                            label: "% Ano Ant.",
                            calc: (m: keyof BomDiaMetrics) =>
                              card.anoAnterior &&
                              card.anoAnterior[m] > 0 &&
                              card.real
                                ? `${((card.real[m] / card.anoAnterior[m]) * 100).toFixed(0)}%`
                                : "0%",
                            color: "text-slate-500 font-bold",
                          },
                          {
                            label: "Gap Meta Dia",
                            calc: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaDia) return 0;
                              const val = card.real[m] - card.metaDia[m];
                              return val > 0 ? `+${val}` : val;
                            },
                            color: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaDia) return "text-slate-600";
                              const val = card.real[m] - card.metaDia[m];
                              return val > 0 ? "text-emerald-600" : (val < 0 ? "text-rose-600" : "text-slate-600");
                            },
                          },
                          {
                            label: "Gap Ano Ant.",
                            calc: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.anoAnterior) return 0;
                              const val = card.real[m] - card.anoAnterior[m];
                              return val > 0 ? `+${val}` : val;
                            },
                            color: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.anoAnterior) return "text-slate-600";
                              const val = card.real[m] - card.anoAnterior[m];
                              return val > 0 ? "text-emerald-600" : (val < 0 ? "text-rose-600" : "text-slate-600");
                            },
                          },
                          {
                            label: "Gap Meta Final",
                            calc: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaFinal) return 0;
                              const val = card.real[m] - card.metaFinal[m];
                              return val > 0 ? `+${val}` : val;
                            },
                            color: (m: keyof BomDiaMetrics) => {
                              if (!card.real || !card.metaFinal) return "text-slate-600";
                              const val = card.real[m] - card.metaFinal[m];
                              return val > 0 ? "text-emerald-600" : (val < 0 ? "text-rose-600" : "text-slate-600");
                            },
                          },
                        ].map((row, idx) => (
                          <tr key={`calc-${idx}`} className="bg-slate-100/50">
                            <td className="py-1.5 font-bold text-[9px] text-slate-400 uppercase">
                              {row.label}
                            </td>
                            <td
                              className={cn(
                                "py-1.5 text-center text-[10px] font-bold",
                                typeof row.color === "function"
                                  ? row.color("insc")
                                  : row.color,
                              )}
                            >
                              {row.calc("insc")}
                            </td>
                            <td
                              className={cn(
                                "py-1.5 text-center text-[10px] font-bold",
                                typeof row.color === "function"
                                  ? row.color("matFin")
                                  : row.color,
                              )}
                            >
                              {row.calc("matFin")}
                            </td>
                            <td
                              className={cn(
                                "py-1.5 text-center text-[10px] font-bold",
                                typeof row.color === "function"
                                  ? row.color("matAcad")
                                  : row.color,
                              )}
                            >
                              {row.calc("matAcad")}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="bg-slate-100 p-2 text-right">
                    <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                      Atualizado: {card.createdAt?.seconds ? new Date(card.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : card.createdAt ? new Date(card.createdAt).toLocaleDateString("pt-BR") : "-"}
                    </span>
                  </div>
                </div>
              ))}
          </div>
        </section>
      )}

      {/* QG Ligações Widget */}
      {widgets.qgLigacoes !== false && qgLigacoes && qgLigacoes.length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-xl font-bold text-slate-900 flex items-center">
              <span className="bg-emerald-100 text-emerald-600 p-2 rounded-xl mr-3">
                <Phone size={20} />
              </span>
              QG Ligações
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
            {qgLigacoes.map((qg) => (
              <div
                key={qg.id}
                className="bg-slate-50 border border-slate-100 rounded-2xl p-4 flex flex-col justify-between items-start"
              >
                <div className="flex items-center space-x-2 text-emerald-600 mb-2 font-bold">
                  <Phone size={16} />
                  <span>{qg.nome}</span>
                </div>
                <div className="text-sm font-semibold text-slate-700">
                  {Array.isArray(qg.diaSemana)
                    ? qg.diaSemana.join(", ")
                    : qg.diaSemana}
                </div>
                <div className="text-xs text-slate-500 font-medium bg-emerald-100/50 px-2 py-1 rounded-md mt-2">
                  {qg.horario}
                </div>
                <div className="w-full text-right mt-2 pt-2 border-t border-slate-200/60">
                  <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                    Atualizado: {qg.createdAt?.seconds ? new Date(qg.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : qg.createdAt ? new Date(qg.createdAt).toLocaleDateString("pt-BR") : "-"}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </section>
      )}

      {/* Forecasts (Complete - All cards) */}
      {widgets.forecast && forecast.filter((f) => !f.oculto).length > 0 && (
        <section className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
          <div className="flex items-center justify-between mb-6">
            <h3 className="text-xl font-bold text-slate-900">
              Forecasts de Captação
            </h3>
            <TrendingUp size={24} className="text-blue-600" />
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-6">
            {[...forecast]
              .filter((f) => !f.oculto)
              .sort((a, b) => a.nome.localeCompare(b.nome))
              .map((f) => {
                const percFech =
                  f.metaFechamento > 0
                    ? ((f.realizado / f.metaFechamento) * 100).toFixed(1)
                    : "0";
                const gapFech = f.realizado - f.metaFechamento;
                const dataFim = new Date(f.dataFim);
                const hoje = new Date();
                const diffTime = dataFim.getTime() - hoje.getTime();
                const diasRestantes = Math.max(
                  1,
                  Math.ceil(diffTime / (1000 * 60 * 60 * 24)),
                );
                const pacing = (Math.abs(gapFech) / diasRestantes).toFixed(1);

                return (
                  <div
                    key={f.id}
                    className="bg-slate-50 p-5 rounded-2xl border border-slate-100"
                  >
                    <div className="flex justify-between items-start mb-4">
                      <div>
                        <h4 className="font-bold text-slate-900">{f.nome}</h4>
                        <p className="text-[10px] text-slate-500 font-medium">
                          Até{" "}
                          {f.dataFim
                            .split("T")[0]
                            .split("-")
                            .reverse()
                            .join("/")}
                        </p>
                      </div>
                      <span
                        className={`text-[10px] font-bold px-2 py-1 rounded-full ${Number(percFech) >= 100 ? "bg-emerald-100 text-emerald-600" : "bg-blue-100 text-blue-600"}`}
                      >
                        {percFech}%
                      </span>
                    </div>

                    <div className="grid grid-cols-2 gap-4 mb-4">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Realizado
                        </p>
                        <p className="text-lg font-bold text-emerald-600">
                          {f.realizado || 0}
                        </p>
                      </div>
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Meta
                        </p>
                        <p className="text-lg font-bold text-slate-700">
                          {f.metaFechamento || 0}
                        </p>
                      </div>
                    </div>

                    <div className="space-y-3 pt-4 border-t border-slate-200/60">
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider pl-1 border-l-2 border-slate-400">
                          Meta Dia YTD
                        </span>
                        <span className="text-xs font-bold text-slate-700">
                          {f.metaDiaYTD || 0}
                        </span>
                      </div>
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider pl-1 border-l-2 border-rose-400">
                          Gap Fechamento
                        </span>
                        <span
                          className={`text-xs font-bold ${gapFech > 0 ? "text-emerald-600" : gapFech < 0 ? "text-rose-600" : "text-slate-600"}`}
                        >
                          {gapFech > 0 ? "+" : ""}
                          {gapFech}
                        </span>
                      </div>
                      <div className="flex justify-between items-center">
                        <span className="text-xs font-bold text-slate-500 uppercase tracking-wider pl-1 border-l-2 border-blue-400">
                          Pacing (por dia)
                        </span>
                        <span className="text-xs font-bold text-blue-600">
                          {pacing}
                        </span>
                      </div>
                      <div className="flex justify-between items-center bg-slate-200/50 p-2 rounded-lg mt-2">
                        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                          Dias Restantes
                        </span>
                        <span className="text-xs font-bold text-slate-800">
                          {diasRestantes}
                        </span>
                      </div>
                      <div className="w-full text-right mt-2 pt-2 border-t border-slate-200/60">
                        <span className="text-[9px] text-slate-400 font-medium uppercase tracking-wider">
                          Atualizado: {f.createdAt?.seconds ? new Date(f.createdAt.seconds * 1000).toLocaleDateString("pt-BR") : f.createdAt ? new Date(f.createdAt).toLocaleDateString("pt-BR") : "-"}
                        </span>
                      </div>
                    </div>
                  </div>
                );
              })}
          </div>
        </section>
      )}

      {widgets.periodo && periodos.length > 0 && (
        <section>
          <div className="flex items-center justify-between mb-4">
            <h3 className="text-xl font-bold text-slate-900">
              Períodos da Captação
            </h3>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {periodos.map((p) => {
              const isActive =
                today >= p.inicioInscricao && today <= p.fimMatFin;
              return (
                <div
                  key={p.id}
                  className={cn(
                    "bg-white p-5 rounded-3xl shadow-sm border transition-all",
                    isActive
                      ? "border-blue-500 ring-4 ring-blue-50"
                      : "border-slate-100",
                  )}
                >
                  <div className="flex items-center justify-between mb-4">
                    <div className="flex items-center space-x-3">
                      <div
                        className={cn(
                          "p-2 rounded-xl",
                          isActive
                            ? "bg-blue-600 text-white"
                            : "bg-blue-100 text-blue-600",
                        )}
                      >
                        <Calendar size={20} />
                      </div>
                      <h4 className="font-bold text-slate-900">{p.nome}</h4>
                    </div>
                    {isActive && (
                      <span className="px-2 py-1 bg-emerald-100 text-emerald-600 text-[10px] font-bold rounded-full uppercase">
                        Ativo
                      </span>
                    )}
                  </div>
                  <div className="space-y-4">
                    <div className="flex justify-between items-center p-3 bg-slate-50 rounded-2xl">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Inscrição
                        </p>
                        <p className="text-xs font-bold text-slate-700">
                          {formatLocalDateString(p.inicioInscricao)} -{" "}
                          {formatLocalDateString(p.fimInscricao)}
                        </p>
                      </div>
                      <span className="text-xs font-bold text-blue-600">
                        {getWorkingDaysBetween(
                          p.inicioInscricao,
                          p.fimInscricao,
                        )}{" "}
                        dias
                      </span>
                    </div>
                    <div className="flex justify-between items-center p-3 bg-slate-50 rounded-2xl">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Mat Fin
                        </p>
                        <p className="text-xs font-bold text-slate-700">
                          {formatLocalDateString(p.inicioMatFin)} -{" "}
                          {formatLocalDateString(p.fimMatFin)}
                        </p>
                      </div>
                      <div className="text-right">
                        <span className="text-xs font-bold text-blue-600 block">
                          {getWorkingDaysBetween(p.inicioMatFin, p.fimMatFin)}{" "}
                          dias úteis
                        </span>
                        <span className="text-[10px] font-bold text-slate-500 block">
                          {getWorkingDaysRemaining(p.fimMatFin)} dias restantes
                        </span>
                      </div>
                    </div>
                    <div className="flex justify-between items-center p-3 bg-slate-50 rounded-2xl">
                      <div>
                        <p className="text-[10px] font-bold text-slate-400 uppercase">
                          Mat Acad
                        </p>
                        <p className="text-xs font-bold text-slate-700">
                          {formatLocalDateString(p.inicioMatAcad)} -{" "}
                          {formatLocalDateString(p.fimMatAcad)}
                        </p>
                      </div>
                      <span className="text-xs font-bold text-blue-600">
                        {getWorkingDaysBetween(p.inicioMatAcad, p.fimMatAcad)}{" "}
                        dias
                      </span>
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {widgets.links && (
        <section>
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between mb-4 gap-4">
            <h3 className="text-xl font-bold text-slate-900">Links Úteis</h3>
            <div className="flex items-center gap-2 w-full sm:w-auto">
              <select
                value={linksFilterLocal}
                onChange={(e) => setLinksFilterLocal(e.target.value)}
                className="px-4 py-2 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm appearance-none"
              >
                <option value="">Todos os Locais</option>
                {Array.from(new Set(links.map(l => l.local).filter(Boolean))).map(local => (
                  <option key={local} value={local}>{local}</option>
                ))}
              </select>
              <div className="relative w-full sm:w-64">
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400">
                  <Search size={18} />
                </span>
                <input
                  type="text"
                  placeholder="Buscar links..."
                  value={linksSearchTerm}
                  onChange={(e) => setLinksSearchTerm(e.target.value)}
                  className="w-full pl-10 pr-4 py-2 bg-white border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none text-sm"
                />
              </div>
            </div>
          </div>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            {links
              .filter(link => 
                (!linksFilterLocal || link.local === linksFilterLocal) &&
                (!linksSearchTerm || 
                  link.nome.toLowerCase().includes(linksSearchTerm.toLowerCase()) || 
                  link.url.toLowerCase().includes(linksSearchTerm.toLowerCase()) || 
                  (link.local && link.local.toLowerCase().includes(linksSearchTerm.toLowerCase()))
                )
              )
              .map((link) => (
              <a
                key={link.id}
                href={link.url}
                target="_blank"
                rel="noopener noreferrer"
                className="bg-white p-4 rounded-2xl shadow-sm border border-slate-100 flex items-center space-x-3 hover:border-blue-200 hover:bg-blue-50/30 transition-all group"
              >
                <div className="p-2 bg-blue-100 text-blue-600 rounded-lg group-hover:bg-blue-600 group-hover:text-white transition-all shrink-0">
                  <ExternalLink size={18} />
                </div>
                <div className="flex flex-col overflow-hidden">
                  <span className="font-bold text-slate-700 truncate">
                    {link.nome}
                  </span>
                  {link.local && (
                    <span className="text-[10px] text-slate-400 font-medium truncate uppercase tracking-wider">
                      {link.local}
                    </span>
                  )}
                </div>
              </a>
            ))}
            {links.length === 0 && (
              <p className="text-slate-400 text-sm italic">
                Nenhum link cadastrado.
              </p>
            )}
          </div>
        </section>
      )}

      {widgets.planner && (
        <section>
          <h3 className="text-xl font-bold text-slate-900 mb-4">
            Planner da Semana
          </h3>
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 xl:grid-cols-7 gap-4">
            {days.map((day) => {
              const tasks = planner.filter((t) => t.dayOfWeek === day);
              return (
                <div
                  key={day}
                  className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden flex flex-col"
                >
                  <div className="bg-slate-50 px-4 py-2 border-b border-slate-100">
                    <span className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                      {day.split("-")[0]}
                    </span>
                  </div>
                  <div className="p-4 flex-1 space-y-2">
                    {tasks.length > 0 ? (
                      tasks.map((task) => (
                        <div
                          key={task.id}
                          className="p-2 bg-blue-50 border-l-4 border-blue-600 rounded-r-lg"
                        >
                          <p className="text-xs font-bold text-blue-900">
                            {task.atendenteName}
                          </p>
                          <p className="text-[10px] text-blue-600 font-medium">
                            {task.baseName}
                          </p>
                        </div>
                      ))
                    ) : (
                      <div className="h-full flex items-center justify-center">
                        <p className="text-xs text-slate-300 italic">Folga</p>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </section>
      )}

      {/* Customization Modal */}
      <AnimatePresence>
        {isCustomizing && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
            <motion.div
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
              exit={{ opacity: 0, scale: 0.95 }}
              className="bg-white rounded-2xl shadow-xl w-full max-w-md overflow-hidden"
            >
              <div className="p-6 border-b border-slate-100 flex justify-between items-center">
                <h3 className="text-xl font-bold text-slate-900">
                  Personalizar Dashboard
                </h3>
                <button
                  onClick={() => setIsCustomizing(false)}
                  className="text-slate-400 hover:text-slate-600"
                >
                  <X size={24} />
                </button>
              </div>
              <div className="p-6 space-y-4">
                <p className="text-sm text-slate-500 mb-4">
                  Escolha quais blocos você deseja visualizar na sua tela
                  principal.
                </p>

                {[
                  {
                    id: "periodo",
                    label: "Períodos da Captação",
                    icon: Calendar,
                  },
                  { id: "bomDia", label: "Bom Dia Captação", icon: Sun },
                  { id: "forecast", label: "Forecasts", icon: TrendingUp },
                  { id: "links", label: "Links Úteis", icon: ExternalLink },
                  { id: "planner", label: "Planner da Semana", icon: Calendar },
                  { id: "qgLigacoes", label: "QG Ligações", icon: Phone },
                  {
                    id: "aniversarios",
                    label: "Aniversariantes do Mês",
                    icon: Cake,
                  },
                ].map((item) => (
                  <button
                    key={item.id}
                    onClick={() => toggleWidget(item.id as any)}
                    className={cn(
                      "w-full flex items-center justify-between p-4 rounded-xl border transition-all",
                      widgets[item.id as keyof typeof widgets]
                        ? "bg-blue-50 border-blue-200 text-blue-900"
                        : "bg-white border-slate-100 text-slate-500",
                    )}
                  >
                    <div className="flex items-center space-x-3">
                      <item.icon size={20} />
                      <span className="font-bold">{item.label}</span>
                    </div>
                    <div
                      className={cn(
                        "w-10 h-6 rounded-full relative transition-all",
                        widgets[item.id as keyof typeof widgets]
                          ? "bg-blue-600"
                          : "bg-slate-200",
                      )}
                    >
                      <div
                        className={cn(
                          "absolute top-1 w-4 h-4 bg-white rounded-full transition-all",
                          widgets[item.id as keyof typeof widgets]
                            ? "left-5"
                            : "left-1",
                        )}
                      />
                    </div>
                  </button>
                ))}
              </div>
              <div className="p-6 bg-slate-50 border-t border-slate-100">
                <button
                  onClick={() => setIsCustomizing(false)}
                  className="w-full bg-slate-900 text-white font-bold py-3 rounded-xl hover:bg-slate-800 transition-all"
                >
                  Concluído
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function CadastroView({
  onToast,
  profile,
  calendarioAcoes = [],
  uniqueUnidades = [],
}: {
  onToast: (m: string, t?: "success" | "error") => void;
  profile: UserProfile;
  calendarioAcoes?: CalendarioAcao[];
  uniqueUnidades?: string[];
}) {
  const handleContatoViaSales = async (contact: any, origem: string) => {
    try {
      await addDoc(collection(db, COLLECTIONS.SALES_CONTACTS), {
        contactId: contact.id,
        nome: contact.nome,
        telefone: contact.telefone,
        curso: contact.cursoInteresse || contact.curso || "Não informado",
        origem,
        createdAt: serverTimestamp(),
      });
      onToast("Contato via Sales registrado com sucesso!", "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar Contato via Sales.", "error");
    }
  };
  
  const [formData, setFormData] = useState({
    acao: "",
    acaoId: "",
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    cursoInteresse: "",
  });
  const [loading, setLoading] = useState(false);
  const [activeForm, setActiveForm] = useState<"lead" | "promotor">("lead");
  const [promotorData, setPromotorData] = useState({
    nome: "",
    email: "",
    cpf: "",
    dataNascimento: "",
    phone: "",
    chavePix: "",
    unidade: "",
  });

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    // Duplicate check
    const cleanCpf = formData.cpf.replace(/\D/g, "");
    const cleanTelefone = formData.telefone.replace(/\D/g, "");

    if (cleanCpf) {
      const qCpf = query(
        collection(db, COLLECTIONS.LEADS),
        where("cpf", "==", cleanCpf),
      );
      const snapCpf = await getDocs(qCpf);
      if (!snapCpf.empty) {
        onToast(
          "Atenção: Este CPF já possui um lead cadastrado no sistema.",
          "error",
        );
        return;
      }
    } else if (cleanTelefone) {
      const qTel = query(
        collection(db, COLLECTIONS.LEADS),
        where("telefone", "==", cleanTelefone),
      );
      const snapTel = await getDocs(qTel);
      if (!snapTel.empty) {
        onToast(
          "Atenção: Este Telefone já possui um lead cadastrado no sistema.",
          "error",
        );
        return;
      }
    }

    setLoading(true);
    try {
      const newLeadData: any = {
        ...formData,
        cpf: cleanCpf,
        telefone: cleanTelefone,
        converted: false,
        createdAt: serverTimestamp(),
        promotorId: profile.uid,
        promotorName: profile.name,
        promotorRole: profile.role,
        unidade: profile.unidade || "",
        servidor: profile.servidor || "principal",
      };

      if (profile.linkadoA) {
        newLeadData.linkadoA = profile.linkadoA;
      }

      await addDoc(collection(db, COLLECTIONS.LEADS), newLeadData);

      if (newLeadData.acaoId && newLeadData.acaoId !== "manual") {
        try {
          const qLeads = query(
            collection(db, COLLECTIONS.LEADS),
            where("acaoId", "==", newLeadData.acaoId),
          );
          const snapLeads = await getDocs(qLeads);
          await updateDoc(
            doc(db, COLLECTIONS.CALENDARIO_ACOES, newLeadData.acaoId),
            {
              leadsFeitos: snapLeads.size,
            },
          );
        } catch (error) {
          console.error("Error auto-updating action leadsCount:", error);
        }
      }

      onToast("Lead cadastrado com sucesso!");
      setFormData({
        acao: "",
        acaoId: "",
        nome: "",
        telefone: "",
        email: "",
        cpf: "",
        cursoInteresse: "",
      });
    } catch (err: any) {
      onToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  const handlePromotorSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    const cleanCpf = promotorData.cpf.replace(/\D/g, "");
    const cleanPhone = promotorData.phone.replace(/\D/g, "");
    const cleanEmail = promotorData.email.trim();

    if (!promotorData.nome || !cleanEmail || !cleanPhone) {
      onToast(
        "Por favor, preencha todos os campos obrigatórios (Nome, Email e Telefone).",
        "error",
      );
      return;
    }

    if (!cleanEmail.includes("@")) {
      onToast("Formato de email inválido.", "error");
      return;
    }

    setLoading(true);
    try {
      // 1. Create promoter in Auth with standard base password using secondaryAuth
      const userCredential = await createUserWithEmailAndPassword(
        secondaryAuth,
        cleanEmail,
        "123456",
      );
      await updateProfile(userCredential.user, {
        displayName: `${promotorData.nome}|comercial`,
      });
      const newUid = userCredential.user.uid;

      // 2. Create profile matching promoter/rua rules
      const profileData: any = {
        uid: newUid,
        name: promotorData.nome,
        email: cleanEmail,
        cpf: cleanCpf,
        dataNascimento: promotorData.dataNascimento,
        role: ROLES.PROMOTOR_RUA, // 'Promotor/rua'
        servidor: "comercial", // specified for commercial
        phone: cleanPhone,
        unidade: promotorData.unidade,
        chavePix: promotorData.chavePix,
        blocked: false,
        mustChangePassword: true,
        linkadoA: profile.uid, // linked to the creator FDV
        createdAt: serverTimestamp(),
        updatedAt: serverTimestamp(),
      };

      // 3. Save profile document
      await setDoc(doc(db, COLLECTIONS.USERS, newUid), profileData);

      // 4. Sign out from secondary auth to avoid trace
      await signOut(secondaryAuth);

      onToast(
        "Promotor/rua cadastrado com sucesso! Senha padrão: 123456",
        "success",
      );
      setPromotorData({
        nome: "",
        email: "",
        cpf: "",
        dataNascimento: "",
        phone: "",
        chavePix: "",
        unidade: "",
      });
      setActiveForm("lead");
    } catch (err: any) {
      console.error("Auth error details (Promoter Registration):", err);
      let errorMsg = err.message;
      if (
        err.code === "auth/email-already-in-use" ||
        err.message?.includes("email-already-in-use")
      ) {
        errorMsg = "Este email já está em uso.";
      } else if (
        err.code === "auth/weak-password" ||
        err.message?.includes("weak-password")
      ) {
        errorMsg =
          "A senha de cadastro padrão deve conter pelo menos 6 caracteres.";
      } else if (
        err.code === "auth/invalid-email" ||
        err.message?.includes("invalid-email")
      ) {
        errorMsg = "Endereço de email inválido.";
      }
      onToast(`Erro ao criar promotor: ${errorMsg}`, "error");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="max-w-2xl mx-auto">
      <div className="bg-white p-8 rounded-3xl shadow-sm border border-slate-100">
        {profile?.role === ROLES.FDV_COMERCIAL && (
          <div className="flex space-x-2 bg-slate-50 p-1.5 rounded-2xl mb-6 border border-slate-100">
            <button
              type="button"
              onClick={() => setActiveForm("lead")}
              className={`flex-1 py-3 text-xs font-bold rounded-xl transition-all cursor-pointer flex items-center justify-center space-x-2 ${
                activeForm === "lead"
                  ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20"
                  : "text-slate-400 hover:text-slate-700"
              }`}
            >
              <UserPlus size={16} />
              <span>Cadastrar Novo Lead</span>
            </button>
            <button
              type="button"
              onClick={() => setActiveForm("promotor")}
              className={`flex-1 py-3 text-xs font-bold rounded-xl transition-all cursor-pointer flex items-center justify-center space-x-2 ${
                activeForm === "promotor"
                  ? "bg-gradient-to-r from-sky-500 to-blue-600 text-white shadow shadow-sky-500/20"
                  : "text-slate-400 hover:text-slate-700"
              }`}
            >
              <Users size={16} />
              <span>Cadastrar Promotor de Rua</span>
            </button>
          </div>
        )}

        {activeForm === "lead" ? (
          <>
            <h3 className="text-2xl font-bold text-slate-900 mb-6">
              Cadastrar Novo Lead
            </h3>
            <form onSubmit={handleSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="md:col-span-2 space-y-2">
                  <label className="block text-sm font-bold text-slate-700">
                    Ação / Origem
                  </label>
                  {calendarioAcoes && calendarioAcoes.length > 0 ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <span className="block text-xs font-semibold text-slate-500 mb-1">
                          Selecionar do Calendário
                        </span>
                        <select
                          value={formData.acaoId || ""}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (val === "manual") {
                              setFormData({
                                ...formData,
                                acaoId: "manual",
                                acao: "",
                              });
                            } else {
                              const matched = calendarioAcoes.find(
                                (a) => a.id === val,
                              );
                              setFormData({
                                ...formData,
                                acaoId: val,
                                acao: matched ? matched.nome : "",
                              });
                            }
                          }}
                          className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm bg-white"
                        >
                          <option value="">Selecione...</option>
                          {calendarioAcoes.map((act) => (
                            <option key={act.id} value={act.id}>
                              {act.nome} ({act.dataInicio})
                            </option>
                          ))}
                          <option value="manual">
                            Outro (Digitar manualmente)
                          </option>
                        </select>
                      </div>
                      {(formData.acaoId === "manual" || !formData.acaoId) && (
                        <div>
                          <span className="block text-xs font-semibold text-slate-500 mb-1">
                            Digitar Nome da Ação/Origem
                          </span>
                          <input
                            type="text"
                            required={
                              !formData.acaoId || formData.acaoId === "manual"
                            }
                            value={formData.acao}
                            onChange={(e) =>
                              setFormData({ ...formData, acao: e.target.value })
                            }
                            className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm"
                            placeholder="Ex: Facebook, Panfletagem, etc."
                          />
                        </div>
                      )}
                    </div>
                  ) : (
                    <input
                      type="text"
                      required
                      value={formData.acao}
                      onChange={(e) =>
                        setFormData({ ...formData, acao: e.target.value })
                      }
                      className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                      placeholder="Ex: Evento Junino, Facebook, etc."
                    />
                  )}
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Nome do Candidato
                  </label>
                  <input
                    type="text"
                    required
                    value={formData.nome}
                    onChange={(e) =>
                      setFormData({ ...formData, nome: e.target.value })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="Nome completo"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Telefone (WhatsApp)
                  </label>
                  <input
                    type="tel"
                    required
                    value={formData.telefone}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        telefone: formatPhone(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="DDD + Número"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    E-mail (Opcional)
                  </label>
                  <input
                    type="email"
                    value={formData.email}
                    onChange={(e) =>
                      setFormData({ ...formData, email: e.target.value })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="exemplo@gmail.com"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    CPF (Opcional)
                  </label>
                  <input
                    type="text"
                    value={formData.cpf}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        cpf: formatCPF(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="000.000.000-00"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Curso de Interesse
                  </label>
                  <input
                    type="text"
                    value={formData.cursoInteresse}
                    onChange={(e) =>
                      setFormData({
                        ...formData,
                        cursoInteresse: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="Ex: Administração, Direito..."
                  />
                </div>
              </div>
              <button
                type="submit"
                disabled={loading}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-200 transition-all disabled:opacity-50 flex items-center justify-center space-x-2"
              >
                <Plus size={20} />
                <span>{loading ? "Salvando..." : "Salvar Lead"}</span>
              </button>
            </form>
          </>
        ) : (
          <>
            <h3 className="text-2xl font-bold text-slate-900 mb-2">
              Cadastrar Promotor de Rua
            </h3>
            <p className="text-xs text-slate-500 mb-6 font-medium">
              Os promotores cadastrados por você ficarão automaticamente
              vinculados ao seu perfil de FDV e herdarão todas as regras de
              visualização do sistema.
            </p>

            <form onSubmit={handlePromotorSubmit} className="space-y-6">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Nome Completo *
                  </label>
                  <input
                    type="text"
                    required
                    value={promotorData.nome}
                    onChange={(e) =>
                      setPromotorData({ ...promotorData, nome: e.target.value })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="Nome completo do promotor"
                  />
                </div>
                <div className="md:col-span-2">
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Email (Google institucional ou pessoal) *
                  </label>
                  <input
                    type="email"
                    required
                    value={promotorData.email}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        email: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="exemplo@gmail.com"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Telefone / WhatsApp *
                  </label>
                  <input
                    type="tel"
                    required
                    value={promotorData.phone}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        phone: formatPhone(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="(00) 00000-0000"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    CPF (Opcional)
                  </label>
                  <input
                    type="text"
                    value={promotorData.cpf}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        cpf: formatCPF(e.target.value),
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="000.000.000-00"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Data de Nascimento (Opcional)
                  </label>
                  <input
                    type="date"
                    value={promotorData.dataNascimento}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        dataNascimento: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm"
                  />
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Unidade *
                  </label>
                  <select
                    required
                    value={promotorData.unidade}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        unidade: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all text-sm bg-white"
                  >
                    <option value="">Selecione uma unidade</option>
                    {uniqueUnidades.map((u) => (
                      <option key={u} value={u}>
                        {u}
                      </option>
                    ))}
                  </select>
                </div>
                <div>
                  <label className="block text-sm font-bold text-slate-700 mb-1">
                    Chave PIX (Opcional)
                  </label>
                  <input
                    type="text"
                    value={promotorData.chavePix}
                    onChange={(e) =>
                      setPromotorData({
                        ...promotorData,
                        chavePix: e.target.value,
                      })
                    }
                    className="w-full px-4 py-3 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none transition-all"
                    placeholder="CPF, E-mail, Telefone ou Aleatória"
                  />
                </div>
              </div>
              <button
                type="submit"
                disabled={loading}
                className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-4 rounded-2xl shadow-lg shadow-blue-200 transition-all disabled:opacity-50 flex items-center justify-center space-x-2"
              >
                <Plus size={20} />
                <span>
                  {loading ? "Cadastrando..." : "Cadastrar Promotor de Rua"}
                </span>
              </button>
            </form>
          </>
        )}
      </div>
    </div>
  );
}

function HistoricoView({
  leads,
  profile,
  onToast,
  users,
  whatsappMessages,
  botConfig,
  onSendBot,
  onMassSendBot,
  gap,
  basesRenovacao,
  calendarioAcoes = [],
  pedidosCursos = [],
}: {
  leads: Lead[];
  profile: UserProfile;
  onToast: (m: string, t?: "success" | "error") => void;
  users: UserProfile[];
  whatsappMessages: WhatsAppMessage[];
  botConfig: BotConfig;
  onSendBot: (tel: string, msg: string, contactName?: string) => void;
  onMassSendBot: (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => void;
  gap: GapEntry[];
  basesRenovacao: BaseEntry[];
  calendarioAcoes?: CalendarioAcao[];
  pedidosCursos?: PedidoCursoEntry[];
}) {
  const [searchTerm, setSearchTerm] = useState("");
  const [courseFilter, setCourseFilter] = useState("");
  const [baseFilter, setBaseFilter] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState("");
  const [promotorFilter, setPromotorFilter] = useState("");
  const [selectorOpen, setSelectorOpen] = useState(false);
  const [selectedLead, setSelectedLead] = useState<Lead | null>(null);
  const [selectedEntries, setSelectedEntries] = useState<string[]>([]);
  const [massSelectorOpen, setMassSelectorOpen] = useState(false);
  const [isAddMsgModalOpen, setIsAddMsgModalOpen] = useState(false);
  const [newMsgData, setNewMsgData] = useState({ modelName: "", texto: "" });
  const [msgLoading, setMsgLoading] = useState(false);
  const [invalidLeadIds, setInvalidLeadIds] = useState<Set<string>>(new Set());
  const [blockedFilter, setBlockedFilter] = useState<
    "all" | "blocked" | "unblocked"
  >("all");

  const [editModalOpen, setEditModalOpen] = useState(false);
  const [editingLead, setEditingLead] = useState<Lead | null>(null);
  const [historicoSubTab, setHistoricoSubTab] = useState<
    "dashboard" | "lista" | "pedidos_cursos"
  >("dashboard");

  const [editFormData, setEditFormData] = useState({
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    cursoInteresse: "",
    acao: "",
    acaoId: "",
  });

  const handleVerificacao = () => {
    const invalidIds = new Set<string>();
    leads.forEach((lead) => {
      let match = false;

      if (
        gap.some(
          (g) =>
            (g.cpf &&
              lead.cpf &&
              g.cpf.replace(/\D/g, "") === lead.cpf.replace(/\D/g, "")) ||
            (g.telefone &&
              lead.telefone &&
              g.telefone.replace(/\D/g, "") ===
                lead.telefone.replace(/\D/g, "")) ||
            g.nome.toLowerCase().trim() === lead.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (
        !match &&
        basesRenovacao.some(
          (b) =>
            (b.cpf &&
              lead.cpf &&
              b.cpf.replace(/\D/g, "") === lead.cpf.replace(/\D/g, "")) ||
            (b.telefone &&
              lead.telefone &&
              b.telefone.replace(/\D/g, "") ===
                lead.telefone.replace(/\D/g, "")) ||
            b.nome.toLowerCase().trim() === lead.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (match) {
        invalidIds.add(lead.id);
      }
    });
    setInvalidLeadIds(invalidIds);
    onToast(
      `Verificação concluída: ${invalidIds.size} leads já estão cadastrados em GAP/Base Líquida.`,
      "success",
    );
  };

  const uniqueCursos = useMemo(() => {
    return Array.from(
      new Set(leads.map((l) => l.cursoInteresse).filter(Boolean)),
    ).sort();
  }, [leads]);

  const uniqueBases = useMemo(() => {
    return Array.from(new Set(leads.map((l) => l.acao).filter(Boolean))).sort();
  }, [leads]);

  const uniqueStatuses = [
    "Pendente",
    "Sem retorno",
    "Interessado",
    "Não Interessado",
    "Convertido",
    "Contato via Sales",
  ];

  const uniquePromotores = useMemo(() => {
    return Array.from(
      new Set(leads.map((l) => l.promotorName).filter(Boolean)),
    ).sort();
  }, [leads]);

  const isAdmin = [
    ROLES.ADMIN_MASTER,
    ROLES.LIDER_FDV,
    ROLES.GESTOR_COMERCIAL,
    ROLES.GESTOR_COMERCIAL_COMERCIAL,
    ROLES.QG,
    ROLES.SALA_MATRICULA,
    ROLES.FDV,
    ROLES.PROMOTOR,
    ROLES.PROMOTOR_RUA,
    ROLES.FDV_COMERCIAL,
  ].includes(profile.role);

  const handleAddCustomMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMsgData.texto.trim()) return;
    setMsgLoading(true);
    try {
      await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
        tipo: "historico",
        texto: newMsgData.texto,
        nome: newMsgData.modelName || undefined,
        createdAt: serverTimestamp(),
      });
      onToast("Mensagem de histórico salva!");
      setNewMsgData({ modelName: "", texto: "" });
      setIsAddMsgModalOpen(false);
    } catch (err: any) {
      console.error("Erro ao salvar mensagem:", err);
      onToast(`Erro ao salvar mensagem: ${err.message}`, "error");
    } finally {
      setMsgLoading(false);
    }
  };

  const handleInsertDefaultHistoricoMessages = async () => {
    try {
      const existing = whatsappMessages.filter((m) => m.tipo === "historico");
      if (existing.length > 0) {
        if (
          !window.confirm(
            "Já existem mensagens para Histórico. Deseja adicionar as mensagens padrões mesmo assim?",
          )
        ) {
          return;
        }
      }

      const defaults = [
        "Olá [nome], tudo bem? Vimos aqui seu interesse no curso de [curso]. Podemos te ajudar?",
        "Oi [nome], aqui é da faculdade! Recebemos seu cadastro sobre o curso de [curso]. Qual o melhor horário para conversarmos?",
        "Olá [nome]! Qual a sua dúvida sobre o curso de [curso]?",
      ];

      for (const texto of defaults) {
        await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
          tipo: "historico",
          texto,
          createdAt: serverTimestamp(),
        });
      }
      onToast("Mensagens padrões de histórico inseridas!");
    } catch (err: any) {
      onToast("Erro ao inserir mensagens padrões.", "error");
    }
  };

  const filteredLeads = useMemo(() => {
    return leads
      .filter((l) => {
        // Gestor Unidade filtering
        if (profile.role === "Gestor Unidade") {
          if (!profile.unidade || l.unidade !== profile.unidade) {
            return false;
          }
        }

        const matchesSearch =
          !searchTerm ||
          l.nome.toLowerCase().includes(searchTerm.toLowerCase()) ||
          l.telefone.includes(searchTerm) ||
          l.acao.toLowerCase().includes(searchTerm.toLowerCase());
        const matchesCourse =
          !courseFilter || l.cursoInteresse === courseFilter;
        const matchesBase =
          baseFilter.length === 0 || baseFilter.includes(l.acao);
        const matchesStatus = !statusFilter || l.status === statusFilter;
        const matchesPromotor =
          !promotorFilter || l.promotorName === promotorFilter;
        const isBlocked = invalidLeadIds.has(l.id);
        const matchesBlocked =
          blockedFilter === "all" ||
          (blockedFilter === "blocked" && isBlocked) ||
          (blockedFilter === "unblocked" && !isBlocked);
        return (
          matchesSearch &&
          matchesCourse &&
          matchesBase &&
          matchesStatus &&
          matchesPromotor &&
          matchesBlocked
        );
      })
      .sort(
        (a, b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0),
      );
  }, [
    leads,
    searchTerm,
    courseFilter,
    baseFilter,
    statusFilter,
    promotorFilter,
    blockedFilter,
    invalidLeadIds,
  ]);

  const stats = useMemo(() => {
    const total = filteredLeads.length;
    const conv = filteredLeads.filter((l) => l.converted).length;
    const userLeads = filteredLeads.filter(
      (l) => l.promotorId === profile.uid,
    ).length;

    // Stats by Course (Top 5)
    const courseGroups: Record<string, number> = {};
    filteredLeads.forEach((l) => {
      const c = l.cursoInteresse || "Não Informado";
      courseGroups[c] = (courseGroups[c] || 0) + 1;
    });
    const byCourse = Object.entries(courseGroups)
      .map(([name, count]) => ({
        name,
        count,
        percentage: total > 0 ? ((count / total) * 100).toFixed(1) : "0",
      }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    // Stats by Status
    const statusGroups: Record<string, number> = {
      Pendente: 0,
      Convertido: 0,
      "Sem retorno": 0,
      Interessado: 0,
      "Não Interessado": 0,
      "Contato via Sales": 0,
    };
    filteredLeads.forEach((l) => {
      const s = l.converted ? "Convertido" : l.status || "Pendente";
      if (statusGroups[s] !== undefined) statusGroups[s] += 1;
      else statusGroups["Pendente"] += 1;
    });
    const byStatus = Object.entries(statusGroups).map(([name, count]) => ({
      name,
      count,
      percentage: total > 0 ? ((count / total) * 100).toFixed(1) : "0",
    }));

    return {
      total,
      conv,
      userLeads,
      rate: total > 0 ? ((conv / total) * 100).toFixed(1) : "0",
      byCourse,
      byStatus,
    };
  }, [filteredLeads, profile]);

  const toggleSelect = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedEntries((prev) => [...prev, id]);
    } else {
      setSelectedEntries((prev) => prev.filter((s) => s !== id));
    }
  };

  const toggleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedEntries(
        filteredLeads.filter((l) => !invalidLeadIds.has(l.id)).map((l) => l.id),
      );
    } else {
      setSelectedEntries([]);
    }
  };

  const handleStatusChange = async (id: string, newStatus: string) => {
    try {
      await updateDoc(doc(db, COLLECTIONS.LEADS, id), { status: newStatus });
      onToast("Status atualizado!");
    } catch (err: any) {
      handleFirestoreError(
        err,
        OperationType.UPDATE,
        `${COLLECTIONS.LEADS}/${id}`,
      );
      onToast("Erro ao atualizar status.", "error");
    }
  };

  const handleContatoViaSales = async (contact: any, origem: string) => {
    try {
      await addDoc(collection(db, COLLECTIONS.SALES_CONTACTS), {
        contactId: contact.id,
        nome: contact.nome,
        telefone: contact.telefone,
        curso: contact.cursoInteresse || contact.curso || 'Não informado',
        origem,
        createdAt: serverTimestamp(),
      });
      onToast("Contato via Sales registrado com sucesso!", "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar Contato via Sales.", "error");
    }
  };

  const handleMoveToGap = async (lead: Lead) => {
    try {
      await addDoc(collection(db, COLLECTIONS.GAP), {
        nome: lead.nome,
        telefone: lead.telefone,
        matAcad: false,
        documentos: {},
        leadId: lead.id,
        createdAt: serverTimestamp(),
      });
      onToast("Candidato movido para o GAP!");
    } catch (err: any) {
      handleFirestoreError(err, OperationType.CREATE, COLLECTIONS.GAP);
      onToast("Erro ao mover para o GAP.", "error");
    }
  };

  const handleDeleteLead = async (id: string) => {
    if (
      !window.confirm("Tem certeza que deseja excluir este lead do histórico?")
    )
      return;
    try {
      await deleteDoc(doc(db, COLLECTIONS.LEADS, id));
      onToast("Lead excluído com sucesso!", "success");
      setSelectedEntries((prev) => prev.filter((s) => s !== id));
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao excluir lead.", "error");
    }
  };

  const handleBulkDelete = async () => {
    if (
      !window.confirm(
        `Tem certeza que deseja excluir ${selectedEntries.length} lead(s) do histórico?`,
      )
    )
      return;
    try {
      const firestoreBatch = writeBatch(db);
      selectedEntries.forEach((id) => {
        firestoreBatch.delete(doc(db, COLLECTIONS.LEADS, id));
      });
      await firestoreBatch.commit();
      onToast(
        `${selectedEntries.length} lead(s) excluído(s) com sucesso!`,
        "success",
      );
      setSelectedEntries([]);
    } catch (err) {
      console.error(err);
      onToast("Erro ao excluir leads em massa.", "error");
    }
  };

  const handleExport = () => {
    const data = filteredLeads.map((l) => ({
      Nome: l.nome,
      Telefone: l.telefone,
      CPF: l.cpf || "",
      Email: l.email || "",
      Curso: l.cursoInteresse || "",
      Acao: l.acao,
      Promotor: l.promotorName,
      Status: l.converted ? "Convertido" : "Pendente",
      Data: l.createdAt?.seconds
        ? new Date(l.createdAt.seconds * 1000).toLocaleDateString()
        : "",
    }));
    exportToExcel(data, "Historico_Leads");
  };

  const handleExportMalaDireta = () => {
    const data = filteredLeads.map((l) => ({
      Nome: l.nome,
      Email: l.email || "",
    }));
    exportToExcel(data, "Mala_Direta_Leads");
  };

  const handleExportSMS = () => {
    const data = filteredLeads.map((l) => {
      let tel = l.telefone.replace(/\D/g, "");
      if (tel.length > 0 && !tel.startsWith("55")) {
        tel = "55" + tel;
      }
      return { Telefone: tel };
    });
    exportToCSV(data, "SMS_Leads");
  };

  const handleEditClick = (lead: Lead) => {
    setEditingLead(lead);
    setEditFormData({
      nome: lead.nome,
      telefone: lead.telefone,
      email: lead.email || "",
      cpf: lead.cpf || "",
      cursoInteresse: lead.cursoInteresse || "",
      acao: lead.acao,
      acaoId: lead.acaoId || "",
    });
    setEditModalOpen(true);
  };

  const handleSaveEdit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingLead) return;
    try {
      const prevAcaoId = editingLead.acaoId;
      const newAcaoId = editFormData.acaoId;

      await updateDoc(doc(db, COLLECTIONS.LEADS, editingLead.id), {
        nome: editFormData.nome,
        telefone: editFormData.telefone,
        cpf: editFormData.cpf,
        cursoInteresse: editFormData.cursoInteresse,
        acao: editFormData.acao,
        acaoId: newAcaoId || "",
      });

      if (prevAcaoId && prevAcaoId !== "manual" && prevAcaoId !== newAcaoId) {
        try {
          const qLeadsOld = query(
            collection(db, COLLECTIONS.LEADS),
            where("acaoId", "==", prevAcaoId),
          );
          const snapOld = await getDocs(qLeadsOld);
          await updateDoc(doc(db, COLLECTIONS.CALENDARIO_ACOES, prevAcaoId), {
            leadsFeitos: snapOld.size,
          });
        } catch (err) {
          console.error(err);
        }
      }

      if (newAcaoId && newAcaoId !== "manual") {
        try {
          const qLeadsNew = query(
            collection(db, COLLECTIONS.LEADS),
            where("acaoId", "==", newAcaoId),
          );
          const snapNew = await getDocs(qLeadsNew);
          await updateDoc(doc(db, COLLECTIONS.CALENDARIO_ACOES, newAcaoId), {
            leadsFeitos: snapNew.size,
          });
        } catch (err) {
          console.error(err);
        }
      }

      onToast("Lead atualizado com sucesso!", "success");
      setEditModalOpen(false);
      setEditingLead(null);
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao editar lead.", "error");
    }
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return undefined;
        };

        const batch = data.map((item) => {
          const rawStatus = String(getVal(item, "Status", "status") || "")
            .trim()
            .toLowerCase();
          const isConverted =
            rawStatus === "convertido" || getVal(item, "converted") === true;

          return {
            nome: String(getVal(item, "Nome", "nome") || "").trim(),
            telefone: String(
              getVal(item, "Telefone", "telefone") || "",
            ).replace(/\D/g, ""),
            cpf: String(getVal(item, "CPF", "cpf") || "").replace(/\D/g, ""),
            cursoInteresse: String(
              getVal(item, "Curso", "cursoInteresse", "curso") || "",
            ).trim(),
            acao: String(
              getVal(item, "Acao", "acao", "Ação", "ação") || "Importação",
            ).trim(),
            promotorId: "import",
            promotorName: String(
              getVal(item, "Promotor", "promotorName") || "Sistema",
            ).trim(),
            converted: isConverted,
            unidade: profile.unidade || "",
            createdAt: serverTimestamp(),
          };
        });

        let imported = 0;
        let skipped = 0;
        const insertedCpfs = new Set();
        const insertedTels = new Set();

        for (const entry of batch) {
          const isDupCpf =
            entry.cpf &&
            (leads.some((l) => l.cpf === entry.cpf) ||
              insertedCpfs.has(entry.cpf));
          const isDupTel =
            entry.telefone &&
            (leads.some((l) => l.telefone === entry.telefone) ||
              insertedTels.has(entry.telefone));

          if (!isDupCpf && !isDupTel) {
            await addDoc(collection(db, COLLECTIONS.LEADS), entry);
            if (entry.cpf) insertedCpfs.add(entry.cpf);
            if (entry.telefone) insertedTels.add(entry.telefone);
            imported++;
          } else {
            skipped++;
          }
        }
        onToast(
          `${imported} leads importados! ${skipped > 0 ? `${skipped} ignorados por duplicidade.` : ""}`,
        );
      } catch (err: any) {
        onToast("Erro ao importar leads.", "error");
      }
    });
  };

  return (
    <div className="space-y-6">
      {/* Tab Navigation */}
      <div className="flex items-center gap-1 bg-white p-1 rounded-2xl shadow-sm border border-slate-100 w-fit">
        <button
          onClick={() => setHistoricoSubTab("dashboard")}
          className={cn(
            "flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all",
            historicoSubTab === "dashboard"
              ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
              : "text-slate-500 hover:bg-slate-50",
          )}
        >
          <BarChart3 size={18} />
          <span>Dashboard</span>
        </button>
        <button
          onClick={() => setHistoricoSubTab("lista")}
          className={cn(
            "flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all",
            historicoSubTab === "lista"
              ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
              : "text-slate-500 hover:bg-slate-50",
          )}
        >
          <List size={18} />
          <span>Lista de Leads</span>
        </button>
        <button
          onClick={() => setHistoricoSubTab("pedidos_cursos")}
          className={cn(
            "flex items-center gap-2 px-6 py-2.5 rounded-xl text-sm font-bold transition-all",
            historicoSubTab === "pedidos_cursos"
              ? "bg-blue-600 text-white shadow-lg shadow-blue-200"
              : "text-slate-500 hover:bg-slate-50",
          )}
        >
          <GraduationCap size={18} />
          <span>Pedidos de Cursos</span>
        </button>
      </div>

      {historicoSubTab === "dashboard" && (
        <div className="space-y-6">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <StatCard
              title="Total de Leads"
              value={stats.total}
              icon={Users}
              color="bg-blue-500"
            />
            <StatCard
              title="Convertidos"
              value={stats.conv}
              icon={CheckCircle2}
              color="bg-emerald-500"
            />
            <StatCard
              title="Taxa de Conv."
              value={`${stats.rate}%`}
              icon={TrendingUp}
              color="bg-purple-500"
            />
            <StatCard
              title="Meus Leads"
              value={stats.userLeads}
              icon={UserPlus}
              color="bg-amber-500"
            />
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                <Target size={18} className="text-blue-500" />
                Status dos Leads
              </h3>
              <div className="space-y-3">
                {stats.byStatus.map((s) => (
                  <div key={s.name} className="space-y-1">
                    <div className="flex justify-between text-xs font-semibold">
                      <span className="text-slate-600 flex items-center gap-1.5">
                        <span
                          className={cn(
                            "w-2 h-2 rounded-full",
                            s.name === "Convertido" && "bg-emerald-400",
                            s.name === "Pendente" && "bg-amber-400",
                            s.name === "Interessado" && "bg-blue-400",
                            s.name === "Não Interessado" && "bg-rose-400",
                            s.name === "Sem retorno" && "bg-slate-400",
                            s.name === "Contato via Sales" && "bg-purple-400",
                          )}
                        />
                        {s.name}
                      </span>
                      <span className="text-slate-800 font-bold">
                        {s.count}{" "}
                        <span className="text-slate-400 font-normal">
                          ({s.percentage}%)
                        </span>
                      </span>
                    </div>
                    <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                      <div
                        className={cn(
                          "h-full rounded-full transition-all",
                          s.name === "Convertido" && "bg-emerald-400",
                          s.name === "Pendente" && "bg-amber-400",
                          s.name === "Interessado" && "bg-blue-400",
                          s.name === "Não Interessado" && "bg-rose-400",
                          s.name === "Sem retorno" && "bg-slate-400",
                            s.name === "Contato via Sales" && "bg-purple-400",
                        )}
                        style={{ width: `${s.percentage}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </div>

            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                <GraduationCap size={18} className="text-blue-500" />
                Cursos de Interesse (Top 5)
              </h3>
              <div className="space-y-3">
                {stats.byCourse.map((p) => (
                  <div key={p.name} className="space-y-1">
                    <div className="flex justify-between text-xs font-semibold">
                      <span className="text-slate-600 truncate max-w-[200px]">
                        {p.name}
                      </span>
                      <span className="text-slate-800 font-bold">
                        {p.count}{" "}
                        <span className="text-slate-400 font-normal">
                          ({p.percentage}%)
                        </span>
                      </span>
                    </div>
                    <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                      <div
                        className="h-full bg-blue-500 rounded-full transition-all"
                        style={{ width: `${p.percentage}%` }}
                      />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {historicoSubTab === "lista" && (
        <>
          <div className="flex justify-between items-center">
            <h2 className="text-2xl font-bold text-slate-800">
              Histórico de Leads
            </h2>
            <div className="flex space-x-2">
              {[ROLES.ADMIN_MASTER, ROLES.LIDER_FDV].includes(profile.role) && (
                <button
                  onClick={handleVerificacao}
                  className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold"
                  title="Verificar se leads existem no GAP ou Base Líquida"
                >
                  <Search size={18} />
                  <span>Verificação</span>
                </button>
              )}
              <button
                onClick={() => setIsAddMsgModalOpen(true)}
                className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold"
              >
                <Plus size={18} />
                <span>Inserir Mensagens</span>
              </button>
              <button
                onClick={handleInsertDefaultHistoricoMessages}
                className="bg-slate-50 text-slate-400 px-3 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-100 transition-all text-[10px] font-bold"
                title="Inserir Mensagens Padrões"
              >
                <MessageSquare size={14} />
              </button>
              <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer">
                <Upload size={18} />
                <span>Importar</span>
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  onChange={handleImport}
                  className="hidden"
                />
              </label>
              <button
                onClick={handleExport}
                className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold"
              >
                <Download size={18} />
                <span>Exportar Excel</span>
              </button>
              <button
                onClick={handleExportMalaDireta}
                className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold shadow-sm"
              >
                <Mail size={18} />
                <span>Mala Direta</span>
              </button>
              <button
                onClick={handleExportSMS}
                className="bg-orange-50 text-orange-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-orange-100 transition-all text-sm font-bold shadow-sm"
              >
                <MessageSquare size={18} />
                <span>SMS (CSV)</span>
              </button>
            </div>
          </div>

          <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="p-6 border-b border-slate-100 flex flex-col xl:flex-row xl:items-center justify-between gap-4">
              <h3 className="text-xl font-bold text-slate-900 whitespace-nowrap font-sans tracking-tight">
                Lista de Leads
              </h3>
              <div className="flex flex-wrap items-center gap-2 w-full xl:w-auto xl:justify-end">
                <div className="relative flex-1 min-w-[200px] xl:flex-none">
                  <Search
                    className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                    size={16}
                  />
                  <input
                    type="text"
                    placeholder="Buscar por nome, telefone..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    className="w-full pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl focus:ring-2 focus:ring-blue-500 outline-none transition-all text-xs"
                  />
                </div>
                <MultiSelect
                  options={uniqueBases}
                  selectedValues={baseFilter}
                  onChange={setBaseFilter}
                  placeholder="Todas as Origens / Ações"
                  allLabel="Todas as Origens"
                />
                <select
                  className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500 max-w-[150px] lg:max-w-[200px] truncate"
                  value={courseFilter}
                  onChange={(e) => setCourseFilter(e.target.value)}
                >
                  <option value="">Todos os Cursos</option>
                  {uniqueCursos.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
                <select
                  className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                  value={statusFilter}
                  onChange={(e) => setStatusFilter(e.target.value)}
                >
                  <option value="">Todos os Status</option>
                  {uniqueStatuses.map((s) => (
                    <option key={s} value={s}>
                      {s}
                    </option>
                  ))}
                </select>
                <select
                  className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                  value={blockedFilter}
                  onChange={(e) => setBlockedFilter(e.target.value as any)}
                >
                  <option value="all">Verificação: Todos</option>
                  <option value="blocked">Verificação: Bloqueados</option>
                  <option value="unblocked">Verificação: Ativos</option>
                </select>
                {isAdmin && (
                  <select
                    className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500 max-w-[150px] lg:max-w-[200px] truncate"
                    value={promotorFilter}
                    onChange={(e) => setPromotorFilter(e.target.value)}
                  >
                    <option value="">Todos os Promotores</option>
                    {uniquePromotores.map((p) => (
                      <option key={p} value={p}>
                        {p}
                      </option>
                    ))}
                  </select>
                )}
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase tracking-wider">
                    <th className="px-6 py-4 w-12">
                      <input
                        type="checkbox"
                        checked={
                          filteredLeads.filter((l) => !invalidLeadIds.has(l.id))
                            .length > 0 &&
                          selectedEntries.length ===
                            filteredLeads.filter(
                              (l) => !invalidLeadIds.has(l.id),
                            ).length
                        }
                        onChange={(e) => toggleSelectAll(e.target.checked)}
                      />
                    </th>
                    <th className="px-3 py-4 w-12 text-slate-400">#</th>
                    <th className="px-6 py-4">Candidato</th>
                    <th className="px-6 py-4">Ação / Origem</th>
                    <th className="px-6 py-4">Promotor</th>
                    <th className="px-6 py-4">Status</th>
                    <th className="px-6 py-4 flex flex-col gap-2">
                      {selectedEntries.length > 0 && botConfig.url && (
                        <button
                          onClick={() => setMassSelectorOpen(true)}
                          className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1"
                        >
                          <Bot size={14} /> Em Massa
                        </button>
                      )}
                      {selectedEntries.length > 0 && (
                        <button
                          onClick={handleBulkDelete}
                          className="text-rose-600 font-bold hover:underline py-1 px-2 bg-rose-50 rounded-lg flex items-center gap-1"
                        >
                          <Trash2 size={14} /> Excluir ({selectedEntries.length}
                          )
                        </button>
                      )}
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {filteredLeads.map((lead, index) => (
                    <tr
                      key={lead.id}
                      className={cn(
                        "hover:bg-slate-50/50 transition-all",
                        invalidLeadIds.has(lead.id) && "bg-rose-50/50",
                      )}
                    >
                      <td className="px-6 py-4">
                        <input
                          type="checkbox"
                          disabled={invalidLeadIds.has(lead.id)}
                          checked={selectedEntries.includes(lead.id)}
                          onChange={(e) =>
                            !invalidLeadIds.has(lead.id) &&
                            toggleSelect(lead.id, e.target.checked)
                          }
                        />
                      </td>
                      <td className="px-3 py-4 text-xs font-bold text-slate-400 font-mono">
                        {index + 1}
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex flex-col">
                          <span className="font-bold text-slate-900">
                            {lead.nome}
                          </span>
                          <span className="text-xs text-slate-500">
                            {formatPhone(lead.telefone)}
                          </span>
                          {lead.cursoInteresse && (
                            <span className="text-xs text-slate-600 font-medium">
                              Curso: {lead.cursoInteresse}
                            </span>
                          )}
                          {lead.empresa && (
                            <span className="text-[11px] text-indigo-600 font-bold mt-0.5 bg-indigo-50/60 border border-indigo-100/40 px-2 py-0.5 rounded-md self-start">
                              Empresa: {lead.empresa}
                            </span>
                          )}
                          {lead.cpf && (
                            <span className="text-xs text-slate-400">
                              CPF: {formatCPF(lead.cpf)}
                            </span>
                          )}
                        </div>
                      </td>
                      <td className="px-6 py-4 text-sm text-slate-600">
                        {lead.acao}
                      </td>
                      <td className="px-6 py-4 text-sm text-slate-600 font-medium">
                        {lead.promotorName}
                      </td>
                      <td className="px-6 py-4">
                        <select
                          value={lead.status || "Pendente"}
                          onChange={(e) =>
                            handleStatusChange(lead.id, e.target.value)
                          }
                          className={cn(
                            "px-3 py-1 rounded-full text-[10px] font-bold uppercase transition-all border-none focus:ring-0",
                            lead.status === "Convertido"
                              ? "bg-emerald-100 text-emerald-600"
                              : lead.status === "Interessado"
                                ? "bg-blue-100 text-blue-600"
                                : lead.status === "Não Interessado"
                                  ? "bg-rose-100 text-rose-600"
                                  : lead.status === "Sem retorno"
                                    ? "bg-slate-100 text-slate-600"
                                    : lead.status === "Contato via Sales" ? "bg-purple-100 text-purple-600"
                                    : "bg-amber-100 text-amber-600",
                          )}
                        >
                          <option value="Pendente">Pendente</option>
                          <option value="Sem retorno">Sem retorno</option>
<option value="Contato via Sales">Contato via Sales</option>
                          <option value="Interessado">Interessado</option>
                          <option value="Não Interessado">
                            Não Interessado
                          </option>
                          <option value="Convertido">Convertido</option>
                        </select>
                      </td>
                      <td className="px-6 py-4">
                        <div className="flex items-center space-x-3">
                          {!invalidLeadIds.has(lead.id) && (
                            <button
                              onClick={() => {
                                setSelectedLead(lead);
                                setSelectorOpen(true);
                              }}
                              className="inline-flex items-center space-x-1 text-emerald-600 font-bold text-sm hover:text-emerald-700"
                            >
                              <MessageSquare size={14} />
                              <span>WhatsApp</span>
                            </button>
                          )}
                          <button
                            onClick={() => handleContatoViaSales(lead, 'Leads')}
                            className="inline-flex items-center space-x-1 text-sky-600 font-bold text-sm hover:text-sky-700 bg-sky-50 px-2 py-1 rounded-lg"
                            title="Registrar Contato via Sales"
                          >
                            <PhoneOutgoing size={14} />
                            <span>Sales</span>
                          </button>
                          {lead.status === "Convertido" && !invalidLeadIds.has(lead.id) && (
                            <button
                              onClick={() => handleMoveToGap(lead)}
                              className="text-purple-600 hover:text-purple-700 font-bold text-sm flex items-center space-x-1"
                              title="Mover para GAP Acadêmico"
                            >
                              <GraduationCap size={14} />
                            </button>
                          )}
                          <button
                            onClick={() => handleEditClick(lead)}
                            className="text-slate-400 hover:text-blue-600 transition-colors"
                            title="Editar Lead"
                          >
                            <Edit2 size={16} />
                          </button>
                          <button
                            onClick={() => handleDeleteLead(lead.id)}
                            className="text-slate-400 hover:text-rose-600 transition-colors"
                            title="Excluir Lead"
                          >
                            <Trash2 size={16} />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                  {filteredLeads.length === 0 && (
                    <tr>
                      <td
                        colSpan={7}
                        className="px-6 py-12 text-center text-slate-400 italic"
                      >
                        Nenhum lead encontrado.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </>
      )}

      {historicoSubTab === "pedidos_cursos" && (
        <div className="space-y-6">
          <div className="flex justify-between items-center bg-white p-6 rounded-2xl shadow-sm border border-slate-100">
            <div>
              <h2 className="text-2xl font-bold text-slate-800">
                Pedidos de Cursos
              </h2>
              <p className="text-slate-500 text-sm mt-1">
                Gere o link e acompanhe os cursos solicitados.
              </p>
            </div>
            <button
              onClick={() => {
                const url = `${window.location.origin}${window.location.pathname}?view=pedido-curso`;
                navigator.clipboard.writeText(url);
                onToast(
                  "Link copiado para a área de transferência!",
                  "success",
                );
              }}
              className="bg-indigo-600 text-white px-6 py-2.5 rounded-xl font-bold hover:bg-indigo-700 transition-colors shadow-lg shadow-indigo-200 flex items-center space-x-2"
            >
              <Copy size={18} />
              <span>Gerar Link do Formulário</span>
            </button>
          </div>

          <div className="bg-white rounded-2xl shadow-sm border border-slate-100 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50 text-slate-500 text-xs uppercase tracking-wider">
                    <th className="p-4 font-bold border-b border-slate-100">
                      Nome
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Telefone
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Curso
                    </th>
                    <th className="p-4 font-bold border-b border-slate-100">
                      Data
                    </th>
                  </tr>
                </thead>
                <tbody className="text-sm divide-y divide-slate-50">
                  {pedidosCursos.length === 0 ? (
                    <tr>
                      <td
                        colSpan={4}
                        className="p-8 text-center text-slate-400"
                      >
                        Nenhum pedido registrado.
                      </td>
                    </tr>
                  ) : (
                    pedidosCursos.map((pedido) => (
                      <tr
                        key={pedido.id}
                        className="hover:bg-slate-50/50 transition-colors"
                      >
                        <td className="p-4 font-medium text-slate-700">
                          {pedido.nome}
                        </td>
                        <td className="p-4 text-slate-600">
                          {pedido.telefone}
                        </td>
                        <td className="p-4 text-slate-800 font-semibold">
                          {pedido.curso}
                        </td>
                        <td className="p-4 text-slate-500">
                          {pedido.createdAt
                            ? new Date(
                                pedido.createdAt.toDate(),
                              ).toLocaleDateString()
                            : "-"}
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      <WhatsAppMessageSelector
        isOpen={selectorOpen}
        onClose={() => setSelectorOpen(false)}
        leadName={selectedLead?.nome || ""}
        leadCurso={selectedLead?.cursoInteresse || ""}
        messages={whatsappMessages.filter((m) => m.tipo === "historico")}
        onSelect={(msg) => {
          if (selectedLead) {
            window.open(getWhatsAppUrl(selectedLead.telefone, msg), "_blank");
          }
        }}
        botConfig={botConfig}
        onSendBot={(msg, contactName) => {
          if (selectedLead) {
            onSendBot(
              selectedLead.telefone,
              Array.isArray(msg) ? msg[0] : msg,
              contactName || selectedLead.nome,
            );
          }
        }}
      />

      <WhatsAppMessageSelector
        isOpen={massSelectorOpen}
        onClose={() => setMassSelectorOpen(false)}
        leadName="Candidatos"
        messages={whatsappMessages.filter((m) => m.tipo === "historico")}
        onSelect={(msg) => {
          // not used for mass send
        }}
        botConfig={botConfig}
        onSendBot={(msgTemplates) => {
          const templates = Array.isArray(msgTemplates) ? msgTemplates : [msgTemplates];
          const selectedLeadObjs = leads.filter(
            (l) => selectedEntries.includes(l.id) && !invalidLeadIds.has(l.id),
          );
          const messagesPayload = selectedLeadObjs.map((l, idx) => {
            const template = templates[idx % templates.length];
            return {
              telefone: l.telefone,
              message: replaceMessageVariables(template, l),
            };
          });
          onMassSendBot(messagesPayload);
          setMassSelectorOpen(false);
          setSelectedEntries([]);
        }}
        forceBotOnly={true}
      />

      <MessageTemplateModal
        isOpen={isAddMsgModalOpen}
        onClose={() => setIsAddMsgModalOpen(false)}
        tipo="historico"
        onToast={onToast}
        availableVariables={[
          { key: "[nome]", label: "Nome do Lead", previewValue: "João Silva" },
          {
            key: "[curso]",
            label: "Curso",
            previewValue: "Engenharia de Software",
          },
          {
            key: "[unidade]",
            label: "Unidade",
            previewValue: "Unidade Central",
          },
          {
            key: "[data_contato]",
            label: "Data",
            previewValue: new Date().toLocaleDateString("pt-BR"),
          },
          { key: "[saudacao]", label: "Saudação", previewValue: "Bom dia" },
        ]}
      />

      {editModalOpen && (
        <div className="fixed inset-0 z-[70] flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm">
          <motion.div
            initial={{ opacity: 0, scale: 0.95 }}
            animate={{ opacity: 1, scale: 1 }}
            exit={{ opacity: 0, scale: 0.95 }}
            className="bg-white rounded-3xl shadow-2xl w-full max-w-lg overflow-hidden"
          >
            <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
              <h3 className="text-xl font-bold text-slate-900">Editar Lead</h3>
              <button
                onClick={() => {
                  setEditModalOpen(false);
                  setEditingLead(null);
                }}
                className="text-slate-400 hover:text-slate-600"
              >
                <X size={24} />
              </button>
            </div>
            <form onSubmit={handleSaveEdit} className="p-6 space-y-4">
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nome
                  </label>
                  <input
                    required
                    value={editFormData.nome}
                    onChange={(e) =>
                      setEditFormData({ ...editFormData, nome: e.target.value })
                    }
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Telefone
                  </label>
                  <input
                    required
                    value={editFormData.telefone}
                    onChange={(e) => {
                      e.target.value = formatPhone(e.target.value);
                      setEditFormData({
                        ...editFormData,
                        telefone: e.target.value,
                      });
                    }}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    E-mail (Opcional)
                  </label>
                  <input
                    type="email"
                    value={editFormData.email}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        email: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                    placeholder="exemplo@gmail.com"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    CPF
                  </label>
                  <input
                    value={editFormData.cpf}
                    onChange={(e) => {
                      e.target.value = formatCPF(e.target.value);
                      setEditFormData({ ...editFormData, cpf: e.target.value });
                    }}
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Curso
                  </label>
                  <input
                    value={editFormData.cursoInteresse}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        cursoInteresse: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                  />
                </div>
                <div className="col-span-2 space-y-2">
                  <label className="block text-xs font-bold text-slate-500">
                    Origem / Ação
                  </label>
                  {calendarioAcoes && calendarioAcoes.length > 0 ? (
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <span className="block text-[10px] font-semibold text-slate-400 mb-1">
                          Selecionar do Calendário
                        </span>
                        <select
                          value={editFormData.acaoId || ""}
                          onChange={(e) => {
                            const val = e.target.value;
                            if (val === "manual") {
                              setEditFormData({
                                ...editFormData,
                                acaoId: "manual",
                                acao: "",
                              });
                            } else {
                              const matched = calendarioAcoes.find(
                                (a) => a.id === val,
                              );
                              setEditFormData({
                                ...editFormData,
                                acaoId: val,
                                acao: matched ? matched.nome : "",
                              });
                            }
                          }}
                          className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm bg-white"
                        >
                          <option value="">Selecione...</option>
                          {calendarioAcoes.map((act) => (
                            <option key={act.id} value={act.id}>
                              {act.nome} ({act.dataInicio})
                            </option>
                          ))}
                          <option value="manual">
                            Outro (Digitar manualmente)
                          </option>
                        </select>
                      </div>
                      {(editFormData.acaoId === "manual" ||
                        !editFormData.acaoId) && (
                        <div>
                          <span className="block text-[10px] font-semibold text-slate-400 mb-1">
                            Digitar Nome da Ação/Origem
                          </span>
                          <input
                            type="text"
                            required={
                              !editFormData.acaoId ||
                              editFormData.acaoId === "manual"
                            }
                            value={editFormData.acao}
                            onChange={(e) =>
                              setEditFormData({
                                ...editFormData,
                                acao: e.target.value,
                              })
                            }
                            className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none text-sm"
                            placeholder="Ex: Facebook, Panfletagem, etc."
                          />
                        </div>
                      )}
                    </div>
                  ) : (
                    <input
                      value={editFormData.acao}
                      onChange={(e) =>
                        setEditFormData({
                          ...editFormData,
                          acao: e.target.value,
                        })
                      }
                      className="w-full px-4 py-2 rounded-xl border border-slate-200 text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                    />
                  )}
                </div>
              </div>
              <button
                type="submit"
                className="w-full bg-blue-600 hover:bg-blue-700 text-white py-3 rounded-xl font-bold transition flex items-center justify-center space-x-2"
              >
                <span>Salvar Alterações</span>
              </button>
            </form>
          </motion.div>
        </div>
      )}
    </div>
  );
}

function BasesView({
  bases,
  onToast,
  whatsappMessages,
  botConfig,
  onSendBot,
  onMassSendBot,
  gap,
  basesRenovacao,
  profile,
}: {
  bases: BaseEntry[];
  onToast: (m: string, t?: "success" | "error") => void;
  whatsappMessages: WhatsAppMessage[];
  botConfig: BotConfig;
  onSendBot: (tel: string, msg: string, contactName?: string) => void;
  onMassSendBot: (
    messages: { telefone: string; message: string; nome?: string }[],
  ) => void;
  gap: GapEntry[];
  basesRenovacao: BaseEntry[];
  profile: UserProfile;
}) {
  const handleContatoViaSales = async (contact: any, origem: string) => {
    try {
      await addDoc(collection(db, COLLECTIONS.SALES_CONTACTS), {
        contactId: contact.id,
        nome: contact.nome,
        telefone: contact.telefone,
        curso: contact.cursoInteresse || contact.curso || "Não informado",
        origem,
        createdAt: serverTimestamp(),
      });
      onToast("Contato via Sales registrado com sucesso!", "success");
    } catch (err: any) {
      console.error(err);
      onToast("Erro ao registrar Contato via Sales.", "error");
    }
  };
  
  const [formData, setFormData] = useState({
    nomeBase: "",
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    curso: "",
    produto: "Graduação" as "Graduação" | "Técnico" | "Pós-graduação",
    numeroOportunidade: "",
    semestre: "",
    periodo: "",
    metodologia: "",
    formaIngresso: "",
    numeroMatricula: "",
  });
  const [loading, setLoading] = useState(false);
  const [searchTerm, setSearchTerm] = useState("");
  const [baseFilter, setBaseFilter] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState("");
  const [produtoFilter, setProdutoFilter] = useState("");
  const [cursoFilter, setCursoFilter] = useState("");
  const [semestreFilter, setSemestreFilter] = useState("");
  const [selectorOpen, setSelectorOpen] = useState(false);
  const [selectedEntry, setSelectedEntry] = useState<BaseEntry | null>(null);
  const [massSelectorOpen, setMassSelectorOpen] = useState(false);
  const [isAddMsgModalOpen, setIsAddMsgModalOpen] = useState(false);
  const [newMsgData, setNewMsgData] = useState({ modelName: "", texto: "" });
  const [invalidBaseIds, setInvalidBaseIds] = useState<Set<string>>(new Set());
  const [blockedFilter, setBlockedFilter] = useState<
    "all" | "blocked" | "unblocked"
  >("all");

  // New States for Sub-tabs and Candidates Editing
  const [basesSubTab, setBasesSubTab] = useState<
    "dashboard" | "lista" | "novo"
  >("dashboard");
  const [editingCandidate, setEditingCandidate] = useState<BaseEntry | null>(
    null,
  );
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editFormData, setEditFormData] = useState({
    nomeBase: "",
    nome: "",
    telefone: "",
    email: "",
    cpf: "",
    curso: "",
    produto: "Graduação" as "Graduação" | "Técnico" | "Pós-graduação",
    numeroOportunidade: "",
    semestre: "",
    periodo: "",
    metodologia: "",
    formaIngresso: "",
    numeroMatricula: "",
    status: "Pendente" as
      | "Pendente"
      | "Interessado"
      | "Convertido"
      | "Não tem interesse"
      | "Sem retorno"
      | "Contato via Sales",
  });

  // Memoized aggregations for Dashboard basic metrics
  const statsByBase = useMemo(() => {
    const groups: {
      [key: string]: {
        total: number;
        converted: number;
        interested: number;
        pending: number;
      };
    } = {};
    bases.forEach((b) => {
      const baseName = b.nomeBase || "Sem Nome";
      if (!groups[baseName]) {
        groups[baseName] = {
          total: 0,
          converted: 0,
          interested: 0,
          pending: 0,
        };
      }
      groups[baseName].total += 1;
      if (b.status === "Convertido") groups[baseName].converted += 1;
      if (b.status === "Interessado") groups[baseName].interested += 1;
      if (b.status === "Pendente") groups[baseName].pending += 1;
    });

    return Object.entries(groups)
      .map(([name, data]) => ({
        name,
        total: data.total,
        converted: data.converted,
        interested: data.interested,
        pending: data.pending,
        conversionRate:
          data.total > 0
            ? ((data.converted / data.total) * 100).toFixed(1)
            : "0",
      }))
      .sort((a, b) => b.total - a.total);
  }, [bases]);

  const statsByProduct = useMemo(() => {
    const groups: { [key: string]: number } = {
      Graduação: 0,
      Técnico: 0,
      "Pós-graduação": 0,
    };
    bases.forEach((b) => {
      const p = b.produto || "Graduação";
      if (groups[p] !== undefined) {
        groups[p] += 1;
      } else {
        groups[p] = 1;
      }
    });
    return Object.entries(groups).map(([name, count]) => ({
      name,
      count,
      percentage:
        bases.length > 0 ? ((count / bases.length) * 100).toFixed(1) : "0",
    }));
  }, [bases]);

  const statsByStatus = useMemo(() => {
    const groups: { [key: string]: number } = {
      Pendente: 0,
      Interessado: 0,
      Convertido: 0,
      "Não tem interesse": 0,
      "Sem retorno": 0,
      "Contato via Sales": 0,
    };
    bases.forEach((b) => {
      const s = b.status || "Pendente";
      if (groups[s] !== undefined) {
        groups[s] += 1;
      }
    });
    return Object.entries(groups).map(([name, count]) => ({
      name,
      count,
      percentage:
        bases.length > 0 ? ((count / bases.length) * 100).toFixed(1) : "0",
    }));
  }, [bases]);

  const handleEditSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingCandidate) return;

    setLoading(true);
    try {
      const cleanCpf = editFormData.cpf
        ? editFormData.cpf.replace(/\D/g, "")
        : "";
      const cleanTelefone = editFormData.telefone.replace(/\D/g, "");

      const updatedData = {
        ...editFormData,
        cpf: cleanCpf,
        telefone: cleanTelefone,
        updatedAt: serverTimestamp(),
      };

      await updateDoc(
        doc(db, COLLECTIONS.BASES, editingCandidate.id),
        updatedData,
      );

      // If conversion status toggled to Convertido, check and sync with GAP
      if (
        editFormData.status === "Convertido" &&
        editingCandidate.status !== "Convertido" &&
        !invalidBaseIds.has(editingCandidate.id)
      ) {
        const q = query(
          collection(db, COLLECTIONS.GAP),
          where("cpf", "==", cleanCpf || ""),
        );
        const snap = await getDocs(q);
        if (snap.empty && cleanCpf) {
          await addDoc(collection(db, COLLECTIONS.GAP), {
            nome: editFormData.nome,
            telefone: cleanTelefone,
            cpf: cleanCpf,
            produto: editFormData.produto,
            numeroOportunidade: editFormData.numeroOportunidade,
            curso: editFormData.curso,
            metodologia: editFormData.metodologia,
            formaIngresso: editFormData.formaIngresso,
            semestre: editFormData.semestre,
            matAcad: false,
            documentos: {},
            createdAt: serverTimestamp(),
          });
          onToast(
            "Candidato atualizado e enviado para o GAP (Convertido)!",
            "success",
          );
        } else {
          onToast("Status atualizado com sucesso!", "success");
        }
      } else {
        onToast("Informações do candidato atualizadas com sucesso!", "success");
      }

      setIsEditModalOpen(false);
      setEditingCandidate(null);
    } catch (err: any) {
      onToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  const handleVerificacao = () => {
    const invalidIds = new Set<string>();
    bases.forEach((base) => {
      let match = false;

      if (
        gap.some(
          (g) =>
            (g.cpf &&
              base.cpf &&
              g.cpf.replace(/\D/g, "") === base.cpf.replace(/\D/g, "")) ||
            (g.telefone &&
              base.telefone &&
              g.telefone.replace(/\D/g, "") ===
                base.telefone.replace(/\D/g, "")) ||
            g.nome.toLowerCase().trim() === base.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (
        !match &&
        basesRenovacao.some(
          (b) =>
            (b.cpf &&
              base.cpf &&
              b.cpf.replace(/\D/g, "") === base.cpf.replace(/\D/g, "")) ||
            (b.telefone &&
              base.telefone &&
              b.telefone.replace(/\D/g, "") ===
                base.telefone.replace(/\D/g, "")) ||
            b.nome.toLowerCase().trim() === base.nome.toLowerCase().trim(),
        )
      ) {
        match = true;
      }

      if (match) {
        invalidIds.add(base.id);
      }
    });
    setInvalidBaseIds(invalidIds);
    onToast(
      `Verificação concluída: ${invalidIds.size} contatos já estão cadastrados em GAP/Base Líquida.`,
      "success",
    );
  };

  const filteredBases = bases.filter((b) => {
    // Gestor Unidade filtering
    if (profile.role === "Gestor Unidade") {
      if (!profile.unidade || b.unidade !== profile.unidade) {
        return false;
      }
    }

    const matchesSearch = b.nome
      .toLowerCase()
      .includes(searchTerm.toLowerCase());
    const matchesBase =
      baseFilter.length === 0 || baseFilter.includes(b.nomeBase);
    const matchesStatus = !statusFilter || b.status === statusFilter;
    const matchesProduto = !produtoFilter || b.produto === produtoFilter;
    const matchesCurso =
      !cursoFilter || b.curso.toLowerCase().includes(cursoFilter.toLowerCase());
    const matchesSemestre =
      !semestreFilter ||
      (b.semestre &&
        b.semestre.toLowerCase().includes(semestreFilter.toLowerCase()));

    const isBlocked = invalidBaseIds.has(b.id);
    const matchesBlocked =
      blockedFilter === "all" ||
      (blockedFilter === "blocked" && isBlocked) ||
      (blockedFilter === "unblocked" && !isBlocked);

    return (
      matchesSearch &&
      matchesBase &&
      matchesStatus &&
      matchesProduto &&
      matchesCurso &&
      matchesSemestre &&
      matchesBlocked
    );
  });
  const uniqueBases = Array.from(new Set(bases.map((b) => b.nomeBase))).sort();
  const uniqueProdutos = ["Graduação", "Técnico", "Pós-graduação"];
  const uniqueCursos = Array.from(new Set(bases.map((b) => b.curso))).sort();
  const uniqueSemestres = Array.from(
    new Set(bases.map((b) => b.semestre).filter(Boolean)),
  ).sort();

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();

    const cleanCpf = formData.cpf ? formData.cpf.replace(/\D/g, "") : "";
    const cleanTelefone = formData.telefone.replace(/\D/g, "");

    const isDuplicate = bases.some(
      (b) =>
        (cleanCpf && b.cpf === cleanCpf) ||
        (!cleanCpf && cleanTelefone && b.telefone === cleanTelefone),
    );

    if (isDuplicate) {
      onToast("Registro já existe na base (verificado CPF/Telefone).", "error");
      return;
    }

    setLoading(true);
    try {
      await addDoc(collection(db, COLLECTIONS.BASES), {
        ...formData,
        status: "Pendente",
        unidade: profile.unidade || "",
        promotorId: profile.uid,
        linkadoA: profile.linkadoA || null,
        createdAt: serverTimestamp(),
      });
      onToast("Registro salvo na base!");
      setFormData({
        nomeBase: "",
        nome: "",
        telefone: "",
        email: "",
        cpf: "",
        curso: "",
        produto: "Graduação",
        numeroOportunidade: "",
        semestre: "",
        periodo: "",
        metodologia: "",
        formaIngresso: "",
        numeroMatricula: "",
      });
    } catch (err: any) {
      onToast(err.message, "error");
    } finally {
      setLoading(false);
    }
  };

  const handleAddCustomMessage = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newMsgData.texto.trim()) return;
    setLoading(true);
    try {
      await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
        tipo: "bases",
        texto: newMsgData.texto,
        nome: newMsgData.modelName || undefined,
        createdAt: serverTimestamp(),
      });
      onToast("Mensagem de base salva!");
      setNewMsgData({ modelName: "", texto: "" });
      setIsAddMsgModalOpen(false);
    } catch (err: any) {
      console.error("Erro ao salvar mensagem:", err);
      onToast(`Erro ao salvar mensagem: ${err.message}`, "error");
    } finally {
      setLoading(false);
    }
  };

  const handleInsertDefaultBasesMessages = async () => {
    try {
      const existing = whatsappMessages.filter((m) => m.tipo === "bases");
      if (existing.length > 0) {
        if (
          !window.confirm(
            "Já existem mensagens para Bases. Deseja adicionar as mensagens padrões mesmo assim?",
          )
        ) {
          return;
        }
      }

      const defaults = [
        "Olá [nome], vi que você tem interesse no curso de [curso]. Vamos tirar suas dúvidas?",
        "Oi [nome], aqui é da faculdade! Recebemos sua solicitação sobre o curso de [curso]. Qual o melhor horário para conversarmos?",
        "Tudo bem, [nome]? Preparamos uma oferta especial para você começar o curso de [curso] ainda este semestre! Vamos lá?",
      ];

      for (const texto of defaults) {
        await addDoc(collection(db, COLLECTIONS.WHATSAPP_MESSAGES), {
          tipo: "bases",
          texto,
          createdAt: serverTimestamp(),
        });
      }
      onToast("Mensagens padrões de base inseridas!");
    } catch (err: any) {
      onToast("Erro ao inserir mensagens padrões.", "error");
    }
  };

  const [selectedEntries, setSelectedEntries] = useState<string[]>([]);

  const handleBulkDelete = async () => {
    if (selectedEntries.length === 0) return;
    if (
      window.confirm(
        `Deseja excluir ${selectedEntries.length} registros selecionados?`,
      )
    ) {
      try {
        for (const id of selectedEntries) {
          await deleteDoc(doc(db, COLLECTIONS.BASES, id));
        }
        onToast(`${selectedEntries.length} registros removidos.`);
        setSelectedEntries([]);
      } catch (err: any) {
        onToast("Erro ao excluir registros.", "error");
      }
    }
  };

  const toggleSelect = (id: string, checked: boolean) => {
    if (checked) {
      setSelectedEntries([...selectedEntries, id]);
    } else {
      setSelectedEntries(selectedEntries.filter((s) => s !== id));
    }
  };

  const toggleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedEntries(
        filteredBases.filter((b) => !invalidBaseIds.has(b.id)).map((b) => b.id),
      );
    } else {
      setSelectedEntries([]);
    }
  };

  const handleStatusChange = async (entry: BaseEntry, status: string) => {
    try {
      await updateDoc(doc(db, COLLECTIONS.BASES, entry.id), { status });

      if (status === "Convertido" && !invalidBaseIds.has(entry.id)) {
        // Logic for transferring to GAP
        const q = query(
          collection(db, COLLECTIONS.GAP),
          where("cpf", "==", entry.cpf || ""),
        );
        const snap = await getDocs(q);
        if (snap.empty && entry.cpf) {
          await addDoc(collection(db, COLLECTIONS.GAP), {
            nome: entry.nome,
            telefone: entry.telefone,
            cpf: entry.cpf,
            produto: entry.produto,
            numeroOportunidade: entry.numeroOportunidade,
            curso: entry.curso,
            metodologia: entry.metodologia,
            formaIngresso: entry.formaIngresso,
            semestre: entry.semestre,
            matAcad: false,
            documentos: {},
            createdAt: serverTimestamp(),
          });
          onToast("Candidato convertido e enviado para GAP!");
        } else {
          onToast("Status atualizado!");
        }
      } else {
        onToast("Status da base atualizado!");
      }
    } catch (err: any) {
      onToast(err.message, "error");
    }
  };

  const handleDeleteBase = async (id: string) => {
    if (window.confirm("Deseja excluir este registro da base?")) {
      try {
        await deleteDoc(doc(db, COLLECTIONS.BASES, id));
        onToast("Registro removido.");
      } catch (err: any) {
        onToast("Erro ao excluir registro.", "error");
      }
    }
  };

  const handleExport = () => {
    const data = filteredBases.map((b) => ({
      Nome: b.nome,
      Telefone: b.telefone,
      Email: b.email || "",
      CPF: b.cpf || "",
      Curso: b.curso,
      Produto: b.produto || "Graduação",
      "Nº Oportunidade": b.numeroOportunidade || "",
      Semestre: b.semestre || "",
      Periodo: b.periodo || "",
      Metodologia: b.metodologia || "",
      "Forma de Ingresso": b.formaIngresso || "",
      "Nº Matrícula": b.numeroMatricula || "",
      Base: b.nomeBase,
      Status: b.status,
      Data: b.createdAt?.seconds
        ? new Date(b.createdAt.seconds * 1000).toLocaleDateString()
        : "",
    }));
    exportToExcel(data, "Base_Candidatos");
  };

  const handleExportMalaDireta = () => {
    const data = filteredBases.map((b) => ({
      Nome: b.nome,
      Email: b.email || "",
    }));
    exportToExcel(data, "Mala_Direta_Bases");
  };

  const handleExportSMS = () => {
    const data = filteredBases.map((b) => {
      let tel = b.telefone.replace(/\D/g, "");
      if (tel.length > 0 && !tel.startsWith("55")) {
        tel = "55" + tel;
      }
      return { Telefone: tel };
    });
    exportToCSV(data, "SMS_Bases");
  };

  const handleImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    importFromExcel(file, async (data) => {
      try {
        const getVal = (row: any, ...keys: string[]) => {
          const rowKeys = Object.keys(row);
          for (const key of keys) {
            const foundKey = rowKeys.find(
              (k) => k.toLowerCase() === key.toLowerCase(),
            );
            if (foundKey && row[foundKey] !== undefined) return row[foundKey];
          }
          return undefined;
        };

        const normalizeProduto = (val: string) => {
          if (!val) return "Graduação";
          const lower = val.trim().toLowerCase();
          if (lower.includes("gradua")) return "Graduação";
          if (lower.includes("tecnic") || lower.includes("técnic"))
            return "Técnico";
          if (lower.includes("pos") || lower.includes("pós"))
            return "Pós-graduação";
          return val;
        };

        const normalizeMetodologia = (val: string) => {
          if (!val) return "";
          const lower = val.trim().toLowerCase();
          if (lower === "ead") return "EAD";
          if (lower === "presencial") return "Presencial";
          if (lower === "semipresencial") return "Semipresencial";
          if (lower === "flex") return "Flex";
          if (lower === "hibrido" || lower === "híbrido") return "Híbrido";
          if (lower === "digital") return "Digital";
          return val;
        };

        const normalizeStatusBase = (val: string) => {
          if (!val) return "Pendente";
          const lower = val.trim().toLowerCase();
          if (lower === "pendente") return "Pendente";
          if (lower === "matriculado") return "Matriculado";
          if (
            lower === "ligacao efetuada" ||
            lower === "ligação efetuada" ||
            lower.includes("liga")
          )
            return "Ligação Efetuada";
          if (lower === "sem interesse" || lower.includes("sem inter"))
            return "Sem Interesse";
          return val.charAt(0).toUpperCase() + val.slice(1);
        };

        const batch = data.map((item) => ({
          nome: String(getVal(item, "Nome", "nome") || "").trim(),
          telefone: String(getVal(item, "Telefone", "telefone") || "").replace(
            /\D/g,
            "",
          ),
          cpf: String(getVal(item, "CPF", "cpf") || "").replace(/\D/g, ""),
          curso: String(getVal(item, "Curso", "curso") || "").trim(),
          produto: normalizeProduto(
            String(getVal(item, "Produto", "produto") || ""),
          ),
          numeroOportunidade: String(
            getVal(
              item,
              "Nº Oportunidade",
              "numeroOportunidade",
              "oportunidade",
            ) || "",
          ).trim(),
          semestre: String(getVal(item, "Semestre", "semestre") || "").trim(),
          periodo: String(
            getVal(item, "Periodo", "periodo", "período") || "",
          ).trim(),
          metodologia: normalizeMetodologia(
            String(getVal(item, "Metodologia", "metodologia") || ""),
          ),
          formaIngresso: String(
            getVal(item, "Forma de Ingresso", "formaIngresso", "ingresso") ||
              "",
          ).trim(),
          numeroMatricula: String(
            getVal(
              item,
              "Nº Matrícula",
              "numeroMatricula",
              "matricula",
              "matrícula",
            ) || "",
          ).trim(),
          nomeBase: String(
            getVal(item, "Base", "nomeBase") || "Importado",
          ).trim(),
          status: normalizeStatusBase(
            String(getVal(item, "Status", "status") || ""),
          ),
          createdAt: serverTimestamp(),
        }));

        let imported = 0;
        let skipped = 0;
        const insertedCpfs = new Set();
        const insertedTels = new Set();

        for (const entry of batch) {
          const isDupCpf =
            entry.cpf &&
            (bases.some((b) => b.cpf === entry.cpf) ||
              insertedCpfs.has(entry.cpf));
          const isDupTel =
            entry.telefone &&
            (bases.some((b) => b.telefone === entry.telefone) ||
              insertedTels.has(entry.telefone));

          if (!isDupCpf && !isDupTel) {
            await addDoc(collection(db, COLLECTIONS.BASES), entry);
            if (entry.cpf) insertedCpfs.add(entry.cpf);
            if (entry.telefone) insertedTels.add(entry.telefone);
            imported++;
          } else {
            skipped++;
          }
        }
        onToast(
          `${imported} registros importados com sucesso! ${skipped > 0 ? `${skipped} ignorados por duplicidade.` : ""}`,
        );
      } catch (err: any) {
        onToast("Erro ao importar dados.", "error");
      }
    });
  };

  return (
    <div className="space-y-6">
      {/* Header and Actions */}
      <div className="flex flex-col md:flex-row justify-between items-center gap-4">
        <div>
          <h2 className="text-2xl font-bold text-slate-800 flex items-center gap-2">
            <Database className="text-blue-600" size={28} />
            Acompanhamento de Bases
          </h2>
          <p className="text-sm text-slate-500">
            Gerencie e analise as bases de captação de candidatos da sua
            unidade.
          </p>
        </div>
        <div className="flex flex-wrap gap-2">
          {[ROLES.ADMIN_MASTER, ROLES.LIDER_FDV].includes(profile.role) && (
            <button
              onClick={handleVerificacao}
              className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold shadow-sm"
              title="Verificar se contatos existem no GAP ou Base Líquida"
            >
              <Search size={18} />
              <span>Verificação</span>
            </button>
          )}
          <button
            onClick={() => setIsAddMsgModalOpen(true)}
            className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold shadow-sm"
          >
            <Plus size={18} />
            <span>Inserir Mensagens</span>
          </button>
          <button
            onClick={handleInsertDefaultBasesMessages}
            className="bg-slate-50 text-slate-400 px-3 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-100 transition-all text-[10px] font-bold shadow-sm"
            title="Inserir Mensagens Padrões"
          >
            <MessageSquare size={14} />
          </button>
          <label className="bg-blue-50 text-blue-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-blue-100 transition-all text-sm font-bold cursor-pointer shadow-sm">
            <Upload size={18} />
            <span>Importar</span>
            <input
              type="file"
              accept=".xlsx, .xls"
              onChange={handleImport}
              className="hidden"
            />
          </label>
          <button
            onClick={handleExport}
            className="bg-slate-100 text-slate-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-slate-200 transition-all text-sm font-bold shadow-sm"
          >
            <Download size={18} />
            <span>Exportar Excel</span>
          </button>
          <button
            onClick={handleExportMalaDireta}
            className="bg-emerald-50 text-emerald-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-emerald-100 transition-all text-sm font-bold shadow-sm"
          >
            <Mail size={18} />
            <span>Mala Direta</span>
          </button>
          <button
            onClick={handleExportSMS}
            className="bg-orange-50 text-orange-600 px-4 py-2 rounded-xl flex items-center space-x-2 hover:bg-orange-100 transition-all text-sm font-bold shadow-sm"
          >
            <MessageSquare size={18} />
            <span>SMS (CSV)</span>
          </button>
        </div>
      </div>

      {/* Elegant Sub-tabs */}
      <div className="flex border-b border-slate-100 gap-2 overflow-x-auto">
        <button
          onClick={() => setBasesSubTab("dashboard")}
          className={cn(
            "px-5 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap",
            basesSubTab === "dashboard"
              ? "border-b-2 border-blue-600 text-blue-600 font-bold"
              : "border-b-2 border-transparent text-slate-500 hover:text-slate-800",
          )}
        >
          <LayoutDashboard size={16} />
          <span>Painel Geral (Dashboard)</span>
        </button>
        <button
          onClick={() => setBasesSubTab("lista")}
          className={cn(
            "px-5 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap",
            basesSubTab === "lista"
              ? "border-b-2 border-blue-600 text-blue-600 font-bold"
              : "border-b-2 border-transparent text-slate-500 hover:text-slate-800",
          )}
        >
          <Database size={16} />
          <span>Lista de Candidatos</span>
        </button>
        <button
          onClick={() => setBasesSubTab("novo")}
          className={cn(
            "px-5 py-3 text-sm font-semibold border-b-2 transition-all flex items-center gap-2 whitespace-nowrap",
            basesSubTab === "novo"
              ? "border-b-2 border-blue-600 text-blue-600 font-bold"
              : "border-b-2 border-transparent text-slate-500 hover:text-slate-800",
          )}
        >
          <UserPlus size={16} />
          <span>Novo Registro</span>
        </button>
      </div>

      {/* Dashboard Sub-tab */}
      {basesSubTab === "dashboard" && (
        <div className="space-y-6" id="bases-dashboard-view">
          {/* Main Hero KPI Stats */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-blue-50 text-blue-600 rounded-xl">
                <Users size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-slate-400 uppercase tracking-wider block">
                  Total de Cadastros
                </span>
                <span className="text-2xl font-black text-slate-800">
                  {bases.length}
                </span>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-emerald-50 text-emerald-600 rounded-xl">
                <CheckCircle2 size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-emerald-500 uppercase tracking-wider block">
                  Convertidos
                </span>
                <div className="flex items-baseline gap-2">
                  <span className="text-2xl font-black text-slate-800">
                    {bases.filter((b) => b.status === "Convertido").length}
                  </span>
                  <span className="text-xs font-bold text-emerald-600">
                    (
                    {bases.length > 0
                      ? (
                          (bases.filter((b) => b.status === "Convertido")
                            .length /
                            bases.length) *
                          100
                        ).toFixed(1)
                      : "0"}
                    %)
                  </span>
                </div>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-blue-50 text-blue-500 rounded-xl">
                <TrendingUp size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-blue-500 uppercase tracking-wider block">
                  Interessados
                </span>
                <div className="flex items-baseline gap-2">
                  <span className="text-2xl font-black text-slate-800">
                    {bases.filter((b) => b.status === "Interessado").length}
                  </span>
                  <span className="text-xs font-bold text-blue-600">
                    (
                    {bases.length > 0
                      ? (
                          (bases.filter((b) => b.status === "Interessado")
                            .length /
                            bases.length) *
                          100
                        ).toFixed(1)
                      : "0"}
                    %)
                  </span>
                </div>
              </div>
            </div>

            <div className="bg-white p-5 rounded-2xl border border-slate-100 shadow-sm flex items-center gap-4">
              <div className="p-3.5 bg-amber-50 text-amber-600 rounded-xl">
                <Clock size={24} />
              </div>
              <div>
                <span className="text-xs font-bold text-amber-500 uppercase tracking-wider block">
                  Pendentes
                </span>
                <div className="flex items-baseline gap-2">
                  <span className="text-2xl font-black text-slate-800">
                    {bases.filter((b) => b.status === "Pendente").length}
                  </span>
                  <span className="text-xs font-bold text-amber-600">
                    (
                    {bases.length > 0
                      ? (
                          (bases.filter((b) => b.status === "Pendente").length /
                            bases.length) *
                          100
                        ).toFixed(1)
                      : "0"}
                    %)
                  </span>
                </div>
              </div>
            </div>
          </div>

          {/* 2-Column Bento Breakdown */}
          <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
            {/* Column 1: Performance por Base */}
            <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
              <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                <Database size={18} className="text-blue-500" />
                Desempenho por Base de Origem
              </h3>
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs">
                  <thead>
                    <tr className="border-b border-slate-100 text-slate-400 font-bold uppercase pb-2">
                      <th className="pb-2">Nome da Base</th>
                      <th className="pb-2 text-center">Registros</th>
                      <th className="pb-2 text-center">Conversões</th>
                      <th className="pb-2 text-right">Conversão (%)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-50">
                    {statsByBase.length === 0 ? (
                      <tr>
                        <td
                          colSpan={4}
                          className="py-4 text-center text-slate-400 italic"
                        >
                          Nenhuma base registrada ainda.
                        </td>
                      </tr>
                    ) : (
                      statsByBase.slice(0, 10).map((b) => (
                        <tr key={b.name} className="hover:bg-slate-50/50">
                          <td className="py-3 font-semibold text-slate-700">
                            {b.name}
                          </td>
                          <td className="py-3 text-center font-bold text-slate-600">
                            {b.total}
                          </td>
                          <td className="py-3 text-center text-emerald-600 font-bold">
                            {b.converted}
                          </td>
                          <td className="py-3 text-right">
                            <span className="inline-block px-2 py-0.5 rounded-full font-black bg-emerald-50 text-emerald-700 text-[10px]">
                              {b.conversionRate}%
                            </span>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Column 2: Status & Product distributions */}
            <div className="space-y-6">
              {/* Distribution of Statuses */}
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <Target size={18} className="text-blue-500" />
                  Distribuição de Status dos Candidatos
                </h3>
                <div className="space-y-3">
                  {statsByStatus.map((s) => (
                    <div key={s.name} className="space-y-1">
                      <div className="flex justify-between text-xs font-semibold">
                        <span className="text-slate-600 flex items-center gap-1.5">
                          <span
                            className={cn(
                              "w-2 h-2 rounded-full",
                              s.name === "Pendente" && "bg-slate-400",
                              s.name === "Interessado" && "bg-blue-400",
                              s.name === "Convertido" && "bg-emerald-400",
                              s.name === "Não tem interesse" && "bg-rose-400",
                              s.name === "Sem retorno" && "bg-orange-400",
                              s.name === "Contato via Sales" && "bg-purple-400",
                            )}
                          />
                          {s.name}
                        </span>
                        <span className="text-slate-800 font-bold">
                          {s.count}{" "}
                          <span className="text-slate-400 font-normal">
                            ({s.percentage}%)
                          </span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                        <div
                          className={cn(
                            "h-full rounded-full transition-all",
                            s.name === "Pendente" && "bg-slate-400",
                            s.name === "Interessado" && "bg-blue-400",
                            s.name === "Convertido" && "bg-emerald-400",
                            s.name === "Não tem interesse" && "bg-rose-400",
                            s.name === "Sem retorno" && "bg-orange-400",
                              s.name === "Contato via Sales" && "bg-purple-400",
                          )}
                          style={{ width: `${s.percentage}%` }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Distribution of Products */}
              <div className="bg-white p-6 rounded-2xl border border-slate-100 shadow-sm">
                <h3 className="text-base font-bold text-slate-800 mb-4 flex items-center gap-2">
                  <GraduationCap size={18} className="text-blue-500" />
                  Distribuição por Produto Acadêmico
                </h3>
                <div className="space-y-3">
                  {statsByProduct.map((p) => (
                    <div key={p.name} className="space-y-1">
                      <div className="flex justify-between text-xs font-semibold">
                        <span className="text-slate-600">{p.name}</span>
                        <span className="text-slate-800 font-bold">
                          {p.count}{" "}
                          <span className="text-slate-400 font-normal">
                            ({p.percentage}%)
                          </span>
                        </span>
                      </div>
                      <div className="w-full bg-slate-100 h-2 rounded-full overflow-hidden">
                        <div
                          className="h-full bg-blue-500 rounded-full transition-all"
                          style={{ width: `${p.percentage}%` }}
                        />
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* New Cadastro Sub-tab */}
      {basesSubTab === "novo" && (
        <div className="max-w-xl mx-auto">
          <div className="bg-white p-6 rounded-3xl shadow-sm border border-slate-100">
            <h3 className="text-xl font-bold text-slate-900 mb-4">
              Novo Registro em Base
            </h3>
            <form onSubmit={handleSubmit} className="space-y-4">
              <input
                placeholder="Nome da Base (Ex: Junho 2024)"
                required
                value={formData.nomeBase}
                onChange={(e) =>
                  setFormData({ ...formData, nomeBase: e.target.value })
                }
                className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
              />
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Nome"
                  required
                  value={formData.nome}
                  onChange={(e) =>
                    setFormData({ ...formData, nome: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="Telefone"
                  required
                  value={formData.telefone}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      telefone: formatPhone(e.target.value),
                    })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Email (Opcional)"
                  type="email"
                  value={formData.email}
                  onChange={(e) =>
                    setFormData({ ...formData, email: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="CPF"
                  value={formData.cpf}
                  onChange={(e) =>
                    setFormData({ ...formData, cpf: formatCPF(e.target.value) })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="N° Oportunidade"
                  required
                  value={formData.numeroOportunidade}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      numeroOportunidade: e.target.value,
                    })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Semestre"
                  required
                  value={formData.semestre}
                  onChange={(e) =>
                    setFormData({ ...formData, semestre: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <select
                  value={formData.produto}
                  onChange={(e) =>
                    setFormData({ ...formData, produto: e.target.value as any })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                >
                  {uniqueProdutos.map((p) => (
                    <option key={p} value={p}>
                      {p}
                    </option>
                  ))}
                </select>
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Metodologia"
                  required
                  value={formData.metodologia}
                  onChange={(e) =>
                    setFormData({ ...formData, metodologia: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="Forma de Ingresso"
                  required
                  value={formData.formaIngresso}
                  onChange={(e) =>
                    setFormData({ ...formData, formaIngresso: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <div className="grid grid-cols-2 gap-4">
                <input
                  placeholder="Período"
                  value={formData.periodo}
                  onChange={(e) =>
                    setFormData({ ...formData, periodo: e.target.value })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
                <input
                  placeholder="Nº Matrícula"
                  value={formData.numeroMatricula}
                  onChange={(e) =>
                    setFormData({
                      ...formData,
                      numeroMatricula: e.target.value,
                    })
                  }
                  className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
                />
              </div>
              <input
                placeholder="Curso"
                required
                value={formData.curso}
                onChange={(e) =>
                  setFormData({ ...formData, curso: e.target.value })
                }
                className="w-full px-4 py-2 rounded-xl border border-slate-200 focus:ring-2 focus:ring-blue-500 outline-none"
              />
              <button
                type="submit"
                disabled={loading}
                className="w-full bg-amber-600 hover:bg-amber-700 text-white font-bold py-3 rounded-xl transition-all disabled:opacity-50"
              >
                {loading ? "Salvando..." : "Adicionar à Base"}
              </button>
            </form>
          </div>
        </div>
      )}

      {/* Candidates List Sub-tab */}
      {basesSubTab === "lista" && (
        <div className="bg-white rounded-3xl shadow-sm border border-slate-100 overflow-hidden">
          <div className="p-6 border-b border-slate-100 flex flex-col md:flex-row md:items-center justify-between gap-4">
            <h3 className="text-xl font-bold text-slate-900">
              Bases a Trabalhar
            </h3>
            <div className="flex flex-wrap gap-2">
              <div className="relative">
                <Search
                  className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400"
                  size={16}
                />
                <input
                  type="text"
                  placeholder="Buscar por nome..."
                  className="pl-9 pr-4 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500 w-48"
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                />
              </div>
              <MultiSelect
                options={uniqueBases}
                selectedValues={baseFilter}
                onChange={setBaseFilter}
                placeholder="Todas as Bases"
                allLabel="Todas as Bases"
              />
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={produtoFilter}
                onChange={(e) => setProdutoFilter(e.target.value)}
              >
                <option value="">Todos os Produtos</option>
                {uniqueProdutos.map((p) => (
                  <option key={p} value={p}>
                    {p}
                  </option>
                ))}
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={cursoFilter}
                onChange={(e) => setCursoFilter(e.target.value)}
              >
                <option value="">Todos os Cursos</option>
                {uniqueCursos.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={semestreFilter}
                onChange={(e) => setSemestreFilter(e.target.value)}
              >
                <option value="">Todos os Semestres</option>
                {uniqueSemestres.map((s) => (
                  <option key={s} value={s}>
                    {s}
                  </option>
                ))}
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={statusFilter}
                onChange={(e) => setStatusFilter(e.target.value)}
              >
                <option value="">Todos Status</option>
                <option value="Pendente">Pendente</option>
                <option value="Interessado">Interessado</option>
                <option value="Convertido">Convertido</option>
                <option value="Não tem interesse">Não tem interesse</option>
                <option value="Sem retorno">Sem retorno</option>
<option value="Contato via Sales">Contato via Sales</option>
              </select>
              <select
                className="px-3 py-2 bg-slate-50 border border-slate-200 rounded-xl text-xs outline-none focus:ring-2 focus:ring-blue-500"
                value={blockedFilter}
                onChange={(e) => setBlockedFilter(e.target.value as any)}
              >
                <option value="all">Verificação: Todos</option>
                <option value="blocked">Verificação: Bloqueados</option>
                <option value="unblocked">Verificação: Ativos</option>
              </select>
            </div>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-xs font-bold uppercase tracking-wider">
                  <th className="px-6 py-4 w-12 text-center">#</th>
                  <th className="px-6 py-4 w-12">
                    <input
                      type="checkbox"
                      checked={
                        filteredBases.filter((b) => !invalidBaseIds.has(b.id))
                          .length > 0 &&
                        selectedEntries.length ===
                          filteredBases.filter((b) => !invalidBaseIds.has(b.id))
                            .length
                      }
                      onChange={(e) => toggleSelectAll(e.target.checked)}
                    />
                  </th>
                  <th className="px-6 py-4">Nome</th>
                  <th className="px-6 py-4">Base</th>
                  <th className="px-6 py-4">Status</th>
                  <th className="px-6 py-4 flex items-center gap-4">
                    {selectedEntries.length > 0 && (
                      <button
                        onClick={handleBulkDelete}
                        className="text-rose-600 font-bold hover:underline"
                      >
                        excluir selecionados
                      </button>
                    )}
                    {selectedEntries.length > 0 && botConfig.url && (
                      <button
                        onClick={() => setMassSelectorOpen(true)}
                        className="text-blue-600 font-bold hover:underline py-1 px-2 bg-blue-50 rounded-lg flex items-center gap-1"
                      >
                        <Bot size={14} /> Em Massa
                      </button>
                    )}
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredBases.map((entry, index) => (
                  <tr
                    key={entry.id}
                    className={cn(
                      "hover:bg-slate-50/50 transition-all",
                      invalidBaseIds.has(entry.id) && "bg-rose-50/50",
                    )}
                  >
                    <td className="px-6 py-4 text-center font-bold text-slate-400 text-xs">
                      {index + 1}
                    </td>
                    <td className="px-6 py-4">
                      <input
                        type="checkbox"
                        disabled={invalidBaseIds.has(entry.id)}
                        checked={selectedEntries.includes(entry.id)}
                        onChange={(e) =>
                          !invalidBaseIds.has(entry.id) &&
                          toggleSelect(entry.id, e.target.checked)
                        }
                      />
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex flex-col">
                        <span className="font-bold text-slate-900">
                          {entry.nome}
                        </span>
                        <span className="text-xs text-slate-500">
                          {entry.curso}
                        </span>
                        <div className="flex items-center space-x-2 mt-1 flex-wrap gap-y-1">
                          {entry.telefone && (
                            <span className="text-[10px] text-slate-400 font-bold">
                              {entry.telefone}
                            </span>
                          )}
                          {entry.cpf && (
                            <span className="text-[10px] text-slate-500 font-bold px-2 py-0.5 bg-slate-100 rounded-full">
                              CPF: {formatCPF(entry.cpf)}
                            </span>
                          )}
                          {entry.semestre && (
                            <span className="text-[10px] text-blue-500 font-bold px-2 py-0.5 bg-blue-50 rounded-full">
                              {entry.semestre}
                            </span>
                          )}
                          {entry.periodo && (
                            <span className="text-[10px] text-purple-500 font-bold px-2 py-0.5 bg-purple-50 rounded-full">
                              {entry.periodo}
                            </span>
                          )}
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4 text-sm text-slate-600">
                      {entry.nomeBase}
                    </td>
                    <td className="px-6 py-4">
                      <select
                        value={entry.status}
                        onChange={(e) =>
                          handleStatusChange(entry, e.target.value)
                        }
                        className={cn(
                          "px-2 py-1 rounded-lg text-xs font-bold outline-none border-none",
                          entry.status === "Pendente" &&
                            "bg-slate-100 text-slate-600",
                          entry.status === "Interessado" &&
                            "bg-blue-100 text-blue-600",
                          entry.status === "Convertido" &&
                            "bg-emerald-100 text-emerald-600",
                          entry.status === "Não tem interesse" &&
                            "bg-rose-100 text-rose-600",
                          entry.status === "Sem retorno" &&
                            "bg-orange-100 text-orange-600",
                          entry.status === "Contato via Sales" &&
                            "bg-purple-100 text-purple-600",
                        )}
                      >
                        <option value="Pendente">Pendente</option>
                        <option value="Interessado">Interessado</option>
                        <option value="Convertido">Convertido</option>
                        <option value="Não tem interesse">
                          Não tem interesse
                        </option>
                        <option value="Sem retorno">Sem retorno</option>
<option value="Contato via Sales">Contato via Sales</option>
                      </select>
                    </td>
                    <td className="px-6 py-4 flex items-center space-x-2">
                      {!invalidBaseIds.has(entry.id) && (
                        <button
                          onClick={() => {
                            setSelectedEntry(entry);
                            setSelectorOpen(true);
                          }}
                          className="text-emerald-600 font-bold text-sm flex items-center space-x-1 hover:text-emerald-700"
                        >
                          <MessageSquare size={14} />
                          <span>WhatsApp</span>
                        </button>
                      )}
                      <button
                        onClick={() => handleContatoViaSales(entry, entry.nomeBase || 'Bases')}
                        className="text-sky-600 font-bold text-sm flex items-center space-x-1 hover:text-sky-700 bg-sky-50 px-2 py-1 rounded-lg ml-2"
                        title="Registrar Contato via Sales"
                      >
                        <PhoneOutgoing size={14} />
                        <span>Sales</span>
                      </button>
                      <button
                        onClick={() => {
                          setEditingCandidate(entry);
                          setEditFormData({
                            nomeBase: entry.nomeBase || "",
                            nome: entry.nome || "",
                            telefone: entry.telefone || "",
                            email: entry.email || "",
                            cpf: entry.cpf || "",
                            curso: entry.curso || "",
                            produto: entry.produto || "Graduação",
                            numeroOportunidade: entry.numeroOportunidade || "",
                            semestre: entry.semestre || "",
                            periodo: entry.periodo || "",
                            metodologia: entry.metodologia || "",
                            formaIngresso: entry.formaIngresso || "",
                            numeroMatricula: entry.numeroMatricula || "",
                            status: entry.status || "Pendente",
                          });
                          setIsEditModalOpen(true);
                        }}
                        className="text-blue-500 hover:text-blue-700 p-2 hover:bg-blue-50 rounded-lg transition-all"
                        title="Editar Candidato"
                      >
                        <Edit2 size={16} />
                      </button>
                      <button
                        onClick={() => handleDeleteBase(entry.id)}
                        className="text-rose-400 hover:text-rose-600 p-2 hover:bg-rose-50 rounded-lg transition-all"
                      >
                        <Trash2 size={16} />
                      </button>
                    </td>
                  </tr>
                ))}
                {filteredBases.length === 0 && (
                  <tr>
                    <td
                      colSpan={6}
                      className="px-6 py-12 text-center text-slate-400 italic"
                    >
                      Nenhum registro encontrado com os filtros aplicados.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Editing Candidate Modal */}
      {isEditModalOpen && editingCandidate && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/40 backdrop-blur-sm p-4 overflow-y-auto">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden my-8 animate-in fade-in zoom-in duration-200">
            {/* Modal Header */}
            <div className="flex justify-between items-center p-6 border-b border-slate-100">
              <h3 className="text-lg font-bold text-slate-800 flex items-center gap-2">
                <Edit2 size={20} className="text-blue-600" />
                Editar Candidato
              </h3>
              <button
                onClick={() => {
                  setIsEditModalOpen(false);
                  setEditingCandidate(null);
                }}
                className="text-slate-400 hover:text-slate-600 p-1 rounded-lg hover:bg-slate-100 transition-colors"
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Form */}
            <form onSubmit={handleEditSubmit} className="p-6 space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="md:col-span-2">
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nome da Base *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.nomeBase}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        nomeBase: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                    placeholder="Ex: Junho 2024"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nome do Candidato *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.nome}
                    onChange={(e) =>
                      setEditFormData({ ...editFormData, nome: e.target.value })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Telefone *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.telefone}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        telefone: formatPhone(e.target.value),
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    E-mail (Opcional)
                  </label>
                  <input
                    type="email"
                    value={editFormData.email}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        email: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                    placeholder="exemplo@gmail.com"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    CPF
                  </label>
                  <input
                    type="text"
                    value={editFormData.cpf}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        cpf: formatCPF(e.target.value),
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Nº Oportunidade
                  </label>
                  <input
                    type="text"
                    value={editFormData.numeroOportunidade}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        numeroOportunidade: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Curso *
                  </label>
                  <input
                    type="text"
                    required
                    value={editFormData.curso}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        curso: e.target.value,
                      })
                    }
                    className="w-full px-4 py-2.5 rounded-xl border border-slate-200 text-sm focus:outline-none focus:border-blue-500"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 mb-1">
                    Produto *
                  </label>
                  <select
                    value={editFormData.produto}
                    onChange={(e) =>
                      setEditFormData({
                        ...editFormData,
                        produto: e.target.value as any,
                      })
                    }
                x���rI� x�Wb��#2� �Gvu� �,t�� �v!���Q�x"{�9�a�2[�+��-R"#�2�>6�d�`>aU����G H�+]2�ps{�������2��0���M4��:K'�4e�˥�lz�������l2�GK�);��Q��?KE���ÕVƗ�R1f'�pV�Y�&�xi�Mb�$K��x�	8>]�8Kʸ�y��@c׳I��Y��g�Y��q4�v�=���u�{�M�$��������G�<�
��[��TU�,j���B�-qK�̳�Qr����Γ5��q�҉:N��G���<)���t$���z|���	�y?�E���`�y���$��VY^M�k؅��2�ǿ�%y<
~�����e����2����Y�&�g��
ucĄ��+�r�Tڽ���~�O�_�̨�5`q���Ӹ���W���?���ٗch���L�݋�O�F�gC�VN�<�>|iH)��N�ϋ�����f�I�gE˱�Ǘ���k��'��艳�Q�v&�y\^�y��Q=������/��>/����e��6��,����l�v&��|ix�t�Le�Yt*�rV�Q�6��(Ee6�%D���o�KCD�+�XT�hr��D�[�&��A�[g/�>M����W���)�p�FY�9y����lr�e�u��sU���?d �1Kd`@~�\U�C�X�|�"/�%.��$b�Q|4vR]50�D'�z��D��YYf��}���*��'i|�~�err�H�N���#6-u%��6���R4������MH��G 	\G^�@v
�	��Q��N�I�$J����*;fN&�����	,�`���z& �˘�P��E<N\j��X�P�,��$\��y4)D��!�y��'�+0�a�Fy U�矞bv<N�l�()��4�]�Y4���8�BHqZ�O�N�P����ȇ�����!R��c�*�\��A6��Iy4�|�H�:���=�`�鰁|��F
-E��������C�j"k��2��4��d��nԚ�۳�,6���HtO�}N>2�I��d������W]V�j����'���2��8����ţ�I�_��'�8f���ӱ3o��"�r15�]s�~6���k��Hã�$8I�p�/�i����X�8*`�zb� �qq�ѡ�u�n�2u� �\�30jj��]�_���x�a;����8�&;a2�"��8+aC9IN׮�O{ �ыL�`�q��8g�G��XT�ɵ���U?)�_�u���,!�S�t��n1�.����s/�1Я��嵓�z�t4�M(�O���Ŗ�x<����0&�(���ٚ?��0�d�S}Hߏ�z��~ �=���U��f�{,'� GY��0���{�OF=���l!� S��^@U;��ɏ�z~o���EW�@g�����t�%�� �bC�k�B�W潟Ɠ����fG�ǜ��R�l���%'�>�:��a,Q�(Op,���E���¸°r��X+͂Y6�@�	vV��%��ڷ��{xDr�sC����j���gq`�oL�1�wx��%��J�(��%���$ ��F��S�q�X�w���#0�����Z�*;Ȣ��A�D�Q���\�>���E����)�8�Gw������=�b��i���I|�J6��5�����"'��z����<4���� �;5o���$)J�I�C��*m6�:��;��i`2��1�҆����=z?�^u�P7�ЛI|��PN����U6�_�a)LN��^�������f�hYS��i��<즟�� �� �Vm=[.*͠�j��d6r��o�6�d�ؗ�=����MIlǟ�<�ω��)b���D1It��B�nL71��<�!'������p]d�:�����\�o�y�w��>ϒ�Sӻ{W���xy��9�hW�z�~>���8G�O�k����T�w��d��tM�Q���-T];�9<BؒVn�(6�3`j�Xj7~H"�ۀ�.*�&C֕p����4[+	�E���"JJ�F[�ʧ� uGǋls�ի�̓��7����W���7w�ll��l�lug4P?a�7-�8����|�w���d�J��yV�R�*��QFV� 	�;�aM�6 d�_ $h���]����+���k�`�?����03�j<Z���1\VsÆQ9<c]�v>QF�9�Ҹ��~�[ކO,�t{9����F�R�M>��ƣ�����S�p� �fE���X*~q�p-�J�`ө����x��yNOȋ�J�:vZ��=@n�v��p`F>��p���ާ)�NI�3.��N���۔i����ƕ!�J���U�s�ױ�"g+�@	b���~%~[ 6|�,P�Q><;���`�ԫU�ӡe�ܾ�;/�B��2�8<z��:�"�̓��$��U9q��M�)�1���4�5���ўZI�e��*[�&�Ë.�|e�Xox����|n���6W�]q������ �[_	p;�Y/�7��^�l����>��\y�W �r]�c.�yB��)Y^f�Ôe9S̟��ɿ�jDn��Ƞ���"C,1���/�;�c�� �T�"E>Ego$�|��1��H��B�F ��d~�/^��&���N��\lV��Q�e� [�U�e-�U�>>��ߊN�j��kl��w�B~����@-r�c5���L�Ws�9P_�z��$��x�03T����zU7fShO~ �O�믍.G�V�Ǯ����4�S�1J�F��M�S�&�9q���JdWy*����Z�B�t�g�.�G�q�K��aĺ&h����w���ض�P塵�/�M}1���5�A��G��~)`8�ݽ�V5�=E�^d@��I����1�c��g<�Q��6F�Y��s�	���(#B�V|�RU��q��mNO�6�����b��k_j��˿�Z&6E<C�<uk:��Nqu�J��$��l�&C�r�* %�zUY�]������1 �P�=�!��l���.u�U���UWq� 5[�f��
f:c?~�#�/�w�&�P��@I���{����l���-O��uQ[&��F{�����7�?lln�Z"X��?��0�݅�5���!UD�y���mΕ|�.��X��B�J�S|�/[�࣡BO�|��i�FzCq%^�+E��$�&L�DA�FrT_�R��	��nM�45�GP�ꄥ�f����aH��6R�b�~܂ҜdHrH��{@�h��`�z�"��r4�p�:�|l�/>l�E�c�v�$g_]���Q25�1<�AX���5��3@����-W� Y���=-��Ģq8�����>zze��75�ڌ-���y� ���`^@7� ���ʛ ����i*�
=�Ɉh�������k#��ܣ��	複���H/��A].x��#��
.r�Y��F�� �aP�fs+�+�C����Zr0|3&���A��h{���B�8��qPk��m�D�|���Kr��O�&���B׆��$���;��P��C�bdC�b#�ծwzU�H���+j���1hOD�/qGƥB�'2��r�U��9�7��8�t���8�t���$�M^t��*�l�R�ca�`,�2:h�NY���y���k΃�V�=Ņ�kg��ה!9�w��|��Ŷ8%l{|2�lH|��� %�P*	�~��3�hp�m_�;�ʫ�����K ��B;c�BZ�t�^�~}���^�N���h�A���~ߋ���ʑ�B0�0'"�7��J�m~Φ���2d�W�hZ�!�E�]�c��>�W�����@��@>(�{�#�z^
���	�ߑo�l�U���B�P��@1��Jݏ�Gmegߠn;�ƺH�U�	B4z�ޏ�^�f���!����6V�
]�0UO8S0A䇭#6
��y�h���d�=�(�o��"0�R(҇��ZC�S7�o�O�3�m���P�2FLEo�}��ښՌ��4�1͊p�O�RT��郞�� k5s���?{�5eB��6��퍭0�D�8�O����3�5E��7X|��PSZ���/�&�Yr��M=�2�ӟD����*���Qr��VǷd��@�1�;���^�a�jmh�)6V{���$�-l�1RS�������$E����k�!7Y�X�Cxx���5o���q�\7	�����mCP�R��Y�o����wө�Ⱦ�_�4���^�sFy�s��3Eky�9�G�������y�Ea-�,3fT���I4���ѕ�Q�.V���ꔚׂ���ZITasI_�
.�4p��*d�4ze�.�g�k��l��"Z�����0��Ȝ؎�F�T	��QVn5([r�&ް��^�Q?OV��g3�<) J��BK	(��E���%�ԀGM���'���|tZ��RȆ؂(ArcH�ͨ��}����a�hՂ	���s鰷T@�4����D�@Uݞ�a\�)XoK0'Ͱ�|���ߢ!E�1�[s�Ѽr�X�AZ���̗��Ոim�u�X��Ҹ��j���+O�/���>��sO�(�9=A1S�����&������|(��ݳ�l�i?�@(�C~bh}蒓Is>,Ou��"��wL��%�-c�[�9�l�u�i���%���R��p���� ��!��
��?����h�򭪜�5bSXgpj������>�ֈ�G"����K�9I�_]���yK��FaY��9�\Y��
[��2�%�@/x(�F��o�����Y˷���nA��N�V�B��1mΊ2K{�;�6h�4�i}nm&�}�w���������������o�V�xad o�ϵ���Ns�8���6��-ꐊk�ӝ�_����s���#�0��5��"U�kQlkf�x���;V�&=@/BV�*�`�#X�u�-ǜ�HL�jfu� x|M�PnghO���kޥ�]��� �i q�(�P'���ܠ�� X�:*ǧ�w�f� =)��3yV����9m���h���l��0��+�d�n�J�~Y$��t51�ԑ�nU���4LܾZ��(��g1݈�E0��Mp��d��(�x��(�&e��b��i��#��D��_.�r�:��nhs�?%Ⲻ8+��5b�t��e?��89 ��Ve<YEKlڊ����h]�? `�V ^�g�RHN�HlAW~�yq����7�����-謤���r���'�U9O���M�qXl�V^SX�W3�2J7�ل[:�G�bq?�y%�rR�bL��/��"�ݺ'Q��TO��d���[]'�4�KWK��~K�^P��*�ƣ��g�5�q\^����41�.�.�A��r)�ct=�xGyv��6#<9�aON��¸_��Iv�GS��
W������ EL�g�H��#���p����V!G!����~�� �udu��*������H����3�8��{�\B6�-DՂ�V�o�nm�#lz�����$� �]����u���<ߑ+J/�g�<�/����m��:�*`��>�@��8WW��Ga�"-S��ۓ��f����Ϣ<VS�ؙ� |}�i��E���1s{4�U��1�6Η��P�;�|7��0M�E>�U�>|�d�up��\��/��r��7���F-\�6�g�h;���i�ɵX�Ȧ�WWV\�K��X�}���c+����W1�м��my���}�$�8����U��`�2���M�)��ȱ��w����%�B�e΃�o:��{O���o���+�qv8�r�vB���G8�.������;@��q����a
��ީ�H�a�����0'7�-��w�	��}���;�<��y�{�	Ջ���B���q&�i������e���ٓ�\P�3NG�w�'`�'�1�}�$��W����te�r�{�� F���{�f��`L8u9�v�?�>�ڭ�ֵE�1l�)�ϖ�7`+7���eAZ�-� e\<後R�گ;�s��v��̠�ݮ��e*���7pM7^&������F��_	EJ�'�Q=�VQ=����?	`/�9ea���.��ΑrI��xLOA-�P����ݭ����o7!E��n����#g��8_������ۗ�5�p����Þ�yW��N��EHu��I�3e��u=��E;�=��5���渍�c����!?y�N�������5s�M�UheL�Є�>A[��n����
����MX�+z�k3��π:���,�W�t��1���2���t�������s��{�=�K��k������+3	]p��g7i-���w<q#�Tw��MQ*�H�L��&��*m9y����~w����^�����0.(���}CR"`|���;g��9cl��kVGp��j�#:̿�}��͢��·_׵��5�־|llA�EFq��v"!?e�OyP�v��
���u�-�Z�c#��V�
����-�f��"��55��Tg5*6�5ԧ�۲w�i��T'��$�#����Z��6��'�����o�oK���s�lӜz*�:�e;ȣ�(=�i�*���jGEUo���iTk���B8y�#�Zubk	�Zfӥ��l�c*�O����E��~�.��H]<�z��C .�&�j��n��t�o�R���NT'�^�Z9���x],=�U��j�26o��#VW�l uhoz=K�d?ď
��X�&������=�g��;����fk24��^�n��r�
m��y��1�}>�pǤ8Q걲io����c�����e�� ��Q(G"�������CwU/\n����Y�;q&�f։����s�Բa�E&1��6�=��=���/s�mg�m��v3}_����6L��'�h3����"<����sO���y��^�]TX�[��c`�;���`��0�!���HFSމ>�-Q�ȴ���w���,�Fxʼ��%[A��LPB�D�3g�~�K�1�ԫm�����ry6g�Xɍq�����2�[:I\2�'u'�`���5O0w�6�vQ�]����(׏�z���X� ��[�=oE��p`Q���
��ܬ�믃��*��z�h�6<���Yh�/���
w���0&���L;Q�~i�㪌B���������<�38��8���ǅ��TP	]ĢUԇ>d�F�b,�Ӫȵ��ʳYi]�`�c�C��cnBK�rw�
l���ltE�����+&T����ӌ�k��wYł�~Hd&R巵!x�s/�����f@����Q��P+ӠJ��<����a��e��#��۶�zT�H�v�nC$[be�N�M�j���T��"k�a���Xݕ2L\���S3�G"w��L�ѿIT�	ˌV�t�]K��C|�w�V�^g�����n�򦛳0�m���]�I3X*vL�3�ժkrt� �_�?	mc��^���P�����,����yn�џ��y��� g_ūڲ�
Y2ߦUd���F<R#1�Km���i�]�+^�Q�T���q^'���QH�JyK_���WR�ŏ����P0�+�Z#T-��>�\֜�=�Hjn��ҥ�\�Ss{��s�f�\�tVͭ�+N��7W�T�Uߠh�\rs.��4[fGM֦yIou���
YɅ�VZQ���Њj�S�VTR�V���\���ՉϨP5=��os�z5liՆ4�Z�J{�K����P�FU���j��ɽ��I���ͫ�h�%�*�Y(�����ӌ�9�� ٚ)�8!t����f��j�uk5T����B� :�9� gw�hyƂ��8���o��A�9�����<��ƭ�,KZ���D�iVeճ��e�U�S�B�!��]�J
�J`�Qq�ИjU��\a��C����#[6(�=�kɾ3>�,�X�~�� a�O�\�]�'��À{O�fc6)�`tx�=+�!��x<H�}y7�'�\�j�$��v���\*j+�]6)pGQz=��P���)`,Ѵ[���xh��q4R��z��H�i���Q��\�ʮo��B��ҡc�v]�2������� f\��=Xt=�J���J5I���.g"I1lg�u����c����?Y�S���]�� &���"zi+�a��7�ptm�z���K��M
�W t�{�r���-@:�3m�B�r��	X��/7�zݒ��*�Mg��d�-�f?��o��IkQ2�T�uPd2ˤ0�n��~H߭�K�b5� t��?V�����䂇5�|����I؋��w�5��d+�x��߆z@T�;��̻�D�l�^:Et�
s99�\^��&�dH��!��G��jz���EfqW=��kUY0�&�Mp`g�^N��`��;��a��Nҫ�k��|c��a�2��I%l�q����埸>o�ѻ���1-�kw�[�V��-�ڵ�AG�Q��d�I]�>�f�E���Հu��WO��]���1��˅�������0ZP�2B6��S��~�V��}Z���X,N��q2����U��PR=�wN��e6��rA�6.���~(d�������`��˘PC���UF{�L˥o;�v�/����N�>��P. ���#��ˆ�/ͥ�½y��d6��}�A��w�pM�v���s�g<�!�t�?L�[�(&I�x!4T� i�1���tx���f�d��	g!9핝����.��l��|;���]\��Y���.ػ"�����@���6D�zH�B�|JG=*���	�J�ذF�V42�;.)��u��ɖ"J?5�J@��kg7G��MZq��7IJ��Po����K1��
����,3��{Z^dS�Ք�.f,z7/��2�����~]{�h��GSjJ�@�H��4��$�2)��z�)�zgZ��~�J��Y�G	��W���r�$X1��ڂ53����.�K��� �G����E��kv|+p�7�lgg�Y����Oa3�:���!�H�QV-m@;#ѿ-�`5�����4*��5ܳ�B/��#���P�sM�-9�4�=ԯ^f�o%Q��O?�6�ؠ��I����$��P�u<κT�.c��ʽ�FD�&x�?��GZyY�6)Ϻ�齉��/���v�Ev�0�(G�*���+&����S��q��quqRZ����)g6��4�cY�qW/b�,�CLG=/H��N�p���\�D��
��	S�U�	)��C�V�A�~co]���V՜�#HF��i��m�����q���>�k�G�c����z!�nA��S��{���_TƂ`e�4�&�(�>C�Y���|2R�ub-*p����T<س����m�<^�=J�ʁ�	m���N��#��s.�l/^�ʬ`�ʐ�Y��p7^��ۄ=��x���~s������C��`�?��S���V16T}���Ӆ��_uV������Ʉ; Q�V��α���2��ԉ��h�yc�b鵝W��U�a����c��1b1�Sb�T�D3]*�;ZS/�e�L\z,YB�Nu]b�իq%�R,4*щ���=c��*��J@��~CL�7t�m�_e/�W�0�G�-�L�O�+ �u�~o�����d�fc�Җ�%��	�������2��Pv���5��1��$۞����j��[�9@��U�!�=��&�C##CK�*֋D������/Zq�/�����|Aqf��,kǮ�Q��Ӿ�ΙT�Ѱ��׎qsǾ�s7���y�7�@�1�6L@��˒�:�;��<Pȷ\ë���P=h�]8�W=����J�+�^4M/��-���(� k�x|>_\q�����{�g�)W	r��@�.��8r���Y�������N8���ާ)�N	��B��B�oG�3�C�Ц�0ʷ'ꔦ�0fb� �v��������1� �=���~,�]Q�'���I[tg6)�������.|W/xa*C=�AE�d4��2��S%�H��zs�/���i�wC�S�ґ������ƈ$���hζ�]��?) l���0p�Q���4��ɫՀɵPX��8 ּ�/�������߮W���A��e��i�:�`u8K�o?ML�]��O\P%$��T�T��d9BAW�4�6#OB��I'rn�?F���M�|2v��x�(f������Q�ق�̊��뺩QW!5�X���U)��&�d>���C@�צ��=B�cS�����k)�����ƭ�^�V�I����!�j��[�����ۢ���fW������z�d\Q�&�ݺ��Y0��8e���N��Gb�:T�?Twr�m�M?ȦѾUR�A����O��ĨI���{�E� �r���j!�'��C*'�]^�{ȁ�ו�sAp`�]�&�A$��$�� ~�+�,�R�K�����\�s�(�/�U͈���F ��ݿqT|������9����i=ꢂ1M���T���ڈ�:����B�����;0��B��@8d~��D�Ɉj�8F$��o����	,���R��.U�!���=c��U%�-P�Beڴ�t5k��}��@�+�o�R�0�Q��!�.Ȭ*�����ˊ�H��r��b^��3���|�����4�C�*Ud{�I���B��_���X��FSg3�	� r�2q���iJ�9D�36F��ӫ:��1���㣎.����N'�O$N��cR��{���u�'��=-��-�URF�Hfw�����Ъ6:�[V�%���*O�!�Nӫ���=颍{ O�VQm�wY.
Z'o��!�s#; ���G3�M+0��PI�,T������r����5����ku-�V���}c\�r�c��7��y��Վ���:(:�3Yw�D�>�;��,�ao��ƋO��;��������m�fy�γ�f���Eh�Ƕ1<�0� ��Ϩ���?���hL[~:K�`}��,)#X�2���j�A�����r�a�cr�� ���'�d�D����_������_�/�1ɀ�`>BC�8��G��з�EbUe6�8h���,���b0O��>7x3u�������ZR��9�g���7:���l���>䉗�Q�")�7E\�D��JEOp���a��Bx�d'H��ٌ����⎇�}���sR��ЕO:MJq~�}��{ɾe���T�%$�W��W!��O�ZpXa�">�MF���De|�`C�y$0C��a�B2���	C#��{1d�='1Ѕ������"����>/��N����iY�O�f�Z)2�4@d_Q)���c ?⯃��O�<�!N@���ۃ�=��O�b�ډm�� U���a�T��Pt��@�iz�c��3��l����z��t�3�Ye_��@T�e��
�����f��� ��-�\ ��ׅ*N,Ȍ��u���zas$���$A����F	[���A���ޙ\�x�L����X�g;(���PT�	.�	6z�%�.8>�ȼ`{x���'@�؈/������˿����1K&]dWz�	��'��1+/�#\��1D`�")�C�q�7���)����Nf�VBqy4��	��"�[}�q�[!�r�\N��O�0˿��n�����(#fҝ�P~����H�i��'��nln�l��]��o7v����Z�@ir�p(M#L��_@9x{��V��{����x�-�e{�o6^�&6�l�lA�~�圾ؒ(5��hz59��`Ɩ���[z�o�9�����6��~�7	�����bo�����`l2���,�������{�yE�>��V�j��y������mnl���_������|0�m�}�ͻ����(�� �Ӡ�"}Ե(	%�#o��@S��s��}4!����$��e�ff�(�	�1��֭׾d_�˿�Z>���� 6aŚT�AZ��KTbu%)��T&p�"��it2�Zzj����Dq�����z�:pN�Uj��W!��n����>�t%�{��w ����Ҫn�\8c?݆3��*��|>'��h�����)�i�;{۶p��R'�V�f��[0���S���z�UF�|Кz���B�@+�e�ÊX���zS"�
�1'��c����-ӱ��[0�v��Z�"�_�xm]Ѧ<���uץ��U#�<� ��@Vy�e�G�v'�IcE[������G��Qx��ʾ�4V�3`�����a����9�/������ӸD��+{h�h�67^m���x���~csw{��s�8�8+^�	Gٗ>�,�ݦKUF=�a;���'γ�0���r��|_�d3i���D���	��0��+��M~�y�o�8�C�K'�^���k[.�i��y'�#>a+Fb�Zc��+���	.�Sa�h�	0�>�F�,CF�b��dI���J���5���R����O�1��D�js�5��E���m��}���^�qn���a���5�,�.�t{O�|�55�ͦyhe�YʞM����z��6Bb�כ+Z��Q�I�L���@Qp��"��b3Y�c��7����N
��%�9�| q_���{xMC�UJ�E��Dj�_� �1��)~�[�F�^f�p%�*|8���'5�A�����!���׳L�d-�z�Ʈ�++C�"��Zh��T�3���}`��Jӌ�gFg��螺���=�{�Y��?䑀�
x�p�Qw�<�$��DL\�����#N���*9�"ذ��+D��W~�����}u��No����n����4+��<N�\�%/�k@n�}Z�}���C"	(�T�>����+�Z�0��x�]� qp �Zm��Gm���~QYD��o��.m�`Sl��6K��k��S�.�*��4𭝯��? P�,��%��檭���͞|�IDrrYJ���hڊ&2$e n��ma��S\���q'=Ŵ��_]�/7�/�ZfD��X��sQi�k���|0��ѻZ��Q� ۾�iw��t ��#l�q2̪V��X�&����ٯ^��A���)W�?������(���^��V���~�!�H��̳��4�}�Tm8k���}����ͳy�8��RGćG�z��o �c���`Avg��{3fs}BI8��;����[�i>�}��9�B�vj�o'W�&���z�����8�C��Vd^]��+Zä}��5Nx�Eߍ0�K�m�R�t�aIL(�b��|̎�>6?!���/?��ӯN�ײ؞�%_�Qh�2h4�d�@��h�g�8�<���,W�P�!RK/�7��&�d�*�+	9QΉ]n�j̢�ܓд���:��/ց�_Km�f����V�76�J��/?5��e��X�֏ɌP[{.�P�g����V�x���ߵZ�xf��=�Z�P���ߦa�y�D�LH�>�i׹�Z���^;��~�W��7ʕ�}�c����Hn�#���ӟ���z�{�����Y��P����|e�ʏs�N�J���^}6�����u���pџ�l1j�W�Z_���t�~���ڃB��n�v>[�ԥ})K8�k\����K,8PՉ���OB�Y�:�4B�OY��\}G���NqEL�SD)��3��Q+T��R�b1�a��C
�k����-���np�Z��R���̰�Q֨[d�����]�>����_ޡ3�UKT"GY��^`��7?���N2�U��T����,9�vfȲ+��5�1B&��!�@�X�a���w:����7� :fo���T�}����B��DW�g�Φ����˔g��t
��a7.�N)�����Y�ޱ����Ʋ&�M� 2��a��������t��k?^B&0�s����/�>"`(a���_<�C?�u��^�A[�2�y���ͳ(/)�r�g�[j(�C|߫���(�}�(���Nޫ�(��sp��O����g{n���g۱~3�	�:͓�0�a����C������h�����&tȁ�!q��^p���d��I�M'	����Ņ��?a�]izὠsY��9�
վ<W��E�j�Po�#�p����f���au�U��;v��Άҕ����!�O��72�������2BE��ƍ����pr�W�=?����M��o���^�<�iY��!_5߹�Ʃ�l��Yxն�7v/Fѳ�G~hc�K�C�+�k|\;��Wf����`�!�B6/9�%��ٺ�L�;��G�H+hޣ@��-/(B[T�5�(�E]Y܄��
Xd�~�err�t�q<����8�An��Uŝ�?�����Z��O�����V�L_ҧ� �>��݌�����
hCf(�|��,�����B.ͫZT\��Zj�VIB��Y�7�l�B�и�H�g����:�C�kJGᝠֲ>�j�w-�B�毪"p6�6�Tu�Ugi_pLgVlLw0d+OR �g���z�C��i�;g�U�����q]�ê��5y�Y���
��kv��ʳWJXX��:�^e`�� ��~�@F�u��3���xX���?Pw$��~"�C:�<Ǵ�1�9��窟�a��~�Mf���ɨ]�յ;�<�r�����wS���*G5S��	.TJ]��졇�H��hl�9���bUD�a�(�{T�*�f����6�Nƀ�>}|�}��������ӧɾ��H���ըo�O��; ����w��n�5��=�8��ǹ�Q�8��m�#�Ɲ^g��d�H���l���6kS���� ߱��
�n<\	W��My�,S���ϸ-d*O���w�˪_;��4_�fI��������wZo
�y�V�>K�~2�[��iQ�z^N��xZ�u��iq���O(�$n��vM�9C�����e���l����2x����()�=�h�Z�ĩ[G(v�F�FTP}dP&)�������P��27�5W�G�Y#�L�U�m�!;ԑg�W��ɉ���L��|pz��@<I&i2��\U��W	�`a/yX`�rniC~���\~6K�N�GJ�!��&�D���oO�D[��dy+���%Db��Q�*����#���9�c��#<�O��Ņ�)��<���7X���U�B�G����vcZ��jw^��t���u�Ak&��AN��Y�K�	j��b�6rP!yiw��ͽ����4��J����J��	�r�V���	@<��M����A��\�ٲ�А�
��K��*�B��ղ�vṛ��&ǀ�Ww�]�����V��d�����$gO�W�����j��?S����q�X4& i�?��JKz�� |�%����G�ϖ�e���v�M7=W�Z$D���a~j��mpS�i�����(��r6�Z���l̍��ｖ%�����qs��,���+>j��׫��Ft��%�0��_K���|2���ilʳ8u�e���XB�%��L/���`6Ee>��@7���r���j#�s���cv��:򓥪�D��"z�8��>�����*e���4�7�e�s�~�������Z�p�RR��O�}�R��7_�B���[�N;��f��B]OW\��5�۲�8�vr�T�}����"��8�����w��-"���Ri��.�ʯ����R�#w��z�Z-�T�ut��r=��gyzo�>�2�r�#z�{�ca�9t<J�wp�W+[o�<�q�7��P��}�]��z�!U��1:;퇖�*������^$��kX��8��ʮ�HC��`��~̳(��O�Db���؇�^�'�i�e���3 �R�>�z��1C��;)r�%]E�!=�7��ͳ�8]Q������?4���{�rb;"LC���)7�Qޔ���>��Q���Z��,VrQ�b'���zƫ=�E��Jj�=M�-��]��O��/2�]�[���*�r��6�v�9�g�?xPk��g�D����Cݸ�5Wk쐲H�E�v�NO���p��r��;�V�]�Wc �%��u�7�G�lL��7�g�{���6^���rj3J)�F�5���C�����O��an�4��������~L�Fհtj�>�Ͱ���C<o�"hPѣ��~���q�GSyE}�����f-Ю+����X�g��Im������ӄ���׸:t�æE���PZ�E�ȶ��G/�=�5�R����9.������܊�F.������Z��� ]�n�h���oOz�ETо���Yu����O��ƻ�П��}�I�t�d��K�h��po?R�����[:�m��>��W?p�״=��i�:�0�s������
�����J�4���P�4�cU�g��P�>��Y��m��X����a���Ƕ��w��>w���Põ΍D���=�<]c�`��9r띄��u���"�T�
��w��k\+P�ǭ��4����>����VH��;sZ0�gk�;vrc�����>��\g�[�_������z�忕A�!�j��;��!��n���vZ|��5ici������ºZ�l�cL�]���'�����?~�1���b�~J�]�$���"�j�Y'�Ĕ�?jY7�Z�S��5��+�^���s���FlU�_}3�[��7�,�/�|E�׿,׸�Q�Iu�����+wD�خ�q(,��#T=�M>�(TE�z��~����U�w0�k�겠�slUV��w�R�Ȕ��)yӆ�7^]$��M�>�;�]ʳl�^�ք	T�ߒTc���O@|���0V�>��;Z�8���,�y�]�-�Yn	�<*��	��[���J_s�S�au�������O��7��t����'��`3W
�H�<�S@~@����������8���v�*7G�/�q����mL�?�=���Щ�:1�-4���>��G�Op��������7������8���:����l�t-���f���~�E2����/B�8��EV�����>	S�H@�*��˭�ŗI9wK-́�4��xе�j�8ӄnz���oj<�7	�J�uO�S�:{�k5bsb�X�_}���墫p��#��n4��[?ld��ΐ����	�RI�7�C�q%�#��~O�-V�M�g�cDs;nwx���_H}(^5U��d�*ZB��F[���>Im_F.�>����%�a�f6�b4�J���?�G���Q���Cy(�U����a�dI�Vg%pZg�,WV�j���.x+����vT��������hs�嗁:��P�ŜZ�"6��pu���p���������DK����M���R�h�,K�_��|���4�v�t�M�]4l4���b�Ѱ�{�0P���_�R�nn*+j����ۺF ��}a��)�>�'�,<�S���1���^�?�5������/�oi�@�yL�O�"��G!��|b�h�~?(bݾx$q|�l_�Õ�����X���~�R'�?5Z��/�~hBB L!&_*eB��>�o?
z��?"ZV)_**��4� ����I���l�Sȗ-��}	�%������/�(pC��!��d�"�ݝrW5i����&����ah~�$Q5�u�h�b�dkm���n�0�k0*��k�� n��+@9el�C��<�]kqu��6>.堦�-��s+�geks���o]�x�dMfQ�q����yL��Ӛ��Gh���/U��~�L9y�Ph��#fb�A��̆�0O�ɨ�e]ƺ������?0Gm��b8�u�Zv[M��ܺ�%��u6r���s?����^���y1��a�ġ;W�N��@��ۋ����i���w(��=���ۡ��۬�.�=�ۙ$@�k�݉����clG�Wҡ����쳻��r5Q���M�o���W"O�'���C�]���S���F#{��'f��j[&����<9���n��-�F�B=����^�:f�Y:qM8�O�K��K<�i��G�QO2���Z�O��ޏ���"ۋ&'i���Ƌ,.��gw���2ׂj��/C��Y�i�̋�sak�*;�_����~�O�0�l��7�w�I2�	��#K%�T:U�o��� �7A�F�@W�	F�Ə�Zzd'rh��ż��P5ƪ��`�:��G�9�{��61��]%�Ҝmफ़H�魧7�E��ԯ�)^w/�&8'.����es#Ŀ�%Z��l��
�3�8H��ĺ��PQqq܉��x�U���$V��4A;���"_<�{
O�\�7ݽ`���$=�<��4b�19���h�9.N=�3�Q^C��
2��[L��a���.@��.��j]d���i4�ةLG#�Х�Eז�pے�2���Jںƪ�����ڵ�邛{�^D� L���GU��!)XJ�a(��u�[z�LyE�w^"�-��j��-���Mx��2�?�$@qX�'�ɐK��L�C_pvv$��lr��ß�<ø��S�7�ҙ6a��;uV�y�?���LT�E�0N��erj �� �i�o�7�,8�[L�Z3�nLur`!����n[6V��Y�$����~Ad�8ϳ����<KFO� ��~O�<�G��H����s���� H�X�j���jH����i*Id4�[W��4�`�r�<E��X�5+�e#����{�����WBr4�b7X��:5N���>��r����0E��}4别 O�
;6=�V�j��r�`5����nĢ$��D�[���~ �V�� �B����q:S�8�p��������6�CI�L�= ί�!��X���f����9�/��n�Q�
�9F}��% ���]�^���d���x�ט��-[���ܡW^vE_�X�GѕST���T��|��ñc7K_]���Tu�s�p�����h+�*�p��ͬA�U���q���ٜ�O�X\���4)��%�!$�<�:�Y B���:��أ��-���Q��ʑ�a`����[~{���Na;�B�h��mp���K�\h�[/��I�D�=
���}`"�d�
��Ҳ��&
@7���*��rDQF� b��ZV�_�A��,����=��S����\��.�nkp�CG)�����	,%\=R��2;��* W�WWpW}X5삩�Ó'_]�W"�-˱�ox�����m�
��8�i��^�4�	��gt%�C���u�=�������i���/5`�J7a�b8�I�����HuR�)nPҔW��dT�v�Y��&k'fǰs��c�� xt%FY�3��;|?���� ��>jen���'��Q�w��M���7�k�W�N9���,���䤠�E~�����'!�d1��T8�����29��8K}x*�@�XW����{��C����|�DI��d�1�_D�4����y��K"�zU>�}��cl���a�qP;`8������!� �IV���sdϐnsntx����ǂY&2����qA�� ˟}I��!��I؇o6Γ�Ǽ~��Gl��P�i3O�Q���n�鏋�@Y�d��o~7y7�`����,aC^�p�w�#6�f,Vt��_]�t�`�{ ���<�Q;�����,f:7+��,.HU�F4&@Ǭ�+1�)�R����Xx�tTeP��N ��u�f�O���}y�Ė��e^G oy�IO����:b׬�g�p��v���~�����ݷ�߽������Y��^��gvϨg�pǢ����Ӯ���Y}����u+�b��Y��޳�L��ͦq�	���xU@�J�ĸO�nW�5�E>�:9����F�k��=4�lO|�Ȣ����\�^*�Kx�0Q�j��1���LEa�唩�>����s�.�+g"i]���=�|�'hV�m��h��)�S�����$��{�%?2��n�g��4��т���8����Uaۣ>�q9�[{g���ҩ�؎|9���dm��~�������'��,�#���;i��[CU�c��V�X��Mudy�o�Z���)d����X�f���Mt����� �n���,�E� Ծq6����qӴ�w�9�t��P�$�����)?w�8�W�P���ö�u"ҦYQ.	,�� �BhԶON0h=+P�4wֈmom���q����+Ub��6�Ƿ�c��[�	#YZhQ*����l£�}�<�; �ns��zO�S�y�F�/�d��
($��qѵP��m�uI�B�{���C�S�؞6��E�N/*���C�p�R��`J!�8qu+��ғS+�vpiii	%5=l�+Ύc(��\[ TL���?u���� W��
ݘ/�-��Ug�$ů�
}k��֐wAS=����,��Bs�/�!«x|�dXg���1��\��(,#��x�m�z��y���f����j����۝��������.cpM�#ug����7T@0s��_�&L�l���}���t]�c7u9������I�N^؀����?4�� wϙ7Q���4@�Ng	pSj��3t����b�'n�Zp̕#�i�R�W�Ql�a舯j&}�U=A�U=�ble��­z����jB�7-�V#�\�o�^@QF��}� �x� �'d�⧹)��M%URұ��H�b};��U����>:�]�VOP��#�Aq�h^���ɼd5,Z�Դ�N��zn����,+�:�S��Q�
zzK�uq��:�i/����l�\��Q��,:���򱪆�N���$�7��֍����ϭB\$輌�4��OɡKR�P"�w�G3��h�f~��ҕl���͠Y)tJ':A��R�.�J7�{{2
�6/��"kktM2M8�m����,�q�����/B�_�P�Q �t�A��P[�7�L�`(Q��!���A��8��P{7�~�/.0Ξ>��Fy�
ۀ�f�D#���H�_�|�a��}�����G����a��w�첿�����������y���J�Ǣ�(��<����l�&JC�ѰV2�s3��?�;��=w�>��{{�nN��X*�2s�V�%jI�Y٭�2/r��m�&�hi���Ja����E�X G=b�vX� �/���$��X>��u0�34v�u�Ǩ��'l��ӟF�5��7��y����'{B��tg�M��>�J�J�T�����za$�68h�ou�Kj8e�iJMɤ������ ^���^�jL�mb��\)���v��L�� �������#�ߟDǼ��n��E��q�#>�)lQG�K��՟'�w�T��Яv�CZ/T+(�Q/�#��l�jz(��+����n��KJo�Z��壶I���j�E�!�[Fv�]a�|:_�t������f���d��Ny��u4� -�ԫ��L� w�/*J&8����,R9#1`�P����㸌$��/��13��0�Τ�I]�)�3�2尚�y��RV}"�y�f����5�k7F��2��xB�o��q�����ra��d��CdhDl輧�7����HJ5y�}ŀ���d�����q���Ynȥ�<E�]�f�<�O���BJ�}f��UG�8�����w�7��U߿}�a,L�ڒY�yds�f��wA׻��z������o�&0���Λ�7��;ow��\��e0�.`~�r�5��K^�:�����f�ϕ���@W�oȞaa�mQ��_t~�'�
	@<1�sHv�&���?�Y}�K[��m�b�$%��e�6J��P	�ɺ�!��+�_��'����L3;JR�vۖK$�u�S�}�C�p"�P�k���/��l�_o,le��'��B��ܷ�.^�)�A6u�䝿C�pMw콸�CV���d����-,�s��
�\^��*�$�^b*��¦:����*^2ʅ���{���5�q�`MV�Dy�qSo��g��j���	����nU+���Ç��p����|.���|��B��{�z��GvD�a�����ӽ��je��mI�!酥�R�zy�ք�b1�'���� ]����$M�+s��`'�6�qO����ve�M��='��S����������^]�H,�]oͬ�-�٫�"�7�"���	�.��蚫�v��a�y��������X�Гڽ��<G,r��n�C=^��4 �	��^,�2v���l�7��N4��	OKڱ�hSs���%�nAwɉo0����^`��N�wT�k�F�5�+���r��,��7���!���0��`i��[�أ&�ó5��^_Jdx�Ɯ6%	�\}����նW�x�6'vEr�k̰W��̛u)^��ޑ�]="E
,�$�	_��̒(��zɽ�-E�k)}���~��Um;�	�&�@-S��p@!�]2�>u��JzO�R@�.�I2�N���l��,t��p��N��[֕7"*g ae+�	Êu5����#nݲ%N�h[B���t�	`J�]3U��K�}�VWV�/�^bt��*:��t,�s|up5��m<���3uC_h5���oDW�����hx��߾ �/Ts�j�-�Ӓ�y��@/"r�����z��Ⱦt�4C+LI�:��t)�y���iG���j�����L=��Z"��U���-i�5��E$�#��Ȫif��~�@+��I��1��-a�eF�'�'�ds�Lx>DX��	���l�%�<M��9��n\Lų�&,5u��Hߊ,/��1��q_Զ�i����Km��%�U�c�����+�]���8t[��Qg�t	-

g��"J���n�`�ŉ��p���$k.F]�m�'�,-�(Aǔ��+�J�+�ݜ�Y꧀6L�T`�A�mF/)\0g�n �߲�Vs#:/i�R�a��J��f��n3��)(�-�9I�7S���t�N���P������{��&#O���	G�>��1�,��^@N7�s&�#G@�Q�sc���Ӯ��ƹ��[�i���b%\ ��*i��W�7[�&�bՅ>[f����U�$�J%#+� ~���Z*	2L*�JWs���}�#�/�Ԏ���[�1���X�������ro]�����]�V2�v���큵	(�$�[=li,&������a�`F��X^f�#d�s��Jԧ�c��4�S�o�<R��˭z~�/yQ%)
k��_)J�5����r��'���n���5Q������m���;
�A}7�o��-�ň6)x��m��L[�?�o�l�6>����Z��ȇ���[Zg��.� >:��5;��T��q�v�2��,�X [M���՘�ߴZO�GMF��]Pt	q+�&φ,*�Иmu���,��y�Iw.3�cSJ��Wy��>�W����$���;��&hz,���4��+l��*�Nq[�i+�m�Ů!��������:x8x�:c��'�5g����N�զ_U�n�	[�YW�)V��TӡS�IR�3�Ɠ����#��	��Cͫ�oY��y�e���m�ƕu9j�JN�P���������X�۸<����z *��w���<�qP�Y��i�'�.n�%��06'/�� �6��1���������9���2�nO�`������J�݄ė�\g���uHztdӒ i����̼X�T��IE��aڈ?���-�Vd�d��s,��-�*IJ��4�5 T�Yv�`m��i���|�Ǯ�f/��7;�G��oW�x�q� ��7����,5�W)j�}��̯��y+�lob6IOA�F��,�����"��ˌ��2�8'R��:�(����̐�.t�.Ր2����(/XMC�4�bEgT���w#6O���	�۽�!V���T(�ے�ۭ'�b���f��*�4���)oEլ&<�N-�_���5`嵂.��y�sr���t��!o�ux�{��4p ��nh����R���Sy�P- ��5 �[@#��G0��ģ풙�J�ݸ;��<(�Z�d��W�E$���A�0��o��W׺��w�%���L�V�G8fNqr�we�y�0��ǈŗ�ujJǶ�����u�[L҈��~�`vj$s�K��z�pЩ!T1���<
���i�]��g�1�$���cG~��-���(d mV�Шëؙ|��L��{�oga�fA���+iά-nԇ]![���/���8�YM2��GIjF����a�v��z���������_���X̀>a��ӳ_�~��!߶���so�P�D�{��?\9"w0�q>���_����X�B��X'��z�!.�@�ͳ����7}��
����#�B���<H�<�RXO�Sc��w�1[�-J���*e��`�ݏ�GgP��j��Ε`~�_������z?�wf}g��C����e^]�0�O]�1�k�̚�F��$��H�x����'f��H-� �k��s�| @���iu�89n�6���I��������ЬY�^μ�@P��M�o��sd����}rf�h��4!pF�q�Rxx [c��R�j����=+ܞ�f{�C*���}������zczɺ��f��nS�P_�B~��Ơ+�`�Xs�G�*%�ړ���9fQ���Ӟa��-wX;�"^NV����D�ևY!~'�.�u^
ܖu�'�{�v��6B��b�<S=�����͗�.O���9F��_�S���ɚ��(��S��Q%
=�A#�=�
\yj})>&ө�A�	��D!�����ڒ���f�w�С&?�v�q��^<�F�]��؀x��3rO�����$�j,����~���Ԉa��
�d"y���r�9�\����.��z	5��K�g��o-V�=�GN�����Ӊ1�ؤZ�a2d�l�.����J�b}�	7,9��YCT���c����t��R#��ڲv���h/]-}�c�^/��c�&:ONř�7�
Ln�4��D���iq�!�V1���w/*���˔g��],�`��Օv�t��$�����$Ï*`��>�K�bX1YM���{%V�!X��X�?�A{e(e����Pe��Cl���AX�ƛ���n���3=\�B ����O�XX%٬���z�"�A���G�H� 0[�Ս��d���������=���ɗ:��w_��J�ܾn�0G��O9i�B�Tщ�v�7ҜP��<1�g	���H1�ׇ,=%���eJ^��S�؉�Y�M��2)Sh����rp��   ���}�nI��{E(Q�HV��KT��<��Q����Q��Z(�Af�23�#3E�9��}��}X`_v�i��/�U2?0�	���a~DdR%U˻��~�����ہ&��Y�C��ڱSK"�_	-����g�K�ر�~��hu���@�M�g��^Y��-���Q��cV��s�s�x����?Pf���0"P��8���B>��d^��Yt��+��.�R�����~6��]��9����N��M���!9W�$��?���� nR9�$F�R�Ei���Xmh�C@�ERp�8���z˔y�}���C�K����<��i�t�b�h���N4	�p%!T���Xl;����md|KPg
�c����`�����􃍋B���tV�B���yQ��/�l�Ũ�E�4ș�R�Z Dxd�ގ�E�M��:�7��n�s�r2�d�\c��`�`囃%:���Q"j`FW������u|�wu[t5��X����0"�ps��V�=��x+S�#�,4;���I(������CB$e��Y?ƽ�	H�f�X���"�Y8��?Y��tvIx��Uv^�!;����z�]�j{�ܻ�+��]^�s���e5��C_[��0��e5�v�=�4�@-��Yt7VI?�C�8��s��1Ώ@�g����ɰ��ڹ1Hs���'L4���+�~}��X���ʁnU�H�٨�K���s�� ��~F�Q])#%��Nֶ3�b	Ѯ"J&��='�[�]�=tm�[Ά�1���[tHY�l��+��N�[H����B�%�,��,hD%��/2��)ϊ?��I�c� )����|��XeP']�NcT�|�Q͊�hU0i/#	bdg�QN��I���l-��]-�(Lk�_��LI'w\�(���T�Z�#
��`#��G�a�pE�R���v�(t����]�i��40�/��t��?��|���O<K�h�j�kmƇs'f�S0�#nW�u2�h����u���xX�&���r�+/�a
ζ�����,4�6����0?.��v6Lݦ7P��ӳ���Bͽ�:��Ӫ^�T����h��	��F3
�G���'5�`}��r�",�ؾ�ON��l�ӻN/V3���Rj��[��iP����G�DC׭��m�c�D�_K��X[	���7�W��T\`3���9||@�F� �I�:���͠�����	���_onL.��-0��V�zÈ����Y1%��4�qԂY]������Y���ɏ��pN`>,NgP��H��[�E1:�K����g%"�7`�|@oX�!��ާ��Y5g�z�^� �d�I��(j�-W|!kt�	9m,"		���,7��5��1$�R��Q'��.X
�z�y��S�p���/� ��=���Iix�Q⻀��u��턱t�[r'R|�u����8���	��b ��8^V��E�O >�[g_ S_�1��5��O�4��Z�ł�_������*��)�N|�X���������i�m�����5�n��!�>������1��f0"�e��廪�xe�=%bzj }�]��R{˳�`g�\?|>Ǎ�.�������i�J��g�*����\�i���QǪl:�?�[1�l�_�Mh�?��\z}�ts ���^l��\g�жX��C5溛��Bw�FNoB_�p��5�֦Ws���p��)&ErH��FRI����2�*���]Ԉ�Q��Ư��՛�4�q9�8'G(~-p���x�g���Q�P��d�X��D>��?�(���b��6�o�R���W�Юn�
���H�~�J=�c��Y���ݨ*6,B`�F�&��"���ʡ2�F���t�_�Ց��.cx(���A����iN���.�z
KzP\��뛁�*sb���SQ�Y��n��|�SO�F��x��ޚw�m��}ҋ9[Ûz�^��V�{Q���e�g�V�ƍ�E�>2��o���t�q��rh�,��A���x�:WxM�f?#/�۪g��L��_�~~��z�e�<��^�2۴� ����ѭ#��|֏a�J��ŉ��Y�Ot�u���Xaq�(�!se3"�=^��1��iq|��*Nh_���x-<:�[x���6�IPZХ�(�0v��"vć˵��\�'��҈��P82��p�Y�� 2SO���n�J�"MD+�:�'�)��4"�HI��6��D�ԩ*<,�,o&�@+���4�X�	���`+�h#KXO2��4�IY�@R�����ZҰ@��H3?�,,��5�����&p�6��kc͡i\�b��m%$s�
��7M��]�4�LH��O��@Gi���r�4���&�	L"5?vp�,��&"����\?n��ɮx�^�!���6`��e��7�k����ﱳ���0�p)�k���f�=����~���=�@�:����e�OK�Q��&U���_1�@��u9~���ip���!��{���˵-^��*�7Y��w��=/��7��P�A��?�d-3mUt5Î�e���������vK���Z�?����ޒ�W1����]X��>�6�z.Ԏ5�2�}̀����7/�@����W"@j�N��z`�#>�������Ւg�b"c��1]Z��y]��C��~���J����o�z���+�W�ƀ�b}�אEM��Ņ�0��t�`ke��E ������e����*n�!�4�Vb��!Έ�P�a�[�QdV�/f`��X/+�z�4���\W�r�AJ�5	���C9C�/��M��GR䔃؇����~��Nb��y�{���޻�'��Y9���R@RL����tQP/�M4o.p�܉53\ja�&~ϊ�ÿT}������=,"3�C�ѫSz�O��������C'\i�?,G���f`5R����Z�U|(���>��qn� ������'�,��<��
j|)���j�ܐz���s����(���օ޼VjQk�-��\	�`���L�&Q{w�B
F��b�12���n�r��9�lD��`�I�?~��˭��)7�!v��Jϋ�|8cI�"�����w�i75��q��&����l/��"ԏ��-���� f]��au�)��ֆ�Nkmi�w�.�P%�fy߇T���K>���E2?��9�G�2~9<d�S��)��^l!����_�������e,2��q����tq(v��yw7��֐�U+��S"L�0�J�xn��r(^Mh� �6��8�Ȏ��S��ayd�?��|�z*���0dV`+��zb��VLGUjT�2yE�d���Ξ���g��e�by���KF���uP�P�a�v��,9��+V&���뉞�8�a?1)i32 �2�	�B�$ܕ)v��_X1�^�K0�(T�C�Ӡ���f�_�
;��덟6'?m��㼻�J����Zy��=KX�&�VhYz��KoU�ɬH��,�%��I��AD[�t���d�1�^=���Z��3$y�3TEn��Q{�q�d���J4��N�X�.�Kx���;��=N�̡N�íp ��� ��fݍ�gU�U�@�`�����b��%���>^@�-uZm�O����[�p�R��Z���0�R�s���wC��e��^�V��Q�/��  -�@:6���昌��3��e�>H/��P��f���Pf�?����b|"��V�����@ڊO���#2��Z�h7vj3��1TlJ7r�Qڧv|�T�ߜ�Ex�7�y�x��Q�0[CB�%N6�ϥaߕ��C;���\>/�<��B�#n-S�>�Y=��$��5���:��zV���jN�K���H榩Ǭ�/�(�6���Ѳ��G3q��zs��x��3�$�*�HV��õ6�����x��60J\y�r��e�FDطE�t��,?˧�$�B!�Lh�k疫�kqV�<�M�����wff�@��w��JƬ
8��69�@��s2�T�E�Vb�*�	�GW������X�E��q�^��T� j�>m�@�CT���J��(��t��&�4�-�$Ul�ꋥ�MZ�R4�t�t�A	�Ow�TI�Di�s���tn�J����4�ȕ~�O>H1V���4��l�2��-pu9�\��#�w�֌��^��fwW����wuc�݆w��ܔo��A���^2�n˰۱k�3�qO6�0�0�&6� ���fܾ2,���o2u�/���r�A�K�.��;���M��F��z��Ks־���y���:�FM�6+v����f�����F��~3q�:�6YQ'4\��.�2�%z�m+H+5j�;��;���vл�]�ʾ�1�6��e\����H8vi�k�[���4�Ǧ�(����q�����L�@��|F}q�ո�cn�?|	���H��t>��a�TIT�9�,��(Jh'�Wׁ�m{�5�6�������ˢe���D��b�^��Lqt��x�����N�r�]����,�6B1$��W8��Ŀ1B�W�΂�5�O�� ZaL͇�����w�W ��w٦Ϸ�ʳ�V�?�N *��^�E]6T���%�Y�5�(T1H�|��H)���F,��O
,��'g�0�����x����`0��W�bL�d\�  O�S���lm�%��Oޭ��ȪI~R�/aX���n���
����!<J��.��$2F�w�Nb�
�y��%H?�K@U�tR������!��),�e��w��ʹ��xJ�\D�f�v@�p������<@��n.x0��h�D~Β�a��'	��5$2�6�q�1��\�"��.��1��b�jjE2C>M
�l/�̪�7v�{ˬ���Z~m�iwb)S��oB��B�m�x3�(���.�V�p�CZ��e��P��P��l$�l8NCm� ��II�5��%	5��&aoT�w���F����	�m�H�kpm"��ֆNLPFb�q��/4�Y����b-�y�zi(���=�C*���/~w,_���s��3~/W�	5#����	!���օ�GB�I+���b��Z�|�ڊ٫$U�Rm�(����|�������̉��0��Il�f"��K4�}>k��c�p*h�M�C�c��(�;�:�sy=�!��g�Ԧ��hAq�UU�5��٢����{��������5�֞EޛW�K����3Xk�;�x�(:��F0�|2���������+����
U2�XSIQ�X�����Up�p���N0�xx�VL��؁�Qd��	$F�}�в5�εh|���Z��j�+��E��,!VWj����t����1�y4LW�]�?T�R�C�/����r�j>Z�qʲ��¡��IC�V��Ve�X������j���U2-G�����}}Ȗ��~�6�E+��5�,��aeq��S�A�鈒�nKI�F��%J���j�h��A�iY����
��6j����ވW�#�ߤ�ҥHw��,�`Y� 3�6١x��QPf���@�(Q�q�((�3FAY k���QPf���H 	u2KJH��4��٨���T	��%�J��c��D��l;a��\AY8��Vٮ�ĳ�@i��
�2_Ai��
J�Ė��5�&���aAi�K�u��XPZfƂ��D2dAY8K�V����MiI���͚e5�2g�CɭՈj�O\#D��S�����k��_-"���I�v���:BV���fP�,f�`n���4%י��&��A�3���yϠ��}%��5ʁe	yР�̅%mb�s�A	�E�� 7m0��L�Qu�35��yS���n�D��pf�ȕ��b�~�����l�8��u�W������n�_��l���l��Ga�����]�4�֡=��W}��/&�x�j��X�$K�̧a���j��?�\G���aeIv>b�~#������� ��D�|���W%'�.^:��������� ���'��u5]�]�_��D�S��)�8g�嘐�|�u%L�w���B!o����~�q9"c5jm�Z�X�Pq��Q~�v湈A���bF{�ڃ���n����"b3c-�d�N�*�(�Y�W�X!�e���+t�0h��9 ��a�^�@�W}��B�2Q��3��p�f%�����k���I�1����qL��/1�?S�P�=��F��/�ǣr�ń�����"��m�9.۠�q��٣z`@����ش4�|�@0h,E�K*y�z�2�f_{8�g�!g���3����	~�?'Aí��t�*4ݲ�U~�b�l_��L섗88��ν%�@�1'���w�r�2���uu�Us�`�Ƅ$��T[�I2(��b4eM(q�W@�Lβ��b1�����ǍQl54.�SJ;���ZG,����}���d��p�q�ˤ����S��)�C���tn!�p�nV�0<`��U_·M�sC�M������� �>�0��Êe�:Gᘾ��,�ٱvIH-$��d�B�9�&��A1�{<�H�:&4�!$��8>J\�+�u����?A%�mX��nO�M��E]�P��B#�T��]��0Y�/H@q�l�h���b�/L;B�Хv��W����H��Wb[��<�����ֱ�=�����
�36�e@VA�c*�\�_DC� ��5�Ro<#�IBKJm��b�
�x��E��3jS^p��+�څ�@�b�����*`r��/`u�V��2c���4گh�2B��}�R�n�g
#��l�[ _��OK���>��1�HL�N�`�zwC�����׺�J-g�(��z���O=�u��z������g��Ι��ިҙ[?3��:��O9�9�Q.#/>2���V������� �R<"UK|
����J�~!���p�����I��4�3�Qd�������\����C؍`�k#[���q�=/�rx����E����c�"�zeG�����a�_B�%����"1dB���A����6��z��A�*�������koŰ�o *@� ,�FtdRAMk�����3�r�[`�|͍1(�G?U	h�̡�vn���s�H��Z;-��t�A�R���=�_��xC5Ա��� &I=2�}�W�Woc�bF��b��f��E��?~�!j$h
��jYcX+	�/�w(�Eܟ]��G��a���YvK�.�MP�kH�_�@S���J�k��7��Kc,���k��'�Ak1O�$Wd>+�UFnؒ��g�9$����{we`�g� �Rde}"\>>e��_;����+f�믿M�'�,��*-�nP�&Bs�"D8�'�u���5p.n��0�#_��Lq��v>o�؈����݊r
��fLv���[l�-/ִT:�涙iP���ny&�EzU.�0܉���=�_�M�=#'Y0�+�����_�\u�@�+<S�����q��d?��}4��AA���8;ɇԨ�ʨ�d=�=X�����xsc���p��O��Mgń|��d4sH�V��%���O|!���\�����bJ25zZ�'��/�!#@�~�����ķnZ��/w���������ۼ~��_�(�Җ"�Q��O~mKBKyB�_�m'�]hߟ��w�Յ��E_߷V�Ow�:���ᳺ�����
.��h��'6�s�<򏵰ͶA�U!B�ت(��q�� ��bz�R��Ό�.E����h�>�;����x^�Z�l,<�[�~IA�a����3�$�n�2s���B�f�B��PxN/Ƚ�&ܨ�d�:?y�|^�#BM8���VX�= w�p4��\�%��J������Td!Q�v��G�o�|
��l\Lg��}ه7Ո���8��#ބ��r��1 ��r3�X�m��� 7��B-�l��+�L��D��x³'��e7�%8+�7Ǭ�{�"��E�����1���H�甈=�0��>O	�J=9�i�$Cn��m�y/u�Xi���$��f3d��ᚗ�5����\��v�}�@e�k3���(i�mE�9τ���n7�7i��0�4+钪*�@��Қ��V�_����	D'G"�=b*NL�q���	�E��d�g$oH-���]����.��ir���:�t�\O��E��9��}���l1J-�
��r��iL�zY�D�����h�e�M�$(�;��%e��-���1kCr���Ywc�7�^��9@t��D`3>e鸌Σ��Z�x�%�Ġ{z�MFM���WY1��������\�0~��zh����)5�·���Y4wAI���4�Oa듧iٿ9��
C�S n'ſ1R���p/f=/���FLWq�"�A�����"�-����2�Ȳ��7���ܦD���7*�;�J�'^Zj�F�m��Q �vw'��'��O"1Ҳy��-�y�^=#<bf��ۣ��}�+���n��h|��G�?�~i��%�66��~�Za�S(M8�4Ȉ�ؕ,��4u\eS�T�Q��ِZDSׂG�r�U��~�����؆���π��mԈ�z�r��z����Lg[�R.�3կ{yH3�u�ه��mg��?����+�%���5rJs��;�3%�r���dN6��[2s8� 8�#U��e����P8c�A7��kv9 N?�~��PSHEt�P�5��F#�xR::���8�.�i�g/�� k_#p��D�);�uzBt{ �m$jo��F�1s_�C�}x&�iI҃v�A�_���$.�x,*��eIFK����H���ua�|�eP���P~L���7�9���d�~�?��9�� �H!�H
{����O�B���/�|:c~_�	4D�gxF�s��s�-� ���A>�M&�业ݫƧ��x\�y^����֤�N�a�����L�rW�BL�5%
t���T������'"�EY��{ǧ��uG�ٔf_�f�c��''�tJc	u]�J��We����wi�S�3.������C6
+�\ğ�9�q�`#_t'�j\�1��X�3���	�#�؃�~���4�����E+�o|sX����t�_���G?�=-Or�i�Y>��p66���ƦJ=_��{k�Pc��|�mb�[0����H��zH�f�����uU֫���K��{��a�=)�fh��[�au��� �A� ��� m�k��Cu��H��O���Y}i�@Z!�@=�]���F]���f�6�C:?��YvZ�N��:��:�Y��ΪƛG�lP�E�<{��vE2 ������!�� 4�%�R)��L,��?O��������f|��[����8�I��#��<c#��X#�\#B[h�ֵ��ɬϠm��F�����������]�O,�X��K- E5��zhx���qR�����K!;s�{�t\�=�I�[� ��6CQ���ݸ:g�$�o�H��bHP�
솷vw0�YF�M�뷪%��A��!�u�n16��:���!O��D�4��xE��k�����iAH��eQ�VA|!��Hh����۝�M���Z��Z���C�.���S��9��Z+��'��L��yIS���z��ƨxF	8�Z{���o�iR�O����_�Q����	is�tR�WY���X%�0�W�bgʫZ����Ż	���2l������.aϣ�p��ao>Df���A����{*��ј����z��,Uև�s��Ui� �#���z�h�O�
��0Z�0e�r8�'"��	HW��e~��K���)��p5�>�5�j���՘rZ�T�9���;�\N�;� �I�b�+h���i�#�g��{"A?G���G����$���mh�4���d1�{����bTuuBnyT��ez�̓#k���uAHff���):�"�㾫��+�>C�0gV+�)�>�A׫�kh�\�}�N�[�?��#`�T`��4�̈́K3����e���]�Q���A�W�d,P�Z �4�YE_0
���{�EQ����=igL��ظcJe�NFv�l+;�ґ����H��Tع�_:'?�����5��1�t��f�Ū�G|���NO�Z���Id�Fh��mO�/��C����64h7� �0�l@�Tk*6SѸ��"Dt*2"�g ����KN�dQF%�rP;�ȷ/��_�@Γ�YR;��Ku��G3q�6�3M����o���Z��Qw�ϲ��͍����|����/OO_�"��z��%	���S��&8L��A�tX�O6��uaTdpw俶n���/�أX�A/�ޒ��8�,q- L(�e����`��UG��8c�l~6�D�I�:��6w��*y{X�/�p�A��ه
���Є�E�������1�>�W�F��Crb՞�1����R#���?�����kؕd��x���f�Z?��S�&a�G�jg��})�Zՠ�Ff�^h�����Ȝ�WI��N�5_�ԵFf|Y��92��*�	��r��gd��#�l�i�IQ�SYG��������|�C1��Z�O�4��y8�����kE<��%���˚0���n�������Z%�a�_я�)R��J�!�^Z>2];�4	����I��Q���Ǐ�{���/z�<{���ы��=�����p崢�~b��zW��-�f��C��� �j�IQxp]k���H��#��R�Z�w	�Tx/k�Y?�+Y���3�����|=�3�0���=\�X�#�S�0��}:�a7����[Gdn1+)�Z�$:�CǓ�Q�ժ�t-�F�1�.�ɲ�@m�|�f��j8�sII6o�-����xE;-Ȫ�2/ybQs����J��V�� 3�8����	�'K��l�|�Tw:~�O0L�F���-���֙D�D�SG�XY�S>�	'��
;6R�(��{,!����"�]���`V|Lp�Ov�ܙ`{����MJ�m(� "����З��q	����u_0^���;ޣ�<I�m?�L���^>>)���G4�䇼b��w�s��F<]���tNHwJ'�f�@�P��jF�[5N+�Z����)�6�j[�[��E��Pl��͞�ɬ:�)g{x��1�5,`kh� Oy��~q�Za�&�_�,�`��.꼮I���~�%�,T5��Q!���"ۣ��Td������zw��/ꗢǍf�'����TD Յ��r"���i�ہgM�%�~��S9Ӭ���j�m�)�|5���:���"�`-��t�҇��<e�S�?���%����)��	��k~Ϟb���'�dj���>�M�t���j+�oWUo�^�8�tF��}�ҲՍ�����@�B_`�\�J�h���T$7��f���5�~�r�5߮LI�Q�]�4GL���}A�� e�F���9�\.�(?%#�t����K�d�a ����<�-�t���J������!G�?~�kV\��fIH��
{��Z=��/蠜��m` ����X�Q,h��5i{�Djj]�
/���4�/v��9�'�c�1±Vk����Kg�W/�P{�eo�����Wd��Y	�i�Ip����ȟ+���U��Օ���,"k��-��&r�ЃF�E�i� m��ǆ�"�K籿���d�-<4͌�@�\4����ѧl�[2,1��u��
�5 q�aT�C��}D/;J��������� �au����IX �~��h�ξ/�&�=$���3�C��3��xD�������/z�?x�����=}����GG��_n��he�� �]טCV��������K�9R�4>�L����z;�i�5\�jɑ��_�}c�ƈ����ع_�����P�T���=a9�,�����?~��j�5���J���֟Ư�CU���G6{�+�k���i���W5ʡ�l4��|m9�]o{v�+���-��lfg��v^p��!=DzT�G	�����"�tj9k Z���\[��c�������:����8M����>&'Ր�5$%�b�z�9bH��Ԁ:݌X�E^�"�|�Wd�"�|�W��+_�_���D��"՝�u�f�_����ޖ�{��n)��
X>0���Z�^k
�����G/�����5��S��:~���}l���eVaYq���~uޣ:��۹_L��	^]�!@2�c���Ί����B��Z`h��]���e?�髏ZnC���=��!� �rI���:����#+!a�[
����H�4Mg�S[���|��^1k�\v#����%���W��4���5'��pDE1S��d��#���7_]yGr-�2};P�O�4�)}8IL���>ņZuM(����a�b9(�C���\"���WW��8(<?�؃ �&�Go��b�f�� W�|�Nរ����`�A�bD)��#��v���f�}u�������H:��kq=!P����Ȟ*� �������R|�R^R��nQ���Fc�c�"��v	�ҋG��^*~�����w���w�^ޱ�D��{��Κ���_h�x���َ���y1�1��6u�a��Y�q9~��J�����>cWv�;���P^�e����5��q���V�~�ݘ���U9����9�����\�î���L0��O�I��w7<�� �X�e��BU ��0��#5���a���E�o�CE��s��ú��Sq��Gb2�ӈId����9�
w���UPѼ+.��ډZ!�O������O�&�ւv|'6��5|��O!�sC捣Y�����!7��m���Q���I����7TF�d�t�6��),�����,�t�<��W�����ɬ�>"���-��e��{g���S�T�>Ҹ��h��܁�u��O3T-�V�)����vcjF�����AJ�/m��]��Jke�k@�5�w���I��)���/#���_���l��)"�د�ݍ�>�l"�m��z&���w$4nG�RT�	0�9Ix���I}r?���EL�BY�|�J�S�4��ڌ�:����Vla���aB�M�����P��4�4�+1X������h�]c��/4Y���ġ�ߌ�iI@4D����j-Z�(��dX�c�jǍ���u��|e��Z8~�.�2�<�R�(���.;!{t�g�hQxб�e�ih#pv�F��Z���Q���i㴰����'{���Zu󩲨L�8�@���%�[]� g��?�_?[@�=?o
	� ���4�l58n�.2=��S+7�,�2M��<��օ���^4XZaL�N���mi]By~>đ���L���BS`���Ѓa4�	˄5�[�$Kggغ�7+�ES�8B�������jPX���\k\�Oe)�-u>_�tc��Ҍp��c}ޡ0�M�	���zA�C)���]LbÅieo��e��0�,��'�sE㐠��1���f���L�\���X%�B�$��6٥5=o��
&�����cK�K��1=��/�H҅#��iW{g�n�~��B%�7�0/�C��Ga���E��2�֗�q�B,M�~u%zҴ�� �-P��Ef�(���YyF(|�{Kq�Z3#iye���S����p��y�#ü�ـ�Nz�-@lz&�m#��]o���������dn99M�����>!�[ٔp������f�Q'.��wo��D>�2�zj1{�E�
�d&U�f��-�ǉ��UF�T�^��4�YO�������?���N'0�?[扊�32M�6����ډA�Gp�2a�D�";�i$�T�N\f�iE�.�8JB�Ezu�)�Fsܝ���X5�ȯ�l�X���aδg�lH�����=2Щ٥��q~Y�g�C��6�^�{�P�m�����F{��1��Y09�h�\:��\$%h�F��L���s��[����Lh�勵;7wI,�H�6H�φ��8.�i�zjX��ƯG��"�"��Y���܁�X �tA�f���{5#� �?����tÍ����b2���.�Ӌ��c����7�~J���~a�ص�'�~p����!$G��-��Z,����~u>NYW6Cl]]�6H��\��̏��%��j-�*[��8N�y]Tu�ǗN�eS2$c��S4�u�Eu&Ǆ8Q2e��#{=�$�1ك�8Vn�8�+-FН��0�ڕ��QɑF�5EN�u}c�.�Z������+�W,D���wPa9kC#a���Cp���>�U����?��A'�̍>��iM?��ӆ�)/"#�����G�G���lDi�<ͺ4T`>]q	��5TN� '1�w���A`L��m��4���"�=�y��7�K��z%r(W��QT���d{�O�!x���+>! 9��8�;�n��9�(�/���R(�~������QɒAX(t!��	H���>�yEvl�*]�?��� +&�!Ih����@�Π�\��HW�a���vSa�~~�8?�V�9dm'H9�&k��[[R:�K��<�� �H��<@�v�zf�(Vgd���5$��՜Z��z=��J�ϳNR(#���/p����dݲI-${=�/��q�=�Ow��v��!�>Bt�a9.��p���#���,�x3�w��4Y0�*-�$��T�z;�1ٗ֒���#���1+uW��1EB�fQB�E��O��V�"؅�6�#ɞ*�d�ϗ�<�����ٻh%���@��\�0$< ��$�>iOx �L�lZ�П" ���L(�D2a}����_ʮ@�H�m$c��G��=B��������?ɕx`����}E8Py߿6b������׊�Wf�5f�8�f�}д��k	�kO��9���77=�烷�>k9r�/8Q�j}�h[��+'C��i�N%�N}����ub���?.eeT�4\���~�Hj�jF]��_QX��GeYx����!�
�
ן�j��6�:�n������L�y�K�� g�%�bHWk�Q
���q�,ʹ�U��lԸ�bbH��m��V�Q:3*��ܵS��2�;�
j8�����0��o��J�fwQ�#�H���A����v�Oˎ ƗQ>D�xA����Ul3�vے�&d}0?�Z��#h�Ũ��a�n73,O�{��@ɤk����x0�_u��q�w�I}~���	r VW�||�zo�>9؊�RMpVh�~�]�PE�T�����V�v�0p�𴨺�)s�n<��:yiN����І1+�~�����F����T�σ���\�iO0K\c��k�e��*����ݛ�2�kv�ս��aݍ���f ��� «|�-@�T�7{>�ҳ!� �whO����L��8{����f~I��G3�m�\,0Sp:�8m^c��t�K����l@��v��� �i�/�]��oߑ�4r�}�EL�� �x5>;[#Kɿ�dw���()i�Z7 X�U� ���S�H�k�@�������,Ʈ)V����7�b�5��åaX��C�*o���I�ħɳ,�w6,�+s��#�3.l�/����z���S�WfI��c�Ay[�<rH��'L�1�b���ik�׀��� �G�nӁ�emb�,�7Bg��M�'����Pϕk�r��9Y
�0�Ugg�¢�q�8���:J�Z�%�$c�����ƕ���T�uܯ�+$_�5�^����C�8���%��c�/	
s�&X��Ϗ�`�C�8�q^�ڪ���P�/��i4�+����vIn$!V� �NY�DA�~��h�(ˢ�`�U�Nf�b���nY�����3v�y{���u~ك��]�d�2�Z4>�9�7͆�<�^=�Rj�5�V������\!{�K\ηƧ%������I���Z.#�@ZW�+�[I��
Y�+�䬼u�P�Q��%�K��3��y�Ď��E��T�qQ^v�hJ��h*h]Q�xL��b&u��⟎­i%����z��Ƃ�Q�]Vt��j��ۘ��<
��?�^4w���t0ɤ4N2mA<8��x�u���!̊�����@P��s~$"ȑ0�>'^dd��6EU� ͽ��{�}���B�9�f�ն*������C�(��Wj�!C�,�r�(,M�_��3�`��[LT$fdه���H� O�ud�a�	Ud��&<#�;�ኢ{60�a1^�:�����c�3{:A��]�+7����Z��Ys�R
Em�5س��Z]����cDS^ 82��S˦�^�	O4��]�8/Hո�2��SN�`Ġ
Q[K�Ai6�/�5�`;6F�u%hf�,y��ɖ�-@AYuƐ�׆G�����/cH�ٛ��z:�9Qko��׹��y�_�%�;�t@D����K��@��>C*��v`����&�e�V�[�YچWg�JL{�J�YZ/n�\�tjEB��%@���R�ۙ��Ԍ�4|��V�� b��th}��j�����f�Y�9��v�/ܹ�ڿ�A�ma��ٹ�<r^�z�*�^�y��
h�SHf��7��z}��v�$��Ss���R,�|�5�)�Y]`P�t�a	Ou]��,�AI:���F������;��P����y�[N;+i��.Y��RlBe�n.s�!V���՛W�[R<mU�;4�'�yCan&��1XV�Ji�R ��r��OR�o�ЭTd�eJ��;>��C��p�{ԏ�3��k(j �	�� �F�����Lk��u\g�+]�x?��ȕ�>�#�z� ��f�1@�D Cm�<���u�[Q�ǁD��㫆y쨣.���:,��@�!���*f��������q��ra6,rp?X5�E@�t�1�+�W�N2���쨒p&�2H�G���8�`	7��w�wYrnSw��v�F@ʸ�hgi黺,N��K'p�=�\�DT�����e1�$�i��V�?�P{+�ʘC`���e�Y�M.s�I[�h6�e��`�W��z�6�Փ�\Ǘ+�`�W��6ͫ3
�!o�8�c�#����j�x%0+�Y�p/�[/vfZ9�$�>�m��[�O��x�][ш*����3��ƍ���4�fS��M�8������+�mq��i���5Jc�rs�.Oj�.�و;�ZP{D� ��sL�&�D��"sN`,q9��q�I^��JF�z>>!u;�h:�T����<+C���i�<W���JKn��$p��UQ"�\�M6���ZD�&2HkP��D�+-O������R�d����NЎ!)+:~�F�����N{�	�K�u}G#�g��4={j�5���v&�vA��.��iw��(�X�m����tX��:��I: y�������3�H��|VX��q���c���g"�����'򰥘�/���6K5D/��vD�or
�s�@;�A�HҒ6��VRKi�'7h��J�DҎ���O��i+#�!r�D���[�&%�)&ER�(}�Ouq�o4^�j����0��a]�;㪚�b������1���Oq2��pZv$-����#nB�4��{�����au\�<Ǵ���eJ��y{aL��\*~��zB/_6#�d��������a�ivՆ8F�ꦐ,l�10K3.�˪���$d�2�qvM쭍껌g�D�/����ˇ�>8�����y�Ϛd|N�{�:;�J¸6s��c�ym�����������o�iY.Z��K���r��<:ˏ��P�(�8����%�;���E�G6��I�������͖��R׷}��i�y��L%KB�)<�V�_�/i��������� tٺ-�Yz�{��=�"y�6d��f�W'��<�8�*��=:u�� %�W`x#�y� ���f�U�RA��ˌ�aۤ���F��P����j��%��� T{�A�呎.�W�	Pq\o���"�D��D-Y��ެ���̓Tƪ�s1�΅��21]P�퐕T�CV� �����f�rJU}��Iճ�`�YK��s��<H}Yi���J+�GVҡ��d%�t����J�)#�o
i�����n`�,��D�97K3���7��n����2S��q5c��Yڌ"J�f.���$�r#q(�rLøS�"'�r��M.���:�Nh�"˔<͍���n��6sZIjs1�Xsdɾ��N�;���6�$YQ��I�;'�wV���PԌw�Y�[GRkn���4m�6ܺ�<��O��e���a.��-0��>obxK�{A����2�ߠ,�J+?81����A��$ѥ�_�d�&�?�6n���2rw�x���:������F;�!�f�>p�䳉7�%0b��r �s	�F��-,��zY^x��D{�_����x��߄�"�Bq�Âzk	&�g�C�)��Xc/R�W �!�7�y���Σ���x��A�2�^|�I˪*��)�����|2�qs�CV:=��	��������|���)�5�55���X�e->������&�:���HN���47�1ʩ�֖��ۯt����	 ���(6�MWܽ�m�Z�~_l��*%��j|jQ+!V0��a��Y��ݭ���|;�MN�J���&A	�b	�>º���څ�k�٥�Dq�准KXr#<�g���������Pq	7���   ���}�n�H���|Ev�wP�X��nw�dY�{t�li$�g>�M)��b5ɒ���3,v�`��9�<���O��6"o�$3�ɒd��&��b2�����k�z��>f_�y�*Pq���֣�oz�)NE�M|P����!/�IH�H�f%?%��Qt5H�
Tc�b%+�J�"ZĚ�.���c�3�����4�t�GMW��њ���z5o:f���S��3�di�����VΉkꡱ�c�g)��I�**��ُ'iV4͸����y�8eѽ�S���bN~�&QL6�,����Ļ���w�I�xO&Iu��z����MS�w�w����˪�{ky񚡓O8td$<���N<:� ���H�z`����C�Ɛ��Y��pc��q�ir��\�Z!QBc�W�0@�1x�{�Lc�݀i`r�FC��(c�!y
�#)9��!�G�B�4(�n4fsyѬ�gy�| jq�D�n�H��0�$�Z��sB*Ck��E<6ڜR��i	�߹@��CB�1�#��k��RCH:Iri	�,�[�Y��]8i�4��=Lڿ&��H�rkg'W�>���>d��i	�K+��N�(L��"{�]��H!�Τ���MP�����&t�*\�N�$=&�U_��g}n�fLDX-c7�h�1Mt��&�]�u�<2;�G��d	]�u��P�Xak�V�3�$��Хô "���`9
�6��Yz@Cˉ�(1�֨�b���	L*���j�dg����\�L�OZu�	����u���8���e<>��Bq��3YO�jD�.¿��5�`�W��d�En�t�8��Ƚ2�".8�I�*�����b�$�O���|󧹶q?ݼuI<��i��T�$�� ��X�Z͜l�屨G��X�
Zܹc_LQ��Z�}lV_q��h�ܘ��\6zc��� ��r�5D�y^),5d~�;;c���f���H�'�	=���s�`b_�s°���qI��1���w���	��ʌ���6�8ʛl�x�0��Tk ������75c�/a�2|������v��|g��)�ƌ�Z�t;�b��~���K�4��2"��ːe��Y^Õ�Ѯ�ڲV�zia}���Mm��9ZZ��ǡj��+�x��g���?���n<G�xc�%�d@� '���v��=:[�(���'F�.��1���O+�Q���g��ʓ6����f��N��C3Ω@s���g�!ڒ���������G^R��y嶩�*�j<�����^�`Z�A������Fm(ZU��;lz��s��:9��Qyy;k�乭��Ε"{�*μiMN}�����=��z�ć����6��c��W@���U��5p��l�P'~����;�Xm�;�A�T5'��y�'J�83�b��w���,�k��^���&L��������T�y�M��^�m��V�b�(����Ө��ɥ�cP��f��՘ǁ�{ێ�����X6ԓ�4X�э�a4ֆ��+��F[�V�Kէ�������ħ�N&>�^���3��W����P��58�5�v(h�YE��R�%�j��SyMW�������q2B��N^��8w�T�
��Y����������R�l?`��;������P@W?�C�ԧ�Ġ&'�F �����*��M����a��W7������������m���쳄7/?r�۶��[�i�q�g;�d�����=�p� �6��d����ݖ�3��'�݀��nہ�u�M6V+���Y-GS�ӿk��]����v�aZ4q�#�jfj�0���d?�t�L��(�a��+f�M�Iڢ��UX�ꨎ�8���d�B��M9Sh-8���z� ��Lm��zc@���Ii)�f�8���R7mzesLEV��߃b���^$e�UB�|��^2��,,�� G����U�d2���G�7��	^�du���Q��"��+���a����Ņ���'��$ǫx�|����8m��,�2K��Ǥh�B��)�0�_\�	�Q����y������;��iV���2�K�ZWDe�D�a�3��E��;z��x��t�=�]�@�0���S]Ľ��W�we���<6bV�|�D�-���,���T�F��!r���<�6�t�͆�%�Z�߆?���<��p1�� <3�HV�W�=���VFe�F�2�.�zri[�q�(V�w�Xz����W{�yk���Ǵ���`皢�8������NߦO�f���i�ő5�!�G��+���e݅�џp:-��Uױ��ٹ"
�8�:�-� ��
n��	���gk���`��߃�\�Y?Gذ�!���Z8��C:���Ū�j�m�TvC-��n [MUxV����tx��&TJ��F��>���9drF���<���A��c_��+r�'����hZ�O�M�M��_>�?[W��h���V{@?�!�G��NY�o`Gj��g�(��_��R���:p���.ʜ�?�Z~���ĝ'��j���n��*��L�<�k����>ي�|�ϓ��,~�h���r}n'@-��Hyd�q1ׇR�Y�`���������d}��"HP��v``³�Ln�i��|A�{&{0�A2	��+	�"��ų��(o%�Y�x�,�p@9�F4��G�1��~�m��R��}hݿ>�3T������7���?Q��i|�o�y ~�g�����
��f��Q�ѷ���1��K��ZY��ި��m���f��58$���+�I�?�q�~sg�V�)X����Y�J���R��$��z'i�{#(����%��C��� �ԏ�OWG��ۤ<�������%��X�W8nlf��V���eg��p���tgԯn�<���26TUqrc<��{(v]!�tˮn;��Ԏf��n������.,-���m_M�U������q\e�Y�����7S�p��c�mD���0ɋ�iƖ5@��F� b(��?�&ѥ���VǆL��.lg/�l`�LG�1�4���,�`�v�tk��E�
�n�y�m0�t�4/��1��V��£<Na�OF8��da��L�� �(�4�;y-��c�n���;>����:Y}��#6�W|�{�cK��)2�9�x�����;΋�]ф�К��*#�di�:E�J��0�=�����_���Kq������V�YC������R�e��=~o &��f�Rƫ���In�岋v�ٰ�s@�����T����ր�M`�*zn��Y��C����c�lǮ�D�d�rr��rs������i<��(��F�4W�5��	F��F�u�����D����D��5��d�����}WLs0&g���^�i��=����=b���� ��)��Ǹ��J��� �����Hu������j8T�Q0w�D����SM�|d��a�:��h�~<b7�5 �.���b=}��V�<U�7�ޥz��p��!>m0"���6ۣ*X�۸���& F��Q� SP�0��ث8>��2%_1$�����q����'�� �vt^4���؄8q�<�=��?���7�ج~������f���r��]s�^�f��
:G����l`뚬�rjfZ9�x�Y]�"BnJ�iLvpEeHK�O������p����E3.�Z�l����Dl2��԰���W�MF�N�>=�4�U�T�JxZ�-ev �����x:�B�Q�����n�hB�e|C���ΛI:O���0�4��3bt�S��a��y��tX$�!+�����ɂ�mȌb��E<J�*-`��xG �|���0�.�i��T����w�I'��͡���j�Kz���*h�Ah��j��)%�0��7�����ķ�Ĩ�����l;u9�̲��uV��3X�s�3�>��St���oݩo��a�fxNK�����9Z:w�^�xM$��g�;<�r.`����R3'��Us]�?�����4F���t��VO�k�W|͔7��JA���c`��R��N����" �F�6L�iHR����Ga�1eNSt�Q�M�6K
���4.%5:���{TZ��HG��ܳ�!��B����rEt����V]��u�ѬB(qӶo��5ْ.}�+m��D���*cf}��>��ذ�L����L����e��n��K�ͧ�m̲��Yim���FF�+�e���'`�O �O�LdK��y'�A�H����	��l���z_�����������@�۹�樻�WI� �'7�j�]� ��fWb)�l��e���Y�n+/�=���+f\idL /1�C�u��C.tƭ+��蚃Y֙15�ǥkcXh����yԫ�ָ�'~�g|m�[�ќ�*����jQC��&�T�:�#�z;��8Ѧ �L����G��F����x4��X��p�L�4�>�ʊ�9��'�9KNp�Q��|Z�11YA5�#���d��Lu&S�Y9#R\�X����p<�r��j>�w�E�0";��������cؕ�S�E����f�/��x���͞�0���4,�p2y�yx�*e>��Ŀƣ	�7�	�ƺ��-/|�>pw`���=��=�S��B��$��Q\�| �����Ů��헓�䄆���e�.4����7tl�px�'���4�|���κ���o�ø0|A�4��=~
��h�w�}�,e��E�{!����by�������v�*��*�!K�+�'�Q����<)��2�0DOj��Y�fJ���$�A�	�XM^�Sm��윥��ɏ�DI��'{��a��̾H�����Ma*�7�I�"�M@]�����]@"��~韫��'��e
�f��>���U1u\8���T�mrXNW�>fi1��x�n�Us���i��Я���7�!��C��
��Ķ��=彡�K����$L�u#���o�d��Wٰ������@o��/�<蛼������UG���>Ҧ�:Fl�}�b��02 a�DoB� &ᴯ����_ �;M#��?n����{p����Kʧ4 "��bt�J%y�
��5�>����@�aR�x�Lc8�����S�`���؟^��f�0{U0m��xT��:���(�pU��D,"�'?X�5ԖDb�����F���*�Ԥ���Nx�In��wN��H�d�ҵ�h��TM�'q���\Rfxlpz&�<�*���c�)T錱P�!��\G��]�
�.8�(�:��7����ӟ��E'�Q(�A@�e�^��ҫ�IX�ӓ�8r^�&_� �񀹌XO�� 2�Zf�m�W�cج����:=��pȉ�k9��2!g��ϋ�-���-�A:L	%��q��<���aJ���Z�;�j��m�RC���U.��P�|����,4�����D�[N_ �O�0�ˊ0}�G�G� K��X�v��AR�b�.F��|4㭶��akw��S:��e'"t)��L��%���/0V|�tl�l��ݝ�����ݗ��������������;?��h����"bXk����$?Mq;�� lݜ�\ڳjP��9x� J\3~1�_�."
��]I,�/��\��rn�Gu�R���pj�d ��F�Ή��$�a?�rv��<9�) ���I.^��e\��22ny�T-�����)��kN��j<������j����V)W���%,�o���!|�� ����<a��12q���*J�Ή�2�G�;2�t3$Q�[�e���ο��د��<gGCf���ua��9M4�s8-NvsVࡩܳ*,(^%��T-�J�z��)ch+C�q$T��;E�z�s�*P���7�LOE
Ս���zߨ{�Ƌ.hG�����c+��̋2����C� ���6����y2�r�� 䫶����|��c־:7e.�8lT`���N����Ѥ�>�^���Rm� 1.����b�9�d����k��`��n�%!�=f��3�a�2xWvW��E��K��'�[��$�PV����\��[�mj�2b�4;	S�7dEsk����_��.���:A��AK{��<sì���{���p�"���נ'��`h�*1E,jv�u@����;�xױ5��{R�H������O�׋\?��L��&cj�w&���	�"��M��N�dr�v�qx�R�F1.�_��YJ�6�ח�h_JN�2�*��86L���|'9����4��n�~X�� �������(�)�~��]vd>з�l U���h�>�tO�t�߇{��(���罯/dͰ3s��~bU������]9�Y�?�ە��өW>t+xj�ye�������c��U~9LعMTG=*C��:%��t�S�K?���5l%OG{(C��o����S���
;����Bz9�IɌ�27d��Й��[+��Yoy��
�φ?9*�-�oX��i2�vHG���[X�/,,������Y	P]rI �_��ȷekߒο"E��U_<��i�iϤ$*(�0Sq	�F�L��8�?Ȓ�`�*����/O�c�5D��ˌ�W�\W��ȝ��؇1R�b�f���U�	ler
T%1Q�g��o7v_>���6P�qg�,Z�LI_� ���!r
H���[� �(?��`0��0/1���0"�1`�Ii�u�]���N[�e�L��	������^���(zI2��΀u����G@]����~/�������ݠ�E�O,H8@5��"�cʑ��"f�REL�E���ԁ9u��є���b�5�J9��- �oj���c2!_����ˍ��jC�JR܀�'�J��n'h�P��Ep�1���x�NNp�->jCw��q�]�'(�ϷG�w����RF�e���MQ���x	��bd�Q�h9��I�:�,6�OTrAx� ���@	_��t�MY�1g��1rT�~n�1:�mҡ����Te���1"=�PR��ޜ��vX�N�~Q������i{��PÃ�E{�?lm���Cs.���-�h��oS����\�凪j7.��	�o��Rs�1�͸x�߬)��c���뙪:V�.�%o��2��~��y7�t���%���]�.0��@�8+<�l��%��Gd�X	 �w;�Qqz��i: 9���� .����=��Q0��͆eW�T�Z��&��`��㇆J8���jj9�.
ҽ��&u/;(��W����K*��������R�x�2N��^��.��̯��R�U9��Uȣ�3�OC�t��U[���M&RRݾ��O�?b�E2N�����/�7���4=D�&ktVnQ��*�Xc	���e�^��kvA�T8Ձ�8͆��0��'�-^7N��m��{)46$�)���S\�B��]�#�_��f�{�]H�S��|��Oq�3$�w�q.�h_����>�����WC�rzW7����%Ń0-?'#C*��",�'c-o5If\�Q5��V�0�`o�QI1*Û��Zx��@�iD���ɚ]�G^����:�k&i�_p�w��s!�i�l���UE2�MW��`�&fD�'��W�[>&�0;s�ߔ�'�v��<���M�4E=2��O<8e�<�R=��݈rgu7⻶�JW�9@$>�<��f��t����m�j�����OM�J7��#'L��y�ԓ��	�E�д�=5�uJ���Ͽ����*]L?x�ؼn�����*T�ɚ:��_Ժy��u��ZX����հ���f	Cn_� Z��㼈�r���A<J&ƪ,_du��!��]��uU�d%�ax �j�j�,]� ���?fa4)�Q�Ŕ^�L���H�&8d�2��a���_*	Z�*�Vy�(�uǡ��G1�ػ�}�Č��*^s��K
���(R�춒�s����_����VEL m���\
Ʉ�$<�����[�7J�KIm?1�ͮ�[���֕W/�eP
7e��!ۜ���4u\ϣw��|-�i��X�j����4
�QﾣD���aG�j��ᘺ�?R����b��j�:�st�RG![�a�{S�,�Q4�c̨�V㩯$�6������rb6�_l����^�'z�����9Z)�Y�<7cY����OX��C�������6�0�p`�21��yM�Am�T�r�ǫ�Z3�W^*jQ׀��9!�uFD�����+bUZ�^��k�Ī�r>�^S��[b5U� ^�)��9�Y4qIbm�\\�
6IQM�3�*p5˄~��x$�Q�b���`l�������|)���K��Y%�e�N�(s_/��Ƿ;�?�o�nH�X�b���	g�Z�$������nE�V<�z��r�U�t�@����ʼ�؎�@f܏��°%+��e�d�n^��lF}�u�@�b����
S��O�6Y��U71������q ��5��E۴C�������٦B��d�DO˙���V��(�'M�*1�٫��}��)�ek��v����L��|ʌ�K�l]�)64��2�m��:�P�U Z·����`8M2�x�-�2~�v�ԧ|�O��v�o�?�͒cT͏c��R��-�.��FPՄ����_�]�����@��>O2jS٣J�����wa����!��ln�ln���}}Q�e��tS��M���_1�&��K7�A#G��Py��E�M븳��_Jz�5PRN���i.� ��M�n�놈3��O��N�W������2�v�?r��������ځ��T�5���U���y�lZ7��l�Gc8�Vo���=�7�!�Z<���f��S�z������X��L���g���NS�9�}ک��|~�m#��F��oE�mw��������C�uEE���\��!�kl_9+��W伡,�=�klF9�A��v���d6��(���]:>��t��v��N��M�����.7��|akog��˭��aE�Ƒ�>Pe8���ѻ����E�������5W���͂2ݡ�XЭ����ݷ�{��Ѵ�\���+gW�Q�����DV*|�JTe���/cJ�(T�%���>7�0�zv�蕗z���U:��A�p'W�L��+�d�#]g�&T7X�F�%<�P�%>&�X�=\I�V��:L����3��}���/&<H�ϩ��<� �soJ8F㉾n��>O��U��Ȕ*����*��G�-�:�ؼ�5t(�b���4��P�H���*����6�d�[�To���#a��❨u ��Y`��Yfpr[��n��d�f��
�z���y5�D�,�}�!uط�������>j/^+��u�P����):��$!Y�4����kf�M^�t���C��
߸3�2��<�bJ@OC�=��2�`W���Qyt�3�)Jٻ�,���tMMGi>w�}eN;�^���W��M�|ͱ����eN�aq�JkT�w�3�}Ѩs*���\
�?eO�5bx�S�}U�ES�W�5Q�g2�{�8n*{�]��m��+��s�ӏ`i��w���R:(���`���&6�w��vPu2V��LEo�w�E�5��
�sr�ߋ8�w�z[S�*l� *w����i�t覑;H�����|�|ӕ��(?��^����SX���EY8��7;ժ|#b�xz������/���^�;�+��{�"<��@6`�!� s�Tª��<��j���.ˏ��x��˕���w�X�u_��<;���Y8�����BkB8!E�v�:�u���Qu,��7�Q�ƩO�)J��PF�j���z�N�0�A:]	
Aj�z��9%P��%�B���F~�{Vs&q���>�3��Q�&zF}?͋��|�(.>��X�zX�4p�R������h�)��z��ӕJ�� f����0w5�ps�5H�i�\;R��gW4ɼC�M�����+`6���2=K����#�B�գT�"Ȩ��!�u?��A��r��#��8�!n���Tr���Q�[���6�G�)��:Oг�j�8mWj��/����}���4��KNQg�g����1��˭�3�� �f� =F�Z-Ѕ�� 6���9�?lQU�`^Y(�m/���"�,zy4��T��`l[ک�<��a��yK�i@�
��"�KY�c����2��N�p�EdY<Ko�R#V\q� ���a�19�h9'��3�8#��V�������q��4�֋�5��3�(�\��MU�q��U��G�`tM�X A����{��S'������c��/B~�NLN�'��vþ)Ytb+��_��x�z��W5�V���]�+u-����ѯ8q��@|���1�*w$�i\0@<#OB<Ԫq�+���l����k���%�NG�j�}h��J˪-���Ɨ�̘�Uoc��A���~���@�nl��PB�5���bO����7����lp	L��3	�YTg���ﱲ^li[Z��*����'ߘNdVɦ:��i��J��=�C{�@)�Y8���"�ۻg]9ҕ|7t��0L=����"9K���ŝ�$,xs����kpy�T�4nlf����U��O7�B��]���\�0�&q2��ϩ��<���m�'����B�,f}F�����.d��/��>vA��DIJvK�{%G[0��YMg&����ʕñ�r���n1���đ�9�X��]ͭ�L�تO�*1��Ap�-H	E�Z_0���nh�^ۓe���P"R�G*+I�;����N76������R�g��^Q�d��ܽ,���4��m`��6չ�7���ƘlS�FI\�� O�>�&�B�?#J:����p�"����C>�Vo+�݂;|���O�v����N�6A2:�I<�_D�(�[�(K'(?�Pp�\�F����n��R\� گ?'h8�vqAR൒�/��t���MG8�`ıVjI�Z21�7�ۙQ�q�^��1�s���R�A0�����k�A�G��F���D�t~q�9��l��� D�b��es��5[�m�[��*��*J�ӣQR���﷙mT���d��ԫ=��"��!�8��r,���\s��?cj!#�D����u�Jg�6bOŷ��4ѧ4��WK�D�҃�q�,$��Ds���P��V����ix�%k���>u@�O�pT�C���q��cX����)t����>50Y�^Ӛ�^�AMBK�.�RyX����b�pw����u�o~:��<*����z'������[�ԡ�jv�({aܽƃw����u�hDuv�jiD���uqgEat��d���F*�M��L-�m�ҩ�9P�CN6Ћ&�zN����IB��q�$�u=��3�,�֋x,�ɲh�z1��ǭ,"C�6dx>i�2�B;�]�+����tw'�xih�\wɑ:�^.qKm�x�<uk�������i�&��a2�|
Sgk���ŀ��`���v!�r�QLJ����9����P�������U>Ԁa���r��G�5��n��a���3�	:,�]\�#������gnQH��I��{���%��v �<g��>Y��G��<�������!݃���_��(�Ѐ�,`'5��D�H�F��%E�~0cW%���aGht2��)��<
�$�B���T�q�W�1�HD���3,@2��`���HwFMt�"�#�7*m�f�N�!�����.�R���=@����K�@U=�@�l
�I���EU��5��T��hEe���/�x&{��i���{��������T���]�u��i�9��&��h.�95/-!��	s=y ؋G�\���ۡ$J��lA�3��@gM����t����Q�]���:�34�ۏ6�n�m*��U����$��i�Q}2��'m���4#�F��k0N��S�U�\�����7���g��Y��v-@B�E Y �,�6��D��[b��Dk�"o:��P��G15��Ktg�c�r����G��Z#qr�.�us���z�@�vc�������Lm��܆~����P�Z前.��;f�W"�K3��KM�W�6�.��m��֥��Ҹ�f%��`c��<�<�/�0����0��d�E������F�t�	��ZA�7�n>�B@k�%�47�%&�ϩҎ����|(�/�ܴ�%��=�i�3r���P0]S��z|U��J�ܥ��h
L~�{����6B����:��?���L�X����� �a7$���?��6lp�_� ����i�%�t���0kgKͷ�6�w��Iwhj�3��&>3hl�s{Z��X��joEm�i��PF6�z~��2^�KNN�)�mлX�
6>}�s���T�%^4���[ �U��!h=�6q!��ӳ�8z��ڈ�o��a~�m����x�"�cp�)c���pM���}r�Z��@]��h������[ϗ����p��fPԨ���e P׽W���P�{dwx��8�mʪ�{ߐ]v�[��$���C�M5����X�˧Tm��=�U�˧�}">J�0��ոb-�tZǧZ&��┬�Ge�$���܍Qr�1QJ(M�[�Q<Ns� y 5΃FA%��9�o��8�Dh.͛�|�73�/���Ӭ�n=6��a�%L�~�6�p�S�%腓� ɇ���5�͝�g�)zx��=8���t�gy߳	����\�p�C����/Z6�I�G�o_i\@#ADj�J�-ܯ�g�J�MS�3���ʻ4�S�r˷��|�\?��$��ɢ�|B�E[�C?�xŜO>Й��	]T��!prO�� ��f���1h!�Q�y����A�(�{��]��ݒ!Dhi��>T�������t^��-$�89���'��c�)�1~6:F�#L�چW��jE�|z��P��Kˋs���.av� ��)��g|:^�f�j�U��i��JL�d5���E� P)�]�u�D!�����t��6;�At����)��1d
Z� Ͳ�ޣ��'��H��mv���L\a�B�'$u�$�����5�G�HQ5�Ǌ.�#(=�6�r p0��Y:<�ʰ�� z�H0��F��>��l�L�>@�NX�:[LKkK���cT��Ga�1�'� �w�7g��D�t�$&"&���_�@��GG�K�bq�ջ�F�1�e �pn�6{& ��c�I�CAj�b��	�����c���*�-��� �z�U�M�1���A���n�]q��G��g��S`+�ss�=���r���%A��p�C��������w�OM^�0�+���ۉ��u7��.�XA�&����qL�w=���X�Y
�Xd�-���G�k���=��T�ǣ���bˏ$��1�es�26[G��8����x��¥Wez�Re��`�ä8/�t�sRt�|V�v������e\�B]W�s��aW��X��Yx�Yg��e�T��C;�3��ޢ��=ю����qȷ+���O����'Øl�e$�lW{�{�:��$yH��S.'B�5�N��zL$a�yַ1L�c��G%E��la����.���,7+�W��4�t�P\L�Q��FV~���:Z.���]�r�� Vm�����uN�b��{=q@Ca�0�����Hxw�b��ʚ*�����4
�ϒ��S�Ѓb��/Sx[���R���ک|�%�Z���0��_0�Zg��0��X�c �v����,]2&�x��2S��j#�P��@TԋyÝ{yU_b'�5�p�4���w4�`CxȂ� t���=!�!��r�bi�Ng�����!����ĉ�,�4�<+�_Ǿ8��x�o����RW�E�B�|+58�o����ͳ��!�[��繎�e~���T��8b��ͥ�W>L�;ݼ'�U���!����&��������Q�� Z,1Շ+ WoC�y����c�S[�}��(ĩ�H��)��{Y}��Bb:-(>;,���t���*�y=;�Ӭ�X>��q�v���M�F!�2�7QJ�]�^Q�]�XM�!��U��3/=I��9M�o��F���0�Ǵ��Q��i�EJk��(�%W��ɵ�J<L�R��ݗ\�6��@њ]��j��,��%���
��U�n��h���MO�pZ�(љ<(s=�V������>�h,f�㈩O:mX��E���i�=d�2�~�H�r�ưdIc��c��tY��.���:�o���M����5?;iC �?���2�!�����΁r��v�������Y�q��H�����J�~F��As�5�C����8m�P��� �0'h"��̫���ejY�OIT��] )nf���΋�e����h�<8�F���W���J�����p�������?z�Y
�[Y^Xف��<,�}�VVV�x���{%x���+��#���O����d�UDV�rm�
�֝e?����.@A/T�'Bq�%W��x��*�gi�$�"^/�	r'�y�����H>O�fi��+7̜�n_�e+�ڦ�S_�[��\���;���b�:�.y�5g{#��Q9�
���տ�I���`��0#)���G�DpV��kSg1��)*�xY�V�l]S���`-@!�K���N{�34��(�8Ji\w��0�t9I��h��/��_�](.��U:)��3Ρ��?}� a�~�������FF�K�&���(���
B�77����q�5Y6נ�� W����������Pͣ�C���&����c!%*l4� ]�Wn��Q��-�h���SAd��0�H��nX��}n��iq꣞a���/o�%�eY\���?T}&?Eym&��\K�i&}�Yv�^a�v��]_}WU�X�0���.Q�mU$Y��'_9����� 1�-�V�����^�Z����s݊%q�ķ��D�D8��A�J�S#��D��m.�J�o,Ɖ�!৏t������>(�5Tc��*`.	�!/D�ZWA)�Ez�,��O�b1�b{=��� }�o �"����TP��B(N&�$gi��'c��Nv����C�@���p;�c�����OIqJ���{u11���Lk#_�YZ^y��+4��ia�`v�{�|���@�����(���/i ;[7.���px˵F@�9��������	���X��Ys�qYh���3���:*�~ �����f�G�cU����i��ex�N7�H�m�Ը2�Q�1�Qc]2^(���|#`}��Diy��̎���/Z/�(���-/�Ѥ�0�d��])KX%nd���ܷ ?����8��J⧭u� ��Q�f$Sϓ����	l���Lv����h�l���lmn�<^l��sd�ze��Lv��CyG0� 5��U��Qg=8�镕N��X[nX� ���g؝�?@���0� ��E�Jv�EW#ؖV[H<t���c�o8@Zr��A�?%�~�E�)��y�/�.g���8p}�O���lO��F�G��n��0��3��.�a��p�-:_H��,�9%�OI�:�g�������'�;�[g�~)pB Ϥ캅ĩA�d+C�hk�m���ݙu������� j�ݵ��Hh_՞�.�o8����Q�/�<�D�?�(�D��̓V����ߣyܣh_�o�c�~C�P�@%f����p�z�����Fog��DT�F�Q��������:}'�H��k9$�s���:�ƽq��Tw �"
ˌ��6�����ʌ�F�(�X��r2K(e{���G2�-RS��^��c%����4� �9�>�ki�1�6��S��6���r �
�.�/��f�G<�O\�/:���������eX�`/�����(S���@�����"c��A�{�F7��ڢ��걆8V��Շ�Ӡ2�,�a���/�4~ yG�i�V���?�3|��zԌ����g/�� �!@�~|��T��Msx�>� *�Ϙⱱd��0� �	Nˆ��N�qVE��,�c�\�Z8L�>N���s�c�����6���C;�����DS�-��2�Z�Nk��yy�܏���(����J���(�T��'��o� ��YC9C5r(���p�g!9���Zz�T��ZP���8v���o�q2Hy/c��{��j��^%�k�e��*�~��|��,����Z�y���A�3c-���(�o���~5^�Y�1��7��y6���D�ҋ��&E�U���G�1���H�2M�+<ի]���!�]�	@e�	7��aJ/:��6a�ŀ�{��|]]^{𺩚a.�C5�|{d����*Z�y�p0�'�Z'�8�?����d����NuÂPu�5"���?^/�q)�~����&%t��g��$,3O��n��KYh���`�~-u���t�I�	��uo�lL
��� 3�y3G�Y$���E<��H����5����P�Ş�c���)�������͜��Gc=�
���g�nD����xV^�S�SPa*���U���i;���<���p�>�c�cU��v���L��`9�R��v�Gk�K���X��%����6}�#j�������(�Z����ld6�|�6$�a��a��0
C&q��Z��:\���>7/~�; Aо�֭wRZ4i��fX`F��U�ڗ��[ҁ�eb��	�1��3_z��v�nP�'�t~@�ąf���j��9�QZ��S��^�	�a������ϊ�6h`l��ݹ�H�v��m�O�I��v怷�����n�
�`�gd���t���V���$V�4g�?G��b��l�T�@�6u��\o���o=��r|6:!P-�������Ȼ���C�J%�H!�Cn�Z���:$P��1-��Q�w��I�m�#�uU��(�����.����Ѻ�8�s�g��Җ6�a�t�9���BJб�u1��x�����qO�xv�+�;�NJ�|}�,�Le=�	*��APV?�)@��ZԬ]}Tn+W|n+[9w�m)��ˇ��A�Hb����C���D��fG-զ�=�՜+W�ܳ�q!�����w�Ւ�i�ܶÃq�_���YX��w� S��W@_e,�ȷVV�TR��mKs�۶�*�kZ������S�������=X'��O Q��p�v!J�� ������<�.>LU�m>	P�!�����rg�|'�v�i9U�X�L e�ɻ)�2&��}\�_Lgԟ��>ڤH���L@�4��"��!�fn(Aoq��p�&<��{K�ƒ!0K�0s���l�1擫��e(c�!o���L�s�<JU�/͐3���У9�y�s�X��|
y��D��Wf0G�I*f��5Π�9u�L�9�h5Y�8'�*��rm�����Ty%�0U��ȩre2N��@��R������=fl2�&C�9�[3�k�R��r����Tm��p,�]ÉU�WR¬'h� ~6�;y��%���)�Ş)�������w����#6�">N
.����N�`�T�܎e�7٤�w[��qJ֑�M�l�|c�еT���gL�,�6wcػ���5[�0Dm ��1m
��y޸�^��L�n	أ��Y� ��˻�Ŋ������Yn&�'4���Mv�����j�����-�<�z˔�c�E	?�M� ��~�ܚkW� x�f��~�t�����[ɦ��sۉy����Qt��9����"���~�0i6�Z��5�9��{]�6�?�E�/�z���~�;캉�b���b8�CV+}W�g9���YfUc<���n:���` ����`�����l��S�ևa�͈Ѡ�)���}:@��=�p��Y}r!�����8��tM�aPb�L�\"���/(��w>���XEE-�pb�WX-N�02��"���������uP���{E��.V�����+N}Kh!v���ō00���bo�~��,x�������=�(�cst�.6^�}�q���K�(��t��]��Z��4}�4�幟�_\�鄦����;�����.=o�3�'�=�U 陉������(�����[8'�Gi��ǶN:C����WM]��g  �۬]|�ŠO����xK��;b�Ť2�Q*O � �A�k�z
�c�l��t�~Ǧ��,�iVt��<9���#��Pw�1z����Tլ�h��y�.f	U����ӝ&V��9��e�
�(�oÜY�<�v��U�fv��×���qf������X���ONU�ů�LӃ(�y~���C��+�(�?cfM�6M�g-����1;�)9�v���!�]���{��V=/|؊@M4���g�t�}��Y����4�ۥ�Zpu�`1w0Y�`�S�]*
A��si{��=��,��jH�w����U;��������P�㏪f��p~��E���q��(9�27���c�H*.`��q�!'� ��\Hӗ��|#�CK����>��ѡ.���PB��p˚ΤXx�߱���C�Ry��>96�� (+3��	v��<�-��d3�uܭ�����#z.(������XY�:�B]����w2*d�Bl��t4�l�b0�	��Q�3�k{~b5�-�O�f�4KH��皴� ��"�;Ag�u��.ں�����iV��������<H�m���c��\���L؉�4G��8�uG�.��C��H�����h�GI���m�`��wU{_w�ml��a��p��xI5��tҪ>w�n*>#DP�R����V$����!J��Xl����c��1�m �M܄����E��ӵQ��T��5�s�Z�+�����j�\}��J<~Ӆu��n�i����J��0�������!R�"�M���EI�07Y��{W}#��|���N�Mj��ɹ�g�늺M����b�o&>6��3N�l�P�H�¦Ӣ��N��O��5$��J�y�Ÿ�jf8.i��͸���]u�i;U��j��2+_���g笨d�^�}�O�GY��#)��H�Q���s���q�~rZ�a�M2�t�pzZ}����,��;X�5���7	F���G�/��\n9x�lT��.��5���r�zZXZ|X���j�.�0�4s���U�,�5�H-g{�ǧh� j�'dp����$��"��	��}H⏃�4��$�F����(�h[gQB�	~�x�Z.�y񉱡>��z�"F���	�^o�
��	{L>�����1��8�ˬ~��ͽEG:sSn߀ y��(���EgU7�}���Ao�qne0F \b��-�ftPu�+�GG>�*Ӓ�Д{AE���q8�(���6S�8"%>�P=r�����ѨWM��#��G'��d'9	i\j	�{���0��.h����^�d�\�z��«�9Lt�la�쐚��ԗ/�n��",���lMC��VCt�?�H0��i�#��^�'�h����C�����j�P�t�lX�B���$�}����4��,I��\'����H�dA��+e��%�NЅ�  ���}�rG�辿"����n|�a�̀HJf[�$��	^�T
dY 
�*H��\�b������Dx5q7����{N>�2��U (Q�2�-�w�<�<�\G҅���%�������v�3��!�`b˼�`Q7�9�$��ʒ�Ho��$HW��ʿY���J�3���K(e���%� ��4��Y���W���).��ǚ��c�?-|[1�Ƣ	l�J�0<�Z�0J#Fo���?���:�̣V4v�:��vW����)s��x�b��������!F��z(B^���4՘/�
�,��ܰ%�ǵ�ޏf��N���Y�{}��d Z'��e�8��1H>�0�gc�Ĳ��U�I�Eq_ƨ(e,�l �L��Ŷ,c5gRU�����C@Z7k�8DJ�G������[ ��]���8�R΍��j�JH���᪨d �
B@s�0�8zX�F��	��$9X���nm`D�7����P����:it.f5͗ҫ�u\"[,�j�	od��*�kj�r���������qP�F�/�RV}�h�����RU���yO�)?Ǟ���(�|�9���T����m(i`��(��5�Oڶ�s��|��蓨�h��ƽΥ�zM���#�U�����ݽI��zA?`�d1J��K�Wt�b�_�����HI~Q�!�"Z�GF�t1f�)��VY��}�n���QJ�eA�r9�nN�A����h����v�R�ŭT�d��R�6�)YE��&4�]���SM�mݻ���v}�����\K'2;D�R�K�&e��h������Ա̯�E��?0���h�'�e���ѭ�L�^L�]}}��vN�Es�n-{�ܢVs������I�=��u����׆}�:�(�7�gOj��lV�9�ٛ&	В�i�<���!��-}�ўF�o*����~"��Oi{ ��ٜґ�،NGA�Q��b���<����Q4���?k��Dܒ���4��RW����M��d�Y�<�.�d���ӫ�A[y��c��b��	kj�1�^@@����!��w�,���7��/��Mc�|dP0��4����ˠ^B����)�7l�8�8!�@����UQY���Ru���j��N���K*�Y�I�����yF�-6����fM�qx�M@�ܪ�E�>
5Y�ޤ�fA6Mٌ��#�R��w�˻U�L&	`���w���ôy�$�	��">{��<���`�V*�?lPL�b(��1z�ϼ��|��<�����:�06�t�m��esiR��S�^l��M��<����[P>�}�/Y���M��M���M�jM-���s��>��{3=r�Z��VVǊ��� �S�T�l��@�ʹ�2�Xl
�� �N�0���>��ZClFIo��`(p�C������W���'�����~�`�SX�ł��������V�	��Mw���;���5��J2���
s�qj�B���z
e�s�����ǔQ#���o��|⤪Ϣ���(2���j��qR��|4b`2	��T�A�e<�W\kGaJ���0���z�n���\ 1�ybp�/Ƚ>��6�����0���9�˄s�:���V��?OM6��w���Q��
�G�A$��&ͥB�^EĞ��w(�j��E��ve�_f�T�a���<`]�2X�u�m*�T��O��n�4�,��!�D?��F��:�Q�bߔ��I�/_��ͬG4��]9�Q^���a �ޕ�f��g���
zF�X/Qt�q ��m�5�gF��J�W0�F8Z:��%�Y'oJx\ftƙ:;V6�$	O�q?&�B/�m�wmv ߷+�
[�*�G�����=��"�A����H0��Gx�D���'$�ꟷ &�!��h��	�ǈ�[�{�����o� ѕ�H�n4*��*m4�5�d��ߒ��90wp�����`�橵0 ��WoG�d���!p7@�g��/�E� ��XL�\�0R:{{��onq72���2g�ۡ_�ёe�\���ZE�Ep~L�\W�������o�^�����{�B0�s�v���<$��^�m�p���[Ό�����y��r�ג���
�>N���pKA���B�6���������oP<��Z?J0u��w����������Ii�������O��v:FY��j8m���L��8�!����mw���rE��&����򕳟�;,'��@�+� *�r��b���W3� "�Nx����$kUy���߶�3��d��2�B�l�Ӥ�"��t���^�����Y�t�����\�5���}q�Qɫz�4q���T7�M��[�]m��P�lr��U�s�u����=+2�kF^���z�hL�����R�������:*��V}�\��e�ws�&��C'���1�i2J#����7�\:�f'�~��Ŭ~���]�{�ӭ/ �	tf�P�5w޳3�9�=�c��Y1���P�;ct�r�f�������6`	��߮wn׵����B�u^�筷TJL����1���ͽ��g�K��,�{mpύ����7ӌR˃��T������5ҕB��IG��D���ǚT��}����<?�Z�0fq���>~"&i�����p]����Ǧ]H�����!�t*��U�nT�b�K�y��-�qB�i�iǪD�+��Ὑ���z���qƾ2���x����% 2�HV���}�W�>�A�����@(���C�O����G���z�YW����W6$%@��j�0�S\Jrj!�
k�K�f����/�?\2�oK�-;���^�;�9������X�U�/�۽�w*�D��`��y@��ε���R��J^�r6ݛ7���$���i&�x�c�NM��G\�a����zƉ��Ր!�9msi��KM�3�I<ڊ�`��1�5���b�g�u�5���g[)�=�['����~C��~_'�3��z#=�A�a�y=�X�E�th��L�(DeИ��!�8�Q��Q�����n/������	Lc�.�q�E��.۳w�yfo?O�j{h�J@��Z�����T�0�k�JL>[�Tg�>c@�o�&�=�&��Q�0�!P�vt�;��g(�|i"{�8�(�~J�������u����?�ȣ�d�`�p���N�)Y��^�1WE@vnmv�u0��O�+<������8Վ[��H}��7L_
�ƭ�p�:<�:|6ǣZLK!j�\5�j� U�l�q��EW���1~4�5(�>[�,�s5U*ʹ8"5Ѽ󔨂u��qL���|��X'z`�P� �h>?ݞ��q�8�gp@��_��޲�WS�Ҩ@�K���e"���p�ZL|v�;�W�.��~�2����������$V0��f7b�[[y9)1�y.���E��yv�I�4�C�����
K:���Y����	�E����9y2E�El����]��wb���c[nǥ6���
V��̚v�<$��4U�l��DV�o�y,mӀ:nz��"Ύ��	�-�z�3v8#޸������,,!�̬Һ�d��a�V1��><*����;�>c3Z�6c��g|�i�3�����M�h����g�F��Ci����>��Ժ��Z��KŻS^oq	(t����SV��u�v_����i'���;�O	A�� Z�
��b��kx[�L�"�;c�G�\�]ŲnP�\��Q�����fPva3�Y)���=]Ux͒c�O%ff_�p/�cdk��H*����l�����Ӹ����:S�s��Gy{h�~�g�g�G�A^y��h�!�]��H�f�[�.m�v5��y�<��F�7���A���?�Um�wAa��d\����нU{Lo�'#��#b��=0�6ٍ�<�9Hߍ{��Y���o�(�]��➝��C���2��{�t{�hg��a����˭��2�x��ܹԸ7fpUn��ا9X@+C����eQ��2�/���	
Wҏ��	��/1�M	گf��%ld�;#�0I�	0���Ic��f�Jț(�D��6F������n�E��W9\�!c��u!�ɡMD-��n�LB�9�@Of��L��¿5-��=/÷��½�@��0:~�|g"��`��xA�Dɨ��
��'�C�ߓ0���X�����ҧy��e!��y!�1�N'�(~w�m� _�j"\ii+�R�-]��}�i,湫E�嚊�I�%Az6'-�+���aB�|���S8e�PZh'$y�ْ^��@˘��Vx�Rb�j�*����'�C�`*��	jU�\�	�c�y���|?��P����:&�y��k�*�q���5MȎ%P2����4�4��K�,�9��䒈�s�X:^5�����v�e�;�B9�P	%���p-u�<�� �ƙ������Jc�*���	��s��:V%@��#u]����M���h����^��ia0*�3&�a0Z` �!�X�w�&��*ث���DD�\_]���D�4�����4��p �dg|�{O���>�05?
����RB_`l	�<�O
��7ҨZ������^9?q���a-N��2U�fV 1/������Hf/���ђ���\.�x�$���6�+��t��y3 �,�~��7�ٌ"rZ�g��3���g�&2�I$i�*k�l�����n�Y�q�+��G2s:s� ����4���wy'Ԇ�כ۹����F���]���������,s�Ԟmؙ�>	&�7�z ���R팩�>n�1�A�$Z �8
(^H^hW�N�$k6�erB5�U���h$nƣI�����6���0��Պ|�I����!I"ߑUcE�f�L,YQ/�/	�
*�G�a��f�g�!�Uc|�b�N���Č(:X�"N�	[�nU6�g��� =Ss ���8A����_FA4Ƈ�A�ձ���Yw���w+ӵnr����h7����I�䛷E��͵eu3�A��4n�3�=K���;`�8P'݀�o�hb ޕ�;�IC�
,�����d���i>�BU� �%ϳr�ኽ�N-�Ί?�Ѹ�Xi,]^4�#�o�]�"���¯`J0��oqB�]˷�y��0	�}��Y�A���1�O��RN�u;P�"�q�R<�9��n��Ƃ ����;eq����ۗJ3��9]��}O�A�&>)���˄�&�-��Yw�cR�QJ0X����b}��6��Uح��ɱl>��8	oJ���˕�ՙ�P����������%�(��3�o��T3p`�_X�5Zfx�[�'ɜ_�k'?:��ү�� �3��e%p���bs)J�U�k~ܝīn��dQ�����r��*Eq]ݲ��}�ȼ�Auk�?s���m�E���R�m�㽃�����2�-Ӌ�/C�x9���S����&S�W�)����<fj��ߋ�(5��r���9,r�ׅ���S��OQ<ym��+ȫeϡ��6Ս�����%�#���\v_��ٵ7�^�#�$���|�ΩJ<+l��u�Z��͟>���*�i�'e�|�r����U���]�ѽ���&�^Ԥ�{铔bєe��hE��wU��1���vIޖ+�p�����ڜ�/���8L>
��z"z�2���]�ӊu������]dާ#L7���5pz-�0|ہ�3����a>�k�"���0L=B�� �$!��O�!`�������&�H�*��Rp��D
i��(��ܞ
�G���������n�a뭤��l��0P��Q��1 ����꟎:�I��]uaU�*W���2���� dX��'�a�V$�K<�Kwa�3�v���kWm�4�?�5s��)�s[�m�=��MTi�N����g�����gU =3Y�
p����|f�\P�������D�Js<[/�D+z�'ZσD�zN2�'�"U���\�*[�@�\�I�4��mnG ��{�fP�EK���v"�VW��J�'Bl7��yb5�\���E��P�X��+�c��L֝.'��S�^�\\��0�P6�4����V��R���u���X�A��S��?�>{�}�L&~Z}t��߁�����5���{PF;t����n��͚e���a����P
d��WS�YS��Y��lCg�'��rI���4�b+ߛB�m�"".�
���[�?�ҼA�b�������^q�I�[БY�]C�[s�����b���%E�5A�=Gj�굊�-;�^`����eXJ7�Bk�;h�X��ߛ��c�i�7��z��9���s��O�l�����Ou� w�n4��������|dh7�0R�g�� $h�&ڻ�Y���"h��X����|r h������Lw�����!��{�L�p�?4�~D�禂l1�? �֠=��Z�B<%�A���8��8����S��̓f�z()�Is+
ҙ�r���= �����eu�_S�5�ٖ�iI�;���mH`��`�g��V�\��&�����ҥ;��/Yjqb5f�5�0'�E��Rհ��X~+��Fe��R�*���0����,���(��H ��)�.��-u��D�����  ���B ί 6ۏ��x3��y�
�ћp�z���m�WۄOk��}�����r������yZ�9��r+�ܠ�ے�gV+�k���ձ�sY���r����=�4�L<����}J��8O���Q�����u蔱-�d�UP����,`����#��\�����{Q@�$�@*��*�Cޟ��WW����,�m��Z�Ѧo�<k�4���R_�Ʉ��m��4��;��q[4����D����Rl�βl�vV�2O[p'[㸕����臃`:�~d:ݓ8c�Ӟ&C�.��p��5tPT Z���y2$K��6�K���H8(1���/,un�J��ҿE�Y�n�Tci��)mklh���V�]ݾ`�]��Q˒%�Cޑ���~⮊i W}����ub�JfNNE������޳�;O�I��7��w�5u�^�XF�:҆�o�/��A0L�aAY��~7�`� JG��ޏ&M�*��.�(������U��QΫq���(Ȧ,
hp�ƯՌ��'镈��2M>̪Cn_�mط48/=#��d���{�ͪ���U��4���\��!��O�a4w�C���M�%Ք�O�*a�3o� ��d:
(}� ���eq`V(ua 3fs�&�;��@��}נ;����`��}ͦ�}�~�}8 Ġ���"��kq�B{������C�	��׃;13���wa��1FaA�$��47�W�\�R�����$��&P���I�x ���~%?O�|�jD�|T �0�8����;��뫚7B� &7��I0�I ��$g����A�4˪f󈪌�MJ��Ȅs|��:t"mJ���\�&Iܧ0��;�o5
�Ъ=��%^x��]�Fg�W��P�'C�-��3�LI��af�, nv�K0{�p| 4hQso�����탗�������pc�U��2������S�U�׭8i�Yk��m}u�\�K��C���O�²���H0��iMA�3Դ%0�M�t�\\����r����uyu�vD=+�le{�9d� {~��+i�l��+�e�Hƹ�b�=�]f�{�U�⿝$׼@��b{S�n����6@??��߀��7p�� L��n }������D���v�LB�꿓(NIH�q:���h���-U��S%���x�6��[����嶐�� ���ɒ���G�X�e�����_�|�9�h(@���آ[�nd?�f��s|"^�txޠZ�q�ؕo4�;D�a�ӑ���~�O@��'�d]��^%�a�gK�p ��������������X�#4Q�~�������l
f�������B$A6>��p��v ip�b�SJ"�.�կ�(S�؅���O��I�\��N�4	F�ɕ/���̯QƷ���M�l��M$��M���D��/֌O gx�*��4��T�?um~�,7��n��3�u�r>71����k����0�N%�'1���Q�$5<�1�\��b�@�:M���j9��Y�����lD�����,x�����I���=@��˵�;���B۩��F�>)M'Wg�ҍ y�?�H�� x��p �����̵G]��p=�`t��Ş�(}���C�8��7�h���~��Y4x�:af�j�dm�Gޙj?�%x,O�����8h��� ���n�d��4#jm�o�^��H�ٕZ�i�Vt���6�י���?<����k�#��,�wT/�*��O�|4?����j!�K�I��gv�4������4@����f?4�DpSB�ưCQF+P=n��Ȩ#.�o|�?���tɑ���Y��+��h���;�X��><
����!����#g;w~e��{�.f���կ�Q�ک��F@٢7prmr�Qe$� ���
�y
�4	��mp�@�����\5�����i�µO3$8��؜@��Q�b�g:�y�G��Y�JT���u)���T����l�]����?�H(G��}�T6s�b�N�_d��8L��,~�3�)�e�po�M�@-�m�97T*����
�E����=��M���\2��@,��	;�dc+:ŭ��@f��%��޽���E��j�E��ȅE��]>��N�-z��h'!��+�{k�tEB��<�ć(�(y�1\��OX�4:\�gB�r}O�m)��& N{��� BЛ��nK����9��:�U!�tO�A����t:F�3� �F�z �'�s�++�q��h�CJ�I�,�&� �;0(�{��{�
 k�5��ج�(�7m��,�E�L{���kN>�ɇ]a����q>	�+���2�����!Y�¢��Y���dS�n�?Qf�4�{^�u2/��g�����3�fcj�R�N��j�ߥ�q���ā�0�fgT!�J6L1n4���rX$@{���*�g����l������@�@H���8¤��a���P*:��Q�Sc
,��i��`����e���Y��#@{#�4�sR�L�� ~��	�	��>��� ��Ȣ9��ce�.����l!�Y�0vN�û�v43���g8qe����6���8�Z[�ĕ@�t�]�����.��̾��ر���'�E�d�j�EK�j�H[��֤����GE�ZҶZ�X���h����6�ԝzu
���d�$L��9��ZT���0����Ɠ �,"�I��;�Z�t���r����4�C��}��&*t�t�^��F���-�^�2Of�G��i�q����L*5��T�;TA��4:p2{����������lM�|ƾ�<���,�g���c�G���3A�|P+5�/TK��{�mN�+��bQ��Ȟ�=sʥ_�	4:q�@�b�"���:���#��Z���W�iiߖԫ~�X-���ک�Q��!�[��=D{��oUk�ƢX��E�$��՜�s,>�<�*�!��Zl1�ūOn����(m�����b��܅�{W�b|Z��$���q��6Ek��7=���ɦq����<yy��,~���&pO�%<���}x���͝�=k�>Q���|��̔�V��zE���/�W�}�4��U���u�2��Й��˦����.U���^.u��r�Ag/��^.�:S��O~�y���v�s�I�f��˥>��j�k�h��bR�s�Y��j)�dͭ���Η�l`*,,8��ȪV4�r���˥�y����ɃU���.b���\9��r+)QA�$�,ʆ0jW9�%܈ѧO�L�\�y3g��Oi~������y1%>�q��*�&��%��(���G�Y]��R�g��m���3�����/�Q)�kVr��8�:&�a�*��P��ͷD��z������hdJ��
�]��N_��O!{g�a~�]za0�,�'��Py9�X�����L�0���f d��1V�gp��Ƶz��$��._?��>��<�ƯJ��+h�Fn�,�gg���l*Qn&�|`~���g�]�S�[�=}�ŭ�p2|G�;]y��Զ`A��8�$(���4I`�]i���D`c��6�Q�{M��w��1{o�]���LS��k'��Ͱ��ee��M�h���q��w�cޜ$�u���n��aSz-��m�hR��a�5��z�+�6cO5�˥� � %B����E�)\�5�	�������-��2nCՖ�������_����*U�?v4U�G"�ӵ'D�17X:��� ��������j�a&�0}��B���"qì��Dס��(��8t�"*�(�zֺG޶�VYU���W�|����;Ʈ-�g"a�!�r���K��B��sZy3��
g�;h���
�=��3�H[YT]���>�^����͓�Y��=�U���o��m��w�}�)u�s�&F���J�y.�b�_����>l�M�0�q,m�&�,Ͷ\M�72����53���w�ſ���Y��N�e7���#���_���F5��Ϙ�	˓0	ǽ($���f�d�<.^�Fb*C�z>��H���EY��m��J	��䤵����r�{�K���$H��:�c�b/H��Y�f0�^2)��}�a�[�%�w�!�~O�S�ML7�w�q�S��G�3�q+���F��h��'�}�\��F�\��85��n�:��0���I�<�/W�o���ax�T{?^��o�^�v�<����i��A�O�Q^�薰����Vm$�z�,���Y*��D��^Ό����;
N�lA��!.z��5-�"\̺l�QC+)5�"r$o�'�������B.U^Y��⃊<6&�1ؒW�S7>�U�j��7�ѫ�����25���K�ccM�Yg����i0�N�.�UM��Ta0�J��5?@$Q�����s�m��#c�9���&Mi�e�eja��15���V@H:=y�SH;DO���1`��n�%W���à�@�~0����B��E���5
�:.SK˾+�s�M/S\��.�ֲ�snC��;��D�g�>������Ti"T��MbJ����H�*d�*s�C���{��ܔ<dCХjAx��D���M>K�{���e����B'��am�_0��&=��率V6����>暼}��|y�z����OO��fav����y����HZc�X�l7`�<�#�:�U��|�7��5ǣ��n��Z��%c�CnA�檗� ��7A 4{�g!>KW#M��J�9\�����K!b�@;n��Kd$n�<��&����/��AKy$&%p�w'���_���$�C[��l�r�2�$g�G�����X՚öU4lS\����]w����]��O��ٺ�y&��T��nE�rRrcU�2J��5�[ٚ���PR��u�Zk�WS}}��\+�+����G�$_��&6#G��=��\���P��!���,�F};�[r���3g��nY�"���t�BU� 0q��U���wJ���xJv8�oK��r�F��O<ܒ���
��h:�,��z�����q!=r��	�Ө�R\L��`�\f
|N(��8��KP%ނX��p?�`e�}ų%Q���������T����C���T���i��g89��d͘\q�<*W��{{Z_3d/I�2��ղK�߾�v��_�nv�l�u���k)�d �}���;�/&��1�rh/�b��=��p�,�x�x��,��a�e#$�5�g�i�ȧ�ƳM,��=��4T�Hp�й�Vdr�M��g�"�u�I�ML��5`R��&�������3���7D�UsF$�NS�B��t��1B�P +�l�->R+wV-re�,�{+(ʅ,��#�������<�M�5���Ț��i»�@�9j�h�e	Š� W�ͩjJ}C���S�� r��%�T��7�y��|�̏���/z�(��S��~�p85���� N�v�9�Z�o.8%���L��@v�82^v���r)g�Y����y�|��r���Q-�m�3,��z�$���3d ⷡ!�V��͞�r��=J$����B�b� }?�B�^��������J{�����ۯFԗR̗��yo8�y;7~���!C��g<,o<��m�%.�� �3�C \�4?�U;sڅ�%L_����m��U�(�����}C�\�j���=.ݱ qJA70��F�a�ֺ��=q�}$<,^��JDY�6�t:�3G��F���}�eq��*x�k�:�Qz�ӯ�(�1C4lQ�R���5h^��%a�t��������x��2�9Q�ob��ٸ�����Fy��W�/B���8hS��ai;7�%]sk|g��֣c�bޔ�F�5�����@�Ρȿ%]�7�����`o�B/�sI�az����o-���鬼�u���`9k�Z�,��(�9����}��z����*�܋CftE�[@�P?�Z֜
��K���J��VI�+.����e��V&s~���i�C�;�On�#m�ٞA�X木�� 3�����O!�F��?1ة���1���AF&��oo"C�b��� �dK;�we��]��yQj���.�٤p�id�m	�q�ť�V���c�݅�#�Z��H-�̲?眾�G�v�⥍/;=����#@t�hI㓄�G�$�ğr�O�A��������hVJ�e���K����Q���^�t���&E5[J32:���ڄ�t��f�/��:��n�"E��w;�y�����30�(�x��q�����$^������hU�"C:E��Y���W�������%��S�ۦg/�l�Ӓ���#F���_�B���q�6i�,�2f<3'��70Ԧ�<ׁ�C4�S�t�]2s��0�F���'��c1U3�r9XŅn.ԛ�t�-z�2�D"S�陭1����R��W��i�|�Gq�`�-�1��i�[ЖbW���'?�4�� s#2:��:<�D\!zUU�+щ
e��9:�-l�W���ğ��q�f��l�_/D-�$�~�g�:�k���F*�XD��t�I��h|GMh���&ѥ2F�����o=6v[`����E`�0PpS`�[�OeԪK�� g���O�1�r䤦�a����� Z���)�pr~��xP88����.`0ϯ~m)�I��\K�'��IQ�3x�q���;������'Z����,(6O��s�1� j�G�^8�6ړ�@���I<N��RaG��!,�9=v��Տ��d��W^�\��:~S#t)_Td0��5�O?K[P6�#�'���������qi�"w��E@ ]��#L�:
�K� �{@a%� �8�!N
�(u��4_vJ�����$����1�Į��̦�z��Eu�<�3�Vq�Ց�e����*G�� ����Y�W.��l�)�뒡�m=c�l�lp��M�$��;�?��4=w�v�X"Rww}#��$��V��A{�f`�Cj� ��ńN������O c~	�2�g�M��'^�����q��Q�3�ѐ�a��4A���e\�cV�3��27�V����u�d�;%�~" O��i��=��G�v�MJ�Z8������6f�~�}x��`f�^;�n�+s���٠�}�K}%/����n	b�>�+>�2���d�0�}���X� �����b��B���{�c�X�0,�7�o9Qs�&ѾU���W�<1�藷�L$A�N�;�$	R*����s15�Zz�#���İ>�� 	!�F�q�&ĔE㘤�߈ + ���/2��xePJ���c���lOP�V�>L&��gO��_����|B���y���}����\����@,��0
��f+\�^?e���e?k���ϋ��by�02Q�Z�(��1	�osE%���+JL�~�D3 :4*ц�n�d�kC�i�������W˓����G��ʜ�lV\&T�u��E�^
jV��5O`o��á�~����Ma�����SD�v���^�su��f���nV<��[W�ύá&5P΢6[:��-�=�ɲ�AN-����י���-�/����+��%9QGB����.U=�ʐ�/����0��؇�Y][�{g��C����h���X~��U�y����Y�1��k@p��]gb�6�D�g�:z'O�(�RW�ʧ���.s�w�u�u�3��^��,��̠���;�ϟv����?���6y��d癞v{�1ýkQ-�q7/�N�_��濬�w��˵��2MT`�3���c��S:�򭝵�����$�^����Hz[T�N�	 w�����l3�2E1R�[_D
��0W���L���Cy\�-Z0R�ќ�\_�^pޢ/TL�J���w=�o|PP�� �r��G_��[����4��$� �@������{��-񕩑A�1�(E$��( ��ŭ,!�$�����6���>�c5b���&NR	`��*?���:1���p�U�<��*"/cX�4�X�{�el��d��ص��&q�Q8j�A�U����%���i���vc�Qک)�8��R�u�)G]�zI�
��*�-�u�VW])X1��Jd��nR>��Ű_��Zw���]��(_a��hr��h�Ʈ.~o��h�|������������n���h�:��B��6�^�}X�|��x��v)xQ��j?�aF֔\��,��*F�@�Zd���N�`�xh����.:7�qP�Y����� ��i2d�^@�Ժ%j��]U`(¹k����{i��c0���X2�]m���*��,H��pz2�2�@���q�.+��s���ɂ�U�@u:{��C����?O# ��y�1;F� T>�rB����6o��A@�;�XK[c�]����+uw�����6����I/ҧ��O1��ބ���{[�D�G�&���ɑ�G��ZDΞ�R�`��g��5U�d5e���r<�(�R�UF�
� �oB�mĚjܬ�&����]A�k#�+����L%��?�=I*z�ʋ���� �Z3	�@=�z@�JG>�c��? I�<C���Z}��<�6���8��F6Ԓ:��a1�`�Bb̸^�,d?))��Y+7%j3M�7G��449}�����Ůp�GR�0�z��%��R�k4k�5l#N�H5s�=�ѿ1��[Kmf(sԬ"@��Y�E7w�	���E�MQK���w�r$�;�ӝg?��c	�/7����C�9s�E��c��M��d�j�Ã�YGqE,��G���T��(|+�����C��t�O���Q?����sQs�=�^���ڠw��`X��f�w:mi*i�,��ߥQzȾ�T��,���*�z[�&��T�ca�[L�ڭ\���N�)��M���n]V��g<|. ����(�0<�9t^�
������xD�NBr�N�n^�3P�_�h轘�'B�4�Y*e'���=c�&�큩�Z�©e�JGi>Ij�&,iT[����Wڅ��>H��st�tl�2��.�0�ɯ��b�1UE�k!����r��.��ۭ��X6�҅Jb�����������cf�����wH��wE'��Je���y�q�[����tF�$���
i�Z�8��!>8��̫(%������,-�Y��Jh֛Rvx8������a�!��A|L�1e�O��}�7��cvHs�!i��2�6:$��4�/yLd���Qr컡�	uQ�I��O�$#HR)F�/��YyH�i�4:����r1��.;��a+�V�8��wv&�ed��/`�0��`�oÝ�Ӥ`�ٶ������F�YO���?��j|�h�q�O�,E��MS��C����w6a���i��f ���I[�#&�)�4���i��4̶"8� �	SO`�I30'��+����>
�XU��L�g�sVA�c�[8��՟I�hu�C�'8aV?�볎-���`����h7����I�� ��E��S��V�.-{a4l�ݭ�& �U�%���g����`��F�/�5EgKd���N�9V���|�C�>�ߦ�lvAhʂ�9f�H���/|��j���Κ̀^IE���ľG��^�p�Fq�C�b߅i�B��n�\Lc�$!U��؝�c���1�VB�W��ƜL�q/`����0Zm?�k�}���(f�7�{N�����Ol��1mJ�:mOyF4:�oJ�ih�أ5�!���_�̤fA4LI^��Z��%�1��)G^��dc����[|'��>�0Դ���R37m��@�����	D�����Q�I9��)�#��H�(��[�c�=5��ʂ�1&X&�W��Q����8ɡ=�o�kX�j�!��n6�I�$������஺��C����C� �Iw��^�K�i7m��J����<P� r�i������Z9�Z��Y0�����eުZ?��P�#�M̾�#������xþ�/�T�z�~�>��:�
H�o0��8	�LD�v�M�8
���~$��x&�__pVM���I�uh�#����#��`��=��gJv�-;���,5���j�1�&^�e��+����w�������r�����3���!\dxD���g�����=�#��(��[G祶�38)�:o�;��|e�n0(ITv�ؐ��UpȤ�5W��Y�5ܛA�lğKR��q
�h�\S�;��{�$*3�WE���[(���]�k�B�dI��Ԕ:�k��߀�Y��7��E��871�^�٠�i�M�<��S~��"�rKl�IQ���
�����ګ��Z� X�\���a��q>�t�~��U�-�Q$�������R%���@I;%�*W(]������{	��bX�;J�VُN�-H�L�<�_�_Ή3DJ9�er\┗,�`L�x 4�V�p3JQt�HA,:�ӊ�'2�����0M/;O;��e��[�=}'|P%I�`��
��e̶�}��~���vC���]�U��lI�4r�TM��eD5�[$>�� �s;��*+�����I�3
u�2%�4~+0pP�p��f!��5�ۃ���1wGޭ�W��HK�i�V�!��`^F:A��g��7�ݥ�GA�L�(�B�[j��]�%��iri�Q�a�d`��u�g�Z�MJ�%�GEI�:*���<���$JxV%�z�>+�������6�{���;ݧ�L��O�/w�G;�ϟv��ܽ��G�b���n�A>�Q���.'
)�g��8۟/L�߸��oF�T�'���*$�VB�d����CAB��� %���VZ����B�1��K���q�_��-s�Ze�Zݴ�?)���*ߊI���,�wx:FY93������5������[%��/봏B��*E���"��΢�����dwT6=68�ye��&�8h+��N���tNeU�{�DY�C�Bu9�j<rb�G5�v[�`����{!%�s�9���xa��t��V�`g�ewso�p�؅ei*�u���s��07|�����iE@F�u�x�\�����y���w��+�U�ㅓ5��q�>ql�����i�0T%f-��r���[RN��-��D��\")���V����Q8T�.jPỨ���w����m+�xO(i&������`�����}�0%�%�$w�CI,�%nIꕲ������l�T��G��5?'�ڂ#�Z;���<C���V�1��u��#J'�\̓�K�RU�i0�Ŕ]y=����K�����Ʉ �rq?�kFk��^q��l��8y���T����ÿ��j}~J���,~Ij���M��I�V��7/A|)��&����|=��e�ZI�K�si.ݪc�M�,.,���Jt�������1	�:;�{�����t2�0Zlc�x��e�|����������o�#�=���R����`1�u���p6����+�Hj+_��m���a)a7����[���w�V�e덲��g�ەu�VF�,���Ð<��ٺ��L��U��aN�!��-Q�+0����?a��SJp
z���[��D6/P�*F.%�8�V.^V�e�
4�Ga�jD��WF��s�;V��0f�{���A�>�"��OԎ*_�rL%����!����E��t{��e[�EoE!b���w�f��S�%�D:LGJ������Z���vI&�t��^n�ߨl�Τ�#�r�Ց�S,]
$����)���Ly��y���b�}c�����G��;��܁�b=t�E�6~��!?���w\eW]��I��zc !����ㅄ��GH�7ز�k���v���G�a����NB��lP"�ܨD'�X�[���д� �)�-Q4�y�k�Cx��n?n9��kQ��!>2P����?   �� *}5cx��Z[��6~ϯ`� �ۑo�lҩ���d�ۇm���"(Z��l$Q!���L��>�q���=�n����}*d��9<7~�X��;�^�k�c
�xX��O�-IV�~��Ђ҆���s��/<,��j4*qA����Drꈁd{���~��+��&�'�2���u����-
2T��l�C���7
��Er�~��ROr&�WÔ�^15���wK*	�3���It�F��(p����S$��ew���헆�Ֆ����*l��s>�(	w� L�|��EA��ISm\����Q�C-U��S��l�y�#�=�/_����	�LoI9V�̝e����E`a�, ��H��'�]��k}�0�K�OA%���b�K�� �d�R���e�:����ϭ�F�����+�|&?A�����!��̣�8�d0�9�b/����m�\.qp��H�ϐ �}�v�$�7D4�~��0蝽�<F�����	�b96��2�'F8C4��D�,[�%W!81/b/��v��6A�9�ERe}a����R��²�w I������9X�»X�o&A����u�����Jĉ@�V%i �u�؂���D\0n��B4��Y|�B�01�|���*�	��U��߯ǧh���^��Ø�����f	�3�.;R�$!\U�AR�1���'$���y����<��_lV��0U Y��"`�ݚ�<�Pz�k�0 vq� �ә��%ɭ�c$sru��k����?C��;�=��Kt_�k�QB�(�_"�@|�z��{z���?�u��m����z���.������.T+/wQ�<�YX��s��n	_x�{I]X�H�1}���=�n:�{����#��\���s4B�L_��nk�`xV4��TreX8������N��	�������taj:�%�Բ�(>h��:�K1.[|���l�E��Ƒd;�].an\X�J,���w��k~"�����Ur�ꀵ�W+1��p��Kp9�f�懘��j��B�TO���p�Y�	��~�uJQ�Gg�����6S���ȓI�B]��P���P;���1�T.	vKK"ͯ�Wi���KC��8����KT��N�i*
�J* aaU���u圀��P.�#L+��)?@`��dJ�o!)�վ�f
�fo��o�bq�0�?kv96���n��zo���?�l\j�ڗ�"����{�����f)��O؃���;�.�R�Z�R����|��	V5�����C�;&��OH�}A>�Q$�����Y���ԛ��vڜ�kS^H��q�5J~�2Xf��->qi�W(P��c��NuWEm*���CH�s�;�Hz�ok�[�W�@㷐���M�� /@	0�B�=M�Ou��G_�qbL�eor���������X�}xH��k`�>=�(e����R����O��T��=υ���2v����[�i>���o��vls�ܷO*9�* ��?U�G1L�լ�Q@�{W���c�B��T߂h�~�ȉ%���8A�jXK�̍��P���ϸ�Nj9��~���3��LNk�vʴxT���'����c*��Ԓ�ȚT��Ȫ}�ٜj��Uq	s�3̯!��rj�L���PQ���Խ�F��rK��3rUr�j��]g��s�ly|���$�q���P�_�g=����|*�Y���@D������� ���y3��U�x�N�&|
r�n�T#��f:X��6�Z+J�(��u�D�(p�I=�q<� �����S]SU�������9��*E�x�P]2�ڥX�w���)�5Ӣ,���uJ���'���ɶuƤ�8h���D4QD#�/Y�,��w�Y�a�R��P�hY5�<�R���,���K^Ґ]8�������*&���&�h
�RCh������\���icUq�Ǎ$U_*�B��X�l B�ʞe[�h�K���F��5,�ؙo�:��E�= �q��7��N�3����[\eǰڨ��p�g��9m��źm����?�;��z\_���)�'͎l+�(��3������a��K(?ۭ�:�GPJ��?�ˠxF����A�)�5�hs���#\^R�x$=�r�-�W١j��h�\�SZ�b�";W�]�<Mv�Cꌘ��[5;L���=t3�o�]~�l���M+8JN-��ء�+|(Kt��;��A=�b����Gl쿀�b]���-����I}�z�B}����<zg/�l�����cޕOf'$������ꣷ�62�ǍY!k[��H;��u9JT���q(dU���h.�K僚�9�ڡ���M�Z��d4�nc��.P/װ�j��̢[U���K�����k@��p칺LҺ���o�V��"j�6���鱄�i�ǉ�H�e�� Ѷ�5ۣd?��G�v�J.��'#��n���
�Hv��Z!ʧ�-CUd!��ʵڪ��<ӭ,vG��m�a��<���A_5��	��>�M���|�S	�S�=-ضIB�A��Z��x�m�7�'���-�1>�t%�ۛ]�LۻT�c��&�h:��M#����o7��@'�@��6�n/������?��Orf��vM��k�U�7�|l��-���4��P����dVͻ�e9pɫ��a��99�ץ2>K����/����'�96L���$\�>|u>ٚ��Th+ٓ���l%����C�W��<�<�U�����HF�P�>��tD.;����~���'�  �� UzU