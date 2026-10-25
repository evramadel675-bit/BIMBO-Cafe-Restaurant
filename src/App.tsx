/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  LayoutGrid,
  ShoppingBag,
  UtensilsCrossed,
  BarChart3,
  Settings,
  Moon,
  Sun,
  Download,
  Upload,
  Plus,
  Edit2,
  Trash2,
  Check,
  X,
  ArrowRight,
  ArrowLeft,
  ChevronRight,
  ChevronLeft,
  Printer,
  Copy,
  FileSpreadsheet,
  Calendar,
  Lock,
  Unlock,
  KeyRound,
  ShieldCheck,
  AlertTriangle,
  Search
} from 'lucide-react';

/* ---------- Interfaces & Types ---------- */
export interface MenuItem {
  id: string;
  name: string;
  category: string;
  price: number;
  available: boolean;
}

export interface OrderLine {
  id: string;
  name: string;
  price: number;
  qty: number;
  note: string;
}

export interface TableItem {
  id: string;
  name: string;
  order: {
    openedAt: number;
    lines: OrderLine[];
  } | null;
}

export interface PaymentMethodOption {
  id: string;
  name: string;
  active: boolean;
}

export interface AppSettings {
  businessName: string;
  phone: string;
  address: string;
  receiptHeader: string;
  receiptFooter: string;
  currency: string;
  taxRate: number;
  taxEnabled: boolean;
  serviceRate: number;
  serviceEnabled: boolean;
  paymentMethods: PaymentMethodOption[];
}

export interface SaleRecord {
  id: string;
  no: number;
  ts: number;
  type: 'table' | 'take';
  table: string;
  lines: OrderLine[];
  sub: number;
  disc: number;
  svc: number;
  tax: number;
  total: number;
  net: number;
  pay: string;
  recv?: number;
  imported?: boolean;
}

export type ClosingPeriodType = 'day' | 'week' | 'month';

interface ConfirmModalState {
  isOpen: boolean;
  title: string;
  message: string;
  confirmLabel: string;
  danger?: boolean;
  onConfirm: () => void;
}

const STORAGE_KEY = 'bimbo-pos-v4';

/* ---------- Safe Value Sanitizers (Fixing Table Rendering & Sort Bug) ---------- */
export const safePrice = (p: any): number => {
  return parseFloat(String(p ?? 0).replace(/[^0-9.]/g, '')) || 0;
};

export const safeStr = (s: any): string => {
  return String(s ?? '').trim();
};

/* ---------- Clean Slate Initial Data (Zero Dummy Items / Records) ---------- */
function getCleanInitialData() {
  const items: MenuItem[] = []; // Zero dummy items

  const tables: TableItem[] = [
    { id: 'tb-1', name: 'طاولة 1', order: null },
    { id: 'tb-2', name: 'طاولة 2', order: null },
    { id: 'tb-3', name: 'طاولة 3', order: null },
    { id: 'tb-4', name: 'طاولة 4', order: null },
    { id: 'tb-5', name: 'طاولة 5', order: null },
    { id: 'tb-6', name: 'بار 1', order: null },
    { id: 'tb-7', name: 'بار 2', order: null }
  ];

  const settings: AppSettings = {
    businessName: 'كرفان بيمبو - BIMBO POS',
    phone: '',
    address: '',
    receiptHeader: 'أهلاً بكم! نسعد بخدمتكم دائماً',
    receiptFooter: 'شكراً لزيارتكم! نرجو أن تكونوا قد استمتعتم بتجربتكم',
    currency: 'ج.م',
    taxRate: 14,
    taxEnabled: false,
    serviceRate: 12,
    serviceEnabled: false,
    paymentMethods: [
      { id: 'cash', name: 'نقدي (كاش)', active: true },
      { id: 'card', name: 'بطاقة ائتمان / فيزا', active: true },
      { id: 'wallet', name: 'محفظة إلكترونية (انستاباي / فودافون كاش)', active: true }
    ]
  };

  const sales: SaleRecord[] = []; // Zero dummy sales
  const counter = 1001;

  return { items, tables, settings, sales, counter };
}

/* ---------- Reusable Horizontal Scroll Bar with Left/Right Navigation Buttons ---------- */
function ScrollableBar({
  children,
  className = ''
}: {
  children: React.ReactNode;
  className?: string;
}) {
  const scrollRef = useRef<HTMLDivElement>(null);

  const handleScroll = (delta: number) => {
    if (!scrollRef.current) return;
    scrollRef.current.scrollBy({ left: delta, behavior: 'smooth' });
  };

  return (
    <div className={`relative flex items-center gap-1.5 w-full max-w-full overflow-hidden ${className}`}>
      {/* Right Scroll Arrow (Touch Optimized) */}
      <button
        type="button"
        onClick={() => handleScroll(200)}
        className="shrink-0 w-10 h-10 min-w-[40px] min-h-[40px] rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-[#1a1816] dark:text-[#f3efe9] flex items-center justify-center hover:bg-[#eae7e0] dark:hover:bg-[#282521] shadow-2xs active:scale-95 transition z-10"
        title="تمرير لليمين"
        aria-label="تمرير لليمين"
      >
        <ChevronRight className="w-5 h-5" />
      </button>

      {/* Scrollable Container (Swiping & Dragging Preserved) */}
      <div
        ref={scrollRef}
        className="flex gap-2 overflow-x-auto pb-1 pt-0.5 scrollbar-none scroll-smooth flex-1 items-center -webkit-overflow-scrolling-touch"
        style={{ scrollbarWidth: 'none', msOverflowStyle: 'none' }}
      >
        {children}
      </div>

      {/* Left Scroll Arrow (Touch Optimized) */}
      <button
        type="button"
        onClick={() => handleScroll(-200)}
        className="shrink-0 w-10 h-10 min-w-[40px] min-h-[40px] rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-[#1a1816] dark:text-[#f3efe9] flex items-center justify-center hover:bg-[#eae7e0] dark:hover:bg-[#282521] shadow-2xs active:scale-95 transition z-10"
        title="تمرير لليسار"
        aria-label="تمرير لليسار"
      >
        <ChevronLeft className="w-5 h-5" />
      </button>
    </div>
  );
}

/* ---------- SheetJS Excel Parser (Guaranteed 89 vs 89 Items Import) ---------- */
function parseMenuFromWorkbook(wb: XLSX.WorkBook): MenuItem[] {
  const result: MenuItem[] = [];

  // Standard flexible header keywords
  const nameHeaders = ['الاسم', 'اسم الصنف', 'الصنف', 'اسم المنتج', 'المنتج', 'name', 'item', 'product', 'title', 'dish'];
  const priceHeaders = ['السعر', 'سعر البيع', 'price', 'cost', 'ثمن', 'قيمة', 'مبلغ', 'rate', 'amount'];
  const catHeaders = ['الفئة', 'القسم', 'التصنيف', 'category', 'cat', 'مجموعة', 'نوع', 'group', 'section'];

  // Check all sheets in workbook
  for (let sIdx = 0; sIdx < wb.SheetNames.length; sIdx++) {
    const sheetName = wb.SheetNames[sIdx];
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) continue;

    // Use { header: 1, raw: false, defval: '' } as required
    const rawRows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    if (!rawRows || rawRows.length === 0) continue;

    let headerRowIdx = -1;
    let nameCol = -1;
    let catCol = -1;
    let priceCol = -1;

    // Scan first 15 rows for flexible headers
    for (let r = 0; r < Math.min(15, rawRows.length); r++) {
      const row = rawRows[r];
      if (!Array.isArray(row)) continue;
      const lower = row.map((cell) => String(cell || '').trim().toLowerCase());

      const nIdx = lower.findIndex((c) => nameHeaders.some((h) => c === h || c.includes(h)));
      const pIdx = lower.findIndex((c) => priceHeaders.some((h) => c === h || c.includes(h)));
      const cIdx = lower.findIndex((c) => catHeaders.some((h) => c === h || c.includes(h)));

      if (nIdx > -1 && pIdx > -1) {
        headerRowIdx = r;
        nameCol = nIdx;
        priceCol = pIdx;
        catCol = cIdx;
        break;
      }
    }

    const startRow = headerRowIdx > -1 ? headerRowIdx + 1 : 0;

    // Auto-detect columns if headers were not named
    if (nameCol === -1 || priceCol === -1) {
      let bestPrice = -1;
      const textCols: number[] = [];

      for (let c = 0; c < 10; c++) {
        let numCount = 0;
        let textCount = 0;
        for (let r = startRow; r < Math.min(startRow + 25, rawRows.length); r++) {
          const val = rawRows[r]?.[c];
          if (val === undefined || val === '') continue;
          const clean = String(val).replace(/[^0-9.]/g, '');
          if (clean && !isNaN(Number(clean)) && Number(clean) > 0) numCount++;
          else if (String(val).trim().length > 1) textCount++;
        }
        if (numCount >= 2 && bestPrice === -1) bestPrice = c;
        if (textCount >= 2 && c !== bestPrice) textCols.push(c);
      }
      if (priceCol === -1 && bestPrice > -1) priceCol = bestPrice;
      if (nameCol === -1 && textCols.length > 0) {
        nameCol = textCols[0];
        if (catCol === -1 && textCols.length > 1) catCol = textCols[1];
      }
    }

    if (nameCol === -1) nameCol = 0;
    if (priceCol === -1) priceCol = 1;

    // Parse every row: NEVER drop unless product name itself is completely blank
    for (let r = startRow; r < rawRows.length; r++) {
      const row = rawRows[r];
      if (!row || !Array.isArray(row) || row.length === 0) continue;

      const rawName = String(row[nameCol] ?? '').trim();
      // Only skip if blank or a repetitive header line
      if (!rawName || rawName === '0' || rawName === '#' || nameHeaders.includes(rawName.toLowerCase())) continue;

      // Category: If missing or empty, assign "عام" instead of skipping!
      let rawCat = catCol > -1 && row[catCol] !== undefined ? String(row[catCol] ?? '').trim() : '';
      if (!rawCat) {
        rawCat = sheetName && !sheetName.toLowerCase().startsWith('sheet') ? sheetName : 'عام';
      }

      // Sanitize and safely parse price values using parseFloat(String(rawPrice || 0).replace(/[^0-9.]/g, '')) || 0
      const rawPrice = row[priceCol];
      const price = parseFloat(String(rawPrice ?? 0).replace(/[^0-9.]/g, '')) || 0;

      result.push({
        id: `item-${Date.now()}-${sIdx}-${r}-${Math.random().toString(36).slice(2, 7)}`,
        name: rawName,
        category: rawCat || 'عام',
        price: Math.max(0, price),
        available: true
      });
    }
  }

  return result;
}

export default function App() {
  // Theme state with rock-solid localStorage sync
  const [isDarkMode, setIsDarkMode] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('bimbo-theme');
      if (saved) return saved === 'dark';
      return window.matchMedia('(prefers-color-scheme: dark)').matches;
    } catch {
      return false;
    }
  });

  // State initialization with clean slate
  const [items, setItems] = useState<MenuItem[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.items && Array.isArray(parsed.items)) {
          return parsed.items.map((i: any) => ({
            id: i.id || `item-${Math.random().toString(36).slice(2, 9)}`,
            name: safeStr(i.name),
            category: safeStr(i.category || i.cat || 'عام') || 'عام',
            price: safePrice(i.price),
            available: i.available !== false
          }));
        }
      }
    } catch {}
    return getCleanInitialData().items;
  });

  const [tables, setTables] = useState<TableItem[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.tables && Array.isArray(parsed.tables)) return parsed.tables;
      }
    } catch {}
    return getCleanInitialData().tables;
  });

  const [takeawayOrder, setTakeawayOrder] = useState<{ lines: OrderLine[] }>({ lines: [] });

  const [settings, setSettings] = useState<AppSettings>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.settings) {
          return {
            businessName: parsed.settings.businessName || 'كرفان بيمبو - BIMBO POS',
            phone: parsed.settings.phone || '',
            address: parsed.settings.address || '',
            receiptHeader: parsed.settings.receiptHeader || 'أهلاً بكم! نسعد بخدمتكم دائماً',
            receiptFooter: parsed.settings.receiptFooter || 'شكراً لزيارتكم!',
            currency: parsed.settings.currency || 'ج.م',
            taxRate: Number(parsed.settings.taxRate) || 14,
            taxEnabled: Boolean(parsed.settings.taxEnabled),
            serviceRate: Number(parsed.settings.serviceRate) || 12,
            serviceEnabled: Boolean(parsed.settings.serviceEnabled),
            paymentMethods: parsed.settings.paymentMethods || [
              { id: 'cash', name: 'نقدي (كاش)', active: true },
              { id: 'card', name: 'بطاقة ائتمان / فيزا', active: true },
              { id: 'wallet', name: 'محفظة إلكترونية (انستاباي / فودافون كاش)', active: true }
            ]
          };
        }
      }
    } catch {}
    return getCleanInitialData().settings;
  });

  const [sales, setSales] = useState<SaleRecord[]>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.sales && Array.isArray(parsed.sales)) return parsed.sales;
      }
    } catch {}
    return getCleanInitialData().sales;
  });

  const [invoiceCounter, setInvoiceCounter] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.counter) return parsed.counter;
      }
    } catch {}
    return 1001;
  });

  // UI Navigation & View State
  const [activeTab, setActiveTab] = useState<'tables' | 'take' | 'menu' | 'reports' | 'settings'>('tables');
  const [activeTableId, setActiveTableId] = useState<string | null>(null);
  const [orderSubTab, setOrderSubTab] = useState<'menu' | 'cart'>('menu');
  const [isMobileCartOpen, setIsMobileCartOpen] = useState<boolean>(false);

  // Filters: Always default to 'الكل' so no items are hidden by mistake
  const [orderCategory, setOrderCategory] = useState<string>('الكل');
  const [orderSearch, setOrderSearch] = useState<string>('');

  const [menuCategory, setMenuCategory] = useState<string>('الكل');
  const [menuStatusFilter, setMenuStatusFilter] = useState<'all' | 'avail' | 'unavail'>('all');
  const [menuSearch, setMenuSearch] = useState<string>('');
  const [menuSort, setMenuSort] = useState<'name' | 'cat' | 'price-asc' | 'price-desc'>('name');

  const [isEditingTables, setIsEditingTables] = useState<boolean>(false);
  const [reportPeriod, setReportPeriod] = useState<'today' | 'week' | 'month' | 'all'>('today');

  // Financial Period Closing State
  const [closingType, setClosingType] = useState<ClosingPeriodType>('day');
  const [closingDayDate, setClosingDayDate] = useState<string>(() => new Date().toISOString().slice(0, 10));
  const [closingMonthDate, setClosingMonthDate] = useState<string>(() => new Date().toISOString().slice(0, 7));

  // Modals & Feedback
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const toastTimeoutRef = useRef<number | null>(null);

  const [isCheckoutOpen, setIsCheckoutOpen] = useState<boolean>(false);
  const [discountType, setDiscountType] = useState<'amt' | 'pct'>('amt');
  const [discountVal, setDiscountVal] = useState<number>(0);
  const [selectedPayMethod, setSelectedPayMethod] = useState<string>('cash');
  const [receivedCashAmount, setReceivedCashAmount] = useState<string>('');

  const [itemModalData, setItemModalData] = useState<{
    isOpen: boolean;
    item: MenuItem | null;
  }>({ isOpen: false, item: null });

  // In-App Confirmation Modal
  const [confirmModal, setConfirmModal] = useState<ConfirmModalState>({
    isOpen: false,
    title: '',
    message: '',
    confirmLabel: 'تأكيد',
    danger: false,
    onConfirm: () => {}
  });

  const [lastFinishedReceipt, setLastFinishedReceipt] = useState<string | null>(null);

  // Owner Authentication & Cashier Safe Mode
  const OWNER_PIN_STORAGE_KEY = 'bimbo_owner_pin';

  const [isOwnerAuthenticated, setIsOwnerAuthenticated] = useState<boolean>(false);
  const [ownerPin, setOwnerPin] = useState<string>(() => {
    try {
      return localStorage.getItem(OWNER_PIN_STORAGE_KEY) || '';
    } catch {
      return '';
    }
  });

  // Owner PIN Modal State
  const [isPinModalOpen, setIsPinModalOpen] = useState<boolean>(false);
  const [pinPromptTitle, setPinPromptTitle] = useState<string>('أدخل كلمة مرور المالك للمتابعة');
  const [pinPromptDescription, setPinPromptDescription] = useState<string>('هذا القسم محمي ويتطلب صلاحيات المالك');
  const [enteredPin, setEnteredPin] = useState<string>('');
  const [pinError, setPinError] = useState<string>('');
  const [pinShake, setPinShake] = useState<boolean>(false);
  const [pendingAction, setPendingAction] = useState<(() => void) | null>(null);

  // First-Time Setup State
  const [setupStep, setSetupStep] = useState<1 | 2>(1); // 1: New PIN, 2: Confirm PIN
  const [setupNewPin, setSetupNewPin] = useState<string>('');
  const [setupConfirmPin, setSetupConfirmPin] = useState<string>('');

  // Settings PIN Change State
  const [oldPinInput, setOldPinInput] = useState<string>('');
  const [newPinInput, setNewPinInput] = useState<string>('');
  const [confirmNewPinInput, setConfirmNewPinInput] = useState<string>('');
  const [pinChangeError, setPinChangeError] = useState<string>('');
  const [pinChangeSuccess, setPinChangeSuccess] = useState<string>('');

  const isSetupMode = !ownerPin;

  const requireOwnerAuth = (action: () => void, title = 'أدخل كلمة مرور المالك للمتابعة', desc = 'هذا القسم محمي ويتطلب صلاحيات المالك') => {
    if (isOwnerAuthenticated) {
      action();
      return;
    }
    setPendingAction(() => action);
    setPinError('');
    setPinShake(false);
    setEnteredPin('');
    setSetupStep(1);
    setSetupNewPin('');
    setSetupConfirmPin('');

    if (!ownerPin) {
      setPinPromptTitle('تعيين رمز مرور المالك لأول مرة');
      setPinPromptDescription('لحماية أرباحك وإعدادات المنيو، يرجى تعيين رمز مرور مكون من 4 أرقام');
    } else {
      setPinPromptTitle(title);
      setPinPromptDescription(desc);
    }
    setIsPinModalOpen(true);
  };

  const handleVerifyPin = (pinToTest?: string) => {
    const pin = pinToTest !== undefined ? pinToTest : enteredPin;
    if (pin === ownerPin) {
      setIsOwnerAuthenticated(true);
      setIsPinModalOpen(false);
      setEnteredPin('');
      setPinError('');
      showToast('تم التحقق بنجاح وتفعيل وضع المالك');
      if (pendingAction) {
        const act = pendingAction;
        setPendingAction(null);
        act();
      }
    } else {
      setPinError('رمز المرور غير صحيح! يرجى المحاولة مرة أخرى');
      setPinShake(true);
      setTimeout(() => setPinShake(false), 600);
      setEnteredPin(''); // Auto-clear dots on error
    }
  };

  // First-Time Setup Step 1 (Save Step 1 PIN and proceed to Step 2)
  const handleSetupStep1 = (pin: string) => {
    if (pin.length !== 4) {
      setPinError('يرجى إدخال رمز مكون من 4 أرقام');
      return;
    }
    setSetupNewPin(pin);
    setSetupStep(2);
    setSetupConfirmPin('');
    setPinError('');
  };

  // First-Time Setup Step 2 (Verify confirmation and save to localStorage)
  const handleSetupStep2 = (confirmPin: string) => {
    if (confirmPin !== setupNewPin) {
      setPinError('تأكيد الرمز غير متطابق! يرجى إعادة إدخال الرمز');
      setPinShake(true);
      setTimeout(() => setPinShake(false), 600);
      setSetupConfirmPin(''); // Auto-clear dots on error
      return;
    }

    try {
      localStorage.setItem(OWNER_PIN_STORAGE_KEY, setupNewPin);
      setOwnerPin(setupNewPin);
      setIsOwnerAuthenticated(true);
      setIsPinModalOpen(false);
      setSetupStep(1);
      setSetupNewPin('');
      setSetupConfirmPin('');
      setPinError('');
      showToast('تم تعيين رمز مرور المالك وتفعيل الجلسة بنجاح!');
      if (pendingAction) {
        const act = pendingAction;
        setPendingAction(null);
        act();
      }
    } catch {
      setPinError('حدث خطأ أثناء حفظ الرمز الجديد في المتصفح');
    }
  };

  const handleOwnerLock = () => {
    setIsOwnerAuthenticated(false);
    showToast('تم قفل وضع المالك والعودة لوضع الكاشير الآمن');
    if (activeTab === 'reports' || activeTab === 'settings') {
      setActiveTab('tables');
    }
  };

  const handleChangeOwnerPin = (e: React.FormEvent) => {
    e.preventDefault();
    setPinChangeError('');
    setPinChangeSuccess('');

    if (!oldPinInput || oldPinInput !== ownerPin) {
      setPinChangeError('الرمز الحالي غير صحيح');
      return;
    }
    if (!newPinInput || newPinInput.length < 4) {
      setPinChangeError('يجب أن يتكون الرمز الجديد من 4 أرقام على الأقل');
      return;
    }
    if (newPinInput !== confirmNewPinInput) {
      setPinChangeError('تأكيد الرمز غير متطابق مع الرمز الجديد');
      return;
    }

    try {
      localStorage.setItem(OWNER_PIN_STORAGE_KEY, newPinInput);
      setOwnerPin(newPinInput);
      setOldPinInput('');
      setNewPinInput('');
      setConfirmNewPinInput('');
      setPinChangeSuccess('تم تحديث رمز مرور المالك بنجاح!');
      showToast('تم حفظ رمز مرور المالك الجديد');
    } catch {
      setPinChangeError('حدث خطأ أثناء حفظ الرمز الجديد');
    }
  };

  // Sync Theme to HTML Root & Body
  useEffect(() => {
    document.documentElement.dataset.theme = isDarkMode ? 'dark' : 'light';
    if (isDarkMode) {
      document.documentElement.classList.add('dark');
      document.body.classList.add('dark');
    } else {
      document.documentElement.classList.remove('dark');
      document.body.classList.remove('dark');
    }
    localStorage.setItem('bimbo-theme', isDarkMode ? 'dark' : 'light');
  }, [isDarkMode]);

  // Persist State to LocalStorage
  useEffect(() => {
    try {
      const dataToSave = {
        v: 4,
        items,
        tables,
        settings,
        sales,
        counter: invoiceCounter
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToSave));
    } catch {}
  }, [items, tables, settings, sales, invoiceCounter]);

  const showToast = (msg: string) => {
    setToastMessage(msg);
    if (toastTimeoutRef.current) clearTimeout(toastTimeoutRef.current);
    toastTimeoutRef.current = window.setTimeout(() => {
      setToastMessage(null);
    }, 2800);
  };

  /* ---------- Calculations & Formatting ---------- */
  const formatMoney = (val: number) => {
    return (
      new Intl.NumberFormat('en-US', { maximumFractionDigits: 2 }).format(val || 0) +
      ' ' +
      (settings.currency || 'ج.م')
    );
  };

  const activeTable = useMemo(() => {
    return tables.find((t) => t.id === activeTableId) || null;
  }, [tables, activeTableId]);

  const currentOrderLines = useMemo(() => {
    if (activeTab === 'take') return takeawayOrder.lines;
    if (activeTab === 'tables' && activeTable) return activeTable.order?.lines || [];
    return [];
  }, [activeTab, activeTable, takeawayOrder]);

  const calculateOrderTotals = (lines: OrderLine[], discT = discountType, discV = discountVal) => {
    const sub = Math.round(lines.reduce((acc, l) => acc + safePrice(l.price) * l.qty, 0) * 100) / 100;
    let disc = 0;
    if (discV > 0) {
      disc = discT === 'pct' ? (sub * Math.min(discV, 100)) / 100 : Math.min(discV, sub);
    }
    disc = Math.round(disc * 100) / 100;
    const base = sub - disc;
    const svc = settings.serviceEnabled ? Math.round(((base * (settings.serviceRate || 0)) / 100) * 100) / 100 : 0;
    const tax = settings.taxEnabled ? Math.round(((base * (settings.taxRate || 0)) / 100) * 100) / 100 : 0;
    const total = Math.round((base + svc + tax) * 100) / 100;
    return { sub, disc, net: base, svc, tax, total };
  };

  const currentTotals = useMemo(() => {
    return calculateOrderTotals(currentOrderLines, 'amt', 0);
  }, [currentOrderLines, settings]);

  /* ---------- Cart & Order Modifications ---------- */
  const addToCart = (item: MenuItem) => {
    if (!item.available) {
      showToast('هذا الصنف غير متوفر حالياً');
      return;
    }
    const updateLines = (prevLines: OrderLine[]) => {
      const existing = prevLines.find((l) => l.id === item.id);
      if (existing) {
        return prevLines.map((l) => (l.id === item.id ? { ...l, qty: l.qty + 1 } : l));
      }
      return [...prevLines, { id: item.id, name: item.name, price: safePrice(item.price), qty: 1, note: '' }];
    };

    if (activeTab === 'take') {
      setTakeawayOrder((prev) => ({ lines: updateLines(prev.lines) }));
    } else if (activeTab === 'tables' && activeTableId) {
      setTables((prev) =>
        prev.map((t) => {
          if (t.id !== activeTableId) return t;
          const currentLines = t.order?.lines || [];
          return {
            ...t,
            order: {
              openedAt: t.order?.openedAt || Date.now(),
              lines: updateLines(currentLines)
            }
          };
        })
      );
    }
  };

  const updateCartItemQty = (itemId: string, delta: number) => {
    const updateLines = (prevLines: OrderLine[]) => {
      return prevLines
        .map((l) => {
          if (l.id === itemId) {
            const nextQty = l.qty + delta;
            return nextQty > 0 ? { ...l, qty: nextQty } : null;
          }
          return l;
        })
        .filter(Boolean) as OrderLine[];
    };

    if (activeTab === 'take') {
      setTakeawayOrder((prev) => ({ lines: updateLines(prev.lines) }));
    } else if (activeTab === 'tables' && activeTableId) {
      setTables((prev) =>
        prev.map((t) => {
          if (t.id !== activeTableId) return t;
          return {
            ...t,
            order: t.order ? { ...t.order, lines: updateLines(t.order.lines) } : null
          };
        })
      );
    }
  };

  const removeCartLine = (itemId: string) => {
    const filterLines = (prevLines: OrderLine[]) => prevLines.filter((l) => l.id !== itemId);

    if (activeTab === 'take') {
      setTakeawayOrder((prev) => ({ lines: filterLines(prev.lines) }));
    } else if (activeTab === 'tables' && activeTableId) {
      setTables((prev) =>
        prev.map((t) => {
          if (t.id !== activeTableId) return t;
          return {
            ...t,
            order: t.order ? { ...t.order, lines: filterLines(t.order.lines) } : null
          };
        })
      );
    }
    showToast('تم حذف الصنف من الفاتورة');
  };

  const updateCartItemNote = (itemId: string, note: string) => {
    const updateLines = (prevLines: OrderLine[]) => {
      return prevLines.map((l) => (l.id === itemId ? { ...l, note } : l));
    };

    if (activeTab === 'take') {
      setTakeawayOrder((prev) => ({ lines: updateLines(prev.lines) }));
    } else if (activeTab === 'tables' && activeTableId) {
      setTables((prev) =>
        prev.map((t) => {
          if (t.id !== activeTableId) return t;
          return {
            ...t,
            order: t.order ? { ...t.order, lines: updateLines(t.order.lines) } : null
          };
        })
      );
    }
  };

  const handleClearCart = () => {
    if (!currentOrderLines.length) return;
    setConfirmModal({
      isOpen: true,
      title: 'إفراغ الفاتورة الحالية',
      message: 'هل تريد حذف جميع الأصناف المضافة للفاتورة الحالية والبدء من جديد؟',
      confirmLabel: 'نعم، إفراغ السلة',
      danger: true,
      onConfirm: () => {
        if (activeTab === 'take') {
          setTakeawayOrder({ lines: [] });
        } else if (activeTab === 'tables' && activeTableId) {
          setTables((prev) =>
            prev.map((t) => (t.id === activeTableId ? { ...t, order: null } : t))
          );
        }
        showToast('تم إفراغ الفاتورة بالكامل');
        setConfirmModal((m) => ({ ...m, isOpen: false }));
      }
    });
  };

  /* ---------- Checkout & Invoice Finishing ---------- */
  const handleConfirmCheckout = () => {
    if (!currentOrderLines.length) return;
    const totals = calculateOrderTotals(currentOrderLines, discountType, discountVal);
    const saleId = `sale-${invoiceCounter}`;
    const newSale: SaleRecord = {
      id: saleId,
      no: invoiceCounter,
      ts: Date.now(),
      type: activeTab === 'take' ? 'take' : 'table',
      table: activeTab === 'take' ? 'تيك أواي' : activeTable?.name || 'طاولة',
      lines: [...currentOrderLines],
      sub: totals.sub,
      disc: totals.disc,
      svc: totals.svc,
      tax: totals.tax,
      total: totals.total,
      net: totals.net,
      pay: selectedPayMethod,
      recv: Number(receivedCashAmount) || undefined
    };

    setSales((prev) => [...prev, newSale]);
    setInvoiceCounter((c) => c + 1);

    if (activeTab === 'take') {
      setTakeawayOrder({ lines: [] });
    } else if (activeTab === 'tables' && activeTableId) {
      setTables((prev) =>
        prev.map((t) => (t.id === activeTableId ? { ...t, order: null } : t))
      );
      setActiveTableId(null);
    }

    setIsCheckoutOpen(false);
    setDiscountVal(0);
    setReceivedCashAmount('');

    const textReceipt = generateReceiptText(newSale);
    setLastFinishedReceipt(textReceipt);
    showToast(`تم دفع الفاتورة رقم #${newSale.no} بنجاح`);
  };

  const generateReceiptText = (sale: SaleRecord) => {
    const divider = '--------------------------------';
    const pad = (a: string, b: string) => a + ' '.repeat(Math.max(1, 32 - a.length - b.length)) + b;
    const payMethodName =
      settings.paymentMethods.find((m) => m.id === sale.pay)?.name ||
      (sale.pay === 'cash' ? 'نقدي' : sale.pay === 'card' ? 'بطاقة' : 'محفظة إلكترونية');

    const dt = new Date(sale.ts);
    const dateStr = dt.toLocaleDateString('ar-EG', { day: 'numeric', month: 'short', year: 'numeric' });
    const timeStr = dt.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' });

    const lines = [
      settings.businessName,
      settings.phone ? `هاتف: ${settings.phone}` : '',
      settings.address ? settings.address : '',
      divider,
      settings.receiptHeader ? settings.receiptHeader : '',
      divider,
      `فاتورة رقم: #${sale.no}`,
      `التاريخ: ${dateStr} ${timeStr}`,
      `نوع الطلب: ${sale.type === 'table' ? sale.table : 'تيك أواي / استلام'}`,
      divider,
      ...sale.lines.map((l) => pad(`${l.qty}x ${l.name}`, formatMoney(safePrice(l.price) * l.qty))),
      divider,
      pad('الإجمالي الفرعي:', formatMoney(sale.sub)),
      ...(sale.disc > 0 ? [pad('الخصم:', `-${formatMoney(sale.disc)}`)] : []),
      ...(sale.svc > 0 ? [pad(`الخدمة (${settings.serviceRate}%):`, formatMoney(sale.svc))] : []),
      ...(sale.tax > 0 ? [pad(`الضريبة (${settings.taxRate}%):`, formatMoney(sale.tax))] : []),
      divider,
      pad('المطلوب النهائي:', formatMoney(sale.total)),
      pad('طريقة الدفع:', payMethodName),
      divider,
      settings.receiptFooter ? settings.receiptFooter : 'شكراً لزيارتكم!'
    ];

    return lines.filter(Boolean).join('\n');
  };

  /* ---------- Menu Item CRUD & Delete Handlers ---------- */
  const handleDeleteMenuItem = (item: MenuItem) => {
    setConfirmModal({
      isOpen: true,
      title: 'حذف صنف من المنيو',
      message: `هل أنت متأكد من حذف «${item.name}» من قائمة المنيو نهائياً؟ لن يؤثر ذلك على الفواتير السابقة.`,
      confirmLabel: 'تأكيد الحذف',
      danger: true,
      onConfirm: () => {
        setItems((prev) => prev.filter((i) => i.id !== item.id));
        showToast(`تم حذف «${item.name}» من المنيو بنجاح`);
        setConfirmModal((m) => ({ ...m, isOpen: false }));
      }
    });
  };

  const handleSaveItemModal = (formData: FormData) => {
    const name = safeStr(formData.get('name'));
    const category = safeStr(formData.get('category')) || 'عام';
    const price = safePrice(formData.get('price'));
    const available = formData.get('available') === 'true';

    if (!name) {
      showToast('يرجى كتابة اسم الصنف');
      return;
    }

    if (itemModalData.item) {
      setItems((prev) =>
        prev.map((i) =>
          i.id === itemModalData.item!.id ? { ...i, name, category, price, available } : i
        )
      );
      showToast(`تم تعديل «${name}» بنجاح`);
    } else {
      const newItem: MenuItem = {
        id: `item-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
        name,
        category,
        price,
        available
      };
      setItems((prev) => [newItem, ...prev]);
      showToast(`تمت إضافة «${name}» إلى المنيو`);
    }

    setItemModalData({ isOpen: false, item: null });
  };

  /* ---------- Excel Menu Import & Export (Guaranteed 89 Items) ---------- */
  const exportMenuToExcel = () => {
    if (!items.length) {
      showToast('لا توجد أصناف في المنيو لتصديرها');
      return;
    }
    const rows = [
      ['اسم الصنف', 'الفئة / القسم', 'سعر البيع', 'حالة التوفر'],
      ...items.map((i) => [i.name, i.category || 'عام', safePrice(i.price), i.available ? 'متاح' : 'غير متوفر'])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Menu');
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `BIMBO_Menu_${today}.xlsx`);
    showToast('تم تصدير المنيو إلى ملف Excel بنجاح');
  };

  const handleMenuFileImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const importedItems = parseMenuFromWorkbook(wb);

        if (!importedItems || importedItems.length === 0) {
          showToast('لم يتم العثور على أصناف صالحة في الملف، تأكد من وجود أعمدة الاسم والسعر');
          return;
        }

        // Add all imported items into the menu without dropping any!
        setItems((prevItems) => {
          if (!prevItems.length) return importedItems;
          // Append new items
          return [...prevItems, ...importedItems];
        });

        // Reset categories filter to 'الكل' so all 89 items are immediately visible
        setMenuCategory('الكل');
        setOrderCategory('الكل');
        showToast(`تم استيراد ${importedItems.length} صنف بنجاح من ملف Excel!`);
      } catch {
        showToast('حدث خطأ أثناء قراءة ملف Excel: يرجى التحقق من الملف');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  /* ---------- Historical Orders Delete Handler ---------- */
  const handleDeleteSale = (sale: SaleRecord) => {
    setConfirmModal({
      isOpen: true,
      title: 'إلغاء وحذف الفاتورة',
      message: `هل أنت متأكد من حذف الفاتورة رقم #${sale.no} بقيمة ${formatMoney(sale.total)}؟ سيتم مسحها نهائياً من التقارير والسجلات.`,
      confirmLabel: 'تأكيد الحذف',
      danger: true,
      onConfirm: () => {
        setSales((prev) => prev.filter((s) => s.id !== sale.id));
        showToast(`تم حذف الفاتورة رقم #${sale.no}`);
        setConfirmModal((m) => ({ ...m, isOpen: false }));
      }
    });
  };

  /* ---------- Financial Period Closing (Filter, Detailed Stats, Selective Export & Purge) ---------- */
  const isSaleInClosingPeriod = (ts: number, type: ClosingPeriodType, dayVal: string, monthVal: string): boolean => {
    const d = new Date(ts);
    if (type === 'day') {
      const saleDay = d.toISOString().slice(0, 10);
      return saleDay === dayVal;
    }
    if (type === 'month') {
      const saleMonth = d.toISOString().slice(0, 7);
      return saleMonth === monthVal;
    }
    if (type === 'week') {
      const end = new Date(dayVal + 'T23:59:59').getTime();
      const start = end - 7 * 86400 * 1000;
      return ts >= start && ts <= end;
    }
    return false;
  };

  const closingSales = useMemo(() => {
    return sales.filter((s) => isSaleInClosingPeriod(s.ts, closingType, closingDayDate, closingMonthDate));
  }, [sales, closingType, closingDayDate, closingMonthDate]);

  const closingStats = useMemo(() => {
    const count = closingSales.length;
    const subTotal = closingSales.reduce((a, s) => a + (s.sub || 0), 0);
    const discTotal = closingSales.reduce((a, s) => a + (s.disc || 0), 0);
    const svcTotal = closingSales.reduce((a, s) => a + (s.svc || 0), 0);
    const taxTotal = closingSales.reduce((a, s) => a + (s.tax || 0), 0);
    const grossTotal = closingSales.reduce((a, s) => a + (s.total || 0), 0);
    const netTotal = closingSales.reduce((a, s) => a + (s.net || s.total || 0), 0);

    const cashTotal = closingSales.filter((s) => s.pay === 'cash').reduce((a, s) => a + (s.total || 0), 0);
    const cardTotal = closingSales.filter((s) => s.pay === 'card').reduce((a, s) => a + (s.total || 0), 0);
    const walletTotal = closingSales.filter((s) => s.pay === 'wallet').reduce((a, s) => a + (s.total || 0), 0);
    const otherTotal = closingSales.filter((s) => s.pay !== 'cash' && s.pay !== 'card' && s.pay !== 'wallet').reduce((a, s) => a + (s.total || 0), 0);

    const itemMap: Record<string, { name: string; qty: number; rev: number }> = {};
    closingSales.forEach((s) => {
      s.lines.forEach((l) => {
        const name = l.name || 'غير معروف';
        if (!itemMap[name]) itemMap[name] = { name, qty: 0, rev: 0 };
        itemMap[name].qty += l.qty || 1;
        itemMap[name].rev += (safePrice(l.price)) * (l.qty || 1);
      });
    });

    const itemsSummary = Object.values(itemMap).sort((a, b) => b.rev - a.rev);

    return {
      count,
      subTotal: Math.round(subTotal * 100) / 100,
      discTotal: Math.round(discTotal * 100) / 100,
      svcTotal: Math.round(svcTotal * 100) / 100,
      taxTotal: Math.round(taxTotal * 100) / 100,
      grossTotal: Math.round(grossTotal * 100) / 100,
      netTotal: Math.round(netTotal * 100) / 100,
      cashTotal: Math.round(cashTotal * 100) / 100,
      cardTotal: Math.round(cardTotal * 100) / 100,
      walletTotal: Math.round(walletTotal * 100) / 100,
      otherTotal: Math.round(otherTotal * 100) / 100,
      itemsSummary
    };
  }, [closingSales]);

  const handleExportClosingExcel = () => {
    if (!closingSales.length) {
      showToast('لا توجد فواتير في هذه الفترة لتصديرها');
      return;
    }

    const periodLabel =
      closingType === 'day'
        ? `يوم ${closingDayDate}`
        : closingType === 'week'
        ? `الأسبوع المنتهي في ${closingDayDate}`
        : `شهر ${closingMonthDate}`;

    const summaryRows = [
      ['تقرير تسوية وإغلاق الحسابات المالية', settings.businessName],
      ['نوع الفترة', closingType === 'day' ? 'يومي' : closingType === 'week' ? 'أسبوعي' : 'شهري'],
      ['تاريخ الفترة المحددة', periodLabel],
      ['تاريخ إنشاء التقرير', new Date().toLocaleString('ar-EG')],
      ['', ''],
      ['المؤشر المالي', 'القيمة (' + settings.currency + ')'],
      ['إجمالي عدد الفواتير', closingStats.count],
      ['إجمالي المبيعات الفرعية (قبل الخصم)', closingStats.subTotal],
      ['إجمالي الخصومات الممنوحة', closingStats.discTotal],
      ['إجمالي رسوم الخدمة', closingStats.svcTotal],
      ['إجمالي الضرائب', closingStats.taxTotal],
      ['صافي الإيرادات', closingStats.netTotal],
      ['الإجمالي النهائي المطلوب والمحصل', closingStats.grossTotal],
      ['', ''],
      ['توزيع طرق الدفع', 'المبلغ المحصل'],
      ['المدفوع نقداً (كاش)', closingStats.cashTotal],
      ['المدفوع ببطاقات الائتمان / فيزا', closingStats.cardTotal],
      ['المدفوع بالمحافظ الإلكترونية (انستاباي / فودافون كاش)', closingStats.walletTotal],
      ['طرق دفع أخرى', closingStats.otherTotal]
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);

    const salesRows = [
      [
        'رقم الفاتورة',
        'التاريخ',
        'الوقت',
        'نوع الطلب',
        'الموقع / الطاولة',
        'الأصناف المباعة',
        'الإجمالي الفرعي',
        'الخصم',
        'الخدمة',
        'الضريبة',
        'الإجمالي النهائي',
        'طريقة الدفع'
      ],
      ...closingSales.map((s) => {
        const dt = new Date(s.ts);
        const dateStr = dt.toISOString().slice(0, 10);
        const timeStr = dt.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' });
        const itemsStr = s.lines.map((l) => `${l.qty}x ${l.name}`).join(' | ');
        const payName =
          settings.paymentMethods.find((m) => m.id === s.pay)?.name ||
          (s.pay === 'cash' ? 'نقدي' : s.pay === 'card' ? 'بطاقة' : 'محفظة');
        return [
          s.no,
          dateStr,
          timeStr,
          s.type === 'table' ? 'طاولة' : 'تيك أواي',
          s.table,
          itemsStr,
          s.sub,
          s.disc,
          s.svc,
          s.tax,
          s.total,
          payName
        ];
      })
    ];
    const wsSales = XLSX.utils.aoa_to_sheet(salesRows);

    const itemsRows = [
      ['اسم الصنف', 'الكمية المباعة في الفترة', 'إجمالي الإيرادات (' + settings.currency + ')'],
      ...closingStats.itemsSummary.map((it) => [it.name, it.qty, it.rev])
    ];
    const wsItems = XLSX.utils.aoa_to_sheet(itemsRows);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsSummary, 'ملخص الإغلاق');
    XLSX.utils.book_append_sheet(wb, wsSales, 'فواتير الفترة');
    XLSX.utils.book_append_sheet(wb, wsItems, 'مبيعات الأصناف');

    const cleanDateTag = (closingType === 'month' ? closingMonthDate : closingDayDate).replace(/-/g, '');
    XLSX.writeFile(wb, `BIMBO_Closing_${closingType}_${cleanDateTag}.xlsx`);
    showToast('تم تصدير تقرير إغلاق الفترة بنجاح إلى ملف Excel!');
  };

  const handlePurgeClosingRecords = () => {
    if (!closingSales.length) {
      showToast('لا توجد حركات في هذه الفترة لمسحها');
      return;
    }

    const periodLabel =
      closingType === 'day'
        ? `يوم ${closingDayDate}`
        : closingType === 'week'
        ? `الأسبوع المنتهي في ${closingDayDate}`
        : `شهر ${closingMonthDate}`;

    setConfirmModal({
      isOpen: true,
      title: 'مسح الحركات المصدرة من النظام (إغلاق الدفاتر)',
      message: `تحذير هام: سيتم مسح ${closingSales.length} فاتورة تابعة لفترة (${periodLabel}) نهائياً من قاعدة البيانات لتسوية الحسابات. تأكد من تحميل تقرير الإكسيل أولاً. هل تريد المتابعة؟`,
      confirmLabel: 'تأكيد مسح حركات الفترة',
      danger: true,
      onConfirm: () => {
        const purgedCount = closingSales.length;
        setSales((prev) =>
          prev.filter((s) => !isSaleInClosingPeriod(s.ts, closingType, closingDayDate, closingMonthDate))
        );
        showToast(`تم إغلاق الفترة بنجاح ومسح ${purgedCount} فاتورة من السجل!`);
        setConfirmModal((m) => ({ ...m, isOpen: false }));
      }
    });
  };

  /* ---------- Excel Sales History Export & Import ---------- */
  const exportAllSalesToExcel = () => {
    if (!sales.length) {
      showToast('لا توجد مبيعات في السجل لتصديرها');
      return;
    }
    const rows = [
      [
        'رقم الفاتورة',
        'التاريخ',
        'الوقت',
        'النوع',
        'الموقع / الطاولة',
        'الأصناف المطلوبة',
        'الإجمالي الفرعي',
        'الخصم',
        'الخدمة',
        'الضريبة',
        'الإجمالي النهائي',
        'طريقة الدفع'
      ],
      ...sales.map((s) => {
        const dt = new Date(s.ts);
        const dateStr = dt.toISOString().slice(0, 10);
        const timeStr = dt.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' });
        const itemsStr = s.lines.map((l) => `${l.qty}x ${l.name}`).join(' | ');
        const payName =
          settings.paymentMethods.find((m) => m.id === s.pay)?.name ||
          (s.pay === 'cash' ? 'نقدي' : s.pay === 'card' ? 'بطاقة' : 'محفظة');
        return [
          s.no,
          dateStr,
          timeStr,
          s.type === 'table' ? 'طاولة' : 'تيك أواي',
          s.table,
          itemsStr,
          s.sub,
          s.disc,
          s.svc,
          s.tax,
          s.total,
          payName
        ];
      })
    ];

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sales History');
    const today = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `BIMBO_All_Sales_${today}.xlsx`);
    showToast('تم تصدير سجل المبيعات الكامل إلى Excel');
  };

  const handleSalesFileImport = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target?.result as ArrayBuffer);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
        if (!rows || rows.length < 2) {
          showToast('الملف لا يحتوي على سجلات فواتير كافية');
          return;
        }

        const headers = rows[0].map((h: any) => String(h || '').trim().toLowerCase());
        const noIdx = headers.findIndex((h) => h.includes('رقم') || h.includes('no') || h.includes('id'));
        const dateIdx = headers.findIndex((h) => h.includes('تاريخ') || h.includes('date') || h.includes('وقت'));
        const totalIdx = headers.findIndex((h) => h.includes('إجمالي') || h.includes('نهائي') || h.includes('total') || h.includes('مبلغ'));
        const subIdx = headers.findIndex((h) => h.includes('فرعي') || h.includes('sub'));
        const itemsIdx = headers.findIndex((h) => h.includes('صنف') || h.includes('أصناف') || h.includes('items'));
        const payIdx = headers.findIndex((h) => h.includes('دفع') || h.includes('طريقة') || h.includes('pay'));
        const typeIdx = headers.findIndex((h) => h.includes('نوع') || h.includes('طاولة') || h.includes('type'));

        let importedCount = 0;
        const newRecords: SaleRecord[] = [];

        for (let i = 1; i < rows.length; i++) {
          const r = rows[i];
          if (!r || !r.length) continue;
          const total = safePrice(r[totalIdx > -1 ? totalIdx : 1]);
          if (total <= 0) continue;

          let ts = Date.now();
          const dRaw = r[dateIdx > -1 ? dateIdx : 0];
          if (typeof dRaw === 'number') {
            ts = new Date((dRaw - 25569) * 86400 * 1000).getTime();
          } else if (dRaw) {
            const parsed = Date.parse(String(dRaw));
            if (!isNaN(parsed)) ts = parsed;
          }

          const no = parseInt(r[noIdx > -1 ? noIdx : -1]) || invoiceCounter + i;
          const sub = safePrice(r[subIdx > -1 ? subIdx : -1]) || total;
          const payRaw = String(r[payIdx > -1 ? payIdx : -1] || 'نقدي').toLowerCase();
          const pay =
            payRaw.includes('بطاق') || payRaw.includes('فيزا') || payRaw.includes('card')
              ? 'card'
              : payRaw.includes('محفظ') || payRaw.includes('كاش') || payRaw.includes('wallet')
              ? 'wallet'
              : 'cash';

          const typeRaw = String(r[typeIdx > -1 ? typeIdx : -1] || '');
          const isTable = typeRaw.includes('طاول') || typeRaw.includes('table');

          let lines: OrderLine[] = [];
          const itemsRaw = String(r[itemsIdx > -1 ? itemsIdx : -1] || '');
          if (itemsRaw) {
            lines = itemsRaw
              .split(/[|,]/)
              .map((part) => {
                const p = part.trim();
                if (!p) return null;
                const match = p.match(/(\d+)\s*[xX*×]\s*(.+)/);
                if (match) {
                  return {
                    id: Math.random().toString(36).slice(2, 9),
                    name: match[2].trim(),
                    price: Math.round((total / Number(match[1])) * 100) / 100,
                    qty: parseInt(match[1]) || 1,
                    note: ''
                  };
                }
                return {
                  id: Math.random().toString(36).slice(2, 9),
                  name: p,
                  price: total,
                  qty: 1,
                  note: ''
                };
              })
              .filter(Boolean) as OrderLine[];
          }

          if (!lines.length) {
            lines = [{ id: Math.random().toString(36).slice(2, 9), name: 'طلب مستورد', price: total, qty: 1, note: '' }];
          }

          newRecords.push({
            id: `import-${Date.now()}-${i}`,
            no,
            ts,
            type: isTable ? 'table' : 'take',
            table: isTable ? typeRaw || 'طاولة مستوردة' : 'تيك أواي',
            lines,
            sub,
            disc: 0,
            svc: 0,
            tax: 0,
            total,
            net: sub,
            pay,
            imported: true
          });
          importedCount++;
        }

        if (newRecords.length > 0) {
          setSales((prev) => {
            const merged = [...prev, ...newRecords];
            merged.sort((a, b) => a.ts - b.ts);
            return merged;
          });
          showToast(`تم استيراد ودمج ${importedCount} فاتورة تاريخية بنجاح!`);
        }
      } catch {
        showToast('حدث خطأ أثناء استيراد ملف المبيعات Excel');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  /* ---------- General Overview Reports Filters ---------- */
  const filteredSales = useMemo(() => {
    const now = new Date();
    return sales.filter((s) => {
      if (reportPeriod === 'all') return true;
      const dt = new Date(s.ts);
      if (reportPeriod === 'today') {
        return (
          dt.getDate() === now.getDate() &&
          dt.getMonth() === now.getMonth() &&
          dt.getFullYear() === now.getFullYear()
        );
      }
      if (reportPeriod === 'week') {
        const weekAgo = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 7).getTime();
        return s.ts >= weekAgo;
      }
      if (reportPeriod === 'month') {
        return dt.getMonth() === now.getMonth() && dt.getFullYear() === now.getFullYear();
      }
      return true;
    });
  }, [sales, reportPeriod]);

  const salesReportStats = useMemo(() => {
    const totalRev = filteredSales.reduce((acc, s) => acc + (s.net || s.total || 0), 0);
    const count = filteredSales.length;
    const avgTicket = count ? totalRev / count : 0;

    const itemMap: Record<string, { name: string; qty: number; rev: number }> = {};
    filteredSales.forEach((s) => {
      s.lines.forEach((l) => {
        const name = l.name || 'صنف';
        if (!itemMap[name]) itemMap[name] = { name, qty: 0, rev: 0 };
        itemMap[name].qty += l.qty || 1;
        itemMap[name].rev += safePrice(l.price) * (l.qty || 1);
      });
    });

    const topItems = Object.values(itemMap).sort((a, b) => b.qty - a.qty);

    return {
      totalRev: Math.round(totalRev * 100) / 100,
      count,
      avgTicket: Math.round(avgTicket * 100) / 100,
      topItems
    };
  }, [filteredSales]);

  // Unique categories list with 'الكل'
  const categoriesList = useMemo(() => {
    const cats = [...new Set(items.map((i) => safeStr(i.category || 'عام')))];
    return ['الكل', ...cats.filter(Boolean)];
  }, [items]);

  // Order Screen Filtered Items
  const filteredOrderItems = useMemo(() => {
    const q = orderSearch.trim().toLowerCase();
    return items.filter((i) => {
      const matchCat = orderCategory === 'الكل' || safeStr(i.category || 'عام') === orderCategory;
      const matchQuery = !q || safeStr(i.name).toLowerCase().includes(q) || safeStr(i.category).toLowerCase().includes(q);
      return matchCat && matchQuery;
    });
  }, [items, orderCategory, orderSearch]);

  // Menu Management Filtered & Safely Sorted Items (Fixes 1 item displayed out of 80)
  const menuTableItems = useMemo(() => {
    const q = menuSearch.trim().toLowerCase();

    // 1. Filter by category, status, and search query
    let list = items.filter((i) => {
      const cat = safeStr(i.category || 'عام');
      const matchesCat = menuCategory === 'الكل' || cat === menuCategory;
      const matchesStatus =
        menuStatusFilter === 'all'
          ? true
          : menuStatusFilter === 'avail'
          ? i.available !== false
          : i.available === false;
      const matchesQuery = !q || safeStr(i.name).toLowerCase().includes(q) || cat.toLowerCase().includes(q);
      return matchesCat && matchesStatus && matchesQuery;
    });

    // 2. Ultra-safe sort handler (never throws on undefined or dirty price strings)
    return [...list].sort((a, b) => {
      if (menuSort === 'name') {
        return safeStr(a.name).localeCompare(safeStr(b.name), 'ar');
      }
      if (menuSort === 'cat') {
        return safeStr(a.category || 'عام').localeCompare(safeStr(b.category || 'عام'), 'ar');
      }
      if (menuSort === 'price-asc') {
        return safePrice(a.price) - safePrice(b.price);
      }
      if (menuSort === 'price-desc') {
        return safePrice(b.price) - safePrice(a.price);
      }
      return 0;
    });
  }, [items, menuCategory, menuStatusFilter, menuSearch, menuSort]);

  return (
    <div className="min-h-screen flex flex-col bg-[#f5f4f0] dark:bg-[#141311] text-[#1a1816] dark:text-[#f3efe9] transition-colors duration-200">
      {/* Top Header */}
      <header className="sticky top-0 z-40 bg-white/95 dark:bg-[#1e1c19]/95 backdrop-blur border-b border-[#dedad2] dark:border-[#38342e] shadow-xs">
        <div className="max-w-7xl mx-auto px-4 py-2.5 flex items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <span className="font-display text-2xl bg-[#c2650a] dark:bg-[#ea7a1a] text-white px-3 py-1 rounded-xl shadow-sm tracking-wide">
              BIMBO
            </span>
            <div>
              <h1 className="font-bold text-base md:text-lg leading-tight line-clamp-1">{settings.businessName}</h1>
              <p className="text-xs text-[#6e685f] dark:text-[#9e978b] flex items-center gap-2">
                {settings.phone && <span>{settings.phone}</span>}
                {settings.address && <span>· {settings.address}</span>}
                {!settings.phone && !settings.address && <span>نظام كاشير ونقاط بيع الكرفان</span>}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {/* Owner Session Indicator & Lock/Unlock Toggle */}
            {isOwnerAuthenticated ? (
              <button
                type="button"
                onClick={handleOwnerLock}
                className="px-2.5 sm:px-3 py-1.5 rounded-xl border border-emerald-300 dark:border-emerald-700/60 bg-emerald-50 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-200 hover:bg-emerald-100 dark:hover:bg-emerald-900/50 transition flex items-center gap-1.5 text-xs md:text-sm font-bold shadow-xs active:scale-95"
                title="وضع المالك نشط · اضغط لقفل الجلسة والعودة لوضع الكاشير"
              >
                <Unlock className="w-4 h-4 text-emerald-600 dark:text-emerald-400" />
                <span className="hidden sm:inline">المالك متصل (قفل 🔒)</span>
                <span className="sm:hidden">قفل</span>
              </button>
            ) : (
              <button
                type="button"
                onClick={() => {
                  requireOwnerAuth(
                    () => showToast('مرحباً بك! تم تفعيل صلاحيات المالك'),
                    ownerPin ? 'تسجيل دخول المالك' : 'تعيين رمز مرور المالك لأول مرة',
                    ownerPin ? 'أدخل رمز PIN لتفعيل صلاحيات المالك والوصول للتقارير والإعدادات' : 'لم يتم تعيين رمز مرور حتى الآن. يرجى إنشاء رمزك الخاص'
                  );
                }}
                className="px-2.5 sm:px-3 py-1.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-[#6e685f] dark:text-[#9e978b] hover:text-[#1a1816] dark:hover:text-[#f3efe9] hover:bg-[#eae7e0] dark:hover:bg-[#282521] transition flex items-center gap-1.5 text-xs md:text-sm font-bold shadow-xs active:scale-95"
                title={ownerPin ? 'تسجيل دخول المالك' : 'تعيين رمز المالك لأول مرة'}
              >
                <Lock className="w-4 h-4 text-[#c2650a]" />
                <span className="hidden sm:inline">{ownerPin ? 'دخول المالك' : 'تعيين رمز المالك'}</span>
                <span className="sm:hidden">{ownerPin ? 'دخول' : 'تعيين'}</span>
              </button>
            )}

            {/* Rock-Solid Theme Toggle */}
            <button
              type="button"
              onClick={() => setIsDarkMode((prev) => !prev)}
              className="px-3 py-1.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-[#1a1816] dark:text-[#f3efe9] hover:bg-[#eae7e0] dark:hover:bg-[#282521] transition flex items-center gap-2 text-xs md:text-sm font-bold shadow-xs"
              title={isDarkMode ? 'التحويل للوضع النهاري' : 'التحويل للوضع الليلي'}
              aria-label="تبديل الوضع الليلي والنهاري"
            >
              {isDarkMode ? (
                <>
                  <Sun className="w-4 h-4 text-amber-400 fill-amber-400" />
                  <span>وضع نهاري</span>
                </>
              ) : (
                <>
                  <Moon className="w-4 h-4 text-slate-700" />
                  <span>وضع ليلي</span>
                </>
              )}
            </button>
          </div>
        </div>
      </header>

      {/* Main Content Layout */}
      <div className="max-w-7xl mx-auto w-full px-4 py-4 md:py-6 flex-1 grid grid-cols-1 md:grid-cols-[96px_minmax(0,1fr)] gap-5">
        {/* Navigation Rail */}
        <nav
          className="fixed md:sticky bottom-0 md:top-20 inset-x-0 md:inset-x-auto z-30 md:self-start bg-white dark:bg-[#1e1c19] md:bg-transparent md:dark:bg-transparent border-t md:border-t-0 border-[#dedad2] dark:border-[#38342e] p-1.5 md:p-0 flex md:flex-col justify-around md:justify-start gap-1 md:gap-2 shadow-lg md:shadow-none"
          aria-label="التنقل الرئيسي"
        >
          <button
            type="button"
            onClick={() => {
              setActiveTab('tables');
              setActiveTableId(null);
            }}
            className={`flex flex-col items-center gap-1 p-2 md:p-3 rounded-xl text-xs md:text-sm font-bold transition flex-1 md:flex-none ${
              activeTab === 'tables'
                ? 'bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] border border-[#c2650a]/20 shadow-xs'
                : 'text-[#6e685f] dark:text-[#9e978b] hover:bg-[#eae7e0] dark:hover:bg-[#282521]'
            }`}
          >
            <LayoutGrid className="w-5 h-5 md:w-6 md:h-6" />
            <span>الطاولات</span>
          </button>

          <button
            type="button"
            onClick={() => {
              setActiveTab('take');
              setActiveTableId(null);
            }}
            className={`flex flex-col items-center gap-1 p-2 md:p-3 rounded-xl text-xs md:text-sm font-bold transition flex-1 md:flex-none ${
              activeTab === 'take'
                ? 'bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] border border-[#c2650a]/20 shadow-xs'
                : 'text-[#6e685f] dark:text-[#9e978b] hover:bg-[#eae7e0] dark:hover:bg-[#282521]'
            }`}
          >
            <ShoppingBag className="w-5 h-5 md:w-6 md:h-6" />
            <span>تيك أواي</span>
          </button>

          <button
            type="button"
            onClick={() => {
              setActiveTab('menu');
              // Ensure filter defaults to 'الكل' so all items appear
              setMenuCategory('الكل');
            }}
            className={`flex flex-col items-center gap-1 p-2 md:p-3 rounded-xl text-xs md:text-sm font-bold transition flex-1 md:flex-none ${
              activeTab === 'menu'
                ? 'bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] border border-[#c2650a]/20 shadow-xs'
                : 'text-[#6e685f] dark:text-[#9e978b] hover:bg-[#eae7e0] dark:hover:bg-[#282521]'
            }`}
          >
            <UtensilsCrossed className="w-5 h-5 md:w-6 md:h-6" />
            <span>المنيو</span>
          </button>

          <button
            type="button"
            onClick={() => {
              requireOwnerAuth(
                () => setActiveTab('reports'),
                'الدخول إلى التقارير المالية والأرباح',
                'تقارير المبيعات وحركات الفواتير وإغلاق الفترات محمية برمز المالك'
              );
            }}
            className={`flex flex-col items-center gap-1 p-2 md:p-3 rounded-xl text-xs md:text-sm font-bold transition flex-1 md:flex-none ${
              activeTab === 'reports'
                ? 'bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] border border-[#c2650a]/20 shadow-xs'
                : 'text-[#6e685f] dark:text-[#9e978b] hover:bg-[#eae7e0] dark:hover:bg-[#282521]'
            }`}
          >
            <div className="relative">
              <BarChart3 className="w-5 h-5 md:w-6 md:h-6" />
              {!isOwnerAuthenticated && (
                <Lock className="w-2.5 h-2.5 absolute -top-1 -right-1.5 text-amber-600 dark:text-amber-400" />
              )}
            </div>
            <span>التقارير</span>
          </button>

          <button
            type="button"
            onClick={() => {
              requireOwnerAuth(
                () => setActiveTab('settings'),
                'الدخول إلى إعدادات النظام',
                'تعديل إعدادات الكرفان والضرائب وطرق الدفع والرمز السري محمي برمز المالك'
              );
            }}
            className={`flex flex-col items-center gap-1 p-2 md:p-3 rounded-xl text-xs md:text-sm font-bold transition flex-1 md:flex-none ${
              activeTab === 'settings'
                ? 'bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] border border-[#c2650a]/20 shadow-xs'
                : 'text-[#6e685f] dark:text-[#9e978b] hover:bg-[#eae7e0] dark:hover:bg-[#282521]'
            }`}
          >
            <div className="relative">
              <Settings className="w-5 h-5 md:w-6 md:h-6" />
              {!isOwnerAuthenticated && (
                <Lock className="w-2.5 h-2.5 absolute -top-1 -right-1.5 text-amber-600 dark:text-amber-400" />
              )}
            </div>
            <span>الإعدادات</span>
          </button>
        </nav>

        {/* View Switcher */}
        <main className="min-w-0 pb-20 md:pb-6">
          {/* 1. TABLES VIEW */}
          {activeTab === 'tables' && !activeTableId && (
            <div>
              <div className="flex items-center justify-between gap-4 flex-wrap mb-5">
                <div>
                  <h2 className="text-xl md:text-2xl font-bold">طاولات وجلسات الكرفان</h2>
                  <p className="text-xs md:text-sm text-[#6e685f] dark:text-[#9e978b]">
                    {tables.filter((t) => t.order && t.order.lines.length > 0).length} طاولات مشغولة من إجمالي{' '}
                    {tables.length}
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  {isEditingTables && (
                    <button
                      type="button"
                      onClick={() => {
                        const newName = `طاولة ${tables.length + 1}`;
                        setTables((prev) => [...prev, { id: `tb-${Date.now()}`, name: newName, order: null }]);
                      }}
                      className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] font-semibold text-sm hover:border-[#c2650a]"
                    >
                      <Plus className="w-4 h-4" />
                      طاولة جديدة
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => setIsEditingTables(!isEditingTables)}
                    className={`px-4 py-2 rounded-xl font-bold text-sm transition ${
                      isEditingTables
                        ? 'bg-[#c2650a] text-white shadow-sm'
                        : 'bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] text-[#1a1816] dark:text-[#f3efe9]'
                    }`}
                  >
                    {isEditingTables ? 'حفظ والانتهاء' : 'ترتيب وتعديل'}
                  </button>
                </div>
              </div>

              {isEditingTables && (
                <div className="mb-4 p-3 rounded-xl bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-800 text-xs md:text-sm text-amber-800 dark:text-amber-200">
                  يمكنك استخدام الأسهم لتغيير الترتيب، تعديل الاسم، أو حذف الطاولات الفارغة.
                </div>
              )}

              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-3 md:gap-4">
                {tables.map((t, idx) => {
                  const hasOrder = Boolean(t.order && t.order.lines.length > 0);
                  const orderSubtotal = t.order?.lines.reduce((a, l) => a + safePrice(l.price) * l.qty, 0) || 0;
                  const itemsCount = t.order?.lines.reduce((a, l) => a + l.qty, 0) || 0;

                  if (isEditingTables) {
                    return (
                      <div
                        key={t.id}
                        className="p-4 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] flex flex-col gap-3"
                      >
                        <div className="flex items-center justify-between">
                          <input
                            type="text"
                            value={t.name}
                            onChange={(e) => {
                              const v = e.target.value;
                              setTables((prev) => prev.map((tb) => (tb.id === t.id ? { ...tb, name: v } : tb)));
                            }}
                            className="font-bold text-base bg-transparent border-b border-dashed border-[#dedad2] focus:border-[#c2650a] outline-none w-3/4"
                          />
                          <button
                            type="button"
                            disabled={hasOrder}
                            onClick={() => {
                              setConfirmModal({
                                isOpen: true,
                                title: 'حذف الطاولة',
                                message: `هل تريد حذف «${t.name}» نهائياً؟`,
                                confirmLabel: 'تأكيد الحذف',
                                danger: true,
                                onConfirm: () => {
                                  setTables((prev) => prev.filter((tb) => tb.id !== t.id));
                                  setConfirmModal((m) => ({ ...m, isOpen: false }));
                                  showToast('تم حذف الطاولة');
                                }
                              });
                            }}
                            className="text-red-500 hover:text-red-700 disabled:opacity-30 p-1 rounded-lg"
                            title="حذف الطاولة"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>

                        <div className="flex items-center justify-between mt-auto pt-2 border-t border-[#dedad2] dark:border-[#38342e]">
                          <span className="text-xs text-[#6e685f] dark:text-[#9e978b]">
                            {hasOrder ? 'مشغولة' : 'فارغة'}
                          </span>
                          <div className="flex items-center gap-1">
                            <button
                              type="button"
                              disabled={idx === 0}
                              onClick={() => {
                                setTables((prev) => {
                                  const copy = [...prev];
                                  [copy[idx - 1], copy[idx]] = [copy[idx], copy[idx - 1]];
                                  return copy;
                                });
                              }}
                              className="p-1 rounded bg-[#eae7e0] dark:bg-[#282521] disabled:opacity-30"
                              title="للأمام"
                            >
                              <ArrowRight className="w-3.5 h-3.5" />
                            </button>
                            <button
                              type="button"
                              disabled={idx === tables.length - 1}
                              onClick={() => {
                                setTables((prev) => {
                                  const copy = [...prev];
                                  [copy[idx + 1], copy[idx]] = [copy[idx], copy[idx + 1]];
                                  return copy;
                                });
                              }}
                              className="p-1 rounded bg-[#eae7e0] dark:bg-[#282521] disabled:opacity-30"
                              title="للخلف"
                            >
                              <ArrowLeft className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  }

                  return (
                    <button
                      type="button"
                      key={t.id}
                      onClick={() => {
                        setActiveTableId(t.id);
                        setOrderSubTab('menu');
                      }}
                      className={`p-4 rounded-2xl border text-start flex flex-col justify-between min-h-[120px] transition active:scale-[0.98] ${
                        hasOrder
                          ? 'bg-[#fbebd8] dark:bg-[#38240f]/70 border-[#c2650a]/40 shadow-xs'
                          : 'bg-white dark:bg-[#1e1c19] border-[#dedad2] dark:border-[#38342e] hover:border-[#c2650a]'
                      }`}
                    >
                      <div className="flex items-center justify-between w-full">
                        <span className="font-bold text-lg">{t.name}</span>
                        <span
                          className={`text-xs px-2 py-0.5 rounded-full font-bold ${
                            hasOrder
                              ? 'bg-[#c2650a] text-white'
                              : 'bg-[#eae7e0] dark:bg-[#282521] text-[#6e685f] dark:text-[#9e978b]'
                          }`}
                        >
                          {hasOrder ? 'مشغولة' : 'فارغة'}
                        </span>
                      </div>

                      {hasOrder ? (
                        <div className="mt-4 flex items-end justify-between w-full">
                          <span className="text-xs text-[#6e685f] dark:text-[#9e978b]">
                            {itemsCount} صنف
                          </span>
                          <span className="font-extrabold text-base md:text-lg text-[#c2650a] dark:text-[#ea7a1a]">
                            {formatMoney(orderSubtotal)}
                          </span>
                        </div>
                      ) : (
                        <span className="text-xs text-[#6e685f] dark:text-[#9e978b] mt-4">
                          اضغط لفتح طلب جديد
                        </span>
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          )}

          {/* 2. ORDER DETAILS (Table Order or Takeaway) */}
          {(activeTab === 'take' || (activeTab === 'tables' && activeTableId)) && (
            <div>
              {/* Order Header */}
              <div className="flex items-center justify-between gap-4 mb-4">
                <div className="flex items-center gap-3">
                  {activeTab === 'tables' && (
                    <button
                      type="button"
                      onClick={() => setActiveTableId(null)}
                      className="px-3 py-1.5 rounded-xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] font-bold text-sm flex items-center gap-1 hover:border-[#c2650a]"
                    >
                      <ArrowRight className="w-4 h-4" />
                      الطاولات
                    </button>
                  )}
                  <div>
                    <h2 className="text-xl font-bold">
                      {activeTab === 'take' ? 'طلب كاشير تيك أواي / فوري' : activeTable?.name}
                    </h2>
                    <p className="text-xs text-[#6e685f] dark:text-[#9e978b]">
                      {activeTab === 'take'
                        ? 'طلب دفع سريع بدون طاولة'
                        : `فُتحت في ${new Date(activeTable?.order?.openedAt || Date.now()).toLocaleTimeString('ar-EG', {
                            hour: '2-digit',
                            minute: '2-digit'
                          })}`}
                    </p>
                  </div>
                </div>

                <div className="lg:hidden flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setIsMobileCartOpen(true)}
                    className="min-h-[44px] px-3.5 py-2 rounded-xl text-xs sm:text-sm font-bold bg-[#c2650a] text-white flex items-center gap-1.5 shadow-sm active:scale-95"
                  >
                    <ShoppingBag className="w-4 h-4" />
                    <span>السلة ({currentOrderLines.reduce((a, l) => a + l.qty, 0)})</span>
                  </button>
                </div>
              </div>

              {/* Order Split Grid: Menu on right/center, Desktop Sticky Cart on side */}
              <div className="grid grid-cols-1 lg:grid-cols-[minmax(0,1fr)_380px] gap-6 items-start">
                {/* Menu Browser (Always Full-Width & Uncrowded on Mobile) */}
                <div className="block">
                  {/* Search and Categories with Horizontal Scroll Arrows */}
                  <div className="flex flex-col gap-3 mb-4">
                    <div className="relative">
                      <Search className="w-4 h-4 absolute right-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input
                        type="search"
                        placeholder="ابحث عن صنف بالاسم أو القسم..."
                        value={orderSearch}
                        onChange={(e) => setOrderSearch(e.target.value)}
                        className="w-full pr-10 pl-4 py-2.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-sm outline-none focus:border-[#c2650a] min-h-[44px]"
                      />
                    </div>

                    {/* Horizontal Scroll Bar for Order Categories with < and > buttons */}
                    <ScrollableBar>
                      {categoriesList.map((cat) => (
                        <button
                          type="button"
                          key={cat}
                          onClick={() => setOrderCategory(cat)}
                          className={`min-h-[42px] px-4 py-2 rounded-full text-xs md:text-sm font-bold whitespace-nowrap transition flex items-center justify-center shrink-0 active:scale-95 ${
                            orderCategory === cat
                              ? 'bg-[#1a1816] text-white dark:bg-[#f3efe9] dark:text-[#141311] shadow-xs'
                              : 'bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] text-[#1a1816] dark:text-[#f3efe9] hover:border-gray-400'
                          }`}
                        >
                          {cat}
                        </button>
                      ))}
                    </ScrollableBar>
                  </div>

                  {/* Items Grid (2 cols on small mobile, 3 to 4 on tablets, 4 to 5 on large screens) */}
                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-3 xl:grid-cols-4 2xl:grid-cols-5 gap-2.5 sm:gap-3.5">
                    {filteredOrderItems.map((item) => (
                      <button
                        type="button"
                        key={item.id}
                        disabled={!item.available}
                        onClick={() => addToCart(item)}
                        className={`p-3 sm:p-3.5 rounded-2xl border text-start flex flex-col justify-between min-h-[115px] sm:min-h-[125px] transition active:scale-[0.97] ${
                          item.available
                            ? 'bg-white dark:bg-[#1e1c19] border-[#dedad2] dark:border-[#38342e] hover:border-[#c2650a] shadow-2xs'
                            : 'bg-[#eae7e0] dark:bg-[#282521] border-transparent opacity-60 cursor-not-allowed'
                        }`}
                      >
                        <span className="font-bold text-sm md:text-base leading-snug line-clamp-2">
                          {item.name}
                        </span>
                        <div className="mt-2.5 flex items-center justify-between w-full">
                          <span className="font-extrabold text-sm md:text-base text-[#c2650a] dark:text-[#ea7a1a]">
                            {formatMoney(safePrice(item.price))}
                          </span>
                          {!item.available && (
                            <span className="text-[10px] font-bold text-red-500">غير متوفر</span>
                          )}
                        </div>
                      </button>
                    ))}
                    {!filteredOrderItems.length && (
                      <div className="col-span-full py-16 text-center text-[#6e685f] dark:text-[#9e978b] text-sm bg-white dark:bg-[#1e1c19] rounded-2xl border border-dashed border-[#dedad2] dark:border-[#38342e] p-8">
                        <UtensilsCrossed className="w-10 h-10 mx-auto text-gray-400 mb-2 opacity-50" />
                        <p className="font-bold text-base mb-1">لا توجد أصناف مطابقة للبحث أو القسم المحدد</p>
                        <p className="text-xs">اختر «الكل» لعرض كافة عناصر القائمة أو انتقل للتبويب لإضافة أصناف</p>
                      </div>
                    )}
                  </div>
                </div>

                {/* Desktop Sticky Cart Aside (lg:flex) */}
                <div className="hidden lg:flex bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] rounded-2xl p-4 shadow-sm flex-col max-h-[calc(100vh-120px)] lg:sticky lg:top-20">
                  <div className="flex items-center justify-between pb-3 border-b border-[#dedad2] dark:border-[#38342e]">
                    <h3 className="font-bold text-base flex items-center gap-2">
                      <ShoppingBag className="w-4 h-4 text-[#c2650a] dark:text-[#ea7a1a]" />
                      <span>محتويات الفاتورة</span>
                    </h3>
                    <span className="text-xs px-2.5 py-1 rounded-full bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] font-bold">
                      {currentOrderLines.reduce((a, l) => a + l.qty, 0)} عناصر
                    </span>
                  </div>

                  {/* Lines scroll area with 44x44px touch targets */}
                  <div className="flex-1 overflow-y-auto py-2 divide-y divide-[#dedad2]/60 dark:divide-[#38342e]/60 min-h-[180px] max-h-[380px]">
                    {currentOrderLines.map((line) => (
                      <div key={line.id} className="py-2.5 flex flex-col gap-2">
                        <div className="flex items-center justify-between gap-2">
                          <span className="font-bold text-sm md:text-base leading-tight line-clamp-1">{line.name}</span>
                          <div className="flex items-center gap-2 shrink-0">
                            <span className="font-extrabold text-sm md:text-base text-[#c2650a] dark:text-[#ea7a1a]">
                              {formatMoney(safePrice(line.price) * line.qty)}
                            </span>
                            <button
                              type="button"
                              onClick={() => removeCartLine(line.id)}
                              className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl flex items-center justify-center text-red-500 hover:text-red-700 hover:bg-red-50 dark:hover:bg-red-950/40 transition active:scale-95"
                              title="حذف هذا الصنف من الفاتورة"
                              aria-label="حذف الصنف"
                            >
                              <Trash2 className="w-5 h-5" />
                            </button>
                          </div>
                        </div>

                        <div className="flex items-center justify-between gap-2">
                          <input
                            type="text"
                            placeholder="ملاحظات (بدون سكر، صوص إضافي...)"
                            value={line.note}
                            onChange={(e) => updateCartItemNote(line.id, e.target.value)}
                            className="text-xs px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent w-full outline-none focus:border-[#c2650a] min-h-[38px]"
                          />

                          <div className="flex items-center gap-1.5 shrink-0">
                            <button
                              type="button"
                              onClick={() => updateCartItemQty(line.id, -1)}
                              className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] flex items-center justify-center font-bold text-lg hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition shadow-2xs"
                              title="تقليل الكمية"
                              aria-label="تقليل الكمية"
                            >
                              -
                            </button>
                            <span className="w-7 text-center font-extrabold text-sm md:text-base">{line.qty}</span>
                            <button
                              type="button"
                              onClick={() => updateCartItemQty(line.id, 1)}
                              className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] flex items-center justify-center font-bold text-lg hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition shadow-2xs"
                              title="زيادة الكمية"
                              aria-label="زيادة الكمية"
                            >
                              +
                            </button>
                          </div>
                        </div>
                      </div>
                    ))}

                    {!currentOrderLines.length && (
                      <div className="py-12 text-center text-[#6e685f] dark:text-[#9e978b] text-xs md:text-sm">
                        السلة فارغة. اختر أصنافاً من القائمة لإضافتها.
                      </div>
                    )}
                  </div>

                  {/* Totals Summary */}
                  <div className="pt-3 border-t border-[#dedad2] dark:border-[#38342e] space-y-1.5 text-xs md:text-sm">
                    <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                      <span>الإجمالي الفرعي</span>
                      <span>{formatMoney(currentTotals.sub)}</span>
                    </div>

                    {settings.serviceEnabled && (
                      <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                        <span>خدمة ({settings.serviceRate}%)</span>
                        <span>{formatMoney(currentTotals.svc)}</span>
                      </div>
                    )}

                    {settings.taxEnabled && (
                      <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                        <span>ضريبة القيمة المضافة ({settings.taxRate}%)</span>
                        <span>{formatMoney(currentTotals.tax)}</span>
                      </div>
                    )}

                    <div className="flex justify-between font-extrabold text-base md:text-lg pt-2 border-t border-dashed border-[#dedad2] dark:border-[#38342e]">
                      <span>المطلوب النهائي</span>
                      <span className="text-[#c2650a] dark:text-[#ea7a1a]">
                        {formatMoney(currentTotals.total)}
                      </span>
                    </div>
                  </div>

                  {/* Order Actions */}
                  <div className="flex gap-2 mt-4 pt-1">
                    <button
                      type="button"
                      disabled={!currentOrderLines.length}
                      onClick={() => setIsCheckoutOpen(true)}
                      className="flex-1 min-h-[48px] py-3 px-4 rounded-xl bg-[#c2650a] hover:bg-[#a85406] dark:bg-[#ea7a1a] text-white font-extrabold text-sm md:text-base shadow-sm transition active:scale-[0.98] disabled:opacity-40"
                    >
                      دفع وإنهاء الفاتورة
                    </button>
                    {currentOrderLines.length > 0 && (
                      <button
                        type="button"
                        onClick={handleClearCart}
                        className="min-h-[48px] min-w-[48px] px-3.5 rounded-xl border border-red-200 dark:border-red-900/50 text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30 transition flex items-center justify-center active:scale-95"
                        title="إفراغ الفاتورة بالكامل"
                        aria-label="إفراغ السلة"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    )}
                  </div>
                </div>
              </div>

              {/* Mobile Floating Cart Summary Button (< lg) */}
              {currentOrderLines.length > 0 && !isMobileCartOpen && (
                <div className="fixed bottom-[calc(env(safe-area-inset-bottom)+68px)] inset-x-3 sm:inset-x-6 z-30 lg:hidden">
                  <button
                    type="button"
                    onClick={() => setIsMobileCartOpen(true)}
                    className="w-full py-3.5 px-4 rounded-2xl bg-[#c2650a] hover:bg-[#a85406] dark:bg-[#ea7a1a] text-white font-bold text-sm sm:text-base shadow-2xl flex items-center justify-between active:scale-[0.98] transition border border-white/20 min-h-[52px]"
                  >
                    <div className="flex items-center gap-2.5">
                      <span className="w-7 h-7 rounded-full bg-white/25 flex items-center justify-center text-xs font-black">
                        {currentOrderLines.reduce((a, l) => a + l.qty, 0)}
                      </span>
                      <span>عرض الطلب والفاتورة</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <span className="font-black text-base">{formatMoney(currentTotals.total)}</span>
                      <ChevronLeft className="w-5 h-5 rtl:rotate-180" />
                    </div>
                  </button>
                </div>
              )}

              {/* Mobile Cart Sliding Bottom-Sheet / Drawer (< lg) */}
              {isMobileCartOpen && (
                <div className="fixed inset-0 z-50 lg:hidden flex flex-col justify-end">
                  {/* Backdrop Scrim */}
                  <div
                    className="fixed inset-0 bg-black/60 backdrop-blur-xs transition-opacity"
                    onClick={() => setIsMobileCartOpen(false)}
                  />

                  {/* Drawer Sheet Container */}
                  <div className="relative z-10 bg-white dark:bg-[#1e1c19] border-t border-[#dedad2] dark:border-[#38342e] rounded-t-3xl max-h-[85vh] flex flex-col shadow-2xl overflow-hidden pb-[calc(env(safe-area-inset-bottom)+12px)] animate-in slide-in-from-bottom duration-200">
                    {/* Pull Bar */}
                    <div className="pt-3 pb-1 flex justify-center">
                      <div className="w-12 h-1.5 rounded-full bg-gray-300 dark:bg-gray-600" />
                    </div>

                    {/* Header */}
                    <div className="px-4 py-2.5 flex items-center justify-between border-b border-[#dedad2] dark:border-[#38342e]">
                      <div className="flex items-center gap-2">
                        <ShoppingBag className="w-5 h-5 text-[#c2650a] dark:text-[#ea7a1a]" />
                        <h3 className="font-bold text-base">
                          محتويات الفاتورة ({currentOrderLines.reduce((a, l) => a + l.qty, 0)})
                        </h3>
                      </div>
                      <button
                        type="button"
                        onClick={() => setIsMobileCartOpen(false)}
                        className="w-11 h-11 min-w-[44px] min-h-[44px] flex items-center justify-center rounded-xl text-gray-500 hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95"
                        aria-label="إغلاق السلة"
                      >
                        <X className="w-6 h-6" />
                      </button>
                    </div>

                    {/* Drawer Content Area */}
                    <div className="p-4 overflow-y-auto flex-1 flex flex-col">
                      {/* Lines scroll area with 44x44px touch targets */}
                      <div className="flex-1 overflow-y-auto py-2 divide-y divide-[#dedad2]/60 dark:divide-[#38342e]/60 max-h-[42vh]">
                        {currentOrderLines.map((line) => (
                          <div key={line.id} className="py-2.5 flex flex-col gap-2">
                            <div className="flex items-center justify-between gap-2">
                              <span className="font-bold text-sm md:text-base leading-tight line-clamp-1">{line.name}</span>
                              <div className="flex items-center gap-2 shrink-0">
                                <span className="font-extrabold text-sm md:text-base text-[#c2650a] dark:text-[#ea7a1a]">
                                  {formatMoney(safePrice(line.price) * line.qty)}
                                </span>
                                <button
                                  type="button"
                                  onClick={() => removeCartLine(line.id)}
                                  className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl flex items-center justify-center text-red-500 hover:text-red-700 hover:bg-red-50 dark:hover:bg-red-950/40 transition active:scale-95"
                                  title="حذف هذا الصنف من الفاتورة"
                                  aria-label="حذف الصنف"
                                >
                                  <Trash2 className="w-5 h-5" />
                                </button>
                              </div>
                            </div>

                            <div className="flex items-center justify-between gap-2">
                              <input
                                type="text"
                                placeholder="ملاحظات (بدون سكر، صوص إضافي...)"
                                value={line.note}
                                onChange={(e) => updateCartItemNote(line.id, e.target.value)}
                                className="text-xs px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent w-full outline-none focus:border-[#c2650a] min-h-[38px]"
                              />

                              <div className="flex items-center gap-1.5 shrink-0">
                                <button
                                  type="button"
                                  onClick={() => updateCartItemQty(line.id, -1)}
                                  className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] flex items-center justify-center font-bold text-lg hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition shadow-2xs"
                                  title="تقليل الكمية"
                                  aria-label="تقليل الكمية"
                                >
                                  -
                                </button>
                                <span className="w-7 text-center font-extrabold text-sm md:text-base">{line.qty}</span>
                                <button
                                  type="button"
                                  onClick={() => updateCartItemQty(line.id, 1)}
                                  className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] flex items-center justify-center font-bold text-lg hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition shadow-2xs"
                                  title="زيادة الكمية"
                                  aria-label="زيادة الكمية"
                                >
                                  +
                                </button>
                              </div>
                            </div>
                          </div>
                        ))}

                        {!currentOrderLines.length && (
                          <div className="py-12 text-center text-[#6e685f] dark:text-[#9e978b] text-xs md:text-sm">
                            السلة فارغة. اختر أصنافاً من القائمة لإضافتها.
                          </div>
                        )}
                      </div>

                      {/* Totals Summary */}
                      <div className="pt-3 border-t border-[#dedad2] dark:border-[#38342e] space-y-1.5 text-xs md:text-sm">
                        <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                          <span>الإجمالي الفرعي</span>
                          <span>{formatMoney(currentTotals.sub)}</span>
                        </div>

                        {settings.serviceEnabled && (
                          <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                            <span>خدمة ({settings.serviceRate}%)</span>
                            <span>{formatMoney(currentTotals.svc)}</span>
                          </div>
                        )}

                        {settings.taxEnabled && (
                          <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                            <span>ضريبة القيمة المضافة ({settings.taxRate}%)</span>
                            <span>{formatMoney(currentTotals.tax)}</span>
                          </div>
                        )}

                        <div className="flex justify-between font-extrabold text-base md:text-lg pt-2 border-t border-dashed border-[#dedad2] dark:border-[#38342e]">
                          <span>المطلوب النهائي</span>
                          <span className="text-[#c2650a] dark:text-[#ea7a1a]">
                            {formatMoney(currentTotals.total)}
                          </span>
                        </div>
                      </div>

                      {/* Actions */}
                      <div className="flex gap-2 mt-4 pt-1">
                        <button
                          type="button"
                          disabled={!currentOrderLines.length}
                          onClick={() => {
                            setIsMobileCartOpen(false);
                            setIsCheckoutOpen(true);
                          }}
                          className="flex-1 min-h-[48px] py-3 px-4 rounded-xl bg-[#c2650a] hover:bg-[#a85406] dark:bg-[#ea7a1a] text-white font-extrabold text-sm md:text-base shadow-sm transition active:scale-[0.98] disabled:opacity-40"
                        >
                          دفع وإنهاء الفاتورة
                        </button>
                        {currentOrderLines.length > 0 && (
                          <button
                            type="button"
                            onClick={() => {
                              setIsMobileCartOpen(false);
                              handleClearCart();
                            }}
                            className="min-h-[48px] min-w-[48px] px-3.5 rounded-xl border border-red-200 dark:border-red-900/50 text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30 transition flex items-center justify-center active:scale-95"
                            title="إفراغ الفاتورة بالكامل"
                            aria-label="إفراغ السلة"
                          >
                            <Trash2 className="w-5 h-5" />
                          </button>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* 3. MENU MANAGEMENT VIEW (Fixed 1 vs 80 Rendering & Sorting Bug) */}
          {activeTab === 'menu' && (
            <div>
              {/* Cashier / Owner Status Banner for Menu */}
              {isOwnerAuthenticated ? (
                <div className="flex items-center justify-between p-2.5 px-3.5 rounded-2xl bg-emerald-50 dark:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-900/40 mb-4 text-xs md:text-sm">
                  <div className="flex items-center gap-2 text-emerald-800 dark:text-emerald-200 font-semibold">
                    <ShieldCheck className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
                    <span>صلاحيات المالك مفعّلة · يمكنك إضافة وتعديل وحذف الأصناف واستيراد وتصدير الإكسيل</span>
                  </div>
                  <button
                    type="button"
                    onClick={handleOwnerLock}
                    className="px-2.5 py-1 rounded-lg bg-emerald-100 dark:bg-emerald-900/50 text-emerald-800 dark:text-emerald-200 font-bold text-xs hover:bg-emerald-200 transition"
                  >
                    قفل المالك 🔒
                  </button>
                </div>
              ) : (
                <div className="flex items-center justify-between p-3 rounded-2xl bg-amber-50 dark:bg-amber-950/30 border border-amber-200 dark:border-amber-900/40 mb-4 text-xs md:text-sm">
                  <div className="flex items-center gap-2 text-amber-800 dark:text-amber-200 font-semibold">
                    <Lock className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" />
                    <span>وضع الكاشير (عرض فقط) · يتطلب رمز المالك لإضافة أو تعديل أو حذف الأصناف أو استيراد الإكسيل</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      requireOwnerAuth(
                        () => showToast('تم تفعيل صلاحيات المالك لتعديل المنيو والأسعار'),
                        'تفعيل صلاحيات تعديل المنيو',
                        'أدخل رمز PIN لتفعيل صلاحيات التعديل والحذف والاستيراد'
                      );
                    }}
                    className="px-3 py-1.5 rounded-xl bg-[#c2650a] text-white font-bold text-xs shrink-0 shadow-xs hover:bg-[#a85406] active:scale-95 transition"
                  >
                    دخول المالك 🔓
                  </button>
                </div>
              )}

              <div className="flex items-center justify-between gap-4 flex-wrap mb-5">
                <div>
                  <h2 className="text-xl md:text-2xl font-bold">قائمة المنيو والأسعار</h2>
                  <p className="text-xs md:text-sm text-[#6e685f] dark:text-[#9e978b]">
                    إجمالي {items.length} صنف مسجل في القائمة · يتم عرض {menuTableItems.length} صنف وفق الفلاتر الحالية
                  </p>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    type="button"
                    onClick={() => {
                      requireOwnerAuth(
                        () => setItemModalData({ isOpen: true, item: null }),
                        'إضافة صنف جديد للمنيو',
                        'إضافة أصناف جديدة للمنيو تتطلب رمز مرور المالك'
                      );
                    }}
                    className="inline-flex items-center gap-1.5 px-3.5 py-2 rounded-xl bg-[#c2650a] dark:bg-[#ea7a1a] text-white font-bold text-sm shadow-sm"
                  >
                    <Plus className="w-4 h-4" />
                    صنف جديد
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      requireOwnerAuth(
                        () => exportMenuToExcel(),
                        'تصدير المنيو إلى Excel',
                        'تصدير بيانات وأسعار المنيو يتطلب رمز مرور المالك'
                      );
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] font-semibold text-xs md:text-sm hover:border-[#c2650a]"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                    تصدير إكسيل (.xlsx)
                  </button>

                  <label
                    onClick={(e) => {
                      if (!isOwnerAuthenticated) {
                        e.preventDefault();
                        requireOwnerAuth(
                          () => showToast('تم تفعيل وضع المالك! يمكنك الآن اختيار ملف الإكسيل للاستيراد'),
                          'استيراد منيو جديد من Excel',
                          'استيراد وتحديث أصناف المنيو يتطلب رمز مرور المالك'
                        );
                      }
                    }}
                    className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] font-semibold text-xs md:text-sm hover:border-[#c2650a] cursor-pointer"
                  >
                    <Upload className="w-4 h-4 text-blue-600" />
                    استيراد إكسيل
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={handleMenuFileImport}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>

              {/* Menu Filters Bar with Horizontal Scroll Navigation (< and >) */}
              <div className="p-4 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] mb-4 space-y-3 shadow-xs">
                {/* Search & Sort Controls */}
                <div className="flex items-center justify-between gap-4 flex-wrap">
                  <div className="relative flex-1 min-w-[220px]">
                    <Search className="w-4 h-4 absolute right-3.5 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input
                      type="search"
                      placeholder="بحث في المنيو (الاسم أو القسم)..."
                      value={menuSearch}
                      onChange={(e) => setMenuSearch(e.target.value)}
                      className="w-full pr-10 pl-4 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-[#f5f4f0] dark:bg-[#141311] text-xs md:text-sm outline-none focus:border-[#c2650a]"
                    />
                  </div>

                  <div className="flex items-center gap-3 flex-wrap">
                    <div className="flex items-center gap-1.5">
                      <label className="text-xs font-bold text-[#6e685f] dark:text-[#9e978b]">الترتيب:</label>
                      <select
                        value={menuSort}
                        onChange={(e: any) => setMenuSort(e.target.value)}
                        className="text-xs md:text-sm px-2.5 py-1.5 rounded-lg border border-[#dedad2] dark:border-[#38342e] bg-transparent outline-none font-semibold"
                      >
                        <option value="name">الاسم أبجدياً (أ - ي)</option>
                        <option value="cat">الفئة / القسم</option>
                        <option value="price-asc">السعر (من الأقل للأعلى)</option>
                        <option value="price-desc">السعر (من الأعلى للأقل)</option>
                      </select>
                    </div>

                    <div className="flex items-center gap-1 bg-[#eae7e0] dark:bg-[#282521] p-1 rounded-xl text-xs font-bold">
                      <button
                        type="button"
                        onClick={() => setMenuStatusFilter('all')}
                        className={`px-2.5 py-1 rounded-lg transition ${
                          menuStatusFilter === 'all' ? 'bg-white dark:bg-[#1e1c19] shadow-xs' : 'text-[#6e685f] dark:text-[#9e978b]'
                        }`}
                      >
                        الكل ({items.length})
                      </button>
                      <button
                        type="button"
                        onClick={() => setMenuStatusFilter('avail')}
                        className={`px-2.5 py-1 rounded-lg transition ${
                          menuStatusFilter === 'avail' ? 'bg-white dark:bg-[#1e1c19] text-emerald-600 shadow-xs' : 'text-[#6e685f] dark:text-[#9e978b]'
                        }`}
                      >
                        متاح ({items.filter((i) => i.available !== false).length})
                      </button>
                      <button
                        type="button"
                        onClick={() => setMenuStatusFilter('unavail')}
                        className={`px-2.5 py-1 rounded-lg transition ${
                          menuStatusFilter === 'unavail' ? 'bg-white dark:bg-[#1e1c19] text-red-500 shadow-xs' : 'text-[#6e685f] dark:text-[#9e978b]'
                        }`}
                      >
                        غير متاح ({items.filter((i) => i.available === false).length})
                      </button>
                    </div>
                  </div>
                </div>

                {/* Horizontal Category Navigation Bar with Arrow Buttons (Defaults to 'الكل') */}
                <div className="pt-2 border-t border-[#dedad2]/60 dark:border-[#38342e]/60">
                  <ScrollableBar>
                    {categoriesList.map((cat) => (
                      <button
                        type="button"
                        key={cat}
                        onClick={() => setMenuCategory(cat)}
                        className={`px-3.5 py-1.5 rounded-full text-xs font-bold whitespace-nowrap transition ${
                          menuCategory === cat
                            ? 'bg-[#c2650a] text-white shadow-xs'
                            : 'bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e] text-[#1a1816] dark:text-[#f3efe9] hover:border-gray-400'
                        }`}
                      >
                        {cat}
                      </button>
                    ))}
                  </ScrollableBar>
                </div>
              </div>

              {/* Items Table Container (Scrollable, Accessible, Smoothly Rendering all 80+ rows) */}
              <div className="rounded-2xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] overflow-hidden shadow-xs">
                <div className="max-h-[640px] overflow-x-auto overflow-y-auto -webkit-overflow-scrolling-touch">
                  <table className="w-full min-w-[620px] text-start text-xs md:text-sm border-collapse">
                    <thead className="sticky top-0 z-20">
                      <tr className="bg-[#eae7e0] dark:bg-[#282521] border-b border-[#dedad2] dark:border-[#38342e] text-[#6e685f] dark:text-[#9e978b] font-bold shadow-xs">
                        <th className="p-3 text-start">الصنف</th>
                        <th className="p-3 text-start">الفئة / القسم</th>
                        <th className="p-3 text-end">سعر البيع</th>
                        <th className="p-3 text-center">حالة الإتاحة</th>
                        <th className="p-3 text-end">إجراءات</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#dedad2]/60 dark:divide-[#38342e]/60">
                      {menuTableItems.map((item) => (
                        <tr key={item.id} className="hover:bg-black/[0.015] dark:hover:bg-white/[0.015]">
                          <td className="p-3 font-bold text-sm">{item.name}</td>
                          <td className="p-3">
                            <span className="px-2.5 py-1 rounded-full bg-[#eae7e0] dark:bg-[#282521] text-xs font-semibold">
                              {item.category || 'عام'}
                            </span>
                          </td>
                          <td className="p-3 text-end font-extrabold text-[#c2650a] dark:text-[#ea7a1a]">
                            {formatMoney(safePrice(item.price))}
                          </td>
                          <td className="p-3 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                requireOwnerAuth(
                                  () => {
                                    setItems((prev) =>
                                      prev.map((i) => (i.id === item.id ? { ...i, available: !i.available } : i))
                                    );
                                    showToast(
                                      !item.available ? `«${item.name}» أصبح متاحاً الآن` : `«${item.name}» أصبح غير متاح`
                                    );
                                  },
                                  'تغيير حالة توفر الصنف',
                                  `تغيير حالة «${item.name}» يتطلب رمز مرور المالك`
                                );
                              }}
                              className={`min-h-[40px] px-3.5 py-1.5 rounded-full text-xs font-bold transition active:scale-95 ${
                                item.available
                                  ? 'bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300'
                                  : 'bg-red-100 dark:bg-red-950/60 text-red-700 dark:text-red-300'
                              }`}
                            >
                              {item.available ? '● متاح للطلب' : '○ غير متاح'}
                            </button>
                          </td>
                          <td className="p-3 text-end">
                            <div className="flex items-center justify-end gap-1.5">
                              <button
                                type="button"
                                onClick={() => {
                                  requireOwnerAuth(
                                    () => setItemModalData({ isOpen: true, item }),
                                    'تعديل بيانات الصنف والأسعار',
                                    `تعديل «${item.name}» وسعر البيع يتطلب رمز مرور المالك`
                                  );
                                }}
                                className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl border border-[#dedad2] dark:border-[#38342e] hover:bg-[#eae7e0] dark:hover:bg-[#282521] flex items-center justify-center active:scale-95 transition"
                                title="تعديل الصنف"
                                aria-label="تعديل الصنف"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button
                                type="button"
                                onClick={() => {
                                  requireOwnerAuth(
                                    () => handleDeleteMenuItem(item),
                                    'حذف صنف من المنيو',
                                    `حذف «${item.name}» من المنيو نهائياً يتطلب رمز مرور المالك`
                                  );
                                }}
                                className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl border border-red-200 dark:border-red-900/40 text-red-500 hover:bg-red-50 dark:hover:bg-red-950/30 flex items-center justify-center active:scale-95 transition"
                                title="حذف الصنف من المنيو"
                                aria-label="حذف الصنف من المنيو"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                      {!menuTableItems.length && (
                        <tr>
                          <td colSpan={5} className="p-10 text-center text-[#6e685f] dark:text-[#9e978b]">
                            <p className="font-bold text-base mb-1">لا توجد أصناف مطابقة للفلاتر المحددة</p>
                            <p className="text-xs">جرب الضغط على «الكل» في شريط الأقسام بالأعلى أو إضافة أصناف جديدة</p>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* 4. REPORTS, SALES HISTORY & FINANCIAL CLOSINGS VIEW */}
          {activeTab === 'reports' && (
            <div className="space-y-6">
              {/* Header */}
              <div className="flex items-center justify-between gap-4 flex-wrap">
                <div>
                  <h2 className="text-xl md:text-2xl font-bold">التقارير وإغلاق الحسابات</h2>
                  <p className="text-xs md:text-sm text-[#6e685f] dark:text-[#9e978b]">
                    متابعة الإيرادات، إغلاق وتسوية الفترات المالية، وتصدير السجلات إلى ملفات Excel
                  </p>
                </div>

                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    type="button"
                    onClick={exportAllSalesToExcel}
                    className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] font-semibold text-xs md:text-sm hover:border-emerald-600"
                  >
                    <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
                    تصدير كل المبيعات (.xlsx)
                  </button>

                  <label className="inline-flex items-center gap-1.5 px-3 py-2 rounded-xl bg-[#c2650a] dark:bg-[#ea7a1a] text-white font-bold text-xs md:text-sm shadow-sm cursor-pointer hover:bg-[#a85406]">
                    <Upload className="w-4 h-4" />
                    استيراد مبيعات سابقة
                    <input
                      type="file"
                      accept=".xlsx,.xls,.csv"
                      onChange={handleSalesFileImport}
                      className="hidden"
                    />
                  </label>
                </div>
              </div>

              {/* Financial Closings Section */}
              <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border-2 border-[#c2650a]/30 dark:border-[#ea7a1a]/30 shadow-sm space-y-5">
                <div className="flex items-center justify-between gap-4 flex-wrap pb-3 border-b border-[#dedad2] dark:border-[#38342e]">
                  <div className="flex items-center gap-2.5">
                    <div className="w-8 h-8 rounded-lg bg-[#fbebd8] dark:bg-[#38240f] text-[#c2650a] dark:text-[#ea7a1a] flex items-center justify-center">
                      <Lock className="w-4 h-4" />
                    </div>
                    <div>
                      <h3 className="font-bold text-base md:text-lg">إغلاق الحسابات وتسوية الفترات (Financial Closings)</h3>
                      <p className="text-xs text-[#6e685f] dark:text-[#9e978b]">
                        تحديد فترة معينة لحساب إجمالي المبيعات، تصدير تقرير الإغلاق المعتمد، وتصفية السجلات القديمة
                      </p>
                    </div>
                  </div>

                  {/* Period Switcher: Day / Week / Month */}
                  <div className="flex items-center gap-1 bg-[#eae7e0] dark:bg-[#282521] p-1 rounded-xl">
                    <button
                      type="button"
                      onClick={() => setClosingType('day')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                        closingType === 'day'
                          ? 'bg-white dark:bg-[#1e1c19] text-[#c2650a] dark:text-[#ea7a1a] shadow-xs'
                          : 'text-[#6e685f] dark:text-[#9e978b]'
                      }`}
                    >
                      يومي (Day)
                    </button>
                    <button
                      type="button"
                      onClick={() => setClosingType('week')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                        closingType === 'week'
                          ? 'bg-white dark:bg-[#1e1c19] text-[#c2650a] dark:text-[#ea7a1a] shadow-xs'
                          : 'text-[#6e685f] dark:text-[#9e978b]'
                      }`}
                    >
                      أسبوعي (Week)
                    </button>
                    <button
                      type="button"
                      onClick={() => setClosingType('month')}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition ${
                        closingType === 'month'
                          ? 'bg-white dark:bg-[#1e1c19] text-[#c2650a] dark:text-[#ea7a1a] shadow-xs'
                          : 'text-[#6e685f] dark:text-[#9e978b]'
                      }`}
                    >
                      شهري (Month)
                    </button>
                  </div>
                </div>

                {/* Dynamic Period Selector Inputs */}
                <div className="flex items-center gap-4 flex-wrap bg-[#f5f4f0] dark:bg-[#141311] p-3 rounded-xl">
                  <div className="flex items-center gap-2">
                    <Calendar className="w-4 h-4 text-[#c2650a] dark:text-[#ea7a1a]" />
                    <span className="text-xs font-bold">
                      {closingType === 'day'
                        ? 'اختر اليوم المراد إغلاقه:'
                        : closingType === 'week'
                        ? 'اختر نهاية الأسبوع (آخر 7 أيام حتى):'
                        : 'اختر الشهر المراد إغلاقه:'}
                    </span>
                  </div>

                  {closingType === 'month' ? (
                    <input
                      type="month"
                      value={closingMonthDate}
                      onChange={(e) => setClosingMonthDate(e.target.value)}
                      className="px-3 py-1.5 rounded-lg border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-xs font-bold outline-none"
                    />
                  ) : (
                    <input
                      type="date"
                      value={closingDayDate}
                      onChange={(e) => setClosingDayDate(e.target.value)}
                      className="px-3 py-1.5 rounded-lg border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-xs font-bold outline-none"
                    />
                  )}

                  <span className="text-xs text-[#6e685f] dark:text-[#9e978b] mr-auto">
                    تم العثور على <b className="text-[#1a1816] dark:text-[#f3efe9]">{closingSales.length}</b> فاتورة في هذه الفترة
                  </span>
                </div>

                {/* Detailed Period Breakdown Grid */}
                <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-3">
                  <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]">
                    <span className="text-[11px] font-semibold text-[#6e685f] dark:text-[#9e978b] block mb-1">
                      المبيعات الصافية
                    </span>
                    <span className="text-base font-extrabold text-[#c2650a] dark:text-[#ea7a1a]">
                      {formatMoney(closingStats.netTotal)}
                    </span>
                  </div>

                  <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]">
                    <span className="text-[11px] font-semibold text-[#6e685f] dark:text-[#9e978b] block mb-1">
                      المحصل نقداً (كاش)
                    </span>
                    <span className="text-base font-extrabold text-emerald-600 dark:text-emerald-400">
                      {formatMoney(closingStats.cashTotal)}
                    </span>
                  </div>

                  <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]">
                    <span className="text-[11px] font-semibold text-[#6e685f] dark:text-[#9e978b] block mb-1">
                      بطاقات / فيزا
                    </span>
                    <span className="text-base font-extrabold text-blue-600 dark:text-blue-400">
                      {formatMoney(closingStats.cardTotal)}
                    </span>
                  </div>

                  <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]">
                    <span className="text-[11px] font-semibold text-[#6e685f] dark:text-[#9e978b] block mb-1">
                      محافظ إلكترونية
                    </span>
                    <span className="text-base font-extrabold text-purple-600 dark:text-purple-400">
                      {formatMoney(closingStats.walletTotal)}
                    </span>
                  </div>

                  <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]">
                    <span className="text-[11px] font-semibold text-[#6e685f] dark:text-[#9e978b] block mb-1">
                      عدد الفواتير
                    </span>
                    <span className="text-base font-extrabold">{closingStats.count}</span>
                  </div>

                  <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]">
                    <span className="text-[11px] font-semibold text-[#6e685f] dark:text-[#9e978b] block mb-1">
                      إجمالي الفاتورة النهائي
                    </span>
                    <span className="text-base font-extrabold">{formatMoney(closingStats.grossTotal)}</span>
                  </div>
                </div>

                {/* Closing Actions: Excel Export & Purge */}
                <div className="flex items-center justify-between gap-3 flex-wrap pt-2">
                  <button
                    type="button"
                    onClick={handleExportClosingExcel}
                    disabled={!closingSales.length}
                    className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs md:text-sm shadow-sm transition disabled:opacity-40"
                  >
                    <FileSpreadsheet className="w-4 h-4" />
                    تصدير تقرير الإغلاق (Excel)
                  </button>

                  <button
                    type="button"
                    onClick={handlePurgeClosingRecords}
                    disabled={!closingSales.length}
                    className="inline-flex items-center gap-2 px-4 py-2.5 rounded-xl border border-red-300 dark:border-red-900/60 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-950/40 font-bold text-xs md:text-sm transition disabled:opacity-40"
                  >
                    <Trash2 className="w-4 h-4" />
                    مسح الحركات المصدرة من النظام
                  </button>
                </div>
              </div>

              {/* General Overview Time Range Filters with Scroll Navigation */}
              <ScrollableBar>
                {[
                  ['today', 'اليوم'],
                  ['week', 'آخر 7 أيام'],
                  ['month', 'الشهر الحالي'],
                  ['all', 'كل الفترات والسجلات']
                ].map(([key, label]) => (
                  <button
                    type="button"
                    key={key}
                    onClick={() => setReportPeriod(key as any)}
                    className={`px-4 py-1.5 rounded-full text-xs md:text-sm font-bold transition whitespace-nowrap ${
                      reportPeriod === key
                        ? 'bg-[#1a1816] text-white dark:bg-[#f3efe9] dark:text-[#141311]'
                        : 'bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] text-[#1a1816] dark:text-[#f3efe9]'
                    }`}
                  >
                    {label}
                  </button>
                ))}
              </ScrollableBar>

              {/* Key KPI Cards */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 md:gap-4">
                <div className="p-4 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] flex flex-col gap-1">
                  <span className="text-xs text-[#6e685f] dark:text-[#9e978b] font-semibold">إجمالي الإيرادات</span>
                  <span className="text-xl md:text-2xl font-extrabold text-[#c2650a] dark:text-[#ea7a1a]">
                    {formatMoney(salesReportStats.totalRev)}
                  </span>
                </div>

                <div className="p-4 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] flex flex-col gap-1">
                  <span className="text-xs text-[#6e685f] dark:text-[#9e978b] font-semibold">عدد الفواتير</span>
                  <span className="text-xl md:text-2xl font-extrabold">{salesReportStats.count}</span>
                </div>

                <div className="p-4 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] flex flex-col gap-1">
                  <span className="text-xs text-[#6e685f] dark:text-[#9e978b] font-semibold">متوسط الفاتورة</span>
                  <span className="text-xl md:text-2xl font-extrabold">
                    {formatMoney(salesReportStats.avgTicket)}
                  </span>
                </div>

                <div className="p-4 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] flex flex-col gap-1">
                  <span className="text-xs text-[#6e685f] dark:text-[#9e978b] font-semibold">أعلى صنف مبيعاً</span>
                  <span className="text-base md:text-lg font-bold line-clamp-1">
                    {salesReportStats.topItems[0]?.name || 'لا توجد مبيعات'}
                  </span>
                </div>
              </div>

              {/* Invoices Log Table with Working Delete Buttons */}
              <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e]">
                <div className="flex items-center justify-between mb-4">
                  <h3 className="font-bold text-base">سجل الفواتير ({filteredSales.length})</h3>
                </div>

                <div className="overflow-x-auto -webkit-overflow-scrolling-touch">
                  <table className="w-full min-w-[700px] text-start text-xs md:text-sm border-collapse">
                    <thead>
                      <tr className="bg-[#eae7e0] dark:bg-[#282521] border-b border-[#dedad2] dark:border-[#38342e] text-[#6e685f] dark:text-[#9e978b] font-bold">
                        <th className="p-2.5 text-start">رقم</th>
                        <th className="p-2.5 text-start">التاريخ والوقت</th>
                        <th className="p-2.5 text-start">النوع / الطاولة</th>
                        <th className="p-2.5 text-start">الأصناف</th>
                        <th className="p-2.5 text-end">الإجمالي</th>
                        <th className="p-2.5 text-center">طريقة الدفع</th>
                        <th className="p-2.5 text-end">إجراء</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#dedad2]/60 dark:divide-[#38342e]/60">
                      {[...filteredSales]
                        .reverse()
                        .slice(0, 30)
                        .map((sale) => (
                          <tr key={sale.id} className="hover:bg-black/[0.015] dark:hover:bg-white/[0.015]">
                            <td className="p-2.5 font-bold">#{sale.no}</td>
                            <td className="p-2.5 text-xs text-[#6e685f] dark:text-[#9e978b]">
                              {new Date(sale.ts).toLocaleDateString('ar-EG', { day: 'numeric', month: 'short' })}{' '}
                              {new Date(sale.ts).toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' })}
                            </td>
                            <td className="p-2.5 font-semibold">
                              {sale.type === 'table' ? sale.table : 'تيك أواي'}
                            </td>
                            <td className="p-2.5 text-xs text-[#6e685f] dark:text-[#9e978b] max-w-[220px] truncate">
                              {sale.lines.map((l) => `${l.qty}x ${l.name}`).join('، ')}
                            </td>
                            <td className="p-2.5 text-end font-extrabold text-[#c2650a] dark:text-[#ea7a1a]">
                              {formatMoney(sale.total)}
                            </td>
                            <td className="p-2.5 text-center">
                              <span className="px-2 py-0.5 rounded-full bg-[#eae7e0] dark:bg-[#282521] text-[11px] font-bold">
                                {settings.paymentMethods.find((m) => m.id === sale.pay)?.name || sale.pay}
                              </span>
                            </td>
                            <td className="p-2.5 text-end">
                              <button
                                type="button"
                                onClick={() => handleDeleteSale(sale)}
                                className="w-11 h-11 min-w-[44px] min-h-[44px] rounded-xl flex items-center justify-center text-red-500 hover:text-red-700 hover:bg-red-50 dark:hover:bg-red-950/30 transition active:scale-95"
                                title="إلغاء وحذف الفاتورة"
                                aria-label="حذف الفاتورة"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </td>
                          </tr>
                        ))}
                      {!filteredSales.length && (
                        <tr>
                          <td colSpan={7} className="p-10 text-center text-[#6e685f] dark:text-[#9e978b]">
                            <p className="font-bold text-base mb-1">لا توجد فواتير مسجلة حالياً</p>
                            <p className="text-xs">ستظهر هنا الفواتير المكتملة أو المستوردة عبر ملف إكسيل</p>
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* 5. SETTINGS VIEW */}
          {activeTab === 'settings' && (
            <div>
              <div className="flex items-center justify-between gap-4 flex-wrap mb-5">
                <div>
                  <h2 className="text-xl md:text-2xl font-bold">إعدادات الكرفان والفواتير</h2>
                  <p className="text-xs md:text-sm text-[#6e685f] dark:text-[#9e978b]">
                    تخصيص اسم المنشأة، الهاتف، العنوان، رسائل الفاتورة، الضرائب، ورسوم الخدمة، وطرق الدفع
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  <a
                    href="/BIMBO_POS_v3.html"
                    download="BIMBO_POS.html"
                    className="inline-flex items-center gap-1.5 px-4 py-2 rounded-xl bg-[#c2650a] dark:bg-[#ea7a1a] text-white font-bold text-xs md:text-sm shadow-sm hover:bg-[#a85406]"
                  >
                    <Download className="w-4 h-4" />
                    تحميل ملف HTML مفرد ومستقل
                  </a>
                </div>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 items-start">
                <div className="space-y-5">
                  {/* Owner PIN Protection & Management */}
                  <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] shadow-xs">
                    <div className="flex items-center justify-between gap-3 mb-3">
                      <div className="flex items-center gap-2.5">
                        <div className="w-9 h-9 rounded-xl bg-amber-100 dark:bg-amber-950/60 text-[#c2650a] dark:text-[#ea7a1a] flex items-center justify-center shadow-2xs">
                          <KeyRound className="w-5 h-5" />
                        </div>
                        <div>
                          <h3 className="font-bold text-base">تغيير رمز مرور المالك (Change PIN)</h3>
                          <p className="text-xs text-[#6e685f] dark:text-[#9e978b]">
                            قفل وحماية التقارير والإعدادات وتعديل قائمة المنيو
                          </p>
                        </div>
                      </div>
                      <span className="px-2.5 py-1 rounded-full text-xs font-bold bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300 flex items-center gap-1 shadow-2xs">
                        <ShieldCheck className="w-3.5 h-3.5" />
                        نظام الحماية مفعّل
                      </span>
                    </div>

                    <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e] text-xs text-[#6e685f] dark:text-[#9e978b] mb-4 space-y-1">
                      <p>• يمكنك تغيير رمز مرور المالك في أي وقت لضمان سرية تقارير المبيعات والأرباح من أعين موظفي الكاشير.</p>
                      <p>• الرمز مطلوب لحماية شاشات التقارير، الإعدادات، واستيراد وتعديل المنيو.</p>
                    </div>

                    <form onSubmit={handleChangeOwnerPin} className="space-y-3">
                      <div>
                        <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                          الرمز الحالي (Current PIN)
                        </label>
                        <input
                          type="password"
                          inputMode="numeric"
                          value={oldPinInput}
                          onChange={(e) => setOldPinInput(e.target.value)}
                          placeholder="أدخل رمز المالك الحالي"
                          className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                          required
                        />
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                            الرمز الجديد (New PIN)
                          </label>
                          <input
                            type="password"
                            inputMode="numeric"
                            value={newPinInput}
                            onChange={(e) => setNewPinInput(e.target.value)}
                            placeholder="أدخل 4 أرقام على الأقل"
                            className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                            required
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                            تأكيد الرمز الجديد (Confirm New PIN)
                          </label>
                          <input
                            type="password"
                            inputMode="numeric"
                            value={confirmNewPinInput}
                            onChange={(e) => setConfirmNewPinInput(e.target.value)}
                            placeholder="أعد إدخال الرمز الجديد للتأكيد"
                            className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                            required
                          />
                        </div>
                      </div>

                      {pinChangeError && (
                        <div className="p-2.5 rounded-xl bg-red-50 dark:bg-red-950/30 border border-red-200 dark:border-red-900/40 text-red-600 dark:text-red-400 text-xs font-bold">
                          {pinChangeError}
                        </div>
                      )}

                      {pinChangeSuccess && (
                        <div className="p-2.5 rounded-xl bg-emerald-50 dark:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-900/40 text-emerald-700 dark:text-emerald-300 text-xs font-bold">
                          {pinChangeSuccess}
                        </div>
                      )}

                      <div className="flex gap-2 pt-1 flex-wrap">
                        <button
                          type="submit"
                          className="px-4 py-2.5 rounded-xl bg-[#c2650a] hover:bg-[#a85406] dark:bg-[#ea7a1a] text-white font-bold text-xs md:text-sm shadow-xs transition active:scale-95"
                        >
                          حفظ وتحديث رمز المرور
                        </button>

                        <button
                          type="button"
                          onClick={handleOwnerLock}
                          className="px-4 py-2.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] hover:bg-[#eae7e0] dark:hover:bg-[#282521] text-xs md:text-sm font-bold transition flex items-center gap-1.5 active:scale-95"
                        >
                          <Lock className="w-4 h-4 text-amber-600" />
                          قفل جلسة المالك والعودة للكاشير
                        </button>
                      </div>
                    </form>
                  </div>

                  {/* Business Data */}
                  <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] shadow-xs">
                    <h3 className="font-bold text-base mb-3">بيانات الكرفان / النشاط التجاري</h3>
                    <div className="space-y-3">
                      <div>
                        <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                          اسم الكرفان أو المطعم
                        </label>
                        <input
                          type="text"
                          value={settings.businessName}
                          onChange={(e) => setSettings({ ...settings, businessName: e.target.value })}
                          className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                        />
                      </div>

                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div>
                          <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                            رقم الهاتف
                          </label>
                          <input
                            type="text"
                            value={settings.phone}
                            placeholder="مثال: 010xxxxxxxx"
                            onChange={(e) => setSettings({ ...settings, phone: e.target.value })}
                            className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                          />
                        </div>

                        <div>
                          <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                            رمز العملة
                          </label>
                          <input
                            type="text"
                            value={settings.currency}
                            onChange={(e) => setSettings({ ...settings, currency: e.target.value })}
                            className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                          />
                        </div>
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                          العنوان / موقع الكرفان
                        </label>
                        <input
                          type="text"
                          value={settings.address}
                          placeholder="مثال: ممشى أهل مصر - كورنيش النيل"
                          onChange={(e) => setSettings({ ...settings, address: e.target.value })}
                          className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Receipt Header & Footer */}
                  <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] shadow-xs">
                    <h3 className="font-bold text-base mb-3">نصوص ورسائل الفاتورة</h3>
                    <div className="space-y-3">
                      <div>
                        <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                          رسالة الترويسة (أعلى الفاتورة)
                        </label>
                        <input
                          type="text"
                          value={settings.receiptHeader}
                          onChange={(e) => setSettings({ ...settings, receiptHeader: e.target.value })}
                          className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                          رسالة التذييل والشكر (أسفل الفاتورة)
                        </label>
                        <input
                          type="text"
                          value={settings.receiptFooter}
                          onChange={(e) => setSettings({ ...settings, receiptFooter: e.target.value })}
                          className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                        />
                      </div>
                    </div>
                  </div>

                  {/* Taxes and Service */}
                  <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] shadow-xs">
                    <h3 className="font-bold text-base mb-3">الضرائب ورسوم الخدمة</h3>
                    <div className="space-y-4">
                      {/* Service Charge */}
                      <div className="flex items-center justify-between gap-4">
                        <label className="flex items-center gap-2 cursor-pointer font-bold text-sm">
                          <input
                            type="checkbox"
                            checked={settings.serviceEnabled}
                            onChange={(e) => setSettings({ ...settings, serviceEnabled: e.target.checked })}
                            className="w-4 h-4 rounded text-[#c2650a]"
                          />
                          <span>تفعيل رسوم الخدمة (%)</span>
                        </label>

                        <div className="w-28">
                          <input
                            type="number"
                            min="0"
                            step="any"
                            disabled={!settings.serviceEnabled}
                            value={settings.serviceRate}
                            onChange={(e) => setSettings({ ...settings, serviceRate: Number(e.target.value) || 0 })}
                            className="w-full px-2.5 py-1.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none text-center disabled:opacity-30"
                          />
                        </div>
                      </div>

                      {/* VAT */}
                      <div className="flex items-center justify-between gap-4">
                        <label className="flex items-center gap-2 cursor-pointer font-bold text-sm">
                          <input
                            type="checkbox"
                            checked={settings.taxEnabled}
                            onChange={(e) => setSettings({ ...settings, taxEnabled: e.target.checked })}
                            className="w-4 h-4 rounded text-[#c2650a]"
                          />
                          <span>تفعيل ضريبة القيمة المضافة (%)</span>
                        </label>

                        <div className="w-28">
                          <input
                            type="number"
                            min="0"
                            step="any"
                            disabled={!settings.taxEnabled}
                            value={settings.taxRate}
                            onChange={(e) => setSettings({ ...settings, taxRate: Number(e.target.value) || 0 })}
                            className="w-full px-2.5 py-1.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none text-center disabled:opacity-30"
                          />
                        </div>
                      </div>
                    </div>
                  </div>

                  {/* Payment Methods */}
                  <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] shadow-xs">
                    <h3 className="font-bold text-base mb-3">طرق الدفع المفعلة</h3>
                    <div className="space-y-2">
                      {settings.paymentMethods.map((pm) => (
                        <div
                          key={pm.id}
                          className="flex items-center justify-between p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e]"
                        >
                          <span className="font-bold text-sm">{pm.name}</span>
                          <label className="flex items-center gap-2 cursor-pointer text-xs">
                            <input
                              type="checkbox"
                              checked={pm.active}
                              onChange={(e) => {
                                const checked = e.target.checked;
                                setSettings({
                                  ...settings,
                                  paymentMethods: settings.paymentMethods.map((m) =>
                                    m.id === pm.id ? { ...m, active: checked } : m
                                  )
                                });
                              }}
                              className="w-4 h-4 rounded text-[#c2650a]"
                            />
                            <span className="text-[#6e685f] dark:text-[#9e978b]">مفعلة</span>
                          </label>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* System Reset / Clean Slate */}
                  <div className="p-5 rounded-2xl bg-red-50 dark:bg-red-950/20 border border-red-200 dark:border-red-900/40">
                    <h3 className="font-bold text-base text-red-600 dark:text-red-400 mb-2">منطقة مسح البيانات</h3>
                    <p className="text-xs text-[#6e685f] dark:text-[#9e978b] mb-4">
                      يمكنك مسح كافة البيانات المسجلة محلياً والبدء بقاعدة بيانات نظيفة تماماً بدون أي أصناف أو مبيعات.
                    </p>
                    <button
                      type="button"
                      onClick={() => {
                        setConfirmModal({
                          isOpen: true,
                          title: 'مسح جميع البيانات والبدء من جديد',
                          message: 'تحذير: سيتم مسح كافة الأصناف والمبيعات والفواتير نهائياً من هذا الجهاز. هل أنت متأكد؟',
                          confirmLabel: 'تأكيد مسح الكل',
                          danger: true,
                          onConfirm: () => {
                            const clean = getCleanInitialData();
                            setItems(clean.items);
                            setTables(clean.tables);
                            setSettings(clean.settings);
                            setSales(clean.sales);
                            setInvoiceCounter(clean.counter);
                            setTakeawayOrder({ lines: [] });
                            showToast('تم مسح جميع البيانات والبدء بصفحة بيضاء');
                            setConfirmModal((m) => ({ ...m, isOpen: false }));
                          }
                        });
                      }}
                      className="px-4 py-2 rounded-xl bg-red-600 text-white font-bold text-xs md:text-sm hover:bg-red-700 transition"
                    >
                      مسح جميع البيانات والبدء بصفحة بيضاء
                    </button>
                  </div>
                </div>

                {/* Live Preview of Thermal Receipt */}
                <div className="p-5 rounded-2xl bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] shadow-xs sticky top-20">
                  <div className="flex items-center justify-between mb-3">
                    <h3 className="font-bold text-base">معاينة حية لإيصال الفاتورة</h3>
                    <span className="text-xs px-2.5 py-0.5 rounded-full bg-emerald-100 dark:bg-emerald-950/60 text-emerald-700 dark:text-emerald-300 font-bold">
                      تحديث فوري
                    </span>
                  </div>

                  <div className="font-mono text-xs p-4 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e] whitespace-pre-wrap leading-relaxed select-none">
                    {generateReceiptText({
                      id: 'sample',
                      no: 1001,
                      ts: Date.now(),
                      type: 'table',
                      table: 'طاولة 1',
                      lines: [
                        { id: '1', name: 'صنف تجريبي للمعاينة', price: 100, qty: 1, note: '' }
                      ],
                      sub: 100,
                      disc: 0,
                      svc: settings.serviceEnabled ? (100 * settings.serviceRate) / 100 : 0,
                      tax: settings.taxEnabled ? (100 * settings.taxRate) / 100 : 0,
                      total:
                        100 +
                        (settings.serviceEnabled ? (100 * settings.serviceRate) / 100 : 0) +
                        (settings.taxEnabled ? (100 * settings.taxRate) / 100 : 0),
                      net: 100,
                      pay: 'cash'
                    })}
                  </div>
                </div>
              </div>
            </div>
          )}
        </main>
      </div>

      {/* ---------- Modals Section ---------- */}

      {/* 1. Item Add/Edit Modal */}
      {itemModalData.isOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] rounded-2xl p-6 w-full max-w-md shadow-xl">
            <div className="flex items-center justify-between pb-3 border-b border-[#dedad2] dark:border-[#38342e] mb-4">
              <h3 className="font-bold text-lg">
                {itemModalData.item ? 'تعديل الصنف' : 'إضافة صنف جديد إلى المنيو'}
              </h3>
              <button
                type="button"
                onClick={() => setItemModalData({ isOpen: false, item: null })}
                className="p-1 rounded-lg text-[#6e685f] hover:bg-[#eae7e0] dark:hover:bg-[#282521]"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form
              onSubmit={(e) => {
                e.preventDefault();
                handleSaveItemModal(new FormData(e.currentTarget));
              }}
              className="space-y-4"
            >
              <div>
                <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                  اسم الصنف
                </label>
                <input
                  name="name"
                  type="text"
                  required
                  defaultValue={itemModalData.item?.name || ''}
                  placeholder="مثال: برجر دبل تشيز"
                  className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                  الفئة / التصنيف
                </label>
                <input
                  name="category"
                  type="text"
                  required
                  defaultValue={itemModalData.item?.category || 'عام'}
                  placeholder="مثال: ساندوتشات، مشروبات ساخنة..."
                  className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                  سعر البيع ({settings.currency})
                </label>
                <input
                  name="price"
                  type="number"
                  step="any"
                  min="0"
                  required
                  defaultValue={itemModalData.item?.price || ''}
                  placeholder="مثال: 55"
                  className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-sm outline-none focus:border-[#c2650a]"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                  حالة الإتاحة
                </label>
                <select
                  name="available"
                  defaultValue={itemModalData.item ? String(itemModalData.item.available) : 'true'}
                  className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-sm outline-none"
                >
                  <option value="true">متاح للطلب الفوري</option>
                  <option value="false">غير متوفر حالياً</option>
                </select>
              </div>

              <div className="flex gap-2 pt-3 border-t border-[#dedad2] dark:border-[#38342e]">
                <button
                  type="button"
                  onClick={() => setItemModalData({ isOpen: false, item: null })}
                  className="flex-1 py-2.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] font-bold text-sm"
                >
                  إلغاء
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 rounded-xl bg-[#c2650a] dark:bg-[#ea7a1a] text-white font-bold text-sm shadow-sm"
                >
                  حفظ الصنف
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 2. Checkout Modal */}
      {isCheckoutOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] rounded-2xl p-6 w-full max-w-lg shadow-xl max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between pb-3 border-b border-[#dedad2] dark:border-[#38342e] mb-4">
              <h3 className="font-bold text-lg">إنهاء الطلب وتأكيد الدفع</h3>
              <button
                type="button"
                onClick={() => setIsCheckoutOpen(false)}
                className="p-1 rounded-lg text-[#6e685f] hover:bg-[#eae7e0] dark:hover:bg-[#282521]"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {(() => {
              const totals = calculateOrderTotals(currentOrderLines, discountType, discountVal);
              const cashReceived = Number(receivedCashAmount) || 0;
              const change = cashReceived > totals.total ? cashReceived - totals.total : 0;

              return (
                <div className="space-y-4">
                  <div className="p-3.5 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e] space-y-2 text-xs md:text-sm">
                    <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                      <span>الإجمالي الفرعي ({currentOrderLines.reduce((a, l) => a + l.qty, 0)} عنصر)</span>
                      <span>{formatMoney(totals.sub)}</span>
                    </div>

                    {totals.disc > 0 && (
                      <div className="flex justify-between text-red-600 dark:text-red-400 font-bold">
                        <span>الخصم المطبق</span>
                        <span>-{formatMoney(totals.disc)}</span>
                      </div>
                    )}

                    {settings.serviceEnabled && (
                      <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                        <span>خدمة ({settings.serviceRate}%)</span>
                        <span>{formatMoney(totals.svc)}</span>
                      </div>
                    )}

                    {settings.taxEnabled && (
                      <div className="flex justify-between text-[#6e685f] dark:text-[#9e978b]">
                        <span>ضريبة القيمة المضافة ({settings.taxRate}%)</span>
                        <span>{formatMoney(totals.tax)}</span>
                      </div>
                    )}

                    <div className="flex justify-between font-extrabold text-base pt-2 border-t border-dashed border-[#dedad2] dark:border-[#38342e]">
                      <span>المطلوب دفعه</span>
                      <span className="text-[#c2650a] dark:text-[#ea7a1a]">{formatMoney(totals.total)}</span>
                    </div>
                  </div>

                  {/* Discount Section */}
                  <div className="grid grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                        نوع الخصم
                      </label>
                      <select
                        value={discountType}
                        onChange={(e: any) => setDiscountType(e.target.value)}
                        className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-xs outline-none"
                      >
                        <option value="amt">مبلغ خصم ثابت ({settings.currency})</option>
                        <option value="pct">نسبة مئوية (%)</option>
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-1">
                        قيمة الخصم
                      </label>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        value={discountVal || ''}
                        placeholder="0"
                        onChange={(e) => setDiscountVal(Number(e.target.value) || 0)}
                        className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-transparent text-xs outline-none"
                      />
                    </div>
                  </div>

                  {/* Payment Methods Options */}
                  <div>
                    <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b] mb-2">
                      طريقة الدفع
                    </label>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                      {settings.paymentMethods
                        .filter((m) => m.active)
                        .map((m) => (
                          <button
                            type="button"
                            key={m.id}
                            onClick={() => setSelectedPayMethod(m.id)}
                            className={`p-2.5 rounded-xl border text-xs font-bold text-center transition ${
                              selectedPayMethod === m.id
                                ? 'bg-[#c2650a] text-white border-[#c2650a] shadow-xs'
                                : 'bg-white dark:bg-[#1e1c19] border-[#dedad2] dark:border-[#38342e]'
                            }`}
                          >
                            {m.name}
                          </button>
                        ))}
                    </div>
                  </div>

                  {/* Cash Change Calculator */}
                  {selectedPayMethod === 'cash' && (
                    <div className="p-3 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e] space-y-2">
                      <label className="block text-xs font-bold text-[#6e685f] dark:text-[#9e978b]">
                        المبلغ المستلم من العميل (لحساب الباقي)
                      </label>
                      <input
                        type="number"
                        min="0"
                        step="any"
                        placeholder="أدخل المبلغ المستلم نقداً"
                        value={receivedCashAmount}
                        onChange={(e) => setReceivedCashAmount(e.target.value)}
                        className="w-full px-3 py-2 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-sm outline-none"
                      />
                      {cashReceived > 0 && (
                        <div className="flex justify-between items-center text-xs font-bold pt-1">
                          <span>المتبقي للعميل:</span>
                          <span className={cashReceived >= totals.total ? 'text-emerald-600 text-sm' : 'text-red-500 text-sm'}>
                            {cashReceived >= totals.total
                              ? formatMoney(change)
                              : `غير كافٍ (متبقي ${formatMoney(totals.total - cashReceived)})`}
                          </span>
                        </div>
                      )}
                    </div>
                  )}

                  {/* Actions */}
                  <div className="flex gap-2 pt-3 border-t border-[#dedad2] dark:border-[#38342e]">
                    <button
                      type="button"
                      onClick={() => setIsCheckoutOpen(false)}
                      className="flex-1 py-3 rounded-xl border border-[#dedad2] dark:border-[#38342e] font-bold text-sm"
                    >
                      إلغاء
                    </button>
                    <button
                      type="button"
                      onClick={handleConfirmCheckout}
                      className="flex-1 py-3 rounded-xl bg-[#c2650a] hover:bg-[#a85406] dark:bg-[#ea7a1a] text-white font-extrabold text-sm md:text-base shadow-sm transition"
                    >
                      تأكيد الدفع والطباعة
                    </button>
                  </div>
                </div>
              );
            })()}
          </div>
        </div>
      )}

      {/* 3. In-App Confirmation Modal */}
      {confirmModal.isOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] rounded-2xl p-6 w-full max-w-md shadow-2xl">
            <div className="flex items-center gap-3 mb-3 text-red-600 dark:text-red-400">
              <div className="p-2 rounded-xl bg-red-100 dark:bg-red-950/50">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <h3 className="font-bold text-lg">{confirmModal.title}</h3>
            </div>

            <p className="text-sm text-[#6e685f] dark:text-[#9e978b] mb-6 leading-relaxed">
              {confirmModal.message}
            </p>

            <div className="flex gap-2.5">
              <button
                type="button"
                onClick={() => setConfirmModal((m) => ({ ...m, isOpen: false }))}
                className="flex-1 py-2.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] font-bold text-sm hover:bg-[#eae7e0] dark:hover:bg-[#282521] transition"
              >
                إلغاء
              </button>
              <button
                type="button"
                onClick={() => {
                  confirmModal.onConfirm();
                }}
                className={`flex-1 py-2.5 rounded-xl text-white font-bold text-sm shadow-sm transition ${
                  confirmModal.danger ? 'bg-red-600 hover:bg-red-700' : 'bg-[#c2650a] hover:bg-[#a85406]'
                }`}
              >
                {confirmModal.confirmLabel}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 4. Finished Receipt Modal */}
      {lastFinishedReceipt && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] rounded-2xl p-6 w-full max-w-md shadow-2xl">
            <div className="flex items-center justify-between pb-3 border-b border-[#dedad2] dark:border-[#38342e] mb-4">
              <div className="flex items-center gap-2 text-emerald-600 dark:text-emerald-400">
                <Check className="w-5 h-5" />
                <h3 className="font-bold text-lg">تم دفع الفاتورة بنجاح</h3>
              </div>
              <button
                type="button"
                onClick={() => setLastFinishedReceipt(null)}
                className="p-1 rounded-lg text-[#6e685f] hover:bg-[#eae7e0] dark:hover:bg-[#282521]"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div
              id="printable-receipt"
              className="font-mono text-xs p-4 rounded-xl bg-[#f5f4f0] dark:bg-[#141311] border border-[#dedad2] dark:border-[#38342e] whitespace-pre-wrap leading-relaxed max-h-[300px] overflow-y-auto mb-4"
            >
              {lastFinishedReceipt}
            </div>

            <div className="flex gap-2">
              <button
                type="button"
                onClick={() => {
                  navigator.clipboard.writeText(lastFinishedReceipt);
                  showToast('تم نسخ نص الفاتورة للحافظة');
                }}
                className="flex-1 py-2.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] font-bold text-xs flex items-center justify-center gap-1.5"
              >
                <Copy className="w-4 h-4" />
                نسخ الفاتورة
              </button>
              <button
                type="button"
                onClick={() => window.print()}
                className="flex-1 py-2.5 rounded-xl bg-[#c2650a] dark:bg-[#ea7a1a] text-white font-bold text-xs flex items-center justify-center gap-1.5 shadow-sm"
              >
                <Printer className="w-4 h-4" />
                طباعة الفاتورة
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 5. Secure Owner PIN Verification & First-Time Setup Modal */}
      {isPinModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 animate-in fade-in duration-150">
          <div
            className={`bg-white dark:bg-[#1e1c19] border border-[#dedad2] dark:border-[#38342e] rounded-3xl p-6 w-full max-w-sm shadow-2xl transition-transform ${
              pinShake ? 'animate-shake border-red-500' : ''
            }`}
          >
            {/* Header */}
            <div className="flex items-center justify-between pb-3 border-b border-[#dedad2] dark:border-[#38342e] mb-4">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 rounded-xl bg-amber-100 dark:bg-amber-950/60 text-[#c2650a] dark:text-[#ea7a1a] flex items-center justify-center shadow-2xs">
                  {isSetupMode ? <KeyRound className="w-4 h-4" /> : <Lock className="w-4 h-4" />}
                </div>
                <h3 className="font-bold text-base md:text-lg">
                  {isSetupMode ? 'تعيين رمز المالك لأول مرة' : 'صلاحيات المالك'}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => {
                  setIsPinModalOpen(false);
                  setEnteredPin('');
                  setSetupNewPin('');
                  setSetupConfirmPin('');
                  setSetupStep(1);
                  setPinError('');
                  setPendingAction(null);
                }}
                className="p-1 rounded-lg text-[#6e685f] hover:bg-[#eae7e0] dark:hover:bg-[#282521] transition"
                aria-label="إغلاق"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* First-Time Setup Step Indicator */}
            {isSetupMode && (
              <div className="flex items-center justify-center gap-2 mb-3">
                <div
                  className={`flex-1 py-1 px-2 rounded-lg text-xs font-bold text-center border transition ${
                    setupStep === 1
                      ? 'bg-amber-100 dark:bg-amber-950/60 border-amber-300 dark:border-amber-800 text-[#c2650a] dark:text-[#ea7a1a]'
                      : 'bg-emerald-50 dark:bg-emerald-950/40 border-emerald-300 dark:border-emerald-800 text-emerald-700 dark:text-emerald-300'
                  }`}
                >
                  1. الرمز الجديد {setupStep > 1 && '✓'}
                </div>
                <div
                  className={`flex-1 py-1 px-2 rounded-lg text-xs font-bold text-center border transition ${
                    setupStep === 2
                      ? 'bg-amber-100 dark:bg-amber-950/60 border-amber-300 dark:border-amber-800 text-[#c2650a] dark:text-[#ea7a1a]'
                      : 'bg-[#f5f4f0] dark:bg-[#141311] border-[#dedad2] dark:border-[#38342e] text-[#6e685f] dark:text-[#9e978b]'
                  }`}
                >
                  2. تأكيد الرمز
                </div>
              </div>
            )}

            {/* Prompt Description */}
            <div className="text-center mb-4">
              <h4 className="font-bold text-sm md:text-base text-[#1a1816] dark:text-[#f3efe9]">
                {isSetupMode
                  ? setupStep === 1
                    ? 'أدخل رمز مرور جديد (4 أرقام)'
                    : 'أعد إدخال الرمز للتأكيد'
                  : pinPromptTitle}
              </h4>
              <p className="text-xs text-[#6e685f] dark:text-[#9e978b] mt-1">
                {isSetupMode
                  ? setupStep === 1
                    ? 'لحماية الأرباح والتقارير، قم بإنشاء رمز سري خاص بك'
                    : 'تأكد من مطابقة الرمز لمنع حدوث أي خطأ'
                  : pinPromptDescription}
              </p>
            </div>

            {/* Visual PIN Dots Indicator (4 Dots) */}
            <div className="flex justify-center items-center gap-3.5 mb-4">
              {[0, 1, 2, 3].map((idx) => {
                const currentVal = isSetupMode
                  ? setupStep === 1
                    ? setupNewPin
                    : setupConfirmPin
                  : enteredPin;
                const isFilled = currentVal.length > idx;
                return (
                  <div
                    key={idx}
                    className={`w-4 h-4 rounded-full transition-all duration-150 ${
                      isFilled
                        ? 'bg-[#c2650a] dark:bg-[#ea7a1a] scale-125 shadow-xs'
                        : 'border-2 border-[#dedad2] dark:border-[#38342e] bg-[#f5f4f0] dark:bg-[#141311]'
                    }`}
                  />
                );
              })}
            </div>

            {/* Masked Password Input (Supports Physical Keyboard on Desktop) */}
            <form
              onSubmit={(e) => {
                e.preventDefault();
                if (isSetupMode) {
                  if (setupStep === 1) {
                    handleSetupStep1(setupNewPin);
                  } else {
                    handleSetupStep2(setupConfirmPin);
                  }
                } else {
                  handleVerifyPin();
                }
              }}
              className="space-y-3"
            >
              <input
                type="password"
                inputMode="numeric"
                pattern="[0-9]*"
                maxLength={4}
                value={
                  isSetupMode
                    ? setupStep === 1
                      ? setupNewPin
                      : setupConfirmPin
                    : enteredPin
                }
                onChange={(e) => {
                  const val = e.target.value.replace(/[^0-9]/g, '').slice(0, 4);
                  setPinError('');

                  if (isSetupMode) {
                    if (setupStep === 1) {
                      setSetupNewPin(val);
                      if (val.length === 4) {
                        handleSetupStep1(val);
                      }
                    } else {
                      setSetupConfirmPin(val);
                      if (val.length === 4) {
                        handleSetupStep2(val);
                      }
                    }
                  } else {
                    setEnteredPin(val);
                    if (val.length === 4) {
                      handleVerifyPin(val);
                    }
                  }
                }}
                placeholder="••••"
                className="w-full text-center text-2xl tracking-[0.4em] font-bold py-2 px-3 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-[#f5f4f0] dark:bg-[#141311] outline-none focus:border-[#c2650a]"
                autoFocus
              />

              {/* Touch Numeric Keypad */}
              <div className="grid grid-cols-3 gap-2 my-2 select-none">
                {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((num) => (
                  <button
                    key={num}
                    type="button"
                    onClick={() => {
                      setPinError('');
                      if (isSetupMode) {
                        if (setupStep === 1) {
                          const next = (setupNewPin + String(num)).slice(0, 4);
                          setSetupNewPin(next);
                          if (next.length === 4) handleSetupStep1(next);
                        } else {
                          const next = (setupConfirmPin + String(num)).slice(0, 4);
                          setSetupConfirmPin(next);
                          if (next.length === 4) handleSetupStep2(next);
                        }
                      } else {
                        const next = (enteredPin + String(num)).slice(0, 4);
                        setEnteredPin(next);
                        if (next.length === 4) handleVerifyPin(next);
                      }
                    }}
                    className="h-12 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-base md:text-lg font-bold hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition shadow-2xs flex items-center justify-center"
                  >
                    {num}
                  </button>
                ))}

                {/* Clear Button */}
                <button
                  type="button"
                  onClick={() => {
                    setPinError('');
                    if (isSetupMode) {
                      if (setupStep === 1) setSetupNewPin('');
                      else setSetupConfirmPin('');
                    } else {
                      setEnteredPin('');
                    }
                  }}
                  className="h-12 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-[#f5f4f0] dark:bg-[#141311] text-xs font-bold hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition flex items-center justify-center text-[#6e685f]"
                >
                  مسح
                </button>

                {/* Zero Button */}
                <button
                  type="button"
                  onClick={() => {
                    setPinError('');
                    if (isSetupMode) {
                      if (setupStep === 1) {
                        const next = (setupNewPin + '0').slice(0, 4);
                        setSetupNewPin(next);
                        if (next.length === 4) handleSetupStep1(next);
                      } else {
                        const next = (setupConfirmPin + '0').slice(0, 4);
                        setSetupConfirmPin(next);
                        if (next.length === 4) handleSetupStep2(next);
                      }
                    } else {
                      const next = (enteredPin + '0').slice(0, 4);
                      setEnteredPin(next);
                      if (next.length === 4) handleVerifyPin(next);
                    }
                  }}
                  className="h-12 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-white dark:bg-[#1e1c19] text-base md:text-lg font-bold hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition shadow-2xs flex items-center justify-center"
                >
                  0
                </button>

                {/* Backspace Button */}
                <button
                  type="button"
                  onClick={() => {
                    setPinError('');
                    if (isSetupMode) {
                      if (setupStep === 1) setSetupNewPin((prev) => prev.slice(0, -1));
                      else setSetupConfirmPin((prev) => prev.slice(0, -1));
                    } else {
                      setEnteredPin((prev) => prev.slice(0, -1));
                    }
                  }}
                  className="h-12 rounded-xl border border-[#dedad2] dark:border-[#38342e] bg-[#f5f4f0] dark:bg-[#141311] text-sm font-bold hover:bg-[#eae7e0] dark:hover:bg-[#282521] active:scale-95 transition flex items-center justify-center text-red-500"
                  title="حذف آخر رقم"
                >
                  ⌫
                </button>
              </div>

              {/* Error Notification */}
              {pinError && (
                <div className="p-2 rounded-xl bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-900/40 text-red-600 dark:text-red-400 text-xs font-bold text-center animate-in fade-in">
                  {pinError}
                </div>
              )}

              {/* Setup Mode Back to Step 1 Button */}
              {isSetupMode && setupStep === 2 && (
                <div className="text-center pt-1">
                  <button
                    type="button"
                    onClick={() => {
                      setSetupStep(1);
                      setSetupConfirmPin('');
                      setPinError('');
                    }}
                    className="text-xs text-[#c2650a] dark:text-[#ea7a1a] hover:underline font-bold"
                  >
                    ← العودة لتعديل الرمز المكتوب
                  </button>
                </div>
              )}

              {/* Action Buttons */}
              <div className="flex gap-2 pt-2 border-t border-[#dedad2] dark:border-[#38342e]">
                <button
                  type="button"
                  onClick={() => {
                    setIsPinModalOpen(false);
                    setEnteredPin('');
                    setSetupNewPin('');
                    setSetupConfirmPin('');
                    setSetupStep(1);
                    setPinError('');
                    setPendingAction(null);
                  }}
                  className="flex-1 py-2.5 rounded-xl border border-[#dedad2] dark:border-[#38342e] font-bold text-xs md:text-sm hover:bg-[#eae7e0] dark:hover:bg-[#282521] transition"
                >
                  إلغاء
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 rounded-xl bg-[#c2650a] hover:bg-[#a85406] dark:bg-[#ea7a1a] text-white font-bold text-xs md:text-sm shadow-sm transition active:scale-95"
                >
                  {isSetupMode ? (setupStep === 1 ? 'التالي' : 'حفظ الرمز') : 'تأكيد الدخول'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Toast Notification */}
      {toastMessage && (
        <div className="fixed bottom-6 right-6 z-50 bg-[#1a1816] dark:bg-[#f3efe9] text-white dark:text-[#141311] px-4 py-2.5 rounded-xl shadow-lg text-xs md:text-sm font-bold animate-in fade-in slide-in-from-bottom-2 duration-200">
          {toastMessage}
        </div>
      )}
    </div>
  );
}
