import React, { useState, useEffect, useRef, useMemo } from 'react';
import { User, Lock, Eye, EyeOff, ArrowLeft, MessageCircle, Activity, Clock, Layers, Bell, BellOff, Download, Trash2, Menu, X, Pen, Building, LayoutDashboard, Wrench, Users, Settings, LogOut, ShieldCheck } from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import * as XLSX from 'xlsx';
import {
  collection,
  doc,
  setDoc,
  getDoc,
  getDocs,
  onSnapshot,
  updateDoc,
  deleteDoc,
  query,
  where,
  orderBy,
  serverTimestamp,
  getDocFromServer,
  writeBatch
} from 'firebase/firestore';
import { db, auth } from './lib/firebase';
import { User as FirebaseUser } from 'firebase/auth';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// Statistics Component for Sections
const SectionStats = ({
  requests,
  activeSection,
  userUnits,
  userRole
}: {
  requests: any[],
  activeSection: string,
  userUnits: string[],
  userRole: string
}) => {
  const currentYear = new Date().getFullYear();
  const currentMonth = new Date().getMonth();

  const stats = useMemo(() => {
    const sectionReqs = activeSection === 'dashboard' ? requests : requests.filter(r => r.sectionId === activeSection);

    const yearReqs = sectionReqs.filter(r => new Date(r.timestamp).getFullYear() === currentYear);
    const monthReqs = sectionReqs.filter(r =>
      new Date(r.timestamp).getMonth() === currentMonth &&
      new Date(r.timestamp).getFullYear() === currentYear
    );

    const pendingReqs = sectionReqs.filter(r => !r.status || r.status === '' || r.status.includes('انتظار'));

    const completedYear = yearReqs.filter(r => r.status?.includes('منجز') || r.status?.includes('مكتمل'));
    const completedMonth = monthReqs.filter(r => r.status?.includes('منجز') || r.status?.includes('مكتمل'));

    const isAdmin = userRole === 'admin';

    const unitReqs = (isAdmin || userUnits.length > 0)
      ? sectionReqs.filter(r =>
        (isAdmin || userUnits.includes(r.assignedUnit)) &&
        (!r.status || (!r.status.includes('قيد العمل') && !r.status.includes('منجز') && !r.status.includes('مكتمل')))
      )
      : [];

    return [
      {
        title: `الطلبات المرفوعة (${currentYear})`,
        count: yearReqs.length,
        icon: <Activity />,
        color: 'from-blue-600 to-blue-800'
      },
      {
        title: `الطلبات لشهر (${new Date().toLocaleDateString('ar-EG', { month: 'long' })})`,
        count: monthReqs.length,
        icon: <Clock />,
        color: 'from-blue-400 to-blue-600'
      },
      {
        title: 'طلبات القسم (جديدة/انتظار)',
        count: pendingReqs.length,
        icon: <Bell />,
        color: 'from-orange-500 to-orange-700',
        pulse: pendingReqs.length > 0
      },
      {
        title: `الطلبات المنجزة (${currentYear})`,
        count: completedYear.length,
        icon: <Layers />,
        color: 'from-green-600 to-green-800'
      },
      {
        title: `المنجزة لشهر (${new Date().toLocaleDateString('ar-EG', { month: 'long' })})`,
        count: completedMonth.length,
        icon: <Activity />,
        color: 'from-green-400 to-green-600'
      },
      {
        title: 'طلبات الوحدة',
        count: unitReqs.length,
        icon: <BellOff />,
        color: 'from-red-600 to-red-800',
        pulse: !isAdmin && unitReqs.length > 0
      },
    ];
  }, [requests, activeSection, userUnits, userRole, currentYear, currentMonth]);

  return (
    <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-6 gap-5 mb-10 rtl">
      {stats.map((s, idx) => (
        <motion.div
          key={idx}
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: idx * 0.1 }}
          className="relative group"
        >
          {s.pulse && (
            <div className={`absolute top-2 left-2 w-3 h-3 rounded-full bg-red-500 animate-ping z-20`}></div>
          )}
          <div className="bg-white p-6 rounded-[2.2rem] shadow-sm border border-gray-100 hover:border-[#1a5e1a]/20 transition-all hover:shadow-[0_20px_50px_rgba(0,0,0,0.06)] hover:-translate-y-1 overflow-hidden relative group-hover:bg-gradient-to-b group-hover:from-white group-hover:to-gray-50/50">
            <div className={`absolute top-0 right-0 w-24 h-24 bg-gradient-to-br ${s.color} opacity-[0.03] -mr-8 -mt-8 rounded-full transition-transform group-hover:scale-150`}></div>
            <div className="flex flex-col items-center gap-4 relative z-10 text-center">
              <div className={`w-14 h-14 bg-gradient-to-br ${s.color} text-white rounded-[1.2rem] flex items-center justify-center shadow-lg transition-all duration-500 group-hover:rotate-12 group-hover:scale-110`}>
                {React.cloneElement(s.icon as any, { size: 28, strokeWidth: 2.5 })}
              </div>
              <div className="space-y-1.5">
                <p className="text-[11px] font-black text-gray-400 leading-tight h-8 flex items-center justify-center uppercase tracking-tighter px-1">{s.title}</p>
                <div className="flex items-baseline justify-center">
                  <h4 className="text-3xl font-black text-gray-900 tracking-tighter">{s.count}</h4>
                </div>

              </div>
            </div>
          </div>
        </motion.div>
      ))}
    </div>
  );
};

// TagInput component for better list management
const TagInput = ({
  tags,
  onAdd,
  onRemove,
  placeholder,
  label
}: {
  tags: string[],
  onAdd: (tag: string) => void,
  onRemove: (idx: number) => void,
  placeholder: string,
  label: string
}) => {
  const [inputValue, setInputValue] = useState('');

  const handleAdd = () => {
    if (inputValue.trim()) {
      onAdd(inputValue.trim());
      setInputValue('');
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Enter' || e.key === ',') {
      e.preventDefault();
      handleAdd();
    }
  };

  return (
    <div className="space-y-2">
      <label className="text-sm font-bold text-gray-600 pr-2">{label}</label>
      <div className="min-h-[100px] p-4 bg-gray-50 border-2 border-gray-100 rounded-2xl focus-within:border-[#1a5e1a] focus-within:bg-white transition-all shadow-inner">
        <div className="flex flex-wrap gap-2 mb-3">
          {tags.map((tag, idx) => (
            <span key={idx} className="bg-white border border-gray-100 px-3 py-1.5 rounded-xl text-[10px] font-black text-[#1a5e1a] flex items-center gap-2 shadow-sm animate-in zoom-in-90 group hover:border-[#1a5e1a] transition-all">
              {tag}
              <button
                type="button"
                onClick={() => onRemove(idx)}
                className="text-gray-300 hover:text-red-500 transition-colors"
              >✕</button>
            </span>
          ))}
          {tags.length === 0 && <span className="text-[10px] text-gray-400 italic">لا توجد عناصر مضافة بعد...</span>}
        </div>
        <div className="flex gap-2">
          <input
            value={inputValue}
            onChange={e => setInputValue(e.target.value)}
            onKeyDown={handleKeyDown}
            className="flex-1 bg-transparent border-none focus:outline-none text-sm font-bold text-gray-800 placeholder:text-gray-300 w-full"
            placeholder={placeholder}
          />
          <button
            type="button"
            onClick={handleAdd}
            className="w-8 h-8 bg-[#1a5e1a] text-white rounded-lg flex items-center justify-center text-lg hover:scale-110 transition-transform shadow-md shrink-0"
          >+</button>
        </div>
      </div>
    </div>
  );
};

export default function App() {
  const [users, setUsers] = useState<any[]>([]);
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState(false);
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [activeSection, setActiveSection] = useState('dashboard');
  const [activeUserSubSection, setActiveUserSubSection] = useState('list'); // 'list' or 'form'
  const [editingUser, setEditingUser] = useState<any>(null);
  const [activeSectionTab, setActiveSectionTab] = useState('list'); // 'list' or 'raise'
  const [editingRequest, setEditingRequest] = useState<any>(null);

  // Add body scroll lock when modal is open
  useEffect(() => {
    if (activeSectionTab === 'raise') {
      document.body.style.overflow = 'hidden';
      window.scrollTo(0, 0);
    } else {
      document.body.style.overflow = 'unset';
    }
    return () => {
      document.body.style.overflow = 'unset';
    };
  }, [activeSectionTab]);
  const [checkingSession, setCheckingSession] = useState(true);
  const [whatsappModal, setWhatsappModal] = useState<{ isOpen: boolean, request: any }>({ isOpen: false, request: null });
  const [selectedRecipientPhones, setSelectedRecipientPhones] = useState<string[]>([]);
  const [whatsappNote, setWhatsappNote] = useState('');
  const [selectedWhatsappUnit, setSelectedWhatsappUnit] = useState('');
  const [requests, setRequests] = useState<any[]>([]);
  const [sections, setSections] = useState<any[]>([]);
  const [globalUnits, setGlobalUnits] = useState<string[]>([]);
  const [whatsappTemplate, setWhatsappTemplate] = useState('');

  const [currentTime, setCurrentTime] = useState(new Date());
  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);
  const usernameRef = useRef<HTMLInputElement>(null);
  const isRequestsInitialized = useRef(false);

  // Notifications State & Logic
  const [notificationPermission, setNotificationPermission] = useState<NotificationPermission>(
    typeof Notification !== 'undefined' ? Notification.permission : 'default'
  );

  const requestNotificationPermission = () => {
    if (!('Notification' in window)) {
      alert('متصفحك لا يدعم الإشعارات أو ربما الإطار الحالي مقيد.');
      return;
    }

    if (Notification.permission === 'denied') {
      alert('لقد تم رفض الاشعارات مسبقاً. يرجى الضغط على أيقونة "القفل" في شريط العنوان وإعادة تعيين "السماح" للبرنامج.');
      return;
    }

    if (Notification.permission === 'granted') {
      new Notification('الاشعارات مفعلة', { body: 'ستصلك التنبيهات تلقائياً عند وجود طلبات جديدة.' });
      return;
    }

    try {
      Notification.requestPermission().then(permission => {
        setNotificationPermission(permission);
        if (permission === 'granted') {
          new Notification('تم تفعيل الإشعارات', {
            body: 'ستصلك تنبيهات عند وجود طلبات صيانة جديدة في أقسامك.',
            icon: 'https://cdn-icons-png.flaticon.com/512/833/833602.png',
            dir: 'rtl'
          });
        }
      }).catch(err => {
        console.error("Permission request failed:", err);
        alert('يجب فتح البرنامج في صفحة مستقلة (New Tab) لتفعيل الإشعارات.');
      });
    } catch (e) {
      // Fallback for older browsers
      Notification.requestPermission(permission => {
        setNotificationPermission(permission);
      });
    }
  };

  const sendLocalNotification = (title: string, body: string) => {
    if (notificationPermission === 'granted') {
      try {
        new Notification(title, {
          body,
          icon: 'https://cdn-icons-png.flaticon.com/512/833/833602.png',
          dir: 'rtl'
        });
      } catch (e) {
        console.error("Notification failed", e);
      }
    }
  };

  const currentUserData = useMemo(() => users.find(u => u.username === username), [users, username]);
  const userSectionPerms = useMemo(() => currentUserData?.permissions[activeSection], [currentUserData, activeSection]);

  const visibleRequests = useMemo(() => {
    if (!currentUserData) return [];
    if (currentUserData.role === 'admin') return requests;

    // Get all section IDs where the user has at least 'access' permission
    const allowedSectionIds = sections
      .filter(s => currentUserData.permissions?.[s.id]?.access)
      .map(s => s.id);

    return requests.filter(r => allowedSectionIds.includes(r.sectionId));
  }, [requests, sections, currentUserData]);

  const [isSyncing, setIsSyncing] = useState(true);

  // Session Duration (adjustable)
  const SESSION_DURATION = 50 * 60 * 1000; // 50 minutes in milliseconds

  // Session Persistence
  useEffect(() => {
    const savedLogin = localStorage.getItem('engineeringLogin');
    if (savedLogin) {
      try {
        const loginData = JSON.parse(savedLogin);
        const currentTime = new Date().getTime();
        const timeDiff = currentTime - loginData.timestamp;

        // Check if session is valid and not expired
        if (timeDiff < loginData.expiryTime && loginData.username) {
          setIsLoggedIn(true);
          setUsername(loginData.username);
        } else {
          // Explictly remove if expired
          localStorage.removeItem('engineeringLogin');
        }
      } catch (e) {
        console.error("Session restoration failed", e);
        localStorage.removeItem('engineeringLogin');
      }
    }
    setCheckingSession(false);
  }, []);

  // Focus username input on login page load
  useEffect(() => {
    if (!isLoggedIn && !checkingSession && usernameRef.current) {
      // Use a small timeout to ensure DOM is ready and any animations are complete
      const timer = setTimeout(() => {
        usernameRef.current?.focus();
      }, 300);
      return () => clearTimeout(timer);
    }
  }, [isLoggedIn, checkingSession]);

  // Connection Test
  useEffect(() => {
    const testConnection = async () => {
      try {
        await getDocFromServer(doc(db, 'test-connection', 'dummy'));
      } catch (error) {
        if (error instanceof Error && error.message.includes('the client is offline')) {
          console.error("Please check your Firebase configuration.");
        }
      }
    };
    testConnection();
  }, []);

  // Sync Logic
  useEffect(() => {
    let syncCount = 0;
    const checkSyncStatus = () => {
      syncCount++;
      if (syncCount >= 4) setIsSyncing(false);
    };

    // 1. Sync Sections (Public)
    const unsubscribeSections = onSnapshot(collection(db, 'sections'), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setSections(data);
      checkSyncStatus();
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'sections'));

    // 2. Sync Settings (Public)
    const unsubscribeSettings = onSnapshot(doc(db, 'settings', 'global'), (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        setGlobalUnits(data.globalUnits || []);
        setWhatsappTemplate(data.whatsappTemplate || '');
      }
      checkSyncStatus();
    }, (err) => handleFirestoreError(err, OperationType.GET, 'settings/global'));

    // 3. Sync Users
    const unsubscribeUsers = onSnapshot(collection(db, 'users'), (snapshot) => {
      const data = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
      setUsers(data);
      checkSyncStatus();
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'users'));

    // 4. Sync Requests
    const unsubscribeRequests = onSnapshot(collection(db, 'requests'), (snapshot) => {
      const data = snapshot.docs.map(doc => {
        const item = doc.data();
        if (!item) return null;
        return {
          ...item,
          docId: doc.id, // Store original Firestore document ID
          timestamp: item.timestamp?.toDate ? item.timestamp.toDate() : new Date(item.timestamp),
          completionTime: item.completionTime?.toDate ? item.completionTime.toDate() : (item.completionTime ? new Date(item.completionTime) : null),
          lastSentDate: item.lastSentDate?.toDate ? item.lastSentDate.toDate() : (item.lastSentDate ? new Date(item.lastSentDate) : null)
        };
      }).filter(Boolean);

      setRequests(data);
      checkSyncStatus();
    }, (err) => handleFirestoreError(err, OperationType.LIST, 'requests'));

    return () => {
      unsubscribeUsers();
      unsubscribeSections();
      unsubscribeRequests();
      unsubscribeSettings();
    };
  }, []); // Remove isAuthReady dependency

  // Seed Data Logic
  useEffect(() => {
    const seed = async () => {
      try {
        const usersSnap = await getDocs(collection(db, 'users'));
        if (usersSnap.empty) {
          console.log('Seeding initial data...');
          const batch = writeBatch(db);

          // Seed Users
          const initialUsers = [
            {
              name: 'مرتضى الجنابي',
              username: 'admin',
              role: 'admin',
              password: '1234',
              phone: '07735559707',
              whatsapp: '07735559707',
              email: 'fhjuj01@gmail.com',
              department: 'المشاريع الهندسية',
              permissions: {
                mamalji: {
                  access: true,
                  raiseRequest: true,
                  viewRequests: true,
                  isUnitLead: true,
                  units: ['الكهرباء'],
                  editWorkStatus: true,
                  editAssignedUnit: true,
                  editDeptNotes: true,
                  editRequestDetails: true,
                  exportReports: true
                },
                hurr: {
                  access: true,
                  raiseRequest: true,
                  viewRequests: true,
                  isUnitLead: false,
                  units: [],
                  editWorkStatus: true,
                  editAssignedUnit: true,
                  editDeptNotes: true,
                  editRequestDetails: true,
                  exportReports: true
                }
              }
            }
          ];
          initialUsers.forEach(u => {
            batch.set(doc(db, 'users', u.username), u);
          });

          // Seed Sections
          const initialSections = [
            {
              id: 'mamalji',
              name: 'المعملجي',
              description: 'إدارة طلبات صيانة قسم المعمل',
              isActive: true,
              questions: [
                { id: 'q1', title: 'البناية', type: 'dropdown', options: ['بناية A', 'بناية B', 'الورشة المركزية'], required: true },
                { id: 'q2', title: 'موقع الصيانة', type: 'dropdown', options: ['الطابق الأول', 'المجمع الخارجي', 'المختبر'], required: true },
                { id: 'q3', title: 'نوع العمل', type: 'text', required: true },
                { id: 'q4', title: 'ملاحظات إضافية', type: 'textarea', required: false }
              ],
              availableStatuses: ['قيد الانتظار', 'قيد العمل (جاري التنفيذ)', 'تم الإنجاز (منجز)', 'مرفوض / ملغي'],
              sectionUnits: ['الكهرباء', 'الميكانيك', 'المدني', 'الاتصالات', 'التبريد', 'المولدات']
            },
            {
              id: 'hurr',
              name: 'الحر',
              description: 'إدارة طلبات صيانة منطقة الحر',
              isActive: true,
              questions: [
                { id: 'h1', title: 'الموقع', type: 'dropdown', options: ['المجمع الشمالي', 'المجمع الجنوبي'], required: true },
                { id: 'h2', title: 'نوع الخلل', type: 'checkbox', options: ['كهرباء', 'ماء', 'تكييف'], required: true },
                { id: 'h3', title: 'وصف المشكلة', type: 'textarea', required: true }
              ],
              availableStatuses: ['قيد الانتظار', 'قيد العمل (جاري التنفيذ)', 'تم الإنجاز (منجز)', 'مرفوض / ملغي'],
              sectionUnits: ['الكهرباء', 'الميكانيك', 'المدني', 'الاتصالات', 'التبريد', 'المولدات']
            }
          ];
          initialSections.forEach(s => {
            batch.set(doc(db, 'sections', s.id), s);
          });

          // Seed Global Settings
          batch.set(doc(db, 'settings', 'global'), {
            globalUnits: ['الكهرباء', 'الميكانيك', 'المدني', 'الاتصالات', 'التبريد', 'المولدات'],
            whatsappTemplate: `*طلب صيانة جديد (إسناد الصيانة الهندسي)*
—————————————
*الموقع:* مجمع {sectionName}
*الوحدة:* {unit}
*مقدم الطلب:* {submitter}
*القسم:* {department}
—————————————
*التفاصيل:*
{details}
{note}
—————————————
*الحالة الحالية:* {status}`
          });

          await batch.commit();
        }
      } catch (err) {
        console.error('Seed Error:', err);
      }
    };
    seed();
  }, []);

  const [activeSettingsSection, setActiveSettingsSection] = useState('');

  // Set default active settings section once sections are loaded
  useEffect(() => {
    if (!activeSettingsSection && sections.length > 0) {
      setActiveSettingsSection(sections[0].id);
    }
  }, [sections, activeSettingsSection]);
  const [isSectionModalOpen, setIsSectionModalOpen] = useState(false);
  const [editingSectionConfig, setEditingSectionConfig] = useState<any>(null);
  const [sectionForm, setSectionForm] = useState({
    id: '',
    name: '',
    description: '',
    isActive: true,
    questions: [] as any[],
    availableStatuses: [] as string[],
    sectionUnits: [] as string[]
  });

  const openSectionModal = (section?: any) => {
    if (section) {
      setEditingSectionConfig(section);
      setSectionForm({
        ...section,
        questions: [...section.questions],
        availableStatuses: [...section.availableStatuses],
        sectionUnits: [...section.sectionUnits]
      });
    } else {
      setEditingSectionConfig(null);
      setSectionForm({
        id: Math.random().toString(36).substr(2, 9),
        name: '',
        description: '',
        isActive: true,
        questions: [
          { id: Math.random().toString(36).substr(2, 9), title: 'السؤال الأول', type: 'text', required: true }
        ],
        availableStatuses: ['قيد الانتظار', 'قيد العمل (جاري التنفيذ)', 'تم الإنجاز (منجز)', 'مرفوض / ملغي'],
        sectionUnits: [...globalUnits]
      });
    }
    setIsSectionModalOpen(true);
  };

  const saveSectionConfig = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      if (editingSectionConfig) {
        await updateDoc(doc(db, 'sections', editingSectionConfig.id), sectionForm);
      } else {
        await setDoc(doc(db, 'sections', sectionForm.id), sectionForm);
        setActiveSettingsSection(sectionForm.id);
      }
      setIsSectionModalOpen(false);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, `sections/${sectionForm.id}`);
    }
  };

  const deleteSection = async (id: string) => {
    if (sections.length <= 1) {
      alert('يجب أن يحتوي النظام على قسم واحد على الأقل.');
      return;
    }
    if (window.confirm('هل أنت متأكد من حذف هذا القسم نهائياً؟ سيؤدي ذلك لفقدان الوصول لجميع طلباته.')) {
      try {
        await deleteDoc(doc(db, 'sections', id));
        const newSections = sections.filter(s => s.id !== id);
        setActiveSettingsSection(newSections[0]?.id || '');
        setIsSectionModalOpen(false);
      } catch (err) {
        handleFirestoreError(err, OperationType.DELETE, `sections/${id}`);
      }
    }
  };

  // Sections State & Logic (Already handled by Firestore listeners)

  // Tracking for changes
  const notifiedRequests = useRef<Set<string>>(new Set());
  const prevRequestStates = useRef<Record<string, string>>({});
  const prevRequestUnits = useRef<Record<string, string>>({});

  // Tracking for stale requests (pending > 24h)
  const notifiedStaleRequests = useRef<Set<string>>(new Set());

  useEffect(() => {
    const currentUserData = users.find(u => u.username === username);

    // Initialize if not done yet
    if (!isRequestsInitialized.current && requests.length > 0) {
      requests.forEach(r => r?.id && notifiedRequests.current.add(r.id));
      prevRequestStates.current = requests.reduce((acc, r) => r?.id ? ({ ...acc, [r.id]: r.status }) : acc, {});
      prevRequestUnits.current = requests.reduce((acc, r) => r?.id ? ({ ...acc, [r.id]: r.assignedUnit }) : acc, {});
      isRequestsInitialized.current = true;
      return;
    }

    if (!currentUserData || requests.length === 0) return;

    const isAdmin = currentUserData.role === 'admin';

    // 1. Check for New Requests
    requests.forEach(r => {
      if (!r || !r.id) return;
      if (!notifiedRequests.current.has(r.id)) {
        // This is a new request
        const userSectionPerms = currentUserData.permissions[r.sectionId];
        const hasAccess = isAdmin || userSectionPerms?.access;

        // Notify if user has access and didn't create it themselves
        if (hasAccess && r.submitterUsername !== username) {
          const sectionName = sections.find(s => s.id === r.sectionId)?.name || 'غير معروف';
          sendLocalNotification(
            'هنالك اشعار عمل جديد في المجمع',
            `قسم: ${sectionName} | مقدم الطلب: ${r.submitterName}`
          );
        }
        notifiedRequests.current.add(r.id);
      }
    });

    // 2. Check for Status or Unit Updates
    requests.forEach(r => {
      if (!r || !r.id) return;
      const oldStatus = prevRequestStates.current[r.id];
      const oldUnit = prevRequestUnits.current[r.id];
      const userSectionPerms = currentUserData.permissions?.[r.sectionId];
      const userUnits = userSectionPerms?.units || [];

      // Notify if assigned to user's unit
      if (oldUnit !== undefined && r.assignedUnit !== oldUnit && r.assignedUnit !== '' && (isAdmin || userUnits.includes(r.assignedUnit))) {
        sendLocalNotification(
          '🎯 تم تعيين طلب لوحدتك',
          `طلب رقم #${r.id.toString().slice(-4)} | الوحدة: ${r.assignedUnit}`
        );
      }

      // Notify if status updated for a request they have access to
      if (oldStatus !== undefined && r.status !== oldStatus && r.status !== '') {
        const hasAccess = isAdmin || userSectionPerms?.access;
        if (hasAccess) {
          sendLocalNotification(
            '🔄 تحديث حالة الطلب',
            `طلب رقم #${r.id.toString().slice(-4)} | الحالة الجديدة: ${r.status}`
          );
        }
      }
    });

    // 3. Check for Stale Requests (Pending > 24h)
    const checkStaleRequests = () => {
      const now = Date.now();
      const delay24h = 24 * 60 * 60 * 1000;

      requests.forEach(r => {
        if (!r?.id) return;
        const isPending = !r.status || r.status.includes('انتظار');
        if (isPending && r.timestamp) {
          const age = now - new Date(r.timestamp).getTime();
          if (age > delay24h && !notifiedStaleRequests.current.has(r.id)) {
            const userSectionPerms = currentUserData.permissions[r.sectionId];
            const isAdminOrLead = isAdmin || (userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(r.assignedUnit));

            if (isAdminOrLead) {
              sendLocalNotification(
                '⚠️ تنبيه: طلب متأخر',
                `الطلب #${r.id.toString().slice(-4)} لا يزال قيد الانتظار منذ أكثر من 24 ساعة.`
              );
              notifiedStaleRequests.current.add(r.id);
            }
          }
        }
      });
    };

    checkStaleRequests(); // Run immediately

    // Update Refs
    prevRequestStates.current = requests.reduce((acc, r) => r?.id ? ({ ...acc, [r.id]: r.status }) : acc, {});
    prevRequestUnits.current = requests.reduce((acc, r) => r?.id ? ({ ...acc, [r.id]: r.assignedUnit }) : acc, {});
  }, [requests, sections, username, users]);

  // Periodic check for stale requests (every hour)
  useEffect(() => {
    const timer = setInterval(() => {
      const now = Date.now();
      const delay24h = 24 * 60 * 60 * 1000;

      requests.forEach(r => {
        const isPending = !r.status || r.status.includes('انتظار');
        if (isPending && r.timestamp) {
          const age = now - new Date(r.timestamp).getTime();
          if (age > delay24h && !notifiedStaleRequests.current.has(r.id)) {
            sendLocalNotification(
              `⚠️ تنبيه: طلب متأخر #${r.id}`,
              `الطلب لا يزال قيد الانتظار منذ أكثر من 24 ساعة. (مجمع: ${sections.find(s => s.id === r.sectionId)?.name})`
            );
            notifiedStaleRequests.current.add(r.id);
          }
        }
      });
    }, 60 * 60 * 1000); // Check every hour

    return () => clearInterval(timer);
  }, [requests, sections]);

  // Auto-request on login
  useEffect(() => {
    if (isLoggedIn && notificationPermission === 'default') {
      requestNotificationPermission();
    }
  }, [isLoggedIn]);

  // Function to toggle section status (Maintenance mode)
  const toggleSectionActive = async (id: string) => {
    const s = sections.find(s => s.id === id);
    if (!s) return;
    try {
      await updateDoc(doc(db, 'sections', id), { isActive: !s.isActive });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, `sections/${id}`);
    }
  };

  const updateSectionName = async (id: string, newName: string) => {
    try {
      await updateDoc(doc(db, 'sections', id), { name: newName });
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, `sections/${id}`);
    }
  };

  // Form State for Adding/Editing User
  const [userForm, setUserForm] = useState({
    name: '',
    username: '',
    role: 'user',
    password: '',
    phone: '',
    whatsapp: '',
    email: '',
    department: '',
    permissions: {} as any
  });

  // Handle Form Open for Add or Edit
  const openUserForm = (user?: any) => {
    if (user) {
      setEditingUser(user);
      setUserForm({ ...user, role: user.role || 'user' });
    } else {
      setEditingUser(null);
      // Initialize with default permissions for all sections
      const defaultPerms = {} as any;
      sections.forEach(s => {
        defaultPerms[s.id] = {
          access: false,
          raiseRequest: false,
          viewRequests: false,
          isUnitLead: false,
          units: [],
          editWorkStatus: false,
          editAssignedUnit: false,
          editDeptNotes: false,
          editRequestDetails: false,
          exportReports: false
        };
      });
      setUserForm({
        name: '',
        username: '',
        role: 'user',
        password: '',
        phone: '',
        whatsapp: '',
        email: '',
        department: '',
        permissions: defaultPerms
      });
    }
    setActiveUserSubSection('form');
  };

  const addQuestion = (sectionId: string) => {
    const newQuestion = {
      id: `q${Date.now()}`,
      title: 'سؤال جديد',
      type: 'text' as const,
      required: false
    };
    setSections(sections.map(s => s.id === sectionId ? {
      ...s,
      questions: [...s.questions, newQuestion]
    } : s));
  };

  const removeQuestion = (sectionId: string, questionId: string) => {
    setSections(sections.map(s => s.id === sectionId ? {
      ...s,
      questions: s.questions.filter(q => q.id !== questionId)
    } : s));
  };

  const updateQuestion = (sectionId: string, questionId: string, updates: any) => {
    setSections(sections.map(s => s.id === sectionId ? {
      ...s,
      questions: s.questions.map(q => q.id === questionId ? { ...q, ...updates } : q)
    } : s));
  };

  const handleSaveUser = async (e: React.FormEvent) => {
    e.preventDefault();
    try {
      if (editingUser) {
        // If username changed, we need to delete old and create new
        if (editingUser.username !== userForm.username) {
          const batch = writeBatch(db);
          batch.delete(doc(db, 'users', editingUser.username));
          batch.set(doc(db, 'users', userForm.username), userForm);
          await batch.commit();
        } else {
          await updateDoc(doc(db, 'users', editingUser.username), userForm);
        }
      } else {
        // Check if username already exists
        const userDoc = await getDoc(doc(db, 'users', userForm.username));
        if (userDoc.exists()) {
          alert('اسم المستخدم موجود مسبقاً، يرجى اختيار اسم آخر.');
          return;
        }
        await setDoc(doc(db, 'users', userForm.username), userForm);
      }
      setActiveUserSubSection('list');
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, `users/${userForm.username}`);
    }
  };

  const handleDeleteUser = async (userId: string, userName: string) => {
    if (userId === 'admin') {
      alert('لا يمكن حذف حساب المدير الأساسي.');
      return;
    }
    if (window.confirm(`هل أنت متأكد من حذف المستخدم "${userName}" نهائياً من قاعدة البيانات؟ لا يمكن التراجع عن هذه العملية.`)) {
      try {
        await deleteDoc(doc(db, 'users', userId));
        alert('✅ تم حذف المستخدم بنجاح.');
      } catch (err) {
        console.error('Delete error:', err);
        alert('❌ فشل في حذف المستخدم. الرجاء التحقق من الاتصال بالإنترنت.');
      }
    }
  };

  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [filterStatus, setFilterStatus] = useState('الكل');

  const filteredRequests = useMemo(() => {
    // Advanced Arabic Normalization for "Fuzzy" matching
    const normalize = (txt: any) =>
      String(txt || '').replace(/[أإآ]/g, 'ا').replace(/ة/g, 'ه').replace(/ى/g, 'ي').toLowerCase().trim();

    return requests.filter(r => {
      if (r.sectionId !== activeSection) return false;

      // Date range filtering
      if (startDate) {
        const start = new Date(startDate);
        start.setHours(0, 0, 0, 0);
        if (r.timestamp < start) return false;
      }
      if (endDate) {
        const end = new Date(endDate);
        end.setHours(23, 59, 59, 999);
        if (r.timestamp > end) return false;
      }

      // Status filtering
      if (filterStatus !== 'الكل') {
        const reqStatus = r.status || 'قيد الانتظار';
        if (reqStatus !== filterStatus) return false;
      }

      if (!searchTerm) return true;

      // Multi-Field & Multi-Term Search Logic
      const normalizedSearch = normalize(searchTerm);
      const searchTerms = normalizedSearch.split(/\s+/).filter(t => t.length > 0);

      // Combine all searchable fields into one normalized string
      const searchableFields = [
        r.id.toString(),
        normalize(r.submitterName),
        normalize(r.submitterDept),
        normalize(r.assignedUnit),
        normalize(r.status || 'قيد الانتظار'),
        normalize(r.sectionNotes),
        ...Object.values(r.answers).map(val => normalize(val))
      ];

      const combinedText = searchableFields.join(' ');

      // A record matches if EVERY term in the search string is found ANYWHERE in its searchable fields
      return searchTerms.every(term => combinedText.includes(term));
    }).sort((a, b) => b.timestamp.getTime() - a.timestamp.getTime());
  }, [requests, activeSection, startDate, endDate, filterStatus, searchTerm]);

  const getFilteredRequests = () => filteredRequests;

  const exportToExcel = (data: any[], fileName: string) => {
    const sectionNames = sections.reduce((acc, s) => ({ ...acc, [s.id]: s.name }), {} as any);

    const formattedData = data.map(r => {
      const section = sections.find(s => s.id === r.sectionId);
      const row: any = {
        'رقم الطلب': r.id,
        'تاريخ الرفع': r.timestamp instanceof Date ? r.timestamp.toLocaleDateString('ar-EG') : new Date(r.timestamp?.seconds * 1000).toLocaleDateString('ar-EG'),
        'وقت الرفع': r.timestamp instanceof Date ? r.timestamp.toLocaleTimeString('ar-EG') : new Date(r.timestamp?.seconds * 1000).toLocaleTimeString('ar-EG'),
        'القسم/المجمع': sectionNames[r.sectionId] || r.sectionId,
        'مقدم الطلب': r.submitterName,
        'القسم الوظيفي': r.submitterDept,
        'الوحدة المكلفة': r.assignedUnit || 'بانتظار التوزيع',
        'الحالة': r.status || 'قيد الانتظار',
        'وقت الإنجاز': r.completionTime ? (r.completionTime instanceof Date ? r.completionTime.toLocaleString('ar-EG') : new Date(r.completionTime?.seconds * 1000).toLocaleString('ar-EG')) : '-',
        'ملاحظات الشعبة': r.sectionNotes || '-',
      };

      // Add dynamic questions
      if (section) {
        section.questions.forEach((q: any) => {
          row[q.title] = r.answers[q.id] || '-';
        });
      }

      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(formattedData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "الطلبات");

    // Auto-size columns (rough approximation)
    const maxWidths = formattedData.reduce((acc: any, row: any) => {
      Object.keys(row).forEach((key, i) => {
        const val = row[key] ? row[key].toString() : '';
        acc[i] = Math.max(acc[i] || 10, val.length + 5);
      });
      return acc;
    }, {});
    worksheet['!cols'] = Object.keys(maxWidths).map(i => ({ wch: maxWidths[i] }));

    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  };

  const handleExportData = (format: 'excel') => {
    const dataToExport = getFilteredRequests();
    if (dataToExport.length === 0) {
      alert('لا توجد بيانات للتصدير وفق الفلاتر الحالية.');
      return;
    }

    if (format === 'excel') {
      const sectionName = sections.find(s => s.id === activeSection)?.name || activeSection;
      exportToExcel(dataToExport, `تقرير_طلبات_${sectionName}_${new Date().toLocaleDateString('ar-EG')}`);
    }
  };

  const exportAllRequestsToExcel = () => {
    if (requests.length === 0) {
      alert('لا توجد طلبات في النظام للتصدير.');
      return;
    }
    exportToExcel(requests, `جميع_طلبات_النظام_${new Date().toLocaleDateString('ar-EG')}`);
  };

  const getUserPendingCount = (sectionId: string) => {
    const user = users.find(u => u.username === username);
    if (!user) return 0;
    const userUnits = user.permissions[sectionId]?.units || [];
    return requests.filter(r =>
      r.sectionId === sectionId &&
      (!r.status || r.status === 'قيد الانتظار') &&
      userUnits.includes(r.assignedUnit)
    ).length;
  };

  const handleRaiseRequest = async (e: React.FormEvent) => {
    e.preventDefault();
    const activeSectionData = sections.find(s => s.id === activeSection);
    if (!activeSectionData) return;

    const formData = new FormData(e.currentTarget as HTMLFormElement);
    const answers: any = {};
    activeSectionData.questions.forEach(q => {
      if (q.type === 'checkbox') {
        answers[q.id] = formData.getAll(q.id).join(' - ');
      } else {
        answers[q.id] = formData.get(q.id);
      }
    });

    const currentUserData = users.find(u => u.username === username);

    const requestId = requests.length > 0 ? Math.max(...requests.map(r => r.id)) + 1 : 1;

    const newRequest = {
      id: requestId,
      sectionId: activeSection,
      timestamp: serverTimestamp(),
      submitterName: currentUserData?.name,
      submitterDept: currentUserData?.department,
      answers: answers,
      assignedUnit: '',
      status: '',
      completionTime: null,
      sectionNotes: '',
      lastSent: ''
    };

    try {
      await setDoc(doc(db, 'requests', requestId.toString()), newRequest);
      setActiveSectionTab('list');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, `requests/${requestId}`);
    }
  };

  const updateRequest = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingRequest) return;

    // Ensure completion time is updated if status changes to completed
    const updatedRequest = { ...editingRequest };
    const oldRequest = requests.find(r => r.id === editingRequest.id);
    if (updatedRequest.status !== oldRequest?.status && (updatedRequest.status?.includes('منجز') || updatedRequest.status?.includes('مكتمل'))) {
      updatedRequest.completionTime = serverTimestamp();
    }

    try {
      // Use Firestore document ID for update
      const docId = (editingRequest as any).docId || editingRequest.id.toString();
      await updateDoc(doc(db, 'requests', docId), {
        ...updatedRequest,
        timestamp: oldRequest.timestamp // Keep original timestamp
      });
      setEditingRequest(null);
    } catch (err) {
      const docId = (editingRequest as any).docId || editingRequest.id.toString();
      handleFirestoreError(err, OperationType.UPDATE, `requests/${docId}`);
    }
  };

  const deleteRequest = async (id: any) => {
    if (!window.confirm('🚨 هل أنت متأكد من رغبتك في حذف هذا الطلب نهائياً من النظام؟ لا يمكن التراجع عن هذه العملية.')) return;

    // Explicitly use the serial ID as the Doc ID
    const docId = id?.toString().trim();
    if (!docId) {
      alert('خطأ: لم يتم تحديد معرف الطلب بشكل صحيح.');
      return;
    }

    console.log('Attempting to delete request with Serial Doc ID:', docId);
    try {
      await deleteDoc(doc(db, 'requests', docId));
      setEditingRequest(null);
      alert('✅ تم حذف الطلب بنجاح من قاعدة البيانات.');
    } catch (err) {
      console.error('Error deleting request:', err);
      alert('❌ فشل الحذف: تأكد من الصلاحيات أو حاول الحذف يدوياً من السيرفر.');
      handleFirestoreError(err, OperationType.DELETE, `requests/${docId}`);
    }
  };

  const sendWhatsApp = (request: any, recipient: string) => {
    if (!request) return;
    const sectionName = sections.find(s => s.id === request.sectionId)?.name;

    let templateMsg = whatsappTemplate;
    templateMsg = templateMsg.replace('{sectionName}', sectionName || '');
    templateMsg = templateMsg.replace('{unit}', request.assignedUnit || '');
    templateMsg = templateMsg.replace('{submitter}', request.submitterName || '');
    templateMsg = templateMsg.replace('{department}', request.submitterDept || '');
    templateMsg = templateMsg.replace('{details}', Object.values(request.answers).join('\n'));
    templateMsg = templateMsg.replace('{note}', '');
    templateMsg = templateMsg.replace('{status}', request.status || 'قيد الانتظار');

    const message = encodeURIComponent(templateMsg);
    window.open(`https://wa.me/${recipient.replace(/\s+/g, '')}?text=${message}`, '_blank');

    // Update last sent info
    const recipientName = users.find(u => u.phone === recipient || u.whatsapp === recipient)?.name || recipient;
    setRequests(requests.map(r => r.id === request.id ? {
      ...r,
      lastSent: `whatsapp: ${recipient} - ${new Date().toLocaleDateString('ar-EG')}`,
      lastSentRecipients: [recipientName],
      lastSentDate: new Date()
    } : r));
    setWhatsappModal({ isOpen: false, request: null });
  };

  const handleSendAdvancedWhatsApp = () => {
    if (!whatsappModal.request || selectedRecipientPhones.length === 0) return;

    const request = whatsappModal.request;
    const sectionName = sections.find(s => s.id === request.sectionId)?.name;
    const recipients = users.filter(u => selectedRecipientPhones.includes(u.phone));
    const recipientNames = recipients.map(u => u.name).join(' و ');

    // Generate message from template
    let templateMsg = whatsappTemplate;
    templateMsg = templateMsg.replace('{sectionName}', sectionName || '');
    templateMsg = templateMsg.replace('{unit}', request.assignedUnit || '');
    templateMsg = templateMsg.replace('{submitter}', request.submitterName || '');
    templateMsg = templateMsg.replace('{department}', request.submitterDept || '');
    templateMsg = templateMsg.replace('{details}', Object.values(request.answers).join('\n'));
    templateMsg = templateMsg.replace('{note}', whatsappNote ? `\n—————————————\n*ملاحظات إضافية:*\n${whatsappNote}` : '');
    templateMsg = templateMsg.replace('{status}', request.status || 'قيد الانتظار');

    const message = encodeURIComponent(templateMsg);

    // 1. WhatsApp Number Formatting (Automatic Iraq Code 964)
    const formatPhoneNumber = (phone: string) => {
      let cleaned = phone.replace(/\D/g, ''); // Remove all non-digits
      if (cleaned.startsWith('0')) {
        cleaned = '964' + cleaned.substring(1);
      } else if (cleaned.startsWith('7') && cleaned.length === 10) {
        cleaned = '964' + cleaned;
      }
      return cleaned;
    };

    const firstRecipient = formatPhoneNumber(selectedRecipientPhones[0]);
    const isMobile = /iPhone|iPad|iPod|Android/i.test(navigator.userAgent);

    // 2. Direct Redirection: Using web.whatsapp.com for Desktop skips the intermediate screen
    // Using api.whatsapp.com for mobile ensures it triggers the app directly
    const finalUrl = isMobile
      ? `https://api.whatsapp.com/send?phone=${firstRecipient}&text=${message}`
      : `https://web.whatsapp.com/send?phone=${firstRecipient}&text=${message}`;

    window.open(finalUrl, '_blank');

    // Update last sent info with names
    setRequests(requests.map(r => r.id === request.id ? {
      ...r,
      lastSent: `${recipientNames} (${new Date().toLocaleDateString('ar-EG')})`,
      lastSentRecipients: recipients.map(u => u.name),
      lastSentDate: new Date()
    } : r));

    setWhatsappModal({ isOpen: false, request: null });
    setSelectedRecipientPhones([]);
    setWhatsappNote('');
  };

  const togglePermission = (sectionId: string, field: string) => {
    setUserForm(prev => {
      const sectionPerms = prev.permissions[sectionId] || {
        access: false,
        raiseRequest: false,
        viewRequests: false,
        isUnitLead: false,
        units: [],
        editWorkStatus: false,
        editAssignedUnit: false,
        editDeptNotes: false,
        editRequestDetails: false,
        exportReports: false
      };

      return {
        ...prev,
        permissions: {
          ...prev.permissions,
          [sectionId]: {
            ...sectionPerms,
            [field]: !sectionPerms[field]
          }
        }
      };
    });
  };

  const toggleUnitPermission = (sectionId: string, unit: string) => {
    setUserForm(prev => {
      const sectionPerms = prev.permissions[sectionId] || {
        access: false,
        raiseRequest: false,
        viewRequests: false,
        isUnitLead: false,
        units: [],
        editWorkStatus: false,
        editAssignedUnit: false,
        editDeptNotes: false,
        editRequestDetails: false,
        exportReports: false
      };

      const currentUnits = [...(sectionPerms.units || [])];
      const index = currentUnits.indexOf(unit);
      if (index > -1) {
        currentUnits.splice(index, 1);
      } else {
        currentUnits.push(unit);
      }

      return {
        ...prev,
        permissions: {
          ...prev.permissions,
          [sectionId]: {
            ...sectionPerms,
            units: currentUnits
          }
        }
      };
    });
  };


  const handleLogout = () => {
    setIsLoggedIn(false);
    setUsername('');
    setPassword('');
    localStorage.removeItem('engineeringLogin');
    setActiveSection('dashboard');
  };

  const handleLogin = (e?: React.FormEvent) => {
    if (e) e.preventDefault();

    // Check against users state instead of hardcoded values
    const foundUser = users.find(u => u.username === username && u.password === password);
    if (foundUser) {
      setIsLoggedIn(true);
      setError(false);

      // Save session with timestamp and 50-minute expiry
      const loginData = {
        username: foundUser.username,
        timestamp: new Date().getTime(),
        expiryTime: SESSION_DURATION
      };
      localStorage.setItem('engineeringLogin', JSON.stringify(loginData));

      // Redirect non-admins to their first allowed section if they don't have access to the current activeSection
      if (sections.length > 0 && foundUser.permissions) {
        const firstAllowed = sections.find(s => foundUser.permissions?.[s.id]?.access);
        if (firstAllowed) {
          setActiveSection(firstAllowed.id);
        } else {
          // If no sections allowed, maybe just show a restricted view or logout
          setActiveSection('restricted');
        }
      }
    } else {
      setError(true);
      setPassword('');
      if (usernameRef.current) usernameRef.current.focus();

      // Hide error after 3 seconds
      setTimeout(() => setError(false), 3000);
    }
  };

  if (isSyncing || checkingSession) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center rtl">
        <div className="flex flex-col items-center gap-4 text-center">
          <div className="w-16 h-16 border-4 border-[#1a5e1a] border-t-transparent rounded-full animate-spin"></div>
          <p className="text-gray-800 font-black text-xl">جاري الاتصال بقاعدة البيانات...</p>
          <p className="text-gray-500 text-sm">برجاء الانتظار قليلاً للمزامنة</p>
        </div>
      </div>
    );
  }

  if (isLoggedIn) {
    return (
      <>
        <div className="min-h-screen bg-slate-50 flex flex-col lg:flex-row rtl overflow-x-hidden">
          {/* Mobile Header */}
          <header className="lg:hidden bg-[#1a5e1a] text-white p-4 flex items-center justify-between sticky top-0 z-[100] shadow-xl w-full border-b border-white/10 backdrop-blur-md">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-white/20 p-1 flex items-center justify-center">
                <img src="https://i.ibb.co/0jpV8nXS/image.png" alt="Logo" className="w-full h-full object-contain rounded-lg" />
              </div>
              <div className="text-right">
                <h1 className="text-sm font-black tracking-tight text-white">اسناد الصيانة الهندسي</h1>
                <p className="text-[8px] font-bold opacity-60">نظام إدارة صيانة مجمع العميد</p>
              </div>
            </div>
            <button
              onClick={() => setIsSidebarOpen(!isSidebarOpen)}
              className="p-2.5 bg-white/10 rounded-xl active:scale-90 transition-all border border-white/20 shadow-inner"
            >
              {isSidebarOpen ? <X size={24} /> : <Menu size={24} />}
            </button>
          </header>

          {/* Sidebar Overlay for Mobile */}
          <AnimatePresence>
            {isSidebarOpen && activeSectionTab !== 'raise' && (
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                onClick={() => setIsSidebarOpen(false)}
                className="lg:hidden fixed inset-0 bg-black/60 backdrop-blur-sm z-[90]"
              />
            )}
          </AnimatePresence>

          {/* Sidebar */}
          <aside
            className={`fixed lg:fixed top-0 right-0 h-screen w-[300px] bg-gradient-to-b from-[#1a5e1a] to-[rgba(0,0,0,0.98)] text-white p-8 flex flex-col z-[110] shadow-[10px_0_50px_rgba(0,0,0,0.3)] transition-transform duration-500 lg:translate-x-0 ${isSidebarOpen ? 'translate-x-0' : 'translate-x-full'} border-l border-white/10 overflow-y-auto shrink-0 backdrop-blur-[20px] ${activeSectionTab === 'raise' ? 'hidden' : ''}`}
          >
            {/* Sidebar Header - Professional Identity */}
            <div className="sidebar-header text-center mb-10 relative cursor-pointer" onClick={() => { setActiveSection('dashboard'); setIsSidebarOpen(false); }}>
              <div className="logo w-[120px] h-[120px] rounded-full mx-auto mb-5 overflow-hidden border-[3px] border-white/20 transition-all duration-300 hover:scale-105 hover:rotate-2 shadow-2xl">
                <img src="https://i.ibb.co/0jpV8nXS/image.png" alt="Logo" className="w-full h-full object-contain p-2" />
              </div>

              <div className="space-y-1">
                <h3 className="text-[1.4rem] font-black tracking-tight">العتبة العباسية</h3>
                <p className="text-[0.95rem] font-bold opacity-90 leading-relaxed">قسم المشاريع الهندسية</p>
                <p className="text-[0.85rem] font-bold opacity-70 leading-relaxed">شعبة صيانة مجموعة العميد</p>
                <div className="h-px w-20 bg-gradient-to-r from-transparent via-white/30 to-transparent mx-auto my-4"></div>
              </div>
            </div>

            <nav className="sidebar-nav flex-1 space-y-2 scrollbar-hide">
              <button
                onClick={() => { setActiveSection('dashboard'); setIsSidebarOpen(false); }}
                className={`nav-item w-full flex items-center px-5 py-4 rounded-[10px] transition-all duration-300 font-medium ${activeSection === 'dashboard' ? 'bg-white/15 border-r-4 border-white active' : 'hover:bg-white/10 hover:-translate-x-2'}`}
              >
                <LayoutDashboard size={20} className="ml-4" />
                <span className="text-[16px]">الرئيسية</span>
              </button>

              <div className="py-2">
                <p className="px-5 text-[11px] font-black text-white/40 uppercase tracking-[0.2em] mb-3 text-right">أقسام مجمع العميد</p>
                <div className="space-y-2">
                  {sections.map(section => {
                    const userData = users.find(u => u.username === username);
                    const userPerms = userData?.permissions[section.id];
                    const isAdmin = userData?.role === 'admin';
                    if (!userPerms?.access && !isAdmin) return null;

                    const isActive = activeSection === section.id;
                    const pendingCount = visibleRequests.filter(r => r.sectionId === section.id && (!r.status || r.status.includes('انتظار') || r.status === '')).length;

                    return (
                      <button
                        key={section.id}
                        onClick={() => { setActiveSection(section.id); setIsSidebarOpen(false); }}
                        className={`nav-item w-full flex items-center px-5 py-4 rounded-[10px] transition-all duration-300 font-medium relative ${isActive ? 'bg-white/15 border-r-4 border-white active' : 'hover:bg-white/10 hover:-translate-x-2'}`}
                      >
                        <div className="flex items-center gap-3 flex-1">
                          <Wrench size={20} className="ml-4" />
                          <span className="text-[16px]">{section.name}</span>
                          {pendingCount > 0 && (
                            <span className="mr-auto bg-red-500 text-white text-[10px] font-black px-2 py-0.5 rounded-full shadow-lg border border-white/20">
                              {pendingCount}
                            </span>
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {users.find(u => u.username === username)?.role === 'admin' && (
                <div className="pt-2 border-t border-white/10 mt-2 space-y-2">
                  <p className="px-5 text-[11px] font-black text-white/40 uppercase tracking-[0.2em] mb-3 text-right">إدارة النظام</p>
                  <button
                    onClick={() => { setActiveSection('users'); setIsSidebarOpen(false); }}
                    className={`nav-item w-full flex items-center px-5 py-4 rounded-[10px] transition-all duration-300 font-medium ${activeSection === 'users' ? 'bg-white/15 border-r-4 border-white active' : 'hover:bg-white/10 hover:-translate-x-2'}`}
                  >
                    <Users size={20} className="ml-4" />
                    <span className="text-[16px]">إدارة المستخدمين</span>
                  </button>
                  <button
                    onClick={() => { setActiveSection('settings'); setIsSidebarOpen(false); }}
                    className={`nav-item w-full flex items-center px-5 py-4 rounded-[10px] transition-all duration-300 font-medium ${activeSection === 'settings' ? 'bg-white/15 border-r-4 border-white active' : 'hover:bg-white/10 hover:-translate-x-2'}`}
                  >
                    <Settings size={20} className="ml-4" />
                    <span className="text-[16px]">إعدادات النظام</span>
                  </button>
                </div>
              )}
            </nav>

            <div className="user-info mt-auto pt-8 border-t border-white/15 text-center space-y-4">
              <div className="space-y-1">
                <div id="userName" className="text-[1.2rem] font-black text-white mb-1 px-4 py-1.5 bg-white/10 rounded-full inline-block">
                  {users.find(u => u.username === username)?.name || username}
                </div>
                <div id="userRole" className="text-[0.95rem] font-bold opacity-90">
                  {users.find(u => u.username === username)?.role === 'admin' ? 'مدير النظام' : 'صلاحية محدودة'}
                </div>
              </div>

              <p className="text-[0.85rem] font-bold opacity-70 leading-relaxed px-4">وحدة الكهرباء والاتصالات</p>
              <p className="text-[0.8rem] font-black opacity-60 px-4 mt-2">Design by MURTADA AL_JANABY</p>

              <button
                onClick={handleLogout}
                className="w-full flex items-center justify-center gap-3 px-6 py-4 bg-red-600/10 hover:bg-red-600 text-white rounded-[15px] transition-all duration-300 font-bold text-[15px] group border border-red-600/20 active:scale-95 mt-4"
              >
                <span>الخروج</span>
                <LogOut size={18} className="group-hover:translate-x-1 transition-transform" />
              </button>
            </div>
          </aside>

          <main className={`flex-1 min-w-0 bg-[#f8fafc] transition-all duration-300 ${activeSectionTab === 'raise' ? 'lg:pr-0 overflow-hidden h-screen' : 'lg:pr-[320px]'}`}>
            <div className={`max-w-7xl mx-auto p-4 md:p-6 lg:p-8 space-y-6 ${activeSectionTab === 'raise' ? 'hidden sm:block opacity-20 pointer-events-none grayscale' : ''}`}>

              {/* Professional Header Section - Optimized & Compact */}
              {activeSection === 'dashboard' && (
                <div className="modern-header p-4 md:p-6 rounded-[2.5rem] flex flex-col xl:flex-row items-center justify-between gap-6 relative overflow-hidden border border-white shadow-[0_20px_60px_-15px_rgba(0,0,0,0.1)] bg-white">
                  <div className="absolute top-0 left-0 right-0 h-1.5 bg-gradient-to-r from-[#1a5e1a] via-[#facc15] to-[#1a5e1a]"></div>
                  <div className="absolute -bottom-24 -right-24 w-64 h-64 bg-green-50 rounded-full blur-3xl opacity-50"></div>

                  {/* Welcome Section */}
                  <div className="welcome-section text-center md:text-right w-full xl:w-auto order-1 xl:order-2 relative z-10">
                    <div className="welcome-badge inline-flex items-center gap-2 px-4 py-1.5 bg-[#1a5e1a]/10 text-[#1a5e1a] text-[11px] font-black rounded-full uppercase tracking-widest mb-3 border border-[#1a5e1a]/20 shadow-sm transition-all hover:bg-[#1a5e1a]/15">
                      <User size={14} fill="currentColor" className="opacity-80" />
                      <span>مرحباً بك عزيزي</span>
                    </div>
                    <h1 className="text-3xl md:text-5xl font-black text-gray-900 tracking-tighter leading-tight flex flex-col md:flex-row items-center md:items-end gap-4">
                      <span className="bg-gradient-to-l from-[#1a5e1a] to-gray-600 bg-clip-text text-transparent">{users.find(u => u.username === username)?.name || username}</span>
                    </h1>
                    <div className="department-info flex items-center justify-center md:justify-start gap-2 mt-3 text-gray-500 font-black text-sm">
                      <ShieldCheck size={16} className="text-[#1a5e1a]" />
                      <span className="tracking-wide font-black">اسناد الصيانة الهندسي</span>
                    </div>
                  </div>

                  {/* Actions Section */}
                  <div className="flex flex-col md:flex-row items-center gap-4 w-full xl:w-auto order-2 xl:order-1 relative z-10 rtl">
                    {/* Time Card - Compact */}
                    <div className="time-card flex items-center gap-3 bg-[#f1f5f9] px-3 py-1.5 rounded-full border border-gray-200/50 shadow-sm min-w-[160px] h-12">
                      <div className="bg-[#1a5e1a] text-white w-9 h-9 rounded-full flex items-center justify-center shadow-md grow-0 shrink-0">
                        <Clock size={18} />
                      </div>
                      <div className="text-right leading-none flex-1">
                        <div className="text-[17px] font-black text-[#1a5e1a] tabular-nums tracking-tighter mb-0.5">
                          {currentTime.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' })}
                        </div>
                        <div className="text-[10px] font-black text-gray-500 uppercase tracking-tight">
                          {currentTime.toLocaleDateString('ar-EG', { weekday: 'long' })}
                        </div>
                      </div>
                    </div>

                    {/* Support Link - Compact */}
                    <a
                      href="https://wa.me/9647735559707"
                      target="_blank"
                      className="h-12 bg-gradient-to-r from-[#2ecc71] to-[#27ae60] px-3 rounded-full flex items-center gap-2.5 text-white shadow-lg transition-all hover:scale-105 active:scale-95 group min-w-[160px]"
                    >
                      <div className="w-9 h-9 bg-white text-[#25D366] rounded-full flex items-center justify-center shadow-md grow-0 shrink-0">
                        <MessageCircle size={20} fill="currentColor" />
                      </div>
                      <div className="text-right flex-1 leading-tight">
                        <p className="text-[14px] font-black tracking-tight whitespace-nowrap">الدعم الفني</p>
                      </div>
                    </a>

                    {/* Notifications Toggle */}
                    <button
                      onClick={requestNotificationPermission}
                      className={`w-12 h-12 rounded-full flex items-center justify-center transition-all bg-white shadow-lg border-2 hover:scale-110 active:scale-95 relative group ${notificationPermission === 'granted' ? 'text-[#1a5e1a] border-[#1a5e1a]/20' : 'text-gray-400 border-gray-100'}`}
                      title={notificationPermission === 'granted' ? "الإشعارات مفعلة" : "تفعيل الإشعارات"}
                    >
                      <div className="relative z-10 transition-transform group-hover:rotate-12">
                        {notificationPermission === 'granted' ? <Bell size={24} fill="currentColor" className="opacity-20 absolute inset-0 scale-125 blur-[1px]" /> : null}
                        {notificationPermission === 'granted' ? <Bell size={24} className="animate-swing" /> : <BellOff size={24} />}
                      </div>
                      {notificationPermission === 'granted' && (
                        <span className="absolute top-2.5 right-3 w-2.5 h-2.5 bg-red-500 rounded-full border-2 border-white shadow-sm animate-pulse"></span>
                      )}
                    </button>
                  </div>
                </div>
              )}


              {/* Dynamic Content Sections */}
              <div className="relative">
                {activeSection === 'dashboard' && (
                  <div className="animate-in fade-in slide-in-from-top-4 duration-700">
                    <SectionStats
                      requests={visibleRequests}
                      activeSection={activeSection}
                      userUnits={[]}
                      userRole={currentUserData?.role || 'user'}
                    />
                  </div>
                )}

                {/* Dashboard View (Always underlying if not active, but hidden/blurred for focus) */}
                <div className={`transition-all duration-500 ${activeSection !== 'dashboard' ? 'opacity-20 blur-sm pointer-events-none scale-95' : 'opacity-100 blur-0'}`}>
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-5 py-4">

                    {/* Whatsapp Section (External) */}
                    <a
                      href="https://wa.me/9647760055149"
                      target="_blank"
                      className="group bg-white p-6 rounded-[2rem] shadow-lg border border-gray-100 hover:border-[#1a5e1a]/30 transition-all hover:scale-[1.03] hover:shadow-2xl flex flex-col items-center text-center space-y-4"
                    >
                      <div className="w-16 h-16 bg-gradient-to-br from-[#25D366] to-[#128C7E] flex items-center justify-center text-white shadow-xl transition-all duration-300 group-hover:scale-110" style={{ borderRadius: '30% 70% 70% 30% / 30% 30% 70% 70%' }}>
                        <MessageCircle size={32} />
                      </div>
                      <div className="space-y-1">
                        <h3 className="font-black text-lg text-gray-800">تواصل مع الذاتية</h3>
                      </div>
                    </a>

                    {/* Section Cards */}
                    {sections.map(section => {
                      const userData = users.find(u => u.username === username);
                      const userPerms = userData?.permissions[section.id];
                      const isAdmin = userData?.role === 'admin';
                      if (!userPerms?.access && !isAdmin) return null;

                      const pendingCount = visibleRequests.filter(r => r.sectionId === section.id && (!r.status || r.status.includes('انتظار') || r.status === '')).length;

                      return (
                        <button
                          key={section.id}
                          onClick={() => setActiveSection(section.id)}
                          className="group bg-white p-6 rounded-[2rem] shadow-lg border border-gray-100 transition-all hover:scale-[1.03] hover:shadow-2xl flex flex-col items-center text-center space-y-4 relative overflow-hidden active:scale-95"
                        >
                          <div className="absolute -top-10 -right-10 w-24 h-24 bg-[#1a5e1a]/5 rounded-full group-hover:scale-150 transition-transform duration-700"></div>

                          {/* Professional Squircle Icon */}
                          <div
                            className={`w-16 h-16 ${section.id === 'mamalji' ? 'bg-gradient-to-br from-orange-400 to-red-600' :
                              section.id === 'hurr' ? 'bg-gradient-to-br from-blue-400 to-indigo-600' :
                                'bg-gradient-to-br from-[#1a5e1a] to-emerald-800'
                              } text-white flex items-center justify-center shadow-lg transition-all duration-300 group-hover:scale-110 group-hover:rotate-6`}
                            style={{ borderRadius: '30% 70% 70% 30% / 30% 30% 70% 70%' }}
                          >
                            <Layers size={32} />
                          </div>

                          <div className="space-y-1 relative z-10">
                            <h3 className="font-black text-[20px] text-gray-800 tracking-tight">{section.name}</h3>
                          </div>

                          {pendingCount > 0 && (
                            <div className="absolute top-4 left-4 w-8 h-8 bg-red-600 text-white rounded-xl flex items-center justify-center font-black text-xs border-2 border-white shadow-lg animate-bounce">
                              {pendingCount}
                            </div>
                          )}
                        </button>
                      )
                    })}

                    {/* Management & Settings Cards (Only for Admin) */}
                    {users.find(u => u.username === username)?.role === 'admin' && (
                      <>
                        <button
                          onClick={() => setActiveSection('users')}
                          className="group bg-white p-6 rounded-[2rem] shadow-lg border border-gray-100 transition-all hover:scale-[1.03] hover:shadow-2xl flex flex-col items-center text-center space-y-4"
                        >
                          <div className="w-16 h-16 bg-gradient-to-br from-purple-500 to-indigo-700 text-white flex items-center justify-center shadow-lg transition-all duration-300 group-hover:scale-110" style={{ borderRadius: '30% 70% 70% 30% / 30% 30% 70% 70%' }}>
                            <User size={32} />
                          </div>
                          <div className="space-y-2">
                            <h3 className="font-black text-[20px] text-gray-800">إدارة المستخدمين</h3>
                          </div>
                        </button>
                        <button
                          onClick={() => setActiveSection('settings')}
                          className="group bg-white p-6 rounded-[2rem] shadow-lg border border-gray-100 transition-all hover:scale-[1.03] hover:shadow-2xl flex flex-col items-center text-center space-y-4"
                        >
                          <div className="w-16 h-16 bg-gradient-to-br from-slate-600 to-slate-900 text-white flex items-center justify-center shadow-lg transition-all duration-300 group-hover:scale-110" style={{ borderRadius: '30% 70% 70% 30% / 30% 30% 70% 70%' }}>
                            <Layers size={32} />
                          </div>
                          <div className="space-y-2">
                            <h3 className="font-black text-[20px] text-gray-800">إعدادات النظام</h3>
                          </div>
                        </button>
                      </>
                    )}
                  </div>
                </div>

                {/* Sub-windows (Active Content) */}
                {activeSection !== 'dashboard' && (
                  <div className={`absolute inset-x-0 -top-4 z-40 animate-in zoom-in-95 fade-in duration-500 ${activeSectionTab === 'raise' ? 'hidden' : ''}`}>
                    <div className="flex items-center justify-between mb-6 bg-white/80 backdrop-blur-md p-4 rounded-3xl border border-white shadow-lg sticky top-0 md:top-4 z-50">
                      <div className="flex items-center gap-4">
                        <h2 className="text-xl font-black text-gray-800 pr-2">
                          {activeSection === 'users' ? 'إدارة المستخدمين' :
                            activeSection === 'settings' ? 'إعدادات النظام' :
                              `قسم ${sections.find(s => s.id === activeSection)?.name}`}
                        </h2>
                      </div>
                      <button
                        onClick={() => setActiveSection('dashboard')}
                        className="flex items-center gap-2 px-6 py-3 bg-red-50 hover:bg-red-100 text-red-600 rounded-2xl font-bold transition-all shadow-sm border border-red-100"
                      >
                        <span>رجوع للرئيسية</span>
                        <ArrowLeft size={18} className="rotate-180" />
                      </button>
                    </div>

                    <div className="bg-white/95 backdrop-blur-sm rounded-[3rem] shadow-2xl border-2 border-white p-6 md:p-10 mb-20 min-h-[60vh]">
                      {/* Content for different sub-sections */}
                      {activeSection === 'users' && users.find(u => u.username === username)?.role === 'admin' && (
                        <div className="space-y-8">
                          {/* User Section Navigation */}
                          <div className="flex flex-col sm:flex-row gap-4 p-2 bg-gray-100 rounded-[2rem] w-full sm:w-fit">
                            <button
                              onClick={() => setActiveUserSubSection('list')}
                              className={`flex-1 sm:flex-none px-8 py-3 rounded-2xl font-bold transition-all ${activeUserSubSection === 'list' ? 'bg-[#1a5e1a] text-white shadow-lg' : 'text-gray-600 hover:bg-gray-200'}`}
                            >
                              إعدادات المستخدمين
                            </button>
                            <button
                              onClick={() => openUserForm()}
                              className={`flex-1 sm:flex-none px-8 py-3 rounded-2xl font-bold transition-all ${activeUserSubSection === 'form' && !editingUser ? 'bg-[#1a5e1a] text-white shadow-lg' : 'text-gray-600 hover:bg-gray-200'}`}
                            >
                              إضافة مستخدم جديد
                            </button>
                          </div>

                          {activeUserSubSection === 'list' ? (
                            <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-gray-100">
                              {/* Desktop Table */}
                              <div className="hidden md:block overflow-x-auto">
                                <table className="w-full text-right border-collapse">
                                  <thead>
                                    <tr className="bg-gradient-to-l from-[#1a5e1a] to-[#2d7d2d] text-white">
                                      <th className="p-4 font-bold border-b border-white/10">ID</th>
                                      <th className="p-4 font-bold border-b border-white/10">الاسم</th>
                                      <th className="p-4 font-bold border-b border-white/10">اسم المستخدم</th>
                                      <th className="p-4 font-bold border-b border-white/10">كلمة السر</th>
                                      <th className="p-4 font-bold border-b border-white/10">الصلاحية (Role)</th>
                                      <th className="p-4 font-bold border-b border-white/10">القسم</th>
                                      <th className="p-4 font-bold border-b border-white/10">الأقسام المتاحة</th>
                                      <th className="p-4 font-bold border-b border-white/10">الإجراءات</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {users.map((user) => (
                                      <tr key={user.id} className="hover:bg-gray-50 transition-colors border-b border-gray-100">
                                        <td className="p-4 text-gray-500 font-mono text-sm">#{user.id}</td>
                                        <td className="p-4 font-bold text-gray-800">{user.name}</td>
                                        <td className="p-4 text-gray-600">{user.username}</td>
                                        <td className="p-4 font-mono text-gray-400">
                                          <span className="bg-gray-100 px-2 py-1 rounded">••••••</span>
                                        </td>
                                        <td className="p-4">
                                          <span className={`px-3 py-1 rounded-full text-[10px] font-black ${user.role === 'admin' ? 'bg-purple-100 text-purple-700' : user.role === 'manager' ? 'bg-blue-100 text-blue-700' : 'bg-gray-100 text-gray-600'}`}>
                                            {user.role === 'admin' ? 'مدير نظام' : user.role === 'manager' ? 'مسؤول قسم' : 'مستخدم'}
                                          </span>
                                        </td>
                                        <td className="p-4 text-gray-600">{user.department}</td>
                                        <td className="p-4">
                                          <div className="flex flex-wrap gap-1">
                                            {sections.map(s => user.permissions[s.id]?.access && (
                                              <span key={s.id} className="bg-green-50 text-green-700 px-2 py-0.5 rounded-full text-xs font-bold border border-green-200">
                                                {s.name}
                                              </span>
                                            ))}
                                          </div>
                                        </td>
                                        <td className="p-4">
                                          <div className="flex items-center gap-2">
                                            <button
                                              onClick={() => openUserForm(user)}
                                              className="w-10 h-10 rounded-full bg-blue-50 text-blue-600 flex items-center justify-center hover:bg-blue-600 hover:text-white transition-all shadow-sm"
                                            >
                                              <span className="text-sm">✏️</span>
                                            </button>
                                            <button
                                              onClick={() => handleDeleteUser(user.username, user.name)}
                                              className="w-10 h-10 rounded-full bg-red-50 text-red-600 flex items-center justify-center hover:bg-red-600 hover:text-white transition-all shadow-sm"
                                            >
                                              <span className="text-sm">🗑️</span>
                                            </button>
                                          </div>
                                        </td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>

                              {/* Mobile Card View */}
                              <div className="md:hidden divide-y divide-gray-100 text-right">
                                {users.map((user) => (
                                  <div key={user.id} className="p-6 space-y-4 hover:bg-gray-50 transition-colors">
                                    <div className="flex justify-between items-start">
                                      <div className="flex items-center gap-3">
                                        <div className="w-12 h-12 rounded-2xl bg-[#1a5e1a]/10 flex items-center justify-center text-[#1a5e1a] font-bold text-xl">
                                          {user.name.charAt(0)}
                                        </div>
                                        <div className="text-right">
                                          <h4 className="font-black text-gray-800">{user.name}</h4>
                                          <p className="text-xs text-gray-400 font-mono">#{user.id} | {user.username}</p>
                                        </div>
                                      </div>
                                      <div className="flex items-center gap-2">
                                        <button
                                          onClick={() => openUserForm(user)}
                                          className="w-10 h-10 bg-blue-100 text-blue-600 rounded-xl flex items-center justify-center shadow-lg"
                                        >
                                          ✏️
                                        </button>
                                        <button
                                          onClick={() => handleDeleteUser(user.username, user.name)}
                                          className="w-10 h-10 bg-red-100 text-red-600 rounded-xl flex items-center justify-center shadow-lg"
                                        >
                                          🗑️
                                        </button>
                                      </div>
                                    </div>
                                    <div className="grid grid-cols-2 gap-4">
                                      <div className="bg-gray-50 p-3 rounded-xl border border-white text-right">
                                        <p className="text-[8px] font-black text-gray-400 uppercase">الصلاحية</p>
                                        <p className="text-xs font-bold text-[#1a5e1a]">{user.role === 'admin' ? 'مدير نظام' : user.role === 'manager' ? 'مسؤول قسم' : 'مستخدم'}</p>
                                      </div>
                                      <div className="bg-gray-50 p-3 rounded-xl border border-white text-right">
                                        <p className="text-[8px] font-black text-gray-400 uppercase">القسم</p>
                                        <p className="text-xs font-bold text-gray-700 truncate">{user.department || 'غير محدد'}</p>
                                      </div>
                                    </div>
                                    <div className="space-y-1 text-right">
                                      <p className="text-[8px] font-black text-gray-400 uppercase">الأقسام المفعلة</p>
                                      <div className="flex flex-wrap gap-1">
                                        {sections.map(s => user.permissions[s.id]?.access && (
                                          <span key={s.id} className="bg-white border border-gray-100 text-[#1a5e1a] px-3 py-1 rounded-lg text-[9px] font-bold shadow-sm">
                                            {s.name}
                                          </span>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          ) : (
                            <form onSubmit={handleSaveUser} className="bg-white p-8 rounded-[2rem] shadow-xl border border-gray-100 space-y-10 animate-in fade-in zoom-in-95 duration-300">
                              <div className="flex justify-between items-center border-b border-gray-100 pb-4">
                                <h2 className="text-2xl font-black text-gray-800 flex items-center gap-3">
                                  <span className="bg-green-100 p-2 rounded-xl text-green-700">👤</span>
                                  {editingUser ? 'تعديل بيانات مستخدم' : 'إنشاء حساب مستخدم جديد'}
                                </h2>
                                <button
                                  type="button"
                                  onClick={() => setActiveUserSubSection('list')}
                                  className="text-gray-400 hover:text-red-500 transition-colors"
                                >
                                  إلغاء وإغلاق ✖️
                                </button>
                              </div>

                              {/* Basic Info Grid */}
                              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">اسم الحساب (الكامل)</label>
                                  <input
                                    required
                                    value={userForm.name || ''}
                                    onChange={e => setUserForm({ ...userForm, name: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all font-semibold"
                                    placeholder="مثال: علي محمد حسن"
                                  />
                                </div>
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">اسم المستخدم للولوج</label>
                                  <input
                                    required
                                    value={userForm.username || ''}
                                    onChange={e => setUserForm({ ...userForm, username: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all font-bold text-blue-600"
                                    placeholder="Example: ali_2024"
                                  />
                                </div>
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">كلمة السر</label>
                                  <input
                                    required
                                    type="password"
                                    value={userForm.password || ''}
                                    onChange={e => setUserForm({ ...userForm, password: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all font-mono"
                                  />
                                </div>
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">القسم الوظيفي</label>
                                  <input
                                    value={userForm.department || ''}
                                    onChange={e => setUserForm({ ...userForm, department: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all"
                                    placeholder="مثال: قسم التربية والتعليم"
                                  />
                                </div>
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">رقم الهاتف (اختياري)</label>
                                  <input
                                    value={userForm.phone || ''}
                                    onChange={e => setUserForm({ ...userForm, phone: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all"
                                  />
                                </div>
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">رابط واتساب (اختياري)</label>
                                  <input
                                    value={userForm.whatsapp || ''}
                                    onChange={e => setUserForm({ ...userForm, whatsapp: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all"
                                  />
                                </div>
                                <div className="space-y-2">
                                  <label className="text-sm font-bold text-gray-600 pr-2">نوع الحساب (Role)</label>
                                  <select
                                    value={userForm.role || 'user'}
                                    onChange={e => setUserForm({ ...userForm, role: e.target.value })}
                                    className="w-full px-4 py-3 rounded-xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:outline-none transition-all font-bold"
                                  >
                                    <option value="user">مستخدم عادي</option>
                                    <option value="manager">مسؤول قسم</option>
                                    <option value="admin">مدير نظام</option>
                                  </select>
                                </div>
                              </div>

                              {/* Permissions Matrix */}
                              <div className="space-y-6">
                                <div className="bg-gray-50 p-4 rounded-2xl border-r-4 border-orange-500">
                                  <h3 className="text-lg font-black text-gray-800">تحديد الصلاحيات المتقدمة للأقسام</h3>
                                  <p className="text-xs text-gray-500 font-bold">بإمكانك تحديد صلاحية كل موظف لكل قسم بشكل منفصل</p>
                                </div>

                                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                                  {sections.map(section => (
                                    <div key={section.id} className="bg-white p-6 rounded-3xl border border-gray-100 shadow-sm hover:shadow-md transition-all">
                                      <div className="flex items-center justify-between mb-4">
                                        <div className="flex items-center gap-3">
                                          <div className="w-10 h-10 rounded-2xl bg-[#1a5e1a]/10 flex items-center justify-center text-[#1a5e1a]">
                                            <Layers size={20} />
                                          </div>
                                          <span className="font-bold text-gray-800">{section.name}</span>
                                        </div>
                                        <label className="relative inline-flex items-center cursor-pointer">
                                          <input
                                            type="checkbox"
                                            className="sr-only peer"
                                            checked={userForm.permissions[section.id]?.access || false}
                                            onChange={() => togglePermission(section.id, 'access')}
                                          />
                                          <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:-translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:right-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-green-600"></div>
                                        </label>
                                      </div>

                                      {userForm.permissions[section.id]?.access && (
                                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 animate-in slide-in-from-top-2 duration-300">
                                          {[
                                            { id: 'raiseRequest', label: 'رفع طلب', icon: '📝' },
                                            { id: 'viewRequests', label: 'مشاهدة الطلبات والبحث', icon: '🔍' },
                                            { id: 'editWorkStatus', label: 'تعديل حالة العمل (الحالة)', icon: '⚡' },
                                            { id: 'editAssignedUnit', label: 'تعديل الوحدة المكلفة', icon: '🏢' },
                                            { id: 'editRequestDetails', label: 'تعديل الطلب (المواقع والنوع)', icon: '📍' },
                                            { id: 'editDeptNotes', label: 'تعديل ملاحظات الشعبة', icon: '📜' },
                                            { id: 'exportReports', label: 'تصدير Excel', icon: '📄' }
                                          ].map(p => (
                                            <label key={p.id} className="flex items-center gap-3 p-4 bg-white rounded-2xl border-2 border-gray-50 hover:border-blue-500 hover:shadow-lg cursor-pointer transition-all group">
                                              <div className="relative flex items-center justify-center">
                                                <input
                                                  type="checkbox"
                                                  className="w-5 h-5 rounded-lg accent-blue-600 transition-all cursor-pointer"
                                                  checked={userForm.permissions[section.id]?.[p.id] || false}
                                                  onChange={() => togglePermission(section.id, p.id)}
                                                />
                                              </div>
                                              <div className="flex flex-col">
                                                <span className="text-[12px] font-black text-gray-800 group-hover:text-blue-600 transition-colors flex items-center gap-2">
                                                  <span>{p.icon}</span>
                                                  {p.label}
                                                </span>
                                              </div>
                                            </label>
                                          ))}

                                          {section.sectionUnits && section.sectionUnits.length > 0 && (
                                            <div className="col-span-2 mt-4 pt-4 border-t border-gray-100">
                                              <h5 className="text-[11px] font-black text-gray-400 uppercase tracking-widest mb-3">صلاحيات الوحدات الفرعية</h5>
                                              <div className="flex flex-wrap gap-2">
                                                {section.sectionUnits.map(unit => (
                                                  <button
                                                    key={unit}
                                                    type="button"
                                                    onClick={() => toggleUnitPermission(section.id, unit)}
                                                    className={`px-4 py-2 rounded-xl text-xs font-bold transition-all border-2 ${userForm.permissions[section.id]?.units?.includes(unit) ? 'bg-orange-500 text-white border-orange-500 shadow-md translate-y-[-1px]' : 'bg-gray-50 text-gray-400 border-transparent hover:border-orange-200'}`}
                                                  >
                                                    {userForm.permissions[section.id]?.units?.includes(unit) && '⭐ '} {unit}
                                                  </button>
                                                ))}
                                              </div>
                                              <label className="flex items-center gap-2 mt-4 p-3 bg-orange-50 rounded-xl cursor-pointer">
                                                <input
                                                  type="checkbox"
                                                  className="w-4 h-4 accent-orange-600"
                                                  checked={userForm.permissions[section.id]?.isUnitLead || false}
                                                  onChange={() => togglePermission(section.id, 'isUnitLead')}
                                                />
                                                <span className="text-[10px] font-black text-orange-800">تفعيل صفة (مسؤول وحدة) فعلياً في الاحصائيات وإدارة الطلبات</span>
                                              </label>
                                            </div>
                                          )}
                                        </div>
                                      )}
                                    </div>
                                  ))}
                                </div>
                              </div>

                              <div className="flex justify-end pt-8 border-t border-gray-100 gap-4">
                                <button
                                  type="button"
                                  onClick={() => setActiveUserSubSection('list')}
                                  className="px-8 py-3 rounded-xl font-bold text-gray-500 hover:bg-gray-100 transition-colors"
                                >
                                  إلغاء
                                </button>
                                <button
                                  type="submit"
                                  className="px-12 py-3 bg-gradient-to-l from-[#1a5e1a] to-[#2d7d2d] text-white rounded-xl font-black shadow-lg hover:shadow-xl hover:-translate-y-1 transition-all"
                                >
                                  {editingUser ? 'حفظ التعديلات' : 'إضافة المستخدم للنظام 🔒'}
                                </button>
                              </div>
                            </form>
                          )}
                        </div>
                      )}

                      {activeSection === 'settings' && users.find(u => u.username === username)?.role === 'admin' && (
                        <div className="space-y-8 animate-in fade-in duration-500 text-right">
                          <div className="flex items-center gap-4 border-b border-gray-100 pb-6 mb-8">
                            <div className="w-14 h-14 bg-slate-800 text-white rounded-2xl flex items-center justify-center text-2xl shadow-lg">
                              ⚙️
                            </div>
                            <div>
                              <h2 className="text-2xl font-black text-gray-800">إعدادات النظام العامة</h2>
                              <p className="text-sm text-gray-500 font-bold">إدارة الأقسام، الأسئلة، وحالة استقبال الطلبات</p>
                            </div>
                          </div>

                          <div className="flex flex-col gap-10">
                            {/* Global Settings & Template Editor */}
                            <div className="bg-green-50 p-6 rounded-[2rem] border-2 border-green-100 shadow-sm space-y-4">
                              <div className="flex items-center justify-between">
                                <h3 className="text-lg font-black text-green-800 flex items-center gap-2">
                                  <MessageCircle size={24} className="text-green-600" />
                                  تخصيص نموذج رسالة WhatsApp
                                </h3>
                                <div className="flex flex-wrap gap-2">
                                  {['{sectionName}', '{unit}', '{submitter}', '{department}', '{details}', '{note}', '{status}'].map(tag => (
                                    <button
                                      key={tag}
                                      type="button"
                                      onClick={() => setWhatsappTemplate(prev => prev + tag)}
                                      className="px-2 py-1 bg-white border border-green-200 rounded-lg text-[10px] font-bold text-green-700 hover:bg-green-100 transition-colors"
                                    >
                                      {tag}
                                    </button>
                                  ))}
                                </div>
                              </div>
                              <p className="text-xs text-green-600 font-bold italic">قم بتعديل نص الرسالة أدناه. استخدم الأزرار أعلاه لإضافة حقول ديناميكية.</p>
                              <textarea
                                value={whatsappTemplate}
                                onChange={(e) => setWhatsappTemplate(e.target.value)}
                                rows={6}
                                className="w-full p-5 rounded-2xl border-2 border-white focus:border-green-500 focus:outline-none transition-all shadow-inner font-mono text-sm leading-relaxed"
                                placeholder="أدخل قالب الرسالة هنا..."
                              ></textarea>
                              <div className="text-[10px] text-gray-400 font-bold text-left">
                                Example: *New Maintenance Request* \n Location: {`{sectionName}`} ...
                              </div>
                            </div>

                            <div className="flex flex-col lg:flex-row gap-8">
                              {/* Sections Sidebar */}
                              <div className="w-full lg:w-[250px] space-y-3">
                                <p className="text-xs font-black text-gray-400 mb-4 px-2 uppercase tracking-widest">الأقسام المتاحة</p>
                                {sections.map(section => (
                                  <div key={section.id} className="relative group">
                                    <button
                                      onClick={() => setActiveSettingsSection(section.id)}
                                      className={`w-full p-4 rounded-2xl flex items-center justify-between transition-all border-2 ${activeSettingsSection === section.id ? 'border-[#1a5e1a] bg-green-50 shadow-md' : 'border-transparent bg-gray-50 hover:bg-gray-100'}`}
                                    >
                                      <span className={`font-bold ${activeSettingsSection === section.id ? 'text-[#1a5e1a]' : 'text-gray-600'}`}>{section.name}</span>
                                      {!section.isActive && <span className="text-[10px] bg-red-500 text-white px-2 py-0.5 rounded-full">إيقاف</span>}
                                    </button>
                                    <button
                                      onClick={() => openSectionModal(section)}
                                      className="absolute -left-2 top-1/2 -translate-y-1/2 p-2 bg-white shadow-md rounded-full text-[10px] opacity-0 group-hover:opacity-100 transition-all border border-gray-100 hover:text-blue-600"
                                      title="تعديل إعدادات القسم"
                                    >
                                      ⚙️
                                    </button>
                                  </div>
                                ))}
                                <button
                                  onClick={() => openSectionModal()}
                                  className="w-full p-4 rounded-2xl border-2 border-dashed border-gray-300 text-gray-400 font-bold hover:border-[#1a5e1a] hover:text-[#1a5e1a] transition-all flex items-center justify-center gap-2 mt-4"
                                >
                                  <span>➕ إضافة قسم جديد</span>
                                </button>
                              </div>

                              {/* Section Information Summary */}
                              <div className="flex-1 bg-white p-8 rounded-[2rem] border border-gray-100 space-y-10 shadow-sm relative overflow-hidden">
                                <div className="absolute top-0 right-0 w-2 h-full bg-[#1a5e1a]"></div>
                                {sections.map(section => section.id === activeSettingsSection && (
                                  <div key={section.id} className="space-y-8 animate-in slide-in-from-left-4 duration-400">
                                    <div className="flex items-center justify-between">
                                      <div>
                                        <h3 className="text-2xl font-black text-gray-800 flex items-center gap-2">
                                          {section.name}
                                          {!section.isActive && <span className="text-xs bg-red-100 text-red-600 px-3 py-1 rounded-full font-bold">متوقف مؤقتاً</span>}
                                        </h3>
                                        <p className="text-gray-500 font-bold text-sm mt-1">{section.description}</p>
                                      </div>
                                      <button
                                        onClick={() => openSectionModal(section)}
                                        className="px-6 py-3 bg-gray-100 text-gray-700 rounded-xl font-bold flex items-center gap-2 hover:bg-gray-200 transition-all"
                                      >
                                        ⚙️ تعديل إعدادات القسم والأسئلة
                                      </button>
                                    </div>

                                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                                      <div className="p-6 bg-gray-50 rounded-2xl border border-gray-100">
                                        <h4 className="text-[10px] uppercase tracking-widest font-black text-gray-400 mb-4 px-2">الوحدات الفنية المستخدمة</h4>
                                        <div className="flex flex-wrap gap-2">
                                          {section.sectionUnits.map(unit => (
                                            <span key={unit} className="px-3 py-1 bg-white rounded-lg text-xs font-bold text-gray-600 border border-gray-200">{unit}</span>
                                          ))}
                                        </div>
                                      </div>
                                      <div className="p-6 bg-gray-50 rounded-2xl border border-gray-100">
                                        <h4 className="text-[10px] uppercase tracking-widest font-black text-gray-400 mb-4 px-2">حالات العمل المبرمجة</h4>
                                        <div className="flex flex-wrap gap-2">
                                          {section.availableStatuses.map(status => (
                                            <span key={status} className="px-3 py-1 bg-white rounded-lg text-xs font-bold text-gray-600 border border-gray-200">{status}</span>
                                          ))}
                                        </div>
                                      </div>
                                    </div>

                                    <div className="space-y-4">
                                      <h4 className="text-[10px] uppercase tracking-widest font-black text-gray-400 px-2">هيكلية استمارة رفع الطلب</h4>
                                      <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                        {section.questions.map((q, idx) => (
                                          <div key={q.id} className="p-4 bg-white border-2 border-gray-50 rounded-xl flex items-center gap-4">
                                            <span className="w-8 h-8 bg-gray-100 rounded-lg flex items-center justify-center text-xs font-black text-gray-400 group-hover:bg-[#1a5e1a] group-hover:text-white transition-all">{idx + 1}</span>
                                            <div>
                                              <p className="text-sm font-black text-gray-800">{q.title} {q.required && <span className="text-red-500">*</span>}</p>
                                              <p className="text-[10px] font-bold text-gray-400">النوع: {q.type === 'text' ? 'نص قصير' : q.type === 'textarea' ? 'وصف طويل' : q.type === 'dropdown' ? 'قائمة منسدلة' : 'خيارات متعددة'}</p>
                                            </div>
                                          </div>
                                        ))}
                                      </div>
                                    </div>
                                  </div>
                                ))}
                              </div>
                            </div>
                          </div>
                        </div>
                      )}

                      {activeSection !== 'dashboard' && activeSection !== 'users' && activeSection !== 'settings' && (
                        <div className={`space-y-8 animate-in fade-in duration-500 text-right ${activeSectionTab === 'raise' ? 'hidden' : ''}`}>
                          {/* Check logic for simple users */}
                          {(() => {
                            const currentUserData = users.find(u => u.username === username);
                            const userSectionPerms = currentUserData?.permissions[activeSection];

                            if (!userSectionPerms?.access && currentUserData?.role !== 'admin') {
                              return (
                                <div className="text-center py-20 bg-white rounded-[2.5rem] shadow-xl border-2 border-red-50">
                                  <div className="text-6xl mb-6">🔒</div>
                                  <h2 className="text-2xl font-black text-gray-800">ليس لديك صلاحية الدخول لهذا القسم</h2>
                                  <p className="text-gray-500 mt-2">يرجى مراجعة مدير النظام لطلب إذن الوصول.</p>
                                  <button
                                    onClick={() => {
                                      const firstAllowed = sections.find(s => currentUserData?.permissions[s.id]?.access);
                                      setActiveSection(firstAllowed ? firstAllowed.id : 'dashboard');
                                    }}
                                    className="mt-8 px-10 py-3 bg-red-600 text-white rounded-xl font-bold"
                                  >
                                    العودة للأقسام المتاحة
                                  </button>
                                </div>
                              );
                            }

                            return (
                              <div className="section-content-wrapper">
                                {/* Section Sub-Navigation */}
                                <div className="flex justify-center md:justify-end items-center mb-8">
                                  <div className="flex flex-wrap gap-4 p-2 bg-gray-100 rounded-2xl w-fit">
                                    {(userSectionPerms?.viewRequests || currentUserData?.role === 'admin') && (
                                      <button
                                        onClick={() => setActiveSectionTab('list')}
                                        className={`px-8 py-2 rounded-xl font-bold transition-all ${activeSectionTab === 'list' ? 'bg-[#1a5e1a] text-white shadow-md' : 'text-gray-600 hover:bg-gray-200'}`}
                                      >
                                        الطلبات
                                      </button>
                                    )}
                                    {(userSectionPerms?.raiseRequest || currentUserData?.role === 'admin') && (
                                      <button
                                        onClick={() => setActiveSectionTab('raise')}
                                        className={`px-8 py-2 rounded-xl font-bold transition-all ${activeSectionTab === 'raise' ? 'bg-[#1a5e1a] text-white shadow-md' : 'text-gray-600 hover:bg-gray-200'}`}
                                      >
                                        رفع طلب
                                      </button>
                                    )}

                                    {currentUserData?.role === 'admin' && (
                                      <div className="flex gap-2 mr-2 border-r border-gray-300 pr-4">
                                        <button
                                          onClick={exportAllRequestsToExcel}
                                          className="px-5 py-2 rounded-xl font-bold bg-[#1a5e1a] text-white hover:bg-green-800 transition-all shadow-md flex items-center gap-2"
                                        >
                                          <span>تصدير</span>
                                          <Download size={18} />
                                        </button>
                                        <button
                                          onClick={async () => {
                                            const sectionName = activeSection === 'dashboard' ? 'جميع الأقسام' : (sections.find(s => s.id === activeSection)?.name || activeSection);
                                            const targetReqs = activeSection === 'dashboard' ? requests : requests.filter(r => r.sectionId === activeSection);

                                            if (targetReqs.length === 0) {
                                              alert('لا توجد طلبات لحذفها في هذا القسم حالياً.');
                                              return;
                                            }

                                            const confirmMsg = activeSection === 'dashboard'
                                              ? `⚠️ تحذير نهائي: أنت على وشك حذف جميع الطلبات في النظام (${targetReqs.length} طلب). هل أنت متأكد؟`
                                              : `⚠️ تحذير نهائي: أنت على وشك مسح جميع طلبات قسم "${sectionName}" نهائياً (عدد: ${targetReqs.length}). هل تريد الاستمرار؟`;

                                            if (window.confirm(confirmMsg)) {
                                              try {
                                                const batch = writeBatch(db);
                                                // Force deletion based on the serial 'id' as the Document Name
                                                targetReqs.forEach((r: any) => {
                                                  const finalDocId = r.id?.toString().trim();
                                                  if (finalDocId) {
                                                    batch.delete(doc(db, 'requests', finalDocId));
                                                  }
                                                });

                                                await batch.commit();
                                                alert(`✅ تم بنجاح تصفير القسم. تم حذف عدد (${targetReqs.length}) طلب.`);
                                              } catch (err: any) {
                                                console.error('Batch delete error:', err);
                                                alert(`❌ فشل في عملية الحذف الجماعي: ${err.message}`);
                                                handleFirestoreError(err, OperationType.DELETE, `requests/batch-by-id`);
                                              }
                                            }
                                          }}
                                          className="px-5 py-2 rounded-xl font-bold bg-red-600 text-white hover:bg-red-700 shadow-md flex items-center gap-2 transform active:scale-95 transition-all"
                                        >
                                          <Trash2 size={18} />
                                          <span className="uppercase tracking-tighter text-[11px] font-black underline decoration-white/30 underline-offset-4 bg-red-700/30 px-2 py-1 rounded-md">تصفير الطلبات</span>
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                </div>

                                {/* Raise Request Modal Content logic moved to separate sibling below main container to avoid hidden parent bug */}
                                <SectionStats
                                  requests={visibleRequests}
                                  activeSection={activeSection}
                                  userUnits={userSectionPerms?.units || []}
                                  userRole={currentUserData?.role || 'user'}
                                />

                                <div className="bg-white rounded-[2rem] shadow-xl overflow-hidden border border-gray-100 animate-in fade-in duration-500">
                                  <div className="p-6 bg-gray-50 border-b border-gray-100 flex flex-col xl:flex-row justify-between items-center gap-4">
                                    <h3 className="text-xl font-black text-gray-800 shrink-0">بيانات طلبات الصيانة - مجمع {sections.find(s => s.id === activeSection)?.name}</h3>
                                    <div className="flex flex-wrap gap-4 items-center justify-end w-full">
                                      {/* Search Input */}
                                      <div className="relative">
                                        <input
                                          className="pr-10 pl-12 py-2 rounded-xl border border-gray-200 text-xs focus:border-[#1a5e1a] focus:outline-none w-56 shadow-inner"
                                          placeholder="بحث شامل (اسم، حالة، تفاصيل)..."
                                          value={searchTerm}
                                          onChange={(e) => setSearchTerm(e.target.value)}
                                        />
                                        <div className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400">🔍</div>
                                        {searchTerm && (
                                          <div className="absolute left-3 top-1/2 -translate-y-1/2 px-2 py-0.5 bg-blue-50 text-blue-600 rounded-md text-[9px] font-bold border border-blue-100 animate-in zoom-in duration-300">
                                            {filteredRequests.length}
                                          </div>
                                        )}
                                      </div>

                                      {/* Status Filter */}
                                      <div className="flex items-center gap-2">
                                        <span className="text-[10px] font-black text-gray-400">الحالة:</span>
                                        <select
                                          value={filterStatus}
                                          onChange={(e) => setFilterStatus(e.target.value)}
                                          className="px-3 py-2 rounded-xl border border-gray-200 text-[10px] font-bold focus:border-[#1a5e1a] focus:outline-none bg-white shadow-sm"
                                        >
                                          <option value="الكل">الكل</option>
                                          <option value="قيد الانتظار">قيد الانتظار</option>
                                          {sections.find(s => s.id === activeSection)?.availableStatuses.map(s => (
                                            s !== 'قيد الانتظار' && <option key={s} value={s}>{s}</option>
                                          ))}
                                        </select>
                                      </div>

                                      {/* Date Filters */}
                                      <div className="flex items-center gap-2">
                                        <div className="flex items-center gap-1">
                                          <span className="text-[10px] font-black text-gray-400">من:</span>
                                          <input
                                            type="date"
                                            value={startDate}
                                            onChange={(e) => setStartDate(e.target.value)}
                                            className="px-2 py-1.5 rounded-lg border border-gray-200 text-[10px] font-bold focus:border-[#1a5e1a] focus:outline-none bg-white shadow-sm"
                                          />
                                        </div>
                                        <div className="flex items-center gap-1">
                                          <span className="text-[10px] font-black text-gray-400">إلى:</span>
                                          <input
                                            type="date"
                                            value={endDate}
                                            onChange={(e) => setEndDate(e.target.value)}
                                            className="px-2 py-1.5 rounded-lg border border-gray-200 text-[10px] font-bold focus:border-[#1a5e1a] focus:outline-none bg-white shadow-sm"
                                          />
                                        </div>
                                        {(startDate || endDate || filterStatus !== 'الكل') && (
                                          <button
                                            onClick={() => { setStartDate(''); setEndDate(''); setFilterStatus('الكل'); }}
                                            className="text-[10px] font-black text-red-500 underline hover:text-red-700 transition-colors"
                                          >
                                            إعادة تعيين الفلاتر
                                          </button>
                                        )}
                                      </div>

                                      <div className="h-8 w-px bg-gray-200 hidden xl:block"></div>

                                      {(userSectionPerms?.exportReports || currentUserData?.role === 'admin') && (
                                        <div className="flex gap-2">
                                          <button
                                            onClick={() => handleExportData('excel')}
                                            className="bg-green-600 text-white px-5 py-2 rounded-xl text-[10px] font-black shadow-md hover:bg-green-700 transition-all flex items-center gap-2"
                                          >
                                            <Download size={14} />
                                            تصدير Excel
                                          </button>
                                        </div>
                                      )}
                                    </div>
                                  </div>

                                  {/* Desktop View */}
                                  <div className="hidden md:block overflow-x-auto">
                                    <table className="w-full text-[13px] text-right border-collapse whitespace-nowrap">
                                      <thead>
                                        <tr className="bg-[#1a5e1a] text-white">
                                          <th className="p-4 font-black border-l border-white/10 uppercase tracking-tighter">ID</th>
                                          <th className="p-4 font-black border-l border-white/10">وقت الرفع</th>
                                          <th className="p-4 font-black border-l border-white/10">المجمع</th>
                                          <th className="p-4 font-black border-l border-white/10">مقدم الطلب</th>
                                          <th className="p-4 font-black border-l border-white/10">القسم</th>
                                          {/* Dynamic Question Columns */}
                                          {sections.find(s => s.id === activeSection)?.questions.map(q => (
                                            <th key={q.id} className="p-4 font-black border-l border-white/10">{q.title}</th>
                                          ))}
                                          <th className="p-4 font-black border-l border-white/10">الوحدة المكلفة</th>
                                          <th className="p-4 font-black border-l border-white/10">حالة العمل</th>
                                          <th className="p-4 font-black border-l border-white/10">وقت الإنجاز</th>
                                          <th className="p-4 font-black border-l border-white/10">ملاحظات الشعبة</th>
                                          <th className="p-4 font-black">آخر إرسال</th>
                                        </tr>
                                      </thead>
                                      <tbody>
                                        {getFilteredRequests().map(req => (
                                          <tr key={req.id} className={`hover:bg-gray-50 border-b border-gray-100 transition-colors ${!req.status || req.status === 'قيد الانتظار' ? 'bg-red-50/50' : ''}`}>
                                            <td className="p-4 font-black text-gray-500 border-l border-gray-50 group-hover:bg-green-50/30 transition-colors">
                                              <div className="flex items-center gap-2">
                                                <div className="w-1.5 h-1.5 rounded-full bg-[#1a5e1a] animate-pulse"></div>
                                                <span className="bg-gray-100 px-2 py-0.5 rounded-md shadow-sm border border-gray-200">#{req.id}</span>
                                              </div>
                                            </td>
                                            <td className="p-4 font-black text-gray-600">
                                              {req.timestamp.toLocaleDateString('ar-EG')} <br />
                                              <span className="text-[11px] opacity-70">{req.timestamp.toLocaleTimeString('ar-EG', { hour: '2-digit', minute: '2-digit' })}</span>
                                            </td>
                                            <td className="p-4 font-black text-[#1a5e1a]">
                                              {sections.find(s => s.id === req.sectionId)?.name}
                                            </td>
                                            <td className="p-4 font-black text-gray-900">{req.submitterName}</td>
                                            <td className="p-4 text-gray-700 font-extrabold">{req.submitterDept}</td>

                                            {/* Dynamic Answer Columns */}
                                            {sections.find(s => s.id === activeSection)?.questions.map(q => (
                                              <td key={q.id} className="p-4 max-w-[200px] overflow-hidden truncate font-bold text-gray-800">
                                                {req.answers[q.id] || '-'}
                                              </td>
                                            ))}

                                            <td className="p-4">
                                              <span className={`font-black px-4 py-1.5 rounded-full text-[12px] ${req.assignedUnit ? 'bg-blue-100 text-blue-700' : 'bg-red-100 text-red-700'}`}>
                                                {req.assignedUnit || 'بانتظار التوزيع'}
                                              </span>
                                            </td>
                                            <td className="p-4 whitespace-nowrap">
                                              <div className={`flex items-center gap-2 font-black text-[13px] ${req.status === 'منجز' ? 'text-green-600' : 'text-red-500 animate-pulse'}`}>
                                                {req.status ? (
                                                  <>
                                                    <div className="w-2.5 h-2.5 rounded-full bg-green-600"></div>
                                                    {req.status}
                                                  </>
                                                ) : (
                                                  <>
                                                    <div className="w-2.5 h-2.5 rounded-full bg-red-600"></div>
                                                    بانتظار المعالجة
                                                  </>
                                                )}
                                              </div>
                                            </td>
                                            <td className="p-4 font-black text-gray-600">
                                              {req.completionTime ? req.completionTime.toLocaleString('ar-EG') : '-'}
                                            </td>
                                            <td className="p-4 text-gray-500 text-[12px] font-bold italic max-w-[150px] truncate">
                                              {req.sectionNotes || 'لا توجد ملاحظات...'}
                                            </td>
                                            <td className="p-4">
                                              <div className="flex items-center gap-2">
                                                {(userSectionPerms?.editWorkStatus || userSectionPerms?.editAssignedUnit || userSectionPerms?.editDeptNotes || userSectionPerms?.editRequestDetails || currentUserData?.role === 'admin' || (userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(req.assignedUnit))) && (
                                                  <button
                                                    onClick={() => setEditingRequest(req)}
                                                    className="p-2.5 bg-blue-50 text-blue-600 rounded-xl hover:bg-blue-600 hover:text-white transition-all shadow-sm border border-blue-100 group"
                                                  >
                                                    {userSectionPerms?.viewRequests &&
                                                      currentUserData?.role !== 'admin' &&
                                                      !userSectionPerms?.editWorkStatus &&
                                                      !userSectionPerms?.editAssignedUnit &&
                                                      !userSectionPerms?.editDeptNotes &&
                                                      !userSectionPerms?.editRequestDetails &&
                                                      !(userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(req.assignedUnit))
                                                      ? <Eye size={18} />
                                                      : <span className="group-hover:scale-110 transition-transform block">✏️</span>}
                                                  </button>
                                                )}
                                                <div className="flex items-center gap-2">
                                                  <div className="flex flex-col items-end gap-0.5">
                                                    {req.lastSentRecipients && req.lastSentRecipients.length > 1 ? (
                                                      <div className="relative group cursor-help">
                                                        <span className="font-black text-gray-800 underline decoration-dotted decoration-green-500">مستلمون متعددون</span>
                                                        <div className="absolute bottom-full left-0 mb-2 hidden group-hover:block z-[150] bg-white border-2 border-green-500 p-3 rounded-2xl shadow-2xl min-w-[180px] animate-in fade-in zoom-in-95 duration-200">
                                                          <div className="text-[11px] font-black text-green-600 mb-2 border-b border-green-100 pb-1">قائمة المستلمين:</div>
                                                          {req.lastSentRecipients.map((name: string, i: number) => (
                                                            <div key={i} className="text-[12px] font-bold text-gray-700 py-1 border-b border-gray-50 last:border-0 flex items-center gap-2">
                                                              <span className="w-1.5 h-1.5 bg-green-500 rounded-full"></span>
                                                              {name}
                                                            </div>
                                                          ))}
                                                        </div>
                                                      </div>
                                                    ) : (
                                                      <span className="font-black text-gray-800 text-[13px]">
                                                        {req.lastSentRecipients?.[0] || req.lastSent?.split('-')[0]?.replace('whatsapp:', '')?.split('(')[0]?.trim() || 'لم يرسل'}
                                                      </span>
                                                    )}
                                                    {(req.lastSentDate || req.lastSent) && (
                                                      <span className="text-[11px] text-gray-400 font-bold italic">
                                                        {req.lastSentDate
                                                          ? req.lastSentDate.toLocaleDateString('ar-EG')
                                                          : (req.lastSent?.includes('(') ? req.lastSent.split('(')[1].split(')')[0] : (req.lastSent?.split('-').pop()?.trim() || ''))}
                                                      </span>
                                                    )}
                                                  </div>
                                                  <button
                                                    onClick={() => setWhatsappModal({ isOpen: true, request: req })}
                                                    className="text-green-600 hover:scale-125 transition-all text-lg p-1 hover:bg-green-50 rounded-lg"
                                                    title="إرسال إشعار جديد"
                                                  >
                                                    📲
                                                  </button>
                                                </div>
                                              </div>
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>

                                  {/* Mobile View */}
                                  <div className="md:hidden divide-y divide-gray-100 text-right">
                                    {getFilteredRequests().map(req => (
                                      <div key={req.id} className={`p-6 space-y-4 hover:bg-gray-50 transition-colors ${!req.status || req.status === 'قيد الانتظار' ? 'bg-red-50/30' : ''}`}>
                                        <div className="flex justify-between items-start">
                                          <div className="flex items-center gap-3">
                                            <div className={`w-12 h-12 rounded-2xl flex items-center justify-center font-black ${req.status === 'منجز' ? 'bg-green-100 text-green-600' : 'bg-red-100 text-red-600'}`}>
                                              {req.status === 'منجز' ? '✓' : '!'}
                                            </div>
                                            <div className="text-right">
                                              <h4 className="font-black text-gray-900 text-lg">{req.submitterName}</h4>
                                              <p className="text-[12px] text-gray-500 font-black">#{req.id} | {req.timestamp.toLocaleDateString('ar-EG')}</p>
                                            </div>
                                          </div>
                                          <div className="flex gap-2">
                                            {(userSectionPerms?.editWorkStatus || userSectionPerms?.editAssignedUnit || userSectionPerms?.editDeptNotes || userSectionPerms?.editRequestDetails || currentUserData?.role === 'admin' || (userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(req.assignedUnit))) && (
                                              <button
                                                onClick={() => setEditingRequest(req)}
                                                className="p-2.5 bg-blue-50 text-blue-600 rounded-xl hover:bg-blue-600 hover:text-white transition-all shadow-sm border border-blue-100 group"
                                              >
                                                {userSectionPerms?.viewRequests &&
                                                  currentUserData?.role !== 'admin' &&
                                                  !userSectionPerms?.editWorkStatus &&
                                                  !userSectionPerms?.editAssignedUnit &&
                                                  !userSectionPerms?.editDeptNotes &&
                                                  !userSectionPerms?.editRequestDetails &&
                                                  !(userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(req.assignedUnit))
                                                  ? <Eye size={18} />
                                                  : <span className="group-hover:scale-110 transition-transform block">✏️</span>}
                                              </button>
                                            )}
                                            <button
                                              onClick={() => setWhatsappModal({ isOpen: true, request: req })}
                                              className="text-green-600 hover:scale-125 transition-all text-lg p-1 hover:bg-green-50 rounded-lg"
                                            >
                                              📲
                                            </button>
                                          </div>
                                        </div>

                                        <div className="grid grid-cols-2 gap-3 text-right">
                                          <div className="bg-white p-3 rounded-xl border border-gray-100">
                                            <p className="text-[10px] text-gray-400 font-black mb-1">الحالة</p>
                                            <span className={`text-xs font-black ${req.status === 'منجز' ? 'text-green-600' : 'text-red-500'}`}>{req.status || 'قيد الانتظار'}</span>
                                          </div>
                                          <div className="bg-white p-3 rounded-xl border border-gray-100">
                                            <p className="text-[10px] text-gray-400 font-black mb-1">المكلف</p>
                                            <span className="text-xs font-black text-blue-700">{req.assignedUnit || 'غير محدد'}</span>
                                          </div>
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              </div>
                            );
                          })()}
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          </main>
        </div>

        <AnimatePresence>
          {activeSectionTab === 'raise' && (users.find(u => u.username === username)?.permissions[activeSection]?.raiseRequest || users.find(u => u.username === username)?.role === 'admin') && (
            <div className="fixed inset-0 z-[500] overflow-y-auto bg-slate-900/60 backdrop-blur-xl flex items-start justify-center p-3 sm:p-4 md:p-8 rtl">
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                className="absolute inset-0 z-0"
                onClick={() => setActiveSectionTab('list')}
              ></motion.div>

              <motion.div
                initial={{ scale: 0.9, opacity: 0, y: 40 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.9, opacity: 0, y: 40 }}
                className="bg-white/95 backdrop-blur-sm rounded-[2rem] md:rounded-[3rem] shadow-2xl border border-white/40 w-full max-w-2xl overflow-hidden relative z-[510] flex flex-col my-auto"
              >
                {/* Modal Header */}
                <div className="px-6 py-6 md:px-10 md:py-8 border-b border-gray-100/50 flex items-center justify-between bg-white/50 backdrop-blur-md relative z-20">
                  <div className="flex items-center gap-4 text-right">
                    <div className="w-12 h-12 md:w-14 md:h-14 bg-[#1a5e1a]/10 text-[#1a5e1a] rounded-2xl flex items-center justify-center shadow-inner border border-[#1a5e1a]/10">
                      <Pen size={24} className="opacity-80" />
                    </div>
                    <div>
                      <h2 className="text-xl md:text-2xl font-black text-gray-900 tracking-tight leading-tight">
                        رفع طلب صيانة جديد
                      </h2>
                      <p className="text-xs md:text-sm font-bold text-[#1a5e1a] flex items-center gap-2 mt-1 opacity-80">
                        <span className="w-2 h-2 bg-[#1a5e1a] rounded-full animate-pulse"></span>
                        مجمع {sections.find(s => s.id === activeSection)?.name}
                      </p>
                    </div>
                  </div>

                  <button
                    onClick={() => setActiveSectionTab('list')}
                    className="w-10 h-10 md:w-12 md:h-12 bg-gray-100 hover:bg-red-50 text-gray-400 hover:text-red-500 rounded-2xl flex items-center justify-center transition-all border border-gray-200/50 group"
                  >
                    <X size={24} className="group-hover:rotate-90 transition-transform" />
                  </button>
                </div>

                {/* Modal Body */}
                <div className="flex-1 overflow-y-auto scrollbar-hide p-6 md:p-10 max-h-[70vh]">
                  {!sections.find(s => s.id === activeSection)?.isActive ? (
                    <div className="text-center space-y-6 py-12 bg-red-50/30 rounded-3xl border border-red-100/20">
                      <div className="w-20 h-20 bg-white text-red-500 rounded-3xl flex items-center justify-center mx-auto text-3xl shadow-xl border border-red-50">⚠️</div>
                      <div className="space-y-2">
                        <h3 className="text-2xl font-black text-red-700">هذا القسم متوقف حالياً</h3>
                        <p className="text-red-600/70 text-base font-bold">يقوم مدير النظام حالياً ببعض أعمال الصيانة، يرجى المحاولة لاحقاً.</p>
                      </div>
                      <button
                        onClick={() => setActiveSectionTab('list')}
                        className="px-10 py-4 bg-red-600 text-white rounded-2xl font-black shadow-lg hover:bg-red-700 transition-all active:scale-95"
                      >
                        الرجوع للطلبات
                      </button>
                    </div>
                  ) : (
                    <form className="space-y-6 md:space-y-8 pb-4" onSubmit={handleRaiseRequest}>
                      <div className="space-y-6">
                        {sections.find(s => s.id === activeSection)?.questions.map(q => (
                          <div key={q.id} className="space-y-3 group">
                            <label className="text-base md:text-lg font-black text-gray-800 pr-1 flex items-center gap-2 group-focus-within:text-[#1a5e1a] transition-colors">
                              <span className="w-2.5 h-2.5 bg-[#1a5e1a]/30 rounded-full group-focus-within:bg-[#1a5e1a] transition-all"></span>
                              {q.title} {q.required && <span className="text-red-500">*</span>}
                            </label>

                            {q.type === 'text' && (
                              <input
                                name={q.id}
                                required={q.required}
                                placeholder={`أدخل ${q.title}...`}
                                className="w-full px-5 py-4 md:px-6 md:py-5 rounded-2xl border-2 border-gray-100 bg-gray-50/30 focus:border-[#1a5e1a] focus:bg-white focus:ring-8 focus:ring-[#1a5e1a]/5 focus:outline-none transition-all font-bold text-gray-900 shadow-sm placeholder:text-gray-300 text-base md:text-lg"
                              />
                            )}

                            {q.type === 'textarea' && (
                              <textarea
                                name={q.id}
                                required={q.required}
                                rows={4}
                                placeholder={`اكتب تفاصيل ${q.title} هنا بوضوح لتسهيل عملية الصيانة...`}
                                className="w-full px-5 py-4 md:px-6 md:py-5 rounded-2xl border-2 border-gray-100 bg-gray-50/30 focus:border-[#1a5e1a] focus:bg-white focus:ring-8 focus:ring-[#1a5e1a]/5 focus:outline-none transition-all font-bold text-gray-900 shadow-sm resize-none placeholder:text-gray-300 text-base md:text-lg"
                              ></textarea>
                            )}

                            {q.type === 'dropdown' && (
                              <div className="relative">
                                <select
                                  name={q.id}
                                  required={q.required}
                                  className="w-full px-5 py-4 md:px-6 md:py-5 rounded-2xl border-2 border-gray-100 bg-gray-50/30 focus:border-[#1a5e1a] focus:bg-white focus:ring-8 focus:ring-[#1a5e1a]/5 focus:outline-none appearance-none cursor-pointer font-bold text-gray-900 shadow-sm text-base md:text-lg"
                                >
                                  <option value="">إختر من القائمة المتاحة...</option>
                                  {q.options?.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                </select>
                                <div className="absolute left-6 top-1/2 -translate-y-1/2 pointer-events-none text-gray-400">
                                  <ArrowLeft size={20} className="-rotate-90 opacity-40" />
                                </div>
                              </div>
                            )}

                            {q.type === 'checkbox' && (
                              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 p-2">
                                {q.options?.map(opt => (
                                  <label key={opt} className="relative flex items-center p-4 rounded-2xl border-2 border-gray-100 hover:border-[#1a5e1a]/30 cursor-pointer transition-all bg-white group/opt has-[:checked]:border-[#1a5e1a] has-[:checked]:bg-[#1a5e1a]/5 shadow-sm">
                                    <input
                                      name={q.id}
                                      value={opt}
                                      type="checkbox"
                                      className="peer w-6 h-6 opacity-0 absolute"
                                    />
                                    <div className="w-6 h-6 border-2 border-gray-200 rounded-lg peer-checked:bg-[#1a5e1a] peer-checked:border-[#1a5e1a] transition-all flex items-center justify-center shrink-0">
                                      <div className="w-2 h-2 bg-white rounded-full opacity-0 peer-checked:opacity-100 transition-opacity"></div>
                                    </div>
                                    <span className="font-bold text-base text-gray-700 pr-4 peer-checked:text-[#1a5e1a] transition-colors">{opt}</span>
                                  </label>
                                ))}
                              </div>
                            )}
                          </div>
                        ))}
                      </div>

                      <div className="flex flex-col sm:flex-row justify-end gap-4 pt-8 border-t border-gray-100/50">
                        <button
                          type="button"
                          onClick={() => setActiveSectionTab('list')}
                          className="px-8 py-4 rounded-2xl font-black text-gray-400 hover:bg-gray-100 hover:text-gray-600 transition-all text-base order-2 sm:order-1"
                        >
                          إغلاق النافذة
                        </button>
                        <button
                          type="submit"
                          className="px-12 py-4 bg-[#1a5e1a] hover:bg-[#2d7d2d] text-white rounded-2xl font-black shadow-xl shadow-[#1a5e1a]/20 hover:shadow-[#1a5e1a]/30 hover:-translate-y-1 active:scale-95 transition-all text-lg order-1 sm:order-2 flex items-center justify-center gap-3"
                        >
                          <span>إرسال الطلب الآن</span>
                          <ArrowLeft size={22} className="rotate-180" />
                        </button>
                      </div>
                    </form>
                  )}
                </div>
              </motion.div>
            </div>
          )}
        </AnimatePresence>

        {/* Section Configuration Modal */}
        {isSectionModalOpen && (
          <div className="fixed inset-0 bg-black/80 backdrop-blur-xl z-[200] flex items-center justify-center p-4 rtl">
            <form onSubmit={saveSectionConfig} className="bg-white rounded-[2.5rem] shadow-2xl w-full max-w-4xl overflow-hidden animate-in zoom-in-95 duration-300 border-4 border-white flex flex-col max-h-[90vh]">
              <div className="p-8 bg-gradient-to-l from-slate-800 to-slate-900 text-white flex justify-between items-center relative">
                <button type="button" onClick={() => setIsSectionModalOpen(false)} className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center hover:bg-white/20 hover:scale-110 transition-all border border-white/10">✕</button>
                <div className="flex items-center gap-4">
                  <div className="w-14 h-14 bg-white/10 rounded-2xl flex items-center justify-center text-2xl shadow-inner border border-white/10">
                    ⚙️
                  </div>
                  <div className="text-right">
                    <h2 className="text-2xl font-black">{editingSectionConfig ? `إعدادات ${editingSectionConfig.name}` : 'إضافة قسم جديد للنظام'}</h2>
                    <p className="text-xs opacity-70 font-bold">تحكم في الحقول، الوحدات، وحالة الاستقبال</p>
                  </div>
                </div>
              </div>

              <div className="p-10 space-y-10 overflow-y-auto custom-scrollbar flex-1 text-right">
                {/* General Info */}
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                  <div className="space-y-2">
                    <label className="text-sm font-bold text-gray-600 pr-2">اسم القسم</label>
                    <input
                      required
                      value={sectionForm.name}
                      onChange={e => setSectionForm({ ...sectionForm, name: e.target.value })}
                      className="w-full px-5 py-4 rounded-2xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:bg-white bg-gray-50 font-black text-gray-800 transition-all"
                      placeholder="مثال: المعملجي"
                    />
                  </div>
                  <div className="space-y-2 lg:col-span-2">
                    <label className="text-sm font-bold text-gray-600 pr-2">وصف القسم (اختياري)</label>
                    <input
                      value={sectionForm.description}
                      onChange={e => setSectionForm({ ...sectionForm, description: e.target.value })}
                      className="w-full px-5 py-4 rounded-2xl border-2 border-gray-100 focus:border-[#1a5e1a] focus:bg-white bg-gray-50 font-bold text-gray-800 transition-all"
                      placeholder="وصف موجز لوظيفة هذا القسم في النظام..."
                    />
                  </div>
                  <div className="space-y-2">
                    <label className="text-sm font-bold text-gray-600 pr-2">حالة الاستقبال</label>
                    <div className="flex items-center justify-between bg-gray-50 p-4 rounded-2xl border-2 border-gray-100">
                      <span className={`text-sm font-black ${sectionForm.isActive ? 'text-green-600' : 'text-red-500'}`}>
                        {sectionForm.isActive ? '✅ يستقبل طلبات' : '🛑 متوقف مؤقتاً'}
                      </span>
                      <label className="relative inline-flex items-center cursor-pointer">
                        <input
                          type="checkbox"
                          className="sr-only peer"
                          checked={sectionForm.isActive}
                          onChange={() => setSectionForm({ ...sectionForm, isActive: !sectionForm.isActive })}
                        />
                        <div className="w-11 h-6 bg-gray-200 peer-focus:outline-none rounded-full peer peer-checked:after:-translate-x-full peer-checked:after:border-white after:content-[''] after:absolute after:top-[2px] after:right-[2px] after:bg-white after:border-gray-300 after:border after:rounded-full after:h-5 after:w-5 after:transition-all peer-checked:bg-green-600"></div>
                      </label>
                    </div>
                  </div>
                  <TagInput
                    label="الوحدات الفنية"
                    tags={sectionForm.sectionUnits}
                    onAdd={tag => setSectionForm({ ...sectionForm, sectionUnits: [...sectionForm.sectionUnits, tag] })}
                    onRemove={idx => setSectionForm({ ...sectionForm, sectionUnits: sectionForm.sectionUnits.filter((_, i) => i !== idx) })}
                    placeholder="اكتب اسم الوحدة واضغط Enter..."
                  />
                  <TagInput
                    label="حالات العمل المبرمجة"
                    tags={sectionForm.availableStatuses}
                    onAdd={tag => setSectionForm({ ...sectionForm, availableStatuses: [...sectionForm.availableStatuses, tag] })}
                    onRemove={idx => setSectionForm({ ...sectionForm, availableStatuses: sectionForm.availableStatuses.filter((_, i) => i !== idx) })}
                    placeholder="أضف حالة جديدة (مثل: مكتمل)..."
                  />
                </div>

                {/* Questions Form Builder */}
                <div className="space-y-6">
                  <div className="flex items-center justify-between border-b-2 border-gray-100 pb-4">
                    <h3 className="text-lg font-black text-slate-800 flex items-center gap-3">
                      <span className="w-8 h-8 bg-slate-800 text-white rounded-lg flex items-center justify-center text-sm">📋</span>
                      برمجة أسئلة الاستمارة
                    </h3>
                    <button
                      type="button"
                      onClick={() => setSectionForm({
                        ...sectionForm,
                        questions: [...sectionForm.questions, { id: Math.random().toString(36).substr(2, 9), title: '', type: 'text', required: false, options: [] }]
                      })}
                      className="px-6 py-2 bg-blue-600 text-white rounded-xl font-black text-xs shadow-lg hover:scale-105 transition-all"
                    >
                      ➕ إضافة حقل جديد
                    </button>
                  </div>

                  <div className="grid grid-cols-1 gap-4">
                    {sectionForm.questions.map((q, idx) => (
                      <div key={q.id} className="p-6 bg-slate-50 rounded-[2rem] border-2 border-white shadow-sm hover:border-blue-200 transition-all group">
                        <div className="flex flex-col lg:flex-row gap-6">
                          <div className="flex items-center gap-4 flex-[2]">
                            <span className="font-black text-slate-300 group-hover:text-blue-500 transition-colors">#{idx + 1}</span>
                            <div className="space-y-1 w-full">
                              <label className="text-[13px] font-black text-gray-500 pr-2 uppercase tracking-wide">عنوان الحقل / السؤال</label>
                              <input
                                required
                                value={q.title}
                                onChange={e => {
                                  const newQs = [...sectionForm.questions];
                                  newQs[idx].title = e.target.value;
                                  setSectionForm({ ...sectionForm, questions: newQs });
                                }}
                                className="w-full px-5 py-3 rounded-xl border-2 border-gray-100 focus:border-blue-500 bg-white font-black text-gray-800 transition-all"
                                placeholder="مثال: البناية، رقم الغرفة، وصف العطل..."
                              />
                            </div>
                          </div>

                          <div className="flex items-center gap-4 flex-1">
                            <div className="space-y-1 w-full">
                              <label className="text-[13px] font-black text-gray-500 pr-2 uppercase tracking-wide">نوع المدخل</label>
                              <select
                                value={q.type}
                                onChange={e => {
                                  const newQs = [...sectionForm.questions];
                                  newQs[idx].type = e.target.value;
                                  setSectionForm({ ...sectionForm, questions: newQs });
                                }}
                                className="w-full px-5 py-3 rounded-xl border-2 border-gray-100 focus:border-blue-500 bg-white font-bold text-gray-700 transition-all"
                              >
                                <option value="text">نص قصير</option>
                                <option value="textarea">نص ممتد (وصف)</option>
                                <option value="dropdown">قائمة خيارات</option>
                                <option value="checkbox">مربعات اختيار</option>
                              </select>
                            </div>
                          </div>

                          <div className="flex items-center gap-6">
                            <label className="flex items-center gap-2 cursor-pointer mt-4">
                              <input
                                type="checkbox"
                                checked={q.required}
                                onChange={e => {
                                  const newQs = [...sectionForm.questions];
                                  newQs[idx].required = e.target.checked;
                                  setSectionForm({ ...sectionForm, questions: newQs });
                                }}
                                className="w-5 h-5 accent-red-500"
                              />
                              <span className="text-xs font-black text-red-600">إلزامي</span>
                            </label>
                            <button
                              type="button"
                              onClick={() => {
                                setSectionForm({
                                  ...sectionForm,
                                  questions: sectionForm.questions.filter((_, i) => i !== idx)
                                });
                              }}
                              className="mt-4 p-3 bg-red-50 text-red-600 rounded-xl hover:bg-red-600 hover:text-white transition-all shadow-sm"
                              title="حذف هذا السؤال"
                            >
                              🗑️
                            </button>
                          </div>
                        </div>

                        {(q.type === 'dropdown' || q.type === 'checkbox') && (
                          <div className="mt-4 pt-4 border-t border-gray-100 animate-in slide-in-from-top-2">
                            <TagInput
                              label="الخيارات المتاحة لهذا السؤال"
                              tags={q.options || []}
                              onAdd={tag => {
                                const newQs = [...sectionForm.questions];
                                newQs[idx].options = [...(newQs[idx].options || []), tag];
                                setSectionForm({ ...sectionForm, questions: newQs });
                              }}
                              onRemove={optIdx => {
                                const newQs = [...sectionForm.questions];
                                newQs[idx].options = newQs[idx].options.filter((_: any, i: number) => i !== optIdx);
                                setSectionForm({ ...sectionForm, questions: newQs });
                              }}
                              placeholder="أضف خياراً واضغط Enter..."
                            />
                          </div>
                        )}
                      </div>
                    ))}
                    {sectionForm.questions.length === 0 && (
                      <div className="text-center py-10 bg-gray-50 rounded-[2rem] border-2 border-dashed border-gray-200">
                        <p className="text-gray-400 font-bold italic">لا توجد أسئلة مبرمجة حالياً. يرجى إضافة سؤال واحد على الأقل.</p>
                      </div>
                    )}
                  </div>
                </div>
              </div>

              <div className="p-8 bg-gray-50 border-t border-gray-100 flex flex-col md:flex-row gap-4 items-center justify-between">
                <div className="flex gap-4 w-full md:w-auto">
                  <button
                    type="button"
                    onClick={() => setIsSectionModalOpen(false)}
                    className="px-10 py-4 bg-white text-gray-500 rounded-2xl font-black border-2 border-gray-200 hover:bg-gray-100 transition-all flex-1 md:flex-none"
                  >إلغاء</button>
                  {editingSectionConfig && (
                    <button
                      type="button"
                      onClick={() => deleteSection(editingSectionConfig.id)}
                      className="px-10 py-4 bg-red-50 text-red-600 rounded-2xl font-black border-2 border-red-100 hover:bg-red-600 hover:text-white transition-all flex items-center gap-2 flex-1 md:flex-none justify-center"
                    >🗑️ حذف القسم</button>
                  )}
                </div>
                <button
                  type="submit"
                  disabled={sectionForm.questions.length === 0}
                  className={`px-16 py-4 rounded-2xl font-black shadow-xl transition-all w-full md:w-auto text-white ${sectionForm.questions.length > 0 ? 'bg-gradient-to-l from-[#1a5e1a] to-[#2d7d2d] hover:scale-105' : 'bg-gray-400 cursor-not-allowed'}`}
                >
                  {editingSectionConfig ? 'حفظ كافة التعديلات 💾' : 'تأكيد إضافة القسم الجديد 🚀'}
                </button>
              </div>
            </form>
          </div>
        )}
        {editingRequest && (() => {
          const currentUserData = users.find(u => u.username === username);
          const userSectionPerms = currentUserData?.permissions[editingRequest.sectionId];
          const isAdmin = currentUserData?.role === 'admin';

          return (
            <div className="fixed inset-0 z-[200] flex items-center justify-center p-4 sm:p-6 md:p-10 rtl">
              <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="absolute inset-0 bg-slate-900/45 backdrop-blur-md transition-opacity"
                onClick={() => setEditingRequest(null)}
              ></motion.div>

              <motion.form
                onSubmit={updateRequest}
                initial={{ scale: 0.95, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                className="bg-[#fcfaf7] rounded-[2.5rem] shadow-[0_50px_100px_-20px_rgba(0,0,0,0.3)] border border-white/60 w-full max-w-2xl max-h-[95vh] overflow-hidden relative z-[210] flex flex-col"
              >
                <div className="p-6 md:p-8 bg-white border-b border-gray-100 flex justify-between items-center relative z-20">
                  <div className="text-right flex items-center gap-4">
                    <div className="w-12 h-12 bg-blue-50 text-blue-700 rounded-2xl flex items-center justify-center shadow-inner border border-blue-100/50">
                      <Activity size={24} />
                    </div>
                    <div>
                      <h2 className="text-xl font-black text-gray-900 tracking-tight">تحديث الطلب #{editingRequest.id}</h2>
                      <p className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mt-0.5">إدارة الحالة والتعليقات الفنية</p>
                    </div>
                  </div>
                  <button
                    type="button"
                    onClick={() => setEditingRequest(null)}
                    className="w-10 h-10 rounded-2xl bg-gray-50 hover:bg-red-50 text-gray-400 hover:text-red-500 flex items-center justify-center transition-all border border-gray-100/50"
                  >
                    <X size={20} />
                  </button>
                </div>

                <div className="flex-1 p-6 md:p-10 space-y-10 text-right overflow-y-auto scrollbar-hide">

                  {/* Section: Request Basic Details */}
                  {(isAdmin || userSectionPerms?.editRequestDetails) && (
                    <div className="space-y-6 p-6 md:p-8 bg-blue-50/40 rounded-[2rem] border border-blue-100/60 shadow-sm relative overflow-hidden group">
                      <div className="absolute top-0 right-0 w-24 h-24 bg-blue-600/5 rounded-full -mr-12 -mt-12 transition-transform group-hover:scale-110"></div>
                      <div className="relative flex items-center gap-3 mb-2">
                        <div className="p-2 bg-blue-600 text-white rounded-xl shadow-md">
                          <Pen size={14} />
                        </div>
                        <h3 className="text-sm font-black text-blue-900 uppercase tracking-widest">تعديل البيانات الأساسية</h3>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 text-right relative">
                        {sections.find(s => s.id === editingRequest.sectionId)?.questions.map(q => (
                          <div key={q.id} className="space-y-2">
                            <label className="text-[13px] font-black text-blue-900/60 pr-2 uppercase tracking-tighter">{q.title}</label>
                            <input
                              type="text"
                              value={editingRequest.answers[q.id] || ''}
                              onChange={e => setEditingRequest({
                                ...editingRequest,
                                answers: { ...editingRequest.answers, [q.id]: e.target.value }
                              })}
                              className="w-full px-5 py-3.5 bg-white rounded-xl border border-blue-100 focus:border-blue-500 focus:ring-4 focus:ring-blue-500/5 font-bold text-sm shadow-sm transition-all outline-none"
                            />
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-8">
                    <div className="space-y-3">
                      <label className="text-[15px] font-black text-gray-600 pr-2 uppercase tracking-widest flex items-center gap-2">
                        <Building size={14} className="text-[#1a5e1a]/60" />
                        الوحدة المكلفة
                      </label>
                      <div className="relative group">
                        <select
                          value={editingRequest.assignedUnit || ''}
                          onChange={e => setEditingRequest({ ...editingRequest, assignedUnit: e.target.value })}
                          disabled={!(isAdmin || userSectionPerms?.editAssignedUnit)}
                          className={`w-full px-5 py-4 rounded-2xl border-2 border-gray-100 bg-white focus:border-[#1a5e1a] focus:ring-8 focus:ring-[#1a5e1a]/5 font-black text-gray-800 transition-all appearance-none cursor-pointer ${!(isAdmin || userSectionPerms?.editAssignedUnit) ? 'cursor-not-allowed opacity-60 bg-gray-50' : ''}`}
                        >
                          <option value="">بانتظار التوزيع</option>
                          {(sections.find(s => s.id === editingRequest.sectionId)?.sectionUnits || globalUnits).map(unit => (
                            <option key={unit} value={unit}>{unit}</option>
                          ))}
                        </select>
                        <div className="absolute left-6 top-1/2 -translate-y-1/2 pointer-events-none text-gray-400 group-hover:text-[#1a5e1a] transition-colors">
                          <ArrowLeft size={16} className="-rotate-90 opacity-40" />
                        </div>
                      </div>
                    </div>

                    <div className="space-y-3">
                      <label className="text-[15px] font-black text-gray-600 pr-2 uppercase tracking-widest flex items-center gap-2">
                        <Activity size={14} className="text-[#1a5e1a]/60" />
                        حالة الإنجاز
                      </label>
                      {(() => {
                        const canEditStatus = isAdmin ||
                          userSectionPerms?.editWorkStatus ||
                          (userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(editingRequest.assignedUnit));

                        return (
                          <div className="space-y-2">
                            <div className="relative group">
                              <select
                                value={editingRequest.status || ''}
                                onChange={e => setEditingRequest({
                                  ...editingRequest,
                                  status: e.target.value
                                })}
                                disabled={!canEditStatus}
                                className={`w-full px-5 py-4 rounded-2xl border-2 border-gray-100 bg-white focus:border-[#1a5e1a] focus:ring-8 focus:ring-[#1a5e1a]/5 font-black text-gray-800 transition-all appearance-none cursor-pointer ${!canEditStatus ? 'cursor-not-allowed opacity-60 bg-gray-50' : ''}`}
                              >
                                <option value="">قيد الانتظار</option>
                                {(sections.find(s => s.id === editingRequest.sectionId)?.availableStatuses || ['قيد العمل (جاري التنفيذ)', 'تم الإنجاز (منجز)', 'مرفوض / ملغي']).map(status => (
                                  <option key={status} value={status}>{status}</option>
                                ))}
                              </select>
                              <div className="absolute left-6 top-1/2 -translate-y-1/2 pointer-events-none text-gray-400 group-hover:text-[#1a5e1a] transition-colors">
                                <ArrowLeft size={16} className="-rotate-90 opacity-40" />
                              </div>
                            </div>
                            {!canEditStatus && userSectionPerms?.isUnitLead && (
                              <p className="text-[10px] text-red-500 font-bold px-2 flex items-center gap-1">
                                <span>⚠️</span> لا تملك صلاحية تعديل الحالة لهذا الطلب
                              </p>
                            )}
                          </div>
                        );
                      })()}
                    </div>
                  </div>

                  <div className="space-y-3">
                    <label className="text-[15px] font-black text-gray-600 pr-2 uppercase tracking-widest flex items-center gap-2">
                      <MessageCircle size={14} className="text-[#1a5e1a]/60" />
                      التفاصيل والتقارير الفنية
                    </label>
                    {(() => {
                      const canEditNotes = isAdmin ||
                        userSectionPerms?.editDeptNotes ||
                        (userSectionPerms?.isUnitLead && userSectionPerms?.units?.includes(editingRequest.assignedUnit));

                      return (
                        <textarea
                          value={editingRequest.sectionNotes || ''}
                          onChange={e => setEditingRequest({ ...editingRequest, sectionNotes: e.target.value })}
                          disabled={!canEditNotes}
                          rows={4}
                          className={`w-full px-6 py-4 rounded-[1.8rem] border-2 border-gray-100 focus:border-[#1a5e1a] focus:ring-8 focus:ring-[#1a5e1a]/5 font-bold resize-none bg-white transition-all ${!canEditNotes ? 'cursor-not-allowed opacity-60 bg-gray-50' : ''}`}
                          placeholder="وثق هنا الإجراءات المتخذة أو قطع الغيار المستبدلة..."
                        ></textarea>
                      );
                    })()}
                  </div>

                  <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pt-6 mt-4 border-t border-gray-100">
                    <div className="flex gap-3 w-full sm:w-auto">
                      <button type="button" onClick={() => setEditingRequest(null)} className="px-8 py-3 rounded-xl font-bold text-gray-400 hover:bg-gray-100 transition-all flex-1 sm:flex-none">إلغاء</button>
                      {isAdmin && (
                        <button
                          type="button"
                          onClick={() => deleteRequest(editingRequest.id)}
                          className="w-12 h-12 bg-red-600 text-white rounded-xl flex items-center justify-center hover:bg-red-700 transition-all shadow-lg border border-red-500"
                          title="حذف الطلب نهائياً"
                        >
                          <Trash2 size={20} />
                        </button>
                      )}
                    </div>
                    <button type="submit" className="w-full sm:w-auto px-12 py-4 bg-[#1a5e1a] hover:bg-[#2d7d2d] text-white rounded-xl font-black shadow-xl shadow-[#1a5e1a]/20 hover:shadow-[#1a5e1a]/40 hover:-translate-y-0.5 transition-all text-sm flex items-center justify-center gap-3">
                      <span>تحديث البيانات</span>
                      <ArrowLeft size={18} className="rotate-180" />
                    </button>
                  </div>
                </div>
              </motion.form>
            </div>
          );
        })()}


        {/* WhatsApp Selection Modal (Professional Engineering Support Design) */}
        {whatsappModal.isOpen && (
          <div className="fixed inset-0 bg-black/80 backdrop-blur-xl z-[200] flex items-center justify-center p-4 rtl">
            <div className="bg-white rounded-[3rem] shadow-[0_35px_60px_-15px_rgba(0,0,0,0.3)] w-full max-w-xl overflow-hidden animate-in slide-in-from-bottom-20 duration-500 border border-white/20">

              {/* Gradient Header */}
              <div className="p-8 bg-gradient-to-l from-[#0f3f0f] via-[#1a5e1a] to-[#25D366] text-white relative">
                <div className="absolute top-0 left-0 w-full h-full bg-[radial-gradient(circle_at_top_right,rgba(255,255,255,0.1),transparent_70%)]"></div>
                <div className="relative flex items-center justify-between">
                  <div className="flex items-center gap-5">
                    <div className="w-16 h-16 bg-white/20 backdrop-blur-md rounded-2xl flex items-center justify-center text-white shadow-[0_10px_25px_rgba(37,211,102,0.3)] border border-white/30 animate-pulse">
                      <MessageCircle size={32} fill="currentColor" />
                    </div>
                    <div className="text-right">
                      <h2 className="text-2xl font-black tracking-tight">إرسال إشعار WhatsApp</h2>
                      <p className="text-xs opacity-80 font-black mt-1">اسناد الصيانة الهندسي - شعبة صيانة مجموعة العميد</p>
                    </div>
                  </div>
                  <button
                    onClick={() => setWhatsappModal({ isOpen: false, request: null })}
                    className="w-10 h-10 rounded-full bg-black/20 flex items-center justify-center hover:bg-black/40 hover:scale-110 transition-all border border-white/10"
                  >
                    <span className="text-xl">✕</span>
                  </button>
                </div>
              </div>

              <div className="p-10 space-y-8 text-right max-h-[65vh] overflow-y-auto custom-scrollbar">

                {/* Request Summary Card */}
                <div className="bg-gradient-to-br from-gray-50 to-gray-100 p-6 rounded-3xl border-2 border-white shadow-inner relative group">
                  <div className="absolute top-4 left-4 bg-green-500 text-white text-[10px] px-3 py-1 rounded-full font-black shadow-md">بيانات الطلب #{whatsappModal.request?.id}</div>
                  <p className="text-[10px] font-black text-gray-400 mb-3 uppercase tracking-widest">تنسيق الرسالة المهني</p>
                  <div className="space-y-2">
                    <h4 className="font-black text-lg text-[#1a5e1a] border-b-2 border-green-200 pb-2 mb-3 inline-block">مجمع {sections.find(s => s.id === whatsappModal.request?.sectionId)?.name}</h4>
                    <div className="grid grid-cols-2 gap-4">
                      <div className="bg-white/60 p-3 rounded-2xl border border-white">
                        <span className="block text-[8px] text-gray-400 font-black">الوحدة</span>
                        <span className="font-bold text-gray-700">{whatsappModal.request?.assignedUnit}</span>
                      </div>
                      <div className="bg-white/60 p-3 rounded-2xl border border-white">
                        <span className="block text-[8px] text-gray-400 font-black">الحالة</span>
                        <span className="font-bold text-orange-600">{whatsappModal.request?.status || 'قيد الانتظار'}</span>
                      </div>
                    </div>
                    <p className="text-[11px] text-gray-500 mt-3 font-bold bg-white p-3 rounded-xl border border-gray-100 leading-relaxed">
                      <span className="text-[#1a5e1a] ml-1">📝 النوع:</span>
                      {Object.values(whatsappModal.request?.answers || {}).slice(0, 2).join(' - ')}
                    </p>
                  </div>
                </div>

                {/* Recipient Selection */}
                <div className="space-y-4">
                  <div className="flex items-center gap-3 pr-2">
                    <div className="w-2 h-6 bg-green-500 rounded-full"></div>
                    <label className="text-sm font-black text-gray-800">1. اختيار الوحدة والمسؤولين</label>
                  </div>

                  <select
                    className="w-full px-5 py-4 rounded-[1.5rem] border-2 border-gray-100 focus:border-[#25D366] focus:outline-none font-bold text-gray-700 bg-gray-50 shadow-sm transition-all"
                    onChange={(e) => setSelectedWhatsappUnit(e.target.value)}
                    value={selectedWhatsappUnit}
                  >
                    <option value="">--- اختر الوحدة الهندسية ---</option>
                    {(sections.find(s => s.id === whatsappModal.request?.sectionId)?.sectionUnits || globalUnits).map(unit => (
                      <option key={unit} value={unit}>
                        {unit} ({users.filter(u => u.permissions[whatsappModal.request?.sectionId]?.units.includes(unit)).length} متاح)
                      </option>
                    ))}
                  </select>

                  <div className="bg-gray-50 rounded-[2rem] p-4 min-h-[150px] border-2 border-white shadow-inner">
                    {!selectedWhatsappUnit ? (
                      <div className="flex flex-col items-center justify-center py-10 opacity-30">
                        <div className="text-4xl mb-2">🔍</div>
                        <p className="text-xs font-black">يرجى تحديد وحدة لعرض الكوادر</p>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        {users.filter(u => u.permissions[whatsappModal.request?.sectionId]?.units.includes(selectedWhatsappUnit)).map(u => (
                          <label
                            key={u.id}
                            className={`flex items-center gap-4 p-4 rounded-2xl border-2 transition-all cursor-pointer group shadow-sm ${selectedRecipientPhones.includes(u.phone) ? 'bg-white border-[#25D366] ring-4 ring-green-50 translate-x-[-5px]' : 'bg-white/60 border-white hover:border-green-200'}`}
                          >
                            <div className={`w-6 h-6 rounded-lg flex items-center justify-center transition-all ${selectedRecipientPhones.includes(u.phone) ? 'bg-[#25D366] text-white rotate-12 scale-110' : 'bg-gray-100'}`}>
                              <input
                                type="checkbox"
                                className="sr-only"
                                checked={selectedRecipientPhones.includes(u.phone)}
                                onChange={() => {
                                  if (selectedRecipientPhones.includes(u.phone)) {
                                    setSelectedRecipientPhones(selectedRecipientPhones.filter(p => p !== u.phone));
                                  } else {
                                    setSelectedRecipientPhones([...selectedRecipientPhones, u.phone]);
                                  }
                                }}
                              />
                              {selectedRecipientPhones.includes(u.phone) ? '✓' : ''}
                            </div>
                            <div className="flex-1 flex justify-between items-center">
                              <div>
                                <p className="font-black text-sm text-gray-800">{u.name}</p>
                                <p className="text-[10px] text-gray-400 font-mono">{u.phone || '000-000-000'}</p>
                              </div>
                              <span className={`text-[8px] px-3 py-1 rounded-full font-black ${u.phone ? 'bg-green-100 text-green-700' : 'bg-red-50 text-red-600'}`}>
                                {u.phone ? 'متاح واتساب' : 'رقم مفقود'}
                              </span>
                            </div>
                          </label>
                        ))}
                        {users.filter(u => u.permissions[whatsappModal.request?.sectionId]?.units.includes(selectedWhatsappUnit)).length === 0 && (
                          <div className="text-center py-8">
                            <p className="text-xs font-bold text-gray-400 italic">لا يوجد مسؤولين مسجلين لهذه الوحدة حالياً</p>
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                </div>

                {/* Notes Section */}
                <div className="space-y-3">
                  <div className="flex items-center gap-3 pr-2">
                    <div className="w-2 h-6 bg-orange-400 rounded-full"></div>
                    <label className="text-sm font-black text-gray-800">2. الإسناد الإضافي والتعليمات</label>
                  </div>
                  <textarea
                    rows={3}
                    className="w-full px-6 py-4 rounded-[1.5rem] border-2 border-gray-100 focus:border-[#25D366] bg-gray-50 focus:bg-white focus:outline-none font-bold text-sm resize-none shadow-sm transition-all"
                    placeholder="توجيهات فنية، رتبة الأولوية، أو ملاحظات للمسؤول..."
                    value={whatsappNote}
                    onChange={(e) => setWhatsappNote(e.target.value)}
                  ></textarea>
                </div>
              </div>

              {/* Footer Actions */}
              <div className="p-8 bg-gray-50 border-t border-gray-100/50 flex gap-4">
                <button
                  onClick={() => setWhatsappModal({ isOpen: false, request: null })}
                  className="flex-1 py-4 bg-white text-gray-500 rounded-2xl font-black border-2 border-gray-200 hover:bg-gray-100 transition-all"
                >
                  إلغاء ✖️
                </button>
                <button
                  onClick={handleSendAdvancedWhatsApp}
                  disabled={selectedRecipientPhones.length === 0}
                  className={`flex-[2] py-4 text-white rounded-2xl font-black shadow-lg transition-all relative overflow-hidden group ${selectedRecipientPhones.length > 0 ? 'bg-gradient-to-r from-[#128C7E] to-[#25D366] hover:-translate-y-1 hover:shadow-green-200/50' : 'bg-gray-300 opacity-50 cursor-not-allowed'}`}
                >
                  <div className="relative z-10 flex items-center justify-center gap-3 uppercase tracking-wider">
                    <span>إرسال عبر واتساب</span>
                    <MessageCircle size={20} fill="currentColor" />
                  </div>
                  <div className="absolute inset-0 bg-white/20 translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-1000 skew-x-[-20deg]"></div>
                </button>
              </div>
            </div>
          </div>
        )}
      </>
    );
  }

  return (
    <div className="login-wrapper min-h-screen w-full flex items-center justify-center relative overflow-hidden"
      style={{
        backgroundImage: "url('https://i.ibb.co/DfMCvycX/image.jpg')",
        backgroundSize: 'cover',
        backgroundPosition: 'center',
        backgroundAttachment: 'fixed'
      }}>

      {/* Overlay Filter */}
      <div className="absolute inset-0 bg-black/50 backdrop-blur-[8px] z-0"></div>

      {/* Login Card */}
      <div className="login-card relative w-[90%] max-w-[450px] bg-[#073b03]/15 backdrop-blur-[20px] rounded-[40px] p-8 md:p-10 shadow-[0_25px_50px_-12px_rgba(255,255,255,0.5)] border border-white/70 z-10 animate-in fade-in zoom-in duration-700">

        {/* Header Section */}
        <div className="flex flex-col items-center mb-8 text-center">
          <div className="w-[120px] h-[120px] md:w-[150px] md:h-[150px] rounded-full overflow-hidden border-4 border-white shadow-2xl mb-4 transition-transform hover:scale-105 hover:rotate-6">
            <img
              src="https://i.ibb.co/0jpV8nXS/image.png"
              alt="Maintenance Support Logo"
              className="w-full h-full object-cover"
            />
          </div>
          <h1 className="text-3xl md:text-[2rem] font-black text-[#eeeeee] mb-1 drop-shadow-[2px_2px_4px_rgba(255,0,0,1)]">
            اسناد الصيانة الهندسي
          </h1>
          <h2 className="bg-white/10 text-white px-5 py-1 rounded-full text-lg md:text-[1.3rem] font-black mb-3">
            العتبة العباسية المقدسة
          </h2>
          <p className="bg-[#1a5e1a] text-white px-4 py-1.5 rounded-full text-xs md:text-[0.95rem] font-bold">
            قسم المشاريع الهندسية - شعبة صيانة مجموعة العميد
          </p>
        </div>

        {/* Form Section */}
        <form onSubmit={handleLogin} className="space-y-6">
          <div className="relative group">
            <div className="absolute right-4 top-1/2 -translate-y-1/2 text-[#1a5e1a] group-focus-within:scale-110 transition-transform">
              <User size={18} />
            </div>
            <input
              ref={usernameRef}
              type="text"
              placeholder="اسم المستخدم"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="w-full pr-12 pl-4 py-3 bg-white/95 border-2 border-gray-200 rounded-2xl focus:outline-none focus:border-[#1a5e1a] transition-all font-semibold text-[#2c3e50]"
            />
          </div>

          <div className="relative group">
            <div className="absolute right-4 top-1/2 -translate-y-1/2 text-[#1a5e1a] group-focus-within:scale-110 transition-transform">
              <Lock size={18} />
            </div>
            <input
              type={showPassword ? 'text' : 'password'}
              placeholder="كلمة المرور"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full pr-12 pl-12 py-3 bg-white/95 border-2 border-gray-200 rounded-2xl focus:outline-none focus:border-[#1a5e1a] transition-all font-semibold text-[#2c3e50]"
            />
            <button
              type="button"
              onClick={() => setShowPassword(!showPassword)}
              className="absolute left-4 top-1/2 -translate-y-1/2 text-[#1a5e1a] hover:scale-110 transition-transform focus:outline-none"
            >
              {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
            </button>
          </div>

          <button
            type="submit"
            className="w-full py-4 bg-gradient-to-r from-[#0f3f0f] to-[#2d7d2d] text-white rounded-2xl font-bold text-lg flex items-center justify-center gap-2 hover:-translate-y-1 hover:shadow-[0_10px_30px_rgba(26,94,26,0.4)] active:translate-y-0 transition-all relative overflow-hidden group shadow-lg"
          >
            <span>تسجيل الدخول</span>
            <ArrowLeft size={20} />
            <div className="absolute inset-0 bg-white/10 translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-1000 skew-x-[-20deg]"></div>
          </button>
        </form>

        {/* Footer Section */}
        <div className="mt-8 flex flex-col items-center gap-4">
          <a
            href="https://wa.me/9647735559707"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-3 bg-gradient-to-br from-[#0f3f0f] to-[#128C7E] px-6 py-2 rounded-full shadow-[0_5px_20px_rgba(37,211,102,0.3)] hover:scale-105 hover:shadow-[0_10px_30px_rgba(37,211,102,0.5)] transition-all group"
          >
            <div className="w-8 h-8 bg-white rounded-full flex items-center justify-center text-[#25D366] group-hover:animate-pulse">
              <MessageCircle size={20} fill="currentColor" />
            </div>
            <div className="text-white text-right">
              <p className="text-sm font-bold leading-tight">الدعم التقني</p>
            </div>
          </a>

          <div className="flex items-center gap-2 text-white/50 text-[10px] md:text-xs">
            <span className="border-l border-white/30 pl-2">DESIGN BY MURTADA AL_JANABY</span>
            <span>© 2026</span>
          </div>
        </div>

        {/* Error Message */}
        {error && (
          <div className="fixed bottom-10 left-1/2 -translate-x-1/2 bg-red-500 text-white px-6 py-3 rounded-full flex items-center gap-2 shadow-2xl animate-in slide-in-from-bottom-5">
            <Lock size={16} />
            <span>اسم المستخدم أو كلمة المرور غير صحيحة!</span>
          </div>
        )}
      </div>
    </div>
  );
}
