// Небольшой набор линейных иконок (inline SVG — без зависимостей).
const base = { width: 16, height: 16, viewBox: "0 0 24 24", fill: "none", stroke: "currentColor",
  strokeWidth: 2, strokeLinecap: "round", strokeLinejoin: "round", "aria-hidden": true };

export const IconSend = () => (
  <svg {...base}><path d="M22 2 11 13" /><path d="M22 2 15 22l-4-9-9-4 20-7z" /></svg>
);
export const IconClock = () => (
  <svg {...base}><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></svg>
);
export const IconFlame = () => (
  <svg {...base}><path d="M12 22c4 0 7-3 7-7 0-5-4-7-5-11-2 3-2 5-4 6-1-1-1-3-2-4-2 3-3 6-3 9 0 4 3 7 7 7z" /></svg>
);
export const IconCopy = () => (
  <svg {...base}><rect x="9" y="9" width="12" height="12" rx="2" /><path d="M5 15V5a2 2 0 0 1 2-2h10" /></svg>
);
export const IconBack = () => (
  <svg {...base}><path d="M15 18 9 12l6-6" /></svg>
);
export const IconRefresh = () => (
  <svg {...base}><path d="M21 12a9 9 0 1 1-3-6.7L21 8" /><path d="M21 3v5h-5" /></svg>
);
export const IconEye = () => (
  <svg {...base}><path d="M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12z" /><circle cx="12" cy="12" r="3" /></svg>
);
export const IconWarn = () => (
  <svg {...base}><path d="M12 3 2 21h20L12 3z" /><path d="M12 10v5M12 18v.5" /></svg>
);
export const IconCheck = () => (
  <svg {...base}><path d="m5 12 5 5 9-10" /></svg>
);
export const IconX = () => (
  <svg {...base}><path d="M18 6 6 18M6 6l12 12" /></svg>
);
export const IconPlus = () => (
  <svg {...base}><path d="M12 5v14M5 12h14" /></svg>
);
export const IconAntenna = () => (
  <svg {...base}><path d="M5 12a7 7 0 0 1 14 0M8.5 12a3.5 3.5 0 0 1 7 0" /><path d="M12 12v10" /></svg>
);
// знак SQANDAL: динамик с двумя волнами (как «🔊 SQANDAL» в подписи постов)
export const IconSpeaker = ({ size = 16 }) => (
  <svg {...base} width={size} height={size}>
    <path d="M4 9.5h3.5L13 5v14l-5.5-4.5H4z" />
    <path d="M16.5 9.2a4 4 0 0 1 0 5.6M19.3 6.4a8 8 0 0 1 0 11.2" />
  </svg>
);
export const IconSearch = () => (
  <svg {...base}><circle cx="11" cy="11" r="6.5" /><path d="m20 20-4.2-4.2" /></svg>
);
export const IconChevron = () => (
  <svg {...base}><path d="m6 9 6 6 6-6" /></svg>
);
export const IconInfo = () => (
  <svg {...base}><circle cx="12" cy="12" r="9" /><path d="M12 11v6M12 7.5v.5" /></svg>
);
